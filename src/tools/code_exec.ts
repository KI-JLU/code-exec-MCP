import { z } from "zod";
import * as fs from "node:fs/promises";
import * as path from "node:path";
import { execFile } from "node:child_process";
import { promisify } from "node:util";
import type { ToolDefinition } from "../server.js";

const execFileAsync = promisify(execFile);

// ---------------------------------------------------------------------------
// Schema
// ---------------------------------------------------------------------------

const inputSchema = z.object({
  files: z
    .array(
      z.object({
        name: z.string().describe("File name; it is available in the sandbox at /work/<name>."),
        content_base64: z.string().describe("The file's bytes, base64 encoded."),
      })
    )
    .max(20)
    .optional()
    .describe(
      "Files to place next to the program, read-only at /work/<name> - a template the user attached, " +
      "a data file to analyse. At most 20 files and 25 MB in total."
    ),
  code: z
    .string()
    .describe(
      "Python source code to execute. " +
      "Stdlib + pandas + numpy + scipy + matplotlib preinstalled. " +
      "No network. /tmp is a 64 MB tmpfs (writable); the working directory /work is not - it holds this program and the files passed in `files`. " +
      "Save plots with plt.savefig('/tmp/plot.png'), then read and base64-encode to return image data. " +
      "For PowerPoint decks, node and pptxgenjs are installed: write the deck script to /tmp and run it " +
      "with subprocess.run(['node', '/tmp/deck.js']), then base64-encode the .pptx it wrote. " +
      "require('pptxgenjs') resolves from anywhere - NODE_PATH is set."
    ),
});

// ---------------------------------------------------------------------------
// Config & interfaces
// ---------------------------------------------------------------------------

export interface CodeExecConfig {
  image: string;
  wallClockMs: number;
  memoryMB: number;
  cpus: number;
  /** Size of the writable /tmp inside the sandbox; a deck, its PDF and the slide PNGs live there. */
  tmpfsMB: number;
  stdoutBytes: number;
  runtime?: string;
}

/**
 * The runtime the sandbox container runs under. gVisor (`runsc`) in production;
 * SANDBOX_RUNTIME="" selects the daemon's default runtime for a machine that
 * has no gVisor - Docker Desktop, a developer's laptop - where the isolation
 * is the developer's own and the point is to exercise the image.
 */
/**
 * A numeric limit from the environment, so a deployment can be tuned without a
 * rebuild. The document toolchain is the reason this exists: LibreOffice needs
 * noticeably more time and memory than a matplotlib plot, and how much more
 * depends on the host - two slow cores under gVisor are not a developer's
 * laptop. An unreadable or non-positive value falls back rather than crippling
 * the sandbox.
 */
function numberFromEnv(name: string, fallback: number): number {
  const raw = process.env[name];
  if (raw === undefined || raw.trim() === "") return fallback;

  const value = Number(raw);
  if (!Number.isFinite(value) || value <= 0) {
    console.error(`code_exec: ignoring invalid ${name}="${raw}", using ${fallback}`);
    return fallback;
  }

  return value;
}

function runtimeFromEnv(): string | undefined {
  const raw = process.env.SANDBOX_RUNTIME;
  if (raw === undefined) return "runsc";
  const trimmed = raw.trim();
  return trimmed === "" ? undefined : trimmed;
}

export const defaultCodeExecConfig: CodeExecConfig = {
  image: process.env.SANDBOX_IMAGE?.trim() || "code-exec-sandbox:latest",
  wallClockMs: numberFromEnv("SANDBOX_WALL_CLOCK_MS", 10000),
  memoryMB: numberFromEnv("SANDBOX_MEMORY_MB", 256),
  cpus: numberFromEnv("SANDBOX_CPUS", 1.0),
  tmpfsMB: numberFromEnv("SANDBOX_TMPFS_MB", 64),
  stdoutBytes: numberFromEnv("SANDBOX_STDOUT_BYTES", 512 * 1024),
  runtime: runtimeFromEnv(),
};

export interface RunnerResult {
  stdout: string;
  stderr: string;
  runError?: Error;
}

export interface CodeExecRunner {
  run(codeDir: string, cfg: CodeExecConfig, abortSignal?: AbortSignal): Promise<RunnerResult>;
}

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

function capBytes(b: string | Buffer, cap: number): string {
  const buf = Buffer.isBuffer(b) ? b : Buffer.from(b, "utf-8");
  if (cap <= 0 || buf.length <= cap) return buf.toString("utf-8");
  return Buffer.concat([buf.subarray(0, cap), Buffer.from("\n[truncated]")]).toString("utf-8");
}

const MAX_INPUT_FILES_BYTES = 25 * 1024 * 1024;

/**
 * Files the caller wants next to the program. They are written into the same
 * directory as code.py, which is bind-mounted read-only at /work, so a file
 * named deck.potx is /work/deck.potx inside the sandbox. Names are reduced to a
 * basename - a caller does not get to choose where in the directory tree its
 * bytes land - and "code.py" is refused, that name is the program's.
 */
function decodeInputFiles(files: { name: string; content_base64: string }[]): { name: string; bytes: Buffer }[] {
  const seen = new Set<string>();
  let total = 0;
  const decoded: { name: string; bytes: Buffer }[] = [];

  for (const file of files) {
    const name = path.basename(file.name.trim());
    if (!name || name === "." || name === ".." || name.startsWith(".") || name === "code.py") {
      throw new Error(`code_exec: invalid file name "${file.name}"`);
    }
    if (seen.has(name)) throw new Error(`code_exec: duplicate file name "${name}"`);
    seen.add(name);

    const bytes = Buffer.from(file.content_base64, "base64");
    total += bytes.length;
    if (total > MAX_INPUT_FILES_BYTES) {
      throw new Error(`code_exec: files too large (max ${MAX_INPUT_FILES_BYTES / 1024 / 1024} MB in total)`);
    }
    decoded.push({ name, bytes });
  }

  return decoded;
}

// ---------------------------------------------------------------------------
// Docker runner
// ---------------------------------------------------------------------------

export class DockerRunner implements CodeExecRunner {
  async run(codeDir: string, cfg: CodeExecConfig, abortSignal?: AbortSignal): Promise<RunnerResult> {
    const args = [
      "run", "--rm",
      "--network=none",
      "--read-only",
      "--cap-drop=ALL",
      "--security-opt=no-new-privileges",
      `--memory=${cfg.memoryMB}m`,
      `--memory-swap=${cfg.memoryMB}m`,
      `--cpus=${cfg.cpus}`,
      `--tmpfs=/tmp:size=${cfg.tmpfsMB}m`,
      "--pids-limit=64",
      "--user=sandboxuser",
      "-v", `${codeDir}:/work:ro`,
      "-w", "/work",
    ];

    if (cfg.runtime) args.push(`--runtime=${cfg.runtime}`);
    args.push(cfg.image, "python", "/work/code.py");

    try {
      const { stdout, stderr } = await execFileAsync("docker", args, {
        maxBuffer: cfg.stdoutBytes * 2,
        signal: abortSignal,
      });
      return {
        stdout: capBytes(stdout, cfg.stdoutBytes),
        stderr: capBytes(stderr, cfg.stdoutBytes),
      };
    } catch (error: unknown) {
      const err = error as NodeJS.ErrnoException & { stdout?: string; stderr?: string };
      return {
        stdout: err.stdout ? capBytes(err.stdout, cfg.stdoutBytes) : "",
        stderr: err.stderr ? capBytes(err.stderr, cfg.stdoutBytes) : "",
        runError: err,
      };
    }
  }
}

// ---------------------------------------------------------------------------
// Tool factory
// ---------------------------------------------------------------------------

const defaultRunner = new DockerRunner();
const defaultIsEnabled = () => true;

export function createCodeExecTool(
  cfg: CodeExecConfig = defaultCodeExecConfig,
  runner: CodeExecRunner = defaultRunner,
  isEnabled: () => boolean = defaultIsEnabled
): ToolDefinition<typeof inputSchema.shape> {
  return {
    name: "code_exec",
    description:
      "Execute Python code in a sandboxed gVisor container. " +
      "No network. Stdlib + pandas + numpy + scipy + matplotlib, plus node and pptxgenjs " +
      "for building PowerPoint decks. " +
      "Use for data analysis, calculations, chart generation and .pptx authoring. " +
      "Save plots and decks to /tmp and base64-encode them for binary output. " +
      "Output capped at 512 KB stdout + stderr.",
    inputSchema,
    handler: async ({ code, files }) => {
      if (!isEnabled()) throw new Error("code_exec: tool is disabled");

      code = code.trim();
      if (!code) throw new Error("code_exec: code is required");
      if (Buffer.byteLength(code, "utf-8") > 256 * 1024) {
        throw new Error("code_exec: code too large (max 256 KB)");
      }

      const inputFiles = decodeInputFiles(files ?? []);

      const baseTmpDir = path.join(process.cwd(), ".tmp");
      await fs.mkdir(baseTmpDir, { recursive: true });
      const dir = await fs.mkdtemp(path.join(baseTmpDir, "sandbox-"));
      // mkdtemp creates the dir 0700, but the sandbox container runs as a
      // different uid (sandboxuser) and must traverse it to read code.py.
      await fs.chmod(dir, 0o755);

      // When this server itself runs inside a container sharing the host's
      // Docker socket (Docker-out-of-Docker), the `docker run -v` source is
      // resolved by the HOST daemon, so it must be a host path — not this
      // container's path. SANDBOX_HOST_TMP is the host dir bind-mounted to
      // baseTmpDir; falls back to baseTmpDir when running directly on a host.
      const hostBaseTmp = process.env.SANDBOX_HOST_TMP || baseTmpDir;
      const hostDir = path.join(hostBaseTmp, path.basename(dir));

      let stdout = "", stderr = "", runError: Error | undefined;
      let durationMs = 0, timedOut = false;

      try {
        await fs.writeFile(path.join(dir, "code.py"), code, { mode: 0o644 });
        for (const file of inputFiles) {
          await fs.writeFile(path.join(dir, file.name), file.bytes, { mode: 0o644 });
        }

        const startTime = Date.now();
        const ac = new AbortController();
        const timeoutId = setTimeout(() => { ac.abort(); timedOut = true; }, cfg.wallClockMs);

        try {
          const result = await runner.run(hostDir, cfg, ac.signal);
          stdout = result.stdout;
          stderr = result.stderr;
          runError = result.runError;
        } finally {
          clearTimeout(timeoutId);
          durationMs = Date.now() - startTime;
        }
      } finally {
        await fs.rm(dir, { recursive: true, force: true });
      }

      const meta: Record<string, unknown> = {
        duration_ms: durationMs,
        stdout_len: Buffer.byteLength(stdout, "utf-8"),
        stderr_len: Buffer.byteLength(stderr, "utf-8"),
        image: cfg.image,
        runtime: cfg.runtime,
        timed_out: timedOut,
      };

      let text = stdout;
      if (timedOut) {
        meta.exit_error = `Execution timed out after ${cfg.wallClockMs} ms`;
        text = `Execution timed out after ${cfg.wallClockMs} ms`;
      } else if (runError) {
        meta.exit_error = runError.message;
        text = stderr ? `${stderr}\n---\n${stdout}` : stdout;
      } else if (stderr.trim() !== "") {
        // A clean exit with something on stderr is usually a subprocess that
        // failed inside the snippet - node refusing to write, soffice choking on
        // a file - which the snippet caught and turned into a generic message.
        // Without the stderr the model sees "exit status 1" and nothing else,
        // and its next attempt is a guess.
        text = `${stdout}\n[stderr]\n${stderr}`;
      }

      return {
        content: [{ type: "text" as const, text: JSON.stringify({ text, meta }, null, 2) }],
        isError: runError !== undefined,
      };
    },
  };
}

export const codeExecTool = createCodeExecTool();
