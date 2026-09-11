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
  code: z
    .string()
    .describe(
      "Python source code to execute. " +
      "Stdlib + pandas + numpy + scipy + matplotlib preinstalled. " +
      "No network. /tmp is a 64 MB tmpfs (writable). " +
      "Save plots with plt.savefig('/tmp/plot.png'), then read and base64-encode to return image data."
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
  stdoutBytes: number;
  runtime?: string;
}

export const defaultCodeExecConfig: CodeExecConfig = {
  image: "code-exec-sandbox:latest",
  wallClockMs: 10000,
  memoryMB: 256,
  cpus: 1.0,
  stdoutBytes: 512 * 1024,
  runtime: "runsc",
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
      "--tmpfs=/tmp:size=64m",
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
      "No network. Stdlib + pandas + numpy + scipy + matplotlib. " +
      "Use for data analysis, calculations, and chart generation. " +
      "Save plots to /tmp and base64-encode them for image output. " +
      "Output capped at 512 KB stdout + stderr.",
    inputSchema,
    handler: async ({ code }) => {
      if (!isEnabled()) throw new Error("code_exec: tool is disabled");

      code = code.trim();
      if (!code) throw new Error("code_exec: code is required");
      if (Buffer.byteLength(code, "utf-8") > 256 * 1024) {
        throw new Error("code_exec: code too large (max 256 KB)");
      }

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
      }

      return {
        content: [{ type: "text" as const, text: JSON.stringify({ text, meta }, null, 2) }],
        isError: runError !== undefined,
      };
    },
  };
}

export const codeExecTool = createCodeExecTool();
