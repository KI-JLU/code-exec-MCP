# code-exec-MCP

An MCP (Model Context Protocol) server that executes Python code in an isolated gVisor sandbox. Supports multiple concurrent clients via Streamable HTTP transport.

## Quick Start

```bash
npm install
npm run build
docker build -f Dockerfile.sandbox -t code-exec-sandbox:latest .
npm start
```

## Project Structure

```
src/
├── index.ts              # Entry point — transport selection, session management
├── server.ts             # Server setup & tool registration
└── tools/
    └── code_exec.ts      # Python sandbox execution tool
Dockerfile                # MCP server image
Dockerfile.sandbox        # Python sandbox image (scientific stack, python-pptx + hawki_slides, node + pptxgenjs, LibreOffice)
sandbox/sitecustomize.py  # In every sandbox Python: self-explaining subprocess errors, documents in /tmp delivered at exit
sandbox/hawki_slides/     # Deck helper on python-pptx: JLU template, drawn styles, attached templates
sandbox/templates/        # JLU-de.potx, JLU-en.potx (copies of templates/)
docker-compose.yml        # Production deployment
docker-entrypoint.sh      # Container startup — fixes .tmp ownership for mcp user
```

## Tool: `code_exec`

Executes a Python snippet in a sandboxed gVisor container.

- No network access
- Read-only filesystem (`/tmp` is writable, 64 MB tmpfs)
- Libraries: `numpy`, `pandas`, `scipy`, `matplotlib`, `python-pptx`; `hawki_slides` for decks; `node` + `pptxgenjs`
- Input files: `files: [{ name, content_base64 }]` on the call, read-only at `/work/<name>`
- Binaries: `soffice` (LibreOffice Impress) and `pdftoppm`, to render a deck back to slide images
- Timeout: 10 seconds
- Output cap: 512 KB stdout + stderr
- Every call is a fresh process and a fresh `/tmp`: nothing carries over between calls
- stderr comes back too: appended under `[stderr]` when the snippet exited cleanly but something it ran did not (a node script failing inside a `try`), in front of stdout when the snippet itself failed
- A failed subprocess explains itself: `print(e)` on a `CalledProcessError` includes the captured stderr (`sandbox/sitecustomize.py`)
- Documents are delivered by leaving them in `/tmp`: every `.pptx`, `.docx`, `.xlsx`, `.csv` and `.pdf` there is printed as a named data URI when the run ends, unless the program printed it already; a PDF next to a `.pptx`/`.docx` of the same name is treated as LibreOffice's conversion step and skipped

### Building a PowerPoint deck

`hawki_slides` ([sandbox/hawki_slides](sandbox/hawki_slides)) is the way in: a
deck in a dozen Python calls on the JLU corporate template (German or English
by language), one of HAWKI's own drawn styles, or a template the user attached.
Layout, fonts and the logo come from the template; the program that writes
the deck does not have to know python-pptx, whose API is exactly what models
got wrong on the first try.

```python
from hawki_slides import Deck

deck = Deck(title="Anthropomorphisierung von LLM", author="HAWKI", lang="de")   # JLU template, German
deck.title("Anthropomorphisierung von LLM", "Warum wir Sprachmodelle vermenschlichen")
deck.bullets("Was bedeutet das?", ["Punkt eins", {"text": "Punkt zwei", "sub": ["Detail"]}], sources=["https://..."])
deck.cards("Vier Signale", [{"heading": "Dialog", "text": "..."}, {"heading": "Ich-Perspektive", "text": "..."}])
deck.two_columns("Hilfreich vs. irreführend", {"heading": "Hilfreich", "items": ["..."]}, {"heading": "Irreführend", "items": ["..."]})
deck.quote("Menschlich genug, um zu helfen.", "Fazit")
deck.closing("Danke · Fragen?")
deck.save("/tmp/anthropomorphisierung.pptx")   # delivered when the run ends - nothing to print
```

Where the look comes from:

| | |
|---|---|
| `Deck(lang="de")` (default) | JLU template, German (`sandbox/templates/JLU-de.potx`) |
| `Deck(lang="en")` | JLU template, English |
| `Deck(template="attached")` | the `.potx`/`.pptx` the caller passed in `files` (at `/work/<name>`), as a template: layouts kept, slides dropped |
| `Deck(style="purple")` | HAWKI's drawn style; also `blue`, `green`, `red`, `slate` |
| `Deck.open("/work/<name>.pptx")` | continue a deck built earlier: its slides stay, new ones go on its layouts (`style=` for a drawn deck) |

Methods: `title`, `bullets`, `cards` (2-6), `two_columns`, `quote`, `closing`,
`image(title, path, caption)`, `save` (a `/tmp` path). Each slide method takes
`notes=` and `sources=`, both land in the speaker notes. `deck.raw` is the
python-pptx `Presentation` for anything else. A template's sample slides are
dropped, but the pictures on them - a corporate template often keeps its logo
there rather than on the layout - are carried onto the slides the deck adds.

`image()` takes a bitmap or an SVG - the SVG is rasterised first (`rsvg-convert`,
else LibreOffice Draw). A bare file name is looked up in `/work`, so a picture
the caller passed in `files` is `image("...", "otter.png")`. `save()` writes the
deck title and author into the file's core properties, which is what
`Deck.open()` reads back.

### Passing files in

`code_exec` takes an optional `files` array (`[{ name, content_base64 }]`, at
most 20 files and 25 MB); each appears read-only at `/work/<name>` next to the
program. That is how an attached template, or a CSV to analyse, reaches the
sandbox.

Raw pptxgenjs (node) and python-pptx are installed too, for the rare slide the
helper cannot express.

### Checking a deck before returning it

pptxgenjs lays slides out in absolute inches, so nothing but a look at the result
catches a text box placed on top of another one. Convert and rasterise in the
same call - the sandbox filesystem does not survive between calls - and read the
PNGs back as images:

```python
import subprocess, glob

subprocess.run(["soffice", "--headless", "--convert-to", "pdf",
                "--outdir", "/tmp", "/tmp/deck.pptx"], check=True)
subprocess.run(["pdftoppm", "-png", "-r", "80", "/tmp/deck.pdf", "/tmp/slide"], check=True)
print(sorted(glob.glob("/tmp/slide*.png")))
```

Build, convert and rasterise for nine slides measures ~0.7 s of the 10 s budget
(without `runsc`; gVisor adds syscall overhead on top).

### Returning a plot

```python
import matplotlib.pyplot as plt
import base64

plt.plot([1, 2, 3])
plt.savefig("/tmp/plot.png", dpi=90)
plt.close()

with open("/tmp/plot.png", "rb") as f:
    b64 = base64.b64encode(f.read()).decode()
print(f"data:image/png;base64,{b64}")
```

## Transport Modes

### stdio (default)

```bash
npm start
```

### Streamable HTTP

```bash
npm run start:http
```

Endpoint: `http://localhost:3001/mcp`  
Health check: `http://localhost:3001/health`

## Docker Deployment

The sandbox image must be built on the host before starting the stack, because the MCP server spawns sandbox containers directly via the Docker socket:

```bash
docker build -f Dockerfile.sandbox -t code-exec-sandbox:latest .
docker compose up -d
```

The MCP endpoint will be available at `http://localhost:3001/mcp`.  
Health check: `http://localhost:3001/health` → `{"status":"ok","uptime_s":…,"sessions":…}`

## Environment Variables

| Variable | Default | Description |
|----------|---------|-------------|
| `MCP_TRANSPORT` | `stdio` | Transport mode: `stdio` or `http` |
| `MCP_PORT` | `3001` | HTTP port for Streamable HTTP transport |
| `SANDBOX_IMAGE` | `code-exec-sandbox:latest` | Image every snippet runs in |
| `SANDBOX_RUNTIME` | `runsc` | Container runtime for the sandbox; set to an empty string for the daemon's default runtime on a machine without gVisor (local development only) |
| `SANDBOX_HOST_TMP` | `<cwd>/.tmp` | Host path of the `.tmp` bind mount, needed when this server itself runs in a container (Docker-out-of-Docker) |
| `SANDBOX_WALL_CLOCK_MS` | `10000` | Hard kill after this long. A deck that is built, converted by LibreOffice and rasterised in one run needs far more than a plot does; HAWKI allows this tool 120 s, so stay under that |
| `SANDBOX_MEMORY_MB` | `256` | Memory limit per run (swap is pinned to the same value). LibreOffice does not fit in the default |
| `SANDBOX_CPUS` | `1.0` | CPU limit per run |
| `SANDBOX_TMPFS_MB` | `64` | Size of the writable `/tmp`, which holds a generated deck, its PDF and the slide PNGs |
| `SANDBOX_STDOUT_BYTES` | `524288` | Cap on captured stdout and stderr |

An unreadable or non-positive value for any of the numeric knobs is ignored with a line on stderr, and the default applies.

## Second server for a new sandbox image

The image with the document toolchain changes behaviour for every caller: a `.pptx`, `.docx`, `.xlsx`, `.csv` or `.pdf` left in `/tmp` is delivered as a data URI when the run ends. A client that does not know that convention - a HAWKI older than the code interpreter's file delivery, or any other consumer on the gateway - passes that base64 straight to its model. So the new image does not replace the old one under the running server. It gets a second server on another port, reached through a gateway alias of its own, while production keeps `:3001` and `code-exec-sandbox:latest`:

```bash
# in a directory of its own, so the production checkout is untouched
docker build -f Dockerfile.sandbox -t code-exec-sandbox:pptx .
docker compose -p code-exec-mcp-next -f docker-compose.next.yml up -d --build
```

`docker-compose.next.yml` runs on port 3002, points `SANDBOX_IMAGE` at the new tag and raises the limits the document toolchain needs. Then add a second MCP server entry in the gateway that points at `http://<host>:3002/mcp`, and the client picks the image by which alias it calls - in HAWKI, `HAWKI_CODE_EXEC_MCP_SERVER`.

**Behind a proxy** (a host without a direct route to the internet) both builds need it passed in; the Dockerfiles declare `http_proxy`/`https_proxy` as build args, so nothing is baked into the image:

```bash
docker build -f Dockerfile.sandbox -t code-exec-sandbox:pptx \
  --build-arg http_proxy=http://10.60.3.254:3128 \
  --build-arg https_proxy=http://10.60.3.254:3128 .

BUILD_HTTP_PROXY=http://10.60.3.254:3128 BUILD_HTTPS_PROXY=http://10.60.3.254:3128 \
  docker compose -p code-exec-mcp-next -f docker-compose.next.yml up -d --build
```

**Status 2026-09-17.** HAWKI production (ki-chat) took path A on 2026-09-15: its `code-exec-mcp` gateway alias was switched from `mcp_gVisor` (`:3001`, `code-exec-sandbox:latest`) to `mcp_gVisor_next` (`:3002`, `code-exec-sandbox:pptx`) together with the HAWKI release v2.3.2.5 that understands the file delivery. Both HAWKI hosts now share the `:3002` server; `:3001` keeps running for the other gateway keys that still hold `mcp_gVisor` (KANBAN MCP, kidevlokal_june26, VS-Code). Whether `:3001` is retired or promoted to the new image is still open (Kanban KI-755 / KI-760).

## Session Management

In HTTP mode each client gets an isolated `McpServer` instance identified by `Mcp-Session-Id` header. Sessions are evicted after **30 minutes** of inactivity. Active session count is visible in `/health`.

## Security

Each code execution runs in a separate gVisor container:

- `--runtime=runsc` — gVisor kernel isolation
- `--network=none` — no network access
- `--read-only` — read-only root filesystem
- `--cap-drop=ALL` — no Linux capabilities
- `--security-opt=no-new-privileges` — prevents privilege escalation
- `--memory=256m` — hard memory limit
- `--cpus=1.0` — hard CPU limit
- `--pids-limit=64` — prevents fork bombs

## License

MIT
