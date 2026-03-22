# PDFForge Tool

PDFForge Tool is a FastAPI-based PDF utility web app with a polished single-page frontend and multiple document workflows in one place.

It includes:

- PDF merge
- PDF split
- PDF to image
- Scan images to PDF
- Add logo
- Add page numbers
- Redact PDF
- PDF summary
- PDF to Word / Excel / PowerPoint

## Highlights

- FastAPI backend with a frontend served from the same app
- Hosting-friendly ASGI entrypoint via `app.py`
- Vercel-ready configuration
- Works on Render, Railway, Docker, and VPS hosting
- Runtime storage auto-detects a writable folder
- Browser auto-open is disabled automatically on hosted environments
- CORS and runtime directory can be configured with environment variables

## Project Structure

```text
.
├── app.py
├── pdf-tool.py
├── index.html
├── requirements.txt
├── vercel.json
├── Dockerfile
├── .env.example
└── README.md
```

## Local Development

### Requirements

- Python 3.12 recommended
- `pip`

### Install

```bash
pip install -r requirements.txt
```

### Run locally

```bash
python pdf-tool.py
```

Or with Uvicorn:

```bash
uvicorn app:app --host 0.0.0.0 --port 8001 --reload
```

Open:

```text
http://localhost:8001
```

## Environment Variables

You can copy `.env.example` and set values as needed.

| Variable | Purpose | Default |
| --- | --- | --- |
| `PDF_TOOL_HOST` | Host for local/server process | `0.0.0.0` |
| `PDF_TOOL_PORT` | Local fallback port | `8001` |
| `PORT` | Hosting platform port override | platform-defined |
| `PDF_TOOL_OPEN_BROWSER` | Auto-open browser for local runs | auto |
| `PDF_TOOL_CORS_ORIGINS` | Comma-separated CORS origins or `*` | `*` |
| `PDF_TOOL_RUNTIME_DIR` | Custom writable folder for uploads/outputs | auto |

## Deploy on Vercel

### Important note

Vercel is supported, but it is best for smaller jobs. Vercel Functions have a `4.5 MB` request/response body limit and Python functions also have bundle-size limits. For larger PDFs or heavier conversion jobs, Render, Railway, or Docker/VPS hosting is a better fit.

Official docs:

- FastAPI on Vercel: https://vercel.com/docs/frameworks/backend/fastapi
- Python runtime: https://vercel.com/docs/functions/runtimes/python
- Function limits: https://vercel.com/docs/functions/limitations
- Import project: https://vercel.com/docs/getting-started-with-vercel/import

### Vercel dashboard settings

- Framework / Application Preset: `Python`
- Root Directory: `./` if this folder is the repo root
- Build Command: `None`
- Output Directory: `N/A`
- Install Command: `pip install -r requirements.txt`

### Deploy steps

1. Push the project to GitHub.
2. Import the repository in Vercel.
3. Keep the preset as `Python`.
4. Confirm the root directory is correct.
5. Deploy.

### Why this project works on Vercel

- `app.py` exposes a top-level ASGI `app`
- `.python-version` pins Python 3.12
- `vercel.json` stays minimal so Vercel can detect the root `app.py` entrypoint cleanly
- Runtime file writes automatically use a writable temp directory on Vercel

## Deploy on Render

Official guide:

- https://render.com/docs/deploy-fastapi

### Recommended settings

- Runtime: `Python`
- Build Command: `pip install -r requirements.txt`
- Start Command: `uvicorn app:app --host 0.0.0.0 --port $PORT`
- Optional Blueprint file included: `render.yaml`

### Steps

1. Create a new Web Service in Render.
2. Connect the GitHub repository.
3. Set the build and start commands above.
4. Add `PDF_TOOL_OPEN_BROWSER=0` as an environment variable.
5. Deploy.

Render is a better choice than Vercel when you expect larger uploads or longer processing time.

## Deploy on Railway

Official guide:

- https://docs.railway.com/guides/fastapi

### Recommended setup for this project

- Deploy the GitHub repo to Railway
- Start command: `uvicorn app:app --host 0.0.0.0 --port $PORT`
- Environment variable: `PDF_TOOL_OPEN_BROWSER=0`

Railway is a good fit when you want a simple managed deployment without Vercel's request body constraints.

## Deploy with Docker

Official FastAPI container docs:

- https://fastapi.tiangolo.com/deployment/docker/

### Build image

```bash
docker build -t pdfforge-tool .
```

### Run container

```bash
docker run --rm -p 8000:8000 -e PDF_TOOL_OPEN_BROWSER=0 pdfforge-tool
```

Open:

```text
http://localhost:8000
```

## Deploy on a VPS

### Simple Uvicorn run

```bash
pip install -r requirements.txt
export PDF_TOOL_OPEN_BROWSER=0
export PORT=8000
uvicorn app:app --host 0.0.0.0 --port 8000
```

### Recommended production setup

- Run behind Nginx or Caddy
- Use HTTPS
- Use a process manager like `systemd`, `supervisor`, or `pm2`
- Set `PDF_TOOL_RUNTIME_DIR` to a persistent writable folder if you do not want temporary generated files inside the project directory

## Hosting Recommendations

- Vercel: best for demos, smaller files, and lightweight usage
- Render: best default choice for production hosting
- Railway: best for quick managed deployment with fewer manual steps
- Docker/VPS: best for full control, larger files, and custom scaling

## Notes About Optional Features

- HTML to PDF may need native libraries depending on the hosting environment
- `pdf2image` may need Poppler when `pypdfium2` is not available
- Translation and advanced document conversion depend on optional Python packages
- On strict serverless platforms, temporary files are ephemeral by design

## Health Check

After deployment, test:

```text
GET /api/health
```

Expected response:

```json
{"message":"PDF Tool API is running."}
```

## Troubleshooting

### Browser tries to open on the server

Set:

```text
PDF_TOOL_OPEN_BROWSER=0
```

### Uploaded files fail on Vercel

This is usually caused by Vercel Function size limits. Move the app to Render, Railway, or Docker/VPS hosting for larger files.

### HTML or advanced conversion feature fails

Install the required native system packages on the host, or use the provided Docker image as a more predictable runtime.

## License / Handoff

If you are selling or handing off this project, the recommended production handoff package is:

- Source code
- `README.md`
- `requirements.txt`
- `Dockerfile`
- `.env.example`
- Hosting instructions for the buyer's preferred platform
