# Document to PPT Generator

Generate PPT from `Word / Markdown / PDF` with a unified pipeline:

1. Parse source document while preserving hierarchy
2. Plan slide narrative with Gemini 3.1 Pro (with strict JSON schema)
3. Generate slide images
4. Render PPT using template-style full-screen visuals
5. Evaluate PPT quality and export score report

## Install

```bash
npm install
```

## Environment

Copy `.env.example` to `.env` and configure:

```env
IMAGE_API_KEY=your_api_key_here
IMAGE_API_BASE_URL=https://www.aigenimage.cn
PORT=3000

ENABLE_AI_IMAGES=true
IMAGE_CONCURRENCY=2
IMAGE_MODEL=gemini-3.1-flash-image-preview
IMAGE_RESOLUTION=2K

ENABLE_PLANNER=true
PLANNER_MODEL=gemini-3.1-pro-preview
PLANNER_API_BASE_URL=https://www.aigenimage.cn:3001
PLANNER_AUTH_TOKEN=
LLM_AUTH_TOKEN=
CLOUDFLARE_WORKER_URL=
LLM_API_KEY=
GOOGLE_API_KEY=
AIWORKFLOW_BACKEND_ENV_PATH=E:\\GitHubWorkSpace\\aiworkflow\\back-end\\.env
PLANNER_CONTENT_MODE=strict
PLANNER_EXPAND_SPARSE_CONTENT=true
PLANNER_USE_GUEST_LOGIN=false
ENABLE_EVALUATION=true

PPT_TEMPLATE_STYLE=true
PPT_KEEP_TEXT=true
PPT_IMAGE_ONLY_MODE=false
PPT_MAX_BULLETS_PER_SLIDE=5
```

### Planner auth notes

- `PLANNER_AUTH_TOKEN` (or `LLM_AUTH_TOKEN`) is used for `/api/llm`.
- If planner-specific token is empty, the planner will also try `IMAGE_API_KEY`.
- If `CLOUDFLARE_WORKER_URL` + (`LLM_API_KEY` or `GOOGLE_API_KEY`) is available, planner will call Gemini via worker proxy first (no `PLANNER_AUTH_TOKEN` required).
- If local `.env` does not contain worker settings, planner can auto-read `AIWORKFLOW_BACKEND_ENV_PATH` (default: `E:\\GitHubWorkSpace\\aiworkflow\\back-end\\.env`).
- If not provided, planner falls back to local heuristic planning.
- `PLANNER_USE_GUEST_LOGIN=true` can auto-login guest, but guest accounts may have zero points.

### Planner content modes

- `PLANNER_CONTENT_MODE=strict` (default): stay as close as possible to source content and wording.
- `PLANNER_CONTENT_MODE=creative`: keep slide order and hierarchy, but allow Gemini 3.1 Pro to lightly polish bullets, summaries, and image direction without adding unsupported facts.
- `PLANNER_EXPAND_SPARSE_CONTENT=true` (default): when a slide is too sparse, planner first fills baseline content and then asks LLM to expand it (worker proxy first, token-based `/api/llm` fallback).

### Quality evaluation system

Each generation can output:

- `<ppt-name>.quality.json`
- `<ppt-name>.quality.md`

Scoring dimensions:

- Content logic (hierarchy/transition/title quality/redundancy checks)
- Layout quality (visual coverage/overflow/layout diversity)
- Image semantic alignment (prompt coverage/alignment/fallback penalty)
- Content richness (sparse slide and low-information penalty)
- Audience fit (takeaway density/readability/final-slide action cue)
- Consistency (title style/transition consistency/cross-slide completeness gap)

## Run Web server

```bash
npm start
```

Open <http://localhost:3000>.

## CLI generation

```bash
npm run generate -- --input input/计算机发展史.docx --output output/计算机发展史-unified.pptx
```

Creative mode:

```bash
npm run generate -- --input input/计算机发展史.docx --output output/计算机发展史-creative.pptx --planner-mode creative
```

## API

`POST /generate-ppt`

- Content-Type: `multipart/form-data`
- Field: `file` (`.md/.docx/.pdf`)
- Optional field: `plannerMode` (`strict` or `creative`)
- Returns: generated `.pptx`

Example:

```bash
curl -X POST http://localhost:3000/generate-ppt \
  -F "file=@input/计算机发展史.docx" \
  -F "plannerMode=creative" \
  --output output/from-api.pptx
```

## Render modes

- Template overlay mode (default): full-screen image + concise text overlay
- Image-only mode: set `PPT_IMAGE_ONLY_MODE=true`

## Node compatibility

- Recommended: Node.js `>=16`
- Includes `Object.hasOwn` polyfill for older runtime compatibility
