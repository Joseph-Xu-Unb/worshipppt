# Worship PPT Generator

This project generates a worship service PowerPoint from:

- a structured worship data JSON file

The PowerPoint generation logic lives in `src/backend/worship.py`. The recommended UI is a React app in `src/frontend/` backed by a small FastAPI service in `src/backend/app.py`, and the backend uses the repository's bundled `data/template.pptx`.

## What It Creates

The generated presentation includes the worship flow used by the script:

- cover slide
- preparation and slogan slides
- worship order slide
- call-to-worship scripture
- praise prayer
- Apostles' Creed
- hymn slides
- intercessory prayer
- theme scripture
- sermon slide
- response hymn
- offering, prayer, Lord's Prayer, doxology, benediction, announcements, and closing slides

Long scripture and hymn sections are split into multiple slides, with up to 4 non-empty lines per slide.

For songs, use explicit `pages` to preserve the source presentation's lyric grouping:

```json
{
  "title": "诗歌名称",
  "pages": [
    { "lines": ["第一句", "第二句", "第三句", "第四句"], "footer": "间奏" },
    { "lines": ["下一段第一句", "下一段第二句"], "footer": "结束" }
  ]
}
```

Use this section structure for each entry in `hymns` and for `response_hymn`.
Each page becomes one slide without further splitting. `footer` is optional and
accepts `间奏`, `前奏`, `尾声`, or `结束`, with optional parentheses. It appears
after that page's lyrics and does not count toward the four lyric lines. Keep
each page to at most four lyric lines when preparing the data. Put markers in
`footer`, rather than `lines`, and omit section-level `lines` when using `pages`.

Existing sections using `lines` still work. Standalone markers (including
parenthesized forms) attach to the preceding page and end that page, even if it
has fewer than four lyric lines. Blank lines are ignored, so they cannot specify
page breaks. A marker before any lyrics is rejected because its position is
ambiguous. Frontend preview and backend generation follow the same rules.

## Requirements

- Python 3.10 or newer
- Node.js 18 or newer

Python packages:

```powershell
pip install -r requirements.txt
```

```powershell
npm install
```

## How To Run

Start the FastAPI backend:

```powershell
python app.py
```

In another terminal, start the React development server:

```powershell
npm run dev
```

Open the URL printed by Vite, usually:

```text
http://127.0.0.1:5173
```

In the web app:

1. Select the worship JSON data file.
2. Enter the service date, for example `2026-05-24`.
3. Click `Generate PPT`.
4. The generated PowerPoint downloads in the browser.

The downloaded file is named:

```text
Worship_<selected-date>.pptx
```

For example:

```text
Worship_2026-05-24.pptx
```

## Production Build

Build the React app:

```powershell
npm run build
```

Then run:

```powershell
python app.py
```

After a build, FastAPI serves the compiled frontend from `dist/` at:

```text
http://127.0.0.1:5000
```

## Docker Compose

Build and run the full app with Docker Compose:

```powershell
docker compose up --build
```

Then open:

```text
http://127.0.0.1:5000
```

This uses the included multi-stage `Dockerfile` to build the React frontend and serve it with the FastAPI backend in a single container.

## Render Deployment

Render is a good fit for this repository because it can run the frontend and FastAPI
backend together as a single Docker web service.

The repository includes [render.yaml](/home/joseph/worshipppt/render.yaml), which tells
Render to:

- create one `free` web service
- build from the included `Dockerfile`
- enable auto-deploys from `main`
- use `/api/health` as the health check

### Deploy on Render

1. Push the repository to GitHub.
2. In Render, choose `New -> Blueprint`.
3. Connect the GitHub repository.
4. Render will detect `render.yaml` and propose one web service named `worshipppt`.
5. Click `Apply`.

When the deploy finishes, open the Render URL for the live app.

### Notes

- Render injects a `PORT` environment variable, and the Docker runtime now honors it automatically.
- This app stores its template and sample files in the repository under `data/`, so no external storage setup is required.
- On Render's free plan, the service spins down after 15 minutes of inactivity and may take about a minute to wake up again.

## JSON Data Format

The JSON file must contain these top-level sections:

```json
{
  "call_to_worship": {
    "title": "Psalm 95:1-7",
    "lines": [
      "1 Come, let us sing for joy...",
      "2 Let us come before him..."
    ]
  },
  "hymns": [
    {
      "title": "Hymn title",
      "lines": [
        "Verse line 1",
        "Verse line 2",
        "",
        "Verse line 3"
      ]
    }
  ],
  "theme_scripture": {
    "title": "John 8:1-11",
    "lines": [
      "1 Scripture line...",
      "2 Scripture line..."
    ]
  },
  "response_hymn": {
    "title": "Response hymn title",
    "lines": [
      "Response line 1",
      "Response line 2"
    ]
  }
}
```

Notes:

- `title` is shown in the relevant slide title.
- `lines` is a list of strings.
- Blank strings in `lines` are ignored when the script splits text across slides.
- `hymns` can contain multiple hymn objects.

## Project Folders

Application code is organized under `src/`:

```text
data/
  2026-05-17.pptx
  template.pptx
  sample_worship_data.json
src/
  backend/
    app.py
    worship.py
  frontend/
    index.html
    main.tsx
    styles.css
```

Generated presentation files still download in the browser, and the backend applies your JSON content to the bundled `data/template.pptx`.

## MCP server

The FastAPI service also exposes a stateless Streamable HTTP MCP endpoint at
`http://127.0.0.1:5000/mcp/`. After deployment, use
`https://worshipppt.onrender.com/mcp/` in an MCP client that supports remote
Streamable HTTP servers. Use the trailing slash. No authentication is required,
just as with the existing generation API.

Available tools:

- `get_sample_worship_data`: returns example worship JSON.
- `validate_worship_data`: checks required sections, titles, and text lines.
- `generate_worship_powerpoint`: accepts `data` and optional `service_date`
  (`YYYY-MM-DD`), and returns `filename`, `mime_type`, `pptx_base64`, and `warnings`.
  The client must decode the base64 result to save the PowerPoint. Large decks
  may exceed a client's tool response limit; the web app remains available for
  direct downloads.

Start the server using the existing `python app.py` command after installing
`requirements.txt`; MCP runs in the same process and Docker image. Generated
files use temporary storage and are removed after encoding the response.
For a different hostname, configure comma-separated `MCP_ALLOWED_HOSTS` and
`MCP_ALLOWED_ORIGINS` environment variables to allow that hostname and origin.
