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

## GitHub Pages

GitHub Pages can publish the React frontend, but it cannot run the FastAPI backend.
For a working Pages deployment, host the backend separately and point the frontend at
that backend during the build.

### 1. Deploy the backend

Deploy the FastAPI app somewhere that can serve:

- `POST /api/validate`
- `POST /api/generate`
- `GET /downloads/template`
- `GET /downloads/sample-json`
- `GET /downloads/sample-ppt`

The backend now enables CORS by default and exposes the `X-Worship-Warnings` header,
so the Pages frontend can call it from another origin.

### 2. Add the repository variable

In GitHub, open `Settings -> Secrets and variables -> Actions -> Variables` and add:

- `VITE_API_BASE_URL`

Example value:

```text
https://your-backend.example.com
```

### 3. Enable GitHub Pages

In `Settings -> Pages`:

- set `Source` to `GitHub Actions`

The repository includes [.github/workflows/deploy-pages.yml](.github/workflows/deploy-pages.yml),
which builds the frontend and deploys `dist/` whenever `main` is updated.

### 4. Push to `main`

Once the backend URL is configured, pushing to `main` will publish the frontend to:

```text
https://joseph-xu-unb.github.io/worshipppt/
```

The workflow uses the repository name as the Vite base path. If you later switch to a
custom domain, update `VITE_BASE_PATH` in the workflow to `/`.

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
