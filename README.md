# Worship PPT Generator

This project generates a worship service PowerPoint from:

- a PowerPoint template file, usually `template.pptx`
- a structured worship data JSON file

The PowerPoint generation logic lives in `src/backend/worship.py`. The recommended UI is a React app in `src/frontend/` backed by a small FastAPI service in `src/backend/app.py`.

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
- PowerPoint template file with the expected slide layout names

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

1. Select the PPT template file, for example `template.pptx`.
2. Select the worship JSON data file.
3. Enter the service date, for example `2026-05-24`.
4. Click `Generate PPT`.
5. The generated PowerPoint downloads in the browser.

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
src/
  backend/
    app.py
    worship.py
  frontend/
    index.html
    main.tsx
    styles.css
```

Generated presentation files still download in the browser, but the backend is responsible for applying your JSON content to the uploaded PowerPoint template.
