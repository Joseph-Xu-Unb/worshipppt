# Worship PPT Generator

This project generates a worship service PowerPoint from:

- a PowerPoint template file, usually `template.pptx`
- a structured worship data JSON file

The PowerPoint generation logic lives in `worship.py`. The recommended UI is now a React 18 web app backed by a small Flask API in `app.py`.

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
- Python packages:

```powershell
pip install -r requirements.txt
```

Frontend packages:

```powershell
npm install
```

## How To Run

Start the Flask backend:

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

Generation logs are written to:

```text
worship_ppt.log
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

After a build, Flask serves the compiled frontend from `dist/` at:

```text
http://127.0.0.1:5000
```

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

## Template Layouts

The script depends on named slide layouts in the PPT template. The layout names are matched case-insensitively.

Expected layout names include:

```text
cover
prepare
Slogan
order
call_to_scripture
praise_prayer
apostles_creed
Hymn
intercessory_prayer
theme_scripture
sermon
response
Offering
praying
lords_prayer
ode_to_the_Trinity
benediction
child_pickup_reminder
announcements_welcome_banner
matters
wishyouwell
```

The script also references one Chinese-named closing layout. If a layout is missing, generation continues and the missing layout is written as a warning in the log.

## Template Placeholder Expectations

The script fills placeholders by placeholder index:

- title placeholders: `0`, `11`
- body placeholders: `1`, `10`, `13`
- page number placeholder: `12`

If a slide layout does not have one of these placeholder indexes, that part of the slide may remain blank.

## Project Folders

The repository ignores generated/local folders:

```text
data/
output/
```

Typical usage:

- keep worship JSON input files in `data/`
- keep generated PPT files in `output/` or beside the selected JSON file
- keep the reusable PowerPoint template in the project root

## Developer Notes

The reusable generation entry point is:

```python
generate_worship_ppt(
    template_path,
    data,
    output_path,
    selected_date=None,
)
```

`ChurchApp` is kept as a compatibility alias for the Tkinter controller class.

The web API endpoint is:

```text
POST /api/generate
```

It expects multipart form fields:

- `template`: PowerPoint `.pptx` file
- `data`: worship `.json` file
- `date`: service date string
