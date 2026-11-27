# Worship PPT Generator

This project generates a worship service PowerPoint from:

- a structured worship data JSON file
- optionally, last week's worship PPT from the church OneDrive archive (automatic
  retrieval, write-back, and share link)

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

### Hymn Page Numbers

Hymn slides (including each song in `hymns` and the `response_hymn`) feature centered page numbers at the bottom in `X/Y` format (e.g., `4/6` indicates slide 4 of a 6-slide hymn), displayed in 28 pt font size. Page numbers are tracked per individual song, resetting for each new hymn. Scripture slides do not include page numbers.

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

### Copying Last Week's Announcement Slides

The web app can optionally take last week's worship PPT and copy its
announcement/matters slides into the generated deck. The copy is performed at the
PowerPoint package level, so it preserves:

- pictures and background images
- tables
- grouped shapes, including nested pictures and text
- charts and their embedded Excel workbooks
- hyperlinks and media relationships

The generated deck keeps its own `matters` layout, so the template's title and
decorations are preserved. Slide notes and layout/master definitions from the
source deck are not copied. The source deck must use the same slide size for
pixel-perfect placement; a warning is returned when it does not.

New decks use a layout named `matters`. Older decks may not, so the server also
matches announcement-related layout/text and the position between the benediction
and closing slides. If detection is still ambiguous, the MCP response returns
`matters_candidates`; the agent can choose the correct 1-based slide numbers and
retry with `last_week_slide_numbers`.

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
2. (Optional) Select last week's worship PPT. A manual upload takes priority over
   OneDrive automatic retrieval.
3. Enter the service date, for example `2026-05-24`.
4. Leave “自动复制上周事项汇报” and “生成后写回 OneDrive 并生成分享链接” checked when
   OneDrive is configured. These are the defaults: the server finds the previous
   Sunday's `<YYYY-MM-DD>.pptx` in the church archive, copies its
   announcement/matters slides (pictures, tables, grouped shapes, and charts),
   writes the new deck back, and creates view and edit links. If the matters slide
   cannot be found or copied, write-back is deferred so you can choose whether to
   continue with a blank matters page or provide the correct last-week file/pages.
5. Click `Preview Slides` to review the generated slides.
6. (Optional) Edit titles, lyrics, reorder, insert, or delete individual slides.
7. Click `Generate PPT`.
8. The generated PowerPoint downloads in the browser. The result area shows the full
   editable share link and the no-login download link when write-back succeeds.
   Forward the editable link to the group chat for pastor/chair review and editing;
   on Sunday the synced local OneDrive folder on the HP/DELL computers contains the
   `<date>.pptx` file.

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

## Automatic OneDrive Retrieval and Write-back

When configured, the server uses a dedicated personal Microsoft account to:

- resolve the church archive share link (the folder is public read/write);
- find the strictly previous Sunday's `<YYYY-MM-DD>.pptx`;
- copy its matters/announcement slides into the new deck;
- write the new deck back to the same folder as `<service-date>.pptx`;
- create anonymous view and edit links and return them as `uploaded_share_url`
  (no-login download) and `uploaded_edit_url` (for pastor/chair review and
  continued editing);
- defer write-back when the matters slide cannot be copied, so the user can choose
  between continuing with a blank matters page and supplying the correct file/pages.

### One-time app registration

1. Sign in to <https://portal.azure.com> with the account that owns the directory
   (a free Azure signup is enough). Open **Microsoft Entra ID → App registrations →
   New registration**.
2. Set **Supported account types** to *Personal Microsoft account users* (or
   *Accounts in any organizational directory and personal Microsoft accounts*).
   Name the app, for example `WORSHIPPPT`.
3. Open **Authentication → Advanced settings → Allow public client flows** and choose
   **Yes**. Device-code flow needs no redirect URI and no client secret.
4. Open **API permissions → Add a permission → Microsoft Graph → Delegated
   permissions**, and add `Files.ReadWrite` and `offline_access`. Do not add
   application permissions.
5. Copy the **Application (client) ID** and **Directory (tenant) ID** from Overview.

### One-time login and Render variables

Run the login script locally and sign in with the dedicated church account:

```powershell
python onedrive_login.py --client-id <client-id> --check
```

Use `--check --check-date 2026-11-22` to verify a specific archived date. The script
prints:

```text
ONEDRIVE_CLIENT_ID=...
ONEDRIVE_TENANT_ID=consumers
ONEDRIVE_REFRESH_TOKEN=...
ONEDRIVE_TOKEN_ISSUED_AT=...
ONEDRIVE_SCOPE=Files.ReadWrite offline_access
ONEDRIVE_ARCHIVE_URL=...
```

Set these in Render. `render.yaml` declares them with `sync: false`, so the values are
not stored in Git. `--write-env onedrive.env` writes a local `0600` file; that file is
gitignored. Never paste the refresh token into chat or commit it.

Optional variables:

- `ONEDRIVE_ARCHIVE_PATH`: path inside the signed-in account's own OneDrive. Path mode
  wins over the share URL. It requires `Files.ReadWrite` if write-back is enabled.
- `ONEDRIVE_AUTO_ENABLED=false`: disable automatic retrieval.
- `ONEDRIVE_UPLOAD_ENABLED=false`: disable write-back and share-link creation.

### Token renewal

Refresh tokens last about 90 days. `/api/onedrive/status` reports `token_age_days`
and `token_status` (`ok`, `renew_soon`, `expired`). On Render's free plan there is
no persistent disk, so rotated tokens are kept in memory only and the environment
variable remains the seed unless you enable the GitHub Actions workflow below.

#### Automatic renewal with GitHub Actions

The workflow `.github/workflows/renew-onedrive-token.yml` runs monthly and can be
started manually from the Actions tab. It reads the current
`ONEDRIVE_CLIENT_ID`, `ONEDRIVE_TENANT_ID`, `ONEDRIVE_SCOPE`, and
`ONEDRIVE_REFRESH_TOKEN` from Render, refreshes the Microsoft token, and writes the
new refresh token plus `ONEDRIVE_TOKEN_ISSUED_AT` back to Render. **No Microsoft
token is stored in GitHub.**

Setup:

1. In Render, open **Account Settings → API Keys** and create a key. Copy it once.
2. In GitHub, open the repository → **Settings → Secrets and variables → Actions**
   and add:
   - `RENDER_API_KEY`: the Render API key from step 1.
   - Optional `RENDER_SERVICE_ID`: a fallback if automatic lookup finds more than one
     `worshipppt` service. The script resolves the service by name first.

   The Render API key must belong to the account/workspace that owns the
   `worshipppt` service. A key from another account sees a 404 for this service.
3. Open **Actions → Renew OneDrive refresh token → Run workflow** to run it once
   manually and confirm it succeeds. The monthly schedule then keeps it fresh.

The workflow uses `scripts/renew_onedrive_token.py`, which updates those two Render
variables via the single-variable API and then triggers a Render deploy so the
service reloads the new token immediately. It prints statuses, never token values.
If the deploy trigger fails, the token is still saved and the next Render deploy or
cold start picks it up; the log tells you to deploy manually. If the workflow fails
with `invalid_grant`, the refresh token was revoked or expired: run
`onedrive_login.py` again, update Render, and the next workflow run will use the new
seed.

GitHub Actions usage for this monthly job is negligible; public repositories do not
consume paid minutes.

### Behavior and file naming

- Archived filenames must contain `YYYY-MM-DD`. The server uploads the new deck as
  exactly `<service-date>.pptx` and replaces an existing same-date file.
- Automatic retrieval only copies when exactly one file matches the previous Sunday.
  Missing or multiple matches return `last_week_archive_candidates`; the caller can
  retry with `last_week_filename` or upload a file manually.
- Auto retrieval/write-back failures never discard the generated deck. Warnings,
  `auto_last_week_error`, and `upload_error` explain what failed.
- If the matters slide cannot be found or copied, upload is deferred
  (`upload_deferred=true`). Muse asks whether to continue with a blank matters page
  (retry with `use_onedrive=false`, `upload_to_onedrive=true`) or to provide the
  correct last-week file/page numbers.
- Older source templates may return `matters_candidates`; retry with
  `last_week_slide_numbers`.
- The MCP result includes `last_week_source`, `last_week_filename`,
  `last_week_archive_candidates`, `uploaded_to_onedrive`, `uploaded_filename`,
  `uploaded_web_url`, `uploaded_view_url`, `uploaded_share_url`,
  `uploaded_edit_url`, `upload_error`, `upload_deferred`, and
  `upload_deferred_reason`. `uploaded_share_url` is a no-login `download=1` link
  that returns the PPTX directly; `uploaded_edit_url` is the anonymous edit link
  for pastor/chair review and continued editing; `uploaded_view_url` is the
  OneDrive view link and may require a Microsoft sign-in for Office previews.

### Costs and security

- Entra app registration and Microsoft Graph calls are free. OneDrive storage quota
  still applies; weekly decks accumulate in the archive owner's OneDrive.
- The delegated token can write to files the dedicated account can access. The code
  only uploads the generated `<date>.pptx` into the configured archive and creates
  view/edit links for that file; it does not delete or rename other files.
- Do not give the account password to Muse or paste it into a chat. The account
  authorizes once through device code.
- If the archive share link is view-only, write-back fails; the archive owner must
  grant edit permission.

### Troubleshooting

- `invalid_grant` / login expired: re-run the login script and update
  `ONEDRIVE_REFRESH_TOKEN` and `ONEDRIVE_TOKEN_ISSUED_AT`.
- 403/404 on the share: verify the link is public and editable, and that the church
  account can open it. If reading works but writing returns 403, add delegated
  `Files.ReadWrite.All` to the app registration, re-run the login script with
  `--scope "Files.ReadWrite.All offline_access"`, and update `ONEDRIVE_SCOPE`; or use
  path mode with a folder inside the account.
- `AADSTS50020` / `AADSTS16000`: the app's supported account types or authority do
  not match the sign-in account. A personal-only app uses `consumers`; an app that
  includes organizational accounts uses the Directory (tenant) ID or `common`.
- `invalid_client` / public-client error: in the app registration, open
  **Authentication → Allow public client flows = Yes**. Device-code flow also does not
  need a client secret.
- `invalid_scope` / consent error: add delegated `Files.ReadWrite` and
  `offline_access` under **API permissions**, then sign in again to consent.
- Missing/multiple date matches: set `last_week_filename`; filenames must include
  `YYYY-MM-DD`.
- Upload session 401 or large-file issues: retry; a simple upload is attempted first.
  Check share permission and OneDrive quota.
- Anonymous sharing disabled: upload may still succeed but `uploaded_share_url` can be
  empty; the response includes `uploaded_web_url`.

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
    content.py
    matters.py
    muse_prompt.py
    onedrive.py
    pptx_inputs.py
    worship.py
  shared/
    church_resources.json
    muse_prompt.txt
  frontend/
    components/
    hooks/
    lib/
    pages/
    App.tsx
    index.html
    main.tsx
    styles.css
onedrive_login.py
tests/
  test_content.py
  test_onedrive.py
  test_onedrive_integration.py
requirements-dev.txt
```

Generated presentation files still download in the browser, and the backend applies your JSON content to the bundled `data/template.pptx`. When OneDrive is configured, the generated file is also written back to the archive and download/edit links are returned.

## AI Assistant Guide (Muse AI)

The web app includes an assistant guide at `/guide/muse`. It allows church volunteers to paste church group chat logs (WeChat or WhatsApp) into an AI prompt that instructs Muse AI or any LLM client to:

1. Extract dates, scriptures, and hymns from group messages.
2. Retrieve complete lyrics from the church's OneDrive lyric library.
3. Automatically validate and generate the presentation via the MCP server (`/mcp/`).
4. By default, copy last week's matters and write the new deck back to OneDrive
   without asking first. The server reads the strictly previous Sunday's
   `<YYYY-MM-DD>.pptx` from the church archive, copies the announcements
   ("matters") slides (including pictures, tables, grouped shapes, charts, and
   embedded chart data), writes the new deck as `<service-date>.pptx`, and returns
   `uploaded_share_url` (no-login download) and `uploaded_edit_url` (for
   pastor/chair review and continued editing). The agent must display both links in
   full and tell the user to forward the editable link to the group chat; on Sunday
   the synced local OneDrive folder on the HP/DELL computers contains the deck.
   Only if last week's matters cannot be found or copied does the agent ask the user
   whether to continue with a blank matters page (retry with `use_onedrive=false`,
   `upload_to_onedrive=true`) or provide the correct last-week file/pages. Older
   templates may return `matters_candidates`; retry with
   `last_week_slide_numbers`. Ambiguous dates retry with `last_week_filename`.
   Users can also ask Muse to save a new song to the church OneDrive lyric library;
   PPT generation does not upload songs automatically.
5. Keep its stored instructions in sync: call `get_muse_prompt` or
   `check_muse_prompt` before generating, replace stale instructions with the
   returned prompt, and pass the current `muse_prompt_version`. Missing or stale
   versions are rejected.

## MCP server

The FastAPI service also exposes a stateless Streamable HTTP MCP endpoint at
`http://127.0.0.1:5000/mcp/`. After deployment, use
`https://worshipppt.onrender.com/mcp/` in an MCP client that supports remote
Streamable HTTP servers. Use the trailing slash. No authentication is required,
just as with the existing generation API.

Available tools:

- `get_sample_worship_data`: returns example worship JSON.
- `validate_worship_data`: checks required sections, titles, and text lines.
- `get_muse_prompt`: returns the current canonical Muse prompt, its version, and
  the action the agent should take.
- `check_muse_prompt`: compares the agent's stored prompt text or version with the
  current canonical prompt and returns `up_to_date`, `current_version`, and
  `current_prompt`.
- `generate_worship_powerpoint`: accepts `data`, optional `service_date`
  (`YYYY-MM-DD`), optional `last_week_pptx_url` or `last_week_pptx_base64`,
  optional `last_week_slide_numbers` (1-based slide numbers), `use_onedrive`
  (default `true`), `last_week_filename` (exact archived filename for ambiguous
  dates), `upload_to_onedrive` (default `true`), and required
  `muse_prompt_version`.
  - Automatic mode resolves the archived `<previous-Sunday>.pptx`, copies its
    matters slides, writes the final deck back as `<service-date>.pptx`, and
    creates anonymous view and edit links. It returns `last_week_source`,
    `last_week_filename`, `auto_last_week_attempted`, `auto_last_week_declined`,
    `auto_last_week_error`, `last_week_archive_candidates`, `uploaded_to_onedrive`,
    `uploaded_filename`, `uploaded_web_url`, `uploaded_view_url`,
    `uploaded_share_url`, `uploaded_edit_url`, `upload_error`, `upload_deferred`,
    and `upload_deferred_reason`, in addition to the existing fields.
  - Default behavior is to copy and write back without asking first. Only when the
    matters slide cannot be found or copied does the agent ask the user whether to
    continue with a blank matters page (retry with `use_onedrive=false`,
    `upload_to_onedrive=true`) or provide the correct file/pages; write-back is
    deferred in that case. `use_onedrive=false` leaves matters blank without a
    manual reminder; `upload_to_onedrive=false` skips the upload and returns the
    deck as base64. `uploaded_share_url` adds `download=1` to the anonymous link so
    it downloads the PPTX without a sign-in; `uploaded_edit_url` is the link to
    forward to the group chat for pastor/chair review and editing;
    `uploaded_view_url` may require a Microsoft sign-in. On Sunday, the synced
    local OneDrive folder on the HP/DELL computers contains the deck.
  - Explicit `last_week_pptx_url`/`last_week_pptx_base64` (or a REST upload)
    overrides automatic retrieval. The URL must be a public `https`
    OneDrive/SharePoint link; uploads and downloads are limited to 50 MB.
  - Older templates may return `matters_candidates`; retry with
    `last_week_slide_numbers=[...]`. Missing or multiple archived dates return
    candidates; retry with `last_week_filename` or upload a file.
  - Auto retrieval and write-back are best-effort: failures still return the deck
    with warnings and an explanatory `upload_error`.
  - When OneDrive is not configured (or automatic mode is disabled), `reminder` and
    `post_generation_prompt` contain the manual copy and browser-download-history
    fallback.
  - Large decks may exceed a client's tool response limit; the web app remains
    available for direct downloads.

### Muse prompt version sync

The canonical Muse instructions live in `src/shared/muse_prompt.txt`, with the
church URLs in `src/shared/church_resources.json`. Both the web guide and the MCP
server render this same file, so there is one source of truth. The server computes
a short hash version from the prompt template and resource URLs.

- Before generating, an agent should call `get_muse_prompt` or `check_muse_prompt`.
  If its stored prompt is older or different, it must replace its stored
  instructions with the returned prompt.
- `generate_worship_powerpoint` requires `muse_prompt_version`. A missing or stale
  version is rejected with an error that includes the current version and the
  update instruction.
- The MCP server cannot read Muse's internal stored prompt, so this is an explicit
  version handshake: the agent must call the check/get tool and self-report the
  version. The server enforces the version, not the agent's private prompt text.
- The `/guide/muse` page loads the same shared prompt, so volunteers always copy
  the current version.

Start the server using the existing `python app.py` command after installing
`requirements.txt`; MCP runs in the same process and Docker image. Generated
files use temporary storage and are removed after encoding the response.
For a different hostname, configure comma-separated `MCP_ALLOWED_HOSTS` and
`MCP_ALLOWED_ORIGINS` environment variables to allow that hostname and origin.

## Tests

Install the test dependency and run the backend suite:

```powershell
pip install -r requirements-dev.txt
python -m unittest discover -s tests -v
```

The backend tests mock OneDrive, so they run offline. They also clear
`ONEDRIVE_*` variables in the test module so a configured shell does not trigger
real network calls.

Build/type-check the frontend:

```powershell
npm run build
```
