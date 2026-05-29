import type { FormEvent } from "react";
import { FilePicker } from "../components/FilePicker";
import { TopBar } from "../components/TopBar";
import { useWorshipBuilder } from "../hooks/useWorshipBuilder";
import { useNavigate } from "react-router-dom";

export function HomePage() {
  const navigate = useNavigate();
  const {
    dataFile,
    serviceDate,
    error,
    validationMessage,
    canPreview,
    setServiceDate,
    handleDataFileChange,
  } = useWorshipBuilder();

  function handleSubmit(event: FormEvent<HTMLFormElement>) {
    event.preventDefault();
    if (!canPreview || error) return;
    navigate("/preview");
  }

  return (
    <main className="app-shell home-shell">
      <TopBar />

      <section className="workspace">
        <div className="masthead">
          <p className="eyebrow">Worship Workflow</p>
          <h1>BSBC Chinese Congregation</h1>
          <p className="summary">
            Upload your worship JSON and generate the final deck through the FastAPI
            backend using the app's bundled PowerPoint template.
          </p>
        </div>

        <form className="builder" onSubmit={handleSubmit}>
          <div className="steps single">
            <FilePicker
              id="worship-data"
              label="Worship Data"
              accept=".json,application/json"
              file={dataFile}
              onChange={(file) => {
                void handleDataFileChange(file);
              }}
              helper="Structured .json"
            />
          </div>

          <div className="date-row">
            <label htmlFor="service-date">
              <span>Service Date</span>
              <input
                id="service-date"
                type="date"
                value={serviceDate}
                onChange={(event) => setServiceDate(event.target.value)}
              />
            </label>
            <button type="submit" disabled={!canPreview || Boolean(error)}>
              Preview Slides
            </button>
          </div>

          <div className="inline-note">
            The server applies your worship data to the built-in
            {" "}
            `data/template.pptx`, so users only need to provide the JSON file.
          </div>

          {error && <div className="notice error">{error}</div>}
          {validationMessage && !error && (
            <div className="notice success">{validationMessage}</div>
          )}
        </form>
      </section>
    </main>
  );
}
