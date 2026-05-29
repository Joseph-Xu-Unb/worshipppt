import React, { useMemo, useState } from "react";
import { createRoot } from "react-dom/client";
import "./styles.css";

const today = new Date().toISOString().slice(0, 10);

function formatSize(file) {
  if (!file) return "";
  if (file.size < 1024 * 1024) return `${Math.round(file.size / 1024)} KB`;
  return `${(file.size / (1024 * 1024)).toFixed(1)} MB`;
}

function FilePicker({ id, label, accept, file, onChange, helper }) {
  return (
    <label className="file-panel" htmlFor={id}>
      <span className="file-label">{label}</span>
      <span className={file ? "file-name selected" : "file-name"}>
        {file ? file.name : "Choose file"}
      </span>
      <span className="file-meta">{file ? formatSize(file) : helper}</span>
      <input
        id={id}
        type="file"
        accept={accept}
        onChange={(event) => onChange(event.target.files?.[0] ?? null)}
      />
    </label>
  );
}

function App() {
  const [templateFile, setTemplateFile] = useState(null);
  const [dataFile, setDataFile] = useState(null);
  const [serviceDate, setServiceDate] = useState(today);
  const [isGenerating, setIsGenerating] = useState(false);
  const [error, setError] = useState("");
  const [result, setResult] = useState(null);

  const canGenerate = useMemo(
    () => templateFile && dataFile && serviceDate && !isGenerating,
    [templateFile, dataFile, serviceDate, isGenerating],
  );

  async function handleSubmit(event) {
    event.preventDefault();
    if (!canGenerate) return;

    setIsGenerating(true);
    setError("");
    setResult(null);

    const formData = new FormData();
    formData.append("template", templateFile);
    formData.append("data", dataFile);
    formData.append("date", serviceDate);

    try {
      const response = await fetch("/api/generate", {
        method: "POST",
        body: formData,
      });

      if (!response.ok) {
        const payload = await response.json().catch(() => ({}));
        throw new Error(payload.error || "Failed to generate PowerPoint.");
      }

      const blob = await response.blob();
      const fileName = `Worship_${serviceDate}.pptx`;
      const url = URL.createObjectURL(blob);
      const link = document.createElement("a");
      link.href = url;
      link.download = fileName;
      link.click();
      URL.revokeObjectURL(url);

      const rawWarnings = response.headers.get("X-Worship-Warnings");
      const warnings = rawWarnings ? JSON.parse(rawWarnings) : [];
      setResult({ fileName, warnings });
    } catch (caughtError) {
      setError(caughtError.message);
    } finally {
      setIsGenerating(false);
    }
  }

  return (
    <main className="app-shell">
      <section className="workspace">
        <div className="masthead">
          <p className="eyebrow">Fredericton BSBC</p>
          <h1>Worship PPT Builder</h1>
          <p className="summary">
            Upload a PowerPoint template and structured worship JSON, then download
            the generated service deck.
          </p>
        </div>

        <form className="builder" onSubmit={handleSubmit}>
          <div className="steps">
            <FilePicker
              id="template"
              label="Template"
              accept=".pptx"
              file={templateFile}
              onChange={setTemplateFile}
              helper="PowerPoint .pptx"
            />
            <FilePicker
              id="worship-data"
              label="Worship Data"
              accept=".json,application/json"
              file={dataFile}
              onChange={setDataFile}
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
            <button type="submit" disabled={!canGenerate}>
              {isGenerating ? "Generating..." : "Generate PPT"}
            </button>
          </div>

          {error && <div className="notice error">{error}</div>}
          {result && (
            <div className="notice success">
              <strong>{result.fileName}</strong> downloaded.
              {result.warnings.length > 0 && (
                <span>{result.warnings.length} template warning(s) were reported.</span>
              )}
            </div>
          )}
        </form>
      </section>

      <aside className="reference">
        <h2>JSON Checklist</h2>
        <ul>
          <li>call_to_worship.title and lines</li>
          <li>hymns array with title and lines</li>
          <li>theme_scripture.title and lines</li>
          <li>response_hymn.title and lines</li>
        </ul>
      </aside>
    </main>
  );
}

createRoot(document.getElementById("root")).render(<App />);
