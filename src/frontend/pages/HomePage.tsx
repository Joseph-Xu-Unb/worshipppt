import { FilePicker } from "../components/FilePicker";
import { PreviewPanel } from "../components/PreviewPanel";
import { useWorshipBuilder } from "../hooks/useWorshipBuilder";

export function HomePage() {
  const {
    templateFile,
    dataFile,
    serviceDate,
    isGenerating,
    error,
    result,
    preview,
    canGenerate,
    setServiceDate,
    setTemplateFile,
    handleDataFileChange,
    handleSubmit,
  } = useWorshipBuilder();

  return (
    <main className="app-shell">
      <section className="workspace">
        <div className="masthead">
          <p className="eyebrow">React + FastAPI</p>
          <h1>Worship PPT Builder</h1>
          <p className="summary">
            Upload your PowerPoint template and worship JSON, then generate the final
            deck through the FastAPI backend while preserving the original template.
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
              helper="Original template.pptx"
            />
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
            <button type="submit" disabled={!canGenerate || Boolean(error)}>
              {isGenerating ? "Generating..." : "Generate PPT"}
            </button>
          </div>

          <div className="inline-note">
            The generated deck now uses your uploaded `template.pptx` through the
            restored backend generator.
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

      <PreviewPanel preview={preview} />
    </main>
  );
}
