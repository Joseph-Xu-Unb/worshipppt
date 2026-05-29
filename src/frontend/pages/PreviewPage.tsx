import { Navigate, useNavigate } from "react-router-dom";
import { PreviewPanel } from "../components/PreviewPanel";
import { TopBar } from "../components/TopBar";
import { useWorshipBuilder } from "../hooks/useWorshipBuilder";

export function PreviewPage() {
  const navigate = useNavigate();
  const {
    error,
    result,
    preview,
    isGenerating,
    generatePpt,
    clearResult,
    resetBuilder,
  } = useWorshipBuilder();

  if (preview.length === 0) {
    return <Navigate to="/" replace />;
  }

  async function handleGenerateClick() {
    clearResult();
    await generatePpt();
  }

  function handleCancelClick() {
    resetBuilder();
    navigate("/", { replace: true });
  }

  return (
    <main className="app-shell preview-shell">
      <TopBar />

      <section className="preview-hero">
        <p className="eyebrow">Review Before Export</p>
        <h1>Slide Preview</h1>
        <p className="summary">
          Review every scripture and hymn slide before generating the final PowerPoint.
          If anything looks off, go back and adjust the JSON.
        </p>
      </section>

      <div className="preview-actions">
        <button type="button" className="secondary-button" onClick={() => navigate("/")}>
          Back To Edit
        </button>
        <button type="button" className="secondary-button" onClick={handleCancelClick}>
          Cancel
        </button>
        <button type="button" onClick={() => void handleGenerateClick()} disabled={isGenerating}>
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

      <PreviewPanel preview={preview} />
    </main>
  );
}
