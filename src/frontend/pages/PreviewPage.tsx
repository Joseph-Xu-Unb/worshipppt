import { Navigate, useNavigate } from "react-router-dom";
import { PreviewPanel } from "../components/PreviewPanel";
import { TopBar } from "../components/TopBar";
import { useWorshipBuilder } from "../hooks/useWorshipBuilder";

export function PreviewPage() {
  const navigate = useNavigate();
  const {
    error,
    dataFile,
    result,
    preview,
    isGenerating,
    generatePpt,
    updatePreviewSlide,
    deletePreviewSlide,
    movePreviewSlide,
    insertPreviewSlide,
    clearResult,
    resetBuilder,
  } = useWorshipBuilder();

  if (!dataFile) {
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
          Edit or delete any scripture or hymn slide before generating the final
          PowerPoint.
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
          {result.mattersCopied > 0 && (
            <span>
              Copied {result.mattersCopied} announcement slide
              {result.mattersCopied === 1 ? "" : "s"} from last week's PPT.
            </span>
          )}
          {result.warnings.length > 0 && (
            <span>{result.warnings.length} warning(s) were reported.</span>
          )}
        </div>
      )}

      <PreviewPanel
        preview={preview}
        onUpdate={updatePreviewSlide}
        onDelete={deletePreviewSlide}
        onMove={movePreviewSlide}
        onInsert={insertPreviewSlide}
      />
    </main>
  );
}
