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
          {result.lastWeek?.source === "onedrive" && result.mattersCopied > 0 && (
            <span>
              已从 OneDrive 文件 <strong>{result.lastWeek.filename}</strong> 复制{" "}
              {result.mattersCopied} 页事项汇报。
            </span>
          )}
          {result.lastWeek?.source !== "onedrive" && result.mattersCopied > 0 && (
            <span>
              Copied {result.mattersCopied} announcement slide
              {result.mattersCopied === 1 ? "" : "s"} from last week's PPT.
            </span>
          )}
          {result.upload?.uploaded && (
            <span className="share-link">
              已写回 OneDrive：<strong>{result.upload.filename}</strong>。
              {result.upload.share_url ? (
                <>
                  <a
                    href={result.upload.share_url}
                    target="_blank"
                    rel="noreferrer"
                  >
                    分享链接
                  </a>
                  <button
                    type="button"
                    className="secondary-button"
                    onClick={() => {
                      if (result.upload?.share_url) {
                        void navigator.clipboard.writeText(result.upload.share_url);
                      }
                    }}
                  >
                    复制链接
                  </button>
                </>
              ) : result.upload.web_url ? (
                <a href={result.upload.web_url} target="_blank" rel="noreferrer">
                  打开文件
                </a>
              ) : null}
            </span>
          )}
          {result.upload?.error && (
            <span>OneDrive 写回失败：{result.upload.error}</span>
          )}
          {result.lastWeek?.auto_error && (
            <span>自动读取 OneDrive 失败：{result.lastWeek.auto_error}</span>
          )}
          {result.lastWeek && result.lastWeek.candidates.length > 0 && (
            <span>
              候选文件：
              {result.lastWeek.candidates.map((candidate) => candidate.name).join("、")}
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
