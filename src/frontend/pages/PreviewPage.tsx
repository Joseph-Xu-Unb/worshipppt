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
    uploadToOneDrive,
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
              {result.upload.edit_url ? (
                <>
                  <span>
                    可编辑分享链接：
                    <a
                      href={result.upload.edit_url}
                      target="_blank"
                      rel="noreferrer"
                    >
                      {result.upload.edit_url}
                    </a>
                  </span>
                  <button
                    type="button"
                    className="secondary-button"
                    onClick={() => {
                      if (result.upload?.edit_url) {
                        void navigator.clipboard.writeText(result.upload.edit_url);
                      }
                    }}
                  >
                    复制可编辑链接
                  </button>
                  <span>
                    请把该链接转发到教会群聊，请牧师和主席 review，并在共享链接上继续修改。
                  </span>
                </>
              ) : (
                <span>
                  未能创建可编辑分享链接；请登录教会 OneDrive 后在归档文件夹中编辑。
                </span>
              )}
              {result.upload.share_url && (
                <>
                  <span>
                    下载链接：
                    <a
                      href={result.upload.share_url}
                      target="_blank"
                      rel="noreferrer"
                    >
                      {result.upload.share_url}
                    </a>
                  </span>
                  <button
                    type="button"
                    className="secondary-button"
                    onClick={() => {
                      if (result.upload?.share_url) {
                        void navigator.clipboard.writeText(result.upload.share_url);
                      }
                    }}
                  >
                    复制下载链接
                  </button>
                </>
              )}
              {!result.upload.edit_url &&
                !result.upload.share_url &&
                result.upload.web_url && (
                  <a href={result.upload.web_url} target="_blank" rel="noreferrer">
                    打开文件
                  </a>
                )}
              {result.upload.view_url &&
                result.upload.view_url !== result.upload.edit_url &&
                result.upload.view_url !== result.upload.share_url && (
                  <a
                    href={result.upload.view_url}
                    target="_blank"
                    rel="noreferrer"
                  >
                    在 OneDrive 查看（可能需登录）
                  </a>
                )}
              <span>
                周日可在同步了该 OneDrive 文件夹的惠普/DELL 电脑上，从本地 OneDrive
                打开 {result.upload.filename}。
              </span>
            </span>
          )}
          {result.upload?.deferred && (
            <span>
              上周事项汇报页未成功复制，OneDrive 写回已暂缓。请取消勾选“自动复制上周事项汇报”后重新生成并写回（事项页留空），或提供上周文件/页码。
            </span>
          )}
          {result.upload?.error && (
            <span>OneDrive 写回失败：{result.upload.error}</span>
          )}
          {result.lastWeek?.auto_declined && result.mattersCopied === 0 && (
            <span>已按选择不复制上周事项汇报。</span>
          )}
          {!uploadToOneDrive && (
            <span>已按选择不写回 OneDrive。</span>
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
