import type { FormEvent } from "react";
import { FilePicker } from "../components/FilePicker";
import { TopBar } from "../components/TopBar";
import { useWorshipBuilder } from "../hooks/useWorshipBuilder";
import { useNavigate } from "react-router-dom";

export function HomePage() {
  const navigate = useNavigate();
  const {
    dataFile,
    lastWeekFile,
    serviceDate,
    error,
    validationMessage,
    canPreview,
    oneDriveStatus,
    useOneDrive,
    uploadToOneDrive,
    setServiceDate,
    setLastWeekFile,
    setUseOneDrive,
    setUploadToOneDrive,
    handleDataFileChange,
  } = useWorshipBuilder();

  const oneDriveMessage = !oneDriveStatus
    ? "OneDrive 状态未知；生成时由服务器决定是否自动读取。"
    : !oneDriveStatus.configured
      ? "OneDrive 未配置；可以手动上传上周 PPT。"
      : oneDriveStatus.token_status === "expired"
        ? "OneDrive 登录已过期，请重新运行登录脚本并更新 Render 环境变量。"
        : oneDriveStatus.token_status === "renew_soon"
          ? `OneDrive 登录即将到期（已使用约 ${oneDriveStatus.token_age_days ?? 0} 天），请计划续期。`
          : oneDriveStatus.error
            ? oneDriveStatus.error
            : oneDriveStatus.mode === "path"
              ? `OneDrive 已连接（路径模式：${oneDriveStatus.archive}）。`
              : "OneDrive 已连接：会自动读取上周文件并写回归档。";
  const oneDriveTone =
    oneDriveStatus?.token_status === "expired"
      ? "error"
      : oneDriveStatus?.token_status === "renew_soon" ||
          (oneDriveStatus?.configured && oneDriveStatus.error)
        ? "warning"
        : "";

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

        <div className={`onedrive-status ${oneDriveTone}`}>
          <strong>OneDrive</strong>
          <span>{oneDriveMessage}</span>
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
            <FilePicker
              id="last-week-pptx"
              label="Last Week's PPT (optional)"
              accept=".pptx,application/vnd.openxmlformats-officedocument.presentationml.presentation"
              file={lastWeekFile}
              onChange={setLastWeekFile}
              helper="可选；手动上传后优先于 OneDrive 自动读取"
            />
          </div>

          <div className="auto-last-week">
            <label className="toggle-row">
              <input
                type="checkbox"
                checked={useOneDrive}
                onChange={(event) => setUseOneDrive(event.target.checked)}
                disabled={oneDriveStatus ? !oneDriveStatus.configured : false}
              />
              <span>
                <strong>自动复制上周事项汇报</strong>
                <small>从教会 OneDrive 归档读取上周（前一个主日）的文件。</small>
              </span>
            </label>
            <label className="toggle-row">
              <input
                type="checkbox"
                checked={uploadToOneDrive}
                onChange={(event) => setUploadToOneDrive(event.target.checked)}
                disabled={
                  oneDriveStatus
                    ? !oneDriveStatus.configured || !oneDriveStatus.upload_enabled
                    : false
                }
              />
              <span>
                <strong>生成后写回 OneDrive 并生成分享链接</strong>
                <small>新文件按 &lt;YYYY-MM-DD&gt;.pptx 保存，分享链接会显示在结果页。</small>
              </span>
            </label>
            {lastWeekFile && (
              <small className="toggle-note">
                已选择手动文件，本次生成会优先使用它。
              </small>
            )}
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
            `data/template.pptx`. When OneDrive is connected, it reads last week's
            file from the church archive, copies its announcement/matters slides
            (including pictures, tables, grouped shapes, and charts), writes the new
            deck back, and returns a share link. Manual upload still takes priority.
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
