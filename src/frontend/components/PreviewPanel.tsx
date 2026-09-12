import { useState } from "react";
import type { WorshipSlidePreview } from "../lib/worship-data";

type PreviewPanelProps = {
  preview: WorshipSlidePreview[];
  onUpdate: (id: string, title: string, lines: string[]) => void;
  onDelete: (id: string) => void;
};

export function PreviewPanel({ preview, onUpdate, onDelete }: PreviewPanelProps) {
  const [editingId, setEditingId] = useState<string | null>(null);
  const [draftTitle, setDraftTitle] = useState("");
  const [draftLines, setDraftLines] = useState("");

  function startEditing(slide: WorshipSlidePreview) {
    setEditingId(slide.id);
    setDraftTitle(slide.title);
    setDraftLines(slide.lines.join("\n"));
  }

  function cancelEditing() {
    setEditingId(null);
    setDraftTitle("");
    setDraftLines("");
  }

  function saveSlide(id: string) {
    const title = draftTitle.trim();
    if (!title) return;

    const lines = draftLines
      .split(/\r?\n/)
      .map((line) => line.trim())
      .filter(Boolean);
    onUpdate(id, title, lines);
    cancelEditing();
  }

  function deleteSlide(slide: WorshipSlidePreview) {
    if (window.confirm(`Delete ${slide.sectionLabel} slide ${slide.slideLabel}?`)) {
      onDelete(slide.id);
      if (editingId === slide.id) cancelEditing();
    }
  }

  return (
    <section className="reference">
      <h2>Slide Preview</h2>
      {preview.length > 0 ? (
        <div className="slide-preview-list">
          {preview.map((item) => (
            <article key={item.id} className="slide-preview-card">
              <div className="slide-preview-head">
                <span className="slide-preview-section">{item.sectionLabel}</span>
                <span className="slide-preview-page">{item.slideLabel}</span>
              </div>
              {editingId === item.id ? (
                <div className="slide-editor">
                  <label>
                    <span>Slide title</span>
                    <input
                      value={draftTitle}
                      onChange={(event) => setDraftTitle(event.target.value)}
                    />
                  </label>
                  <label>
                    <span>Slide text (one line per row)</span>
                    <textarea
                      rows={6}
                      value={draftLines}
                      onChange={(event) => setDraftLines(event.target.value)}
                    />
                  </label>
                  <div className="slide-card-actions">
                    <button type="button" className="secondary-button" onClick={cancelEditing}>
                      Cancel
                    </button>
                    <button
                      type="button"
                      onClick={() => saveSlide(item.id)}
                      disabled={!draftTitle.trim()}
                    >
                      Save
                    </button>
                  </div>
                </div>
              ) : (
                <>
                  <h3>{item.title}</h3>
                  <div className="slide-preview-lines">
                    {item.lines.map((line, index) => (
                      <p key={`${item.id}-${index}`}>{line}</p>
                    ))}
                  </div>
                  <div className="slide-card-actions">
                    <button type="button" className="secondary-button" onClick={() => startEditing(item)}>
                      Edit
                    </button>
                    <button type="button" className="delete-button" onClick={() => deleteSlide(item)}>
                      Delete
                    </button>
                  </div>
                </>
              )}
            </article>
          ))}
        </div>
      ) : (
        <p className="placeholder-copy">
          All preview slides have been deleted. Generate the PowerPoint to export
          only the remaining standard template slides, or go back to reload the JSON.
        </p>
      )}
    </section>
  );
}
