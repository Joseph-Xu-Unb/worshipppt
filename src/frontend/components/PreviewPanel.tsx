import type { WorshipSlidePreview } from "../lib/worship-data";

type PreviewPanelProps = {
  preview: WorshipSlidePreview[];
};

export function PreviewPanel({ preview }: PreviewPanelProps) {
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
              <h3>{item.title}</h3>
              <div className="slide-preview-lines">
                {item.lines.map((line, index) => (
                  <p key={`${item.id}-${index}`}>{line}</p>
                ))}
              </div>
            </article>
          ))}
        </div>
      ) : (
        <p className="placeholder-copy">
          Choose a valid JSON file to preview the paged scripture and hymn slides
          before export.
        </p>
      )}
    </section>
  );
}
