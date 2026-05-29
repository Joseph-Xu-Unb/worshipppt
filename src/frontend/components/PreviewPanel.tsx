import type { WorshipSectionPreview } from "../lib/worship-data";

type PreviewPanelProps = {
  preview: WorshipSectionPreview[];
};

export function PreviewPanel({ preview }: PreviewPanelProps) {
  return (
    <aside className="reference">
      <h2>Deck Preview</h2>
      {preview.length > 0 ? (
        <ul>
          {preview.map((item) => (
            <li key={`${item.label}-${item.title}`}>
              <strong>{item.label}</strong>: {item.title} ({item.lineCount} lines)
            </li>
          ))}
        </ul>
      ) : (
        <p className="placeholder-copy">
          Choose a valid JSON file to preview the worship sections before export.
        </p>
      )}
    </aside>
  );
}
