import type { ChangeEvent } from "react";
import { formatSize, type SelectedFile } from "../lib/worship-data";

type FilePickerProps = {
  id: string;
  label: string;
  accept: string;
  file: SelectedFile;
  onChange: (file: SelectedFile) => void;
  helper: string;
};

export function FilePicker({
  id,
  label,
  accept,
  file,
  onChange,
  helper,
}: FilePickerProps) {
  function handleChange(event: ChangeEvent<HTMLInputElement>) {
    onChange(event.target.files?.[0] ?? null);
  }

  return (
    <label className="file-panel" htmlFor={id}>
      <span className="file-label">{label}</span>
      <div className="file-picker-row">
        <span className="file-action">{file ? "Replace JSON File" : "Choose JSON File"}</span>
        <span className={file ? "file-name selected" : "file-name"}>
          {file ? file.name : "No file selected"}
        </span>
      </div>
      <span className="file-meta">{file ? formatSize(file) : helper}</span>
      <input id={id} type="file" accept={accept} onChange={handleChange} />
    </label>
  );
}
