import type { WorshipData } from "../presentation";

export type SelectedFile = File | null;

export type GenerationResult = {
  fileName: string;
  warnings: string[];
};

export type WorshipSectionPreview = {
  label: string;
  title: string;
  lineCount: number;
};

export function formatSize(file: SelectedFile): string {
  if (!file) return "";
  if (file.size < 1024 * 1024) return `${Math.round(file.size / 1024)} KB`;
  return `${(file.size / (1024 * 1024)).toFixed(1)} MB`;
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null;
}

function isStringArray(value: unknown): value is string[] {
  return Array.isArray(value) && value.every((item) => typeof item === "string");
}

function parseSection(value: unknown, key: string) {
  if (!isRecord(value)) {
    throw new Error(`${key} must be an object.`);
  }

  const title = value.title;
  const lines = value.lines;

  if (typeof title !== "string" || !title.trim()) {
    throw new Error(`${key}.title must be a non-empty string.`);
  }

  if (!isStringArray(lines) || lines.length === 0) {
    throw new Error(`${key}.lines must be a non-empty list of strings.`);
  }

  return {
    title: title.trim(),
    lines,
  };
}

export function parseWorshipData(raw: string): WorshipData {
  let parsed: unknown;

  try {
    parsed = JSON.parse(raw);
  } catch {
    throw new Error("Invalid JSON file.");
  }

  if (!isRecord(parsed)) {
    throw new Error("JSON root must be an object.");
  }

  const hymnsValue = parsed.hymns;
  if (!Array.isArray(hymnsValue) || hymnsValue.length === 0) {
    throw new Error("hymns must be a non-empty array.");
  }

  return {
    call_to_worship: parseSection(parsed.call_to_worship, "call_to_worship"),
    hymns: hymnsValue.map((hymn, index) => parseSection(hymn, `hymns[${index}]`)),
    theme_scripture: parseSection(parsed.theme_scripture, "theme_scripture"),
    response_hymn: parseSection(parsed.response_hymn, "response_hymn"),
  };
}

export function buildPreview(data: WorshipData): WorshipSectionPreview[] {
  return [
    {
      label: "Call To Worship",
      title: data.call_to_worship.title,
      lineCount: data.call_to_worship.lines.filter((line) => line.trim()).length,
    },
    ...data.hymns.map((hymn, index) => ({
      label: `Hymn ${index + 1}`,
      title: hymn.title,
      lineCount: hymn.lines.filter((line) => line.trim()).length,
    })),
    {
      label: "Theme Scripture",
      title: data.theme_scripture.title,
      lineCount: data.theme_scripture.lines.filter((line) => line.trim()).length,
    },
    {
      label: "Response Hymn",
      title: data.response_hymn.title,
      lineCount: data.response_hymn.lines.filter((line) => line.trim()).length,
    },
  ];
}
