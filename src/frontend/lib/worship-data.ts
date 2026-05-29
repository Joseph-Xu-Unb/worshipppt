export type WorshipSection = {
  title: string;
  lines: string[];
};

export type HymnSection = {
  title: string;
  lines: string[];
};

export type WorshipData = {
  call_to_worship: WorshipSection;
  hymns: HymnSection[];
  theme_scripture: WorshipSection;
  response_hymn: WorshipSection;
};

export type SelectedFile = File | null;

export type GenerationResult = {
  fileName: string;
  warnings: string[];
};

export type WorshipSlidePreview = {
  id: string;
  sectionLabel: string;
  title: string;
  slideLabel: string;
  lines: string[];
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

function chunkLines(lines: string[], size = 4): string[][] {
  const normalized = lines.map((line) => line.trim()).filter(Boolean);
  const chunks: string[][] = [];

  for (let index = 0; index < normalized.length; index += size) {
    chunks.push(normalized.slice(index, index + size));
  }

  return chunks;
}

function buildSectionSlides(
  sectionLabel: string,
  title: string,
  lines: string[],
): WorshipSlidePreview[] {
  const chunks = chunkLines(lines);

  return chunks.map((chunk, index) => ({
    id: `${sectionLabel}-${title}-${index + 1}`,
    sectionLabel,
    title,
    slideLabel: `${index + 1} / ${chunks.length}`,
    lines: chunk,
  }));
}

export function buildPreview(data: WorshipData): WorshipSlidePreview[] {
  return [
    ...buildSectionSlides(
      "Call To Worship",
      `宣召经文：${data.call_to_worship.title}`,
      data.call_to_worship.lines,
    ),
    ...data.hymns.flatMap((hymn, index) =>
      buildSectionSlides(`Hymn ${index + 1}`, `Hymn${index + 1}: ${hymn.title}`, hymn.lines),
    ),
    ...buildSectionSlides(
      "Theme Scripture",
      `主题经文：${data.theme_scripture.title}`,
      data.theme_scripture.lines,
    ),
    ...buildSectionSlides(
      "Response Hymn",
      `回应诗：${data.response_hymn.title}`,
      data.response_hymn.lines,
    ),
  ];
}
