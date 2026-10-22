export type LyricPage = {
  lines: string[];
  footer?: string;
};

export type WorshipSection = {
  title: string;
  lines: string[];
  pages?: LyricPage[];
};

export type HymnSection = WorshipSection;

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
  kind: "call_to_worship" | "hymn" | "theme_scripture" | "response_hymn";
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

function isMarker(line: string): boolean {
  return /^(?:间奏|前奏|尾声|结束|[（(]\s*(?:间奏|前奏|尾声|结束)\s*[）)])$/.test(line.trim());
}

function pageLines(page: LyricPage): string[] {
  const lines = page.lines.map((line) => line.trim());
  return page.footer ? [...lines, page.footer.trim()] : lines;
}

function parseSection(value: unknown, key: string): WorshipSection {
  if (!isRecord(value)) {
    throw new Error(`${key} must be an object.`);
  }

  const title = value.title;
  const lines = value.lines;

  if (typeof title !== "string" || !title.trim()) {
    throw new Error(`${key}.title must be a non-empty string.`);
  }

  if ("pages" in value) {
    if (!Array.isArray(value.pages) || value.pages.length === 0) {
      throw new Error(`${key}.pages must be a non-empty list.`);
    }
    const pages = value.pages.map((page: unknown, index: number): LyricPage => {
      const path = `${key}.pages[${index}]`;
      if (!isRecord(page) || !isStringArray(page.lines) || page.lines.length === 0
        || page.lines.some((line) => !line.trim() || isMarker(line))) {
        throw new Error(`${path}.lines must contain non-empty lyric strings; put markers in footer.`);
      }
      if ("footer" in page && (typeof page.footer !== "string" || !isMarker(page.footer))) {
        throw new Error(`${path}.footer must be 间奏, 前奏, 尾声 or 结束 (optional parentheses).`);
      }
      return { lines: page.lines, ...(typeof page.footer === "string" ? { footer: page.footer } : {}) };
    });
    const flattened = pages.flatMap(pageLines);
    const originalLines = pages.flatMap((page) => [
      ...page.lines, ...(page.footer ? [page.footer] : []),
    ]);
    if ("lines" in value && (!isStringArray(lines)
      || JSON.stringify(lines) !== JSON.stringify(originalLines))) {
      throw new Error(`${key}.lines conflicts with pages; omit lines when using pages.`);
    }
    return { title: title.trim(), lines: flattened, pages };
  }

  if (!isStringArray(lines) || lines.length === 0 || !lines.some((line) => line.trim())) {
    throw new Error(`${key}.lines must be a non-empty list of strings.`);
  }
  chunkLines(lines);

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
  let closed = false;

  for (const line of normalized) {
    const previous = chunks[chunks.length - 1];
    if (isMarker(line)) {
      if (!previous) {
        throw new Error("A lyric marker must follow lyrics; its position is unknown.");
      }
      previous.push(line);
      closed = true;
    } else {
      if (!previous || closed || previous.length >= size) {
        chunks.push([line]);
      } else {
        previous.push(line);
      }
      closed = false;
    }
  }

  return chunks;
}

function buildSectionSlides(
  kind: WorshipSlidePreview["kind"],
  sectionLabel: string,
  title: string,
  section: WorshipSection,
): WorshipSlidePreview[] {
  const chunks = section.pages ? section.pages.map(pageLines) : chunkLines(section.lines);

  return chunks.map((chunk, index) => ({
    id: `${sectionLabel}-${title}-${index + 1}`,
    kind,
    sectionLabel,
    title,
    slideLabel: `${index + 1} / ${chunks.length}`,
    lines: chunk,
  }));
}

export function buildPreview(data: WorshipData): WorshipSlidePreview[] {
  return [
    ...buildSectionSlides(
      "call_to_worship",
      "Call To Worship",
      `宣召经文：${data.call_to_worship.title}`,
      data.call_to_worship,
    ),
    ...data.hymns.flatMap((hymn, index) =>
      buildSectionSlides(
        "hymn",
        `Hymn ${index + 1}`,
        `Hymn${index + 1}: ${hymn.title}`,
        hymn,
      ),
    ),
    ...buildSectionSlides(
      "theme_scripture",
      "Theme Scripture",
      `主题经文：${data.theme_scripture.title}`,
      data.theme_scripture,
    ),
    ...buildSectionSlides(
      "response_hymn",
      "Response Hymn",
      `回应诗：${data.response_hymn.title}`,
      data.response_hymn,
    ),
  ];
}
