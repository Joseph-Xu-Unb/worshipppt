import {
  createContext,
  useContext,
  useEffect,
  useMemo,
  useState,
  type ReactNode,
} from "react";
import {
  buildPreview,
  parseWorshipData,
  type GenerationResult,
  type LastWeekMetadata,
  type OneDriveStatus,
  type SelectedFile,
  type UploadMetadata,
  type WorshipSlidePreview,
} from "../lib/worship-data";
import { apiUrl } from "../lib/app-config";

const today = new Date().toISOString().slice(0, 10);

type WorshipBuilderContextValue = {
  dataFile: SelectedFile;
  lastWeekFile: SelectedFile;
  serviceDate: string;
  isGenerating: boolean;
  error: string;
  validationMessage: string;
  result: GenerationResult | null;
  preview: WorshipSlidePreview[];
  canPreview: boolean;
  oneDriveStatus: OneDriveStatus | null;
  useOneDrive: boolean;
  uploadToOneDrive: boolean;
  setServiceDate: (value: string) => void;
  setLastWeekFile: (file: SelectedFile) => void;
  setUseOneDrive: (value: boolean) => void;
  setUploadToOneDrive: (value: boolean) => void;
  handleDataFileChange: (file: SelectedFile) => Promise<void>;
  updatePreviewSlide: (id: string, title: string, lines: string[]) => void;
  deletePreviewSlide: (id: string) => void;
  movePreviewSlide: (id: string, direction: -1 | 1) => void;
  insertPreviewSlide: (afterId: string) => string;
  generatePpt: () => Promise<boolean>;
  clearResult: () => void;
  resetBuilder: () => void;
};

const WorshipBuilderContext = createContext<WorshipBuilderContextValue | null>(null);

async function readResponseError(response: Response, fallback: string) {
  const contentType = response.headers.get("content-type") || "";

  if (contentType.includes("application/json")) {
    const payload = (await response.json().catch(() => ({}))) as {
      error?: string;
      detail?: string;
      message?: string;
    };
    return payload.error || payload.detail || payload.message || fallback;
  }

  const text = await response.text().catch(() => "");
  return text.trim() || fallback;
}

function readJsonHeader<T>(headers: Headers, name: string): T | undefined {
  const raw = headers.get(name);
  if (!raw) return undefined;
  try {
    return JSON.parse(raw) as T;
  } catch {
    return undefined;
  }
}

export function WorshipBuilderProvider({ children }: { children: ReactNode }) {
  const [dataFile, setDataFile] = useState<SelectedFile>(null);
  const [lastWeekFile, setLastWeekFile] = useState<SelectedFile>(null);
  const [serviceDate, setServiceDate] = useState(today);
  const [isGenerating, setIsGenerating] = useState(false);
  const [error, setError] = useState("");
  const [validationMessage, setValidationMessage] = useState("");
  const [result, setResult] = useState<GenerationResult | null>(null);
  const [preview, setPreview] = useState<WorshipSlidePreview[]>([]);
  const [oneDriveStatus, setOneDriveStatus] = useState<OneDriveStatus | null>(null);
  const [useOneDrive, setUseOneDrive] = useState(true);
  const [uploadToOneDrive, setUploadToOneDrive] = useState(true);

  useEffect(() => {
    let cancelled = false;
    const controller = new AbortController();
    const timer = window.setTimeout(() => controller.abort(), 8000);

    fetch(apiUrl("/api/onedrive/status"), {
      cache: "no-store",
      signal: controller.signal,
    })
      .then(async (response) => {
        if (!response.ok) throw new Error("OneDrive status unavailable");
        return (await response.json()) as OneDriveStatus;
      })
      .then((status) => {
        if (cancelled) return;
        setOneDriveStatus(status);
        if (!status.configured) {
          setUseOneDrive(false);
          setUploadToOneDrive(false);
        } else {
          setUseOneDrive(status.auto_enabled);
          setUploadToOneDrive(status.upload_enabled);
        }
      })
      .catch(() => {
        if (!cancelled) setOneDriveStatus(null);
      })
      .finally(() => window.clearTimeout(timer));

    return () => {
      cancelled = true;
      controller.abort();
      window.clearTimeout(timer);
    };
  }, []);

  const canPreview = useMemo(
    () => Boolean(dataFile && serviceDate && preview.length > 0 && !isGenerating),
    [dataFile, serviceDate, preview.length, isGenerating],
  );

  async function readWorshipData(file: File) {
    const raw = await file.text();
    return parseWorshipData(raw);
  }

  async function handleDataFileChange(file: SelectedFile) {
    setDataFile(file);
    setResult(null);
    setValidationMessage("");

    if (!file) {
      setError("");
      setPreview([]);
      return;
    }

    try {
      const data = await readWorshipData(file);

      const formData = new FormData();
      formData.append("data", file);
      const response = await fetch(apiUrl("/api/validate"), {
        method: "POST",
        body: formData,
      });

      if (!response.ok) {
        throw new Error(await readResponseError(response, "Validation failed."));
      }

      setError("");
      setPreview(buildPreview(data));
      setValidationMessage("JSON file validated successfully.");
    } catch (caughtError) {
      setPreview([]);
      setValidationMessage("");
      setError(caughtError instanceof Error ? caughtError.message : "Invalid JSON file.");
    }
  }

  async function generatePpt() {
    if (!dataFile || !serviceDate || isGenerating) return false;

    setIsGenerating(true);
    setError("");
    setResult(null);

    try {
      await readWorshipData(dataFile);

      const formData = new FormData();
      formData.append("data", dataFile);
      formData.append("date", serviceDate);
      formData.append("slides", JSON.stringify(preview));
      formData.append(
        "use_onedrive",
        String(lastWeekFile ? false : useOneDrive),
      );
      formData.append("upload_to_onedrive", String(uploadToOneDrive));
      if (lastWeekFile) {
        formData.append("last_week_pptx", lastWeekFile);
      }

      const response = await fetch(apiUrl("/api/generate"), {
        method: "POST",
        body: formData,
      });

      if (!response.ok) {
        throw new Error(
          await readResponseError(response, "Failed to generate PowerPoint."),
        );
      }

      const blob = await response.blob();
      const fileName = `Worship_${serviceDate}.pptx`;
      const url = URL.createObjectURL(blob);
      const link = document.createElement("a");
      link.href = url;
      link.download = fileName;
      link.click();
      URL.revokeObjectURL(url);

      const rawWarnings = response.headers.get("X-Worship-Warnings");
      const warnings = rawWarnings ? (JSON.parse(rawWarnings) as string[]) : [];
      const mattersCopied =
        Number(response.headers.get("X-Worship-Matters-Copied") || "0") || 0;
      const lastWeek = readJsonHeader<LastWeekMetadata>(
        response.headers,
        "X-Worship-Last-Week-Metadata",
      );
      const upload = readJsonHeader<UploadMetadata>(
        response.headers,
        "X-Worship-Upload-Metadata",
      );
      setResult({ fileName, warnings, mattersCopied, lastWeek, upload });
      return true;
    } catch (caughtError) {
      setError(
        caughtError instanceof Error ? caughtError.message : "Failed to generate PowerPoint.",
      );
      return false;
    } finally {
      setIsGenerating(false);
    }
  }

  function clearResult() {
    setResult(null);
  }

  function updatePreviewSlide(id: string, title: string, lines: string[]) {
    setPreview((current) =>
      current.map((slide) =>
        slide.id === id ? { ...slide, title, lines } : slide,
      ),
    );
    setResult(null);
  }

  function relabelSlides(slides: WorshipSlidePreview[]) {
    const totals = slides.reduce<Record<string, number>>((counts, slide) => {
      counts[slide.sectionLabel] = (counts[slide.sectionLabel] || 0) + 1;
      return counts;
    }, {});
    const positions: Record<string, number> = {};

    return slides.map((slide) => {
      positions[slide.sectionLabel] = (positions[slide.sectionLabel] || 0) + 1;
      return {
        ...slide,
        slideLabel: `${positions[slide.sectionLabel]} / ${totals[slide.sectionLabel]}`,
      };
    });
  }

  function deletePreviewSlide(id: string) {
    setPreview((current) => relabelSlides(current.filter((slide) => slide.id !== id)));
    setResult(null);
  }

  function movePreviewSlide(id: string, direction: -1 | 1) {
    setPreview((current) => {
      const sourceIndex = current.findIndex((slide) => slide.id === id);
      if (sourceIndex < 0) return current;

      const sectionLabel = current[sourceIndex].sectionLabel;
      const targetIndex = sourceIndex + direction;
      if (
        targetIndex < 0
        || targetIndex >= current.length
        || current[targetIndex].sectionLabel !== sectionLabel
      ) {
        return current;
      }

      const reordered = [...current];
      [reordered[sourceIndex], reordered[targetIndex]] = [
        reordered[targetIndex],
        reordered[sourceIndex],
      ];
      return relabelSlides(reordered);
    });
    setResult(null);
  }

  function insertPreviewSlide(afterId: string) {
    const newId = `inserted-${Date.now()}-${Math.random().toString(36).slice(2)}`;
    setPreview((current) => {
      const sourceIndex = current.findIndex((slide) => slide.id === afterId);
      if (sourceIndex < 0) return current;

      const source = current[sourceIndex];
      const inserted: WorshipSlidePreview = {
        ...source,
        id: newId,
        lines: [],
      };
      const next = [...current];
      next.splice(sourceIndex + 1, 0, inserted);
      return relabelSlides(next);
    });
    setResult(null);
    return newId;
  }

  function resetBuilder() {
    setDataFile(null);
    setLastWeekFile(null);
    setServiceDate(today);
    setError("");
    setValidationMessage("");
    setResult(null);
    setPreview([]);
    setUseOneDrive(oneDriveStatus ? oneDriveStatus.auto_enabled : true);
    setUploadToOneDrive(oneDriveStatus ? oneDriveStatus.upload_enabled : true);
  }

  return (
    <WorshipBuilderContext.Provider
      value={{
        dataFile,
        lastWeekFile,
        serviceDate,
        isGenerating,
        error,
        validationMessage,
        result,
        preview,
        canPreview,
        oneDriveStatus,
        useOneDrive,
        uploadToOneDrive,
        setServiceDate,
        setLastWeekFile,
        setUseOneDrive,
        setUploadToOneDrive,
        handleDataFileChange,
        updatePreviewSlide,
        deletePreviewSlide,
        movePreviewSlide,
        insertPreviewSlide,
        generatePpt,
        clearResult,
        resetBuilder,
      }}
    >
      {children}
    </WorshipBuilderContext.Provider>
  );
}

export function useWorshipBuilder() {
  const context = useContext(WorshipBuilderContext);

  if (!context) {
    throw new Error("useWorshipBuilder must be used within WorshipBuilderProvider.");
  }

  return context;
}
