import {
  createContext,
  useContext,
  useMemo,
  useState,
  type ReactNode,
} from "react";
import {
  buildPreview,
  parseWorshipData,
  type GenerationResult,
  type SelectedFile,
  type WorshipSlidePreview,
} from "../lib/worship-data";
import { apiUrl } from "../lib/app-config";

const today = new Date().toISOString().slice(0, 10);

type WorshipBuilderContextValue = {
  dataFile: SelectedFile;
  serviceDate: string;
  isGenerating: boolean;
  error: string;
  validationMessage: string;
  result: GenerationResult | null;
  preview: WorshipSlidePreview[];
  canPreview: boolean;
  setServiceDate: (value: string) => void;
  handleDataFileChange: (file: SelectedFile) => Promise<void>;
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

export function WorshipBuilderProvider({ children }: { children: ReactNode }) {
  const [dataFile, setDataFile] = useState<SelectedFile>(null);
  const [serviceDate, setServiceDate] = useState(today);
  const [isGenerating, setIsGenerating] = useState(false);
  const [error, setError] = useState("");
  const [validationMessage, setValidationMessage] = useState("");
  const [result, setResult] = useState<GenerationResult | null>(null);
  const [preview, setPreview] = useState<WorshipSlidePreview[]>([]);

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
      setResult({ fileName, warnings });
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

  function resetBuilder() {
    setDataFile(null);
    setServiceDate(today);
    setError("");
    setValidationMessage("");
    setResult(null);
    setPreview([]);
  }

  return (
    <WorshipBuilderContext.Provider
      value={{
        dataFile,
        serviceDate,
        isGenerating,
        error,
        validationMessage,
        result,
        preview,
        canPreview,
        setServiceDate,
        handleDataFileChange,
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
