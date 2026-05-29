import { startTransition, useMemo, useState, type FormEvent } from "react";
import {
  buildPreview,
  parseWorshipData,
  type GenerationResult,
  type SelectedFile,
  type WorshipSectionPreview,
} from "../lib/worship-data";

const today = new Date().toISOString().slice(0, 10);

type UseWorshipBuilderResult = {
  templateFile: SelectedFile;
  dataFile: SelectedFile;
  serviceDate: string;
  isGenerating: boolean;
  error: string;
  result: GenerationResult | null;
  preview: WorshipSectionPreview[];
  canGenerate: boolean;
  setServiceDate: (value: string) => void;
  setTemplateFile: (file: SelectedFile) => void;
  handleDataFileChange: (file: SelectedFile) => Promise<void>;
  handleSubmit: (event: FormEvent<HTMLFormElement>) => Promise<void>;
};

export function useWorshipBuilder(): UseWorshipBuilderResult {
  const [templateFile, setTemplateFile] = useState<SelectedFile>(null);
  const [dataFile, setDataFile] = useState<SelectedFile>(null);
  const [serviceDate, setServiceDate] = useState(today);
  const [isGenerating, setIsGenerating] = useState(false);
  const [error, setError] = useState("");
  const [result, setResult] = useState<GenerationResult | null>(null);
  const [preview, setPreview] = useState<WorshipSectionPreview[]>([]);

  const canGenerate = useMemo(
    () => Boolean(templateFile && dataFile && serviceDate && !isGenerating),
    [templateFile, dataFile, serviceDate, isGenerating],
  );

  async function readWorshipData(file: File) {
    const raw = await file.text();
    return parseWorshipData(raw);
  }

  async function handleDataFileChange(file: SelectedFile) {
    setDataFile(file);
    setResult(null);

    if (!file) {
      setError("");
      startTransition(() => setPreview([]));
      return;
    }

    try {
      const data = await readWorshipData(file);
      setError("");
      startTransition(() => setPreview(buildPreview(data)));
    } catch (caughtError) {
      startTransition(() => setPreview([]));
      setError(caughtError instanceof Error ? caughtError.message : "Invalid JSON file.");
    }
  }

  async function handleSubmit(event: FormEvent<HTMLFormElement>) {
    event.preventDefault();
    if (!templateFile || !dataFile || !serviceDate || isGenerating) return;

    setIsGenerating(true);
    setError("");
    setResult(null);

    try {
      await readWorshipData(dataFile);

      const formData = new FormData();
      formData.append("template", templateFile);
      formData.append("data", dataFile);
      formData.append("date", serviceDate);

      const response = await fetch("/api/generate", {
        method: "POST",
        body: formData,
      });

      if (!response.ok) {
        const payload = (await response.json().catch(() => ({}))) as { error?: string };
        throw new Error(payload.error || "Failed to generate PowerPoint.");
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
    } catch (caughtError) {
      setError(
        caughtError instanceof Error ? caughtError.message : "Failed to generate PowerPoint.",
      );
    } finally {
      setIsGenerating(false);
    }
  }

  return {
    templateFile,
    dataFile,
    serviceDate,
    isGenerating,
    error,
    result,
    preview,
    canGenerate,
    setServiceDate,
    setTemplateFile,
    handleDataFileChange,
    handleSubmit,
  };
}
