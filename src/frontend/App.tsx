import { Navigate, Route, Routes } from "react-router-dom";
import { WorshipBuilderProvider } from "./hooks/useWorshipBuilder";
import { GuidePage } from "./pages/GuidePage";
import { HomePage } from "./pages/HomePage";
import { PreviewPage } from "./pages/PreviewPage";

export function App() {
  return (
    <WorshipBuilderProvider>
      <Routes>
        <Route path="/" element={<HomePage />} />
        <Route path="/preview" element={<PreviewPage />} />
        <Route path="/guide" element={<GuidePage />} />
        <Route path="*" element={<Navigate to="/" replace />} />
      </Routes>
    </WorshipBuilderProvider>
  );
}
