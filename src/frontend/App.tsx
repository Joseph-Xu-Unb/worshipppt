import { Navigate, Route, Routes } from "react-router-dom";
import { WorshipBuilderProvider } from "./hooks/useWorshipBuilder";
import { MuseGuidePage } from "./pages/MuseGuidePage";
import { HomePage } from "./pages/HomePage";
import { PreviewPage } from "./pages/PreviewPage";

export function App() {
  return (
    <WorshipBuilderProvider>
      <Routes>
        <Route path="/" element={<HomePage />} />
        <Route path="/preview" element={<PreviewPage />} />
        <Route path="/guide" element={<Navigate to="/guide/muse" replace />} />
        <Route path="/guide/muse" element={<MuseGuidePage />} />
        <Route path="*" element={<Navigate to="/" replace />} />
      </Routes>
    </WorshipBuilderProvider>
  );
}
