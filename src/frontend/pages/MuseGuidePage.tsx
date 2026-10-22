import { ProjectGuide } from "../components/ProjectGuide";
import { TopBar } from "../components/TopBar";

export function MuseGuidePage() {
  return (
    <main className="app-shell simple-shell">
      <TopBar />
      <ProjectGuide provider="muse" />
    </main>
  );
}
