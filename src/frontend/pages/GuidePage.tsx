import { ProjectGuide } from "../components/ProjectGuide";
import { TopBar } from "../components/TopBar";

export function GuidePage() {
  return (
    <main className="app-shell simple-shell">
      <TopBar />
      <ProjectGuide />
    </main>
  );
}
