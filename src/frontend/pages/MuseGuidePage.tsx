import { MuseGuide } from "../components/MuseGuide";
import { TopBar } from "../components/TopBar";

export function MuseGuidePage() {
  return (
    <main className="app-shell simple-shell">
      <TopBar />
      <MuseGuide />
    </main>
  );
}
