import { Link } from "react-router-dom";
import { downloadUrl } from "../lib/app-config";

export function TopBar() {
  return (
    <header className="topbar">
      <Link className="brand" to="/">
        BSBC Chinese Congregation
      </Link>
      <nav className="resource-nav" aria-label="Resources">
        <a href={downloadUrl("/downloads/template")}>Template</a>
        <a href={downloadUrl("/downloads/sample-json")}>Sample JSON</a>
        <Link to="/guide">ChatGPT Guide</Link>
      </nav>
    </header>
  );
}
