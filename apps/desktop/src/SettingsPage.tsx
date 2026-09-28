import { ArrowLeft } from "lucide-react";
import { useRef, type ReactNode } from "react";
export type SettingsSection = "connections" | "ai" | "about";
const sections: Array<{ id: SettingsSection; label: string }> = [{ id: "connections", label: "Connections" }, { id: "ai", label: "AI settings" }, { id: "about", label: "About" }];

export function SettingsPage({ section, onSection, onBack, connections, aiTarget, about }: { section: SettingsSection; onSection(value: SettingsSection): void; onBack(): void; connections: ReactNode; aiTarget(node: HTMLDivElement | null): void; about: ReactNode }): React.JSX.Element {
  const tabs = useRef<Array<HTMLButtonElement | null>>([]);
  return <section className="settings-page" aria-labelledby="settings-title">
    <div className="section-heading settings-heading"><button type="button" className="settings-back" aria-label="Back to workspace" title="Back to workspace" onClick={onBack}><ArrowLeft size={18} aria-hidden="true"/></button><h2 id="settings-title">Settings</h2></div>
    <nav className="settings-tabs" role="tablist" aria-label="Settings sections">{sections.map((item, index) => <button key={item.id} ref={node => { tabs.current[index] = node; }} id={`settings-tab-${item.id}`} role="tab" aria-selected={section === item.id} aria-controls={`settings-${item.id}`} tabIndex={section === item.id ? 0 : -1} onClick={() => onSection(item.id)} onKeyDown={event => {
      if (!["ArrowLeft", "ArrowRight", "Home", "End"].includes(event.key)) return;
      event.preventDefault();
      const next = event.key === "Home" ? 0 : event.key === "End" ? sections.length - 1 : (index + (event.key === "ArrowRight" ? 1 : -1) + sections.length) % sections.length;
      onSection(sections[next].id); tabs.current[next]?.focus();
    }}>{item.label}</button>)}</nav>
    <div className="settings-content">
      <div id="settings-connections" role="tabpanel" aria-labelledby="settings-tab-connections" hidden={section !== "connections"}>{connections}</div>
      <div id="settings-ai" role="tabpanel" aria-labelledby="settings-tab-ai" hidden={section !== "ai"} ref={aiTarget}/>
      <div id="settings-about" role="tabpanel" aria-labelledby="settings-tab-about" hidden={section !== "about"}>{about}</div>
    </div>
  </section>;
}
