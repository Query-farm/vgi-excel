import React from "react";
import { createRoot } from "react-dom/client";
import { App } from "./App";
import { initializeTelemetry, TelemetryErrorBoundary } from "./telemetry";
import "./functions";
import "./styles.css";

function render(): void {
  const root = document.getElementById("root");
  if (root) createRoot(root).render(<React.StrictMode><TelemetryErrorBoundary fallback={<main className="recovery"><section><h1>Cupola needs to restart</h1><p>Close and reopen the Cupola task pane.</p></section></main>}><App/></TelemetryErrorBoundary></React.StrictMode>);
}

initializeTelemetry();
if (typeof Office !== "undefined") Office.onReady(render);
else render();
