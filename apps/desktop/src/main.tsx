import React from "react";
import ReactDOM from "react-dom/client";
import { App } from "./App";
import { initializeTelemetry, TelemetryErrorBoundary } from "./telemetry";
import "./styles.css";

initializeTelemetry();
ReactDOM.createRoot(document.getElementById("root")!).render(<React.StrictMode><TelemetryErrorBoundary fallback={<main className="recovery"><section><h1>Cupola needs to restart</h1><p>Close this window and open Cupola from the Excel ribbon again.</p></section></main>}><App /></TelemetryErrorBoundary></React.StrictMode>);
