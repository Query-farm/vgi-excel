import React from "react";
import { createRoot } from "react-dom/client";
import { App } from "./App";
import "./functions";
import "./styles.css";

function render(): void {
  const root = document.getElementById("root");
  if (root) createRoot(root).render(<React.StrictMode><App/></React.StrictMode>);
}

if (typeof Office !== "undefined") Office.onReady(render);
else render();
