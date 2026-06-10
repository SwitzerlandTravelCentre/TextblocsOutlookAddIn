import React from "react";
import { createRoot } from "react-dom/client";
import App from "./App";
import "./styles.css";

const render = () => {
  const root = document.getElementById("root");

  if (!root) {
    throw new Error("Task pane root element was not found.");
  }

  createRoot(root).render(
    <React.StrictMode>
      <App />
    </React.StrictMode>
  );
};

if (typeof Office !== "undefined") {
  Office.onReady(() => render());
} else {
  render();
}
