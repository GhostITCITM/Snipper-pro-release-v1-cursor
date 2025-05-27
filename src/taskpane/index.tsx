import React from "react";
import { createRoot } from "react-dom/client";
import App from "./App";

Office.onReady(() => {
  console.log("Office is ready. Starting React app...");

  const container = document.getElementById("container");
  if (!container) {
    throw new Error("Container element not found");
  }

  const root = createRoot(container);
  root.render(<App />);
});
