import * as React from "react";
import { createRoot } from "react-dom/client";
import App from "./App";

/* global console, document, Excel, Office */

const rootElement: HTMLElement = document.getElementById("container");
const root = createRoot(rootElement);

Office.onReady((info) => {
  root.render(<App />);
});
