import * as React from "react";
import { createRoot } from "react-dom/client";
import { App } from "./App";
import "./styles/app.css";

// No React.StrictMode: it double-invokes effects, which would open the mock workbook
// twice and double-subscribe to host events.
createRoot(document.getElementById("root")!).render(<App />);
