import React from "react";
import ReactDOM from "react-dom/client";
import { PowerProvider } from "./PowerProvider";
import App from "./App";

ReactDOM.createRoot(
    document.getElementById("root") as HTMLElement
).render(
    <React.StrictMode>
        <PowerProvider>
            <App />
        </PowerProvider>
    </React.StrictMode>
);