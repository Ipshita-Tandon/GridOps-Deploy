import { createRoot } from "react-dom/client";
import App from "./App.tsx";
import "./index.css";
import { registerLicense } from "@syncfusion/ej2-base";

registerLicense(
  "ORg4AjUWIQA/Gnt2XFhhQlJHfVhdWHxLflFzVWFTel96dFZWESFaRnZdR11lSXtTdEBkXHhdeXNUTWJV"
);

createRoot(document.getElementById("root")!).render(<App />);
