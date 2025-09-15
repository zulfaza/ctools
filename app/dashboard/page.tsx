import { Metadata } from "next";
import { ExcelProcessor } from "./excel-processor";

export const metadata: Metadata = {
  title: "Dashboard - Excel Metrics Processor",
  description:
    "Upload and process Excel files to generate comprehensive metrics and analysis for livestream data.",
};

export default function Page() {
  return <ExcelProcessor />;
}
