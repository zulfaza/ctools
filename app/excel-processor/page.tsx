"use client";

import { useState } from "react";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { Upload, Download, FileSpreadsheet } from "lucide-react";

export default function ExcelProcessorPage() {
  const [file, setFile] = useState<File | null>(null);
  const [processing, setProcessing] = useState(false);
  const [downloadUrl, setDownloadUrl] = useState<string | null>(null);

  const handleFileSelect = (event: React.ChangeEvent<HTMLInputElement>) => {
    const selectedFile = event.target.files?.[0];
    if (selectedFile) {
      // Validate file type
      if (
        !selectedFile.name.toLowerCase().endsWith(".xlsx") &&
        !selectedFile.name.toLowerCase().endsWith(".xls")
      ) {
        alert("Please select an Excel file (.xlsx or .xls)");
        return;
      }
      setFile(selectedFile);
      setDownloadUrl(null); // Clear previous download
    }
  };

  const processFile = async () => {
    if (!file) return;

    setProcessing(true);
    try {
      const formData = new FormData();
      formData.append("file", file);

      const response = await fetch("/api/process-excel", {
        method: "POST",
        body: formData,
      });

      if (!response.ok) {
        throw new Error("Failed to process file");
      }

      const blob = await response.blob();
      const url = URL.createObjectURL(blob);
      setDownloadUrl(url);
    } catch (error) {
      console.error("Error processing file:", error);
      alert("Error processing file. Please try again.");
    } finally {
      setProcessing(false);
    }
  };

  const downloadProcessedFile = () => {
    if (downloadUrl) {
      const link = document.createElement("a");
      link.href = downloadUrl;
      link.download = `processed_${file?.name || "metrics.xlsx"}`;
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
    }
  };

  return (
    <div className="container mx-auto px-4 py-8 max-w-2xl">
      <div className="space-y-6">
        <div className="text-center">
          <FileSpreadsheet className="mx-auto h-12 w-12 text-blue-600 mb-4" />
          <h1 className="text-3xl font-bold text-gray-900 dark:text-white">
            Excel Metrics Processor
          </h1>
          <p className="text-gray-600 dark:text-gray-300 mt-2">
            Upload your Excel file to generate metrics with raw data, helper
            calculations, daily summaries, and overall summary statistics
          </p>
        </div>

        <div className="border-2 border-dashed border-gray-300 dark:border-gray-600 rounded-lg p-6">
          <div className="text-center">
            <Upload className="mx-auto h-8 w-8 text-gray-400 mb-4" />
            <div className="space-y-4">
              <Input
                type="file"
                accept=".xlsx,.xls"
                onChange={handleFileSelect}
                className="cursor-pointer"
              />
              {file && (
                <div className="text-sm text-gray-600 dark:text-gray-300">
                  Selected: {file.name} ({(file.size / 1024 / 1024).toFixed(2)}{" "}
                  MB)
                </div>
              )}
            </div>
          </div>
        </div>

        <div className="space-y-4">
          <Button
            onClick={processFile}
            disabled={!file || processing}
            className="w-full"
            size="lg"
          >
            {processing ? (
              <>
                <div className="animate-spin rounded-full h-4 w-4 border-b-2 border-white mr-2"></div>
                Processing...
              </>
            ) : (
              <>
                <FileSpreadsheet className="mr-2 h-4 w-4" />
                Process Excel File
              </>
            )}
          </Button>

          {downloadUrl && (
            <Button
              onClick={downloadProcessedFile}
              variant="outline"
              className="w-full"
              size="lg"
            >
              <Download className="mr-2 h-4 w-4" />
              Download Processed File
            </Button>
          )}
        </div>

        <div className="bg-gray-50 dark:bg-gray-800 rounded-lg p-4">
          <h3 className="font-semibold text-gray-900 dark:text-white mb-2">
            What this tool does:
          </h3>
          <ul className="text-sm text-gray-600 dark:text-gray-300 space-y-1">
            <li>
              • Creates 4 sheets: raw data (unchanged), clean data (normalized),
              metrics (daily summaries), and summary (overall averages)
            </li>
            <li>
              • Generates formulas with absolute ranges based on your data
              length
            </li>
            <li>• Processes Indonesian livestream metrics data</li>
            <li>
              • All numeric outputs are raw numbers (format in Excel later)
            </li>
          </ul>
        </div>
      </div>
    </div>
  );
}
