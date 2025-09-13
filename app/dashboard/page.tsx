"use client";

import {
  Breadcrumb,
  BreadcrumbItem,
  BreadcrumbLink,
  BreadcrumbList,
  BreadcrumbPage,
  BreadcrumbSeparator,
} from "@/components/ui/breadcrumb";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { Separator } from "@/components/ui/separator";
import { SidebarInset, SidebarTrigger } from "@/components/ui/sidebar";
import { Upload, Download, FileSpreadsheet, AlertCircle } from "lucide-react";
import { useState } from "react";

export default function Page() {
  const [file, setFile] = useState<File | null>(null);
  const [isProcessing, setIsProcessing] = useState(false);
  const [error, setError] = useState<string | null>(null);

  const handleFileSelect = (event: React.ChangeEvent<HTMLInputElement>) => {
    const selectedFile = event.target.files?.[0];
    if (selectedFile) {
      // Validate file type
      if (
        !selectedFile.name.endsWith(".xlsx") &&
        !selectedFile.name.endsWith(".xls")
      ) {
        setError("Please select an Excel file (.xlsx or .xls)");
        setFile(null);
        return;
      }
      setFile(selectedFile);
      setError(null);
    }
  };

  const handleProcess = async () => {
    if (!file) return;

    setIsProcessing(true);
    setError(null);

    try {
      const formData = new FormData();
      formData.append("file", file);

      const response = await fetch("/api/process-excel", {
        method: "POST",
        body: formData,
      });

      if (!response.ok) {
        const errorData = await response.json();
        throw new Error(errorData.error || "Failed to process file");
      }

      // Download the processed file
      const blob = await response.blob();
      const url = window.URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.style.display = "none";
      a.href = url;
      a.download = `processed_${file.name}`;
      document.body.appendChild(a);
      a.click();
      window.URL.revokeObjectURL(url);
      document.body.removeChild(a);

      // Reset form
      setFile(null);
      const fileInput = document.getElementById(
        "file-input",
      ) as HTMLInputElement;
      if (fileInput) fileInput.value = "";
    } catch (err) {
      setError(err instanceof Error ? err.message : "An error occurred");
    } finally {
      setIsProcessing(false);
    }
  };

  return (
    <SidebarInset>
      <header className="flex h-16 shrink-0 items-center gap-2">
        <div className="flex items-center gap-2 px-4">
          <SidebarTrigger className="-ml-1" />
          <Separator
            orientation="vertical"
            className="mr-2 data-[orientation=vertical]:h-4"
          />
          <Breadcrumb>
            <BreadcrumbList>
              <BreadcrumbItem className="hidden md:block">
                <BreadcrumbLink href="/">Home</BreadcrumbLink>
              </BreadcrumbItem>
              <BreadcrumbSeparator className="hidden md:block" />
              <BreadcrumbItem>
                <BreadcrumbPage>Dashboard</BreadcrumbPage>
              </BreadcrumbItem>
            </BreadcrumbList>
          </Breadcrumb>
        </div>
      </header>

      <div className="flex flex-1 flex-col gap-6 p-6">
        {/* Excel Processor Card */}
        <div className="rounded-xl border bg-card p-6">
          <div className="flex items-center gap-3 mb-6">
            <FileSpreadsheet className="h-6 w-6 text-primary" />
            <div>
              <h2 className="text-2xl font-semibold">
                Excel Metrics Processor
              </h2>
              <p className="text-muted-foreground">
                Upload your livestream data to generate comprehensive metrics
                and analysis
              </p>
            </div>
          </div>

          <div className="space-y-4">
            {/* File Upload */}
            <div className="space-y-2">
              <label htmlFor="file-input" className="text-sm font-medium">
                Select Excel File
              </label>
              <div className="flex items-center gap-3">
                <Input
                  id="file-input"
                  type="file"
                  accept=".xlsx,.xls"
                  onChange={handleFileSelect}
                  className="flex-1"
                />
                <Button
                  onClick={handleProcess}
                  disabled={!file || isProcessing}
                  className="min-w-[120px]"
                >
                  {isProcessing ? (
                    <>
                      <div className="mr-2 h-4 w-4 animate-spin rounded-full border-2 border-current border-t-transparent" />
                      Processing...
                    </>
                  ) : (
                    <>
                      <Download className="mr-2 h-4 w-4" />
                      Process
                    </>
                  )}
                </Button>
              </div>
            </div>

            {/* File Info */}
            {file && (
              <div className="rounded-lg bg-muted p-3">
                <div className="flex items-center gap-2">
                  <Upload className="h-4 w-4 text-green-600" />
                  <span className="text-sm font-medium">Selected file:</span>
                  <span className="text-sm">{file.name}</span>
                  <span className="text-xs text-muted-foreground">
                    ({(file.size / 1024).toFixed(1)} KB)
                  </span>
                </div>
              </div>
            )}

            {/* Error Display */}
            {error && (
              <div className="rounded-lg bg-destructive/10 border border-destructive/20 p-3">
                <div className="flex items-center gap-2">
                  <AlertCircle className="h-4 w-4 text-destructive" />
                  <span className="text-sm text-destructive font-medium">
                    Error:
                  </span>
                  <span className="text-sm text-destructive">{error}</span>
                </div>
              </div>
            )}
          </div>
        </div>

        {/* Instructions Card */}
        <div className="rounded-xl border bg-card p-6">
          <h3 className="text-lg font-semibold mb-4">How to Use</h3>
          <div className="space-y-3 text-sm text-muted-foreground">
            <div className="flex items-start gap-3">
              <div className="rounded-full bg-primary/10 text-primary w-6 h-6 flex items-center justify-center text-xs font-medium mt-0.5">
                1
              </div>
              <div>
                <strong>Upload your Excel file</strong> containing livestream
                data with the expected headers: Livestream, Start time,
                Duration, Gross revenue, Direct GMV, Items sold, etc.
              </div>
            </div>
            <div className="flex items-start gap-3">
              <div className="rounded-full bg-primary/10 text-primary w-6 h-6 flex items-center justify-center text-xs font-medium mt-0.5">
                2
              </div>
              <div>
                <strong>Click Process</strong> to generate a new workbook with
                three sheets: raw data (unchanged), metrics (summary analysis),
                and helper (hidden calculations).
              </div>
            </div>
            <div className="flex items-start gap-3">
              <div className="rounded-full bg-primary/10 text-primary w-6 h-6 flex items-center justify-center text-xs font-medium mt-0.5">
                3
              </div>
              <div>
                <strong>Download the processed file</strong> with comprehensive
                metrics including revenue analysis, engagement stats, and prime
                time calculations.
              </div>
            </div>
          </div>
        </div>

        {/* Features Overview */}
        <div className="grid md:grid-cols-3 gap-4">
          <div className="rounded-xl border bg-card p-4">
            <h4 className="font-semibold mb-2">📊 Metrics Analysis</h4>
            <p className="text-sm text-muted-foreground">
              Comprehensive daily metrics including revenue stats, engagement
              rates, and performance indicators.
            </p>
          </div>
          <div className="rounded-xl border bg-card p-4">
            <h4 className="font-semibold mb-2">⏰ Prime Time Detection</h4>
            <p className="text-sm text-muted-foreground">
              Automatically calculates optimal streaming times based on revenue,
              CTR, and CTOR performance.
            </p>
          </div>
          <div className="rounded-xl border bg-card p-4">
            <h4 className="font-semibold mb-2">📈 Formula-Based</h4>
            <p className="text-sm text-muted-foreground">
              All calculations use Excel formulas with absolute ranges, ensuring
              accuracy and transparency.
            </p>
          </div>
        </div>
      </div>
    </SidebarInset>
  );
}
