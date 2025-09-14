"use client";

import { Button } from "@/components/ui/button";
import { Separator } from "@/components/ui/separator";
import { SidebarInset, SidebarTrigger } from "@/components/ui/sidebar";
import {
  Upload,
  Download,
  FileSpreadsheet,
  AlertCircle,
  X,
  Plus,
} from "lucide-react";
import { useState, useRef, useCallback } from "react";

interface FileWithId {
  id: string;
  file: File;
  status: "pending" | "processing" | "completed" | "error";
  error?: string;
}

export default function Page() {
  const [files, setFiles] = useState<FileWithId[]>([]);
  const [isProcessing, setIsProcessing] = useState(false);
  const [isDragOver, setIsDragOver] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const fileInputRef = useRef<HTMLInputElement>(null);

  const validateFile = (file: File): string | null => {
    if (!file.name.endsWith(".xlsx") && !file.name.endsWith(".xls")) {
      return "Please select an Excel file (.xlsx or .xls)";
    }
    if (file.size > 50 * 1024 * 1024) {
      // 50MB limit
      return "File size must be less than 50MB";
    }
    return null;
  };

  const addFiles = useCallback(
    (newFiles: File[]) => {
      const validFiles: FileWithId[] = [];
      const errors: string[] = [];

      newFiles.forEach((file) => {
        const validationError = validateFile(file);
        if (validationError) {
          errors.push(`${file.name}: ${validationError}`);
        } else if (
          !files.some(
            (f) => f.file.name === file.name && f.file.size === file.size,
          )
        ) {
          validFiles.push({
            id: `${file.name}-${Date.now()}-${Math.random()}`,
            file,
            status: "pending",
          });
        }
      });

      if (errors.length > 0) {
        setError(errors.join("; "));
      } else {
        setError(null);
      }

      if (validFiles.length > 0) {
        setFiles((prev) => [...prev, ...validFiles]);
      }
    },
    [files],
  );

  const handleFileSelect = (event: React.ChangeEvent<HTMLInputElement>) => {
    const selectedFiles = Array.from(event.target.files || []);
    if (selectedFiles.length > 0) {
      addFiles(selectedFiles);
    }
  };

  const handleDragOver = useCallback((e: React.DragEvent) => {
    e.preventDefault();
    setIsDragOver(true);
  }, []);

  const handleDragLeave = useCallback((e: React.DragEvent) => {
    e.preventDefault();
    setIsDragOver(false);
  }, []);

  const handleDrop = useCallback(
    (e: React.DragEvent) => {
      e.preventDefault();
      setIsDragOver(false);

      const droppedFiles = Array.from(e.dataTransfer.files);
      addFiles(droppedFiles);
    },
    [addFiles],
  );

  const removeFile = (id: string) => {
    setFiles((prev) => prev.filter((f) => f.id !== id));
  };

  const clearAllFiles = () => {
    setFiles([]);
    setError(null);
  };

  const handleProcess = async () => {
    if (files.length === 0) return;

    setIsProcessing(true);
    setError(null);

    try {
      const formData = new FormData();

      // Add all files to form data
      files.forEach((fileItem, index) => {
        formData.append(`file${index}`, fileItem.file);
      });

      const response = await fetch("/api/process-excel", {
        method: "POST",
        body: formData,
      });

      if (!response.ok) {
        const errorData = await response.json();
        throw new Error(errorData.error || "Failed to process files");
      }

      // Download the processed file(s)
      const blob = await response.blob();
      const url = window.URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.style.display = "none";
      a.href = url;

      if (files.length === 1) {
        a.download = `processed_${files[0].file.name}`;
      } else {
        a.download = "processed_excel_files.zip";
      }

      document.body.appendChild(a);
      a.click();
      window.URL.revokeObjectURL(url);
      document.body.removeChild(a);

      // Reset form
      clearAllFiles();
      if (fileInputRef.current) {
        fileInputRef.current.value = "";
      }
    } catch (err) {
      setError(err instanceof Error ? err.message : "An error occurred");
    } finally {
      setIsProcessing(false);
    }
  };

  return (
    <SidebarInset>
      <div className="flex flex-1 flex-col gap-6 p-6">
        <div className="rounded-xl border bg-card p-6">
          <div className="flex items-center gap-3 mb-6">
            <FileSpreadsheet className="h-6 w-6 text-primary" />
            <div>
              <h2 className="text-2xl font-semibold">
                Excel Metrics Processor
              </h2>
              <p className="text-muted-foreground">
                Upload single or multiple Excel files to generate comprehensive
                metrics and analysis
              </p>
            </div>
          </div>

          <div className="space-y-4">
            {/* Drag and Drop Upload Area */}
            <div className="space-y-2">
              <label className="text-sm font-medium">Select Excel Files</label>
              <div
                className={`relative border-2 border-dashed rounded-lg transition-colors ${
                  isDragOver
                    ? "border-primary bg-primary/5"
                    : "border-muted-foreground/25 hover:border-primary/50"
                }`}
                onDragOver={handleDragOver}
                onDragLeave={handleDragLeave}
                onDrop={handleDrop}
              >
                <input
                  ref={fileInputRef}
                  type="file"
                  accept=".xlsx,.xls"
                  multiple
                  onChange={handleFileSelect}
                  className="absolute inset-0 w-full h-full opacity-0 cursor-pointer"
                />
                <div className="flex flex-col items-center justify-center py-12 px-6">
                  <div className="rounded-full bg-primary/10 p-3 mb-4">
                    <Upload className="h-6 w-6 text-primary" />
                  </div>
                  <div className="text-center">
                    <p className="text-lg font-medium mb-1">
                      Drop your Excel files here
                    </p>
                    <p className="text-sm text-muted-foreground mb-4">
                      or click to browse files
                    </p>
                    <Button
                      variant="outline"
                      size="sm"
                      className="pointer-events-none"
                    >
                      <Plus className="mr-2 h-4 w-4" />
                      Choose Files
                    </Button>
                  </div>
                </div>
              </div>
            </div>

            {/* File List */}
            {files.length > 0 && (
              <div className="space-y-2">
                <div className="flex items-center justify-between">
                  <h3 className="text-sm font-medium">
                    Selected Files ({files.length})
                  </h3>
                  <Button
                    variant="ghost"
                    size="sm"
                    onClick={clearAllFiles}
                    className="text-xs"
                  >
                    Clear All
                  </Button>
                </div>
                <div className="space-y-2 max-h-40 overflow-y-auto">
                  {files.map((fileItem) => (
                    <div
                      key={fileItem.id}
                      className="flex items-center gap-3 p-3 rounded-lg bg-muted/50 border"
                    >
                      <FileSpreadsheet className="h-4 w-4 text-green-600 flex-shrink-0" />
                      <div className="flex-1 min-w-0">
                        <p className="text-sm font-medium truncate">
                          {fileItem.file.name}
                        </p>
                        <p className="text-xs text-muted-foreground">
                          {(fileItem.file.size / 1024).toFixed(1)} KB
                        </p>
                      </div>
                      <Button
                        variant="ghost"
                        size="sm"
                        onClick={() => removeFile(fileItem.id)}
                        className="h-8 w-8 p-0 flex-shrink-0"
                      >
                        <X className="h-4 w-4" />
                      </Button>
                    </div>
                  ))}
                </div>
              </div>
            )}

            {/* Process Button */}
            {files.length > 0 && (
              <Button
                onClick={handleProcess}
                disabled={isProcessing}
                className="w-full"
                size="lg"
              >
                {isProcessing ? (
                  <>
                    <div className="mr-2 h-4 w-4 animate-spin rounded-full border-2 border-current border-t-transparent" />
                    Processing {files.length} file
                    {files.length !== 1 ? "s" : ""}...
                  </>
                ) : (
                  <>
                    <Download className="mr-2 h-4 w-4" />
                    Process {files.length} file{files.length !== 1 ? "s" : ""}
                  </>
                )}
              </Button>
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
                <strong>Upload your Excel file(s)</strong> containing livestream
                data with the expected headers: Livestream, Start time,
                Duration, Gross revenue, Direct GMV, Items sold, etc. You can
                select multiple files at once.
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
                <strong>Download the processed file(s)</strong> with
                comprehensive metrics including revenue analysis, engagement
                stats, and prime time calculations. Multiple files are returned
                as a ZIP archive.
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
