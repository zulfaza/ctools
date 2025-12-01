import { NextRequest, NextResponse } from "next/server";
import ExcelJS from "exceljs";
import JSZip from "jszip";
import Papa from "papaparse";
import {
  ExcelFormat,
  detectExcelFormat,
  detectDataLength,
  metricsHeaders,
  tiktokNewRawDataHeaders,
  shopeeNewRawDataHeaders,
  generateRawDataFormulas,
  generateMetricsFormulas,
  buildDataRanges,
  cleanAndCopyData,
  applyPercentageFormatting,
  applyCurrencyFormatting,
  applyDurationFormatting,
  applyMetricsFormatting,
  detectCurrencyFromData,
  getAdvancedCurrencyFormat,
  summaryHeaders,
  summaryMetrics,
  generateSummaryFormulas,
  applySummaryFormatting,
  trendHeaders,
  generateTrendFormulas,
  applyTrendFormatting,
} from "@/lib/excel-helpers";
import { TZDate } from "@date-fns/tz";

const MONTH_MAP_STRING = [
  "Jan",
  "Feb",
  "Mar",
  "Apr",
  "May",
  "Jun",
  "Jul",
  "Aug",
  "Sep",
  "Oct",
  "Nov",
  "Dec",
];

async function convertCsvToExcelWorkbook(
  csvBuffer: ArrayBuffer,
): Promise<ExcelJS.Workbook> {
  const csvText = new TextDecoder().decode(csvBuffer);
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet("Sheet1");

  return new Promise((resolve, reject) => {
    Papa.parse(csvText, {
      header: false,
      skipEmptyLines: true,
      complete: (results) => {
        try {
          // Add data to worksheet
          results.data.forEach((row: unknown, rowIndex: number) => {
            const worksheetRow = worksheet.getRow(rowIndex + 1);
            (row as string[]).forEach((cellValue: string, colIndex: number) => {
              worksheetRow.getCell(colIndex + 1).value = cellValue;
            });
            worksheetRow.commit();
          });
          resolve(workbook);
        } catch (error) {
          reject(error);
        }
      },
      error: (error: Error) => {
        reject(error);
      },
    });
  });
}

async function processExcelFile(
  file: File,
): Promise<{ buffer: ArrayBuffer; filename: string }> {
  // Read the uploaded file
  const buffer = await file.arrayBuffer();
  let inputWorkbook: ExcelJS.Workbook;

  // Check if it's a CSV file
  if (file.name.toLowerCase().endsWith(".csv")) {
    // Convert CSV to Excel workbook
    inputWorkbook = await convertCsvToExcelWorkbook(buffer);
  } else {
    // Load Excel file
    inputWorkbook = new ExcelJS.Workbook();
    await inputWorkbook.xlsx.load(buffer);
  }

  // Get the first worksheet (assuming it contains the data)
  const inputWorksheet = inputWorkbook.getWorksheet(1);
  if (!inputWorksheet) {
    throw new Error("No worksheet found in the file");
  }

  // Detect Excel format first
  const format = detectExcelFormat(inputWorksheet);
  if (format === ExcelFormat.UNSUPPORTED) {
    throw new Error(
      "Unsupported file format. Supported formats: TikTok Livestream, Shopee Livestream (Monthly/Daily CSV)",
    );
  }

  // Detect data start row and length
  const { startRow, lastRow } = detectDataLength(
    inputWorksheet,
    undefined,
    format,
  );

  console.log(
    `Data starts at row ${startRow}, last row = ${lastRow}, ${lastRow - startRow + 1} data rows`,
  );

  // Create new workbook with 3 sheets
  const outputWorkbook = new ExcelJS.Workbook();

  // Sheet 1: Raw data (copy original data as-is)
  const rawDataSheet = outputWorkbook.addWorksheet("raw data");

  // Sheet 2: Clean data (cleaned and standardized data starting at row 2)
  const cleanDataSheet = outputWorkbook.addWorksheet("clean data");

  // Copy all data from input worksheet
  inputWorksheet.eachRow((row, rowNumber) => {
    const newRow = rawDataSheet.getRow(rowNumber);
    row.eachCell((cell, colNumber) => {
      const newCell = newRow.getCell(colNumber);
      newCell.value = cell.value;
      // Copy basic styling if present
      if (cell.style) {
        newCell.style = cell.style;
      }
    });
    newRow.commit();
  });

  // Clean and copy data to clean data sheet
  cleanAndCopyData(inputWorksheet, cleanDataSheet, startRow, lastRow, format);

  // Calculate clean data dimensions (always starts at row 2)
  const cleanDataRows = lastRow - startRow + 1;
  const cleanLastRow = cleanDataRows + 1; // +1 because clean data starts at row 2

  // Detect Shopee format variant for conditional formula generation
  let isShopeeDailyFormat = false;
  if (format === ExcelFormat.SHOPEE_MONTHLY) {
    const headerRow = startRow - 1;
    const cellD = inputWorksheet.getCell(`D${headerRow}`);
    const valueD = cellD.value?.toString().toLowerCase() || "";
    isShopeeDailyFormat = !(valueD.includes("nama livestream") || valueD.includes("livestream"));
  }

  // Add new headers to clean data sheet based on format
  const formatHeaders =
    format === ExcelFormat.SHOPEE_MONTHLY
      ? shopeeNewRawDataHeaders
      : tiktokNewRawDataHeaders;
  const startColumns =
    format === ExcelFormat.SHOPEE_MONTHLY
      ? ["R", "S", "T", "U", "V", "W"]
      : ["X", "Y", "Z", "AA", "AB", "AC"];

  formatHeaders.forEach((header, index) => {
    const col = startColumns[index];
    cleanDataSheet.getCell(`${col}1`).value = header;
  });

  // Add formulas for the new columns in clean data (using row 2+ as data start)
  if (cleanLastRow > 1) {
    for (let row = 2; row <= cleanLastRow; row++) {
      let formulas: Record<string, string>;
      if (format === ExcelFormat.SHOPEE_MONTHLY) {
        // Generate Shopee formulas based on variant
        if (isShopeeDailyFormat) {
          // Daily CSV format: Start date from Periode Data, no start/end time
          // Extract date directly from Periode Data column (A) - format is DD-MM-YYYY
          const periodeDataCell = cleanDataSheet.getCell(`A${row}`);
          const periodeDataValue = periodeDataCell.value;
          
          let startDateValue: string = "";
          if (periodeDataValue) {
            if (typeof periodeDataValue === "string") {
              // String like "01-11-2025" -> parse as DD-MM-YYYY (1st November 2025)
              const datePart = periodeDataValue.split(" ")[0]; // In case there's time part
              const [day, month, year] = datePart.split("-");
              if (day && month && year) {
                // Parse as DD-MM-YYYY format
                const dateObj = new Date(parseInt(year), parseInt(month) - 1, parseInt(day));
                // Use TZDate to ensure correct timezone handling
                startDateValue = new TZDate(dateObj, "Asia/Jakarta").toISOString().split("T")[0];
              } else {
                startDateValue = datePart;
              }
            } else if (periodeDataValue instanceof Date) {
              startDateValue = new TZDate(periodeDataValue, "Asia/Jakarta").toISOString().split("T")[0];
            } else if (typeof periodeDataValue === "number") {
              // Excel serial date
              const dateObj = new Date((periodeDataValue - 25569) * 86400 * 1000); // Convert Excel serial to JS date
              startDateValue = new TZDate(dateObj, "Asia/Jakarta").toISOString().split("T")[0];
            }
          }
          
          // Set Start Date (R) directly
          const rCell = cleanDataSheet.getCell(`R${row}`);
          rCell.value = startDateValue;
          
          // Use formulas for other columns
          formulas = {
            S: `""`, // Start Time (not used for daily)
            T: `""`, // End Time (not used for daily)
            U: `IFERROR(IF(R${row}<>"",WEEKNUM(R${row},2),""),"")`, // Week in Year
            V: `IFERROR(IF(R${row}<>"",MONTH(R${row}),""),"")`, // Month Index
            W: `0`, // GMV/hour (not applicable for daily format)
          };
        } else {
          // Monthly format: Extract Start Date, Start Time, and calculate End Time directly from raw data
          const rawDataRow = startRow + (row - 2); // Map clean row to original raw data row
          
          // Get Start Time value from raw data (column E)
          const startTimeCell = inputWorksheet.getCell(`E${rawDataRow}`);
          const startTimeValue = startTimeCell.value;
          
          // Extract Start Date (R)
          let startDateValue: number | string = "";
          if (startTimeValue) {
            if (startTimeValue instanceof Date) {
              startDateValue = new TZDate(startTimeValue, "Asia/Jakarta").toISOString().split("T")[0];
            } else if (typeof startTimeValue === "number") {
              // Excel serial number - use integer part (date only)
              startDateValue = new TZDate(startTimeValue, "Asia/Jakarta").toISOString().split("T")[0];
            } else if (typeof startTimeValue === "string") {
              // String like "29-11-2025 12:58" -> extract "29-11-2025"
              const datePart = startTimeValue.split(" ")[0];
              // String like "01-11-2025" -> parse as DD-MM-YYYY (1st November 2025)
              const [day, month, year] = datePart.split("-");
              if (day && month && year) {
                // Parse as DD-MM-YYYY format
                const dateObj = new Date(parseInt(year), parseInt(month) - 1, parseInt(day));
                // Use TZDate to ensure correct timezone handling
                startDateValue = new TZDate(dateObj, "Asia/Jakarta").toISOString().split("T")[0];
              } else{
                startDateValue = datePart ;
              }
             
            }
          }
          
          // Extract Start Time (S) - "29-11-2025 12:58" -> "12:58"
          let startTimeOnly: number | null = null;
          if (startTimeValue) {
            if (startTimeValue instanceof Date) {
              // Extract time portion as fraction of day
              const hours = startTimeValue.getHours();
              const minutes = startTimeValue.getMinutes();
              startTimeOnly = (hours * 60 + minutes) / 1440; // Convert to fraction of day
            } else if (typeof startTimeValue === "number") {
              // Excel serial number - extract fractional part (time)
              startTimeOnly = startTimeValue - Math.floor(startTimeValue);
            } else if (typeof startTimeValue === "string") {
              // String like "29-11-2025 12:58" -> extract "12:58"
              const timePart = startTimeValue.split(" ")[1];
              if (timePart) {
                const [hours, minutes] = timePart.split(":").map(Number);
                if (!isNaN(hours) && !isNaN(minutes)) {
                  startTimeOnly = (hours * 60 + minutes) / 1440; // Convert to fraction of day
                }
              }
            }
          }
          
          // Calculate End Time (T) = Start Time + Duration
          // Use null check instead of truthy check to handle 00:00 (which is 0, falsy)
          let endTimeValue: number | string = "";
          const durasiCell = inputWorksheet.getCell(`F${rawDataRow}`);
          const durasiValue = durasiCell.value;
          
          if (startTimeOnly !== null && durasiValue) {
            let durasiInDays = 0;
            if (typeof durasiValue === "number") {
              // Durasi in seconds, convert to days
              durasiInDays = durasiValue / 86400;
            } else if (typeof durasiValue === "string") {
              // Durasi in time format "HH:MM:SS" or "HH:MM"
              const timeParts = durasiValue.split(":");
              if (timeParts.length >= 2) {
                const hours = parseInt(timeParts[0]) || 0;
                const minutes = parseInt(timeParts[1]) || 0;
                const seconds = timeParts[2] ? parseInt(timeParts[2]) : 0;
                durasiInDays = (hours * 3600 + minutes * 60 + seconds) / 86400;
              }
            }
            endTimeValue = startTimeOnly + durasiInDays;
          } else if (startTimeOnly !== null) {
            // If no duration, end time equals start time
            endTimeValue = startTimeOnly;
          }
          
          // Set Start Date (R), Start Time (S), and End Time (T) directly
          const rCell = cleanDataSheet.getCell(`R${row}`);
          rCell.value = startDateValue;
          
          const sCell = cleanDataSheet.getCell(`S${row}`);
          sCell.value = startTimeOnly !== null ? startTimeOnly : "";
          sCell.numFmt = "HH:mm"; // Format as 24-hour time
          
          const tCell = cleanDataSheet.getCell(`T${row}`);
          tCell.value = endTimeValue;
          tCell.numFmt = "HH:mm"; // Format as 24-hour time
          
          // Use formulas for other columns
          formulas = {
            U: `IFERROR(IF(R${row}<>"",WEEKNUM(R${row},2),""),"")`, // Week in Year
            V: `IFERROR(IF(R${row}<>"",MONTH(R${row}),""),"")`, // Month Index
            // GMV/hour = Penjualan(Pesanan Siap Dikirim) / Durasi (in hours)
            W: `IFERROR(IF('raw data'!F${rawDataRow}>0,IF(ISNUMBER('raw data'!F${rawDataRow}),D${row}/('raw data'!F${rawDataRow}/3600),D${row}/('raw data'!F${rawDataRow}*24)),0),0)`, // GMV/hour
          };
        }
      } else {
        formulas = generateRawDataFormulas(row, format);
      }

      // Set formulas for each new column
      Object.entries(formulas).forEach(([col, formula]) => {
        const cell = cleanDataSheet.getCell(`${col}${row}`);
        cell.value = { formula };
      });
    }
  }

  // Sheet 3: Metrics (summary table)
  const metricsSheet = outputWorkbook.addWorksheet("metrics");

  // Add metrics headers
  metricsHeaders.forEach((header, index) => {
    metricsSheet.getCell(1, index + 1).value = header;
  });

  // Add metrics formulas (now referencing clean data)
  // Count unique dates from the clean data (column B - Start time)
  let uniqueDatesCount = 0;
  const seenDates = new Set<string>();
  const uniqueMonths = new Set<number>();

  if (cleanLastRow > 1) {
    const dataRanges = buildDataRanges(2, cleanLastRow, format); // Clean data always starts at row 2

    // Scan through clean data to count unique dates based on format
    // For Shopee, use the Start Date column (R) which is calculated from Periode Data
    const dateColumn = format === ExcelFormat.SHOPEE_MONTHLY ? "R" : "B";
    for (let row = 2; row <= cleanLastRow; row++) {
      const dateCell = cleanDataSheet.getCell(`${dateColumn}${row}`);
      if (dateCell.value && dateCell.value !== "") {
        // Convert datetime to date string

        let dateValue;
        if (dateCell.value instanceof Date) {
          uniqueMonths.add(dateCell.value.getMonth());
          dateValue = dateCell.value.toISOString().split("T")[0];
        } else if (
          typeof dateCell.value === "string" ||
          typeof dateCell.value === "number"
        ) {
          const dateObj =
            typeof dateCell.value === "string" //use trinary to fix type error
              ? new TZDate(dateCell.value)
              : new Date(dateCell.value);

          if (!isNaN(dateObj.getTime())) {
            uniqueMonths.add(dateObj.getMonth());
            dateValue = dateObj.toISOString().split("T")[0];
          }
        }
        if (dateValue && !seenDates.has(dateValue)) {
          seenDates.add(dateValue);
          uniqueDatesCount++;
        }
      }
    }

    const metricsLastRow = uniqueDatesCount + 1; // +1 because metrics starts at row 2

    // Special case for B2 - unique dates formula
    metricsSheet.getCell("B2").value = {
      formula: `IFERROR(SORT(UNIQUE(FILTER(${dataRanges.startDateR},${dataRanges.startDateR}<>""))),"")`,
    };

    // Add formulas only for rows that have unique dates (rows 2 to metricsLastRow)
    for (let row = 2; row <= metricsLastRow; row++) {
      const formulas = generateMetricsFormulas(row, 2, cleanLastRow, format); // Clean data starts at row 2
      // Set formulas for each column (skip B since it's handled specially)
      Object.entries(formulas).forEach(([col, formula]) => {
        if (col !== "B" || row === 2) {
          // Only set B2 once
          const cell = metricsSheet.getCell(`${col}${row}`);
          cell.value = { formula };
        }
      });
    }
  }

  // Detect currency from original data
  const detectedCurrency = detectCurrencyFromData(
    inputWorksheet,
    startRow,
    lastRow,
  );

  if (cleanLastRow > 1) {
    const metricsFormatLastRow = uniqueDatesCount;
    // Apply percentage formatting to all percentage columns based on format definition
    applyPercentageFormatting(cleanDataSheet, 2, cleanLastRow, format);
    // Apply currency formatting to revenue columns based on format definition
    applyCurrencyFormatting(cleanDataSheet, 2, cleanLastRow, detectedCurrency, format);
    // Apply duration formatting to duration columns (decimal hours with 2 decimal places)
    applyDurationFormatting(cleanDataSheet, 2, cleanLastRow, format);
    // Format GMV/Hour column as currency
    const currencyFormat = getAdvancedCurrencyFormat(detectedCurrency);
    if (format === ExcelFormat.TIKTOK_LIVESTREAM) {
      // Format GMV/Hour column (AC) as currency for TikTok format
      for (let row = 2; row <= cleanLastRow; row++) {
        const acCell = cleanDataSheet.getCell(`AC${row}`);
        acCell.numFmt = currencyFormat;
      }
    } else if (format === ExcelFormat.SHOPEE_MONTHLY) {
      // Format GMV/hour column (W) as currency for Shopee format
      for (let row = 2; row <= cleanLastRow; row++) {
        const wCell = cleanDataSheet.getCell(`W${row}`);
        wCell.numFmt = currencyFormat;
      }
      // Format Start Date column (R) as date for Shopee format
      // Use DD-MM-YYYY format to match the input data format
      for (let row = 2; row <= cleanLastRow; row++) {
        const rCell = cleanDataSheet.getCell(`R${row}`);
        rCell.numFmt = "dd-mm-yyyy"; // Format as DD-MM-YYYY
      }
    }

    if (metricsFormatLastRow > 1) {
      applyMetricsFormatting(
        metricsSheet,
        2,
        metricsFormatLastRow,
        detectedCurrency,
      );
    }
  }

  // Sheet 4: Summary (overall metrics)
  const summarySheet = outputWorkbook.addWorksheet("summary");

  // Add summary headers
  summaryHeaders.forEach((header, index) => {
    summarySheet.getCell(1, index + 1).value = header;
  });

  // Add summary metrics labels and formulas
  if (cleanLastRow > 1) {
    const summaryFormulas = generateSummaryFormulas(2, cleanLastRow, format);

    summaryMetrics.forEach((metric, index) => {
      const row = index + 2; // Start from row 2
      summarySheet.getCell(`A${row}`).value = metric;
      summarySheet.getCell(`B${row}`).value = {
        formula: summaryFormulas[metric],
      };
    });

    // Apply summary formatting
    applySummaryFormatting(summarySheet, detectedCurrency);
  }

  // Sheet 5: Trend (monthly metrics)
  const trendSheet = outputWorkbook.addWorksheet("trend");

  // Add trend headers starting at B2
  trendHeaders.forEach((header, index) => {
    trendSheet.getCell(1, index + 1).value = header;
  });

  if (cleanLastRow > 1) {
    const sortedMonths = Array.from(uniqueMonths).sort();
    let row = 0;
    // Add months to column A starting from A2
    sortedMonths.forEach((month, index) => {
      row = index + 2;

      trendSheet.getCell(`A${row}`).value = month + 1;
      trendSheet.getCell(`B${row}`).value = MONTH_MAP_STRING[month];

      // Add trend formulas for this month
      const formulas = generateTrendFormulas(row, 2, cleanLastRow, format);
      Object.entries(formulas).forEach(([col, formula]) => {
        const cell = trendSheet.getCell(`${col}${row}`);
        cell.value = { formula };
      });
    });

    applyTrendFormatting(trendSheet, 2, row);
  }

  // Generate the output Excel buffer
  const outputBuffer = await outputWorkbook.xlsx.writeBuffer();

  // Always output as XLSX format
  const baseName = file.name.replace(/\.(xlsx?|csv)$/i, "");
  return {
    buffer: outputBuffer,
    filename: `${baseName}-processed.xlsx`,
  };
}

export async function POST(request: NextRequest) {
  try {
    const formData = await request.formData();

    // Get all files from form data
    const files: File[] = [];
    for (const [key, value] of formData.entries()) {
      if (value instanceof File && key.startsWith("file")) {
        files.push(value);
      }
    }

    // Also check for single file (backward compatibility)
    const singleFile = formData.get("file") as File;
    if (singleFile && !files.length) {
      files.push(singleFile);
    }

    if (!files.length) {
      return NextResponse.json({ error: "No files provided" }, { status: 400 });
    }

    // Process single file case
    if (files.length === 1) {
      const result = await processExcelFile(files[0]);
      return new NextResponse(result.buffer, {
        status: 200,
        headers: {
          "Content-Type":
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
          "Content-Disposition": `attachment; filename="${result.filename}"`,
        },
      });
    }

    // Process multiple files case - create zip
    const zip = new JSZip();
    const processedFiles: { buffer: ArrayBuffer; filename: string }[] = [];

    // Process all files
    for (const file of files) {
      try {
        const result = await processExcelFile(file);
        processedFiles.push(result);
      } catch (error) {
        console.error(`Error processing file ${file.name}:`, error);
        // Add error file to results with error message
        processedFiles.push({
          buffer: new TextEncoder().encode(
            `Error processing ${file.name}: ${error instanceof Error ? error.message : "Unknown error"}`,
          ).buffer,
          filename: `ERROR_${file.name}.txt`,
        });
      }
    }

    // Add all processed files to zip
    for (const { buffer, filename } of processedFiles) {
      zip.file(filename, buffer);
    }

    // Generate zip buffer
    const zipBuffer = await zip.generateAsync({ type: "arraybuffer" });

    // Return the zip file
    return new NextResponse(zipBuffer, {
      status: 200,
      headers: {
        "Content-Type": "application/zip",
        "Content-Disposition": `attachment; filename="processed_excel_files.zip"`,
      },
    });
  } catch (error) {
    console.error("Error processing Excel files:", error);
    return NextResponse.json(
      { error: "Failed to process Excel files" },
      { status: 500 },
    );
  }
}
