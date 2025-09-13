import { NextRequest, NextResponse } from "next/server";
import ExcelJS from "exceljs";
import {
  detectDataLength,
  metricsHeaders,
  newRawDataHeaders,
  generateRawDataFormulas,
  generateMetricsFormulas,
  buildDataRanges,
  cleanAndCopyData,
  applyCTRCTORFormatting,
  applyCurrencyFormatting,
  applyMetricsFormatting,
  detectCurrencyFromData,
} from "@/lib/excel-helpers";

export async function POST(request: NextRequest) {
  try {
    const formData = await request.formData();
    const file = formData.get("file") as File;

    if (!file) {
      return NextResponse.json({ error: "No file provided" }, { status: 400 });
    }

    // Read the uploaded Excel file
    const buffer = await file.arrayBuffer();
    const inputWorkbook = new ExcelJS.Workbook();
    await inputWorkbook.xlsx.load(buffer);

    // Get the first worksheet (assuming it contains the data)
    const inputWorksheet = inputWorkbook.getWorksheet(1);
    if (!inputWorksheet) {
      return NextResponse.json(
        { error: "No worksheet found in the file" },
        { status: 400 },
      );
    }

    // Detect data start row and length
    const { startRow, lastRow } = detectDataLength(inputWorksheet);

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
    cleanAndCopyData(inputWorksheet, cleanDataSheet, startRow, lastRow);

    // Calculate clean data dimensions (always starts at row 2)
    const cleanDataRows = lastRow - startRow + 1;
    const cleanLastRow = cleanDataRows + 1; // +1 because clean data starts at row 2

    // Add new headers to clean data sheet (columns X, Y, Z, AA)
    newRawDataHeaders.forEach((header, index) => {
      const columns = ["X", "Y", "Z", "AA"];
      const col = columns[index];
      cleanDataSheet.getCell(`${col}1`).value = header;
    });

    // Add formulas for the new columns in clean data (using row 2+ as data start)
    if (cleanLastRow > 1) {
      for (let row = 2; row <= cleanLastRow; row++) {
        const formulas = generateRawDataFormulas(row);

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
    const seenDates = new Set();

    if (cleanLastRow > 1) {
      const dataRanges = buildDataRanges(2, cleanLastRow); // Clean data always starts at row 2

      // Scan through clean data to count unique dates
      for (let row = 2; row <= cleanLastRow; row++) {
        const startTimeCell = cleanDataSheet.getCell(`B${row}`);
        if (startTimeCell.value && startTimeCell.value !== "") {
          // Convert datetime to date string
          let dateValue;
          if (startTimeCell.value instanceof Date) {
            dateValue = startTimeCell.value.toISOString().split("T")[0];
          } else {
            // Handle other date formats
            const dateObj = new Date(startTimeCell.value as string | number);
            if (!isNaN(dateObj.getTime())) {
              dateValue = dateObj.toISOString().split("T")[0];
            }
          }

          if (dateValue && !seenDates.has(dateValue)) {
            seenDates.add(dateValue);
            uniqueDatesCount++;
          }
        }
      }

      const metricsLastRow = uniqueDatesCount; // +1 for header row

      // Special case for B2 - unique dates formula
      metricsSheet.getCell("B2").value = {
        formula: `IFERROR(SORT(UNIQUE(FILTER(${dataRanges.startDateR},${dataRanges.startDateR}<>""))),"")`,
      };

      // Add formulas only for rows that have unique dates (rows 2 to metricsLastRow)
      for (let row = 2; row <= metricsLastRow; row++) {
        const formulas = generateMetricsFormulas(row, 2, cleanLastRow); // Clean data starts at row 2
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

    // Apply cell formatting
    console.log("Applying cell formatting...");

    // Detect currency from original data
    const detectedCurrency = detectCurrencyFromData(
      inputWorksheet,
      startRow,
      lastRow,
    );
    console.log(`Detected currency: ${detectedCurrency}`);

    // Format clean data sheet
    if (cleanLastRow > 1) {
      // Apply CTR/CTOR percentage formatting (columns V and W)
      applyCTRCTORFormatting(cleanDataSheet, 2, cleanLastRow);

      // Apply currency formatting to revenue columns
      applyCurrencyFormatting(
        cleanDataSheet,
        2,
        cleanLastRow,
        detectedCurrency,
      );
    }

    // Format metrics sheet
    if (cleanLastRow > 1) {
      const metricsFormatLastRow = uniqueDatesCount;
      if (metricsFormatLastRow > 1) {
        applyMetricsFormatting(
          metricsSheet,
          2,
          metricsFormatLastRow,
          detectedCurrency,
        );
      }
    }

    // Generate the output Excel buffer
    const outputBuffer = await outputWorkbook.xlsx.writeBuffer();

    // Return the processed file
    return new NextResponse(outputBuffer, {
      status: 200,
      headers: {
        "Content-Type":
          "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        "Content-Disposition": `attachment; filename="processed_${file.name}"`,
      },
    });
  } catch (error) {
    console.error("Error processing Excel file:", error);
    return NextResponse.json(
      { error: "Failed to process Excel file" },
      { status: 500 },
    );
  }
}
