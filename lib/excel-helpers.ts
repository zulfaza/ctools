import { Worksheet } from "exceljs";

export enum ExcelFormat {
  TIKTOK_LIVESTREAM = "TIKTOK_LIVESTREAM",
  SHOPEE_MONTHLY = "SHOPEE_MONTHLY",
  UNSUPPORTED = "UNSUPPORTED",
}

/**
 * Column cleaning rule type
 */
export type ColumnCleaningRule = "currency" | "percentage" | "numeric" | "text" | "duration";

/**
 * Column cleaning rules mapping (0-based column index -> cleaning rule)
 */
export type ColumnCleaningRules = Record<number, ColumnCleaningRule>;

/**
 * Detection result for format detection
 */
export interface FormatDetectionResult {
  formatId: ExcelFormat;
  startRow: number;
}

/**
 * Excel format definition interface.
 * 
 * To add a new format:
 * 1. Create an object implementing this interface with all required properties
 * 2. Register it using `registerExcelFormat()` or add it to the `EXCEL_FORMATS` array
 * 3. Add the format ID to the `ExcelFormat` enum
 * 
 * Example:
 * const myFormat: ExcelFormatDefinition = {
 *   id: ExcelFormat.MY_FORMAT,
 *   headers: ["Header1", "Header2"],
 *   columnCleaningRules: { 0: "text", 1: "numeric" },
 *   newRawDataHeaders: ["Date", "Time"],
 *   newRawDataColumns: ["X", "Y"],
 *   detect: (worksheet) => { return 2; },
 *   buildDataRanges: (startRow, lastRow) => ({ startDateR: "...", ... }),
 *   generateRawDataFormulas: (row) => ({ X: "FORMULA", ... }),
 *   generateMetricsFormulas: (row, startRow, lastRow) => ({ A: "FORMULA", ... }),
 *   generateSummaryFormulas: (startRow, lastRow) => ({ "Metric": "FORMULA", ... }),
 *   generateTrendFormulas: (row, startRow, lastRow) => ({ C: "FORMULA", ... }),
 * };
 */
export interface ExcelFormatDefinition {
  /** Unique format identifier */
  id: ExcelFormat;
  
  /** Column headers for this format */
  headers: readonly string[];
  
  /** Column cleaning rules (0-based index -> cleaning rule type) */
  columnCleaningRules: ColumnCleaningRules;
  
  /** Headers for new raw data columns to append */
  newRawDataHeaders: readonly string[];
  
  /** Column letters where new raw data headers should be placed */
  newRawDataColumns: readonly string[];
  
  /**
   * Detect if this format matches the worksheet and return the data start row.
   * Returns the start row number if detected, or null/undefined if not detected.
   */
  detect: (worksheet: Worksheet) => number | null;
  
  /**
   * Build data range strings for formulas
   */
  buildDataRanges: (startRow: number, lastRow: number) => DataRangeStrings;
  
  /**
   * Generate raw data formulas for a specific row
   */
  generateRawDataFormulas: (row: number) => Record<string, string>;
  
  /**
   * Generate metrics formulas for a specific row
   */
  generateMetricsFormulas: (
    row: number,
    startRow: number,
    lastRow: number,
  ) => Record<string, string>;
  
  /**
   * Generate summary formulas
   */
  generateSummaryFormulas: (
    startRow: number,
    lastRow: number,
  ) => Record<string, string>;
  
  /**
   * Generate trend formulas for a specific row
   */
  generateTrendFormulas: (
    row: number,
    startRow: number,
    lastRow: number,
  ) => Record<string, string>;
}

export interface RangeStrings {
  dateR: string;
  timeTextR: string;
  timeSerialR: string;
  weekR: string;
  dayR: string;
  gmvR: string;
  itemsR: string;
  avgViewDurR: string;
  ctrR: string;
  ctorR: string;
  likesR: string;
  commentsR: string;
  sharesR: string;
  startSerialR: string;
}

export interface DataRangeStrings {
  startDateR: string;
  startTimeR: string;
  endTimeR: string;
  weekInYearR: string;
  montIndex: string;
  grossR: string;
  directR: string;
  itemsR: string;
  avgViewR: string;
  likesR: string;
  commentsR: string;
  sharesR: string;
  ctrR: string;
  ctorR: string;
  gmv1kShowsR: string;
}

/**
 * Registry of all registered Excel formats
 */
const EXCEL_FORMATS_REGISTRY: ExcelFormatDefinition[] = [];

/**
 * Register a new Excel format definition.
 * Formats are checked in registration order during detection.
 * 
 * @param formatDef The format definition to register
 */
export function registerExcelFormat(formatDef: ExcelFormatDefinition): void {
  // Prevent duplicate registrations
  if (EXCEL_FORMATS_REGISTRY.some((f) => f.id === formatDef.id)) {
    console.warn(
      `Format ${formatDef.id} is already registered. Skipping duplicate registration.`,
    );
    return;
  }
  EXCEL_FORMATS_REGISTRY.push(formatDef);
}

/**
 * Get a format definition by its ID
 * 
 * @param formatId The format identifier
 * @returns The format definition or undefined if not found
 */
export function getFormatDefinition(
  formatId: ExcelFormat,
): ExcelFormatDefinition | undefined {
  return EXCEL_FORMATS_REGISTRY.find((f) => f.id === formatId);
}

/**
 * Get all registered format definitions
 * 
 * @returns Array of all registered format definitions
 */
export function getAvailableFormats(): ExcelFormatDefinition[] {
  return [...EXCEL_FORMATS_REGISTRY];
}

/**
 * Detect the Excel format from a worksheet by checking all registered formats
 * 
 * @param worksheet The worksheet to detect
 * @returns Detection result with format ID and start row, or UNSUPPORTED if no match
 */
function detectFormat(
  worksheet: Worksheet,
): FormatDetectionResult {
  for (const formatDef of EXCEL_FORMATS_REGISTRY) {
    const startRow = formatDef.detect(worksheet);
    if (startRow !== null && startRow !== undefined) {
      return {
        formatId: formatDef.id,
        startRow,
      };
    }
  }
  return {
    formatId: ExcelFormat.UNSUPPORTED,
    startRow: 2, // Default fallback
  };
}

// Import format definitions from separate files
import { tiktokFormat } from "./excel-formats/tiktok";
import { shopeeFormat } from "./excel-formats/shopee";

// Register built-in formats
registerExcelFormat(tiktokFormat);
registerExcelFormat(shopeeFormat);

// Export format-specific constants for backward compatibility
// Note: The actual exported constants are defined later in the file (around line 839+)
// We reference the internal versions here for the format definitions

// Backward compatibility exports will be defined after the constants are exported

export function buildDataRanges(
  startRow: number,
  lastRow: number,
  format: ExcelFormat = ExcelFormat.TIKTOK_LIVESTREAM,
): DataRangeStrings {
  const formatDef = getFormatDefinition(format);
  if (!formatDef) {
    // Fallback to TikTok format if format not found
    const defaultFormat = getFormatDefinition(ExcelFormat.TIKTOK_LIVESTREAM);
    if (!defaultFormat) {
      throw new Error(`Format ${format} not found and no default format available`);
    }
    return defaultFormat.buildDataRanges(startRow, lastRow);
  }
  return formatDef.buildDataRanges(startRow, lastRow);
}

export function detectExcelFormat(worksheet: Worksheet): ExcelFormat {
  const result = detectFormat(worksheet);
  return result.formatId;
}

export function detectDataStartRow(
  worksheet: Worksheet,
  format?: ExcelFormat,
): number {
  if (format) {
    // If format is provided, use its detection to find start row
    const formatDef = getFormatDefinition(format);
    if (formatDef) {
      const startRow = formatDef.detect(worksheet);
      if (startRow !== null && startRow !== undefined) {
        return startRow;
      }
    }
    // If format provided but detection fails, default to row 2
    return 2;
  }
  
  // If no format provided, use full detection
  const result = detectFormat(worksheet);
  return result.startRow;
}

export function detectDataLength(
  worksheet: Worksheet,
  startRow?: number,
  format?: ExcelFormat,
): { startRow: number; lastRow: number } {
  const detectedFormat = format || detectExcelFormat(worksheet);
  const dataStartRow =
    startRow || detectDataStartRow(worksheet, detectedFormat);
  let dataRows = 0;
  let row = dataStartRow;

  while (true) {
    const cellA = worksheet.getCell(`A${row}`);
    const cellB = worksheet.getCell(`B${row}`);

    // Check if both key columns are empty based on format
    if (
      (!cellA.value || cellA.value === "") &&
      (!cellB.value || cellB.value === "")
    ) {
      break;
    }

    dataRows++;
    row++;
  }

  return {
    startRow: dataStartRow,
    lastRow: dataStartRow + dataRows - 1,
  };
}

// Import format constants for re-export
import {
  tiktokHeaders,
  tiktokColumnCleaningRules,
  tiktokNewRawDataHeaders,
} from "./excel-formats/tiktok";
import {
  shopeeHeaders,
  shopeeColumnCleaningRules,
  shopeeNewRawDataHeaders,
} from "./excel-formats/shopee";

// Re-export format constants for backward compatibility
export {
  tiktokHeaders,
  tiktokColumnCleaningRules,
  tiktokNewRawDataHeaders,
  shopeeHeaders,
  shopeeColumnCleaningRules,
  shopeeNewRawDataHeaders,
};

// Backward compatibility
export const inputHeaders = tiktokHeaders;

// Data cleaning functions
export function cleanCurrencyValue(value: unknown): number {
  if (typeof value === "number") return value;
  if (!value) return 0;

  const str = value.toString();
  // Remove currency symbols, thousand separators, and extra spaces
  const cleaned = str
    .replace(/[Rp$€¥£₹₽¢₪₨₱₦₡₵₴₸₺₼₹]/g, "") // Currency symbols
    .replace(/[,.']/g, "") // Thousand separators and decimal points
    .replace(/\s+/g, "") // Spaces
    .trim();

  const number = parseFloat(cleaned);
  return isNaN(number) ? 0 : number;
}

export function cleanPercentageValue(value: unknown): number {
  if (typeof value === "number") return value;
  if (!value) return 0;

  const str = value.toString();
  // Remove percentage sign, spaces, and normalize decimal separators
  const cleaned = str
    .replace(/%/g, "")
    .replace(/,/g, ".")
    .replace(/\s+/g, "")
    .trim();
  const number = parseFloat(cleaned);

  if (isNaN(number)) return 0;

  // If value was originally a percentage string, convert to decimal
  // Check if original string had % sign (before cleaning)
  if (str.includes("%")) {
    return number / 100;
  }

  // If the number is > 1 and looks like it might be a percentage (e.g., "15" meaning 15%)
  // but we can't be sure, so we'll assume it's already a decimal if no % sign
  // This handles cases where percentages are stored as decimals (0.15) vs strings ("15%")
  return number;
}

export function cleanNumericValue(value: unknown): number {
  if (typeof value === "number") return value;
  if (!value) return 0;

  const str = value.toString();
  const cleaned = str.replace(/[,\s]/g, "").trim();
  const number = parseFloat(cleaned);
  return isNaN(number) ? 0 : number;
}

export function cleanTextValue(value: unknown): string {
  if (!value) return "";
  return value.toString().trim();
}

/**
 * Convert duration from time format (e.g., "4:02:02") to decimal hours (e.g., 4.03)
 * @param value The duration value (can be string like "4:02:02" or number)
 * @returns Decimal hours as a number
 */
export function cleanDurationValue(value: unknown): number {
  if (typeof value === "number") {
    // If already a number, assume it's already in decimal hours format
    return value;
  }
  if (!value) return 0;

  const str = value.toString().trim();
  if (!str) return 0;

  // Handle time format like "4:02:02" (hours:minutes:seconds)
  const timePattern = /^(\d+):(\d+):(\d+)$/;
  const match = str.match(timePattern);
  
  if (match) {
    const hours = parseInt(match[1], 10) || 0;
    const minutes = parseInt(match[2], 10) || 0;
    const seconds = parseInt(match[3], 10) || 0;
    
    // Convert to decimal hours: hours + (minutes/60) + (seconds/3600)
    const decimalHours = hours + (minutes / 60) + (seconds / 3600);
    // Round to 2 decimal places
    return Math.round(decimalHours * 100) / 100;
  }

  // If not in time format, try to parse as number
  const cleaned = str.replace(/[,\s]/g, "").trim();
  const number = parseFloat(cleaned);
  return isNaN(number) ? 0 : number;
}

// Backward compatibility
export const columnCleaningRules = tiktokColumnCleaningRules;
export const newRawDataHeaders = tiktokNewRawDataHeaders;

export const helperHeaders = [
  "Date serial",
  "Time text",
  "Time serial",
  "ISO Week",
  "Day name",
  "Direct GMV number",
  "Items sold",
  "Avg view duration seconds",
  "CTR decimal",
  "CTOR decimal",
  "Likes",
  "Comments",
  "Shares",
  "Start serial",
];

export const metricsHeaders = [
  "Day",
  "Date",
  "Week",
  "Jumlah Sesi",
  "Avg Revenue",
  "Sum",
  "Median CTR",
  "AVG CTR",
  "Median CTOR",
  "AVG CTOR",
  "Avg, view duration",
  "Sum Share",
  "Sum Comment",
  "Sum Likes",
  "sum Viewers",
  "Median GMV/1K shows",
];

export const summaryHeaders = ["Metric", "Value"];

export const summaryMetrics = [
  "Avg GMV/Sesi",
  "Avg GMV/Day",
  "Avg CTR",
  "Avg CTOR",
  "Avg Viewers",
  "Avg Like",
  "Avg Comment",
  "Avg Share",
];

export const trendHeaders = [
  "Month Index",
  "Month",
  "Avg CTR",
  "Avg CTOR",
  "Avg Viewers",
  "Avg Like",
  "Avg Comment",
  "Avg Share",
];

export function generateRawDataFormulas(
  row: number,
  format: ExcelFormat = ExcelFormat.TIKTOK_LIVESTREAM,
) {
  const formatDef = getFormatDefinition(format);
  if (!formatDef) {
    // Fallback to TikTok format if format not found
    const defaultFormat = getFormatDefinition(ExcelFormat.TIKTOK_LIVESTREAM);
    if (!defaultFormat) {
      throw new Error(`Format ${format} not found and no default format available`);
    }
    return defaultFormat.generateRawDataFormulas(row);
  }
  return formatDef.generateRawDataFormulas(row);
}

export function generateMetricsFormulas(
  row: number,
  startRow: number,
  lastRow: number,
  format: ExcelFormat = ExcelFormat.TIKTOK_LIVESTREAM,
) {
  const formatDef = getFormatDefinition(format);
  if (!formatDef) {
    // Fallback to TikTok format if format not found
    const defaultFormat = getFormatDefinition(ExcelFormat.TIKTOK_LIVESTREAM);
    if (!defaultFormat) {
      throw new Error(`Format ${format} not found and no default format available`);
    }
    return defaultFormat.generateMetricsFormulas(row, startRow, lastRow);
  }
  return formatDef.generateMetricsFormulas(row, startRow, lastRow);
}

export function cleanAndCopyData(
  inputWorksheet: Worksheet,
  cleanDataSheet: Worksheet,
  startRow: number,
  lastRow: number,
  format: ExcelFormat = ExcelFormat.TIKTOK_LIVESTREAM,
): void {
  const formatDef = getFormatDefinition(format);
  if (!formatDef) {
    // Fallback to TikTok format if format not found
    const defaultFormat = getFormatDefinition(ExcelFormat.TIKTOK_LIVESTREAM);
    if (!defaultFormat) {
      throw new Error(`Format ${format} not found and no default format available`);
    }
    return cleanAndCopyData(inputWorksheet, cleanDataSheet, startRow, lastRow, ExcelFormat.TIKTOK_LIVESTREAM);
  }

  const headers = formatDef.headers;
  const cleaningRules = formatDef.columnCleaningRules;

  // Copy headers first (row 1)
  headers.forEach((header, colIndex) => {
    cleanDataSheet.getCell(1, colIndex + 1).value = header;
  });

  // Copy and clean data starting from row 2 in clean sheet
  let cleanRow = 2;
  for (let rawRow = startRow; rawRow <= lastRow; rawRow++) {
    headers.forEach((_, colIndex) => {
      const rawCell = inputWorksheet.getCell(rawRow, colIndex + 1);
      const cleanCell = cleanDataSheet.getCell(cleanRow, colIndex + 1);

      const rawValue = rawCell.value;
      const cleaningRule = cleaningRules[colIndex];

      let cleanedValue: string | number;
      switch (cleaningRule) {
        case "currency":
          cleanedValue = cleanCurrencyValue(rawValue);
          break;
        case "percentage":
          cleanedValue = cleanPercentageValue(rawValue);
          break;
        case "numeric":
          cleanedValue = cleanNumericValue(rawValue);
          break;
        case "duration":
          cleanedValue = cleanDurationValue(rawValue);
          break;
        default:
          cleanedValue = cleanTextValue(rawValue);
          break;
      }

      cleanCell.value = cleanedValue;
    });
    cleanRow++;
  }
}

// Cell formatting functions
export function getAdvancedCurrencyFormat(currency: string = "IDR"): string {
  const formats: Record<string, string> = {
    USD: "$#,##0.00_);[Red]($#,##0.00)",
    EUR: "€#,##0.00_);[Red](€#,##0.00)",
    GBP: "£#,##0.00_);[Red](£#,##0.00)",
    JPY: "¥#,##0_);[Red](¥#,##0)",
    IDR: "Rp#,##0.00_);[Red](Rp#,##0.00)",
    SGD: "S$#,##0.00_);[Red](S$#,##0.00)",
    MYR: "RM#,##0.00_);[Red](RM#,##0.00)",
    THB: "฿#,##0.00_);[Red](฿#,##0.00)",
  };

  return formats[currency.toUpperCase()] || "Rp#,##0.00_);[Red](Rp#,##0.00)";
}

export function generateSummaryFormulas(
  startRow: number,
  lastRow: number,
  format: ExcelFormat = ExcelFormat.TIKTOK_LIVESTREAM,
): Record<string, string> {
  const formatDef = getFormatDefinition(format);
  if (!formatDef) {
    // Fallback to TikTok format if format not found
    const defaultFormat = getFormatDefinition(ExcelFormat.TIKTOK_LIVESTREAM);
    if (!defaultFormat) {
      throw new Error(`Format ${format} not found and no default format available`);
    }
    return defaultFormat.generateSummaryFormulas(startRow, lastRow);
  }
  return formatDef.generateSummaryFormulas(startRow, lastRow);
}

export function generateTrendFormulas(
  row: number,
  startRow: number,
  lastRow: number,
  format: ExcelFormat = ExcelFormat.TIKTOK_LIVESTREAM,
) {
  const formatDef = getFormatDefinition(format);
  if (!formatDef) {
    // Fallback to TikTok format if format not found
    const defaultFormat = getFormatDefinition(ExcelFormat.TIKTOK_LIVESTREAM);
    if (!defaultFormat) {
      throw new Error(`Format ${format} not found and no default format available`);
    }
    return defaultFormat.generateTrendFormulas(row, startRow, lastRow);
  }
  return formatDef.generateTrendFormulas(row, startRow, lastRow);
}

export function applyCTRCTORFormatting(
  worksheet: Worksheet,
  startRow: number,
  lastRow: number,
): void {
  // CTR is column V (22nd column - 0-based index 21)
  // CTOR is column W (23rd column - 0-based index 22)
  const ctrColumnIndex = 22; // Column V (CTR)
  const ctorColumnIndex = 23; // Column W (CTOR)

  for (let row = startRow; row <= lastRow; row++) {
    // Format CTR column
    const ctrCell = worksheet.getCell(row, ctrColumnIndex);
    ctrCell.numFmt = "0.00%";

    // Format CTOR column
    const ctorCell = worksheet.getCell(row, ctorColumnIndex);
    ctorCell.numFmt = "0.00%";
  }
}

/**
 * Apply percentage formatting to all percentage columns based on format definition.
 * This function reads the columnCleaningRules from the format definition and applies
 * percentage formatting to all columns marked as "percentage".
 * 
 * @param worksheet The worksheet to format
 * @param startRow The first data row (1-based)
 * @param lastRow The last data row (1-based)
 * @param format The Excel format to use for determining percentage columns
 */
export function applyPercentageFormatting(
  worksheet: Worksheet,
  startRow: number,
  lastRow: number,
  format: ExcelFormat = ExcelFormat.TIKTOK_LIVESTREAM,
): void {
  const formatDef = getFormatDefinition(format);
  if (!formatDef) {
    // Fallback to hardcoded columns if format not found
    applyCTRCTORFormatting(worksheet, startRow, lastRow);
    return;
  }

  // Find all percentage columns from the format definition
  const percentageColumns: number[] = [];
  Object.entries(formatDef.columnCleaningRules).forEach(([colIndexStr, rule]) => {
    if (rule === "percentage") {
      // Convert 0-based index to 1-based Excel column index
      const colIndex = parseInt(colIndexStr, 10);
      percentageColumns.push(colIndex + 1);
    }
  });

  // Apply percentage formatting to all percentage columns
  for (let row = startRow; row <= lastRow; row++) {
    percentageColumns.forEach((colIndex) => {
      const cell = worksheet.getCell(row, colIndex);
      cell.numFmt = "0.00%";
    });
  }
}

export function applyCurrencyFormatting(
  worksheet: Worksheet,
  startRow: number,
  lastRow: number,
  currency: string = "IDR",
  format: ExcelFormat = ExcelFormat.TIKTOK_LIVESTREAM,
): void {
  const formatDef = getFormatDefinition(format);
  const currencyFormat = getAdvancedCurrencyFormat(currency);

  let currencyColumns: number[] = [];

  if (formatDef) {
    // Find all currency columns from the format definition
    Object.entries(formatDef.columnCleaningRules).forEach(([colIndexStr, rule]) => {
      if (rule === "currency") {
        // Convert 0-based index to 1-based Excel column index
        const colIndex = parseInt(colIndexStr, 10);
        currencyColumns.push(colIndex + 1);
      }
    });
  } else {
    // Fallback to hardcoded columns for TikTok format if format not found
    // Currency columns based on TikTok format:
    // 3: Gross revenue (column D)
    // 4: Direct GMV (column E)
    // 7: Avg. price (column H)
    // 9: GMV/1K shows (column J)
    // 10: GMV/1K views (column K)
    currencyColumns = [4, 5, 8, 10, 11]; // 1-based column indices
  }

  for (let row = startRow; row <= lastRow; row++) {
    currencyColumns.forEach((colIndex) => {
      const cell = worksheet.getCell(row, colIndex);
      cell.numFmt = currencyFormat;
    });
  }
}

/**
 * Apply duration formatting to all duration columns based on format definition.
 * Duration columns are formatted as numbers with 2 decimal places (e.g., 4.03).
 * 
 * @param worksheet The worksheet to format
 * @param startRow The first data row (1-based)
 * @param lastRow The last data row (1-based)
 * @param format The Excel format to use for determining duration columns
 */
export function applyDurationFormatting(
  worksheet: Worksheet,
  startRow: number,
  lastRow: number,
  format: ExcelFormat = ExcelFormat.TIKTOK_LIVESTREAM,
): void {
  const formatDef = getFormatDefinition(format);
  if (!formatDef) {
    return;
  }

  // Find all duration columns from the format definition
  const durationColumns: number[] = [];
  Object.entries(formatDef.columnCleaningRules).forEach(([colIndexStr, rule]) => {
    if (rule === "duration") {
      // Convert 0-based index to 1-based Excel column index
      const colIndex = parseInt(colIndexStr, 10);
      durationColumns.push(colIndex + 1);
    }
  });

  // Apply number format with 2 decimal places to duration columns
  for (let row = startRow; row <= lastRow; row++) {
    durationColumns.forEach((colIndex) => {
      const cell = worksheet.getCell(row, colIndex);
      cell.numFmt = "0.00"; // Format as number with 2 decimal places
    });
  }
}

export function applyMetricsFormatting(
  worksheet: Worksheet,
  startRow: number,
  lastRow: number,
  currency: string = "IDR",
): void {
  if (lastRow < startRow) return;

  const currencyFormat = getAdvancedCurrencyFormat(currency);

  for (let row = startRow; row <= lastRow; row++) {
    // Revenue columns (E, F - Avg Revenue, Sum)
    ["E", "F"].forEach((col) => {
      const cell = worksheet.getCell(`${col}${row}`);
      cell.numFmt = currencyFormat;
    });

    // CTR columns (G, H - Median CTR, AVG CTR)
    ["G", "H"].forEach((col) => {
      const cell = worksheet.getCell(`${col}${row}`);
      cell.numFmt = "0.00%";
    });

    // CTOR columns (I, J - Median CTOR, AVG CTOR)
    ["I", "J"].forEach((col) => {
      const cell = worksheet.getCell(`${col}${row}`);
      cell.numFmt = "0.00%";
    });
  }
}

export function detectCurrencyFromData(
  worksheet: Worksheet,
  startRow: number,
  lastRow: number,
): string {
  // Check a few currency columns for currency symbols to detect currency type
  const currencyColumns = [4, 5]; // Gross revenue and Direct GMV columns

  for (let row = startRow; row <= Math.min(startRow + 5, lastRow); row++) {
    for (const colIndex of currencyColumns) {
      const cell = worksheet.getCell(row, colIndex);
      const value = cell.value?.toString() || "";

      // Check for currency symbols in the original data
      if (value.includes("$")) return "USD";
      if (value.includes("€")) return "EUR";
      if (value.includes("£")) return "GBP";
      if (value.includes("¥")) return "JPY";
      if (value.includes("S$")) return "SGD";
      if (value.includes("RM")) return "MYR";
      if (value.includes("฿")) return "THB";
      if (value.includes("Rp")) return "IDR";
    }
  }

  // Default to IDR (Indonesian Rupiah)
  return "IDR";
}

export function applySummaryFormatting(
  worksheet: Worksheet,
  currency: string = "IDR",
): void {
  const currencyFormat = getAdvancedCurrencyFormat(currency);

  // Format Avg GMV/Sesi (row 2) and Avg GMV/Day (row 3) as currency
  worksheet.getCell("B2").numFmt = currencyFormat; // Avg GMV/Sesi
  worksheet.getCell("B3").numFmt = currencyFormat; // Avg GMV/Day

  // Format Avg CTR (row 4) and Avg CTOR (row 5) as percentage
  worksheet.getCell("B4").numFmt = "0.00%"; // Avg CTR
  worksheet.getCell("B5").numFmt = "0.00%"; // Avg CTOR

  // Format Avg Viewers, Likes, Comments, Shares as numbers (rows 6-9)
  ["B6", "B7", "B8", "B9"].forEach((cell) => {
    worksheet.getCell(cell).numFmt = "#,##0";
  });

  // Style the headers
  ["A1", "B1"].forEach((cell) => {
    const headerCell = worksheet.getCell(cell);
    headerCell.font = { bold: true };
    headerCell.fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "FFE0E0E0" },
    };
  });

  // Style the metric names
  for (let i = 2; i <= 9; i++) {
    const metricCell = worksheet.getCell(`A${i}`);
    metricCell.font = { bold: true };
  }
}

export function applyTrendFormatting(
  worksheet: Worksheet,
  startRow: number,
  lastRow: number,
): void {
  if (lastRow < startRow) return;
  for (let row = startRow; row <= lastRow; row++) {
    ["C", "D"].forEach((col) => {
      const cell = worksheet.getCell(`${col}${row}`);
      cell.numFmt = "0.00%";
    });
  }
}
