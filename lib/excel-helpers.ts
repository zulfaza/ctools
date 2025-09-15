import { Worksheet } from "exceljs";

export enum ExcelFormat {
  TIKTOK_LIVESTREAM = "TIKTOK_LIVESTREAM",
  SHOPEE_DAILY = "SHOPEE_DAILY",
  UNSUPPORTED = "UNSUPPORTED",
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
}

export function buildDataRanges(
  startRow: number,
  lastRow: number,
  format: ExcelFormat = ExcelFormat.TIKTOK_LIVESTREAM,
): DataRangeStrings {
  const seg = (c: string) => `\$${c}\$${startRow}:\$${c}\$${lastRow}`;

  if (format === ExcelFormat.SHOPEE_DAILY) {
    return {
      startDateR: `'clean data'!${seg("AE")}`, // Date from Data Period
      startTimeR: `'clean data'!${seg("AF")}`, // Time (placeholder)
      endTimeR: `'clean data'!${seg("AG")}`, // End time (placeholder)
      weekInYearR: `'clean data'!${seg("AH")}`, // Week number
      montIndex: `'clean data'!${seg("AI")}`, // Month
      grossR: `'clean data'!${seg("C")}`, // Sales(Placed Order)
      directR: `'clean data'!${seg("D")}`, // Sales(Confirmed Order)
      itemsR: `'clean data'!${seg("G")}`, // Total Items Sold(Placed Order)
      avgViewR: `'clean data'!${seg("K")}`, // Avg. Viewing Duration
      likesR: `'clean data'!${seg("X")}`, // Total Likes
      commentsR: `'clean data'!${seg("Z")}`, // Total Comments
      sharesR: `'clean data'!${seg("Y")}`, // Total Shares
      ctrR: `'clean data'!${seg("O")}`, // CTR
      ctorR: `'clean data'!${seg("P")}`, // Click to Order Rate(Placed Order)
    };
  }

  // TikTok format (default)
  return {
    startDateR: `'clean data'!${seg("X")}`,
    startTimeR: `'clean data'!${seg("Y")}`,
    endTimeR: `'clean data'!${seg("Z")}`,
    weekInYearR: `'clean data'!${seg("AA")}`,
    montIndex: `'clean data'!${seg("AB")}`,
    grossR: `'clean data'!${seg("D")}`,
    directR: `'clean data'!${seg("E")}`,
    itemsR: `'clean data'!${seg("F")}`,
    avgViewR: `'clean data'!${seg("P")}`,
    likesR: `'clean data'!${seg("Q")}`,
    commentsR: `'clean data'!${seg("R")}`,
    sharesR: `'clean data'!${seg("S")}`,
    ctrR: `'clean data'!${seg("V")}`,
    ctorR: `'clean data'!${seg("W")}`,
  };
}

export function detectExcelFormat(worksheet: Worksheet): ExcelFormat {
  // Check first 10 rows for format indicators
  for (let row = 1; row <= 10; row++) {
    const cellA = worksheet.getCell(`A${row}`);
    const cellB = worksheet.getCell(`B${row}`);
    const cellC = worksheet.getCell(`C${row}`);

    const valueA = cellA.value?.toString().toLowerCase() || "";
    const valueB = cellB.value?.toString().toLowerCase() || "";
    const valueC = cellC.value?.toString().toLowerCase() || "";

    // Check for TikTok Livestream format
    if (
      valueA.includes("livestream") &&
      valueB.includes("start time") &&
      valueC.includes("duration")
    ) {
      return ExcelFormat.TIKTOK_LIVESTREAM;
    }

    // Check for Shopee Daily format
    if (
      valueA.includes("data period") &&
      valueB.includes("user id") &&
      (valueC.includes("sales") || valueC.includes("placed order"))
    ) {
      return ExcelFormat.SHOPEE_DAILY;
    }
  }

  return ExcelFormat.UNSUPPORTED;
}

export function detectDataStartRow(
  worksheet: Worksheet,
  format?: ExcelFormat,
): number {
  const detectedFormat = format || detectExcelFormat(worksheet);

  for (let row = 1; row <= 10; row++) {
    const cellA = worksheet.getCell(`A${row}`);
    const cellB = worksheet.getCell(`B${row}`);
    const cellC = worksheet.getCell(`C${row}`);

    const valueA = cellA.value?.toString().toLowerCase() || "";
    const valueB = cellB.value?.toString().toLowerCase() || "";
    const valueC = cellC.value?.toString().toLowerCase() || "";

    if (detectedFormat === ExcelFormat.TIKTOK_LIVESTREAM) {
      // Check for TikTok headers
      if (
        valueA.includes("livestream") &&
        valueB.includes("start time") &&
        valueC.includes("duration")
      ) {
        return row + 1; // Data starts in the next row
      }
    } else if (detectedFormat === ExcelFormat.SHOPEE_DAILY) {
      // Check for Shopee headers
      if (
        valueA.includes("data period") &&
        valueB.includes("user id") &&
        (valueC.includes("sales") || valueC.includes("placed order"))
      ) {
        return row + 1; // Data starts in the next row
      }
    }
  }

  return 2; // Default to row 2 if headers not found
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

// TikTok Livestream format headers
export const tiktokHeaders = [
  "Livestream",
  "Start time",
  "Duration",
  "Gross revenue",
  "Direct GMV",
  "Items sold",
  "Customers",
  "Avg, price",
  "Orders paid for",
  "GMV/1K shows",
  "GMV/1K views",
  "Views",
  "Viewers",
  "Peak viewers",
  "New followers",
  "Avg, view duration",
  "Likes",
  "Comments",
  "Shares",
  "Product impressions",
  "Product clicks",
  "CTR",
  "CTOR",
];

// Shopee Daily format headers
export const shopeeHeaders = [
  "Data Period",
  "User Id",
  "Sales(Placed Order)",
  "Sales(Confirmed Order)",
  "Orders(Placed Order)",
  "Orders(Confirmed Order)",
  "Total Items Sold(Placed Order)",
  "Total Items Sold(Confirmed Order)",
  "Total Viewers",
  "Engaged Viewers",
  "Avg. Viewing Duration",
  "Buyers(Placed Order)",
  "Buyers(Confirmed Order)",
  "Total ATC",
  "CTR",
  "Click to Order Rate(Placed Order)",
  "Click to Order Rate(Confirmed Order)",
  "ABS(Placed Order)",
  "ABS(Confirmed Order)",
  "GPM(Placed Order)",
  "GPM(Confirmed Order)",
  "Total Views",
  "PCU",
  "Total Likes",
  "Total Shares",
  "Total Comments",
  "Live New Followers",
  "Shop Voucher Claimed",
  "Special Live Voucher Claimed",
  "Coins Claimed",
];

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
  const cleaned = str.replace(/%/g, "").replace(/,/g, ".").trim();
  const number = parseFloat(cleaned);

  if (isNaN(number)) return 0;

  // If value was originally a percentage string, convert to decimal
  if (str.includes("%")) {
    return number / 100;
  }

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

// TikTok column cleaning rules (0-based index)
export const tiktokColumnCleaningRules = {
  // Currency columns
  3: "currency", // Gross revenue
  4: "currency", // Direct GMV
  7: "currency", // Avg. price
  9: "currency", // GMV/1K shows
  10: "currency", // GMV/1K views

  // Numeric columns
  2: "numeric", // Duration
  5: "numeric", // Items sold
  6: "numeric", // Customers
  8: "numeric", // Orders paid for
  11: "numeric", // Views
  12: "numeric", // Viewers
  13: "numeric", // Peak viewers
  14: "numeric", // New followers
  15: "numeric", // Avg. view duration
  16: "numeric", // Likes
  17: "numeric", // Comments
  18: "numeric", // Shares
  19: "numeric", // Product impressions
  20: "numeric", // Product clicks

  // Percentage columns
  21: "percentage", // CTR
  22: "percentage", // CTOR
} as const;

// Shopee column cleaning rules (0-based index)
export const shopeeColumnCleaningRules = {
  // Currency columns
  2: "currency", // Sales(Placed Order)
  3: "currency", // Sales(Confirmed Order)
  17: "currency", // ABS(Placed Order)
  18: "currency", // ABS(Confirmed Order)
  19: "currency", // GPM(Placed Order)
  20: "currency", // GPM(Confirmed Order)

  // Numeric columns
  4: "numeric", // Orders(Placed Order)
  5: "numeric", // Orders(Confirmed Order)
  6: "numeric", // Total Items Sold(Placed Order)
  7: "numeric", // Total Items Sold(Confirmed Order)
  8: "numeric", // Total Viewers
  9: "numeric", // Engaged Viewers
  10: "numeric", // Avg. Viewing Duration
  11: "numeric", // Buyers(Placed Order)
  12: "numeric", // Buyers(Confirmed Order)
  13: "numeric", // Total ATC
  21: "numeric", // Total Views
  22: "numeric", // PCU
  23: "numeric", // Total Likes
  24: "numeric", // Total Shares
  25: "numeric", // Total Comments
  26: "numeric", // Live New Followers
  29: "numeric", // Coins Claimed

  // Percentage columns
  14: "percentage", // CTR
  15: "percentage", // Click to Order Rate(Placed Order)
  16: "percentage", // Click to Order Rate(Confirmed Order)

  // Text columns
  1: "text", // User Id
  27: "text", // Shop Voucher Claimed
  28: "text", // Special Live Voucher Claimed
} as const;

// Backward compatibility
export const columnCleaningRules = tiktokColumnCleaningRules;

// TikTok new headers to append to raw data (columns X, Y, Z, AA, AB)
export const tiktokNewRawDataHeaders = [
  "Start Date",
  "Start Time",
  "End Time",
  "Week in Year",
  "Month",
];

// Shopee new headers to append to raw data (columns AE, AF, AG, AH, AI)
export const shopeeNewRawDataHeaders = [
  "Date",
  "Time", // Placeholder
  "End Time", // Placeholder
  "Week in Year",
  "Month",
];

// Backward compatibility
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
  "Max Revenue",
  "Min Revenue",
  "Avg Revenue",
  "Sum",
  "Prime time based on Max Revenue",
  "Max Item Sold (pcs)",
  "Min Item Sold (pcs)",
  "Avg Item Sold (pcs)",
  "Avg Duration (Seconds)",
  "Prime time based on Max by CTR",
  "Max CTR",
  "Min CTR",
  "AVG CTR",
  "Prime time based on Max by CTOR",
  "Max CTOR",
  "Min CT0R",
  "AVG CTOR",
  "Avg, view duration",
  "Max Likes",
  "Min Likes",
  "AVG Likes",
  "Max Comment",
  "Min Comment",
  "AVG Comment",
  "Max Share",
  "Min Share",
  "AVG Share",
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
  if (format === ExcelFormat.SHOPEE_DAILY) {
    return {
      AE: `TEXT(DATEVALUE(A${row}),"yyyy-mm-dd")`, // Date from Data Period
      AF: `""`, // Time placeholder
      AG: `""`, // End Time placeholder
      AH: `WEEKNUM(AE${row},2)`, // Week in Year
      AI: `MONTH(AE${row})`, // Month
    };
  }

  // TikTok format (default)
  return {
    X: `TEXT(B${row},"yyyy-mm-dd")`, // Start Date
    Y: `TEXT(B${row},"hh:mm")`, // Start Time
    Z: `TEXT(B${row}+(C${row}/86400),"hh:mm")`, // End Time
    AA: `WEEKNUM(X${row},2)`, // Week in Year
    AB: `MONTH(X${row})`, // Month
  };
}

export function generateMetricsFormulas(
  row: number,
  startRow: number,
  lastRow: number,
  format: ExcelFormat = ExcelFormat.TIKTOK_LIVESTREAM,
) {
  const dataRanges = buildDataRanges(startRow, lastRow, format);

  if (format === ExcelFormat.SHOPEE_DAILY) {
    return {
      A: `TEXT(B${row},"dddd")`, // Day
      B: `IFERROR(SORT(UNIQUE(FILTER(${dataRanges.startDateR},${dataRanges.startDateR}<>""))),"")`, // Date
      C: `XLOOKUP(B${row},${dataRanges.startDateR},${dataRanges.weekInYearR})`, // Week
      D: `COUNTIF(${dataRanges.startDateR},B${row})`, // Sessions
      E: `MAXIFS(${dataRanges.grossR},${dataRanges.startDateR},B${row})`, // Max Sales(Placed)
      F: `MINIFS(${dataRanges.grossR},${dataRanges.startDateR},B${row})`, // Min Sales(Placed)
      G: `ROUND(AVERAGEIFS(${dataRanges.grossR},${dataRanges.startDateR},B${row}),0)`, // Avg Sales(Placed)
      H: `SUMIFS(${dataRanges.grossR},${dataRanges.startDateR},B${row})`, // Sum Sales(Placed)
      I: `""`, // Prime time placeholder
      J: `MAXIFS(${dataRanges.itemsR},${dataRanges.startDateR},B${row})`, // Max Items Sold
      K: `MINIFS(${dataRanges.itemsR},${dataRanges.startDateR},B${row})`, // Min Items Sold
      L: `ROUND(AVERAGEIFS(${dataRanges.itemsR},${dataRanges.startDateR},B${row}),0)`, // Avg Items Sold
      M: `ROUND(AVERAGEIFS(${dataRanges.avgViewR},${dataRanges.startDateR},B${row}),0)`, // Avg View Duration
      N: `""`, // Prime time by CTR placeholder
      O: `MAXIFS(${dataRanges.ctrR},${dataRanges.startDateR},B${row})`, // Max CTR
      P: `MINIFS(${dataRanges.ctrR},${dataRanges.startDateR},B${row})`, // Min CTR
      Q: `ROUND(AVERAGEIFS(${dataRanges.ctrR},${dataRanges.startDateR},B${row}),4)`, // AVG CTR
      R: `""`, // Prime time by CTOR placeholder
      S: `MAXIFS(${dataRanges.ctorR},${dataRanges.startDateR},B${row})`, // Max CTOR
      T: `MINIFS(${dataRanges.ctorR},${dataRanges.startDateR},B${row})`, // Min CTOR
      U: `ROUND(AVERAGEIFS(${dataRanges.ctorR},${dataRanges.startDateR},B${row}),4)`, // AVG CTOR
      V: `ROUND(AVERAGEIFS(${dataRanges.avgViewR},${dataRanges.startDateR},B${row}),0)`, // Avg view duration
      W: `MAXIFS(${dataRanges.likesR},${dataRanges.startDateR},B${row})`, // Max Likes
      X: `MINIFS(${dataRanges.likesR},${dataRanges.startDateR},B${row})`, // Min Likes
      Y: `ROUND(AVERAGEIFS(${dataRanges.likesR},${dataRanges.startDateR},B${row}),0)`, // AVG Likes
      Z: `MAXIFS(${dataRanges.commentsR},${dataRanges.startDateR},B${row})`, // Max Comment
      AA: `MINIFS(${dataRanges.commentsR},${dataRanges.startDateR},B${row})`, // Min Comment
      AB: `ROUND(AVERAGEIFS(${dataRanges.commentsR},${dataRanges.startDateR},B${row}),0)`, // AVG Comment
      AC: `MAXIFS(${dataRanges.sharesR},${dataRanges.startDateR},B${row})`, // Max Share
      AD: `MINIFS(${dataRanges.sharesR},${dataRanges.startDateR},B${row})`, // Min Share
      AE: `ROUND(AVERAGEIFS(${dataRanges.sharesR},${dataRanges.startDateR},B${row}),0)`, // AVG Share
    };
  }

  // TikTok format (default)
  return {
    A: `TEXT(B${row},"dddd")`, // Day
    B: `IFERROR(SORT(UNIQUE(FILTER(${dataRanges.startDateR},${dataRanges.startDateR}<>""))),"")`, // Date (only for B2)
    C: `XLOOKUP(B${row},${dataRanges.startDateR},${dataRanges.weekInYearR})`, // Week
    D: `COUNTIF(${dataRanges.startDateR},B${row})`, // Jumlah Sesi
    E: `MAXIFS(${dataRanges.grossR},${dataRanges.startDateR},B${row})`, // Max Revenue
    F: `MINIFS(${dataRanges.grossR},${dataRanges.startDateR},B${row})`, // Min Revenue
    G: `ROUND(AVERAGEIFS(${dataRanges.grossR},${dataRanges.startDateR},B${row}),0)`, // Avg Revenue
    H: `SUMIFS(${dataRanges.grossR},${dataRanges.startDateR},B${row})`, // Sum
    I: `TEXTJOIN(", ",TRUE, FILTER(${dataRanges.startTimeR}, (${dataRanges.startDateR}=B${row})*(${dataRanges.grossR}=E${row})))`, // Prime time by Max Revenue
    J: `MAXIFS(${dataRanges.itemsR},${dataRanges.startDateR},B${row})`, // Max Item Sold
    K: `MINIFS(${dataRanges.itemsR},${dataRanges.startDateR},B${row})`, // Min Item Sold
    L: `ROUND(AVERAGEIFS(${dataRanges.itemsR},${dataRanges.startDateR},B${row}),0)`, // Avg Item Sold
    M: `ROUND(AVERAGEIFS(${dataRanges.avgViewR},${dataRanges.startDateR},B${row}),0)`, // Avg Duration
    N: `TEXTJOIN(", ",TRUE, FILTER(${dataRanges.startTimeR}, (${dataRanges.startDateR}=B${row})*(${dataRanges.ctrR}=O${row})))`, // Prime time by Max CTR
    O: `MAXIFS(${dataRanges.ctrR},${dataRanges.startDateR},B${row})`, // Max CTR
    P: `MINIFS(${dataRanges.ctrR},${dataRanges.startDateR},B${row})`, // Min CTR
    Q: `ROUND(AVERAGEIFS(${dataRanges.ctrR},${dataRanges.startDateR},B${row}),4)`, // AVG CTR
    R: `TEXTJOIN(", ",TRUE, FILTER(${dataRanges.startTimeR}, (${dataRanges.startDateR}=B${row})*(${dataRanges.ctorR}=S${row})))`, // Prime time by Max CTOR
    S: `MAXIFS(${dataRanges.ctorR},${dataRanges.startDateR},B${row})`, // Max CTOR
    T: `MINIFS(${dataRanges.ctorR},${dataRanges.startDateR},B${row})`, // Min CTOR
    U: `ROUND(AVERAGEIFS(${dataRanges.ctorR},${dataRanges.startDateR},B${row}),4)`, // AVG CTOR
    V: `ROUND(AVERAGEIFS(${dataRanges.avgViewR},${dataRanges.startDateR},B${row}),0)`, // Avg view duration
    W: `MAXIFS(${dataRanges.likesR},${dataRanges.startDateR},B${row})`, // Max Likes
    X: `MINIFS(${dataRanges.likesR},${dataRanges.startDateR},B${row})`, // Min Likes
    Y: `ROUND(AVERAGEIFS(${dataRanges.likesR},${dataRanges.startDateR},B${row}),0)`, // AVG Likes
    Z: `MAXIFS(${dataRanges.commentsR},${dataRanges.startDateR},B${row})`, // Max Comment
    AA: `MINIFS(${dataRanges.commentsR},${dataRanges.startDateR},B${row})`, // Min Comment
    AB: `ROUND(AVERAGEIFS(${dataRanges.commentsR},${dataRanges.startDateR},B${row}),0)`, // AVG Comment
    AC: `MAXIFS(${dataRanges.sharesR},${dataRanges.startDateR},B${row})`, // Max Share
    AD: `MINIFS(${dataRanges.sharesR},${dataRanges.startDateR},B${row})`, // Min Share
    AE: `ROUND(AVERAGEIFS(${dataRanges.sharesR},${dataRanges.startDateR},B${row}),0)`, // AVG Share
  };
}

export function cleanAndCopyData(
  inputWorksheet: Worksheet,
  cleanDataSheet: Worksheet,
  startRow: number,
  lastRow: number,
  format: ExcelFormat = ExcelFormat.TIKTOK_LIVESTREAM,
): void {
  const headers =
    format === ExcelFormat.SHOPEE_DAILY ? shopeeHeaders : tiktokHeaders;
  const cleaningRules =
    format === ExcelFormat.SHOPEE_DAILY
      ? shopeeColumnCleaningRules
      : tiktokColumnCleaningRules;

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
      const cleaningRule =
        cleaningRules[colIndex as keyof typeof cleaningRules];

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
  const dataRanges = buildDataRanges(startRow, lastRow, format);

  if (format === ExcelFormat.SHOPEE_DAILY) {
    return {
      "Avg Sales/Session": `ROUND(AVERAGE(${dataRanges.grossR}),0)`,
      "Avg Sales/Day": `ROUND(AVERAGE(metrics!H:H),0)`,
      "Avg CTR": `ROUND(AVERAGE(${dataRanges.ctrR}),4)`,
      "Avg CTOR": `ROUND(AVERAGE(${dataRanges.ctorR}),4)`,
      "Avg Viewers": `ROUND(AVERAGE('clean data'!$I$${startRow}:$I$${lastRow}),0)`, // Total Viewers column
      "Avg Like": `ROUND(AVERAGE(${dataRanges.likesR}),0)`,
      "Avg Comment": `ROUND(AVERAGE(${dataRanges.commentsR}),0)`,
      "Avg Share": `ROUND(AVERAGE(${dataRanges.sharesR}),0)`,
    };
  }

  // TikTok format (default)
  return {
    "Avg GMV/Sesi": `ROUND(AVERAGE(${dataRanges.directR}),0)`,
    "Avg GMV/Day": `ROUND(AVERAGE(metrics!H:H),0)`, // Average of Sum column from metrics
    "Avg CTR": `ROUND(AVERAGE(${dataRanges.ctrR}),4)`,
    "Avg CTOR": `ROUND(AVERAGE(${dataRanges.ctorR}),4)`,
    "Avg Viewers": `ROUND(AVERAGE('clean data'!$M$${startRow}:$M$${lastRow}),0)`, // Viewers column
    "Avg Like": `ROUND(AVERAGE(${dataRanges.likesR}),0)`,
    "Avg Comment": `ROUND(AVERAGE(${dataRanges.commentsR}),0)`,
    "Avg Share": `ROUND(AVERAGE(${dataRanges.sharesR}),0)`,
  };
}

export function generateTrendFormulas(
  row: number,
  startRow: number,
  lastRow: number,
  format: ExcelFormat = ExcelFormat.TIKTOK_LIVESTREAM,
) {
  const dataRanges = buildDataRanges(startRow, lastRow, format);

  if (format === ExcelFormat.SHOPEE_DAILY) {
    return {
      C: `ROUND(AVERAGEIFS(${dataRanges.ctrR},${dataRanges.montIndex},A${row}),4)`, // Avg CTR
      D: `ROUND(AVERAGEIFS(${dataRanges.ctorR},${dataRanges.montIndex},A${row}),4)`, // Avg CTOR
      E: `ROUND(AVERAGEIFS('clean data'!$I$${startRow}:$I$${lastRow},${dataRanges.montIndex},A${row}),0)`, // Avg Viewers
      F: `ROUND(AVERAGEIFS(${dataRanges.likesR},${dataRanges.montIndex},A${row}),0)`, // Avg Like
      G: `ROUND(AVERAGEIFS(${dataRanges.commentsR},${dataRanges.montIndex},A${row}),0)`, // Avg Comment
      H: `ROUND(AVERAGEIFS(${dataRanges.sharesR},${dataRanges.montIndex},A${row}),0)`, // Avg Share
    };
  }

  // TikTok format (default)
  return {
    C: `ROUND(AVERAGEIFS(${dataRanges.ctrR},${dataRanges.montIndex},A${row}),4)`, // Avg CTR
    D: `ROUND(AVERAGEIFS(${dataRanges.ctorR},${dataRanges.montIndex},A${row}),4)`, // Avg CTOR
    E: `ROUND(AVERAGEIFS(${dataRanges.avgViewR},${dataRanges.montIndex},A${row}),0)`, // Avg Viewers
    F: `ROUND(AVERAGEIFS(${dataRanges.likesR},${dataRanges.montIndex},A${row}),0)`, // Avg Like
    G: `ROUND(AVERAGEIFS(${dataRanges.commentsR},${dataRanges.montIndex},A${row}),0)`, // Avg Comment
    H: `ROUND(AVERAGEIFS(${dataRanges.sharesR},${dataRanges.montIndex},A${row}),0)`, // Avg Share
  };
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

export function applyCurrencyFormatting(
  worksheet: Worksheet,
  startRow: number,
  lastRow: number,
  currency: string = "IDR",
): void {
  // Currency columns based on columnCleaningRules:
  // 3: Gross revenue (column D)
  // 4: Direct GMV (column E)
  // 7: Avg. price (column H)
  // 9: GMV/1K shows (column J)
  // 10: GMV/1K views (column K)
  const currencyColumns = [4, 5, 8, 10, 11]; // 1-based column indices

  const currencyFormat = getAdvancedCurrencyFormat(currency);

  for (let row = startRow; row <= lastRow; row++) {
    currencyColumns.forEach((colIndex) => {
      const cell = worksheet.getCell(row, colIndex);
      cell.numFmt = currencyFormat;
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
    // Revenue columns (E, F, G, H - Max Revenue, Min Revenue, Avg Revenue, Sum)
    ["E", "F", "G", "H"].forEach((col) => {
      const cell = worksheet.getCell(`${col}${row}`);
      cell.numFmt = currencyFormat;
    });

    // CTR columns (O, P, Q - Max CTR, Min CTR, AVG CTR)
    ["O", "P", "Q"].forEach((col) => {
      const cell = worksheet.getCell(`${col}${row}`);
      cell.numFmt = "0.00%";
    });

    // CTOR columns (S, T, U - Max CTOR, Min CTOR, AVG CTOR)
    ["S", "T", "U"].forEach((col) => {
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
