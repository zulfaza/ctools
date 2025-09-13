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
): DataRangeStrings {
  const seg = (c: string) => `\$${c}\$${startRow}:\$${c}\$${lastRow}`;
  return {
    startDateR: `'clean data'!${seg("X")}`,
    startTimeR: `'clean data'!${seg("Y")}`,
    endTimeR: `'clean data'!${seg("Z")}`,
    weekInYearR: `'clean data'!${seg("AA")}`,
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

// eslint-disable-next-line @typescript-eslint/no-explicit-any
export function detectDataStartRow(worksheet: any): number {
  for (let row = 1; row <= 10; row++) {
    // Check first 10 rows
    const cellA = worksheet.getCell(`A${row}`);
    const cellB = worksheet.getCell(`B${row}`);
    const cellC = worksheet.getCell(`C${row}`);

    const valueA = cellA.value?.toString().toLowerCase() || "";
    const valueB = cellB.value?.toString().toLowerCase() || "";
    const valueC = cellC.value?.toString().toLowerCase() || "";

    // Check if this row contains the expected headers
    if (
      valueA.includes("livestream") &&
      valueB.includes("start time") &&
      valueC.includes("duration")
    ) {
      return row + 1; // Data starts in the next row
    }
  }

  return 2; // Default to row 2 if headers not found
}

export function detectDataLength(
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  worksheet: any,
  startRow?: number,
): { startRow: number; lastRow: number } {
  const dataStartRow = startRow || detectDataStartRow(worksheet);
  let dataRows = 0;
  let row = dataStartRow;

  while (true) {
    const livestreamCell = worksheet.getCell(`A${row}`);
    const startTimeCell = worksheet.getCell(`B${row}`);

    // Check if both Livestream and Start time are empty
    if (
      (!livestreamCell.value || livestreamCell.value === "") &&
      (!startTimeCell.value || startTimeCell.value === "")
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

export const inputHeaders = [
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

// Define which columns need what type of cleaning
export const columnCleaningRules = {
  // Currency columns (0-based index)
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

  // Text columns (everything else defaults to text)
} as const;

// New headers to append to raw data (columns X, Y, Z, AA)
export const newRawDataHeaders = [
  "Start Date",
  "Start Time",
  "End Time",
  "Week in Year",
];

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

export function generateRawDataFormulas(row: number) {
  return {
    X: `TEXT(B${row},"yyyy-mm-dd")`, // Start Date
    Y: `TEXT(B${row},"hh:mm")`, // Start Time
    Z: `TEXT(B${row}+(C${row}/86400),"hh:mm")`, // End Time
    AA: `WEEKNUM(X${row},2)`, // Week in Year
  };
}

export function generateMetricsFormulas(
  row: number,
  startRow: number,
  lastRow: number,
) {
  const dataRanges = buildDataRanges(startRow, lastRow);

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
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  inputWorksheet: any,
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  cleanDataSheet: any,
  startRow: number,
  lastRow: number,
): void {
  // Copy headers first (row 1)
  inputHeaders.forEach((header, colIndex) => {
    cleanDataSheet.getCell(1, colIndex + 1).value = header;
  });

  // Copy and clean data starting from row 2 in clean sheet
  let cleanRow = 2;
  for (let rawRow = startRow; rawRow <= lastRow; rawRow++) {
    inputHeaders.forEach((_, colIndex) => {
      const rawCell = inputWorksheet.getCell(rawRow, colIndex + 1);
      const cleanCell = cleanDataSheet.getCell(cleanRow, colIndex + 1);

      const rawValue = rawCell.value;
      const cleaningRule =
        columnCleaningRules[colIndex as keyof typeof columnCleaningRules];

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

export function applyCTRCTORFormatting(
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  worksheet: any,
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
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  worksheet: any,
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
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  worksheet: any,
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
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  worksheet: any,
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
