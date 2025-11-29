import { Worksheet } from "exceljs";
import {
  ExcelFormat,
  ExcelFormatDefinition,
  DataRangeStrings,
  ColumnCleaningRules,
} from "../excel-helpers";

// Shopee Monthly format headers
export const shopeeHeaders = [
  "Periode Data",
  "User Id",
  "No.",
  "Nama Livestream",
  "Start Time",
  "Durasi",
  "Penonton Aktif",
  "Komentar",
  "Tambah ke Keranjang",
  "Durasi Rata-Rata Menonton",
  "Penonton",
  "Pesanan(Pesanan Dibuat)",
  "Pesanan(Pesanan Siap Dikirim)",
  "Produk Terjual(Pesanan Dibuat)",
  "Produk Terjual(Pesanan Siap Dikirim)",
  "Penjualan(Pesanan Dibuat)",
  "Penjualan(Pesanan Siap Dikirim)",
] as const;

// Shopee Monthly column cleaning rules (0-based index)
export const shopeeColumnCleaningRules: ColumnCleaningRules = {
  // Currency columns
  15: "currency", // Penjualan(Pesanan Dibuat)
  16: "currency", // Penjualan(Pesanan Siap Dikirim)

  // Numeric columns
  2: "numeric", // No.
  6: "numeric", // Penonton Aktif
  7: "numeric", // Komentar
  8: "numeric", // Tambah ke Keranjang
  10: "numeric", // Penonton
  11: "numeric", // Pesanan(Pesanan Dibuat)
  12: "numeric", // Pesanan(Pesanan Siap Dikirim)
  13: "numeric", // Produk Terjual(Pesanan Dibuat)
  14: "numeric", // Produk Terjual(Pesanan Siap Dikirim)

  // Duration columns (convert time format to decimal hours)
  5: "duration", // Durasi (convert "4:02:02" to 4.03)
  9: "duration", // Durasi Rata-Rata Menonton (convert time format to decimal hours)

  // Text columns (default for remaining)
  0: "text", // Periode Data
  1: "text", // User Id
  3: "text", // Nama Livestream
  4: "text", // Start Time
};

// Shopee Monthly new headers to append to raw data (column R)
export const shopeeNewRawDataHeaders = [
  "GMV/hour",
] as const;

const shopeeNewRawDataColumns = ["R"] as const;

/**
 * Shopee Monthly format definition
 */
export const shopeeFormat: ExcelFormatDefinition = {
  id: "SHOPEE_MONTHLY" as ExcelFormat,
  headers: shopeeHeaders,
  columnCleaningRules: shopeeColumnCleaningRules,
  newRawDataHeaders: shopeeNewRawDataHeaders,
  newRawDataColumns: shopeeNewRawDataColumns,
  detect: (worksheet: Worksheet): number | null => {
    // Check first 10 rows for format indicators
    for (let row = 1; row <= 10; row++) {
      const cellA = worksheet.getCell(`A${row}`);
      const cellD = worksheet.getCell(`D${row}`);

      const valueA = cellA.value?.toString().toLowerCase() || "";
      const valueD = cellD.value?.toString().toLowerCase() || "";

      // Check for Shopee Monthly format
      // Look for "Periode Data" in column A and "Nama Livestream" in column D
      if (
        (valueA.includes("periode data") || valueA.includes("data period")) &&
        (valueD.includes("nama livestream") || valueD.includes("livestream"))
      ) {
        return row + 1; // Data starts in the next row
      }
    }
    return null;
  },
  buildDataRanges: (startRow: number, lastRow: number): DataRangeStrings => {
    const seg = (c: string) => `\$${c}\$${startRow}:\$${c}\$${lastRow}`;
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
      gmv1kShowsR: `'clean data'!${seg("C")}`, // Placeholder (Shopee doesn't have GMV/1K shows)
    };
  },
  generateRawDataFormulas: (row: number): Record<string, string> => {
    return {
      // GMV/hour = Penjualan(Pesanan Siap Dikirim) / Durasi (in hours)
      // Q = column 17 (Penjualan(Pesanan Siap Dikirim), 0-based index 16)
      // F = column 6 (Durasi, 0-based index 5)
      // Durasi is now stored as decimal hours (e.g., 4.03), so we can directly divide
      R: `IFERROR(Q${row}/F${row},0)`, // GMV/hour
    };
  },
  generateMetricsFormulas: (
    row: number,
    startRow: number,
    lastRow: number,
  ): Record<string, string> => {
    const dataRanges = shopeeFormat.buildDataRanges(startRow, lastRow);
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
  },
  generateSummaryFormulas: (
    startRow: number,
    lastRow: number,
  ): Record<string, string> => {
    const dataRanges = shopeeFormat.buildDataRanges(startRow, lastRow);
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
  },
  generateTrendFormulas: (
    row: number,
    startRow: number,
    lastRow: number,
  ): Record<string, string> => {
    const dataRanges = shopeeFormat.buildDataRanges(startRow, lastRow);
    return {
      C: `ROUND(AVERAGEIFS(${dataRanges.ctrR},${dataRanges.montIndex},A${row}),4)`, // Avg CTR
      D: `ROUND(AVERAGEIFS(${dataRanges.ctorR},${dataRanges.montIndex},A${row}),4)`, // Avg CTOR
      E: `ROUND(AVERAGEIFS('clean data'!$I$${startRow}:$I$${lastRow},${dataRanges.montIndex},A${row}),0)`, // Avg Viewers
      F: `ROUND(AVERAGEIFS(${dataRanges.likesR},${dataRanges.montIndex},A${row}),0)`, // Avg Like
      G: `ROUND(AVERAGEIFS(${dataRanges.commentsR},${dataRanges.montIndex},A${row}),0)`, // Avg Comment
      H: `ROUND(AVERAGEIFS(${dataRanges.sharesR},${dataRanges.montIndex},A${row}),0)`, // Avg Share
    };
  },
};


