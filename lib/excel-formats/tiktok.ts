import { Worksheet } from "exceljs";
import {
  ExcelFormat,
  ExcelFormatDefinition,
  DataRangeStrings,
  ColumnCleaningRules,
} from "../excel-helpers";

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
] as const;

// TikTok column cleaning rules (0-based index)
export const tiktokColumnCleaningRules: ColumnCleaningRules = {
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
};

// TikTok new headers to append to raw data (columns X, Y, Z, AA, AB, AC)
export const tiktokNewRawDataHeaders = [
  "Start Date",
  "Start Time",
  "End Time",
  "Week in Year",
  "Month",
  "GMV/Hour",
] as const;

const tiktokNewRawDataColumns = ["X", "Y", "Z", "AA", "AB", "AC"] as const;

/**
 * TikTok Livestream format definition
 */
export const tiktokFormat: ExcelFormatDefinition = {
  id: "TIKTOK_LIVESTREAM" as ExcelFormat,
  headers: tiktokHeaders,
  columnCleaningRules: tiktokColumnCleaningRules,
  newRawDataHeaders: tiktokNewRawDataHeaders,
  newRawDataColumns: tiktokNewRawDataColumns,
  detect: (worksheet: Worksheet): number | null => {
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
        (valueA.includes("livestream") ||
          valueA.includes("streaming langsung")) &&
        (valueB.includes("start time") || valueB.includes("waktu mulai")) &&
        (valueC.includes("duration") || valueC.includes("durasi"))
      ) {
        return row + 1; // Data starts in the next row
      }
    }
    return null;
  },
  buildDataRanges: (startRow: number, lastRow: number): DataRangeStrings => {
    const seg = (c: string) => `\$${c}\$${startRow}:\$${c}\$${lastRow}`;
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
      gmv1kShowsR: `'clean data'!${seg("J")}`,
    };
  },
  generateRawDataFormulas: (row: number): Record<string, string> => {
    return {
      X: `TEXT(B${row},"yyyy-mm-dd")`, // Start Date
      Y: `TEXT(B${row},"hh:mm")`, // Start Time
      Z: `TEXT(B${row}+(C${row}/86400),"hh:mm")`, // End Time
      AA: `WEEKNUM(X${row},2)`, // Week in Year
      AB: `MONTH(X${row})`, // Month
      AC: `D${row}/(C${row}/3600)`, // GMV/Hour
    };
  },
  generateMetricsFormulas: (
    row: number,
    startRow: number,
    lastRow: number,
  ): Record<string, string> => {
    const dataRanges = tiktokFormat.buildDataRanges(startRow, lastRow);
    return {
      A: `TEXT(B${row},"dddd")`, // Day
      B: `IFERROR(SORT(UNIQUE(FILTER(${dataRanges.startDateR},${dataRanges.startDateR}<>""))),"")`, // Date (only for B2)
      C: `XLOOKUP(B${row},${dataRanges.startDateR},${dataRanges.weekInYearR})`, // Week
      D: `COUNTIF(${dataRanges.startDateR},B${row})`, // Jumlah Sesi
      E: `ROUND(AVERAGEIFS(${dataRanges.grossR},${dataRanges.startDateR},B${row}),0)`, // Avg Revenue
      F: `SUMIFS(${dataRanges.grossR},${dataRanges.startDateR},B${row})`, // Sum
      G: `MEDIAN(FILTER(${dataRanges.ctrR},${dataRanges.startDateR}=B${row}))`, // Median CTR
      H: `ROUND(AVERAGEIFS(${dataRanges.ctrR},${dataRanges.startDateR},B${row}),4)`, // AVG CTR
      I: `MEDIAN(FILTER(${dataRanges.ctorR},${dataRanges.startDateR}=B${row}))`, // Median CTOR
      J: `ROUND(AVERAGEIFS(${dataRanges.ctorR},${dataRanges.startDateR},B${row}),4)`, // AVG CTOR
      K: `ROUND(AVERAGEIFS(${dataRanges.avgViewR},${dataRanges.startDateR},B${row}),0)`, // Avg view duration
      L: `SUMIFS(${dataRanges.sharesR},${dataRanges.startDateR},B${row})`, // Sum Share
      M: `SUMIFS(${dataRanges.commentsR},${dataRanges.startDateR},B${row})`, // Sum Comment
      N: `SUMIFS(${dataRanges.likesR},${dataRanges.startDateR},B${row})`, // Sum Likes
      O: `SUMIFS('clean data'!$M$${startRow}:$M$${lastRow},${dataRanges.startDateR},B${row})`, // sum Viewers
      P: `MEDIAN(FILTER(${dataRanges.gmv1kShowsR},${dataRanges.startDateR}=B${row}))`, // Median GMV/1K shows
    };
  },
  generateSummaryFormulas: (
    startRow: number,
    lastRow: number,
  ): Record<string, string> => {
    const dataRanges = tiktokFormat.buildDataRanges(startRow, lastRow);
    return {
      "Avg GMV/Sesi": `ROUND(AVERAGE(${dataRanges.directR}),0)`,
      "Avg GMV/Day": `ROUND(AVERAGE(metrics!F:F),0)`, // Average of Sum column from metrics (now column F)
      "Avg CTR": `ROUND(AVERAGE(${dataRanges.ctrR}),4)`,
      "Avg CTOR": `ROUND(AVERAGE(${dataRanges.ctorR}),4)`,
      "Avg Viewers": `ROUND(AVERAGE('clean data'!$M$${startRow}:$M$${lastRow}),0)`, // Viewers column
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
    const dataRanges = tiktokFormat.buildDataRanges(startRow, lastRow);
    return {
      C: `ROUND(AVERAGEIFS(${dataRanges.ctrR},${dataRanges.montIndex},A${row}),4)`, // Avg CTR
      D: `ROUND(AVERAGEIFS(${dataRanges.ctorR},${dataRanges.montIndex},A${row}),4)`, // Avg CTOR
      E: `ROUND(AVERAGEIFS(${dataRanges.avgViewR},${dataRanges.montIndex},A${row}),0)`, // Avg Viewers
      F: `ROUND(AVERAGEIFS(${dataRanges.likesR},${dataRanges.montIndex},A${row}),0)`, // Avg Like
      G: `ROUND(AVERAGEIFS(${dataRanges.commentsR},${dataRanges.montIndex},A${row}),0)`, // Avg Comment
      H: `ROUND(AVERAGEIFS(${dataRanges.sharesR},${dataRanges.montIndex},A${row}),0)`, // Avg Share
    };
  },
};


