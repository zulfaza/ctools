import { Worksheet } from "exceljs";
import {
  ExcelFormat,
  ExcelFormatDefinition,
  DataRangeStrings,
  ColumnCleaningRules,
} from "../excel-helpers";

// Canonical Shopee format headers (supports both Monthly livestream and Daily CSV)
export const shopeeHeaders = [
  "Periode Data", // 0: Date
  "User Id", // 1: User ID
  "Penjualan(Pesanan Dibuat)", // 2: GMV Placed
  "Penjualan(Pesanan Siap Dikirim)", // 3: GMV Shipped
  "Pesanan(Pesanan Dibuat)", // 4: Orders Placed
  "Pesanan(Pesanan Siap Dikirim)", // 5: Orders Shipped
  "Produk Terjual(Pesanan Dibuat)", // 6: Items Sold Placed
  "Produk Terjual(Pesanan Siap Dikirim)", // 7: Items Sold Shipped
  "Penonton", // 8: Viewers
  "Penonton Aktif", // 9: Active Viewers
  "Rata-rata durasi ditonton", // 10: Avg Viewing Duration
  "Persentase Klik", // 11: CTR
  "Pesanan per Klik(Pesanan Dibuat)", // 12: CTOR Placed
  "Pesanan per Klik(Pesanan Siap Dikirim)", // 13: CTOR Shipped
  "Suka", // 14: Likes
  "Share", // 15: Shares
  "Komentar", // 16: Comments
] as const;

// Shopee canonical column cleaning rules (0-based index)
export const shopeeColumnCleaningRules: ColumnCleaningRules = {
  // Currency columns
  2: "currency", // Penjualan(Pesanan Dibuat) - GMV Placed
  3: "currency", // Penjualan(Pesanan Siap Dikirim) - GMV Shipped

  // Numeric columns
  4: "numeric", // Pesanan(Pesanan Dibuat) - Orders Placed
  5: "numeric", // Pesanan(Pesanan Siap Dikirim) - Orders Shipped
  6: "numeric", // Produk Terjual(Pesanan Dibuat) - Items Sold Placed
  7: "numeric", // Produk Terjual(Pesanan Siap Dikirim) - Items Sold Shipped
  8: "numeric", // Penonton - Viewers
  9: "numeric", // Penonton Aktif - Active Viewers
  14: "numeric", // Suka - Likes
  15: "numeric", // Share - Shares
  16: "numeric", // Komentar - Comments

  // Text columns for duration (keep as-is, don't convert)
  10: "text", // Rata-rata durasi ditonton - Avg Viewing Duration (keep original format)

  // Percentage columns
  11: "percentage", // Persentase Klik - CTR
  12: "percentage", // Pesanan per Klik(Pesanan Dibuat) - CTOR Placed
  13: "percentage", // Pesanan per Klik(Pesanan Siap Dikirim) - CTOR Shipped

  // Text columns (default for remaining)
  0: "text", // Periode Data - Date
  1: "text", // User Id
};

// Shopee new headers to append to raw data
// For daily CSV: Start Date, Week in Year, Month, GMV/hour (no Start/End Time)
// For monthly: Start Date, Start Time, End Time, Week in Year, Month, GMV/hour
export const shopeeNewRawDataHeaders = [
  "Start Date",
  "Start Time",
  "End Time",
  "Week in Year",
  "Month",
  "GMV/hour",
] as const;

const shopeeNewRawDataColumns = ["R", "S", "T", "U", "V", "W"] as const;

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
      const valueA = cellA.value?.toString().toLowerCase() || "";

      // Check for "Periode Data" in column A (common to both formats)
      if (!valueA.includes("periode data") && !valueA.includes("data period")) {
        continue;
      }

      // Check for old Shopee Monthly livestream format
      // Look for "Nama Livestream" in column D
      const cellD = worksheet.getCell(`D${row}`);
      const valueD = cellD.value?.toString().toLowerCase() || "";
      if (valueD.includes("nama livestream") || valueD.includes("livestream")) {
        return row + 1; // Data starts in the next row
      }

      // Check for new Shopee Daily CSV format
      // Look for distinctive headers like "Penjualan(Pesanan Dibuat)" in column C
      const cellC = worksheet.getCell(`C${row}`);
      const valueC = cellC.value?.toString().toLowerCase() || "";
      if (
        valueC.includes("penjualan") &&
        (valueC.includes("pesanan dibuat") || valueC.includes("placed order"))
      ) {
        // This is likely the daily CSV format
        // Check if row+1 has actual headers (row 2 in CSV)
        const nextRow = row + 1;
        const nextCellA = worksheet.getCell(`A${nextRow}`);
        const nextValueA = nextCellA.value?.toString().toLowerCase() || "";
        if (
          nextValueA.includes("periode data") ||
          nextValueA.includes("data period")
        ) {
          // Row+1 is the header row, data starts at row+2
          return nextRow + 1;
        }
        // Otherwise, data starts right after this header row
        return row + 1;
      }
    }
    return null;
  },
  buildDataRanges: (startRow: number, lastRow: number): DataRangeStrings => {
    const seg = (c: string) => `\$${c}\$${startRow}:\$${c}\$${lastRow}`;
    return {
      startDateR: `'clean data'!${seg("R")}`, // Start Date (from Periode Data)
      startTimeR: `'clean data'!${seg("S")}`, // Start Time
      endTimeR: `'clean data'!${seg("T")}`, // End Time
      weekInYearR: `'clean data'!${seg("U")}`, // Week in Year
      montIndex: `'clean data'!${seg("V")}`, // Month Index
      grossR: `'clean data'!${seg("C")}`, // Penjualan(Pesanan Dibuat) - GMV Placed
      directR: `'clean data'!${seg("D")}`, // Penjualan(Pesanan Siap Dikirim) - GMV Shipped
      itemsR: `'clean data'!${seg("G")}`, // Produk Terjual(Pesanan Dibuat) - Items Sold Placed
      avgViewR: `'clean data'!${seg("K")}`, // Rata-rata durasi ditonton - Avg Viewing Duration
      likesR: `'clean data'!${seg("O")}`, // Suka - Likes
      commentsR: `'clean data'!${seg("Q")}`, // Komentar - Comments
      sharesR: `'clean data'!${seg("P")}`, // Share - Shares
      ctrR: `'clean data'!${seg("L")}`, // Persentase Klik - CTR
      ctorR: `'clean data'!${seg("M")}`, // Pesanan per Klik(Pesanan Dibuat) - CTOR Placed
      gmv1kShowsR: `'clean data'!${seg("C")}`, // Placeholder (Shopee doesn't have GMV/1K shows)
    };
  },
  generateRawDataFormulas: (row: number, isDailyFormat?: boolean): Record<string, string> => {
    if (isDailyFormat) {
      // Daily CSV format: Get start date from Periode Data, no start/end time
      return {
        // Start Date: Extract date from Periode Data (column A)
        // Handle DD-MM-YYYY format (common in Indonesian date formats)
        R: `IFERROR(IF(ISNUMBER(A${row}),A${row},IF(ISERROR(DATEVALUE(A${row})),DATE(MID(A${row},7,4),MID(A${row},4,2),LEFT(A${row},2)),DATEVALUE(A${row}))),"")`, // Start Date
        // Start Time: Empty for daily format
        S: `""`, // Start Time (not used for daily)
        // End Time: Empty for daily format
        T: `""`, // End Time (not used for daily)
        // Week in Year: Calculate week number from Start Date
        U: `IFERROR(IF(R${row}<>"",WEEKNUM(R${row},2),""),"")`, // Week in Year
        // Month: Extract month from Start Date
        V: `IFERROR(IF(R${row}<>"",MONTH(R${row}),""),"")`, // Month Index
        // GMV/hour: Not applicable for daily format (no duration), set to 0
        W: `0`, // GMV/hour (not applicable for daily format)
      };
    } else {
      // Monthly format: Get start date from Start Time column (column E in old format, which maps to a different column)
      // Note: In the old format, Start Time is in column 4 (index 4), but we need to check the original worksheet
      // For now, we'll extract from Periode Data and Start Time if available
      // Actually, looking at the old format, "Start Time" column contains datetime, so we can extract date from it
      // But since we're in clean data, we don't have the original Start Time column anymore
      // We need to get it from the original mapping - but we can't access that here
      // Let's use a formula that checks if there's a datetime in column A (Periode Data might have datetime)
      return {
        // Start Date: For monthly, try to extract from Periode Data (which might contain datetime)
        // If Periode Data has datetime, extract date part; otherwise parse as date
        R: `IFERROR(IF(ISNUMBER(A${row}),INT(A${row}),IF(ISERROR(DATEVALUE(A${row})),DATE(MID(A${row},7,4),MID(A${row},4,2),LEFT(A${row},2)),INT(DATEVALUE(A${row})))),"")`, // Start Date
        // Start Time: Extract time from Periode Data if it's a datetime, otherwise default to 00:00
        S: `IFERROR(IF(ISNUMBER(A${row}),MOD(A${row},1),TIMEVALUE(A${row})),TIME(0,0,0))`, // Start Time
        // End Time: Start Time + Duration from "Durasi" column
        // Note: In old format, Durasi is column 5 (index 5), but in clean data we don't have it directly
        // We'll need to calculate from Rata-rata durasi ditonton (column K) if available
        // Actually, old format has "Durasi" which is different from "Durasi Rata-Rata Menonton"
        // For now, use Rata-rata durasi ditonton (K) converted to hours
        T: `IFERROR(IF(AND(K${row}<>"",K${row}<>0),S${row}+(TIMEVALUE("00:"&SUBSTITUTE(K${row},":",""))*24),S${row}),S${row})`, // End Time
        // Week in Year: Calculate week number from Start Date
        U: `IFERROR(IF(R${row}<>"",WEEKNUM(R${row},2),""),"")`, // Week in Year
        // Month: Extract month from Start Date
        V: `IFERROR(IF(R${row}<>"",MONTH(R${row}),""),"")`, // Month Index
        // GMV/hour = Penjualan(Pesanan Siap Dikirim) / Duration
        // For monthly, we need to get duration from the original "Durasi" column
        // Since we don't have it in clean data, we'll skip this or use a placeholder
        W: `0`, // GMV/hour (would need original Durasi column)
      };
    }
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
      "Avg Viewers": `ROUND(AVERAGE('clean data'!$I$${startRow}:$I$${lastRow}),0)`, // Penonton column (I)
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


