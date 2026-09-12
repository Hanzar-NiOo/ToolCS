// function createSummaryTableStructure(sheet, startRow)
// {
//   try
//   {
//     sheet.getRange(startRow, 1, 1, 6).setValues
//     ([
//       ["Use case", "Transaction Counts", "", "Amount", "", ""]
//     ]);

//     sheet.getRange(startRow + 1, 1, 1, 6).setValues
//     ([
//       ["", "LRD Transactions", "USD Transactions", "LRD Value Processed", "LRD Value equilvalent USD(184LRD/USD)", "USD Value Processed"]
//     ]);

//     // Perform Cell Merging to match the layout perfectly
//     sheet.getRange(startRow, 1, 2, 1).merge();
//     sheet.getRange(startRow, 2, 1, 2).merge();
//     sheet.getRange(startRow, 4, 1, 3).merge();
//     sheet.getRange(startRow + 5, 2, 1, 2).merge();
//     sheet.getRange(startRow + 5, 4, 1, 3).merge();

//     // Write Row Labels
//     sheet.getRange(startRow + 2, 1).setValue("P2P");
//     sheet.getRange(startRow + 3, 1).setValue("G2P");
//     sheet.getRange(startRow + 4, 1).setValue("Total");
//     sheet.getRange(startRow + 5, 1).setValue("Grand Total");

//     // Style the Header Blocks (Blue Fill, White Text, Center Aligned)
//     const headerRange = sheet.getRange(startRow, 1, 2, 6);
//     headerRange.setBackground("#3474b4")
//               .setFontColor("#ffffff")
//               .setFontWeight("bold")
//               .setHorizontalAlignment("center")
//               .setVerticalAlignment("middle");

//     // Apply Bold formatting to the "Total" and "Grand Total"
//     sheet.getRange(startRow + 4, 1).setFontWeight("bold");
//     sheet.getRange(startRow + 5, 1).setFontWeight("bold");

//     // Apply clean, uniform borders around and inside the entire summary table grid
//     const totalTableRange = sheet.getRange(startRow, 1, 6, 6);
//     totalTableRange.setBorder(
//       true,
//       true,
//       true,
//       true,
//       true,
//       true,
//       "#000000",
//       SpreadsheetApp.BorderStyle.SOLID
//     );
//   }
//   catch (e)
//   {
//     throw new Error("Failed to create the summary table: " + e.message);
//   }
// }

function createSummaryTableStructure(sheet, startRow)
{
  try
  {
    sheet.getRange(startRow, 1, 1, 6).setValues
    ([
      ["Use case", "Transaction Counts", "", "Amount", "", ""]
    ]);

    sheet.getRange(startRow + 1, 1, 1, 6).setValues
    ([
      ["", "LRD Transactions", "USD Transactions", "LRD Value Processed", "LRD Value equilvalent USD(184LRD/USD)", "USD Value Processed"]
    ]);

    // Perform Cell Merging to match the layout perfectly
    sheet.getRange(startRow, 1, 2, 1).merge();
    sheet.getRange(startRow, 2, 1, 2).merge();
    sheet.getRange(startRow, 4, 1, 3).merge();
    sheet.getRange(startRow + 5, 2, 1, 2).merge();
    sheet.getRange(startRow + 5, 4, 1, 3).merge();

    // Write Row Labels
    sheet.getRange(startRow + 2, 1).setValue("P2P");
    sheet.getRange(startRow + 3, 1).setValue("G2P");
    sheet.getRange(startRow + 4, 1).setValue("Total");
    sheet.getRange(startRow + 5, 1).setValue("Grand Total");

    // Apply typography per style guide BEFORE the color override below
    applyTextStyle(sheet.getRange(startRow, 1, 1, 6), "heading1");      // top header row
    applyTextStyle(sheet.getRange(startRow + 1, 1, 1, 6), "heading2");  // sub-header row
    applyTextStyle(sheet.getRange(startRow + 2, 1), "body");            // P2P
    applyTextStyle(sheet.getRange(startRow + 3, 1), "body");            // G2P
    applyTextStyle(sheet.getRange(startRow + 4, 1), "heading3", { bold: true }); // Total
    applyTextStyle(sheet.getRange(startRow + 5, 1), "heading3", { bold: true }); // Grand Total

    // Style the Header Blocks (Blue Fill, White Text, Center Aligned)
    // NOTE: setFontColor here intentionally overrides heading1/heading2 colors to keep white-on-blue
    const headerRange = sheet.getRange(startRow, 1, 2, 6);
    headerRange.setBackground("#3474b4")
              .setFontColor("#ffffff")
              .setFontWeight("bold")
              .setHorizontalAlignment("center")
              .setVerticalAlignment("middle");

    // Apply clean, uniform borders around and inside the entire summary table grid
    const totalTableRange = sheet.getRange(startRow, 1, 6, 6);
    totalTableRange.setBorder(
      true, true, true, true, true, true,
      "#000000",
      SpreadsheetApp.BorderStyle.SOLID
    );
  }
  catch (e)
  {
    throw new Error("Failed to create the summary table: " + e.message);
  }
}

// function addAbortedTransactionsBox(sheet, startRow, abortedCount)
// {
//   try
//   {
//     // Header cell
//     sheet.getRange(startRow, 1).setValue("Aborted Transactions");
//     sheet.getRange(startRow, 1).setFontWeight("bold");

//     // Value cell
//     sheet.getRange(startRow + 1, 1).setValue(abortedCount);
//     sheet.getRange(startRow + 1, 1).setHorizontalAlignment("right");

//     // Border around the box
//     sheet.getRange(startRow, 1, 2, 1).setBorder(
//       true,
//       true,
//       true,
//       true,
//       true,
//       true,
//       "#000000",
//       SpreadsheetApp.BorderStyle.SOLID
//     );
//   }
//   catch (e)
//   {
//     throw new Error("Failed to create the aborted transactions box: " + e.message);
//   }
// }

function addAbortedTransactionsBox(sheet, startRow, abortedCount)
{
  try
  {
    // Header cell
    sheet.getRange(startRow, 1).setValue("Aborted Transactions");
    applyTextStyle(sheet.getRange(startRow, 1), "heading3", { bold: true });

    // Value cell
    sheet.getRange(startRow + 1, 1).setValue(abortedCount);
    applyTextStyle(sheet.getRange(startRow + 1, 1), "body");
    sheet.getRange(startRow + 1, 1).setHorizontalAlignment("right");

    // Border around the box
    sheet.getRange(startRow, 1, 2, 1).setBorder(
      true, true, true, true, true, true,
      "#000000",
      SpreadsheetApp.BorderStyle.SOLID
    );
  }
  catch (e)
  {
    throw new Error("Failed to create the aborted transactions box: " + e.message);
  }
}