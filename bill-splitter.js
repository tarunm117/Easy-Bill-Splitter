function buildBillTableFull() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  // Inputs
  const numFood = sheet.getRange("B1").getValue();
  const numAlcohol = sheet.getRange("B2").getValue();
  const numPeople = sheet.getRange("B3").getValue();
  const discountTotal = sheet.getRange("B4").getValue();   // Overall discount applied to the total bill
  const discountFood = sheet.getRange("B5").getValue();     // Discount specifically for food items
  const discountAlcohol = sheet.getRange("B6").getValue();  // Discount specifically for alcohol items

  // NEW: Read tax amounts and service charge from cells E1:E3
  const foodTaxAmount = sheet.getRange("E1").getValue();     // Absolute amount of food tax
  const alcoholTaxAmount = sheet.getRange("E2").getValue();  // Absolute amount of alcohol tax
  const serviceChargeAmount = sheet.getRange("E3").getValue(); // Absolute amount of service charge

  // Clear old output
  sheet.getRange("A8:Z200").clearContent();

  let startRow = 8;

  // --- Headers: Food ---
  sheet.getRange(startRow, 1).setValue("People");
  for (let i = 1; i <= numFood; i++) {
    sheet.getRange(startRow, 1 + i).setValue("Food " + i);
  }

  // --- Headers: Alcohol ---
  const alcoholStartCol = 1 + numFood + 1;
  for (let i = 1; i <= numAlcohol; i++) {
    sheet.getRange(startRow, alcoholStartCol + i - 1).setValue("Alcohol " + i);
  }

  // --- People rows (consumption matrix) ---
  for (let i = 1; i <= numPeople; i++) {
    sheet.getRange(startRow + i, 1).setValue("Person " + i);
    for (let col = 2; col <= numFood + numAlcohol + 1; col++) {
      sheet.getRange(startRow + i, col).setValue(0); // initialize with 0 consumption
    }
  }

  // --- Totals row for consumption ---
  const totalRow = startRow + numPeople + 1;
  sheet.getRange(totalRow, 1).setValue("Total");
  for (let col = 2; col <= numFood + numAlcohol + 1; col++) {
    let colLetter = columnToLetter(col);
    sheet.getRange(totalRow, col).setFormula(`=SUM(${colLetter}${startRow + 1}:${colLetter}${startRow + numPeople})`);
  }

  // --- Menu price rows ---
  const priceRow = totalRow + 2;
  sheet.getRange(priceRow, 1).setValue("Menu price");
  for (let i = 1; i <= numFood; i++) {
    sheet.getRange(priceRow, 1 + i).setValue(0); // placeholder
  }
  for (let i = 1; i <= numAlcohol; i++) {
    sheet.getRange(priceRow, alcoholStartCol + i - 1).setValue(0); // placeholder
  }

  // --- Cost summary (Pre-Tax & Post-Tax totals) ---
  const costRowStart = priceRow + 2; // Adjusted as Final Price row is removed

  // NEW LAYOUT FOR COST SUMMARY headers
  sheet.getRange(costRowStart, 1).setValue("Cost");
  sheet.getRange(costRowStart, 2).setValue("Total Food");
  sheet.getRange(costRowStart, 3).setValue("Total Alcohol");
  sheet.getRange(costRowStart, 4).setValue("Final Total");

  let foodMenuPriceRange = sheet.getRange(priceRow, 2, 1, numFood).getA1Notation(); // Range of Menu Prices for Food
  let alcoholMenuPriceRange = sheet.getRange(priceRow, alcoholStartCol, 1, numAlcohol).getA1Notation(); // Range of Menu Prices for Alcohol

  // Pre Tax Row
  sheet.getRange(costRowStart + 1, 1).setValue("Pre Tax");
  sheet.getRange(costRowStart + 1, 2).setFormula(`=SUM(${foodMenuPriceRange})`); // Total Food (Pre Tax)
  sheet.getRange(costRowStart + 1, 3).setFormula(`=SUM(${alcoholMenuPriceRange})`); // Total Alcohol (Pre Tax)
  sheet.getRange(costRowStart + 1, 4).setFormula(`=B${costRowStart + 1}+C${costRowStart + 1}`); // Final Total (Pre Tax)

  // Post Tax Row
  sheet.getRange(costRowStart + 2, 1).setValue("Post Tax");
  sheet.getRange(costRowStart + 2, 2).setFormula(`=B${costRowStart + 1}+$E$1`); // Total Food (Post Tax) = Pre Tax Food + Food Tax Amount
  sheet.getRange(costRowStart + 2, 3).setFormula(`=C${costRowStart + 1}+$E$2`); // Total Alcohol (Post Tax) = Pre Tax Alcohol + Alcohol Tax Amount
  // Final Total (Post Tax, BEFORE overall discount is applied) - this is the base for distributing effectiveTotalDiscount
  sheet.getRange(costRowStart + 2, 4).setFormula(`=B${costRowStart + 2}+C${costRowStart + 2}`);

  // New row for Final Total BEFORE discount (including service charge)
  sheet.getRange(costRowStart + 3, 1).setValue("Final Total (pre discount)");
  sheet.getRange(costRowStart + 3, 4).setFormula(`=D${costRowStart + 2} + $E$3`); // Adds service charge to Post-Tax total

  // New row for Final Total AFTER discount
  sheet.getRange(costRowStart + 4, 1).setValue("Final Total (After Discount)");
  sheet.getRange(costRowStart + 4, 4).setFormula(`=D${costRowStart + 3} - ($B$4 + $B$5 + $B$6)`);

  // --- People share table ---
  const peopleRowStart = costRowStart + 6; // Adjusted starting row based on new Cost Summary height
  // NEW headers: Removed "Service Charge", added "Food Tax" and "Alcohol Tax"
  sheet.getRange(peopleRowStart, 1, 1, 7).setValues([["People","Food","Alcohol","Food Tax","Alcohol Tax","Discount","Total"]]);

  // References to Cost Summary totals for proportional distribution
  const totalFoodMenuPriceRef = `$B$${costRowStart + 1}`; // Total Food (Pre Tax)
  const totalAlcoholMenuPriceRef = `$C$${costRowStart + 1}`; // Total Alcohol (Pre Tax)
  const totalPostTaxBeforeOverallDiscountRef = `$D$${costRowStart + 2}`; // Total bill (Post Tax, BEFORE overall discount)

  for (let i = 1; i <= numPeople; i++) {
    let row = peopleRowStart + i;
    sheet.getRange(row, 1).setFormula(`=A${startRow + i}`);

    // Food share (Column B in people share table) - based on menu price
    let foodShareFormulas = [];
    for (let j = 1; j <= numFood; j++) {
      let consCell = sheet.getRange(startRow + i, 1 + j).getA1Notation();
      let totCell = sheet.getRange(totalRow, 1 + j).getA1Notation();
      let menuPriceCell = sheet.getRange(priceRow, 1 + j).getA1Notation();
      foodShareFormulas.push(`IF(${totCell}>0,${consCell}/${totCell}*${menuPriceCell},0)`);
    }
    sheet.getRange(row, 2).setFormula("=" + foodShareFormulas.join("+"));

    // Alcohol share (Column C in people share table) - based on menu price
    let alcoholShareFormulas = [];
    for (let j = 1; j <= numAlcohol; j++) {
      let consCell = sheet.getRange(startRow + i, alcoholStartCol + j - 1).getA1Notation();
      let totCell = sheet.getRange(totalRow, alcoholStartCol + j - 1).getA1Notation();
      let menuPriceCell = sheet.getRange(priceRow, alcoholStartCol + j - 1).getA1Notation();
      alcoholShareFormulas.push(`IF(${totCell}>0,${consCell}/${totCell}*${menuPriceCell},0)`);
    }
    sheet.getRange(row, 3).setFormula("=" + alcoholShareFormulas.join("+"));

    // NEW: Food Tax Share (Column D in people share table)
    // Proportional to person's food menu price share
    sheet.getRange(row, 4).setFormula(
      `=IF(${totalFoodMenuPriceRef}>0, (B${row}/${totalFoodMenuPriceRef})*$E$1, 0)`
    );

    // NEW: Alcohol Tax Share (Column E in people share table)
    // Proportional to person's alcohol menu price share
    sheet.getRange(row, 5).setFormula(
      `=IF(${totalAlcoholMenuPriceRef}>0, (C${row}/${totalAlcoholMenuPriceRef})*$E$2, 0)`
    );

    // Discount formula (Column F in people share table)
    // Combines food-specific, alcohol-specific, and the effective total discount (B4-E3).
    // The effectiveTotalDiscount is applied against the total of (person's food + alcohol + food tax + alcohol tax)
    // relative to the overall total bill (post tax, before overall discount)
    // FIX 2: Corrected formula to use string literals "($B$4-$E$3)" instead of undefined variables (B4-E3)
    sheet.getRange(row, 6).setFormula(
      `=-(B${row}*$B$5/$B${peopleRowStart + numPeople + 1} + ` +
      `C${row}*$B$6/$C${peopleRowStart + numPeople + 1} + ` +
      `(B${row}+C${row}+D${row}+E${row})*($B$4-$E$3)/${totalPostTaxBeforeOverallDiscountRef})`
    );

    // Total = Food (menu) + Alcohol (menu) + Food Tax + Alcohol Tax + Discount (Column G in people share table)
    sheet.getRange(row, 7).setFormula(`=B${row}+C${row}+D${row}+E${row}+F${row}`);
  }

  // Totals row for People share table
  let totalsRow = peopleRowStart + numPeople + 1;
  sheet.getRange(totalsRow, 1).setValue("Total");
  sheet.getRange(totalsRow, 2).setFormula(`=SUM(B${peopleRowStart + 1}:B${peopleRowStart + numPeople})`); // Total Food Share (menu)
  sheet.getRange(totalsRow, 3).setFormula(`=SUM(C${peopleRowStart + 1}:C${peopleRowStart + numPeople})`); // Total Alcohol Share (menu)
  sheet.getRange(totalsRow, 4).setFormula(`=SUM(D${peopleRowStart + 1}:D${peopleRowStart + numPeople})`); // Total Food Tax Share
  sheet.getRange(totalsRow, 5).setFormula(`=SUM(E${peopleRowStart + 1}:E${peopleRowStart + numPeople})`); // Total Alcohol Tax Share
  sheet.getRange(totalsRow, 6).setFormula(`=SUM(F${peopleRowStart + 1}:F${peopleRowStart + numPeople})`); // Total Discount
  sheet.getRange(totalsRow, 7).setFormula(`=SUM(G${peopleRowStart + 1}:G${peopleRowStart + numPeople})`); // Grand Total (should match Cost Summary's Final Total After Discount)

   // --- Final Summary Table (Names and Totals Only) ---
  const summaryTableStartRow = totalsRow + 2; // Start 2 rows below the detailed table's total

  // Set headers for the new table
  sheet.getRange(summaryTableStartRow, 1, 1, 2).setValues([["People", "Final Total"]]);

  // Loop through each person to create formula-based references
  for (let i = 1; i <= numPeople; i++) {
    let sourceRow = peopleRowStart + i; // The row in the detailed table we're copying from
    let targetRow = summaryTableStartRow + i; // The row in our new summary table

    // Set formula for the name (e.g., =A25)
    sheet.getRange(targetRow, 1).setFormula(`=A${sourceRow}`);
    
    // Set formula for the final total (e.g., =G25)
    sheet.getRange(targetRow, 2).setFormula(`=G${sourceRow}`);
  }
}

// Helper to convert column number to letter
function columnToLetter(column) {
  let temp, letter = '';
  while (column > 0) {
    temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    column = (column - temp - 1) / 26;
  }
  return letter;
}
