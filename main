/**
 * Calculates monthly usage costs applying discounted and standard pricing
 * with a global cost threshold.
 *
 * Designed to run inside a Google Spreadsheet.
 */
function calculateMonthlyConsumption() {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = spreadsheet.getActiveSheet();
    const ui = SpreadsheetApp.getUi();

    ui.alert('Starting monthly consumption calculation');

    // Pricing sheet name is read dynamically from cell S1 (row 1, column 19)
    const pricingSheetName = sheet.getRange(1, 19).getValue();
    const pricingSheet = spreadsheet.getSheetByName(pricingSheetName);

    if (!pricingSheet) {
      throw new Error('Pricing sheet not found');
    }

    // Clear previous output (W2:AD)
    sheet.getRange("W2:AD" + sheet.getLastRow()).clearContent();

    // Read input data
    const usageData = sheet.getRange(2, 14, sheet.getLastRow() - 1, 3).getValues(); // SKU, Usage, Country
    const pricingData = pricingSheet.getRange(2, 1, pricingSheet.getLastRow() - 1, 3).getValues(); // Country, Discounted, Standard

    // Build price map: { country: [discountedPrice, standardPrice] }
    const priceMap = {};
    pricingData.forEach(row => {
      priceMap[row[0]] = [parseFloat(row[1]) || 0, parseFloat(row[2]) || 0];
    });

    let copiedData = [];
    let appliedPrices = [];
    let discountedCosts = [];
    let standardCosts = [];

    let lastCountry = null;
    let lastDiscountedPrice = null;
    let lastStandardPrice = null;

    // Process rows
    usageData.forEach(row => {
      const sku = row[0];
      const usage = parseFloat(row[1]) || 0;
      const country = row[2];

      copiedData.push([sku, usage, country]);

      if (country === lastCountry) {
        appliedPrices.push([lastDiscountedPrice, lastStandardPrice]);
      } else {
        const prices = priceMap[country] || [0, 0];
        appliedPrices.push(prices);
        lastDiscountedPrice = prices[0];
        lastStandardPrice = prices[1];
        lastCountry = country;
      }

      discountedCosts.push([usage * lastDiscountedPrice]);
      standardCosts.push([usage * lastStandardPrice]);
    });

    // Write intermediate results
    sheet.getRange(2, 23, copiedData.length, 3).setValues(copiedData);       // W:Y
    sheet.getRange(2, 26, appliedPrices.length, 2).setValues(appliedPrices); // Z:AA
    sheet.getRange(2, 28, discountedCosts.length, 1).setValues(discountedCosts); // AB
    sheet.getRange(2, 29, standardCosts.length, 1).setValues(standardCosts);     // AC

    // Threshold logic
    const threshold = parseFloat(sheet.getRange("S4").getValue()) || 0;
    const usages = sheet.getRange(2, 24, copiedData.length, 1).getValues(); // X
    const discountPrices = sheet.getRange(2, 26, copiedData.length, 1).getValues(); // Z
    const standardPrices = sheet.getRange(2, 27, copiedData.length, 1).getValues(); // AA

    let cumulativeTotal = 0;
    let cumulativeResults = [];
    let thresholdExceeded = false;

    // Process bottom-up
    for (let i = copiedData.length - 1; i >= 0; i--) {
      const usage = parseFloat(usages[i][0]) || 0;
      const discountPrice = parseFloat(discountPrices[i][0]) || 0;
      const standardPrice = parseFloat(standardPrices[i][0]) || 0;

      let rowCost = 0;

      if (!thresholdExceeded) {
        const discountedCost = usage * discountPrice;

        if (cumulativeTotal + discountedCost <= threshold) {
          rowCost = discountedCost;
        } else {
          const remainingBudget = threshold - cumulativeTotal;
          const discountedUsage = discountPrice > 0 ? remainingBudget / discountPrice : 0;
          const standardUsage = usage - discountedUsage;

          rowCost =
            discountedUsage * discountPrice +
            standardUsage * standardPrice;

          thresholdExceeded = true;
        }
      } else {
        rowCost = usage * standardPrice;
      }

      cumulativeTotal += rowCost;
      cumulativeResults.unshift([cumulativeTotal]);
    }

    // Write cumulative totals
    sheet.getRange(2, 30, cumulativeResults.length, 1).setValues(cumulativeResults); // AD

    ui.alert('Calculation completed successfully');

  } catch (error) {
    SpreadsheetApp.getUi().alert('Error: ' + error.message);
    Logger.log(error);
  }
}
