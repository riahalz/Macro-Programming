function calculateCommission() {
  
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  
  // Salespeople names and total sales
  var salesData = [
    ['Riya H', 10170],
    ['Sneha R', 11700],
    ['Vaibhavi K', 10335],
    ['Misaki', 6010],
    ['Takumi', 11585]
  ];

  // Headings
  sheet.getRange(1, 1).setValue('Salesperson Name');
  sheet.getRange(1, 2).setValue('Total Sales ($)');
  sheet.getRange(1, 3).setValue('Commission ($)');

  // Tiered commission rates
  var commissionRates = [
    { min: 0, max: 4999, rate: 0.05 },
    { min: 5000, max: 9999, rate: 0.10 },
    { min: 10000, max: Infinity, rate: 0.15 }
  ];

  // Calculate commission for each salesperson
  for (var i = 0; i < salesData.length; i++) {
    var salesperson = salesData[i][0];
    var totalSales = salesData[i][1];
    var commission = 0;

    for (var j = 0; j < commissionRates.length; j++) {
      var tier = commissionRates[j];
      if (totalSales > tier.min) {
        var salesInTier = Math.min(totalSales, tier.max) - tier.min;
        commission += salesInTier * tier.rate;
        totalSales -= salesInTier;
      }
    }

    // Print salesperson, total sales, commission
    sheet.getRange(i + 2, 1).setValue(salesperson);
    sheet.getRange(i + 2, 2).setValue(salesData[i][1]);
    sheet.getRange(i + 2, 3).setValue(commission.toFixed(2));
  }
}
