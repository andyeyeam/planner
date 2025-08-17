function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Retirement Planning')
    .addItem('Forecast Plan', 'createMonthsList')
    .addToUi();
}

function createMonthsList() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Get People sheet to find birth dates
  var peopleSheet = ss.getSheetByName('People');
  if (!peopleSheet) {
    SpreadsheetApp.getUi().alert('People sheet not found!');
    return;
  }
  
  // Get people data from People sheet (A = name, B = birth date)
  var peopleData = [];
  var peopleLastRow = peopleSheet.getLastRow();
  if (peopleLastRow >= 2) {
    var peopleRange = peopleSheet.getRange('A2:B' + peopleLastRow).getValues();
    for (var i = 0; i < peopleRange.length; i++) {
      var name = peopleRange[i][0];
      var birthDate = peopleRange[i][1];
      
      if (name && name.toString().trim() !== '') {
        peopleData.push({
          name: name.toString(),
          birthDate: new Date(birthDate)
        });
      }
    }
  }
  
  // For backward compatibility, get first two birth dates
  var andrewBirth = peopleData.length > 0 ? peopleData[0].birthDate : null;
  var alisonBirth = peopleData.length > 1 ? peopleData[1].birthDate : null;
  
  // Calculate 90th birthdays
  var andrew90th = new Date(andrewBirth.getFullYear() + 90, andrewBirth.getMonth(), andrewBirth.getDate());
  var alison90th = new Date(alisonBirth.getFullYear() + 90, alisonBirth.getMonth(), alisonBirth.getDate());
  
  // Find the later of the two 90th birthdays
  var endDate = andrew90th > alison90th ? andrew90th : alison90th;

  // Get Plan sheet to find growth rate
  var planSheet = ss.getSheetByName('Plan');
  if (!planSheet) {
    SpreadsheetApp.getUi().alert('Plan sheet not found!');
    return;
  }
  
  // Search for growth rates, state pension value, and inflation rate in Plan sheet
  var planData = planSheet.getRange('A:B').getValues();
  var stocksGrowthRate = null;
  var cashGrowthRate = null;
  var currentStatePension = null;
  var annualInflationRate = null;
  
  for (var i = 0; i < planData.length; i++) {
    if (planData[i][0]) {
      var propertyName = planData[i][0].toString().toLowerCase();
      if (propertyName.includes('stocks and shares growth')) {
        stocksGrowthRate = planData[i][1];
      }
      if (propertyName.includes('cash growth')) {
        cashGrowthRate = planData[i][1];
      }
      if (propertyName.includes('current state pension')) {
        currentStatePension = planData[i][1];
      }
      if (propertyName.includes('annual inflation rate')) {
        annualInflationRate = planData[i][1];
      }
    }
  }
  
  if (stocksGrowthRate === null) {
    SpreadsheetApp.getUi().alert('Stocks and Shares growth rate not found in Plan sheet!');
    return;
  }
  
  if (cashGrowthRate === null) {
    SpreadsheetApp.getUi().alert('Cash growth rate not found in Plan sheet!');
    return;
  }
  
  if (currentStatePension === null) {
    SpreadsheetApp.getUi().alert('Current State Pension value not found in Plan sheet!');
    return;
  }
  
  if (annualInflationRate === null) {
    SpreadsheetApp.getUi().alert('Annual Inflation Rate not found in Plan sheet!');
    return;
  }
  
  var monthlyStocksGrowthRate = stocksGrowthRate / 12; // Convert annual to monthly
  var monthlyCashGrowthRate = cashGrowthRate / 12; // Convert annual to monthly
  
  
  // Get Stocks and Shares sheet to find current values
  var stocksSharesSheet = ss.getSheetByName('Stocks and Shares');
  if (!stocksSharesSheet) {
    SpreadsheetApp.getUi().alert('Stocks and Shares sheet not found!');
    return;
  }
  
  // Get current values from Stocks and Shares table - now including columns D and E
  var stocksData = stocksSharesSheet.getRange('A2:E' + stocksSharesSheet.getLastRow()).getValues();
  
  // Calculate total current tax-free and taxable values, and collect contribution data
  var totalTaxFree = 0;
  var totalTaxable = 0;
  var contributionData = [];
  
  for (var i = 0; i < stocksData.length; i++) {
    var title = stocksData[i][0];
    var taxFreeValue = stocksData[i][1] || 0;
    var taxableValue = stocksData[i][2] || 0;
    var monthlyContribution = stocksData[i][3] || 0;
    var lastMonth = stocksData[i][4];
    
    if (title && title.toString().trim() !== '') {
      totalTaxFree += Number(taxFreeValue);
      totalTaxable += Number(taxableValue);
      
      // Collect contribution data for this stock/share
      if (Number(monthlyContribution) > 0 && lastMonth) {
        contributionData.push({
          title: title.toString(),
          monthlyContribution: Number(monthlyContribution),
          lastMonth: new Date(lastMonth),
          monthlyTaxFree: Math.round((Number(monthlyContribution) * 0.25) * 100) / 100,
          monthlyTaxable: Math.round((Number(monthlyContribution) * 0.75) * 100) / 100
        });
      }
    }
  }
  
  // Get Cash sheet to find current cash values
  var cashSheet = ss.getSheetByName('Cash');
  if (!cashSheet) {
    SpreadsheetApp.getUi().alert('Cash sheet not found!');
    return;
  }
  
  // Get current cash values from column B - sum all numeric values
  var cashData = cashSheet.getRange('B:B').getValues();
  var totalCash = 0;
  
  for (var i = 0; i < cashData.length; i++) {
    var cashValue = cashData[i][0];
    if (cashValue && !isNaN(Number(cashValue))) {
      totalCash += Number(cashValue);
    }
  }
  
  // Get Final Salary Pension sheet data
  var finalSalaryPensionSheet = ss.getSheetByName('Final Salary Pension');
  if (!finalSalaryPensionSheet) {
    SpreadsheetApp.getUi().alert('Final Salary Pension sheet not found!');
    return;
  }
  
  // Get pension data from row 3 onwards (A = title, B = monthly income, C = first month)
  var pensionData = [];
  var lastRow = finalSalaryPensionSheet.getLastRow();
  if (lastRow >= 3) {
    var pensionRange = finalSalaryPensionSheet.getRange('A3:C' + lastRow).getValues();
    for (var i = 0; i < pensionRange.length; i++) {
      var title = pensionRange[i][0];
      var monthlyIncome = pensionRange[i][1];
      var firstMonth = pensionRange[i][2];
      
      if (title && title.toString().trim() !== '') {
        pensionData.push({
          title: title.toString(),
          monthlyIncome: Number(monthlyIncome) || 0,
          firstMonth: new Date(firstMonth)
        });
      }
    }
  }
  
  // Get Occasional Income sheet data
  var occasionalIncomeSheet = ss.getSheetByName('Occasional Income');
  if (!occasionalIncomeSheet) {
    SpreadsheetApp.getUi().alert('Occasional Income sheet not found!');
    return;
  }
  
  // Get occasional income data (A = title, B = value, C = month)
  var occasionalIncomeData = [];
  var occasionalLastRow = occasionalIncomeSheet.getLastRow();
  if (occasionalLastRow >= 2) {
    var occasionalRange = occasionalIncomeSheet.getRange('A2:C' + occasionalLastRow).getValues();
    for (var i = 0; i < occasionalRange.length; i++) {
      var title = occasionalRange[i][0];
      var value = occasionalRange[i][1];
      var month = occasionalRange[i][2];
      
      if (title && title.toString().trim() !== '') {
        occasionalIncomeData.push({
          title: title.toString(),
          value: Number(value) || 0,
          month: new Date(month)
        });
      }
    }
  }
  
  // Start from current month
  var currentDate = new Date();
  var startDate = new Date(currentDate.getFullYear(), currentDate.getMonth(), 1);
  
  // Create array of forecast data
  var forecastData = [];
  var date = new Date(startDate);
  var monthIndex = 0;
  
  // Track current values for growth calculation
  var currentTaxFree = totalTaxFree;
  var currentTaxable = totalTaxable;
  var currentCash = totalCash;
  
  // Track current pension values for growth calculation
  var currentPensionValues = [];
  for (var p = 0; p < pensionData.length; p++) {
    currentPensionValues.push(pensionData[p].monthlyIncome);
  }
  
  // Track current state pension value for inflation adjustment
  var currentStatePensionValue = currentStatePension;
  
  while (date <= endDate) {
    // Format as "Month Year" (e.g., "August 2025")
    var monthNames = ["January", "February", "March", "April", "May", "June",
                     "July", "August", "September", "October", "November", "December"];
    var monthStr = monthNames[date.getMonth()] + " " + date.getFullYear();
    
    // Apply annual inflation to state pension in April
    if (date.getMonth() === 3 && monthIndex > 0) { // April is month 3 (0-indexed)
      currentStatePensionValue = Math.round((currentStatePensionValue * (1 + annualInflationRate)) * 100) / 100;
    }
    
    // Create row array starting with basic data
    var rowData = [];
    
    // Calculate monthly contributions for this month
    var monthlyTaxFreeContributions = 0;
    var monthlyTaxableContributions = 0;
    
    for (var c = 0; c < contributionData.length; c++) {
      var contribution = contributionData[c];
      // Check if contribution end month is >= current forecast month
      if (contribution.lastMonth >= date) {
        monthlyTaxFreeContributions += contribution.monthlyTaxFree;
        monthlyTaxableContributions += contribution.monthlyTaxable;
      }
    }
    
    // For first month, use starting values; for others, apply growth to previous month
    if (monthIndex === 0) {
      currentTaxFree += monthlyTaxFreeContributions;
      currentTaxable += monthlyTaxableContributions;
      var stocksTotal = Math.round((currentTaxFree + currentTaxable) * 100) / 100;
      rowData = [monthStr, currentTaxFree, currentTaxable, stocksTotal, currentCash];
    } else {
      // Apply monthly growth to previous month's values
      currentTaxFree = Math.round((currentTaxFree * (1 + monthlyStocksGrowthRate)) * 100) / 100;
      currentTaxable = Math.round((currentTaxable * (1 + monthlyStocksGrowthRate)) * 100) / 100;
      currentCash = Math.round((currentCash * (1 + monthlyCashGrowthRate)) * 100) / 100;
      
      // Add monthly contributions after growth
      currentTaxFree += monthlyTaxFreeContributions;
      currentTaxable += monthlyTaxableContributions;
      
      var stocksTotal = Math.round((currentTaxFree + currentTaxable) * 100) / 100;
      rowData = [monthStr, currentTaxFree, currentTaxable, stocksTotal, currentCash];
    }
    
    // Add pension values for this month
    for (var p = 0; p < pensionData.length; p++) {
      var pension = pensionData[p];
      // Check if current forecast month is >= pension start month
      if (date >= pension.firstMonth) {
        // Apply growth to pension values for months after the first
        if (monthIndex > 0) {
          currentPensionValues[p] = Math.round((currentPensionValues[p] * (1 + monthlyCashGrowthRate)) * 100) / 100;
        }
        rowData.push(currentPensionValues[p]);
      } else {
        rowData.push(0); // No pension payment yet
      }
    }
    
    // Add occasional income for this month
    var occasionalIncomeValue = 0;
    for (var o = 0; o < occasionalIncomeData.length; o++) {
      var occasionalIncome = occasionalIncomeData[o];
      // Check if the month matches (year and month only, ignore day)
      if (date.getFullYear() === occasionalIncome.month.getFullYear() && 
          date.getMonth() === occasionalIncome.month.getMonth()) {
        occasionalIncomeValue += occasionalIncome.value;
      }
    }
    rowData.push(occasionalIncomeValue);
    
    // Add state pension values for each person
    for (var sp = 0; sp < peopleData.length; sp++) {
      var person = peopleData[sp];
      // Calculate 67th birthday
      var statePensionAge = new Date(person.birthDate.getFullYear() + 67, person.birthDate.getMonth(), person.birthDate.getDate());
      
      // Check if current forecast month is >= state pension age
      if (date >= statePensionAge) {
        rowData.push(currentStatePensionValue);
      } else {
        rowData.push(0); // No state pension payment yet
      }
    }
    
    forecastData.push(rowData);
    
    // Move to next month
    date.setMonth(date.getMonth() + 1);
    monthIndex++;
  }
  
  // Get or create Forecast sheet
  var forecastSheet = ss.getSheetByName('Forecast');
  if (!forecastSheet) {
    forecastSheet = ss.insertSheet('Forecast');
  }
  
  // Set headers
  forecastSheet.getRange('A1').setValue('Month');
  forecastSheet.getRange('B1').setValue('Stocks & Shares Tax Free');
  forecastSheet.getRange('C1').setValue('Stocks & Shares Taxable');
  forecastSheet.getRange('D1').setValue('Stocks & Shares Total');
  forecastSheet.getRange('E1').setValue('Cash');
  
  // Add pension column headers
  for (var p = 0; p < pensionData.length; p++) {
    var colIndex = 6 + p; // Start from column F (6)
    forecastSheet.getRange(1, colIndex).setValue(pensionData[p].title);
  }
  
  // Add Occasional Income column header
  var occasionalIncomeColIndex = 6 + pensionData.length;
  forecastSheet.getRange(1, occasionalIncomeColIndex).setValue('Occasional Income');
  
  // Add State Pension column headers for each person
  for (var sp = 0; sp < peopleData.length; sp++) {
    var statePensionColIndex = 7 + pensionData.length + sp;
    var headerTitle = peopleData[sp].name + ' State Pension';
    forecastSheet.getRange(1, statePensionColIndex).setValue(headerTitle);
  }
  
  // Format header row as title (including all columns)
  var totalColumns = 6 + pensionData.length + peopleData.length; // +1 for occasional income + state pensions
  var headerRange = forecastSheet.getRange(1, 1, 1, totalColumns);
  headerRange.setFontWeight('bold');
  headerRange.setFontSize(12);
  headerRange.setBackground('#4a90e2');
  headerRange.setFontColor('white');
  headerRange.setHorizontalAlignment('center');
  headerRange.setWrap(true);
  
  // Set column widths for better formatting
  forecastSheet.setColumnWidth(1, 150); // Month column
  forecastSheet.setColumnWidth(2, 200); // Stocks & Shares Tax Free column  
  forecastSheet.setColumnWidth(3, 200); // Stocks & Shares Taxable column
  forecastSheet.setColumnWidth(4, 200); // Stocks & Shares Total column
  forecastSheet.setColumnWidth(5, 150); // Cash column
  
  // Set column widths for pension columns
  for (var p = 0; p < pensionData.length; p++) {
    var colIndex = 6 + p;
    forecastSheet.setColumnWidth(colIndex, 150); // Pension columns
  }
  
  // Set column width for Occasional Income column
  forecastSheet.setColumnWidth(occasionalIncomeColIndex, 150);
  
  // Set column widths for State Pension columns
  for (var sp = 0; sp < peopleData.length; sp++) {
    var statePensionColIndex = 7 + pensionData.length + sp;
    forecastSheet.setColumnWidth(statePensionColIndex, 150); // State Pension columns
  }
  
  // Freeze row 1 and column 1
  forecastSheet.setFrozenRows(1);
  forecastSheet.setFrozenColumns(1);
  
  // Clear existing data (except headers)
  var totalColumns = 6 + pensionData.length + peopleData.length; // +1 for occasional income + state pensions
  if (forecastSheet.getLastRow() > 1) {
    forecastSheet.getRange(2, 1, forecastSheet.getLastRow() - 1, totalColumns).clearContent();
  }
  
  // Write forecast data starting from row 2
  if (forecastData.length > 0) {
    forecastSheet.getRange(2, 1, forecastData.length, totalColumns).setValues(forecastData);
    
    // Format currency columns (columns B, C, D, E, all pension columns, occasional income, and state pensions)
    var currencyColumns = 5 + pensionData.length + peopleData.length; // B, C, D, E, pension columns + occasional income + state pensions
    forecastSheet.getRange(2, 2, forecastData.length, currencyColumns).setNumberFormat('£#,##0.00');
    
    // Center align all data columns from B onwards
    forecastSheet.getRange(2, 2, forecastData.length, currencyColumns).setHorizontalAlignment('center');
  }
  
  // Show completion message
  var monthNames = ["January", "February", "March", "April", "May", "June",
                   "July", "August", "September", "October", "November", "December"];
  SpreadsheetApp.getUi().alert('Forecast plan created successfully!\n\nGenerated ' + forecastData.length + ' months from ' + 
                              monthNames[startDate.getMonth()] + ' ' + startDate.getFullYear() + 
                              ' to ' + monthNames[endDate.getMonth()] + ' ' + endDate.getFullYear() + 
                              '\n\nStarting values:\nTax Free: £' + totalTaxFree.toFixed(2) + 
                              '\nTaxable: £' + totalTaxable.toFixed(2) + 
                              '\nCash: £' + totalCash.toFixed(2) + 
                              '\n\nGrowth rates:\nStocks & Shares: ' + (monthlyStocksGrowthRate * 100).toFixed(3) + '% monthly' +
                              '\nCash: ' + (monthlyCashGrowthRate * 100).toFixed(3) + '% monthly');
}