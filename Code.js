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
  
  // Get people data from People sheet (A = name, B = birth date, E = last month of income)
  var peopleData = [];
  var peopleNames = []; // Track names to prevent duplicates
  var latestLastMonthOfIncome = null; // Track the latest "Last Month of Income" date
  var peopleLastRow = peopleSheet.getLastRow();
  if (peopleLastRow >= 2) {
    var peopleRange = peopleSheet.getRange('A2:E' + peopleLastRow).getValues();
    for (var i = 0; i < peopleRange.length; i++) {
      var name = peopleRange[i][0];
      var birthDate = peopleRange[i][1];
      var lastMonthOfIncome = peopleRange[i][4]; // Column E is index 4
      
      if (name && name.toString().trim() !== '' && birthDate) {
        var cleanName = name.toString().trim();
        // Check if name already exists to prevent duplicates
        if (peopleNames.indexOf(cleanName) === -1) {
          peopleData.push({
            name: cleanName,
            birthDate: new Date(birthDate),
            lastMonthOfIncome: lastMonthOfIncome ? new Date(lastMonthOfIncome) : null
          });
          peopleNames.push(cleanName);
          
          // Track the latest "Last Month of Income" date
          if (lastMonthOfIncome) {
            var incomeDate = new Date(lastMonthOfIncome);
            if (!latestLastMonthOfIncome || incomeDate > latestLastMonthOfIncome) {
              latestLastMonthOfIncome = incomeDate;
            }
          }
        }
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
  var monthlyInflationRate = annualInflationRate / 12; // Convert annual to monthly
  
  
  // Get Stocks and Shares sheet to find current values
  var stocksSharesSheet = ss.getSheetByName('Stocks and Shares');
  if (!stocksSharesSheet) {
    SpreadsheetApp.getUi().alert('Stocks and Shares sheet not found!');
    return;
  }
  
  // Get current values from Stocks and Shares table - now including columns D, E and F
  var stocksData = stocksSharesSheet.getRange('A2:F' + stocksSharesSheet.getLastRow()).getValues();
  
  // Calculate total current tax-free and taxable values, and collect contribution data
  var totalTaxFree = 0;
  var totalTaxable = 0;
  var contributionData = [];
  
  for (var i = 0; i < stocksData.length; i++) {
    var title = stocksData[i][0];
    var currentValue = stocksData[i][1] || 0;  // Column B
    var taxFreeValue = stocksData[i][2] || 0;  // Column C - Tax Free
    var taxableValue = stocksData[i][3] || 0;  // Column D - Taxable
    var monthlyContribution = stocksData[i][4] || 0;  // Column E
    var lastMonth = stocksData[i][5];  // Column F
    
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
  
  // Get Retirement Income sheet to find Annual Income property
  var retirementIncomeSheet = ss.getSheetByName('Retirement Income');
  var annualIncomeValue = 0;
  if (retirementIncomeSheet) {
    // Search for Annual Income property in Retirement Income sheet
    var retirementIncomeData = retirementIncomeSheet.getRange('A:B').getValues();
    for (var i = 0; i < retirementIncomeData.length; i++) {
      if (retirementIncomeData[i][0]) {
        var propertyName = retirementIncomeData[i][0].toString().toLowerCase();
        if (propertyName.includes('annual income')) {
          annualIncomeValue = retirementIncomeData[i][1] || 0;
          break;
        }
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
  
  // Create array of forecast data with just months
  var forecastData = [];
  var date = new Date(startDate);
  
  while (date <= endDate) {
    // Format as "Month Year" (e.g., "August 2025")
    var monthNames = ["January", "February", "March", "April", "May", "June",
                     "July", "August", "September", "October", "November", "December"];
    var monthStr = monthNames[date.getMonth()] + " " + date.getFullYear();
    
    // Create row with just month data - all other values will be empty
    var rowData = [monthStr];
    
    forecastData.push(rowData);
    
    // Move to next month
    date.setMonth(date.getMonth() + 1);
  }
  
  // Get or create Capital sheet (formerly Forecast)
  var forecastSheet = ss.getSheetByName('Capital');
  if (!forecastSheet) {
    forecastSheet = ss.insertSheet('Capital');
  }
  
  // Get or create Income sheet
  var incomeSheet = ss.getSheetByName('Income');
  if (!incomeSheet) {
    incomeSheet = ss.insertSheet('Income');
  }
  
  // Set headers for Capital sheet (only non-income columns)
  forecastSheet.getRange('A1').setValue('Month');
  forecastSheet.getRange('B1').setValue('Stocks & Shares Tax Free');
  forecastSheet.getRange('C1').setValue('Stocks & Shares Taxable');
  forecastSheet.getRange('D1').setValue('Cash');
  
  // Set headers for Income sheet
  incomeSheet.getRange('A1').setValue('Month');
  var incomeColIndex = 2;
  
  // Add From Assets column header to Income sheet
  incomeSheet.getRange(1, incomeColIndex).setValue('From Assets');
  incomeColIndex++;
  
  // Add pension column headers to Income sheet
  for (var p = 0; p < pensionData.length; p++) {
    incomeSheet.getRange(1, incomeColIndex).setValue(pensionData[p].title);
    incomeColIndex++;
  }
  
  // Add State Pension column headers for each person to Income sheet
  for (var sp = 0; sp < peopleData.length; sp++) {
    var headerTitle = peopleData[sp].name + ' State Pension';
    incomeSheet.getRange(1, incomeColIndex).setValue(headerTitle);
    incomeColIndex++;
  }
  
  // Format header row for Capital sheet (only 4 columns now)
  var forecastColumns = 4;
  var forecastHeaderRange = forecastSheet.getRange(1, 1, 1, forecastColumns);
  forecastHeaderRange.setFontWeight('bold');
  forecastHeaderRange.setFontSize(12);
  forecastHeaderRange.setBackground('#4a90e2');
  forecastHeaderRange.setFontColor('white');
  forecastHeaderRange.setHorizontalAlignment('center');
  forecastHeaderRange.setWrap(true);
  
  // Format header row for Income sheet 
  var incomeColumns = 1 + 1 + pensionData.length + peopleData.length; // Month + From Assets + Pensions + State Pensions
  var incomeHeaderRange = incomeSheet.getRange(1, 1, 1, incomeColumns);
  incomeHeaderRange.setFontWeight('bold');
  incomeHeaderRange.setFontSize(12);
  incomeHeaderRange.setBackground('#4a90e2');
  incomeHeaderRange.setFontColor('white');
  incomeHeaderRange.setHorizontalAlignment('center');
  incomeHeaderRange.setWrap(true);
  
  // Set column widths for Capital sheet
  forecastSheet.setColumnWidth(1, 150); // Month column
  forecastSheet.setColumnWidth(2, 200); // Stocks & Shares Tax Free column  
  forecastSheet.setColumnWidth(3, 200); // Stocks & Shares Taxable column
  forecastSheet.setColumnWidth(4, 150); // Cash column (now includes occasional income)
  
  // Set column widths for Income sheet
  incomeSheet.setColumnWidth(1, 150); // Month column
  incomeSheet.setColumnWidth(2, 150); // From Assets column
  var incomeColIndex = 3;
  
  // Set column widths for pension columns in Income sheet
  for (var p = 0; p < pensionData.length; p++) {
    incomeSheet.setColumnWidth(incomeColIndex, 150); // Pension columns
    incomeColIndex++;
  }
  
  // Set column widths for State Pension columns in Income sheet
  for (var sp = 0; sp < peopleData.length; sp++) {
    incomeSheet.setColumnWidth(incomeColIndex, 150); // State Pension columns
    incomeColIndex++;
  }
  
  // Freeze row 1 and column 1 for both sheets
  forecastSheet.setFrozenRows(1);
  forecastSheet.setFrozenColumns(1);
  incomeSheet.setFrozenRows(1);
  incomeSheet.setFrozenColumns(1);
  
  // Clear existing data (except headers) for both sheets
  var forecastColumns = 4;
  var incomeColumns = 1 + 1 + pensionData.length + peopleData.length; // Month + From Assets + Pensions + State Pensions
  
  if (forecastSheet.getLastRow() > 1) {
    var forecastClearRange = forecastSheet.getRange(2, 1, forecastSheet.getLastRow() - 1, forecastColumns);
    forecastClearRange.clearContent();
    forecastClearRange.clearNote();
  }
  
  if (incomeSheet.getLastRow() > 1) {
    var incomeClearRange = incomeSheet.getRange(2, 1, incomeSheet.getLastRow() - 1, incomeColumns);
    incomeClearRange.clearContent();
    incomeClearRange.clearNote();
  }
  
  // Write forecast data starting from row 2
  if (forecastData.length > 0) {
    // Prepare separate data arrays for Capital and Income sheets - just months
    var forecastSheetData = [];
    var incomeSheetData = [];
    
    for (var i = 0; i < forecastData.length; i++) {
      var row = forecastData[i];
      
      // Capital sheet data: Month only (other columns will be empty)
      var forecastRow = [row[0]];
      forecastSheetData.push(forecastRow);
      
      // Income sheet data: Month only (other columns will be empty)
      var incomeRow = [row[0]];
      incomeSheetData.push(incomeRow);
    }
    
    // Write data to Capital sheet (only month column)
    forecastSheet.getRange(2, 1, forecastSheetData.length, 1).setValues(forecastSheetData);
    
    // Write data to Income sheet (only month column)
    incomeSheet.getRange(2, 1, incomeSheetData.length, 1).setValues(incomeSheetData);
    
    // Set Final Salary Pension values based on First Month matching
    for (var p = 0; p < pensionData.length; p++) {
      var pension = pensionData[p];
      var pensionColumnIndex = 3 + p; // Pension columns start at column 3 (after Month and From Assets)
      
      var pensionStartFound = false;
      var currentPensionValue = pension.monthlyIncome;
      
      // Process all forecast months for this pension
      for (var i = 0; i < forecastData.length; i++) {
        var forecastDate = new Date(startDate);
        forecastDate.setMonth(startDate.getMonth() + i);
        
        // Check if forecast month/year is before pension first month/year
        if (forecastDate.getFullYear() < pension.firstMonth.getFullYear() || 
            (forecastDate.getFullYear() === pension.firstMonth.getFullYear() && 
             forecastDate.getMonth() < pension.firstMonth.getMonth())) {
          // Set to zero for months before pension starts
          incomeSheet.getRange(i + 2, pensionColumnIndex).setValue(0);
        }
        // Check if forecast month/year matches pension first month/year
        else if (forecastDate.getFullYear() === pension.firstMonth.getFullYear() && 
                 forecastDate.getMonth() === pension.firstMonth.getMonth()) {
          // First month - use original value
          incomeSheet.getRange(i + 2, pensionColumnIndex).setValue(currentPensionValue);
          pensionStartFound = true;
        }
        // For months after pension starts
        else if (pensionStartFound) {
          // Apply monthly inflation growth
          currentPensionValue = Math.round((currentPensionValue * (1 + monthlyInflationRate)) * 100) / 100;
          incomeSheet.getRange(i + 2, pensionColumnIndex).setValue(currentPensionValue);
        }
      }
    }
    
    // Format pension columns as currency on Income sheet
    for (var p = 0; p < pensionData.length; p++) {
      var pensionColumnIndex = 3 + p;
      incomeSheet.getRange(2, pensionColumnIndex, forecastData.length, 1).setNumberFormat('£#,##0.00');
      incomeSheet.getRange(2, pensionColumnIndex, forecastData.length, 1).setHorizontalAlignment('center');
    }
    
    // Set Stocks & Shares Tax Free for first month using totalTaxFree value
    forecastSheet.getRange(2, 2).setValue(totalTaxFree);
    
    // Set Stocks & Shares Taxable for first month using totalTaxable value
    forecastSheet.getRange(2, 3).setValue(totalTaxable);
    
    // Set Cash for first month using totalCash value
    forecastSheet.getRange(2, 4).setValue(totalCash);
    
    // Format the cells as currency
    forecastSheet.getRange(2, 2).setNumberFormat('£#,##0.00');
    forecastSheet.getRange(2, 2).setHorizontalAlignment('center');
    forecastSheet.getRange(2, 3).setNumberFormat('£#,##0.00');
    forecastSheet.getRange(2, 3).setHorizontalAlignment('center');
    forecastSheet.getRange(2, 4).setNumberFormat('£#,##0.00');
    forecastSheet.getRange(2, 4).setHorizontalAlignment('center');
  }
  
  // Show completion message
  var monthNames = ["January", "February", "March", "April", "May", "June",
                   "July", "August", "September", "October", "November", "December"];
  SpreadsheetApp.getUi().alert('Forecast plan created successfully!\n\nGenerated ' + forecastData.length + ' months from ' + 
                              monthNames[startDate.getMonth()] + ' ' + startDate.getFullYear() + 
                              ' to ' + monthNames[endDate.getMonth()] + ' ' + endDate.getFullYear() + 
                              '\n\nCapital and Income sheets have been created with month rows and column headers.');
}