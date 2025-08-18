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
      rowData = [monthStr, currentTaxFree, currentTaxable, currentCash];
    } else {
      // Apply monthly growth to previous month's values
      currentTaxFree = Math.round((currentTaxFree * (1 + monthlyStocksGrowthRate)) * 100) / 100;
      currentTaxable = Math.round((currentTaxable * (1 + monthlyStocksGrowthRate)) * 100) / 100;
      currentCash = Math.round((currentCash * (1 + monthlyCashGrowthRate)) * 100) / 100;
      
      // Add monthly contributions after growth
      currentTaxFree += monthlyTaxFreeContributions;
      currentTaxable += monthlyTaxableContributions;
      
      rowData = [monthStr, currentTaxFree, currentTaxable, currentCash];
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
    
    // Add occasional income to cash for this month
    var occasionalIncomeValue = 0;
    var occasionalIncomeTitles = [];
    for (var o = 0; o < occasionalIncomeData.length; o++) {
      var occasionalIncome = occasionalIncomeData[o];
      // Check if the month matches (year and month only, ignore day)
      if (date.getFullYear() === occasionalIncome.month.getFullYear() && 
          date.getMonth() === occasionalIncome.month.getMonth()) {
        occasionalIncomeValue += occasionalIncome.value;
        occasionalIncomeTitles.push(occasionalIncome.title);
      }
    }
    
    // Add occasional income to cash value
    var totalCashForMonth = Math.round((currentCash + occasionalIncomeValue) * 100) / 100;
    
    // Update rowData to include cash with occasional income
    rowData[3] = totalCashForMonth; // Cash is now in position 3 (column D)
    
    // Store occasional income info for this row to add comments later
    if (occasionalIncomeValue > 0) {
      rowData.occasionalIncomeComment = occasionalIncomeTitles.join(', ');
    }
    
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
    forecastSheet.getRange(2, 1, forecastSheet.getLastRow() - 1, forecastColumns).clearContent();
  }
  
  if (incomeSheet.getLastRow() > 1) {
    incomeSheet.getRange(2, 1, incomeSheet.getLastRow() - 1, incomeColumns).clearContent();
  }
  
  // Write forecast data starting from row 2
  if (forecastData.length > 0) {
    // Prepare separate data arrays for Forecast and Income sheets
    var forecastSheetData = [];
    var incomeSheetData = [];
    
    for (var i = 0; i < forecastData.length; i++) {
      var row = forecastData[i];
      
      // Capital sheet data: Month, Tax Free, Taxable, Cash (first 4 columns)
      var forecastRow = [row[0], row[1], row[2], row[3]];
      if (row.occasionalIncomeComment) {
        forecastRow.occasionalIncomeComment = row.occasionalIncomeComment;
      }
      forecastSheetData.push(forecastRow);
      
      // Income sheet data: Month + From Assets + pension columns + state pension columns
      var incomeRow = [row[0]]; // Start with month
      
      // Calculate From Assets value for this specific month
      // First calculate the underlying growth value (regardless of whether it will be shown)
      var underlyingFromAssetsValue;
      if (i === 0) {
        underlyingFromAssetsValue = annualIncomeValue; // First month equals exact Annual Income
      } else {
        // Calculate value based on first month plus accumulated monthly inflation
        underlyingFromAssetsValue = annualIncomeValue * Math.pow(1 + monthlyInflationRate, i);
        underlyingFromAssetsValue = Math.round(underlyingFromAssetsValue * 100) / 100; // Round to 2 decimal places
      }
      
      // Now check if this month should show zero due to Last Month of Income restriction
      var fromAssetsValue;
      if (latestLastMonthOfIncome) {
        // Compare year and month only (ignore day)
        var forecastYear = date.getFullYear();
        var forecastMonth = date.getMonth();
        var latestIncomeYear = latestLastMonthOfIncome.getFullYear();
        var latestIncomeMonth = latestLastMonthOfIncome.getMonth();
        
        // If forecast month/year is <= latest income month/year, force to zero
        if (forecastYear < latestIncomeYear || (forecastYear === latestIncomeYear && forecastMonth <= latestIncomeMonth)) {
          fromAssetsValue = 0; // Force to zero for months before or equal to latest Last Month of Income
        } else {
          fromAssetsValue = underlyingFromAssetsValue; // Use calculated value for months after
        }
      } else {
        // If no Last Month of Income data found, use calculated value for all months
        fromAssetsValue = underlyingFromAssetsValue;
      }
      
      incomeRow.push(fromAssetsValue);
      
      // Add pension values (columns 4+ in original data)
      for (var p = 0; p < pensionData.length; p++) {
        incomeRow.push(row[4 + p]);
      }
      
      // Add state pension values 
      for (var sp = 0; sp < peopleData.length; sp++) {
        incomeRow.push(row[4 + pensionData.length + sp]);
      }
      
      incomeSheetData.push(incomeRow);
    }
    
    // Write data to Capital sheet
    forecastSheet.getRange(2, 1, forecastSheetData.length, forecastColumns).setValues(forecastSheetData);
    
    // Write data to Income sheet
    incomeSheet.getRange(2, 1, incomeSheetData.length, incomeColumns).setValues(incomeSheetData);
    
    // Add comments to Cash column cells that have occasional income
    for (var i = 0; i < forecastSheetData.length; i++) {
      if (forecastSheetData[i].occasionalIncomeComment) {
        var cashCell = forecastSheet.getRange(i + 2, 4); // Row i+2 (since we start from row 2), column 4 (Cash column)
        cashCell.setNote(forecastSheetData[i].occasionalIncomeComment);
      }
    }
    
    // Format currency columns for Capital sheet (columns B, C, D)
    forecastSheet.getRange(2, 2, forecastSheetData.length, 3).setNumberFormat('£#,##0.00');
    forecastSheet.getRange(2, 2, forecastSheetData.length, 3).setHorizontalAlignment('center');
    
    // Format currency columns for Income sheet (all columns except Month)
    if (incomeColumns > 1) {
      incomeSheet.getRange(2, 2, incomeSheetData.length, incomeColumns - 1).setNumberFormat('£#,##0.00');
      incomeSheet.getRange(2, 2, incomeSheetData.length, incomeColumns - 1).setHorizontalAlignment('center');
    }
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