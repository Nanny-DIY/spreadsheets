function generatePaystub() {
  var SETTING_SHEET_NAME = 'Payroll settings and hourly tracking';
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getActiveSheet();
  if (!sheet || !sheet.getName() || sheet.getName() !== SETTING_SHEET_NAME) {
    SpreadsheetApp.getUi().alert('To generate a paystub, you must select a cell in the row representing the pay period.' +
                                'To do this, you must be on the sheet called \'' + SETTING_SHEET_NAME + '\'. But you are currently ' +
                                'on the sheet called \'' + sheet.getName() + '\'.');
    return;
  }
  
  var activeCell = sheet.getActiveCell();
  if (!activeCell) {
    SpreadsheetApp.getUi().alert('To generate a paystub, you must select a cell in the row representing the pay period.' +
                                 'No cell is active right now, so you must select one.');
    return;
  }
    
  // Get all relevant data
  var row = activeCell.getRow();
  var startDate = Date.parse(sheet.getRange(row, 2).getValue());
  var endDate = Date.parse(sheet.getRange(row, 3).getValue());
  var normalHoursWorked = sheet.getRange(row, 4).getValue();
  var paidSickTime = sheet.getRange(row, 5).getValue();
  var paidTimeOff = sheet.getRange(row, 6).getValue();
  var sumNormalHoursWorked = sheet.getRange(row, 7).getValue();
  var sumPaidSickTime = sheet.getRange(row, 8).getValue();
  var sumPaidTimeOff = sheet.getRange(row, 9).getValue();
  
  if (!startDate || !endDate || isNaN(startDate) || isNaN(endDate)) {
    SpreadsheetApp.getUi().alert('To generate a paystub, you must select a cell in the row representing the pay period. ' +
                                'These rows are in the \'Time tracking\' section of this spreadsheet.');
    return;
  }
  
  var startDateDate = new Date(startDate);
  var endDateDate = new Date(endDate);
  
  var nannyName = sheet.getRange('D6').getValue();
  var nannyAddress = sheet.getRange('D7').getValue();
  var nannyCity = sheet.getRange('D8').getValue();
  
  var employerName = sheet.getRange('D11').getValue();
  var employerAddress = sheet.getRange('D12').getValue();
  var employerCity = sheet.getRange('D13').getValue();
  var employerChild = sheet.getRange('D14').getValue();
    
  var medicareEmployeeRate = sheet.getRange('D17').getValue();
  var socialSecurityEmployeeRate = sheet.getRange('D18').getValue();
  var waStateFamilyEmployeeRate = sheet.getRange('D19').getValue();
  var medicareEmployerRate = sheet.getRange('D22').getValue();
  var socialSecurityEmployerRate = sheet.getRange('D23').getValue();
  var federalUnemploymentEmployerRate = sheet.getRange('D24').getValue();
  var maxFederalUnemploymentAmount = sheet.getRange('F24').getValue();
  var waStateUnemploymentEmployerRate = sheet.getRange('D25').getValue();
  var waStateFamilyEmployerRate = sheet.getRange('D26').getValue();
  var hourlyPay = sheet.getRange('D29').getValue();
  var generateSeattlePaidSickTimeFields = sheet.getRange('D32').isChecked();
  var rolloverSeattlePaidSickTimeHours = sheet.getRange('D33').getValue();
  
  // Create a new sheet  
  var paystubSheetName = 'Paystub ' + Number(endDateDate.getMonth()+1) + '/' + endDateDate.getDate() + '/' + endDateDate.getYear();
  var existingSheet = spreadsheet.getSheetByName(paystubSheetName);
  if (existingSheet) {
    SpreadsheetApp.getUi().alert('A paystub for the date has already been generated. Generating another one.');
    paystubSheetName = 'Paystub ' + Number(endDateDate.getMonth()+1) + '/' + endDateDate.getDate() + '/' + endDateDate.getYear() + ' ' + Date.now();
  }
  paystubSheet = spreadsheet.insertSheet(paystubSheetName);
  
  // Fill the data in the new paystub sheet
  paystubSheet.getRange('A1').setValue(nannyName);
  paystubSheet.getRange('A1').setFontWeight('bold');
  paystubSheet.getRange('A2').setValue(nannyAddress);
  paystubSheet.getRange('A3').setValue(nannyCity);
  
  paystubSheet.getRange('E1').setValue('Payroll date');
  paystubSheet.getRange('E1').setHorizontalAlignment('right');
  paystubSheet.getRange('F1').setValue(Number(endDateDate.getMonth()+1) + '/' + endDateDate.getDate() + '/' + endDateDate.getYear());
  paystubSheet.getRange('F1').setHorizontalAlignment('left');
  paystubSheet.getRange('E2').setValue('YTD hours on previous paycheck');
  paystubSheet.getRange('E2').setHorizontalAlignment('right');
  paystubSheet.getRange('F2').setValue(sumNormalHoursWorked + sumPaidSickTime + sumPaidTimeOff - (normalHoursWorked + paidSickTime + paidTimeOff));
  paystubSheet.getRange('F2').setHorizontalAlignment('left');
  paystubSheet.getRange('E3').setValue('Service dates');
  paystubSheet.getRange('E3').setHorizontalAlignment('right');
  paystubSheet.getRange('F3').setValue(Number(startDateDate.getMonth()+1) + '/' + startDateDate.getDate() + '/' + startDateDate.getYear() + '-' +
    Number(endDateDate.getMonth()+1) + '/' + endDateDate.getDate() + '/' + endDateDate.getYear());
  paystubSheet.getRange('F3').setHorizontalAlignment('left');
  
  paystubSheet.getRange('A5').setValue('Childcare for ' + employerChild);
  paystubSheet.getRange('A5').setFontWeight('bold');
  paystubSheet.getRange('E5').setValue('This Pay Period');
  paystubSheet.getRange('E5').setFontWeight('bold');
  paystubSheet.getRange('G5').setValue('Year To Date');
  paystubSheet.getRange('G5').setFontWeight('bold');
  paystubSheet.getRange('C6').setValue('Rate');
  paystubSheet.getRange('E6').setValue('Hours');
  paystubSheet.getRange('F6').setValue('Total');
  paystubSheet.getRange('G6').setValue('Hours');
  paystubSheet.getRange('H6').setValue('Total');
  
  paystubSheet.getRange('A7').setValue('Employee Wages');
  paystubSheet.getRange('B8').setValue('Hourly');
  paystubSheet.getRange('C8').setValue(hourlyPay);  
  var hoursThisPayPeriod = normalHoursWorked + paidSickTime + paidTimeOff;
  var overallHours = sumNormalHoursWorked + sumPaidSickTime + sumPaidTimeOff;
  paystubSheet.getRange('E8').setValue(hoursThisPayPeriod);
  paystubSheet.getRange('F8').setValue('=C8*E8');
  paystubSheet.getRange('G8').setValue(overallHours);
  paystubSheet.getRange('H8').setValue('=C8*G8');
  
  paystubSheet.getRange('B9').setValue('Total Gross Pay');
  paystubSheet.getRange('F9').setValue('=F8');
  paystubSheet.getRange('H9').setValue('=H8');
  var currentGrossPay = paystubSheet.getRange('F9').getValue();
  var overallGrossPay = paystubSheet.getRange('H9').getValue();
  
  paystubSheet.getRange('A10').setValue('Employee Taxes and Adjustments');
  paystubSheet.getRange('B11').setValue('Medicare Employee');
  paystubSheet.getRange('C11').setValue(medicareEmployeeRate);
  paystubSheet.getRange('F11').setValue('=F8*C11');
  paystubSheet.getRange('H11').setValue('=H8*C11');
  paystubSheet.getRange('B12').setValue('Social Security Employee');
  paystubSheet.getRange('C12').setValue(socialSecurityEmployeeRate);
  paystubSheet.getRange('F12').setValue('=F8*C12');
  paystubSheet.getRange('H12').setValue('=H8*C12');
  paystubSheet.getRange('B13').setValue('WA Paid Family and Medical Leave');
  paystubSheet.getRange('C13').setValue(waStateFamilyEmployeeRate);
  paystubSheet.getRange('F13').setValue('=F8*C13');
  paystubSheet.getRange('H13').setValue('=H8*C13');
  paystubSheet.getRange('B14').setValue('Total Taxes Withheld');
  paystubSheet.getRange('F14').setValue('=F8*C13');
  paystubSheet.getRange('H14').setValue('=H8*C13');
  
  paystubSheet.getRange('A15').setValue('Net Pay');
  paystubSheet.getRange('F15').setValue('=F8-F14');
  paystubSheet.getRange('F15').setFontWeight('bold');
  paystubSheet.getRange('H15').setValue('=H8-H14');
  
  paystubSheet.getRange('A16').setValue('Employer Taxes and Contributions');
  paystubSheet.getRange('B17').setValue('Medicare Employer');
  paystubSheet.getRange('C17').setValue(medicareEmployerRate);
  paystubSheet.getRange('F17').setValue('=F8*C17');
  
  paystubSheet.getRange('H17').setValue('=H8*C17');
  paystubSheet.getRange('B18').setValue('Social Security Employer');
  paystubSheet.getRange('C18').setValue(socialSecurityEmployerRate);
  paystubSheet.getRange('F18').setValue('=F8*C18');
  paystubSheet.getRange('H18').setValue('=H8*C18');
  paystubSheet.getRange('B19').setValue('Federal Unemployment');
  paystubSheet.getRange('C19').setValue(federalUnemploymentEmployerRate);
  if (overallGrossPay <= maxFederalUnemploymentAmount) {
    paystubSheet.getRange('F19').setValue('=F8*C19');
    paystubSheet.getRange('H19').setValue('=H8*C19');
  } else {
    if (overallGrossPay - currentGrossPay < maxFederalUnemploymentAmount) {
      var fractionalAmountToBeTaxed = maxFederalUnemploymentAmount - (overallGrossPay - currentGrossPay);
      paystubSheet.getRange('F19').setValue('=C19*' + fractionalAmountToBeTaxed.toString());
    } else {
      paystubSheet.getRange('F19').setValue('=F8*C19');
    }
    paystubSheet.getRange('H19').setValue('=C19*' + maxFederalUnemploymentAmount.toString());
  }
  paystubSheet.getRange('B20').setValue('WA State Unemployment');
  paystubSheet.getRange('C20').setValue(waStateUnemploymentEmployerRate); 
  paystubSheet.getRange('F20').setValue('=F8*C20');
  paystubSheet.getRange('H20').setValue('=H8*C20');
  paystubSheet.getRange('B21').setValue('WA Paid Family and Medical Leave\nEmployer');
  paystubSheet.getRange('C21').setValue(waStateFamilyEmployerRate);
  paystubSheet.getRange('F21').setValue('=F8*C21');
  paystubSheet.getRange('H21').setValue('=H8*C21');
  
  // Household employer
  paystubSheet.getRange('A26').setValue('Household employer');
  paystubSheet.getRange('A26').setFontWeight('bold');
  paystubSheet.getRange('A27').setValue('=\'' + SETTING_SHEET_NAME + '\'!D11');
  paystubSheet.getRange('A28').setValue('=\'' + SETTING_SHEET_NAME + '\'!D12');
  paystubSheet.getRange('A29').setValue('=\'' + SETTING_SHEET_NAME + '\'!D13');
  
  // draw the border
  paystubSheet.getRange('A5:H21').setBorder(true, true, true, true, null, null);
  paystubSheet.getRange('E5:F21').setBorder(true, true, true, true, null, null);
  paystubSheet.getRange('A5:H6').setBorder(true, true, true, true, null, null);  
  
  // set number formats
  paystubSheet.getRange('C8').setNumberFormat('$0.00');
  paystubSheet.getRange('C11:C21').setNumberFormat('0.0000%');
  paystubSheet.getRange('F8:F21').setNumberFormat('$0.00');
  paystubSheet.getRange('H8:H21').setNumberFormat('$0.00');
  
  // resize column
  paystubSheet.autoResizeColumn(2);
  
  paystubSheet.getRange('A23').setValue('Paid vacation this period (hours):');
  paystubSheet.getRange('C23').setValue(paidTimeOff);
  paystubSheet.getRange('A24').setValue('Paid vacation this year (hours):');
  paystubSheet.getRange('C24').setValue(sumPaidTimeOff);
  
  // if Seattle paid sick and safe time is selected, populate the required fields
  if (generateSeattlePaidSickTimeFields) {
    paystubSheet.getRange('E23').setValue('Paid Sick and Safe Time (PSST) available (hours):');
    paystubSheet.getRange('H23').setValue(rolloverSeattlePaidSickTimeHours + sumNormalHoursWorked / 40 - sumPaidSickTime);
    paystubSheet.getRange('E24').setValue('PSST accrued since last notice (hours):');
    paystubSheet.getRange('H24').setValue(normalHoursWorked / 40);
    paystubSheet.getRange('E25').setValue('PSST reduced (hours):');
    paystubSheet.getRange('H25').setValue(paidSickTime / 40);
    paystubSheet.getRange('E26').setValue('PSST used this year (hours):');  
    paystubSheet.getRange('H26').setValue(sumPaidSickTime);
  }
};
