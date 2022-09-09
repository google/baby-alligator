/***********************************************************************
Copyright 2022 Google LLC
Licensed under the Apache License, Version 2.0 (the "License");
you may not use this file except in compliance with the License.
You may obtain a copy of the License at
    https://www.apache.org/licenses/LICENSE-2.0
Unless required by applicable law or agreed to in writing, software
distributed under the License is distributed on an "AS IS" BASIS,
WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
See the License for the specific language governing permissions and
limitations under the License.
************************************************************************/

/**
 * Sets up and formats the needed sheets in the Spreadsheet.
 * @private
 */
/**
 * Sets up and formats the needed sheets in the Spreadsheet.
 * @private
 */
function initSpreadsheet_() {
  let sheetIndex = 1;

  if (!instructionSheet) {
    let row = 0;
    doc.insertSheet(INSTRUCTION_SHEET_NAME, sheetIndex);
    instructionSheet = doc.getSheetByName(INSTRUCTION_SHEET_NAME);
    instructionSheet.setTabColor('red');
    instructionSheet.setRowHeight(1, 60);
    instructionSheet.getRange(++row, 1)
        .setValue('Instructions')
        .setBackground('#FFA500')
        .setFontWeight('bold')
        .setFontSize(18)
        .setHorizontalAlignment('center')
        .setVerticalAlignment('middle');

    instructionSheet.getRange(++row, 1).setValue(
        'In the top menu click on "Collect GBP Data" and select "Collect accounts". This action will retrieve all your accounts and fill the Accounts sheet.');
    row++;

    instructionSheet.getRange(++row, 1).setFontWeight('bold').setValue(
        'First run:');

    instructionSheet.getRange(++row, 1).setValue(
        'Open now the "Configuration" sheet and fill as follows:');
    instructionSheet.getRange(++row, 1).setValue(
        '   "Weeks of retention": default to 52, you can edit this value to increase or reduce the retention period.');
    instructionSheet.getRange(++row, 1).setValue(
        '   "Account name": Select from the drop down menu the account you want to include. The value for "AccountGroup" will be automatically filled.');
    instructionSheet.getRange(++row, 1).setValue(
        '    "FilterStatus": Select from the drop down menu the status of the locations to retrieve');
    instructionSheet.getRange(++row, 1).setValue(
        '    "FilterRegion": Manually insert the country code (1 per row) of the locations to retrieve. In case you need to retrieve locations from several countries create 1 entry per country');
    row++;

    instructionSheet.getRange(++row, 1).setFontWeight('bold').setValue(
        'NOTE: Due to size limitations in the spreadsheet be careful while editing "Weeks of retention". If you have 900+ locations we recommend not to increase it and, if possible, to reduce it');
    row++;

    instructionSheet.getRange(++row, 1).setValue(
        'Now click again on "Collect GBP Data" and select "Collect locations and insights". This will start the process to retrieve all the locations and insights for your filtered accounts within the defined retention period. This process could last up to a few days.');

    instructionSheet.getRange(++row, 1).setValue(
        'Once completed a weekly trigger will update the insights with new data while erasing the oldest entries.');
    row++;

    instructionSheet.getRange(++row, 1).setFontWeight('bold').setValue(
        'NOTE: Do NOT edit the configuration without clicking on "Reset and use new config" since your configuration will be updated only for new insights entries and your data will be incoherent.');
    row++;

    instructionSheet.getRange(++row, 1).setFontWeight('bold').setValue(
        'Change your configuration');
    instructionSheet.getRange(++row, 1).setValue(
        'In case you want to update your configuration to perform a different analysis, you have to:');
    instructionSheet.getRange(++row, 1).setValue(
        '   Update your "Configuration" sheet');
    instructionSheet.getRange(++row, 1).setValue(
        '   Click on "Reset and use new config" and confirm your choice. ');
    instructionSheet.getRange(++row, 1).setValue(
        'You will notice that all your previous data will be wiped out and the new data will be retrieved.');

    instructionSheet.autoResizeColumns(1, 3);
  }
  sheetIndex++;

  if (!accountSheet) {
    doc.insertSheet(ACC_SHEET_NAME, sheetIndex);
    accountSheet = doc.getSheetByName(ACC_SHEET_NAME);
    accountSheet.setTabColor('yellow');
    ACCOUNT_FIELDS_MAP.forEach(
        (pos, value) =>
            accountSheet.getRange(1, pos).setValue(value).setFontWeight(
                'bold'));
    accountSheet.getRange(1, 1, 1, ACCOUNT_FIELDS_MAP.size)
        .setBackground('#F4C7C3');
    accountSheet.autoResizeColumns(1, ACCOUNT_FIELDS_MAP.size);
  }
  sheetIndex++;

  if (!configSheet) {
    const HEADER_ROW = 4;
    doc.insertSheet(CONFIG_SHEET_NAME, sheetIndex);
    configSheet = doc.getSheetByName(CONFIG_SHEET_NAME);
    configSheet.setTabColor('orange');
    CONFIG_FIELDS_MAP.forEach(
        (pos, value) => configSheet.getRange(HEADER_ROW, pos)
                            .setValue(value)
                            .setFontWeight('bold'));
    configSheet.getRange(HEADER_ROW, 1, 1, CONFIG_FIELDS_MAP.size)
        .setBackground('#FF6D01');

    const rule = SpreadsheetApp.newDataValidation()
                     .requireNumberGreaterThan(0)
                     .setHelpText('Minimum retention period is 1 week')
                     .build();
    configSheet.getRange(1, 1).setDataValidation(rule).setValue(
        DEFAULT_RETENTION_PERIOD_IN_WEEKS);
    configSheet.getRange(1, 2)
        .setValue('Weeks of retention')
        .setFontWeight('bold');
    configSheet.autoResizeColumns(HEADER_ROW, CONFIG_FIELDS_MAP.size);
  }
  sheetIndex++;

  if (!locSheet) {
    doc.insertSheet(LOC_SHEET_NAME, sheetIndex);
    locSheet = doc.getSheetByName(LOC_SHEET_NAME);
    locSheet.setTabColor('blue');
    LOCATION_FIELDS_MAP.forEach(
        (pos, value) =>
            locSheet.getRange(1, pos).setValue(value).setFontWeight('bold'));
    locSheet.getRange(1, 1, 1, LOCATION_FIELDS_MAP.size)
        .setBackground('#9AC0CD');
    locSheet.autoResizeColumns(1, LOCATION_FIELDS_MAP.size);
  }
  sheetIndex++;

  if (!insSheet) {
    doc.insertSheet(INS_SHEET_NAME, sheetIndex);
    insSheet = doc.getSheetByName(INS_SHEET_NAME);
    insSheet.setTabColor('green');
    INSIGHT_FIELDS_MAP.forEach(
        (pos, value) =>
            insSheet.getRange(1, pos).setValue(value).setFontWeight('bold'));
    insSheet.getRange(1, 1, 1, INSIGHT_FIELDS_MAP.size)
        .setBackground('#AFDD9C');
    insSheet.autoResizeColumns(1, INSIGHT_FIELDS_MAP.size);
  }
  sheetIndex++;

  if (!logSheet) {
    doc.insertSheet(LOG_SHEET_NAME, sheetIndex);
    logSheet = doc.getSheetByName(LOG_SHEET_NAME);
    logSheet.setTabColor('black');
    LOG_FIELDS_MAP.forEach(
        (pos, value) =>
            logSheet.getRange(1, pos).setValue(value).setFontWeight('bold'));
    logSheet.getRange(1, 1, 1, LOG_FIELDS_MAP.size).setBackground('#F5F5AE');
    logSheet.autoResizeColumns(1, LOG_FIELDS_MAP.size);
  }
  SpreadsheetApp.flush();
}

/**
 * Initialization function to structure the spreadsheet, if needed.
 * @private
 */
function init_() {
  if (!accountSheet || !configSheet || !locSheet || !insSheet || !logSheet ||
      !instructionSheet) {
    // We need setup and format the spreadsheet
    initSpreadsheet_();
  }
  doc.setActiveSheet(configSheet);
}
# baby-alligator
# baby-alligator
