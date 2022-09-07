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
 * @fileoverview Utility functions used by the methods from the main file.
 */


/**
 * Calls a Google API using different methods and supporting pagination.
 * @param{string} url: the url to invoke
 * @param{string} methodType: the method of the HTTPS request
 * @param{?boolean} withPagination : should return an array of item(s) or just
 * the result as is
 * @param{?object} requestBody : The body of the request
 * @return {?object} Either the result of the HTTPS invocation as is or wrapped
 *     in an array.
 * @private
 */
function callApi_(url, methodType, withPagination, requestBody) {
  const headers = {
    'Content-Type': 'application/json',
    'Accept': 'application/json',
    'Authorization': `Bearer ${ScriptApp.getOAuthToken()}`
  };
  let options = {
    method: methodType,
    headers: headers,
    muteHttpExceptions: true
  };
  if (!!requestBody) {
    options.payload = JSON.stringify(requestBody);
  }
  let results = [];
  let result = {};
  try {
    // Do-while loop to support pagination.
    do {
      let callUrl = !!result && result['nextPageToken'] ?
          `${url}&pageToken=${result['nextPageToken']}` :
          url;
      let res = UrlFetchApp.fetch(callUrl, options);
      result = JSON.parse(res);
      if (withPagination) {
        results.push(result);
      } else {
        return result;
      }
    } while (result['nextPageToken']);
    return results;
  } catch (e) {
    customLog_('API call error: ' + e.toString());
    throw e;
  }
}

/**
 * Appends a new entry (with timestamp) in the Log sheet.
 * @param{string} message: Message to print
 * @public
 */
function customLog_(message) {
  const logRow = [
    Utilities.formatDate(new Date(), 'Europe/Rome', 'yyyy-MM-dd HH:mm:ss'),
    message
  ];
  SpreadsheetApp.getActiveSpreadsheet()
      .getSheetByName(LOG_SHEET_NAME)
      .appendRow(logRow);
}

/**
 * Removes older rows from the Log sheet (if necessary) until there are 1500
 * left.
 * @public
 */
function shortenLogSheet() {
  const MAX_ROW_SIZE = 1500;
  logSheet.setTabColor('black');
  LOG_FIELDS_MAP.forEach(
      (key, value) =>
          logSheet.getRange(1, key).setValue(value).setFontWeight('bold'));
  logSheet.getRange(1, 1, 1, 3).setBackground('#F5F5AE');
  if (logSheet.getMaxRows() > MAX_ROW_SIZE) {
    logSheet.deleteRows(2, logSheet.getMaxRows() - MAX_ROW_SIZE);
  }
}

/**
 * This function will whipe the data of a given sheet and restore the headers
 * (if provided)
 * @param{string} sheetName: Name of the sheet that contains the column we want
 * to clean
 * @param{?object} columnsHeader: Index of the column that we want to clean
 * @public
 */
function cleanSheet(sheetName, columnsHeader) {
  const inputSheet =
      SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (columnsHeader.length > 1) {
    inputSheet.clearContents().appendRow(columnsHeader);
  }
}

/**
 * This function will whipe the data of all columns of the given sheet starting
 * from the selected index (default = 1)
 * @param{string} sheetName: Name of the sheet that contains the column we want
 * to clean
 * @param{number=} columnIndex: Index of the column that we want to clean
 * @public
 */
function cleanSheetColumnsFromIndex(sheetName, columnIndex = 1) {
  const inputSheet =
      SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  inputSheet.getRange(2, columnIndex, inputSheet.getLastRow()).clearContent();
}

/**
 * Utility function to calculate the date difference in number of days
 *
 * It expects the input dates to be objects that can be parsed using the Date
 * constructor.
 * @param{!object} firstDate: First date used to compare.
 * @param{!object} secondDate: Second date used to compare.
 * @return{number} number of days that separate the 2 input dates (absolute
 * value).
 * @public
 */
function dateDifferenceInDaysObjects(firstDate, secondDate) {
  const diffTime = Math.abs(secondDate - firstDate);
  return Math.ceil(diffTime / MILLISECONDS_PER_DAY);
}


/**
 * Utility function to calculate the date difference in number of days
 *
 * It expects the input dates to be strings that can be parsed using the Date
 * constructor.
 * @param{string} firstDateString: First date used to compare.
 * @param{string} secondDateString: Second date used to compare.
 * @return{number} number of days that separate the 2 input dates (absolute
 * value).
 * @public
 */
function dateDifferenceInDays(firstDateString, secondDateString) {
  return dateDifferenceInDaysObjects(
      new Date(firstDateString), new Date(secondDateString));
}

/**
 * Utility function to shift a date by the given number of days
 *
 * It expects the input date as an object and the number of days to shift.
 * If the number of days is positive then the date will be shifted towards the
 * future, otherwise towards the past.
 * @param{!object} dateObject: Date that we want to shift
 * @param{number=} numberOfDays: days to shift the date. Default value is 0.
 * @return{!object} Date object shifted by numberOfDays days .
 * @public
 */
function shiftDateByDays(dateObject, numberOfDays = 0) {
  if (!numberOfDays || numberOfDays === 0) {
    return dateObject;
  }
  return new Date(dateObject.getTime() + (numberOfDays * MILLISECONDS_PER_DAY));
}

/**
 * Function that sets to '' the value of every element in a column.
 *
 * it allows to specify the starting row for more accuracy (leave headers
 * intact) however it will whipe the data from all the values below.
 * @param{string} sheetName: Name of the sheet that contains the column we want
 * to clean
 * @param{number} columnIndex: Index of the column that we want to clean
 * @param{number=} startingRow: Index of the first row of column that we want to
 * clean. Default value is 2
 * @public
 */
function cleanDataFromColumn(sheetName, columnIndex, startingRow = 2) {
  let sheetObj =
      SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);

  for (let rowIndex = startingRow; rowIndex <= sheetObj.getLastRow();
       rowIndex++) {
    sheetObj.getRange(rowIndex, columnIndex).setValue('');
  }
}

/**
 * Function to remove hours, minutes and seconds from a date String
 * @param{string} dateString: The date from which you want to remove hours,
 * minutes and seconds
 * @return{!object} Date object with the follogin pattern: yyyy-MM-dd 01:00:00 .
 * @public
 */
function cleanDateString(dateString) {
  return cleanDateObject(new Date(dateString));
}

/**
 * Function to remove hours, minutes and seconds from a date object
 * @param{!object} dateObject: The date from which you want to remove hours,
 * minutes and seconds
 * @return{!object} Date object with the follogin pattern: yyyy-MM-dd 01:00:00 .
 * @public
 */
function cleanDateObject(dateObject) {
  dateObject.setHours(0);
  dateObject.setSeconds(0);
  dateObject.setMinutes(0);
  return dateObject;
}

/**
 * Checks if the trigger is currently active
 * @param{string} handlerFunction: Name of the function that the trigger handles
 * @return {boolean} TRUE if the passed function has a trigger enabled, FALSE
 *     otherwise.
 * @public
 */
function checkTriggerIsActive(handlerFunction) {
  return !!handlerFunction &&
      ScriptApp.getProjectTriggers().some(
          trig => trig.getHandlerFunction() === handlerFunction);
}

/**
 * Function loops through all triggers and delete the trigger identified by
 * input handler function.
 * @param{string} handlerFunction: Name of the function that the trigger handles
 * @public
 */
function deleteTargetTrigger(handlerFunction) {
  if (!!handlerFunction) {
    let triggers = ScriptApp.getProjectTriggers();
    for (let i = 0; i < triggers.length; i++) {
      if (triggers[i].getHandlerFunction() === handlerFunction) {
        ScriptApp.deleteTrigger(triggers[i]);
        break;
      }
    }
  }
}

/**
 * Function which deletes all enabled triggers.
 * @public
 */
function deleteAllTriggers() {
  let triggers = ScriptApp.getProjectTriggers();
  for (let i = 0; i < triggers.length; i++) {
    ScriptApp.deleteTrigger(triggers[i]);
  }
}


/**
 * Function to retrieve the day of week AppScript constant from a date object
 * @param{!object} dateObj: the date from which it is required to extract the
 * AppScript day of the week constant
 * @return {?Enum} The appscript enum for week day (throw exception if no match
 *     is found for provided input).
 * @public
 */
function getDayOfWeekFromDate(dateObj) {
  if (!!dateObj) {
    switch (dateObj.getDay()) {
      case 0:
        return ScriptApp.WeekDay.SUNDAY;
      case 1:
        return ScriptApp.WeekDay.MONDAY;
      case 2:
        return ScriptApp.WeekDay.TUESDAY;
      case 3:
        return ScriptApp.WeekDay.WEDNESDAY;
      case 4:
        return ScriptApp.WeekDay.THURSDAY;
      case 5:
        return ScriptApp.WeekDay.FRIDAY;
      case 6:
        return ScriptApp.WeekDay.SATURDAY;
      default:
        throw `No day found for input ${dateObj}`;
    }
  }
  throw `Provided date object to retrieve the week day was null!`;
}

/**
 * Function to create a dropdown menu
 * @param{!Array<string>} dropList: List of elements to display in the dropdown
 * menu
 * @param{!object} outputSheet: Sheet where the dropdown will be created
 * @param{number} outputColumn: Column index where the dropdown will be created
 * @param{number=} startingRow: first row of the column where the dropdown will
 * be created (from: startingRow to: end of document). default value =
 * DEFAULT_FIRST_ROW ( = 2)
 * @public
 */
function createDropDownMenu(
    dropList, outputSheet, outputColumn, startingRow = DEFAULT_FIRST_ROW) {
  let validationRule =
      SpreadsheetApp.newDataValidation().requireValueInList(dropList).build();
  let dropRange = outputSheet.getRange(startingRow, outputColumn, 100);
  dropRange.setDataValidation(validationRule);
}

/**
 * Function to retrieve the values of a single configuration entry
 * @param{number} rowIndex: index of the row from which the configuration values
 * have to be retrieved
 * @return{!object} obj: object containing for every attribute its value for the
 * given row index.
 * @public
 */
function getTargetConfigurationRow(rowIndex) {
  return {
    accountName:
        configSheet.getRange(rowIndex, CONFIG_FIELDS_MAP.get('AccountName'))
            .getValue(),
    accountNumber:
        configSheet.getRange(rowIndex, CONFIG_FIELDS_MAP.get('AccountGroup'))
            .getValue(),
    filterStatus:
        configSheet.getRange(rowIndex, CONFIG_FIELDS_MAP.get('FilterStatus'))
            .getValue(),
    filterCountry:
        configSheet.getRange(rowIndex, CONFIG_FIELDS_MAP.get('FilterRegion'))
            .getValue()
  };
}
