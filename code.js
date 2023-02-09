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
 * @fileoverview Main functions to collect data from the Google Business Profile
 * API and store it in the sheets of the attached Google Sheet.
 */

// START - INITIAL TRIGGERS //
/**
 * Function tha cleans the configuration lastLocationUpdate column and
 * initializes the year location trigger. Needed to decouple the locationRetry
 * trigger from the deletition of column lastLocationUpdate (which is used to
 * keep track of the status of the retrieval of the locations)
 * @public
 */
function initializeYearlyDownload_() {
  deleteAllTriggers();
  cleanSheet(INS_SHEET_NAME, Array.from(INSIGHT_FIELDS_MAP.keys()));
  cleanSheetColumnsFromIndex(
      CONFIG_SHEET_NAME, CONFIG_FIELDS_MAP.get(LAST_LOCATION_UPDATE_COLUMN));
  cleanSheetColumnsFromIndex(
      LOC_SHEET_NAME, LOCATION_FIELDS_MAP.get(LAST_INSIGHT_UPDATE_COLUMN));

  locationsYearRetryTrigger_();
}

/**
 * Function tha cleans the configuration lastLocationUpdate column and
 * initializes the weekly location trigger. Needed to decouple the locationRetry
 * trigger from the deletition of column lastLocationUpdate (which is used to
 * keep track of the status of the retrieval of the locations)
 * @private
 */
function initializeWeeklyDownload_() {
  shortenLogSheet();
  locationsWeeklyRetryTrigger_();
}
// END - INITIAL TRIGGERS //

// START - YEAR FILTER CHAIN - Locations + insights //


/**
 * Function which set and activate an hourly weekly trigger for the retry of
 * insights download. This trigger is meant to be removed as soon as all the
 * insights have been downloaded
 * @private
 */
function locationsYearRetryTrigger_() {
  cleanSheet(LOC_SHEET_NAME, Array.from(LOCATION_FIELDS_MAP.keys()));
  cleanDataFromColumn(
      CONFIG_SHEET_NAME, CONFIG_FIELDS_MAP.get(LAST_LOCATION_UPDATE_COLUMN));
  customLog_('Locations sheet cleaned up, ready to be repopulated.');

  ScriptApp.newTrigger(LOCATIONS_YEAR_RETRY_TRIGGERED_FUNCTION)
      .timeBased()
      .everyHours(1)
      .create();

  // trigger the function immediately since otherwise we have to wait 1 hour
  getLocationsYearWithRetry_();
}

/**
 * Function that runs the getLocations_ method and, once completed, removes the
 * retry year trigger
 * @private
 */
function getLocationsYearWithRetry_() {
  // if no result is provided by getInsights_ due to timeout or other errors,
  // isCompleted is initialized to false
  let isCompleted = getLocations_() || false;
  if (isCompleted) {
    deleteTargetTrigger(LOCATIONS_YEAR_RETRY_TRIGGERED_FUNCTION);
    insightsYearRetryTrigger_();
  }
}

/**
 * Function which set and activate an hourly year trigger for the retry of
 * insights download. This trigger is meant to be removed as soon as all the
 * insights have been downloaded
 * @private
 */
function insightsYearRetryTrigger_() {
  ScriptApp.newTrigger(INSIGHTS_YEAR_RETRY_TRIGGERED_FUNCTION)
      .timeBased()
      .everyHours(1)
      .create();

  // trigger the function immediately since otherwise we have to wait 1 hour
  getInsightsYearWithRetry_();
}

/**
 * Function that runs the getInsights_ method and, once completed, removes the
 * insights retry year trigger and creates the weekly trigger.
 * @private
 */
function getInsightsYearWithRetry_() {
  // if no result is provided by getInsights_ due to timeout or other errors,
  // isCompleted is initialized to false
  let isCompleted = getInsights_(true) || false;

  if (isCompleted) {
    // Once all locations and insights have been retrieved for year trigger, it
    // deletes all triggers and starts the weekly one.
    deleteAllTriggers();

    let today =
        new Date(configSheet
                     .getRange(
                         CONFIG_SHEET_FIRST_ROW,
                         CONFIG_FIELDS_MAP.get(LAST_LOCATION_UPDATE_COLUMN))
                     .getValue()
                     .replace('d-', ''));
    customLog_('Initializing weekly trigger');

    today = cleanDateObject(today);
    let yesterday = shiftDateByDays(today, -1);
    let weekDay = getDayOfWeekFromDate(yesterday);
    weeklyTrigger_(weekDay);
  }
}

// END - YEAR FILTER CHAIN - Locations + insights //


// START - WEEKLY FILTER CHAIN - Locations + insights //

/**
 * Function which set and activate an hourly weekly trigger for the retry of
 * insights download. This trigger is meant to be removed as soon as all the
 * insights have been downloaded
 * @private
 */
function locationsWeeklyRetryTrigger_() {
  cleanSheet(LOC_SHEET_NAME, Array.from(LOCATION_FIELDS_MAP.keys()));
  cleanDataFromColumn(
      CONFIG_SHEET_NAME, CONFIG_FIELDS_MAP.get(LAST_LOCATION_UPDATE_COLUMN));
  customLog_('Locations sheet cleaned up, ready to be repopulated.');

  ScriptApp.newTrigger(LOCATIONS_WEEKLY_RETRY_TRIGGERED_FUNCTION)
      .timeBased()
      .everyHours(1)
      .create();

  // trigger the function immediately since otherwise we have to wait 1 hour
  getLocationsWeeklyWithRetry_();
}

/**
 * Function that runs the getLocations_ method and, once completed, removes the
 * retry weekly trigger
 * @private
 */
function getLocationsWeeklyWithRetry_() {
  // if no result is provided by getInsights_ due to timeout or other errors,
  // isCompleted is initialized to false
  let isCompleted = getLocations_() || false;
  if (isCompleted) {
    deleteTargetTrigger(LOCATIONS_WEEKLY_RETRY_TRIGGERED_FUNCTION);
    insightsWeeklyRetryTrigger_();
  }
}

/**
 * Function which set and activate an hourly weekly trigger for the retry of
 * insights download. This trigger is meant to be removed as soon as all the
 * insights have been downloaded.
 * @private
 */
function insightsWeeklyRetryTrigger_() {
  ScriptApp.newTrigger(INSIGHTS_WEEKLY_RETRY_TRIGGERED_FUNCTION)
      .timeBased()
      .everyHours(1)
      .create();

  // trigger the function immediately since otherwise we have to wait 1 hour
  geInsightsWeeklyWithRetry_();
}

/**
 * Function that runs the getInsights_ method and, once completed, removes the
 * insights retry weekly trigger
 * @private
 */
function geInsightsWeeklyWithRetry_() {
  // if no result is provided by getInsights_ due to timeout or other errors,
  // isCompleted is initialized to false
  let isCompleted = getInsights_(false) || false;
  if (isCompleted) {
    // After all insights have been retrieved we have to remove the oldest week
    // from the insights sheet
    cleanInsightsOldestWeek_();

    // Delete temp retry trigger, leave weekly trigger
    deleteTargetTrigger(INSIGHTS_WEEKLY_RETRY_TRIGGERED_FUNCTION);
  }
}

// START - WEEKLY FILTER CHAIN - Locations + insights //


/**
 * Function which builds the custom menu 'Collect GBP data' in the Trix and
 * creates (if missing) the sheets.
 * @param{?object} e: event, automatic input provided by the listener that
 * contains relevant information.
 */
function onOpen(e) {
  init_();

  SpreadsheetApp.getUi()
      .createMenu('Collect GBP data')
      .addItem('Collect accounts', 'getAccounts')
      .addItem('Collect locations and insights', 'initializeYearlyDownload_')
      //.addItem('Clear all triggers (for development purpose)',
      //'deleteAllTriggers')
      .addItem('Reset and use new config', 'resetAndRestart')
      .addToUi();
}

/**
 * Event listener that updates 'Configuration' sheet column 'AccountGroup' once
 * selected (or removed) the account name.
 * @param{!object} e: event, automatic input provided by the listener that
 * contains relevant information.
 */
function onEdit(e) {
  const ACCOUNT_SHEET_NAME_COLUMN = 1;
  const ACCOUNT_SHEET_NAMEGROUP_COLUMN = 2;
  const CONFIG_SHEET_GROUP_COLUMN = 2;

  // Verify that the action has been performed on Configuration sheet on an
  // input field (not on the header) and on the proper column
  if (SpreadsheetApp.getActiveSheet().getName() === CONFIG_SHEET_NAME &&
      e.range.rowStart > 1 &&
      e.range.columnStart === ACCOUNT_SHEET_NAME_COLUMN) {
    // New value from the user that triggered this function
    const newAccountName = e.value;

    // Cell that needs to be updated
    const targetCell =
        configSheet.getRange(e.range.rowStart, CONFIG_SHEET_GROUP_COLUMN);

    if (!!newAccountName) {
      const accountValues =
          accountSheet
              .getRange(
                  2, ACCOUNT_SHEET_NAME_COLUMN, accountSheet.getLastRow(),
                  ACCOUNT_SHEET_NAMEGROUP_COLUMN)
              .getValues();  // Proper range made of columns 'Name' and
                             // 'AccountName' of Accounts sheet

      // Search the groupname in Accounts sheet and retrieve the name to fill
      // the Configuration cell
      for (let row = 0; row < accountValues.length; row++) {
        if (accountValues[row][ACCOUNT_SHEET_NAMEGROUP_COLUMN - 1] ===
            newAccountName) {  // Since it is a matrix index instead of the Trix
                               // column index I need to use the -1
          targetCell.setValue(
              accountValues[row]
                           [ACCOUNT_SHEET_NAME_COLUMN -
                            1]);  // Since it is a matrix index instead of the
                                  // Trix column index I need to use the -1
          break;
        }
      }
    } else {
      // set to blank the value of the cell if the account name has been deleted
      targetCell.setValue('');
    }
  }
}

 
    /**
     * Function which set and activate a weekly trigger for the
     * 'initializeWeeklyDownload_' function to be run every Monday from 5:00 am
     * to 6:00 am.
     * @param{number} weekDay: Appscript constant for the weekday configuration
     * of the trigger trigger.
     * @private
     */
    function weeklyTrigger_(weekDay) {
      // pick current hour and, if it is possible to shift it back by 1 without
      // changing day, then we do so
      let hours = new Date().getHours();
      hours = hours > 1 ? hours - 1 : hours;

      ScriptApp.newTrigger(WEEKLY_TRIGGERED_FUNCTION)
          .timeBased()
          .onWeekDay(weekDay)
          .atHour(hours)
          .create();
    }

/**
 * Collects Account level data from Google Business Profile API and stores it in
 * the Accounts sheet.
 * @public
 */
function getAccounts() {
  configSheet
      .getRange(
          CONFIG_SHEET_FIRST_ROW, 1, configSheet.getLastRow(),
          configSheet.getLastColumn())
      .clearContent();
  shortenLogSheet();
  customLog_('Accounts list update started.');
  const accountsUrl =
      'https://mybusinessaccountmanagement.googleapis.com/v1/accounts';
  accountSheet
      .getRange(
          DEFAULT_FIRST_ROW, 1, accountSheet.getLastRow(),
          accountSheet.getLastColumn())
      .clearContent();
  const accountsResult = callApi_(accountsUrl, 'GET');
  for (let accountIndex in accountsResult['accounts']) {
    const newRow = [];
    for (let field of ACCOUNT_FIELDS_MAP.keys()) {
      newRow.push(accountsResult['accounts'][accountIndex][field]);
    }
    accountSheet.appendRow(newRow);
  }

  let accDropList =
      accountSheet
          .getRange(
              DEFAULT_FIRST_ROW, ACCOUNT_FIELDS_MAP.get('accountName'),
              accountSheet.getLastRow())
          .getValues();
  createDropDownMenu(accDropList, configSheet, 1, CONFIG_SHEET_FIRST_ROW);
  createDropDownMenu(
      STATUS_FILTER_VALUES, configSheet, 3, CONFIG_SHEET_FIRST_ROW);

  customLog_(`Accounts update completed ${
      accountsResult['accounts'].length} total accounts).`);
}

/**
 * This function verifies that all locations have been processed (their insights
 * have been downlaoded). it checks that the lastUpdate value of the locations
 * matches the pattern used for the processed locations (d-yyyy-MM-dd).
 * @return{boolean}: true if all insights have been retrieved for each
 * locations, false otherwise
 * @private
 */
function haveAllLocationsBeenProcessed_() {
  if (locSheet.getLastRow() > 1) {
    const lastInsightsUpdate =
        locSheet
            .getRange(
                DEFAULT_FIRST_ROW,
                LOCATION_FIELDS_MAP.get(LAST_INSIGHT_UPDATE_COLUMN),
                locSheet.getLastRow() - 1)
            .getValues();
    return lastInsightsUpdate.every(date => DATE_REGEX.test(date[0]));
  }
  return false;
}

/**
 * Get account details and collects the corresponding Location data.
 * @return{boolean}: true if all locations have been retrieved for each account,
 * false otherwise
 * @public
 */
function getLocations_() {
  const locationsUrl =
      'https://mybusinessbusinessinformation.googleapis.com/v1/%account%/locations?readMask=name,storeCode,title,categories,storefrontAddress,openInfo';

  // Since GBP data could have a few days of delay we set the time range
  // up to previous week in order to avoid retrieval of partial data
  const processedDate = `d-${
      Utilities.formatDate(
          shiftDateByDays(new Date(), -7), 'Europe/Rome', 'yyyy-MM-dd')}`;
  customLog_('Locations and insights update started.');

  for (let j = CONFIG_SHEET_FIRST_ROW; j <= configSheet.getLastRow(); j++) {
    // If the account has no lastLocationUpdate it has to processed
    if (!DATE_REGEX.test(
            configSheet
                .getRange(j, CONFIG_FIELDS_MAP.get(LAST_LOCATION_UPDATE_COLUMN))
                .getValue())) {
      const configuration = getTargetConfigurationRow(j);

      let restAPIurlCall = locationsUrl;

      // Manage filtering via url query parameters. It is build according to
      // user configurations.
      if ((!!configuration.filterStatus &&
           configuration.filterStatus !== NO_STATUS_FILTER) ||
          !!configuration.filterCountry) {
        let filter = '&filter=';
        let nextFilter = '';
        if (!!configuration.filterStatus &&
            configuration.filterStatus !== NO_STATUS_FILTER) {
          filter += `openInfo.status=%22${configuration.filterStatus}%22`;
          nextFilter = '+AND+';
        }
        if (!!configuration.filterCountry) {
          filter += `${nextFilter}storefrontAddress.regionCode=%22${
              configuration.filterCountry}%22`;
          nextFilter = '+AND+';
        }
        restAPIurlCall += filter;
      }

      // Locations list for the account.
      const locResults = callApi_(
          restAPIurlCall.replace('%account%', configuration.accountNumber),
          'GET', true);
      const accLocRows = [];
      for (let p in locResults) {
        let page = locResults[p];
        for (let locIndex in page['locations']) {
          let loc = page['locations'][locIndex];
          if (!!loc) {
            let regionCode = loc['storefrontAddress'] ?
                loc['storefrontAddress']['regionCode'] :
                'N/A';
            let status = loc['openInfo'] ? loc['openInfo']['status'] : 'N/A';
            // Collect Location info we want to store and show.
            accLocRows.push([
              configuration.accountNumber, loc['name'], loc['title'],
              loc['storeCode'], status, regionCode,
              loc['categories']['primaryCategory']['displayName']
            ]);
          }
        }
      }  // end of locResults loop

      customLog_(`Total locations for account "${
          configuration.accountName}": '${accLocRows.length}`);
      if (accLocRows.length > 0) {
        locSheet
            .getRange(
                locSheet.getLastRow() + 1, 1, accLocRows.length,
                accLocRows[0].length)
            .setValues(accLocRows);
      }
      configSheet
          .getRange(j, CONFIG_FIELDS_MAP.get(LAST_LOCATION_UPDATE_COLUMN))
          .setValue(processedDate);
    }
  }
  customLog_(
      'Locations list update completed - now looping through locations for insights');

  for (let j = CONFIG_SHEET_FIRST_ROW; j <= configSheet.getLastRow(); j++) {
    // If even one account is missing then the getLocations_ function has to be
    // executed again by the retry trigger
    if (!DATE_REGEX.test(
            configSheet
                .getRange(j, CONFIG_FIELDS_MAP.get(LAST_LOCATION_UPDATE_COLUMN))
                .getValue())) {
      return false;
    }
  }
  return true;
}

/**
 * Loops through Locations and collects corresponding Insights data for the time
 * range [retention weeks * 7, 8 days ago]. It can either process one whole year
 * (checkForYear = true) or just the last week (checkForYear = false) At every
 * iteration it updates the lastUpdate value of the related location. If there
 * are still weeks to retrieve the value of lastUpdate will be in the following
 * pattern: yyyy-MM-dd If all the weeks have been processed for the insights of
 * a given location the pattern of lastUpdate will be: d-yyyy-MM-dd.
 * @param{boolean} checkForYear: indicates if this is the first iteration
 * (yearly) or the weekly one.
 * @return{boolean}: true if all insights have been retrieved for each
 * locations, false otherwise
 * @public
 */
function getInsights_(checkForYear) {
  const locData = locSheet.getDataRange().getValues();

  // retrieve the date to be used as 'today' from the updated date of location
  // in the configuration sheet. The reason is that since this jobs can last
  // quite a lot, the date could change and this would introduce missmatches.
  let today =
      new Date(configSheet
                   .getRange(
                       CONFIG_SHEET_FIRST_ROW,
                       CONFIG_FIELDS_MAP.get(LAST_LOCATION_UPDATE_COLUMN))
                   .getValue()
                   .replace('d-', ''));
  today = cleanDateObject(today);

  const numberOfWeeks =
      checkForYear ? configSheet.getRange(1, 1).getValue() : 1;

  if (typeof numberOfWeeks !== 'number' || numberOfWeeks < 1) {
    const ui = SpreadsheetApp.getUi();
    ui.alert(
        'ERROR!',
        'Provided number of weeks must be numeric and greater than 0. Aborting execution.',
        ui.ButtonSet.OK);
    customLog_(
        'ERROR! Provided number of weeks must be numeric and > 0. Aborting the job');
    return false;
  }

  // If all locations have been processed then I can return true
  if (haveAllLocationsBeenProcessed_()) {
    customLog_('All insights had already been processed - job completed!');
    return true;
  }

  // By parametrizing the value of 'firstDayWeekTimeIntervalObj' this function
  // manages both yearly and weekly updates
  let firstDayWeekTimeIntervalObj =
      shiftDateByDays(today, -(7 * numberOfWeeks));
  // lastDayWeekTimeIntervalObj is obtained by shifting
  // firstDayWeekTimeIntervalObj by 7 days. These variables are tightly coupled
  // and will always represent 1 week
  let lastDayWeekTimeIntervalObj =
      shiftDateByDays(firstDayWeekTimeIntervalObj, 7);

  // On the first iteration of year trigger we could require a week interval
  // shift in case previous trigger run didn't complete the insights update
  let firstIterationOfYear = checkForYear;

  // Shift by 1 week until I reach my target date
  do {
    customLog_(`collect insights for startDate: ${
        Utilities.formatDate(
            firstDayWeekTimeIntervalObj, 'Europe/Rome', 'yyyy-MM-dd')}`);
    let locationIndex = 1;

    // Loop through the locations array
    while (locationIndex < locData.length) {
      // If there is a match it means that the location has been updated for all
      // the required weeks and we can ignore said location
      while (locationIndex < locData.length &&
             DATE_REGEX.test(
                 locData[locationIndex]
                        [LOCATION_FIELDS_MAP.get(LAST_INSIGHT_UPDATE_COLUMN) -
                         1])) {
        locationIndex++;
      }

      // Return if every locations' insights have been updated
      if (locationIndex === locData.length) {
        customLog_(
            `All insights have been collected for all locations for start date ${
                Utilities.formatDate(
                    firstDayWeekTimeIntervalObj, 'Europe/Rome',
                    'yyyy-MM-dd')}`);
        continue;
      }

      // Retrieve the next location to be processed according to already
      // processed locations and trigger type
      let nextLocationToProcess = identifyNextLocationToProcess_(
          locData, locationIndex, firstIterationOfYear,
          lastDayWeekTimeIntervalObj);
      locationIndex = +nextLocationToProcess.locationIndex;
      lastDayWeekTimeIntervalObj =
          nextLocationToProcess.lastDayWeekTimeIntervalObj;

      // bucket of locations which insights have to be fetched
      let bucketInformation =
          createLocationsToFetchBucket_(locData, locationIndex);
      let locationsToFetch = bucketInformation.locationsToFetch;
      let accountName = bucketInformation.accountName;

      // Retrieve insights
      collectBatchInsights_(
          firstDayWeekTimeIntervalObj, lastDayWeekTimeIntervalObj, accountName,
          locationsToFetch, locationIndex, 1, checkForYear, today);

      // update index location
      locationIndex += Object.keys(locationsToFetch).length;

      // update trigger and return if every locations' insights have been
      // updated
      if (locationIndex === locData.length - 1) {
        customLog_(
            `All insights have been collected for all locations for start date ${
                Utilities.formatDate(
                    firstDayWeekTimeIntervalObj, 'Europe/Rome',
                    'yyyy-MM-dd')}`);
        continue;
      }

    }  // end of for loop that iterates the locations

    // Update the week time interval
    firstDayWeekTimeIntervalObj = lastDayWeekTimeIntervalObj;
    lastDayWeekTimeIntervalObj =
        shiftDateByDays(firstDayWeekTimeIntervalObj, 7);

    SpreadsheetApp.flush();

  } while (lastDayWeekTimeIntervalObj.getTime() <= today.getTime());

  if (haveAllLocationsBeenProcessed_()) {
    customLog_('Insights update completed successfully - job completed!');
    return true;
  } else {
    customLog_('Insights update NOT completed - job will resume in 1 hour');
  }
  return false;
}

/**
 * Collects Insights data from the API and appends it in the insights sheet.
 * It updates the locations sheet lastUpdate column in order to keep track of
 * the current progress.
 * @param{!object} firstDayWeekObj: First date of the week range
 * @param{!object} lastDayWeekObj: Last date of the week range
 * @param{string} accountName: Name of the account to which the provided
 * locations belong
 * @param{!object} locationsToFetch: Data structure of the locations that we
 * need to fetch
 * @param{string} nextLocationRow: Cursor of the row of the next location to be
 * retrieved
 * @param{number} retryCount: Counter for the number of exceptions that happen
 * in the HTTPS invocation
 * @param{boolean} checkForYear: indicates if this is the first iteration
 * (yearly) or the weekly one.
 * @param{!object} lastDateTimeRange: last date of the overall timerange for the
 * insights.
 *
 * @private
 */
function collectBatchInsights_(
    firstDayWeekObj, lastDayWeekObj, accountName, locationsToFetch,
    nextLocationRow, retryCount, checkForYear, lastDateTimeRange) {
  const insights = retrieveLocationInsights_(
      firstDayWeekObj, lastDayWeekObj, accountName, locationsToFetch,
      retryCount, nextLocationRow, collectBatchInsights_, checkForYear,
      lastDateTimeRange);
  const insRows = [];

  // array of entries to append on the insight sheet
  for (let e in insights) {
    let entry = insights[e];
    insRows.push([
      accountName,
      entry['locationName'],
      entry['storeCode'],
      entry['Region'],
      entry['Status'],
      entry['category'],
      entry['TimeZone'],
      entry['queriesDirect'],
      entry['queriesIndirect'],
      entry['queriesChain'],
      entry['viewsMaps'],
      entry['viewsSearch'],
      entry['actionsWebsite'],
      entry['actionsPhone'],
      entry['actionsDrivingDirections'],
      entry['photosViewMerchant'],
      entry['photosViewsCustomers'],
      entry['photosCountMerchant'],
      entry['photosCountCustomers'],
      entry['localPostViewsSearch'],
      entry['localPostActionsCallToAction'],
      entry['startWeekDate'],
      entry['endWeekDate']
    ]);
  }

  // Append entries
  if (insRows.length > 0) {
    insSheet
        .getRange(
            insSheet.getLastRow() + 1, 1, insRows.length, insRows[0].length)
        .setValues(insRows);
  }

  const newDates = [];
  let lastDateTimeRangeAsString =
      Utilities.formatDate(lastDateTimeRange, 'Europe/Rome', 'yyyy-MM-dd');
  let lastDayWeek =
      Utilities.formatDate(lastDayWeekObj, 'Europe/Rome', 'yyyy-MM-dd');

  // Pick the target date to fill the Locations sheet: If the difference between
  // the end week date and today is less than 7 days it means that it has
  // reached the current week and that we can use Today as target date.
  // Otherwise use lastDayWeek
  let targetDate =
      (dateDifferenceInDays(lastDayWeek, lastDateTimeRangeAsString) > 7 ?
           lastDayWeek :
           `d-${lastDateTimeRangeAsString}`);
  for (let t = 0; t < Object.keys(locationsToFetch).length; t++) {
    newDates.push([targetDate]);
  }
  // Update locations sheet to track the status of the insights download.
  locSheet
      .getRange(
          nextLocationRow + 1,
          LOCATION_FIELDS_MAP.get(LAST_INSIGHT_UPDATE_COLUMN),
          Object.keys(locationsToFetch).length, 1)
      .setValues(newDates);
}

/**
 * Function that will retrieve the insights of up to 5 locations that belong to
 * the same account for a given week (delimited by firstDayWeekObj and
 * lastDayWeekObj)
 * @param{!object} firstDayWeekObj: First date of the week range
 * @param{!object} lastDayWeekObj: Last date of the week range
 * @param{string} accountName: Name of the account to which the provided
 * locations belong
 * @param{!object} locationsToFetch: Data structure of the locations that we
 * need to fetch
 * @param{number} retryCount: Counter for the number of exceptions that happen
 * in the HTTPS invocation
 * @param{string} nextLocationRow: Cursor of the row of the next location to be
 * retrieved
 * @param{string} retryCallback: callback function to perform the retry in case
 * of errors
 * @param{boolean} checkForYear: indicates if this is the first iteration
 * (yearly) or the weekly one.
 * @return{!Array<!object>} Array of structured objects that will be converted
 * into rows in the insights sheet
 * @param{!object} lastDateTimeRange: last date of the overall timerange for the
 * insights.
 * @private
 */
function retrieveLocationInsights_(
    firstDayWeekObj, lastDayWeekObj, accountName, locationsToFetch, retryCount,
    nextLocationRow, retryCallback, checkForYear, lastDateTimeRange) {
  const fullLocationsNames =
      Object.keys(locationsToFetch).map(loc => `${accountName}/${loc}`);

  const firstDayWeek = Utilities.formatDate(
      firstDayWeekObj, 'Europe/Rome', 'yyyy-MM-dd\'T\'HH:mm:ss.SSS\'Z\'');
  const lastDayWeek = Utilities.formatDate(
      lastDayWeekObj, 'Europe/Rome', 'yyyy-MM-dd\'T\'HH:mm:ss.SSS\'Z\'');
  const locationMetrics = [];
  const body = {
    'locationNames': fullLocationsNames,
    'basicRequest': {
      'metricRequests': [{'metric': 'ALL', 'options': 'AGGREGATED_TOTAL'}],
      'timeRange': {'startTime': firstDayWeek, 'endTime': lastDayWeek}
    }
  };
  const revUrl = `https://mybusiness.googleapis.com/v4/${
      accountName}/locations:reportInsights`;

  const result = callApi_(revUrl, 'POST', true, body);
  // log errors and perform retry if possible
  let page = {};
  for (let p in result) {
    page = result[p];
    if (page['error']) {
      customLog_(page['error']['message']);
      if (retryCount < MAX_BATCH_RETRY) {
        if ( Object.keys(locationsToFetch).length === 0) {
          customLog_(`Ignored Retry #${retryCount} due to no locations left`);
          return []; // Empty array, no element to add
        }  
        // Best effort retry approach, removing first location of batch each
        // time.
        delete locationsToFetch[Object.keys(locationsToFetch)[0]];
        retryCount += 1;
        customLog_(`Retry #${retryCount} with ${
            Object.keys(locationsToFetch).length} locations, starting with ${
            Object.keys(locationsToFetch)[0]}`);
        Utilities.sleep(1000 * retryCount);
        return retryCallback(
            firstDayWeekObj, lastDayWeekObj, accountName, locationsToFetch,
            nextLocationRow, retryCount, checkForYear, lastDateTimeRange);
      }
    }
  }
  const inss = page['locationMetrics'];
  const metricsArray = [
    'queriesDirect', 'queriesIndirect', 'queriesChain', 'viewsMaps',
    'viewsSearch', 'actionsWebsite', 'actionsPhone', 'actionsDrivingDirections',
    'photosViewMerchant', 'photosViewsCustomers', 'photosCountMerchant',
    'photosCountCustomers', 'localPostViewsSearch',
    'localPostActionsCallToAction'
  ];
  for (let i in inss) {
    let ins = inss[i];
    let timeZone = ins['timeZone'];
    let locKey = `locations${ins['locationName'].split('/locations')[1]}`;
    locationMetrics.push({
      'locationName': locationsToFetch[locKey][2],
      'storeCode': locationsToFetch ? locationsToFetch[locKey][3] : 'N/A',
      'Region': locationsToFetch ? locationsToFetch[locKey][5] : 'N/A',
      'Status': locationsToFetch ? locationsToFetch[locKey][4] : 'N/A',
      'category': locationsToFetch ? locationsToFetch[locKey][6] : 'N/A',
      'TimeZone': timeZone,
      'startWeekDate':
          Utilities.formatDate(firstDayWeekObj, 'Europe/Rome', 'yyyy-MM-dd'),
      'endWeekDate':
          Utilities.formatDate(lastDayWeekObj, 'Europe/Rome', 'yyyy-MM-dd')
    });

    for (let index in metricsArray) {
      locationMetrics[locationMetrics.length - 1][metricsArray[index]] =
          ins['metricValues'] ?
          ins['metricValues'][index]['totalValue']['value'] :
          'N/A';
    }
  }
  return locationMetrics;
}

/**
 * Function that will clear all content from insights and lcoation and restart
 * the retrieval of the information
 * @public
 */
function resetAndRestart() {
  const proceed = showConfirmationResetAlert_();
  if (proceed) {
    // Delete sheets Locations and Inights
    cleanSheet(INS_SHEET_NAME, Array.from(INSIGHT_FIELDS_MAP.keys()));
    cleanSheet(LOC_SHEET_NAME, Array.from(LOCATION_FIELDS_MAP.keys()));
    // Force the yearly download
    initializeYearlyDownload_();
  }
}

/**
 * Function that will prompt modal confirmation to delete all current data and
 * download again
 * @return{boolean} choice of the user whether to proceed or not by deleting all
 * data and downloading it again.
 * @private
 */
function showConfirmationResetAlert_() {
  const ui = SpreadsheetApp.getUi();
  const result = ui.alert(
      'Please confirm',
      `Are you sure you want to DELETE ALL current data and use the new config?\nThis will trigger the download of last ${
          configSheet.getRange(1, 1)
              .getValue()} weeks data for the given account(s)`,
      ui.ButtonSet.YES_NO);

  // Process the user's response.
  if (result === ui.Button.YES) {
    // User clicked "Yes".
    return true;
  } else {
    // User clicked "No" or X in the title bar.
    return false;
  }
}

/**
 * This function will remove from insights sheed all the entries that belong to
 * the oldest week
 * @private
 */
function cleanInsightsOldestWeek_() {
  let index = 1;
  const startWeekEntries =
      insSheet
          .getRange(
              DEFAULT_FIRST_ROW, INSIGHT_FIELDS_MAP.get('Start Week Date'),
              insSheet.getLastRow(), 1)
          .getValues();

  if ((startWeekEntries.length > 0) && (!!startWeekEntries[0][0])) {
    let oldest_startWeek = startWeekEntries[0][0];
    while (oldest_startWeek === startWeekEntries[index][0]) {
      index++;
    }
    insSheet.deleteRows(DEFAULT_FIRST_ROW, index);
  }
}

/**
 * This function will find the next location that needs to be processed
 * @param{!Array<!object>} locData: list of locations to be processed
 * @param{number} locationIndex: starting index of the locations. It will be
 * updated in order to match the next item to be processed
 * @param{boolean} firstIterationOfYear: identifies if this is the first
 * iteration of yearly trigger or not.
 * @param{!object} lastDayWeekTimeIntervalObj: end date of the time interval. It
 * will be updated according to the next item to be processed.
 * @return{!object} obj: updated values for locationIndex and
 * lastDayWeekTimeIntervalObj
 * @private
 */
function identifyNextLocationToProcess_(
    locData, locationIndex, firstIterationOfYear, lastDayWeekTimeIntervalObj) {
  let currentLastInsightUpdateDate =
      locData[locationIndex][LOCATION_FIELDS_MAP.get(LAST_INSIGHT_UPDATE_COLUMN) - 1];  // I retrieve from the map the column number so I need to shift it back by 1 to match the array position of the field

  // If we have at least one update and is the first year trigger iteration
  // we check for required week interval updates
  if (currentLastInsightUpdateDate !== '' && firstIterationOfYear) {
    // At least one location has been updated.
    let nextLocationItem = locationIndex + 1;
    let currentLastInsightUpdateDateString = Utilities.formatDate(
        currentLastInsightUpdateDate, 'Europe/Rome', 'yyyy-MM-dd');
    let nextLocationItemDateObj = new Date(
        locData[nextLocationItem]
               [LOCATION_FIELDS_MAP.get(LAST_INSIGHT_UPDATE_COLUMN) - 1]);

    while (nextLocationItem < locData.length - 1 &&
           (currentLastInsightUpdateDateString ===
                Utilities.formatDate(
                    nextLocationItemDateObj, 'Europe/Rome', 'yyyy-MM-dd') ||
            locData[nextLocationItem]
                   [LOCATION_FIELDS_MAP.get(LAST_INSIGHT_UPDATE_COLUMN) - 1] ===
                '')) {
      nextLocationItem++;
      nextLocationItemDateObj =
          locData[nextLocationItem][LOCATION_FIELDS_MAP.get(LAST_INSIGHT_UPDATE_COLUMN) - 1] !==
              '' ?
          new Date(locData[nextLocationItem]
                          [LOCATION_FIELDS_MAP.get(LAST_INSIGHT_UPDATE_COLUMN) -
                           1]) :
          firstDayWeekTimeIntervalObj;
    }

    // Some locations have been updated to a later date than others. I start
    // updating the non-updated ones (the ones at the bottom of the list)
    if (nextLocationItem !== locData.length - 1) {
      locationIndex = nextLocationItem;
      currentLastInsightUpdateDate =
          locData[locationIndex][LOCATION_FIELDS_MAP.get(LAST_INSIGHT_UPDATE_COLUMN) - 1];
    }
    // Update the week time interval so that it matches what I need to
    // update my target date to
    firstDayWeekTimeIntervalObj = new Date(currentLastInsightUpdateDate);
    lastDayWeekTimeIntervalObj =
        shiftDateByDays(firstDayWeekTimeIntervalObj, 7);

    // We need to check for previous unfinished updates only on the first
    // iteration.
    firstIterationOfYear = false;
  }
  return {
    locationIndex: locationIndex,
    lastDayWeekTimeIntervalObj: lastDayWeekTimeIntervalObj
  };
}

/**
 * Function create the bucket of locations to fetch and their related account
 * name for the getInsights_ function
 * @param{!Array<!object>} locData: list of locations to be processed
 * @param{number} locationIndex: index of the first location to be fetched
 * @return{!object} obj: array of locations to fetch and their account name
 * @private
 */
function createLocationsToFetchBucket_(locData, locationIndex) {
  let locationsToFetch = [];

  let subsetIndex = 1;
  let accountName =
      locData[locationIndex][LOCATION_FIELDS_MAP.get('account') - 1];
  locationsToFetch[locData[locationIndex]
                          [LOCATION_FIELDS_MAP.get('name') - 1]] =
      locData[locationIndex];

  while (subsetIndex < MAX_LOCS_IN_BATCH &&
         (subsetIndex + locationIndex) < locData.length &&
         accountName ===
             locData[locationIndex + subsetIndex]
                    [LOCATION_FIELDS_MAP.get('account') - 1]) {
    // I retrieve from the map the column number so I need to shift it back
    // by 1 to match the array position of the field
    locationsToFetch[locData[locationIndex + subsetIndex]
                            [LOCATION_FIELDS_MAP.get('name') - 1]] =
        locData[locationIndex + subsetIndex];
    subsetIndex++;
  }
  return {locationsToFetch: locationsToFetch, accountName: accountName};
}
