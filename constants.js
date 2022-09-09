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
 * @fileoverview All global constants and variables used by this project.
 * AppScript doesn't allow for global const variables hence they are all defined
 * with "let".
 */

let MILLISECONDS_PER_DAY = 1000 * 60 * 60 *
    24;  // this variable is used to calculate the date ranges passed to the API
         // call to retrieve location's insights.

// How many locations to be considered in batch to download insights from. Best
// observed compromise for the number of API calls at a time and performance in
// writing the stack of locations on the spreadsheet.
let MAX_LOCS_IN_BATCH = 5;

// Starting row for every sheet since there're headers that shall not be
// considered.
let DEFAULT_FIRST_ROW = 2;
let MAX_BATCH_RETRY = 10;
let DEFAULT_RETENTION_PERIOD_IN_WEEKS = 52;
let DATE_REGEX = new RegExp('d-[0-9]{4}-[0-9]{2}-[0-9]{2}');

// Global constants to keep the sheet's names in the Google Spreadsheet.
let INSTRUCTION_SHEET_NAME = 'Instructions';
let CONFIG_SHEET_NAME = 'Configuration';
let ACC_SHEET_NAME = 'Accounts';
let LOC_SHEET_NAME = 'Locations';
let INS_SHEET_NAME = 'Insights';
let LOG_SHEET_NAME = 'Log';

let CONFIG_SHEET_FIRST_ROW = 5;

let doc = SpreadsheetApp.getActive();

let instructionSheet = doc.getSheetByName(INSTRUCTION_SHEET_NAME);
let configSheet = doc.getSheetByName(CONFIG_SHEET_NAME);
let accountSheet = doc.getSheetByName(ACC_SHEET_NAME);
let locSheet = doc.getSheetByName(LOC_SHEET_NAME);
let insSheet = doc.getSheetByName(INS_SHEET_NAME);
let logSheet = doc.getSheetByName(LOG_SHEET_NAME);

let LAST_LOCATION_UPDATE_COLUMN = 'lastLocationUpdate';
let LAST_INSIGHT_UPDATE_COLUMN = 'lastInsightsUpdate';

// The following variables are used to create the header names of the sheets and
// drop-down menus.
let NO_STATUS_FILTER = '-none-';
let STATUS_FILTER_VALUES = [
  NO_STATUS_FILTER, 'OPEN_FOR_BUSINESS_UNSPECIFIED', 'OPEN',
  'CLOSED_PERMANENTLY', 'CLOSED_TEMPORARILY'
];

let CONFIG_FIELDS_MAP = new Map()
                            .set('AccountName', 1)
                            .set('AccountGroup', 2)
                            .set('FilterStatus', 3)
                            .set('FilterRegion', 4)
                            .set(LAST_LOCATION_UPDATE_COLUMN, 5);

let ACCOUNT_FIELDS_MAP = new Map()
                             .set('name', 1)
                             .set('accountName', 2)
                             .set('type', 3)
                             .set('role', 4)
                             .set('permissionLevel', 5);

let LOCATION_FIELDS_MAP = new Map()
                              .set('account', 1)
                              .set('name', 2)
                              .set('locationName', 3)
                              .set('storeCode', 4)
                              .set('Status', 5)
                              .set('region', 6)
                              .set('categoryName', 7)
                              .set(LAST_INSIGHT_UPDATE_COLUMN, 8);

let INSIGHT_FIELDS_MAP = new Map()
                             .set('account', 1)
                             .set('Location Name', 2)
                             .set('StoreCode', 3)
                             .set('Region', 4)
                             .set('Status', 5)
                             .set('Category Name', 6)
                             .set('Time Zone', 7)
                             .set('Queries Direct', 8)
                             .set('Queries Inirect', 9)
                             .set('Queries Chain', 10)
                             .set('Views Maps', 11)
                             .set('Views Search', 12)
                             .set('Actions Website', 13)
                             .set('Actions Phone', 14)
                             .set('Actions Driving Directions', 15)
                             .set('Photos View Merchant', 16)
                             .set('Photos Views Customers', 17)
                             .set('Photos Count Merchant', 18)
                             .set('Photos Count Customers', 19)
                             .set('Local Post Views Search', 20)
                             .set('Local Post Actions Call To Action', 21)
                             .set('Start Week Date', 22)
                             .set('End Week Date', 23);

let LOG_FIELDS_MAP = new Map().set('Time', 1).set('Message', 2);

// Trigger functions
let WEEKLY_TRIGGERED_FUNCTION = 'initializeWeeklyDownload_';
let INSIGHTS_YEAR_RETRY_TRIGGERED_FUNCTION = 'getInsightsYearWithRetry_';
let LOCATIONS_YEAR_RETRY_TRIGGERED_FUNCTION = 'getLocationsYearWithRetry_';
let LOCATIONS_WEEKLY_RETRY_TRIGGERED_FUNCTION = 'getLocationsWeeklyWithRetry_';
let INSIGHTS_WEEKLY_RETRY_TRIGGERED_FUNCTION = 'geInsightsWeeklyWithRetry_';
# baby-alligator
