DISCLAIMER: This is not an officially supported Google product.

# Baby Alligator

## Table of Contents

1.  [About](#About)
2.  [Setup](#Setup)
    1.  [Google Cloud](#Google-Cloud)
        1.  [Request access to the APIs](#Request-access-to-the-APIs)
    2.  [Spreadsheet](#Spreadsheet)
3.  [First run](#First-run)
    1.  [Change your configuration](#Change-your-configuration)
4.  [Wrapping up](#Wrapping-up)

## About

This tool Allows you to retrieve and analyze your stores insights from Google
Business Profile in a set timeframe.

--------------------------------------------------------------------------------

## Setup

### Google Cloud

Before you can send requests to the Business Profile APIs, you need to
create/own a GCP (Google Cloud Platform) project. You can go
[here](https://console.cloud.google.com/getting-started) to create your new GCP project.
To create a new project click <button>Create project</button>, enter a name, and
click Create.

Once you have access to your GCP project you need to use the Google API Console
to request access to the Business Profile APIs for that project.

In order to run the solution you should first enable your account as **Test
User**:

-   Go to [Google API Console](https://console.cloud.google.com) and
    select the project you created for use with Business Profile.
-   On the left side menu click on **API & Services**
-   On the left side menu click on **Oauth consent screen**
-   In the **Test users** section click on **Add User** and insert the email of
    the Google account used to access Google services.

#### Request access to the APIs

To enable your project and access the APIs, be sure to complete the
[prerequisites](https://developers.google.com/my-business/content/prereqs#request-access).

***NOTE: Even if this solution relies on a GCP project to execute API
requests*** ***it doesn't require a billing account.***

### Spreadsheet

In order to create your spreadsheet you need to:

-   Open your [Google drive](https://drive.google.com/drive/u/1/my-drive)
-   Click on the *plus icon* <button>+</button>
-   Click on **Google Sheets** and then on **Blank spreadsheet**

You now need to add the script files to your spreadsheet:

-   Click on **Extensions** on the top menu
-   Select **Apps Script**
-   Click on **Editor** on the left menu
-   Click on the *plus icon* <button>+</button> next to Files and select
    **Script**
-   Name this file **code** (beware that the full file name will be `code.gs`)
-   Copy the content of `code.js` in `code.gs`
-   Repeat this step for each other js files (`utilities.js`,
    `spreadsheetInitialization.js` and `constants.js`)

In order to enable the APIs required by the script you need to:

-   On the left side menu select **Project Settings** and tick the option **Show
    "appsscript.json" manifest file in editor**.
-   In the editor select the **appsscripts.json** file and add the following
    variable:

    ```
    "oauthScopes": ["https://www.googleapis.com/auth/business.manage", "https://www.googleapis.com/auth/script.external_request", "https://www.googleapis.com/auth/spreadsheets.currentonly", "https://www.googleapis.com/auth/script.scriptapp"]
    ```

--------------------------------------------------------------------------------

## First run

Once you open/refresh the your spreadsheet, the scripts will automatically
generate all the required tabs (if missing).

In the top menu you will now be able to see a new item <button>Collect GBP
Data</button>. Click it and select *Collect accounts*. This action will retrieve
all your accounts and fill the **Accounts** sheet.

Open now the **Configuration** sheet and fill as follows:

-   *Weeks of retention*: default to 52, you can edit this value to increase or
    reduce the retention period.
-   *Account name*: Select from the drop down menu the account you want to
    include. The value for *AccountGroup* will be automatically filled.
-   *FilterStatus*: Select from the drop down menu the status of the locations
    to retrieve
-   *FilterRegion*: Manually insert the country code (1 per row) of the
    locations to retrieve. In case you need to retrieve locations from several
    countries create 1 entry per country

***NOTE: Due to size limitations in the spreadsheet be careful while editing
"Weeks of retention". If you have 900+ locations we recommend not to increase it
and, if possible, to reduce it.***

Now click again on <button>Collect GBP Data</button> and select *Collect
locations and insights*. This will start the process to retrieve all the
locations and insights for your filtered accounts within the defined retention
period. This process could last up to a few days.

Once completed a weekly trigger will update the insights with new data while
erasing the oldest entries.

***NOTE: Do NOT edit the configuration without clicking on "Reset and use new
config" since your configuration will be updated only for new insights entries
and your data will be incoherent.***

### Change your configuration

In case you want to update your configuration to perform a different analysis,
you have to:

-   Update your *Configuration* sheet
-   Click on <button>Collect GBP Data</button> and select *Reset and use new
    config* and confirm your choice. You will notice that all your previous data
    will be wiped out and the new data will be retrieved.

--------------------------------------------------------------------------------

## Wrapping up

You now have a running Baby Alligator instance. Closing the spreadsheet will not
lose any data nor stop the run of Baby Alligator.

The duration of insights retrieval is proportional to the number of locations:
The more locations you have the more insights you have to retrieve and the
longer it takes to complete the process, which could last a few days. Don't
worry! You can check the status of the insights in the insights tab and verify
that they are being gradually retrieved.

Remember to always click on <button>Collect GBP Data</button> and select *Reset
and use new config* to apply any configuration change.
