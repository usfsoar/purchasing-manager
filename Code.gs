/**
 * @typedef {Object} Column A data object describing a named column.
 * @prop {number} index 1-based index of the column in the sheet.
 * @prop {?string} name Name of the column.
 */

/**
 * Global options object.
 * @constant
 * @readonly
 * @global
 */
var OPTS = {
  /** Named Ranges throughout the spreadsheet. */
  NAMED_RANGES: {
    /** Range containing the email addresses of approved officers. 
     * 1 column, 5 rows (no header). */
    APPROVED_OFFICERS: 'ApprovedOfficers',
    /** Range containing the names of all project-specific sheets.
     * 1 column, 12 rows (no header). */
    PROJECT_SHEETS: 'ProjectSheets',
    PROJECT_NAMES_TO_SHEETS: 'ProjectNamesToSheets',
  },
  /** Custom Menu labels. */
  CUSTOM_MENU: {
    NAME: 'SOAR Purchasing',
  },
  /** The number of header rows in the project sheets. */
  NUM_HEADER_ROWS: 2,
  /** 
   * Relevant columns in the project sheets, as 1-based indexes.
   * @enum {Column}
   * @typedef {1|2|3|4|5|6|7|8|9|10|11|12|13|14|15|16|17|18|19|20|21|22} EnumITEM_COLUMNS
   */
  ITEM_COLUMNS: {
    STATUS: {index: 1, name: 'Status'},
    NAME: {index: 2, name: 'Name'},
    SUPPLIER: {index: 3, name: 'Supplier'},
    PRODUCT_NUM: {index: 4, name: 'Product Number'},
    LINK: {index: 5, name: 'Link'},
    UNIT_PRICE: {index: 6, name: 'Unit Price'},
    QUANTITY: {index: 7, name: 'Quantity'},
    SHIPPING: {index: 8, name: 'Shipping Price'},
    TOTAL_PRICE: {index: 9, name: 'Total Price'},
    CATEGORY: {index: 10, name: 'Category'},
    REQUEST_COMMENTS: {index: 11, name: 'Request Comments'},
    REQUEST_EMAIL: {index: 12, name: 'Requestor Email'},
    REQUEST_DATE: {index: 13, name: 'Request Date'},
    OFFICER_EMAIL: {index: 14, name: 'Financial Officer Email'},
    OFFICER_COMMENTS: {index: 15, name: 'Financial Officer Comments'},
    ACCOUNT: {index: 16, name: 'Purchasing Account'},
    REQUEST_ID: {index: 17, name: 'Request ID'},
    SUBMIT_DATE: {index: 18, name: 'Submit Date'},
    UPDATE_DATE: {index: 19, name: 'Update Date'},
    ARRIVE_DATE: {index: 20, name: 'Arrival Date'},
    RECIEVE_EMAIL: {index: 21, name: 'Recieve Date'},
    RECIEVE_DATE: {index: 22, name: 'Reciever Email'},
  },
  /** Options relating to the user interface. */
  UI: {
    /** Typical toast length in seconds. */
    TOAST_DURATION: 5,
    TOAST_TITLES: {
      ERROR: 'Error!',
      SUCCESS: 'Completed',
      WARNING: 'Alert!',
      INFO: 'Note',
    },
    SLACK_ID_PROMPT: 'Looks like this is your first time using the SOAR purchasing database. Please enter your Slack ID (NOT YOUR USERNAME!) found in your Slack profile, in the dropdown menu. For more details see https://drive.google.com/open?id=1Q1PleYhE1i0A5VFyjKqyLswom3NQuXcn.',
    FULL_NAME_PROMPT: 'Great, thank you! Please also enter your full name. You won\'t have to do this next time.'
  },
  /** Default values for items. */
  DEFAULT_VALUES: {
    ACCOUNT_NAME: getNamedRangeValues("Accounts")[0],
    CATEGORY: "Uncategorized"
  },
  /** Names of sheets in the Spreadsheet */
  SHEET_NAMES: {
    USERS: "Stored Slack IDs"
  },
  /** Slack API pieces */
  SLACK: {
    /** Webhook linked to the purchasing channel */
    WEBHOOKS: {
      PURCHASING: 'https://hooks.slack.com/services/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',
      RECIEVING: 'https://hooks.slack.com/services/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'
    }
  },
};

/**
 * @typedef {Object} Status A data object describing a possible item status.
 * @prop {string} text The textual name of the status.
 * @prop {string[]} allowedPrevious Allowed previous statuses as their text properties.
 * @prop {Object} actionText Menu item text.
 * @prop {?string} actionText.selected Menu item text for marking just selected.
 * @prop {?string} actionText.all Menu item text for marking all possible.
 * @prop {Object} slack Data for sending Slack notifications.
 * @prop {string} slack.messageTemplate Template for sending Slack messages.
 * @prop {string} slack.channelWebhook Webhook to send Slack messages to.
 * @prop {Object} columns Columns to input data into.
 * @prop {?Column} columns.user Column to input attribution email address into.
 * @prop {?Column} columns.date Column to input action date into.
 * @prop {?(Column[])} requiredColumns Optional required columns needed to perform actions.
 * @prop {?(Column[])} reccomendedColumns Optional reccomended columns desired to perform actions.
 */

/**
 * Information about each possible item status.
 * @constant
 * @readonly
 * @global
 * @enum {Status}
 */
var STATUSES_DATA = {
  CREATED: {
    text: '',
    allowedPrevious: [],
    actionText: {},
    slack: {
      messageTemplate: '',
      channelWebhook: OPTS.SLACK.WEBHOOKS.PURCHASING,
    },
    columns: {
      user: null,
      date: null,
    }
  },
  NEW: {
    text: 'New',
    allowedPrevious: ['', 'Awaiting Info'],
    actionText: {
      selected: 'Submit selected new items',
      all: 'Submit all new items',
    },
    slack: {
      messageTemplate: '',
      channelWebhook: OPTS.SLACK.WEBHOOKS.PURCHASING,
    },
    columns: {
      user: OPTS.ITEM_COLUMNS.REQUEST_EMAIL,
      date: OPTS.ITEM_COLUMNS.REQUEST_DATE,
    },
    requiredColumns: [
      OPTS.ITEM_COLUMNS.NAME,
      OPTS.ITEM_COLUMNS.SUPPLIER,
      OPTS.ITEM_COLUMNS.LINK,
      OPTS.ITEM_COLUMNS.PRICE,
      OPTS.ITEM_COLUMNS.QUANTITY,
    ]
  },
  SUBMITTED: {
    text: 'Submitted',
    allowedPrevious: ['New'],
    actionText: {
      selected: 'Mark selected items as submitted',
    },
    slack: {
      messageTemplate: '',
      channelWebhook: OPTS.SLACK.WEBHOOKS.PURCHASING,
    },
    columns: {
      user: OPTS.ITEM_COLUMNS.OFFICER_EMAIL,
      date: OPTS.ITEM_COLUMNS.SUBMIT_DATE,
    },
    reccomendedColumns: [
      OPTS.ITEM_COLUMNS.ACCOUNT,
      OPTS.ITEM_COLUMNS.CATEGORY,
    ]
  },
  APPROVED: {
    text: 'Approved',
    allowedPrevious: ['Submitted'],
    actionText: {
      selected: 'Mark selected items as approved',
    },
    slack: {
      messageTemplate: '',
      channelWebhook: OPTS.SLACK.WEBHOOKS.PURCHASING,
    },
    columns: {
      user: null,
      date: OPTS.ITEM_COLUMNS.UPDATE_DATE,
    }
  },
  AWAITING_PICKUP: {
    text: 'Awaiting Pickup',
    allowedPrevious: ['Submitted', 'Approved'],
    actionText: {
      selected: 'Mark selected items as awaiting pickup',
    },
    slack: {
      messageTemplate: '',
      channelWebhook: OPTS.SLACK.WEBHOOKS.RECIEVING,
    },
    columns: {
      user: null,
      date: OPTS.ITEM_COLUMNS.ARRIVE_DATE,
    }
  },
  RECIEVED: {
    text: 'Recieved',
    allowedPrevious: ['Awaiting Pickup'],
    actionText: {
      selected: 'Mark selected items as recieved',
    },
    slack: {
      messageTemplate: '',
      channelWebhook: OPTS.SLACK.WEBHOOKS.RECIEVING,
    },
    columns: {
      user: OPTS.ITEM_COLUMNS.RECIEVE_EMAIL,
      date: OPTS.ITEM_COLUMNS.RECIEVE_DATE,
    }
  },
  DENIED: {
    text: 'Denied',
    allowedPrevious: ['New', 'Submitted', 'Approved', 'Awaiting Info'],
    actionText: {
      selected: 'Deny selected items',
    },
    slack: {
      messageTemplate: '',
      channelWebhook: OPTS.SLACK.WEBHOOKS.PURCHASING,
    },
    columns: {
      user: OPTS.ITEM_COLUMNS.UPDATE_DATE,
      date: OPTS.ITEM_COLUMNS.OFFICER_EMAIL,
    },
    requiredColumns: [
      OPTS.ITEM_COLUMNS.OFFICER_COMMENTS
    ]
  },
  AWAITING_INFO: {
    text: 'Awating Info',
    allowedPrevious: ['New', 'Submitted', 'Denied', 'Approved', 'Recieved'],
    actionText: {
      selected: 'Request more information for selected items'
    },
    slack: {
      messageTemplate: '',
      channelWebhook: OPTS.SLACK.WEBHOOKS.PURCHASING,
    },
    columns: {
      user: OPTS.ITEM_COLUMNS.UPDATE_DATE,
      date: OPTS.ITEM_COLUMNS.OFFICER_EMAIL,
    },
    requiredColumns: [
      OPTS.ITEM_COLUMNS.OFFICER_COMMENTS
    ]
  }
};

/**
 * Global that represents whether the user is authorized as a financial officer.
 */
function onOpen() {
  buildAndAddCustomMenu();
}

/**
 * Verify whether or not the current user is an approved financial officer.
 * After first run, uses cache to avoid having to pull the range again.
 * @returns {boolean} true if the current user is approved.
 */
var verifyFinancialOfficer = (function() {
  var cache = {
    verified: /** @type {?boolean} */  (null),
  };

  return function() {
    if(cache.verified === null) {
      cache.verified = false;
      var email = Session.getActiveUser().getEmail();
      
      if(email !== '' 
        && getNamedRangeValues(OPTS.NAMED_RANGES.APPROVED_OFFICERS)
            .indexOf(email) !== -1
      ) {
        cache.verified = true;
      }
    }

    return /** @type {boolean} */ (cache.verified);
  };
})();

/**
 * Get the list of non-empty values in the named range.
 * @returns {string[]} Unordered array of values, flattened into a 1-dimensional
 * array.
 */
function getNamedRangeValues(rangeName) {
  var valuesGrid = SpreadsheetApp
      .getActiveSpreadsheet()
      .getRange(rangeName)
      .getValues();

  // Flatten and remove empty values
  var valuesArray = [].concat.apply([], valuesGrid)
      .filter(function (value) {
        return value !== '';
      });
  
  return valuesArray;
}

/**
 * Build the custom SOAR Purchasing menu and add it to the user interface.
 */
function buildAndAddCustomMenu() {
  // Use yourFunction.name because it requires a string and this is a little
  // more reusable than just hardcoding the name

  var customMenu = SpreadsheetApp.getUi()
    .createMenu(OPTS.CUSTOM_MENU.NAME)
    .addItem(STATUSES_DATA.NEW.ACTION_TEXT.ALL, markAllNew.name)
    .addItem(STATUSES_DATA.NEW.ACTION_TEXT.SELECTED, markSelectedNew.name)
    .addSeparator()
    .addItem(STATUSES_DATA.RECIEVED.ACTION_TEXT.SELECTED, markSelectedRecieved.name);

  if (verifyFinancialOfficer()) {
    customMenu
      .addSeparator()
      .addItem(STATUSES_DATA.SUBMITTED.ACTION_TEXT.SELECTED, markSelectedSubmitted.name)
      .addItem(STATUSES_DATA.APPROVED.ACTION_TEXT.SELECTED, markSelectedApproved.name)
      .addItem(STATUSES_DATA.AWAITING_PICKUP.ACTION_TEXT.SELECTED, markSelectedAwaitingPickup.name)
      .addSeparator()
      .addItem(STATUSES_DATA.AWAITING_INFO.ACTION_TEXT.SELECTED, markSelectedAwaitingInfo.name)
      .addItem(STATUSES_DATA.DENIED.ACTION_TEXT.SELECTED, markSelectedDenied.name);
  }

  customMenu.addToUi();
}

/** Show the user an error message. */
function errorNotification(message) {
  SpreadsheetApp.getActiveSpreadsheet().toast(
    message,
    OPTS.UI.TOAST_TITLES.ERROR,
    OPTS.UI.TOAST_DURATION);
}

/** Show the user a warning message. */
function warnNotification(message) {
  SpreadsheetApp.getActiveSpreadsheet().toast(
    message,
    OPTS.UI.TOAST_TITLES.WARNING,
    OPTS.UI.TOAST_DURATION);
}

/** Show a log message and log it. For debugging. */
function log(message) {
  Logger.log(message);
  SpreadsheetApp.getActiveSpreadsheet().toast(
    message,
    OPTS.UI.TOAST_TITLES.INFO,
    OPTS.UI.TOAST_DURATION);
}

/** Show the user a success message. */
function successNotification(message) {
  SpreadsheetApp.getActiveSpreadsheet().toast(
    message,
    OPTS.UI.TOAST_TITLES.SUCCESS,
    OPTS.UI.TOAST_DURATION);
}

/**
 * Returns the current user's slack ID from the storage sheet, or prompts for it,
 * or returns it from the local cache if it's been asked before.
 * @returns {string} The user's Slack ID.
 */
var getCurrentUserSlackId = (function() {
  var cache = {
    slackId: /** @type {?string} */  (null),
  };

  return function() {
    if(cache.slackId === null) {
      var userSheet = SpreadsheetApp
          .getActiveSpreadsheet()
          .getSheetByName(OPTS.SHEET_NAMES.USERS);
      var userData = userSheet.getDataRange().getValues();
      var email = Session.getActiveUser().getEmail();
      
      for(var i = 1; i < userData.length; i++) {
        if(userData[i][0] === email) {
          cache.slackId = userData[i][1];
          return cache.slackId;
        }
      }

      cache.slackId = ui.prompt(OPTS.UI.SLACK_ID_PROMPT);
      userSheet.appendRow([email, cache.slackId]);
      return cache.slackId;
    }

    return /** @type {string} */ (cache.slackId);
  };
})();

/**
 * Returns the Slack ID that matches the email address provided.
 * @param {string} email
 * @returns {?string} The Slack ID or null if no match.
 */
function getSlackIdByEmail(email) {
  var userSheet = SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName(OPTS.SHEET_NAMES.USERS);
  var userData = userSheet.getDataRange().getValues();
  
  for(var i = 1; i < userData.length; i++) {
    if(userData[i][0] === email) {
      return userData[i][1];
    }
  }

  return null;
}

/**
 * Returns the current user's slack ID from the storage sheet, or prompts for it,
 * or returns it from the local cache if it's been asked before.
 * @returns {string} The user's Slack ID.
 */
var getCurrentUserFullName = (function() {
  var cache = {
    fullName: /** @type {?string} */  (null),
  };

  return function() {
    if(cache.fullName === null) {
      var userSheet = SpreadsheetApp
          .getActiveSpreadsheet()
          .getSheetByName(OPTS.SHEET_NAMES.USERS);
      var userData = userSheet.getDataRange().getValues();
      var email = Session.getActiveUser().getEmail();
      
      for(var i = 1; i < userData.length; i++) {
        if(userData[i][0] === email) {
          cache.fullName = userData[i][2];
          return cache.fullName;
        }
      }

      cache.fullName = ui.prompt(OPTS.UI.FULL_NAME_PROMPT);
      userSheet.appendRow([email, cache.fullName]);
      return cache.fullName;
    }

    return /** @type {string} */ (cache.fullName);
  };
})();

/**
 * Turn an array into a human-readable list.
 * @param {string[]} listArray Array to make a list from.
 * @param {string} [conjunction='and'] Conjunction to use at the end of the list.
 * @param {boolean} [noOxfordComma] If true, won't add an Oxford comma.
 * @returns {string} A nicely formatted list, ie: 'One, Two, and Three'.
 */
function makeListFromArray(listArray, conjunction, noOxfordComma) {
  /**
   * The oxford comma, or an empty string if `noOxfordComma` is true or the array
   * is too short.
   */
  var oxfordComma = (noOxfordComma || (listArray.length <= 2)) ? '' : ',';
  conjunction = conjunction === undefined ? 'and' : conjunction;

  return listArray.reduce(function(finalString, listItem, index) {
    switch(index) {
      case 0:
        return listItem;
      case listArray.length - 1:
        return finalString + oxfordComma + ' ' + conjunction + ' ' + listItem;
      default:
        return finalString + ', ' + listItem;
    }
  });
}

/**
 * Truncates the string if it's longer than `chars` and adds "..." to the end.
 * @param {string} longString The string to shorten.
 * @param {number} chars The maximum number of characters in the final string.
 * @returns {string} The truncated string.
 */
function truncateString(longString, chars) {
  if(longString.length > chars) {
    return longString.slice(0, chars - 4) + "...";
  }
  return longString;
}

/**
 * Get the ranges of all the currently selected rows in the active sheet.
 * @returns {GoogleAppsScript.Spreadsheet.Range[]} Array of selected ranges,
 * expanded to cover entire width of data in the sheet.
 */
function getSelectedRows() {
  var activeSheet = SpreadsheetApp.getActiveSheet();
  var selections = activeSheet.getSelection()
      .getActiveRangeList().getRanges();
  var lastColumn = activeSheet.getLastColumn();

  // Expand selections to width of spreadsheet
  var expandedSelections = selections.map(function(selectionRange) {
    var selectionStartRow = selectionRange.getRow();
    var selectionNumRows = selectionRange.getNumRows();
    if(selectionStartRow === 1) {
      selectionStartRow++;
      selectionNumRows--;
    }
    if(selectionStartRow === 2) {
      selectionStartRow++;
      selectionNumRows--;
    }
    return activeSheet.getRange(selectionStartRow, 1, selectionNumRows, lastColumn);
  });

  return expandedSelections;
}

/**
 * Get the range of all data in the active sheet.
 * @returns {GoogleAppsScript.Spreadsheet.Range[]} Array with one Range.
 */
function getAllRows() {
  var activeSheet = SpreadsheetApp.getActiveSheet();
  var lastColumnInSheet = activeSheet.getLastColumn();
  var firstNonHeaderRow = OPTS.NUM_HEADER_ROWS + 1;

  var nameColumnValues = getColumnRange(OPTS.ITEM_COLUMNS.NAME).getValues();

  /** The number of the last row in the sheet that has a value for Name. */
  var lastRowWithData = firstNonHeaderRow;

  nameColumnValues.forEach(function(name, index) {
    if(name.toString().trim() !== '') lastRowWithData = index + firstNonHeaderRow;
  });

  var numNonHeaderRowsWithData = lastRowWithData - OPTS.NUM_HEADER_ROWS;

  return [
    activeSheet
      .getRange(firstNonHeaderRow, 1, numNonHeaderRowsWithData, lastColumnInSheet)
  ];
}

/**
 * Checks if the current sheet is in the list of project sheets. If not,
 * shows a message in the UI and returns false.
 * @returns {boolean} True if a project sheet is active.
 */
function checkIfProjectSheet() {
  var currentSheetName = SpreadsheetApp.getActiveSheet().getName();
  var projectSheetNames = getNamedRangeValues(OPTS.NAMED_RANGES.PROJECT_SHEETS);

  if(projectSheetNames.indexOf(currentSheetName) === -1) {
    errorNotification('This action may only be performed in a project sheet');
    return false;
  }

  return true;
}

/**
 * Get the range of an entire column in the active project sheet, minus headers.
 * @param {number} columnNumber The number of the column to get.
 * @returns {GoogleAppsScript.Spreadsheet.Range} The range of the column.
 */
function getColumnRange(columnNumber) {
  var activeSheet = SpreadsheetApp.getActiveSheet();
  var firstNonHeaderRow = OPTS.NUM_HEADER_ROWS + 1;
  var numNonHeaderRows = activeSheet.getLastRow() - OPTS.NUM_HEADER_ROWS;

  return activeSheet
    .getRange(firstNonHeaderRow, columnNumber, numNonHeaderRows, 1);
}

/**
 * Mark all of the items in the currently selected rows as `newStatus` if they
 * are currently one of the allowed previous statuses, and also fill in the date and
 * attribution columns.
 * @param {Status} newStatus The object representing the status to change the
 * selected items to.
 * @param {boolean} [markAll] If true, mark all possible rows, else mark selected.
 * @returns {void}
 */
function markItems(newStatus, markAll) {
  if(!checkIfProjectSheet() || !getCurrentUserFullName() || !getCurrentUserSlackId()) return;

  /** All the ranges in the sheet if `markAll` is set, else just the selected. */
  var selectedRanges = markAll ? getAllRows() : getSelectedRows();
  var numMarked = 0;
  var currentUser = Session.getActiveUser().getEmail();
  var currentDate = new Date();

  // We would filter out all the rows with disallowed current statuses here, 
  // rather than skipping them in both of these loops, but that would require
  // modifying the ranges, which is much more time-intensive than just skipping.

  // Loop through every row in every range and validate them, throwing the flag
  // if any are invalid. This is a separate loop from the actual modification
  // loop because if it were the same, we could modify some of the data before
  // seeing that other data is invalid, which would not be the expected behavior.
  // No need to alert the user on fail; validateRow will do that itself.
  var allRowsValid = true;

  selectionsLoop: for(var i = 0; i < selectedRanges.length; i++) {
    var rangeValues = selectedRanges[i].getValues();

    for(var j = 0; j < selectedRanges.length; j++) {
      /** Array of row values. */
      var row = rangeValues[j];
      // If current status is not in allowed statuses, don't verify, just skip
      // minus 1 to convert from 1-based Sheets column number to 0-based array index
      if(!isCurrentStatusAllowed(row[OPTS.ITEM_COLUMNS.STATUS - 1].toString(), newStatus)) continue;

      // Otherwise validate. If a single row is invalid, quit both loops
      if(!validateRow(row, newStatus)) {
        allRowsValid = false;
        break selectionsLoop;
      }
    }
  }

  if(allRowsValid) {
    // Cache the entire columns, to avoid making dozens of calls to the server
    var statusColumn = getColumnRange(OPTS.ITEM_COLUMNS.STATUS);
    var statusColumnValues = statusColumn.getValues();

    var userColumn, dateColumn, userColumnValues, dateColumnValues;
    if(userColumnNum !== null) {
      userColumn = getColumnRange(userColumnNum);
      userColumnValues = userColumn.getValues();
    }
    if(dateColumnNum !== null) {
      dateColumn = getColumnRange(dateColumnNum);
      dateColumnValues = dateColumn.getValues();
    }

    var accountColumn, categoryColumn, accountColumnValues, categoryColumnValues;
    var fillDefaultValues = newStatus === OPTS.STATUSES.SUBMITTED
        || newStatus === OPTS.STATUSES.APPROVED
        || newStatus === OPTS.STATUSES.AWAITING_PICKUP;
    if(fillDefaultValues) {
      accountColumn = getColumnRange(OPTS.ITEM_COLUMNS.ACCOUNT);
      categoryColumn = getColumnRange(OPTS.ITEM_COLUMNS.CATEGORY);
      accountColumnValues = accountColumn.getValues();
      categoryColumnValues = categoryColumn.getValues();
    }

    // Loop through the ranges
    for(var k = 0; k < selectedRanges.length; k++) {
      var range = selectedRanges[k];
      var rangeStartIndex = range.getRow() - 1;

      // Loop through the rows in the range
      for (var l = 0; l < range.getNumRows(); l++) {
        /** The index (not number) of the current row in the spreadsheet. */
        var currentSheetRowIndex = rangeStartIndex + l;
        /** 
         * The index of the current value row in the spreadsheet, with the first
         * row after the headers being 0. 
         */
        var currentValuesRowIndex = currentSheetRowIndex - OPTS.NUM_HEADER_ROWS;

        // If this row's status is not in allowed statuses, don't verify, just skip
        if(!isCurrentStatusAllowed(statusColumnValues[currentValuesRowIndex][0].toString(), newStatus)) continue;

        // Update values in local cache
        // These ranges don't include the header, so 0 in the range is
        // OPTS.NUM_HEADER_ROWS in the spreadsheet
        statusColumnValues[currentValuesRowIndex][0] = newStatus;

        if(userColumnNum !== null) {
          userColumnValues[currentValuesRowIndex][0] = currentUser;
        }
        if(dateColumnNum !== null) {
          dateColumnValues[currentValuesRowIndex][0] = currentDate;
        }

        if(fillDefaultValues) {
          if(accountColumnValues[currentValuesRowIndex][0].toString() === '') {
            accountColumnValues[currentValuesRowIndex][0] =
                OPTS.DEFAULT_VALUES.ACCOUNT_NAME;
          }
          if(categoryColumnValues[currentValuesRowIndex][0].toString() === '') {
            categoryColumnValues[currentValuesRowIndex][0] =
                OPTS.DEFAULT_VALUES.CATEGORY;
          }
        }

        numMarked++;
      }
    }

    // Write the cached values
    statusColumn.setValues(statusColumnValues);

    if(userColumnNum !== null) userColumn.setValues(userColumnValues);
    if(dateColumnNum !== null) dateColumn.setValues(dateColumnValues);

    if(fillDefaultValues) {
      accountColumn.setValues(accountColumnValues);
      categoryColumn.setValues(categoryColumnValues);
    }

    /** All of the possible 'from' statuses, but with double quotes around them. */
    var quotedFromStatuses = OPTS.ALLOWED_PREV_STATUSES[newStatus].map(wrapInDoubleQuotes);

    SpreadsheetApp.getActiveSpreadsheet().toast(
        numMarked + ' items marked from '
          + makeListFromArray(quotedFromStatuses, 'or')
          + ' to "' + newStatus + '."',
        'Completed',
        OPTS.UI.TOAST_DURATION);
  }
}

/**
 * Check if the current status of the row is in the valid statuses list.
 * @param {string} rowCurrentStatus The current status of the row.
 * @param {Status} newStatus The status object to check for changing to.
 * @returns {boolean} True if the current status of the row allows it to be
 * changed to the newStatus.
 */
function isCurrentStatusAllowed(rowCurrentStatus, newStatus) {
  var currentStatus = rowCurrentStatus.trim();
  return newStatus.allowedPrevious.indexOf(currentStatus) !== -1;
}

/**
 * Wrap a string with double quotes.
 * @param {string} stringToWrap The string to be wrapped in quotes.
 * @returns {string} `stringToWrap`, but with quotes around it. If it's an
 * empty string, returns a wrapped space character.
 */
function wrapInDoubleQuotes(stringToWrap) {
  if(stringToWrap === '') stringToWrap = ' ';
  return '"' + stringToWrap + '"';
}

/**
 * Check if the given row has data and the data is valid for the desired operation.
 * If the validation fails, alerts the user. Does not check row statuses;
 * rows with incorrect statuses are skipped silently.
 * @param {(string|number|Date)[]} rowValues rowValues The current data for the row.
 * @param {string} newStatus The new status of the row for testing against.
 * @returns {boolean} True if the row is valid and can be submitted.
 */
function validateRow(rowValues, newStatus) {
  switch(newStatus.text) {
    case OPTS.STATUSES.NEW:
      return validateNewRow(rowValues);

    case OPTS.STATUSES.DENIED:
    case OPTS.STATUSES.AWAITING_INFO:
      return validateDenyingOrAwaitingInfoRow(rowValues);

    case OPTS.STATUSES.SUBMITTED:
      return validateSubmittingRow(rowValues);

    default:
      return true;
  }
}

/**
 * Check if the given row has data and the data is valid for submitting as a new
 * item. If the validation fails, alerts the user. Does not check row statuses;
 * rows with incorrect statuses are skipped silently.
 * @param {(string|number|Date)[]} rowValues The current data for the row.
 * @returns {boolean} True if the row is valid and can be submitted.
 * @todo Add default category
 */
function validateNewRow(rowValues) {
  var missingString = 'Could not submit items because one or more items is missing a';
  if(rowValues[OPTS.ITEM_COLUMNS.NAME - 1] === '') {
    errorNotification(missingString + ' name.');
    return false;
  }
  if(rowValues[OPTS.ITEM_COLUMNS.SUPPLIER - 1] === '') {
    errorNotification(missingString + ' supplier.');
    return false;
  }
  if(rowValues[OPTS.ITEM_COLUMNS.LINK - 1] === '') {
    errorNotification(missingString + ' link.');
    return false;
  }
  if(rowValues[OPTS.ITEM_COLUMNS.UNIT_PRICE - 1] === '') {
    errorNotification(missingString + ' price.');
    return false;
  }
  if(!rowValues[OPTS.ITEM_COLUMNS.QUANTITY - 1]) {
    errorNotification(missingString + ' quantity.');
    return false;
  }
  return true;
}

/**
 * Check if the given row has data and the data is valid for submitting to SG.
 * If the validation fails, alerts the user. Does not check row statuses;
 * rows with incorrect statuses are skipped silently.
 * @param {(string|number|Date)[]} rowValues The current data for the row.
 * @returns {boolean} True if the row is valid and can be submitted.
 */
function validateSubmittingRow(rowValues) {
  if(rowValues[OPTS.ITEM_COLUMNS.ACCOUNT - 1] === '') {
    warnNotification('One or more items are missing a value for account. Defaulted to "' + OPTS.DEFAULT_ACCOUNT_NAME + '".');
  }
  if(rowValues[OPTS.ITEM_COLUMNS.REQUEST_ID - 1] === '') {
    warnNotification('One or more items are missing a request ID. They will be marked as submitted anyway.');
  }
  return true;
}

/**
 * Check if the given row has data and the data is valid for marking as rejected
 * or as awaiting information.
 * If the validation fails, alerts the user. Does not check row statuses;
 * rows with incorrect statuses are skipped silently.
 * @param {(string|number|Date)[]} rowValues The current data for the row.
 * @returns {boolean} True if the row is valid and can be marked.
 */
function validateDenyingOrAwaitingInfoRow(rowValues) {
  if(rowValues[OPTS.ITEM_COLUMNS.OFFICER_COMMENTS - 1] === '') {
    errorNotification('Could not perform this action because one or more items is missing a comment, which is required.');
    return false;
  }
  return true;
}

/**
 * Send a message to the Slack channel.
 * @param {string} message The message to send.
 */
function sendSlackMessage(message, webhook) {
  var messageData = {
    text: message
  };
  var requestOptions = {
    method: 'post',
    payload: JSON.stringify(messageData),
    contentType: 'application/json'
  };
  UrlFetchApp.fetch(webhook, requestOptions);
}

/**
 * 
 * @param newStatus 
 * @param numMarked 
 * @param relatedPeople  COULD BE FINANCIAL OFFICERS!
 * @param projectSheetName 
 */
function slackNotifyItems(newStatus, numMarked, relatedPeople, projectSheetName) {
  var userFullName = getCurrentUserFullName();
  var targetSlackIds = makeListFromArray(relatedPeople.map(function (email) {
    return '<@' + getSlackIdByEmail + '>';
  }), '');
  var itemsWord = numMarked > 1 ? 'items' : 'item';
  var projectName = getProjectNameFromSheetName(projectSheetName);
  var targetWebHook = OPTS.SLACK.WEBHOOKS.PURCHASING;

  var message = '';
  switch(newStatus) {
    case OPTS.STATUSES.NEW:
      message = 
  }

}

/**
 * 
 * @param sheetName 
 */
function getProjectNameFromSheetName(sheetName) {
  var projectsData = SpreadsheetApp
    .getActiveSpreadsheet()
    .getRange(OPTS.NAMED_RANGES.PROJECT_NAMES_TO_SHEETS)
    .getValues();

  for(var i = 0; i < projectsData.length; i++) {
    if(projectsData[i][1] === sheetName) return projectsData[i][0];
  }

  return '_Error: Project Not Found_';
}

/** Mark the selected items in the sheet as new. */
function markSelectedNew() {
  markItems(
      OPTS.STATUSES.NEW,
      OPTS.ITEM_COLUMNS.REQUEST_EMAIL,
      OPTS.ITEM_COLUMNS.REQUEST_DATE,
      false);
}

/** Mark all possible items in the sheet as new. */
function markAllNew() {
  markItems(
      OPTS.STATUSES.NEW,
      OPTS.ITEM_COLUMNS.REQUEST_EMAIL,
      OPTS.ITEM_COLUMNS.REQUEST_DATE,
      true);
}

/** Mark selected items in the sheet as recieved. */
function markSelectedRecieved() {
  markItems(
      OPTS.STATUSES.RECIEVED,
      OPTS.ITEM_COLUMNS.RECIEVE_EMAIL,
      OPTS.ITEM_COLUMNS.RECIEVE_DATE,
      false);
}

/** Mark selected items in the sheet as submitted. */
function markSelectedSubmitted() {
  markItems(
      OPTS.STATUSES.SUBMITTED,
      OPTS.ITEM_COLUMNS.OFFICER_EMAIL,
      OPTS.ITEM_COLUMNS.SUBMIT_DATE,
      false);
}

/** Mark selected items in the sheet as approved. */
function markSelectedApproved() {
  markItems(
      OPTS.STATUSES.APPROVED,
      null,
      OPTS.ITEM_COLUMNS.UPDATE_DATE,
      false);
}

/** Mark selected items in the sheet as arrived / awaiting pickup. */
function markSelectedAwaitingPickup() {
  markItems(
      OPTS.STATUSES.AWAITING_PICKUP,
      null,
      OPTS.ITEM_COLUMNS.ARRIVE_DATE,
      false);
}

/** Mark selected items in the sheet as awaiting info. */
function markSelectedAwaitingInfo() {
  markItems(
      OPTS.STATUSES.AWAITING_INFO,
      OPTS.ITEM_COLUMNS.OFFICER_EMAIL,
      OPTS.ITEM_COLUMNS.UPDATE_DATE,
      false);
}

/** Mark selected items in the sheet as denied. */
function markSelectedDenied() {
  markItems(
      OPTS.STATUSES.DENIED,
      OPTS.ITEM_COLUMNS.OFFICER_EMAIL,
      OPTS.ITEM_COLUMNS.UPDATE_DATE,
      false);
}