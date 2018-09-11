/**
 * @typedef {Object} Column A data object describing a named column.
 * @prop {number} index 1-based index of the column in the sheet.
 * @prop {?string} name Name of the column.
 */

/**
 * Secret values. DO NOT PUSH TO GITHUB.
 * @constant
 * @readonly
 * @global
 * @type {Object}
 */
var SECRET_OPTS = getSecretOpts();

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
    /** Range containing the YES or NO values for whether to notify the
     * financial officer with the same index. 1 column, 5 rows (no header). */
    NOTIFY_APPROVED_OFFICERS: 'NotifyApprovedOfficers',
    /** Range containing the names of all project-specific sheets.
     * 1 column, 12 rows (no header). */
    PROJECT_SHEETS: 'ProjectSheets',
    PROJECT_NAMES_TO_SHEETS: 'ProjectNamesToSheets',
    STATUSES: 'Statuses',
  },
  /** Custom Menu info. */
  CUSTOM_MENU: {
    NAME: 'SOAR Purchasing',
  },
  FAST_FORWARD_MENU: {
    NAME: 'Fast-Forward'
  },
  /** The number of header rows in the project sheets. */
  NUM_HEADER_ROWS: 2,
  /**
   * Relevant columns in the project sheets, as 1-based indexes.
   * @enum {Column}
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
    SLACK_ID_PROMPT: 'Looks like this is your first time using the SOAR purchasing database. Please enter your Slack Member ID # (NOT your username!) found in your Slack profile, in the dropdown menu. For more details, see detailed instructions at:\nhttps://drive.google.com/open?id=1Q1PleYhE1i0A5VFyjKqyLswom3NQuXcn.',
    FULL_NAME_PROMPT: 'Great, thank you! Please also enter your full name. You won\'t have to do this next time.'
  },
  /** Default values for items. */
  DEFAULT_VALUES: {
    ACCOUNT_NAME: getNamedRangeValues("Accounts")[0],
    CATEGORY: "Uncategorized"
  },
  /** Names of sheets in the Spreadsheet */
  SHEET_NAMES: {
    USERS: "Users"
  },
  /** Slack API pieces */
  SLACK: {
    KYBER_TASK_REACTION: ':ballot_box_with_check:',
    /**
     * Possible cases for target users to tag in messages.
     * @enum {string}
     * @typedef {'CHANNEL'|'REQUESTORS'|'OFFICERS'} EnumTARGET_USERS
     */
    TARGET_USERS: {
      /** The entire channel. */
      CHANNEL: 'CHANNEL',
      /**
       * Just the people who requested said items (can be multiple if multiple)
       * items are affected.
       */
      REQUESTORS: 'REQUESTORS',
      /** Just all the listed Financial Officers. */
      OFFICERS: 'OFFICERS',
    }
  },
  /** Number of adjacent officer columns in the project sheets. */
  NUM_OFFICER_COLS: 7
};

/**
 * @typedef {Object} Status A data object describing a possible item status.
 * @prop {string} text The textual name of the status.
 * @prop {string[]} allowedPrevious Allowed previous statuses as their text properties.
 * @prop {Object} actionText Menu item text.
 * @prop {?string} actionText.selected Menu item text for marking just selected.
 * @prop {?string} actionText.all Menu item text for marking all possible.
 * @prop {?string} actionText.fastForward Menu item text for fast-forwarding items. 
 * @prop {Object} slack Data for sending Slack notifications.
 * @prop {string[]} slack.messageTemplates Templates for sending Slack messages.
 * Will send a Slack message per string. Will replace {emoji} with the emoji,
 * {userTags} with the target user tags, {userFullName} with full name of
 * submitter, {numMarked} with the number of items marked, {projectName} with
 * the name of the project, and {projectSheetUrl} with the link to the project
 * sheet.
 * @prop {string[]} slack.channelWebhooks Webhooks to send Slack messages to.
 * Will only tag targetUsers in the first channel provided, to avoid annoying.
 * @prop {string} slack.emoji Emoji to send with slack message.
 * @prop {EnumTARGET_USERS} slack.targetUsers String representing a user group
 * to tag in Slack messages (only in the first channel the message is sent to).
 * @prop {Object} columns Columns to input data into.
 * @prop {?Column} columns.user Column to input attribution email address into.
 * @prop {?Column} columns.date Column to input action date into.
 * @prop {Object} fastForwardColumns Columns to auto-fill upon fast-forwarding.
 * @prop {?(Column[])} fastForwardColumns.user Column to input attribution email address into.
 * @prop {?(Column[])} fastForwardColumns.date Column to input action date into.
 * @prop {?(Column[])} requiredColumns Optional required columns needed to perform actions.
 * @prop {?(Column[])} reccomendedColumns Optional reccomended columns desired to perform actions.
 * @prop {?boolean} fillInDefaults If true, will fill default values for Account
 * and Cateegory when those are applied.
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
    actionText: {
    },
    slack: {},
    columns: {
      user: null,
      date: null,
    },
    officersOnly: false,
  },
  NEW: {
    text: 'New',
    allowedPrevious: ['', 'Awaiting Info'],
    actionText: {
      fastForward: 'New',
      selected: 'Submit selected new items',
      all: 'Submit all new items',
    },
    slack: {
      emoji: ':large_blue_circle:',
      targetUsers: OPTS.SLACK.TARGET_USERS.OFFICERS,
      messageTemplates: [
        '{userTags} React with ' + OPTS.SLACK.KYBER_TASK_REACTION + ' to the following message if you\'re going to review / submit these items:',
        '{emoji} {userFullName} has submitted {numMarked} new item{plural} to be purchased for {projectName}. *<{projectSheetUrl}|View Items>*'
      ],
      channelWebhooks: [SECRET_OPTS.SLACK.WEBHOOKS.PURCHASING],
    },
    columns: {
      user: OPTS.ITEM_COLUMNS.REQUEST_EMAIL,
      date: OPTS.ITEM_COLUMNS.REQUEST_DATE,
    },
    fastForwardColumns: {
      user: [],
      date: [],
    },
    requiredColumns: [
      OPTS.ITEM_COLUMNS.NAME,
      OPTS.ITEM_COLUMNS.SUPPLIER,
      OPTS.ITEM_COLUMNS.LINK,
      OPTS.ITEM_COLUMNS.UNIT_PRICE,
      OPTS.ITEM_COLUMNS.QUANTITY,
    ],
    officersOnly: false,
  },
  SUBMITTED: {
    text: 'Submitted',
    allowedPrevious: ['New'],
    actionText: {
      fastForward: 'Submitted',
      selected: 'Mark selected items as submitted',
    },
    slack: {
      emoji: ':white_circle:',
      targetUsers: OPTS.SLACK.TARGET_USERS.REQUESTORS,
      messageTemplates: [
        '{emoji} {userTags} {userFullName} marked {numMarked} item{plural} for {projectName} as *submitted* to Student Government. *<{projectSheetUrl}|View Items>*'
      ],
      channelWebhooks: [SECRET_OPTS.SLACK.WEBHOOKS.PURCHASING],
    },
    columns: {
      user: OPTS.ITEM_COLUMNS.OFFICER_EMAIL,
      date: OPTS.ITEM_COLUMNS.SUBMIT_DATE,
    },
    fastForwardColumns: {
      user: [
        OPTS.ITEM_COLUMNS.REQUEST_EMAIL,
      ],
      date: [
        OPTS.ITEM_COLUMNS.REQUEST_DATE,
      ],
    },
    reccomendedColumns: [
      OPTS.ITEM_COLUMNS.ACCOUNT,
      OPTS.ITEM_COLUMNS.CATEGORY,
    ],
    fillInDefaults: true,
    officersOnly: true,
  },
  APPROVED: {
    text: 'Ordered',
    allowedPrevious: ['Submitted'],
    actionText: {
      fastForward: 'Ordered',
      selected: 'Mark selected items as ordered by SBS',
    },
    slack: {
      emoji: ':white_circle:',
      targetUsers: OPTS.SLACK.TARGET_USERS.REQUESTORS,
      messageTemplates: [
        '{emoji} {userTags} {userFullName} marked {numMarked} item{plural} for {projectName} as *ordered* by Student Government. *<{projectSheetUrl}|View Items>*'
      ],
      channelWebhooks: [SECRET_OPTS.SLACK.WEBHOOKS.PURCHASING],
    },
    columns: {
      user: null,
      date: OPTS.ITEM_COLUMNS.UPDATE_DATE,
    },
    fastForwardColumns: {
      user: [
        OPTS.ITEM_COLUMNS.REQUEST_EMAIL,
        OPTS.ITEM_COLUMNS.OFFICER_EMAIL
      ],
      date: [
        OPTS.ITEM_COLUMNS.REQUEST_DATE,
        OPTS.ITEM_COLUMNS.SUBMIT_DATE,
      ],
    },
    fillInDefaults: true,
    officersOnly: true,
  },
  AWAITING_PICKUP: {
    text: 'Awaiting Pickup',
    allowedPrevious: ['Submitted', 'Ordered'],
    actionText: {
      fastForward: 'Awaiting Pickup',
      selected: 'Mark selected items as awaiting pickup',
    },
    slack: {
      emoji: ':large_blue_circle:',
      targetUsers: OPTS.SLACK.TARGET_USERS.CHANNEL,
      messageTemplates: [
        '{userTags} React with ' + OPTS.SLACK.KYBER_TASK_REACTION + ' to the following message if you\'re going to pickup these items:',
        '{emoji} {userFullName} marked {numMarked} item{plural} for {projectName} as awaiting pickup from Student Business Services in MSC4300. *<{projectSheetUrl}|View Items>*'
      ],
      channelWebhooks: [SECRET_OPTS.SLACK.WEBHOOKS.RECIEVING],
    },
    columns: {
      user: null,
      date: OPTS.ITEM_COLUMNS.ARRIVE_DATE,
    },
    fastForwardColumns: {
      user: [
        OPTS.ITEM_COLUMNS.REQUEST_EMAIL,
        OPTS.ITEM_COLUMNS.OFFICER_EMAIL
      ],
      date: [
        OPTS.ITEM_COLUMNS.REQUEST_DATE,
        OPTS.ITEM_COLUMNS.SUBMIT_DATE,
        OPTS.ITEM_COLUMNS.UPDATE_DATE,
      ],
    },
    fillInDefaults: true,
    officersOnly: true,
  },
  RECEIVED: {
    text: 'Received',
    allowedPrevious: ['Awaiting Pickup'],
    actionText: {
      fastForward: 'Received',
      selected: 'Mark selected items as received (picked up)',
    },
    slack: {
      emoji: ':heavy_check_mark:',
      targetUsers: OPTS.SLACK.TARGET_USERS.REQUESTORS,
      messageTemplates: [
        '{emoji} {userTags} {userFullName} marked {numMarked} item{plural} for {projectName} as received (picked up). *<{projectSheetUrl}|View Items>*'
      ],
      channelWebhooks: [SECRET_OPTS.SLACK.WEBHOOKS.PURCHASING, SECRET_OPTS.SLACK.WEBHOOKS.RECIEVING],
    },
    columns: {
      user: OPTS.ITEM_COLUMNS.RECIEVE_EMAIL,
      date: OPTS.ITEM_COLUMNS.RECIEVE_DATE,
    },
    fastForwardColumns: {
      user: [
        OPTS.ITEM_COLUMNS.REQUEST_EMAIL,
        OPTS.ITEM_COLUMNS.OFFICER_EMAIL,
      ],
      date: [
        OPTS.ITEM_COLUMNS.REQUEST_DATE,
        OPTS.ITEM_COLUMNS.SUBMIT_DATE,
        OPTS.ITEM_COLUMNS.UPDATE_DATE,
        OPTS.ITEM_COLUMNS.ARRIVE_DATE,
      ],
    },
    officersOnly: false,
  },
  DENIED: {
    text: 'Denied',
    allowedPrevious: ['New', 'Submitted', 'Ordered', 'Awaiting Info'],
    actionText: {
      fastForward: 'Denied',
      selected: 'Deny selected items',
    },
    slack: {
      emoji: ':red_circle:',
      targetUsers: OPTS.SLACK.TARGET_USERS.REQUESTORS,
      messageTemplates: [
        '{emoji} {userTags} {userFullName} *denied* {numMarked} item{plural} for {projectName} (_see comments in database_). *<{projectSheetUrl}|View Items>*'
      ],
      channelWebhooks: [SECRET_OPTS.SLACK.WEBHOOKS.PURCHASING],
    },
    columns: {
      user: OPTS.ITEM_COLUMNS.OFFICER_EMAIL,
      date: OPTS.ITEM_COLUMNS.UPDATE_DATE,
    },
    fastForwardColumns: {
      user: [
        OPTS.ITEM_COLUMNS.REQUEST_EMAIL
      ],
      date: [
        OPTS.ITEM_COLUMNS.REQUEST_DATE      
      ],
    },
    requiredColumns: [
      OPTS.ITEM_COLUMNS.OFFICER_COMMENTS
    ],
    officersOnly: true,
  },
  AWAITING_INFO: {
    text: 'Awaiting Info',
    allowedPrevious: ['New', 'Submitted', 'Denied', 'Ordered', 'Received'],
    actionText: {
      fastForward: 'Awaiting Info',
      selected: 'Request more information for selected items'
    },
    slack: {
      emoji: ':black_circle:',
      targetUsers: OPTS.SLACK.TARGET_USERS.REQUESTORS,
      messageTemplates: [
        '{emoji} {userTags} {userFullName} requested for info for {numMarked} item{plural} for {projectName} (_see comments in database_). *<{projectSheetUrl}|View Items>*'
      ],
      channelWebhooks: [SECRET_OPTS.SLACK.WEBHOOKS.PURCHASING],
    },
    columns: {
      user: OPTS.ITEM_COLUMNS.OFFICER_EMAIL,
      date: OPTS.ITEM_COLUMNS.UPDATE_DATE,
    },
    fastForwardColumns: {
      user: [
        OPTS.ITEM_COLUMNS.REQUEST_EMAIL
      ],
      date: [
        OPTS.ITEM_COLUMNS.REQUEST_DATE       
      ],
    },
    requiredColumns: [
      OPTS.ITEM_COLUMNS.OFFICER_COMMENTS
    ],
    officersOnly: true,
  }
};

/**
 * Build normal strings from the status' templates.
 * @param {Status} statusData Data for the target status.
 * @param {string} userFullName Full Name of the current user.
 * @param {string[]} requestors Emails of people who requested the items
 * affected by this action.
 * @param {number} numMarked Number of items affected by this action.
 * @param {string} projectName Name of the relevant project.
 * @param {string} projectSheetUrl Link to the relevant project's sheet in the
 * database.
 * @param {boolean} [dontTagUsers] If truthy, won't add user tags.
 * @returns {string[]} Filled in message strings.
 */
function buildSlackMessages(
    statusData,
    userFullName,
    requestors,
    numMarked,
    projectName,
    projectSheetUrl,
    dontTagUsers) {

  if(!dontTagUsers) {
    var targetUserTagsString = '';
    switch(statusData.slack.targetUsers) {
      case OPTS.SLACK.TARGET_USERS.CHANNEL:
        targetUserTagsString = '<!channel>';
        break;

      case OPTS.SLACK.TARGET_USERS.OFFICERS:
        /** Array of booleans with indexes that match officers. Only false if NO. */
        var officerNotifyOptions =
            getNamedRangeValues(OPTS.NAMED_RANGES.NOTIFY_APPROVED_OFFICERS)
            .map(function(value) {return value !== 'NO';});
        /** Emails of all the officers that do get notified. */
        var officerEmails =
            getNamedRangeValues(OPTS.NAMED_RANGES.APPROVED_OFFICERS)
            .filter(function(email, index) {return officerNotifyOptions[index];});
        var officerUserTags = officerEmails.map(getSlackTagByEmail)
            .filter(function(slackTag) {return slackTag != '';});
        targetUserTagsString = makeListFromArray(officerUserTags, 'or');
        break;

      case OPTS.SLACK.TARGET_USERS.REQUESTORS:
        var requestorUserTags = requestors.map(getSlackTagByEmail);
        targetUserTagsString = makeListFromArray(requestorUserTags, '');
    }
  }

  return statusData.slack.messageTemplates.map(function(template, index) {
    return template
        .replace('{emoji}', statusData.slack.emoji)
        .replace('{userTags}', !dontTagUsers ? (targetUserTagsString + ':') : '')
        .replace('{userFullName}', userFullName)
        .replace('{numMarked}', numMarked.toString())
        .replace('{projectName}', projectName)
        .replace('{projectSheetUrl}', projectSheetUrl)
        .replace('{plural}', numMarked !== 1 ? 's' : '');
  });
}

/**
 * Global that represents whether the user is authorized as a financial officer.
 */
function onOpen() {
  buildAndAddCustomMenu();
}

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
    .addItem(STATUSES_DATA.NEW.actionText.all, markAllNew.name)
    .addItem(STATUSES_DATA.NEW.actionText.selected, markSelectedNew.name);

  var fastFowardMenu = null;

  if (verifyFinancialOfficer()) {
    customMenu
      .addSeparator()
      .addItem(STATUSES_DATA.SUBMITTED.actionText.selected, markSelectedSubmitted.name)
      .addItem(STATUSES_DATA.APPROVED.actionText.selected, markSelectedApproved.name)
      .addItem(STATUSES_DATA.AWAITING_PICKUP.actionText.selected, markSelectedAwaitingPickup.name)
      .addSeparator()
      .addItem(STATUSES_DATA.AWAITING_INFO.actionText.selected, markSelectedAwaitingInfo.name)
      .addItem(STATUSES_DATA.DENIED.actionText.selected, markSelectedDenied.name);

    fastFowardMenu = SpreadsheetApp.getUi()
      .createMenu(OPTS.FAST_FORWARD_MENU.NAME)
      .addItem(STATUSES_DATA.NEW.actionText.fastForward, fastForwardSelectedNew.name)
      .addItem(STATUSES_DATA.SUBMITTED.actionText.fastForward, fastForwardSelectedSubmitted.name)
      .addItem(STATUSES_DATA.APPROVED.actionText.fastForward, fastForwardSelectedApproved.name)
      .addItem(STATUSES_DATA.AWAITING_INFO.actionText.fastForward, fastForwardSelectedAwaitingInfo.name)
      .addItem(STATUSES_DATA.DENIED.actionText.fastForward, fastForwardSelectedDenied.name)
      .addItem(STATUSES_DATA.AWAITING_PICKUP.actionText.fastForward, fastForwardSelectedAwaitingPickup.name)
      .addItem(STATUSES_DATA.RECEIVED.actionText.fastForward, fastForwardSelectedReceived.name);
  }

  customMenu
      .addSeparator()
      .addItem(STATUSES_DATA.RECEIVED.actionText.selected, markSelectedReceived.name);

  if(verifyAdmin()) {
    customMenu
      .addSeparator()
      .addItem('Refresh protections', protectRanges.name);
  }

  customMenu.addToUi();
  if (verifyFinancialOfficer()) fastFowardMenu.addToUi();
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
 * Returns the current user's information from the storage sheet, or (prompt)s for it,
 * or returns it from the local cache if it's been asked before.
 * @returns {{slackId:string,fullName:string,email:string,isFinancialOfficer:boolean}}
 * Information about the user.
 */
var getCurrentUserInfo = (function() {
  var currentEmail = Session.getActiveUser().getEmail();
  var cache = {
    slackId: /** @type {?string} */  (null),
    fullName: /** @type {?string} */  (null),
    isFinancialOfficer: verifyFinancialOfficer(currentEmail),
    email: currentEmail,
  };

  return function() {
    if(cache.email !== null && (cache.slackId === null || cache.fullName == null)) {
      var userSheet = SpreadsheetApp
          .getActiveSpreadsheet()
          .getSheetByName(OPTS.SHEET_NAMES.USERS);
      var userData = userSheet.getDataRange().getValues();

      var userDataFound = false;
      for(var i = 1; i < userData.length; i++) {
        if(userData[i][0] === cache.email) {
          cache.slackId = userData[i][1];
          cache.fullName = userData[i][2];
          userDataFound = true;
          break;
        }
      }

      if(!userDataFound) {
        while(!cache.slackId) cache.slackId = SpreadsheetApp.getUi().prompt(OPTS.UI.SLACK_ID_PROMPT).getResponseText();
        while(!cache.fullName) cache.fullName = SpreadsheetApp.getUi().prompt(OPTS.UI.FULL_NAME_PROMPT).getResponseText();
        userSheet.appendRow([cache.email, cache.slackId, cache.fullName]);
      }
    }

    return cache;
  };
})();


/**
 * Verify whether or not the email provided is one of an approved financial officer.
 * After first run, uses cache to avoid having to pull the range again.
 * @param {?string} [email] Email of the user to check. If no email provided,
 * uses current user email (if possible; if not returns false).
 * @returns {boolean} true if the user is a financial officer.
 */
function verifyFinancialOfficer(email) {
  if(!email) email = Session.getActiveUser().getEmail() || null;
  if(email && getNamedRangeValues(OPTS.NAMED_RANGES.APPROVED_OFFICERS)
      .indexOf(email) !== -1) {
    return true;
  }
  return false;
}

/**
 * Verify whether or not the current user is the admin.
 * @returns {boolean} true if the current user is an admin.
 */
function verifyAdmin() {
  if(Session.getActiveUser().getEmail() === SECRET_OPTS.ADMIN_EMAIL) return true;
  return false;
}

/**
 * Returns the Slack ID that matches the email address provided.
 * @param {string} email Email address of the person to look for.
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
 * Wrapper for `getSlackIdByEmail` that adds tagging formatting.
 * @param {string} email Email address of the person to look for.
 * @returns {string} The Slack ID or '' if no match.
 */
function getSlackTagByEmail(email) {
  var slackId = getSlackIdByEmail(email);
  return slackId ? '<@' + getSlackIdByEmail(email) + '>' : '';
}

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

  var nameColumnValues = getColumnRange(OPTS.ITEM_COLUMNS.NAME.index).getValues();

  /** The number of the last row in the sheet that has a value for Name. */
  var lastRowWithData = firstNonHeaderRow;

  nameColumnValues.forEach(function(name, index) {
    if(name.toString().trim() !== '') lastRowWithData = index + firstNonHeaderRow;
  });

  var numNonHeaderRowsWithData = lastRowWithData - OPTS.NUM_HEADER_ROWS;

  return [
    activeSheet.getRange(firstNonHeaderRow, 1, numNonHeaderRowsWithData, lastColumnInSheet)
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

  return activeSheet.getRange(firstNonHeaderRow, columnNumber, numNonHeaderRows, 1);
}

/**
 * Mark all of the items in the currently selected rows as `newStatus` if they
 * are currently one of the allowed previous statuses, and also fill in the date and
 * attribution columns.
 * @param {Status} newStatus The object representing the status to change the
 * selected items to.
 * @param {boolean} [markAll=false] If truthy, mark all possible rows, else mark
 * selected.
 * @returns {void}
 */
function markItems(newStatus, markAll) {
  if(!checkIfProjectSheet() || (newStatus.officersOnly && !verifyFinancialOfficer())) return;

  /** All the ranges in the sheet if `markAll` is set, else just the selected. */
  var selectedRanges = markAll ? getAllRows() : getSelectedRows();

  var numMarked = 0;
  var currentUser = getCurrentUserInfo();
  var currentUserEmail = currentUser.email;
  var currentUserFullName = currentUser.fullName;
  var currentDate = new Date();

  var currentSheet = SpreadsheetApp.getActiveSheet();
  var projectName = getProjectNameFromSheetName(currentSheet.getSheetName());
  var projectSheetUrl = SpreadsheetApp.getActiveSpreadsheet().getUrl()
      + '#gid=' + currentSheet.getSheetId();
  var itemRequestors = /** @type {string[]} */ ([]);

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
    var range = selectedRanges[i];
    var rangeValues = range.getValues();

    for(var j = 0; j < range.getNumRows(); j++) {
      /** Array of row values. */
      var row = rangeValues[j];
      // If current status is not in allowed statuses, don't verify, just skip
      // minus 1 to convert from 1-based Sheets column number to 0-based array index
      if(!isCurrentStatusAllowed(row[OPTS.ITEM_COLUMNS.STATUS.index - 1].toString(), newStatus)) continue;

      // Otherwise validate. If a single row is invalid, quit both loo[p]s
      if(!validateRow(row, newStatus)) {
        allRowsValid = false;
        break selectionsLoop;
      }
    }
  }

  if(allRowsValid) {
    // Cache the entire columns, to avoid making dozens of calls to the server
    var statusColumn = getColumnRange(OPTS.ITEM_COLUMNS.STATUS.index);
    var statusColumnValues = statusColumn.getValues();

    var userColumn, dateColumn, userColumnValues, dateColumnValues;
    if(newStatus.columns.user) {
      userColumn = getColumnRange(newStatus.columns.user.index);
      userColumnValues = userColumn.getValues();
    }
    if(newStatus.columns.date) {
      dateColumn = getColumnRange(newStatus.columns.date.index);
      dateColumnValues = dateColumn.getValues();
    }

    var accountColumn, categoryColumn, accountColumnValues, categoryColumnValues;
    if(newStatus.fillInDefaults) {
      accountColumn = getColumnRange(OPTS.ITEM_COLUMNS.ACCOUNT.index);
      categoryColumn = getColumnRange(OPTS.ITEM_COLUMNS.CATEGORY.index);
      accountColumnValues = accountColumn.getValues();
      categoryColumnValues = categoryColumn.getValues();
    }

    // Read (not modify, so no need for range) the requestor data for notifying
    var requestorColumnValues;
    if(newStatus.slack.targetUsers === OPTS.SLACK.TARGET_USERS.REQUESTORS) {
      requestorColumnValues = getColumnRange(OPTS.ITEM_COLUMNS.REQUEST_EMAIL.index).getValues();
    }

    // Loop through the ranges
    for(var k = 0; k < selectedRanges.length; k++) {
      var range = selectedRanges[k];
      var rangeStartIndex = range.getRow() - 1;

      // Loop through the rows in the range
      rowLoop: for (var l = 0; l < range.getNumRows(); l++) {
        /** The index (not number) of the current row in the spreadsheet. */
        var currentSheetRowIndex = rangeStartIndex + l;
        /**
         * The index of the current value row in the spreadsheet, with the first
         * row after the headers being 0.
         */
        var currentValuesRowIndex = currentSheetRowIndex - OPTS.NUM_HEADER_ROWS;

        // If this row's status is not in allowed statuses, don't verify, just skip
        var currentStatusText = statusColumnValues[currentValuesRowIndex][0].toString();
        if(!isCurrentStatusAllowed(currentStatusText, newStatus)) continue rowLoop;

        // Update values in local cache
        // These ranges don't include the header, so 0 in the range is
        // OPTS.NUM_HEADER_ROWS in the spreadsheet
        statusColumnValues[currentValuesRowIndex][0] = newStatus.text;

        if(newStatus.columns.user) {
          userColumnValues[currentValuesRowIndex][0] = currentUserEmail;
        }
        if(newStatus.columns.date) {
          dateColumnValues[currentValuesRowIndex][0] = currentDate;
        }

        if(newStatus.fillInDefaults) {
          if(accountColumnValues[currentValuesRowIndex][0].toString() === '') {
            accountColumnValues[currentValuesRowIndex][0] =
                OPTS.DEFAULT_VALUES.ACCOUNT_NAME;
          }
          if(categoryColumnValues[currentValuesRowIndex][0].toString() === '') {
            categoryColumnValues[currentValuesRowIndex][0] =
                OPTS.DEFAULT_VALUES.CATEGORY;
          }
        }

        // Save the requestor data for notifying; avoid duplicates
        if(newStatus.slack.targetUsers === OPTS.SLACK.TARGET_USERS.REQUESTORS) {
          pushIfNewAndTruthy(itemRequestors, requestorColumnValues[currentValuesRowIndex][0].toString());
        }

        numMarked++;
      }
    }

    // Write the cached values
    statusColumn.setValues(statusColumnValues);

    if(newStatus.columns.user) userColumn.setValues(userColumnValues);
    if(newStatus.columns.date) dateColumn.setValues(dateColumnValues);

    if(newStatus.fillInDefaults) {
      accountColumn.setValues(accountColumnValues);
      categoryColumn.setValues(categoryColumnValues);
    }

    /** All of the possible 'from' statuses, but with double quotes around them. */
    var quotedFromStatuses = newStatus.allowedPrevious.map(wrapInDoubleQuotes);

    successNotification(numMarked + ' items marked from '
        + makeListFromArray(quotedFromStatuses, 'or')
        + ' to "' + newStatus.text + '."');

    if(numMarked !== 0) {
      slackNotifyItems(
        newStatus,
        currentUserFullName,
        itemRequestors,
        numMarked,
        projectName,
        projectSheetUrl
      );
    }
  }
}

/**
 * Fast-forward all of the items in the currently selected rows to `newStatus`,
 * filling in the date and attribution columns but not notifying on Slack. Allows
 * for skipping statuses
 * @param {Status} newStatus The object representing the status to fast-forward the
 * selected items to.
 * @returns {void}
 */
function fastForwardItems(newStatus) {
  if(!checkIfProjectSheet() || !verifyFinancialOfficer()) return;

  var selectedRanges = getSelectedRows();

  var numMarked = 0;
  var currentOfficer = getCurrentUserInfo();
  var currentOfficerEmail = currentOfficer.email;
  var currentDate = new Date();

  // Cache the entire columns, to avoid making dozens of calls to the server
  var statusColumn = getColumnRange(OPTS.ITEM_COLUMNS.STATUS.index);
  var statusColumnValues = statusColumn.getValues();

  // Fetch normal columns to update
  var userColumn, dateColumn, userColumnValues, dateColumnValues;
  if(newStatus.columns.user) {
    userColumn = getColumnRange(newStatus.columns.user.index);
    userColumnValues = userColumn.getValues();
  }
  if(newStatus.columns.date) {
    dateColumn = getColumnRange(newStatus.columns.date.index);
    dateColumnValues = dateColumn.getValues();
  }

  // Fetch default columns to fill if empty
  var accountColumn, categoryColumn, accountColumnValues, categoryColumnValues;
  if(newStatus.fillInDefaults) {
    accountColumn = getColumnRange(OPTS.ITEM_COLUMNS.ACCOUNT.index);
    categoryColumn = getColumnRange(OPTS.ITEM_COLUMNS.CATEGORY.index);
    accountColumnValues = accountColumn.getValues();
    categoryColumnValues = categoryColumn.getValues();
  }

  // Fetch fast-forward columns to fill if empty
  var pastUserColumns, pastDateColumns, pastUserColumnsValues, pastDateColumnsValues;
  pastUserColumns = newStatus.fastForwardColumns.user.map(function(ffCol) {
    return getColumnRange(ffCol.index);
  });
  pastDateColumns = newStatus.fastForwardColumns.date.map(function(ffCol) {
    return getColumnRange(ffCol.index);
  });
  pastUserColumnsValues = pastUserColumns.map(function(colRange) {
    return colRange.getValues();
  });
  pastDateColumnsValues = pastDateColumns.map(function(colRange) {
    return colRange.getValues();
  });

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

      // Update values in local cache
      // These ranges don't include the header, so 0 in the range is
      // OPTS.NUM_HEADER_ROWS in the spreadsheet
      statusColumnValues[currentValuesRowIndex][0] = newStatus.text;

      if(newStatus.columns.user) {
        userColumnValues[currentValuesRowIndex][0] = currentOfficerEmail;
      }
      if(newStatus.columns.date) {
        dateColumnValues[currentValuesRowIndex][0] = currentDate;
      }

      if(newStatus.fillInDefaults) {
        if(accountColumnValues[currentValuesRowIndex][0].toString() === '') {
          accountColumnValues[currentValuesRowIndex][0] =
              OPTS.DEFAULT_VALUES.ACCOUNT_NAME;
        }
        if(categoryColumnValues[currentValuesRowIndex][0].toString() === '') {
          categoryColumnValues[currentValuesRowIndex][0] =
              OPTS.DEFAULT_VALUES.CATEGORY;
        }
      }

      // If any of the past columns are blank, fill them in with current info
      pastUserColumnsValues.forEach(function(columnValues) {
        if(columnValues[currentValuesRowIndex][0].toString() === '') {
          columnValues[currentValuesRowIndex][0] = currentOfficerEmail;
        }
      });
      pastDateColumnsValues.forEach(function(columnValues) {
        if(columnValues[currentValuesRowIndex][0].toString() === '') {
          columnValues[currentValuesRowIndex][0] = currentDate;
        }
      });

      numMarked++;
    }
  }

  // Write the cached values
  statusColumn.setValues(statusColumnValues);

  if(newStatus.columns.user) userColumn.setValues(userColumnValues);
  if(newStatus.columns.date) dateColumn.setValues(dateColumnValues);

  if(newStatus.fillInDefaults) {
    accountColumn.setValues(accountColumnValues);
    categoryColumn.setValues(categoryColumnValues);
  }

  pastUserColumns.forEach(function(columnRange, index) {
    columnRange.setValues(pastUserColumnsValues[index]);
  });
  pastDateColumns.forEach(function(columnRange, index) {
    columnRange.setValues(pastDateColumnsValues[index]);
  });

  successNotification(numMarked + ' items fast-forwarded to "' + newStatus.text + '."');
}

/**
 * Push `potentialNewItem` to `arr` if it's not already in `arr`. Returns modified
 * `arr`.
 * @param {[]} arr
 * @param {*} potentialNewItem
 * @return {[]}
 */
function pushIfNewAndTruthy(arr, potentialNewItem) {
  if(arr.indexOf(potentialNewItem) === -1 && potentialNewItem) arr.push(potentialNewItem);
  return arr;
}

/**
 * Check if the current status of the row is in the valid statuses list.
 * @param {string} currentStatusText The current status of the row.
 * @param {Status} newStatus The status object to check for changing to.
 * @returns {boolean} True if the current status of the row allows it to be
 * changed to the newStatus.
 */
function isCurrentStatusAllowed(currentStatusText, newStatus) {
  var currentStatusTrimmed = currentStatusText.trim();
  return newStatus.allowedPrevious.indexOf(currentStatusTrimmed) !== -1;
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
 * @param {Status} newStatus The new status of the row for testing against.
 * @returns {boolean} True if the row is valid and can be submitted.
 */
function validateRow(rowValues, newStatus) {
  var column, columnIndex;

  for(var i = 0; newStatus.reccomendedColumns && i < newStatus.reccomendedColumns.length; i++) {
    column = newStatus.reccomendedColumns[i];
    columnIndex = column.index - 1;
    if(rowValues[columnIndex] === '') warnNotification('One or more items is missing a value for "' + column.name + '". Will mark anyway with default value.');
  }

  for(var j = 0; newStatus.requiredColumns && j < newStatus.requiredColumns.length; j++) {
    column = newStatus.requiredColumns[j];
    columnIndex = column.index - 1;
    if(rowValues[columnIndex] === '') {
      errorNotification('Cannot submit: one or more items is missing a value for "' + column.name + '". This value is required.');
      return false;
    }
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
 * Build normal strings from the status' templates.
 * @param {Status} statusData Data for the target status.
 * @param {string} userFullName Full Name of the current user.
 * @param {string[]} requestors Emails of people who requested the items
 * affected by this action.
 * @param {number} numMarked Number of items affected by this action.
 * @param {string} projectName Name of the relevant projec.
 * @param {string} projectSheetUrl Link to the relevant project's sheet in the
 * databse.
 * @returns {string[]} Filled in message strings.
 */
function slackNotifyItems(
  statusData,
  userFullName,
  requestors,
  numMarked,
  projectName,
  projectSheetUrl) {
  statusData.slack.channelWebhooks.forEach(function(webhook, index) {
    var messages = [];
    if(index === 0) {
      messages = buildSlackMessages(
          statusData,
          userFullName,
          requestors,
          numMarked,
          projectName,
          projectSheetUrl);
    } else {
      messages = buildSlackMessages(
          statusData,
          userFullName,
          requestors,
          numMarked,
          projectName,
          projectSheetUrl,
          true);
    }

    messages.forEach(function(message) {sendSlackMessage(message, webhook);});
  });
}

/**
 * Get the full name of the project that matches the name of the sheet.
 * @param {string} sheetName Name of the project's sheet.
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
  markItems(STATUSES_DATA.NEW);
}

/** Mark all possible items in the sheet as new. */
function markAllNew() {
  markItems(STATUSES_DATA.NEW, true);
}

/** Mark selected items in the sheet as recieved. */
function markSelectedReceived() {
  markItems(STATUSES_DATA.RECEIVED);
}

/** Mark selected items in the sheet as submitted. */
function markSelectedSubmitted() {
  markItems(STATUSES_DATA.SUBMITTED);
}

/** Mark selected items in the sheet as approved. */
function markSelectedApproved() {
  markItems(STATUSES_DATA.APPROVED);
}

/** Mark selected items in the sheet as arrived / awaiting pickup. */
function markSelectedAwaitingPickup() {
  markItems(STATUSES_DATA.AWAITING_PICKUP);
}

/** Mark selected items in the sheet as awaiting info. */
function markSelectedAwaitingInfo() {
  markItems(STATUSES_DATA.AWAITING_INFO);
}

/** Mark selected items in the sheet as denied. */
function markSelectedDenied() {
  markItems(STATUSES_DATA.DENIED);
}

/** Fast-forward the selected items in the sheet to new. */
function fastForwardSelectedNew() {
  fastForwardItems(STATUSES_DATA.NEW);
}

/** Fast-forward selected items in the sheet to recieved. */
function fastForwardSelectedReceived() {
  fastForwardItems(STATUSES_DATA.RECEIVED);
}

/** Fast-forward selected items in the sheet to submitted. */
function fastForwardSelectedSubmitted() {
  fastForwardItems(STATUSES_DATA.SUBMITTED);
}

/** Fast-forward selected items in the sheet to approved. */
function fastForwardSelectedApproved() {
  fastForwardItems(STATUSES_DATA.APPROVED);
}

/** Fast-forward selected items in the sheet to arrived / awaiting pickup. */
function fastForwardSelectedAwaitingPickup() {
  fastForwardItems(STATUSES_DATA.AWAITING_PICKUP);
}

/** Fast-forward selected items in the sheet to awaiting info. */
function fastForwardSelectedAwaitingInfo() {
  fastForwardItems(STATUSES_DATA.AWAITING_INFO);
}

/** Fast-forward selected items in the sheet to denied. */
function fastForwardSelectedDenied() {
  fastForwardItems(STATUSES_DATA.DENIED);
}

/** Reinstate / update all the protected ranges. */
function protectRanges() {
  if(!verifyAdmin()) return;

  var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  var financialOfficers = getNamedRangeValues(OPTS.NAMED_RANGES.APPROVED_OFFICERS);
  var projectSheetNames = getNamedRangeValues(OPTS.NAMED_RANGES.PROJECT_SHEETS);
  var admin = SECRET_OPTS.ADMIN_EMAIL;
  var userDataSheetName = OPTS.SHEET_NAMES.USERS;

  var adminProtectDescription = 'This part of the sheet can only be edited by the admin.';
  var officerProtectDescription = 'This part of the sheet can only be edited by Financial Officers.';

  SpreadsheetApp.getActiveSpreadsheet().getProtections(SpreadsheetApp.ProtectionType.RANGE)
      .forEach(function(protection) {
        protection.remove();
      });

  SpreadsheetApp.getActiveSpreadsheet().getProtections(SpreadsheetApp.ProtectionType.SHEET)
      .forEach(function(protection) {
        protection.remove();
      });

  sheets.forEach(function(sheet) {
    var sheetName = sheet.getName();

    if(projectSheetNames.indexOf(sheetName) !== -1) {
      // Lock certain sections of project sheets (only the headers and formula-driven parts)
      var numDataRows = sheet.getLastRow() - OPTS.NUM_HEADER_ROWS;

      var headerRangeProtection = sheet.getRange(1, 1, OPTS.NUM_HEADER_ROWS, sheet.getLastColumn()).protect();
      var calculatedPriceColumnProtection = sheet.getRange(1, OPTS.ITEM_COLUMNS.TOTAL_PRICE.index,numDataRows, 1).protect();
      var financialOfficerRangeProtection = sheet.getRange(3, OPTS.ITEM_COLUMNS.OFFICER_EMAIL.index, numDataRows, OPTS.NUM_OFFICER_COLS).protect();

      headerRangeProtection.removeEditors(headerRangeProtection.getEditors());
      calculatedPriceColumnProtection.removeEditors(calculatedPriceColumnProtection.getEditors());
      financialOfficerRangeProtection.removeEditors(financialOfficerRangeProtection.getEditors());

      headerRangeProtection.setDescription(adminProtectDescription);
      calculatedPriceColumnProtection.setDescription(adminProtectDescription);
      financialOfficerRangeProtection.setDescription(officerProtectDescription);

      headerRangeProtection.addEditor(admin);
      calculatedPriceColumnProtection.addEditor(admin);
      financialOfficerRangeProtection.addEditors(financialOfficers);

    } else if(sheetName !== userDataSheetName) {
      // Lock the entire sheet if not the user data sheet
      var sheetProtection = sheet.protect();
      sheetProtection.removeEditors(sheetProtection.getEditors());
      sheetProtection.addEditors(financialOfficers);
    }
  });

  // Protect the statuses, since they need to match the values in the script
  var statusesProtection = SpreadsheetApp.getActiveSpreadsheet()
      .getRangeByName(OPTS.NAMED_RANGES.STATUSES).protect();
  statusesProtection.setDescription(adminProtectDescription);
  statusesProtection.removeEditors(statusesProtection.getEditors());
  statusesProtection.addEditor(admin);
}