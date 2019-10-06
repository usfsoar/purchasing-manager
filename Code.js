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
    receive_EMAIL: {index: 21, name: 'receive Date'},
    receive_DATE: {index: 22, name: 'receiver Email'},
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
    USERS: "Users",
    PURCHASING_TEMPLATE: "Purchasing Sheet Template",
    MAIN_DASHBOARD: "Main Dashboard"
  },
  DASHBOARD_CELLS: {
    TOTAL_BUDGET: {
      row: 4, // 1-based index!
      column: 3,
    },
    TOTAL_EXPENSES: {
      row: 4,
      column: 4
    },
    PURCHASES_FOLDER: {
      row: 61,
      column: 5
    }
  },
  /** Slack API pieces */
  SLACK: {
    CHECK_MARK_EMOJI: ':heavy_check_mark:',
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
    },
    STATUS_SLASH_COMMAND: 'budgetstatus',
    ITEM_LIST_ACTION_NAME_LEGACY: 'listItems',
    ITEM_LIST_ACTION_NAME: 'showItemList',
    SOAR_ICON: 'http://www.usfsoar.com/wp-content/uploads/2018/09/595bae9a-c1f9-4b46-880e-dc6d4e1d0dac.png'
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
        '{emoji} {userTags} {userFullName} has submitted {numMarked} new item{plural} to be purchased for {projectName}.'
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
      OPTS.ITEM_COLUMNS.UNIT_PRICE,
      OPTS.ITEM_COLUMNS.QUANTITY,
      OPTS.ITEM_COLUMNS.CATEGORY
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
        '{emoji} {userTags} {userFullName} marked {numMarked} item{plural} for {projectName} as *submitted* to Student Government.'
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
      OPTS.ITEM_COLUMNS.REQUEST_ID
    ],
    fillInDefaults: true,
    officersOnly: true,
  },
  APPROVED: {
    text: 'Ordered',
    allowedPrevious: ['Submitted'],
    actionText: {
      fastForward: 'Ordered',
      selected: 'Mark selected items as ordered',
    },
    slack: {
      emoji: ':white_circle:',
      targetUsers: OPTS.SLACK.TARGET_USERS.REQUESTORS,
      messageTemplates: [
        '{emoji} {userTags} {userFullName} marked {numMarked} item{plural} for {projectName} as *ordered*.'
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
        '{emoji} {userFullName} marked {numMarked} item{plural} for {projectName} as awaiting pickup (usually in MSC 4300). _React with ' + OPTS.SLACK.CHECK_MARK_EMOJI + ' if you\'re going to pick them up._'
      ],
      channelWebhooks: [SECRET_OPTS.SLACK.WEBHOOKS.PURCHASING],
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
    allowedPrevious: ['Awaiting Pickup', 'Submitted', 'Ordered'],
    actionText: {
      fastForward: 'Received',
      selected: 'Mark selected items as received (picked up)',
    },
    slack: {
      emoji: ':heavy_check_mark:',
      targetUsers: OPTS.SLACK.TARGET_USERS.REQUESTORS,
      messageTemplates: [
        '{emoji} {userTags} {userFullName} marked {numMarked} item{plural} for {projectName} as received (picked up).'
      ],
      channelWebhooks: [SECRET_OPTS.SLACK.WEBHOOKS.PURCHASING],
    },
    columns: {
      user: OPTS.ITEM_COLUMNS.receive_EMAIL,
      date: OPTS.ITEM_COLUMNS.receive_DATE,
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
        '{emoji} {userTags} {userFullName} *denied* {numMarked} item{plural} for {projectName} (_see comments in database_).'
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
        '{emoji} {userTags} {userFullName} requested more info for {numMarked} item{plural} for {projectName} (_see comments in database_). Update the information, then resubmit as new items.'
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
  RECEIVED_REIMBURSE: {
    text: 'Received - Awaiting Reimbursement',
    allowedPrevious: ['New', 'Submitted', 'Ordered', 'received', 'Awaiting Pickup', 'Awaiting Info'],
    actionText: {
      fastForward: 'Received - Awaiting Reimbursement',
      selected: 'Mark selected items received and request reimbursement'
    },
    slack: {
      emoji: ':heavy_dollar_sign:',
      targetUsers: OPTS.SLACK.TARGET_USERS.OFFICERS,
      messageTemplates: [
        '{emoji} {userTags} {userFullName} marked {numMarked} item{plural} as received for {projectName} and requested reimbursement for them.'
      ],
      channelWebhooks: [SECRET_OPTS.SLACK.WEBHOOKS.PURCHASING],
    },
    columns: {
      date: OPTS.ITEM_COLUMNS.receive_DATE,
    },
    requiredColumns: [
      OPTS.ITEM_COLUMNS.REQUEST_COMMENTS
    ],
    officersOnly: false,
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
  },
  REIMBURSED: {
    text: 'Reimbursed',
    allowedPrevious: ['Received - Awaiting Reimbursement', 'Received'],
    actionText: {
      fastForward: 'Reimbursed',
      selected: 'Mark selected items as reimbursed'
    },
    slack: {
      emoji: ':heavy_dollar_sign:',
      targetUsers: OPTS.SLACK.TARGET_USERS.REQUESTORS,
      messageTemplates: [
        '{emoji} {userTags} {userFullName} sent reimbursement for {numMarked} item{plural}.'
      ],
      channelWebhooks: [SECRET_OPTS.SLACK.WEBHOOKS.PURCHASING],
    },
    columns: {
      date: OPTS.ITEM_COLUMNS.UPDATE_DATE,
    },
    requiredColumns: [],
    officersOnly: true,
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
  }
};

var TEST_STATUS = {
  text: 'Test',
  allowedPrevious: ['', 'Test'],
  actionText: {
    fastForward: 'Test',
    selected: 'Test update item',
  },
  slack: {
    emoji: ':checkered_flag:',
    targetUsers: OPTS.SLACK.TARGET_USERS.CHANNEL,
    messageTemplates: [
      '{emoji} {userTags} {userFullName} marked {numMarked} item{plural} for {projectName} as *test* by TEsting.'
    ],
    channelWebhooks: [SECRET_OPTS.SLACK.WEBHOOKS.DEV],
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
};

// Handle post request from Slack
function doPost(e) {
  var message = {
    response_type: "ephemeral",
    replace_original: false,
    text: "Error: command not found.",
  };

  if(e.parameter.command == "/budgetstatus") {
    // If the budgetStatus command, send the budgetStatus message
    var text = e.parameter.text;
    var sheetName = getSheetNameFromProjectName(text, true);
    if(sheetName !== null) text = sheetName;
    message = buildProjectStatusSlackMessage(text);

  } else if(e.parameter.payload) {
    // Else maybe it's an interactive message command. Parse the payload and check.
    var payload = JSON.parse(e.parameter.payload);
    console.log(payload);

    if(payload.type === "interactive_message"
        && payload.actions) {
      if(payload.actions[0].name === OPTS.SLACK.ITEM_LIST_ACTION_NAME_LEGACY) {
        var parsedText = payload.actions[0].value;

        message = {
          response_type: "ephemeral",
          replace_original: false,
          text: parsedText,
        };
      } else {
        console.log("so far so good...");
        message = JSON.parse(payload.actions[0].value);
        console.log(message);
      }
    }
  }

  return ContentService.createTextOutput(JSON.stringify(message)).setMimeType(ContentService.MimeType.JSON);
}

/**
 * Replace all occurrences of a string within another string.
 * @param {string} search The string to search in. Not modified.
 * @param {string} replacement The search term to match.
 * @returns {string}
 * @todo Make NOT a prototype modifier (only acceptable because this is GAS and
 * the JS never gets updated anyway).
 */
String.prototype.replaceAll = function(search, replacement) {
  var target = this;
  return target.replace(new RegExp(search, 'g'), replacement);
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
        .replaceAll('{emoji}', statusData.slack.emoji)
        .replaceAll('{userTags}', !dontTagUsers ? (targetUserTagsString + ':') : '')
        .replaceAll('{userFullName}', userFullName)
        .replaceAll('{numMarked}', numMarked.toString())
        .replaceAll('{projectName}', projectName)
        .replaceAll('{projectSheetUrl}', projectSheetUrl)
        .replaceAll('{plural}', numMarked !== 1 ? 's' : '');
  });
}

/**
 * Build a project status message to be used in Slack, with current information
 * about the given project.
 * @param {string} project The name of the project.
 * @returns {Object} A valid Slack message with attachments.
 */
function buildProjectStatusSlackMessage(project) {
  var projectSheetName = getSheetNameFromProjectName(project, true) || project;
  if(!checkIfProjectSheet(projectSheetName)) return {
    "response_type": "ephemeral",
    "text": "Sorry, I don't recognize that project."
  };
  var projectName = getProjectNameFromSheetName(project);

  var dashboardSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(projectSheetName + " Dashboard");
  var totalBudget = dashboardSheet.getRange(
      OPTS.DASHBOARD_CELLS.TOTAL_BUDGET.row,
      OPTS.DASHBOARD_CELLS.TOTAL_BUDGET.column).getValue();
  var totalExpenses = dashboardSheet.getRange(
      OPTS.DASHBOARD_CELLS.TOTAL_EXPENSES.row,
      OPTS.DASHBOARD_CELLS.TOTAL_EXPENSES.column).getValue();

  var budgetRemaining = (totalBudget - totalExpenses).toFixed(2);
  var percentBudgetRemaining = (budgetRemaining / totalBudget * 100).toFixed(0);

  var dashboardSheetUrl = SpreadsheetApp.getActiveSpreadsheet().getUrl() + "#gid=" + dashboardSheet.getSheetId();

  var projectSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(projectSheetName);
  var projectSheetUrl = SpreadsheetApp.getActiveSpreadsheet().getUrl() + "#gid=" + projectSheet.getSheetId();

  return {
        "response_type": "in_channel",
        "attachments": [
            {
                "fallback": "The " + projectName + " project has $" + budgetRemaining + " (or " + percentBudgetRemaining + "%) remaining, out of a total annual budget of $" + totalBudget.toFixed(2) + ". For more details, see the <" + dashboardSheetUrl + "|project dashboard>.",
                "color": dashboardSheet.getTabColor(),
                "title": projectName + " Budget Status",
                "text": "This is the latest budget information from the SOAR Purchasing Database:",
                "fields": [
                  {
                    "title": "Total Budget",
                    "value": "$" + totalBudget.toFixed(2),
                    "short": true
                  },
                  {
                    "title": "Percent Remaining",
                    "value": percentBudgetRemaining + "%",
                    "short": true
                  },
                  {
                    "title": "Total Expenses",
                    "value": "$" + totalExpenses.toFixed(2),
                    "short": true
                  },
                  {
                    "title": "Amount Remaining",
                    "value": "$" + budgetRemaining,
                    "short": true
                  }
                ],
                "footer": "SOAR Purchasing Database",
                "footer_icon": OPTS.SLACK.SOAR_ICON,
                "ts": new Date().getTime() / 1000,
                "actions": [{
                    "type": "button",
                    "text": "Open Dashboard ↗",
                    "url": dashboardSheetUrl
                  }, {
                    "type": "button",
                    "text": "Open Purchasing Sheet ↗",
                    "url": projectSheetUrl
                  }
                ]
            }
        ]
      };
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
  var testMenu = null;

  if (verifyFinancialOfficer()) {
    customMenu
      .addSeparator()
      .addItem(STATUSES_DATA.SUBMITTED.actionText.selected, markSelectedSubmitted.name)
      .addItem(STATUSES_DATA.APPROVED.actionText.selected, markSelectedApproved.name)
      .addItem(STATUSES_DATA.AWAITING_PICKUP.actionText.selected, markSelectedAwaitingPickup.name)
      .addSeparator()
      .addItem(STATUSES_DATA.AWAITING_INFO.actionText.selected, markSelectedAwaitingInfo.name)
      .addItem(STATUSES_DATA.DENIED.actionText.selected, markSelectedDenied.name)
      .addItem(STATUSES_DATA.REIMBURSED.actionText.selected, markSelectedReimbursed.name)
      .addSeparator()
      .addItem("Send to new purchasing sheet", sendSelectedToSheet.name);

    fastFowardMenu = SpreadsheetApp.getUi()
      .createMenu(OPTS.FAST_FORWARD_MENU.NAME)
      .addItem(STATUSES_DATA.NEW.actionText.fastForward, fastForwardSelectedNew.name)
      .addItem(STATUSES_DATA.SUBMITTED.actionText.fastForward, fastForwardSelectedSubmitted.name)
      .addItem(STATUSES_DATA.APPROVED.actionText.fastForward, fastForwardSelectedApproved.name)
      .addItem(STATUSES_DATA.AWAITING_INFO.actionText.fastForward, fastForwardSelectedAwaitingInfo.name)
      .addItem(STATUSES_DATA.DENIED.actionText.fastForward, fastForwardSelectedDenied.name)
      .addItem(STATUSES_DATA.AWAITING_PICKUP.actionText.fastForward, fastForwardSelectedAwaitingPickup.name)
      .addItem(STATUSES_DATA.RECEIVED.actionText.fastForward, fastForwardSelectedReceived.name)
      .addItem(STATUSES_DATA.RECEIVED_REIMBURSE.actionText.fastForward, fastForwardSelectedReceivedReimburse.name)
      .addItem(STATUSES_DATA.REIMBURSED.actionText.fastForward, fastForwardSelectedReimbursed.name);
  }

  customMenu
      .addSeparator()
      .addItem(STATUSES_DATA.RECEIVED.actionText.selected, markSelectedReceived.name)
      .addSeparator()
      .addItem(STATUSES_DATA.RECEIVED_REIMBURSE.actionText.selected, markSelectedReceivedReimburse.name);

  if(verifyAdmin()) {
    customMenu
      .addSeparator()
      .addItem('Refresh protections', protectRanges.name);
    testMenu = SpreadsheetApp.getUi()
      .createMenu("Test")
      .addItem('Test Update', testUpdateItem.name);
  }

  customMenu.addToUi();
  if (verifyFinancialOfficer()) fastFowardMenu.addToUi();
  if (verifyAdmin()) testMenu.addToUi();
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
 * @returns {{slackId:string,fullName:string,email:string,isFinancialOfficer:boolean,phone?:string}}
 * Information about the user.
 */
var getCurrentUserInfo = (function() {
  var currentEmail = Session.getActiveUser().getEmail();
  var cache = {
    slackId: /** @type {?string} */  (null),
    fullName: /** @type {?string} */  (null),
    isFinancialOfficer: verifyFinancialOfficer(currentEmail),
    email: currentEmail,
    phone: /** @type {?string} */ (null)
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
          cache.phone = userData[i][3] || null;
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
 * @param {boolean} pad If true, will add padding to end of string to make it
 * the target length.
 * @param {boolean} padOnly If true, won't truncate ever, will just pad.
 * @returns {string} The truncated string.
 */
function truncateString(longString, chars, pad, padOnly) {
  longString = longString.toString();

  if(longString.length > chars && !padOnly) {
    longString = longString.slice(0, chars - 4) + "...";
  }
  if(pad) {
    var padding = chars - longString.length > 0 ? chars - longString.length : 0;
    for(var i = 0; i < padding; i++) longString += " ";
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
 * @param {?string} sheetName Name of the sheet to check. If empty, uses current sheet.
 * @returns {boolean} True if a project sheet is active.
 */
function checkIfProjectSheet(sheetName) {
  var currentSheetName = sheetName || SpreadsheetApp.getActiveSheet().getName();

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

    /* List of items, for sending to Slack */
    var items = [];

    // Loop through the ranges
    for(var k = 0; k < selectedRanges.length; k++) {
      var range = selectedRanges[k];
      var rangeStartIndex = range.getRow() - 1;
      var rangeValues = range.getValues();

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

        items.push({
          name:              rangeValues[l][OPTS.ITEM_COLUMNS.NAME.index - 1],
          quantity:          rangeValues[l][OPTS.ITEM_COLUMNS.QUANTITY.index - 1],
          totalPrice:        rangeValues[l][OPTS.ITEM_COLUMNS.TOTAL_PRICE.index - 1],
          unitPrice:         rangeValues[l][OPTS.ITEM_COLUMNS.UNIT_PRICE.index - 1],
          category:          rangeValues[l][OPTS.ITEM_COLUMNS.CATEGORY.index - 1],
          requestorComments: rangeValues[l][OPTS.ITEM_COLUMNS.REQUEST_COMMENTS.index - 1],
          officerComments:   rangeValues[l][OPTS.ITEM_COLUMNS.OFFICER_COMMENTS.index - 1],
          supplier:          rangeValues[l][OPTS.ITEM_COLUMNS.SUPPLIER.index - 1],
          productNum:        rangeValues[l][OPTS.ITEM_COLUMNS.PRODUCT_NUM.index - 1],
          link:              rangeValues[l][OPTS.ITEM_COLUMNS.LINK.index - 1]
        });
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

    successNotification(items.length + ' items marked from '
        + makeListFromArray(quotedFromStatuses, 'or')
        + ' to "' + newStatus.text + '."');

    var projectColor = currentSheet.getTabColor();

    if(items.length !== 0) {
      slackNotifyItems(
        newStatus,
        currentUserFullName,
        itemRequestors,
        items,
        projectName,
        projectSheetUrl,
        projectColor
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
 * @param {Object} messageData The message to send, according to the Slack API.
 */
function sendSlackMessage(messageData, webhook) {
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
 * @param {Object[]} itemsMarked Data about all items affected by this action.
 * @param {string} itemsMarked.name
 * @param {string} itemsMarked.quantity
 * @param {number} itemsMarked.totalPrice
 * @param {number} itemsMarked.unitPrice
 * @param {string} itemsMarked.category
 * @param {string} itemsMarked.requestorComments
 * @param {string} itemsMarked.officerComments
 * @param {string} itemsMarked.supplier
 * @param {string} itemsMarked.productNum
 * @param {string} itemsMarked.link
 * @param {string} projectName Name of the relevant projec.
 * @param {string} projectSheetUrl Link to the relevant project's sheet in the
 * databse.
 * @param {string} projectColor Tab color of the project's sheet.
 */
function slackNotifyItems(
  statusData,
  userFullName,
  requestors,
  itemsMarked,
  projectName,
  projectSheetUrl,
  projectColor) {
  statusData.slack.channelWebhooks.forEach(function(webhook, index) {
    var messages = [];
    Logger.log(itemsMarked);
    if(index === 0) {
      messages = buildSlackMessages(
          statusData,
          userFullName,
          requestors,
          itemsMarked.length,
          projectName,
          projectSheetUrl);
    } else {
      messages = buildSlackMessages(
          statusData,
          userFullName,
          requestors,
          itemsMarked.length,
          projectName,
          projectSheetUrl,
          true);
    }

    messages = messages.map(function(messageText) {return {text: messageText};});
    messages[messages.length - 1].attachments = [
      {
        callback_id: "itemNotification",
        fallback: "<" + projectSheetUrl + "|View Items>",
        actions: [
          buildItemListSlackAttachment(
              itemsMarked,
              projectName,
              projectSheetUrl,
              userFullName,
              statusData.text,
              projectColor),
          {
            type: "button",
            text: "Open Sheet ↗",
            url: projectSheetUrl
          }
        ],
        color: projectColor
      }
    ];

    messages.forEach(function(message) {sendSlackMessage(message, webhook);});
  });
}

/**
 * Build a Slack button attachment that sends a request to show the full item
 * list on click.
 * @param {Object[]} items Data about all items to list.
 * @param {string} items.name
 * @param {string} items.quantity
 * @param {number} items.totalPrice
 * @param {number} items.unitPrice
 * @param {string} items.category
 * @param {string} items.requestorComments
 * @param {string} items.officerComments
 * @param {string} items.supplier
 * @param {string} items.productNum
 * @param {string} items.link
 * @param {string} projectName
 * @param {string} projectSheetUrl
 * @param {string} user
 * @param {string} action
 * @param {string} projectColor
 */
function buildItemListSlackAttachment(items, projectName, projectSheetUrl, user, action, projectColor) {
  var attachment = {
    type: "button",
    text: "List Items",
    name: OPTS.SLACK.ITEM_LIST_ACTION_NAME,
    /** JSON to parse as return message later */
    value: ''
  };

  /*var sep = "   ";

  items.forEach(function(currentItem) {
    var currentItemString =
        truncateString(currentItem.name, 35, true) + sep +
        truncateString(currentItem.quantity, 3, true, true) + sep +
        truncateString("$" + currentItem.totalPrice.toFixed(2), 5, true, true) + sep +
        currentItem.category + "{NL}";
    if(currentItem.requestorComments) currentItemString += "{TAB}Note: \"" + currentItem.requestorComments + "\"{NL}";
    if(currentItem.officerComments) currentItemString += "{TAB}Officer Note: \"" + currentItem.officerComments + "\"{NL}";
    attachment.value += currentItemString;
  });

  attachment.value = truncateString(attachment.value, 2000);*/

  var itemListMessage = {
    response_type: "ephemeral",
    replace_original: false,
    text: "Here are all the items that were affected by that action:",
    attachments: [],
    parse: "full",
    mrkdwn: true
  };

  var itemsByCategory = {};
  items.forEach(function(item) {
    if(!itemsByCategory[item.category]) itemsByCategory[item.category] = [];
    itemsByCategory[item.category].push(item);
  });

  Object.getOwnPropertyNames(itemsByCategory).forEach(function(category) {
    var categoryAttachment = {
      author_name: user + " - " + action,
      title: category,
      title_link: projectSheetUrl,
      color: projectColor,
      fields: itemsByCategory[category].map(function(item) {
        var itemField = {
          title: truncateString(item.name, 45),
          value: "$" + item.totalPrice.toFixed(2) + "\n\t (" + item.quantity + "x @ $" + item.unitPrice.toFixed(2) + "/e)",
          short: "true"
        };

        if(item.supplier || item.productNum) itemField.value += "\n\t";
        // Links seem to be broken :(
        //if(item.link && (item.supplier || item.productNum)) itemField.value += "<" + item.link + "|";
        if(item.productNum) itemField.value += "`#" + item.productNum + "`";
        if(item.supplier) itemField.value += " from " + item.supplier;
        //if(item.link && (item.supplier || item.productNum)) itemField.value += ">";

        if(item.requestorComments) itemField.value += "\n Requestor Comment: \n> _" + item.requestorComments + "_";
        if(item.officerComments) itemField.value += "\n Officer Comment: \n> _" + item.officerComments + "_";

        return itemField;
      }),
      footer: projectName,
      footer_icon: OPTS.SLACK.SOAR_ICON,
      mrkdwn_in: ["fields"]
    };

    itemListMessage.attachments.push(categoryAttachment);
  });

  attachment.value = JSON.stringify(itemListMessage);

  if(attachment.value.length >= 2000) {
    itemListMessage.text = "Sorry, there were too many items to list. Open the project sheet to view them instead. https://github.com/usfsoar/purchasing-manager/issues/6";
    itemListMessage.attachments = [];
    attachment.value = JSON.stringify(itemListMessage);
  }

  return attachment;
}

/**
 * Get the full name of the project that matches the name of the sheet.
 * @param {string} sheetName Name of the project's sheet.
 * @param {boolean} nullIfMissing If present, will return null if the value is not found.
 * @returns {?string} String, unless nullIfMissing is specified.
 */
function getProjectNameFromSheetName(sheetName, nullIfMissing) {
  var projectsData = SpreadsheetApp
    .getActiveSpreadsheet()
    .getRange(OPTS.NAMED_RANGES.PROJECT_NAMES_TO_SHEETS)
    .getValues();

  for(var i = 0; i < projectsData.length; i++) {
    if(projectsData[i][1] === sheetName) return projectsData[i][0];
  }

  if(nullIfMissing) return null;
  return '_Error: Project Not Found_';
}

/**
 * Get the sheet name that matches the project.
 * @param {string} projectName Name of the project.
 * @param {boolean} nullIfMissing If present, will return null if the value is not found.
 * @returns {?string} String, unless nullIfMissing is specified.
 */
function getSheetNameFromProjectName(projectName, nullIfMissing) {
  var projectsData = SpreadsheetApp
    .getActiveSpreadsheet()
    .getRange(OPTS.NAMED_RANGES.PROJECT_NAMES_TO_SHEETS)
    .getValues();

  for(var i = 0; i < projectsData.length; i++) {
    if(projectsData[i][0] === projectName) return projectsData[i][1];
  }

  if(nullIfMissing) return null;
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

/** Mark selected items in the sheet as received. */
function markSelectedReceived() {
  markItems(STATUSES_DATA.RECEIVED);
}

/** Mark selected items in the sheet as received and request reimbursement. */
function markSelectedReceivedReimburse() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.alert('Confirm',
    'NOTE: Reimbursements are not guarunteed and MUST be preapproved. Items must be received before reimbursement will be sent. If at all possible, items should be purchased by a financial officer. You are required to put your PayPal email address in the "Requestor Comments" field, and only the original item requestor can be reimbursed. Are you sure you want to continue?',
    ui.ButtonSet.OK_CANCEL);
  if (response === ui.Button.CANCEL)
    return;
  markItems(STATUSES_DATA.RECEIVED_REIMBURSE);
}

/** Mark selected items in the sheet as reimbursed. */
function markSelectedReimbursed() {
  markItems(STATUSES_DATA.REIMBURSED);
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

/** Fast-forward selected items in the sheet to received. */
function fastForwardSelectedReceived() {
  fastForwardItems(STATUSES_DATA.RECEIVED);
}

/** Fast-forward selected items in the sheet to received and request reimbursement. */
function fastForwardSelectedReceivedReimburse() {
  fastForwardItems(STATUSES_DATA.RECEIVED_REIMBURSE);
}

/** Fast-forward selected items in the sheet to reimbursed. */
function fastForwardSelectedReimbursed() {
  fastForwardItems(STATUSES_DATA.REIMBURSED);
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

/** Test post a message in the DEV channel. */
function testUpdateItem() {
  markItems(TEST_STATUS);
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

  /*SpreadsheetApp.getActiveSpreadsheet().getProtections(SpreadsheetApp.ProtectionType.RANGE)
      .forEach(function(protection) {
        protection.remove();
      });*/

  SpreadsheetApp.getActiveSpreadsheet().getProtections(SpreadsheetApp.ProtectionType.SHEET)
      .forEach(function(protection) {
        protection.remove();
      });

  sheets.forEach(function(sheet) {
    var sheetName = sheet.getName();

    if(projectSheetNames.indexOf(sheetName) !== -1) {
      // The below code was taking too long so now we maintain range protections manually

      // Lock certain sections of project sheets (only the headers and formula-driven parts)
      /*var numDataRows = sheet.getLastRow() - OPTS.NUM_HEADER_ROWS;

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
      financialOfficerRangeProtection.addEditors(financialOfficers);*/
    } else if(sheetName !== userDataSheetName) {
      // Lock the entire sheet if not the user data sheet
      var sheetProtection = sheet.protect();
      sheetProtection.removeEditors(sheetProtection.getEditors());
      sheetProtection.addEditors(financialOfficers);
      successNotification("Updated protections for " + sheetName);
    }
  });

  // Protect the statuses, since they need to match the values in the script
  /*var statusesProtection = SpreadsheetApp.getActiveSpreadsheet()
      .getRangeByName(OPTS.NAMED_RANGES.STATUSES).protect();
  statusesProtection.setDescription(adminProtectDescription);
  statusesProtection.removeEditors(statusesProtection.getEditors());
  statusesProtection.addEditor(admin);*/
}

/**
 * Show option to open the folder or the file.
 * @param {GoogleAppsScript.Drive.File} spreadsheet
 * @param {GoogleAppsScript.Drive.Folder} folder
 */
function openFile(spreadsheet, folder) {
  var spreadsheetId = spreadsheet.getId();
  var folderId = folder.getId();
  var fileUrl = "https://docs.google.com/spreadsheets/d/"+spreadsheetId;
  var folderUrl = "https://drive.google.com/drive/u/2/folders/"+folderId;
  var html = "Succesfully sent items to sheet.<br><a target='_blank' href='" + folderUrl + "'>Open Purchasing Sheets Folder</a><br><a target='_blank' href='" + fileUrl + "'>Open The New Purchasing Sheet</a>";
  var userInterface = HtmlService.createHtmlOutput(html);
  SpreadsheetApp.getUi().showModalDialog(userInterface, "Open Sheet");
}

/** Send the selected items to a new purchasing sheet. */
function sendSelectedToSheet() {
  if(!checkIfProjectSheet() || !verifyFinancialOfficer()) return;
  var selectedRanges = getSelectedRows();

  var totalRowCount = selectedRanges.reduce(function(total, currentRange) {
    return total + currentRange;
  }, 0);
  if(totalRowCount > 12 || totalRowCount < 1) {
    errorNotification("Can only send 1-12 rows at a time to a purchasing sheet.");
    return;
  }

  var currentSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  var purchasesFolderId = currentSpreadsheet
    .getSheetByName(OPTS.SHEET_NAMES.MAIN_DASHBOARD)
    .getRange(
      OPTS.DASHBOARD_CELLS.PURCHASES_FOLDER.row, 
      OPTS.DASHBOARD_CELLS.PURCHASES_FOLDER.column
    )
    .getValue();
  var targetFolder = DriveApp.getFolderById(purchasesFolderId);

  var currentSheet = currentSpreadsheet.getActiveSheet();
  var template = currentSpreadsheet.getSheetByName(OPTS.SHEET_NAMES.PURCHASING_TEMPLATE);
  var newSheet = template.copyTo(currentSpreadsheet);

  var projectName = getProjectNameFromSheetName(currentSheet.getSheetName());
  
  var officer = getCurrentUserInfo();
  newSheet.getRange("I6").setValue(officer.fullName);
  newSheet.getRange("I7").setValue(officer.email);
  newSheet.getRange("I8").setValue(officer.phone);

  newSheet.getRange("F14").setValue(projectName);

  var needBy = moment().add(2, "weeks").format("MM/DD/YY");
  newSheet.getRange("M38").setValue(needBy);

  var vendor = selectedRanges[0].getValues()[0][OPTS.ITEM_COLUMNS.SUPPLIER.index - 1];
  newSheet.getRange("J42").setValue(vendor);

  var allHaveSameVendor = true;
  var allNew = true;
  var index = 50;
  selectedRanges.forEach(function(range) {
    range.getValues().forEach(function(row) {
      if(row[OPTS.ITEM_COLUMNS.SUPPLIER.index - 1] !== vendor) {
        allHaveSameVendor = false;
      }
      if(row[OPTS.ITEM_COLUMNS.STATUS.index - 1] !== STATUSES_DATA.NEW.text) {
        allNew = false;
      }
      var itemName = row[OPTS.ITEM_COLUMNS.NAME.index - 1];
      newSheet.getRange(index, 2).setValue(itemName);
      var link = row[OPTS.ITEM_COLUMNS.LINK.index - 1];
      newSheet.getRange(index, 8).setValue(link);
      var qty = row[OPTS.ITEM_COLUMNS.QUANTITY.index - 1];
      newSheet.getRange(index, 13).setValue(qty);
      var unitPrice = row[OPTS.ITEM_COLUMNS.UNIT_PRICE.index - 1];
      newSheet.getRange(index, 15).setValue(unitPrice);
      index++;
    });
  });

  if(!allNew) {
    errorNotification("One or more items was not 'New'!")
    currentSpreadsheet.deleteSheet(newSheet);
    return;
  }

  if(!allHaveSameVendor) {
    errorNotification("The items selected do not all have the same vendor!");
    currentSpreadsheet.deleteSheet(newSheet);
    return;
  }

  var sheetName = moment().format("YY-MM-DD") + " - " + projectName + " - " + vendor;
  newSheet.setName(sheetName);

  var newSpreadsheet = SpreadsheetApp.create(sheetName);
  var file = DriveApp.getFileById(newSpreadsheet.getId());
  var parents = file.getParents();
  while(parents.hasNext()) {
    parents.next().removeFile(file);
  }
  targetFolder.addFile(file);

  file.setName(sheetName);
  newSheet.copyTo(newSpreadsheet);
  newSpreadsheet.deleteSheet(newSpreadsheet.getSheetByName("Sheet1"));
  currentSpreadsheet.deleteSheet(newSheet);
  openFile(file, targetFolder);
}