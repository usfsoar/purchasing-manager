import { getNamedRangeValues } from "./spreadsheet_utils";

/** @file Global configuration. */
export default {
  /** Named Ranges throughout the spreadsheet. */
  NAMED_RANGES: {
    /** Range containing the email addresses of approved officers.
     * 1 column, 5 rows (no header). */
    APPROVED_OFFICERS: "ApprovedOfficers",
    /** Range containing the YES or NO values for whether to notify the
     * financial officer with the same index. 1 column, 5 rows (no header). */
    NOTIFY_APPROVED_OFFICERS: "NotifyApprovedOfficers",
    /** Range containing the names of all project-specific sheets.
     * 1 column, 12 rows (no header). */
    PROJECT_SHEETS: "ProjectSheets",
    PROJECT_NAMES_TO_SHEETS: "ProjectNamesToSheets",
    STATUSES: "Statuses",
    /** Range containing the ID of the purchasing sheets folder. 1 cell. */
    PURCHASING_SHEETS_FOLDER_ID: "PurchasingSheetsFolderID",
  },
  /** Custom Menu info. */
  CUSTOM_MENU: {
    NAME: "SOAR Purchasing",
  },
  FAST_FORWARD_MENU: {
    NAME: "Fast-Forward",
  },
  /** The number of header rows in the project sheets. */
  NUM_HEADER_ROWS: 2,
  /**
   * Relevant columns in the project sheets, as 1-based indexes.
   * @enum {Column}
   */
  ITEM_COLUMNS: {
    STATUS: { index: 1, name: "Status" },
    NAME: { index: 2, name: "Name" },
    SUPPLIER: { index: 3, name: "Supplier" },
    PRODUCT_NUM: { index: 4, name: "Product Number" },
    LINK: { index: 5, name: "Link" },
    UNIT_PRICE: { index: 6, name: "Unit Price" },
    QUANTITY: { index: 7, name: "Quantity" },
    SHIPPING: { index: 8, name: "Shipping Price" },
    TOTAL_PRICE: { index: 9, name: "Total Price" },
    CATEGORY: { index: 10, name: "Category" },
    REQUEST_COMMENTS: { index: 11, name: "Request Comments" },
    REQUEST_EMAIL: { index: 12, name: "Requestor Email" },
    REQUEST_DATE: { index: 13, name: "Request Date" },
    OFFICER_EMAIL: { index: 14, name: "Financial Officer Email" },
    OFFICER_COMMENTS: { index: 15, name: "Financial Officer Comments" },
    ACCOUNT: { index: 16, name: "Purchasing Account" },
    REQUEST_ID: { index: 17, name: "Request ID" },
    SUBMIT_DATE: { index: 18, name: "Submit Date" },
    UPDATE_DATE: { index: 19, name: "Update Date" },
    ARRIVE_DATE: { index: 20, name: "Arrival Date" },
    RECEIVE_EMAIL: { index: 21, name: "receive Date" },
    RECEIVE_DATE: { index: 22, name: "receiver Email" },
  },
  /** Options relating to the user interface. */
  UI: {
    /** Typical toast length in seconds. */
    TOAST_DURATION: 5,
    TOAST_TITLES: {
      ERROR: "Error!",
      SUCCESS: "Completed",
      WARNING: "Alert!",
      INFO: "Note",
    },
    SLACK_ID_PROMPT:
      "Looks like this is your first time using the SOAR purchasing database. Please enter your Slack Member ID # (NOT your username!) found in your Slack profile, in the dropdown menu. For more details, see detailed instructions at:\nhttps://drive.google.com/open?id=1Q1PleYhE1i0A5VFyjKqyLswom3NQuXcn.",
    FULL_NAME_PROMPT:
      "Great, thank you! Please also enter your full name. You won't have to do this next time.",
  },
  /** Default values for items. */
  DEFAULT_VALUES: {
    ACCOUNT_NAME: getNamedRangeValues("Accounts")[0],
    CATEGORY: "Uncategorized",
  },
  /** Names of sheets in the Spreadsheet */
  SHEET_NAMES: {
    USERS: "Users",
    PURCHASING_TEMPLATE: "Purchasing Sheet Template",
    MAIN_DASHBOARD: "Main Dashboard",
  },
  DASHBOARD_CELLS: {
    TOTAL_BUDGET: {
      row: 4, // 1-based index!
      column: 3,
    },
    TOTAL_EXPENSES: {
      row: 4,
      column: 4,
    },
    PROJECT_DESCRIPTION: {
      row: 11,
      column: 3,
    },
  },
  /** Slack API pieces */
  SLACK: {
    CHECK_MARK_EMOJI: ":heavy_check_mark:",
    /** Possible cases for target users to tag in messages. */
    TARGET_USERS: {
      /** The entire channel. */
      CHANNEL: "CHANNEL",
      /**
       * Just the people who requested said items (can be multiple if multiple)
       * items are affected.
       */
      REQUESTORS: "REQUESTORS",
      /** Just all the listed Financial Officers. */
      OFFICERS: "OFFICERS",
    },
    STATUS_SLASH_COMMAND: "budgetstatus",
    ITEM_LIST_ACTION_NAME_LEGACY: "listItems",
    ITEM_LIST_ACTION_NAME: "showItemList",
    SOAR_ICON:
      "http://www.usfsoar.com/wp-content/uploads/2018/09/595bae9a-c1f9-4b46-880e-dc6d4e1d0dac.png",
  },
  /** Number of adjacent officer columns in the project sheets. */
  NUM_OFFICER_COLS: 7,
} as const;
