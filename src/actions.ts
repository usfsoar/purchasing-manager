import { getCurrentUserInfo, verifyFinancialOfficer } from "./authorization";
import OPTS from "./config";
import { Status } from "./interfaces";
import * as notifications from "./notifications";
import { slackNotifyItems } from "./slack_utils";
import {
  checkIfProjectSheet,
  getAllRows,
  getColumnRange,
  getNamedRangeValues,
  getProjectNameFromSheetName,
  getSelectedRows,
} from "./spreadsheet_utils";
import STATUSES_DATA, { isStatusAllowedAsNext } from "./statuses_config";
import { makeListFromArray, wrapInDoubleQuotes } from "./utils";

/**
 * Check if the given row has data and the data is valid for the desired operation.
 * If the validation fails, alerts the user. Does not check row statuses;
 * rows with incorrect statuses are skipped silently.
 * @param rowValues The current data for the row, in order.
 * @param newStatus The new status of the row for testing against.
 * @return `true` if the row is valid and can be submitted.
 */
function validateRow(
  rowValues: (string | number | Date)[],
  newStatus: Status
): boolean {
  const missingReccomendedColumns =
    newStatus.reccomendedColumns
      ?.filter((column) => rowValues[column.index - 1] === "")
      .map((column) => column.name) ?? [];

  if (missingReccomendedColumns.length > 1) {
    notifications.warn(
      `One or more items is missing a value for "${makeListFromArray(
        missingReccomendedColumns
      )}". Will mark anyway with default value.`
    );
  }

  const missingRequiredColumns =
    newStatus.requiredColumns
      ?.filter((column) => rowValues[column.index - 1] === "")
      .map((column) => column.name) ?? [];

  if (missingRequiredColumns.length > 1) {
    notifications.error(
      `Cannot submit: one or more items is missing values for ${makeListFromArray(
        missingRequiredColumns
      )}. This value is required.`
    );
    return false;
  }
  return true;
}

/**
 * Mark all of the items in the currently selected rows as `newStatus` if they
 * are currently one of the allowed previous statuses, and also fill in the date
 * and attribution columns.
 * @param newStatus The object representing the status to change the
 * selected items to.
 * @param markAll If truthy, mark all possible rows, else mark
 * selected.
 */
function markItems(newStatus: Status, markAll = false): void {
  if (
    !checkIfProjectSheet() ||
    (newStatus.officersOnly && !verifyFinancialOfficer())
  ) {
    return;
  }

  const rangesToMark = markAll ? getAllRows() : getSelectedRows();

  const currentUser = getCurrentUserInfo();
  const currentDate = new Date();

  const currentSheet = SpreadsheetApp.getActiveSheet();
  const projectName = getProjectNameFromSheetName(currentSheet.getSheetName());
  const projectSheetUrl = `${SpreadsheetApp.getActiveSpreadsheet().getUrl()}#gid=${currentSheet.getSheetId()}`;
  const itemRequestors = new Set<string>();

  // We would filter out all the rows with disallowed current statuses here,
  // rather than skipping them in both of these loops, but that would require
  // modifying the ranges, which is much more time-intensive than just skipping.

  for (const range of rangesToMark) {
    const rangeValues = range.getValues();
    for (const row of rangeValues) {
      // If current status is not in allowed statuses, don't verify, just skip.
      if (
        !isStatusAllowedAsNext(
          row[OPTS.ITEM_COLUMNS.STATUS.index - 1].toString(),
          newStatus
        )
      ) {
        continue;
      }

      // No need to alert the user on fail; validateRow will do that itself.
      // Otherwise validate. If a single row is invalid, quit
      if (!validateRow(row, newStatus)) return;
    }
  }

  // Caching column values to write in bulk at the end:
  let userColumn;
  let dateColumn;
  let userColumnValues;
  let dateColumnValues;
  if (newStatus.columns.user) {
    userColumn = getColumnRange(newStatus.columns.user.index);
    userColumnValues = getColumnRange(newStatus.columns.user.index).getValues();
  }
  if (newStatus.columns.date) {
    dateColumn = getColumnRange(newStatus.columns.date.index);
    dateColumnValues = getColumnRange(newStatus.columns.date.index).getValues();
  }

  let accountColumn;
  let categoryColumn;
  let accountColumnValues;
  let categoryColumnValues;
  if (newStatus.fillInDefaults) {
    accountColumn = getColumnRange(OPTS.ITEM_COLUMNS.ACCOUNT.index);
    categoryColumn = getColumnRange(OPTS.ITEM_COLUMNS.CATEGORY.index);
    accountColumnValues = accountColumn.getValues();
    categoryColumnValues = categoryColumn.getValues();
  }

  // Read (not modify, so no need for range) the requestor data for notifying
  const requestorColumnValues =
    newStatus.slack.targetUsers === OPTS.SLACK.TARGET_USERS.REQUESTORS
      ? getColumnRange(OPTS.ITEM_COLUMNS.REQUEST_EMAIL.index).getValues()
      : null;

  /* List of items, for sending to Slack */
  const items = [];

  // Cache the entire columns, to avoid making dozens of calls to the server
  const statusColumn = getColumnRange(OPTS.ITEM_COLUMNS.STATUS.index);
  const statusColumnValues = statusColumn.getValues();

  // Loop through the ranges
  for (const range of rangesToMark) {
    const rangeStartIndex = range.getRow() - 1;
    const rangeValues = range.getValues();
    const rangeLength = range.getNumRows();

    // Loop through the rows in the range
    for (let l = 0; l < rangeLength; l++) {
      /** The index (not number) of the current row in the spreadsheet. */
      const currentSheetRowIndex = rangeStartIndex + l;
      /**
       * The index of the current value row in the spreadsheet, with the first
       * row after the headers being 0.
       */
      const currentValuesRowIndex = currentSheetRowIndex - OPTS.NUM_HEADER_ROWS;

      // If this row's status is not in allowed statuses, don't verify, just
      // skip
      if (
        !isStatusAllowedAsNext(
          statusColumnValues[currentValuesRowIndex][0].toString(),
          newStatus
        )
      ) {
        continue;
      }

      // Update values in local cache
      // These ranges don't include the header, so 0 in the range is
      // OPTS.NUM_HEADER_ROWS in the spreadsheet
      statusColumnValues[currentValuesRowIndex][0] = newStatus.text;

      if (newStatus.columns.user && userColumnValues) {
        userColumnValues[currentValuesRowIndex][0] = currentUser.email;
      }
      if (newStatus.columns.date && dateColumnValues) {
        dateColumnValues[currentValuesRowIndex][0] = currentDate;
      }

      if (
        newStatus.fillInDefaults &&
        accountColumnValues &&
        categoryColumnValues
      ) {
        if (accountColumnValues[currentValuesRowIndex][0].toString() === "") {
          accountColumnValues[currentValuesRowIndex][0] = getNamedRangeValues(
            OPTS.NAMED_RANGES.ACCOUNTS
          )[0];
        }
        if (categoryColumnValues[currentValuesRowIndex][0].toString() === "") {
          categoryColumnValues[currentValuesRowIndex][0] =
            OPTS.DEFAULT_VALUES.CATEGORY;
        }
      }

      // Save the requestor data for notifying; avoid duplicates
      if (
        newStatus.slack.targetUsers === OPTS.SLACK.TARGET_USERS.REQUESTORS &&
        requestorColumnValues
      ) {
        itemRequestors.add(requestorColumnValues[currentValuesRowIndex][0]);
      }

      items.push({
        name: rangeValues[l][OPTS.ITEM_COLUMNS.NAME.index - 1],
        quantity: rangeValues[l][OPTS.ITEM_COLUMNS.QUANTITY.index - 1],
        totalPrice: rangeValues[l][OPTS.ITEM_COLUMNS.TOTAL_PRICE.index - 1],
        unitPrice: rangeValues[l][OPTS.ITEM_COLUMNS.UNIT_PRICE.index - 1],
        category: rangeValues[l][OPTS.ITEM_COLUMNS.CATEGORY.index - 1],
        requestorComments:
          rangeValues[l][OPTS.ITEM_COLUMNS.REQUEST_COMMENTS.index - 1],
        officerComments:
          rangeValues[l][OPTS.ITEM_COLUMNS.OFFICER_COMMENTS.index - 1],
        supplier: rangeValues[l][OPTS.ITEM_COLUMNS.SUPPLIER.index - 1],
        productNum: rangeValues[l][OPTS.ITEM_COLUMNS.PRODUCT_NUM.index - 1],
        link: rangeValues[l][OPTS.ITEM_COLUMNS.LINK.index - 1],
      });
    }
  }

  // Write the cached values
  statusColumn.setValues(statusColumnValues);

  if (newStatus.columns.user && userColumn && userColumnValues) {
    getColumnRange(newStatus.columns.user.index).setValues(userColumnValues);
  }
  if (newStatus.columns.date && dateColumn && dateColumnValues) {
    getColumnRange(newStatus.columns.date.index).setValues(dateColumnValues);
  }

  if (
    newStatus.fillInDefaults &&
    accountColumn &&
    accountColumnValues &&
    categoryColumn &&
    categoryColumnValues
  ) {
    accountColumn.setValues(accountColumnValues);
    categoryColumn.setValues(categoryColumnValues);
  }

  if (items.length > 0) {
    notifications.success(
      `${items.length} items marked from ${makeListFromArray(
        newStatus.allowedPrevious.map(wrapInDoubleQuotes),
        "or"
      )} to "${newStatus.text}."`
    );

    slackNotifyItems(
      newStatus,
      currentUser.fullName,
      Array.from(itemRequestors.values()),
      items,
      projectName ?? "_Error: project not found._",
      projectSheetUrl,
      currentSheet.getTabColor() ?? "#000000"
    );
  } else {
    notifications.error("No valid items selected for that action.");
  }
}

/**
 * Fast-forward all of the items in the currently selected rows to `newStatus`,
 * filling in the date and attribution columns but not notifying on Slack.
 * Allows for skipping statuses
 * @param newStatus The object representing the status to fast-forward the
 * selected items to.
 */
function fastForwardItems(newStatus: Status): void {
  if (!checkIfProjectSheet() || !verifyFinancialOfficer()) return;

  const selectedRanges = getSelectedRows();

  let numMarked = 0;
  const currentOfficer = getCurrentUserInfo();
  const currentOfficerEmail = currentOfficer.email;
  const currentDate = new Date();

  // Cache the entire columns, to avoid making dozens of calls to the server
  const statusColumn = getColumnRange(OPTS.ITEM_COLUMNS.STATUS.index);
  const statusColumnValues = statusColumn.getValues();

  // Fetch normal columns to update
  let userColumn;
  let dateColumn;
  let userColumnValues;
  let dateColumnValues;
  if (newStatus.columns.user) {
    userColumn = getColumnRange(newStatus.columns.user.index);
    userColumnValues = userColumn.getValues();
  }
  if (newStatus.columns.date) {
    dateColumn = getColumnRange(newStatus.columns.date.index);
    dateColumnValues = dateColumn.getValues();
  }

  // Fetch default columns to fill if empty
  let accountColumn;
  let categoryColumn;
  let accountColumnValues;
  let categoryColumnValues;
  if (newStatus.fillInDefaults) {
    accountColumn = getColumnRange(OPTS.ITEM_COLUMNS.ACCOUNT.index);
    categoryColumn = getColumnRange(OPTS.ITEM_COLUMNS.CATEGORY.index);
    accountColumnValues = accountColumn.getValues();
    categoryColumnValues = categoryColumn.getValues();
  }

  // Fetch fast-forward columns to fill if empty
  const pastUserColumns = (
    newStatus.fastForwardColumns?.user ?? []
  ).map((ffCol) => getColumnRange(ffCol.index));
  const pastDateColumns = (
    newStatus.fastForwardColumns?.date ?? []
  ).map((ffCol) => getColumnRange(ffCol.index));
  const pastUserColumnsValues = pastUserColumns.map((colRange) =>
    colRange.getValues()
  );
  const pastDateColumnsValues = pastDateColumns.map((colRange) =>
    colRange.getValues()
  );

  // Loop through the ranges
  for (const range of selectedRanges) {
    const rangeStartIndex = range.getRow() - 1;

    // Loop through the rows in the range
    for (let l = 0; l < range.getNumRows(); l++) {
      /** The index (not number) of the current row in the spreadsheet. */
      const currentSheetRowIndex = rangeStartIndex + l;
      /**
       * The index of the current value row in the spreadsheet, with the first
       * row after the headers being 0.
       */
      const currentValuesRowIndex = currentSheetRowIndex - OPTS.NUM_HEADER_ROWS;

      // Update values in local cache
      // These ranges don't include the header, so 0 in the range is
      // OPTS.NUM_HEADER_ROWS in the spreadsheet
      statusColumnValues[currentValuesRowIndex][0] = newStatus.text;

      if (newStatus.columns.user && userColumnValues) {
        userColumnValues[currentValuesRowIndex][0] = currentOfficerEmail;
      }
      if (newStatus.columns.date && dateColumnValues) {
        dateColumnValues[currentValuesRowIndex][0] = currentDate;
      }

      if (
        newStatus.fillInDefaults &&
        accountColumnValues &&
        categoryColumnValues
      ) {
        if (accountColumnValues[currentValuesRowIndex][0].toString() === "") {
          accountColumnValues[currentValuesRowIndex][0] = getNamedRangeValues(
            OPTS.NAMED_RANGES.ACCOUNTS
          )[0];
        }
        if (categoryColumnValues[currentValuesRowIndex][0].toString() === "") {
          categoryColumnValues[currentValuesRowIndex][0] =
            OPTS.DEFAULT_VALUES.CATEGORY;
        }
      }

      // If any of the past columns are blank, fill them in with current info
      pastUserColumnsValues.forEach((columnValues) => {
        if (columnValues[currentValuesRowIndex][0].toString() === "") {
          columnValues[currentValuesRowIndex][0] = currentOfficerEmail;
        }
      });
      pastDateColumnsValues.forEach((columnValues) => {
        if (columnValues[currentValuesRowIndex][0].toString() === "") {
          columnValues[currentValuesRowIndex][0] = currentDate;
        }
      });

      numMarked++;
    }
  }

  // Write the cached values
  statusColumn.setValues(statusColumnValues);

  if (newStatus.columns.user && userColumn && userColumnValues) {
    userColumn.setValues(userColumnValues);
  }
  if (newStatus.columns.date && dateColumn && dateColumnValues) {
    dateColumn.setValues(dateColumnValues);
  }

  if (
    newStatus.fillInDefaults &&
    accountColumn &&
    categoryColumn &&
    accountColumnValues &&
    categoryColumnValues
  ) {
    accountColumn.setValues(accountColumnValues);
    categoryColumn.setValues(categoryColumnValues);
  }

  pastUserColumns.forEach((columnRange, index) =>
    columnRange.setValues(pastUserColumnsValues[index])
  );
  pastDateColumns.forEach((columnRange, index) =>
    columnRange.setValues(pastDateColumnsValues[index])
  );

  notifications.success(
    `${numMarked} items fast-forwarded to "${newStatus.text}."`
  );
}

/** Mark selected items in the sheet as received and request reimbursement. */
export function markSelectedReceivedReimburse(): void {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    "Confirm",
    'NOTE: Reimbursements are not guarunteed and MUST be preapproved. Items must be received before reimbursement will be sent. If at all possible, items should be purchased by a financial officer. You are required to put your PayPal email address in the "Requestor Comments" field, and only the original item requestor can be reimbursed. Are you sure you want to continue?',
    ui.ButtonSet.OK_CANCEL
  );
  if (response === ui.Button.CANCEL) return;
  markItems(STATUSES_DATA.RECEIVED_REIMBURSE);
}

/** Mark selected items in the sheet as reimbursed. */
export function markSelectedReimbursed(): void {
  markItems(STATUSES_DATA.REIMBURSED);
}

/** Mark selected items in the sheet as submitted. */
export function markSelectedSubmitted(): void {
  markItems(STATUSES_DATA.SUBMITTED);
}

/** Mark selected items in the sheet as approved. */
export function markSelectedApproved(): void {
  markItems(STATUSES_DATA.APPROVED);
}

/** Mark selected items in the sheet as arrived / awaiting pickup. */
export function markSelectedAwaitingPickup(): void {
  markItems(STATUSES_DATA.AWAITING_PICKUP);
}

/** Mark selected items in the sheet as awaiting info. */
export function markSelectedAwaitingInfo(): void {
  markItems(STATUSES_DATA.AWAITING_INFO);
}

/** Mark selected items in the sheet as denied. */
export function markSelectedDenied(): void {
  markItems(STATUSES_DATA.DENIED);
}

/** Fast-forward the selected items in the sheet to new. */
export function fastForwardSelectedNew(): void {
  fastForwardItems(STATUSES_DATA.NEW);
}

/** Fast-forward selected items in the sheet to received. */
export function fastForwardSelectedReceived(): void {
  fastForwardItems(STATUSES_DATA.RECEIVED);
}

/** Fast-forward selected items in the sheet to received and request reimbursement. */
export function fastForwardSelectedReceivedReimburse(): void {
  fastForwardItems(STATUSES_DATA.RECEIVED_REIMBURSE);
}

/** Fast-forward selected items in the sheet to reimbursed. */
export function fastForwardSelectedReimbursed(): void {
  fastForwardItems(STATUSES_DATA.REIMBURSED);
}

/** Fast-forward selected items in the sheet to submitted. */
export function fastForwardSelectedSubmitted(): void {
  fastForwardItems(STATUSES_DATA.SUBMITTED);
}

/** Fast-forward selected items in the sheet to approved. */
export function fastForwardSelectedApproved(): void {
  fastForwardItems(STATUSES_DATA.APPROVED);
}

/** Fast-forward selected items in the sheet to arrived / awaiting pickup. */
export function fastForwardSelectedAwaitingPickup(): void {
  fastForwardItems(STATUSES_DATA.AWAITING_PICKUP);
}

/** Fast-forward selected items in the sheet to awaiting info. */
export function fastForwardSelectedAwaitingInfo(): void {
  fastForwardItems(STATUSES_DATA.AWAITING_INFO);
}

/** Fast-forward selected items in the sheet to denied. */
export function fastForwardSelectedDenied(): void {
  fastForwardItems(STATUSES_DATA.DENIED);
}

/** Mark the selected items in the sheet as new. */
export function markSelectedNew(): void {
  markItems(STATUSES_DATA.NEW);
}

/** Mark all possible items in the sheet as new. */
export function markAllNew(): void {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    "Confirm",
    "Are you sure you want to mark ALL unmarked items in this project as new?",
    ui.ButtonSet.OK_CANCEL
  );
  if (response === ui.Button.CANCEL) return;
  markItems(STATUSES_DATA.NEW, true);
}

/** Mark selected items in the sheet as received. */
export function markSelectedReceived(): void {
  markItems(STATUSES_DATA.RECEIVED);
}
