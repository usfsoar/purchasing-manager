import { DateTime } from "luxon";
import { getCurrentUserInfo, verifyFinancialOfficer } from "./authorization";
import OPTS from "./config";
import * as notifications from "./notifications";
import {
  checkIfProjectSheet,
  getNamedRangeValues,
  getProjectNameFromSheetName,
  getSelectedRows,
} from "./spreadsheet_utils";
import STATUSES_DATA from "./statuses_config";
import { escapeSingleQuotes, escapeSpaces } from "./utils";

/**
 * Show option to open the folder or the file.
 */
export function openFile(
  spreadsheet: GoogleAppsScript.Drive.File,
  folder: GoogleAppsScript.Drive.Folder,
  vendorName: string
): void {
  const vendor =
    escapeSpaces(
      encodeURIComponent(escapeSingleQuotes(vendorName.toUpperCase()))
    ) || "VENDOR+NAME";
  const spreadsheetId = spreadsheet.getId();
  const folderId = folder.getId();
  const fileUrl = `https://docs.google.com/spreadsheets/d/${spreadsheetId}`;
  const folderUrl = `https://drive.google.com/drive/u/2/folders/${folderId}`;
  const currentUserEmail = escapeSingleQuotes(
    encodeURIComponent(getCurrentUserInfo().email)
  );
  const emailTemplateLink = `https://mail.google.com/mail/u/0/?view=cm&fs=1&to=sg-rmdpurchase@usf.edu&authuser=${currentUserEmail}&su=SOCIETY+OF+AERONAUTICS+AND+ROCKETRY,+${vendor}&body=Please+see+attached+purchasing+form.&tf=1`;
  const html = `<div style='font-family: sans-serif;'>Successfully sent items to sheet.<br><a target='_blank' href='${folderUrl}'>Open Purchasing Sheets Folder</a><br><a target='_blank' href='${fileUrl}'>Open The New Purchasing Sheet</a><br /><br /><strong>Remember:</strong> Attach the form to an email sent from your @mail.usf.edu email adress to sg-rmdpurchase@usf.edu with subject \\"SOCIETY OF AERONAUTICS AND ROCKETRY, VENDOR NAME, EVENT DATE (if applicable)\\". Only one form per email! <br><br> <a  href='${emailTemplateLink}' target='_blank'><strong>Start Email in Gmail</strong></a></div>`;
  const userInterface = HtmlService.createHtmlOutput(html);
  SpreadsheetApp.getUi().showModalDialog(userInterface, "Open Sheet");
}

/** Send the selected items to a new purchasing sheet. */
export function sendSelectedToSheet(): void {
  if (!checkIfProjectSheet() || !verifyFinancialOfficer()) return;
  const selectedRanges = getSelectedRows();

  const totalRowCount = selectedRanges.reduce((total, currentRange) => {
    return total + currentRange.getNumRows();
  }, 0);

  if (totalRowCount > 12 || totalRowCount < 1) {
    notifications.error(
      "Can only send 1-12 rows at a time to a purchasing sheet."
    );
    return;
  }

  const currentSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  const purchasesFolderId = getNamedRangeValues(
    OPTS.NAMED_RANGES.PURCHASING_SHEETS_FOLDER_ID
  )[0];
  const targetFolder = DriveApp.getFolderById(purchasesFolderId);

  const currentSheet = currentSpreadsheet.getActiveSheet();
  const template = currentSpreadsheet.getSheetByName(
    OPTS.SHEET_NAMES.PURCHASING_TEMPLATE
  );
  if (template === null) throw new Error("No purchasing sheet template found.");
  const newSheet = template.copyTo(currentSpreadsheet);

  const projectSheetName = currentSheet.getSheetName();
  const projectName = getProjectNameFromSheetName(projectSheetName);
  const dashboardSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(
    projectSheetName + " Dashboard"
  );
  if (dashboardSheet === null) {
    throw new Error("No dashboard sheet found for project.");
  }
  const projectDescription = dashboardSheet
    .getRange(
      OPTS.DASHBOARD_CELLS.PROJECT_DESCRIPTION.row,
      OPTS.DASHBOARD_CELLS.PROJECT_DESCRIPTION.column
    )
    .getValue();
  const officer = getCurrentUserInfo();
  newSheet.getRange("I6").setValue(officer.fullName);
  newSheet.getRange("I7").setValue(officer.email);
  newSheet.getRange("I8").setValue(officer.phone);

  newSheet.getRange("F14").setValue(projectName);
  newSheet.getRange("A21").setValue(projectDescription);

  const needBy = DateTime.local().plus({ weeks: 2 }).toFormat("MM/dd/yy");
  newSheet.getRange("M38").setValue(needBy);

  const vendor = selectedRanges[0].getValues()[0][
    OPTS.ITEM_COLUMNS.SUPPLIER.index - 1
  ];
  newSheet.getRange("J42").setValue(vendor);

  let allHaveSameVendor = true;
  let allNew = true;
  let index = 50;
  selectedRanges.forEach((range) => {
    range.getValues().forEach((row) => {
      if (row[OPTS.ITEM_COLUMNS.SUPPLIER.index - 1] !== vendor) {
        allHaveSameVendor = false;
      }
      if (row[OPTS.ITEM_COLUMNS.STATUS.index - 1] !== STATUSES_DATA.NEW.text) {
        allNew = false;
      }
      const itemName = row[OPTS.ITEM_COLUMNS.NAME.index - 1];
      newSheet.getRange(index, 2).setValue(itemName);
      const link = row[OPTS.ITEM_COLUMNS.LINK.index - 1];
      newSheet.getRange(index, 8).setValue(link);
      const qty = row[OPTS.ITEM_COLUMNS.QUANTITY.index - 1];
      newSheet.getRange(index, 13).setValue(qty);
      const unitPrice = row[OPTS.ITEM_COLUMNS.UNIT_PRICE.index - 1];
      newSheet.getRange(index, 15).setValue(unitPrice);
      index++;
    });
  });

  if (!allNew) {
    notifications.error("One or more items was not 'New'!");
    currentSpreadsheet.deleteSheet(newSheet);
    return;
  }

  if (!allHaveSameVendor) {
    notifications.error("The items selected do not all have the same vendor!");
    currentSpreadsheet.deleteSheet(newSheet);
    return;
  }

  const sheetName = `${DateTime.local().toFormat(
    "yy-MM-dd"
  )} - ${projectName} - ${vendor}`;
  newSheet.setName(sheetName);

  const newSpreadsheet = SpreadsheetApp.create(sheetName);
  const file = DriveApp.getFileById(newSpreadsheet.getId());
  const parents = file.getParents();
  while (parents.hasNext()) {
    parents.next().removeFile(file);
  }
  targetFolder.addFile(file);

  file.setName(sheetName);
  newSheet.copyTo(newSpreadsheet);
  const sheet1 = newSpreadsheet.getSheetByName("Sheet1");
  if (sheet1) {
    newSpreadsheet.deleteSheet(sheet1);
  }
  currentSpreadsheet.deleteSheet(newSheet);
  openFile(file, targetFolder, vendor);
}
