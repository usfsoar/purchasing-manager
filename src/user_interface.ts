import * as actions from "./actions";
import {
  getCurrentUserInfo,
  verifyAdmin,
  verifyFinancialOfficer,
} from "./authorization";
import OPTS from "./config";
import { sendSelectedToSheet } from "./purchasing";
import { protectRanges } from "./spreadsheet_utils";
import STATUSES_DATA from "./statuses_config";
import { escapeSingleQuotes, escapeSpaces } from "./utils";

/**
 * Add the item to the menu if the text is defined.
 * @param menu The menu to add the item to.
 * @param text The text to display on the menu item.
 * @param action Reference to the function to call when the item is clicked.
 */
function addMenuItem(
  menu: GoogleAppsScript.Base.Menu,
  text: string | undefined,
  action: Function
): void {
  if (text !== undefined) {
    menu.addItem(text, action.name);
  }
}

/**
 * Build the custom SOAR Purchasing menu and add it to the user interface.
 */
export function buildAndAddCustomMenu(): void {
  // Use yourFunction.name because it requires a string and this is a little
  // more reusable than just hardcoding the name

  const customMenu = SpreadsheetApp.getUi().createMenu(OPTS.CUSTOM_MENU.NAME);

  addMenuItem(customMenu, STATUSES_DATA.NEW.actionText.all, actions.markAllNew);
  addMenuItem(
    customMenu,
    STATUSES_DATA.NEW.actionText.selected,
    actions.markSelectedNew
  );

  let fastFowardMenu = null;

  if (verifyFinancialOfficer()) {
    customMenu.addSeparator();
    addMenuItem(
      customMenu,
      STATUSES_DATA.SUBMITTED.actionText.selected,
      actions.markSelectedSubmitted
    );
    addMenuItem(
      customMenu,
      STATUSES_DATA.APPROVED.actionText.selected,
      actions.markSelectedApproved
    );
    addMenuItem(
      customMenu,
      STATUSES_DATA.AWAITING_PICKUP.actionText.selected,
      actions.markSelectedAwaitingPickup
    );
    customMenu.addSeparator();
    addMenuItem(
      customMenu,
      STATUSES_DATA.AWAITING_INFO.actionText.selected,
      actions.markSelectedAwaitingInfo
    );
    addMenuItem(
      customMenu,
      STATUSES_DATA.DENIED.actionText.selected,
      actions.markSelectedDenied
    );
    addMenuItem(
      customMenu,
      STATUSES_DATA.REIMBURSED.actionText.selected,
      actions.markSelectedReimbursed
    );
    customMenu.addSeparator();
    addMenuItem(
      customMenu,
      "Send to new purchasing sheet",
      sendSelectedToSheet
    );

    fastFowardMenu = SpreadsheetApp.getUi().createMenu(
      OPTS.FAST_FORWARD_MENU.NAME
    );
    addMenuItem(
      fastFowardMenu,
      STATUSES_DATA.NEW.actionText.fastForward,
      actions.fastForwardSelectedNew
    );
    addMenuItem(
      fastFowardMenu,
      STATUSES_DATA.SUBMITTED.actionText.fastForward,
      actions.fastForwardSelectedSubmitted
    );
    addMenuItem(
      fastFowardMenu,
      STATUSES_DATA.APPROVED.actionText.fastForward,
      actions.fastForwardSelectedApproved
    );
    addMenuItem(
      fastFowardMenu,
      STATUSES_DATA.AWAITING_INFO.actionText.fastForward,
      actions.fastForwardSelectedAwaitingInfo
    );
    addMenuItem(
      fastFowardMenu,
      STATUSES_DATA.DENIED.actionText.fastForward,
      actions.fastForwardSelectedDenied
    );
    addMenuItem(
      fastFowardMenu,
      STATUSES_DATA.AWAITING_PICKUP.actionText.fastForward,
      actions.fastForwardSelectedAwaitingPickup
    );
    addMenuItem(
      fastFowardMenu,
      STATUSES_DATA.RECEIVED.actionText.fastForward,
      actions.fastForwardSelectedReceived
    );
    addMenuItem(
      fastFowardMenu,
      STATUSES_DATA.RECEIVED_REIMBURSE.actionText.fastForward,
      actions.fastForwardSelectedReceivedReimburse
    );
    addMenuItem(
      fastFowardMenu,
      STATUSES_DATA.REIMBURSED.actionText.fastForward,
      actions.fastForwardSelectedReimbursed
    );
  }

  customMenu.addSeparator();
  addMenuItem(
    customMenu,
    STATUSES_DATA.RECEIVED.actionText.selected,
    actions.markSelectedReceived
  );
  customMenu.addSeparator();
  addMenuItem(
    customMenu,
    STATUSES_DATA.RECEIVED_REIMBURSE.actionText.selected,
    actions.markSelectedReceivedReimburse
  );

  if (verifyAdmin()) {
    customMenu.addSeparator();
    addMenuItem(customMenu, "Refresh protections", protectRanges);
  }

  customMenu.addToUi();
  if (verifyFinancialOfficer()) fastFowardMenu?.addToUi();
}

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
  const fileUrl = "https://docs.google.com/spreadsheets/d/" + spreadsheetId;
  const folderUrl = "https://drive.google.com/drive/u/2/folders/" + folderId;
  const currentUserEmail = escapeSingleQuotes(
    encodeURIComponent(getCurrentUserInfo().email)
  );
  const emailTemplateLink =
    "https://mail.google.com/mail/u/0/?view=cm&fs=1&to=sg-rmdpurchase@usf.edu&authuser=" +
    currentUserEmail +
    "&su=SOCIETY+OF+AERONAUTICS+AND+ROCKETRY,+" +
    vendor +
    "&body=Please+see+attached+purchasing+form.&tf=1";
  const html =
    "<div style='font-family: sans-serif;'>Successfully sent items to sheet.<br><a target='_blank' href='" +
    folderUrl +
    "'>Open Purchasing Sheets Folder</a><br><a target='_blank' href='" +
    fileUrl +
    "'>Open The New Purchasing Sheet</a><br /><br /><strong>Remember:</strong> Attach the form to an email sent from your @mail.usf.edu email adress to sg-rmdpurchase@usf.edu with subject \"SOCIETY OF AERONAUTICS AND ROCKETRY, VENDOR NAME, EVENT DATE (if applicable)\". Only one form per email! <br><br> <a  href='" +
    emailTemplateLink +
    "' target='_blank'><strong>Start Email in Gmail</strong></a></div>";
  const userInterface = HtmlService.createHtmlOutput(html);
  SpreadsheetApp.getUi().showModalDialog(userInterface, "Open Sheet");
}
