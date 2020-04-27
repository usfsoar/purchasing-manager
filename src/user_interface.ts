import * as actions from "./actions";
import { verifyAdmin, verifyFinancialOfficer } from "./authorization";
import { protectRanges } from "./protect_ranges";
import { sendSelectedToSheet } from "./purchasing";
import STATUSES_DATA from "./statuses_config";

/**
 * Add the item to the menu if the text is defined.
 * @param menu The menu to add the item to.
 * @param text The text to display on the menu item.
 * @param action Reference to the function to call when the item is clicked.
 * IMPORTANT! For this to work, the function must be statically added to the
 * global object, like with the simple triggers in `index.ts`.
 */
function addMenuItem(
  menu: GoogleAppsScript.Base.Menu,
  text: string | undefined,
  action: () => void
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

  const customMenu = SpreadsheetApp.getUi().createMenu("🚀 SOAR Purchasing");

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
      "➡️ Send to new purchasing sheet",
      sendSelectedToSheet
    );

    fastFowardMenu = SpreadsheetApp.getUi().createMenu("Fast-Forward");
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
