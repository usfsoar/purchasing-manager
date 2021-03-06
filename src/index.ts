import * as actions from "./actions";
import OPTS from "./config";
import { protectRanges } from "./protect_ranges";
import { sendSelectedToSheet } from "./purchasing";
import { buildProjectStatusSlackMessage } from "./slack_utils";
import { getSheetNameFromProjectName } from "./spreadsheet_utils";
import { buildAndAddCustomMenu } from "./user_interface";

/**
 * Called automatically when Google receives a Post request (this would be from
 * Slack).
 */
export function doPost(
  e: GoogleAppsScript.Events.DoPost & {
    parameter: {
      command?: string;
      text?: string;
      payload?: string;
    };
  }
): GoogleAppsScript.Content.TextOutput {
  let message: Record<string, unknown> = {
    response_type: "ephemeral",
    replace_original: false,
    text: "Error: command not found.",
  };

  if (e.parameter.command == "/budgetstatus") {
    // If the budgetStatus command, send the budgetStatus message
    let text = e.parameter.text;
    if (!text)
      return ContentService.createTextOutput("Error: Invalid project name.");
    const sheetName = getSheetNameFromProjectName(text, true);
    if (sheetName !== null) text = sheetName;
    message = buildProjectStatusSlackMessage(text);
  } else if (e.parameter.payload) {
    // Else maybe it's an interactive message command. Parse the payload and
    // check.
    const payload = JSON.parse(e.parameter.payload);

    if (payload.type === "interactive_message" && payload.actions) {
      if (payload.actions[0].name === OPTS.SLACK.ITEM_LIST_ACTION_NAME_LEGACY) {
        const parsedText = payload.actions[0].value;

        message = {
          response_type: "ephemeral",
          replace_original: false,
          text: parsedText,
        };
      } else {
        message = JSON.parse(payload.actions[0].value);
      }
    }
  }

  return ContentService.createTextOutput(JSON.stringify(message)).setMimeType(
    ContentService.MimeType.JSON
  );
}

/**
 * Global that represents whether the user is authorized as a financial officer.
 */
export function onOpen(): void {
  buildAndAddCustomMenu();
}

global.onOpen = onOpen;
global.doPost = doPost;

// Note: Don't change this into a loop or some other dynamic assignment. It must
// be an explicit static assignment for each one or Webpack won't see it.
global.markSelectedReceivedReimburse = actions.markSelectedReceivedReimburse;
global.markSelectedReimbursed = actions.markSelectedReimbursed;
global.markSelectedSubmitted = actions.markSelectedSubmitted;
global.markSelectedApproved = actions.markSelectedApproved;
global.markSelectedAwaitingPickup = actions.markSelectedAwaitingPickup;
global.markSelectedAwaitingInfo = actions.markSelectedAwaitingInfo;
global.markSelectedDenied = actions.markSelectedDenied;
global.markSelectedNew = actions.markSelectedNew;
global.markAllNew = actions.markAllNew;
global.markSelectedReceived = actions.markSelectedReceived;
global.fastForwardSelectedNew = actions.fastForwardSelectedNew;
global.fastForwardSelectedReceived = actions.fastForwardSelectedReceived;
global.fastForwardSelectedReceivedReimburse =
  actions.fastForwardSelectedReceivedReimburse;
global.fastForwardSelectedReimbursed = actions.fastForwardSelectedReimbursed;
global.fastForwardSelectedSubmitted = actions.fastForwardSelectedSubmitted;
global.fastForwardSelectedApproved = actions.fastForwardSelectedApproved;
global.fastForwardSelectedAwaitingPickup =
  actions.fastForwardSelectedAwaitingPickup;
global.fastForwardSelectedAwaitingInfo =
  actions.fastForwardSelectedAwaitingInfo;
global.fastForwardSelectedDenied = actions.fastForwardSelectedDenied;
global.protectRanges = protectRanges;
global.sendSelectedToSheet = sendSelectedToSheet;
