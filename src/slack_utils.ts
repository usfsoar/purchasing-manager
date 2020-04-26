import OPTS from "./config";
import { Item, Status } from "./interfaces";
import {
  checkIfProjectSheet,
  getNamedRangeValues,
  getProjectNameFromSheetName,
  getSheetNameFromProjectName,
} from "./spreadsheet_utils";
import { makeListFromArray, replaceAll, truncateString } from "./utils";

/**
 * Returns the Slack ID that matches the email address provided.
 * @param email Email address of the person to look for.
 * @return The Slack ID or null if no match.
 */
function getSlackIdByEmail(email: string): string | null {
  const userSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(
    OPTS.SHEET_NAMES.USERS
  );
  if (!userSheet) throw new Error("User data sheet not found.");

  const userData = userSheet.getDataRange().getValues();

  for (let i = 1; i < userData.length; i++) {
    if (userData[i][0] === email) {
      return userData[i][1];
    }
  }

  return null;
}

/**
 * Wrapper for `getSlackIdByEmail` that adds tagging formatting.
 * @param email Email address of the person to look for.
 * @return The Slack ID or '' if no match.
 */
function getSlackTagByEmail(email: string): string {
  const slackId = getSlackIdByEmail(email);
  return slackId ? "<@" + getSlackIdByEmail(email) + ">" : "";
}

/**
 * Build a project status message to be used in Slack, with current information
 * about the given project.
 * @param project The name of the project.
 * @return A valid Slack message with attachments.
 */
export function buildProjectStatusSlackMessage(
  project: string
): Record<string, unknown> {
  const projectSheetName =
    getSheetNameFromProjectName(project, true) || project;

  if (!checkIfProjectSheet(projectSheetName)) {
    return {
      response_type: "ephemeral",
      text: "Sorry, I don't recognize that project.",
    };
  }

  const projectName = getProjectNameFromSheetName(project);
  const dashboardSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(
    `${projectSheetName} Dashboard`
  );

  if (!dashboardSheet) {
    return {
      response_type: "ephemeral",
      text: "I couldn't find the dashboard sheet for that project.",
    };
  }

  const totalBudget = dashboardSheet
    .getRange(
      OPTS.DASHBOARD_CELLS.TOTAL_BUDGET.row,
      OPTS.DASHBOARD_CELLS.TOTAL_BUDGET.column
    )
    .getValue();
  const totalExpenses = dashboardSheet
    .getRange(
      OPTS.DASHBOARD_CELLS.TOTAL_EXPENSES.row,
      OPTS.DASHBOARD_CELLS.TOTAL_EXPENSES.column
    )
    .getValue();

  const budgetRemaining = (totalBudget - totalExpenses).toFixed(2);
  const percentBudgetRemaining = (
    ((totalBudget - totalExpenses) / totalBudget) *
    100
  ).toFixed(0);

  const dashboardSheetUrl =
    SpreadsheetApp.getActiveSpreadsheet().getUrl() +
    "#gid=" +
    dashboardSheet.getSheetId();

  const projectSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(
    projectSheetName
  );

  const projectSheetUrl =
    projectSheet &&
    SpreadsheetApp.getActiveSpreadsheet().getUrl() +
      "#gid=" +
      projectSheet.getSheetId();

  const actions = [
    {
      type: "button",
      text: "Open Dashboard ↗",
      url: dashboardSheetUrl,
    },
  ];

  if (projectSheetUrl !== null) {
    actions.push({
      type: "button",
      text: "Open Purchasing Sheet ↗",
      url: projectSheetUrl,
    });
  }

  return {
    response_type: "in_channel",
    attachments: [
      {
        fallback: `The ${projectName} project has $${budgetRemaining} (or ${percentBudgetRemaining}% remaining, out of a total annual budget of ${totalBudget.toFixed(
          2
        )}) . For more details, see the <${dashboardSheetUrl}|project dashboard>.`,
        color: dashboardSheet.getTabColor(),
        title: `${projectName} Budget Status`,
        text:
          "This is the latest budget information from the SOAR Purchasing Database:",
        fields: [
          {
            title: "Total Budget",
            value: `$${totalBudget.toFixed(2)}`,
            short: true,
          },
          {
            title: "Percent Remaining",
            value: `${percentBudgetRemaining}%`,
            short: true,
          },
          {
            title: "Total Expenses",
            value: `$${totalExpenses.toFixed(2)}`,
            short: true,
          },
          {
            title: "Amount Remaining",
            value: `$${budgetRemaining}`,
            short: true,
          },
        ],
        footer: "SOAR Purchasing Database",
        footer_icon: OPTS.SLACK.SOAR_ICON,
        ts: new Date().getTime() / 1000,
        actions,
      },
    ],
  };
}

/**
 * Build normal strings from the status' templates.
 * @param statusData Data for the target status.
 * @param userFullName Full Name of the current user.
 * @param requestors Emails of people who requested the items affected by this
 * action, so they can be notified of the change.
 * @param numMarked Number of items affected by this action.
 * @param projectName Name of the relevant project.
 * @param projectSheetUrl Link to the relevant project's sheet in the database.
 * @param dontTagUsers If `true`, won't add user tags.
 * @return Filled in message strings.
 */
export function buildSlackMessages(
  statusData: Status,
  userFullName: string,
  requestors: string[],
  numMarked: number,
  projectName: string,
  projectSheetUrl: string,
  dontTagUsers = false
): string[] {
  if (
    statusData.slack.messageTemplates === undefined ||
    statusData.slack.messageTemplates.length === 0
  ) {
    return [];
  }

  let targetUserTagsString = "";
  if (!dontTagUsers) {
    switch (statusData.slack.targetUsers) {
      case OPTS.SLACK.TARGET_USERS.CHANNEL:
        targetUserTagsString = "<!channel>";
        break;

      case OPTS.SLACK.TARGET_USERS.OFFICERS: {
        /**
         * Array of booleans with indexes that match officers. Only false if NO.
         */
        const officerNotifyOptions = getNamedRangeValues(
          OPTS.NAMED_RANGES.NOTIFY_APPROVED_OFFICERS
        ).map(function (value) {
          return value !== "NO";
        });
        /** Emails of all the officers that do get notified. */
        const officerEmails = getNamedRangeValues(
          OPTS.NAMED_RANGES.APPROVED_OFFICERS
        ).filter(function (_, index) {
          return officerNotifyOptions[index];
        });
        const officerUserTags = officerEmails
          .map(getSlackTagByEmail)
          .filter(function (slackTag) {
            return slackTag != "";
          });
        targetUserTagsString = makeListFromArray(officerUserTags, "or");
        break;
      }
      case OPTS.SLACK.TARGET_USERS.REQUESTORS: {
        const requestorUserTags = requestors.map(getSlackTagByEmail);
        targetUserTagsString = makeListFromArray(requestorUserTags, "");
      }
    }
  }

  return statusData.slack.messageTemplates.map((template) =>
    replaceAll(
      template,
      new Map([
        ["{emoji}", statusData.slack.emoji ?? ""],
        ["{userTags}", !dontTagUsers ? targetUserTagsString + ":" : ""],
        ["{userFullName}", userFullName],
        ["{numMarked}", numMarked.toString()],
        ["{projectName}", projectName],
        ["{projectSheetUrl}", projectSheetUrl],
        ["{plural}", numMarked !== 1 ? "s" : ""],
      ])
    )
  );
}

/**
 * Build a Slack button attachment that sends a request to show the full item
 * list on click.
 */
function buildItemListSlackAttachment(
  items: Item[],
  projectName: string,
  projectSheetUrl: string,
  user: string,
  action: string,
  projectColor: string
): Record<string, unknown> {
  const attachment = {
    type: "button",
    text: "List Items",
    name: OPTS.SLACK.ITEM_LIST_ACTION_NAME,
    /** JSON to parse as return message later */
    value: "",
  };

  const itemListMessage = {
    response_type: "ephemeral",
    replace_original: false,
    text: "Here are all the items that were affected by that action:",
    attachments: [] as Record<string, unknown>[],
    parse: "full",
    mrkdwn: true,
  };

  const itemsByCategory = items.reduce<Record<string, Item[]>>(
    (result, item) => {
      if (!result[item.category]) result[item.category] = [];
      result[item.category].push(item);
      return result;
    },
    {}
  );

  itemListMessage.attachments = Object.getOwnPropertyNames(itemsByCategory).map(
    (category) => {
      const categoryAttachment = {
        author_name: user + " - " + action,
        title: category,
        title_link: projectSheetUrl,
        color: projectColor,
        fields: itemsByCategory[category].map((item) => {
          const totalPrice =
            typeof item.totalPrice === "number"
              ? item.totalPrice.toFixed(2)
              : "UNKNOWN";
          const itemField = {
            title: truncateString(item.name, 45),
            value:
              "$" +
              totalPrice +
              "\n\t (" +
              item.quantity +
              "x @ $" +
              item.unitPrice.toFixed(2) +
              "/e)",
            short: "true",
          };

          if (item.supplier || item.productNum) itemField.value += "\n\t";
          // Links seem to be broken :(
          // if(item.link && (item.supplier || item.productNum)) itemField.value += "<" + item.link + "|";
          if (item.productNum) itemField.value += "`#" + item.productNum + "`";
          if (item.supplier) itemField.value += " from " + item.supplier;
          // if(item.link && (item.supplier || item.productNum)) itemField.value += ">";

          if (item.requestorComments) {
            itemField.value +=
              "\n Requestor Comment: \n> _" + item.requestorComments + "_";
          }
          if (item.officerComments) {
            itemField.value +=
              "\n Officer Comment: \n> _" + item.officerComments + "_";
          }

          return itemField;
        }),
        footer: projectName,
        footer_icon: OPTS.SLACK.SOAR_ICON,
        mrkdwn_in: ["fields"],
      };

      return categoryAttachment;
    }
  );

  attachment.value = JSON.stringify(itemListMessage);

  if (attachment.value.length >= 2000) {
    itemListMessage.text =
      "Sorry, there were too many items to list. Open the project sheet to view them instead. https://github.com/usfsoar/purchasing-manager/issues/6";
    itemListMessage.attachments = [];
    attachment.value = JSON.stringify(itemListMessage);
  }

  return attachment;
}

/**
 * Send a message to the Slack channel.
 * @param messageData The message to send, according to the Slack API.
 * @param webhook The webhook URL to send the message to.
 */
function sendSlackMessage(
  messageData: Record<string, unknown>,
  webhook: string
): void {
  const requestOptions = {
    method: "post",
    payload: JSON.stringify(messageData),
    contentType: "application/json",
  } as const;
  UrlFetchApp.fetch(webhook, requestOptions);
}

/**
 * Build normal strings from the status' templates.
 * @param statusData Data for the target status.
 * @param userFullName Full Name of the current user.
 * @param requestors Emails of people who requested the items affected by this
 * action.
 * @param itemsMarked Data about all items affected by this action.
 * @param projectName Name of the relevant projec.
 * @param projectSheetUrl Link to the relevant project's sheet in the database.
 * @param projectColor Tab color of the project's sheet.
 */
export function slackNotifyItems(
  statusData: Status,
  userFullName: string,
  requestors: string[],
  itemsMarked: Item[],
  projectName: string,
  projectSheetUrl: string,
  projectColor: string
): void {
  statusData.slack.channelWebhooks?.forEach((webhook, index) => {
    let messages: string[] = [];
    Logger.log(itemsMarked);
    if (index === 0) {
      messages = buildSlackMessages(
        statusData,
        userFullName,
        requestors,
        itemsMarked.length,
        projectName,
        projectSheetUrl
      );
    } else {
      messages = buildSlackMessages(
        statusData,
        userFullName,
        requestors,
        itemsMarked.length,
        projectName,
        projectSheetUrl,
        true
      );
    }

    const messagesWithAttachments: Array<{
      text: string;
      attachments?: Record<string, unknown>[];
    }> = messages.map((messageText) => ({ text: messageText }));
    messagesWithAttachments[messagesWithAttachments.length - 1].attachments = [
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
            projectColor
          ),
          {
            type: "button",
            text: "Open Sheet ↗",
            url: projectSheetUrl,
          },
        ],
        color: projectColor,
      },
    ];

    messagesWithAttachments.forEach((msg) => sendSlackMessage(msg, webhook));
  });
}
