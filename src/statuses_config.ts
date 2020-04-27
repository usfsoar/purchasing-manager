import OPTS from "./config";
import { Status } from "./interfaces";
import SECRET_OPTS from "./secret_config";

/**
 * Information about each possible item status.
 */
const STATUSES: Record<string, Status> = {
  CREATED: {
    text: "",
    allowedPrevious: [],
    actionText: {},
    slack: {},
    columns: {
      user: null,
      date: null,
    },
    officersOnly: false,
  },
  NEW: {
    text: "New",
    allowedPrevious: ["", "Awaiting Info"],
    actionText: {
      fastForward: "New",
      selected: "✔️ Submit selected new items",
      all: "Submit all new items",
    },
    slack: {
      emoji: ":new:",
      targetUsers: OPTS.SLACK.TARGET_USERS.OFFICERS,
      messageTemplates: [
        "{emoji} {userTags} {userFullName} has submitted {numMarked} new item{plural} to be purchased for {projectName}.",
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
      OPTS.ITEM_COLUMNS.CATEGORY,
    ],
    officersOnly: false,
  },
  SUBMITTED: {
    text: "Submitted",
    allowedPrevious: ["New"],
    actionText: {
      fastForward: "Submitted",
      selected: "🛒 Mark selected items as submitted",
    },
    slack: {
      emoji: ":usf:",
      targetUsers: OPTS.SLACK.TARGET_USERS.REQUESTORS,
      messageTemplates: [
        "{emoji} {userTags} {userFullName} marked {numMarked} item{plural} for {projectName} as *submitted* to Student Government.",
      ],
      channelWebhooks: [SECRET_OPTS.SLACK.WEBHOOKS.PURCHASING],
    },
    columns: {
      user: OPTS.ITEM_COLUMNS.OFFICER_EMAIL,
      date: OPTS.ITEM_COLUMNS.SUBMIT_DATE,
    },
    fastForwardColumns: {
      user: [OPTS.ITEM_COLUMNS.REQUEST_EMAIL],
      date: [OPTS.ITEM_COLUMNS.REQUEST_DATE],
    },
    reccomendedColumns: [OPTS.ITEM_COLUMNS.ACCOUNT, OPTS.ITEM_COLUMNS.CATEGORY],
    fillInDefaults: true,
    officersOnly: true,
  },
  APPROVED: {
    text: "Ordered",
    allowedPrevious: ["Submitted", "New"],
    actionText: {
      fastForward: "Ordered",
      selected: "Mark selected items as ordered",
    },
    slack: {
      emoji: ":white_check_mark:",
      targetUsers: OPTS.SLACK.TARGET_USERS.REQUESTORS,
      messageTemplates: [
        "{emoji} {userTags} {userFullName} marked {numMarked} item{plural} for {projectName} as *ordered*.",
      ],
      channelWebhooks: [SECRET_OPTS.SLACK.WEBHOOKS.PURCHASING],
    },
    columns: {
      user: null,
      date: OPTS.ITEM_COLUMNS.UPDATE_DATE,
    },
    fastForwardColumns: {
      user: [OPTS.ITEM_COLUMNS.REQUEST_EMAIL, OPTS.ITEM_COLUMNS.OFFICER_EMAIL],
      date: [OPTS.ITEM_COLUMNS.REQUEST_DATE, OPTS.ITEM_COLUMNS.SUBMIT_DATE],
    },
    fillInDefaults: true,
    officersOnly: true,
  },
  AWAITING_PICKUP: {
    text: "Awaiting Pickup",
    allowedPrevious: ["Submitted", "Ordered"],
    actionText: {
      fastForward: "Awaiting Pickup",
      selected: "Mark selected items as awaiting pickup",
    },
    slack: {
      emoji: ":package:",
      targetUsers: OPTS.SLACK.TARGET_USERS.CHANNEL,
      messageTemplates: [
        "{emoji} {userFullName} marked {numMarked} item{plural} for {projectName} as awaiting pickup (usually in MSC 4300). _React with " +
          OPTS.SLACK.CHECK_MARK_EMOJI +
          " if you're going to pick them up._",
      ],
      channelWebhooks: [SECRET_OPTS.SLACK.WEBHOOKS.PURCHASING],
    },
    columns: {
      user: null,
      date: OPTS.ITEM_COLUMNS.ARRIVE_DATE,
    },
    fastForwardColumns: {
      user: [OPTS.ITEM_COLUMNS.REQUEST_EMAIL, OPTS.ITEM_COLUMNS.OFFICER_EMAIL],
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
    text: "Received",
    allowedPrevious: ["Awaiting Pickup", "Submitted", "Ordered"],
    actionText: {
      fastForward: "Received",
      selected: "📦 Mark selected items as received (picked up)",
    },
    slack: {
      emoji: ":heavy_check_mark:",
      targetUsers: OPTS.SLACK.TARGET_USERS.REQUESTORS,
      messageTemplates: [
        "{emoji} {userTags} {userFullName} marked {numMarked} item{plural} for {projectName} as received (picked up).",
      ],
      channelWebhooks: [SECRET_OPTS.SLACK.WEBHOOKS.PURCHASING],
    },
    columns: {
      user: OPTS.ITEM_COLUMNS.RECEIVE_EMAIL,
      date: OPTS.ITEM_COLUMNS.RECEIVE_DATE,
    },
    fastForwardColumns: {
      user: [OPTS.ITEM_COLUMNS.REQUEST_EMAIL, OPTS.ITEM_COLUMNS.OFFICER_EMAIL],
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
    text: "Denied",
    allowedPrevious: ["New", "Submitted", "Ordered", "Awaiting Info"],
    actionText: {
      fastForward: "Denied",
      selected: "Deny selected items",
    },
    slack: {
      emoji: ":x:",
      targetUsers: OPTS.SLACK.TARGET_USERS.REQUESTORS,
      messageTemplates: [
        "{emoji} {userTags} {userFullName} *denied* {numMarked} item{plural} for {projectName} (_see comments in database_).",
      ],
      channelWebhooks: [SECRET_OPTS.SLACK.WEBHOOKS.PURCHASING],
    },
    columns: {
      user: OPTS.ITEM_COLUMNS.OFFICER_EMAIL,
      date: OPTS.ITEM_COLUMNS.UPDATE_DATE,
    },
    fastForwardColumns: {
      user: [OPTS.ITEM_COLUMNS.REQUEST_EMAIL],
      date: [OPTS.ITEM_COLUMNS.REQUEST_DATE],
    },
    requiredColumns: [OPTS.ITEM_COLUMNS.OFFICER_COMMENTS],
    officersOnly: true,
  },
  AWAITING_INFO: {
    text: "Awaiting Info",
    allowedPrevious: ["New", "Submitted", "Denied", "Ordered", "Received"],
    actionText: {
      fastForward: "Awaiting Info",
      selected: "Request more information for selected items",
    },
    slack: {
      emoji: ":exclamation:",
      targetUsers: OPTS.SLACK.TARGET_USERS.REQUESTORS,
      messageTemplates: [
        "{emoji} {userTags} {userFullName} requested more info for {numMarked} item{plural} for {projectName} (_see comments in database_). Update the information, then resubmit as new items.",
      ],
      channelWebhooks: [SECRET_OPTS.SLACK.WEBHOOKS.PURCHASING],
    },
    columns: {
      user: OPTS.ITEM_COLUMNS.OFFICER_EMAIL,
      date: OPTS.ITEM_COLUMNS.UPDATE_DATE,
    },
    fastForwardColumns: {
      user: [OPTS.ITEM_COLUMNS.REQUEST_EMAIL],
      date: [OPTS.ITEM_COLUMNS.REQUEST_DATE],
    },
    requiredColumns: [OPTS.ITEM_COLUMNS.OFFICER_COMMENTS],
    officersOnly: true,
  },
  RECEIVED_REIMBURSE: {
    text: "Received - Awaiting Reimbursement",
    allowedPrevious: [
      "",
      "New",
      "Submitted",
      "Ordered",
      "Received",
      "Awaiting Pickup",
      "Awaiting Info",
    ],
    actionText: {
      fastForward: "Received - Awaiting Reimbursement",
      selected: "Mark selected items received and request reimbursement",
    },
    slack: {
      emoji: ":heavy_dollar_sign:",
      targetUsers: OPTS.SLACK.TARGET_USERS.OFFICERS,
      messageTemplates: [
        "{emoji} {userTags} {userFullName} marked {numMarked} item{plural} as received for {projectName} and requested reimbursement for them.",
      ],
      channelWebhooks: [SECRET_OPTS.SLACK.WEBHOOKS.PURCHASING],
    },
    columns: {
      date: OPTS.ITEM_COLUMNS.RECEIVE_DATE,
      user: null,
    },
    requiredColumns: [OPTS.ITEM_COLUMNS.REQUEST_COMMENTS],
    officersOnly: false,
    fastForwardColumns: {
      user: [OPTS.ITEM_COLUMNS.REQUEST_EMAIL, OPTS.ITEM_COLUMNS.OFFICER_EMAIL],
      date: [
        OPTS.ITEM_COLUMNS.REQUEST_DATE,
        OPTS.ITEM_COLUMNS.SUBMIT_DATE,
        OPTS.ITEM_COLUMNS.UPDATE_DATE,
        OPTS.ITEM_COLUMNS.ARRIVE_DATE,
      ],
    },
  },
  REIMBURSED: {
    text: "Reimbursed",
    allowedPrevious: ["Received - Awaiting Reimbursement", "Received"],
    actionText: {
      fastForward: "Reimbursed",
      selected: "Mark selected items as reimbursed",
    },
    slack: {
      emoji: ":money_with_wings:",
      targetUsers: OPTS.SLACK.TARGET_USERS.REQUESTORS,
      messageTemplates: [
        "{emoji} {userTags} {userFullName} t reimbursement for {numMarked} item{plural}.",
      ],
      channelWebhooks: [SECRET_OPTS.SLACK.WEBHOOKS.PURCHASING],
    },
    columns: {
      date: OPTS.ITEM_COLUMNS.UPDATE_DATE,
      user: null,
    },
    requiredColumns: [],
    officersOnly: true,
    fastForwardColumns: {
      user: [OPTS.ITEM_COLUMNS.REQUEST_EMAIL, OPTS.ITEM_COLUMNS.OFFICER_EMAIL],
      date: [
        OPTS.ITEM_COLUMNS.REQUEST_DATE,
        OPTS.ITEM_COLUMNS.SUBMIT_DATE,
        OPTS.ITEM_COLUMNS.UPDATE_DATE,
        OPTS.ITEM_COLUMNS.ARRIVE_DATE,
      ],
    },
  },
};

export default STATUSES;

export const TEST_STATUS = {
  text: "Test",
  allowedPrevious: ["", "Test"],
  actionText: {
    fastForward: "Test",
    selected: "Test update item",
  },
  slack: {
    emoji: ":checkered_flag:",
    targetUsers: OPTS.SLACK.TARGET_USERS.CHANNEL,
    messageTemplates: [
      "{emoji} {userTags} {userFullName} marked {numMarked} item{plural} for {projectName} as *test* by TEsting.",
    ],
    channelWebhooks: [SECRET_OPTS.SLACK.WEBHOOKS.DEV],
  },
  columns: {
    user: null,
    date: OPTS.ITEM_COLUMNS.UPDATE_DATE,
  },
  fastForwardColumns: {
    user: [OPTS.ITEM_COLUMNS.REQUEST_EMAIL, OPTS.ITEM_COLUMNS.OFFICER_EMAIL],
    date: [OPTS.ITEM_COLUMNS.REQUEST_DATE, OPTS.ITEM_COLUMNS.SUBMIT_DATE],
  },
  fillInDefaults: true,
  officersOnly: true,
};

/**
 * Check to see if the row is allowed to tbe changes to `newStatus`.
 * @param {string} currentStatusText The current status of the row.
 * @param {Status} newStatus The status object to check for changing to.
 * @return {boolean} True if the current status of the row allows it to be
 * changed to the newStatus.
 */
export function isNewStatusAllowed(
  currentStatusText: string,
  newStatus: Status
): boolean {
  return newStatus.allowedPrevious.includes(currentStatusText.trim());
}
