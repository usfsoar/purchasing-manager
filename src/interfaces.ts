import OPTS from "./config";

/** A data object describing a named column. */
export interface Column {
  /** 1-based index of the column in the sheet. */
  index: number;
  /** Name of the column. */
  name?: string;
}

/** A data object describing a possible item status. */
export interface Status {
  /** The textual name of the status. */
  text: string;
  /** Allowed previous statuses identified by their `text` properties. */
  allowedPrevious: string[];
  /** Menu item text. */
  actionText: {
    /** Menu item text for marking just selected. */
    selected?: string;
    /** Menu item text for marking all possible. */
    all?: string;
    /** Menu item text for fast-forwarding items. */
    fastForward?: string;
  };
  /** Data for sending Slack notifications. */
  slack: {
    /**
     * Templates for sending Slack messages.
     * Will send a Slack message per string. Will replace {emoji} with the emoji,
     * {userTags} with the target user tags, {userFullName} with full name of
     * submitter, {numMarked} with the number of items marked, {projectName} with
     * the name of the project, and {projectSheetUrl} with the link to the project
     * sheet.
     */
    messageTemplates?: string[];
    /**
     * Webhooks to send Slack messages to.
     * Will only tag targetUsers in the first channel provided, to avoid
     * annoying.
     */
    channelWebhooks?: string[];
    /** Emoji to send with slack message. */
    emoji?: string;
    /**
     * String representing a user group to tag in Slack messages (only in the
     * first channel the message is sent to).
     */
    targetUsers?: keyof typeof OPTS["SLACK"]["TARGET_USERS"];
  };
  /** Columns to input data into. */
  columns: {
    /** Column to input attribution email address into. */
    user: Column | null;
    /** Column to input action date into. */
    date: Column | null;
  };
  /** Columns to auto-fill upon fast-forwarding. */
  fastForwardColumns?: {
    /** Column to input attribution email address into. */
    user?: Column[];
    /** Column to input action date into. */
    date?: Column[];
  };
  /** Optional required columns needed to perform actions. */
  requiredColumns?: Column[];
  /** Optional reccomended columns desired to perform actions. */
  reccomendedColumns?: Column[];
  /**
   * If true, will fill default values for Account and Cateegory when those are
   * applied.
   */
  fillInDefaults?: boolean;
  /* Set to `true` to make the status accessible only to Financial Officers. **/
  officersOnly: boolean;
}

export interface User {
  slackId: string;
  fullName: string;
  isFinancialOfficer: boolean;
  /** Email will be empty string if the app doesn't have access. */
  email: string;
  phone?: string;
}

/** Metadata about all items affected by an action. */
export interface Item {
  name: string;
  quantity: string;
  totalPrice: number;
  unitPrice: number;
  category: string;
  requestorComments: string;
  officerComments: string;
  supplier: string;
  productNum: string;
  link: string;
}
