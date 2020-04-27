import OPTS from "./config";

/** Show the user an error message. */
export function error(message: string): void {
  SpreadsheetApp.getActiveSpreadsheet().toast(
    message,
    OPTS.UI.TOAST_TITLES.ERROR,
    OPTS.UI.TOAST_DURATION
  );
}

/** Show the user a warning message. */
export function warn(message: string): void {
  SpreadsheetApp.getActiveSpreadsheet().toast(
    message,
    OPTS.UI.TOAST_TITLES.WARNING,
    OPTS.UI.TOAST_DURATION
  );
}

/** Show a log message and log it. For debugging. */
export function log(message: string): void {
  Logger.log(message);
  SpreadsheetApp.getActiveSpreadsheet().toast(
    message,
    OPTS.UI.TOAST_TITLES.INFO,
    OPTS.UI.TOAST_DURATION
  );
}

/** Show the user a success message. */
export function success(message: string): void {
  SpreadsheetApp.getActiveSpreadsheet().toast(
    message,
    OPTS.UI.TOAST_TITLES.SUCCESS,
    OPTS.UI.TOAST_DURATION
  );
}
