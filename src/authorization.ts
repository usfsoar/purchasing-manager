import OPTS from "./config";
import { User } from "./interfaces";
import SECRET_OPTS from "./secret_config";
import { getNamedRangeValues } from "./spreadsheet_utils";

/**
 * Verify whether or not the email provided is one of an approved financial
 * officer.
 * After first run, uses cache to avoid having to pull the range again.
 * @param email Email of the user to check. If not provided, uses current user's
 * email (sometimes Google does not allow accessing this - in that case, returns
 * `false`).
 * @return `true` if the user is a financial officer.
 */
export function verifyFinancialOfficer(
  email: string | null = Session.getActiveUser().getEmail()
): boolean {
  return (
    email !== null &&
    getNamedRangeValues(OPTS.NAMED_RANGES.APPROVED_OFFICERS).includes(email)
  );
}

/**
 * Verify whether or not the current user is the admin.
 * @return `true` if the current user is an admin.
 */
export function verifyAdmin(): boolean {
  return Session.getActiveUser().getEmail() === SECRET_OPTS.ADMIN_EMAIL;
}

/**
 * Given the cached user object, checks if it's complete.
 * @param cache The cache object.
 */
function isUserCacheComplete(cache: Partial<User>): cache is User {
  return (
    cache.email !== undefined &&
    cache.fullName !== undefined &&
    cache.isFinancialOfficer !== undefined &&
    cache.slackId !== undefined
  );
}

/**
 * Returns the current user's information from the storage sheet, or (prompt)s
 * for it, or returns it from the local cache if it's been asked before.
 * @returns Information about the current user.
 */
export const getCurrentUserInfo = (function (): () => User {
  const currentEmail = Session.getActiveUser().getEmail();
  const cache: Partial<User> & Pick<User, "email" | "isFinancialOfficer"> = {
    email: currentEmail,
    isFinancialOfficer: verifyFinancialOfficer(currentEmail),
  };

  return (): User => {
    if (isUserCacheComplete(cache)) return cache;

    const userSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(
      OPTS.SHEET_NAMES.USERS
    );

    if (userSheet === null) throw new Error("No user data sheet found.");

    const userData = userSheet.getDataRange().getValues();

    let userDataFound = false;
    for (let i = 1; i < userData.length; i++) {
      if (userData[i][0] === cache.email) {
        cache.slackId = userData[i][1];
        cache.fullName = userData[i][2];
        cache.phone = userData[i][3] || undefined;
        userDataFound = true;
        break;
      }
    }

    if (!userDataFound) {
      while (!cache.slackId) {
        cache.slackId = SpreadsheetApp.getUi()
          .prompt(OPTS.UI.SLACK_ID_PROMPT)
          .getResponseText();
      }
      while (!cache.fullName) {
        cache.fullName = SpreadsheetApp.getUi()
          .prompt(OPTS.UI.FULL_NAME_PROMPT)
          .getResponseText();
      }
      userSheet.appendRow([cache.email, cache.slackId, cache.fullName]);
    }

    if (isUserCacheComplete(cache)) return cache;
    throw new Error("Failed to get user data.");
  };
})();
