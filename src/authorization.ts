import OPTS from "./config";
import { User } from "./interfaces";
import SECRET_OPTS from "./secret_config";
import { getNamedRangeValues } from "./spreadsheet_utils";

/**
 * Attempt to get the current user's email. Tries both the active and effective
 * users, and if neither works (because Apps Script doesn't always give
 * permission to access users' emails), returns `null`.
 */
export function getActiveUserEmail(): string | null {
  return (
    Session.getActiveUser().getEmail() ||
    Session.getEffectiveUser().getEmail() ||
    null
  );
}

/**
 * Verify whether or not the email provided is one of an approved financial
 * officer.
 * After first run, uses cache to avoid having to pull the range again.
 * @param email Email of the user to check. If not provided, uses current user's
 * email (sometimes Google does not allow accessing this - in that case, returns
 * `false` always).
 * @return `true` if the user is a financial officer.
 */
export function verifyFinancialOfficer(
  email: string | null = getActiveUserEmail()
): boolean {
  return (
    email !== null &&
    getNamedRangeValues(OPTS.NAMED_RANGES.APPROVED_OFFICERS).includes(email)
  );
}

/**
 * Verify whether or not the current user is the admin.
 * @return `true` if the current user is the admin.
 */
export function verifyAdmin(): boolean {
  return getActiveUserEmail() === SECRET_OPTS.ADMIN_EMAIL;
}

type UserCache = Partial<User> & Pick<User, "email" | "isFinancialOfficer">;

/**
 * Given the cached user object, checks if it's complete.
 * @param cache The cache object.
 */
function isUserCacheComplete(cache: UserCache): cache is User {
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
export const getCurrentUserInfo = ((): (() => User) => {
  const currentEmail = getActiveUserEmail();
  const cache: UserCache = {
    email: currentEmail ?? "",
    isFinancialOfficer: verifyFinancialOfficer(currentEmail),
  };

  return (): User => {
    if (isUserCacheComplete(cache)) return cache;

    const userSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(
      OPTS.SHEET_NAMES.USERS
    );

    if (userSheet === null) throw new Error("No user data sheet found.");

    const userData = userSheet
      .getDataRange()
      .getValues()
      .find((user) => user[0] === cache.email);

    if (userData !== undefined) {
      cache.slackId = userData[1];
      cache.fullName = userData[2];
      cache.phone = userData[3] || undefined; // use || to work with empty string
    } else {
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
    else throw new Error("Failed to get user data.");
  };
})();
