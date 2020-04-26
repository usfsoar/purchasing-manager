import { verifyAdmin } from "./authorization";
import OPTS from "./config";
import * as notifications from "./notifications";
import { getNamedRangeValues } from "./spreadsheet_utils";

/** Reinstate / update all the protected ranges. */
export function protectRanges(): void {
  if (!verifyAdmin()) return;
  const sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  const financialOfficers = getNamedRangeValues(
    OPTS.NAMED_RANGES.APPROVED_OFFICERS
  );
  const projectSheetNames = getNamedRangeValues(
    OPTS.NAMED_RANGES.PROJECT_SHEETS
  );
  const userDataSheetName = OPTS.SHEET_NAMES.USERS;
  SpreadsheetApp.getActiveSpreadsheet()
    .getProtections(SpreadsheetApp.ProtectionType.SHEET)
    .forEach((protection) => protection.remove());
  sheets.forEach((sheet) => {
    const sheetName = sheet.getName();
    if (
      projectSheetNames.indexOf(sheetName) === -1 &&
      sheetName !== userDataSheetName
    ) {
      // Lock the entire sheet if not the user data sheet or a project sheet
      const sheetProtection = sheet.protect();
      sheetProtection.removeEditors(sheetProtection.getEditors());
      sheetProtection.addEditors(financialOfficers);
      notifications.successNotification("Updated protections for " + sheetName);
    }
  });
}
