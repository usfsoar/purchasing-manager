import { verifyAdmin } from "./authorization";
import OPTS from "./config";
import * as notifications from "./notifications";
import { getNamedRangeValues } from "./spreadsheet_utils";

/** Reinstate / update all the protected ranges. */
export function protectRanges(): void {
  if (!verifyAdmin()) return;

  const financialOfficers = getNamedRangeValues(
    OPTS.NAMED_RANGES.APPROVED_OFFICERS
  );
  const projectSheetNames = getNamedRangeValues(
    OPTS.NAMED_RANGES.PROJECT_SHEETS
  );

  SpreadsheetApp.getActiveSpreadsheet()
    .getProtections(SpreadsheetApp.ProtectionType.SHEET)
    .forEach((protection) => protection.remove());

  SpreadsheetApp.getActiveSpreadsheet()
    .getSheets()
    .forEach((sheet) => {
      const sheetName = sheet.getName();
      if (
        // We can't lock the user data sheet or nobody new can enter their info
        sheetName !== OPTS.SHEET_NAMES.USERS &&
        projectSheetNames.includes(sheetName)
      ) {
        // Lock the entire sheet if not the user data sheet or a project sheet
        const sheetProtection = sheet.protect();
        sheetProtection
          .removeEditors(sheetProtection.getEditors())
          .addEditors(financialOfficers);
        notifications.success(`Updated protections for ${sheetName}`);
      }
    });
}
