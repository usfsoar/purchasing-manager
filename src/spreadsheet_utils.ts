import { verifyAdmin } from "./authorization";
import OPTS from "./config";
import * as nofications from "./notifications";

/**
 * Get the sheet name that matches the project.
 * @param projectName Name of the project.
 * @param nullIfMissing If present, will return null if the value is not found.
 * @return String, unless `nullIfMissing` is specified.
 */
export function getSheetNameFromProjectName(
  projectName: string,
  nullIfMissing = false
): string | null {
  const projectsData = SpreadsheetApp.getActiveSpreadsheet()
    .getRange(OPTS.NAMED_RANGES.PROJECT_NAMES_TO_SHEETS)
    .getValues();

  for (let i = 0; i < projectsData.length; i++) {
    if (projectsData[i][0] === projectName) return projectsData[i][1];
  }

  if (nullIfMissing) return null;
  return "_Error: Project Not Found_";
}

/**
 * Get the full name of the project that matches the name of the sheet.
 * @param sheetName Name of the project's sheet.
 */
export function getProjectNameFromSheetName(sheetName: string): string | null {
  const projectsData = SpreadsheetApp.getActiveSpreadsheet()
    .getRange(OPTS.NAMED_RANGES.PROJECT_NAMES_TO_SHEETS)
    .getValues();

  for (let i = 0; i < projectsData.length; i++) {
    if (projectsData[i][1] === sheetName) return projectsData[i][0];
  }

  return null;
}

/**
 * Get the list of non-empty values in the named range.
 * @param rangeName
 * @return Unordered array of values, flattened into a 1-dimensional array.
 */
export function getNamedRangeValues(rangeName: string): string[] {
  const valuesGrid = SpreadsheetApp.getActiveSpreadsheet()
    .getRange(rangeName)
    .getValues();

  // Flatten and remove empty values
  const valuesArray = valuesGrid
    .map((row) => row[0])
    .filter((value) => value !== "");

  return valuesArray;
}

/**
 * Checks if the current sheet is in the list of project sheets. If not,
 * shows a message in the UI and returns false.
 * @param sheetName Name of the sheet to check. If empty, uses current
 * sheet.
 * @return True if a project sheet is active.
 */
export function checkIfProjectSheet(sheetName?: string): boolean {
  const currentSheetName =
    sheetName || SpreadsheetApp.getActiveSheet().getName();

  const projectSheetNames = getNamedRangeValues(
    OPTS.NAMED_RANGES.PROJECT_SHEETS
  );

  if (projectSheetNames.indexOf(currentSheetName) === -1) {
    nofications.errorNotification(
      "This action may only be performed in a project sheet"
    );
    return false;
  }

  return true;
}

/**
 * Get the range of an entire column in the active project sheet, minus headers.
 * @param columnNumber The number of the column to get.
 * @return The range of the column.
 */
export function getColumnRange(
  columnNumber: number
): GoogleAppsScript.Spreadsheet.Range {
  const activeSheet = SpreadsheetApp.getActiveSheet();
  const firstNonHeaderRow = OPTS.NUM_HEADER_ROWS + 1;
  const numNonHeaderRows = activeSheet.getLastRow() - OPTS.NUM_HEADER_ROWS;

  return activeSheet.getRange(
    firstNonHeaderRow,
    columnNumber,
    numNonHeaderRows,
    1
  );
}

/**
 * Get the ranges of all the currently selected rows in the active sheet.
 * @return Array of selected ranges, expanded to cover entire width of data in
 * the sheet.
 */
export function getSelectedRows(): GoogleAppsScript.Spreadsheet.Range[] {
  const activeSheet = SpreadsheetApp.getActiveSheet();

  const selections =
    activeSheet.getSelection().getActiveRangeList()?.getRanges() ?? [];
  const lastColumn = activeSheet.getLastColumn();

  // Expand selections to width of spreadsheet
  const expandedSelections = selections.map((selectionRange) => {
    let selectionStartRow = selectionRange.getRow();
    let selectionNumRows = selectionRange.getNumRows();
    if (selectionStartRow === 1) {
      selectionStartRow++;
      selectionNumRows--;
    }
    if (selectionStartRow === 2) {
      selectionStartRow++;
      selectionNumRows--;
    }
    return activeSheet.getRange(
      selectionStartRow,
      1,
      selectionNumRows,
      lastColumn
    );
  });

  return expandedSelections;
}

/**
 * Get the range of all data in the active sheet.
 * @return Array with one `Range`.
 */
export function getAllRows(): GoogleAppsScript.Spreadsheet.Range[] {
  const activeSheet = SpreadsheetApp.getActiveSheet();
  const lastColumnInSheet = activeSheet.getLastColumn();
  const firstNonHeaderRow = OPTS.NUM_HEADER_ROWS + 1;

  const nameColumnValues = getColumnRange(
    OPTS.ITEM_COLUMNS.NAME.index
  ).getValues();

  /** The number of the last row in the sheet that has a value for Name. */
  let lastRowWithData = firstNonHeaderRow;

  nameColumnValues.forEach((name, index) => {
    if (name.toString().trim() !== "") {
      lastRowWithData = index + firstNonHeaderRow;
    }
  });

  const numNonHeaderRowsWithData = lastRowWithData - OPTS.NUM_HEADER_ROWS;

  return [
    activeSheet.getRange(
      firstNonHeaderRow,
      1,
      numNonHeaderRowsWithData,
      lastColumnInSheet
    ),
  ];
}

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
      nofications.successNotification("Updated protections for " + sheetName);
    }
  });
}
