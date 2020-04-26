/**
 * Replace all occurrences of each key in replacements with the corresponding
 * value, in order. Subsequent replacements can modify previous ones.
 * @param original The string to search in.
 * @param replacements The map of substitutes to use.
 */
export function replaceAll(
  original: string,
  replacements: Map<string, string>
): string {
  return Array.from(replacements.entries()).reduce<string>(
    (result, [replacement, search]) =>
      result.replace(new RegExp(search, "g"), replacement),
    original
  );
}

/**
 * Turn an array into a human-readable list.
 * @param arr Array to make a list from.
 * @param conjunction Conjunction to use at the end of the list.
 * @param noOxfordComma If true, won't add an Oxford comma.
 * @return A nicely formatted list, ie: 'One, Two, and Three'.
 */
export function makeListFromArray(
  arr: string[],
  conjunction = "and",
  noOxfordComma = false
): string {
  const oxfordComma = noOxfordComma || arr.length <= 2 ? "" : ",";

  return arr.reduce((result, element, index) => {
    switch (index) {
      case 0:
        return element;
      case arr.length - 1:
        return `${result}${oxfordComma} ${conjunction}${element}`;
      default:
        return `${result}, ${element}`;
    }
  }, "");
}

/**
 * Push `potentialNewItem` to `arr` if it's not already in `arr`. Returns
 * modified `arr`.
 * @param arr The array to push to.
 * @param potentialNewItem The new item to check for.
 */
export function pushIfNewAndTruthy<T>(arr: T[], potentialNewItem: T): T[] {
  if (arr.indexOf(potentialNewItem) === -1 && potentialNewItem) {
    arr.push(potentialNewItem);
  }
  return arr;
}

/**
 * Wrap a string with double quotes.
 * @param {string} stringToWrap The string to be wrapped in quotes.
 * @return {string} `stringToWrap`, but with quotes around it. If it's an
 * empty string, returns a wrapped space character.
 */
export function wrapInDoubleQuotes(stringToWrap: string): string {
  return `"${stringToWrap || " "}"`;
}

/**
 * Truncates the string if it's longer than `chars` and adds "..." to the end.
 * @param longString The string to shorten.
 * @param chars The maximum number of characters in the final string.
 * @param pad If true, will add padding to end of string to make it
 * the target length.
 * @param padOnly If true, won't truncate ever, will just pad.
 * @return The truncated string.
 */
export function truncateString(
  longString: string,
  chars: number,
  pad = false,
  padOnly = false
): string {
  longString = longString.toString();

  if (longString.length > chars && !padOnly) {
    longString = longString.slice(0, chars - 4) + "...";
  }
  if (pad) {
    const padding =
      chars - longString.length > 0 ? chars - longString.length : 0;
    for (let i = 0; i < padding; i++) longString += " ";
  }
  return longString;
}

/**
 * Escape single quotes in the string for a URL.
 */
export function escapeSingleQuotes(unescaped: string): string {
  return unescaped.replace(/'/g, "%27");
}

/**
 * Escape spaces (replacing `%20` with `+`) in a URI encoded string.
 */
export function escapeSpaces(unescaped: string): string {
  return unescaped.replace(/%20/g, "+");
}
