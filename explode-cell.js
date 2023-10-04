/**
 * Constants and Variables
 * INDEX_TO_REPLACE: The index of the column to replace, determined by calling columnNameToIndex.
 * COLUMN_START and COLUMN_END: The start and end columns for the range to be processed.
 * MULTI: A flag to indicate whether to disassemble multiple rows or not.
 * INDEXES_TO_REPLACE: An empty array mapped with columnNameToIndex, but it's not actually doing anything useful as it stands.
 * 
 * 
 * ** Functions
 * columnNameToIndex(columnName, zeroBased = true): Converts a column name (like 'A', 'B', 'AA', etc.) to a zero-based index number.
 * 
 * duplicateRowReplacingField(sheet, range, value): Duplicates a row in a Google Sheet, replacing the value at the index specified by INDEX_TO_REPLACE with a new value.
 * 
 * main(): The main function that iterates through all rows in the active sheet and calls disassembleRow on each row.
 * 
 * disassembleRow(sheet, range, index = null): Takes a row and a column index. It splits the value in the cell at the given index by commas and duplicates the row for each split value, replacing the original cell with each new value.
 */

//example values, use your own
const INDEX_TO_REPLACE   = columnNameToIndex('D');
const COLUMN_START       = 'A';
const COLUMN_END         = 'E';
const MULTI              = false;
const INDEXES_TO_REPLACE = [].map(columnNameToIndex);

function columnNameToIndex(columnName, zeroBased = true ) {
    let index = 0;
    for (let i = 0; i < columnName.length; i++) {
      const char = columnName[i];
      index = index * 26 + (char.charCodeAt(0) - 'A'.charCodeAt(0) + 1);
    } 

    if( zeroBased ) index -= 1;
    
    return index;
}

/**
 * @param {SpreadsheetApp.Sheet} sheet
 * @param {SpreadsheetApp.Range} range
 */
function duplicateRowReplacingField(sheet, range, value) {
    let tab = Array.from(range);
    tab[INDEX_TO_REPLACE] = value;
    sheet.appendRow(tab);
}

function main() {
    let sheet = SpreadsheetApp.getActiveSheet();
    let rows = sheet.getMaxRows();

    for (let i = 2; i <= rows; i++) {
        if( MULTI ){
            disassembleRow(sheet, sheet.getRange(`${COLUMN_START}${i}:${COLUMN_END}${i}`))
        } else {
            INDEXES_TO_REPLACE.forEach( idx => disassembleRow(sheet, sheet.getRange(`${COLUMN_START}${i}:${COLUMN_ENDT}${i}`), idx ) )
        }
    }
}

/**
 * @param {SpreadsheetApp.Sheet} sheet
 * @param {SpreadsheetApp.Range} range
 */
function disassembleRow(sheet, range, index = null) {
    let idx     = index ?? INDEX_TO_REPLACE;
    let values  = range.getValues();
    let numbers = values[0][idx].toString().split(",");

    if (!(numbers.length > 1)) return false;

    for (let i = 1; i < numbers.length; i++) {
        duplicateRowReplacingField(sheet, values[0], numbers[i]);
    }

    let cell = range.getCell(1, (idx + 1));
    cell.setValue(numbers[0])
    return true;
}
