/*
 * yarn run dev {TARGET_BOOK(FILE)} {TARGET_SHEET} {TARGET_NAME} {TARGET_COLUMN_INDEX}
 *
 * *** EXAMPLE *** 
 * yarn run dev text.xlsx sheet1 FOO_USER 2
 * or
 * node sample text.xlsx sheet1 FOO_USER 2
 */

const XLSX = require("xlsx");
const Utils = XLSX.utils;

// arguments
const book = XLSX.readFile(process.argv[2]);
const sheet = book.Sheets[process.argv[3]];
const TARGET_NAME = process.argv[4];
const TARGET_COLUMN_INDEX = process.argv[5];

const range = sheet["!ref"];
const decodeRange = Utils.decode_range(range);

for (let rowIndex = decodeRange.s.r; rowIndex <= decodeRange.e.r; rowIndex++) {
  const address = Utils.encode_cell({ r: rowIndex, c: TARGET_COLUMN_INDEX});
  if(typeof sheet[address] !== 'undefined' &&
     typeof sheet[address].h !== 'undefined' ) {
    if (sheet[address].h.match(TARGET_NAME)) {
      for (let colIndex = decodeRange.s.c; colIndex <= decodeRange.e.c; colIndex++) {
        const address = Utils.encode_cell({ r: rowIndex, c:colIndex });
        const cell = sheet[address];
        if(typeof cell !== 'undefined') {
          console.log(cell.h);
        }
      }
    }
  }
}
