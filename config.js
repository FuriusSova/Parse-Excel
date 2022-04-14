const xlsx = require("xlsx");
const readlineSync = require('readline-sync');
const fs = require("fs");

let options = {};
let sheets = {};
let cells = [];
let lines = [];

const data = fs.readdirSync("./filesForParsing");

if (data.length < 1) {
    readlineSync.question('There are no files in the folder: filesForParsing\nPress enter to exit');
    process.exit(-1);
} else if (data.length > 1) {
    readlineSync.question('You are trying to parse more than 1 file\nPress enter to exit');
    process.exit(-1);
}

options.workBook = xlsx.readFile(`filesForParsing/${data[0]}`);
for (let i = 0; i < Object.keys(options.workBook.Sheets).length; i++) {
    sheets[i] = Object.keys(options.workBook.Sheets)[i];
    console.log(`Number ${i} : ${Object.keys(options.workBook.Sheets)[i]}\n`);
}
options.nameOfSheet = sheets[readlineSync.question('Enter the number of sheet you want to parse: ')];
options.workSheet = options.workBook.Sheets[options.nameOfSheet];
if (!options.workSheet || Object.keys(options.workSheet).includes("!ref") == false) {
    readlineSync.question('The sheet is empty, or you choose not existing sheet\nPress enter to exit');
    process.exit(-1);
}
for (let cell in options.workSheet) {
    const cellAsString = cell.toString();
    cells.push(cellAsString.slice(0, 1));
    if (!cellAsString.includes("!")) lines.push(+cellAsString.slice(1));
}
options.lines = +readlineSync.question('Enter number of line, which you want to parse the latest (including) (0 - parse all lines): ');
if (options.lines == NaN || (!lines.includes(options.lines) && options.lines !== 0)) {
    readlineSync.question('You enter an invalid line number, or the line under this number is empty\nPress enter to exit');
    process.exit(-1);
}
options.columns = readlineSync.question('Enter column letter, which you want to parse the latest (including) (leave empty for parsing all columns): ').toUpperCase()
options.lastLetter = options.columns == "" ? options.workSheet["!ref"][3] : options.columns;
if (options.lastLetter !== "" && !cells.includes(options.lastLetter)) {
    readlineSync.question('You enter an invalid column letter, please enter column letter, which exists in your sheet\nPress enter to exit');
    process.exit(-1);
}


module.exports = options;

