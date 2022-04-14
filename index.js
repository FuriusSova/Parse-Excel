const xlsx = require("xlsx");
const fs = require("fs");
const { Parser } = require('json2csv');
const readlineSync = require('readline-sync');

const options = require("./config");

const json2csvParser = new Parser();

const posts = [];
let post = {};
const headers = [];
let counter = 0;
let checker;
let str = ``;

const data = fs.readdirSync("./parsedFiles");

for (const file of data) {
    fs.unlinkSync(`./parsedFiles/${file}`);
};

//Parsing xlsx
for (let cell in options.workSheet) {
    const cellAsString = cell.toString();

    if (cellAsString[0] === "A") checker = false;
    if (checker == true) continue;

    if (cellAsString[1] !== "r" && cellAsString[1] !== "m") {
        if (cellAsString.slice(1) == 1) {
            headers.push(options.workSheet[cell].v);
        } else if (cellAsString[0] !== options.lastLetter) {
            post[headers[counter]] = options.workSheet[cell].v.replace(/\r?\n/g, "  ");
            counter++;
        } else if (cellAsString[0] === options.lastLetter) {
            post[headers[counter]] = options.workSheet[cell].v.replace(/\r?\n/g, "  ");
            posts.push(post);
            post = {};
            counter = 0;
            checker = true;
        }
        if (options.lines !== 0 && cellAsString.slice(1) > options.lines) break;
    };
}

posts.forEach((elem) => {
    let keys = Object.keys(elem);
    keys.forEach((key) => {
        str += `${key} : ${elem[key]}\n`
    })
    str += "\n\n"
})

//Creating JSON file
fs.appendFileSync('./parsedFiles/data.json', `${JSON.stringify(posts)}`);

//Creating TXT file
fs.appendFileSync('./parsedFiles/data.txt', str);

//Creating CSV file
const csv = json2csvParser.parse(posts);
fs.appendFileSync('./parsedFiles/data.csv', csv, err => {
    if (err) {
        console.error(err)
        return
    } else {
        console.log("csv")
    }
})

readlineSync.question('Files have been parsed. Look into parsedFiles folder\nPress enter to exit');