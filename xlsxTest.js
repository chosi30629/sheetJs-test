const XLSX = require("xlsx");
const moment = require("moment");

var workbook = XLSX.readFile("test.xlsx", { cellDates: true, dateNF: "YYYY-MM-DD hh:mm:ss" });
// console.log(workbook.SheetNames);

var firstSheetName = workbook.SheetNames[0];
var workSheet = workbook.Sheets[firstSheetName];

console.log(workSheet);

// for (const work in workSheet) {
//     if (workSheet[work].hasOwnProperty("v") && workSheet[work].t === "d") {
//         workSheet[work].v = moment(workSheet[work].w);
//     }
// }

var addressOfCell = "A1";
var deSiredCell = workSheet[addressOfCell];
// console.log(deSiredCell);

// console.log(XLSX.utils.sheet_to_html(workSheet));
// console.log(XLSX.utils.sheet_to_json(workSheet, { raw: false, defval: "" }));
console.log(XLSX.utils.sheet_to_json(workSheet, { raw: true }));