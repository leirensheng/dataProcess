let fs = require('fs')
const XLSX = require("xlsx");
const path = require("path");
let config = JSON.parse(fs.readFileSync("./config.json"));
let keyMap = require("./keyMap");

async function getExcelJson(getFromExistFile= false) {
  let readFromExcel = getFromExistFile && !fs.existsSync(path.resolve("./excel.json"));
  if (readFromExcel) {
    const filepath = path.resolve(config.fileName);
    const workbook = XLSX.readFile(filepath);
    const { Sheets } = workbook;
    const res = XLSX.utils.sheet_to_json(Sheets[config.sheetName], {
      raw: false,
    });
    formatId(res);
    return res;
  } else {
    let res = JSON.parse(fs.readFileSync(path.resolve("./excel.json")));
    formatId(res);
    return res;
  }
}

function formatId(arr) {
  for (let i = 0; i < arr.length; i++) {
    let key = keyMap["表格id"];
    let obj = arr[i];
    let id = obj[key];
    if (typeof id === "string") {
      obj[key] = Number(obj[key].trim());
    }
  }
}

module.exports = getExcelJson