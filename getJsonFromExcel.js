const fs = require("fs");
const path = require("path");
const XLSX = require("xlsx");
let config = JSON.parse(fs.readFileSync("./config.json"));
 


function getExcelJson() {
  const filepath = path.resolve(config.fileName);
  const workbook = XLSX.readFile(filepath);
  const { Sheets } = workbook;
  const res = XLSX.utils.sheet_to_json(Sheets[config.sheetName], {
    raw: false,
  });
  return res;
}


json =   getExcelJson();
json.forEach(one=> {
  one.percent=one.percent.replace('%','')
})
fs.writeFileSync('./tempData/htmlContent',JSON.stringify(json,null,4))