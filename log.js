const chalk = require("chalk");
const { getNameLength } = require("./utils");

const handleLog = (val) => {
  let nameLength = getNameLength(val);
  let template = "=";
  let oneNum = Math.ceil((80 - nameLength) / 2);
  let oneStr = Array.from({ length: oneNum }, () => template).join("");
  return "\n" + oneStr + val + oneStr;
};

const log = (val) => {
  let str = handleLog(val);
  console.log(str);
};

let arr = ["red", "green", "yellow", "pink"];

arr.forEach((color) => {
  log[color] = (val) => {
    let realColor = color;
    let str =
      typeof val === "object" ? JSON.stringify(val, null, 4) : handleLog(val);
    if (color === "pink") realColor = "magentaBright";
    console.log(chalk[realColor](str));
  };
});
module.exports = log;
