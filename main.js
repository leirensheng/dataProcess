let json;
let log = require("./log");
const fs = require("fs");
const path = require("path");
const inquirer = require("inquirer");
let config = JSON.parse(fs.readFileSync("./config.json"));
let startIndex = config.index;
let utils = require("./utils");
const chalk = require("chalk");
let getDefaultValueFromContent = require("./getDefaultValueFromContent");
let outputNewSheet = require("./generateExcel");
let keyMap = require("./keyMap");
let getExcelJson = require('./getExcel')

async function getContent(page) {
  let obj = {};
  try {
    await page.waitForSelector(".article-content", { timeout: 5000 });
    await page.waitForFunction(() => {
      let dom = document.querySelector(".article-content");
      return dom.innerHTML;
    });
    obj = await page.$eval(".article-content", (el) => {
      let str = el.innerHTML;
      // 去掉回车
      str = str.replace(/\n/g, "");
      // 去掉空标签
      str = str.replace(
        /<(span|p|div|strong)[^>]*?>(&nbsp;)?<\/(p|strong|div|span)>/g,
        ""
      );
      // 去掉注释
      str = str.replace(/<!--.*?-->/g, "");
      return {
        content: str,
        innerText: el.innerText,
      };
    });
  } catch (e) {
    obj = {
      content: "",
      innerText: "",
    };
  }
  let tdAttachment = await page.$eval(
    "tr:nth-child(3) td:nth-child(2)",
    (el) => el.innerText
  );
  return { ...obj, tdAttachment };
}

function getHasHandleNextIndex() {
  let i = 0;
  let obj = json[i];
  while (obj && obj["处理人"]) {
    i++;
    obj = json[i];
  }
  return i;
}

function save() {
  fs.writeFileSync(
    path.resolve("./config.json"),
    JSON.stringify(config, null, 4)
  );
  fs.writeFileSync(path.resolve("./excel.json"), JSON.stringify(json, null, 4));
  try {
    let handledNextIndex = getHasHandleNextIndex();
    let jsonForExcel = JSON.parse(
      JSON.stringify(json.slice(0, handledNextIndex))
    );
    jsonForExcel.forEach((one, index) => (one.index = index));
    outputNewSheet(jsonForExcel, config);
  } catch (e) {
    console.log(e);
    log.yellow("文件已打开，未写入成功，请关闭output.xlsx，继续录入下一条，录完下一条会补写进去");
  }
}

async function closeTab(page) {
  await page.waitForFunction(() => {
    let arr = [...document.querySelectorAll(".el-icon-close")];
    let lastOne = arr.pop();
    lastOne.click();
    return true;
  });
}

function isNoContent(obj, key) {
  return !(obj[key] || "").trim();
}

function getProgress(index) {
  return `【进度】${index}/${json.length}≈${(
    (index / json.length) *
    100
  ).toFixed(2)}%`;
}

function getSameSnapshotIdRows(index) {
  let snapshotId = json[index]["公告snapshot_id"];
  let i = 1;
  let curIndex = index + 1;
  let obj = json[curIndex];
  if (!obj) return i;
  while (obj["公告snapshot_id"] === snapshotId) {
    i++;
    curIndex++;
    obj = json[curIndex];
  }
  return i;
}

function removeSubRows(obj, index) {
  log.pink("当前处于编辑模式");
  let id = obj[keyMap["表格id"]];
  let arr = String(id).split(".");
  if (arr.length === 2) {
    throw new Error("当前index不对");
  } else {
    let sameSnapshotIdRows = getSameSnapshotIdRows(index);
    if (sameSnapshotIdRows !== 1) {
      json.splice(index + 1, sameSnapshotIdRows - 1);
    }
  }
}

function checkDuplicate(obj, index) {
  if (index === 0) return false;
  let preSnap = json[index - 1];
  if (obj["公告标题"] === preSnap["公告标题"]) {
    obj[keyMap["备注"]] = `与${preSnap["公告snapshot_id"]}重复`;
    log.yellow(`与${preSnap["公告snapshot_id"]}重复`);
    return true;
  }
  return false;
}

function addHandler(obj) {
  obj["处理人"] = config.userInfo.handler;
  obj["处理时间"] = utils.formatDate(new Date());
}

async function handleOneSnapshot(content, index, tdAttachment, innerText) {
  let obj = json[index];

  log(`【snapshotId】${obj["公告snapshot_id"]}====${getProgress(index)}`);
  log.yellow(obj);

  let isEdit = obj["处理人"];
  // 编辑的时候不考虑和上一条重复与否
  if (isEdit) {
    // 先删除有小数的同个公告的行
    removeSubRows(obj, index);
  } else {
    let isDuplicate = checkDuplicate(obj, index);
    if (isDuplicate) {
      addHandler(obj);
      return 1;
    }
  }
  addHandler(obj);

  let startTime = Date.now();

  let template = { ...obj };
  let defaultValue = getDefaultValueFromContent(
    content,
    true,
    tdAttachment,
    innerText
  );
  let {
    allNum,
    failNum,
    successNum,
    successRate,
    hasAttachment,
    attachmentName,
    tdAttachmentNum,
  } = defaultValue;

  log("从正文解析");
  log.pink(defaultValue);
  const { rows, attachmentNum } = await inquirer.prompt([
    {
      type: "number",
      name: "rows",
      default: () => {
        if (tdAttachmentNum) {
          return tdAttachmentNum;
        }
        if (attachmentName) {
          return attachmentName.split("|").length;
        }
        let oneType =
          /(全部合格)|(全部不合格)/.test(content) ||
          successNum === 0 ||
          failNum === 0 ||
          successRate === "100%";
        if (oneType) return 1;

        let notEqual = allNum !== failNum || allNum !== successNum;
        return notEqual ? 2 : 1;
      },
      validate: (val) => (isNaN(val) ? "请输入数字！（按⬆重新输入）" : true),
      message: "当前公告在excel中会有多少行？",
    },
    {
      type: "number",
      name: "attachmentNum",
      default: () => (hasAttachment ? tdAttachmentNum || 1 : 0),
      validate: (val) => (isNaN(val) ? "请输入数字！（按⬆重新输入）" : true),
      message: "附件数量✉️",
    },
  ]);

  let curIndex = index;
  for (let i = 0; i < rows; i++) {
    await addOneRow({
      row: i + 1,
      index: curIndex,
      content,
      attachmentNum,
      defaultValue,
      rows,
    });
    if (i !== rows - 1) {
      let curId = json[curIndex][keyMap["表格id"]];
      let newId = curId + 0.01;
      let newObj = {
        ...template,
        公布总抽检批次数: obj["公布总抽检批次数"],
        公布合格率: obj["公布合格率"],
        发布主体: obj["发布主体"],
        [keyMap["表格id"]]: newId,
      };
      curIndex++;
      json.splice(curIndex, 0, newObj);
    }
  }
  let useTime = (Date.now() - startTime) / 1000;

  log("【处理结果】");
  let res = json.slice(index, curIndex + 1);
  log.green(res);
  log(`${chalk.green("✅")}${obj["公告snapshot_id"]}，用时${useTime}秒`);
  return rows;
}

async function addOneRow({
  row,
  index,
  content,
  attachmentNum,
  defaultValue,
  rows,
}) {
  let obj = json[index];
  if (rows > 1) {
    log(
      `【snapshotId】${obj["公告snapshot_id"]}第${row}行======${getProgress(
        index
      )}`
    );
  }
  let {
    allNum,
    failNum,
    successNum,
    successRate,
    attachmentName,
    publisher,
  } = defaultValue;

  let curAttachmentName = attachmentName;
  if (attachmentName && attachmentName.indexOf("|") !== -1) {
    let arr = attachmentName.split("|");
    curAttachmentName = arr[row - 1];
  }

  let promptConfig = [
    {
      key: "源文件类型【1-EXCEL，2-PDF，3-WORD，4-HTML，5-jpg】",
      type: "rawlist",
      message: "源文件类型",
      default: () => (attachmentNum ? "1" : "4"),
      choices: [
        { name: "EXCEL", value: "1" },
        { name: "PDF", value: "2" },
        { name: "WORD", value: "3" },
        { name: "HTML", value: "4" },
        { name: "jpg", value: "5" },
      ],
    },
    {
      key: "发布主体",
      default: publisher,
      when: () => isNoContent(obj, "发布主体") && row === 1,
    },
    {
      key: "公布表格表名",
      default: curAttachmentName,
    },
    {
      key: "公布表格合格条目数【抽检表格条目数】",
      message: "合格数✔️",
      type: "number",
      default: successNum,
      when: (answer) => /未发现(不符合)|(不合格)/.test(answer["公布表格表名"])||!/(不符合)|(不合格)/.test(answer["公布表格表名"]),
      validate: (val) => (isNaN(val) ? "请输入数字！（按⬆重新输入）" : true),
    },
    {
      key: "公布表格不合格条目数【抽检表格条目数】",
      message: "不合格数❌",
      type: "number",
      default: failNum,
      validate: (val) => (isNaN(val) ? "请输入数字！（按⬆重新输入）" : true),
      when: (answer) =>
        !/[^不]合格|(^合格)|[^不]符合|(^符合)/.test(answer["公布表格表名"]),
    },
    {
      key: "公布总抽检批次数",
      default: row === 1 ? allNum : obj["公布总抽检批次数"],
      type: "number",
      validate: (val) => (isNaN(val) ? "请输入数字！（按⬆重新输入）" : true),
    },
    {
      key: "公布合格率",
      default: row === 1 ? successRate : obj["公布合格率"],
      when:/[^不]合格率/.test(content),
      validate: (val) =>
        val.match(/(\d+)(\.(\d)+)?%/) ? true : "请输入百分数！（按⬆重新输入）",
    },
    {
      key: "备注（缺少表格的，需要注明表格）",
      default: (answer) =>
        !answer["公布表格表名"] && !attachmentNum ? "无表格" : "",
    },
    {
      key: "是否存在合并单元格(1-是，2-否)",
      type: "rawlist",
      default: "2",
      choices: [
        { name: "是", value: "1" },
        { name: "否", value: "2" },
      ],
    },
  ];

  promptConfig.forEach((one) => {
    one.message = one.message || one.key;
    one.name = one.key;
    if (!one.type === "number") {
      one.filter = (val) => {
        return val.trim();
      };
    }
  });
  const res = await inquirer.prompt(promptConfig);

  Object.keys(res).forEach((key) => {
    let val = res[key];
    // 剔除NaN
    if (val !== 0 && val === val) {
      obj[key] = val;
    }
  });

  if (attachmentNum) {
    obj[
      "公告附件数量（仅填EXCEL、PDF、WORD附件的数量，无附件不需要填）"
    ] = attachmentNum;
  }
}

async function login(targetUrl, page) {
  await page.goto(targetUrl);
  try {
    await page.waitForFunction(() => location.href.indexOf("/login") !== -1, {
      timeout: config.loginTimeout,
    });
    log("需要登录");
    // await page.waitForSelector('[placeholder=Username]')
    await page.type("[placeholder=Username]", config.userInfo.username);
    await page.type("[type=password]", config.userInfo.password);
    log("请在页面输入验证码");
    await inputCode(page);
  } catch (e) {
    log("不需要登录");
  }
}

async function inputCode(page) {
  await page.waitForFunction(
    () => {
      let code = document.querySelector("[placeholder=验证码]").value;
      return code.length === 4;
    },
    { timeout: 60000 }
  );
  await page.click("button");
  try {
    await page.waitForFunction(() => location.href.indexOf("/login") === -1, {
      timeout: 1000,
    });
  } catch (e) {
    log.red("验证码错误");
    await page.$eval("[placeholder=验证码]", (el) => (el.value = ""));
    await inputCode(page);
  }
}


async function start(page) {
  let targetUrl = config.origin + "/dc-admin#/inspection/spiderArticle/list";
  await login(targetUrl, page);
  json = await getExcelJson();
  if (json.length < startIndex) {
    throw new Error("输入的index对应的数据不存在，请检查输入的index");
  }
  for (let i = startIndex; i < json.length; ) {
    let obj = json[i];
    await page.goto(
      config.origin +
        "/dc-admin#/inspection/inspection/snap/" +
        obj["公告snapshot_id"]
    );

    let { content, tdAttachment, innerText } = await getContent(page);
    let rows = await handleOneSnapshot(content, i, tdAttachment, innerText);

    config.index = i + rows;
    config.lastId = json[i][keyMap["表格id"]];
    config.lastSnapshotId = json[i]["公告snapshot_id"];
    save();
    await closeTab(page);
    i += rows;
  }
  log.green("所有数据处理完毕!!!");
}

module.exports = start;
