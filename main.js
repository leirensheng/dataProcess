let json;
let log = require("./log");
const fs = require("fs");
const path = require("path");
const inquirer = require("inquirer");
const XLSX = require("xlsx");
let config = JSON.parse(fs.readFileSync("./config.json"));
let startIndex = config.index;
let utils = require("./utils");
const chalk = require("chalk");

function outputNewSheet(res) {
  const wb = XLSX.utils.book_new();
  const sheet = XLSX.utils.json_to_sheet(res);
  XLSX.utils.book_append_sheet(wb, sheet, config.sheetName);
  XLSX.writeFile(wb, path.resolve(__dirname, "./output.xlsx"));
}

async function getContent(page) {
  try {
    await page.waitForSelector(".article-content", { timeout: 5000 });
    await page.waitForFunction(() => {
      let dom = document.querySelector(".article-content");
      return dom.innerHTML;
    });
    content = await page.$eval(".article-content", (el) => {
      let str = el.innerHTML;
      // 去掉回车
      str = str.replace(/\n/g, "");
      // 去掉空标签
      str = str.replace(
        /<(span|p|div|strong)[^>]*?><\/(p|strong|div|span)>/g,
        ""
      );
      return str;
    });
  } catch (e) {
    content = "";
  }
  return content;
}

function getHasHandleIndex() {
  let i = 0;
  let obj = json[i];
  while (obj["处理人"]) {
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
    let handledIndex = getHasHandleIndex();
    let jsonForExcel = JSON.parse(JSON.stringify(json.slice(0, handledIndex)));
    jsonForExcel.forEach((one, index) => (one.index = index));
    outputNewSheet(jsonForExcel);
  } catch (e) {
    log.yellow(
      `\n========文件已经打开，未能写入成功，下次保存数据写入!==========`
    );
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
  return `进度：${index}/${json.length}≈${((index / json.length) * 100).toFixed(
    2
  )}%`;
}

function getSameSnapshotIdRows(index) {
  let snapshotId = json[index]["公告snapshot_id"];
  let i = 1;
  let curIndex = index + 1;
  let obj = json[curIndex];
  while (obj["公告snapshot_id"] === snapshotId) {
    i++;
    curIndex++;
    obj = json[curIndex];
  }
  return i;
}

// 上一条公告的最后一行
function getPreSnap(index, snapshotId) {
  if (!snapshotId) {
    return {};
  }
  let curIndex = index - 1;
  let obj = json[curIndex];
  while (obj["公告snapshot_id"] !== snapshotId) {
    curIndex--;
    obj = json[curIndex];
  }
  return obj;
}

function removeSubRows(obj, index) {
  log.pink("===================当前处于编辑模式==========================");
  let id =
    obj[
      "表格id（如有多个附件，请插入1行，在原表格id上加.1,如1.1,1.2，不要合并任何单元格）"
    ];
  let arr = String(id).split(".");
  if (arr.length === 2) {
    throw new Error("当前index不对");
  } else {
    let sameSnapshotIdRows = getSameSnapshotIdRows(index);
    if (sameSnapshotIdRows !== 1) {
      console.log(`=======删除${sameSnapshotIdRows - 1}条=========`);
      json.splice(index + 1, sameSnapshotIdRows - 1);
    }
  }
}

function checkDuplicate(obj, index) {
  let preSnap = getPreSnap(index, config.lastSnapshotId);
  if (obj["公告标题"] === preSnap["公告标题"]) {
    obj[
      "备注（缺少表格的，需要注明表格）"
    ] = `与${preSnap["公告snapshot_id"]}重复`;
    log.yellow(`========与${preSnap["公告snapshot_id"]}重复==========`);
    return true;
  }
  return false;
}

function addHandler(obj) {
  obj["处理人"] = config.userInfo.handler;
  obj["处理时间"] = utils.formatDate(new Date());
}

async function handleOneSnapshot(content, index) {
  let obj = json[index];

  console.log(
    `\n========📣snapshotId：${obj["公告snapshot_id"]}====${getProgress(
      index
    )} ========上一条${config.lastSnapshotId}📣============`
  );
  log.green(obj);

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
  let defaultValue = getDefaultValueFromContent(content);
  let {
    allNum,
    failNum,
    successNum,
    successRate,
    hasAttachment,
    attachmentName,
  } = defaultValue;

  console.log("\n===========从正文解析===========");
  log.pink(defaultValue);
  const { rows, attachmentNum } = await inquirer.prompt([
    {
      type: "number",
      name: "rows",
      default: () => {
        let oneType =
          /(全部合格)|(全部不合格)/.test(content) ||
          successNum === 0 ||
          failNum === 0 ||
          successRate === "100%";
        if (oneType) return 1;
        if (attachmentName) {
          return attachmentName.split("|").length;
        }
        let notEqual = allNum !== failNum || allNum !== successNum;
        return notEqual ? 2 : 1;
      },
      // validate:(val)=> typeof val === 'number',
      message: "当前公告有多少行？",
    },
    {
      type: "number",
      name: "attachmentNum",
      default: () => (hasAttachment ? 1 : 0),
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
    });
    if (i !== rows - 1) {
      let curId =
        obj[
          "表格id（如有多个附件，请插入1行，在原表格id上加.1,如1.1,1.2，不要合并任何单元格）"
        ];
      let newId = curId + 0.01;
      let newObj = {
        ...template,
        发布主体: obj["发布主体"],
        "表格id（如有多个附件，请插入1行，在原表格id上加.1,如1.1,1.2，不要合并任何单元格）": newId,
      };
      curIndex++;
      json.splice(curIndex, 0, newObj);
    }
  }
  let useTime = (Date.now() - startTime) / 1000;

  console.log(
    `========================${chalk.green("✅")}${
      obj["公告snapshot_id"]
    }，用时${useTime}秒 =================================`
  );
  return rows;
}

function formatTableName(str) {
  return str.trim().replace(/(:|：)/, "");
}

function getDefaultValueFromContent(content) {
  fs.writeFileSync(path.resolve("./tempData/htmlContent"), content);
  let allNum,
    failNum,
    successNum,
    successRate,
    hasAttachment,
    attachmentName,
    publisher;
  let publisherRes = content.match(/([^，。；原]+?)(发布|通报)/);
  if (publisherRes) {
    publisher = publisherRes[1].trim().replace(/(官网|网站)/, "");
  }
  let allRes = content.match(/抽(取|查)(了)?[^，。]*?(\d+)(批次|组)/);
  if (allRes) {
    allNum = Number(allRes[3]);
  }

  let failRes = content.match(/不合格(产品|样品)?(\d+)(批次|组)/);
  if (failRes) {
    failNum = Number(failRes[2]);
  } else {
    failRes = content.match(/(\d+)(组|批次)([^，。；]*?)不合格/);
    if (failRes) {
      failNum = Number(failRes[1]);
    }
  }

  let successRes = content.match(/[^不]合格(产品|样品)?(\d+)(批次|组)/);
  if (successRes) {
    successNum = Number(successRes[2]);
  }

  let successRateRes = content.match(/[^不]合格率(为)?(.*?%)/);
  if (successRateRes) {
    successRate = successRateRes[2];
  }

  let attachmentRes = content.match(/附件：/);
  if (attachmentRes) {
    hasAttachment = true;
    let attachmentNameRes = content.match(
      /附件：(.*?)\.(xlsx|doc|docx|pdf|xls)/
    );
    if (attachmentNameRes) {
      attachmentName = attachmentNameRes[1].replace(/<.*?>/g, "");
    }
  } else {
    let attachmentNameRes;
    let arr = [];
    // 情况1：<p>表格名</p><table
    let reg = />([^<]{1,32})<\/(p|div|strong)><table/g;
    while ((attachmentNameRes = reg.exec(content))) {
      console.log("=======push=========", attachmentNameRes[1]);
      arr.push(formatTableName(attachmentNameRes[1]));
    }
    // 情况2：<div><strong>表格名</strong></div><div><table
    if (!attachmentName) {
      reg = />([^<]{1,32})(<\/(div|p|strong)>)?<\/(div|p)><div[^>]*?><table/g;
      while ((attachmentNameRes = reg.exec(content))) {
        console.log("=======push=========", attachmentNameRes[1]);
        arr.push(formatTableName(attachmentNameRes[1]));
      }
    }

    attachmentName = arr.join("|");
  }

  return {
    allNum,
    failNum,
    successNum,
    successRate,
    hasAttachment,
    attachmentName,
    publisher,
  };
}

async function addOneRow({ row, index, content, attachmentNum, defaultValue }) {
  let obj = json[index];
  console.log(
    `\n===========snapshotId：${
      obj["公告snapshot_id"]
    }第${row}行======${getProgress(index)} ==========`
  );
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
      key: "发布主体",
      default: publisher,
      when: () => isNoContent(obj, "发布主体") && row === 1,
    },
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
      key: "公布表格表名",
      default: curAttachmentName,
    },
    {
      key: "公布表格合格条目数【抽检表格条目数】",
      message: "合格数✔️",
      type: "number",
      default: successNum,
      when: (answer) => !/(不符合)|(不合格)/.test(answer["公布表格表名"]),
    },
    {
      key: "公布表格不合格条目数【抽检表格条目数】",
      message: "不合格数❌",
      type: "number",
      default: failNum,
      when: (answer) =>
        !/[^不]合格|(^合格)|[^不]符合|(^符合)/.test(answer["公布表格表名"]),
    },
    {
      key: "公布总抽检批次数",
      default: allNum,
      type: "number",
    },
    {
      key: "公布合格率",
      default: successRate,
      when: () => successRate,
    },
    {
      key: "备注（缺少表格的，需要注明表格）",
    },
    {
      key: "是否存在合并单元格(1-是，2-否)",
      default: "2",
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
    if (val !== 0 && !isNaN(val)) {
      obj[key] = val;
    }
  });

  console.log(obj);
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
      timeout: 1000,
    });
    console.log("========需要登录=========");
    // await page.waitForSelector('[placeholder=Username]')
    await page.type("[placeholder=Username]", config.userInfo.username);
    await page.type("[type=password]", config.userInfo.password);
    console.log("========请在页面输入验证码=========");
    await inputCode(page);
  } catch (e) {
    console.log("========不需要登录=========");
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
      timeout: 500,
    });
  } catch (e) {
    log.red("===========验证码错误============");
    await page.$eval("[placeholder=验证码]", (el) => (el.value = ""));
    await inputCode(page);
  }
}

function formatId(arr) {
  for (let i = 0; i < arr.length; i++) {
    let key =
      "表格id（如有多个附件，请插入1行，在原表格id上加.1,如1.1,1.2，不要合并任何单元格）";
    let obj = arr[i];
    let id = obj[key];
    if (typeof id === "string") {
      obj[key] = Number(obj[key].trim());
    }
  }
}

async function getExcelJson() {
  let readFromExcel = !fs.existsSync(path.resolve("./excel.json"));
  if (readFromExcel) {
    const filepath = path.resolve(__dirname, "./tempData", config.fileName);
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

async function start(page) {
  let targetUrl = config.origin + "/dc-admin#/inspection/spiderArticle/list";
  await login(targetUrl, page);
  json = await getExcelJson();
  for (let i = startIndex; i < json.length; ) {
    let obj = json[i];
    await page.goto(
      config.origin +
        "/dc-admin#/inspection/inspection/snap/" +
        obj["公告snapshot_id"]
    );

    let content = await getContent(page);
    let rows = await handleOneSnapshot(content, i);

    config.index = i + rows;
    config.lastId =
      json[i][
        "表格id（如有多个附件，请插入1行，在原表格id上加.1,如1.1,1.2，不要合并任何单元格）"
      ];
    config.lastSnapshotId = json[i]["公告snapshot_id"];
    save();
    await closeTab(page);
    i += rows;
  }
  log.green("========== 所有数据处理完毕!!! =============");
}

module.exports = start;
