const fs = require("fs");
const path = require("path");
function formatTableName(str) {
  return str.trim().replace(/(:|：)/, "");
}

function getDefaultValueFromContent(content, isWriteFile = true, tdAttachment) {
  fs.writeFileSync(path.resolve("./tempData/htmlContent"), content);
  let allNum,
    failNum,
    successNum,
    successRate,
    hasAttachment,
    tdAttachmentNum,
    attachmentName,
    publisher;
  let publisherRes = content.match(/([^，>。；原]+?局)(官网|网站)?(发布|通报|公布)/);
  if (publisherRes) {
    publisher = publisherRes[1].trim()
  } else {
    publisherRes = content.match(/从(原)?([^，>。；]+?局)(官网)?获悉/);
    publisherRes && (publisher = publisherRes[2]);
  }
  // let allRes = content.match(/抽(取|查)(了)?[^，。：]*?(\d+)(个)?(批次|组)/);
  let allRes = content.match(/抽(取|查|检)(了)?([^，。：]*?)(\d+)(个)?(批次|组)/);

  if (allRes) {
    // console.log(allRes[1],allRes[3],allRes[4]);
    allNum = Number(allRes[4]);
  }

  let failRes = content.match(/不合格(产品|样品)?(\d+)(个)?(批次|组)/);
  if (failRes) {
    failNum = Number(failRes[2]);
  } else {
    failRes = content.match(/(\d+)(组|批次)([^，。；]*?)不合格/);
    if (failRes) {
      failNum = Number(failRes[1]);
    }
  }

  let successRes = content.match(/[^不]合格(产品|样品)?(\d+)(个)?(批次|组)/);
  if (successRes) {
    successNum = Number(successRes[2]);
  } else {
    successRes = content.match(/(\d+)(组|批次)([^，。；不]*?)合格/);
    successRes && (successNum = Number(successRes[1]));
    //  console.log(successRes);
  }

  let successRateRes = content.match(/[^不]合格率(为)?((\d|\.)+(%|％))/);
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
  } else if (tdAttachment) {
    attachmentName = tdAttachment.split('\n').join('|');
    tdAttachmentNum = tdAttachment.split('\n').length
    tdAttachment && (hasAttachment = true);
  } else {
    let attachmentNameRes;
    let arr = [];
    // 情况1：<p>表格名</p><table
    let reg = />([^<]{1,60})<\/(p|div|strong)><table/g;
    while ((attachmentNameRes = reg.exec(content))) {
      // console.log('情况1');
      arr.push(formatTableName(attachmentNameRes[1]));
    }
    // 情况2：<div><strong>表格名</strong></div><div><table
    if (!arr.length) {
      reg = />([^<]{1,60})(<\/(div|p|strong)>)?<\/(div|p)>(<div[^>]*?>)?<table/g;
      while ((attachmentNameRes = reg.exec(content))) {
      // console.log('情况2');

        // console.log("=======push=========", attachmentNameRes[1]);
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
    tdAttachmentNum
  };
}

module.exports = getDefaultValueFromContent;
