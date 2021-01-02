const puppeteer = require("puppeteer-core");
const fs = require("fs");
const findChrome = require("./node_modules/carlo/lib/find_chrome");
const start = require("./main.js");
const config = JSON.parse(fs.readFileSync("./config.json"));
const log = require('./log')

async function init() {
  let findChromePath = await findChrome({});
  let executablePath = findChromePath.executablePath;
  const browser = await puppeteer.launch({
    executablePath,
    headless: false,
    defaultViewport: config.defaultViewport,
    userDataDir: "./tempData/data", //可以重用数据，cookie 和缓存
    // args: ['--start-fullscreen'], //全屏打开页面
    // slowMo:10
  });
  const page = await browser.newPage();
  await page.setRequestInterception(true);
  page.on("request", (request) => {
    if ([""].includes(request.resourceType())) request.abort();
    else request.continue();
  });
  // page.on("console", (msg) => {
  //   for (let i = 0; i < msg.args().length; ++i)
  //     console.log(`-----浏览器输出------: ${msg.args()[i]}`);
  // });
  try {
    await start(page);
  } catch (e) {
    log.red(`❌${e.message}❌`);
    console.log(e);
  }
}
init();
