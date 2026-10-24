const fs = require("fs");
const path = require("path");
const {
  chromium,
} = require("/Users/zouxiaoli/.cache/codex-runtimes/codex-primary-runtime/dependencies/node/node_modules/playwright");

const CHROME = "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome";
const ROOT = path.resolve(__dirname, "..");
const DEMO = `file://${path.join(ROOT, "ai询价demo.html")}`;
const OUT_DIR = path.join(ROOT, ".qa/ai_price_check_doc/pc_current");
const QUERY = "普通硅酸盐水泥 P.O42.5 广州 2026年9月";
const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

fs.mkdirSync(OUT_DIR, { recursive: true });

async function captureMain(page, name) {
  const file = path.join(OUT_DIR, `${name}.png`);
  await page.locator(".workspace-main").screenshot({
    path: file,
    animations: "disabled",
  });
  console.log("shot", path.relative(ROOT, file));
}

async function openPriceCheck(page) {
  await page.goto(DEMO);
  await sleep(600);
  await page.click("#openInquiryBtn");
  await sleep(300);
  await page.click('[data-mode="AI查价"]');
  await sleep(250);
}

async function submitQuery(page) {
  await page.fill("#promptInput", QUERY);
  await page.keyboard.press("Enter");
  await page.waitForSelector(".price-check-loading-head", { timeout: 5000 });
}

(async () => {
  const browser = await chromium.launch({ executablePath: CHROME });
  const page = await browser.newPage({
    viewport: { width: 1440, height: 760 },
    deviceScaleFactor: 2,
  });

  await openPriceCheck(page);
  await submitQuery(page);
  await sleep(650);
  await captureMain(page, "08_quick_loading_current");

  await openPriceCheck(page);
  await page.click("#reasoningModeBtn");
  await page.click("#deepReasoningOption");
  await page.click("#deepModeConfirmBtn");
  await submitQuery(page);
  await sleep(1050);
  await captureMain(page, "09_deep_loading_current");

  await page.close();
  await browser.close();
  console.log("price check loading capture done");
})().catch((error) => {
  console.error(error);
  process.exit(1);
});
