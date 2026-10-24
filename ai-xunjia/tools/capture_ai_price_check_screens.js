const fs = require("fs");
const path = require("path");
const {
  chromium,
} = require("/Users/zouxiaoli/.cache/codex-runtimes/codex-primary-runtime/dependencies/node/node_modules/playwright");

const CHROME =
  "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome";
const DEMO = "file:///Users/zouxiaoli/Documents/ChatGPT/ai询价/ai询价demo.html";
const ROOT = path.resolve(__dirname, "..");
const OUT_DIR = path.join(ROOT, ".qa/ai_price_check_doc/pc");
const QUERY = "普通硅酸盐水泥 P.O42.5 广州 2026年9月";
const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

fs.mkdirSync(OUT_DIR, { recursive: true });

async function shoot(page, name) {
  const file = path.join(OUT_DIR, `${name}.png`);
  await page.screenshot({ path: file });
  console.log("shot", path.relative(ROOT, file));
}

async function openPriceCheck(page) {
  await page.click("#newChatBtn");
  await sleep(350);
  await page.click('[data-mode="AI查价"]');
  await sleep(250);
}

async function submitQuery(page) {
  await page.fill("#promptInput", QUERY);
  await page.keyboard.press("Enter");
}

(async () => {
  const browser = await chromium.launch({ executablePath: CHROME });
  const page = await browser.newPage({
    viewport: { width: 1440, height: 900 },
    deviceScaleFactor: 2,
  });

  await page.goto(DEMO);
  await sleep(600);
  await page.click("#openInquiryBtn");
  await sleep(400);
  await page.click('[data-mode="AI查价"]');
  await sleep(300);
  await shoot(page, "01_price_check_default");

  await page.click("#reasoningModeBtn");
  await sleep(300);
  await shoot(page, "02_mode_menu");

  await page.click("#deepReasoningOption");
  await sleep(350);
  await shoot(page, "03_deep_confirm");

  await page.click("#deepModeCancelBtn");
  await openPriceCheck(page);
  await submitQuery(page);
  await sleep(350);
  await shoot(page, "04_quick_thinking");
  await page.waitForSelector(".price-check-result-turn", { timeout: 5000 });
  await sleep(2400);
  await shoot(page, "05_quick_result");

  await openPriceCheck(page);
  await page.click("#reasoningModeBtn");
  await page.click("#deepReasoningOption");
  await page.click("#deepModeConfirmBtn");
  await sleep(2400);
  await submitQuery(page);
  await sleep(1250);
  await shoot(page, "06_deep_thinking");
  await page.waitForSelector(".price-check-result-turn", { timeout: 8000 });
  await sleep(2400);
  await shoot(page, "07_deep_result");

  await page.close();
  await browser.close();
  console.log("price check capture done");
})().catch((error) => {
  console.error(error);
  process.exit(1);
});
