const path = require("path");
const { chromium } = require("/Users/zouxiaoli/.cache/codex-runtimes/codex-primary-runtime/dependencies/node/node_modules/playwright");

const ROOT = path.resolve(__dirname, "..");
const SOURCE = `file://${path.join(ROOT, "tools/thinking_design.html")}`;
const OUTPUT = path.join(ROOT, ".qa/ai_xunjia_doc/pc/23_thinking_design.png");

(async () => {
  const browser = await chromium.launch({
    executablePath: "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome",
    headless: true,
  });
  const page = await browser.newPage({
    viewport: { width: 1600, height: 1000 },
    deviceScaleFactor: 2,
  });
  await page.goto(SOURCE);
  await page.waitForLoadState("networkidle");
  await page.screenshot({ path: OUTPUT, fullPage: false });
  await browser.close();
  console.log(OUTPUT);
})().catch((error) => {
  console.error(error);
  process.exit(1);
});
