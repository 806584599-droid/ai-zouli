const fs = require("fs");
const path = require("path");
const {
  chromium,
} = require("/Users/zouxiaoli/.cache/codex-runtimes/codex-primary-runtime/dependencies/node/node_modules/playwright");

const CHROME =
  "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome";
const DEMO = "file:///Users/zouxiaoli/Documents/ChatGPT/ai询价/ai询价demo.html";
const ROOT = path.resolve(__dirname, "..");
const PC_DIR = path.join(ROOT, ".qa/ai_xunjia_doc/pc");
const M_DIR = path.join(ROOT, ".qa/ai_xunjia_doc/m");

const RESULT_WAIT = 5200;
const SHORT_WAIT = 900;

fs.mkdirSync(PC_DIR, { recursive: true });
fs.mkdirSync(M_DIR, { recursive: true });

const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

async function shoot(page, dir, name) {
  const file = path.join(dir, `${name}.png`);
  await page.screenshot({ path: file });
  console.log("shot", path.relative(ROOT, file));
}

async function submit(page, text) {
  await page.fill("#promptInput", text);
  await page.keyboard.press("Enter");
}

async function scrollTo(page, selector, block = "start") {
  await page.evaluate(
    ({ selector, block }) => {
      document.querySelector(selector)?.scrollIntoView({ block, behavior: "instant" });
    },
    { selector, block }
  );
  await sleep(250);
}

async function scrollToLast(page, selector, block = "end") {
  await page.evaluate(
    ({ selector, block }) => {
      const nodes = document.querySelectorAll(selector);
      nodes[nodes.length - 1]?.scrollIntoView({ block, behavior: "instant" });
    },
    { selector, block }
  );
  await sleep(250);
}

async function scrollToUserMessage(page, text) {
  await page.evaluate((message) => {
    const target = [...document.querySelectorAll(".user-message")].find((node) =>
      node.textContent.includes(message)
    );
    target?.scrollIntoView({ block: "start", behavior: "instant" });
  }, text);
  await sleep(250);
}

async function newChat(page) {
  await page.click("#newChatBtn");
  await sleep(350);
}

// 项目阶段提示已后置：先出询价建议结果，再在结果下方弹出阶段选择。
async function waitForOutcome(page) {
  await page.waitForSelector(
    "#phaseChoicePanel, .price-value, .no-data-message, #incompleteChoicePanel",
    { timeout: 25000 }
  );
}

async function resolveOutcome(page) {
  await waitForOutcome(page);
  if (await page.isVisible(".price-value")) {
    try {
      await page.waitForSelector("#phaseChoicePanel", { timeout: 4000 });
    } catch (error) {
      // 该分支没有阶段提示时继续
    }
  }
  if (await page.isVisible("#phaseChoicePanel")) {
    await page.click('[data-phase="施工图预算"]');
    await page.waitForSelector(".price-value", { timeout: 15000 });
  }
}

async function capturePc(browser) {
  const page = await browser.newPage({
    viewport: { width: 1440, height: 900 },
    deviceScaleFactor: 2,
  });
  await page.goto(DEMO);
  await sleep(600);
  await shoot(page, PC_DIR, "01_home");

  await page.click("#openInquiryBtn");
  await sleep(500);
  await shoot(page, PC_DIR, "02_workspace_empty");

  await submit(page, "广东广州，2026年9月，普通硅酸盐水泥 P.O42.5，数量200吨");
  await sleep(1000);
  await shoot(page, PC_DIR, "03_thinking_prepare");
  await sleep(1100);
  await shoot(page, PC_DIR, "04_thinking_steps");
  await page.waitForSelector(".price-value", { timeout: 25000 });
  await page.waitForSelector("#phaseChoicePanel", { timeout: 25000 });
  await sleep(500);
  await shoot(page, PC_DIR, "18_phase_question");
  await resolveOutcome(page);
  await sleep(500);
  await scrollTo(page, "[data-inquiry-result]");
  await shoot(page, PC_DIR, "05_result_cement");

  await scrollTo(page, "[data-inquiry-result]");
  await page.click('[data-result-feedback="like"]');
  await sleep(400);
  await shoot(page, PC_DIR, "06_result_feedback");

  await submit(page, "供应商联系方式");
  await sleep(SHORT_WAIT);
  await scrollToLast(page, ".agent-turn");
  await shoot(page, PC_DIR, "07_supplier_contact");

  await submit(page, "你给的供应商电话不准确");
  await sleep(SHORT_WAIT);
  await scrollToLast(page, ".agent-turn");
  await shoot(page, PC_DIR, "08_supplier_unavailable");

  await page.evaluate(() => {
    window.__manualJumpCalls = [];
    window.open = (url) => {
      window.__manualJumpCalls.push(String(url));
      return null;
    };
  });
  await page.click("#supplierContactManualBtn");
  await sleep(600);
  const manualJumpCalls = await page.evaluate(() => window.__manualJumpCalls);
  if (manualJumpCalls.at(-1) !== "https://xunjia.zjtcn.com/release.html") {
    throw new Error(`转人工跳转地址异常：${manualJumpCalls.join(", ")}`);
  }
  await shoot(page, PC_DIR, "16_manual_inquiry_jump");

  await newChat(page);
  await submit(page, "广东中山，2026年9月，PE直通管件PE110，数量650个");
  await resolveOutcome(page);
  await sleep(500);
  await scrollTo(page, "[data-inquiry-result]");
  await shoot(page, PC_DIR, "09_result_insufficient");

  await newChat(page);
  await submit(page, "PE管DN110，数量4600米");
  await resolveOutcome(page);
  await sleep(500);
  await scrollTo(page, "[data-inquiry-result]");
  await shoot(page, PC_DIR, "21_insufficient_conditions");

  await newChat(page);
  await submit(page, "广东广州，2026年9月，特种调节阀 DN1200，数量2台");
  await resolveOutcome(page);
  await sleep(500);
  await scrollTo(page, ".no-data-message");
  await shoot(page, PC_DIR, "10_no_data");

  await newChat(page);
  await submit(page, "特种调节阀，数量2台");
  await resolveOutcome(page);
  await sleep(500);
  await scrollToLast(page, "#incompleteChoicePanel", "end");
  await shoot(page, PC_DIR, "11_incomplete_spec");

  await newChat(page);
  await submit(page, "今天天气怎么样");
  await sleep(SHORT_WAIT);
  await shoot(page, PC_DIR, "12_invalid_input");

  await newChat(page);
  await submit(page, "广东广州，2026年9月，普通硅酸盐水泥 P.O42.5，数量200吨");
  await resolveOutcome(page);
  await submit(page, "地区改成湖南");
  await resolveOutcome(page);
  await sleep(600);
  await scrollToUserMessage(page, "地区改成湖南");
  await shoot(page, PC_DIR, "24_condition_override");

  await newChat(page);
  await submit(page, "广东广州，2026年9月，普通硅酸盐水泥 P.O42.5，数量200吨");
  await resolveOutcome(page);
  await sleep(500);
  await page.click("#retryBtn");
  await page.waitForSelector(".result-refresh-note", { timeout: 25000 });
  await sleep(600);
  await scrollToLast(page, "[data-inquiry-result]");
  await shoot(page, PC_DIR, "22_result_refreshed");

  await newChat(page);
  await submit(page, "广东广州，2026年9月，供水管 PE管DN110，数量4600米");
  await resolveOutcome(page);
  await sleep(500);
  await page.click("#exportBtn");
  await sleep(900);
  await page.click("#stampBtn");
  await sleep(700);
  await shoot(page, PC_DIR, "13_preview_seal");
  await page.click("#stampBtn");
  await sleep(200);
  await shoot(page, PC_DIR, "13b_preview_unsealed");
  await sleep(2800);

  await page.click("#closePreviewBtn");
  await sleep(400);
  await page.click('[data-nav="history"]');
  await sleep(500);
  await shoot(page, PC_DIR, "14_history");

  await page.click(".history-item-more");
  await sleep(300);
  await page.click(".history-item-delete");
  await sleep(500);
  await shoot(page, PC_DIR, "15_delete_confirm");

  await page.close();
}

async function captureMobile(browser) {
  const page = await browser.newPage({
    viewport: { width: 390, height: 844 },
    deviceScaleFactor: 2,
  });
  await page.goto(DEMO);
  await sleep(600);
  await shoot(page, M_DIR, "01_home");

  await page.click("#openInquiryBtn");
  await sleep(500);
  await shoot(page, M_DIR, "02_workspace_empty");

  await submit(page, "广东广州，2026年9月，普通硅酸盐水泥 P.O42.5，数量200吨");
  await sleep(2100);
  await shoot(page, M_DIR, "03_thinking_steps");
  await sleep(RESULT_WAIT);
  await shoot(page, M_DIR, "04_result_cement");

  await newChat(page);
  await submit(page, "广东中山，2026年9月，PE直通管件PE110，数量650个");
  await sleep(RESULT_WAIT);
  await shoot(page, M_DIR, "05_result_insufficient");

  await newChat(page);
  await submit(page, "广东广州，2026年9月，螺纹钢 HRB400E Φ20，数量85吨");
  await sleep(RESULT_WAIT);
  await page.click("#exportBtn");
  await sleep(900);
  await shoot(page, M_DIR, "06_preview_seal");

  await page.click("#closePreviewBtn");
  await sleep(400);
  await newChat(page);
  await submit(page, "广东广州，2026年9月，特种调节阀 DN1200，数量2台");
  await sleep(RESULT_WAIT);
  await shoot(page, M_DIR, "07_no_data");

  await page.close();
}

(async () => {
  const only = process.argv[2] || "all";
  const browser = await chromium.launch({ executablePath: CHROME });
  if (only === "all" || only === "pc") await capturePc(browser);
  if (only === "all" || only === "m") await captureMobile(browser);
  await browser.close();
  console.log("capture done");
})().catch((error) => {
  console.error(error);
  process.exit(1);
});
