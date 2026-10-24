const fs = require("fs");
const path = require("path");
const { createCanvas, GlobalFonts } = require("@napi-rs/canvas");

const WIDTH = 2200;
const HEIGHT = 2700;
const canvas = createCanvas(WIDTH, HEIGHT);
const ctx = canvas.getContext("2d");

GlobalFonts.registerFromPath("/System/Library/Fonts/PingFang.ttc", "PingFang SC");

const palette = {
  background: "#F5F8FC",
  panel: "#FFFFFF",
  text: "#18314B",
  muted: "#5E7288",
  line: "#7894B2",
  blueFill: "#EAF3FF",
  blueStroke: "#2B70C2",
  amberFill: "#FFF6E7",
  amberStroke: "#D58A1B",
  greenFill: "#EAF8F0",
  greenStroke: "#409B69",
  redFill: "#FFF1F0",
  redStroke: "#D35C57",
  grayFill: "#F0F3F7",
  grayStroke: "#8497AA"
};

function roundedRect(x, y, w, h, radius) {
  ctx.beginPath();
  ctx.moveTo(x + radius, y);
  ctx.lineTo(x + w - radius, y);
  ctx.quadraticCurveTo(x + w, y, x + w, y + radius);
  ctx.lineTo(x + w, y + h - radius);
  ctx.quadraticCurveTo(x + w, y + h, x + w - radius, y + h);
  ctx.lineTo(x + radius, y + h);
  ctx.quadraticCurveTo(x, y + h, x, y + h - radius);
  ctx.lineTo(x, y + radius);
  ctx.quadraticCurveTo(x, y, x + radius, y);
  ctx.closePath();
}

function drawTextLines(lines, centerX, startY, lineHeight, font, fill) {
  ctx.font = font;
  ctx.fillStyle = fill;
  ctx.textAlign = "center";
  ctx.textBaseline = "middle";
  lines.forEach((line, index) => {
    ctx.fillText(line, centerX, startY + index * lineHeight);
  });
}

function drawNode(node) {
  const styleMap = {
    main: { fill: palette.blueFill, stroke: palette.blueStroke },
    decision: { fill: palette.amberFill, stroke: palette.amberStroke },
    success: { fill: palette.greenFill, stroke: palette.greenStroke },
    warning: { fill: palette.redFill, stroke: palette.redStroke },
    neutral: { fill: palette.grayFill, stroke: palette.grayStroke }
  };
  const style = styleMap[node.type] || styleMap.main;
  ctx.save();
  ctx.shadowColor = "rgba(38, 67, 96, 0.10)";
  ctx.shadowBlur = 10;
  ctx.shadowOffsetY = 4;
  roundedRect(node.x, node.y, node.w, node.h, 12);
  ctx.fillStyle = style.fill;
  ctx.fill();
  ctx.shadowColor = "transparent";
  ctx.strokeStyle = style.stroke;
  ctx.lineWidth = 2.2;
  ctx.stroke();
  ctx.restore();

  const titleLines = String(node.title || "").split("\n");
  const subtitleLines = String(node.subtitle || "").split("\n");
  const titleFont = `${node.bold ? "600" : "500"} ${node.titleSize || 27}px 'PingFang SC'`;
  const subtitleFont = `400 ${node.subtitleSize || 20}px 'PingFang SC'`;
  const titleLineHeight = (node.titleSize || 27) + 8;
  const subtitleLineHeight = (node.subtitleSize || 20) + 8;
  const gap = subtitleLines.length ? 10 : 0;
  const totalHeight = titleLines.length * titleLineHeight + gap + subtitleLines.length * subtitleLineHeight;
  const top = node.y + (node.h - totalHeight) / 2;
  drawTextLines(titleLines, node.x + node.w / 2, top + titleLineHeight / 2, titleLineHeight, titleFont, palette.text);
  if (subtitleLines.length) {
    drawTextLines(
      subtitleLines,
      node.x + node.w / 2,
      top + titleLines.length * titleLineHeight + gap + subtitleLineHeight / 2,
      subtitleLineHeight,
      subtitleFont,
      palette.muted
    );
  }
}

function drawArrowHead(x, y, angle, color = palette.line) {
  const size = 14;
  ctx.save();
  ctx.translate(x, y);
  ctx.rotate(angle);
  ctx.beginPath();
  ctx.moveTo(0, 0);
  ctx.lineTo(-size, -size * 0.55);
  ctx.lineTo(-size, size * 0.55);
  ctx.closePath();
  ctx.fillStyle = color;
  ctx.fill();
  ctx.restore();
}

function drawArrow(points, options = {}) {
  const color = options.color || palette.line;
  ctx.save();
  ctx.strokeStyle = color;
  ctx.lineWidth = options.width || 3.5;
  ctx.setLineDash(options.dash || []);
  ctx.lineJoin = "round";
  ctx.lineCap = "round";
  ctx.beginPath();
  ctx.moveTo(points[0].x, points[0].y);
  for (let index = 1; index < points.length; index += 1) {
    ctx.lineTo(points[index].x, points[index].y);
  }
  ctx.stroke();
  ctx.restore();
  const beforeLast = points[points.length - 2];
  const last = points[points.length - 1];
  drawArrowHead(last.x, last.y, Math.atan2(last.y - beforeLast.y, last.x - beforeLast.x), color);
}

function drawEdgeLabel(text, x, y, options = {}) {
  if (!text) return;
  ctx.font = "500 19px 'PingFang SC'";
  const width = ctx.measureText(text).width + 22;
  const height = 32;
  roundedRect(x - width / 2, y - height / 2, width, height, 7);
  ctx.fillStyle = options.fill || "#FFFFFF";
  ctx.fill();
  ctx.strokeStyle = options.stroke || "#D8E2EC";
  ctx.lineWidth = 1.5;
  ctx.stroke();
  ctx.fillStyle = options.color || palette.muted;
  ctx.textAlign = "center";
  ctx.textBaseline = "middle";
  ctx.fillText(text, x, y + 1);
}

function drawModuleFrame({ x, y, w, h, index, title, description, accent }) {
  ctx.save();
  ctx.shadowColor = "rgba(37, 68, 98, 0.10)";
  ctx.shadowBlur = 14;
  ctx.shadowOffsetY = 5;
  roundedRect(x, y, w, h, 18);
  ctx.fillStyle = palette.panel;
  ctx.fill();
  ctx.shadowColor = "transparent";
  ctx.strokeStyle = "#D7E2EE";
  ctx.lineWidth = 2;
  ctx.stroke();
  ctx.restore();

  roundedRect(x, y, 220, h, 18);
  ctx.fillStyle = accent.fill;
  ctx.fill();
  ctx.strokeStyle = accent.stroke;
  ctx.lineWidth = 2;
  ctx.stroke();

  ctx.strokeStyle = accent.stroke;
  ctx.lineWidth = 5;
  ctx.beginPath();
  ctx.moveTo(x + 218, y + 1);
  ctx.lineTo(x + 218, y + h - 1);
  ctx.stroke();

  ctx.font = "600 24px 'PingFang SC'";
  ctx.fillStyle = accent.stroke;
  ctx.textAlign = "left";
  ctx.textBaseline = "middle";
  ctx.fillText(`模块 ${index}`, x + 28, y + 48);

  const titleLines = String(title).split("\n");
  ctx.font = "600 31px 'PingFang SC'";
  ctx.fillStyle = palette.text;
  titleLines.forEach((line, lineIndex) => {
    ctx.fillText(line, x + 28, y + 106 + lineIndex * 39);
  });

  ctx.font = "400 19px 'PingFang SC'";
  ctx.fillStyle = palette.muted;
  const descLines = String(description).split("\n");
  descLines.forEach((line, lineIndex) => {
    ctx.fillText(line, x + 28, y + 173 + lineIndex * 28);
  });
}

function centerTop(node) {
  return { x: node.x + node.w / 2, y: node.y };
}

function centerBottom(node) {
  return { x: node.x + node.w / 2, y: node.y + node.h };
}

function centerLeft(node) {
  return { x: node.x, y: node.y + node.h / 2 };
}

function centerRight(node) {
  return { x: node.x + node.w, y: node.y + node.h / 2 };
}

function horizontalChain(nodes, labels = []) {
  for (let index = 0; index < nodes.length - 1; index += 1) {
    const from = centerRight(nodes[index]);
    const to = centerLeft(nodes[index + 1]);
    drawArrow([from, to]);
    if (labels[index]) {
      drawEdgeLabel(labels[index], (from.x + to.x) / 2, from.y - 28);
    }
  }
}

const modules = [
  { x: 90, y: 150, w: 2020, h: 390 },
  { x: 90, y: 580, w: 2020, h: 390 },
  { x: 90, y: 1010, w: 2020, h: 590 },
  { x: 90, y: 1640, w: 2020, h: 390 },
  { x: 90, y: 2070, w: 2020, h: 470 }
];

const accents = [
  { fill: "#EEF5FF", stroke: "#2B70C2" },
  { fill: "#FFF7E8", stroke: "#D58A1B" },
  { fill: "#EEF8F2", stroke: "#409B69" },
  { fill: "#EEF5FF", stroke: "#2B70C2" },
  { fill: "#FFF1F0", stroke: "#D35C57" }
];

const nodes = {
  input: {
    x: 350, y: 245, w: 360, h: 130,
    title: "输入或覆盖询价条件",
    subtitle: "变更字段覆盖旧值\n生成新 requestId",
    type: "main",
    bold: true
  },
  intent: {
    x: 785, y: 245, w: 360, h: 130,
    title: "意图识别",
    subtitle: "判断是否为材料询价需求",
    type: "decision",
    bold: true
  },
  standardize: {
    x: 1220, y: 245, w: 360, h: 130,
    title: "材料标准化",
    subtitle: "生成 cid、subcid\nstdname、features",
    type: "main",
    bold: true
  },
  specCheck: {
    x: 1655, y: 245, w: 365, h: 130,
    title: "规格完整性校验",
    subtitle: "不足时补充并确认",
    type: "decision",
    bold: true
  },
  invalid: {
    x: 785, y: 425, w: 360, h: 82,
    title: "非材料询价：提示补充条件",
    subtitle: "不进入查询、不扣次数",
    type: "warning",
    titleSize: 22,
    subtitleSize: 18
  },
  supplement: {
    x: 1655, y: 425, w: 365, h: 82,
    title: "缺规格：补充并确认",
    subtitle: "确认后回到材料标准化",
    type: "warning",
    titleSize: 22,
    subtitleSize: 18
  },

  history: {
    x: 350, y: 675, w: 430, h: 130,
    title: "读取历史询价候选",
    subtitle: "核心时间 → 材料名称 → 询价 ID\n推荐前 3 条",
    type: "main",
    bold: true
  },
  historyCheck: {
    x: 865, y: 675, w: 350, h: 130,
    title: "历史询价有结果？",
    subtitle: "命中则作为价格基础",
    type: "decision",
    bold: true
  },
  market: {
    x: 1300, y: 675, w: 350, h: 130,
    title: "回退造价通市场价",
    subtitle: "历史无结果时自动切换",
    type: "warning",
    bold: true
  },
  priceBase: {
    x: 1735, y: 675, w: 285, h: 130,
    title: "确定价格基础",
    subtitle: "历史价或市场价",
    type: "success",
    bold: true
  },
  noData: {
    x: 1300, y: 850, w: 350, h: 78,
    title: "无可用数据：调整时间/区域/转人工",
    subtitle: "",
    type: "warning",
    titleSize: 20
  },

  mapping: {
    x: 350, y: 1135, w: 420, h: 150,
    title: "五级优先级匹配",
    subtitle: "询价映射表\n命中即终止",
    type: "main",
    bold: true
  },
  coefficientName: {
    x: 865, y: 1135, w: 350, h: 150,
    title: "得到系数名称",
    subtitle: "t_coefficient.name",
    type: "success",
    bold: true
  },
  coefficientLookup: {
    x: 1300, y: 1135, w: 350, h: 150,
    title: "匹配具体 COEF",
    subtitle: "地区、城市、月份\n匹配系数表",
    type: "decision",
    bold: true
  },
  calculatedCoefficient: {
    x: 1735, y: 1135, w: 285, h: 150,
    title: "取得系数",
    subtitle: "进入价格计算",
    type: "success",
    bold: true
  },
  rule1: {
    x: 350, y: 1340, w: 300, h: 82,
    title: "规则 1",
    subtitle: "stdname + features 包含",
    type: "neutral",
    titleSize: 23,
    subtitleSize: 17
  },
  rule2: {
    x: 690, y: 1340, w: 300, h: 82,
    title: "规则 2",
    subtitle: "stdname + features 为空",
    type: "neutral",
    titleSize: 23,
    subtitleSize: 17
  },
  rule3: {
    x: 1030, y: 1340, w: 300, h: 82,
    title: "规则 3",
    subtitle: "subcid + 其余为空",
    type: "neutral",
    titleSize: 23,
    subtitleSize: 17
  },
  rule4: {
    x: 1370, y: 1340, w: 300, h: 82,
    title: "规则 4",
    subtitle: "cid + 其余为空",
    type: "neutral",
    titleSize: 23,
    subtitleSize: 17
  },
  rule5: {
    x: 1710, y: 1340, w: 310, h: 82,
    title: "规则 5",
    subtitle: "cid = 32 默认系数",
    type: "warning",
    titleSize: 23,
    subtitleSize: 17
  },
  pendingRule: {
    x: 1300, y: 1465, w: 720, h: 80,
    title: "系数管理后台：地区、城市、月份维度系数值由数据部门维护并下发",
    subtitle: "",
    type: "neutral",
    titleSize: 21
  },

  calculate: {
    x: 350, y: 1775, w: 430, h: 150,
    title: "计算建议价",
    subtitle: "历史价 × t_coefficient\n或市场价 × t_coefficient",
    type: "success",
    bold: true
  },
  result: {
    x: 865, y: 1775, w: 450, h: 150,
    title: "输出询价建议",
    subtitle: "材料、规格、建议价\n供应商、报价备注",
    type: "main",
    bold: true
  },
  actions: {
    x: 1400, y: 1775, w: 620, h: 150,
    title: "结果操作",
    subtitle: "复制、点赞、点踩、重新生成\n导出盖章报价",
    type: "main",
    bold: true
  },

  phaseNote: {
    x: 350, y: 1948, w: 960, h: 74,
    title: "项目阶段提示后置：结果生成后采集，未选择默认施工图预算，选择后按阶段系数更新建议价",
    subtitle: "",
    type: "neutral",
    titleSize: 20
  },

  contact: {
    x: 350, y: 2205, w: 430, h: 125,
    title: "追问供应商联系方式",
    subtitle: "返回联系人和电话",
    type: "success",
    bold: true
  },
  contactIssue: {
    x: 865, y: 2205, w: 450, h: 125,
    title: "联系方式不可用",
    subtitle: "提示记录较久或换号\n进入转人工入口",
    type: "warning",
    bold: true
  },
  priceSource: {
    x: 1400, y: 2205, w: 290, h: 125,
    title: "追问价格来源",
    subtitle: "默认不回复",
    type: "neutral",
    bold: true
  },
  quota: {
    x: 1775, y: 2205, w: 245, h: 125,
    title: "次数管理",
    subtitle: "AI 询价 10 条/天\nAI 查价 50 条/天",
    type: "neutral",
    titleSize: 22,
    bold: true
  },
  manual: {
    x: 865, y: 2385, w: 450, h: 105,
    title: "跳转造价通人工询价页面",
    subtitle: "页面承接登录、联系方式\n与询价要求填写",
    type: "success",
    bold: true,
    titleSize: 24,
    subtitleSize: 18
  }
};

ctx.fillStyle = palette.background;
ctx.fillRect(0, 0, WIDTH, HEIGHT);

ctx.font = "600 48px 'PingFang SC'";
ctx.fillStyle = palette.text;
ctx.textAlign = "left";
ctx.textBaseline = "middle";
ctx.fillText("AI 询价功能模块流程", 90, 62);
ctx.font = "400 22px 'PingFang SC'";
ctx.fillStyle = palette.muted;
ctx.fillText("按模块拆分：输入与标准化 → 数据选择 → 系数匹配 → 建议价与结果 → 操作与异常", 92, 108);

drawModuleFrame({
  ...modules[0],
  index: 1,
  title: "输入与\n标准化",
  description: "条件覆盖与二次询价\n沉淀标准材料字段",
  accent: accents[0]
});
drawModuleFrame({
  ...modules[1],
  index: 2,
  title: "数据选择与\n候选推荐",
  description: "历史询价优先\n无结果切换市场价",
  accent: accents[1]
});
drawModuleFrame({
  ...modules[2],
  index: 3,
  title: "系数\n匹配",
  description: "映射表五级优先\n命中即终止",
  accent: accents[2]
});
drawModuleFrame({
  ...modules[3],
  index: 4,
  title: "建议价与\n结果",
  description: "价格转换\n结果后采集阶段",
  accent: accents[3]
});
drawModuleFrame({
  ...modules[4],
  index: 5,
  title: "用户操作与\n异常处理",
  description: "结果反馈、盖章撤销\n和人工询价跳转",
  accent: accents[4]
});

Object.values(nodes).forEach(drawNode);

horizontalChain([nodes.input, nodes.intent, nodes.standardize, nodes.specCheck], ["有效", "成功", "校验"]);
drawArrow([centerBottom(nodes.intent), { x: centerBottom(nodes.intent).x, y: 430 }, centerTop(nodes.invalid)], { color: palette.redStroke });
drawEdgeLabel("无效", centerBottom(nodes.intent).x + 46, 405, { color: palette.redStroke, stroke: "#EDC5C2" });
drawArrow([centerBottom(nodes.specCheck), { x: centerBottom(nodes.specCheck).x, y: 430 }, centerTop(nodes.supplement)], { color: palette.amberStroke });
drawEdgeLabel("不完整", centerBottom(nodes.specCheck).x + 56, 405, { color: palette.amberStroke, stroke: "#E7C58F" });

horizontalChain([nodes.history, nodes.historyCheck, nodes.market, nodes.priceBase], ["候选", "无结果", "命中"]);
drawArrow([centerBottom(nodes.market), { x: centerBottom(nodes.market).x, y: 860 }, centerTop(nodes.noData)], { color: palette.redStroke, dash: [8, 6] });
drawEdgeLabel("无数据", centerBottom(nodes.market).x - 46, 842, { color: palette.redStroke, stroke: "#EDC5C2" });

horizontalChain([nodes.mapping, nodes.coefficientName, nodes.coefficientLookup, nodes.calculatedCoefficient], ["命中", "系数名称", "命中"]);
drawArrow([centerBottom(nodes.mapping), { x: centerBottom(nodes.mapping).x, y: 1328 }, { x: nodes.rule1.x, y: 1328 }, { x: nodes.rule1.x, y: centerTop(nodes.rule1).y - 2 }], { color: palette.greenStroke, dash: [8, 6] });
drawArrow([centerBottom(nodes.calculatedCoefficient), { x: centerBottom(nodes.calculatedCoefficient).x, y: 1625 }, { x: nodes.result.x, y: 1625 }, { x: nodes.result.x, y: centerTop(nodes.result).y - 2 }], { color: palette.greenStroke });
drawEdgeLabel("取得系数", centerBottom(nodes.calculatedCoefficient).x - 78, 1610, { color: "#2F7D58", stroke: "#B6DEC8" });

horizontalChain([nodes.calculate, nodes.result, nodes.actions], ["公式计算", "生成结果"]);
drawArrow([centerBottom(nodes.result), centerTop(nodes.phaseNote)], { color: palette.greenStroke });
drawArrow([{ x: 1420, y: nodes.actions.y + nodes.actions.h }, { x: 1420, y: 2130 }, centerTop(nodes.contact)], { color: palette.redStroke, dash: [8, 6] });
drawEdgeLabel("进入后续分支", 1440, 2085, { color: palette.redStroke, stroke: "#EDC5C2" });

drawArrow([centerRight(nodes.calculate), { x: 835, y: centerRight(nodes.calculate).y }, centerLeft(nodes.result)]);
drawArrow([centerRight(nodes.result), centerLeft(nodes.actions)]);

drawArrow([centerRight(nodes.contact), centerLeft(nodes.contactIssue)], { color: palette.redStroke });
drawEdgeLabel("不可用", 820, centerRight(nodes.contact).y - 28, { color: palette.redStroke, stroke: "#EDC5C2" });
drawArrow([centerBottom(nodes.contactIssue), { x: centerBottom(nodes.contactIssue).x, y: 2370 }, centerTop(nodes.manual)], { color: palette.redStroke });

drawArrow([{ x: 1545, y: nodes.actions.y + nodes.actions.h }, centerTop(nodes.priceSource)], { color: palette.grayStroke, dash: [8, 6] });
drawEdgeLabel("追问价格来源", 1560, 2100);
drawArrow([centerRight(nodes.priceSource), centerLeft(nodes.quota)], { color: palette.grayStroke });
drawEdgeLabel("次数", 1718, centerRight(nodes.priceSource).y - 25);

roundedRect(90, 2570, 2020, 96, 12);
ctx.fillStyle = "#FFF8EC";
ctx.fill();
ctx.strokeStyle = "#E5B566";
ctx.lineWidth = 2;
ctx.stroke();
ctx.font = "500 19px 'PingFang SC'";
ctx.fillStyle = "#A75B12";
ctx.textAlign = "left";
ctx.textBaseline = "middle";
ctx.fillText("说明：建议价 = 价格基础（历史询价优先，无结果取造价通市场价）× 时间系数 × 项目阶段系数。", 125, 2602);
ctx.fillText("项目阶段在询价建议结果生成后采集，未选择时默认施工图预算；系数映射表与系数值由数据部门在管理后台维护并下发。", 125, 2638);

const outputPaths = [
  path.join(__dirname, "ai询价业务流程图.png"),
  path.join(__dirname, "generated/lanhu_demo_based_big_sheets/00_AI询价业务流程图.png")
];
const buffer = canvas.toBuffer("image/png");
for (const outputPath of outputPaths) {
  fs.mkdirSync(path.dirname(outputPath), { recursive: true });
  fs.writeFileSync(outputPath, buffer);
  console.log(outputPath);
}
