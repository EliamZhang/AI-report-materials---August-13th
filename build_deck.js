// 银行流水汇报 PPT 重设计 —— Fundo 风格（白底 + 蓝色主色 + 卡片化）
// 生成 银行流水汇报_v2.pptx（LAYOUT_WIDE 13.333x7.5）
const pptxgen = require('C:/Users/zhangyuliang02/AppData/Roaming/npm/node_modules/ppt-mcp/node_modules/pptxgenjs');
const fs = require('fs');

const SW = 13.333;
const SH = 7.5;
const SAFE = { left: 0.40, right: 0.40, top: 0.32, bottom: 0.28 };
const CONTENT_W = SW - SAFE.left - SAFE.right;
const CONTENT_H = SH - SAFE.top - SAFE.bottom;

const C = {
  blue: '0878E3',
  blueLight: 'EAF4FF',
  blueDark: '0B3D91',
  black: '202124',
  gray: '858A91',
  border: 'DFE3E8',
  surface: 'F7F8FA',
  red: 'E60012',
  redBg: 'FDEBEC',
  dark: '202020',
  white: 'FFFFFF',
};
const FONT = 'Microsoft YaHei';
const FONT_EN = 'Arial';

// ---------- 边界检查 ----------
function assertInBounds(name, x, y, w, h, { allowBleed = false } = {}) {
  const eps = 0.01;
  const minX = allowBleed ? 0 : SAFE.left;
  const minY = allowBleed ? 0 : SAFE.top;
  const maxX = allowBleed ? SW : SW - SAFE.right;
  const maxY = allowBleed ? SH : SH - SAFE.bottom;
  if (![x, y, w, h].every(Number.isFinite) || w <= 0 || h <= 0) {
    throw new Error(`${name}: invalid geometry x=${x}, y=${y}, w=${w}, h=${h}`);
  }
  if (x < minX - eps || y < minY - eps || x + w > maxX + eps || y + h > maxY + eps) {
    throw new Error(`${name}: outside safe area - x=${x}, y=${y}, w=${w}, h=${h}`);
  }
}

// ---------- 通用组件 ----------
function addHeader(slide, num, title, subtitle) {
  const y = 0.32;
  const numStr = String(num).padStart(2, '0');
  assertInBounds(`header-num-${num}`, SAFE.left, y, 0.55, 0.5);
  slide.addText(numStr, {
    x: SAFE.left, y, w: 0.6, h: 0.5, fontFace: FONT_EN, fontSize: 28, bold: true, color: C.blue, align: 'left', margin: 0,
  });
  let tx = SAFE.left + 0.72;
  const tw = CONTENT_W - 0.72 - 0.02;
  assertInBounds(`header-title-${num}`, tx, y, tw, 0.6);
  slide.addText(title, {
    x: tx, y, w: tw, h: 0.6, fontFace: FONT, fontSize: 26, bold: true, color: C.black, margin: 0,
  });
  if (subtitle) {
    const sy = 0.95;
    slide.addText(subtitle, {
      x: SAFE.left, y: sy, w: CONTENT_W, h: 0.36, fontFace: FONT, fontSize: 12.5, color: C.gray, margin: 0,
    });
    return sy + 0.36 + 0.12;
  }
  return y + 0.6 + 0.14;
}

function addFooter(slide, pageNo) {
  slide.addText('AI 银行流水识别 · 2026 年 8 月', {
    x: SAFE.left, y: SH - 0.24, w: 4, h: 0.2, fontFace: FONT, fontSize: 8.5, color: C.gray, margin: 0,
  });
  slide.addText(String(pageNo).padStart(2, '0'), {
    x: SW - SAFE.right - 0.5, y: SH - 0.24, w: 0.5, h: 0.2, fontFace: FONT_EN, fontSize: 8.5, color: C.gray, align: 'right', margin: 0,
  });
}

function addCard(slide, x, y, w, h, { fill = C.white, border = C.border, radius = 0.06 } = {}) {
  assertInBounds(`card`, x, y, w, h);
  return slide.addShape('roundRect', {
    x, y, w, h, fill: { color: fill }, line: { color: border, width: 1 }, rectRadius: radius,
  });
}

function addTag(slide, x, y, w, text, { fill = C.blueLight, color = C.blue, bold = true, fontSize = 9 } = {}) {
  const h = 0.26;
  assertInBounds(`tag`, x, y, w, h);
  slide.addShape('roundRect', {
    x, y, w, h, fill: { color: fill }, line: { color: fill, width: 0 }, rectRadius: h / 2,
  });
  slide.addText(text, {
    x, y, w, h, fontFace: FONT, fontSize, bold, color, align: 'center', margin: 0, valign: 'middle',
  });
}

function addNumBadge(slide, x, y, n, { color = C.blue, bg = C.blueLight } = {}) {
  const d = 0.34;
  assertInBounds(`badge-${n}`, x, y, d, d);
  slide.addShape('ellipse', { x, y, w: d, h: d, fill: { color: bg }, line: { color: bg, width: 0 } });
  slide.addText(String(n), {
    x, y, w: d, h: d, fontFace: FONT_EN, fontSize: 13, bold: true, color, align: 'center', valign: 'middle', margin: 0,
  });
}

function addLeftRule(slide, x, y, w, color = C.blue) {
  slide.addShape('rect', { x, y, w, h: 0.05, fill: { color }, line: { color, width: 0 } });
}

// 表格行辅助：把二维数组转成 pptxgenjs 单元格格式，rowOpts 为每行 {text, options} 样式
function tableRows(rows, rowOpts) {
  return rows.map((r, i) => r.map((cell) => ({ text: cell, options: rowOpts[i] || {} })));
}

function txt(opt) {
  const { text, ...rest } = opt;
  return [{ text, options: Object.assign({ fontFace: FONT }, rest) }];
}

// ============================================================
const pptx = new pptxgen();
pptx.layout = 'LAYOUT_WIDE';
pptx.author = 'AI Report';
pptx.title = 'AI 银行流水识别方案';

// ---------------- Slide 1 封面 ----------------
{
  const s = pptx.addSlide();
  // 全出血深蓝背景
  s.addShape('rect', { x: 0, y: 0, w: SW, h: SH, fill: { color: '0B2E59' }, line: { color: '0B2E59', width: 0 } });
  // 顶部细线
  s.addShape('rect', { x: 0, y: 0, w: SW, h: 0.06, fill: { color: C.blue }, line: { color: C.blue, width: 0 } });
  // 左侧竖条装饰
  s.addShape('rect', { x: 0, y: 0, w: 0.12, h: SH, fill: { color: '0B3D91' }, line: { color: '0B3D91', width: 0 } });

  const cx = 0.9;
  s.addText('AI · BANK STATEMENT ANALYSIS', {
    x: cx, y: 1.35, w: 8, h: 0.32, fontFace: FONT_EN, fontSize: 12, bold: true, color: '7FB3F5', charSpacing: 2, margin: 0,
  });
  s.addText('银行流水识别与 AI 分类', {
    x: cx, y: 1.75, w: 11.5, h: 1.0, fontFace: FONT, fontSize: 40, bold: true, color: C.white, margin: 0,
  });
  s.addText('基于交易分类、商户知识库与 AI 质量闭环的银行流水加工方案', {
    x: cx, y: 2.85, w: 11.5, h: 0.5, fontFace: FONT, fontSize: 16, color: 'B8CCE8', margin: 0,
  });
  s.addText('2026 年 8 月', {
    x: cx, y: 3.5, w: 4, h: 0.35, fontFace: FONT, fontSize: 13, color: '8FA9CC', margin: 0,
  });

  // 底部能力标签
  const tags = ['7 大分类引擎', '商户知识库', 'AI 质量闭环', 'Liability 识别'];
  const tw = 1.9;
  const gap = 0.22;
  let tx0 = cx;
  tags.forEach((t, i) => {
    s.addShape('roundRect', {
      x: tx0, y: 5.2, w: tw, h: 0.44, fill: { color: '12407C' }, line: { color: '2C5FA3', width: 1 }, rectRadius: 0.22,
    });
    s.addText(t, {
      x: tx0, y: 5.2, w: tw, h: 0.44, fontFace: FONT, fontSize: 11.5, color: 'DCE8F8', align: 'center', valign: 'middle', margin: 0,
    });
    tx0 += tw + gap;
  });
}

// ---------------- Slide 2 识别解决什么问题 ----------------
{
  const s = pptx.addSlide();
  let y0 = addHeader(s, 1, '银行流水识别解决什么问题', '原始流水 → 结构化数据 → 收入评估、负债识别与偿付能力计算，支撑信审决策');

  // 流程链
  const steps = ['用户申请', '银行卡绑定', '数据解析', '交易对手与分类识别', 'Serviceability 计算', '信审决策'];
  const n = steps.length;
  const boxW = 1.78;
  const gap = 0.14;
  const totalW = n * boxW + (n - 1) * gap;
  const x0 = SAFE.left + (CONTENT_W - totalW) / 2;
  steps.forEach((st, i) => {
    const x = x0 + i * (boxW + gap);
    const y = y0 + 0.06;
    addCard(s, x, y, boxW, 0.5, { border: i === n - 1 ? C.blue : C.border });
    s.addText(txt({ text: st, x, y, w: boxW, h: 0.5, fontSize: 11, bold: i === n - 1, color: i === n - 1 ? C.blue : C.black, align: 'center', valign: 'middle', margin: 0 }));
    if (i < n - 1) {
      s.addShape('rect', { x: x + boxW + 0.02, y: y + 0.24, w: 0.10, h: 0.02, fill: { color: C.gray }, line: { color: C.gray, width: 0 } });
    }
  });

  let y1 = y0 + 0.72;
  // 左：公式卡
  addCard(s, SAFE.left, y1, 4.4, 1.62, { fill: C.blueLight, border: C.blue });
  addLeftRule(s, SAFE.left + 0.22, y1 + 0.2, 0.5, C.blue);
  s.addText(txt({ text: 'Serviceability 核心公式', x: SAFE.left + 0.22, y: y1 + 0.3, w: 3.9, h: 0.3, fontSize: 13, bold: true, color: C.black, margin: 0 }));
  s.addText(txt({ text: '收入 − 必要支出 − 负债还款 = 可支配盈余', x: SAFE.left + 0.22, y: y1 + 0.72, w: 4.0, h: 0.4, fontSize: 13.5, bold: true, color: C.blue, margin: 0 }));
  s.addText(txt({ text: '流水识别准确性直接影响收入、支出与负债的计算结果，进而影响最终信审判断。', x: SAFE.left + 0.22, y: y1 + 1.18, w: 4.0, h: 0.36, fontSize: 10.5, color: C.gray, margin: 0 }));

  // 右：监管依据
  addCard(s, SAFE.left + 4.65, y1, 7.86, 1.62);
  s.addText(txt({ text: '监管依据：Responsible Lending', x: SAFE.left + 4.85, y: y1 + 0.16, w: 5, h: 0.3, fontSize: 13, bold: true, color: C.black, margin: 0 }));
  s.addText(txt({ text: '《National Consumer Credit Protection Act 2009》与 ASIC 监管要求：', x: SAFE.left + 4.85, y: y1 + 0.5, w: 7.5, h: 0.28, fontSize: 10, color: C.gray, margin: 0 }));
  const regs = ['合理了解客户财务状况与借款目的', '核实收入、必要支出、现有负债及可预见变化', '判断能否在不造成重大经济困难下履行还款', '不得提供"不适合"的信贷产品'];
  regs.forEach((r, i) => {
    const rx = SAFE.left + 4.85 + (i % 2) * 3.8;
    const ry = y1 + 0.86 + Math.floor(i / 2) * 0.34;
    s.addShape('ellipse', { x: rx, y: ry + 0.07, w: 0.12, h: 0.12, fill: { color: C.blue }, line: { color: C.blue, width: 0 } });
    s.addText(txt({ text: r, x: rx + 0.18, y: ry, w: 3.6, h: 0.28, fontSize: 10, color: C.black, margin: 0 }));
  });

  let y2 = y1 + 1.8;
  const cards3 = [
    { t: '交易对手与分类识别', d: '识别实际交易对手、机构属性与交易类别，建立结构化识别结果。' },
    { t: 'Serviceability 计算', d: '基于识别结果判断客户能否在重大经济压力下维持还款能力。' },
    { t: '信贷决策支持', d: '识别准确性直接决定收入、支出与负债核算，影响最终信审判断。' },
  ];
  const cw = (CONTENT_W - 2 * 0.2) / 3;
  cards3.forEach((c, i) => {
    const x = SAFE.left + i * (cw + 0.2);
    addCard(s, x, y2, cw, 1.06);
    addNumBadge(s, x + 0.2, y2 + 0.22, i + 1);
    s.addText(txt({ text: c.t, x: x + 0.7, y: y2 + 0.18, w: cw - 0.9, h: 0.3, fontSize: 13, bold: true, color: C.black, margin: 0 }));
    s.addText(txt({ text: c.d, x: x + 0.22, y: y2 + 0.58, w: cw - 0.44, h: 0.42, fontSize: 10, color: C.gray, margin: 0 }));
  });
}

// ---------------- Slide 3 自建能力的 9 个痛点 ----------------
{
  const s = pptx.addSlide();
  let y0 = addHeader(s, 2, '为什么自建分类能力', '当前交易分类与对手识别高度依赖 illion，存在以下痛点');
  const items = [
    { t: '分类逻辑不透明', d: '规则与模型是黑盒，漏识别或误分类时难以定位具体原因。' },
    { t: '无法快速调整', d: '规则与口径调整依赖外部供应商，难以快速响应业务与合规变化。' },
    { t: '分类粒度不足', d: 'SACC / Non-SACC / BNPL / LOC / Wage Advance 等类型需进一步拆分。' },
    { t: '长尾商户覆盖不足', d: '小型商户、新兴平台与特殊描述易落入 Other 或错误分类。' },
    { t: '对手识别能力有限', d: '无法准确识别具体商户或贷款机构，下游归属计算难以开展。' },
    { t: '字段与口径缺控制权', d: 'IDP 与 Serviceability 所需字段及公式、窗口无法按内部标准配置。' },
    { t: '效果难以系统监控', d: '无法掌握模型版本与知识库变化，难以建立效果评估与回溯机制。' },
    { t: '多数据源不一致', d: 'Open Banking 接入后，多供应商字段与分类体系差异增加复杂度。' },
    { t: '受制于外部供应商', d: '长期依赖外部限制模型优化、策略调整与产品创新的自主性。' },
  ];
  const cols = 3;
  const gap = 0.18;
  const cw = (CONTENT_W - gap * (cols - 1)) / cols;
  const ch = 1.24;
  const rows = 3;
  const rgap = 0.16;
  let y = y0 + 0.02;
  items.forEach((it, i) => {
    const col = i % cols;
    const row = Math.floor(i / cols);
    const x = SAFE.left + col * (cw + gap);
    const yy = y + row * (ch + rgap);
    addCard(s, x, yy, cw, ch);
    addNumBadge(s, x + 0.18, yy + 0.16, i + 1);
    s.addText(txt({ text: it.t, x: x + 0.62, y: yy + 0.17, w: cw - 0.8, h: 0.28, fontSize: 12.5, bold: true, color: C.black, margin: 0 }));
    s.addText(txt({ text: it.d, x: x + 0.2, y: yy + 0.56, w: cw - 0.4, h: 0.62, fontSize: 9.8, color: C.gray, margin: 0 }));
  });

  const by = y + rows * (ch + rgap) - rgap + 0.1;
  addCard(s, SAFE.left, by, CONTENT_W, 0.62, { fill: C.blueLight, border: C.blue });
  s.addText(txt({
    text: '结论：需要逐步建设自主的交易分类模型与商户知识库，将对手识别、分类规则与计算逻辑沉淀为内部能力，提升可解释性、可配置性与迭代效率。',
    x: SAFE.left + 0.22, y: by + 0.1, w: CONTENT_W - 0.44, h: 0.42, fontSize: 11.5, bold: true, color: C.blue, margin: 0,
  }));
}

// ---------------- Slide 4 需要完成哪些加工 ----------------
{
  const s = pptx.addSlide();
  let y0 = addHeader(s, 3, '银行流水识别需要完成哪些加工', '两个维度：交易明细识别 + 用户汇总指标');
  const cardH = 5.02;

  // 左侧：交易明细维度
  addCard(s, SAFE.left, y0, 6.2, cardH);
  s.addText(txt({ text: '维度 1 · 交易明细', x: SAFE.left + 0.22, y: y0 + 0.14, w: 4, h: 0.3, fontSize: 14, bold: true, color: C.black, margin: 0 }));
  s.addText(txt({ text: '原始交易文本 → 结构化信息', x: SAFE.left + 0.22, y: y0 + 0.46, w: 4, h: 0.24, fontSize: 10, color: C.gray, margin: 0 }));

  // 示例卡（原始）
  const exY = y0 + 0.8;
  addCard(s, SAFE.left + 0.22, exY, 5.76, 0.92, { fill: C.surface, border: C.border });
  s.addText(txt({ text: '原始交易示例', x: SAFE.left + 0.4, y: exY + 0.08, w: 2, h: 0.24, fontSize: 9.5, bold: true, color: C.gray, margin: 0 }));
  const raw = '2026-03-26 · 156 澳元 · Debit\nBILL PAY Fair Go Finance DT.58oa72 SACC 239';
  s.addText(txt({ text: raw, x: SAFE.left + 0.4, y: exY + 0.34, w: 5.4, h: 0.5, fontSize: 10, fontFace: FONT_EN, color: C.black, margin: 0 }));

  // A 交易对手识别
  let ax = SAFE.left + 0.22;
  let ay = exY + 1.06;
  addNumBadge(s, ax, ay + 0.02, 1, { color: C.blue, bg: C.blueLight });
  s.addText(txt({ text: 'A · 交易对手识别', x: ax + 0.42, y: ay - 0.02, w: 3, h: 0.28, fontSize: 12, bold: true, color: C.black, margin: 0 }));
  ay += 0.32;
  const oppRows = [
    ['对手名称', 'Fair Go Finance'],
    ['主体类型', '贷款机构'],
    ['机构属性', '小额信贷机构（SACC）'],
  ];
  oppRows.forEach((r, i) => {
    const ry = ay + i * 0.32;
    s.addText(txt({ text: r[0], x: ax + 0.05, y: ry, w: 1.0, h: 0.26, fontSize: 9.8, bold: true, color: C.gray, margin: 0 }));
    s.addText(txt({ text: r[1], x: ax + 1.1, y: ry, w: 3.4, h: 0.26, fontSize: 10, color: C.black, margin: 0 }));
  });
  ay += 3 * 0.32 + 0.08;
  s.addText(txt({ text: '常见对手：商户企业（Woolworths / Telstra）、贷款机构（Fair Go Finance）、政府机构（Services Australia / ATO）、个人账户。', x: ax, y: ay, w: 5.8, h: 0.32, fontSize: 9, color: C.gray, margin: 0 }));

  // B 交易分类识别
  let bx = SAFE.left + 0.22;
  let byy = ay + 0.42;
  addNumBadge(s, bx, byy + 0.02, 2, { color: C.blue, bg: C.blueLight });
  s.addText(txt({ text: 'B · 交易分类识别', x: bx + 0.42, y: byy - 0.02, w: 3, h: 0.28, fontSize: 12, bold: true, color: C.black, margin: 0 }));
  byy += 0.34;
  addCard(s, bx + 0.05, byy, 2.6, 0.5, { fill: C.blueLight, border: C.blue });
  s.addText(txt({ text: '一级分类：贷款及还款', x: bx + 0.15, y: byy + 0.06, w: 2.4, h: 0.38, fontSize: 10, bold: true, color: C.blue, align: 'center', valign: 'middle', margin: 0 }));
  addCard(s, bx + 2.8, byy, 2.6, 0.5, { fill: C.blueLight, border: C.blue });
  s.addText(txt({ text: '二级分类：SACC Loans', x: bx + 2.9, y: byy + 0.06, w: 2.4, h: 0.38, fontSize: 10, bold: true, color: C.blue, align: 'center', valign: 'middle', margin: 0 }));
  byy += 0.6;
  s.addText(txt({ text: '常见分类：工资收入 · 赌博支出 · 房租 · 水电通信 · 内部转账 · 外部个人转账', x: bx + 0.05, y: byy, w: 5.9, h: 0.28, fontSize: 9, color: C.gray, margin: 0 }));

  // 右侧：用户汇总维度
  addCard(s, SAFE.left + 6.4, y0, 6.1, cardH);
  s.addText(txt({ text: '维度 2 · 用户汇总', x: SAFE.left + 6.62, y: y0 + 0.14, w: 4, h: 0.3, fontSize: 14, bold: true, color: C.black, margin: 0 }));
  s.addText(txt({ text: '按客户、账户与观察周期汇总，形成用户级负债与还款指标', x: SAFE.left + 6.62, y: y0 + 0.46, w: 5.7, h: 0.24, fontSize: 10, color: C.gray, margin: 0 }));

  let uy = y0 + 0.8;
  const agg1 = [
    ['Fair Go Finance', '每两周还款 156 澳元'],
    ['Credit Corp', '每月还款 90 澳元'],
    ['其他贷款还款', '78 澳元'],
  ];
  agg1.forEach((r, i) => {
    const ry = uy + i * 0.38;
    s.addShape('ellipse', { x: SAFE.left + 6.66, y: ry + 0.08, w: 0.12, h: 0.12, fill: { color: C.blue }, line: { color: C.blue, width: 0 } });
    s.addText(txt({ text: r[0], x: SAFE.left + 6.88, y: ry, w: 2.4, h: 0.3, fontSize: 10.5, bold: true, color: C.black, margin: 0 }));
    s.addText(txt({ text: r[1], x: SAFE.left + 9.4, y: ry, w: 3.0, h: 0.3, fontSize: 10.5, color: C.gray, align: 'right', margin: 0 }));
  });
  uy += 3 * 0.38 + 0.14;
  s.addText(txt({ text: '汇总输出（最近 90 天）', x: SAFE.left + 6.62, y: uy, w: 3, h: 0.28, fontSize: 11, bold: true, color: C.black, margin: 0 }));
  uy += 0.34;
  const agg2 = [
    ['当前贷款机构数', '2 家'],
    ['借款总额', '2,300 澳元'],
    ['还款总额', '1,284 澳元'],
    ['平均月还款', '428 澳元'],
    ['预计待还金额', '1,016 澳元'],
    ['还款次数', '9 次'],
    ['存在 SACC 贷款', '是'],
    ['存在多头借贷', '是'],
    ['持续还款 / 还款中断', '是 / 否'],
  ];
  const aggCols = 2;
  const aggCW = 2.85;
  agg2.forEach((r, i) => {
    const col = i % aggCols;
    const row = Math.floor(i / aggCols);
    const rx = SAFE.left + 6.62 + col * (aggCW + 0.1);
    const ry = uy + row * 0.4;
    addCard(s, rx, ry, aggCW, 0.32, { fill: C.surface, border: C.border });
    s.addText(txt({ text: r[0], x: rx + 0.12, y: ry + 0.01, w: 1.8, h: 0.3, fontSize: 9, color: C.gray, margin: 0 }));
    s.addText(txt({ text: r[1], x: rx + 1.5, y: ry + 0.01, w: 1.25, h: 0.3, fontSize: 9.5, bold: true, color: C.black, align: 'right', margin: 0 }));
  });
}

// ---------------- Slide 5 用户负债汇总（90 天） ----------------
{
  const s = pptx.addSlide();
  let y0 = addHeader(s, 4, '用户负债汇总（最近 90 天）', '识别到的贷款机构：Fair Go Finance 每两周 156 澳元 · Credit Corp 每月 90 澳元 · 其他贷款还款 78 澳元');

  // KPI 条带（7 个）
  const kpis = [
    { v: '2 家', l: '当前贷款机构' },
    { v: '2,300 澳元', l: '最近 90 天借款总额' },
    { v: '1,284 澳元', l: '最近 90 天还款总额' },
    { v: '428 澳元', l: '平均月还款' },
    { v: '1,016 澳元', l: '预计待还金额' },
    { v: '9 次', l: '还款次数' },
  ];
  const kcols = 6;
  const kgap = 0.16;
  const kw = (CONTENT_W - kgap * (kcols - 1)) / kcols;
  let kY = y0 + 0.05;
  kpis.forEach((k, i) => {
    const x = SAFE.left + i * (kw + kgap);
    addCard(s, x, kY, kw, 1.16);
    s.addText(txt({ text: k.v, x: x + 0.12, y: kY + 0.2, w: kw - 0.24, h: 0.42, fontSize: 20, bold: true, color: i === 4 ? C.red : C.blue, align: 'center', margin: 0 }));
    s.addText(txt({ text: k.l, x: x + 0.12, y: kY + 0.7, w: kw - 0.24, h: 0.3, fontSize: 9.5, color: C.gray, align: 'center', margin: 0 }));
  });

  // 判断结论
  let yb = kY + 1.32;
  addCard(s, SAFE.left, yb, CONTENT_W, 0.66, { fill: C.blueLight, border: C.blue });
  const verdict = '判断结论：存在 SACC 贷款 · 存在多头借贷 · 存在持续还款 · 不存在还款中断';
  s.addText(txt({ text: verdict, x: SAFE.left + 0.22, y: yb + 0.14, w: CONTENT_W - 0.44, h: 0.4, fontSize: 13, bold: true, color: C.blue, margin: 0 }));

  // 贷款明细三卡
  let yc = yb + 0.8;
  const loans = [
    { n: 'Fair Go Finance', t: 'SACC 贷款', d: '每两周还款 156 澳元', extra: '最近 90 天还款 6 次' },
    { n: 'Credit Corp', t: '债务催收', d: '每月还款 90 澳元', extra: '最近 90 天持续还款' },
    { n: '其他贷款还款', t: '多项还款交易', d: '合计 78 澳元', extra: '归入其他负债还款' },
  ];
  const lw = (CONTENT_W - 2 * 0.18) / 3;
  loans.forEach((lo, i) => {
    const x = SAFE.left + i * (lw + 0.18);
    addCard(s, x, yc, lw, 2.1);
    addNumBadge(s, x + 0.2, yc + 0.28, i + 1);
    s.addText(txt({ text: lo.n, x: x + 0.62, y: yc + 0.24, w: lw - 0.85, h: 0.3, fontSize: 14, bold: true, color: C.black, margin: 0 }));
    addTag(s, x + 0.62, yc + 0.62, 1.5, lo.t);
    s.addText(txt({ text: lo.d, x: x + 0.24, y: yc + 1.05, w: lw - 0.48, h: 0.32, fontSize: 12, bold: true, color: C.blue, margin: 0 }));
    s.addText(txt({ text: lo.extra, x: x + 0.24, y: yc + 1.44, w: lw - 0.48, h: 0.28, fontSize: 9.8, color: C.gray, margin: 0 }));
  });

  // 注释（页底）
  s.addText(txt({ text: '注：预计待还金额根据借款金额与已还金额估算，可能涉及利息、费用、重复借款及观察期外还款，仅作参考。', x: SAFE.left, y: SH - 0.28 - 0.3, w: CONTENT_W, h: 0.26, fontSize: 9, color: C.gray, margin: 0 }));
}

// ---------------- Slide 6 单笔负债详情 ----------------
{
  const s = pptx.addSlide();
  let y0 = addHeader(s, 5, '单笔负债详情：Fair Go Finance', '贷款明细、还款频率与贷款状态识别');

  // 左：详情表
  addCard(s, SAFE.left, y0, 7.0, 5.0);
  s.addText(txt({ text: '贷款信息', x: SAFE.left + 0.24, y: y0 + 0.18, w: 3, h: 0.32, fontSize: 15, bold: true, color: C.black, margin: 0 }));
  const rows = [
    ['贷款机构', 'Fair Go Finance'],
    ['贷款类型', 'SACC 贷款'],
    ['借款金额', '1,500 澳元'],
    ['最近 90 天已还', '936 澳元'],
    ['预计待还金额', '564 澳元'],
    ['还款频率 / 单次金额', '每两周一次 · 156 澳元'],
    ['最近 90 天还款次数', '6 次'],
    ['贷款状态', '在贷'],
    ['扣款失败 / 还款中断', '否 / 否'],
  ];
  let ry = y0 + 0.66;
  rows.forEach((r, i) => {
    const bg = i % 2 === 0 ? C.white : C.surface;
    s.addShape('rect', { x: SAFE.left + 0.24, y: ry, w: 6.52, h: 0.46, fill: { color: bg }, line: { color: bg, width: 0 } });
    s.addText(txt({ text: r[0], x: SAFE.left + 0.24, y: ry, w: 2.3, h: 0.46, fontSize: 11, bold: true, color: C.gray, margin: 0, valign: 'middle' }));
    s.addText(txt({ text: r[1], x: SAFE.left + 2.7, y: ry, w: 4.0, h: 0.46, fontSize: 11.5, color: C.black, margin: 0, valign: 'middle' }));
    ry += 0.46;
  });

  // 右：状态说明
  let rx0 = SAFE.left + 7.3;
  addCard(s, rx0, y0, 5.23, 5.0);
  s.addText(txt({ text: '贷款状态判定', x: rx0 + 0.24, y: y0 + 0.18, w: 4, h: 0.32, fontSize: 15, bold: true, color: C.black, margin: 0 }));
  const states = [
    { t: '在贷', d: '贷款尚未结清，近期仍存在正常还款。', c: C.blue, bg: C.blueLight, cur: true },
    { t: '结清', d: '贷款已完成还款，近期不再产生相关还款交易。', c: C.gray, bg: C.surface, cur: false },
    { t: '违约中', d: '持续逾期、长期停止还款或多次扣款失败等异常。', c: C.red, bg: C.redBg, cur: false },
    { t: '状态不明', d: '观察周期不足或交易信息不完整，暂时无法准确判断。', c: C.gray, bg: C.surface, cur: false },
  ];
  let sy = y0 + 0.66;
  states.forEach((st) => {
    addCard(s, rx0 + 0.24, sy, 4.75, 0.88, { fill: st.bg, border: st.bg });
    addTag(s, rx0 + 0.42, sy + 0.14, 0.72, st.t, { fill: st.c, color: C.white, fontSize: 9.5 });
    if (st.cur) {
      addTag(s, rx0 + 1.3, sy + 0.14, 1.5, '当前状态', { fill: C.white, color: C.blue, fontSize: 8.5 });
    }
    s.addText(txt({ text: st.d, x: rx0 + 0.42, y: sy + 0.5, w: 4.4, h: 0.32, fontSize: 9.8, color: C.black, margin: 0 }));
    sy += 0.88 + 0.16;
  });
}

// ---------------- Slide 7 AI 应用：7 引擎流水线 + 知识库 ----------------
{
  const s = pptx.addSlide();
  let y0 = addHeader(s, 6, 'AI 在流水处理中的应用', '7 个分类引擎依次处理，后续引擎发现更明确证据时可修正前序分类，最终以最匹配业务语义为准');

  // 引擎流程条
  const engines = ['Init', 'Dishonour', 'Transfer', 'Income', 'Liability', 'Other Credit', 'Fee'];
  const en = engines.length;
  const egap = 0.1;
  const ew = (CONTENT_W - egap * (en - 1)) / en;
  let ey = y0 + 0.02;
  engines.forEach((e, i) => {
    const x = SAFE.left + i * (ew + egap);
    addCard(s, x, ey, ew, 0.5, { border: i === 4 ? C.blue : C.border });
    s.addText(txt({ text: `${i + 1} · ${e}`, x, y: ey, w: ew, h: 0.5, fontSize: 10.5, bold: true, color: i === 4 ? C.blue : C.black, align: 'center', valign: 'middle', margin: 0 }));
  });
  // 第 8 个：质量闭环
  addCard(s, SAFE.left + CONTENT_W / 2 - 1.1, ey + 0.62, 2.2, 0.44, { fill: C.blueLight, border: C.blue });
  s.addText(txt({ text: '8 · 质量评估与优化闭环', x: SAFE.left + CONTENT_W / 2 - 1.1, y: ey + 0.62, w: 2.2, h: 0.44, fontSize: 10, bold: true, color: C.blue, align: 'center', valign: 'middle', margin: 0 }));
  s.addShape('rect', { x: SAFE.left + CONTENT_W / 2 - 0.35, y: ey + 0.5, w: 0.7, h: 0.12, fill: { color: C.blue }, line: { color: C.blue, width: 0 } });

  let y1 = ey + 1.26;
  // 两条知识库维护流水线
  const pipes = [
    {
      t: '流水线一：澳洲官方注册商户数据',
      steps: ['定期拉取注册商户数据', '提取法定 / 商业 / 交易名称', '与现有知识库增量比对', '新增、更新、下线处理', '清洗无效关键词', 'AI 联网验证业务与类别', '更新标准名 / 关键词 / 链接'],
    },
    {
      t: '流水线二：线上未识别商户交易',
      steps: ['定期拉取未识别流水', '筛选未打上商户的交易', 'AI 联网核验真实性', '提取标准名与关键词', '去重合并新别名', '回写商户知识库', '下一轮 Init 分类自动生效'],
    },
  ];
  const pw = (CONTENT_W - 0.18) / 2;
  pipes.forEach((p, pi) => {
    const x = SAFE.left + pi * (pw + 0.18);
    addCard(s, x, y1, pw, 1.5);
    s.addText(txt({ text: p.t, x: x + 0.2, y: y1 + 0.1, w: pw - 0.4, h: 0.26, fontSize: 11.5, bold: true, color: C.black, margin: 0 }));
    s.addText(txt({ text: p.steps.join('  →  '), x: x + 0.2, y: y1 + 0.42, w: pw - 0.4, h: 1.0, fontSize: 9.5, color: C.gray, margin: 0 }));
  });

  // 知识库沉淀示例表
  let y2 = y1 + 1.68;
  s.addText(txt({ text: '商户知识库沉淀示例', x: SAFE.left, y: y2, w: 4, h: 0.28, fontSize: 12, bold: true, color: C.black, margin: 0 }));
  y2 += 0.34;
  const kbRows = [
    ['原始交易描述', '标准商户', '关键词 / 别名', '类别'],
    ['MCDONALDS BRISBANE', "McDonald's", 'MCDONALDS、MCDONALD\'S', 'Dining Out'],
    ['CHEMIST WAREHOUSE 123', 'Chemist Warehouse', 'CHEMIST WAREHOUSE', 'Health'],
    ['AMZNPRIMEA*', 'Amazon Prime', 'AMZNPRIMEA*、Amazon Prime', 'Subscription TV'],
    ['OPENAI *CHATGPT', 'OpenAI', 'OPENAI *CHATGPT、OpenAI', 'Information'],
  ];
  const kbW = [3.6, 2.9, 3.9, 2.2];
  const kbRowOpts = [
    { fill: { color: C.blue }, bold: true, color: C.white },
    { fill: { color: C.surface } }, { fill: { color: C.white } }, { fill: { color: C.surface } }, { fill: { color: C.white } },
  ];
  s.addTable(tableRows(kbRows, kbRowOpts), {
    x: SAFE.left, y: y2, w: 12.6, colW: kbW, rowH: [0.36, 0.4, 0.4, 0.4, 0.4],
    fontFace: FONT, fontSize: 9.5, color: C.black, border: { pt: 0.5, color: C.border },
    align: 'left', valign: 'middle', margin: 0.08,
    autoPage: false,
  });
}

// ---------------- Slide 8 Dishonour / Transfer / Income ----------------
{
  const s = pptx.addSlide();
  let y0 = addHeader(s, 7, 'AI 应用：退票 / 转账 / 收入识别', 'Dishonour Engine 与 Transfer Engine 使用明确规则，Income Engine 综合文本与行为证据');

  // 三个引擎卡
  const engs = [
    {
      t: 'Dishonour Engine · 退票识别',
      n: '1',
      d: '识别退票、撤销、失败扣款及直接扣款退回，标记为退票类交易。使用明确交易文本规则，无需实时 AI 判断；新退票描述通过评估闭环补充为规则。',
    },
    {
      t: 'Transfer Engine · 资金搬运识别',
      n: '2',
      d: '识别内部转账与外部转账，避免资金搬运被误分类为消费、收入或负债。综合转账文本、交易对手、方向与排除规则判断；误判场景由 AI 审计给出优化建议。',
    },
    {
      t: 'Income Engine · 收入识别',
      n: '3',
      d: '识别工资、薪酬包装、政府补助与零工收入。综合交易方向、金额、付款方重复次数、金额稳定性与周 / 双周 / 月规律周期判断。',
    },
  ];
  const ew2 = (CONTENT_W - 2 * 0.18) / 3;
  let ey2 = y0 + 0.02;
  engs.forEach((e, i) => {
    const x = SAFE.left + i * (ew2 + 0.18);
    addCard(s, x, ey2, ew2, 1.9);
    addNumBadge(s, x + 0.2, ey2 + 0.2, Number(e.n));
    s.addText(txt({ text: e.t.split(' · ')[0], x: x + 0.62, y: ey2 + 0.16, w: ew2 - 0.85, h: 0.3, fontSize: 13, bold: true, color: C.black, margin: 0 }));
    s.addText(txt({ text: e.t.split(' · ')[1], x: x + 0.2, y: ey2 + 0.52, w: ew2 - 0.4, h: 0.26, fontSize: 10, color: C.blue, margin: 0 }));
    s.addText(txt({ text: e.d, x: x + 0.2, y: ey2 + 0.84, w: ew2 - 0.4, h: 1.0, fontSize: 9.8, color: C.gray, margin: 0 }));
  });

  // Income 知识库表
  let y3 = ey2 + 2.08;
  s.addText(txt({ text: 'Income 知识库示例：文本模式 + 行为证据共同成立才归为收入', x: SAFE.left, y: y3, w: 12, h: 0.3, fontSize: 12.5, bold: true, color: C.black, margin: 0 }));
  y3 += 0.38;
  const incRows = [
    ['场景', '交易描述与行为线索', '命中的知识', '结果'],
    ['普通工资', 'PAYROLL ACME PTY LTD，同一付款方每两周入账，金额稳定', '工资文本模式 + 双周周期 + 稳定金额阈值', 'salary_payg'],
    ['薪酬包装', 'SALARY PACKAGING ...', '薪酬包装模式优先于普通工资', 'salary_packaging'],
    ['政府补助', 'CENTRELINK ...', 'Centrelink 专项模式', 'centrelink'],
    ['零工收入', 'UBER ... 入账，不含退款 / 贷款 / 转账信号', '零工模式 + 入账方向 + 排除规则', 'self_employed_gig'],
    ['非收入排除', 'LOAN DEPOSIT、退款或内部转账描述', '贷款 / 退款 / 转账负向模式', '不归为收入'],
  ];
  const incW = [1.5, 5.0, 4.1, 2.0];
  const incRowOpts = [
    { fill: { color: C.blue }, bold: true, color: C.white },
    { fill: { color: C.white } }, { fill: { color: C.surface } }, { fill: { color: C.white } }, { fill: { color: C.surface } }, { fill: { color: C.white } },
  ];
  s.addTable(tableRows(incRows, incRowOpts), {
    x: SAFE.left, y: y3, w: 12.6, colW: incW, rowH: [0.36, 0.6, 0.6, 0.6, 0.6, 0.6],
    fontFace: FONT, fontSize: 9.5, color: C.black, border: { pt: 0.5, color: C.border },
    align: 'left', valign: 'middle', margin: 0.08,
    autoPage: false,
  });
}

// ---------------- Slide 9 Liability Engine 知识库 ----------------
{
  const s = pptx.addSlide();
  let y0 = addHeader(s, 8, 'Liability Engine：贷款 / 信用卡 / 催收识别', '先识别交易对手与产品类型，再将同类还款交易归集为负债流');

  const liabRows = [
    ['知识库 / 场景', '交易描述示例', '识别逻辑', '结果'],
    ['BNPL 交易对手库', 'DIRECT DEBIT AFTERPAY', 'Afterpay 命中 BNPL 关键词', 'Afterpay · BNPL'],
    ['信用卡还款规则库', 'BPAY ... ANZ CARDS', '银行 / 账户 / 信用卡文本模式匹配', 'Credit Card Repayment'],
    ['房贷规则库', 'MORTGAGE REPAYMENT', 'Mortgage、Home Loan 关键词', 'Home Loan'],
    ['车贷规则库', '车贷机构或专用描述扣款', '车贷标记 + 交易对手规则', 'Car Loan'],
    ['催收规则库', '已知催收机构名称或关键词', '催收机构关键词与专项标记', 'Debt Collection'],
    ['债务整合规则库', '债务整合机构或描述', '债务整合关键词与专项标记', 'Debt Consolidation'],
    ['退票规则库', 'Return ... Loan Repayment', '贷款还款被退回的文本模式', 'Dishonours'],
    ['透支规则库', 'OVERDRAWN ...', '透支专用关键词', 'Overdrawn'],
    ['通用贷款兜底', '未命中机构但含完整 LOAN', '专项规则未命中后的兜底', 'Generic Loans · Non SACC'],
  ];
  const liabW = [2.3, 3.1, 4.6, 2.6];
  const liabH = [0.36, 0.46, 0.46, 0.46, 0.46, 0.46, 0.46, 0.46, 0.46, 0.46];
  const liabTotalH = liabH.reduce((a, b) => a + b, 0);
  const liabRowOpts = [
    { fill: { color: C.blue }, bold: true, color: C.white },
    { fill: { color: C.surface } }, { fill: { color: C.white } }, { fill: { color: C.surface } }, { fill: { color: C.white } }, { fill: { color: C.surface } },
    { fill: { color: C.white } }, { fill: { color: C.surface } }, { fill: { color: C.white } }, { fill: { color: C.surface } },
  ];
  s.addTable(tableRows(liabRows, liabRowOpts), {
    x: SAFE.left, y: y0 + 0.02, w: 12.6, colW: liabW, rowH: liabH,
    fontFace: FONT, fontSize: 9.5, color: C.black, border: { pt: 0.5, color: C.border },
    align: 'left', valign: 'middle', margin: 0.08,
    autoPage: false,
  });

  // 特例处理说明
  const ny = y0 + 0.02 + liabTotalH + 0.2;
  addCard(s, SAFE.left, ny, CONTENT_W, 0.85, { fill: C.blueLight, border: C.blue });
  s.addText(txt({
    text: '复杂交易特例处理：',
    x: SAFE.left + 0.22, y: ny + 0.08, w: 2, h: 0.24, fontSize: 10.5, bold: true, color: C.blue, margin: 0,
  }));
  s.addText(txt({
    text: 'Credit Corp 按子产品线索判断为 BNPL / 循环额度 / 个人贷款；Cash Converters 含门店、终端或卡尾号消费证据时归为零售，含贷款合同编号时保留贷款含义；Afterpay 限额信息用于额度分析而非单笔判定。',
    x: SAFE.left + 0.22, y: ny + 0.34, w: CONTENT_W - 0.44, h: 0.45, fontSize: 9.5, color: C.black, margin: 0,
  }));
}

// ---------------- Slide 10 All Other Credit / Fee / 质量闭环 ----------------
{
  const s = pptx.addSlide();
  let y0 = addHeader(s, 9, '兜底引擎与 AI 质量闭环', 'All Other Credit 与 Fee 引擎兜底未识别交易，质量闭环持续监控与优化');

  // 两个引擎卡
  const engs2 = [
    { t: 'All Other Credit · 未识别入账兜底', n: '6', d: '归集未被收入、转账、负债引擎识别的入账交易。确保不遗漏可识别入账；高频落入该类的交易提示对应知识库仍有补充空间。' },
    { t: 'Fee Engine · 费用识别', n: '7', d: '识别 ATM、账户、国际交易、透支、退票与现金预支费用。使用费用描述与交易模式规则，在流水线最后完成归类，新费用类型可经闭环补充。' },
  ];
  const ew3 = (CONTENT_W - 0.18) / 2;
  let ey3 = y0 + 0.02;
  engs2.forEach((e, i) => {
    const x = SAFE.left + i * (ew3 + 0.18);
    addCard(s, x, ey3, ew3, 1.5);
    addNumBadge(s, x + 0.2, ey3 + 0.2, Number(e.n));
    s.addText(txt({ text: e.t.split(' · ')[0], x: x + 0.62, y: ey3 + 0.16, w: ew3 - 0.85, h: 0.3, fontSize: 13.5, bold: true, color: C.black, margin: 0 }));
    s.addText(txt({ text: e.t.split(' · ')[1], x: x + 0.2, y: ey3 + 0.52, w: ew3 - 0.4, h: 0.26, fontSize: 10, color: C.blue, margin: 0 }));
    s.addText(txt({ text: e.d, x: x + 0.2, y: ey3 + 0.84, w: ew3 - 0.4, h: 0.6, fontSize: 10, color: C.gray, margin: 0 }));
  });

  // 质量闭环
  let y4 = ey3 + 1.68;
  addCard(s, SAFE.left, y4, CONTENT_W, 2.7, { fill: C.surface, border: C.border });
  s.addText(txt({ text: '质量评估闭环：AI 质检 Agent 监控分类波动并输出诊断报告', x: SAFE.left + 0.22, y: y4 + 0.14, w: 9, h: 0.3, fontSize: 12.5, bold: true, color: C.black, margin: 0 }));
  const qc = [
    { t: '引擎规则问题', d: '规则或优先级配置错误，导致误改或漏分类' },
    { t: '知识库不足', d: '缺失商户、关键词或机构，落入兜底分类' },
    { t: '数据本身异常', d: '描述格式变化、字段变更或上游数据异常' },
    { t: '多引擎冲突', d: '引擎规则冲突，最终分类被错误覆盖' },
    { t: '正常波动', d: '可接受的分类变化，无需处理' },
  ];
  const qw = (CONTENT_W - 4 * 0.14 - 0.44) / 5;
  let qy = y4 + 0.56;
  qc.forEach((q, i) => {
    const x = SAFE.left + 0.22 + i * (qw + 0.14);
    addCard(s, x, qy, qw, 1.5, { fill: C.white, border: C.border });
    addNumBadge(s, x + 0.14, qy + 0.14, i + 1);
    s.addText(txt({ text: q.t, x: x + 0.14, y: qy + 0.55, w: qw - 0.28, h: 0.5, fontSize: 10.5, bold: true, color: C.black, margin: 0 }));
    s.addText(txt({ text: q.d, x: x + 0.14, y: qy + 0.85, w: qw - 0.28, h: 0.6, fontSize: 8.8, color: C.gray, margin: 0 }));
  });
  const fy = qy + 1.62;
  s.addText(txt({
    text: '处置：AI 输出按优先级排序的诊断报告，标明问题引擎、需检查的规则、缺失的商户或关键词，以及应补充的知识库 / 规则表；业务审核后写回，下一轮分类生效。',
    x: SAFE.left + 0.22, y: fy, w: CONTENT_W - 0.44, h: 0.34, fontSize: 10, bold: true, color: C.blue, margin: 0,
  }));
}

// ---------------- Slide 11 Liability 分类效果（OOT 对比） ----------------
{
  const s = pptx.addSlide();
  let y0 = addHeader(s, 10, 'Liability 分类效果：OOT 样本对比', '35 个申请 case · 44,241 条交易 · 2026/07/20–21，对比基准为 illion 外部标签');

  // 三个指标卡
  const kpis = [
    { v: '5,680 条', l: '我方识别贷款', s: '覆盖率 12.8%' },
    { v: '5,363 条', l: 'illion 识别贷款', s: '覆盖率 12.1%' },
    { v: '92.0%', l: '共认贷款分类一致率', s: '双方案例标签一致性' },
  ];
  const kw2 = (CONTENT_W - 2 * 0.16) / 3;
  let kY = y0 + 0.02;
  kpis.forEach((k, i) => {
    const x = SAFE.left + i * (kw2 + 0.16);
    addCard(s, x, kY, kw2, 0.68);
    s.addText(txt({ text: k.v, x: x + 0.12, y: kY + 0.06, w: kw2 - 0.24, h: 0.32, fontSize: 17, bold: true, color: C.blue, align: 'center', margin: 0 }));
    s.addText(txt({ text: k.l, x: x + 0.12, y: kY + 0.4, w: kw2 - 0.24, h: 0.22, fontSize: 9.5, bold: true, color: C.black, align: 'center', margin: 0 }));
    s.addText(txt({ text: k.s, x: x + 0.12, y: kY + 0.58, w: kw2 - 0.24, h: 0.18, fontSize: 8, color: C.gray, align: 'center', margin: 0 }));
  });

  // 覆盖率对比（横向条形图）+ 韦恩集合并列
  let yb2 = kY + 0.76;
  addCard(s, SAFE.left, yb2, 5.6, 2.28);
  s.addText(txt({ text: '贷款识别覆盖率对比', x: SAFE.left + 0.22, y: yb2 + 0.12, w: 5, h: 0.28, fontSize: 12.5, bold: true, color: C.black, margin: 0 }));
  s.addChart(pptx.ChartType.bar, [
    {
      name: '覆盖率',
      labels: ['我方', 'illion'],
      values: [12.8, 12.1],
    },
  ], {
    x: SAFE.left + 0.22, y: yb2 + 0.46, w: 5.1, h: 1.6,
    chartColors: [C.blue, '9AB8E0'],
    showValue: true, dataLabelPosition: 'outEnd', dataLabelColor: C.black, dataLabelFontSize: 11,
    catAxisLabelColor: C.black, catAxisLabelFontSize: 10, catAxisLabelFontFace: FONT,
    catGridLine: { style: 'none' },
    valAxisHidden: true,
    valGridLine: { style: 'none' },
    showLegend: false, showTitle: false,
    barDir: 'bar',
  });

  // 并集覆盖卡
  addCard(s, SAFE.left + 5.8, yb2, 6.7, 2.28);
  s.addText(txt({ text: '识别覆盖交集（韦恩图数据）', x: SAFE.left + 6.02, y: yb2 + 0.12, w: 5, h: 0.28, fontSize: 12.5, bold: true, color: C.black, margin: 0 }));
  // 简化韦恩：两个重叠圆
  const ovY = yb2 + 0.48;
  s.addShape('ellipse', { x: SAFE.left + 6.35, y: ovY, w: 1.45, h: 1.45, fill: { color: C.blue, transparency: 35 }, line: { color: C.blue, width: 1.5 } });
  s.addShape('ellipse', { x: SAFE.left + 7.4, y: ovY, w: 1.45, h: 1.45, fill: { color: '9AB8E0', transparency: 35 }, line: { color: '5B84C4', width: 1.5 } });
  s.addText(txt({ text: '我方 5,680', x: SAFE.left + 6.35, y: ovY + 0.6, w: 1.45, h: 0.28, fontSize: 9, bold: true, color: C.blue, align: 'center', margin: 0 }));
  s.addText(txt({ text: '共认 5,284', x: SAFE.left + 7.05, y: ovY + 0.6, w: 1.25, h: 0.28, fontSize: 9, bold: true, color: C.black, align: 'center', margin: 0 }));
  s.addText(txt({ text: 'illion 5,363', x: SAFE.left + 7.4, y: ovY + 0.6, w: 1.45, h: 0.28, fontSize: 9, bold: true, color: '4A6FA5', align: 'center', margin: 0 }));
  // 左右独有说明
  s.addText(txt({ text: '我方独有 396 条：BNPL · EWA · Home Loan', x: SAFE.left + 6.02, y: ovY + 1.52, w: 3.3, h: 0.22, fontSize: 9, bold: true, color: C.blue, margin: 0 }));
  s.addText(txt({ text: 'illion 独有 79 条：知识库缺失或归属差异', x: SAFE.left + 9.42, y: ovY + 1.52, w: 3.2, h: 0.22, fontSize: 9, bold: true, color: C.gray, margin: 0 }));
  s.addText(txt({ text: '我方漏识别更少，整体覆盖更优', x: SAFE.left + 6.02, y: ovY + 1.76, w: 6.5, h: 0.24, fontSize: 10, bold: true, color: C.black, margin: 0 }));

  // 分类矩阵表
  let y5 = yb2 + 2.4;
  s.addText(txt({ text: '共认贷款标签一致性矩阵（5,284 条）', x: SAFE.left, y: y5, w: 8, h: 0.26, fontSize: 11.5, bold: true, color: C.black, margin: 0 }));
  y5 += 0.3;
  const mRows = [
    ['illion 分类 \\ 我方', '我方 SACC', '我方 Non-SACC', '我方 Unknown', '小计'],
    ['illion SACC', '262 · 60.0%', '91 · 20.8%', '80 · 18.3%', '437'],
    ['illion Non-SACC', '106 · 2.2%', '4,601 · 93.4%', '144 · 2.9%', '4,926'],
    ['我方小计', '377', '5,072', '224', '5,284'],
  ];
  const mW = [3.3, 2.4, 2.6, 2.4, 1.9];
  const mH = [0.32, 0.38, 0.38, 0.38];
  const mRowOpts = [
    { fill: { color: C.blue }, bold: true, color: C.white },
    { fill: { color: C.white } }, { fill: { color: C.surface } }, { fill: { color: C.blueLight }, bold: true },
  ];
  s.addTable(tableRows(mRows, mRowOpts), {
    x: SAFE.left, y: y5, w: 12.6, colW: mW, rowH: mH,
    fontFace: FONT, fontSize: 10, color: C.black, border: { pt: 0.5, color: C.border },
    align: 'center', valign: 'middle', margin: 0.06,
    autoPage: false,
  });

  // 结论
  const ny5 = y5 + mH.reduce((a, b) => a + b, 0) + 0.1;
  addCard(s, SAFE.left, ny5, CONTENT_W, 0.44, { fill: C.blueLight, border: C.blue });
  s.addText(txt({
    text: '解读：illion 的 SACC 标记中 39.1% 被我方修正（过分类明显）；我方 Non-SACC 与 illion 一致率 93.4%，分类更稳健。',
    x: SAFE.left + 0.22, y: ny5 + 0.06, w: CONTENT_W - 0.44, h: 0.32, fontSize: 10.5, bold: true, color: C.blue, margin: 0,
  }));
}

// ---------------- Slide 12 典型案例 ----------------
{
  const s = pptx.addSlide();
  let y0 = addHeader(s, 11, '典型案例：覆盖 → 一致性 → 精度', 'OOT 对比分析（35 个申请 case）中三类差异化价值');

  const cases = [
    {
      tag: '覆盖优势',
      tagC: C.blue,
      tagBg: C.blueLight,
      title: 'Case 1 · Home Loan 被 illion 漏识别',
      q: '交易 "Home loan - Receipt ... To Orange Everyday" 被 illion 标为 Internal Transfer（非贷款），完全遗漏住房贷款。',
      a: '我方基于 "Home loan" 产品关键词判定为贷款，细分为 Non-SACC Loans（Home Loan）。',
      v: '价值：通过产品词库识别捕捉 illion 覆盖盲区，扩大贷款识别面。',
    },
    {
      tag: '一致性',
      tagC: C.black,
      tagBg: C.surface,
      title: 'Case 2 · 同一贷款流内 illion 标签自相矛盾（MoneySpot）',
      q: '放款 750 澳元、等额还款 49.09 澳元 × 10；illion 前 6 笔标 SACC、后 5 笔标 Non-SACC。',
      a: '我方通过放款与还款匹配识别为单一贷款流，按小额等额高频特征统一判定为 SACC Loans。',
      v: '价值：验证贷款流粒度分类一致性，纠正 illion 碎片化误判。',
    },
    {
      tag: '精度',
      tagC: C.red,
      tagBg: C.redBg,
      title: 'Case 3 · LOC 产品被 illion 泛化为 SACC（Credit24）',
      q: '同一用户 19 笔 Credit24 交易，illion 全部标为 SACC Loans，忽略产品类型差异。',
      a: '我方结合交易模式与机构产品库识别为循环授信（LOC），归入 Non-SACC Loans。',
      v: '价值：细粒度产品区分，避免非 SACC 产品误判为高风险 SACC，为风控提供准确信号。',
    },
  ];
  const ch2 = 5.1;
  const cw3 = (CONTENT_W - 2 * 0.18) / 3;
  let cy = y0 + 0.02;
  cases.forEach((cs, i) => {
    const x = SAFE.left + i * (cw3 + 0.18);
    addCard(s, x, cy, cw3, ch2);
    addTag(s, x + 0.2, cy + 0.18, 1.1, cs.tag, { fill: cs.tagBg, color: cs.tagC, fontSize: 9.5 });
    s.addText(txt({ text: cs.title, x: x + 0.2, y: cy + 0.56, w: cw3 - 0.4, h: 0.62, fontSize: 12.5, bold: true, color: C.black, margin: 0 }));
    const qy2 = cy + 1.28;
    s.addText(txt({ text: '问题', x: x + 0.2, y: qy2, w: 1, h: 0.26, fontSize: 10.5, bold: true, color: C.red, margin: 0 }));
    s.addText(txt({ text: cs.q, x: x + 0.2, y: qy2 + 0.26, w: cw3 - 0.4, h: 0.85, fontSize: 9.8, color: C.black, margin: 0 }));
    const ay3 = qy2 + 1.18;
    s.addText(txt({ text: '我方逻辑', x: x + 0.2, y: ay3, w: 1.2, h: 0.26, fontSize: 10.5, bold: true, color: C.blue, margin: 0 }));
    s.addText(txt({ text: cs.a, x: x + 0.2, y: ay3 + 0.26, w: cw3 - 0.4, h: 0.85, fontSize: 9.8, color: C.black, margin: 0 }));
    const vy = ay3 + 1.18;
    addCard(s, x + 0.2, vy, cw3 - 0.4, 0.92, { fill: cs.tagBg, border: cs.tagBg });
    s.addText(txt({ text: cs.v, x: x + 0.3, y: vy + 0.14, w: cw3 - 0.6, h: 0.64, fontSize: 9.8, bold: true, color: cs.tagC, margin: 0 }));
  });
}

// ---------------- Slide 13 总结与价值 ----------------
{
  const s = pptx.addSlide();
  let y0 = addHeader(s, 12, '总结与价值', '自主交易分类能力带来的业务与工程价值');

  const values = [
    { t: '业务价值', d: '收入、负债、支出识别更准确，Serviceability 计算更可靠，支撑 Responsible Lending 合规要求。', n: '1' },
    { t: '能力自主', d: '沉淀商户知识库与规则，脱离供应商黑盒，快速响应业务、风险与合规变化。', n: '2' },
    { t: 'AI 赋能', d: '7 引擎流水线 + 知识库自动维护 + 质量闭环，AI 辅助实时判断与规则优化。', n: '3' },
    { t: '效果验证', d: 'OOT 覆盖 12.8% vs 12.1%，共认一致率 92.0%，SACC 修正率 39.1%，Non-SACC 一致率 93.4%。', n: '4' },
  ];
  const cw4 = (CONTENT_W - 3 * 0.18) / 4;
  let cy4 = y0 + 0.02;
  values.forEach((v, i) => {
    const x = SAFE.left + i * (cw4 + 0.18);
    addCard(s, x, cy4, cw4, 2.3);
    addNumBadge(s, x + 0.2, cy4 + 0.2, Number(v.n));
    s.addText(txt({ text: v.t, x: x + 0.2, y: cy4 + 0.62, w: cw4 - 0.4, h: 0.3, fontSize: 14, bold: true, color: C.black, margin: 0 }));
    s.addText(txt({ text: v.d, x: x + 0.2, y: cy4 + 1.0, w: cw4 - 0.4, h: 1.2, fontSize: 10.5, color: C.gray, margin: 0 }));
  });

  // 深色结语面板
  const py = cy4 + 2.5;
  addCard(s, SAFE.left, py, CONTENT_W, 1.9, { fill: C.dark, border: C.dark });
  s.addText(txt({ text: '下一步', x: SAFE.left + 0.3, y: py + 0.18, w: 2, h: 0.3, fontSize: 13, bold: true, color: C.blue, margin: 0 }));
  const nexts = [
    '持续扩充商户与负债知识库',
    '接入 Open Banking 多源数据',
    '建立效果监控与模型版本管理',
  ];
  nexts.forEach((n, i) => {
    const x = SAFE.left + 0.3 + i * 4.1;
    s.addShape('ellipse', { x, y: py + 0.7, w: 0.14, h: 0.14, fill: { color: C.blue }, line: { color: C.blue, width: 0 } });
    s.addText(txt({ text: n, x: x + 0.22, y: py + 0.62, w: 3.7, h: 0.3, fontSize: 11, color: C.white, margin: 0 }));
  });
  const ny6 = py + 1.15;
  s.addText(txt({
    text: '通过自主模型提升分类可解释性、可配置性与迭代效率，降低长期供应商依赖。',
    x: SAFE.left + 0.3, y: ny6, w: CONTENT_W - 0.6, h: 0.5, fontSize: 10.5, color: 'B0B6BD', margin: 0,
  }));
}

// ---------- 输出 ----------
const outPath = '银行流水汇报_v2.pptx';
pptx.writeFile({ fileName: outPath })
  .then(() => console.log('written:', outPath, fs.existsSync(outPath) ? fs.statSync(outPath).size : 0))
  .catch((e) => { console.error('WRITE FAIL', e); process.exit(1); });
