const pptxgen = require("pptxgenjs");

const pres = new pptxgen();
pres.layout = "LAYOUT_16x9";
pres.author = "Gary Chen";
pres.title = "Gary Chen 五年 AI Operator 戰略地圖";

// Gary Chen Brand Color Palette — 嚴格只用 Navy + Electric Blue + 灰階
const C = {
  primary: "091F33",       // Gary Navy - 頂部條、主強調、CTA
  accent: "1E77E9",        // Electric Blue - 卡片邊條、第二強調
  primaryDark: "050F1A",   // Midnight Navy - 章節分隔頁深底
  nearBlack: "0A0E14",     // 主文字
  darkGray: "5A6573",      // 副文字
  midGray: "8A95A5",       // 說明、邊框
  lightText: "C4CCD6",     // 深底頁面副文字
  white: "FFFFFF",         // 白色背景
  offWhite: "F8FAFB",      // 微灰白卡片底
  lightGray: "EEF2F6",     // 淺灰背景、Insight Bar
  divider: "D8DEE5",       // 分隔線
  warning: "C8463C",       // 隱藏警示色 - 僅 risk matrix / 狀態紅燈 / 負成長
};
const FONT = "Noto Sans TC";

const mkShadow = () => ({ type: "outer", color: "000000", blur: 4, offset: 2, angle: 135, opacity: 0.08 });

// Helper: Navy 頂條 + 標題
function addTitle(s, text, opts = {}) {
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.05, fill: { color: C.primary } });
  s.addText(text, {
    x: 0.8, y: 0.35, w: 8.4, h: 0.9,
    fontSize: 24, fontFace: FONT, color: C.nearBlack, bold: true, margin: 0, valign: "top"
  });
  if (!opts.noDivider) {
    s.addShape(pres.shapes.RECTANGLE, {
      x: 0.8, y: 1.28, w: 8.4, h: 0.025,
      fill: { color: C.divider }
    });
  }
}

// Helper: Insight Bar
function addInsight(s, text, opts = {}) {
  const y = opts.y || 4.85;
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.8, y: y, w: 8.4, h: 0.55,
    fill: { color: C.lightGray }
  });
  s.addText(text, {
    x: 1.05, y: y, w: 8, h: 0.55,
    fontSize: 13, fontFace: FONT, color: C.nearBlack,
    bold: true, valign: "middle", margin: 0
  });
}

// ============ SLIDE 1: 封面 ============
{
  const s = pres.addSlide();
  s.background = { color: C.lightGray };
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.05, fill: { color: C.primary } });

  // 裝飾幾何
  s.addShape(pres.shapes.OVAL, { x: 7.5, y: -0.8, w: 4, h: 4, fill: { color: C.white, transparency: 60 } });
  s.addShape(pres.shapes.OVAL, { x: 8.5, y: 3.5, w: 3, h: 3, fill: { color: C.white, transparency: 50 } });

  s.addText("Gary Chen 五年 AI Operator 戰略地圖", {
    x: 0.8, y: 1.2, w: 8, h: 1.4,
    fontSize: 40, fontFace: FONT, color: C.nearBlack, bold: true, margin: 0
  });
  s.addShape(pres.shapes.RECTANGLE, { x: 0.8, y: 2.85, w: 3, h: 0.04, fill: { color: C.accent } });
  s.addText("從「AI 工具使用者」升級為「AI 系統打造者」", {
    x: 0.8, y: 3.15, w: 8, h: 0.6,
    fontSize: 20, fontFace: FONT, color: C.darkGray, italic: true, margin: 0
  });
  s.addText("Personal AI Operator Roadmap", {
    x: 0.8, y: 3.7, w: 8, h: 0.4,
    fontSize: 14, fontFace: FONT, color: C.midGray, margin: 0
  });
  s.addText("2026 年 5 月", {
    x: 0.8, y: 4.8, w: 4, h: 0.4,
    fontSize: 12, fontFace: FONT, color: C.midGray, margin: 0
  });
}

// ============ SLIDE 2: 核心洞見（三卡片）============
{
  const s = pres.addSlide();
  s.background = { color: C.white };
  addTitle(s, "我已擁有三個多數知識工作者一輩子拿不到的關鍵槓桿");

  const assets = [
    {
      num: "1", title: "判斷力",
      sub: "Editorial Taste", color: C.primary,
      items: ["十年領域知識累積", "對好內容的辨識度", "拒絕平庸的審美"]
    },
    {
      num: "2", title: "工作流",
      sub: "Personal Workflow", color: C.accent,
      items: ["可重複的研究 / 寫作 / 出貨流程", "AI 工具串接已成肌肉記憶", "每次迭代都被記錄"]
    },
    {
      num: "3", title: "受眾資產",
      sub: "Audience Capital", color: C.primary,
      items: ["長期累積的訂閱者", "高訊雜比的回饋圈", "具備分發能力的個人通路"]
    },
  ];

  assets.forEach((a, i) => {
    const x = 0.8 + i * 3.1;
    s.addShape(pres.shapes.RECTANGLE, { x, y: 1.55, w: 2.8, h: 3.0, fill: { color: C.offWhite }, shadow: mkShadow() });
    s.addShape(pres.shapes.RECTANGLE, { x, y: 1.55, w: 2.8, h: 0.06, fill: { color: a.color } });

    s.addShape(pres.shapes.OVAL, { x: x + 0.15, y: 1.75, w: 0.45, h: 0.45, fill: { color: a.color } });
    s.addText(a.num, {
      x: x + 0.15, y: 1.75, w: 0.45, h: 0.45,
      fontSize: 16, fontFace: FONT, color: C.white, bold: true, align: "center", valign: "middle", margin: 0
    });
    s.addText(a.title, {
      x: x + 0.7, y: 1.75, w: 1.9, h: 0.25,
      fontSize: 16, fontFace: FONT, color: C.nearBlack, bold: true, margin: 0
    });
    s.addText(a.sub, {
      x: x + 0.7, y: 2.0, w: 1.9, h: 0.25,
      fontSize: 12, fontFace: FONT, color: a.color, margin: 0
    });

    s.addText(
      a.items.map((item, j) => ({
        text: item,
        options: { bullet: true, breakLine: j < a.items.length - 1, fontSize: 12 }
      })),
      {
        x: x + 0.2, y: 2.5, w: 2.4, h: 1.8,
        fontFace: FONT, color: C.darkGray, valign: "top", margin: 0, paraSpaceAfter: 6
      }
    );
  });

  addInsight(s, "三者合起來 → 「判斷 × 工作流 × 受眾」是 AI 時代極少數無法被模型取代的個人槓桿");
}

// ============ SLIDE 3: 願景架構（五步流程）============
{
  const s = pres.addSlide();
  s.background = { color: C.lightGray };
  addTitle(s, "AI Operator 工作流：把模糊意圖系統化為可重複、可委派、可規模化的輸出");

  const steps = [
    { label: "意圖", desc: "釐清「為誰\n解決什麼問題」", color: C.primary },
    { label: "拆解", desc: "切成可委派的\n子任務", color: C.accent },
    { label: "委派", desc: "Agents / SDK\n並行執行", color: C.primary },
    { label: "驗證", desc: "判斷力 + 品味\n篩選輸出", color: C.accent },
    { label: "整合", desc: "封裝為 Workflow\n可重複使用", color: C.primary },
  ];

  steps.forEach((st, i) => {
    const x = 0.5 + i * 1.9;
    s.addShape(pres.shapes.RECTANGLE, { x, y: 1.7, w: 1.65, h: 2.2, fill: { color: C.white }, shadow: mkShadow() });
    s.addShape(pres.shapes.RECTANGLE, { x, y: 1.7, w: 1.65, h: 0.06, fill: { color: st.color } });

    s.addText(st.label, {
      x, y: 1.95, w: 1.65, h: 0.4,
      fontSize: 15, fontFace: FONT, color: C.nearBlack, bold: true, align: "center", margin: 0
    });
    s.addText(st.desc, {
      x, y: 2.45, w: 1.65, h: 0.8,
      fontSize: 11, fontFace: FONT, color: C.darkGray, align: "center", margin: 0
    });

    if (i < steps.length - 1) {
      s.addText("→", {
        x: x + 1.65, y: 2.3, w: 0.25, h: 0.4,
        fontSize: 18, fontFace: FONT, color: C.midGray, align: "center", valign: "middle", margin: 0
      });
    }
  });

  s.addText([
    { text: "對標：", options: { bold: true, fontSize: 12, color: C.midGray } },
    { text: "Linear、Vercel、Cursor 的內部 PRD 流程", options: { fontSize: 12, color: C.midGray } },
    { text: " — 但落到個人尺度，工具門檻已經消失", options: { bold: true, fontSize: 12, color: C.accent } },
  ], {
    x: 0.8, y: 4.2, w: 8.4, h: 0.4, fontFace: FONT, margin: 0
  });

  addInsight(s, "AI 時代的差距不在「會不會用工具」，而在「能不能把工作流系統化」");
}

// ============ SLIDE 4: 第一階段 ============
{
  const s = pres.addSlide();
  s.background = { color: C.white };
  addTitle(s, "第一階段（現在～1 年）：建立個人 AI Stack，把所有手動工作收斂成可重用元件");

  s.addShape(pres.shapes.RECTANGLE, { x: 0.8, y: 1.5, w: 1.8, h: 0.35, fill: { color: C.primary } });
  s.addText("Phase 1：AI Stack", {
    x: 0.8, y: 1.5, w: 1.8, h: 0.35,
    fontSize: 11, fontFace: FONT, color: C.white, bold: true, align: "center", valign: "middle", margin: 0
  });

  // 左：已擁有
  s.addShape(pres.shapes.RECTANGLE, { x: 0.8, y: 2.1, w: 4, h: 2.4, fill: { color: C.offWhite }, shadow: mkShadow() });
  s.addShape(pres.shapes.RECTANGLE, { x: 0.8, y: 2.1, w: 4, h: 0.05, fill: { color: C.accent } });
  s.addText("已擁有 ✓", {
    x: 1.1, y: 2.3, w: 3.5, h: 0.3,
    fontSize: 15, fontFace: FONT, color: C.accent, bold: true, margin: 0
  });
  s.addText([
    { text: "Claude / Cursor / Codex 日常使用", options: { bullet: true, breakLine: true, fontSize: 13 } },
    { text: "個人寫作 / 研究的反覆嘗試", options: { bullet: true, fontSize: 13 } },
  ], {
    x: 1.1, y: 2.7, w: 3.5, h: 1.5,
    fontFace: FONT, color: C.nearBlack, valign: "top", margin: 0, paraSpaceAfter: 8
  });

  // 右：需要補齊
  s.addShape(pres.shapes.RECTANGLE, { x: 5.2, y: 2.1, w: 4, h: 2.4, fill: { color: C.offWhite }, shadow: mkShadow() });
  s.addShape(pres.shapes.RECTANGLE, { x: 5.2, y: 2.1, w: 4, h: 0.05, fill: { color: C.primary } });
  s.addText("需要補齊 →", {
    x: 5.5, y: 2.3, w: 3.5, h: 0.3,
    fontSize: 15, fontFace: FONT, color: C.primary, bold: true, margin: 0
  });
  s.addText([
    { text: "個人化 Agents（研究、編輯、發布）", options: { bullet: true, breakLine: true, fontSize: 13 } },
    { text: "可重用的 Prompt Library", options: { bullet: true, breakLine: true, fontSize: 13 } },
    { text: "MCP Servers 連接私人資料", options: { bullet: true, breakLine: true, fontSize: 13 } },
    { text: "工作流自動化（Zapier / Make / 自寫）", options: { bullet: true, fontSize: 13 } },
  ], {
    x: 5.5, y: 2.7, w: 3.5, h: 1.5,
    fontFace: FONT, color: C.nearBlack, valign: "top", margin: 0, paraSpaceAfter: 6
  });

  addInsight(s, "目標：打開電腦就能在 5 分鐘內啟動「研究 → 寫作 → 發布」的完整 AI 流水線");
}

// ============ SLIDE 5: 第二階段 ============
{
  const s = pres.addSlide();
  s.background = { color: C.lightGray };
  addTitle(s, "第二階段（1～2 年）：把個人工作流產品化，從「我能做」變成「系統幫我做」");

  s.addShape(pres.shapes.RECTANGLE, { x: 0.8, y: 1.5, w: 1.8, h: 0.35, fill: { color: C.accent } });
  s.addText("Phase 2：Productize", {
    x: 0.8, y: 1.5, w: 1.8, h: 0.35,
    fontSize: 11, fontFace: FONT, color: C.white, bold: true, align: "center", valign: "middle", margin: 0
  });

  // 左卡片：Workflow Library
  s.addShape(pres.shapes.RECTANGLE, { x: 0.8, y: 2.1, w: 4, h: 2.4, fill: { color: C.white }, shadow: mkShadow() });
  s.addText("Workflow Library", {
    x: 1.1, y: 2.25, w: 3.5, h: 0.3,
    fontSize: 15, fontFace: FONT, color: C.nearBlack, bold: true, margin: 0
  });
  s.addText([
    { text: "可重複呼叫的工作流：", options: { bold: true, breakLine: true, fontSize: 13 } },
    { text: "• 研究報告生成（30 分鐘 → 5 分鐘）", options: { breakLine: true, fontSize: 12 } },
    { text: "• 內容草稿產出（2 小時 → 20 分鐘）", options: { breakLine: true, fontSize: 12 } },
    { text: "", options: { breakLine: true, fontSize: 8 } },
    { text: "節省時間：", options: { bold: true, breakLine: true, fontSize: 13 } },
    { text: "每週 15+ 小時", options: { fontSize: 18, bold: true, color: C.accent } },
  ], {
    x: 1.1, y: 2.65, w: 3.5, h: 1.6,
    fontFace: FONT, color: C.nearBlack, valign: "top", margin: 0
  });

  // 右卡片：Personal API
  s.addShape(pres.shapes.RECTANGLE, { x: 5.2, y: 2.1, w: 4, h: 2.4, fill: { color: C.white }, shadow: mkShadow() });
  s.addText("Personal API", {
    x: 5.5, y: 2.25, w: 3.5, h: 0.3,
    fontSize: 15, fontFace: FONT, color: C.nearBlack, bold: true, margin: 0
  });
  s.addText([
    { text: "個人服務介面：", options: { bold: true, breakLine: true, fontSize: 13 } },
    { text: "• /research <topic> → 結構化報告", options: { breakLine: true, fontSize: 12 } },
    { text: "• /draft <brief> → 風格化草稿", options: { breakLine: true, fontSize: 12 } },
    { text: "", options: { breakLine: true, fontSize: 8 } },
    { text: "對外開放：", options: { bold: true, breakLine: true, fontSize: 13 } },
    { text: "Subscribers 私域試用 →", options: { fontSize: 16, bold: true, color: C.accent } },
  ], {
    x: 5.5, y: 2.65, w: 3.5, h: 1.6,
    fontFace: FONT, color: C.nearBlack, valign: "top", margin: 0
  });

  // 中間箭頭
  s.addText("→", {
    x: 4.8, y: 2.8, w: 0.4, h: 0.5,
    fontSize: 24, fontFace: FONT, color: C.primary, align: "center", valign: "middle", margin: 0
  });

  addInsight(s, "從「靠技能換時間」到「靠系統換槓桿」— 這一步是 Operator 與 Worker 的分水嶺");
}

// ============ SLIDE 6: 第三階段 ============
{
  const s = pres.addSlide();
  s.background = { color: C.white };
  addTitle(s, "第三階段（2～3 年）：個人 IP 系統化後，受眾會主動把問題送上門");

  s.addShape(pres.shapes.RECTANGLE, { x: 0.8, y: 1.5, w: 2.0, h: 0.35, fill: { color: C.primary } });
  s.addText("Phase 3：IP Marketplace", {
    x: 0.8, y: 1.5, w: 2.0, h: 0.35,
    fontSize: 11, fontFace: FONT, color: C.white, bold: true, align: "center", valign: "middle", margin: 0
  });

  // 左：個人輸出
  s.addText("個人持續產出", {
    x: 0.8, y: 2.1, w: 3.2, h: 0.3,
    fontSize: 14, fontFace: FONT, color: C.darkGray, bold: true, margin: 0
  });
  const outputItems = ["週報 / 月報", "深度分析", "工具評測", "工作流模板"];
  outputItems.forEach((d, i) => {
    const y = 2.55 + i * 0.5;
    s.addShape(pres.shapes.RECTANGLE, { x: 0.8, y, w: 2.8, h: 0.38, fill: { color: C.lightGray } });
    s.addText("• " + d, {
      x: 1.0, y, w: 2.4, h: 0.38,
      fontSize: 13, fontFace: FONT, color: C.nearBlack, valign: "middle", margin: 0
    });
  });

  // 中間箭頭
  s.addText("→", {
    x: 3.8, y: 2.9, w: 0.5, h: 0.5,
    fontSize: 28, fontFace: FONT, color: C.accent, align: "center", valign: "middle", margin: 0
  });

  // 右：受眾需求 2x3
  s.addText("受眾主動需求", {
    x: 4.5, y: 2.1, w: 5, h: 0.3,
    fontSize: 14, fontFace: FONT, color: C.darkGray, bold: true, margin: 0
  });
  const products = ["諮詢", "工作坊", "Cohort", "顧問", "代寫", "授權"];
  products.forEach((p, i) => {
    const col = i % 3;
    const row = Math.floor(i / 3);
    const x = 4.5 + col * 1.8;
    const y = 2.55 + row * 1.0;
    s.addShape(pres.shapes.RECTANGLE, { x, y, w: 1.55, h: 0.75, fill: { color: C.offWhite }, shadow: mkShadow() });
    s.addShape(pres.shapes.RECTANGLE, { x, y, w: 1.55, h: 0.05, fill: { color: C.accent } });
    s.addText(p, {
      x, y, w: 1.55, h: 0.75,
      fontSize: 15, fontFace: FONT, color: C.nearBlack, bold: true, align: "center", valign: "middle", margin: 0
    });
  });

  addInsight(s, "個人 IP 變成「需求 Marketplace」，客戶主動找上門，而不再需要主動推銷");
}

// ============ SLIDE 7: 第四階段 ============
{
  const s = pres.addSlide();
  s.background = { color: C.lightGray };
  addTitle(s, "第四階段（3～5 年）：從 Operator 升級為 Operator-of-Operators");

  // 漏斗
  s.addShape(pres.shapes.RECTANGLE, { x: 2.5, y: 1.6, w: 5, h: 0.7, fill: { color: C.offWhite }, shadow: mkShadow() });
  s.addText("受 眾", {
    x: 2.5, y: 1.6, w: 5, h: 0.7,
    fontSize: 20, fontFace: FONT, color: C.nearBlack, bold: true, align: "center", valign: "middle", margin: 0
  });

  s.addText("↓", {
    x: 4.5, y: 2.3, w: 1, h: 0.4,
    fontSize: 24, fontFace: FONT, color: C.primary, align: "center", margin: 0
  });

  s.addShape(pres.shapes.RECTANGLE, { x: 3, y: 2.7, w: 4, h: 0.7, fill: { color: C.primary } });
  s.addText("Gary Chen Operator OS", {
    x: 3, y: 2.7, w: 4, h: 0.7,
    fontSize: 18, fontFace: FONT, color: C.white, bold: true, align: "center", valign: "middle", margin: 0
  });

  s.addText("↓", {
    x: 4.5, y: 3.4, w: 1, h: 0.4,
    fontSize: 24, fontFace: FONT, color: C.primary, align: "center", margin: 0
  });

  s.addShape(pres.shapes.RECTANGLE, { x: 1.5, y: 3.8, w: 7, h: 0.7, fill: { color: C.offWhite }, shadow: mkShadow() });
  s.addText("一個小團隊  +  數十條自動化工作流", {
    x: 1.5, y: 3.8, w: 7, h: 0.7,
    fontSize: 20, fontFace: FONT, color: C.nearBlack, bold: true, align: "center", valign: "middle", margin: 0
  });

  addInsight(s, "個人 → 工作流 → 小團隊運營工作流 → 個人成為「微型 Operating System」");
}

// ============ SLIDE 8: AI 時代的個人護城河 ============
{
  const s = pres.addSlide();
  s.background = { color: C.lightGray };
  addTitle(s, "AI 時代的個人護城河不是模型也不是數據 — 是判斷力 × 工作流 × 受眾");

  // 左：我擁有的
  s.addShape(pres.shapes.RECTANGLE, { x: 0.8, y: 1.55, w: 4, h: 2.8, fill: { color: C.white }, shadow: mkShadow() });
  s.addShape(pres.shapes.RECTANGLE, { x: 0.8, y: 1.55, w: 4, h: 0.06, fill: { color: C.accent } });
  s.addText("Operator 看到的", {
    x: 1.1, y: 1.8, w: 3.5, h: 0.3,
    fontSize: 15, fontFace: FONT, color: C.accent, bold: true, margin: 0
  });
  s.addText([
    { text: "可重複的 Workflow", options: { bullet: true, breakLine: true, fontSize: 13 } },
    { text: "對輸出的審美底線", options: { bullet: true, breakLine: true, fontSize: 13 } },
    { text: "受眾的真實回饋圈", options: { bullet: true, breakLine: true, fontSize: 13 } },
    { text: "犯過錯的歷史經驗", options: { bullet: true, fontSize: 13 } },
  ], {
    x: 1.1, y: 2.25, w: 3.5, h: 1.8,
    fontFace: FONT, color: C.nearBlack, valign: "top", margin: 0, paraSpaceAfter: 8
  });

  // vs
  s.addText("vs", {
    x: 4.6, y: 2.6, w: 0.8, h: 0.5,
    fontSize: 18, fontFace: FONT, color: C.midGray, bold: true, align: "center", valign: "middle", margin: 0
  });

  // 右：模型看到的
  s.addShape(pres.shapes.RECTANGLE, { x: 5.2, y: 1.55, w: 4, h: 2.8, fill: { color: C.white }, shadow: mkShadow() });
  s.addShape(pres.shapes.RECTANGLE, { x: 5.2, y: 1.55, w: 4, h: 0.06, fill: { color: C.midGray } });
  s.addText("純粹靠模型的人看到的", {
    x: 5.5, y: 1.8, w: 3.5, h: 0.3,
    fontSize: 15, fontFace: FONT, color: C.midGray, bold: true, margin: 0
  });
  s.addText([
    { text: "通用輸出（人人都拿得到）", options: { bullet: true, breakLine: true, fontSize: 13 } },
    { text: "不知道好壞", options: { bullet: true, breakLine: true, fontSize: 13, color: C.warning } },
    { text: "沒有迭代回饋", options: { bullet: true, breakLine: true, fontSize: 13, color: C.warning } },
    { text: "下一個版本就被取代", options: { bullet: true, fontSize: 13, color: C.warning } },
  ], {
    x: 5.5, y: 2.25, w: 3.5, h: 1.8,
    fontFace: FONT, color: C.nearBlack, valign: "top", margin: 0, paraSpaceAfter: 8
  });

  addInsight(s, "模型每六個月翻一次 — 但「判斷 × 工作流 × 受眾」每年只會變得更深");
}

// ============ SLIDE 9: 殺手級應用 ============
{
  const s = pres.addSlide();
  s.background = { color: C.white };
  addTitle(s, "最有潛力的個人槓桿：「研究 → 寫作 → 影響力」三段式自動化");

  // 場景流
  s.addShape(pres.shapes.RECTANGLE, { x: 0.8, y: 1.55, w: 2.5, h: 1.8, fill: { color: C.lightGray }, shadow: mkShadow() });
  s.addText("看到一個議題", {
    x: 0.8, y: 1.65, w: 2.5, h: 0.3,
    fontSize: 12, fontFace: FONT, color: C.primary, bold: true, align: "center", margin: 0
  });
  s.addText("30 分鐘", {
    x: 0.8, y: 2.0, w: 2.5, h: 0.6,
    fontSize: 32, fontFace: FONT, color: C.primary, bold: true, align: "center", valign: "middle", margin: 0
  });
  s.addText("研究 + 結構化", {
    x: 0.8, y: 2.7, w: 2.5, h: 0.3,
    fontSize: 11, fontFace: FONT, color: C.darkGray, align: "center", margin: 0
  });

  s.addText("→", {
    x: 3.3, y: 2.1, w: 0.5, h: 0.5,
    fontSize: 24, fontFace: FONT, color: C.primary, align: "center", valign: "middle", margin: 0
  });

  // AI 草稿
  s.addShape(pres.shapes.RECTANGLE, { x: 3.8, y: 1.55, w: 2.6, h: 1.8, fill: { color: C.lightGray }, shadow: mkShadow() });
  s.addText("AI 寫第一版", {
    x: 3.8, y: 1.65, w: 2.6, h: 0.3,
    fontSize: 12, fontFace: FONT, color: C.accent, bold: true, align: "center", margin: 0
  });
  s.addText("20 分鐘", {
    x: 3.8, y: 2.0, w: 2.6, h: 0.6,
    fontSize: 32, fontFace: FONT, color: C.accent, bold: true, align: "center", valign: "middle", margin: 0
  });
  s.addText("我審稿 + 改寫", {
    x: 3.8, y: 2.7, w: 2.6, h: 0.3,
    fontSize: 11, fontFace: FONT, color: C.darkGray, align: "center", margin: 0
  });

  s.addText("→", {
    x: 6.4, y: 2.1, w: 0.5, h: 0.5,
    fontSize: 24, fontFace: FONT, color: C.primary, align: "center", valign: "middle", margin: 0
  });

  // 發布觸達
  s.addShape(pres.shapes.RECTANGLE, { x: 6.9, y: 1.55, w: 2.3, h: 1.8, fill: { color: C.lightGray }, shadow: mkShadow() });
  s.addText("一鍵發布", {
    x: 6.9, y: 1.65, w: 2.3, h: 0.3,
    fontSize: 12, fontFace: FONT, color: C.accent, bold: true, align: "center", margin: 0
  });
  s.addText("Substack + X", {
    x: 6.9, y: 2.1, w: 2.3, h: 0.4,
    fontSize: 16, fontFace: FONT, color: C.nearBlack, bold: true, align: "center", margin: 0
  });
  s.addText("觸達 10 萬+\n訂閱者", {
    x: 6.9, y: 2.55, w: 2.3, h: 0.5,
    fontSize: 12, fontFace: FONT, color: C.darkGray, align: "center", margin: 0
  });

  // 底部標語
  s.addShape(pres.shapes.RECTANGLE, { x: 0.8, y: 3.65, w: 8.4, h: 0.4, fill: { color: C.primary } });
  s.addText("研究系統化 × AI 加速 × 受眾通路  =  個人前所未有的影響力槓桿", {
    x: 0.8, y: 3.65, w: 8.4, h: 0.4,
    fontSize: 13, fontFace: FONT, color: C.white, bold: true, align: "center", valign: "middle", margin: 0
  });

  addInsight(s, "從「一週寫一篇」到「一天寫一篇且品質更高」 — 純粹靠工作流結構，不靠加班", { y: 4.35 });
}

// ============ SLIDE 10: 路線圖 ============
{
  const s = pres.addSlide();
  s.background = { color: C.white };
  addTitle(s, "五年路線圖：從 AI Stack 到 Operator OS");

  const phases = [
    { year: "現在～1年", label: "AI Stack", desc: "建立個人\n工作流元件", color: C.primary },
    { year: "1～2年",   label: "Productize", desc: "工作流 →\n產品化", color: C.accent },
    { year: "2～3年",   label: "IP Marketplace", desc: "受眾主動\n上門", color: C.primary },
    { year: "3～5年",   label: "Operator OS", desc: "個人變成\n微型作業系統", color: C.nearBlack },
  ];

  s.addShape(pres.shapes.RECTANGLE, { x: 0.8, y: 2.65, w: 8.4, h: 0.04, fill: { color: C.lightGray } });

  phases.forEach((p, i) => {
    const x = 0.8 + i * 2.3;
    s.addShape(pres.shapes.OVAL, { x: x + 0.8, y: 2.5, w: 0.3, h: 0.3, fill: { color: p.color } });

    s.addText(p.year, {
      x, y: 1.5, w: 2.0, h: 0.3,
      fontSize: 12, fontFace: FONT, color: p.color, bold: true, align: "center", margin: 0
    });

    s.addShape(pres.shapes.RECTANGLE, { x, y: 3.0, w: 2.0, h: 1.5, fill: { color: C.offWhite }, shadow: mkShadow() });
    s.addShape(pres.shapes.RECTANGLE, { x, y: 3.0, w: 2.0, h: 0.05, fill: { color: p.color } });
    s.addText(p.label, {
      x, y: 3.2, w: 2.0, h: 0.35,
      fontSize: 14, fontFace: FONT, color: C.nearBlack, bold: true, align: "center", margin: 0
    });
    s.addText(p.desc, {
      x, y: 3.6, w: 2.0, h: 0.7,
      fontSize: 11, fontFace: FONT, color: C.darkGray, align: "center", margin: 0
    });
  });

  addInsight(s, "每個階段都以前一階段的工作流資產為基礎 → 護城河隨時間越來越深");
}

// ============ SLIDE 11: 結論 ============
{
  const s = pres.addSlide();
  s.background = { color: C.lightGray };
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.05, fill: { color: C.primary } });

  s.addShape(pres.shapes.OVAL, { x: -1, y: -1, w: 4, h: 4, fill: { color: C.white, transparency: 50 } });
  s.addShape(pres.shapes.OVAL, { x: 8, y: 3, w: 3.5, h: 3.5, fill: { color: C.white, transparency: 50 } });

  s.addText("Gary Chen 五年後可能會變成：", {
    x: 0.8, y: 0.8, w: 8.4, h: 0.5,
    fontSize: 18, fontFace: FONT, color: C.darkGray, margin: 0
  });
  s.addText("AI 時代的個人 Operator 範式", {
    x: 0.8, y: 1.4, w: 8.4, h: 0.9,
    fontSize: 36, fontFace: FONT, color: C.nearBlack, bold: true, margin: 0
  });

  s.addShape(pres.shapes.RECTANGLE, { x: 0.8, y: 2.5, w: 3, h: 0.04, fill: { color: C.accent } });

  s.addText("因為我同時擁有：", {
    x: 0.8, y: 2.8, w: 8.4, h: 0.4,
    fontSize: 16, fontFace: FONT, color: C.midGray, margin: 0
  });

  const points = [
    { num: "1", text: "判斷力 — 十年累積的領域知識與審美底線", color: C.primary },
    { num: "2", text: "工作流 — 把意圖系統化為可重用、可委派的元件", color: C.accent },
    { num: "3", text: "受眾資產 — 真實的回饋圈與分發通路", color: C.primary },
  ];

  points.forEach((p, i) => {
    const y = 3.35 + i * 0.55;
    s.addShape(pres.shapes.OVAL, { x: 0.8, y, w: 0.4, h: 0.4, fill: { color: p.color } });
    s.addText(p.num, {
      x: 0.8, y, w: 0.4, h: 0.4,
      fontSize: 14, fontFace: FONT, color: C.white, bold: true, align: "center", valign: "middle", margin: 0
    });
    s.addText(p.text, {
      x: 1.4, y, w: 7.8, h: 0.4,
      fontSize: 15, fontFace: FONT, color: C.nearBlack, valign: "middle", margin: 0
    });
  });

  addInsight(s, "AI 時代的稀缺不是「會用工具」— 而是「在工具之上有判斷、有品味、有受眾」。我現在的位置非常接近。", { y: 5.0 });
}

// Save
pres.writeFile({ fileName: "gary-chen-operator-roadmap.pptx" }).then(() => {
  console.log("Done!");
}).catch(err => console.error(err));
