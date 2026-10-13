# Slide Layout Templates（pptxgenjs 版型範本庫）

本文件取代對「教育訓練簡報 Slide Library.pptx」實體檔案的依賴，將 138 張版型的佈局邏輯萃取為 pptxgenjs 程式碼範本。所有版型均基於 `LAYOUT_16x9`（10" × 5.625"）、遵循 brand-colors.md 色彩規範。

共用基礎設施包含：品牌色常量、`addTitle`、`addInsight`（可選洞見條）、`addBottomBar`（多形態底部條）、`addChapterMarker`（輕量章節標記）、`addCard`、`addNumberCircle`。

---

## 目錄

1. [共用基礎設施](#共用基礎設施)
2. [封面與分隔頁](#封面與分隔頁)
3. [議程與目標](#議程與目標)
4. [數據圖表](#數據圖表)
5. [概念對比](#概念對比)
6. [流程與步驟](#流程與步驟)
7. [漏斗與篩選](#漏斗與篩選)
8. [時間軸與路線圖](#時間軸與路線圖)
9. [KPI 大數字卡片](#kpi-大數字卡片)
10. [表格與清單](#表格與清單)
11. [引言與洞見](#引言與洞見)
12. [團隊與總覽](#團隊與總覽)
13. [版型選擇速查表](#版型選擇速查表)

---

## 共用基礎設施

每份簡報的起手式。所有 helper 函式和品牌常量都在這裡定義，後續版型直接呼叫即可。

```javascript
const pptxgen = require("pptxgenjs");
const pres = new pptxgen();
pres.layout = "LAYOUT_16x9";

// ═══ 品牌色（嚴格只用這些，見 brand-colors.md）═══
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
const FONT = "Noto Sans TC";  // 統一字體（中英共用）

// ═══ 陰影（每次呼叫產生新物件，避免 pptxgenjs 共用陷阱）═══
const mkShadow = () => ({
  type: "outer", color: "000000", blur: 4, offset: 2, angle: 135, opacity: 0.07
});

// ═══ 標題列（每頁必備）═══
// 包含：品牌 Navy 頂條 + 結論句標題 + 分隔線
function addTitle(s, text, opts = {}) {
  // 頂部品牌 Navy 條（4-5px 高，全寬）
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 0, w: 10, h: 0.05, fill: { color: C.primary }
  });
  // 結論句標題（佔頁面頂端 15-20%）
  s.addText(text, {
    x: 0.8, y: 0.35, w: 8.4, h: 0.9,
    fontSize: 24, fontFace: FONT, color: C.nearBlack,
    bold: true, margin: 0, valign: "middle"
  });
  // 分隔線
  if (!opts.noDivider) {
    s.addShape(pres.shapes.RECTANGLE, {
      x: 0.8, y: 1.15, w: 8.4, h: 0.025,
      fill: { color: C.divider }
    });
  }
}

// ═══ 底部洞見條 Insight Bar（可選，非必備）═══
// 適用於：資訊量大的分析頁，需要一句話收束重點
// 不適用於：標題已包含洞見的頁面、封面、分隔頁、KPI 頁、UX 展示頁
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

// ═══ 底部多形態條（Insight Bar 以外的底部選項）═══
// style: "highlight" — Navy 底強調條，全寬貼底（page footer 風格），用於全簡報最關鍵結論
// style: "quote"     — 引言條，內縮浮島，帶引號和人物標注
// style: "source"    — 來源標注條，內縮浮島，灰色小字
function addBottomBar(s, text, opts = {}) {
  const style = opts.style || "highlight";
  // highlight 全寬貼底（5.625 - 0.6 = 5.025）；quote/source 維持浮島留白
  const y = opts.y !== undefined ? opts.y : (style === "highlight" ? 5.025 : 4.85);

  if (style === "highlight") {
    // Navy 底強調條（全寬貼底）
    s.addShape(pres.shapes.RECTANGLE, {
      x: 0, y: y, w: 10, h: 0.6,
      fill: { color: C.primary }
    });
    s.addText(text, {
      x: 0.8, y: y, w: 8.4, h: 0.6,
      fontSize: 13, fontFace: FONT, color: C.white,
      bold: true, valign: "middle", margin: 0
    });
  } else if (style === "quote") {
    // 引言條
    s.addShape(pres.shapes.RECTANGLE, {
      x: 0.8, y: y, w: 8.4, h: 0.6,
      fill: { color: C.lightGray }
    });
    s.addText(`"${opts.quoteText || text}"`, {
      x: 1.05, y: y, w: 6.5, h: 0.6,
      fontSize: 12, fontFace: FONT, color: C.darkGray,
      italic: true, valign: "middle", margin: 0
    });
    if (opts.author) {
      s.addText(`— ${opts.author}`, {
        x: 7.6, y: y, w: 1.6, h: 0.6,
        fontSize: 10, fontFace: FONT, color: C.midGray,
        valign: "middle", align: "right", margin: 0
      });
    }
  } else if (style === "source") {
    // 來源標注
    s.addText(`Source: ${text}`, {
      x: 0.8, y: y + 0.1, w: 8.4, h: 0.35,
      fontSize: 10, fontFace: FONT, color: C.midGray,
      valign: "middle", margin: 0
    });
  }
}

// ═══ 輕量章節標記（嵌入內容頁左上角）═══
// 不佔用一整頁，在內容頁左上角顯示當前章節編號
// 適用於 5 頁以上的簡報，幫助讀者導航
function addChapterMarker(s, num) {
  // 圈號
  s.addShape(pres.shapes.OVAL, {
    x: 0.25, y: 0.2, w: 0.38, h: 0.38,
    line: { color: C.darkGray, width: 1.2 }
  });
  s.addText(String(num), {
    x: 0.25, y: 0.2, w: 0.38, h: 0.38,
    fontSize: 14, fontFace: FONT, color: C.darkGray,
    bold: true, align: "center", valign: "middle", margin: 0
  });
}

// ═══ 卡片容器（帶頂邊條 + 陰影）═══
function addCard(s, x, y, w, h, opts = {}) {
  const borderColor = opts.borderColor || C.accent;
  s.addShape(pres.shapes.RECTANGLE, {
    x, y, w, h,
    fill: { color: opts.fill || C.offWhite },
    shadow: mkShadow()
  });
  s.addShape(pres.shapes.RECTANGLE, {
    x, y, w, h: 0.06,
    fill: { color: borderColor }
  });
}

// ═══ 編號圈（用於步驟、列表）═══
function addNumberCircle(s, x, y, num, color) {
  color = color || C.primary;
  s.addShape(pres.shapes.OVAL, {
    x, y, w: 0.45, h: 0.45,
    fill: { color }
  });
  s.addText(String(num), {
    x, y, w: 0.45, h: 0.45,
    fontSize: 16, fontFace: FONT, color: C.white,
    bold: true, align: "center", valign: "middle", margin: 0
  });
}
```

---

## 封面與分隔頁

### 封面頁（Cover）
原 Slide Library slide1。淺灰底 + 品牌 Navy 條 + 大標題 + 副標題。

```javascript
function slideCover(title, subtitle, date) {
  const s = pres.addSlide();
  s.background = { color: C.lightGray };
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 0, w: 10, h: 0.05, fill: { color: C.primary }
  });
  // 裝飾圓形（增加視覺層次）
  s.addShape(pres.shapes.OVAL, {
    x: 7.5, y: -0.8, w: 4, h: 4,
    fill: { color: C.white, transparency: 60 }
  });
  // 主標題
  s.addText(title, {
    x: 0.8, y: 1.2, w: 8, h: 1.4,
    fontSize: 40, fontFace: FONT, color: C.nearBlack,
    bold: true, margin: 0
  });
  // 紅色裝飾線
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.8, y: 2.85, w: 3, h: 0.04, fill: { color: C.primary }
  });
  // 副標題
  if (subtitle) {
    s.addText(subtitle, {
      x: 0.8, y: 3.15, w: 8, h: 0.6,
      fontSize: 20, fontFace: FONT, color: C.darkGray,
      italic: true, margin: 0
    });
  }
  // 日期
  if (date) {
    s.addText(date, {
      x: 0.8, y: 4.8, w: 4, h: 0.4,
      fontSize: 12, fontFace: FONT, color: C.midGray, margin: 0
    });
  }
  return s;
}
```

### 章節分隔頁（Section Divider）
原 Slide Library slide3/12/26/44/77。深色底或淺灰底 + 大字章節標題。

```javascript
function slideSectionDivider(sectionNum, sectionTitle) {
  const s = pres.addSlide();
  s.background = { color: C.lightGray };
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 0, w: 10, h: 0.05, fill: { color: C.primary }
  });
  // 章節編號
  if (sectionNum) {
    s.addText(String(sectionNum).padStart(2, "0"), {
      x: 0.8, y: 1.5, w: 2, h: 1,
      fontSize: 60, fontFace: FONT, color: C.primary,
      bold: true, margin: 0
    });
  }
  // 章節標題
  s.addText(sectionTitle, {
    x: 0.8, y: 2.5, w: 8.4, h: 1.5,
    fontSize: 32, fontFace: FONT, color: C.nearBlack,
    bold: true, margin: 0
  });
  return s;
}
```

---

## 議程與目標

### 議程列表（Agenda List）
原 slide4。左側議程清單 + 右側詳細說明。

```javascript
// items = [{ title: "議題", detail: "說明" }, ...]
function slideAgendaList(pageTitle, items) {
  const s = pres.addSlide();
  s.background = { color: C.white };
  addTitle(s, pageTitle);

  items.forEach((item, i) => {
    const y = 1.55 + i * 0.6;
    // 編號
    addNumberCircle(s, 0.8, y, i + 1, C.primary);
    // 議題名稱
    s.addText(item.title, {
      x: 1.4, y: y, w: 3.5, h: 0.45,
      fontSize: 14, fontFace: FONT, color: C.nearBlack,
      bold: true, valign: "middle", margin: 0
    });
    // 說明（右側）
    if (item.detail) {
      s.addText(item.detail, {
        x: 5.2, y: y, w: 4, h: 0.45,
        fontSize: 12, fontFace: FONT, color: C.darkGray,
        valign: "middle", margin: 0
      });
    }
  });
  return s;
}
```

### 議程網格（Agenda Grid）
原 slide6。3×3 編號網格，每格一個議程項。

```javascript
// items = ["議程一", "議程二", ...] (最多 9 項)
function slideAgendaGrid(pageTitle, items) {
  const s = pres.addSlide();
  s.background = { color: C.white };
  addTitle(s, pageTitle);

  const cols = 3, startX = 0.8, startY = 1.55;
  const cardW = 2.7, cardH = 1.2, gapX = 0.15, gapY = 0.15;

  items.forEach((item, i) => {
    const col = i % cols;
    const row = Math.floor(i / cols);
    const x = startX + col * (cardW + gapX);
    const y = startY + row * (cardH + gapY);

    addCard(s, x, y, cardW, cardH, { borderColor: C.primary });
    // 編號
    s.addText(String(i + 1), {
      x: x + 0.15, y: y + 0.2, w: 0.4, h: 0.4,
      fontSize: 20, fontFace: FONT, color: C.primary,
      bold: true, margin: 0
    });
    // 文字
    s.addText(item, {
      x: x + 0.6, y: y + 0.2, w: cardW - 0.8, h: 0.8,
      fontSize: 12, fontFace: FONT, color: C.nearBlack,
      valign: "top", margin: 0
    });
  });
  return s;
}
```

### 工作坊議程表（Workshop Agenda）
原 slide5。時間-主題-成果三欄表格。

```javascript
// rows = [{ time: "09:00-10:00", topic: "主題", outcome: "預期成果" }, ...]
function slideWorkshopAgenda(pageTitle, rows) {
  const s = pres.addSlide();
  s.background = { color: C.white };
  addTitle(s, pageTitle);

  // 表頭
  const headers = ["時間", "主題", "預期成果"];
  const colX = [0.8, 2.8, 6.5];
  const colW = [1.8, 3.5, 2.7];
  headers.forEach((h, i) => {
    s.addText(h, {
      x: colX[i], y: 1.5, w: colW[i], h: 0.35,
      fontSize: 12, fontFace: FONT, color: C.white,
      bold: true, valign: "middle", margin: 0,
      fill: { color: C.primary }
    });
  });

  rows.forEach((row, i) => {
    const y = 1.9 + i * 0.5;
    const bgColor = i % 2 === 0 ? C.white : C.lightGray;
    [row.time, row.topic, row.outcome].forEach((text, j) => {
      s.addShape(pres.shapes.RECTANGLE, {
        x: colX[j], y, w: colW[j], h: 0.45,
        fill: { color: bgColor }
      });
      s.addText(text, {
        x: colX[j] + 0.1, y, w: colW[j] - 0.2, h: 0.45,
        fontSize: 11, fontFace: FONT,
        color: j === 0 ? C.primary : C.nearBlack,
        bold: j === 0, valign: "middle", margin: 0
      });
    });
  });
  return s;
}
```

### 後續步驟（Next Steps）
原 slide10/11。5 步驟卡片式佈局。

```javascript
// steps = [{ name: "步驟名", detail: "說明", owner: "負責人", deadline: "日期" }, ...]
function slideNextSteps(pageTitle, steps) {
  const s = pres.addSlide();
  s.background = { color: C.white };
  addTitle(s, pageTitle);

  steps.forEach((step, i) => {
    const y = 1.55 + i * 0.7;
    // 編號圈
    addNumberCircle(s, 0.8, y + 0.05, i + 1, C.primary);
    // 步驟名稱
    s.addText(step.name, {
      x: 1.4, y: y, w: 2.5, h: 0.35,
      fontSize: 14, fontFace: FONT, color: C.nearBlack,
      bold: true, valign: "middle", margin: 0
    });
    // 說明
    s.addText(step.detail, {
      x: 4.0, y: y, w: 3.5, h: 0.55,
      fontSize: 11, fontFace: FONT, color: C.darkGray,
      valign: "top", margin: 0
    });
    // 負責人 + 截止日期（如有）
    if (step.owner || step.deadline) {
      const meta = [step.owner, step.deadline].filter(Boolean).join(" | ");
      s.addText(meta, {
        x: 7.6, y: y, w: 1.6, h: 0.35,
        fontSize: 10, fontFace: FONT, color: C.midGray,
        valign: "middle", margin: 0
      });
    }
    // 分隔線
    if (i < steps.length - 1) {
      s.addShape(pres.shapes.RECTANGLE, {
        x: 0.8, y: y + 0.6, w: 8.4, h: 0.015,
        fill: { color: C.divider }
      });
    }
  });
  return s;
}
```

---

## 數據圖表

### 柱狀圖對比（Bar Chart）
原 slide13。多柱狀圖 + 類別標籤。

```javascript
function slideBarChart(pageTitle, chartData, insightText) {
  const s = pres.addSlide();
  s.background = { color: C.white };
  addTitle(s, pageTitle);

  s.addChart(pres.charts.BAR, chartData, {
    x: 0.8, y: 1.5, w: 8.4, h: 2.8, barDir: "col",
    chartColors: [C.primary, C.accent, C.midGray],
    chartArea: { fill: { color: C.white } },
    catAxisLabelColor: C.darkGray,
    valAxisLabelColor: C.darkGray,
    valGridLine: { color: C.lightGray, size: 0.5 },
    catGridLine: { style: "none" },
    showValue: true,
    dataLabelPosition: "outEnd",
    dataLabelColor: C.nearBlack,
    showLegend: chartData.length > 1,
    legendPos: "b",
  });

  if (insightText) addInsight(s, insightText);
  return s;
}
```

### 2×2 象限矩陣（Quadrant）
原 slide20。四象限矩陣分析。

```javascript
// quadrants = [{ label, items }, ...] 順序：左上、右上、左下、右下
// axes = { x: "X 軸標籤", y: "Y 軸標籤" }
function slideQuadrant(pageTitle, quadrants, axes, insightText) {
  const s = pres.addSlide();
  s.background = { color: C.white };
  addTitle(s, pageTitle);

  const qW = 3.8, qH = 1.6;
  const positions = [
    { x: 1.6, y: 1.5 },  // 左上
    { x: 5.6, y: 1.5 },  // 右上
    { x: 1.6, y: 3.2 },  // 左下
    { x: 5.6, y: 3.2 },  // 右下
  ];
  const colors = [C.primary, C.accent, C.midGray, C.darkGray];

  // 軸線
  s.addShape(pres.shapes.RECTANGLE, {
    x: 1.5, y: 3.1, w: 7.8, h: 0.02, fill: { color: C.divider }
  });
  s.addShape(pres.shapes.RECTANGLE, {
    x: 5.5, y: 1.4, w: 0.02, h: 3.5, fill: { color: C.divider }
  });

  // 軸標籤
  if (axes) {
    s.addText(axes.y || "", {
      x: 0.3, y: 2.5, w: 1, h: 0.4,
      fontSize: 10, fontFace: FONT, color: C.midGray,
      align: "center", margin: 0
    });
    s.addText(axes.x || "", {
      x: 4.5, y: 4.9, w: 2, h: 0.3,
      fontSize: 10, fontFace: FONT, color: C.midGray,
      align: "center", margin: 0
    });
  }

  quadrants.forEach((q, i) => {
    const pos = positions[i];
    s.addShape(pres.shapes.RECTANGLE, {
      x: pos.x, y: pos.y, w: qW, h: qH,
      fill: { color: C.lightGray, transparency: 50 }
    });
    s.addText(q.label, {
      x: pos.x + 0.15, y: pos.y + 0.1, w: qW - 0.3, h: 0.35,
      fontSize: 13, fontFace: FONT, color: colors[i],
      bold: true, margin: 0
    });
    if (q.items) {
      s.addText(
        q.items.map((item, j) => ({
          text: item,
          options: { bullet: true, breakLine: j < q.items.length - 1, fontSize: 11 }
        })),
        {
          x: pos.x + 0.15, y: pos.y + 0.5, w: qW - 0.3, h: qH - 0.6,
          fontFace: FONT, color: C.darkGray, valign: "top", margin: 0
        }
      );
    }
  });

  if (insightText) addInsight(s, insightText);
  return s;
}
```

### 風險矩陣（Risk Matrix）
原 slide14。左側風險清單 + 右側機率/影響矩陣。

```javascript
// risks = [{ name, probability: "高/中/低", impact: "高/中/低", desc }, ...]
function slideRiskMatrix(pageTitle, risks, insightText) {
  const s = pres.addSlide();
  s.background = { color: C.white };
  addTitle(s, pageTitle);

  // 左側風險清單
  risks.forEach((r, i) => {
    const y = 1.55 + i * 0.6;
    addNumberCircle(s, 0.8, y, i + 1, C.primary);
    s.addText(r.name, {
      x: 1.4, y, w: 3, h: 0.3,
      fontSize: 13, fontFace: FONT, color: C.nearBlack,
      bold: true, margin: 0
    });
    s.addText(r.desc || "", {
      x: 1.4, y: y + 0.25, w: 3, h: 0.3,
      fontSize: 10, fontFace: FONT, color: C.darkGray, margin: 0
    });
  });

  // 右側矩陣網格
  const gridX = 5.5, gridY = 1.5, cellW = 1.3, cellH = 0.9;
  const labels = ["低", "中", "高"];
  // 軸標籤
  s.addText("影響 →", {
    x: gridX + 0.5, y: gridY + 3 * cellH + 0.1, w: 3, h: 0.3,
    fontSize: 10, fontFace: FONT, color: C.midGray, margin: 0
  });
  s.addText("機率 →", {
    x: gridX - 0.8, y: gridY + 0.5, w: 0.7, h: 2,
    fontSize: 10, fontFace: FONT, color: C.midGray,
    margin: 0, valign: "middle"
  });
  // 網格
  for (let r = 0; r < 3; r++) {
    for (let c = 0; c < 3; c++) {
      const severity = r + c; // 0=低低, 4=高高
      const bgColor = severity >= 3 ? "FBDDD8" : severity >= 2 ? "E5EAF0" : C.lightGray;
      s.addShape(pres.shapes.RECTANGLE, {
        x: gridX + c * cellW, y: gridY + (2 - r) * cellH,
        w: cellW, h: cellH,
        fill: { color: bgColor },
        line: { color: C.divider, width: 0.5 }
      });
    }
  }
  // 標記風險位置
  risks.forEach((r, i) => {
    const probMap = { "低": 0, "中": 1, "高": 2 };
    const impMap = { "低": 0, "中": 1, "高": 2 };
    const cx = gridX + (impMap[r.impact] || 1) * cellW + cellW / 2 - 0.15;
    const cy = gridY + (2 - (probMap[r.probability] || 1)) * cellH + cellH / 2 - 0.15;
    addNumberCircle(s, cx, cy, i + 1, C.primary);
  });

  if (insightText) addInsight(s, insightText);
  return s;
}
```

---

## 概念對比

### 通用左右分欄（Split Layout）
通用的左右不對稱分欄佈局，適用範圍遠超「對比」：截圖+文字、圖表+解讀、象限圖+Takeaway、組織架構+治理原則等。這是與三段式平級的核心佈局。

```javascript
// leftContent 和 rightContent 各自是一個函式，接收 (s, x, y, w, h) 繪製區域內容
// splitRatio: 左欄佔比，預設 0.5（50:50），可設 0.4（40:60）或 0.6（60:40）
function slideSplitLayout(pageTitle, leftContent, rightContent, opts = {}) {
  const s = pres.addSlide();
  s.background = { color: opts.bg || C.white };
  addTitle(s, pageTitle);
  if (opts.chapterNum) addChapterMarker(s, opts.chapterNum);

  const splitRatio = opts.splitRatio || 0.5;
  const contentY = 1.45;
  const contentH = opts.bottomText ? 3.2 : 3.7;
  const gap = 0.3;
  const totalW = 8.4;
  const leftW = totalW * splitRatio - gap / 2;
  const rightW = totalW * (1 - splitRatio) - gap / 2;
  const leftX = 0.8;
  const rightX = leftX + leftW + gap;

  // 繪製左右欄內容（由呼叫者自行定義）
  if (typeof leftContent === "function") {
    leftContent(s, leftX, contentY, leftW, contentH);
  }
  if (typeof rightContent === "function") {
    rightContent(s, rightX, contentY, rightW, contentH);
  }

  // 可選：中間分隔線
  if (opts.divider !== false) {
    s.addShape(pres.shapes.RECTANGLE, {
      x: leftX + leftW + gap / 2 - 0.01, y: contentY,
      w: 0.02, h: contentH,
      fill: { color: C.divider }
    });
  }

  // 底部區域（可選）
  if (opts.bottomText) {
    if (opts.bottomStyle === "highlight") {
      addBottomBar(s, opts.bottomText, { style: "highlight" });
    } else if (opts.bottomStyle === "source") {
      addBottomBar(s, opts.bottomText, { style: "source" });
    } else if (opts.bottomStyle === "quote") {
      addBottomBar(s, opts.bottomText, { style: "quote", author: opts.bottomAuthor });
    } else {
      addInsight(s, opts.bottomText);
    }
  }
  return s;
}
```

### 左右對比（Comparison / Old vs New）
原 slide17/33/34/37。兩欄對比佈局，適用 Before/After、新舊對比、方案比較。

```javascript
// left = { title, items: [], color }
// right = { title, items: [], color }
function slideComparison(pageTitle, left, right, insightText) {
  const s = pres.addSlide();
  s.background = { color: C.white };
  addTitle(s, pageTitle);

  const panels = [
    { data: left,  x: 0.8, color: left.color || C.midGray },
    { data: right, x: 5.2, color: right.color || C.primary },
  ];

  panels.forEach(p => {
    addCard(s, p.x, 1.55, 4, 2.8, { borderColor: p.color });
    // 標題
    s.addText(p.data.title, {
      x: p.x + 0.25, y: 1.75, w: 3.5, h: 0.35,
      fontSize: 15, fontFace: FONT, color: p.color,
      bold: true, margin: 0
    });
    // 項目
    if (p.data.items) {
      s.addText(
        p.data.items.map((item, j) => ({
          text: item,
          options: { bullet: true, breakLine: j < p.data.items.length - 1, fontSize: 12 }
        })),
        {
          x: p.x + 0.25, y: 2.2, w: 3.5, h: 1.8,
          fontFace: FONT, color: C.nearBlack,
          valign: "top", margin: 0, paraSpaceAfter: 6
        }
      );
    }
  });

  // 中間箭頭
  s.addText("→", {
    x: 4.6, y: 2.6, w: 0.8, h: 0.5,
    fontSize: 24, fontFace: FONT, color: C.primary,
    align: "center", valign: "middle", margin: 0
  });

  if (insightText) addInsight(s, insightText);
  return s;
}
```

### From/To 多項對照（From/To Multi）
原 slide34。多行「From → To」變化對比。

```javascript
// pairs = [{ from: "舊做法", to: "新做法" }, ...]
function slideFromTo(pageTitle, subtitle, pairs, insightText) {
  const s = pres.addSlide();
  s.background = { color: C.white };
  addTitle(s, pageTitle);

  if (subtitle) {
    s.addText(subtitle, {
      x: 0.8, y: 1.4, w: 8.4, h: 0.35,
      fontSize: 14, fontFace: FONT, color: C.darkGray,
      bold: true, margin: 0
    });
  }

  // 欄標題
  s.addText("From", {
    x: 0.8, y: 1.85, w: 3.8, h: 0.3,
    fontSize: 12, fontFace: FONT, color: C.midGray,
    bold: true, margin: 0
  });
  s.addText("To", {
    x: 5.5, y: 1.85, w: 3.8, h: 0.3,
    fontSize: 12, fontFace: FONT, color: C.primary,
    bold: true, margin: 0
  });

  pairs.forEach((pair, i) => {
    const y = 2.25 + i * 0.55;
    // From
    s.addShape(pres.shapes.RECTANGLE, {
      x: 0.8, y, w: 3.8, h: 0.45,
      fill: { color: C.lightGray }
    });
    s.addText(pair.from, {
      x: 1.0, y, w: 3.4, h: 0.45,
      fontSize: 11, fontFace: FONT, color: C.darkGray,
      valign: "middle", margin: 0
    });
    // 箭頭
    s.addText("→", {
      x: 4.7, y, w: 0.6, h: 0.45,
      fontSize: 16, fontFace: FONT, color: C.primary,
      align: "center", valign: "middle", margin: 0
    });
    // To
    s.addShape(pres.shapes.RECTANGLE, {
      x: 5.5, y, w: 3.8, h: 0.45,
      fill: { color: C.offWhite }
    });
    s.addText(pair.to, {
      x: 5.7, y, w: 3.4, h: 0.45,
      fontSize: 11, fontFace: FONT, color: C.nearBlack,
      bold: true, valign: "middle", margin: 0
    });
  });

  if (insightText) addInsight(s, insightText);
  return s;
}
```

---

## 流程與步驟

### 三步驟線性流程（3 Flow Linear）
原 slide46/47。三欄等寬流程卡片 + 箭頭連接。

```javascript
// steps = [{ num, title, items: [] }, ...]
function slide3Flow(pageTitle, steps, insightText) {
  const s = pres.addSlide();
  s.background = { color: C.white };
  addTitle(s, pageTitle);

  const startX = 0.8, cardW = 2.6, gap = 0.3;

  steps.forEach((step, i) => {
    const x = startX + i * (cardW + gap);
    // 步驟編號圓形 + 標題
    s.addShape(pres.shapes.RECTANGLE, {
      x, y: 1.55, w: cardW, h: 0.45,
      fill: { color: C.primary }
    });
    s.addText(`${step.num || i + 1}. ${step.title}`, {
      x, y: 1.55, w: cardW, h: 0.45,
      fontSize: 13, fontFace: FONT, color: C.white,
      bold: true, align: "center", valign: "middle", margin: 0
    });
    // 內容卡片
    addCard(s, x, 2.1, cardW, 2.2, { borderColor: C.accent });
    if (step.items) {
      s.addText(
        step.items.map((item, j) => ({
          text: item,
          options: { bullet: true, breakLine: j < step.items.length - 1, fontSize: 11 }
        })),
        {
          x: x + 0.15, y: 2.3, w: cardW - 0.3, h: 1.8,
          fontFace: FONT, color: C.nearBlack,
          valign: "top", margin: 0, paraSpaceAfter: 4
        }
      );
    }
    // 箭頭
    if (i < steps.length - 1) {
      s.addText("→", {
        x: x + cardW, y: 2.8, w: gap, h: 0.5,
        fontSize: 18, fontFace: FONT, color: C.primary,
        align: "center", valign: "middle", margin: 0
      });
    }
  });

  if (insightText) addInsight(s, insightText);
  return s;
}
```

### 五步驟路徑（5 Steps）
原 slide80/81。五步驟水平路徑，適用於執行計畫。

```javascript
// steps = [{ label, desc }, ...]
function slide5Steps(pageTitle, steps, insightText) {
  const s = pres.addSlide();
  s.background = { color: C.white };
  addTitle(s, pageTitle);

  const startX = 0.8, stepW = 1.68, gap = 0.1;

  // 連接線
  s.addShape(pres.shapes.RECTANGLE, {
    x: startX + 0.3, y: 2.15, w: 8, h: 0.03,
    fill: { color: C.divider }
  });

  steps.forEach((step, i) => {
    const x = startX + i * (stepW + gap);
    // 節點圓
    addNumberCircle(s, x + stepW / 2 - 0.22, 1.95, i + 1, C.primary);
    // 卡片
    addCard(s, x, 2.5, stepW, 1.8, { borderColor: C.primary });
    s.addText(step.label, {
      x: x + 0.1, y: 2.65, w: stepW - 0.2, h: 0.35,
      fontSize: 12, fontFace: FONT, color: C.nearBlack,
      bold: true, align: "center", margin: 0
    });
    s.addText(step.desc || "", {
      x: x + 0.1, y: 3.05, w: stepW - 0.2, h: 1.0,
      fontSize: 10, fontFace: FONT, color: C.darkGray,
      align: "center", margin: 0
    });
  });

  if (insightText) addInsight(s, insightText);
  return s;
}
```

### 流程深入展開（Process Deep-Dive）
原 slide56-58。四步驟橫向流程 + 每步詳細說明。

```javascript
// steps = [{ label: "A", title, detail }, ...]
function slideProcessDeepDive(pageTitle, steps, insightText) {
  const s = pres.addSlide();
  s.background = { color: C.white };
  addTitle(s, pageTitle);

  // 橫向流程線
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.8, y: 2.35, w: 8.4, h: 0.025, fill: { color: C.divider }
  });

  const stepW = 8.4 / steps.length;
  steps.forEach((step, i) => {
    const x = 0.8 + i * stepW;
    // 步驟標籤
    s.addShape(pres.shapes.OVAL, {
      x: x + stepW / 2 - 0.25, y: 1.85, w: 0.5, h: 0.5,
      fill: { color: C.primary }
    });
    s.addText(step.label || String.fromCharCode(65 + i), {
      x: x + stepW / 2 - 0.25, y: 1.85, w: 0.5, h: 0.5,
      fontSize: 16, fontFace: FONT, color: C.white,
      bold: true, align: "center", valign: "middle", margin: 0
    });
    // 標題 + 說明
    s.addText(step.title, {
      x, y: 2.55, w: stepW - 0.1, h: 0.35,
      fontSize: 13, fontFace: FONT, color: C.nearBlack,
      bold: true, margin: 0
    });
    s.addText(step.detail || "", {
      x, y: 2.95, w: stepW - 0.1, h: 1.4,
      fontSize: 10, fontFace: FONT, color: C.darkGray,
      valign: "top", margin: 0
    });
  });

  if (insightText) addInsight(s, insightText);
  return s;
}
```

### 六支柱框架（Six Pillars）
原 slide67。六個等寬支柱並排。

```javascript
// pillars = [{ title, desc }, ...] (最多 6 個)
function slidePillars(pageTitle, pillars, insightText) {
  const s = pres.addSlide();
  s.background = { color: C.white };
  addTitle(s, pageTitle);

  const n = pillars.length;
  const totalW = 8.4, gap = 0.15;
  const pillarW = (totalW - (n - 1) * gap) / n;

  pillars.forEach((p, i) => {
    const x = 0.8 + i * (pillarW + gap);
    addCard(s, x, 1.55, pillarW, 3.0, { borderColor: C.primary });
    // 編號
    addNumberCircle(s, x + 0.1, 1.7, i + 1, C.primary);
    // 標題
    s.addText(p.title, {
      x: x + 0.6, y: 1.7, w: pillarW - 0.75, h: 0.4,
      fontSize: 12, fontFace: FONT, color: C.nearBlack,
      bold: true, valign: "middle", margin: 0
    });
    // 說明
    s.addText(p.desc || "", {
      x: x + 0.15, y: 2.25, w: pillarW - 0.3, h: 2.1,
      fontSize: 10, fontFace: FONT, color: C.darkGray,
      valign: "top", margin: 0
    });
  });

  if (insightText) addInsight(s, insightText);
  return s;
}
```

---

## 漏斗與篩選

### 四步漏斗（4 Step Funnel）
原 slide72/76。從寬到窄的漏斗圖。

```javascript
// steps = [{ label, value, desc }, ...]
function slideFunnel(pageTitle, steps, insightText) {
  const s = pres.addSlide();
  s.background = { color: C.white };
  addTitle(s, pageTitle);

  const n = steps.length;
  const maxW = 8.4, minW = 2.5;

  steps.forEach((step, i) => {
    const w = maxW - i * ((maxW - minW) / (n - 1));
    const x = 0.8 + (8.4 - w) / 2; // 置中
    const y = 1.55 + i * 0.85;
    const colors = [C.primary, C.accent, C.midGray, C.nearBlack];

    s.addShape(pres.shapes.RECTANGLE, {
      x, y, w, h: 0.7,
      fill: { color: colors[i % colors.length] }
    });
    // 步驟標題（白字）
    s.addText(`${step.label}${step.value ? "  " + step.value : ""}`, {
      x, y, w, h: 0.7,
      fontSize: 14, fontFace: FONT, color: C.white,
      bold: true, align: "center", valign: "middle", margin: 0
    });
    // 描述（右側）
    if (step.desc) {
      s.addText(step.desc, {
        x: x + w + 0.2, y, w: 10 - (x + w) - 0.5, h: 0.7,
        fontSize: 10, fontFace: FONT, color: C.darkGray,
        valign: "middle", margin: 0
      });
    }
  });

  if (insightText) addInsight(s, insightText);
  return s;
}
```

---

## 時間軸與路線圖

### 時間軸（Timeline）
原 slide98。水平時間軸 + 里程碑。

```javascript
// milestones = [{ date, title, desc }, ...]
function slideTimeline(pageTitle, milestones, insightText) {
  const s = pres.addSlide();
  s.background = { color: C.white };
  addTitle(s, pageTitle);

  const n = milestones.length;
  const lineY = 2.8, startX = 1.0, endX = 9.0;
  const totalW = endX - startX;

  // 時間軸線
  s.addShape(pres.shapes.RECTANGLE, {
    x: startX, y: lineY, w: totalW, h: 0.03, fill: { color: C.divider }
  });

  milestones.forEach((m, i) => {
    const x = startX + i * (totalW / (n - 1 || 1));
    const isAbove = i % 2 === 0;

    // 節點圓
    s.addShape(pres.shapes.OVAL, {
      x: x - 0.12, y: lineY - 0.12, w: 0.24, h: 0.24,
      fill: { color: C.primary }
    });

    // 垂直連接線
    const connY = isAbove ? lineY - 0.8 : lineY + 0.25;
    s.addShape(pres.shapes.RECTANGLE, {
      x: x - 0.01, y: isAbove ? connY + 0.6 : lineY + 0.12,
      w: 0.02, h: 0.3,
      fill: { color: C.divider }
    });

    // 日期
    s.addText(m.date, {
      x: x - 0.7, y: isAbove ? 1.55 : 3.1, w: 1.4, h: 0.3,
      fontSize: 10, fontFace: FONT, color: C.primary,
      bold: true, align: "center", margin: 0
    });
    // 標題
    s.addText(m.title, {
      x: x - 0.7, y: isAbove ? 1.85 : 3.4, w: 1.4, h: 0.3,
      fontSize: 11, fontFace: FONT, color: C.nearBlack,
      bold: true, align: "center", margin: 0
    });
    // 說明
    if (m.desc) {
      s.addText(m.desc, {
        x: x - 0.7, y: isAbove ? 2.15 : 3.7, w: 1.4, h: 0.5,
        fontSize: 9, fontFace: FONT, color: C.darkGray,
        align: "center", margin: 0
      });
    }
  });

  if (insightText) addInsight(s, insightText);
  return s;
}
```

### 四階段路線圖（4 Horizons）
原 slide78/79。階段式橫向展開 + 卡片。

```javascript
// phases = [{ year, label, desc, color }, ...]
function slideHorizons(pageTitle, phases, insightText) {
  const s = pres.addSlide();
  s.background = { color: C.white };
  addTitle(s, pageTitle);

  // 時間軸線
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.8, y: 2.15, w: 8.4, h: 0.03, fill: { color: C.lightGray }
  });

  phases.forEach((p, i) => {
    const x = 0.8 + i * 2.3;
    const color = p.color || C.primary;
    // 節點圓
    s.addShape(pres.shapes.OVAL, {
      x: x + 0.8, y: 2.0, w: 0.3, h: 0.3,
      fill: { color }
    });
    // 年份標籤
    s.addText(p.year, {
      x, y: 1.5, w: 2.0, h: 0.3,
      fontSize: 12, fontFace: FONT, color,
      bold: true, align: "center", margin: 0
    });
    // 卡片
    addCard(s, x, 2.5, 2.0, 1.8, { borderColor: color });
    s.addText(p.label, {
      x, y: 2.7, w: 2.0, h: 0.35,
      fontSize: 13, fontFace: FONT, color: C.nearBlack,
      bold: true, align: "center", margin: 0
    });
    s.addText(p.desc || "", {
      x, y: 3.1, w: 2.0, h: 1.0,
      fontSize: 10, fontFace: FONT, color: C.darkGray,
      align: "center", margin: 0
    });
  });

  if (insightText) addInsight(s, insightText);
  return s;
}
```

---

## KPI 大數字卡片

### KPI 卡片排列
原設計取自 golden-standard.md。大數字 + 標籤 + 卡片。

```javascript
// kpis = [{ value: "600 元", label: "用戶終身價值", color }, ...]
function slideKPI(pageTitle, kpis, insightText) {
  const s = pres.addSlide();
  s.background = { color: C.white };
  addTitle(s, pageTitle);

  const n = kpis.length;
  const totalW = 8.4, gap = 0.2;
  const cardW = (totalW - (n - 1) * gap) / n;

  kpis.forEach((kpi, i) => {
    const x = 0.8 + i * (cardW + gap);
    const color = kpi.color || C.primary;

    addCard(s, x, 1.55, cardW, 2.5, { borderColor: color });
    // 大數字
    s.addText(kpi.value, {
      x, y: 1.9, w: cardW, h: 1.0,
      fontSize: 36, fontFace: FONT, color,
      bold: true, align: "center", valign: "middle", margin: 0
    });
    // 標籤
    s.addText(kpi.label, {
      x, y: 3.0, w: cardW, h: 0.5,
      fontSize: 12, fontFace: FONT, color: C.darkGray,
      align: "center", valign: "top", margin: 0
    });
    // 副標籤
    if (kpi.sub) {
      s.addText(kpi.sub, {
        x, y: 3.45, w: cardW, h: 0.4,
        fontSize: 10, fontFace: FONT, color: C.midGray,
        align: "center", margin: 0
      });
    }
  });

  if (insightText) addInsight(s, insightText);
  return s;
}
```

---

## 表格與清單

### 資料表格（Data Table）
原 slide131/132。多欄表格 + 斑馬紋。

```javascript
// headers = ["欄位1", "欄位2", ...]
// rows = [["值1", "值2", ...], ...]
function slideTable(pageTitle, headers, rows, insightText) {
  const s = pres.addSlide();
  s.background = { color: C.white };
  addTitle(s, pageTitle);

  const tableData = [];
  // 表頭
  tableData.push(
    headers.map(h => ({
      text: h,
      options: {
        fill: { color: C.primary }, color: C.white,
        bold: true, fontSize: 12, fontFace: FONT,
        valign: "middle"
      }
    }))
  );
  // 資料列
  rows.forEach((row, i) => {
    tableData.push(
      row.map(cell => ({
        text: String(cell),
        options: {
          fill: { color: i % 2 === 0 ? C.white : C.lightGray },
          color: C.nearBlack, fontSize: 11, fontFace: FONT,
          valign: "middle"
        }
      }))
    );
  });

  const colW = headers.map(() => 8.4 / headers.length);
  s.addTable(tableData, {
    x: 0.8, y: 1.5, w: 8.4,
    colW,
    border: { pt: 0.5, color: C.divider },
    autoPage: false,
  });

  if (insightText) addInsight(s, insightText);
  return s;
}
```

### 專案狀態追蹤（Project Status）
原 slide133。多欄狀態表 + 紅黃綠燈號。

```javascript
// projects = [{ name, phase, status: "green/yellow/red", owner, notes }, ...]
function slideProjectStatus(pageTitle, projects, insightText) {
  const s = pres.addSlide();
  s.background = { color: C.white };
  addTitle(s, pageTitle);

  const headers = ["專案", "階段", "狀態", "負責人", "備註"];
  const colW = [2.2, 1.5, 0.8, 1.2, 2.7];
  const statusColors = { green: "4CAF50", yellow: "FFC107", red: C.warning };

  // 表頭
  headers.forEach((h, i) => {
    let x = 0.8;
    for (let j = 0; j < i; j++) x += colW[j];
    s.addShape(pres.shapes.RECTANGLE, {
      x, y: 1.5, w: colW[i], h: 0.4,
      fill: { color: C.primary }
    });
    s.addText(h, {
      x, y: 1.5, w: colW[i], h: 0.4,
      fontSize: 11, fontFace: FONT, color: C.white,
      bold: true, align: "center", valign: "middle", margin: 0
    });
  });

  // 資料列
  projects.forEach((p, rowIdx) => {
    const y = 1.95 + rowIdx * 0.5;
    const bg = rowIdx % 2 === 0 ? C.white : C.lightGray;
    const vals = [p.name, p.phase, null, p.owner, p.notes];

    let x = 0.8;
    vals.forEach((val, i) => {
      s.addShape(pres.shapes.RECTANGLE, {
        x, y, w: colW[i], h: 0.45,
        fill: { color: bg }
      });
      if (i === 2) {
        // 狀態燈號
        s.addShape(pres.shapes.OVAL, {
          x: x + colW[i] / 2 - 0.12, y: y + 0.1, w: 0.24, h: 0.24,
          fill: { color: statusColors[p.status] || C.midGray }
        });
      } else {
        s.addText(val || "", {
          x: x + 0.05, y, w: colW[i] - 0.1, h: 0.45,
          fontSize: 10, fontFace: FONT,
          color: i === 0 ? C.nearBlack : C.darkGray,
          bold: i === 0, valign: "middle", margin: 0
        });
      }
      x += colW[i];
    });
  });

  if (insightText) addInsight(s, insightText);
  return s;
}
```

---

## 引言與洞見

### 引言頁（Quotes）
原 slide129。引言框 + 作者歸屬。

```javascript
// quotes = [{ text, author }, ...]
function slideQuotes(pageTitle, quotes, insightText) {
  const s = pres.addSlide();
  s.background = { color: C.white };
  addTitle(s, pageTitle);

  const n = quotes.length;
  const totalW = 8.4, gap = 0.3;
  const cardW = (totalW - (n - 1) * gap) / n;

  quotes.forEach((q, i) => {
    const x = 0.8 + i * (cardW + gap);
    addCard(s, x, 1.55, cardW, 2.8, { borderColor: C.primary });

    // 引號裝飾
    s.addText(""", {
      x: x + 0.1, y: 1.65, w: 0.5, h: 0.5,
      fontSize: 40, fontFace: FONT, color: C.primary,
      bold: true, margin: 0
    });
    // 引言文字
    s.addText(q.text, {
      x: x + 0.2, y: 2.1, w: cardW - 0.4, h: 1.5,
      fontSize: 12, fontFace: FONT, color: C.nearBlack,
      italic: true, valign: "top", margin: 0
    });
    // 作者
    if (q.author) {
      s.addText(`— ${q.author}`, {
        x: x + 0.2, y: 3.7, w: cardW - 0.4, h: 0.4,
        fontSize: 10, fontFace: FONT, color: C.midGray,
        margin: 0
      });
    }
  });

  if (insightText) addInsight(s, insightText);
  return s;
}
```

### 引言 + 洞見（Quotes + Insights）
原 slide130。左側引言 + 右側灰色面板洞見。

```javascript
function slideQuoteInsight(pageTitle, quote, insights, metrics) {
  const s = pres.addSlide();
  s.background = { color: C.white };
  addTitle(s, pageTitle);

  // 右側灰色面板
  s.addShape(pres.shapes.RECTANGLE, {
    x: 5.8, y: 1.4, w: 3.7, h: 4.0,
    fill: { color: C.lightGray }
  });

  // 左側引言
  s.addText(""", {
    x: 0.8, y: 1.5, w: 0.5, h: 0.5,
    fontSize: 40, fontFace: FONT, color: C.re