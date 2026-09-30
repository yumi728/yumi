// 第一部スライド（seminar-part1.pptx）の生成スクリプト
// デザインは .claude/skills/voice-deck の規定に従う。座標は 1920×1080px 基準で書き、px() でインチに直す。
// 使い方：node build.js          （全スライド）
//         node build.js sample   （確認用：スライド2・26・28だけ）
const fsys = require("fs");
const path = require("path");
const pptxgen = require("pptxgenjs");
const JSZip = require("jszip");

const MODE = process.argv[2] || "full";
const ASSETS = path.join(__dirname, "assets");
const OUT = path.join(__dirname, MODE === "sample" ? "seminar-part1-sample.pptx" : "seminar-part1.pptx");
const SCRIPT_SIZES = JSON.parse(fsys.readFileSync(path.join(ASSETS, "script-sizes.json"), "utf8"));

// ---- 単位 ----
const px = (v) => v / 144; // 1920px = 13.333in
const fs = (v) => v / 2; // 文字サイズ px → pt

// ---- トンマナ（voice-deck ルール7）----
const C = {
  head: "897D74",
  body: "6F635B",
  gold: "F5DE81",
  aqua: "E3F7FD",
  purple: "8B5FBF",
  blue: "4F9CCC",
  band: "FFFBEA",
  pink: "EBA3D3",
  pinkTint: "FCEEF7",
  aquaGroup: "8EC3E6",
  lavender: "B3A6EE",
  white: "FFFFFF",
};
// 書体：游明朝（Windows・Mac の両方に標準で入っている）。太字は bold で指定する
const F = {
  heavy: { face: "游明朝", bold: true },
  semi: { face: "游明朝", bold: true },
  reg: { face: "游明朝", bold: false },
};

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE";
pres.title = "心を掴んで“離さない”極意（第一部）";

// ---- 文字 ----
// 文字列中の <br> を改行に、{{…}} を紫の強調にしたランの配列を作る
function runs(str, base) {
  const out = [];
  const lines = str.split("<br>");
  lines.forEach((line, li) => {
    const parts = line.split(/(\{\{.*?\}\})/).filter((s) => s !== "");
    if (parts.length === 0) parts.push("");
    parts.forEach((p, pi) => {
      const em = p.startsWith("{{");
      const opt = { ...base };
      if (em) opt.color = C.purple;
      if (pi === parts.length - 1 && li < lines.length - 1) opt.breakLine = true;
      out.push({ text: em ? p.slice(2, -2) : p, options: opt });
    });
  });
  return out;
}

function baseRun(o) {
  const f = o.font || F.reg;
  return { fontFace: f.face, bold: f.bold, fontSize: fs(o.size || 32), color: o.color || C.body, lang: "ja-JP" };
}

function boxOpts(o) {
  return {
    isTextBox: true,
    margin: 0,
    align: o.align || "left",
    valign: o.valign || "top",
    lineSpacingMultiple: o.lh || 1.3,
    paraSpaceAfter: o.para || 0,
    x: px(o.x), y: px(o.y), w: px(o.w), h: px(o.h),
    fit: "none",
  };
}

// 1つの文字枠（<br> と {{強調}} が使える）
function rich(slide, str, o) {
  slide.addText(runs(str, baseRun(o)), boxOpts(o));
}

// 文字の大きさや太さが違う段落を1つの枠に重ねる。parts: [{t, size, font, color}]
function stack(slide, parts, o) {
  const arr = [];
  parts.forEach((p, i) => {
    const r = runs(p.t, baseRun({ ...o, ...p }));
    if (i < parts.length - 1) r[r.length - 1].options.breakLine = true;
    arr.push(...r);
  });
  slide.addText(arr, boxOpts(o));
}

// ---- 図形 ----
function rrect(slide, x, y, w, h, r, fill, line) {
  slide.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: px(x), y: px(y), w: px(w), h: px(h),
    rectRadius: px(Math.min(r, Math.min(w, h) / 2)),
    fill: fill ? { color: fill } : { type: "none" },
    line: line ? { color: line.color, width: line.width, dashType: line.dash || "solid" } : { type: "none" },
  });
}

function oval(slide, x, y, w, h, fill, line) {
  slide.addShape(pres.shapes.OVAL, {
    x: px(x), y: px(y), w: px(w), h: px(h),
    fill: fill ? { color: fill } : { type: "none" },
    line: line ? { color: line.color, width: line.width } : { type: "none" },
  });
}

function img(slide, file, x, y, w, h) {
  slide.addImage({ path: path.join(ASSETS, file), x: px(x), y: px(y), w: px(w), h: px(h) });
}

// 筆記体の英単語（画像）。高さを決めて、幅は縦横比から出す。cx を渡すと中央ぞろえ
function script(slide, key, { x, cx, y, h }) {
  const [iw, ih] = SCRIPT_SIZES[key];
  const w = (iw / ih) * h;
  img(slide, `script-${key}.png`, cx !== undefined ? cx - w / 2 : x, y, w, h);
}

function bg(slide) {
  img(slide, "bg-gradient.png", 0, 0, 1920, 1080);
}

function goldDots(s, dots) {
  dots.forEach(([x, y, r]) => oval(s, x - r, y - r, r * 2, r * 2, C.gold));
}

// ---- 型（voice-deck ルール5）----
const X0 = 200; // 本編の中身の左端
const XW = 1520; // 本編の中身の幅
const Y0 = 240; // 見出しの点線より下の、中身の上端

// 枠と白パネル＋左上の見出し＋金の点線（本編の型すべて）
function framed(title) {
  const s = pres.addSlide();
  bg(s);
  rrect(s, 56, 56, 1808, 968, 32, C.white);
  if (title) {
    rich(s, title, { font: F.heavy, size: 60, color: C.head, x: 224, y: 96, w: 1560, h: 84, valign: "middle", lh: 1.0 });
    s.addShape(pres.shapes.LINE, { x: px(100), y: px(204), w: px(1720), h: 0, line: { color: C.gold, width: 2, dashType: "sysDot" } });
  }
  return s;
}

// 表紙
function cover(title, subtitle) {
  const s = pres.addSlide();
  bg(s);
  s.addShape(pres.shapes.RECTANGLE, { x: px(200), y: px(150), w: px(1520), h: px(780), fill: { color: C.white }, line: { type: "none" } });
  s.addShape(pres.shapes.RECTANGLE, { x: px(236), y: px(186), w: px(1448), h: px(708), fill: { type: "none" }, line: { color: C.gold, width: 2 } });
  script(s, "communication", { x: 330, y: 270, h: 110 });
  rich(s, title, { font: F.heavy, size: 104, color: C.head, x: 236, y: 420, w: 1448, h: 150, align: "center", valign: "middle", lh: 1.0 });
  img(s, "deco-line.png", 735, 600, 450, 45);
  rich(s, subtitle, { font: F.semi, size: 40, color: C.head, x: 236, y: 680, w: 1448, h: 70, align: "center", valign: "middle", lh: 1.0 });
  goldDots(s, [[1640, 230, 9], [1600, 262, 5], [1668, 290, 4], [280, 860, 8], [318, 836, 4], [262, 818, 3]]);
  return s;
}

// 目次：左にグラデーションの帯と筆記体の Agenda、右に丸数字つきの一覧
function agenda(items) {
  const s = pres.addSlide();
  s.background = { color: C.white };
  img(s, "bg-band.png", 0, 0, 560, 1080);
  script(s, "agenda", { cx: 280, y: 470, h: 130 });
  const y0 = 540 - (items.length * 130) / 2 + 20;
  items.forEach((t, i) => {
    const y = y0 + i * 130;
    oval(s, 760, y, 76, 76, C.gold);
    rich(s, String(i + 1), { font: F.heavy, size: 44, color: C.head, x: 760, y, w: 76, h: 76, align: "center", valign: "middle", lh: 1.0 });
    rich(s, t, { font: F.semi, size: 40, color: C.head, x: 880, y: y - 10, w: 960, h: 96, valign: "middle", lh: 1.0 });
  });
  goldDots(s, [[1700, 880, 8], [1740, 910, 4], [1664, 930, 3], [1760, 850, 3]]);
  return s;
}

// 章扉：白パネルの中央に Chapter → 章名 → 飾り線（リードは構成案にないため置かない）
function chapter(key, name) {
  const s = framed(null);
  script(s, key, { cx: 960, y: 330, h: 120 });
  rich(s, name, { font: F.heavy, size: 60, color: C.head, x: 160, y: 480, w: 1600, h: 100, align: "center", valign: "middle", lh: 1.0 });
  img(s, "deco-line.png", 735, 610, 450, 45);
  goldDots(s, [[180, 180, 8], [214, 206, 4], [168, 222, 3], [1740, 880, 8], [1706, 906, 4], [1760, 846, 3]]);
  return s;
}

// メッセージ：グラデーション背景の中央に白パネル → 小見出し → 飾り線 → 大きな一文
function message(small, big, extra) {
  const s = pres.addSlide();
  bg(s);
  rrect(s, 240, 160, 1440, 760, 32, C.white);
  if (small) {
    rich(s, small, { font: F.semi, size: 40, color: C.head, x: 240, y: 262, w: 1440, h: 60, align: "center", valign: "middle", lh: 1.0 });
    img(s, "deco-line.png", 735, 340, 450, 45);
  }
  const bigTop = small ? 420 : 300;
  const bigH = extra ? 300 : 820 - bigTop;
  rich(s, big, { font: F.heavy, size: 80, color: C.head, x: 300, y: bigTop, w: 1320, h: bigH, align: "center", valign: "middle", lh: 1.4 });
  goldDots(s, [[300, 230, 7], [330, 262, 4], [1620, 850, 8], [1590, 874, 4], [1648, 820, 3]]);
  return s;
}

// カードの中身。label は白いピル形のラベル（上部中央）
function card(s, x, y, w, h, label, parts, o = {}) {
  rrect(s, x, y, w, h, 28, o.fill || C.aqua, o.line);
  let top = y + 24;
  if (label) {
    const lw = Math.max(200, Math.round(ems(label) * 30 + 80));
    rrect(s, x + (w - lw) / 2, y + 36, lw, 56, 28, o.pill || C.white);
    rich(s, label, { font: F.semi, size: 30, color: C.head, x: x + (w - lw) / 2, y: y + 36, w: lw, h: 56, align: "center", valign: "middle", lh: 1.0 });
    top = y + 110;
  }
  stack(s, parts, { x: x + 24, y: top, w: w - 48, h: y + h - 24 - top, align: o.align || "center", valign: "middle", lh: o.lh || 1.4, para: o.para || 0 });
}

// 結論帯（淡い金色の地に金の点線の枠）
function band(s, str, y, h = 84, x = X0, w = XW) {
  rrect(s, x, y, w, h, 20, C.band, { color: C.gold, width: 1.5, dash: "sysDot" });
  rich(s, str, { font: F.semi, size: 32, color: C.body, x: x + 24, y, w: w - 48, h, align: "center", valign: "middle", lh: 1.3 });
}

// 選択肢の行：水色の横長の行、左端に白い丸の記号
function rowBase(s, y, h, mark, x = X0, w = XW) {
  rrect(s, x, y, w, h, 24, C.aqua);
  oval(s, x + 24, y + (h - 56) / 2, 56, 56, C.white);
  if (mark === "dot") oval(s, x + 44, y + h / 2 - 8, 16, 16, C.gold);
  else rich(s, String(mark), { font: F.heavy, size: 30, color: C.head, x: x + 24, y: y + (h - 56) / 2, w: 56, h: 56, align: "center", valign: "middle", lh: 1.0 });
}

// Step のスライドの右の列の行（左にアーチがあるため）
const SX = 780, SW = 1720 - 780;
function stepRow(s, y, h, str, o = {}) {
  rowBase(s, y, h, "dot", SX, SW);
  rich(s, str, { font: o.font || F.semi, size: o.size || 40, color: C.body, x: SX + 110, y, w: SW - 140, h, valign: "middle", lh: o.lh || 1.35 });
}

function rows(s, items, { y0 = Y0, h = 90, gap = 16, mark = "dot", size = 32 } = {}) {
  items.forEach((t, i) => {
    const y = y0 + i * (h + gap);
    rowBase(s, y, h, mark === "num" ? i + 1 : mark);
    rich(s, t, { font: F.reg, size, color: C.body, x: X0 + 110, y, w: XW - 140, h, valign: "middle", lh: 1.25 });
  });
  return y0 + items.length * (h + gap);
}

// 左右に並べる行（左：発言、右：説明）
function pairRows(s, pairs, { y0 = Y0 - 8, h = 104, gap = 14, leftW = 600 } = {}) {
  pairs.forEach(([l, r], i) => {
    const y = y0 + i * (h + gap);
    rowBase(s, y, h, i + 1);
    rich(s, l, { font: F.semi, size: 32, color: C.body, x: X0 + 104, y, w: leftW, h, valign: "middle", lh: 1.2 });
    s.addShape(pres.shapes.RIGHT_ARROW, { x: px(X0 + 124 + leftW), y: px(y + h / 2 - 14), w: px(44), h: px(28), fill: { color: C.gold }, line: { type: "none" } });
    rich(s, r, { font: F.reg, size: 32, color: C.body, x: X0 + 190 + leftW, y, w: XW - 214 - leftW, h, valign: "middle", lh: 1.2 });
  });
  return y0 + pairs.length * (h + gap);
}

// 全角=1、半角=0.55 として、文字列の見た目の幅（文字数）を数える
function ems(str) {
  let n = 0;
  for (const ch of str.replace(/\{\{|\}\}/g, "")) n += /[\x20-\x7e]/.test(ch) ? 0.55 : 1;
  return n;
}

// ---- 会話スライド（26・31 で共通の配置）----
const TALK = { top: 236, slot: 114, bubbleH: 84, size: 32, pad: 36, circle: 64, leftCircleX: 200, rightCircleX: 1656, gap: 28 };

function speaker(s, who, cx, cy) {
  oval(s, cx, cy - TALK.circle / 2, TALK.circle, TALK.circle, who === "A" ? C.aquaGroup : C.pink);
  rich(s, who, { font: F.heavy, size: 30, color: C.white, x: cx, y: cy - TALK.circle / 2, w: TALK.circle, h: TALK.circle, align: "center", valign: "middle", lh: 1.0 });
}

function bubble(s, line, idx) {
  const cy = TALK.top + idx * TALK.slot + TALK.slot / 2;
  const isA = line.who === "A";
  const w = Math.round(ems(line.text) * TALK.size + TALK.pad * 2);
  const fill = isA ? C.aqua : C.pinkTint;
  let bx;
  if (isA) {
    speaker(s, "A", TALK.leftCircleX, cy);
    bx = TALK.leftCircleX + TALK.circle + TALK.gap;
    s.addShape(pres.shapes.ISOSCELES_TRIANGLE, { x: px(bx - 18), y: px(cy - 12), w: px(24), h: px(24), rotate: 270, fill: { color: fill }, line: { type: "none" } });
  } else {
    speaker(s, "私", TALK.rightCircleX, cy);
    bx = TALK.rightCircleX - TALK.gap - w;
    s.addShape(pres.shapes.ISOSCELES_TRIANGLE, { x: px(bx + w - 6), y: px(cy - 12), w: px(24), h: px(24), rotate: 90, fill: { color: fill }, line: { type: "none" } });
  }
  rrect(s, bx, cy - TALK.bubbleH / 2, w, TALK.bubbleH, 28, fill);
  rich(s, line.text, { font: F.reg, size: TALK.size, color: C.body, x: bx, y: cy - TALK.bubbleH / 2, w, h: TALK.bubbleH, align: "center", valign: "middle", lh: 1.0 });
  if (line.tag) {
    const tw = Math.round(ems(line.tag) * 26 + 40);
    rrect(s, bx + w + 24, cy - 22, tw, 44, 22, C.white, { color: C.lavender, width: 1.5 });
    rich(s, line.tag, { font: F.semi, size: 26, color: C.head, x: bx + w + 24, y: cy - 22, w: tw, h: 44, align: "center", valign: "middle", lh: 1.0 });
  }
}

// 会話を1行ずつ増やしたスライドを並べる。perPage 行でページを改める
function conversation(title, lines, notes, perPage = 6) {
  for (let n = 1; n <= lines.length; n++) {
    const s = framed(title);
    const start = Math.floor((n - 1) / perPage) * perPage;
    for (let i = start; i < n; i++) bubble(s, lines[i], i - start);
    const cur = lines[n - 1];
    const next = lines[n] ? `次の行（${lines[n].who}）：${lines[n].text}` : "会話はここまで";
    s.addNotes(`${notes}\n\n【段階表示 ${n}/${lines.length}】\n今表示した行（${cur.who}）：${cur.text}${cur.tag ? `（${cur.tag}）` : ""}\n${next}`);
  }
}

// ---- 図の部品 ----
// 壁のアイコン（白い丸の中にレンガ）
function wallIcon(s, x, y, d) {
  oval(s, x, y, d, d, C.white);
  const bw = d * 0.2, bh = d * 0.12, g = d * 0.03;
  const rowsDef = [[0, 3], [bw / 2 + g / 2, 2], [0, 3]];
  const totalW = bw * 3 + g * 2, x0 = x + (d - totalW) / 2, y0 = y + (d - (bh * 3 + g * 2)) / 2;
  rowsDef.forEach(([off, n], r) => {
    for (let i = 0; i < n; i++) rrect(s, x0 + off + i * (bw + g), y0 + r * (bh + g), bw, bh, 3, C.lavender);
  });
}

// 6つの壁（2列×3段）。cards: [{name, quotes, em}]
function wallGrid(s, cards) {
  const cw = 740, ch = 218, gx = 40, gy = 16;
  cards.forEach((c, i) => {
    const x = X0 + (i % 2) * (cw + gx), y = Y0 - 4 + Math.floor(i / 2) * (ch + gy);
    rrect(s, x, y, cw, ch, 28, C.aqua, c.em ? { color: C.purple, width: 2.25 } : null);
    wallIcon(s, x + 28, y + (ch - 112) / 2, 112);
    const parts = [{ t: c.name, size: 36, font: F.heavy, color: C.head }];
    if (c.quotes) parts.push({ t: c.quotes, size: 30, font: F.reg, color: C.body });
    stack(s, parts, { x: x + 168, y: y + 10, w: cw - 190, h: ch - 20, valign: "middle", lh: 1.12, para: 4 });
  });
}

// ストップウォッチ（数字なし）
function stopwatch(s, cx, cy, d) {
  rrect(s, cx - d * 0.09, cy - d / 2 - d * 0.16, d * 0.18, d * 0.12, 6, C.head);
  oval(s, cx - d / 2, cy - d / 2, d, d, C.white, { color: C.head, width: 5 });
  s.addShape(pres.shapes.LINE, { x: px(cx), y: px(cy - d * 0.32), w: 0, h: px(d * 0.32), line: { color: C.purple, width: 4 } });
  oval(s, cx - 7, cy - 7, 14, 14, C.purple);
}

// Step のアーチ（上が半円の水色の枠）
function arch(s, key, name, color) {
  const x = 230, w = 460, top = 238, bottom = 930;
  oval(s, x, top, w, w, C.aqua, { color, width: 3 });
  s.addShape(pres.shapes.RECTANGLE, { x: px(x), y: px(top + w / 2), w: px(w), h: px(bottom - top - w / 2), fill: { color: C.aqua }, line: { color, width: 3 } });
  // 半円と長方形のつなぎ目の線を隠す
  s.addShape(pres.shapes.RECTANGLE, { x: px(x + 3), y: px(top + w / 2 - 6), w: px(w - 6), h: px(14), fill: { color: C.aqua }, line: { type: "none" } });
  script(s, key, { cx: x + w / 2, y: 360, h: 110 });
  rich(s, name, { font: F.heavy, size: 48, color: C.head, x, y: 520, w, h: 200, align: "center", valign: "middle", lh: 1.3 });
}

// ---- 発表者ノートの章見出し ----
const CH = {
  1: "【① 今日のゴールと3つのお約束｜約2分30秒】",
  2: "【② 自己紹介｜約1分30秒】",
  3: "【③ こんなすれ違い、ありませんか？｜約3分】",
  4: "【④ 心を掴むことを阻む6つの壁｜約2分】",
  5: "【⑤ 頑張るほど、距離ができてしまう｜約1分】",
  6: "【⑥「はなさない」トークで得られること｜約1分30秒】",
  7: "【⑦ 大切なのは、話す内容より土壌づくり｜約1分30秒】",
  8: "【⑧ 心を掴む鍵は、無意識に好感度をためること｜約3分】",
  9: "【⑨「はなさないトーク」の3つの道具｜約2分】",
  10: "【⑩ 今日から使える3ステップ｜約7分】",
  11: "【⑪「聞いているつもり」の落とし穴｜約5分】",
  12: "【⑫ 会話を広げる究極の技：KMB｜約6分】",
  13: "【⑬ まとめと、このあとの実践ワーク｜約3分（ほかに予備 約1分）】",
  W: "【実践ワーク：聞き上手を体験しよう｜別枠・約10分（時間は当日の進行で調整）】",
};
const note = (s, ch, body) => s.addNotes(`${CH[ch]}\n${body}`);

// ---- スライド ----
function s01() {
  const s = cover("心を掴んで“離さない”極意", "“はなさない”トーク交流会");
  note(s, 1, "・タイトル：心を掴んで“離さない”極意（副題：“はなさない”トーク交流会）\n・次：章扉 ①");
}

function door(n, name, ch, body) {
  const s = chapter(`chapter${String(n).padStart(2, "0")}`, name);
  note(s, ch, `（章扉）\n${body}`);
}

function s02() {
  const s = framed("本日のゴール");
  const items = [
    ["Point01", "頑張って{{「話す」}}を<br>手放す"],
    ["Point02", "{{「心を掴む」}}<br>コミュニケーションを知る"],
    ["Point03", "明日からの人間関係が<br>{{楽しみになる}}"],
  ];
  const cw = 500, gap = 40, x0 = (1920 - (cw * 3 + gap * 2)) / 2;
  items.forEach(([label, body], i) => card(s, x0 + i * (cw + gap), 268, cw, 580, label, [{ t: body, size: 36, font: F.heavy }], { lh: 1.5 }));
  note(s, 1, "・今日のゴール：頑張って「話す」を手放す／「心を掴む」コミュニケーションを知る／明日からの人間関係が楽しみになる\n・次：今日の流れ");
}

function s03() {
  const s = agenda(["ミニセミナー「心を掴んで離さない極意」", "実践ワーク", "もう一つのミニセミナー「響く話し方」", "お知らせ"]);
  note(s, 1, "・今日の流れ：ミニセミナー「心を掴んで離さない極意」／実践ワーク／もう一つのミニセミナー「響く話し方」／お知らせ\n・第二部の詳細は伏せる");
}

function s04() {
  const s = framed("心地よい場作りのための「3つのお約束」");
  const items = [
    ["01.", "否定しない<br>ジャッジしない", "どんな意見も<br>「まずは受け止める」<br>安心な場に。"],
    ["02.", "ここだけの話", "安心して話せるよう、<br>守秘義務を守りましょう。"],
    ["03.", "時間を意識", "全員が楽しく話せるよう、<br>時間オーバーに注意♡"],
  ];
  const cw = 500, gap = 40, x0 = (1920 - (cw * 3 + gap * 2)) / 2;
  items.forEach(([label, head, body], i) =>
    card(s, x0 + i * (cw + gap), 268, cw, 580, label, [
      { t: head, size: 40, font: F.heavy, color: C.head },
      { t: body, size: 30, font: F.reg },
    ], { lh: 1.45, para: 18 }));
  note(s, 1, "・3つのお約束：否定しない・ジャッジしない／ここだけの話／時間を意識");
}

function s05() {
  // 講師が自分で入れるため、レイアウトのみ（角かっこは入力欄の目印）
  const s = framed("自己紹介");
  oval(s, 260, 300, 440, 440, C.aqua);
  rich(s, "［写真］", { font: F.reg, size: 32, color: C.head, x: 260, y: 300, w: 440, h: 440, align: "center", valign: "middle", lh: 1.0 });
  rich(s, "［お名前］", { font: F.heavy, size: 60, color: C.head, x: 800, y: 290, w: 900, h: 90, valign: "middle", lh: 1.0 });
  rich(s, "［肩書き］", { font: F.semi, size: 32, color: C.head, x: 800, y: 390, w: 900, h: 56, valign: "middle", lh: 1.0 });
  rrect(s, 800, 480, 900, 360, 28, C.aqua);
  rich(s, "［経歴］", { font: F.reg, size: 32, color: C.body, x: 840, y: 510, w: 820, h: 300, valign: "top", lh: 1.5 });
  note(s, 2, "・経歴と、心理学とスピリチュアルを学んで変わったこと\n・最後に口頭で一言：「今日は、初対面の緊張や、話しすぎの後悔を軽くする時間」\n※このスライドの文字は講師が入力する（［ ］は入力欄の目印）");
}

function s06() {
  const s = message(null, "“はなさない”トーク<br>とは？");
  note(s, 3, "・「はなさない」トークとは？：自分は話さないのに、相手の心を掴んで「離さない」");
}

// すれ違い（7〜9）：上に願い、下に現実、見出しは Communication Error
function mismatch(wish, real, body) {
  const s = framed("Communication Error");
  rrect(s, X0, 250, XW, 200, 28, C.aqua);
  rich(s, wish, { font: F.heavy, size: 40, color: C.head, x: X0 + 40, y: 250, w: XW - 80, h: 200, align: "center", valign: "middle", lh: 1.45 });
  s.addShape(pres.shapes.ISOSCELES_TRIANGLE, { x: px(960 - 36), y: px(478), w: px(72), h: px(52), rotate: 180, fill: { color: C.gold }, line: { type: "none" } });
  rrect(s, X0, 556, XW, 330, 28, C.white, { color: C.pink, width: 2.25 });
  rich(s, real, { font: F.reg, size: 36, color: C.body, x: X0 + 40, y: 556, w: XW - 80, h: 330, align: "center", valign: "middle", lh: 1.55 });
  note(s, 3, body);
  return s;
}

function s07() {
  mismatch("初対面でも自然に打ち解けて、<br>「また会いたい」と思ってもらいたい",
    "緊張して自分のことを話しすぎてしまったり、<br>逆に無言になってしまったりして、<br>{{なんか噛み合わないまま終わってしまう}}",
    "・すれ違い① 初対面：自然に打ち解けたいのに、緊張して話しすぎる。または無言になる");
}
function s08() {
  mismatch("「この人に話してよかった」と思ってもらえる、<br>信頼される存在でいたい",
    "気づいたら自分ばかり話していたり、<br>アドバイスしすぎてしまったりして、<br>{{「なんか重かったかな…」と後から後悔する}}",
    "・すれ違い② 信頼される存在でいたいのに、自分ばかり話す。アドバイスしすぎて、後から「重かったかな」と後悔する");
}
function s09() {
  mismatch("大切な人とわかり合いたい。<br>ちゃんと気持ちを伝えて、相手の気持ちも聞きたい",
    "言い方が悪かったのか、なぜか険悪になってしまう。<br>言いたいことが伝わらないまま、<br>{{お互いモヤモヤして終わる}}",
    "・すれ違い③ 大切な人とわかり合いたいのに、言い方が悪かったのか、なぜか険悪になる");
}

function s10() {
  const s = framed("どれか、思い当たりますか？");
  const items = ["初対面", "信頼される<br>存在でいたい", "大切な人と<br>わかり合いたい"];
  const cw = 500, gap = 40, x0 = (1920 - (cw * 3 + gap * 2)) / 2;
  items.forEach((t, i) => {
    const x = x0 + i * (cw + gap);
    rrect(s, x, 250, cw, 440, 28, C.aqua);
    oval(s, x + cw / 2 - 60, 290, 120, 120, C.gold);
    rich(s, String(i + 1), { font: F.heavy, size: 64, color: C.head, x: x + cw / 2 - 60, y: 290, w: 120, h: 120, align: "center", valign: "middle", lh: 1.0 });
    rich(s, t, { font: F.heavy, size: 40, color: C.head, x: x + 24, y: 440, w: cw - 48, h: 220, align: "center", valign: "middle", lh: 1.4 });
  });
  band(s, "相手の心を掴むコツは、頑張って{{『話す』ことではない}}", 740, 100);
  note(s, 3, "・「どれか、思い当たりますか？」と一言聞く\n・相手の心を掴むコツは、頑張って「話す」ことではない");
}

const WALLS = [
  { name: "好感・親しみの壁", quotes: "「なんとなく好きじゃない」<br>「苦手なタイプかも・・・」" },
  { name: "安心・信頼の壁", quotes: "「この人に話して大丈夫かな」<br>「何考えているかわからないな」" },
  { name: "タイミング・余裕の壁", quotes: "「今は話したくない」<br>「それどころじゃない」" },
  { name: "納得・共感の壁", quotes: "「言っていることはわかる。<br>けどなんか違う」<br>「なんだか押し付けてくるなぁ」" },
  { name: "過去の経験・思い込みの壁", quotes: "「どうせわかってもらえない」<br>「この人は〇〇な人であるはず」" },
  { name: "価値観・文化の壁", quotes: "「今の若者は・・・」<br>「沖縄の人って・・・」" },
];
const WALL_NOTE = `・好感・親しみの壁：「なんとなく好きじゃない」「苦手なタイプかも」
・安心・信頼の壁：「この人に話して大丈夫かな」「何考えているかわからないな」
・タイミング・余裕の壁：「今は話したくない」「それどころじゃない」
・納得・共感の壁：「言っていることはわかる。けどなんか違う」「なんだか押し付けてくるなぁ」
・過去の経験・思い込みの壁：「どうせわかってもらえない」「この人は〇〇な人であるはず」
・価値観・文化の壁：「今の若者は…」`;

function s11() {
  const s = framed("コミュニケーションを阻む「壁」");
  wallGrid(s, WALLS);
  note(s, 4, WALL_NOTE);
}

function s12() {
  const s = framed("改善しようと・・・");
  const steps = [
    ["01.", "説明する<br>説得する<br>話し方を変える<br>事前に準備をする"],
    ["02.", "諦める<br>憤る"],
    ["03.", "{{余計に距離ができる}}<br>{{噛み合わなくなる}}"],
  ];
  const cw = 440, arrowW = 100, x0 = (1920 - (cw * 3 + arrowW * 2)) / 2, y = 290, h = 560;
  steps.forEach(([num, body], i) => {
    const x = x0 + i * (cw + arrowW);
    rrect(s, x, y, cw, h, 28, C.aqua);
    rich(s, num, { font: F.heavy, size: 56, color: C.blue, x: x + 40, y: y + 36, w: cw - 80, h: 80, valign: "middle", lh: 1.0 });
    rich(s, body, { font: F.semi, size: 36, color: C.body, x: x + 40, y: y + 130, w: cw - 80, h: h - 170, valign: "middle", lh: 1.6 });
    if (i < 2) s.addShape(pres.shapes.RIGHT_ARROW, { x: px(x + cw + 22), y: px(y + h / 2 - 24), w: px(56), h: px(48), fill: { color: C.gold }, line: { type: "none" } });
  });
  note(s, 5, "・改善しようとして、説明する。説得する。話し方を変える。事前に準備する\n・うまくいかないと、諦める。憤る\n・結果：余計に距離ができる。噛み合わなくなる");
}

function s13() {
  const s = framed("“はなさない”トークで得られること");
  wallGrid(s, WALLS);
  note(s, 6, "・6つの壁を、もう一度見せる\n" + WALL_NOTE);
}

function s14() {
  const s = framed("“はなさない”トークで得られること");
  wallGrid(s, [
    { name: "{{好感・親しみを得られる}}", em: true },
    { name: "{{安心・信頼を得られる}}", em: true },
    { name: "タイミング・余裕の壁", quotes: "「今は話したくない」<br>「それどころじゃない」" },
    { name: "{{納得・共感される}}", em: true },
    { name: "過去の経験<br>思い込みの壁を{{越える}}", em: true },
    { name: "価値観・文化の壁を<br>{{越える}}", em: true },
  ]);
  note(s, 6, "・好感・親しみが得られる\n・安心・信頼が得られる\n・納得・共感される\n・タイミング・余裕、過去の経験・思い込み、価値観・文化の壁を越えられる");
}

function s15() {
  const s = framed("大切なのは、話す内容より【土壌づくり】");
  const end = rows(s, [
    "どんなに言葉を磨いても、土壌が整っていなければ花は咲かない",
    "「うまく話せない」のは、頑張る方向が少しズレていただけ",
    "相手の心に「この人と話したい」という土壌ができてはじめて、言葉は届く",
    "今日お伝えするのは、拍子抜けするほど簡単なテクニック",
    "シンプルで、だからこそ誰とでも、今日からすぐ使えること",
  ], { h: 88, gap: 16 });
  band(s, "なんだ、これだけで{{良かったんだ〜}}", end + 20, 96);
  note(s, 7, "・どんなに言葉を磨いても、土壌が整っていなければ花は咲かない\n・「うまく話せない」のは、頑張る方向が少しズレていただけ\n・「この人と話したい」という土壌ができてはじめて、言葉は届く\n・今日お伝えするのは、拍子抜けするほど簡単なテクニック。だからこそ、誰とでも、今日からすぐ使える");
}

function s16() {
  const s = message("理由は説明できないのに、なんとなく好きな人", "「なんとなく好き」と<br>思った経験は<br>ありませんか？");
  note(s, 8, "・「なんとなくこの人、好き」と思った経験はないか。理由は説明できないのに\n・参加者に一言聞く");
}

function s17() {
  const s = framed("コミュニケーションの鍵");
  rich(s, "無意識に{{like}}をためる", { font: F.heavy, size: 104, color: C.head, x: X0, y: 300, w: XW, h: 260, align: "center", valign: "middle", lh: 1.0 });
  img(s, "deco-line.png", 735, 580, 450, 45);
  band(s, "「なんとなく好き」「分かってくれる」<br>と自然に思われることを{{目指す！！}}", 680, 170);
  note(s, 8, "・目指すのは、「なんとなく好き」「分かってくれる」と自然に思われること");
}

function s18() {
  const s = framed("無意識にアプローチする");
  // 図：氷山（水面の上＝顕在意識、下＝潜在意識）
  const fx = 230, fw = 640, water = 420, bottom = 920;
  rrect(s, fx, water, fw, bottom - water, 28, C.aqua);
  s.addShape(pres.shapes.RECTANGLE, { x: px(fx), y: px(water), w: px(fw), h: px(40), fill: { color: C.aqua }, line: { type: "none" } });
  s.addShape(pres.shapes.LINE, { x: px(fx - 20), y: px(water), w: px(fw + 40), h: 0, line: { color: C.aquaGroup, width: 3 } });
  s.addShape(pres.shapes.ISOSCELES_TRIANGLE, { x: px(fx + fw / 2 - 110), y: px(280), w: px(220), h: px(water - 280), fill: { color: C.white }, line: { color: C.aquaGroup, width: 2.5 } });
  s.addShape(pres.shapes.ISOSCELES_TRIANGLE, { x: px(fx + 70), y: px(water), w: px(fw - 140), h: px(bottom - water - 50), rotate: 180, fill: { color: C.white }, line: { color: C.aquaGroup, width: 2.5 } });
  // 右：説明
  stack(s, [
    { t: "Conscious Mind", size: 36, font: F.semi, color: C.head },
    { t: "（顕在意識）", size: 32, font: F.reg, color: C.head },
  ], { x: 960, y: 260, w: 760, h: 140, valign: "middle", lh: 1.3 });
  stack(s, [
    { t: "Unconscious Mind", size: 48, font: F.heavy, color: C.head },
    { t: "（潜在意識）", size: 40, font: F.semi, color: C.head },
  ], { x: 960, y: 470, w: 760, h: 180, valign: "middle", lh: 1.3 });
  band(s, "無意識に対してアプローチする方が<br>{{速い＋強力！}}", 720, 180, 960, 760);
  s.addShape(pres.shapes.LINE, { x: px(900), y: px(330), w: px(46), h: 0, line: { color: C.gold, width: 3 } });
  s.addShape(pres.shapes.LINE, { x: px(900), y: px(560), w: px(46), h: 0, line: { color: C.gold, width: 3 } });
  note(s, 8, "・意識（顕在意識）と、無意識（潜在意識）がある\n・意識に働きかけるより、無意識に働きかけるほうが、速くて強力\n・土壌に、好感度をためていくイメージ（⑦とつなげる）");
}

function tools(title, withBand) {
  const s = framed(title);
  const items = [["01", "うなずき"], ["02", "相槌<br>（アイヅチ）"], ["03", "受容ワード"]];
  const cw = 500, gap = 40, x0 = (1920 - (cw * 3 + gap * 2)) / 2;
  items.forEach(([label, t], i) => card(s, x0 + i * (cw + gap), 260, cw, 440, label, [{ t, size: 48, font: F.heavy, color: C.head }], { lh: 1.35 }));
  if (withBand) band(s, withBand, 740, 150);
  return s;
}

function s19() {
  const s = tools("はなさないトークとは");
  note(s, 9, "・うなずき、相槌（アイヅチ）、受容ワード\n・話す量ではなく、聞き方で好感度をためる");
}

function s20() {
  const s = tools("はなさないトークとは", "{{聞かれるまでは説明しない！}}<br>（講座やセミナーは例外）");
  note(s, 9, "・聞かれるまでは説明しない（講座やセミナーは例外）");
}

function stepSlide(key, name, color) {
  const s = framed("はなさないトークを日常に");
  arch(s, key, name, color);
  return s;
}

function s21() {
  const s = stepSlide("step1", "うなずき", C.aquaGroup);
  stepRow(s, 300, 160, "無意識下に<br>{{承認がたまる}}");
  stepRow(s, 500, 230, "出来てない場合<br>「この人聞いてない」と<br>思われる可能性大", { font: F.reg, size: 36 });
  note(s, 10, "Step1 うなずき（約2分）\n・無意識下に承認がたまる\n・できていないと、「この人、聞いていない」と思われやすい");
}

function s22() {
  const s = stepSlide("step2", "相槌<br>（アイヅチ）", C.pink);
  stepRow(s, 300, 160, "あまり声を出さずに<br>相手の間による");
  stepRow(s, 500, 230, "{{ハ行を意識！}}");
  note(s, 10, "Step2 相槌（約2分）\n・あまり声を出さず、相手の間に合わせる\n・ハ行を意識する（例：はい／へえ／ふーん／ほう）");
}

function s23() {
  const s = stepSlide("step3", "受容ワード", C.lavender);
  stepRow(s, 300, 160, "{{そのまま受容する}}");
  stepRow(s, 500, 230, "・そうですよね<br>・分かります<br>・ありますよね<br>・そうなんだ", { font: F.reg, size: 34, lh: 1.25 });
  note(s, 10, "Step3 受容ワード（約2分）\n・そのまま受け入れる（そうですよね／分かります／ありますよね／そうなんだ）");
}

function s24() {
  const s = message("はなさないトークを日常に", "とにかく{{ジャッジしない}}<br>ことがカギ！");
  note(s, 10, "・とにかく、ジャッジしないことがカギ（約1分）");
}

function s25() {
  const s = message("「簡単そう」に見えますよね？", "「自分はできている」と<br>思った方は？");
  note(s, 11, "・「簡単そう」「自分はできている」と思ったところで、次の会話を見てもらう");
}

const TALK_26 = [
  { who: "私", text: "先週、師匠の還暦祝いに行ったんです" },
  { who: "A", text: "おぉ、楽しかったんですね。" },
  { who: "私", text: "はい笑" },
  { who: "A", text: "どういうところが楽しかったんですか？" },
  { who: "私", text: "んー4、50人かな。人が集まって、懐かしい人にも会えて" },
  { who: "A", text: "そんなに人が集まったんですね。すごいですね" },
  { who: "私", text: "すごいですよね。" },
  { who: "A", text: "なんの先生なんですか" },
  { who: "私", text: "大学の教授で、" },
  { who: "A", text: "大学、何の専門の教授ですか？" },
  { who: "私", text: "アジア社会学の先生ですね。" },
  { who: "A", text: "おぉ、私の大学の教授も…" },
];

function s26() {
  conversation("ある日の会話", TALK_26, `${CH[11]}
・「簡単そう」「自分はできている」と思ったところで、次の会話を見てもらう
・実演：「私」役と「A」役は事前に決めておく
・1行ずつ表示する。6行で次のページに切り替わる`);
}

function s27() {
  const s = message("ある日の会話", "どこが<br>ズレていたでしょう？");
  note(s, 11, "・「どこがズレていたでしょう？」と参加者に投げかける（約1分）");
}

function s28() {
  const s = framed("ズレの正体");
  const end = pairRows(s, [
    ["「楽しかったんですね」", "相手が言っていない気持ちを、<br>先に決めつけている"],
    ["「どういうところが<br>楽しかったんですか？」", "質問で、聞き手の知りたい方向へ<br>引っ張っている"],
    ["「そんなに人が集まったんですね。<br>すごいですね」", "文章に言い換え、評価が入っている"],
    ["「なんの先生なんですか」<br>「何の専門ですか」", "気持ちから情報へ。尋問のようになる"],
    ["「私の大学の教授も…」", "聞き手の話にすり替わる"],
  ]);
  band(s, "結果：「懐かしい人にも会えて」が、{{拾われないまま終わる}}", end + 8, 80);
  note(s, 11, `解説（約2分）
・「楽しかったんですね」：相手が言っていない気持ちを、先に決めつけている
・「どういうところが楽しかったんですか？」：質問で、聞き手の知りたい方向へ引っ張っている
・「そんなに人が集まったんですね。すごいですね」：文章に言い換え、評価（ジャッジ）が入っている
・「なんの先生なんですか」「何の専門ですか」：話題が、気持ちから情報へ移り、質問が続いて尋問のようになる
・「私の大学の教授も…」：いつのまにか、聞き手の話にすり替わる
・結果：話し手が触れた「懐かしい人にも会えて」が、拾われないまま終わる`);
  const m = message("ズレの正体", "Aさんに{{悪気はない}}。<br>よくある聞き方");
  note(m, 11, "・Aさんに悪気はない。よくある聞き方\n・つなぎ（約30秒）：じゃあ、どうすればいいか → 次の技へ");
}

function kmbSub(s) {
  rich(s, "{{KMB}}（極限まで短くバックトラッキング）", { font: F.heavy, size: 40, color: C.head, x: X0, y: 236, w: XW, h: 70, valign: "middle", lh: 1.0 });
}

function s29() {
  const s = framed("会話を広げる究極の技！！");
  kmbSub(s);
  rrect(s, X0, 370, 380, 150, 28, C.aqua);
  rich(s, "単語", { font: F.heavy, size: 56, color: C.head, x: X0, y: 370, w: 380, h: 150, align: "center", valign: "middle", lh: 1.0 });
  rich(s, "+", { font: F.heavy, size: 72, color: C.head, x: X0 + 380, y: 370, w: 100, h: 150, align: "center", valign: "middle", lh: 1.0 });
  rrect(s, X0 + 480, 370, 380, 150, 28, C.aqua);
  rich(s, "（助詞）", { font: F.heavy, size: 56, color: C.head, x: X0 + 480, y: 370, w: 380, h: 150, align: "center", valign: "middle", lh: 1.0 });
  rrect(s, X0 + 940, 370, 580, 150, 28, C.white, { color: C.purple, width: 2.25 });
  rich(s, "{{一文字も}}<br>{{変えずに！}}", { font: F.heavy, size: 44, color: C.body, x: X0 + 940, y: 370, w: 580, h: 150, align: "center", valign: "middle", lh: 1.3 });
  const ex = ["・鹿児島へ！", "・映画に！", "・喧嘩を！", "・鍋で！"];
  const pw = 350, pg = (XW - pw * 4) / 3;
  ex.forEach((t, i) => {
    rrect(s, X0 + i * (pw + pg), 600, pw, 110, 55, C.white, { color: C.gold, width: 2.25 });
    rich(s, t, { font: F.semi, size: 40, color: C.body, x: X0 + i * (pw + pg), y: 600, w: pw, h: 110, align: "center", valign: "middle", lh: 1.0 });
  });
  band(s, "相手からすると{{「自分の言葉」}}", 780, 110);
  note(s, 12, "説明（約3分）\n・KMB＝極限まで短くバックトラッキング\n・単語＋（助詞）だけを、一文字も変えずに返す（鹿児島へ！／映画に！／喧嘩を！／鍋で！）\n・相手からすると、「自分の言葉」が返ってくる");
}

function s30() {
  const s = framed("究極の技をお伝えします！！");
  kmbSub(s);
  const end = pairRows(s, [
    ["「昨日眠れなかったの」", "「昨日は眠れなかったんですね」"],
    ["「このプロジェクト、厳しいな」", "「厳しい状況なんですね」"],
    ["「この商品は高すぎ」", "「高すぎると感じていらっしゃるんですね」"],
  ], { y0: 340, h: 110, gap: 20, leftW: 560 });
  band(s, "{{オウム返しはNG}}", end + 20, 110);
  note(s, 12, "・文章にして言い換えるオウム返しはNG（例：「昨日眠れなかったの」→「昨日は眠れなかったんですね」）\n・短く返す");
}

const TALK_31 = [
  { who: "私", text: "先週、師匠の還暦祝いに行ったんです" },
  { who: "A", text: "おぉ、還暦祝い", tag: "感嘆＋KMB" },
  { who: "私", text: "そうなんです。4、50人集まってくれて" },
  { who: "A", text: "えー！！4、50人！？", tag: "感嘆＋KMB" },
  { who: "私", text: "はい。懐かしい人にも会えて" },
  { who: "A", text: "うんうん。そうだったんですねぇ", tag: "うなずき＋受容ワード" },
  { who: "私", text: "大学の頃の先輩に、10年ぶりに会えたんです" },
  { who: "A", text: "10年ぶりかぁ", tag: "KMB" },
  { who: "私", text: "話し始めたら、一気に当時に戻った感じで嬉しくて" },
  { who: "A", text: "わぁそれは嬉しいですよね！", tag: "共感ワード" },
];

function s31() {
  conversation("同じ場面を、技を組み合わせて", TALK_31, `${CH[12]}
同じ場面を、技を組み合わせて実演（約2分）
・気持ちは、相手が口にしてから受け取る（ズレる版は、相手が言う前に「楽しかったんですね」と決めつけていた）
・質問をしなくても、相手が自分から続きを話す
・最後まで、相手の話のまま
・実演では、うなずきを大きく、声は控えめに
・1行ずつ表示する。6行で次のページに切り替わる`);
}

function s32() {
  const s = framed("KMBの使いどころ");
  rows(s, [
    "気持ちは、相手が口にしてから受け取る",
    "ここぞという言葉に使う",
    "ほかの技と組み合わせる",
    "{{質問しなくても、相手が続きを話す}}",
  ], { y0: 270, h: 130, gap: 26, size: 40 });
  note(s, 12, "補足（約1分）\n・KMBは、ここぞという言葉に使う。ほかの技と組み合わせる\n・気持ちは、相手が口にしてから受け取る\n・質問をしなくても、相手が自分から続きを話す");
}

function s33() {
  const s = framed("まとめ");
  const y1 = 250, h = 110, g = 22;
  rowBase(s, y1, h, "dot");
  rich(s, "「うまく話す」より{{【土壌を作る】}}", { font: F.semi, size: 40, color: C.body, x: X0 + 110, y: y1, w: XW - 140, h, valign: "middle", lh: 1.0 });
  const y2 = y1 + h + g;
  rowBase(s, y2, h, "dot");
  rich(s, "4つの道具：", { font: F.semi, size: 40, color: C.body, x: X0 + 110, y: y2, w: 280, h, valign: "middle", lh: 1.0 });
  ["うなずき", "相槌", "受容ワード", "KMB"].forEach((t, i) => {
    const pw = 250, x = X0 + 390 + i * (pw + 24);
    rrect(s, x, y2 + 22, pw, h - 44, 33, C.white);
    rich(s, t, { font: F.semi, size: 34, color: C.head, x, y: y2 + 22, w: pw, h: h - 44, align: "center", valign: "middle", lh: 1.0 });
  });
  const y3 = y2 + h + g;
  rowBase(s, y3, h, "dot");
  rich(s, "説明せず、ジャッジせず、聞く。それだけで好感度はたまる", { font: F.semi, size: 40, color: C.body, x: X0 + 110, y: y3, w: XW - 140, h, valign: "middle", lh: 1.0 });
  band(s, "簡単なスキルで、{{人間関係は変わる}}", y3 + h + 60, 120);
  note(s, 13, "・今日の道具：うなずき、相槌、受容ワード、KMB\n・説明せず、ジャッジせず、聞く。それだけで好感度はたまる\n・簡単なスキルで、人間関係は変わる");
}

function s34() {
  const s = framed("実践してみましょう！");
  rrect(s, 1440, 110, 280, 64, 32, C.white, { color: C.gold, width: 2.25 });
  rich(s, "WORK TIME", { font: F.semi, size: 32, color: C.head, x: 1440, y: 110, w: 280, h: 64, align: "center", valign: "middle", lh: 1.0 });
  rowBase(s, 270, 120, "dot");
  rich(s, "うなずき・アイヅチ・受容ワード＆KMBを使って", { font: F.semi, size: 40, color: C.body, x: X0 + 110, y: 270, w: XW - 140, h: 120, valign: "middle", lh: 1.0 });
  card(s, X0, 440, XW, 420, "テーマ", [{ t: "{{「自己紹介〜私はこういう人です〜」}}", size: 60, font: F.heavy }]);
  note(s, 13, "・実践のやり方：うなずき・相槌・受容ワード・KMBを使う\n・テーマは「自己紹介〜私はこういう人です〜」\n・このあと、実際に体験する（時間は当日の進行で調整）");
}

function s35() {
  const s = framed("顔と体に　カメラと目線を合わせる");
  // 図：画面（全体）と、話し相手の顔
  const sx = 230, sy = 300, sw = 620, sh = 420;
  rrect(s, sx, sy, sw, sh, 24, C.aqua, { color: C.aquaGroup, width: 3 });
  rrect(s, sx + sw / 2 - 90, sy + sh + 10, 180, 30, 8, C.aquaGroup);
  oval(s, sx + sw / 2 - 70, sy + 80, 140, 140, C.white, { color: C.head, width: 2.5 });
  rrect(s, sx + sw / 2 - 130, sy + 240, 260, 150, 70, C.white, { color: C.head, width: 2.5 });
  oval(s, sx + sw / 2 - 150, sy + 50, 300, 190, null, { color: C.purple, width: 3 });
  const rx = 940, rw = 780;
  rrect(s, rx, 300, rw, 190, 24, C.aqua);
  rich(s, "意識では{{話し相手の顔だけ}}に<br>集中しているつもり", { font: F.semi, size: 40, color: C.body, x: rx + 40, y: 300, w: rw - 80, h: 190, valign: "middle", lh: 1.4 });
  rrect(s, rx, 530, rw, 190, 24, C.aqua);
  rich(s, "でも無意識では{{画面全体の情報}}を<br>察知している", { font: F.semi, size: 40, color: C.body, x: rx + 40, y: 530, w: rw - 80, h: 190, valign: "middle", lh: 1.4 });
  note(s, 13, "・カメラと目線を合わせる（意識では相手の顔だけを見ているつもりでも、無意識は画面全体を察知している）\n・このあと、実際に体験する\n（予備 約1分）");
}

function roles(title, left, right, noteBody) {
  const s = framed(title);
  const cw = 620, y = 290, h = 520;
  [[X0, left], [X0 + XW - cw, right]].forEach(([x, t]) => {
    rrect(s, x, y, cw, h, 28, C.aqua);
    rich(s, t, { font: F.heavy, size: 80, color: C.head, x, y, w: cw, h, align: "center", valign: "middle", lh: 1.0 });
  });
  stopwatch(s, 960, y + h / 2 + 20, 180);
  note(s, "W", noteBody);
}

function s36() {
  roles("1回目", "話し手", "聞き手", "1回目：話し手と聞き手に分かれる（目安 約4分）\n・タイマーは別に動かす");
}
function s37() {
  roles("交代", "{{聞き手}}", "{{話し手}}", "交代：役割を入れ替える（目安 約4分）\n・タイマーは別に動かす");
}

function s38() {
  const s = message(null, "今の体験を、<br>覚えておいてください", true);
  rich(s, "感想の共有は、第二部の冒頭で行います", { font: F.reg, size: 36, color: C.body, x: 300, y: 620, w: 1320, h: 60, align: "center", valign: "middle", lh: 1.0 });
  script(s, "end", { cx: 960, y: 720, h: 110 });
  note(s, "W", "締め（約2分）\n・「今の体験を、覚えておいてください」\n・感想の共有は、第二部の冒頭で行う\n・次のセミナーへの一言（詳細は伏せる）");
}

// ---- 組み立て ----
function buildFull() {
  s01();
  door(1, "今日のゴールと3つのお約束", 1, "・今日のゴール：頑張って「話す」を手放す／「心を掴む」コミュニケーションを知る／明日からの人間関係が楽しみになる\n・3つのお約束：否定しない・ジャッジしない／ここだけの話／時間を意識");
  s02(); s03(); s04();
  door(2, "自己紹介", 2, "・経歴と、心理学とスピリチュアルを学んで変わったこと");
  s05();
  door(3, "こんなすれ違い、ありませんか？", 3, "・3つのすれ違いを見せ、「どれか、思い当たりますか？」と聞く");
  s06(); s07(); s08(); s09(); s10();
  door(4, "心を掴むことを阻む6つの壁", 4, "・6つの壁を紹介する");
  s11();
  door(5, "頑張るほど、距離ができてしまう", 5, "・改善しようとするほど、距離ができてしまう");
  s12();
  door(6, "「はなさない」トークで得られること", 6, "・好感・親しみ、安心・信頼、納得・共感が得られ、3つの壁を越えられる");
  s13(); s14();
  door(7, "大切なのは、話す内容より土壌づくり", 7, "・話す内容より、土壌づくりが大切");
  s15();
  door(8, "心を掴む鍵は、無意識に好感度をためること", 8, "・無意識に好感度をためる");
  s16(); s17(); s18();
  door(9, "「はなさないトーク」の3つの道具", 9, "・うなずき、相槌（アイヅチ）、受容ワード");
  s19(); s20();
  door(10, "今日から使える3ステップ", 10, "・Step1 うなずき／Step2 相槌／Step3 受容ワード／とにかく、ジャッジしないことがカギ");
  s21(); s22(); s23(); s24();
  door(11, "「聞いているつもり」の落とし穴", 11, "・「簡単そう」「自分はできている」と思ったところで、次の会話を見てもらう");
  s25(); s26(); s27(); s28();
  door(12, "会話を広げる究極の技：KMB", 12, "・KMB＝極限まで短くバックトラッキング");
  s29(); s30(); s31(); s32();
  door(13, "まとめと、このあとの実践ワーク", 13, "・今日の道具：うなずき、相槌、受容ワード、KMB");
  s33(); s34(); s35();
  const w = chapter("work", "実践ワーク：聞き上手を体験しよう");
  note(w, "W", "（章扉）\n・1回目：話し手と聞き手に分かれる／交代／締め");
  s36(); s37(); s38();
}

// pptxgenjs は東アジアの書体に簡体字中国語の文字コード（-122）を書くため、日本語（-128）に直して保存する
async function save() {
  const buf = await pres.write({ outputType: "nodebuffer" });
  const zip = await JSZip.loadAsync(buf);
  for (const name of Object.keys(zip.files)) {
    if (!/^ppt\/(slides|notesSlides|slideLayouts|slideMasters)\/.*\.xml$/.test(name)) continue;
    const xml = await zip.file(name).async("string");
    zip.file(name, xml.replace(/(<a:ea [^>]*charset=")-122"/g, '$1-128"'));
  }
  fsys.writeFileSync(OUT, await zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" }));
  console.log("wrote", OUT, "slides:", pres.slides.length);
}

if (MODE === "sample") {
  s02(); s26(); s28();
} else {
  buildFull();
}
save();
