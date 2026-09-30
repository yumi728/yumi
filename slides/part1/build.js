// 第一部スライド（seminar-part1.pptx）の生成スクリプト
// デザインは .claude/skills/voice-deck の規定に従う。座標は 1920×1080px 基準で書き、px() でインチに直す。
// 使い方：node build.js [sample]   （sample を付けると、確認用の3枚だけを作る）
const path = require("path");
const pptxgen = require("pptxgenjs");

const MODE = process.argv[2] || "full";
const ASSETS = path.join(__dirname, "assets");
const OUT = path.join(__dirname, MODE === "sample" ? "seminar-part1-sample.pptx" : "seminar-part1.pptx");

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
const F = {
  heavy: "Shippori Mincho B1 ExtraBold",
  semi: "Shippori Mincho B1 SemiBold",
  reg: "Shippori Mincho B1",
  script: "Allura",
};

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE";
pres.title = "心を掴んで“離さない”極意（第一部）";

// ---- 共通部品 ----
function text(slide, str, o) {
  slide.addText(str, {
    isTextBox: true,
    margin: 0,
    fontFace: o.font || F.reg,
    fontSize: fs(o.size || 32),
    color: o.color || C.body,
    align: o.align || "left",
    valign: o.valign || "top",
    lineSpacingMultiple: o.lh || 1.3,
    x: px(o.x), y: px(o.y), w: px(o.w), h: px(o.h),
    bold: false,
    fit: "none",
    wrap: true,
  });
}

// 文字列中の <br> を改行に、{{…}} を紫の強調にしたランの配列を作る
function runs(str, base) {
  const out = [];
  const lines = str.split("<br>");
  lines.forEach((line, li) => {
    const parts = line.split(/(\{\{.*?\}\})/).filter((s) => s !== "");
    parts.forEach((p, pi) => {
      const em = p.startsWith("{{");
      const t = em ? p.slice(2, -2) : p;
      const opt = { ...base };
      if (em) opt.color = C.purple;
      if (pi === parts.length - 1 && li < lines.length - 1) opt.breakLine = true;
      out.push({ text: t, options: opt });
    });
    if (parts.length === 0 && li < lines.length - 1) out.push({ text: "", options: { ...base, breakLine: true } });
  });
  return out;
}

function rich(slide, str, o) {
  const base = { fontFace: o.font || F.reg, fontSize: fs(o.size || 32), color: o.color || C.body };
  slide.addText(runs(str, base), {
    isTextBox: true,
    margin: 0,
    align: o.align || "left",
    valign: o.valign || "top",
    lineSpacingMultiple: o.lh || 1.3,
    x: px(o.x), y: px(o.y), w: px(o.w), h: px(o.h),
    fit: "none",
  });
}

function rrect(slide, x, y, w, h, r, fill, line) {
  slide.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: px(x), y: px(y), w: px(w), h: px(h),
    rectRadius: px(Math.min(r, Math.min(w, h) / 2)),
    fill: { color: fill },
    line: line ? { color: line.color, width: line.width, dashType: line.dash || "solid" } : { type: "none" },
  });
}

function bg(slide) {
  slide.addImage({ path: path.join(ASSETS, "bg-gradient.png"), x: 0, y: 0, w: px(1920), h: px(1080) });
}

// 枠と白パネル＋左上の見出し＋金の点線（本編の型すべて）
function framed(title) {
  const s = pres.addSlide();
  bg(s);
  rrect(s, 56, 56, 1808, 968, 32, C.white);
  text(s, title, { font: F.heavy, size: 60, color: C.head, x: 224, y: 96, w: 1500, h: 84, valign: "middle", lh: 1.0 });
  s.addShape(pres.shapes.LINE, {
    x: px(100), y: px(204), w: px(1720), h: 0,
    line: { color: C.gold, width: 2, dashType: "sysDot" },
  });
  return s;
}

function goldDots(s, dots) {
  dots.forEach(([x, y, r]) => s.addShape(pres.shapes.OVAL, { x: px(x - r), y: px(y - r), w: px(r * 2), h: px(r * 2), fill: { color: C.gold }, line: { type: "none" } }));
}

// メッセージ型：グラデーション背景の中央に白パネル → 小見出し → 飾り線 → 大きな一文
function message(small, big, dots) {
  const s = pres.addSlide();
  bg(s);
  rrect(s, 240, 160, 1440, 760, 32, C.white);
  text(s, small, { font: F.semi, size: 40, color: C.head, x: 240, y: 270, w: 1440, h: 60, align: "center", valign: "middle", lh: 1.0 });
  s.addImage({ path: path.join(ASSETS, "deco-line.png"), x: px(735), y: px(345), w: px(450), h: px(45) });
  rich(s, big, { font: F.heavy, size: 80, color: C.head, x: 320, y: 430, w: 1280, h: 380, align: "center", valign: "middle", lh: 1.4 });
  goldDots(s, dots || [[300, 230, 7], [330, 262, 4], [1620, 850, 8], [1590, 874, 4], [1648, 820, 3]]);
  return s;
}

// 全角=1、半角=0.55 として、文字列の見た目の幅（文字数）を数える
function ems(str) {
  let n = 0;
  for (const ch of str.replace(/\{\{|\}\}/g, "")) n += /[\x20-\x7e]/.test(ch) ? 0.55 : 1;
  return n;
}

// ---- 会話スライド（26・31 で共通の配置）----
const TALK = {
  top: 236, slot: 114, bubbleH: 84, size: 32, pad: 36, circle: 64,
  leftCircleX: 200, rightCircleX: 1656, gap: 28,
};

function speaker(s, who, cx, cy) {
  const isA = who === "A";
  s.addShape(pres.shapes.OVAL, {
    x: px(cx), y: px(cy - TALK.circle / 2), w: px(TALK.circle), h: px(TALK.circle),
    fill: { color: isA ? C.aquaGroup : C.pink }, line: { type: "none" },
  });
  text(s, who, { font: F.heavy, size: 30, color: C.white, x: cx, y: cy - TALK.circle / 2, w: TALK.circle, h: TALK.circle, align: "center", valign: "middle", lh: 1.0 });
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
    // しっぽ（左向き）
    s.addShape(pres.shapes.ISOSCELES_TRIANGLE, { x: px(bx - 18), y: px(cy - 12), w: px(24), h: px(24), rotate: 270, fill: { color: fill }, line: { type: "none" } });
  } else {
    speaker(s, "私", TALK.rightCircleX, cy);
    bx = TALK.rightCircleX - TALK.gap - w;
    s.addShape(pres.shapes.ISOSCELES_TRIANGLE, { x: px(bx + w - 6), y: px(cy - 12), w: px(24), h: px(24), rotate: 90, fill: { color: fill }, line: { type: "none" } });
  }
  rrect(s, bx, cy - TALK.bubbleH / 2, w, TALK.bubbleH, 28, fill);
  text(s, line.text, { font: F.reg, size: TALK.size, color: C.body, x: bx, y: cy - TALK.bubbleH / 2, w, h: TALK.bubbleH, align: "center", valign: "middle", lh: 1.0 });
  if (line.tag) {
    const tw = Math.round(ems(line.tag) * 26 + 40);
    rrect(s, bx + w + 24, cy - 22, tw, 44, 22, C.white, { color: C.lavender, width: 1.5 });
    text(s, line.tag, { font: F.semi, size: 26, color: C.head, x: bx + w + 24, y: cy - 22, w: tw, h: 44, align: "center", valign: "middle", lh: 1.0 });
  }
}

// 会話を1行ずつ増やしたスライドを並べる。perPage 行でページを改める
function conversation(title, lines, notes, perPage = 6) {
  for (let n = 1; n <= lines.length; n++) {
    const s = framed(title);
    const page = Math.floor((n - 1) / perPage);
    const start = page * perPage;
    for (let i = start; i < n; i++) bubble(s, lines[i], i - start);
    const next = lines[n] ? `次の行（${lines[n].who === "A" ? "A" : "私"}）：${lines[n].text}` : "会話はここまで";
    s.addNotes(`${notes}\n\n【段階表示 ${n}/${lines.length}】\n今表示した行（${lines[n - 1].who}）：${lines[n - 1].text}\n${next}`);
  }
}

// ---- スライド定義 ----
const NOTE = {
  ch1: "【① 今日のゴールと3つのお約束｜約2分30秒】",
  ch11: "【⑪「聞いているつもり」の落とし穴｜約5分】",
};

function slide02() {
  const s = framed("本日のゴール");
  const cards = [
    ["Point01", "頑張って{{「話す」}}を<br>手放す"],
    ["Point02", "{{「心を掴む」}}<br>コミュニケーションを知る"],
    ["Point03", "明日からの人間関係が<br>{{楽しみになる}}"],
  ];
  const cw = 500, gap = 40, x0 = (1920 - (cw * 3 + gap * 2)) / 2, y0 = 268, ch = 580;
  cards.forEach(([label, body], i) => {
    const x = x0 + i * (cw + gap);
    rrect(s, x, y0, cw, ch, 28, C.aqua);
    rrect(s, x + (cw - 220) / 2, y0 + 44, 220, 56, 28, C.white);
    text(s, label, { font: F.semi, size: 30, color: C.head, x: x + (cw - 220) / 2, y: y0 + 44, w: 220, h: 56, align: "center", valign: "middle", lh: 1.0 });
    rich(s, body, { font: F.heavy, size: 36, color: C.body, x: x + 20, y: y0 + 130, w: cw - 40, h: ch - 170, align: "center", valign: "middle", lh: 1.5 });
  });
  s.addNotes(`${NOTE.ch1}
・今日のゴール：頑張って「話す」を手放す／「心を掴む」コミュニケーションを知る／明日からの人間関係が楽しみになる
・次：今日の流れ（スライド3）`);
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

function slide26() {
  conversation("ある日の会話", TALK_26, `${NOTE.ch11}
・「簡単そう」「自分はできている」と思ったところで、次の会話を見てもらう
・実演：「私」役と「A」役は事前に決めておく
・1行ずつ表示する。6行で次のページに切り替わる`);
}

function slide28() {
  const s = framed("ズレの正体");
  const rows = [
    ["「楽しかったんですね」", "相手が言っていない気持ちを、<br>先に決めつけている"],
    ["「どういうところが<br>楽しかったんですか？」", "質問で、聞き手の知りたい方向へ<br>引っ張っている"],
    ["「そんなに人が集まったんですね。<br>すごいですね」", "文章に言い換え、評価が入っている"],
    ["「なんの先生なんですか」<br>「何の専門ですか」", "気持ちから情報へ。尋問のようになる"],
    ["「私の大学の教授も…」", "聞き手の話にすり替わる"],
  ];
  const x0 = 200, w = 1520, rh = 104, gap = 14, y0 = 232;
  rows.forEach(([said, why], i) => {
    const y = y0 + i * (rh + gap);
    rrect(s, x0, y, w, rh, 24, C.aqua);
    s.addShape(pres.shapes.OVAL, { x: px(x0 + 24), y: px(y + (rh - 56) / 2), w: px(56), h: px(56), fill: { color: C.white }, line: { type: "none" } });
    text(s, String(i + 1), { font: F.heavy, size: 30, color: C.head, x: x0 + 24, y: y + (rh - 56) / 2, w: 56, h: 56, align: "center", valign: "middle", lh: 1.0 });
    rich(s, said, { font: F.semi, size: 32, color: C.body, x: x0 + 104, y, w: 600, h: rh, valign: "middle", lh: 1.2 });
    // 左右の区切り（金の矢印）
    s.addShape(pres.shapes.RIGHT_ARROW, { x: px(x0 + 720), y: px(y + rh / 2 - 14), w: px(44), h: px(28), fill: { color: C.gold }, line: { type: "none" } });
    rich(s, why, { font: F.reg, size: 32, color: C.body, x: x0 + 790, y, w: 700, h: rh, valign: "middle", lh: 1.2 });
  });
  // 結論帯
  const by = y0 + 5 * (rh + gap) + 8;
  rrect(s, x0, by, w, 80, 20, C.band, { color: C.gold, width: 1.5, dash: "sysDot" });
  rich(s, "結果：「懐かしい人にも会えて」が、{{拾われないまま終わる}}", { font: F.semi, size: 32, color: C.body, x: x0, y: by, w, h: 80, align: "center", valign: "middle", lh: 1.0 });
  s.addNotes(`${NOTE.ch11}
解説（約2分）
・「楽しかったんですね」：相手が言っていない気持ちを、先に決めつけている
・「どういうところが楽しかったんですか？」：質問で、聞き手の知りたい方向へ引っ張っている
・「そんなに人が集まったんですね。すごいですね」：文章に言い換え、評価（ジャッジ）が入っている
・「なんの先生なんですか」「何の専門ですか」：話題が、気持ちから情報へ移り、質問が続いて尋問のようになる
・「私の大学の教授も…」：いつのまにか、聞き手の話にすり替わる
・結果：話し手が触れた「懐かしい人にも会えて」が、拾われないまま終わる`);

  const m = message("ズレの正体", "Aさんに{{悪気はない}}。<br>よくある聞き方");
  m.addNotes(`${NOTE.ch11}
・Aさんに悪気はない。よくある聞き方
・つなぎ（約30秒）：じゃあ、どうすればいいか → 次の技へ`);
}

// ---- 組み立て ----
if (MODE === "sample") {
  slide02();
  slide26();
  slide28();
}

pres.writeFile({ fileName: OUT }).then((f) => console.log("wrote", f));
