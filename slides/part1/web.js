// build.js の描画を、claude.ai の Slides 形式（1スライド＝1つの HTML の <section>）に書き出すための代わりの pres。
// build.js が使う pptxgenjs の機能（図形・文字・画像・ノート）だけを、同じ座標（1920×1080px）で HTML に置き換える。
const fsys = require("fs");
const path = require("path");

const IN = 144; // 1in = 144px（1920px = 13.333in）
const P = (v) => Math.round(v * IN * 10) / 10;
const esc = (t) => String(t).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
const hex = (c) => "#" + c;

// 筆記体の英単語（画像ファイル名 → 文字）
const SCRIPT_WORDS = {
  communication: "Communication", agenda: "Agenda", work: "Work", end: "END",
  step1: "Step 1", step2: "Step 2", step3: "Step 3",
};
for (let n = 1; n <= 13; n++) SCRIPT_WORDS[`chapter${String(n).padStart(2, "0")}`] = `Chapter ${String(n).padStart(2, "0")}`;

const FONT_BODY = "'Shippori Mincho B1', Georgia, serif";
const FONT_SCRIPT = "'Allura', 'Brush Script MT', cursive";

// 手書き風の飾り線（deco-line.png と同じ形を SVG で描く）
function decoSvg(x, y, w, h) {
  const sx = w / 450, sy = h / 45;
  const pts = [];
  for (let i = 0; i <= 60; i++) {
    const t = (i / 60) * 2 * Math.PI;
    pts.push([60 + 19 * Math.cos(t + Math.PI) + i * 0.09, 24 + 13 * Math.sin(t + Math.PI)]);
  }
  const [x0, y0] = pts[pts.length - 1];
  for (let i = 1; i <= 80; i++) {
    const u = i / 80;
    pts.push([x0 + u * (410 - x0), y0 - 11 * Math.sin(u * Math.PI * 1.15) * (1 - u * 0.4)]);
  }
  const d = pts.map(([px, py], i) => `${i ? "L" : "M"}${(px * sx).toFixed(1)} ${(py * sy).toFixed(1)}`).join(" ");
  return `<svg aria-label="" viewBox="0 0 ${w} ${h}" width="${w}" height="${h}" style="position:absolute;left:${x}px;top:${y}px;width:${w}px;height:${h}px"><path d="${d}" fill="none" stroke="#897D74" stroke-width="2.4" stroke-linecap="round"/></svg>`;
}

class WebSlide {
  constructor(id) {
    this.id = id;
    this.bg = null;
    this.parts = [];
    this.notes = "";
  }
  set background(v) { this.bg = hex(v.color); }
  addNotes(t) { this.notes = t; }

  addImage(o) {
    const f = path.basename(o.path, ".png");
    const x = P(o.x), y = P(o.y), w = P(o.w), h = P(o.h);
    if (f === "bg-gradient") {
      this.bg = "linear-gradient(to top right, #A6BEF7 0%, #D6CDF8 38%, #E8CCF4 68%, #FBBCEC 100%)";
      this.parts.push(`<div style="position:absolute;left:0px;top:0px;width:1920px;height:1080px;background:radial-gradient(circle at 0% 0%, #E6FAFA 0%, rgba(230,250,250,0) 55%)"></div>`);
    } else if (f === "bg-band") {
      this.parts.push(`<div style="position:absolute;left:${x}px;top:${y}px;width:${w}px;height:${h}px;background:linear-gradient(to top right, #A6BEF7 0%, #D6CDF8 45%, #E6FAFA 100%)"></div>`);
    } else if (f === "deco-line") {
      this.parts.push(decoSvg(x, y, w, h));
    } else if (f.startsWith("script-")) {
      const word = SCRIPT_WORDS[f.slice(7)];
      const size = Math.round(h * 1.15);
      this.parts.push(`<p style="position:absolute;left:${x - 200}px;top:${y - h * 0.12}px;width:${w + 400}px;height:${h * 1.3}px;font-family:${FONT_SCRIPT};font-size:${size}px;line-height:1.1;color:#897D74;text-align:center;white-space:nowrap">${esc(word)}</p>`);
    }
  }

  addShape(type, o) {
    const x = P(o.x), y = P(o.y), w = P(o.w), h = P(o.h);
    const fill = o.fill && o.fill.color ? hex(o.fill.color) : null;
    const ln = o.line && o.line.color ? o.line : null;
    const bw = ln ? Math.max(1, Math.round(ln.width * 2)) : 0;
    const bstyle = ln ? (ln.dashType === "sysDot" ? "dotted" : "solid") : null;
    const border = ln ? `border:${bw}px ${bstyle} ${hex(ln.color)};` : "";
    const pos = `position:absolute;left:${x}px;top:${y}px;width:${w}px;height:${h}px;`;
    switch (type) {
      case "roundRect": {
        const r = o.rectRadius ? P(o.rectRadius) : 0;
        this.parts.push(`<div style="${pos}${fill ? `background:${fill};` : ""}${border}border-radius:${r}px"></div>`);
        break;
      }
      case "rect":
        this.parts.push(`<div style="${pos}${fill ? `background:${fill};` : ""}${border}"></div>`);
        break;
      case "ellipse":
        this.parts.push(`<x-shape kind="ellipse" style="${pos}background:${fill || "transparent"};${border}"></x-shape>`);
        break;
      case "rightArrow":
        this.parts.push(`<x-shape kind="arrow-right" style="${pos}background:${fill}"></x-shape>`);
        break;
      case "line": {
        if (h === 0) {
          this.parts.push(`<hr style="position:absolute;left:${x}px;top:${y - bw / 2}px;width:${w}px;height:0px;border-top:${bw}px ${bstyle} ${hex(ln.color)}">`);
        } else {
          this.parts.push(`<svg aria-label="" viewBox="0 0 ${bw * 2} ${h}" width="${bw * 2}" height="${h}" style="position:absolute;left:${x - bw}px;top:${y}px;width:${bw * 2}px;height:${h}px"><line x1="${bw}" y1="0" x2="${bw}" y2="${h}" stroke="${hex(ln.color)}" stroke-width="${bw}" stroke-linecap="round"/></svg>`);
        }
        break;
      }
      case "triangle": {
        // rotate 0=上向き、90=右、180=下、270=左
        const r = ((o.rotate || 0) % 360 + 360) % 360;
        const pad = bw;
        const W = w + pad * 2, H = h + pad * 2;
        const L = pad, T = pad, R = pad + w, B = pad + h, CX = pad + w / 2, CY = pad + h / 2;
        const pts = { 0: [[CX, T], [R, B], [L, B]], 90: [[L, T], [R, CY], [L, B]], 180: [[L, T], [R, T], [CX, B]], 270: [[R, T], [R, B], [L, CY]] }[r];
        const stroke = ln ? ` stroke="${hex(ln.color)}" stroke-width="${bw}" stroke-linejoin="round"` : "";
        this.parts.push(`<svg aria-label="" viewBox="0 0 ${W} ${H}" width="${W}" height="${H}" style="position:absolute;left:${x - pad}px;top:${y - pad}px;width:${W}px;height:${H}px"><polygon points="${pts.map((p) => p.join(",")).join(" ")}" fill="${fill || "none"}"${stroke}/></svg>`);
        break;
      }
      default:
        throw new Error("unsupported shape " + type);
    }
  }

  addText(runsOrText, o) {
    const runsArr = typeof runsOrText === "string" ? [{ text: runsOrText, options: o }] : runsOrText;
    // breakLine ごとに行（<p>）に分ける
    const lines = [[]];
    runsArr.forEach((r) => {
      lines[lines.length - 1].push(r);
      if (r.options && r.options.breakLine) lines.push([]);
    });
    if (lines[lines.length - 1].length === 0 && lines.length > 1) lines.pop();
    const x = P(o.x), y = P(o.y), w = P(o.w), h = P(o.h);
    const jc = { top: "start", middle: "center", bottom: "end" }[o.valign || "top"];
    const lh = ((o.lineSpacingMultiple || 1) * 1.2).toFixed(2);
    const gap = o.paraSpaceAfter ? Math.round(o.paraSpaceAfter * 2) : 0;
    const ps = lines.map((line) => {
      const f = (line[0] && line[0].options) || {};
      const size = Math.round((f.fontSize || 16) * 2);
      const weight = f.weight || (f.bold ? 700 : 400);
      const color = f.color ? hex(f.color) : "#6F635B";
      const inner = line.map((r) => {
        const c = r.options && r.options.color && r.options.color !== f.color ? r.options.color : null;
        const t = esc(r.text);
        return c ? `<span style="color:${hex(c)}">${t}</span>` : t;
      }).join("") || "&#160;";
      return `<p style="font-size:${size}px;font-weight:${weight};line-height:${lh};color:${color};text-align:${o.align || "left"}">${inner}</p>`;
    });
    this.parts.push(`<div style="position:absolute;left:${x}px;top:${y}px;width:${w}px;height:${h}px;display:flex;flex-direction:column;justify-content:${jc};gap:${gap}px">${ps.join("")}</div>`);
  }

  html() {
    const notes = this.notes ? `\n<aside>${esc(this.notes).slice(0, 3900)}</aside>` : "";
    return `<section id="${this.id}" style="background:${this.bg || "#FFFFFF"};font-family:${FONT_BODY};color:#6F635B">\n${this.parts.join("\n")}${notes}\n</section>\n`;
  }
}

class WebPres {
  constructor() {
    this.slides = [];
    this.shapes = {
      ROUNDED_RECTANGLE: "roundRect", RECTANGLE: "rect", OVAL: "ellipse", LINE: "line",
      ISOSCELES_TRIANGLE: "triangle", RIGHT_ARROW: "rightArrow",
    };
    this.sections = [];
  }
  addSlide() {
    const s = new WebSlide(`s${String(this.slides.length + 1).padStart(2, "0")}`);
    this.slides.push(s);
    return s;
  }
  // 章の始まり（目次のアウトライン用）
  markSection(description) {
    this.sections.push({ description, index: this.slides.length });
  }
  writeProject(root, title) {
    const dir = path.join(root, "project", "slides");
    fsys.mkdirSync(dir, { recursive: true });
    for (const f of fsys.readdirSync(dir)) fsys.unlinkSync(path.join(dir, f));
    const sections = {};
    this.sections.forEach((s, i) => { sections[`c${String(i + 1).padStart(2, "0")}`] = { description: s.description, start: this.slides[s.index].id }; });
    const deck = {
      v: 4,
      createdOnFiles: { v: 1, at: new Date().toISOString().replace(/\.\d+Z$/, "Z") },
      lists: "css",
      title,
      cover: this.slides[0].id,
      order: this.slides.map((s) => s.id),
      sections,
      faces: {
        "shippori-mincho-b1": { family: "Shippori Mincho B1", href: "https://fonts.googleapis.com/css2?family=Shippori+Mincho+B1:wght@400;600;800&display=swap" },
        allura: { family: "Allura", href: "https://fonts.googleapis.com/css2?family=Allura&display=swap" },
      },
      designSystems: [],
    };
    fsys.writeFileSync(path.join(root, "project", "deck.json"), JSON.stringify(deck, null, 1));
    this.slides.forEach((s) => fsys.writeFileSync(path.join(dir, `${s.id}.html`), s.html()));
    return deck;
  }
}

module.exports = { WebPres };
