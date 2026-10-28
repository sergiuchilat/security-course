const pptxgen = require("pptxgenjs");
const React = require("react");
const { renderToStaticMarkup } = require("react-dom/server");
const sharp = require("sharp");
const fa = require("react-icons/fa");
const { applyTheme } = require("/Users/sergiuchilat/.claude/skills/synced/e25047ed-7c97-4849-9da1-c59bbd817501_b100ab20-30b2-4fd4-bf29-fbdff3299f26/pptx/scripts/apply_theme.js");

const OUT = "../Case_Study_Moldova_eGov_Breach_2026.pptx";

const THEME = {
  name: "Security Course",
  headFontFace: "Cambria",
  bodyFontFace: "Calibri",
  colors: {
    dk1: "1B2430", lt1: "FFFFFF", dk2: "0F1B2D", lt2: "EEF2F6",
    accent1: "F2A541", accent2: "1B998B", accent3: "D64550",
    accent4: "3D5A80", accent5: "6B7A8C", accent6: "8E6C8A",
    hlink: "1B998B", folHlink: "8E6C8A",
  },
};
const H = THEME.colors;

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE"; // 13.333 x 7.5
pres.title = "Case study – Moldova e-Government breach, October 2026";
pres.author = "Web Security Course";
pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
const C = pres.SchemeColor;

// ---------- icons ----------
async function icon(Comp, hex, size = 256) {
  const svg = renderToStaticMarkup(React.createElement(Comp, { color: "#" + hex, size: String(size) }));
  const buf = await sharp(Buffer.from(svg)).resize(size, size).png().toBuffer();
  return "image/png;base64," + buf.toString("base64");
}

// icon inside a filled circle
async function iconCircle(slide, Comp, x, y, d, fill, iconHex, name) {
  slide.addShape(pres.shapes.OVAL, { x, y, w: d, h: d, fill: { color: fill }, line: { type: "none" }, objectName: name + " circle" });
  const pad = d * 0.25;
  slide.addImage({ data: await icon(Comp, iconHex), x: x + pad, y: y + pad, w: d - 2 * pad, h: d - 2 * pad, objectName: name + " icon", altText: name });
}

// ---------- layouts ----------
const FOOTER = "Case study: Moldova e-Government breach, October 2026";

pres.defineSlideMaster({
  title: "TITLE_DARK",
  background: { color: H.dk2 },
  objects: [
    { placeholder: { options: { name: "title", type: "title", x: 0.8, y: 2.3, w: 11.7, h: 1.8, fontFace: THEME.headFontFace, fontSize: 48, bold: true, color: C.background1, align: "left", valign: "bottom", margin: 0 }, text: "Title" } },
    { placeholder: { options: { name: "body", type: "body", x: 0.8, y: 4.3, w: 11.7, h: 1.2, fontFace: THEME.bodyFontFace, fontSize: 22, color: C.accent1, valign: "top", margin: 0 }, text: "Subtitle" } },
  ],
});

pres.defineSlideMaster({
  title: "SECTION_DARK",
  background: { color: H.dk2 },
  objects: [
    { placeholder: { options: { name: "title", type: "title", x: 3.8, y: 2.9, w: 6.5, h: 1.0, fontFace: THEME.headFontFace, fontSize: 44, bold: true, color: C.background1, align: "left", valign: "middle", margin: 0 }, text: "Section" } },
    { placeholder: { options: { name: "body", type: "body", x: 3.8, y: 3.95, w: 6.5, h: 0.8, fontFace: THEME.bodyFontFace, fontSize: 20, color: C.accent1, align: "left", valign: "top", margin: 0 }, text: "Subtitle" } },
  ],
});

pres.defineSlideMaster({
  title: "CONTENT",
  background: { color: H.lt1 },
  margin: [0.5, 0.6, 0.6, 0.6],
  objects: [
    { text: { text: FOOTER, options: { x: 0.6, y: 7.0, w: 8, h: 0.3, fontFace: THEME.bodyFontFace, fontSize: 10, color: C.accent5, margin: 0 } } },
    { placeholder: { options: { name: "title", type: "title", x: 0.6, y: 0.35, w: 12.1, h: 0.9, fontFace: THEME.headFontFace, fontSize: 36, bold: true, color: C.text2, align: "left", valign: "middle", margin: 0 }, text: "Title" } },
  ],
  slideNumber: { x: 12.2, y: 7.0, w: 0.5, h: 0.3, fontFace: THEME.bodyFontFace, fontSize: 10, color: H.accent5, align: "right" },
});


// Explicit formatting on every placeholder, so viewers that ignore layout styles
// (Keynote, Quick Look, Google Slides) still render titles correctly.
const TITLE_FMT = { fontFace: THEME.headFontFace, bold: true, align: "left", margin: 0 };
const BODY_FMT = { fontFace: THEME.bodyFontFace, bold: false, align: "left", margin: 0 };

function titleSlide(section, title, sub, iconData, iconName) {
  const s = pres.addSlide({ masterName: "TITLE_DARK", sectionTitle: section });
  s.addText(title, Object.assign({ placeholder: "title", x: 0.8, y: 2.3, w: 11.7, h: 1.8, fontSize: 48, color: C.background1, valign: "bottom" }, TITLE_FMT));
  s.addText(sub, Object.assign({ placeholder: "body", x: 0.8, y: 4.3, w: 11.7, h: 1.2, fontSize: 22, color: C.accent1, valign: "top" }, BODY_FMT));
  s.addImage({ data: iconData, x: 0.8, y: 1.0, w: 0.9, h: 0.9, objectName: iconName + " icon", altText: iconName });
  return s;
}

async function dividerSlide(section, num, title, sub, ic, iconName) {
  const s = pres.addSlide({ masterName: "SECTION_DARK", sectionTitle: section });
  tb(s, num, { x: 0.8, y: 2.35, w: 2.7, h: 2.2, fontSize: 130, bold: true, color: C.accent1, fontFace: THEME.headFontFace, valign: "middle" });
  tb(s, `PART ${num} OF 05`, { x: 3.8, y: 2.45, w: 6.5, h: 0.4, fontSize: 14, bold: true, charSpacing: 3, color: C.accent5 });
  s.addText(title, Object.assign({ placeholder: "title", x: 3.8, y: 2.9, w: 6.5, h: 1.0, fontSize: 44, color: C.background1, valign: "middle" }, TITLE_FMT));
  s.addText(sub, Object.assign({ placeholder: "body", x: 3.8, y: 3.95, w: 6.5, h: 0.8, fontSize: 20, color: C.accent1, valign: "top" }, BODY_FMT));
  await iconCircle(s, ic, 10.7, 2.65, 1.8, C.accent4, H.accent1, iconName);
  return s;
}

function contentSlide(title, section) {
  const s = pres.addSlide({ masterName: "CONTENT", sectionTitle: section });
  s.addText(title, Object.assign({ placeholder: "title", x: 0.6, y: 0.35, w: 12.1, h: 0.9, fontSize: 36, color: C.text2, valign: "middle" }, TITLE_FMT));
  return s;
}

function tb(slide, text, opts) {
  slide.addText(text, Object.assign({ isTextBox: true, fontFace: THEME.bodyFontFace, margin: 0, valign: "top" }, opts));
}

function card(slide, x, y, w, h, name, fill = C.background2) {
  slide.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, rectRadius: 0.12, fill: { color: fill }, line: { type: "none" }, objectName: name });
}

(async () => {
  // ================= INTRO =================
  const S0 = "Introduction";
  pres.addSection({ title: S0 });

  // 1. Title
  {
    const s = titleSlide(S0, "Case study: the Moldova e-Government breach", "October 2026  ·  What happened, what the risks are, and how to prevent it", await icon(fa.FaUserShield, H.accent1), "Shield");
    tb(s, "Based on official statements and press reports up to 28 October 2026. Findings are preliminary.", { x: 0.8, y: 5.6, w: 11.7, h: 0.4, fontSize: 14, italic: true, color: C.accent5 });
    s.addNotes("This is a real incident that happened in our country a few days ago. Everything in this deck comes from public statements by the e-Governance Agency (AGE), STISC and press reports up to 28 October 2026. The investigation is ongoing, so numbers may change. We focus on lessons, not on blame.");
  }

  // 2. Agenda
  {
    const s = contentSlide("Today's agenda", S0);
    const items = [
      [fa.FaCalendarAlt, "What happened", "Numbers, timeline, data exposed"],
      [fa.FaSearch, "How it happened", "What is publicly known about the attack"],
      [fa.FaExclamationTriangle, "The risks", "For citizens and for the state"],
      [fa.FaShieldAlt, "How to prevent it", "Citizens, developers, organizations"],
      [fa.FaChalkboardTeacher, "Discussion", "Class exercise and key takeaways"],
    ];
    for (let i = 0; i < items.length; i++) {
      const y = 1.55 + i * 1.03;
      const [ic, head, sub] = items[i];
      card(s, 0.6, y, 12.1, 0.88, "Agenda row " + (i + 1));
      tb(s, `0${i + 1}`, { x: 0.85, y, w: 0.8, h: 0.88, fontSize: 28, bold: true, color: C.accent1, fontFace: THEME.headFontFace, valign: "middle" });
      await iconCircle(s, ic, 1.75, y + 0.14, 0.6, C.text2, H.accent1, head);
      tb(s, head, { x: 2.6, y, w: 4.2, h: 0.88, fontSize: 20, bold: true, color: C.text2, valign: "middle" });
      tb(s, sub, { x: 6.8, y, w: 5.7, h: 0.88, fontSize: 16, color: C.accent5, valign: "middle" });
    }
    s.addNotes("About 60 minutes. Ask at the start: who has the EVO app? Who has checked their data access history this week?");
  }

  // ================= PART 1 =================
  const S1 = "What happened";
  pres.addSection({ title: S1 });
  await dividerSlide(S1, "01", "What happened", "19–23 October 2026: data about one in three Moldovans", fa.FaCalendarAlt, "Calendar");

  // 3. Big numbers
  {
    const s = contentSlide("The breach in numbers", S1);
    const stats = [
      ["~1 million", "personal ID numbers (IDNP) queried by the attackers", C.accent4],
      ["~750,000", "people whose data sets were extracted", C.accent3],
      ["1 in 3", "Moldovans affected (population ~2.4 million)", C.accent1],
      ["82 + 30", "public services and information systems taken offline", C.accent2],
    ];
    for (let i = 0; i < stats.length; i++) {
      const x = 0.6 + i * 3.1;
      card(s, x, 1.65, 2.8, 3.2, "Stat card " + (i + 1));
      tb(s, stats[i][0], { x: x + 0.25, y: 1.95, w: 2.4, h: 1.0, fontSize: 36, bold: true, color: stats[i][2], fontFace: THEME.headFontFace, valign: "middle" });
      tb(s, stats[i][1], { x: x + 0.25, y: 3.1, w: 2.35, h: 1.5, fontSize: 16, color: C.text1 });
    }
    tb(s, [
      { text: "5 days ", options: { bold: true, color: C.accent3 } },
      { text: "passed between the start of the attack (19 October) and the moment it was stopped (23 October, 11:30). ", options: { color: C.text1 } },
      { text: "Source: AGE and STISC, preliminary findings.", options: { italic: true, color: C.accent5 } },
    ], { x: 0.6, y: 5.3, w: 12.1, h: 0.9, fontSize: 17 });
    s.addNotes("Numbers are preliminary (AGE statement of 27 October, STISC via Moldpres). About one million IDNPs were queried, data sets for about 750,000 were extracted. Not every person had every category of data exposed. Emphasize the 5 days: the attack ran for days before it was stopped.");
  }

  // 4. Timeline
  {
    const s = contentSlide("Timeline", S1);
    const ev = [
      ["19–23 Oct", "~1 million IDNPs queried, data extracted"],
      ["23 Oct, 11:30", "Extraction stopped. 82 services and 30+ systems disconnected"],
      ["24 Oct", "Attack announced publicly"],
      ["25 Oct, 02:00", "New forced-access attempt blocked"],
      ["27 Oct, 07:00", "New extraction attempts fail. AGE publishes details"],
      ["28 Oct", "AGE director steps down. Two MPass login methods paused"],
    ];
    const w = 12.1 / ev.length;
    s.addShape(pres.shapes.LINE, { x: 0.6 + w / 2, y: 2.55, w: 12.1 - w, h: 0, line: { color: H.accent5, width: 2 }, objectName: "Timeline line" });
    for (let i = 0; i < ev.length; i++) {
      const cx = 0.6 + i * w + w / 2;
      const red = i < 2;
      tb(s, ev[i][0], { x: cx - w / 2 + 0.05, y: 1.6, w: w - 0.1, h: 0.6, fontSize: 15, bold: true, color: red ? C.accent3 : C.text2, align: "center", valign: "bottom" });
      s.addShape(pres.shapes.OVAL, { x: cx - 0.16, y: 2.39, w: 0.32, h: 0.32, fill: { color: red ? C.accent3 : C.accent2 }, line: { color: H.lt1, width: 2 }, objectName: "Timeline dot " + (i + 1) });
      card(s, cx - w / 2 + 0.07, 3.0, w - 0.14, 1.6, "Timeline card " + (i + 1));
      tb(s, ev[i][1], { x: cx - w / 2 + 0.2, y: 3.15, w: w - 0.4, h: 1.35, fontSize: 14, color: C.text1 });
    }
    tb(s, [
      { text: "Also on 28 October: ", options: { bold: true, color: C.text2 } },
      { text: "the Prime Minister said Moldova is the target of “coordinated and large-scale cyberattacks” and that possible involvement of foreign state actors is being examined. No attribution has been made public.", options: { color: C.text1 } },
    ], { x: 0.6, y: 4.95, w: 12.1, h: 1.0, fontSize: 16 });
    s.addNotes("Timeline from the AGE statement (27 October) and press (Moldpres, The Record, NewsMaker, Pro TV). The two MPass methods temporarily disabled were two-step authentication and EVOSign login. Note the attackers came back twice after being stopped: attackers do not give up after the first block.");
  }

  // 5. Data exposed
  {
    const s = contentSlide("What data was exposed", S1);
    const cats = [
      [fa.FaIdCard, "Identity", "Name, date and place of birth, IDNP, photo", C.accent4, H.accent4],
      [fa.FaUsers, "Family", "Data about first-degree relatives", C.accent4, H.accent4],
      [fa.FaHeartbeat, "Health and social", "Insurance status, family doctor, benefits, pension, disability", C.accent4, H.accent4],
      [fa.FaCarSide, "Documents and assets", "Driving licence, vehicles, companies owned or managed", C.accent4, H.accent4],
      [fa.FaPiggyBank, "Financial", "IBAN, tax debts, MPay payment history", C.accent3, H.accent3],
      [fa.FaCheckCircle, "Not affected", "Bank card data. No changes found in state registers", C.accent2, H.accent2],
    ];
    for (let i = 0; i < cats.length; i++) {
      const col = i % 3, row = Math.floor(i / 3);
      const x = 0.6 + col * 4.15, y = 1.55 + row * 2.45;
      const [ic, head, txt, colr] = cats[i];
      card(s, x, y, 3.8, 2.2, head + " card");
      await iconCircle(s, ic, x + 0.3, y + 0.3, 0.7, colr, "FFFFFF", head);
      tb(s, head, { x: x + 1.15, y: y + 0.3, w: 2.5, h: 0.7, fontSize: 19, bold: true, color: i === 5 ? C.accent2 : C.text2, valign: "middle" });
      tb(s, txt, { x: x + 0.3, y: y + 1.15, w: 3.25, h: 0.95, fontSize: 15, color: C.text1 });
    }
    tb(s, "Not every affected person had every category exposed (AGE, STISC).", { x: 0.6, y: 6.5, w: 12.1, h: 0.35, fontSize: 13, italic: true, color: C.accent5 });
    s.addNotes("Categories from the AGE statement. Stress what was NOT affected: card data is not stored by the government platform, and knowing an IBAN does not allow anyone to take money from an account (STISC). But combined, this data is a perfect toolkit for social engineering, which is the topic of part 3.");
  }

  // ================= PART 2 =================
  const S2 = "How it happened";
  pres.addSection({ title: S2 });
  await dividerSlide(S2, "02", "How it happened", "What is publicly known so far", fa.FaSearch, "Search");

  // 6. Attack chain
  {
    const s = contentSlide("The attack, as described by the authorities", S2);
    const nodes = [
      [fa.FaUserSecret, "Attackers", "Automated requests from the internet", C.accent3],
      [fa.FaGlobe, "Public e-services", "Reachable without authentication", C.accent1],
      [fa.FaDatabase, "Interconnected state registers", "Data returned for each IDNP", C.accent4],
    ];
    const xs = [0.6, 4.75, 8.9];
    for (let i = 0; i < 3; i++) {
      const [ic, head, sub, col] = nodes[i];
      card(s, xs[i], 1.55, 3.8, 2.0, head + " node");
      await iconCircle(s, ic, xs[i] + 0.3, 1.85, 0.85, col, "FFFFFF", head);
      tb(s, head, { x: xs[i] + 1.35, y: 1.8, w: 2.3, h: 0.8, fontSize: 18, bold: true, color: C.text2, valign: "middle" });
      tb(s, sub, { x: xs[i] + 1.35, y: 2.6, w: 2.3, h: 0.8, fontSize: 14, color: C.text1 });
      if (i < 2) s.addShape(pres.shapes.LINE, { x: xs[i] + 3.85, y: 2.55, w: 0.3, h: 0, line: { color: H.accent5, width: 2.5, endArrowType: "triangle" }, objectName: "Chain arrow " + (i + 1) });
    }
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 4.75, y: 3.75, w: 7.95, h: 0.55, rectRadius: 0.27, fill: { color: C.accent1 }, line: { type: "none" }, objectName: "Loop label" });
    tb(s, "↻  Data obtained was used to make further extractions", { x: 4.75, y: 3.75, w: 7.95, h: 0.55, fontSize: 15, bold: true, color: C.text2, align: "center", valign: "middle" });
    const facts = [
      ["Entry point: ", "vulnerabilities in services accessible without authentication.", "AGE"],
      ["Chaining: ", "the data obtained was then used for further extractions from connected systems.", "AGE"],
      ["Scale: ", "about 1 million IDNPs queried; about 11 million requests reported.", "AGE, The Record"],
      ["Audit trail: ", "one e-service recorded the wrong legal basis, so some people saw misleading entries in their EVO access history.", "AGE"],
    ];
    tb(s, facts.map((f, i) => ({ text: "", options: {} })).length ? facts.flatMap((f, i) => [
      { text: f[0], options: { bold: true, color: C.text2, bullet: true } },
      { text: f[1] + " ", options: { color: C.text1 } },
      { text: `(${f[2]})`, options: { italic: true, color: C.accent5, breakLine: i < facts.length - 1 } },
    ]) : [], { x: 0.6, y: 4.6, w: 12.1, h: 2.2, fontSize: 15, paraSpaceAfter: 6 });
    s.addNotes("Only what was officially stated or reported. Full technical details are not public yet. Risky Bulletin reported that the entry point was a logging service exposed online without authentication; AGE describes it as services accessible without authentication. The key idea: no password was cracked. The attackers simply asked public services for data, and the services answered.");
  }

  // 7. Lessons / weaknesses
  {
    const s = contentSlide("Four weaknesses the facts point to", S2);
    const w = [
      [fa.FaLock, "No login in front of personal data", "Services returned data about any person without checking who was asking.", "A01 Broken Access Control"],
      [fa.FaTachometerAlt, "No effective limits", "Around a million lookups over five days were possible before the attack was stopped.", "A09 Logging & Alerting"],
      [fa.FaLink, "Too much data per query", "One query unlocked data from many connected registers: identity, family, money.", "A06 Insecure Design"],
      [fa.FaFileAlt, "Unreliable audit trail", "A wrong legal basis was logged, so the access history misled citizens.", "A09 Logging & Alerting"],
    ];
    for (let i = 0; i < w.length; i++) {
      const col = i % 2, row = Math.floor(i / 2);
      const x = 0.6 + col * 6.2, y = 1.55 + row * 2.45;
      const [ic, head, txt, tag] = w[i];
      card(s, x, y, 5.9, 2.2, head + " card");
      await iconCircle(s, ic, x + 0.3, y + 0.3, 0.75, C.text2, H.accent1, head);
      tb(s, head, { x: x + 1.25, y: y + 0.3, w: 4.4, h: 0.75, fontSize: 18, bold: true, color: C.text2, valign: "middle" });
      tb(s, txt, { x: x + 0.3, y: y + 1.15, w: 5.3, h: 0.6, fontSize: 15, color: C.text1 });
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: x + 0.3, y: y + 1.72, w: 3.0, h: 0.34, rectRadius: 0.17, fill: { color: C.accent3 }, line: { type: "none" }, objectName: head + " tag" });
      tb(s, "OWASP " + tag, { x: x + 0.3, y: y + 1.72, w: 3.0, h: 0.34, fontSize: 11, bold: true, color: C.background1, align: "center", valign: "middle" });
    }
    tb(s, "This is our analysis of the public information, not an official finding.", { x: 0.6, y: 6.5, w: 12.1, h: 0.35, fontSize: 13, italic: true, color: C.accent5 });
    s.addNotes("Connect each weakness to the OWASP Top 10 from Lesson 1. Be honest with students: these are reasonable conclusions from the public facts, not the official investigation result. A01 is number one in the OWASP list for exactly this reason: it is very common.");
  }

  // ================= PART 3 =================
  const S3 = "The risks";
  pres.addSection({ title: S3 });
  await dividerSlide(S3, "03", "The risks", "Stolen data does not expire", fa.FaExclamationTriangle, "Warning");

  // 8. Risks for citizens
  {
    const s = contentSlide("What criminals can do with this data", S3);
    const r = [
      [fa.FaPhoneAlt, "Fake bank calls", "They know your IBAN and recent payments, so the call sounds real."],
      [fa.FaUsers, "“Your son is in trouble”", "They know your relatives' names and can invent a believable emergency."],
      [fa.FaUserTie, "Fake officials", "Tax debts, cars, pension: perfect details for a fake inspector or ministry."],
      [fa.FaHeartbeat, "Targeting the vulnerable", "Pension and disability data shows who is easiest to pressure."],
      [fa.FaIdCard, "Identity fraud", "Name, IDNP, birth date and photo help impersonate you."],
      [fa.FaHistory, "Risk for years", "You can change a password, but not your birth date, IDNP or family."],
    ];
    for (let i = 0; i < r.length; i++) {
      const col = i % 3, row = Math.floor(i / 3);
      const x = 0.6 + col * 4.15, y = 1.55 + row * 2.6;
      const [ic, head, txt] = r[i];
      card(s, x, y, 3.8, 2.35, head + " card");
      await iconCircle(s, ic, x + 0.3, y + 0.3, 0.7, C.accent3, "FFFFFF", head);
      tb(s, head, { x: x + 1.15, y: y + 0.3, w: 2.5, h: 0.7, fontSize: 17, bold: true, color: C.text2, valign: "middle" });
      tb(s, txt, { x: x + 0.3, y: y + 1.15, w: 3.25, h: 1.05, fontSize: 15, color: C.text1 });
    }
    s.addNotes("The government itself warned that the data can be used to impersonate banks or public institutions to obtain passwords, codes or money. The last card is the most important idea: personal data leaks are permanent. Ask students: which of these would work best on your grandparents?");
  }

  // 9. Anatomy of a scam
  {
    const s = contentSlide("Anatomy of a scam built on leaked data", S3);
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.6, y: 1.5, w: 5.6, h: 5.0, rectRadius: 0.3, fill: { color: C.text2 }, line: { type: "none" }, objectName: "Phone frame" });
    tb(s, "SMS  ·  necunoscut", { x: 0.9, y: 1.7, w: 5.0, h: 0.35, fontSize: 12, color: C.accent5, align: "center" });
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.9, y: 2.2, w: 5.0, h: 3.9, rectRadius: 0.2, fill: { color: C.background2 }, line: { type: "none" }, objectName: "Message bubble" });
    tb(s, [
      { text: "Stimată doamnă ", options: { color: C.text1 } },
      { text: "Maria Popescu", options: { bold: true, color: C.accent3 } },
      { text: ", am observat o plată suspectă din contul ", options: { color: C.text1 } },
      { text: "MD24…7731", options: { bold: true, color: C.accent3 } },
      { text: " după achitarea ", options: { color: C.text1 } },
      { text: "impozitului auto din 12.10", options: { bold: true, color: C.accent3 } },
      { text: ". Pentru a bloca operațiunea ", options: { color: C.text1 } },
      { text: "în 15 minute", options: { bold: true, color: C.accent1 } },
      { text: ", ", options: { color: C.text1 } },
      { text: "comunicați codul primit prin SMS", options: { bold: true, color: C.accent1 } },
      { text: " operatorului care vă va suna.", options: { color: C.text1 } },
    ], { x: 1.15, y: 2.4, w: 4.5, h: 3.5, fontSize: 17 });
    const flags = [
      [C.accent3, "Real data from the leak", "Name, IBAN, a real MPay payment: it feels legitimate."],
      [C.accent1, "Urgency", "“15 minutes” leaves no time to think or check."],
      [C.accent1, "Asks for a code", "No bank or institution ever asks for your SMS code."],
      [C.accent2, "The right reaction", "Hang up. Call the bank on the number printed on your card."],
    ];
    for (let i = 0; i < flags.length; i++) {
      const y = 1.5 + i * 1.27;
      card(s, 6.6, y, 6.1, 1.1, flags[i][1] + " flag");
      s.addShape(pres.shapes.OVAL, { x: 6.8, y: y + 0.35, w: 0.4, h: 0.4, fill: { color: flags[i][0] }, line: { type: "none" }, objectName: flags[i][1] + " dot" });
      tb(s, flags[i][1], { x: 7.4, y: y + 0.1, w: 5.1, h: 0.4, fontSize: 17, bold: true, color: C.text2 });
      tb(s, flags[i][2], { x: 7.4, y: y + 0.52, w: 5.1, h: 0.5, fontSize: 15, color: C.text1 });
    }
    s.addNotes("This message is invented for teaching, but it uses exactly the categories of leaked data: name, IBAN and MPay payment history. Read it out loud and ask students what makes it convincing. Then go through the red flags. Recommend they show this slide to their parents and grandparents.");
  }

  // 10. Context
  {
    const s = contentSlide("Not an isolated incident", S3);
    const st = [
      ["2nd", "major breach of Moldovan state systems in 2026. The healthcare database was breached in April", "Risky Bulletin", C.accent3],
      ["2,600", "of 8,200 cyber incidents analyzed in the EU hit public administration, the most targeted sector", "ENISA, September 2026", C.accent1],
      ["Hybrid", "authorities see the attack as possibly part of a wider hybrid campaign against Moldova", "Government of Moldova", C.accent4],
    ];
    for (let i = 0; i < st.length; i++) {
      const x = 0.6 + i * 4.15;
      card(s, x, 1.6, 3.8, 3.6, "Context card " + (i + 1));
      tb(s, st[i][0], { x: x + 0.3, y: 1.85, w: 3.2, h: 1.1, fontSize: 48, bold: true, color: st[i][3], fontFace: THEME.headFontFace, valign: "middle" });
      tb(s, st[i][1], { x: x + 0.3, y: 3.05, w: 3.2, h: 1.6, fontSize: 15, color: C.text1 });
      tb(s, st[i][2], { x: x + 0.3, y: 4.7, w: 3.2, h: 0.35, fontSize: 11, italic: true, color: C.accent5 });
    }
    tb(s, [
      { text: "Why governments? ", options: { bold: true, color: C.text2 } },
      { text: "They hold complete, verified data about everyone, connect many old and new systems, and are a political target.", options: { color: C.text1 } },
    ], { x: 0.6, y: 5.6, w: 12.1, h: 0.9, fontSize: 17 });
    s.addNotes("Context only. Do not speculate about who is behind the attack: no attribution has been made public. Point out that public administration is the most attacked sector in the EU too, so this is not a Moldovan-only problem.");
  }

  // ================= PART 4 =================
  const S4 = "How to prevent it";
  pres.addSection({ title: S4 });
  await dividerSlide(S4, "04", "How to prevent it", "Citizens, developers and organizations", fa.FaShieldAlt, "Shield");

  // 11. Citizens
  {
    const s = contentSlide("What you can do as a citizen, today", S4);
    const a = [
      [fa.FaMobileAlt, "Check your data access history", "EVO app → Contul tău → Istoricul accesării datelor"],
      [fa.FaKey, "Never share passwords or SMS codes", "No bank, police officer or ministry will ever ask for them"],
      [fa.FaPhoneAlt, "Hang up and call back", "Use the official number from the card or website, not the one from the message"],
      [fa.FaUsers, "Agree on a family code word", "Ask for it when someone “from the family” urgently needs money"],
      [fa.FaBell, "Turn on bank notifications", "Report any transaction you do not recognize immediately"],
      [fa.FaEnvelope, "Do not open unsolicited links", "And never approve payments requested by strangers"],
    ];
    for (let i = 0; i < a.length; i++) {
      const col = i % 2, row = Math.floor(i / 2);
      const x = 0.6 + col * 6.2, y = 1.55 + row * 1.7;
      const [ic, head, txt] = a[i];
      card(s, x, y, 5.9, 1.5, head + " card");
      await iconCircle(s, ic, x + 0.25, y + 0.3, 0.85, C.accent2, "FFFFFF", head);
      tb(s, head, { x: x + 1.3, y: y + 0.2, w: 4.4, h: 0.5, fontSize: 17, bold: true, color: C.text2, valign: "middle" });
      tb(s, txt, { x: x + 1.3, y: y + 0.72, w: 4.4, h: 0.65, fontSize: 14, color: C.text1 });
    }
    s.addNotes("Most of these are the official recommendations from AGE and STISC. The family code word is extra advice, simple and very effective against “family emergency” scams. Homework idea: each student checks their EVO access history and helps one older relative do the same.");
  }

  // 12. Developers & organizations
  {
    const s = contentSlide("What developers and organizations must do", S4);
    const cols = [
      ["Prevent", fa.FaLock, C.accent4, [
        "Authenticate every endpoint that returns personal data",
        "Authorize per record and per purpose: who may see what, and why",
        "Return the minimum data needed; mask IDNP and IBAN",
        "Least privilege between connected systems",
        "Keep an inventory: no forgotten or test services online",
      ]],
      ["Detect and respond", fa.FaBell, C.accent3, [
        "Rate limits and quotas per client and per API key",
        "Alerts on unusual volume: a million lookups is never normal",
        "Reliable audit logs: who accessed what, when and why",
        "A tested incident response plan",
        "Regular audits and penetration tests",
      ]],
    ];
    for (let c = 0; c < 2; c++) {
      const x = 0.6 + c * 6.2;
      const [head, ic, col, items] = cols[c];
      card(s, x, 1.55, 5.9, 4.3, head + " card");
      await iconCircle(s, ic, x + 0.3, 1.8, 0.75, col, "FFFFFF", head);
      tb(s, head, { x: x + 1.25, y: 1.8, w: 4.4, h: 0.75, fontSize: 22, bold: true, color: C.text2, valign: "middle" });
      tb(s, items.map((t, i) => ({ text: t, options: { bullet: true, breakLine: i < items.length - 1 } })),
        { x: x + 0.35, y: 2.8, w: 5.3, h: 3.6, fontSize: 16, color: C.text1, paraSpaceAfter: 10 });
    }
    s.addNotes("Each line on the left would have closed the door; each line on the right would have caught the attack much earlier. Ask: which single control would have helped the most? Usually the answer is authentication plus per-client rate limits and alerting.");
  }

  // 13. Incident response
  {
    const s = contentSlide("The response: what worked, what is unknown", S4);
    const cols = [
      ["What worked", fa.FaCheckCircle, C.accent2, H.accent2, [
        "The attack was detected and stopped",
        "Fast containment: 82 services and 30+ systems offline",
        "Public disclosure the next day, with clear advice",
        "Joint investigation: AGE, STISC, ASC, SIS, police",
        "Services return only after new security checks",
      ]],
      ["Open questions", fa.FaQuestionCircle, C.accent1, H.accent1, [
        "Why could these services be used without login?",
        "Why did five days of mass queries not trigger alarms sooner?",
        "Do other systems share the same design?",
        "Exactly how many people were affected?",
        "Who is behind the attack?",
      ]],
    ];
    for (let c = 0; c < 2; c++) {
      const x = 0.6 + c * 6.2;
      const [head, ic, col, hex, items] = cols[c];
      card(s, x, 1.55, 5.9, 4.3, head + " card");
      s.addImage({ data: await icon(ic, hex), x: x + 0.3, y: 1.8, w: 0.65, h: 0.65, objectName: head + " icon", altText: head });
      tb(s, head, { x: x + 1.15, y: 1.8, w: 4.5, h: 0.65, fontSize: 22, bold: true, color: col, valign: "middle" });
      tb(s, items.map((t, i) => ({ text: t, options: { bullet: true, breakLine: i < items.length - 1 } })),
        { x: x + 0.35, y: 2.75, w: 5.3, h: 3.7, fontSize: 16, color: C.text1, paraSpaceAfter: 10 });
    }
    s.addNotes("Be fair: the response after detection was fast and transparent. The open questions are not accusations; they are the questions any security team asks after an incident. The investigation will answer some of them.");
  }

  // 14. Link to labs
  {
    const s = contentSlide("You will practice exactly this in our labs", S4);
    const rows = [
      ["Lab 1", "Threat model and attack surface: which endpoints are public and what they expose"],
      ["Lab 2", "Input validation and safe database queries"],
      ["Lab 3", "Authentication, per-record authorization (IDOR) and rate limiting on login"],
      ["Lab 4", "Logs and Fail2ban: detecting and blocking abusive clients automatically"],
      ["Lab 5", "Nginx rate limiting, HTTPS, security headers, secrets kept out of code"],
    ];
    for (let i = 0; i < rows.length; i++) {
      const y = 1.55 + i * 0.98;
      card(s, 0.6, y, 12.1, 0.82, rows[i][0] + " row");
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.8, y: y + 0.16, w: 1.15, h: 0.5, rectRadius: 0.1, fill: { color: i < 3 ? C.accent2 : C.accent4 }, line: { type: "none" }, objectName: rows[i][0] + " badge" });
      tb(s, rows[i][0], { x: 0.8, y: y + 0.16, w: 1.15, h: 0.5, fontSize: 16, bold: true, color: C.background1, align: "center", valign: "middle" });
      tb(s, rows[i][1], { x: 2.2, y, w: 10.3, h: 0.82, fontSize: 17, color: C.text1, valign: "middle" });
    }
    tb(s, [
      { text: "The Notes app has the same kind of flaw: ", options: { bold: true, color: C.text2 } },
      { text: "in Lab 3 you will find an endpoint that returns other people's notes to anyone who asks.", options: { color: C.text1 } },
    ], { x: 0.6, y: 6.5, w: 12.1, h: 0.45, fontSize: 16 });
    s.addNotes("Make the link explicit: the IDOR in the Notes app (GET /api/notes/:id without an owner check) is the small version of what happened here. The defenses students implement in the labs are the same ones listed on the previous slides.");
  }

  // ================= PART 5 =================
  const S5 = "Discussion";
  pres.addSection({ title: S5 });
  await dividerSlide(S5, "05", "Discussion", "Think like the security team", fa.FaChalkboardTeacher, "Teacher");

  // 15. Class exercise
  {
    const s = contentSlide("Class exercise: you are the security team (20 min)", S5);
    const q = [
      ["Map the attack surface", "Which public services return personal data? Which need login and which do not?"],
      ["Choose three controls", "If you could deploy only three defenses tomorrow, which ones and why?"],
      ["Write one alert rule", "What exact condition should have raised an alarm on 19 October?"],
      ["Notify the citizens", "Write a short, honest SMS to affected people: what happened and what to do."],
    ];
    for (let i = 0; i < q.length; i++) {
      const col = i % 2, row = Math.floor(i / 2);
      const x = 0.6 + col * 6.2, y = 1.55 + row * 2.45;
      card(s, x, y, 5.9, 2.2, q[i][0] + " card");
      tb(s, String(i + 1), { x: x + 0.3, y: y + 0.2, w: 0.8, h: 1.0, fontSize: 48, bold: true, color: C.accent1, fontFace: THEME.headFontFace, valign: "middle" });
      tb(s, q[i][0], { x: x + 1.2, y: y + 0.3, w: 4.5, h: 0.6, fontSize: 19, bold: true, color: C.text2, valign: "middle" });
      tb(s, q[i][1], { x: x + 1.2, y: y + 0.95, w: 4.5, h: 1.1, fontSize: 15, color: C.text1 });
    }
    tb(s, "Groups of 3–4 students. Each group presents one answer in 2 minutes.", { x: 0.6, y: 6.5, w: 12.1, h: 0.35, fontSize: 14, italic: true, color: C.accent5 });
    s.addNotes("Example of a good alert rule: more than 100 different IDNPs queried by the same client or IP within one hour. Example of a good SMS: short, no links, tells people how to check EVO and which number to call. Discuss why an SMS with a link would itself look like phishing.");
  }

  // 16. Key takeaways
  {
    const s = contentSlide("Key takeaways", S5);
    const k = [
      [fa.FaLock, "No authentication = public data", "If an endpoint answers anyone, assume attackers will ask it a million times."],
      [fa.FaTachometerAlt, "Limits and alerts are not optional", "Rate limiting and monitoring turn days of silent theft into minutes."],
      [fa.FaLink, "Connected systems multiply the damage", "Give each system only the data it really needs."],
      [fa.FaUsers, "People are the final target", "Leaked data powers scams for years. Teach your family to verify, not trust."],
    ];
    for (let i = 0; i < k.length; i++) {
      const y = 1.55 + i * 1.27;
      card(s, 0.6, y, 12.1, 1.1, k[i][1] + " row");
      await iconCircle(s, k[i][0], 0.8, y + 0.17, 0.76, C.text2, H.accent1, k[i][1]);
      tb(s, k[i][1], { x: 1.85, y: y + 0.08, w: 10.6, h: 0.5, fontSize: 19, bold: true, color: C.text2, valign: "middle" });
      tb(s, k[i][2], { x: 1.85, y: y + 0.55, w: 10.6, h: 0.45, fontSize: 15, color: C.text1, valign: "middle" });
    }
    s.addNotes("Close with the human side: the most important action this week is helping older relatives recognize scams that use their real data.");
  }

  // 17. Sources
  {
    const s = contentSlide("Sources", S5);
    const src = [
      ["AGE, official statement, 27 Oct 2026", "egov.md/ro/node/41145"],
      ["Moldpres, STISC findings, 28 Oct 2026", "moldpres.md"],
      ["The Record, 28 Oct 2026", "therecord.media/moldova-investigates-possible-foreign-involvement-cyberattacks"],
      ["Risky Bulletin, 28 Oct 2026", "news.risky.biz"],
      ["Diez.md, how to check your data, 27 Oct 2026", "diez.md"],
      ["NewsMaker, MPass changes, 28 Oct 2026", "newsmaker.md"],
      ["Pro TV Chișinău, AGE leadership change, 28 Oct 2026", "protv.md"],
      ["ENISA Threat Landscape, 22 Sep 2026 (via Moldpres)", "enisa.europa.eu"],
    ];
    tb(s, src.flatMap((r, i) => [
      { text: r[0] + "  ", options: { bold: true, color: C.text2, bullet: true } },
      { text: r[1], options: { color: C.accent5, breakLine: i < src.length - 1 } },
    ]), { x: 0.6, y: 1.55, w: 12.1, h: 4.6, fontSize: 15, paraSpaceAfter: 8 });
    tb(s, "Information as of 28 October 2026. The investigation is ongoing and figures are preliminary.", { x: 0.6, y: 6.4, w: 12.1, h: 0.4, fontSize: 13, italic: true, color: C.accent5 });
    s.addNotes("Encourage students to read the official AGE statement themselves and to compare how different outlets reported the same facts (for example, 'a quarter' versus 'a third' of the population).");
  }

  // 18. Closing
  {
    const s = titleSlide(S5, "Check your EVO access history today", "Then help one older relative do the same  ·  Questions?", await icon(fa.FaMobileAlt, H.accent1), "Phone");
    s.addNotes("Practical homework: check your own access history in EVO and talk to one older family member about scam calls that use real personal data.");
  }

  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  console.log("written", OUT);
})().catch((e) => { console.error(e); process.exit(1); });
