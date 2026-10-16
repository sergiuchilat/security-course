const pptxgen = require("pptxgenjs");
const React = require("react");
const { renderToStaticMarkup } = require("react-dom/server");
const sharp = require("sharp");
const fa = require("react-icons/fa");
const { applyTheme } = require("/Users/sergiuchilat/.claude/skills/synced/e25047ed-7c97-4849-9da1-c59bbd817501_b100ab20-30b2-4fd4-bf29-fbdff3299f26/pptx/scripts/apply_theme.js");

const OUT = "../Lesson01_Introduction_to_Security.pptx";

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
pres.title = "Lesson 1 – Introduction to Security";
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
const FOOTER = "Web Application & Infrastructure Security  ·  Lesson 1";

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
  tb(s, `PART ${num} OF 06`, { x: 3.8, y: 2.45, w: 6.5, h: 0.4, fontSize: 14, bold: true, charSpacing: 3, color: C.accent5 });
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
  const S1 = "Introduction";
  pres.addSection({ title: S1 });

  // 1. Title
  {
    const s = titleSlide(S1, "Introduction to Web & Infrastructure Security", "Lesson 1  ·  Core concepts, threats, principles and how this course works", await icon(fa.FaShieldAlt, H.accent1), "Shield");
    s.addNotes("Welcome. This course has two halves: infrastructure security (servers, networks, containers) and application security (the web app itself). Today is the conceptual foundation for all the labs.");
  }

  // 2. Agenda
  {
    const s = contentSlide("Today's agenda", S1);
    const items = [
      [fa.FaExclamationTriangle, "Why security matters", "Real costs and real numbers"],
      [fa.FaLock, "Core concepts", "CIA triad, AAA, threat · vulnerability · risk"],
      [fa.FaUserSecret, "Attackers and attack surface", "Who attacks, and where a web system is exposed"],
      [fa.FaLayerGroup, "Security principles", "Defense in depth, least privilege, and more"],
      [fa.FaBug, "OWASP Top 10 and CVE/CVSS", "How vulnerabilities are classified and scored"],
      [fa.FaBalanceScale, "Ethics and course rules", "What is allowed, labs, grading, homework"],
    ];
    const colX = [0.6, 6.85], rowY = [1.65, 3.4, 5.15];
    for (let i = 0; i < items.length; i++) {
      const x = colX[Math.floor(i / 3)], y = rowY[i % 3];
      const [ic, head, sub] = items[i];
      card(s, x, y, 5.85, 1.45, "Agenda card " + (i + 1));
      await iconCircle(s, ic, x + 0.3, y + 0.3, 0.85, C.text2, H.accent1, "Agenda " + (i + 1));
      tb(s, `${String(i + 1).padStart(2, "0")}  ${head}`, { x: x + 1.4, y: y + 0.3, w: 4.2, h: 0.45, fontSize: 20, bold: true, color: C.text2 });
      tb(s, sub, { x: x + 1.4, y: y + 0.8, w: 4.2, h: 0.4, fontSize: 15, color: C.accent5 });
    }
    s.addNotes("About 90 minutes. Ask students early: has anyone had an account hacked? It makes the topic personal.");
  }

  // 3. Why security matters
  {
    const s = contentSlide("Why security matters", S1);
    const stats = [
      ["$4.88M", "average cost of a data breach", "IBM Cost of a Data Breach, 2024", C.accent3],
      ["68%", "of breaches involve a human element", "Verizon DBIR, 2024", C.accent1],
      ["40,000+", "new CVEs published in one year", "NVD, 2024", C.accent2],
    ];
    for (let i = 0; i < 3; i++) {
      const x = 0.6 + i * 4.15;
      card(s, x, 1.7, 3.8, 3.3, "Stat card " + (i + 1));
      tb(s, stats[i][0], { x: x + 0.35, y: 2.0, w: 3.1, h: 1.2, fontSize: 52, bold: true, color: stats[i][3], fontFace: THEME.headFontFace, valign: "middle" });
      tb(s, stats[i][1], { x: x + 0.35, y: 3.3, w: 3.1, h: 0.9, fontSize: 18, color: C.text1 });
      tb(s, stats[i][2], { x: x + 0.35, y: 4.4, w: 3.1, h: 0.35, fontSize: 11, italic: true, color: C.accent5 });
    }
    tb(s, [
      { text: "Every web app is exposed. ", options: { bold: true, color: C.text2 } },
      { text: "Automated bots scan the whole internet continuously. A new server with SSH open usually gets its first login attempts within minutes, not days.", options: { color: C.text1 } },
    ], { x: 0.6, y: 5.45, w: 12.1, h: 1.0, fontSize: 18 });
    s.addNotes("The numbers are from public annual reports (IBM/Ponemon, Verizon DBIR, NIST NVD). Update them each year. The point: it is not 'if' but 'when', and most incidents involve people (phishing, weak passwords, misconfiguration), not movie-style hacking.");
  }

  // ================= CONCEPTS =================
  const S2 = "Core concepts";
  pres.addSection({ title: S2 });

  // 4. Section divider
  {
    await dividerSlide(S2, "02", "Core concepts", "The vocabulary we will use in every lab", fa.FaLock, "Lock");
  }

  // 5. CIA triad
  {
    const s = contentSlide("The CIA triad", S2);
    const cia = [
      [fa.FaEyeSlash, "Confidentiality", "Only authorized people can read the data.", "Attack: stolen database, sniffed password", "Defense: encryption (TLS), access control", C.accent4, H.accent4],
      [fa.FaCheckDouble, "Integrity", "Data is not changed without authorization.", "Attack: modified price in a request, defaced site", "Defense: validation, hashes, signatures", C.accent2, H.accent2],
      [fa.FaServer, "Availability", "The system works when users need it.", "Attack: DoS/DDoS, ransomware", "Defense: rate limiting, backups, redundancy", C.accent1, H.accent1],
    ];
    for (let i = 0; i < 3; i++) {
      const x = 0.6 + i * 4.15;
      const [ic, head, def, att, dfn, col] = cia[i];
      card(s, x, 1.6, 3.8, 4.9, head + " card");
      await iconCircle(s, ic, x + 0.35, 1.9, 0.95, col, "FFFFFF", head);
      tb(s, head, { x: x + 0.35, y: 3.0, w: 3.1, h: 0.5, fontSize: 24, bold: true, color: C.text2, fontFace: THEME.headFontFace });
      tb(s, def, { x: x + 0.35, y: 3.55, w: 3.1, h: 0.8, fontSize: 16, color: C.text1 });
      tb(s, [
        { text: att, options: { bullet: true, breakLine: true, paraSpaceAfter: 6 } },
        { text: dfn, options: { bullet: true } },
      ], { x: x + 0.35, y: 4.5, w: 3.2, h: 1.8, fontSize: 15, color: C.accent5 });
    }
    s.addNotes("Every security problem in this course breaks at least one of these three. Exercise: ask students which letter is broken by (a) a leaked password list, (b) changing a grade in the database, (c) a DDoS on the exam platform.");
  }

  // 6. AAA
  {
    const s = contentSlide("Beyond CIA: who are you, and what may you do?", S2);
    const aaa = [
      [fa.FaFingerprint, "Authentication", "Proving who you are", "Passwords, keys, MFA, certificates"],
      [fa.FaUserCheck, "Authorization", "Deciding what you are allowed to do", "Roles, permissions, ownership checks"],
      [fa.FaHistory, "Accounting / Auditing", "Recording who did what, and when", "Logs, audit trails, monitoring"],
    ];
    for (let i = 0; i < 3; i++) {
      const y = 1.6 + i * 1.55;
      const [ic, head, what, how] = aaa[i];
      await iconCircle(s, ic, 0.6, y + 0.1, 1.0, C.text2, H.accent1, head);
      tb(s, head, { x: 1.9, y: y + 0.05, w: 4.6, h: 0.5, fontSize: 22, bold: true, color: C.text2 });
      tb(s, what, { x: 1.9, y: y + 0.6, w: 4.6, h: 0.5, fontSize: 16, color: C.text1 });
      tb(s, how, { x: 6.7, y: y + 0.3, w: 6.0, h: 0.6, fontSize: 16, italic: true, color: C.accent5, valign: "middle" });
    }
    card(s, 0.6, 6.15, 12.1, 0.7, "Key point card", C.background2);
    tb(s, [
      { text: "Common mistake: ", options: { bold: true, color: C.accent3 } },
      { text: "checking that a user is logged in (authentication) but not that the record belongs to them (authorization).", options: { color: C.text1 } },
    ], { x: 0.9, y: 6.15, w: 11.6, h: 0.7, fontSize: 16, valign: "middle" });
    s.addNotes("AAA is the backbone of Lab 3. Broken access control is #1 in the OWASP Top 10, and it is almost always the authorization step that is missing.");
  }

  // 7. Threat / vulnerability / risk
  {
    const s = contentSlide("Threat, vulnerability, risk", S2);
    const defs = [
      ["Asset", "What we protect", "User database, server, reputation", C.accent4],
      ["Vulnerability", "A weakness", "SSH open with password login allowed", C.accent1],
      ["Threat", "Who or what can exploit it", "Bots trying common passwords", C.accent3],
      ["Risk", "Chance × damage", "Server takeover, data leak", C.accent2],
    ];
    for (let i = 0; i < 4; i++) {
      const x = 0.6 + i * 3.1;
      const [head, what, ex, col] = defs[i];
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y: 1.6, w: 2.8, h: 0.75, rectRadius: 0.12, fill: { color: col }, line: { type: "none" }, objectName: head + " header" });
      tb(s, head, { x, y: 1.6, w: 2.8, h: 0.75, fontSize: 22, bold: true, color: C.background1, align: "center", valign: "middle" });
      card(s, x, 2.5, 2.8, 2.2, head + " card");
      tb(s, what, { x: x + 0.25, y: 2.7, w: 2.3, h: 0.7, fontSize: 17, bold: true, color: C.text2 });
      tb(s, "e.g. " + ex, { x: x + 0.25, y: 3.45, w: 2.3, h: 1.1, fontSize: 15, italic: true, color: C.accent5 });
    }
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.6, y: 5.1, w: 5.4, h: 1.3, rectRadius: 0.15, fill: { color: C.text2 }, line: { type: "none" }, objectName: "Formula box" });
    tb(s, "Risk = Likelihood × Impact", { x: 0.6, y: 5.1, w: 5.4, h: 1.3, fontSize: 26, bold: true, color: C.accent1, align: "center", valign: "middle", fontFace: THEME.headFontFace });
    tb(s, [
      { text: "Control (fix): ", options: { bold: true, color: C.accent2 } },
      { text: "SSH keys only, Fail2ban, firewall. ", options: { color: C.text1, breakLine: true } },
      { text: "We rarely remove a threat. We reduce the vulnerability or the impact, and that lowers the risk.", options: { color: C.accent5 } },
    ], { x: 6.4, y: 5.1, w: 6.3, h: 1.3, fontSize: 16, valign: "middle" });
    s.addNotes("Use the SSH example because Lab 4 implements exactly this control (Fail2ban), on a real public server. Ask students for another example from a web app: a login form with no rate limit (vulnerability), credential-stuffing bots (threat), account takeover (risk).");
  }

  // 8. Who attacks
  {
    const s = contentSlide("Who attacks, and why", S2);
    const who = [
      [fa.FaRobot, "Automated bots", "Scan everything, try default passwords and known CVEs", "Very high"],
      [fa.FaMoneyBillWave, "Cybercriminals", "Money: ransomware, stolen data, fraud", "High"],
      [fa.FaUserTie, "Insiders", "Employees or ex-employees, by mistake or on purpose", "Medium"],
      [fa.FaBullhorn, "Hacktivists", "Political or social message, defacement, DDoS", "Medium"],
      [fa.FaFlag, "State-sponsored groups", "Espionage, sabotage, long-term access", "Low, high impact"],
    ];
    tb(s, "Attacker", { x: 1.6, y: 1.5, w: 3.4, h: 0.4, fontSize: 14, bold: true, color: C.accent5 });
    tb(s, "Motivation and methods", { x: 5.1, y: 1.5, w: 5.0, h: 0.4, fontSize: 14, bold: true, color: C.accent5 });
    tb(s, "How often you meet them", { x: 10.2, y: 1.5, w: 2.5, h: 0.4, fontSize: 14, bold: true, color: C.accent5 });
    for (let i = 0; i < who.length; i++) {
      const y = 2.0 + i * 0.95;
      const [ic, head, mot, freq] = who[i];
      card(s, 0.6, y, 12.1, 0.8, head + " row");
      await iconCircle(s, ic, 0.75, y + 0.1, 0.6, C.text2, H.accent1, head);
      tb(s, head, { x: 1.6, y, w: 3.4, h: 0.8, fontSize: 18, bold: true, color: C.text2, valign: "middle" });
      tb(s, mot, { x: 5.1, y, w: 5.0, h: 0.8, fontSize: 15, color: C.text1, valign: "middle" });
      tb(s, freq, { x: 10.2, y, w: 2.4, h: 0.8, fontSize: 15, bold: true, color: i === 0 ? C.accent3 : C.accent5, valign: "middle" });
    }
    s.addNotes("Most students imagine a targeted hacker. In reality, a small site will mostly see bots: automated, untargeted and constant. That is why basic hygiene (updates, strong auth, firewall) stops most attacks.");
  }

  // 9. Attack surface
  {
    const s = contentSlide("Attack surface of a web system", S2);
    const nodes = [
      [fa.FaLaptop, "Browser", "XSS, phishing, stolen cookies", "APP", C.accent2],
      [fa.FaGlobe, "Network", "Sniffing, MitM, DNS attacks", "SERVER", C.accent4],
      [fa.FaFire, "Firewall", "Open ports, weak rules", "SERVER", C.accent4],
      [fa.FaServer, "Web server", "Old versions, DoS, misconfig", "SERVER", C.accent4],
      [fa.FaCode, "Application", "Injection, broken auth, IDOR", "APP", C.accent2],
      [fa.FaDatabase, "Database", "Data leaks, weak credentials", "BOTH", C.accent1],
    ];
    const w = 1.75, gap = (12.1 - 6 * w) / 5;
    for (let i = 0; i < nodes.length; i++) {
      const x = 0.6 + i * (w + gap);
      const [ic, head, risk, tag, col] = nodes[i];
      card(s, x, 1.7, w, 3.9, head + " node");
      await iconCircle(s, ic, x + (w - 0.9) / 2, 1.95, 0.9, C.text2, H.accent1, head);
      tb(s, head, { x: x + 0.1, y: 3.0, w: w - 0.2, h: 0.55, fontSize: 16, bold: true, color: C.text2, align: "center", valign: "middle" });
      tb(s, risk, { x: x + 0.15, y: 3.6, w: w - 0.3, h: 1.2, fontSize: 14, color: C.text1, align: "center" });
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: x + 0.3, y: 4.95, w: w - 0.6, h: 0.4, rectRadius: 0.2, fill: { color: col }, line: { type: "none" }, objectName: head + " tag" });
      tb(s, tag, { x: x + 0.3, y: 4.95, w: w - 0.6, h: 0.4, fontSize: 12, bold: true, color: C.background1, align: "center", valign: "middle" });
      if (i < nodes.length - 1) {
        s.addShape(pres.shapes.LINE, { x: x + w + 0.04, y: 2.4, w: gap - 0.08, h: 0, line: { color: H.accent5, width: 2, endArrowType: "triangle" }, objectName: "Arrow " + (i + 1) });
      }
    }
    tb(s, [
      { text: "Attack surface ", options: { bold: true, color: C.text2 } },
      { text: "= every point where an attacker can send input or reach the system. Our job: make it ", options: { color: C.text1 } },
      { text: "smaller", options: { bold: true, color: C.accent3 } },
      { text: " and defend each point. APP = Part 1 (Labs 1–3), SERVER = Part 2 (Labs 4–5).", options: { color: C.text1 } },
    ], { x: 0.6, y: 5.95, w: 12.1, h: 0.85, fontSize: 16 });
    s.addNotes("Follow one HTTP request from the browser to the database. Every hop is a place where something can go wrong, and every hop belongs to one of the two parts of the course.");
  }

  // ================= PRINCIPLES =================
  const S3 = "Security principles";
  pres.addSection({ title: S3 });

  // 10. Defense in depth
  {
    const s = contentSlide("Defense in depth", S3);
    const layers = [
      ["Policies & people", C.accent5],
      ["Network: firewall, segmentation", C.accent4],
      ["Host: hardening, updates", C.accent2],
      ["Application: validation, auth", C.accent1],
      ["Data: encryption, backups", C.accent3],
    ];
    for (let i = 0; i < layers.length; i++) {
      const inset = i * 0.4;
      const x = 0.6 + inset, y = 1.55 + inset * 0.9, w = 6.4 - 2 * inset, h = 5.1 - 2 * inset * 0.9;
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, rectRadius: 0.15, fill: { color: layers[i][1] }, line: { color: H.lt1, width: 2 }, objectName: "Layer " + (i + 1) });
      tb(s, layers[i][0], { x: x + 0.2, y: y + 0.05, w: w - 0.4, h: 0.36, fontSize: 14, bold: true, color: C.background1, align: "center", valign: "middle" });
    }
    s.addImage({ data: await icon(fa.FaGem, "FFFFFF"), x: 3.4, y: 4.25, w: 0.8, h: 0.8, objectName: "Data gem icon", altText: "Valuable data at the center" });
    tb(s, "No single control is perfect.", { x: 7.5, y: 1.7, w: 5.2, h: 0.6, fontSize: 24, bold: true, color: C.text2, fontFace: THEME.headFontFace });
    tb(s, "Stack several independent layers, so an attacker who gets through one still meets the next.", { x: 7.5, y: 2.4, w: 5.2, h: 1.0, fontSize: 17, color: C.text1 });
    tb(s, [
      { text: "Example: SQL injection", options: { bold: true, color: C.text2, breakLine: true } },
      { text: "Input validation rejects bad data (Lab 2)", options: { bullet: true, breakLine: true } },
      { text: "Parameterized queries fix the code (Lab 2)", options: { bullet: true, breakLine: true } },
      { text: "DB user has minimal rights (least privilege)", options: { bullet: true, breakLine: true } },
      { text: "Logs show the attempts (Lab 4)", options: { bullet: true } },
    ], { x: 7.5, y: 3.6, w: 5.2, h: 2.6, fontSize: 16, color: C.text1, paraSpaceAfter: 6 });
    s.addNotes("Castle analogy: moat, walls, guards, vault. Input validation is not a substitute for parameterized queries; it is an extra layer.");
  }

  // 11. Principles grid
  {
    const s = contentSlide("Six principles to remember", S3);
    const pr = [
      [fa.FaKey, "Least privilege", "Give each user and process only the rights it needs, nothing more."],
      [fa.FaCompressArrowsAlt, "Minimize attack surface", "Close unused ports, remove unused services, packages and features."],
      [fa.FaCog, "Secure by default", "Default config must be safe: no default passwords, debug mode off."],
      [fa.FaFilter, "Never trust input", "Validate everything from users, headers, files and other services."],
      [fa.FaDoorClosed, "Fail securely", "On error, deny access and show no stack traces or internals."],
      [fa.FaSyncAlt, "Keep it updated", "Patch OS, libraries and containers. Old software = known CVEs."],
    ];
    for (let i = 0; i < pr.length; i++) {
      const col = i % 3, row = Math.floor(i / 3);
      const x = 0.6 + col * 4.15, y = 1.6 + row * 2.6;
      const [ic, head, txt] = pr[i];
      card(s, x, y, 3.8, 2.3, head + " card");
      await iconCircle(s, ic, x + 0.3, y + 0.3, 0.75, C.text2, H.accent1, head);
      tb(s, head, { x: x + 1.2, y: y + 0.3, w: 2.45, h: 0.75, fontSize: 18, bold: true, color: C.text2, valign: "middle" });
      tb(s, txt, { x: x + 0.3, y: y + 1.2, w: 3.25, h: 1.0, fontSize: 15, color: C.text1 });
    }
    s.addNotes("These principles come back in every lab. Lab 4 (securing the Linux server) is basically: least privilege + minimize attack surface + updates.");
  }

  // ================= VULNERABILITIES =================
  const S4 = "Classifying vulnerabilities";
  pres.addSection({ title: S4 });

  // 12. OWASP Top 10
  {
    const s = contentSlide("OWASP Top 10:2025 – the most critical web risks", S4);
    const top = [
      ["A01", "Broken Access Control", "Lab 3"],
      ["A02", "Security Misconfiguration", "Labs 4, 5"],
      ["A03", "Software Supply Chain Failures", "Lab 5"],
      ["A04", "Cryptographic Failures", "Labs 3, 5"],
      ["A05", "Injection", "Lab 2"],
      ["A06", "Insecure Design", "Lab 1"],
      ["A07", "Authentication Failures", "Lab 3"],
      ["A08", "Software or Data Integrity Failures", "Lab 5"],
      ["A09", "Security Logging & Alerting Failures", "Lab 4"],
      ["A10", "Mishandling of Exceptional Conditions", "Lab 2"],
    ];
    for (let i = 0; i < top.length; i++) {
      const col = Math.floor(i / 5), row = i % 5;
      const x = 0.6 + col * 6.2, y = 1.55 + row * 0.98;
      const [id, name, lab] = top[i];
      card(s, x, y, 5.9, 0.82, id + " row");
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: x + 0.15, y: y + 0.14, w: 0.85, h: 0.54, rectRadius: 0.1, fill: { color: i < 3 ? C.accent3 : C.text2 }, line: { type: "none" }, objectName: id + " badge" });
      tb(s, id, { x: x + 0.15, y: y + 0.14, w: 0.85, h: 0.54, fontSize: 15, bold: true, color: C.background1, align: "center", valign: "middle" });
      tb(s, name, { x: x + 1.2, y, w: 3.4, h: 0.82, fontSize: 16, bold: true, color: C.text2, valign: "middle" });
      tb(s, lab, { x: x + 4.6, y, w: 1.15, h: 0.82, fontSize: 13, color: C.accent2, bold: true, align: "right", valign: "middle" });
    }
    tb(s, "owasp.org/Top10  ·  right column: the lab where we practice it", { x: 0.6, y: 6.5, w: 12.1, h: 0.3, fontSize: 11, italic: true, color: C.accent5 });
    s.addNotes("OWASP is a non-profit foundation. The Top 10 is an awareness document built from real data from many organizations. Broken access control has been #1 since 2021. Check owasp.org/Top10 for the current edition before class.");
  }

  // 13. CVE & CVSS
  {
    const s = contentSlide("CVE and CVSS: naming and scoring vulnerabilities", S4);
    card(s, 0.6, 1.6, 5.6, 2.4, "CVE card");
    tb(s, "CVE = a unique ID", { x: 0.9, y: 1.8, w: 5.0, h: 0.5, fontSize: 20, bold: true, color: C.text2 });
    tb(s, [
      { text: "CVE-", options: { color: C.text2 } },
      { text: "2021", options: { color: C.accent2 } },
      { text: "-", options: { color: C.text2 } },
      { text: "44228", options: { color: C.accent3 } },
    ], { x: 0.9, y: 2.35, w: 5.0, h: 0.7, fontSize: 30, bold: true, fontFace: "Courier New" });
    tb(s, [
      { text: "year", options: { color: C.accent2, bold: true } },
      { text: " + ", options: { color: C.accent5 } },
      { text: "number", options: { color: C.accent3, bold: true } },
      { text: "  ·  this one is Log4Shell (Apache Log4j)", options: { color: C.accent5 } },
    ], { x: 0.9, y: 3.15, w: 5.0, h: 0.6, fontSize: 15 });

    card(s, 6.6, 1.6, 6.1, 2.4, "CVSS card");
    tb(s, "CVSS = severity score 0–10", { x: 6.9, y: 1.8, w: 5.5, h: 0.5, fontSize: 20, bold: true, color: C.text2 });
    tb(s, "Based on: how it is reached (network or local), how hard it is, privileges needed, user interaction, and the impact on C, I and A.", { x: 6.9, y: 2.35, w: 5.5, h: 1.4, fontSize: 15, color: C.text1 });

    const sev = [["None", "0.0", "A3AEBB"], ["Low", "0.1–3.9", "1B998B"], ["Medium", "4.0–6.9", "F2A541"], ["High", "7.0–8.9", "E07A3F"], ["Critical", "9.0–10", "D64550"]];
    const bw = 12.1 / sev.length;
    for (let i = 0; i < sev.length; i++) {
      const x = 0.6 + i * bw;
      s.addShape(pres.shapes.RECTANGLE, { x, y: 4.45, w: bw, h: 1.0, fill: { color: sev[i][2] }, line: { color: H.lt1, width: 2 }, objectName: sev[i][0] + " severity" });
      tb(s, sev[i][0], { x, y: 4.5, w: bw, h: 0.5, fontSize: 18, bold: true, color: C.background1, align: "center", valign: "middle" });
      tb(s, sev[i][1], { x, y: 4.95, w: bw, h: 0.4, fontSize: 14, color: C.background1, align: "center", valign: "middle" });
    }
    tb(s, [
      { text: "In Lab 5 ", options: { bold: true, color: C.text2 } },
      { text: "npm audit runs in your GitHub Actions pipeline and lists known CVEs in the app's dependencies. You decide what to fix first: start with Critical and High.", options: { color: C.text1 } },
    ], { x: 0.6, y: 5.8, w: 12.1, h: 0.9, fontSize: 16 });
    s.addNotes("Show nvd.nist.gov and search CVE-2021-44228 live. Log4Shell scored 10.0: remote, no authentication, full compromise. Explain that a high CVSS score does not always mean high risk for us: it depends on whether the vulnerable component is actually exposed.");
  }

  // ================= COURSE =================
  const S5 = "How this course works";
  pres.addSection({ title: S5 });

  // 15. Section divider
  {
    await dividerSlide(S5, "06", "How this course works", "5 labs  ·  1 project  ·  1 rule: only attack what you own", fa.FaFlask, "Lab flask");
  }

  // 16. The project
  {
    const s = contentSlide("One app, all semester: secure the Notes app", S5);
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 3.75, y: 1.55, w: 8.95, h: 2.95, rectRadius: 0.15, fill: { color: H.lt1, transparency: 100 }, line: { color: H.accent4, width: 1.5, dashType: "dash" }, objectName: "Linux server boundary" });
    tb(s, "Your DigitalOcean Linux server  ·  everything runs in Docker", { x: 4.0, y: 4.05, w: 8.45, h: 0.35, fontSize: 13, bold: true, color: C.accent4 });
    const nodes = [
      [fa.FaReact, "React", "Frontend, runs in the browser", 0.6, 2.5],
      [fa.FaServer, "Nginx", "Web server, HTTPS", 4.0, 2.45],
      [fa.FaCubes, "NestJS", "API / backend", 7.0, 2.45],
      [fa.FaDatabase, "PostgreSQL", "Database", 10.0, 2.45],
    ];
    for (let i = 0; i < nodes.length; i++) {
      const [ic, head, sub, x, w] = nodes[i];
      card(s, x, 1.8, w, 2.1, head + " node");
      await iconCircle(s, ic, x + (w - 0.8) / 2, 1.95, 0.8, C.text2, H.accent1, head);
      tb(s, head, { x: x + 0.1, y: 2.85, w: w - 0.2, h: 0.45, fontSize: 18, bold: true, color: C.text2, align: "center", valign: "middle" });
      tb(s, sub, { x: x + 0.1, y: 3.3, w: w - 0.2, h: 0.5, fontSize: 14, color: C.text1, align: "center" });
      if (i < nodes.length - 1) {
        const x1 = x + w + 0.05, x2 = nodes[i + 1][3] - 0.05;
        s.addShape(pres.shapes.LINE, { x: x1, y: 2.85, w: x2 - x1, h: 0, line: { color: H.accent5, width: 2, endArrowType: "triangle" }, objectName: "Flow arrow " + (i + 1) });
      }
    }
    tb(s, "What you will secure:", { x: 0.6, y: 4.75, w: 12.1, h: 0.4, fontSize: 16, bold: true, color: C.text2 });
    const feats = ["Database queries", "Login & passwords", "Access to notes", "Linux server", "Nginx & HTTPS"];
    const fw = (12.1 - 4 * 0.25) / 5;
    for (let i = 0; i < feats.length; i++) {
      const x = 0.6 + i * (fw + 0.25);
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y: 5.25, w: fw, h: 0.6, rectRadius: 0.3, fill: { color: C.accent2 }, line: { type: "none" }, objectName: feats[i] + " chip" });
      tb(s, feats[i], { x, y: 5.25, w: fw, h: 0.6, fontSize: 15, bold: true, color: C.background1, align: "center", valign: "middle" });
    }
    tb(s, [
      { text: "You get a working app with hidden vulnerabilities. ", options: { bold: true, color: C.text2 } },
      { text: "You find them, fix them, then deploy the app securely with GitHub Actions.", options: { color: C.text1 } },
    ], { x: 0.6, y: 6.15, w: 12.1, h: 0.6, fontSize: 17 });
    s.addNotes("This is the thread of the whole course. Students do not have to build the app: they receive it ready, with vulnerabilities hidden on purpose. Labs 1 to 3 secure the app, labs 4 and 5 secure the server and the web server. At the end each student has the app live on their own DigitalOcean server, deployed automatically by GitHub Actions. Vulnerable versions run only on the laptop, never on the public server.");
  }

  // 17. Course map
  {
    const s = contentSlide("5 labs, one project", S5);
    const labs = [
      ["Run the app & find the risks", "Notes app (NestJS + React + PostgreSQL) in Docker  ·  threat model  ·  secrets in .env  ·  database not exposed", "APP"],
      ["Fix injection: SQL injection & XSS", "Find both in the code  ·  parameterized queries  ·  input validation  ·  safe rendering in React", "APP"],
      ["Secure login & access", "Password hashing (bcrypt)  ·  JWT  ·  only your own notes (IDOR)  ·  login rate limiting", "APP"],
      ["Secure the Linux server", "DigitalOcean server  ·  SSH keys only  ·  ufw firewall  ·  Fail2ban  ·  automatic updates", "SERVER"],
      ["Secure the web server & deploy", "Nginx  ·  HTTPS (Let's Encrypt)  ·  security headers  ·  rate limiting  ·  GitHub Actions + npm audit", "SERVER"],
    ];
    for (let i = 0; i < labs.length; i++) {
      const y = 1.5 + i * 1.03;
      const [title, topics, part] = labs[i];
      const col = part === "APP" ? C.accent2 : C.accent4;
      card(s, 0.6, y, 12.1, 0.9, `Lab ${i + 1} row`);
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.75, y: y + 0.17, w: 1.05, h: 0.56, rectRadius: 0.1, fill: { color: col }, line: { type: "none" }, objectName: `Lab ${i + 1} badge` });
      tb(s, `Lab ${i + 1}`, { x: 0.75, y: y + 0.17, w: 1.05, h: 0.56, fontSize: 16, bold: true, color: C.background1, align: "center", valign: "middle" });
      tb(s, title, { x: 2.05, y: y + 0.08, w: 8.9, h: 0.42, fontSize: 18, bold: true, color: C.text2, valign: "middle" });
      tb(s, topics, { x: 2.05, y: y + 0.48, w: 10.4, h: 0.36, fontSize: 14, color: C.accent5, valign: "middle" });
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 11.3, y: y + 0.12, w: 1.2, h: 0.34, rectRadius: 0.17, fill: { color: col }, line: { type: "none" }, objectName: `Lab ${i + 1} part tag` });
      tb(s, part, { x: 11.3, y: y + 0.12, w: 1.2, h: 0.34, fontSize: 11, bold: true, color: C.background1, align: "center", valign: "middle" });
    }
    s.addNotes("Labs 1 to 3 secure the application code (APP). Labs 4 and 5 secure the Linux server and the web server, and deploy the app with GitHub Actions (SERVER). Each lab is one or two sessions, depending on the group.");
  }

  // 18. Lab workflow
  {
    const s = contentSlide("Every lab follows the same cycle", S5);
    const steps = [
      [fa.FaTools, "Run", "Start the app or the server"],
      [fa.FaSearch, "Find", "Discover the weakness"],
      [fa.FaBug, "Exploit", "Show the impact, on localhost only"],
      [fa.FaShieldAlt, "Fix", "Change the code or the server config"],
      [fa.FaClipboardCheck, "Verify & report", "Prove the fix works, before/after evidence"],
    ];
    const w = 2.2, gap = (12.1 - 5 * w) / 4;
    for (let i = 0; i < steps.length; i++) {
      const x = 0.6 + i * (w + gap);
      const [ic, head, txt] = steps[i];
      const fix = i >= 3;
      card(s, x, 1.7, w, 3.3, head + " step", fix ? C.text2 : C.background2);
      await iconCircle(s, ic, x + (w - 0.9) / 2, 1.95, 0.9, fix ? C.accent1 : C.text2, fix ? H.dk2 : H.accent1, head);
      tb(s, `${i + 1}. ${head}`, { x: x + 0.1, y: 3.05, w: w - 0.2, h: 0.5, fontSize: 18, bold: true, color: fix ? C.accent1 : C.text2, align: "center", valign: "middle" });
      tb(s, txt, { x: x + 0.2, y: 3.6, w: w - 0.4, h: 1.2, fontSize: 14, color: fix ? C.background1 : C.text1, align: "center" });
      if (i < steps.length - 1) {
        s.addShape(pres.shapes.LINE, { x: x + w + 0.04, y: 2.4, w: gap - 0.08, h: 0, line: { color: H.accent5, width: 2, endArrowType: "triangle" }, objectName: "Step arrow " + (i + 1) });
      }
    }
    tb(s, [
      { text: "Most of your grade comes from steps 4 and 5. ", options: { bold: true, color: C.text2 } },
      { text: "Breaking something is easy. Explaining why it broke and fixing it properly is the real skill, and it is what employers pay for.", options: { color: C.text1 } },
    ], { x: 0.6, y: 5.45, w: 12.1, h: 1.0, fontSize: 17 });
    s.addNotes("Emphasize the defensive goal of the course. Students who only run tools without explaining or fixing anything get a low grade.");
  }

  // 19. Ethics
  {
    const s = contentSlide("Ethics and law: the most important slide", S5);
    const cols = [
      ["Allowed", fa.FaCheckCircle, C.accent2, H.accent2, [
        "Your own laptop and DigitalOcean server",
        "Intentionally vulnerable practice apps",
        "Systems where you have written permission",
        "CTF platforms, within their rules",
        "Reporting bugs through responsible disclosure",
      ]],
      ["Not allowed", fa.FaTimesCircle, C.accent3, H.accent3, [
        "Scanning the university or any public network",
        "Testing a real website \"just to see\"",
        "Keeping or sharing captured passwords or data",
        "Using lab techniques outside the lab",
        "Running tools you do not understand",
      ]],
    ];
    for (let c = 0; c < 2; c++) {
      const x = 0.6 + c * 6.2;
      const [head, ic, col, hex, items] = cols[c];
      card(s, x, 1.55, 5.9, 3.6, head + " card");
      s.addImage({ data: await icon(ic, hex), x: x + 0.3, y: 1.8, w: 0.6, h: 0.6, objectName: head + " icon", altText: head });
      tb(s, head, { x: x + 1.1, y: 1.8, w: 4.5, h: 0.6, fontSize: 24, bold: true, color: col, valign: "middle" });
      tb(s, items.map((t, i) => ({ text: t, options: { bullet: true, breakLine: i < items.length - 1 } })),
        { x: x + 0.35, y: 2.65, w: 5.3, h: 2.9, fontSize: 16, color: C.text1, paraSpaceAfter: 8 });
    }
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.6, y: 5.5, w: 12.1, h: 0.85, rectRadius: 0.12, fill: { color: C.text2 }, line: { type: "none" }, objectName: "Law box" });
    tb(s, "Unauthorized access to computer systems is a criminal offense in Moldova, in the EU (Directive 2013/40/EU) and almost everywhere else. Good intentions are not a defense.",
      { x: 0.9, y: 5.5, w: 11.5, h: 0.85, fontSize: 16, color: C.background1, valign: "middle" });
    s.addNotes("Have students sign the lab code-of-conduct agreement today. Explain the difference between a penetration tester and a criminal: written authorization and scope. The tools are the same.");
  }

  // 20. Grading
  {
    const s = contentSlide("Grading", S5);
    const g = [
      ["60%", "Lab reports", "Findings, impact, fix, before/after evidence", C.accent2],
      ["20%", "Final assessment", "Present your secured app, live on your own server", C.accent1],
      ["20%", "Quiz / oral defense", "Explain the concepts behind one of your labs", C.accent4],
    ];
    for (let i = 0; i < 3; i++) {
      const x = 0.6 + i * 4.15;
      card(s, x, 1.7, 3.8, 3.3, g[i][1] + " card");
      tb(s, g[i][0], { x: x + 0.35, y: 1.95, w: 3.1, h: 1.3, fontSize: 66, bold: true, color: g[i][3], fontFace: THEME.headFontFace, valign: "middle" });
      tb(s, g[i][1], { x: x + 0.35, y: 3.35, w: 3.1, h: 0.5, fontSize: 22, bold: true, color: C.text2 });
      tb(s, g[i][2], { x: x + 0.35, y: 3.95, w: 3.1, h: 1.1, fontSize: 15, color: C.text1 });
    }
    tb(s, [
      { text: "Report deadline: ", options: { bold: true, color: C.text2 } },
      { text: "before the next lab. Late reports and copied reports are penalized. Use the report template on the course page.", options: { color: C.text1 } },
    ], { x: 0.6, y: 5.4, w: 12.1, h: 0.8, fontSize: 16 });
    s.addNotes("Adjust the percentages and deadlines to the university regulations.");
  }

  // 21. Homework
  {
    const s = contentSlide("Before Lab 1", S5);
    const hw = [
      [fa.FaDownload, "Prepare your machine", "Linux laptop or VM with Docker, Node.js 22, Git and VS Code. Create GitHub and DigitalOcean accounts."],
      [fa.FaBook, "Read", "OWASP Top 10 overview (owasp.org/Top10). Pick one item and be ready to explain it in 2 minutes."],
      [fa.FaPencilAlt, "Think", "For a notes app with login and a database: list 5 assets, 3 threats and 1 thing to protect first."],
    ];
    for (let i = 0; i < hw.length; i++) {
      const y = 1.65 + i * 1.6;
      const [ic, head, txt] = hw[i];
      card(s, 0.6, y, 12.1, 1.35, head + " task");
      tb(s, String(i + 1), { x: 0.85, y, w: 0.7, h: 1.35, fontSize: 48, bold: true, color: C.accent1, fontFace: THEME.headFontFace, valign: "middle" });
      await iconCircle(s, ic, 1.7, y + 0.25, 0.85, C.text2, H.accent1, head);
      tb(s, head, { x: 2.9, y: y + 0.15, w: 9.5, h: 0.5, fontSize: 20, bold: true, color: C.text2 });
      tb(s, txt, { x: 2.9, y: y + 0.65, w: 9.5, h: 0.6, fontSize: 16, color: C.text1 });
    }
    s.addNotes("Lab 1 starts by running the app in Docker, so Docker must already work: docker run hello-world should succeed. Offer a short help session for students who get stuck.");
  }

  // 22. Closing
  {
    const s = titleSlide(S5, "\u201CSecurity is a process, not a product.\u201D", "Bruce Schneier  ·  Questions?", await icon(fa.FaQuestionCircle, H.accent1), "Question");
    s.addNotes("Close with the idea that security is continuous: patching, monitoring and improving, which is exactly what the labs will practice.");
  }

  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  console.log("written", OUT);
})().catch((e) => { console.error(e); process.exit(1); });
