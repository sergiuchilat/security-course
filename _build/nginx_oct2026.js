const pptxgen = require("pptxgenjs");
const React = require("react");
const { renderToStaticMarkup } = require("react-dom/server");
const sharp = require("sharp");
const fa = require("react-icons/fa");
const { applyTheme } = require("/Users/sergiuchilat/.claude/skills/synced/e25047ed-7c97-4849-9da1-c59bbd817501_b100ab20-30b2-4fd4-bf29-fbdff3299f26/pptx/scripts/apply_theme.js");

const OUT = "../Talk_NGINX_October_2026.pptx";

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
pres.title = "NGINX critical vulnerabilities – Case Study (October 2026)";
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
const FOOTER = "Web Application & Infrastructure Security  ·  Case study: NGINX, October 2026";

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

// Speaker notes în română, în ordinea slide-urilor; nota în engleză rămâne dedesubt.
const NOTES_RO = [
  // 1
  "Durată: 10–15 minute. Pornim de la alerta oficială AL-2026-150 a Agenției de Securitate Cibernetică din Singapore (CSA), din 21 octombrie 2026. Întrebare de deschidere: „Cine dintre voi a instalat deja NGINX?” Toți: în Lab 5 e serverul nostru web, iar aplicația Notițe rulează frontend-ul pe NGINX.",
  // 2
  "Pe scurt: trei bug-uri de memorie în NGINX, toate cu scor CVSS v4.0 9.2 (critic) dat de F5. Atacatorul nu are nevoie de cont: trimite cereri HTTP speciale. Rezultatul sigur e căderea procesului worker (DoS); execuția de cod e posibilă doar dacă ASLR e dezactivat sau ocolit.\n\nCondiția cheie: bug-urile se declanșează doar cu anumite configurații, care NU sunt cele implicite. De asta același bug are 9.2 la F5 și „medium” pe nginx.org. E o diferență bună de discutat: scorul descrie cel mai rău caz, nu serverul vostru.",
  // 3
  "NGINX e peste tot: e unul dintre cele mai folosite servere web din lume, și mai ales stă în fața altor aplicații ca reverse proxy, load balancer, terminație TLS. În Kubernetes, NGINX Ingress Controller e „ușa de intrare” în cluster.\n\nAfectate: NGINX Open Source (până la 1.30.5 și 1.31.6), NGINX Plus R33–R37, Ingress Controller, Gateway Fabric, Instance Manager, F5 WAF și F5 DoS for NGINX. Ironic: chiar produsele de protecție (WAF, DoS) conțin bug-ul, pentru că sunt construite peste NGINX.\n\nLegătura cu cursul: în Lab 5 instalăm NGINX pe Ubuntu cu apt și îl folosim ca reverse proxy pentru aplicația Notițe.",
  // 4
  "Cum lucrează NGINX: un proces master (rulează ca root, citește configurația, deschide porturile 80/443) și mai mulți workeri (rulează ca utilizatorul www-data/nginx) care servesc efectiv cererile. Dacă un worker cade, master-ul pornește imediat altul.\n\nDe ce contează: un singur crash trece aproape neobservat, dar un atacator care trimite cereri în buclă ține workerii într-un ciclu de căderi = DoS. Iar dacă reușește corupția controlată a memoriei, rulează cod în worker, care vede tot traficul deja decriptat (parole, cookie-uri, token-uri).\n\nASLR (Address Space Layout Randomization) = sistemul de operare pune memoria procesului la adrese aleatorii la fiecare pornire. Atacatorul nu știe unde să sară, deci execuția de cod devine foarte grea. E activ implicit pe Linux. E un exemplu perfect de apărare în adâncime: bug-ul există, dar un strat de protecție îl face mult mai greu de exploatat.",
  // 5
  "Cele trei bug-uri și configurația care le face posibile:\n\n1) CVE-2026-15176: depășire de buffer în heap când un location combină alias cu rewrite și folosește variabila $request_filename (în set, return sau în formatul de log). Găsit intern de F5.\n\n2) CVE-2026-97205: use-after-free când sunt activate AIO (citire asincronă de fișiere) și SSI (Server Side Includes, pagini cu <!--# include virtual -->). Raportat de cercetătorul Cyberlanc3r.\n\n3) CVE-2026-93479: depășire de buffer la directiva map cu expresii regulate, când variabilele de captură ($1, $2) sunt folosite în root sau în log înaintea variabilei produse de map. Raportat de Matt Suiche (Tolmo) și Mateusz Gierblinski.\n\nExemplele de pe slide sunt simplificate, scop didactic. Ideea: combinații de directive rare, dar perfect valide, pe care cineva le-a scris cândva și nu le-a mai revizuit.",
  // 6
  "Acesta e mesajul central al prezentării: „versiune vulnerabilă” nu înseamnă automat „server exploatabil”.\n\nPentru ca atacul să meargă trebuie îndeplinite toate trei: versiunea e afectată, configurația conține combinația periculoasă, iar pentru execuție de cod ASLR trebuie ocolit. Fără configurația respectivă, atacatorul nu are ce declanșa.\n\nDe aici diferența de scor: F5 dă 9.2 (CVSS v4.0, cel mai rău caz), CVSS v3.1 dă 8.1, iar echipa nginx.org le-a marcat „medium”, pentru că majoritatea instalărilor nu folosesc acele combinații.\n\nAtenție însă: asta NU e un motiv să nu actualizezi. Configurațiile se schimbă, cineva adaugă mâine un alias, iar analiza „noi nu suntem afectați” poate fi greșită. Actualizarea rămâne regula.",
  // 7
  "Ce obține atacatorul, în ordinea probabilității:\n1) DoS: workerii cad în buclă, site-ul devine lent sau indisponibil. Asta funcționează garantat dacă există configurația vulnerabilă.\n2) Execuție de cod: doar cu ASLR ocolit, de obicei combinând bug-ul cu o scurgere de informație despre adresele din memorie.\n3) Dacă reușește: NGINX vede tot traficul decriptat, are acces la cheile TLS încărcate în memorie și la tot ce se află în spatele lui în rețeaua internă.\n\nCIA: disponibilitatea e sigur afectată; confidențialitatea și integritatea doar în scenariul cu execuție de cod.",
  // 8
  "Cronologie:\n- 14 octombrie: F5 publică trei advisory-uri (K000162416, K000162607, K000162608) și nginx.org lansează 1.30.6 (stable) și 1.31.7 (mainline). Aceleași versiuni repară încă două probleme: injecție de corp de cerere în HTTP/3 (CVE-2026-95101) și falsificarea sesiunilor TLS (CVE-2026-97629).\n- 19 octombrie: presa de specialitate (BleepingComputer) preia subiectul.\n- 21 octombrie: CSA Singapore publică alerta și cere actualizare imediată.\n- Până pe 20–22 octombrie nu a fost raportată exploatare reală.\n\nComparație cu PHP2Shell din aceeași săptămână: aici dezvăluirea a fost coordonată, fără scurgeri, iar patch-ul a apărut odată cu anunțul. Așa ar trebui să arate procesul.\n\nContext pentru întrebări: F5 a fost spartă în 2025 de un actor statal care a furat cod sursă; de aceea orice bug F5 primește multă atenție.",
  // 9
  "Versiunile reparate: NGINX Open Source 1.30.6 sau 1.31.7; NGINX Plus 37.0.5.1 sau 37.1.0.1; Ingress Controller 5.7.0; Gateway Fabric 2.8.0; Instance Manager 2.24.0; F5 WAF for NGINX 5.16.0 sau 5.14.2. Unele ramuri vechi nu primesc patch: trebuie mutat pe o ramură suportată.\n\nDacă nu poți actualiza imediat, schimbi configurația: elimini alias din location sau $request_filename; dezactivezi AIO sau SSI (atenție, fără AIO pot scădea performanțele pentru fișiere mari); folosești capturi cu nume în map și doar în blocul cu expresia regulată.\n\nDetaliu practic: pe Ubuntu/Debian, pachetul din apt poate avea un număr de versiune mai mic (de exemplu 1.24) dar să conțină corecția prin backport. Verificați anunțul distribuției (USN pentru Ubuntu), nu doar numărul versiunii.",
  // 10
  "Exercițiu practic, 2–3 minute. Cum verifici un server:\n- nginx -v → versiunea;\n- nginx -T → toată configurația activă, inclusiv fișierele incluse; o filtrezi cu grep după alias, request_filename, aio, ssi, map;\n- nginx -t înainte de reload, ca să nu strici serverul cu o configurație greșită.\n\nÎntrebare pentru studenți: aplicația noastră Notițe folosește imaginea nginx:1.27-alpine. E vulnerabilă?\nRăspuns: versiunea e în intervalul afectat, dar configurația (labs/notes-app/frontend/nginx.conf) are doar root, try_files și proxy_pass: niciun alias, SSI, AIO sau map. Deci cele trei bug-uri nu pot fi declanșate. Totuși, versiunea 1.27 e veche și are alte bug-uri reparate de atunci, așa că într-un proiect real am actualiza imaginea.",
  // 11
  "Două incidente în aceeași săptămână, două lecții diferite.\nPHP2Shell: configurație implicită, deci aproape toată lumea afectată; dezvăluirea s-a scurs, atacurile au început înaintea patch-ului.\nNGINX: configurații rare, deci puțini afectați realmente; dezvăluire coordonată, patch disponibil din prima zi, fără exploatare cunoscută.\n\nCe e comun: amândouă sunt bug-uri de memorie în cod C (use-after-free, buffer overflow), în componente de infrastructură pe care aplicația noastră se bazează fără să le vadă. Lecțiile: inventar, actualizare rapidă, configurație minimă (activezi doar ce folosești), straturi de apărare (ASLR, privilegii minime, izolare).",
  // 12
  "Încheiere și discuție: „Dacă ai 50 de servere NGINX, cum afli în 30 de minute care au configurația vulnerabilă?” Răspunsuri bune: nginx -T pe fiecare server prin SSH sau Ansible, căutare în repository-ul cu configurații (dacă e ținut în Git), scanner de vulnerabilități, inventar de imagini Docker.\n\nSurse: alerta CSA AL-2026-150; advisory-urile F5 K000162416, K000162607, K000162608; nginx.org/en/security_advisories.html; BleepingComputer, Cybersecurity News, eSecurity Planet.",
];
let noteIdx = 0;
function notes(slide, en) {
  slide.addNotes(NOTES_RO[noteIdx++] + "\n\n— English —\n" + en);
}

function tb(slide, text, opts) {
  slide.addText(text, Object.assign({ isTextBox: true, fontFace: THEME.bodyFontFace, margin: 0, valign: "top" }, opts));
}

function card(slide, x, y, w, h, name, fill = C.background2) {
  slide.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, rectRadius: 0.12, fill: { color: fill }, line: { type: "none" }, objectName: name });
}



(async () => {
  const S1 = "What happened";
  pres.addSection({ title: S1 });

  // 1. Title
  {
    const s = titleSlide(S1, "NGINX: three critical bugs, one dangerous config away", "CSA alert AL-2026-150  ·  CVE-2026-15176, CVE-2026-97205, CVE-2026-93479  ·  October 2026", await icon(fa.FaServer, H.accent1), "Server");
    notes(s, "10-15 minute talk based on Singapore CSA alert AL-2026-150 (21 Oct 2026). Opening question: who has installed NGINX? Everyone: it is our web server in Lab 5, and the Notes app frontend runs on it.");
  }

  // 2. At a glance
  {
    const s = contentSlide("The alert in one slide", S1);
    const stats = [
      ["3", "memory-corruption bugs in NGINX", C.accent3],
      ["9.2", "CVSS v4.0 by F5 (8.1 in v3.1, “medium” on nginx.org)", C.accent1],
      ["0", "credentials needed: remote, unauthenticated HTTP requests", C.accent4],
      ["7", "F5 product families affected, from Open Source to WAF", C.accent2],
    ];
    for (let i = 0; i < 4; i++) {
      const x = 0.6 + i * 3.1;
      card(s, x, 1.6, 2.8, 2.9, "Stat card " + (i + 1));
      tb(s, stats[i][0], { x: x + 0.25, y: 1.85, w: 2.3, h: 1.0, fontSize: 48, bold: true, color: stats[i][2], fontFace: THEME.headFontFace, valign: "middle" });
      tb(s, stats[i][1], { x: x + 0.25, y: 3.0, w: 2.3, h: 1.3, fontSize: 16, color: C.text1 });
    }
    card(s, 0.6, 4.85, 12.1, 1.75, "Summary card");
    await iconCircle(s, fa.FaCogs, 0.95, 5.25, 0.95, C.accent3, "FFFFFF", "Config");
    tb(s, [
      { text: "The catch: only non-default configurations are vulnerable. ", options: { bold: true, color: C.text2 } },
      { text: "A crafted request crashes the worker process (DoS). Remote code execution is possible only if ASLR is disabled or bypassed.", options: { color: C.text1 } },
    ], { x: 2.2, y: 4.85, w: 10.2, h: 1.75, fontSize: 17, valign: "middle" });
    notes(s, "Same bugs, three scores: 9.2 (F5, CVSS v4), 8.1 (CVSS v3.1), medium (nginx.org). A CVSS score describes the worst case, not your server.");
  }

  // 3. Who is affected
  {
    const s = contentSlide("Who was affected", S1);
    const groups = [
      [fa.FaGlobe, "Websites and APIs", "NGINX Open Source up to 1.30.5 (stable) and 1.31.6 (mainline): web servers, reverse proxies, TLS termination."],
      [fa.FaBuilding, "Companies on NGINX Plus", "NGINX Plus R33 to R37, the commercial edition used by large enterprises."],
      [fa.FaCubes, "Kubernetes clusters", "Ingress Controller 5.6.0–5.6.3, Gateway Fabric 1.5–1.6 and 2.1–2.7: the front door of the cluster."],
      [fa.FaShieldAlt, "Security products", "F5 WAF for NGINX 5.9–5.15 and F5 DoS for NGINX 4.8–4.10. The protection layer itself was vulnerable."],
    ];
    for (let i = 0; i < 4; i++) {
      const x = 0.6 + (i % 2) * 6.15, y = 1.6 + Math.floor(i / 2) * 2.0;
      const [ic, head, txt] = groups[i];
      card(s, x, y, 5.95, 1.75, head + " card");
      await iconCircle(s, ic, x + 0.3, y + 0.35, 0.95, C.text2, H.accent1, head);
      tb(s, head, { x: x + 1.5, y: y + 0.2, w: 4.25, h: 0.45, fontSize: 20, bold: true, color: C.text2 });
      tb(s, txt, { x: x + 1.5, y: y + 0.65, w: 4.25, h: 1.0, fontSize: 14, color: C.text1 });
    }
    tb(s, [
      { text: "Closer to home: ", options: { bold: true, color: C.accent3 } },
      { text: "Lab 5 puts NGINX in front of the Notes app, and the app's frontend container runs nginx:1.27-alpine.", options: { color: C.text1 } },
    ], { x: 0.6, y: 5.75, w: 12.1, h: 0.8, fontSize: 18, valign: "middle" });
    notes(s, "NGINX is one of the most used web servers and sits in front of other apps as reverse proxy, load balancer and TLS terminator. Also listed: NGINX Instance Manager 2.18-2.23.");
  }

  // ================= HOW IT WORKS =================
  const S2 = "How it works";
  pres.addSection({ title: S2 });

  // 4. Master / workers
  {
    const s = contentSlide("Background: master, workers and ASLR", S2);
    // master
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.6, y: 2.45, w: 2.6, h: 1.5, rectRadius: 0.12, fill: { color: C.text2 }, line: { type: "none" }, objectName: "Master box" });
    tb(s, "Master", { x: 0.75, y: 2.6, w: 2.3, h: 0.5, fontSize: 20, bold: true, color: C.background1, align: "center" });
    tb(s, "root · reads config · opens :80 :443", { x: 0.75, y: 3.1, w: 2.3, h: 0.75, fontSize: 13, color: C.background1, align: "center" });
    const wk = [["Worker 1", "serving requests", C.text2], ["Worker 2", "crashed → restarted", C.accent3], ["Worker 3", "serving requests", C.text2]];
    for (let i = 0; i < 3; i++) {
      const y = 1.55 + i * 1.1;
      const ty = y + 0.45, up = ty < 3.2;
      s.addShape(pres.shapes.LINE, { x: 3.25, y: up ? ty : 3.2, w: 0.9, h: Math.max(Math.abs(ty - 3.2), 0.01), flipV: up, line: { color: H.accent5, width: 1.5, endArrowType: "triangle" }, objectName: "Spawn arrow " + (i + 1) });
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 4.2, y, w: 2.6, h: 0.9, rectRadius: 0.12, fill: { color: wk[i][2] }, line: { type: "none" }, objectName: wk[i][0] });
      tb(s, wk[i][0], { x: 4.3, y: y + 0.08, w: 2.4, h: 0.4, fontSize: 16, bold: true, color: C.background1, align: "center" });
      tb(s, wk[i][1], { x: 4.3, y: y + 0.48, w: 2.4, h: 0.35, fontSize: 13, color: C.background1, align: "center" });
    }
    tb(s, "www-data · sees decrypted traffic", { x: 4.2, y: 4.85, w: 2.6, h: 0.4, fontSize: 12, italic: true, color: C.accent5, align: "center" });
    const pts = [
      [fa.FaRedo, "Crash = restart", "One crash is barely visible. A loop of crafted requests keeps workers crashing: denial of service."],
      [fa.FaRandom, "ASLR", "Linux loads code and heap at random addresses on every start. The attacker does not know where to jump, so turning memory corruption into code execution is much harder."],
    ];
    for (let i = 0; i < 2; i++) {
      const y = 1.55 + i * 2.0;
      const [ic, head, txt] = pts[i];
      card(s, 7.3, y, 5.4, 1.8, head + " card");
      await iconCircle(s, ic, 7.55, y + 0.4, 0.9, C.text2, H.accent1, head);
      tb(s, head, { x: 8.7, y: y + 0.2, w: 3.8, h: 0.42, fontSize: 19, bold: true, color: C.text2 });
      tb(s, txt, { x: 8.7, y: y + 0.65, w: 3.8, h: 1.1, fontSize: 14, color: C.text1 });
    }
    tb(s, [
      { text: "Defense in depth in action: ", options: { bold: true, color: C.accent2 } },
      { text: "the bug is real, but an OS-level layer (ASLR) turns a likely RCE into a likely crash.", options: { color: C.text1 } },
    ], { x: 0.6, y: 5.75, w: 12.1, h: 0.8, fontSize: 18, valign: "middle" });
    notes(s, "Master runs as root and spawns workers as www-data/nginx. Workers handle requests, so a worker compromise sees decrypted TLS traffic. ASLR is enabled by default on Linux (check: cat /proc/sys/kernel/randomize_va_space = 2).");
  }

  // 5. The three bugs
  {
    const s = contentSlide("Three bugs, three dangerous configs", S2);
    const bugs = [
      ["CVE-2026-15176", "Heap overflow · alias + rewrite", "location /img/ {\n  alias /data/images/;\n  rewrite ^ /x;\n  set $f $request_filename;\n}", "Found internally by F5", C.accent3],
      ["CVE-2026-97205", "Use-after-free · AIO + SSI", "aio on;\nssi on;\n\n<!--# include\n   virtual=\"/menu\" -->", "Reported by Cyberlanc3r", C.accent1],
      ["CVE-2026-93479", "Heap overflow · map + regex", "map $uri $site {\n  ~^/(.+)/ $1;\n}\nroot /srv/$1/$site;", "Matt Suiche (Tolmo), M. Gierblinski", C.accent4],
    ];
    for (let i = 0; i < 3; i++) {
      const x = 0.6 + i * 4.15;
      const [cve, what, code, who, col] = bugs[i];
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y: 1.55, w: 3.8, h: 0.95, rectRadius: 0.12, fill: { color: col }, line: { type: "none" }, objectName: cve + " header" });
      tb(s, cve, { x: x + 0.2, y: 1.62, w: 3.4, h: 0.45, fontSize: 19, bold: true, color: C.background1 });
      tb(s, what, { x: x + 0.2, y: 2.05, w: 3.4, h: 0.4, fontSize: 14, color: C.background1 });
      card(s, x, 2.65, 3.8, 2.0, cve + " code", C.text2);
      tb(s, code, { x: x + 0.25, y: 2.85, w: 3.4, h: 1.65, fontSize: 13, color: C.background1, fontFace: "Courier New" });
      tb(s, who, { x, y: 4.8, w: 3.8, h: 0.4, fontSize: 13, italic: true, color: C.accent5 });
    }
    tb(s, [
      { text: "Common pattern: ", options: { bold: true, color: C.text2 } },
      { text: "rare but valid combinations of directives. Simplified examples for teaching.", options: { color: C.text1 } },
    ], { x: 0.6, y: 5.5, w: 12.1, h: 0.6, fontSize: 17, valign: "middle" });
    notes(s, "15176: location with alias + rewrite and $request_filename in set/return/log_format. 97205: AIO enabled and an SSI page with include virtual. 93479: regex map captures ($1) used in root or log_format before the map's own output variable. Snippets are simplified illustrations, not exact triggers.");
  }

  // 6. Vulnerable vs exploitable
  {
    const s = contentSlide("Vulnerable version ≠ exploitable server", S2);
    const conds = [
      [fa.FaCodeBranch, "Affected version", "e.g. 1.30.5 or 1.31.6", C.accent4],
      [fa.FaFileCode, "Dangerous config", "alias + rewrite, AIO + SSI, regex map", C.accent1],
      [fa.FaUnlock, "ASLR bypassed", "only needed for code execution", C.accent3],
    ];
    for (let i = 0; i < 3; i++) {
      const x = 0.6 + i * 4.15;
      const [ic, head, sub, col] = conds[i];
      card(s, x, 1.6, 3.6, 2.4, head + " card");
      await iconCircle(s, ic, x + 1.35, 1.8, 0.9, col, "FFFFFF", head);
      tb(s, head, { x: x + 0.2, y: 2.85, w: 3.2, h: 0.45, fontSize: 20, bold: true, color: C.text2, align: "center" });
      tb(s, sub, { x: x + 0.2, y: 3.35, w: 3.2, h: 0.5, fontSize: 14, color: C.accent5, align: "center" });
      if (i < 2) tb(s, "+", { x: x + 3.6, y: 2.4, w: 0.55, h: 0.8, fontSize: 40, bold: true, color: C.accent5, align: "center", valign: "middle" });
    }
    const sc = [["9.2", "F5, CVSS v4.0", C.accent3], ["8.1", "CVSS v3.1", C.accent1], ["medium", "nginx.org", C.accent2]];
    tb(s, "One bug, three ratings", { x: 0.6, y: 4.35, w: 5, h: 0.45, fontSize: 18, bold: true, color: C.text2 });
    for (let i = 0; i < 3; i++) {
      const x = 0.6 + i * 2.15;
      card(s, x, 4.9, 1.95, 1.5, "Score " + sc[i][1]);
      tb(s, sc[i][0], { x, y: 5.0, w: 1.95, h: 0.75, fontSize: i === 2 ? 26 : 34, bold: true, color: sc[i][2], align: "center", valign: "middle", fontFace: THEME.headFontFace });
      tb(s, sc[i][1], { x, y: 5.8, w: 1.95, h: 0.4, fontSize: 13, color: C.accent5, align: "center" });
    }
    card(s, 7.2, 4.35, 5.5, 2.05, "Warning card", C.text2);
    tb(s, [
      { text: "Still patch. ", options: { bold: true, color: C.accent1 } },
      { text: "Configs change: someone adds an alias tomorrow. “We are not affected” is an analysis that can be wrong; an update is not.", options: { color: C.background1 } },
    ], { x: 7.5, y: 4.35, w: 5.0, h: 2.05, fontSize: 17, valign: "middle" });
    notes(s, "Central message: all conditions must hold. A CVSS score is the worst case for the component; your risk depends on your configuration. But treat 'not affected' conclusions with caution and patch anyway.");
  }

  // 7. Impact
  {
    const s = contentSlide("What an attacker gets", S2);
    const imp = [
      [fa.FaPowerOff, "Denial of service", "Workers crash in a loop; the site slows down or goes offline. Reliable whenever the vulnerable config exists.", C.accent1, "Likely"],
      [fa.FaTerminal, "Code execution", "Needs ASLR disabled or bypassed, usually by chaining with a memory-leak bug.", C.accent3, "Hard"],
      [fa.FaUserSecret, "If RCE succeeds", "NGINX sees decrypted traffic and TLS keys, and reaches every backend behind it.", C.accent4, "High impact"],
    ];
    for (let i = 0; i < 3; i++) {
      const x = 0.6 + i * 4.15;
      const [ic, head, txt, col, tag] = imp[i];
      card(s, x, 1.6, 3.8, 3.9, head + " card");
      await iconCircle(s, ic, x + 0.35, 1.9, 0.95, col, "FFFFFF", head);
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: x + 1.6, y: 2.15, w: 1.85, h: 0.45, rectRadius: 0.1, fill: { color: C.background1 }, line: { type: "none" }, objectName: head + " tag" });
      tb(s, tag, { x: x + 1.6, y: 2.15, w: 1.85, h: 0.45, fontSize: 13, bold: true, color: col, align: "center", valign: "middle" });
      tb(s, head, { x: x + 0.35, y: 3.0, w: 3.1, h: 0.5, fontSize: 22, bold: true, color: C.text2, fontFace: THEME.headFontFace });
      tb(s, txt, { x: x + 0.35, y: 3.55, w: 3.15, h: 1.8, fontSize: 15, color: C.text1 });
    }
    tb(s, [
      { text: "CIA triad: ", options: { bold: true, color: C.text2 } },
      { text: "Availability is always at risk. Confidentiality and Integrity only in the code-execution scenario.", options: { color: C.text1 } },
    ], { x: 0.6, y: 5.85, w: 12.1, h: 0.7, fontSize: 18, valign: "middle" });
    notes(s, "Order by likelihood: DoS is reliable, RCE is hard, but a compromised reverse proxy is a very valuable foothold.");
  }

  // ================= RESPONSE =================
  const S3 = "Response";
  pres.addSection({ title: S3 });

  // 8. Timeline
  {
    const s = contentSlide("Timeline: a coordinated disclosure", S3);
    const ev = [
      ["Before Oct 14", "Researchers report privately to F5; one bug found by F5's own team", C.accent2],
      ["Oct 14", "F5 advisories K000162416 / 162607 / 162608. nginx 1.30.6 and 1.31.7 released, fixing these 3 + 2 more (HTTP/3, TLS tickets)", C.accent4],
      ["Oct 19", "Security press picks it up (BleepingComputer and others)", C.accent1],
      ["Oct 21", "Singapore CSA publishes alert AL-2026-150: patch immediately", C.accent3],
      ["As of Oct 22", "No exploitation in the wild reported, no public exploit known", C.accent2],
    ];
    const x0 = 2.9;
    s.addShape(pres.shapes.LINE, { x: x0 + 0.14, y: 1.75, w: 0, h: 4.3, line: { color: H.accent5, width: 2 }, objectName: "Timeline line" });
    for (let i = 0; i < ev.length; i++) {
      const y = 1.6 + i * 0.95;
      const [when, what, col] = ev[i];
      tb(s, when, { x: 0.6, y, w: 2.1, h: 0.6, fontSize: 16, bold: true, color: C.text2, align: "right", valign: "middle" });
      s.addShape(pres.shapes.OVAL, { x: x0, y: y + 0.16, w: 0.28, h: 0.28, fill: { color: col }, line: { color: H.lt1, width: 2 }, objectName: "Dot " + (i + 1) });
      tb(s, what, { x: x0 + 0.6, y, w: 9.2, h: 0.6, fontSize: 16, color: C.text1, valign: "middle" });
    }
    tb(s, [
      { text: "Compare with PHP2Shell the same week: ", options: { bold: true, color: C.text2 } },
      { text: "there the details leaked and attacks started before the patch. Here the patch came first.", options: { color: C.text1 } },
    ], { x: 0.6, y: 6.25, w: 12.1, h: 0.5, fontSize: 16, valign: "middle" });
    notes(s, "Advisories published Oct 14 (CVE records 14-15 Oct). The CSA alert came a week later. Context: F5 itself was breached by a state actor in 2025, so F5 vulnerabilities get close attention.");
  }

  // 9. The fix
  {
    const s = contentSlide("How it was fixed", S3);
    const vers = [["NGINX Open Source", "1.30.6 / 1.31.7"], ["NGINX Plus", "37.0.5.1 / 37.1.0.1"], ["Ingress Controller", "5.7.0"], ["Gateway Fabric", "2.8.0"], ["Instance Manager", "2.24.0"], ["F5 WAF for NGINX", "5.16.0 / 5.14.2"]];
    tb(s, "Patched releases", { x: 0.6, y: 1.5, w: 5.8, h: 0.45, fontSize: 18, bold: true, color: C.text2 });
    for (let i = 0; i < vers.length; i++) {
      const y = 2.05 + i * 0.72;
      card(s, 0.6, y, 5.8, 0.6, vers[i][0]);
      tb(s, vers[i][0], { x: 0.85, y, w: 2.7, h: 0.6, fontSize: 15, color: C.accent5, valign: "middle" });
      tb(s, vers[i][1], { x: 3.55, y, w: 2.75, h: 0.6, fontSize: 18, bold: true, color: C.accent2, valign: "middle", fontFace: THEME.headFontFace });
    }
    tb(s, "If you cannot update today: change the config", { x: 7.0, y: 1.5, w: 5.7, h: 0.45, fontSize: 18, bold: true, color: C.text2 });
    const wa = [
      [fa.FaFolderOpen, "CVE-2026-15176", "Avoid alias inside location; avoid $request_filename"],
      [fa.FaToggleOff, "CVE-2026-97205", "Turn off AIO or SSI (AIO off may slow large static files)"],
      [fa.FaTag, "CVE-2026-93479", "Use named captures in map, only inside the regex block"],
    ];
    for (let i = 0; i < 3; i++) {
      const y = 2.05 + i * 1.2;
      const [ic, head, txt] = wa[i];
      await iconCircle(s, ic, 7.0, y + 0.05, 0.8, C.text2, H.accent1, head);
      tb(s, head, { x: 8.05, y, w: 4.65, h: 0.42, fontSize: 18, bold: true, color: C.text2 });
      tb(s, txt, { x: 8.05, y: y + 0.42, w: 4.65, h: 0.6, fontSize: 14, color: C.text1 });
    }
    tb(s, "Ubuntu/Debian: apt may ship an older version number with the fix backported. Check the distro advisory (USN), not only nginx -v.", { x: 7.0, y: 5.75, w: 5.7, h: 0.8, fontSize: 13, italic: true, color: C.accent3 });
    notes(s, "Old unsupported branches get no fix; move to a supported branch. Also: NGINX Ingress Controller users should move to 5.7.0 or the relevant LTS update.");
  }

  // 10. Check your server
  {
    const s = contentSlide("Exercise: is our server exploitable?", S3);
    card(s, 0.6, 1.55, 6.3, 4.4, "Commands card", C.text2);
    tb(s, "Check in 2 minutes", { x: 0.9, y: 1.75, w: 5.7, h: 0.45, fontSize: 18, bold: true, color: C.accent1 });
    tb(s, "nginx -v\n\nsudo nginx -T | grep -nE \\\n  'alias|request_filename|aio|ssi|map '\n\nsudo nginx -t && sudo systemctl reload nginx", { x: 0.9, y: 2.35, w: 5.85, h: 3.4, fontSize: 14, color: C.background1, fontFace: "Courier New" });
    card(s, 7.3, 1.55, 5.4, 4.4, "Notes app card");
    await iconCircle(s, fa.FaDocker, 7.6, 1.85, 0.9, C.accent4, "FFFFFF", "Docker");
    tb(s, "Our Notes app", { x: 8.75, y: 1.95, w: 3.8, h: 0.45, fontSize: 20, bold: true, color: C.text2 });
    tb(s, "frontend: nginx:1.27-alpine", { x: 8.75, y: 2.4, w: 3.8, h: 0.4, fontSize: 14, color: C.accent5, fontFace: "Courier New" });
    tb(s, [
      { text: "Version in the affected range?", options: { bold: true, color: C.text2, breakLine: true } },
      { text: "Dangerous directives in nginx.conf?", options: { bold: true, color: C.text2, breakLine: true } },
      { text: "So: exploitable, or not? And should we update anyway?", options: { bold: true, color: C.accent3 } },
    ], { x: 7.6, y: 3.1, w: 4.9, h: 2.6, fontSize: 16, paraSpaceAfter: 12 });
    tb(s, "nginx -T prints the full active configuration, including every included file.", { x: 0.6, y: 6.15, w: 12.1, h: 0.5, fontSize: 15, italic: true, color: C.accent5 });
    notes(s, "Answer: 1.27 is in the affected version range, but labs/notes-app/frontend/nginx.conf only uses root, try_files and proxy_pass: no alias, SSI, AIO or map, so these three bugs cannot be triggered. Still, 1.27 is old and has other fixed bugs, so a real project would update the image.");
  }

  // 11. Comparison
  {
    const s = contentSlide("Same week, two lessons", S3);
    const rows = [
      ["", "PHP2Shell", "NGINX (AL-2026-150)"],
      ["Trigger", "Default settings", "Rare, non-default config"],
      ["Who is exposed", "Almost every PHP site", "Few servers in practice"],
      ["Disclosure", "Leaked, attacks before patch", "Coordinated, patch first"],
      ["Main risk", "Code execution", "Denial of service"],
      ["Bug class", "Use-after-free in C", "Overflow / use-after-free in C"],
    ];
    const hdr = () => ({ bold: true, color: H.lt1, fill: { color: H.dk2 }, fontSize: 16 });
    s.addTable(rows.map((r, i) => r.map((t, j) => ({ text: t, options: i === 0 ? hdr() : { fontSize: 15, color: j === 0 ? H.dk2 : H.dk1, bold: j === 0, fill: { color: i % 2 ? H.lt2 : H.lt1 } } }))), { x: 0.6, y: 1.55, w: 12.1, colW: [2.9, 4.6, 4.6], rowH: 0.55, fontFace: THEME.bodyFontFace, border: { type: "none" }, valign: "middle", margin: [0, 0.15, 0, 0.15], objectName: "Comparison table" });
    await iconCircle(s, fa.FaLightbulb, 0.6, 5.2, 0.9, C.accent1, "FFFFFF", "Lesson");
    tb(s, [
      { text: "Common lesson: ", options: { bold: true, color: C.text2 } },
      { text: "memory bugs in C infrastructure we never see. Inventory, fast updates, minimal config (enable only what you use), and layers like ASLR and least privilege.", options: { color: C.text1 } },
    ], { x: 1.75, y: 5.05, w: 10.95, h: 1.2, fontSize: 17, valign: "middle" });
    notes(s, "Two incidents in the same week, opposite profiles. Use the table to ask students which one they would patch first on a server that runs both (answer: PHP, because default config and active attacks).");
  }

  // 12. Closing
  {
    const s = titleSlide(S3, "Enable only what you use.", "Discussion: you run 50 NGINX servers. How do you find the vulnerable configs in 30 minutes?", await icon(fa.FaQuestionCircle, H.accent1), "Question");
    notes(s, "Sources: CSA alert AL-2026-150; F5 advisories K000162416, K000162607, K000162608; nginx.org security advisories; BleepingComputer, Cybersecurity News, eSecurity Planet.");
  }

  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  console.log("written", OUT);
})().catch((e) => { console.error(e); process.exit(1); });
