const pptxgen = require("pptxgenjs");
const React = require("react");
const { renderToStaticMarkup } = require("react-dom/server");
const sharp = require("sharp");
const fa = require("react-icons/fa");
const { applyTheme } = require("/Users/sergiuchilat/.claude/skills/synced/e25047ed-7c97-4849-9da1-c59bbd817501_b100ab20-30b2-4fd4-bf29-fbdff3299f26/pptx/scripts/apply_theme.js");

const OUT = "../Talk_PHP2Shell_October_2026.pptx";

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
pres.title = "PHP2Shell – Case Study (October 2026)";
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
const FOOTER = "Web Application & Infrastructure Security  ·  Case study: PHP2Shell";

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
  "Durată: 10–15 minute. Începe cu o întrebare către studenți: „Codul vostru e perfect: fără SQL injection, fără XSS. Poate fi totuși spart site-ul?” Răspunsul: da, dacă e vulnerabil chiar limbajul pe care rulează. Exact asta s-a întâmplat cu PHP pe 21 octombrie 2026.",
  // 2
  "RCE (Remote Code Execution) e cea mai gravă categorie: atacatorul rulează comenzi pe server cu drepturile utilizatorului serverului web (www-data). Numele PHP2Shell înseamnă „de la o cerere PHP la un shell pe server”.\n\nCifrele: scor CVSS 9.2 (oficial; postările scurse înainte de publicare spuneau 10.0, greșit). Nu e nevoie de cont, parolă sau vreun click din partea victimei. PHP rulează pe circa 70% din site-urile la care se știe limbajul de pe server. Toate versiunile de la 5.4 până la 8.5.11 sunt afectate.\n\nPe lângă execuția de cod, atacatorul poate citi memoria altor cereri: parole, token-uri de sesiune, chei API.",
  // 3
  "Afectat e orice server care rulează PHP și răspunde la cereri web: PHP-FPM, Apache mod_php, php-cgi, FrankenPHP, LiteSpeed. Prioritatea maximă a fost mod_php, pentru că atacul funcționează înainte să ruleze vreo linie din codul aplicației. Scripturile PHP rulate din linia de comandă (cron) nu sunt ținta, fiindcă acolo nu există o cerere HTTP controlată de atacator.\n\nExemple concrete: Wiz a demonstrat execuție de cod pe WordPress, Drupal și Joomla. Furnizorii de firewall-uri (pfSense, OPNsense, Cisco, Palo Alto) au publicat buletine pentru produsele lor. Universitatea din Washington a tăiat accesul la rețea (null route) pentru serverele nepatchuite până la actualizare.\n\nMesajul principal: codul aplicației poate fi 100% corect și totuși serverul e compromis, pentru că bug-ul e mai jos, în runtime.",
  // 4
  "Dacă se poate, desenează schema pe tablă.\n\nPHP-FPM și Apache nu pornesc un proces nou pentru fiecare cerere, ar fi prea lent. Țin un set de procese („workeri”) și fiecare servește mii de cereri la rând. Memoria din interiorul unui worker rămâne de la o cerere la alta, iar PHP trebuie să o curețe la final (pasul 4).\n\nPresupunerea pe care se bazează totul: curățarea se termină mereu corect, deci cererea B nu vede nimic din cererea A. PHP2Shell strică exact această presupunere. Ambele bug-uri sunt în pasul marcat cu roșu.\n\nLegătura cu cursul: în Lab 5 lucrăm cu Nginx + PHP-FPM / backend exact pe acest model.",
  // 5
  "Primul bug. Atacatorul trimite un POST multipart/form-data uriaș, fără antetul Content-Length (cu Transfer-Encoding: chunked). Așa, limita post_max_size nu se aplică. Mii de părți și antete umplu memoria.\n\nPHP rămâne fără memorie exact când datele cererii sunt pe jumătate construite și abandonează cererea. Dar rutina de curățare rulează oricum și tratează datele incomplete ca fiind valide: eliberează și folosește pointeri invalizi. Cu o așezare atentă a memoriei (heap), corupția devine execuție de cod.\n\nExemplul de pe slide e simplificat, scop didactic. PoC-ul public folosea un POST multipart cu un nume de fișier extrem de lung.\n\nIdeea pentru studenți: căile de eroare sunt rar testate. Codul care tratează „memorie insuficientă” aproape nu rulează în viața reală, așa că bug-urile de acolo supraviețuiesc ani de zile. Pe mod_php e nevoie de o singură cerere per încercare; pentru PHP-FPM nu s-a arătat un exploit complet pentru acest CVE.",
  // 6
  "Al doilea bug e între două cereri. Cererea 1 („producătorul”) provoacă o eroare fatală: timeout, memorie epuizată, un require eșuat. PHP se oprește înainte să reseteze o parte din starea globală (erori, sesiune, backtrace), iar acea stare arată spre memorie deja eliberată.\n\nCererea 2 („consumatorul”) ajunge la același worker. Ceva rulează înainte de inițializarea normală, de exemplu progresul upload-ului din sesiune sau un avertisment de la o setare INI din .htaccess, și citește pointerul vechi: use-after-free, deci execuție de cod sau scurgere de memorie.\n\nAtacatorul nu controlează la ce worker ajunge, așa că trimite multe perechi de cereri până nimerește.\n\nDetaliu important din advisory: ascunderea erorilor (display_errors=Off, error_reporting, operatorul @) NU ajută. Starea veche e folosită chiar dacă nu se afișează și nu se loghează nimic. Deci „logurile mele sunt curate” nu e o dovadă că serverul e în siguranță.",
  // 7
  "Ce obține atacatorul:\n1) Execuție de cod ca www-data. Pasul tipic următor: un webshell, adică un fișier PHP ascuns prin care revine oricând.\n2) Chiar fără execuție de cod: citește din memorie parolele, sesiunile și cheile API ale altor utilizatori.\n3) De acolo mai departe: citește config.php, ia parola bazei de date, descarcă toată baza. Circa 394 000 de servere PHP au și SSH expus, deci se poate muta și pe alte sisteme.\n\nLegătura cu Lecția 1: vectorul CVSS AV:N/AC:H/AT:P/PR:N/UI:N/VC:H/VI:H/VA:H. Atac prin rețea, fără privilegii, fără interacțiunea utilizatorului, impact mare pe toate trei: confidențialitate, integritate, disponibilitate. Complexitatea e „High” pentru că depinde de așezarea memoriei și de ce worker primește cererea; de aceea 9.2 și nu 10.\n\nÎntrebare pentru studenți: care literă din triada CIA e încălcată de fiecare dintre cele trei carduri?",
  // 8
  "E o lecție despre dezvăluirea coordonată (coordinated disclosure). În mod normal: cercetătorul raportează în privat, producătorul repară, distribuțiile Linux pregătesc pachetele și toți publică în aceeași zi, ca administratorii să poată actualiza înainte să atace cineva.\n\nAici: Wiz Research găsește bug-urile în august (folosind și modele AI) și demonstrează atacul pe WordPress și Drupal. Pe 21 septembrie echipa PHP confirmă și începe lucrul. Pe 20 octombrie identificatorii CVE apar pe Hacker News înainte de termen. În noaptea de 21 octombrie o arhivă cu detalii apare pe VirusTotal, iar dimineața Patchstack și Cloudflare raportează primele încercări de atac, cu ore înainte să existe patch-ul. Practic o fereastră de zero-day.\n\nLa 13:52 UTC PHP publică versiunile reparate; CISA, CERT-EU și NCSC (Marea Britanie) emit alerte urgente; apar PoC-uri publice, unele nefuncționale.\n\nAtenție la formulare: la momentul prezentării nu se știe încă câte atacuri au reușit.",
  // 9
  "Versiunile reparate, publicate pe 21 octombrie: 8.5.12, 8.4.27, 8.3.36, 8.2.35. PHP 8.1 și mai vechi nu mai sunt suportate: nu primesc versiune oficială, doar corecții la nivel de cod sursă pentru 7.4–8.1 sau pachete de la distribuții (Ubuntu, Red Hat etc.).\n\nCe schimbă patch-ul, pe scurt: curățarea nu mai lucrează cu date pe jumătate construite, starea globală e resetată între cereri, iar un worker care atinge limita de memorie e oprit, nu continuă.\n\nActualizarea e SINGURA soluție completă; advisory-ul spune clar că nu există o configurare care să elimine complet problema.\n\nCapcană practică: „php -v” din terminal poate fi alt PHP decât cel care servește site-ul (FPM, container, panoul de hosting). Și după actualizare trebuie repornite procesele, altfel binarul vechi rulează în continuare în memorie.",
  // 10
  "Dacă nu poți actualiza imediat, există măsuri temporare. Ele reduc riscul, nu îl elimină.\n\nÎn php.ini sau în configurația pool-ului FPM, apoi restart: dezactivezi session.upload_progress, fatal_error_backtraces (pe 8.5), limitezi mărimea cererilor; opțional pm.max_requests=1, adică fiecare worker servește o singură cerere (scump ca performanță). Setările trebuie puse în configurație, nu cu ini_set() în cod: bug-ul se declanșează înainte să pornească scriptul.\n\nÎn fața serverului: WAF sau reverse proxy care limitează dimensiunea cererilor și numărul de părți multipart. Izolare: acces doar prin VPN sau listă de IP-uri, ori oprești serviciul (recomandarea CERT-EU). Verificare: cauți înapoi de pe 20 octombrie fișiere noi în directorul site-ului, procese lansate de php-fpm, trafic de ieșire neobișnuit.\n\nCe NU ajută: ascunderea erorilor sau a versiunii PHP. Securitatea prin obscuritate nu e protecție.\n\nÎntrebare pentru studenți: care dintre acestea sunt „apărare în adâncime”? (WAF, izolare, configurare = straturi; patch-ul = reparația reală.)",
  // 11
  "Lecțiile pentru noi:\n1) Să știi ce rulezi. Fă inventarul fiecărui PHP: servere, containere, echipamente. Nu poți actualiza ce nu știi că există.\n2) Versiunile ieșite din suport sunt un risc. Circa 880 000 de servere publice rulau versiuni care nu au primit nicio corecție (PHP 5, 7, 8.0, 8.1).\n3) Actualizarea e un proces: rapid, imagini de container reconstruite automat, procese repornite.\n4) Apărare în adâncime: cod sigur nu ajunge. Privilegii minime pentru www-data, firewall, WAF, monitorizare.\n\nÎn laboratoare: Lab 4 (unattended-upgrades pe server) și Lab 5 (limite în Nginx, reverse proxy) sunt exact aceste controale.\n\nCifrele de expunere sunt estimări din scanări ale internetului (Mallory.ai / Phoenix Security, Modat ~4,4 milioane de servicii PHP), le prezentăm ca ordine de mărime. Merită spus și că AI a fost folosit atât pentru găsirea bug-urilor, cât și, la câteva ore după scurgere, pentru încercarea de a le reproduce: timpul dintre publicare și atac se scurtează.",
  // 12
  "Încheiere și discuție: „Cum ați găsi într-o oră fiecare instalare PHP dintr-o organizație?” Răspunsuri bune: scanare de rețea, inventar de pachete pe servere, scanarea imaginilor de container, întrebat furnizorii de echipamente, căutat antetul X-Powered-By.\n\nSurse: advisory-urile PHP GHSA-w8hw-w9fx-j79w și GHSA-6wmf-4rv5-p3rp; alerta CISA din 21.10.2026; Rapid7, Qualys, Google Cloud, Recorded Future, SANS ISC, Infosecurity Magazine, UW-IT. Unele detalii (numărul exact de atacuri reușite) erau încă în evoluție pe 22.10.2026.",
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
    const s = titleSlide(S1, "PHP2Shell: when the language itself is vulnerable", "CVE-2026-109618 & CVE-2026-109619  ·  Case study, October 2026", await icon(fa.FaBug, H.accent1), "Bug");
    notes(s, "10-15 minute talk. Start with a question: 'Your code is perfect, no SQL injection, no XSS. Can your site still be hacked?' Answer: yes, if the runtime under it is broken. That is exactly what happened to PHP on 21 October 2026.");
  }

  // 2. Key facts
  {
    const s = contentSlide("PHP2Shell in one slide", S1);
    const stats = [
      ["9.2", "CVSS v4.0, Critical (both CVEs)", C.accent3],
      ["0", "credentials or clicks needed: remote and unauthenticated", C.accent1],
      ["~70%", "of websites with a known back-end language run PHP (W3Techs)", C.accent4],
      ["5.4 → 8.5.11", "every version in this range is vulnerable", C.accent2],
    ];
    for (let i = 0; i < 4; i++) {
      const x = 0.6 + i * 3.1;
      card(s, x, 1.6, 2.8, 2.9, "Stat card " + (i + 1));
      tb(s, stats[i][0], { x: x + 0.25, y: 1.85, w: 2.3, h: 1.0, fontSize: i === 3 ? 30 : 48, bold: true, color: stats[i][2], fontFace: THEME.headFontFace, valign: "middle" });
      tb(s, stats[i][1], { x: x + 0.25, y: 3.0, w: 2.3, h: 1.3, fontSize: 16, color: C.text1 });
    }
    card(s, 0.6, 4.85, 12.1, 1.75, "Summary card");
    await iconCircle(s, fa.FaTerminal, 0.95, 5.25, 0.95, C.accent3, "FFFFFF", "Terminal");
    tb(s, [
      { text: "Two memory-safety bugs in the PHP interpreter (written in C). ", options: { bold: true, color: C.text2 } },
      { text: "With specially crafted HTTP requests an attacker can run their own code on the server, or read memory from other users' requests: passwords, session tokens, API keys.", options: { color: C.text1 } },
    ], { x: 2.2, y: 4.85, w: 10.2, h: 1.75, fontSize: 17, valign: "middle" });
    notes(s, "RCE = Remote Code Execution, the worst category: the attacker runs commands with the rights of the web server user (www-data). The nickname PHP2Shell means 'from a PHP request to a shell on the server'. Official score is 9.2; leaked posts before release claimed 10.0, which was wrong.");
  }

  // 3. Who is exposed
  {
    const s = contentSlide("Who was affected", S1);
    const groups = [
      [fa.FaWordpress, "CMS and shops", "WordPress, WooCommerce, Drupal, Joomla, Magento / Adobe Commerce. RCE demonstrated on WordPress, Drupal and Joomla."],
      [fa.FaCloud, "Hosting and cloud", "Shared hosting, Hostinger, AWS, Alibaba Cloud. Google App Engine and Cloud Run patched images; GKE and custom containers had to be rebuilt."],
      [fa.FaNetworkWired, "Devices with a web UI", "Firewalls and appliances that use PHP: pfSense, OPNsense, Cisco and Palo Alto products published bulletins."],
      [fa.FaUniversity, "Universities, companies", "Nextcloud, phpMyAdmin, internal portals. Univ. of Washington cut network access (null route) to unpatched servers."],
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
      { text: "Key idea: ", options: { bold: true, color: C.accent3 } },
      { text: "the application code could be 100% correct and still be hacked. The bug is in the runtime, below the application.", options: { color: C.text1 } },
    ], { x: 0.6, y: 5.75, w: 12.1, h: 0.8, fontSize: 18, valign: "middle" });
    notes(s, "Every server that runs PHP and answers web requests: PHP-FPM, Apache mod_php, php-cgi, FrankenPHP, LiteSpeed. mod_php was the highest priority because the exploit fires before any application code runs. PHP CLI scripts (cron jobs) are not the target, since there is no attacker-controlled HTTP request.");
  }

  // ================= HOW IT WORKS =================
  const S2 = "How it works";
  pres.addSection({ title: S2 });

  // 4. Background: worker reuse
  {
    const s = contentSlide("Background: PHP workers are reused", S2);
    const steps = [
      ["1", "Request A", "arrives at worker #3"],
      ["2", "Startup", "PHP sets up request state"],
      ["3", "Your code runs", "index.php, WordPress…"],
      ["4", "Cleanup", "memory freed, state reset"],
      ["5", "Request B", "same worker #3, next user"],
    ];
    const w = 2.15, gap = 0.34;
    for (let i = 0; i < steps.length; i++) {
      const x = 0.6 + i * (w + gap);
      const col = i === 3 ? C.accent3 : C.text2;
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y: 1.75, w, h: 1.5, rectRadius: 0.12, fill: { color: col }, line: { type: "none" }, objectName: "Step " + steps[i][0] });
      tb(s, steps[i][1], { x: x + 0.15, y: 1.9, w: w - 0.3, h: 0.5, fontSize: 19, bold: true, color: C.background1, align: "center" });
      tb(s, steps[i][2], { x: x + 0.15, y: 2.45, w: w - 0.3, h: 0.7, fontSize: 14, color: C.background1, align: "center" });
      if (i < steps.length - 1) s.addShape(pres.shapes.LINE, { x: x + w + 0.04, y: 2.5, w: gap - 0.08, h: 0, line: { color: H.accent5, width: 2, endArrowType: "triangle" }, objectName: "Arrow " + (i + 1) });
    }
    card(s, 0.6, 3.75, 12.1, 2.75, "Explanation card");
    tb(s, [
      { text: "Why reuse? ", options: { bold: true, color: C.text2 } },
      { text: "Starting a new process for each request is slow, so PHP-FPM and Apache keep a pool of workers and reuse them for thousands of requests.", options: { color: C.text1, breakLine: true } },
      { text: " ", options: { breakLine: true, fontSize: 8 } },
      { text: "The assumption: ", options: { bold: true, color: C.text2 } },
      { text: "step 4 (cleanup) always finishes cleanly, so request B never sees anything from request A.", options: { color: C.text1, breakLine: true } },
      { text: " ", options: { breakLine: true, fontSize: 8 } },
      { text: "PHP2Shell breaks exactly this assumption. ", options: { bold: true, color: C.accent3 } },
      { text: "Both bugs live in the red box.", options: { color: C.text1 } },
    ], { x: 0.95, y: 3.95, w: 11.4, h: 2.4, fontSize: 18, valign: "middle" });
    notes(s, "Draw this on the board if possible. Connect to Lab 5: PHP-FPM/Nginx work exactly like this. A worker is a process; memory inside it survives from one request to the next unless PHP resets it.");
  }

  // 5. Bug 1
  {
    const s = contentSlide("Bug 1 · CVE-2026-109618 · out of memory", S2);
    const flow = [
      [fa.FaPaperPlane, "Attacker sends", "a huge multipart/form-data POST, without Content-Length, so post_max_size is not enforced"],
      [fa.FaMemory, "PHP runs out of memory", "while the request data is only half initialized, and abandons the request"],
      [fa.FaBroom, "Cleanup runs anyway", "and treats the half-built data as valid: it frees and uses garbage pointers"],
      [fa.FaSkullCrossbones, "Memory corruption", "a carefully shaped heap turns this into code execution"],
    ];
    for (let i = 0; i < 4; i++) {
      const y = 1.55 + i * 1.12;
      const [ic, head, txt] = flow[i];
      await iconCircle(s, ic, 0.6, y + 0.05, 0.85, i === 3 ? C.accent3 : C.text2, H.accent1, head);
      tb(s, head, { x: 1.7, y, w: 5.6, h: 0.42, fontSize: 19, bold: true, color: C.text2 });
      tb(s, txt, { x: 1.7, y: y + 0.42, w: 5.6, h: 0.6, fontSize: 14, color: C.text1 });
    }
    card(s, 7.7, 1.55, 5.0, 4.4, "Request example card", C.text2);
    tb(s, "What a trigger looks like", { x: 8.0, y: 1.75, w: 4.4, h: 0.45, fontSize: 18, bold: true, color: C.accent1 });
    tb(s, "POST /index.php HTTP/1.1\nTransfer-Encoding: chunked\nContent-Type: multipart/form-data;\n  boundary=X\n\n--X\nContent-Disposition: form-data;\n  name=\"f\"; filename=\"AAAA…AAAA\"\n(thousands of parts and headers)", { x: 8.0, y: 2.3, w: 4.45, h: 3.5, fontSize: 13, color: C.background1, fontFace: "Courier New" });
    tb(s, [
      { text: "Bug class: ", options: { bold: true, color: C.text2 } },
      { text: "use-after-free / uninitialized pointer (CWE-416, CWE-824) in error handling (CWE-755). One request per attempt on mod_php.", options: { color: C.text1 } },
    ], { x: 0.6, y: 6.1, w: 12.1, h: 0.7, fontSize: 15, valign: "middle" });
    notes(s, "Simplified example; the public PoC used a multipart POST with an extremely long filename. The point for students: error paths are rarely tested. Code that handles 'out of memory' runs almost never in normal life, so bugs there survive for years. On mod_php the attack works before application code even starts. For PHP-FPM no full exploit was shown for this specific CVE.");
  }

  // 6. Bug 2
  {
    const s = contentSlide("Bug 2 · CVE-2026-109619 · stale state", S2);
    const cols = [
      ["Request 1: the producer", C.accent1, [
        "Triggers a fatal error: timeout, out of memory, failed require…",
        "PHP stops before resetting some global state (errors, session, backtrace)",
        "That state now points to memory that has been freed",
      ]],
      ["Request 2: the consumer", C.accent3, [
        "Arrives at the same worker",
        "Something runs before normal init: session upload progress, an INI warning from .htaccess",
        "It reads the stale pointer: use-after-free → RCE or memory leak",
      ]],
    ];
    for (let i = 0; i < 2; i++) {
      const x = 0.6 + i * 6.3;
      const [head, col, items] = cols[i];
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y: 1.55, w: 5.8, h: 0.7, rectRadius: 0.12, fill: { color: col }, line: { type: "none" }, objectName: head + " header" });
      tb(s, head, { x, y: 1.55, w: 5.8, h: 0.7, fontSize: 20, bold: true, color: C.background1, align: "center", valign: "middle" });
      card(s, x, 2.4, 5.8, 2.55, head + " card");
      tb(s, items.map((t, j) => ({ text: t, options: { bullet: true, breakLine: j < items.length - 1, paraSpaceAfter: 8 } })), { x: x + 0.3, y: 2.6, w: 5.25, h: 2.2, fontSize: 16, color: C.text1 });
    }
    s.addShape(pres.shapes.LINE, { x: 6.45, y: 1.9, w: 0.4, h: 0, line: { color: H.accent5, width: 2, endArrowType: "triangle" }, objectName: "Producer to consumer arrow" });
    card(s, 0.6, 5.2, 12.1, 1.4, "Default settings card");
    await iconCircle(s, fa.FaCog, 0.9, 5.42, 0.95, C.text2, H.accent1, "Settings");
    tb(s, [
      { text: "Enabled by default: ", options: { bold: true, color: C.accent3 } },
      { text: "session.upload_progress (since PHP 5.4, 2012) and fatal_error_backtraces (PHP 8.5). Confirmed RCE on Apache mod_php and on PHP-FPM.", options: { color: C.text1 } },
    ], { x: 2.15, y: 5.2, w: 10.3, h: 1.4, fontSize: 17, valign: "middle" });
    notes(s, "This is the cross-request bug. Two requests have to land on the same worker; the attacker just sends many pairs until it works. Important detail from the advisory: hiding errors with display_errors=Off, error_reporting or @ does NOT help. The stale data is used even if nothing is printed or logged. So 'my logs are clean' is not proof of safety.");
  }

  // 7. Impact
  {
    const s = contentSlide("What an attacker gets", S2);
    const imp = [
      [fa.FaTerminal, "Code execution", "Commands run as the web server user (www-data). Typical next step: drop a webshell for persistent access.", C.accent3],
      [fa.FaKey, "Memory disclosure", "Even without RCE: read other users' passwords, session tokens, API keys from worker memory.", C.accent1],
      [fa.FaDatabase, "Data and pivoting", "Read config.php → database credentials → full database. ~394,000 PHP hosts also expose SSH.", C.accent4],
    ];
    for (let i = 0; i < 3; i++) {
      const x = 0.6 + i * 4.15;
      const [ic, head, txt, col] = imp[i];
      card(s, x, 1.6, 3.8, 3.7, head + " card");
      await iconCircle(s, ic, x + 0.35, 1.9, 0.95, col, "FFFFFF", head);
      tb(s, head, { x: x + 0.35, y: 3.0, w: 3.1, h: 0.5, fontSize: 22, bold: true, color: C.text2, fontFace: THEME.headFontFace });
      tb(s, txt, { x: x + 0.35, y: 3.55, w: 3.15, h: 1.6, fontSize: 15, color: C.text1 });
    }
    tb(s, [
      { text: "CIA triad: ", options: { bold: true, color: C.text2 } },
      { text: "all three broken. Confidentiality (leaked secrets), Integrity (modified files, webshells), Availability (crashed workers).", options: { color: C.text1 } },
    ], { x: 0.6, y: 5.65, w: 12.1, h: 0.8, fontSize: 18, valign: "middle" });
    notes(s, "Link back to Lesson 1: CVSS vector AV:N/AC:H/AT:P/PR:N/UI:N/VC:H/VI:H/VA:H. Network attack, no privileges, no user interaction, high impact on C, I and A. AC:H (high complexity) because it depends on heap layout and worker scheduling, which is why it is 9.2 and not 10.");
  }

  // ================= RESPONSE =================
  const S3 = "Response";
  pres.addSection({ title: S3 });

  // 8. Timeline
  {
    const s = contentSlide("Timeline: a disclosure that leaked", S3);
    const ev = [
      ["Aug 19–24", "Wiz Research finds the bugs (with AI help), proves RCE on WordPress and Drupal, reports privately", C.accent2],
      ["Sep 21", "PHP security team confirms and starts writing fixes. Distros get patches in advance", C.accent2],
      ["Oct 20", "CVE IDs leak on Hacker News before the embargo. PHP announces release for next day", C.accent1],
      ["Oct 21 · 02:29", "Archive with exploit details appears on VirusTotal", C.accent3],
      ["Oct 21 · 07:34", "Patchstack reports first attacks via hosting partners. Cloudflare sees attempts", C.accent3],
      ["Oct 21 · 13:52", "PHP releases fixes. CISA, CERT-EU and UK NCSC issue urgent alerts. Public PoCs appear", C.accent4],
    ];
    const x0 = 2.9;
    s.addShape(pres.shapes.LINE, { x: x0 + 0.14, y: 1.65, w: 0, h: 5.0, line: { color: H.accent5, width: 2 }, objectName: "Timeline line" });
    for (let i = 0; i < ev.length; i++) {
      const y = 1.55 + i * 0.86;
      const [when, what, col] = ev[i];
      tb(s, when, { x: 0.6, y, w: 2.1, h: 0.5, fontSize: 16, bold: true, color: C.text2, align: "right", valign: "middle" });
      s.addShape(pres.shapes.OVAL, { x: x0, y: y + 0.11, w: 0.28, h: 0.28, fill: { color: col }, line: { color: H.lt1, width: 2 }, objectName: "Dot " + (i + 1) });
      tb(s, what, { x: x0 + 0.6, y, w: 9.2, h: 0.5, fontSize: 16, color: C.text1, valign: "middle" });
    }
    notes(s, "This is a lesson in coordinated disclosure. Normally: researcher reports privately, vendor fixes, distros prepare packages, everyone publishes on the same day. Here the CVE IDs leaked a day early and an archive with details hit VirusTotal, so attacks started hours BEFORE the patch existed. That is effectively a zero-day window. Note: reports of attacks are early; the number of successful compromises was not confirmed at the time.");
  }

  // 9. The fix
  {
    const s = contentSlide("How it was fixed", S3);
    const vers = [["8.5", "8.5.12"], ["8.4", "8.4.27"], ["8.3", "8.3.36"], ["8.2", "8.2.35"]];
    tb(s, "Patched releases, 21 October 2026", { x: 0.6, y: 1.5, w: 5.6, h: 0.45, fontSize: 18, bold: true, color: C.text2 });
    for (let i = 0; i < 4; i++) {
      const y = 2.05 + i * 0.85;
      card(s, 0.6, y, 5.6, 0.7, "Version " + vers[i][0]);
      tb(s, "PHP " + vers[i][0], { x: 0.9, y, w: 2.2, h: 0.7, fontSize: 18, color: C.accent5, valign: "middle" });
      tb(s, vers[i][1], { x: 3.1, y, w: 2.9, h: 0.7, fontSize: 24, bold: true, color: C.accent2, valign: "middle", fontFace: THEME.headFontFace });
    }
    tb(s, "PHP 8.1 and older (EOL): no official release. Only source backports for 7.4–8.1, or a vendor/distro package.", { x: 0.6, y: 5.5, w: 5.6, h: 1.0, fontSize: 15, italic: true, color: C.accent3 });

    tb(s, "What the patch changes", { x: 6.85, y: 1.5, w: 5.85, h: 0.45, fontSize: 18, bold: true, color: C.text2 });
    const fixes = [
      [fa.FaShieldAlt, "Cleanup checks state", "Out-of-memory paths no longer free half-built request data"],
      [fa.FaUndo, "Reset between requests", "Global error, session and backtrace state is cleared before reuse"],
      [fa.FaPowerOff, "Hard memory limit", "A worker at its memory limit is now killed instead of continuing"],
    ];
    for (let i = 0; i < 3; i++) {
      const y = 2.05 + i * 1.2;
      const [ic, head, txt] = fixes[i];
      await iconCircle(s, ic, 6.85, y + 0.05, 0.8, C.text2, H.accent1, head);
      tb(s, head, { x: 7.9, y, w: 4.8, h: 0.42, fontSize: 18, bold: true, color: C.text2 });
      tb(s, txt, { x: 7.9, y: y + 0.42, w: 4.8, h: 0.6, fontSize: 14, color: C.text1 });
    }
    tb(s, "Fixes by the PHP core team (arnaud-lb, iluuu1994, bukka, ndossche). Advisories: GHSA-w8hw-w9fx-j79w, GHSA-6wmf-4rv5-p3rp", { x: 6.85, y: 5.75, w: 5.85, h: 0.75, fontSize: 12, color: C.accent5 });
    notes(s, "Updating is the ONLY complete fix; the advisory says there is no complete configuration workaround. Practical gotcha: 'php -v' on the command line may be a different PHP than the one serving the website (FPM, a container, a hosting panel). After updating you must restart the workers, otherwise the old binary keeps running in memory.");
  }

  // 10. If you cannot patch today
  {
    const s = contentSlide("If you cannot patch today", S3);
    tb(s, "Temporary mitigations: they reduce risk, they do not remove it", { x: 0.6, y: 1.45, w: 12.1, h: 0.45, fontSize: 18, italic: true, color: C.accent5 });
    card(s, 0.6, 2.1, 6.2, 3.9, "php.ini card", C.text2);
    tb(s, "php.ini / FPM pool, then restart", { x: 0.9, y: 2.3, w: 5.6, h: 0.45, fontSize: 18, bold: true, color: C.accent1 });
    tb(s, "session.upload_progress.enabled = 0\nfatal_error_backtraces = 0   ; PHP 8.5\npost_max_size = 8M\nupload_max_filesize = 4M\npm.max_requests = 1          ; optional", { x: 0.9, y: 2.9, w: 5.75, h: 2.9, fontSize: 15, color: C.background1, fontFace: "Courier New" });
    const other = [
      [fa.FaFilter, "WAF / reverse proxy", "Limit request size and number of multipart parts; block multipart where no uploads are used"],
      [fa.FaBan, "Isolate", "Restrict to VPN or an IP allowlist, or take the service offline (CERT-EU advice)"],
      [fa.FaSearch, "Hunt", "Look back from Oct 20: unknown files in the web root, child processes of php-fpm, odd outbound traffic"],
    ];
    for (let i = 0; i < 3; i++) {
      const y = 2.1 + i * 1.33;
      const [ic, head, txt] = other[i];
      await iconCircle(s, ic, 7.2, y + 0.05, 0.8, C.text2, H.accent1, head);
      tb(s, head, { x: 8.25, y, w: 4.45, h: 0.42, fontSize: 18, bold: true, color: C.text2 });
      tb(s, txt, { x: 8.25, y: y + 0.42, w: 4.45, h: 0.85, fontSize: 14, color: C.text1 });
    }
    tb(s, [
      { text: "Does not help: ", options: { bold: true, color: C.accent3 } },
      { text: "display_errors=Off, the @ operator, hiding the PHP version (expose_php=Off). Obscurity is not protection.", options: { color: C.text1 } },
    ], { x: 0.6, y: 6.2, w: 12.1, h: 0.6, fontSize: 16, valign: "middle" });
    notes(s, "Settings must be set in php.ini or the pool config; ini_set() in code is too late because the bug fires before the script starts. Ask students: which of these is defense in depth? (WAF + isolation + config hardening are layers; the patch is the real fix.)");
  }

  // 11. Lessons
  {
    const s = contentSlide("Lessons for this course", S3);
    const les = [
      [fa.FaListAlt, "Know what you run", "Inventory every PHP: servers, containers, appliances. You cannot patch what you do not know about."],
      [fa.FaHourglassEnd, "End-of-life is a risk", "~880,000 public hosts ran PHP branches that got no fix at all (PHP 5, 7, 8.0, 8.1)."],
      [fa.FaSyncAlt, "Patching is a process", "Fast updates, base images that rebuild automatically, worker restarts."],
      [fa.FaLayerGroup, "Defense in depth", "Secure app code is not enough. Least privilege for www-data, firewall, WAF, monitoring."],
    ];
    for (let i = 0; i < 4; i++) {
      const x = 0.6 + (i % 2) * 6.15, y = 1.6 + Math.floor(i / 2) * 2.15;
      const [ic, head, txt] = les[i];
      card(s, x, y, 5.95, 1.9, head + " card");
      await iconCircle(s, ic, x + 0.3, y + 0.45, 0.95, C.text2, H.accent1, head);
      tb(s, head, { x: x + 1.5, y: y + 0.25, w: 4.25, h: 0.45, fontSize: 20, bold: true, color: C.text2 });
      tb(s, txt, { x: x + 1.5, y: y + 0.75, w: 4.25, h: 1.0, fontSize: 15, color: C.text1 });
    }
    tb(s, [
      { text: "In our labs: ", options: { bold: true, color: C.accent2 } },
      { text: "Lab 4 (unattended-upgrades on the server) and Lab 5 (Nginx limits, reverse proxy) are exactly these controls.", options: { color: C.text1 } },
    ], { x: 0.6, y: 6.05, w: 12.1, h: 0.7, fontSize: 17, valign: "middle" });
    notes(s, "Exposure numbers are estimates from internet scans (Mallory.ai / Phoenix Security, Modat ~4.4 million PHP services). Treat them as directional. Also worth mentioning: AI tools were used both to find these bugs and, within hours of the leak, to try to reproduce them. The time between disclosure and attack keeps shrinking.");
  }

  // 12. Closing
  {
    const s = titleSlide(S3, "Your code was fine. Your server was not.", "Discussion: how would you find every PHP install in your organisation in one hour?", await icon(fa.FaQuestionCircle, H.accent1), "Question");
    notes(s, "Sources: PHP advisories GHSA-w8hw-w9fx-j79w and GHSA-6wmf-4rv5-p3rp; CISA alert 21 Oct 2026; Rapid7, Qualys, Google Cloud, Recorded Future, SANS ISC, Infosecurity Magazine, Univ. of Washington IT. Some details (exact attack counts) were still developing on 22 Oct 2026.");
  }

  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  console.log("written", OUT);
})().catch((e) => { console.error(e); process.exit(1); });
