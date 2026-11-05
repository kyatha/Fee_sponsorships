/* ===== PhD Fee Bridge: Google Sheet backend =====
 * Setup (once):
 *   1. In your Google Sheet: Extensions > Apps Script. Paste this whole file into Code.gs and save.
 *   2. Choose the function "setup" in the toolbar and click Run. Approve the permissions.
 *      (Optional) Run "seedDemo" to add the fictional example applicants.
 *   3. Deploy > New deployment > type "Web app". Execute as: Me. Who has access: Anyone. Deploy.
 *   4. Copy the Web app URL (ends in /exec) into config.js on the website.
 * After you edit this file later: Deploy > Manage deployments > edit > Version: New version.
 * Only you can open the Sheet. The website only ever receives the public fields (see pub()).
 */

function setup() {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName("Applications") || ss.insertSheet("Applications");
  sh.getRange(1, 1, sh.getMaxRows(), COLS.length).setNumberFormat("@");
  sh.getRange(1, 1, 1, COLS.length).setValues([COLS]).setFontWeight("bold");
  sh.setFrozenRows(1);
  var lg = ss.getSheetByName("Log") || ss.insertSheet("Log");
  lg.getRange(1, 1, 1, 4).setValues([["time", "appNo", "event", "detail"]]).setFontWeight("bold");
  lg.setFrozenRows(1);
}

function seedDemo() {
  var db = sheetDb();
  seedRows(Date.now()).forEach(function (o) { db.add(o); });
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try { return json(Object.assign({ ok: true }, handle(JSON.parse(e.postData.contents), sheetDb()))); }
  catch (err) { return json({ ok: false, error: String(err && err.message || err) }); }
  finally { lock.releaseLock(); }
}
function doGet() { return doPost({ postData: { contents: '{"action":"list"}' } }); }
function json(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

function sheetDb() {
  var ss = SpreadsheetApp.getActive(), sh = ss.getSheetByName("Applications"), lg = ss.getSheetByName("Log");
  var v = sh.getDataRange().getDisplayValues(), h = v.shift();
  var list = v.map(function (r, i) {
    var o = { _row: i + 2 }; h.forEach(function (k, j) { o[k] = r[j]; }); return o;
  }).filter(function (o) { return o.id; });
  // Store everything as plain text: stops typed text being read as a formula, and keeps dates and IDs as typed.
  function safe(x) { var s = x === undefined || x === null ? "" : String(x); return s === "" ? "" : "'" + s; }
  function rowOf(o) { return COLS.map(function (k) { return safe(o[k]); }); }
  return {
    all: function () { return list; },
    now: function () { return Date.now(); },
    save: function (o) { sh.getRange(o._row, 1, 1, COLS.length).setValues([rowOf(o)]); },
    add: function (o) { sh.appendRow(rowOf(o)); o._row = sh.getLastRow(); list.push(o); },
    log: function (no, ev, d) { lg.appendRow([new Date().toISOString(), no, ev, safe(d)]); }
  };
}

/* ===== Shared rules: the same code runs in Google Apps Script (Sheet mode) and in the browser (preview mode). ===== */
var HOLD_MS = 24 * 3600e3;
var COLS = ["id", "appNo", "createdAt", "name", "email", "umassId", "country", "field", "program", "level", "deadline", "waiver", "statement",
  "status", "pledgeCode", "sponsorName", "sponsorEmail", "sponsorAnon", "pledgedAt", "payMethod", "payRef", "paidAt", "releases"];
/* For now the program is open to applicants from the 54 African countries. */
var COUNTRIES = ["Algeria", "Angola", "Benin", "Botswana", "Burkina Faso", "Burundi", "Cabo Verde", "Cameroon", "Central African Republic", "Chad", "Comoros", "Democratic Republic of the Congo", "Republic of the Congo", "Côte d'Ivoire", "Djibouti", "Egypt", "Equatorial Guinea", "Eritrea", "Eswatini", "Ethiopia", "Gabon", "The Gambia", "Ghana", "Guinea", "Guinea-Bissau", "Kenya", "Lesotho", "Liberia", "Libya", "Madagascar", "Malawi", "Mali", "Mauritania", "Mauritius", "Morocco", "Mozambique", "Namibia", "Niger", "Nigeria", "Rwanda", "São Tomé and Príncipe", "Senegal", "Seychelles", "Sierra Leone", "Somalia", "South Africa", "South Sudan", "Sudan", "Tanzania", "Togo", "Tunisia", "Uganda", "Zambia", "Zimbabwe"];
var ALPH = "ABCDEFGHJKLMNPQRSTUVWXYZ23456789";
function rnd(n) { var s = ""; for (var i = 0; i < n; i++) s += ALPH[Math.floor(Math.random() * ALPH.length)]; return s; }
function wordCount(t) { return (String(t || "").trim().match(/\S+/g) || []).length; }
function lc(s) { return String(s || "").trim().toLowerCase(); }
function fail(m) { throw new Error(m); }
function isAnon(o) { return String(o.sponsorAnon) === "true"; }

/* What anyone can see. Email, UMass applicant ID, application number and pledge code never leave through here. */
function pub(o) {
  return { id: o.id, name: o.name, country: o.country, field: o.field, program: o.program, level: o.level, deadline: o.deadline,
    waiver: o.waiver, statement: o.statement, status: o.status,
    sponsor: o.status === "open" || isAnon(o) ? "" : o.sponsorName, sponsorAnon: o.status !== "open" && isAnon(o),
    pledgedAt: Number(o.pledgedAt) || 0, releases: Number(o.releases) || 0, payRef: o.status === "paid" ? o.payRef : "" };
}
function sponsorView(o) { var v = pub(o); v.appNo = o.appNo; v.umassId = o.umassId; v.code = o.pledgeCode; v.payMethod = o.payMethod || ""; return v; }
function clearPledge(o) {
  o.status = "open"; o.pledgeCode = ""; o.sponsorName = ""; o.sponsorEmail = ""; o.sponsorAnon = ""; o.pledgedAt = "";
  o.releases = (Number(o.releases) || 0) + 1;
}
function expireHolds(db) {
  var t = db.now();
  db.all().forEach(function (o) {
    if (o.status === "pledged" && t - Number(o.pledgedAt) >= HOLD_MS) { db.log(o.appNo, "auto-released after 24h", o.pledgeCode); clearPledge(o); db.save(o); }
  });
}

function handle(req, db) {
  expireHolds(db);
  var rs = db.all(), a = req.action;
  function byCode() {
    var code = String(req.code || "").trim().toUpperCase();
    var o = rs.filter(function (r) { return r.pledgeCode && r.pledgeCode === code && lc(r.sponsorEmail) === lc(req.email); })[0];
    if (!o) fail("No pledge matches that code and email. Unpaid pledges are released after 24 hours.");
    return o;
  }
  if (a === "list") return { students: rs.map(pub), now: db.now() };

  if (a === "apply") {
    var d = req.data || {}, keys = ["name", "email", "umassId", "country", "field", "program", "level", "deadline", "waiver", "statement"];
    keys.forEach(function (k) { if (!String(d[k] || "").trim()) fail("Missing " + k + "."); });
    if (String(d.email).indexOf("@") < 0) fail("Check the email address.");
    if (COUNTRIES.indexOf(String(d.country).trim()) < 0) fail("For now, applications are open to students from African countries.");
    if (wordCount(d.statement) > 150) fail("The statement is over 150 words.");
    var max = rs.reduce(function (m, r) { return Math.max(m, Number(String(r.appNo).split("-")[2]) || 0); }, 4299);
    var o = { id: rnd(8), appNo: "FB-26-" + (max + 1), createdAt: db.now(), status: "open", releases: 0 };
    keys.forEach(function (k) { o[k] = String(d[k]).trim().slice(0, k === "statement" ? 2000 : 200); });
    db.add(o); db.log(o.appNo, "applied", "");
    return { appNo: o.appNo };
  }
  if (a === "pledge") {
    var p = rs.filter(function (r) { return r.id === req.id; })[0];
    if (!p) fail("That applicant is no longer listed.");
    if (p.status !== "open") fail("Someone else has just pledged for this applicant.");
    if (String(req.email || "").indexOf("@") < 0) fail("Add your email so we can send your receipt.");
    var anon = !!req.anon;
    if (!anon && !String(req.name || "").trim()) fail("Add your name, or choose to stay anonymous.");
    if (rs.some(function (r) { return r.status === "pledged" && lc(r.sponsorEmail) === lc(req.email); })) fail("You already have a pledge on hold. Pay or release it first.");
    p.status = "pledged"; p.pledgeCode = "PL-" + rnd(4); p.sponsorName = String(req.name || "").trim().slice(0, 120);
    p.sponsorEmail = String(req.email).trim(); p.sponsorAnon = String(anon); p.pledgedAt = db.now();
    db.save(p); db.log(p.appNo, "pledged", p.pledgeCode);
    return sponsorView(p);
  }
  if (a === "sponsor") return sponsorView(byCode());
  if (a === "pay") {
    var q = byCode();
    if (q.status === "pledged") {
      q.status = "paid"; q.payMethod = String(req.method || "Credit card").slice(0, 40); q.payRef = "SIM-" + rnd(6); q.paidAt = db.now();
      db.save(q); db.log(q.appNo, "paid (simulated)", q.payRef + " " + q.payMethod);
    }
    return sponsorView(q);
  }
  if (a === "release") {
    var r = byCode();
    if (r.status !== "pledged") fail("This pledge is already paid.");
    db.log(r.appNo, "released by sponsor", r.pledgeCode); clearPledge(r); db.save(r);
    return {};
  }
  if (a === "track") {
    var no = String(req.appNo || "").trim().toUpperCase();
    var t = rs.filter(function (x) { return x.appNo === no && lc(x.email) === lc(req.email); })[0];
    if (!t) fail("No application matches that number and email.");
    var v = pub(t); v.appNo = t.appNo; v.payMethod = t.status === "paid" ? t.payMethod : "";
    return v;
  }
  fail("Unknown action.");
}

/* Fictional example applicants. */
function seedRows(now) {
  var H = 3600e3;
  function row(o) { var r = {}; COLS.forEach(function (k) { r[k] = o[k] === undefined ? "" : String(o[k]); }); return r; }
  return [
    row({ id: "amina001", appNo: "FB-26-4127", createdAt: now - 90 * H, name: "Amina Okafor", email: "amina@example.com", umassId: "UMA-118204", country: "Nigeria", field: "Microbiology", program: "Microbiology", level: "PhD", deadline: "2026-12-01", waiver: "Doesn't qualify for a waiver", status: "open", releases: 0,
      statement: "I want to understand how antibiotic resistance genes move between bacteria in hospital wastewater. During my master's at the University of Ibadan I sampled drains from three teaching hospitals in Oyo State and found resistance genes at levels well above what local guidelines assume. Very little of this monitoring happens in West Africa, so treatment choices are often made without local data.\n\nAt UMass Amherst I hope to join a lab that combines field sampling with genomic sequencing, and build a low-cost surveillance method a regional lab could run with basic equipment.\n\nThe application fee is the last step I cannot cover myself this cycle." }),
    row({ id: "wanjiru1", appNo: "FB-26-4131", createdAt: now - 80 * H, name: "Wanjiru Kamau", email: "wanjiru@example.com", umassId: "UMA-118377", country: "Kenya", field: "Environmental conservation", program: "Environmental Conservation", level: "PhD", deadline: "2026-11-20", waiver: "Doesn't qualify for a waiver",
      status: "pledged", pledgeCode: "PL-H7KQ", sponsorName: "Lena Brandt", sponsorEmail: "lena@example.com", sponsorAnon: "false", pledgedAt: now - 17 * H, releases: 1,
      statement: "I study how shifting rainfall in Kenya's Rift Valley changes which crops smallholder farmers can still grow. For two years I worked with a county extension office mapping maize failures against weather station records.\n\nAt UMass I want to model land use and rainfall together so extension officers can give advice a season ahead. I'm drawn to the department's work on working landscapes, where conservation and farming share the same ground.\n\nI grew up on one of these farms, and I want my research to reach the people making planting decisions." }),
    row({ id: "kwame001", appNo: "FB-26-4102", createdAt: now - 120 * H, name: "Kwame Mensah", email: "kwame@example.com", umassId: "UMA-117950", country: "Ghana", field: "Machine learning", program: "Computer Science", level: "PhD", deadline: "2026-12-15", waiver: "Doesn't qualify for a waiver",
      status: "paid", pledgeCode: "PL-R2MX", sponsorName: "", sponsorEmail: "anon@example.com", sponsorAnon: "true", pledgedAt: now - 50 * H, payMethod: "Credit card", payRef: "SIM-7F3K2Q", paidAt: now - 49 * H, releases: 0,
      statement: "Speech recognition works well in English and poorly in Twi, Ewe, Ga and most other Ghanaian languages. I want to build models that learn from small amounts of recorded speech, so a language with a few hundred hours of audio can still get a usable system.\n\nAt my current job in Accra I helped collect a Twi speech dataset with a nonprofit, and saw that good data mattered more than model size. At UMass I'd like to work on self-supervised learning and careful dataset design for low-resource languages, and release open models for health hotlines and farm advisory services." }),
    row({ id: "fatou001", appNo: "FB-26-4140", createdAt: now - 70 * H, name: "Fatou Diallo", email: "fatou@example.com", umassId: "UMA-118512", country: "Senegal", field: "Public health", program: "Community Health Education: Public Health", level: "PhD", deadline: "2026-11-15", waiver: "Waiver request was declined", status: "open", releases: 0,
      statement: "In rural Senegal, many women give birth without a skilled attendant because the nearest clinic is hours away. I worked for two years as a field coordinator for a maternal health program in the Kolda region, tracking which women reached care and which did not.\n\nAt UMass I want to study how travel time, cost and trust in clinics combine to keep women from care, and design interventions the regional health authority can afford.\n\nI'm the first in my family to apply to a PhD." }),
    row({ id: "tendai01", appNo: "FB-26-4144", createdAt: now - 60 * H, name: "Tendai Moyo", email: "tendai@example.com", umassId: "UMA-118640", country: "Zimbabwe", field: "Polymer science", program: "Polymer Science and Engineering", level: "MS/PhD", deadline: "2026-12-05", waiver: "Doesn't qualify for a waiver",
      status: "pledged", pledgeCode: "PL-9TWB", sponsorName: "", sponsorEmail: "kim@example.com", sponsorAnon: "true", pledgedAt: now - 22.5 * H, releases: 0,
      statement: "I want to design polymer binders that make sodium-ion batteries last longer. Sodium is cheap and plentiful, which could make grid storage affordable in countries like Zimbabwe, where power cuts are a daily problem, but the electrodes crack and fade quickly.\n\nIn my undergraduate thesis at the University of Zimbabwe I tested three water-based binders and saw one roughly double the cycle life. UMass Amherst's polymer science program is where I most want to take this further.\n\nAfter the PhD I want to help build battery research capacity back home." }),
  ];
}
