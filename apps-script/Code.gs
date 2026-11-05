/* ===== PhD Fee Bridge: Google Sheet backend =====
 * Setup (once):
 *   1. In your Google Sheet: Extensions > Apps Script. Paste this whole file into Code.gs and save.
 *   2. Choose the function "setup" in the toolbar and click Run. Approve the permissions.
 *      This creates the tabs and a daily reminder job for committed sponsors.
 *      (Optional) Run "seedDemo" to add the fictional example data.
 *   3. Deploy > New deployment > type "Web app". Execute as: Me. Who has access: Anyone. Deploy.
 *   4. Copy the Web app URL (ends in /exec) into config.js on the website.
 * After you edit this file later: Deploy > Manage deployments > edit > Version: New version.
 * Only you can open the Sheet. The website only receives public fields (see pubApp in the shared rules).
 */
var SITE_URL = "https://kyatha.github.io/Fee_sponsorships/";
var SEND_EMAILS = true;   // set to false to stop all emails (they come from your Gmail; ~100/day limit on free accounts)

function setup() {
  var ss = SpreadsheetApp.getActive();
  Object.keys(TABLES).forEach(function (name) {
    var sh = ss.getSheetByName(name) || ss.insertSheet(name);
    sh.getRange(1, 1, 1, TABLES[name].length).setValues([TABLES[name]]).setFontWeight("bold");
    sh.setFrozenRows(1);
  });
  var lg = ss.getSheetByName("Log") || ss.insertSheet("Log");
  lg.getRange(1, 1, 1, 4).setValues([["time", "appNo", "event", "detail"]]).setFontWeight("bold");
  lg.setFrozenRows(1);
  ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === "dailyReminders") ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger("dailyReminders").timeBased().everyDays(1).atHour(9).create();
}

function seedDemo() {
  var db = sheetDb(), seed = seedData(Date.now());
  Object.keys(seed).forEach(function (name) { seed[name].forEach(function (o) { db.table(name).add(o); }); });
}

/* Runs every morning: releases expired holds and pools, and sends season reminders to committed sponsors. */
function dailyReminders() {
  var db = sheetDb();
  expireHolds(db);
  remind(db);
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
  var ss = SpreadsheetApp.getActive(), cache = {};
  // Everything is stored as plain text: typed text can't become a formula, and dates and IDs stay as typed.
  function safe(x) { var s = x === undefined || x === null ? "" : String(x); return s === "" ? "" : "'" + s; }
  function table(name) {
    if (cache[name]) return cache[name];
    var cols = TABLES[name], sh = ss.getSheetByName(name), v = sh.getDataRange().getDisplayValues(), h = v.shift();
    var list = v.map(function (r, i) { var o = { _row: i + 2 }; h.forEach(function (k, j) { o[k] = r[j]; }); return o; })
      .filter(function (o) { return o[cols[0]]; });
    function rowOf(o) { return cols.map(function (k) { return safe(o[k]); }); }
    return cache[name] = {
      all: function () { return list; },
      save: function (o) { sh.getRange(o._row, 1, 1, cols.length).setValues([rowOf(o)]); },
      add: function (o) { sh.appendRow(rowOf(o)); o._row = sh.getLastRow(); list.push(o); }
    };
  }
  return {
    table: table,
    now: function () { return Date.now(); },
    site: SITE_URL,
    log: function (no, ev, d) { ss.getSheetByName("Log").appendRow([new Date().toISOString(), safe(no), ev, safe(d)]); },
    mail: function (to, subject, body) {
      if (!SEND_EMAILS || String(to).indexOf("@example.com") > 0) return;
      try { MailApp.sendEmail({ to: to, subject: subject, body: body + "\n\nPhD Fee Bridge (independent, not affiliated with UMass Amherst)", name: "PhD Fee Bridge" }); }
      catch (e) { ss.getSheetByName("Log").appendRow([new Date().toISOString(), "", "email failed", safe(to + " " + e)]); }
    }
  };
}

/* ===== Shared rules: the same code runs in Google Apps Script (Sheet mode) and in the browser (preview mode). ===== */
var FEE = 90;                    // UMass Amherst graduate application fee, USD
var MIN_SHARE = 5;               // smallest pool contribution, USD
var HOLD_MS = 24 * 3600e3;       // time to pay once a fee is fully pledged
var POOL_MS = 7 * 24 * 3600e3;   // time for a pool to fill before its pledges are released

var TABLES = {
  Applications: ["id", "appNo", "createdAt", "name", "email", "umassId", "country", "field", "program", "level", "deadline", "waiver", "statement", "status", "paidAt", "releases"],
  Pledges: ["code", "appId", "createdAt", "sponsorName", "sponsorEmail", "sponsorAnon", "amount", "status", "dueAt", "payMethod", "payRef", "paidAt"],
  Sponsors: ["code", "createdAt", "name", "email", "anon", "students", "cycles", "prefField", "prefCountry", "status", "reminded"]
};
/* For now the program is open to applicants from the 54 African countries. */
var COUNTRIES = ["Algeria", "Angola", "Benin", "Botswana", "Burkina Faso", "Burundi", "Cabo Verde", "Cameroon", "Central African Republic", "Chad", "Comoros", "Democratic Republic of the Congo", "Republic of the Congo", "Côte d'Ivoire", "Djibouti", "Egypt", "Equatorial Guinea", "Eritrea", "Eswatini", "Ethiopia", "Gabon", "The Gambia", "Ghana", "Guinea", "Guinea-Bissau", "Kenya", "Lesotho", "Liberia", "Libya", "Madagascar", "Malawi", "Mali", "Mauritania", "Mauritius", "Morocco", "Mozambique", "Namibia", "Niger", "Nigeria", "Rwanda", "São Tomé and Príncipe", "Senegal", "Seychelles", "Sierra Leone", "Somalia", "South Africa", "South Sudan", "Sudan", "Tanzania", "Togo", "Tunisia", "Uganda", "Zambia", "Zimbabwe"];

var LEVELS = ["PhD", "MS/PhD"];
var ALPH = "ABCDEFGHJKLMNPQRSTUVWXYZ23456789";
function rnd(n) { var s = ""; for (var i = 0; i < n; i++) s += ALPH[Math.floor(Math.random() * ALPH.length)]; return s; }
function wordCount(t) { return (String(t || "").trim().match(/\S+/g) || []).length; }
function lc(s) { return String(s || "").trim().toLowerCase(); }
function fail(m) { throw new Error(m); }
function yes(v) { return String(v) === "true"; }
function num(v) { return Number(v) || 0; }
function sum(list, f) { return list.reduce(function (a, x) { return a + f(x); }, 0); }
function money(n) { return "$" + Number(n).toFixed(0); }

/* A cycle is one UMass application season, September to January, named by the fall it leads to.
   Sep 2026 – Jan 2027 is the "Fall 2027" cycle. Between February and August, the next cycle is the upcoming one. */
function cycleOf(t) { var d = new Date(t), m = d.getUTCMonth() + 1, y = d.getUTCFullYear(); return m >= 2 ? y + 1 : y; }
function inSeason(t) { var m = new Date(t).getUTCMonth() + 1; return m >= 9 || m === 1; }
function matchesCycle(s, c) { return s.status === "active" && (s.cycles === "every" || num(s.cycles) === c); }

function pledgesFor(db, appId) { return db.table("Pledges").all().filter(function (p) { return p.appId === appId && p.status !== "released"; }); }
function poolOf(db, app) {
  var ps = pledgesFor(db, app.id);
  var committed = sum(ps, function (p) { return num(p.amount); }), paid = sum(ps.filter(function (p) { return p.status === "paid"; }), function (p) { return num(p.amount); });
  var held = ps.filter(function (p) { return p.status === "held"; }), due = ps.filter(function (p) { return p.status === "due"; });
  var start = held.length ? Math.min.apply(null, held.map(function (p) { return num(p.createdAt); })) : 0;
  return {
    committed: committed, paid: paid, left: Math.max(0, FEE - committed),
    endsAt: start ? start + POOL_MS : 0,
    dueBy: due.length ? Math.min.apply(null, due.map(function (p) { return num(p.dueAt); })) : 0,
    contributors: ps.map(function (p) { return { name: yes(p.sponsorAnon) ? "" : p.sponsorName, amount: num(p.amount), status: p.status }; })
  };
}
function statusOf(pool) { return pool.paid >= FEE ? "paid" : pool.committed >= FEE ? "funding" : pool.committed > 0 ? "pooling" : "open"; }

/* What anyone can see. Emails, UMass applicant IDs, application numbers and codes never leave through here. */
function pubApp(db, a) {
  var pool = poolOf(db, a);
  return { id: a.id, name: a.name, country: a.country, field: a.field, program: a.program, level: a.level, deadline: a.deadline,
    waiver: a.waiver, statement: a.statement, status: statusOf(pool), pool: pool, releases: num(a.releases) };
}
function pledgeView(db, p) {
  var a = db.table("Applications").all().filter(function (x) { return x.id === p.appId; })[0];
  var v = { code: p.code, amount: num(p.amount), status: p.status, createdAt: num(p.createdAt), dueAt: num(p.dueAt),
    payMethod: p.payMethod, payRef: p.payRef, anon: yes(p.sponsorAnon), app: a ? pubApp(db, a) : null };
  if (a && (p.status === "due" || p.status === "paid")) { v.appNo = a.appNo; v.umassId = a.umassId; }
  return v;
}
function usedThisCycle(db, email, c) {
  return sum(db.table("Pledges").all().filter(function (p) { return lc(p.sponsorEmail) === lc(email) && p.status !== "released" && cycleOf(num(p.createdAt)) === c; }),
    function (p) { return num(p.amount); }) / FEE;
}
function commitmentOf(db, email, c) {
  return db.table("Sponsors").all().filter(function (s) { return lc(s.email) === lc(email) && matchesCycle(s, c); })[0];
}
function sponsorView(db, s) {
  var c = cycleOf(db.now());
  return { code: s.code, name: s.name, anon: yes(s.anon), students: num(s.students), cycles: s.cycles, prefField: s.prefField, prefCountry: s.prefCountry,
    status: s.status, cycle: c, used: Math.round(usedThisCycle(db, s.email, c) * 10) / 10, countsThisCycle: matchesCycle(s, c) };
}

function syncApp(db, a) {
  var st = statusOf(poolOf(db, a));
  if (a.status !== st) { a.status = st; if (st === "paid") a.paidAt = db.now(); db.table("Applications").save(a); }
}
function releasePledge(db, p, why) {
  var a = db.table("Applications").all().filter(function (x) { return x.id === p.appId; })[0];
  db.log(a ? a.appNo : "", why, p.code + " " + money(p.amount));
  p.status = "released"; db.table("Pledges").save(p);
  if (a) { a.releases = num(a.releases) + 1; db.table("Applications").save(a); syncApp(db, a); }
}
function expireHolds(db) {
  var t = db.now(), P = db.table("Pledges");
  P.all().forEach(function (p) { if (p.status === "due" && t >= num(p.dueAt)) releasePledge(db, p, "released: not paid within 24 hours"); });
  db.table("Applications").all().forEach(function (a) {
    var pool = poolOf(db, a), closeToDeadline = new Date(a.deadline + "T23:59:59Z").getTime() - t < 2 * 864e5;
    if (pool.endsAt && (t >= pool.endsAt || closeToDeadline)) {
      pledgesFor(db, a.id).filter(function (p) { return p.status === "held"; }).forEach(function (p) {
        releasePledge(db, p, "released: pool didn't fill in time");
        db.mail(p.sponsorEmail, "Your pledge for " + a.name + " was released",
          "The pool for " + a.name + "'s UMass Amherst application fee didn't fill in time, so your " + money(p.amount) + " pledge (" + p.code + ") was released. Nothing was charged.");
      });
    }
  });
}

function handle(req, db) {
  expireHolds(db);
  var A = db.table("Applications"), P = db.table("Pledges"), S = db.table("Sponsors"), t = db.now(), act = req.action;
  function app(id) { var a = A.all().filter(function (x) { return x.id === id; })[0]; if (!a) fail("That applicant is no longer listed."); return a; }
  function pledgeByCode() {
    var code = String(req.code || "").trim().toUpperCase();
    var p = P.all().filter(function (x) { return x.code === code && lc(x.sponsorEmail) === lc(req.email); })[0];
    if (!p) fail("No pledge matches that code and email.");
    return p;
  }
  function sponsorByCode() {
    var code = String(req.code || "").trim().toUpperCase();
    var s = S.all().filter(function (x) { return x.code === code && lc(x.email) === lc(req.email); })[0];
    if (!s) fail("No sponsor sign-up matches that code and email.");
    return s;
  }

  if (act === "list") {
    var c = cycleOf(t);
    var promised = sum(S.all().filter(function (s) { return matchesCycle(s, c); }), function (s) { return Math.max(0, num(s.students) - usedThisCycle(db, s.email, c)); });
    return { students: A.all().map(function (a) { return pubApp(db, a); }), promised: Math.floor(promised), cycle: c, inSeason: inSeason(t), now: t };
  }

  if (act === "apply") {
    var d = req.data || {}, keys = ["name", "email", "umassId", "country", "field", "program", "level", "deadline", "waiver", "statement"];
    keys.forEach(function (k) { if (!String(d[k] || "").trim()) fail("Missing " + k + "."); });
    if (String(d.email).indexOf("@") < 0) fail("Check the email address.");
    if (LEVELS.indexOf(String(d.level).trim()) < 0) fail("Choose PhD or MS/PhD.");
    if (COUNTRIES.indexOf(String(d.country).trim()) < 0) fail("For now, applications are open to students from African countries.");
    if (wordCount(d.statement) > 150) fail("The statement is over 150 words.");
    var max = A.all().reduce(function (m, r) { return Math.max(m, num(String(r.appNo).split("-")[2])); }, 4299);
    var o = { id: rnd(8), appNo: "FB-26-" + (max + 1), createdAt: t, status: "open", releases: 0 };
    keys.forEach(function (k) { o[k] = String(d[k]).trim().slice(0, k === "statement" ? 2000 : 200); });
    A.add(o); db.log(o.appNo, "applied", o.country + ", " + o.program);
    db.mail(o.email, "Your PhD Fee Bridge application number: " + o.appNo,
      "Hi " + o.name.split(" ")[0] + ",\n\nYour request is on the sponsor board. Your private application number is " + o.appNo + ". Use it with this email address to track your application.\n\n" + db.site + "#track");
    S.all().filter(function (s) { return matchesCycle(s, cycleOf(t)) && ((s.prefField && lc(o.field).indexOf(lc(s.prefField)) >= 0) || (s.prefCountry && s.prefCountry === o.country)); })
      .forEach(function (s) { db.mail(s.email, "A new applicant matches what you asked for", o.name + " from " + o.country + " (" + o.level + ", " + o.program + ") is looking for help with their $" + FEE + " UMass Amherst application fee.\n\n" + db.site); });
    return { appNo: o.appNo };
  }

  if (act === "pledge") {
    var a = app(req.id), pool = poolOf(db, a);
    if (statusOf(pool) === "paid" || statusOf(pool) === "funding") fail("This fee is already fully pledged.");
    var amount = Math.round(num(req.amount));
    if (amount < MIN_SHARE && amount !== pool.left) fail("The smallest amount you can chip in is $" + MIN_SHARE + ".");
    if (amount > pool.left) fail("Only " + money(pool.left) + " is left on this fee.");
    if (String(req.email || "").indexOf("@") < 0) fail("Add your email so we can send your pledge code.");
    var anon = !!req.anon;
    if (!anon && !String(req.name || "").trim()) fail("Add your name, or choose to stay anonymous.");
    var mine = P.all().filter(function (p) { return lc(p.sponsorEmail) === lc(req.email) && (p.status === "held" || p.status === "due"); });
    if (mine.some(function (p) { return p.appId === a.id; })) fail("You already have a pledge on this applicant.");
    var commitment = commitmentOf(db, req.email, cycleOf(t)), limit = commitment ? Math.max(1, num(commitment.students)) : 1;
    if (mine.length >= limit) fail(limit === 1 ? "You already have a pledge open. Pay or release it first, or sign up as a committed sponsor to hold more." : "You already have " + limit + " pledges open, which is your limit this cycle.");
    var p = { code: "PL-" + rnd(4), appId: a.id, createdAt: t, sponsorName: String(req.name || "").trim().slice(0, 120), sponsorEmail: String(req.email).trim(),
      sponsorAnon: String(anon), amount: amount, status: "held", dueAt: "", payMethod: "", payRef: "", paidAt: "" };
    P.add(p); db.log(a.appNo, "pledged", p.code + " " + money(amount));
    db.mail(p.sponsorEmail, "Your pledge code: " + p.code, "Thank you for pledging " + money(amount) + " toward " + a.name + "'s UMass Amherst application fee.\n\nYour pledge code is " + p.code + ". Use it with this email address on the sponsor board to pay or release it.\n\n" + db.site);
    if (pool.committed + amount >= FEE) {
      pledgesFor(db, a.id).filter(function (x) { return x.status === "held"; }).forEach(function (x) {
        x.status = "due"; x.dueAt = t + HOLD_MS; P.save(x);
        db.mail(x.sponsorEmail, "Time to pay: " + a.name + "'s fee is fully pledged", "The " + money(FEE) + " fee for " + a.name + " is now fully pledged. Please pay your " + money(x.amount) + " share within 24 hours with pledge code " + x.code + ".\n\n" + db.site);
      });
      db.log(a.appNo, "fully pledged", "");
    } else {
      db.mail(a.email, "Someone chipped in toward your fee", "A sponsor pledged " + money(amount) + " toward your UMass Amherst application fee. " + money(FEE - pool.committed - amount) + " is still needed.");
    }
    syncApp(db, a);
    return pledgeView(db, p);
  }
  if (act === "pledgeLookup") return pledgeView(db, pledgeByCode());
  if (act === "pay") {
    var q = pledgeByCode();
    if (q.status === "held") fail("This pool isn't full yet. You'll be able to pay once it reaches " + money(FEE) + ".");
    if (q.status === "released") fail("This pledge was released, so there's nothing to pay.");
    if (q.status === "due") {
      q.status = "paid"; q.payMethod = String(req.method || "Credit card").slice(0, 40); q.payRef = "SIM-" + rnd(6); q.paidAt = t; P.save(q);
      var qa = app(q.appId); db.log(qa.appNo, "share paid (simulated)", q.code + " " + money(q.amount) + " " + q.payMethod); syncApp(db, qa);
      if (qa.status === "paid") db.mail(qa.email, "Your UMass Amherst application fee is covered", "Good news: sponsors have covered your " + money(FEE) + " application fee. Application " + qa.appNo + ".");
    }
    return pledgeView(db, q);
  }
  if (act === "release") {
    var r = pledgeByCode();
    if (r.status === "paid") fail("This pledge is already paid.");
    if (r.status === "released") return {};
    releasePledge(db, r, "released by sponsor");
    return {};
  }
  if (act === "track") {
    var no = String(req.appNo || "").trim().toUpperCase();
    var ta = A.all().filter(function (x) { return x.appNo === no && lc(x.email) === lc(req.email); })[0];
    if (!ta) fail("No application matches that number and email.");
    var tv = pubApp(db, ta); tv.appNo = ta.appNo; return tv;
  }

  if (act === "commit") {
    var c2 = cycleOf(t), n = Math.round(num(req.students));
    if (n < 1 || n > 50) fail("Choose between 1 and 50 students.");
    if (["every", String(c2), String(c2 + 1)].indexOf(String(req.cycles)) < 0) fail("Choose a cycle.");
    if (String(req.email || "").indexOf("@") < 0) fail("Add your email so we can send reminders.");
    if (!req.anon && !String(req.name || "").trim()) fail("Add your name, or choose to stay anonymous.");
    if (S.all().some(function (s) { return lc(s.email) === lc(req.email) && s.status === "active"; })) fail("You're already signed up with this email. Use your sponsor code to make changes.");
    var s = { code: "SP-" + rnd(4), createdAt: t, name: String(req.name || "").trim().slice(0, 120), email: String(req.email).trim(), anon: String(!!req.anon),
      students: n, cycles: String(req.cycles), prefField: String(req.prefField || "").trim().slice(0, 80), prefCountry: COUNTRIES.indexOf(req.prefCountry) >= 0 ? req.prefCountry : "", status: "active", reminded: "" };
    S.add(s); db.log("", "sponsor signed up", s.code + " " + n + " students, cycle " + s.cycles);
    db.mail(s.email, "You're a committed sponsor: " + s.code, "Thank you for committing to cover " + n + " student" + (n > 1 ? "s'" : "'s") + " application fee" + (n > 1 ? "s" : "") + ". We'll remind you when the application season opens (September) and again in December.\n\nYour sponsor code is " + s.code + ". Use it with this email to change or stop your commitment at " + db.site + "#sponsor");
    return sponsorView(db, s);
  }
  if (act === "commitLookup") return sponsorView(db, sponsorByCode());
  if (act === "commitUpdate") {
    var u = sponsorByCode(), c3 = cycleOf(t);
    if (req.stop) { u.status = "stopped"; S.save(u); db.log("", "sponsor stopped", u.code); return sponsorView(db, u); }
    var n2 = Math.round(num(req.students));
    if (n2 < 1 || n2 > 50) fail("Choose between 1 and 50 students.");
    if (["every", String(c3), String(c3 + 1)].indexOf(String(req.cycles)) < 0) fail("Choose a cycle.");
    u.students = n2; u.cycles = String(req.cycles); u.status = "active"; S.save(u); db.log("", "sponsor updated", u.code + " " + n2 + " students, cycle " + u.cycles);
    return sponsorView(db, u);
  }
  fail("Unknown action.");
}

/* Reminder emails to committed sponsors. Runs daily in Sheet mode. */
function remind(db) {
  var t = db.now(), c = cycleOf(t), m = new Date(t).getUTCMonth() + 1, S = db.table("Sponsors");
  if (!inSeason(t)) return 0;
  var open = db.table("Applications").all().filter(function (a) { var st = pubApp(db, a).status; return st === "open" || st === "pooling"; }).length, sent = 0;
  S.all().filter(function (s) { return matchesCycle(s, c); }).forEach(function (s) {
    var tags = String(s.reminded || "").split(",").filter(String), left = Math.max(0, num(s.students) - usedThisCycle(db, s.email, c));
    function send(tag, subject, body) { if (tags.indexOf(tag) >= 0 || left <= 0) return; db.mail(s.email, subject, body + "\n\nYour sponsor code is " + s.code + " if you want to change or stop your commitment: " + db.site + "#sponsor"); tags.push(tag); sent++; }
    send("open-" + c, "Fall " + c + " applications are open", "The UMass Amherst application season is here. You committed to cover " + num(s.students) + " student" + (num(s.students) > 1 ? "s" : "") + "' fees this cycle, and " + open + " applicants are waiting.\n\n" + db.site);
    if (m === 12 || m === 1) send("dec-" + c, "Many UMass deadlines are in December", "Most PhD deadlines fall in December and early January. " + open + " applicants still need help, and you have " + Math.ceil(left) + " of your " + num(s.students) + " left this cycle.\n\n" + db.site);
    s.reminded = tags.join(","); S.save(s);
  });
  return sent;
}

/* Fictional example data. */
function seedData(now) {
  var H = 3600e3, c = cycleOf(now);
  function fill(cols, o) { var r = {}; cols.forEach(function (k) { r[k] = o[k] === undefined ? "" : String(o[k]); }); return r; }
  var apps = [
    { id: "amina001", appNo: "FB-26-4127", createdAt: now - 90 * H, name: "Amina Okafor", email: "amina@example.com", umassId: "UMA-118204", country: "Nigeria", field: "Microbiology", program: "Microbiology", level: "PhD", deadline: "2026-12-01", waiver: "Doesn't qualify for a waiver", status: "open", releases: 0,
      statement: "I want to understand how antibiotic resistance genes move between bacteria in hospital wastewater. During my master's at the University of Ibadan I sampled drains from three teaching hospitals in Oyo State and found resistance genes at levels well above what local guidelines assume. Very little of this monitoring happens in West Africa, so treatment choices are often made without local data.\n\nAt UMass Amherst I hope to join a lab that combines field sampling with genomic sequencing, and build a low-cost surveillance method a regional lab could run with basic equipment.\n\nThe application fee is the last step I cannot cover myself this cycle." },
    { id: "wanjiru1", appNo: "FB-26-4131", createdAt: now - 80 * H, name: "Wanjiru Kamau", email: "wanjiru@example.com", umassId: "UMA-118377", country: "Kenya", field: "Environmental conservation", program: "Environmental Conservation", level: "PhD", deadline: "2026-11-20", waiver: "Doesn't qualify for a waiver", status: "funding", releases: 1,
      statement: "I study how shifting rainfall in Kenya's Rift Valley changes which crops smallholder farmers can still grow. For two years I worked with a county extension office mapping maize failures against weather station records.\n\nAt UMass I want to model land use and rainfall together so extension officers can give advice a season ahead. I'm drawn to the department's work on working landscapes, where conservation and farming share the same ground.\n\nI grew up on one of these farms, and I want my research to reach the people making planting decisions." },
    { id: "kwame001", appNo: "FB-26-4102", createdAt: now - 120 * H, name: "Kwame Mensah", email: "kwame@example.com", umassId: "UMA-117950", country: "Ghana", field: "Machine learning", program: "Computer Science", level: "PhD", deadline: "2026-12-15", waiver: "Doesn't qualify for a waiver", status: "paid", paidAt: now - 49 * H, releases: 0,
      statement: "Speech recognition works well in English and poorly in Twi, Ewe, Ga and most other Ghanaian languages. I want to build models that learn from small amounts of recorded speech, so a language with a few hundred hours of audio can still get a usable system.\n\nAt my current job in Accra I helped collect a Twi speech dataset with a nonprofit, and saw that good data mattered more than model size. At UMass I'd like to work on self-supervised learning and careful dataset design for low-resource languages, and release open models for health hotlines and farm advisory services." },
    { id: "fatou001", appNo: "FB-26-4140", createdAt: now - 70 * H, name: "Fatou Diallo", email: "fatou@example.com", umassId: "UMA-118512", country: "Senegal", field: "Public health", program: "Community Health Education: Public Health", level: "PhD", deadline: "2026-12-10", waiver: "Waiver request was declined", status: "pooling", releases: 0,
      statement: "In rural Senegal, many women give birth without a skilled attendant because the nearest clinic is hours away. I worked for two years as a field coordinator for a maternal health program in the Kolda region, tracking which women reached care and which did not.\n\nAt UMass I want to study how travel time, cost and trust in clinics combine to keep women from care, and design interventions the regional health authority can afford.\n\nI'm the first in my family to apply to a PhD." },
    { id: "tendai01", appNo: "FB-26-4144", createdAt: now - 60 * H, name: "Tendai Moyo", email: "tendai@example.com", umassId: "UMA-118640", country: "Zimbabwe", field: "Polymer science", program: "Polymer Science and Engineering", level: "MS/PhD", deadline: "2026-12-05", waiver: "Doesn't qualify for a waiver", status: "funding", releases: 0,
      statement: "I want to design polymer binders that make sodium-ion batteries last longer. Sodium is cheap and plentiful, which could make grid storage affordable in countries like Zimbabwe, where power cuts are a daily problem, but the electrodes crack and fade quickly.\n\nIn my undergraduate thesis at the University of Zimbabwe I tested three water-based binders and saw one roughly double the cycle life. UMass Amherst's polymer science program is where I most want to take this further.\n\nAfter the PhD I want to help build battery research capacity back home." },
    { id: "yonas001", appNo: "FB-26-4151", createdAt: now - 30 * H, name: "Yonas Bekele", email: "yonas@example.com", umassId: "UMA-118802", country: "Ethiopia", field: "Applied mathematics", program: "Mathematics", level: "PhD", deadline: "2026-12-12", waiver: "Doesn't qualify for a waiver", status: "open", releases: 0,
      statement: "I want to use mathematical models to plan where new health posts should go in rural Ethiopia, so that more people live within a two-hour walk of basic care.\n\nAs a lecturer at Bahir Dar University I built a small optimisation model for one zone, using road and population data, and the regional health bureau used it to argue for two new sites.\n\nAt UMass I hope to deepen my training in optimisation and probability, and to work on methods that handle missing and uncertain data, which is the normal situation where I work." },
    { id: "nadia001", appNo: "FB-26-4155", createdAt: now - 20 * H, name: "Nadia Benali", email: "nadia@example.com", umassId: "UMA-118870", country: "Morocco", field: "Linguistics", program: "Linguistics", level: "PhD", deadline: "2027-01-02", waiver: "Waiver request was declined", status: "pooling", releases: 0,
      statement: "Moroccan Darija and Tamazight are spoken by millions but are rarely written down or studied formally. I want to document how speakers switch between Darija, Tamazight and French in everyday conversation, and what that reveals about the grammar of each.\n\nFor my master's at Mohammed V University I recorded forty hours of conversation in Rabat markets and transcribed half of it. UMass Amherst's linguistics department is known for formal theory grounded in fieldwork, which is exactly the combination I want." }
  ].map(function (o) { return fill(TABLES.Applications, o); });
  var pledges = [
    { code: "PL-H7KQ", appId: "wanjiru1", createdAt: now - 17 * H, sponsorName: "Lena Brandt", sponsorEmail: "lena@example.com", sponsorAnon: "false", amount: 90, status: "due", dueAt: now + 7 * H },
    { code: "PL-R2MX", appId: "kwame001", createdAt: now - 50 * H, sponsorName: "", sponsorEmail: "anon@example.com", sponsorAnon: "true", amount: 90, status: "paid", dueAt: now - 26 * H, payMethod: "Credit card", payRef: "SIM-7F3K2Q", paidAt: now - 49 * H },
    { code: "PL-4MZT", appId: "fatou001", createdAt: now - 30 * H, sponsorName: "Grace Achieng", sponsorEmail: "grace@example.com", sponsorAnon: "false", amount: 50, status: "held" },
    { code: "PL-8XQD", appId: "fatou001", createdAt: now - 5 * H, sponsorName: "", sponsorEmail: "sam@example.com", sponsorAnon: "true", amount: 10, status: "held" },
    { code: "PL-9TWB", appId: "tendai01", createdAt: now - 40 * H, sponsorName: "", sponsorEmail: "kim@example.com", sponsorAnon: "true", amount: 40, status: "paid", dueAt: now + 1.5 * H, payMethod: "Venmo", payRef: "SIM-Q8D2LM", paidAt: now - 10 * H },
    { code: "PL-K3VR", appId: "tendai01", createdAt: now - 22.5 * H, sponsorName: "David Okello", sponsorEmail: "david@example.com", sponsorAnon: "false", amount: 50, status: "due", dueAt: now + 1.5 * H },
    { code: "PL-N5CE", appId: "nadia001", createdAt: now - 12 * H, sponsorName: "", sponsorEmail: "pat@example.com", sponsorAnon: "true", amount: 30, status: "held" }
  ].map(function (o) { return fill(TABLES.Pledges, o); });
  var sponsors = [
    { code: "SP-L3NA", createdAt: now - 200 * H, name: "Lena Brandt", email: "lena@example.com", anon: "false", students: 3, cycles: String(c), status: "active" },
    { code: "SP-W8OB", createdAt: now - 300 * H, name: "Kofi Owusu", email: "kofi@example.com", anon: "false", students: 2, cycles: "every", status: "active" },
    { code: "SP-T4XA", createdAt: now - 100 * H, name: "", email: "quiet@example.com", anon: "true", students: 1, cycles: String(c + 1), status: "active" }
  ].map(function (o) { return fill(TABLES.Sponsors, o); });
  return { Applications: apps, Pledges: pledges, Sponsors: sponsors };
}
