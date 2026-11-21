var SHEET = "Marks";
var USERS = "Users";
var LABEL = {o:"మనకి ఓటు వేస్తారు", p:"మన ఓటు కాదు", d1:"ఊరు మారిపోయారు", d2:"మృతి చెందారు", d3:"బయట ఊరిలో ఉంటున్నారు"};
var HEAD = ["వార్డు", "క్ర.సం.", "పేరు", "ఓటరు ఐడి", "స్థితి", "కోడ్", "సమయం", "పార్టీ", "నమోదు చేసినవారు", "వారి మొబైల్"];
var UHEAD = ["పేరు", "మొబైల్ నంబరు", "మొదటి లాగిన్", "చివరి లాగిన్", "లాగిన్ల సంఖ్య", "చివరి వార్డు"];

function sh_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var s = ss.getSheetByName(SHEET);
  if (!s) {
    s = ss.insertSheet(SHEET);
    s.appendRow(HEAD);
  } else {
    for (var k = 8; k <= HEAD.length; k++) {
      if (!s.getRange(1, k).getValue()) s.getRange(1, k).setValue(HEAD[k - 1]);
    }
  }
  return s;
}

function users_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var s = ss.getSheetByName(USERS);
  if (!s) {
    s = ss.insertSheet(USERS);
    s.appendRow(UHEAD);
    s.getRange("B:B").setNumberFormat("@");
  }
  return s;
}

// షీట్ లో ఫార్ములా లాగా మారకుండా చూసుకుంటుంది
function safe_(x) {
  x = String(x == null ? "" : x);
  return /^[=+\-@]/.test(x) ? "'" + x : x;
}

function doGet(e) {
  var w = String(e.parameter.ward || "");
  var v = sh_().getDataRange().getValues(), m = {}, p = {}, b = {};
  for (var i = 1; i < v.length; i++) {
    if (String(v[i][0]) === w) {
      if (v[i][5]) m[v[i][1]] = v[i][5];
      if (v[i][7]) p[v[i][1]] = v[i][7];
      if (v[i][8] && (v[i][5] || v[i][7])) b[v[i][1]] = v[i][8];
    }
  }
  return ContentService.createTextOutput(JSON.stringify({m: m, p: p, b: b})).setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  var d = JSON.parse(e.postData.contents);
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    if (d.action === "login") {
      login_(d);
    } else {
      mark_(d);
    }
  } finally {
    lock.releaseLock();
  }
  return ContentService.createTextOutput("ok");
}

function login_(d) {
  var mobile = String(d.mobile || "").replace(/\D/g, "").slice(0, 10);
  if (!mobile) return;
  var name = safe_(String(d.name || "").slice(0, 60));
  var s = users_(), v = s.getDataRange().getValues(), r = -1, now = new Date();
  for (var i = 1; i < v.length; i++) {
    if (String(v[i][1]) === mobile) { r = i + 1; break; }
  }
  if (r > 0) {
    var cnt = Number(v[r - 1][4]) || 0;
    s.getRange(r, 1, 1, 6).setValues([[name, mobile, v[r - 1][2], now, cnt + 1, String(d.ward || "")]]);
  } else {
    s.appendRow([name, mobile, now, now, 1, String(d.ward || "")]);
    s.getRange(s.getLastRow(), 2).setNumberFormat("@").setValue(mobile);
  }
}

function mark_(d) {
  var s = sh_(), v = s.getDataRange().getValues(), r = -1;
  for (var i = 1; i < v.length; i++) {
    if (String(v[i][0]) === String(d.ward) && String(v[i][1]) === String(d.sn)) { r = i + 1; break; }
  }
  var code = d.code || "";
  // పార్టీ పంపకపోతే (పాత పేజీ) ఉన్న పార్టీని అలాగే ఉంచుతాం
  var party = (d.party === undefined) ? (r > 0 ? String(v[r - 1][7] || "") : "") : String(d.party || "");
  // ఎవరు నమోదు చేశారు (పాత పేజీ పంపకపోతే ఉన్నదే ఉంచుతాం)
  var by = (d.by === undefined) ? (r > 0 ? String(v[r - 1][8] || "") : "") : safe_(String(d.by || "").slice(0, 60));
  var mob = (d.mobile === undefined) ? (r > 0 ? String(v[r - 1][9] || "") : "") : String(d.mobile || "").replace(/\D/g, "").slice(0, 10);
  if (!code && !party) {
    if (r > 0) s.deleteRow(r);
  } else {
    var row = [String(d.ward), d.sn, d.name, d.vid, LABEL[code] || code, code, new Date(), party, by, mob];
    if (r > 0) s.getRange(r, 1, 1, 10).setValues([row]); else s.appendRow(row);
  }
}
