var SHEET = "Marks";
var LABEL = {o:"మన ఓటరు", p:"ప్రత్యర్థి", d1:"ఊరు మారిపోయారు", d2:"మృతి చెందారు", d3:"బయట ఊరిలో ఉంటున్నారు"};

function sh_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var s = ss.getSheetByName(SHEET);
  if (!s) {
    s = ss.insertSheet(SHEET);
    s.appendRow(["వార్డు", "క్ర.సం.", "పేరు", "ఓటరు ఐడి", "స్థితి", "కోడ్", "సమయం"]);
  }
  return s;
}

function doGet(e) {
  var w = String(e.parameter.ward || "");
  var v = sh_().getDataRange().getValues(), m = {};
  for (var i = 1; i < v.length; i++) {
    if (String(v[i][0]) === w) m[v[i][1]] = v[i][5];
  }
  return ContentService.createTextOutput(JSON.stringify(m)).setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  var d = JSON.parse(e.postData.contents);
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var s = sh_(), v = s.getDataRange().getValues(), r = -1;
    for (var i = 1; i < v.length; i++) {
      if (String(v[i][0]) === String(d.ward) && String(v[i][1]) === String(d.sn)) { r = i + 1; break; }
    }
    if (!d.code) {
      if (r > 0) s.deleteRow(r);
    } else {
      var row = [String(d.ward), d.sn, d.name, d.vid, LABEL[d.code] || d.code, d.code, new Date()];
      if (r > 0) s.getRange(r, 1, 1, 7).setValues([row]); else s.appendRow(row);
    }
  } finally {
    lock.releaseLock();
  }
  return ContentService.createTextOutput("ok");
}
