// Paste in Code.gs. Files in this project: Code.gs, html.html, css.html, js.html
var SHEET_ID = '1QJeht491pxs7k39zw6jAFGujlQKfcEb97BOnEQSr5gA';      // paste the ID of your Google Sheet (the long text in its URL between /d/ and /edit)
var NOTIFY_EMAIL = 'violentamanny08@gmail.com';  // email for new-lead alerts (blank = the account that deploys this)

function doGet(e) {
  var t = HtmlService.createTemplateFromFile('html');
  var s = (e && e.parameter && (e.parameter.src || e.parameter.utm_source)) || 'direct';
  t.src = String(s).replace(/[^a-zA-Z0-9_-]/g, '').slice(0, 30) || 'direct';
  t.booking = typeof BOOKING_URL === 'string' ? BOOKING_URL : '';
  return t.evaluate()
    .setTitle('Join Manny & Annalyn | Home Partners PH')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(name) {
  return HtmlService.createHtmlOutputFromFile(name).getContent();
}

// Called by the page when someone finishes the quiz
function saveLead(d) {
  var sh = SpreadsheetApp.openById(SHEET_ID).getSheets()[0];
  sh.getRange('C:C').setNumberFormat('@');
  if (sh.getLastRow() === 0) sh.appendRow(['Time','Name','Mobile','Email','City','Score','Answers']);
  sh.appendRow([new Date(), d.name, d.phone, d.email, d.city, d.score, (d.answers || []).join(' | ')]);
  sendLeadEmail(d);
  return 'ok';
}
