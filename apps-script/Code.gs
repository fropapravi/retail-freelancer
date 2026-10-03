/**
 * Praveen Kumar B – website database (Google Sheet)
 * Saves project briefs and client ratings from the website,
 * emails you when something new arrives, and serves approved ratings back to the site.
 *
 * Setup: Extensions → Apps Script in your Google Sheet, paste this file,
 * run  setup  once, then Deploy → New deployment → Web app
 * (Execute as: Me · Who has access: Anyone). Copy the Web App URL.
 */

const SHEET_ID = '1ER_FmRBIca1txUI3Oj2CEDyFoFlaknH2kCzZjS8sUEU';
const NOTIFY_EMAIL = 'praveenkumarb114@gmail.com'; // where new briefs/ratings are emailed

function setup() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const make = (name, headers) => {
    const sh = ss.getSheetByName(name) || ss.insertSheet(name);
    if (sh.getLastRow() === 0) {
      sh.appendRow(headers);
      sh.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#1b1b1e').setFontColor('#c9a86a');
      sh.setFrozenRows(1);
    }
    return sh;
  };
  make('Briefs', ['Received', 'Need drawn', 'Units', 'Client has', 'Timeline', 'Project / contact', 'Status', 'Page']);
  const r = make('Ratings', ['Received', 'Stars', 'Name', 'Company / role', 'Comments', 'Approved (Yes/No)']);
  // dropdown for approval
  const rule = SpreadsheetApp.newDataValidation().requireValueInList(['Yes', 'No'], true).build();
  r.getRange('F2:F1000').setDataValidation(rule);
  const b = ss.getSheetByName('Briefs');
  const st = SpreadsheetApp.newDataValidation().requireValueInList(['New', 'Quoted', 'Won', 'Lost'], true).build();
  b.getRange('G2:G1000').setDataValidation(st);
  const def = ss.getSheetByName('Sheet1');
  if (def && def.getLastRow() === 0 && ss.getSheets().length > 2) ss.deleteSheet(def);
}

function clean_(v, max) {
  return String(v == null ? '' : v).replace(/[\r\n]+/g, ' ').replace(/^[=+\-@]/, "'").slice(0, max || 500);
}

function doPost(e) {
  try {
    const d = JSON.parse(e.postData.contents || '{}');
    const ss = SpreadsheetApp.openById(SHEET_ID);
    if (d.type === 'brief') {
      ss.getSheetByName('Briefs').appendRow([new Date(), clean_(d.need, 300), Number(d.units) || '', clean_(d.have, 80),
        clean_(d.timeline, 80), clean_(d.contact, 300), 'New', clean_(d.page, 200)]);
      MailApp.sendEmail(NOTIFY_EMAIL, 'New project brief: ' + clean_(d.contact, 80),
        'Need drawn: ' + d.need + '\nUnits: ' + d.units + '\nClient has: ' + d.have + '\nTimeline: ' + d.timeline + '\nProject / contact: ' + d.contact);
    } else if (d.type === 'rating') {
      const stars = Math.max(1, Math.min(5, Number(d.stars) || 0));
      ss.getSheetByName('Ratings').appendRow([new Date(), stars, clean_(d.name, 100), clean_(d.company, 150), clean_(d.comments, 800), 'No']);
      MailApp.sendEmail(NOTIFY_EMAIL, 'New rating: ' + stars + '/5 from ' + clean_(d.name, 80),
        d.comments + '\n\nOpen your sheet and set "Approved" to Yes to show it on the website.');
    }
    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}

function doGet(e) {
  if (e && e.parameter && e.parameter.action === 'ratings') {
    const sh = SpreadsheetApp.openById(SHEET_ID).getSheetByName('Ratings');
    const rows = sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, 6).getValues() : [];
    const list = rows.filter(r => String(r[5]).toLowerCase() === 'yes').reverse()
      .map(r => ({ stars: r[1], name: r[2], company: r[3], comments: r[4] }));
    return json_(list);
  }
  return json_({ ok: true, service: 'Praveen website database' });
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
