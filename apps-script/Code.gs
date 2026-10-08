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
  const BH = ['Received', 'Need drawn', 'Units', 'Client has', 'Timeline', 'Brand / project', 'Name', 'Email', 'WhatsApp', 'Requirement details', 'Status', 'Page'];
  const bs = make('Briefs', BH);
  bs.getRange(1, 1, 1, BH.length).setValues([BH]).setFontWeight('bold').setBackground('#1b1b1e').setFontColor('#c9a86a');
  bs.setFrozenRows(1);
  const r = make('Ratings', ['Received', 'Stars', 'Name', 'Company / role', 'Comments', 'Approved (Yes/No)']);
  // dropdown for approval
  const rule = SpreadsheetApp.newDataValidation().requireValueInList(['Yes', 'No'], true).build();
  r.getRange('F2:F1000').setDataValidation(rule);
  const b = ss.getSheetByName('Briefs');
  const st = SpreadsheetApp.newDataValidation().requireValueInList(['New', 'Quoted', 'Won', 'Lost'], true).build();
  b.getRange('G2:G1000').clearDataValidations();
  b.getRange('K2:K1000').setDataValidation(st);
  setupVisitors_(ss);
  const def = ss.getSheetByName('Sheet1');
  if (def && def.getLastRow() === 0 && ss.getSheets().length > 2) ss.deleteSheet(def);
}

// ----- Visitors -----
function setupVisitors_(ss) {
  ss = ss || SpreadsheetApp.openById(SHEET_ID);
  const H = ['Time', 'Visitor ID', 'New / Returning', 'Country', 'City (time zone)', 'Device', 'Language', 'Came from', 'Campaign', 'Page'];
  let v = ss.getSheetByName('Visitors');
  if (!v) { v = ss.insertSheet('Visitors'); }
  if (v.getLastRow() === 0) v.appendRow(H);
  v.getRange(1, 1, 1, H.length).setValues([H]).setFontWeight('bold').setBackground('#1b1b1e').setFontColor('#c9a86a');
  v.setFrozenRows(1);
  let s = ss.getSheetByName('Visitor Summary');
  if (!s) s = ss.insertSheet('Visitor Summary');
  s.clear();
  const rows = [
    ['Website visitors', ''],
    ['Total visits', '=COUNTA(Visitors!A2:A)'],
    ['Unique visitors', '=IFERROR(COUNTUNIQUE(Visitors!B2:B),0)'],
    ['Visits today', '=COUNTIFS(Visitors!A2:A,">="&TODAY())'],
    ['Visits last 7 days', '=COUNTIFS(Visitors!A2:A,">="&TODAY()-6)'],
    ['Visits last 30 days', '=COUNTIFS(Visitors!A2:A,">="&TODAY()-29)'],
    ['Returning visits', '=COUNTIF(Visitors!C2:C,"Returning")'],
    ['Mobile visits', '=COUNTIF(Visitors!F2:F,"Mobile")'],
    ['', ''],
    ['Visits by country', '']
  ];
  s.getRange(1, 1, rows.length, 2).setValues(rows);
  s.getRange('A12').setFormula('=IFERROR(QUERY(Visitors!D2:D,"select D, count(D) where D <> \'\' group by D order by count(D) desc label D \'Country\', count(D) \'Visits\'",0),"No visits yet")');
  s.getRange('D10').setValue('Came from');
  s.getRange('D12').setFormula('=IFERROR(QUERY(Visitors!H2:H,"select H, count(H) where H <> \'\' group by H order by count(H) desc label H \'Source\', count(H) \'Visits\'",0),"")');
  s.getRange('A1').setFontSize(14).setFontWeight('bold');
  s.getRange('A2:A8').setFontWeight('bold'); s.getRange('A10:D10').setFontWeight('bold');
  s.getRange('B2:B8').setFontSize(12).setFontWeight('bold').setFontColor('#8a6d2f');
  s.setColumnWidth(1, 180); s.setColumnWidth(4, 160);
  return v;
}

function tzCountry_(tz) {
  const M = { 'Asia/Dubai': 'UAE', 'Asia/Muscat': 'Oman', 'Asia/Qatar': 'Qatar', 'Asia/Riyadh': 'Saudi Arabia', 'Asia/Kuwait': 'Kuwait', 'Asia/Bahrain': 'Bahrain',
    'Asia/Kolkata': 'India', 'Asia/Calcutta': 'India', 'Asia/Karachi': 'Pakistan', 'Asia/Dhaka': 'Bangladesh', 'Asia/Colombo': 'Sri Lanka',
    'Asia/Singapore': 'Singapore', 'Asia/Hong_Kong': 'Hong Kong', 'Asia/Shanghai': 'China', 'Asia/Tokyo': 'Japan', 'Asia/Seoul': 'South Korea',
    'Asia/Bangkok': 'Thailand', 'Asia/Jakarta': 'Indonesia', 'Asia/Kuala_Lumpur': 'Malaysia', 'Asia/Manila': 'Philippines', 'Asia/Taipei': 'Taiwan',
    'Asia/Jerusalem': 'Israel', 'Asia/Tel_Aviv': 'Israel', 'Asia/Beirut': 'Lebanon', 'Asia/Amman': 'Jordan', 'Africa/Cairo': 'Egypt', 'Europe/Istanbul': 'Turkey',
    'Europe/London': 'UK', 'Europe/Dublin': 'Ireland', 'Europe/Paris': 'France', 'Europe/Berlin': 'Germany', 'Europe/Rome': 'Italy', 'Europe/Madrid': 'Spain',
    'Europe/Lisbon': 'Portugal', 'Europe/Amsterdam': 'Netherlands', 'Europe/Brussels': 'Belgium', 'Europe/Zurich': 'Switzerland', 'Europe/Vienna': 'Austria',
    'Europe/Stockholm': 'Sweden', 'Europe/Copenhagen': 'Denmark', 'Europe/Oslo': 'Norway', 'Europe/Helsinki': 'Finland', 'Europe/Warsaw': 'Poland',
    'Europe/Prague': 'Czechia', 'Europe/Athens': 'Greece', 'Europe/Monaco': 'Monaco', 'Europe/Luxembourg': 'Luxembourg', 'Europe/Moscow': 'Russia',
    'America/New_York': 'USA', 'America/Chicago': 'USA', 'America/Denver': 'USA', 'America/Los_Angeles': 'USA', 'America/Phoenix': 'USA',
    'America/Toronto': 'Canada', 'America/Vancouver': 'Canada', 'America/Mexico_City': 'Mexico', 'America/Sao_Paulo': 'Brazil', 'America/Bogota': 'Colombia',
    'America/Argentina/Buenos_Aires': 'Argentina', 'America/Santiago': 'Chile', 'America/Lima': 'Peru',
    'Australia/Sydney': 'Australia', 'Australia/Melbourne': 'Australia', 'Australia/Perth': 'Australia', 'Pacific/Auckland': 'New Zealand',
    'Africa/Johannesburg': 'South Africa', 'Africa/Lagos': 'Nigeria', 'Africa/Nairobi': 'Kenya', 'Africa/Casablanca': 'Morocco', 'Africa/Algiers': 'Algeria' };
  if (M[tz]) return M[tz];
  if (/^America\//.test(tz)) return 'Americas (' + tz.split('/').pop().replace(/_/g, ' ') + ')';
  if (/^Europe\//.test(tz)) return 'Europe (' + tz.split('/').pop().replace(/_/g, ' ') + ')';
  return tz || 'Unknown';
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
        clean_(d.timeline, 80), clean_(d.project, 200), clean_(d.name, 100), clean_(d.email, 150), clean_(d.whatsapp, 40),
        String(d.details == null ? '' : d.details).replace(/^[=+\-@]/, "'").slice(0, 2000), 'New', clean_(d.page, 200)]);
      const body = 'New project request from your website\n\n' +
        'Name: ' + d.name + '\nEmail: ' + d.email + '\nWhatsApp: ' + d.whatsapp + '\nBrand / project: ' + (d.project || '-') +
        '\n\nNeed drawn: ' + d.need + '\nUnits: ' + d.units + '\nClient has: ' + d.have + '\nTimeline: ' + d.timeline +
        '\n\nRequirement details:\n' + (d.details || '-') + '\n\nReply to this email to answer the client directly.';
      const opt = { name: 'Retail Freelancer website' };
      if (/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(String(d.email || ''))) opt.replyTo = String(d.email);
      MailApp.sendEmail(NOTIFY_EMAIL, 'New project request: ' + clean_(d.name, 60) + (d.project ? ' – ' + clean_(d.project, 60) : ''), body, opt);
    } else if (d.type === 'visit') {
      const sh = ss.getSheetByName('Visitors') || setupVisitors_(ss);
      const tz = clean_(d.tz, 60);
      sh.appendRow([new Date(), clean_(d.visitor, 20), d.returning ? 'Returning' : 'New', tzCountry_(tz), tz.split('/').pop().replace(/_/g, ' '),
        clean_(d.device, 10), clean_(d.lang, 15), clean_(d.referrer, 80) || 'Direct', clean_(d.source, 60), clean_(d.page, 200)]);
      return json_({ ok: true });
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
