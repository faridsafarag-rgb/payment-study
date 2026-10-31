/**
 * Payment Decisions Study – Google Sheets receiver (v5)
 * Paste into Extensions > Apps Script of the results spreadsheet, then
 * Deploy > Manage deployments > Edit > Version: New version > Deploy
 * (keeps the same /exec link).
 *
 * Tabs (created / upgraded automatically):
 *   Dashboard             – live rates with vs. without a pre-request message, factors cited
 *   Participant Summary   – one row per participant (updated as they progress)
 *   Request Responses     – one row per participant per request (append-only)
 *   Analysis Dataset      – the same requests, numeric 1/0 coding (append-only)
 * If a tab still has the headers of an older version, it is upgraded when empty,
 * or renamed "<tab> (old)" when it already holds data, so nothing is ever mixed.
 */

const VERSION = 'v5';
const SUMMARY_SHEET  = 'Participant Summary';
const RESPONSE_SHEET = 'Request Responses';
const ANALYSIS_SHEET = 'Analysis Dataset';
const DASHBOARD_SHEET = 'Dashboard';

const SUMMARY_HEADERS = [
  'participant_id','study_version','start_time','end_time','duration_sec','consent','completion_status','device',
  'requests_completed','accept_count','reject_count','n_correct','accuracy_pct','false_acceptances','false_rejections',
  'accepts_with_message','accepts_without_message','false_acceptances_with_message','false_acceptances_without_message',
  'A1_instapay_frequency','A3_transaction_confidence','A4_online_scam_experience','A5_scam_experience',
  'B1_verify_sender','B2_amount_importance','age_range',
  'fraudulent_scenarios','message_scenarios','presentation_order','last_saved'
];

const FACTOR_KEYS = ['f_sender_name','f_notification_sms','f_history','f_note','f_amount','f_urgency'];
const FACTOR_LABELS = ['Name of the sender','Notification/SMS received before the request','History with this account',
  'Note attached to the request','The amount requested','Urgency of the request'];
const FACTOR_VARS = ['Factor_Sender_Name','Factor_Notification_SMS','Factor_History','Factor_Note','Factor_Amount','Factor_Urgency'];

const RESPONSE_HEADERS = [
  'timestamp','participant_id','request_id','response_id','position','scenario_label',
  'ground_truth','decision','correct','error_type','decision_time','response_time_ms','reason_time_ms',
  'decision_question','selected_factors','n_factors'
].concat(FACTOR_KEYS, [
  'display_name','wallet_address','registered_name','history_with_sender','sender_identity_consistent',
  'amount_egp','amount_level','urgency','presentation','note','reason_for_transfer','reference_number','expiry_date',
  'details_collapsed','block_clicked',
  'message_shown','message_type','message_channel','message_sender','message_text','message_view_ms'
]);

const ANALYSIS_HEADERS = [
  'Participant_ID','Request_ID','Response_ID','Position','Age_Range',
  'Decision','Fraudulent','Correct','False_Acceptance','False_Rejection',
  'Sender_Identity_Consistent','Amount_High','Urgency','Presentation_Official',
  'Message_Shown','Message_Type','Amount_EGP','Response_Time_ms','Details_Collapsed','Block_Clicked','N_Factors'
].concat(FACTOR_VARS);

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const data = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (Array.isArray(data.responses) && data.responses.length) appendResponses_(ss, data.responses);
    ensureDashboard_(ss);
    if (data.participant) {
      data.participant.last_saved = new Date().toISOString();
      upsertSummary_(getSheet_(ss, SUMMARY_SHEET, SUMMARY_HEADERS), data.participant);
      if (data.participant.age_range) backfillAge_(ss, data.participant.participant_id, data.participant.age_range);
    }
    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function doGet() {
  return json_({ ok: true, status: 'Payment study receiver is running (' + VERSION + ')' });
}

/* ---------- writers ---------- */

function appendResponses_(ss, rows) {
  const rs = getSheet_(ss, RESPONSE_SHEET, RESPONSE_HEADERS);
  const as = getSheet_(ss, ANALYSIS_SHEET, ANALYSIS_HEADERS);
  const idCol = RESPONSE_HEADERS.indexOf('response_id') + 1;
  const seen = new Set();
  if (rs.getLastRow() > 1) rs.getRange(2, idCol, rs.getLastRow() - 1, 1).getValues().forEach(r => seen.add(r[0]));
  const fresh = rows.filter(r => r.response_id && !seen.has(r.response_id)); // never overwrite or duplicate
  if (!fresh.length) return;
  rs.getRange(rs.getLastRow() + 1, 1, fresh.length, RESPONSE_HEADERS.length)
    .setValues(fresh.map(r => RESPONSE_HEADERS.map(h => clean_(r[h]))));
  as.getRange(as.getLastRow() + 1, 1, fresh.length, ANALYSIS_HEADERS.length)
    .setValues(fresh.map(r => ANALYSIS_HEADERS.map(h => clean_(analysisValue_(r, h)))));
}

function analysisValue_(r, h) {
  const yes = (v, t) => (v === t ? 1 : 0);
  const fi = FACTOR_VARS.indexOf(h);
  if (fi !== -1) return r[FACTOR_KEYS[fi]] === 1 ? 1 : 0;   // 1 = ticked, 0 = not ticked
  switch (h) {
    case 'Participant_ID': return r.participant_id;
    case 'Request_ID': return r.request_id;
    case 'Response_ID': return r.response_id;
    case 'Position': return r.position;
    case 'Age_Range': return r.age_range || '';               // filled in when the participant finishes
    case 'Decision': return yes(r.decision, 'Accept');        // 1 = Accept, 0 = Reject
    case 'Fraudulent': return yes(r.ground_truth, 'Fraudulent');
    case 'Correct': return r.correct;
    case 'False_Acceptance': return yes(r.error_type, 'False acceptance');
    case 'False_Rejection': return yes(r.error_type, 'False rejection');
    case 'Sender_Identity_Consistent': return yes(r.sender_identity_consistent, 'Yes');
    case 'Amount_High': return yes(r.amount_level, 'High');
    case 'Urgency': return yes(r.urgency, 'Yes');
    case 'Presentation_Official': return yes(r.presentation, 'Official');
    case 'Message_Shown': return yes(r.message_shown, 'Yes');
    case 'Message_Type': return r.message_type;
    case 'Amount_EGP': return r.amount_egp;
    case 'Response_Time_ms': return r.response_time_ms;
    case 'Details_Collapsed': return r.details_collapsed ? 1 : 0;
    case 'Block_Clicked': return r.block_clicked ? 1 : 0;
    case 'N_Factors': return r.n_factors;
  }
  return '';
}

function upsertSummary_(sh, row) {
  const values = SUMMARY_HEADERS.map(h => clean_(row[h]));
  const last = sh.getLastRow();
  if (last > 1) {
    const ids = sh.getRange(2, 1, last - 1, 1).getValues().map(r => r[0]);
    const i = ids.indexOf(row.participant_id);
    if (i !== -1) {
      const rank = { 'Started': 0, 'In progress': 1, 'Completed': 2 };   // never move backwards
      const cur = sh.getRange(i + 2, SUMMARY_HEADERS.indexOf('completion_status') + 1).getValue();
      const curDone = Number(sh.getRange(i + 2, SUMMARY_HEADERS.indexOf('requests_completed') + 1).getValue()) || 0;
      if ((rank[row.completion_status] || 0) < (rank[cur] || 0)) return;
      if ((Number(row.requests_completed) || 0) < curDone) return;
      sh.getRange(i + 2, 1, 1, values.length).setValues([values]);
      return;
    }
  }
  sh.appendRow(values);
}

function backfillAge_(ss, pid, age) {
  const as = ss.getSheetByName(ANALYSIS_SHEET);
  if (!as || as.getLastRow() < 2) return;
  const ageCol = ANALYSIS_HEADERS.indexOf('Age_Range') + 1;
  const ids = as.getRange(2, 1, as.getLastRow() - 1, 1).getValues();
  ids.forEach((r, i) => { if (r[0] === pid) as.getRange(i + 2, ageCol, 1, 1).setValues([[age]]); });
}

/* ---------- dashboard (live formulas) ---------- */

function ensureDashboard_(ss) {
  let sh = ss.getSheetByName(DASHBOARD_SHEET);
  if (sh && sh.getRange(1, 6, 1, 1).getValues()[0][0] === VERSION) return;
  getSheet_(ss, SUMMARY_SHEET, SUMMARY_HEADERS);
  getSheet_(ss, ANALYSIS_SHEET, ANALYSIS_HEADERS);
  if (sh) sh.clear(); else sh = ss.insertSheet(DASHBOARD_SHEET);
  const A = "'" + ANALYSIS_SHEET + "'!";
  const letter = n => { let s = ''; while (n) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = (n - m - 1) / 26; } return s; };
  const col = h => { const L = letter(ANALYSIS_HEADERS.indexOf(h) + 1); return A + L + '2:' + L; };
  const scol = h => { const L = letter(SUMMARY_HEADERS.indexOf(h) + 1); return "'" + SUMMARY_SHEET + "'!" + L + '2:' + L; };
  const D = col('Decision'), F = col('Fraudulent'), M = col('Message_Shown'), C = col('Correct'), NS = col('Factor_Notification_SMS');
  const avg = (v, conds) => conds ? '=IFERROR(AVERAGEIFS(' + v + ',' + conds + '),"–")' : '=IFERROR(AVERAGE(' + v + '),"–")';
  const rows = [
    ['Payment study dashboard (updates automatically)', '', '', ''],
    ['', '', '', ''],
    ['Participants completed', '=COUNTIF(' + scol('completion_status') + ',"Completed")', '', ''],
    ['Request responses saved', '=COUNT(' + D + ')', '', ''],
    ['', '', '', ''],
    ['Rates (share of requests)', 'With message', 'Without message', 'All'],
    ['Acceptance rate – all requests', avg(D, M + ',1'), avg(D, M + ',0'), avg(D, '')],
    ['Acceptance of FRAUDULENT requests (false acceptance)', avg(D, M + ',1,' + F + ',1'), avg(D, M + ',0,' + F + ',1'), avg(D, F + ',1')],
    ['Acceptance of GENUINE requests', avg(D, M + ',1,' + F + ',0'), avg(D, M + ',0,' + F + ',0'), avg(D, F + ',0')],
    ['Decision accuracy', avg(C, M + ',1'), avg(C, M + ',0'), avg(C, '')],
    ['Ticked "Notification/SMS received before the request"', avg(NS, M + ',1'), avg(NS, M + ',0'), avg(NS, '')],
    ['', '', '', ''],
    ['What made people decide (share of decisions)', 'When accepted', 'When rejected', 'All']
  ];
  FACTOR_VARS.forEach((v, i) => { const c = col(v); rows.push([FACTOR_LABELS[i], avg(c, D + ',1'), avg(c, D + ',0'), avg(c, '')]); });
  sh.getRange(1, 1, rows.length, 4).setValues(rows);
  sh.getRange(1, 6, 1, 1).setValues([[VERSION]]);
  sh.getRange('A1').setFontWeight('bold').setFontSize(14);
  sh.getRange('A6:D6').setFontWeight('bold'); sh.getRange('A13:D13').setFontWeight('bold');
  sh.getRange('B7:D11').setNumberFormat('0%'); sh.getRange('B14:D19').setNumberFormat('0%');
  sh.setColumnWidth(1, 430); sh.setColumnWidths(2, 3, 130);
}

/* ---------- helpers ---------- */

function getSheet_(ss, name, headers) {
  let sh = ss.getSheetByName(name);
  if (sh && sh.getLastRow() >= 1) {
    const cur = sh.getRange(1, 1, 1, Math.max(headers.length, sh.getLastColumn())).getValues()[0];
    const same = headers.every((h, i) => cur[i] === h) && cur.slice(headers.length).every(v => v === '');
    if (!same) {
      if (sh.getLastRow() === 1) { sh.clear(); }                       // empty tab: just upgrade headers
      else { sh.setName(name + ' (old ' + Utilities.formatDate(new Date(), 'GMT', 'yyyy-MM-dd HHmm') + ')'); sh = null; }
    }
  }
  if (!sh) sh = ss.insertSheet(name);
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}

function clean_(v) {
  if (v === undefined || v === null) return '';
  if (typeof v === 'number' || typeof v === 'boolean') return v;
  let s = String(v).slice(0, 2000);
  if (/^[=+\-@]/.test(s)) s = "'" + s; // stop text being run as a formula
  return s;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
