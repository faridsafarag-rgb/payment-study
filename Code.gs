/**
 * Payment Decisions Study – Google Sheets receiver (v4)
 * Paste into Extensions > Apps Script of the results spreadsheet, then
 * Deploy > Manage deployments > Edit > Version: New version > Deploy
 * (keeps the same /exec link).
 *
 * Three tabs are created automatically:
 *   Participant Summary   – one row per participant (updated as they progress)
 *   Request Responses     – one row per participant per request (append-only)
 *   Analysis Dataset      – the same requests, numeric 1/0 coding (append-only)
 *   Dashboard             – live acceptance rates with vs. without a pre-request message
 */

const SUMMARY_SHEET  = 'Participant Summary';
const RESPONSE_SHEET = 'Request Responses';
const ANALYSIS_SHEET = 'Analysis Dataset';

const SUMMARY_HEADERS = [
  'participant_id','study_version','start_time','end_time','duration_sec','consent','completion_status','device',
  'age_range',
  'A1_usage_frequency','A2_usage_duration','A3_transaction_confidence','A4_suspicious_requests_seen','A5_scam_experience',
  'B1_verify_sender','B2_amount_importance','B3_urgency_influence','B4_inspect_details','B5_detection_confidence',
  'requests_completed','accept_count','reject_count','n_correct','accuracy_pct','false_acceptances','false_rejections',
  'accepts_with_message','accepts_without_message','false_acceptances_with_message','false_acceptances_without_message',
  'fraudulent_scenarios','message_scenarios','presentation_order','last_saved'
];

const ACCEPT_KEYS = ['acc_recognized_sender','acc_amount_reasonable','acc_notification','acc_note_legitimate',
  'acc_trusted_sender','acc_seemed_urgent','acc_familiar_type','acc_nothing_suspicious','acc_presentation'];
const REJECT_KEYS = ['rej_unrecognized_sender','rej_amount_suspicious','rej_notification','rej_note_suspicious',
  'rej_distrust_sender','rej_seemed_unusual','rej_timing','rej_uncomfortable','rej_urgency','rej_presentation'];
const FACTORS = ['Notification','Sender_Name','Amount','Note','Relationship','Timing','Urgency','Presentation',
  'Familiarity','No_Red_Flags','Unusual','Discomfort'];

const RESPONSE_HEADERS = [
  'timestamp','participant_id','request_id','response_id','position','scenario_label',
  'ground_truth','decision','correct','error_type','decision_time','response_time_ms','reason_time_ms',
  'followup_question','selected_reasons','n_reasons',
  'display_name','wallet_address','registered_name','history_with_sender','sender_identity_consistent',
  'amount_egp','amount_level','urgency','presentation','notification_condition','note','expiry',
  'message_shown','message_type','message_channel','message_sender','message_text','message_view_ms','message_influence'
].concat(ACCEPT_KEYS, REJECT_KEYS);

const ANALYSIS_HEADERS = [
  'Participant_ID','Request_ID','Response_ID','Position','Age_Range',
  'Decision','Fraudulent','Correct','False_Acceptance','False_Rejection',
  'Sender_Identity_Consistent','Amount_High','Urgency','Presentation_Official','Notification_InApp',
  'Amount_EGP','Response_Time_ms','N_Reasons',
  'Message_Shown','Message_Type','Message_Influence_Code',
  'Msg_More_Likely_Accept','Msg_More_Likely_Reject','Msg_No_Influence','Msg_Not_Sure'
].concat(FACTORS.map(f => 'Reason_' + f));

const DASHBOARD_SHEET = 'Dashboard';
const MSG_OPTIONS = ['Yes, it made me more likely to accept.','Yes, it made me more likely to reject.',
                     'No, it did not influence my decision.','I am not sure.'];

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const data = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (Array.isArray(data.responses) && data.responses.length) {
      appendResponses_(ss, data.responses);
    }
    ensureDashboard_(ss);
    if (data.participant) {
      data.participant.last_saved = new Date().toISOString();
      upsertSummary_(getSheet_(ss, SUMMARY_SHEET, SUMMARY_HEADERS), data.participant);
    }
    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function doGet() {
  return json_({ ok: true, status: 'Payment study receiver is running (v4)' });
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
  switch (h) {
    case 'Participant_ID': return r.participant_id;
    case 'Request_ID': return r.request_id;
    case 'Response_ID': return r.response_id;
    case 'Position': return r.position;
    case 'Age_Range': return r.age_range;
    case 'Decision': return yes(r.decision, 'Accept');            // 1 = Accept, 0 = Reject
    case 'Fraudulent': return yes(r.ground_truth, 'Fraudulent');  // 1 = Fraudulent, 0 = Genuine
    case 'Correct': return r.correct;
    case 'False_Acceptance': return yes(r.error_type, 'False acceptance');
    case 'False_Rejection': return yes(r.error_type, 'False rejection');
    case 'Sender_Identity_Consistent': return yes(r.sender_identity_consistent, 'Yes');
    case 'Amount_High': return yes(r.amount_level, 'High');
    case 'Urgency': return yes(r.urgency, 'Yes');
    case 'Presentation_Official': return yes(r.presentation, 'Official');
    case 'Notification_InApp': return 1;                          // held constant in this design
    case 'Amount_EGP': return r.amount_egp;
    case 'Response_Time_ms': return r.response_time_ms;
    case 'N_Reasons': return r.n_reasons;
    case 'Message_Shown': return yes(r.message_shown, 'Yes');     // 1 = message before request, 0 = none
    case 'Message_Type': return r.message_type;
    case 'Message_Influence_Code': {                              // 1 accept, 2 reject, 3 none, 4 not sure, blank = no message
      const i = MSG_OPTIONS.indexOf(r.message_influence); return i === -1 ? '' : i + 1; }
    case 'Msg_More_Likely_Accept': return r.message_shown === 'Yes' ? yes(r.message_influence, MSG_OPTIONS[0]) : '';
    case 'Msg_More_Likely_Reject': return r.message_shown === 'Yes' ? yes(r.message_influence, MSG_OPTIONS[1]) : '';
    case 'Msg_No_Influence':       return r.message_shown === 'Yes' ? yes(r.message_influence, MSG_OPTIONS[2]) : '';
    case 'Msg_Not_Sure':           return r.message_shown === 'Yes' ? yes(r.message_influence, MSG_OPTIONS[3]) : '';
    default: return r[h] === 1 ? 1 : 0;                           // Reason_* columns
  }
}

function upsertSummary_(sh, row) {
  const values = SUMMARY_HEADERS.map(h => clean_(row[h]));
  const last = sh.getLastRow();
  if (last > 1) {
    const ids = sh.getRange(2, 1, last - 1, 1).getValues().map(r => r[0]);
    const i = ids.indexOf(row.participant_id);
    if (i !== -1) {
      // never move a participant backwards (e.g. a late retry of an older save)
      const rank = { 'Started': 0, 'In progress': 1, 'Completed': 2 };
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

/* ---------- dashboard (live formulas, built once) ---------- */

function ensureDashboard_(ss) {
  if (ss.getSheetByName(DASHBOARD_SHEET)) return;
  getSheet_(ss, SUMMARY_SHEET, SUMMARY_HEADERS);
  getSheet_(ss, ANALYSIS_SHEET, ANALYSIS_HEADERS);
  const sh = ss.insertSheet(DASHBOARD_SHEET);
  const A = "'" + ANALYSIS_SHEET + "'!";
  const col = h => { let n = ANALYSIS_HEADERS.indexOf(h) + 1, s = ''; while (n) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = (n - m - 1) / 26; } return A + s + '2:' + s; };
  const D = col('Decision'), F = col('Fraudulent'), M = col('Message_Shown'), T = col('Reason_Relationship'),
        FA = col('False_Acceptance'), C = col('Correct'), INF = col('Message_Influence_Code');
  const avg = (v, conds) => conds ? '=IFERROR(AVERAGEIFS(' + v + ',' + conds + '),"–")' : '=IFERROR(AVERAGE(' + v + '),"–")';
  const rows = [
    ['Payment study dashboard (updates automatically)', '', '', ''],
    ['', '', '', ''],
    ['Participants completed', "=COUNTIF('" + SUMMARY_SHEET + "'!G2:G,\"Completed\")", '', ''],
    ['Request responses saved', '=COUNT(' + D + ')', '', ''],
    ['', '', '', ''],
    ['Rates (share of requests)', 'With message', 'Without message', 'All'],
    ['Acceptance rate – all requests', avg(D, M + ',1'), avg(D, M + ',0'), avg(D, '')],
    ['Acceptance of FRAUDULENT requests (false acceptance)', avg(D, M + ',1,' + F + ',1'), avg(D, M + ',0,' + F + ',1'), avg(D, F + ',1')],
    ['Acceptance of GENUINE requests', avg(D, M + ',1,' + F + ',0'), avg(D, M + ',0,' + F + ',0'), avg(D, F + ',0')],
    ['Decision accuracy', avg(C, M + ',1'), avg(C, M + ',0'), avg(C, '')],
    ['Accepted because "I trusted the sender" (share of accepts)', avg(T, M + ',1,' + D + ',1'), avg(T, M + ',0,' + D + ',1'), avg(T, D + ',1')],
    ['', '', '', ''],
    ['Self-reported message influence (requests with a message)', 'Count', '', ''],
    [MSG_OPTIONS[0], '=COUNTIF(' + INF + ',1)', '', ''],
    [MSG_OPTIONS[1], '=COUNTIF(' + INF + ',2)', '', ''],
    [MSG_OPTIONS[2], '=COUNTIF(' + INF + ',3)', '', ''],
    [MSG_OPTIONS[3], '=COUNTIF(' + INF + ',4)', '', '']
  ];
  sh.getRange(1, 1, rows.length, 4).setValues(rows);
  sh.getRange('A1').setFontWeight('bold').setFontSize(14);
  sh.getRange('A6:D6').setFontWeight('bold'); sh.getRange('A13:B13').setFontWeight('bold');
  sh.getRange('B7:D11').setNumberFormat('0%');
  sh.setColumnWidth(1, 430); sh.setColumnWidths(2, 3, 130);
}

/* ---------- helpers ---------- */

function getSheet_(ss, name, headers) {
  let sh = ss.getSheetByName(name);
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
