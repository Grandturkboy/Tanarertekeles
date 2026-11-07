function exportResponses() {
  scheduleNext_('exportResponses', 7 * 60 * 1000);   // safety net: fires only if this run gets killed or errors out

  const start = Date.now();
  const MAX_MS = 3 * 60 * 1000;
  const props = PropertiesService.getScriptProperties();
  const header = ['class', 'response_id', 'timestamp', 'email', 'teacher', 'question_no', 'answer'];

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const forms = ss.getSheetByName('forms').getDataRange().getValues().slice(1);
  const itemRows = ss.getSheetByName('items').getDataRange().getValues().slice(1);

  const lookup = {};   // "formId|itemId" -> [teacher, question_no]
  itemRows.forEach(([cls, formId, itemId, teacher, qNo]) => {
    lookup[formId + '|' + itemId] = [teacher, qNo];
  });

  const sheet = ss.getSheetByName('responses') || ss.insertSheet('responses');
  const log = ss.getSheetByName('export_log') || ss.insertSheet('export_log');
  const next = Number(props.getProperty('exportNext') || 0);
  if (next === 0) {                       // fresh export: start from empty tabs
    sheet.clear();
    sheet.getRange('A:G').setNumberFormat('@');   // plain text, so comments starting with "=" aren't treated as formulas
    sheet.appendRow(header);
    log.clear();
    log.appendRow(['class', 'form_id', 'responses_in_form', 'responses_with_answers', 'unmapped_extra']);
  }

  for (let i = next; i < forms.length; i++) {
    if (Date.now() - start > MAX_MS) {
      scheduleNext_('exportResponses', 60 * 1000);   // replace the safety net with a quick restart
      Logger.log('Stopped early at form ' + i + ' of ' + forms.length + '; next run scheduled in about a minute.');
      return;
    }
    const [cls, formId] = forms[i];
    const out = [];
    let withAnswers = 0, unmappedExtra = 0;

    const responses = FormApp.openById(formId).getResponses();
    responses.forEach(resp => {
      let email = '', unmapped = 0;
      const rows = [];
      resp.getItemResponses().forEach(ir => {
        const hit = lookup[formId + '|' + ir.getItem().getId()];
        if (hit) rows.push([cls, resp.getId(), resp.getTimestamp().toISOString(), '', hit[0], hit[1], ir.getResponse()]);
        else { email = ir.getResponse(); unmapped++; }
      });
      if (unmapped > 1) unmappedExtra += unmapped - 1;   // more than the one email field = something is wrong
      if (rows.length) withAnswers++;
      rows.forEach(r => { r[3] = email; out.push(r); });
    });

    if (out.length) sheet.getRange(sheet.getLastRow() + 1, 1, out.length, header.length).setValues(out);
    log.appendRow([cls, formId, responses.length, withAnswers, unmappedExtra]);
    props.setProperty('exportNext', String(i + 1));
  }

  props.deleteProperty('exportNext');
  clearTriggers_('exportResponses');   // everything is exported: no more restarts
  Logger.log('Export finished. Verifying...');
  verifyExport();
}

function resetExport() {   // cancels a running chain and makes the next export start from the beginning
  clearTriggers_('exportResponses');
  PropertiesService.getScriptProperties().deleteProperty('exportNext');
}

function verifyExport() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const log = ss.getSheetByName('export_log').getDataRange().getValues().slice(1);
  const rows = ss.getSheetByName('responses').getDataRange().getValues().slice(1);

  const ids = {};   // class -> set of response ids actually in the export
  rows.forEach(r => { (ids[String(r[0])] = ids[String(r[0])] || new Set()).add(r[1]); });

  let problems = 0, total = 0;
  log.forEach(([cls, formId, inForm, withAns, extra]) => {
    const got = ids[String(cls)] ? ids[String(cls)].size : 0;
    total += got;
    if (got !== Number(withAns) || Number(extra) > 0) {
      problems++;
      Logger.log(`PROBLEM ${cls}: form has ${inForm} responses (${withAns} with answers), export has ${got}, unmapped extras: ${extra}`);
    }
  });
  Logger.log(problems ? problems + ' class(es) with problems.' : `All good: ${total} responses across ${log.length} forms match.`);
}
