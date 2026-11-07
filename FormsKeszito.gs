const CONFIG = {
  titlePrefix: 'Tanárértékelés – ',
  description: [
    'Kedves Diákok!',
    'Kérünk benneteket, hogy töltsétek ki az alábbi, tanáraitokra vonatkozó kérdőívet. Az értékelés során az 1-es minden esetben a legnegatívabb, az 5-ös pedig a legpozitívabb választ jelenti.',
    'Ha valamelyik tanár nem tanít titeket, az adott részt kérjük hagyjátok ki.',
    'A kérdőív végén lehetőségetek van szöveges visszajelzést is adni. Kérjük, hogy ezt kulturált, tiszteletteljes formában tegyétek meg – olyan módon, ahogyan ti is szívesen fogadnátok. A bántó vagy vulgáris megjegyzéseket töröljük.',
    'Köszönjük az együttműködéseteket!',
    'Az EGYOSZ',
  ].join('\n\n'),
  emailTitle: 'Iskolai e-mail-címed',

  // type: 'scale' -> 1-5 scale with end labels, 'paragraph' -> free text
  questions: [
    { text: '1. Mennyire érthetően adja le a tananyagot a tanár?', low: 'egyáltalán nem', high: 'teljes mértékben' },
    { text: '2. Mennyire felkészülten érkezik a tanár az órára?', low: 'egyáltalán nem', high: 'teljes mértékben' },
    { text: '3. Mennyire világosak a tanár értékelési szempontjai, és mennyire tartja be ezeket?', low: 'egyáltalán nem', high: 'teljes mértékben' },
    { text: '4. Mennyire kezdi és fejezi be pontosan az órákat?', low: 'soha', high: 'mindig' },
    { text: '5. Mennyire könnyű a tanárral kommunikálni? Mennyire elérhető az iskolán kívül?', low: 'egyáltalán nem', high: 'teljes mértékben' },
    { text: '6. Ha valamit nem értesz órán, vagy kérdésed van, mennyire kapsz segítséget a tanártól?', low: 'egyáltalán nem', high: 'teljes mértékben' },
    { text: '7. Mennyire egyszerű teljesíteni a tantárgyat az adott tanárnál?', low: 'nagyon nehéz', high: 'nagyon könnyű' },
    { text: '8. Mennyire motivál a tanár az adott tantárgy tanulására?', low: 'egyáltalán nem', high: 'teljes mértékben' },
    { text: '9. Összességében milyen jegyet adnál a tanárodnak?', low: 'elégtelen', high: 'jeles' },
    { text: '10. Ha bármilyen észrevételed, javaslatod van a tanárodnak, kérlek írd le!', type: 'paragraph' },
  ],
};

function createForms() {
  scheduleNext_('createForms', 7 * 60 * 1000);   // safety net, first line of the function

  const start = Date.now();
  const MAX_MS = 4 * 60 * 1000;   // checked before each teacher section, which takes well under 2 minutes

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const teachersSheet = ss.getSheetByName('teachers') || ss.getSheets()[0];
  const formsSheet = getOrCreate_(ss, 'forms', ['class', 'form_id', 'edit_url', 'share_url', 'teachers_done']);
  const itemsSheet = getOrCreate_(ss, 'items', ['class', 'form_id', 'item_id', 'teacher', 'question_no']);

  // progress per class, read from the forms tab
  const progress = {};
  formsSheet.getDataRange().getValues().slice(1).forEach((r, i) => {
    progress[String(r[0])] = { row: i + 2, formId: r[1], done: Number(r[4]) || 0 };
  });

  // group teachers by class
  const byClass = {};
  teachersSheet.getDataRange().getValues().slice(1).forEach(([cls, teacher]) => {
    if (!cls || !teacher) return;
    (byClass[String(cls)] = byClass[String(cls)] || []).push(String(teacher));
  });

  for (const [cls, teachers] of Object.entries(byClass)) {
    let p = progress[cls];
    if (p && p.done >= teachers.length) continue;          // class finished

    let form;
    if (!p) {
      // new class: create the form with the intro section and register it immediately
      form = FormApp.create(CONFIG.titlePrefix + cls);
      form.setDescription(CONFIG.description);
      form.setCollectEmail(false);
      form.addTextItem()
        .setTitle(CONFIG.emailTitle)
        .setValidation(FormApp.createTextValidation().requireTextIsEmail().build())
        .setRequired(true);
      formsSheet.appendRow([cls, form.getId(), form.getEditUrl(), form.getPublishedUrl(), 0]);
      p = { row: formsSheet.getLastRow(), formId: form.getId(), done: 0 };
    } else {
      // resuming: drop anything left over from a teacher section that was cut off midway
      form = FormApp.openById(p.formId);
      const expected = 1 + p.done * (CONFIG.questions.length + 1);   // email item + (page break + questions) per finished teacher
      const items = form.getItems();
      for (let i = items.length - 1; i >= expected; i--) form.deleteItem(i);
    }

    for (let k = p.done; k < teachers.length; k++) {
      if (Date.now() - start > MAX_MS) {
        scheduleNext_('createForms', 60 * 1000);
        Logger.log('Stopped early to stay under the time limit; next run scheduled in about a minute.');
        return;
      }
      const teacher = teachers[k];
      form.addPageBreakItem().setTitle(teacher);

      const itemRows = [];
      CONFIG.questions.forEach((q, i) => {
        let item;
        if (q.type === 'paragraph') {
          item = form.addParagraphTextItem().setTitle(q.text).setRequired(false);
        } else {
          item = form.addScaleItem()
            .setTitle(q.text)
            .setBounds(1, 5)
            .setLabels(q.low, q.high)
            .setRequired(false);
        }
        itemRows.push([cls, form.getId(), item.getId(), teacher, i + 1]);
      });

      // section is complete: record its items first, then bump the counter
      itemsSheet.getRange(itemsSheet.getLastRow() + 1, 1, itemRows.length, 5).setValues(itemRows);
      formsSheet.getRange(p.row, 5).setValue(k + 1);
    }
  }
  clearTriggers_('createForms');
  Logger.log('All classes done.');
}

function getOrCreate_(ss, name, header) {
  const sheet = ss.getSheetByName(name) || ss.insertSheet(name);
  if (sheet.getLastRow() === 0) sheet.appendRow(header);
  return sheet;
}

function clearTriggers_(handler) {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === handler)
    .forEach(t => ScriptApp.deleteTrigger(t));
}

function scheduleNext_(handler, ms) {
  clearTriggers_(handler);
  ScriptApp.newTrigger(handler).timeBased().after(ms).create();
}
