/**
 * משימות בקול – צד השרת (Google Apps Script)
 * דף ב-GitHub Pages ← הקלטה/טקסט/שיחה ← Gemini ← מייל ל-Outlook ← Power Automate ← To Do
 *
 * הגדרות נדרשות ב-Project Settings ← Script Properties:
 *   GEMINI_API_KEY – מפתח מ-Google AI Studio
 *   OUTLOOK_EMAIL  – כתובת ה-Outlook של העבודה
 *   APP_TOKEN      – קוד גישה סודי שתבחרי (אותו קוד מקלידים בדף בפעם הראשונה)
 */

const MODEL = 'gemini-flash-latest'; // אפשר להחליף לדגם ספציפי אם תרצי
const FALLBACK_MODELS = ['gemini-flash-lite-latest']; // כשהדגם הראשי עמוס (503)
const TZ = 'Asia/Jerusalem';
const DAYS = ['ראשון', 'שני', 'שלישי', 'רביעי', 'חמישי', 'שישי', 'שבת'];

const TASK_SCHEMA = {
  type: 'OBJECT',
  properties: {
    title: { type: 'STRING' },
    notes: { type: 'STRING' },
    dueDate: { type: 'STRING' },
    importance: { type: 'STRING', enum: ['normal', 'high'] }
  },
  required: ['title', 'notes', 'dueDate', 'importance']
};

function doGet() {
  return ContentService.createTextOutput('משימות בקול – השרת פועל.');
}

/** כל הבקשות מהדף (GitHub Pages) מגיעות לכאן */
function doPost(e) {
  let out;
  try {
    const req = JSON.parse(e.postData.contents);
    const token = PropertiesService.getScriptProperties().getProperty('APP_TOKEN');
    if (!token) throw apiErr_('SETUP', 'חסר APP_TOKEN ב-Script Properties.');
    if (req.token !== token) throw apiErr_('TOKEN', 'קוד הגישה שגוי.');
    const actions = {
      processAudio: processAudio, processText: processText, chatTurn: chatTurn,
      createTasks: createTasks, listTasks: listTasks, listAreas: listAreas, saveAreas: saveAreas,
      sendTasks: sendTasks, // ישן – נשאר זמין לדפים שעוד לא התעדכנו
      ping: function () { return 'ok'; }
    };
    const fn = actions[req.action];
    if (!fn) throw apiErr_('ACTION', 'פעולה לא מוכרת.');
    out = { ok: true, result: fn.apply(null, req.args || []) };
  } catch (err) {
    out = { ok: false, error: err.code || 'ERROR', message: err.message || String(err) };
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

function apiErr_(code, msg) {
  const e = new Error(msg);
  e.code = code;
  return e;
}

function getSettings_() {
  const p = PropertiesService.getScriptProperties();
  const key = p.getProperty('GEMINI_API_KEY');
  const to = p.getProperty('OUTLOOK_EMAIL');
  if (!key || !to) {
    throw new Error('חסרות הגדרות: יש להוסיף GEMINI_API_KEY ו-OUTLOOK_EMAIL ב-Script Properties.');
  }
  return { key: key, to: to };
}

function dateContext_() {
  const now = new Date();
  const today = Utilities.formatDate(now, TZ, 'yyyy-MM-dd');
  const dayIdx = Number(Utilities.formatDate(now, TZ, 'u')) % 7; // 0 = ראשון
  return 'היום יום ' + DAYS[dayIdx] + ', ' + today + '. חשב תאריכים יחסיים לפי זה: ' +
    '"מחר" = היום הבא; "ביום חמישי" = יום חמישי הקרוב; "בעוד שבוע" = שבעה ימים מהיום; "סוף השבוע" = יום שישי הקרוב.';
}

/* ================= הקלטה מהירה ================= */

function processAudio(base64, mimeType) {
  return extractTasks_([
    { text: quickPrompt_(true) },
    { inline_data: { mime_type: mimeType || 'audio/wav', data: base64 } }
  ]);
}

function processText(text) {
  const result = extractTasks_([{ text: quickPrompt_(false) + '\n\nהטקסט:\n' + text }]);
  result.transcript = '';
  return result;
}

function quickPrompt_(isAudio) {
  return [
    isAudio
      ? 'מצורפת הקלטה קולית בעברית שבה אני מכתיבה משימות לעבודה. תמלל אותה במדויק לשדה transcript.'
      : 'להלן טקסט שבו כתבתי משימות לעבודה.',
    'חלץ את כל המשימות. ייתכן שיש משימה אחת או כמה.',
    'לכל משימה:',
    '- title: כותרת קצרה וברורה, עד 10 מילים, שמתחילה בפועל (למשל "להתקשר לרונית בנוגע לשיבוץ").',
    '- notes: פרטים נוספים שנאמרו, או מחרוזת ריקה.',
    '- dueDate: תאריך יעד בפורמט YYYY-MM-DD אם נאמר מועד, אחרת מחרוזת ריקה.',
    '- importance: "high" אם נאמר שהמשימה דחופה או חשובה, אחרת "normal".',
    dateContext_(),
    'אל תמציא משימות או מועדים שלא נאמרו.'
  ].join('\n');
}

function extractTasks_(parts) {
  const out = geminiJson_({
    contents: [{ role: 'user', parts: parts }],
    generationConfig: {
      responseMimeType: 'application/json',
      responseSchema: {
        type: 'OBJECT',
        properties: { transcript: { type: 'STRING' }, tasks: { type: 'ARRAY', items: TASK_SCHEMA } },
        required: ['transcript', 'tasks']
      }
    }
  });
  return {
    transcript: out.transcript || '',
    tasks: (out.tasks || []).map(clean_).filter(function (t) { return t.title; })
  };
}

/* ================= מצב שיחה ================= */

function chatSystem_() {
  return [
    'אתה עוזר אישי לניהול משימות בעבודה. המשתמשת כותבת או מקליטה בעברית מה עובר לה בראש, לפעמים בצורה לא מסודרת.',
    'התפקיד שלך:',
    '- לעזור לה להבין מה בדיוק צריך לעשות, ולפרק דברים גדולים או מעורפלים למשימות קונקרטיות ונפרדות.',
    '- כשחסר מידע חשוב, לשאול: מה תאריך היעד, כמה זה דחוף, האם זה תלוי במישהו אחר, מה הצעד הראשון.',
    '- לשאול שאלה אחת או שתיים בכל פעם, בקצרה ובבהירות. לא להרצות ולא לתת עצות שלא ביקשה.',
    '- אם היא מתלבטת או מרגישה עומס, לעזור לה לסדר עדיפויות בשאלות פשוטות.',
    '- לענות בעברית, בפנייה בלשון נקבה, בטון חם וענייני ובמשפטים קצרים.',
    'בכל תשובה מלא את השדות:',
    '- userText: אם ההודעה האחרונה שלה הייתה הקלטה, התמלול המדויק שלה. אם הייתה טקסט, החזר אותו כמו שהוא.',
    '- reply: מה שאתה עונה לה.',
    '- tasks: הרשימה המלאה והעדכנית של כל המשימות שגובשו עד עכשיו בשיחה, לא רק החדשות. ' +
      'title קצר שמתחיל בפועל; notes עם פרטים; dueDate בפורמט YYYY-MM-DD או ריק; importance "high" או "normal".',
    '- done: true רק כשכל המשימות ברורות ואין לך עוד שאלות. אז אמור לה בקצרה שהרשימה מוכנה לשליחה.',
    'אל תמציא מועדים או פרטים שלא נאמרו. אם אמרה שאין תאריך יעד, השאר ריק ואל תשאל שוב.',
    dateContext_()
  ].join('\n');
}

/** תור אחד בשיחה. history = [{role:'user'|'model', text}] */
function chatTurn(history, text, audioB64, mimeType) {
  const contents = (history || []).map(function (m) {
    return { role: m.role === 'model' ? 'model' : 'user', parts: [{ text: String(m.text || '') }] };
  });
  const parts = [];
  if (text) parts.push({ text: text });
  if (audioB64) parts.push({ inline_data: { mime_type: mimeType || 'audio/wav', data: audioB64 } });
  if (!parts.length) throw new Error('ההודעה ריקה.');
  contents.push({ role: 'user', parts: parts });

  const out = geminiJson_({
    systemInstruction: { parts: [{ text: chatSystem_() }] },
    contents: contents,
    generationConfig: {
      responseMimeType: 'application/json',
      responseSchema: {
        type: 'OBJECT',
        properties: {
          userText: { type: 'STRING' },
          reply: { type: 'STRING' },
          tasks: { type: 'ARRAY', items: TASK_SCHEMA },
          done: { type: 'BOOLEAN' }
        },
        required: ['userText', 'reply', 'tasks', 'done']
      }
    }
  });

  return {
    userText: out.userText || text || '',
    reply: out.reply || '',
    tasks: (out.tasks || []).map(clean_).filter(function (t) { return t.title; }),
    done: !!out.done
  };
}

/* ================= משותף ================= */

/** קוראת ל-Gemini. בעומס (429/5xx) מנסה שוב, ואז עוברת לדגם הגיבוי */
function geminiJson_(request) {
  const settings = getSettings_();
  const payload = JSON.stringify(request);
  const models = [MODEL].concat(FALLBACK_MODELS);
  let body = '';
  let ok = false;

  for (let m = 0; m < models.length && !ok; m++) {
    for (let attempt = 0; attempt < 2; attempt++) {
      const res = UrlFetchApp.fetch(
        'https://generativelanguage.googleapis.com/v1beta/models/' + models[m] + ':generateContent', {
          method: 'post',
          contentType: 'application/json',
          headers: { 'x-goog-api-key': settings.key },
          payload: payload,
          muteHttpExceptions: true
        });
      const code = res.getResponseCode();
      body = res.getContentText();
      if (code === 200) { ok = true; break; }
      if (code === 404 && m > 0) break; // דגם הגיבוי לא קיים – מדלגים
      if (code !== 429 && code < 500) {
        throw new Error('Gemini החזיר שגיאה (' + code + '): ' + body.slice(0, 300));
      }
      if (attempt === 0) Utilities.sleep(2000);
    }
  }
  if (!ok) throw new Error('Gemini עמוס כרגע. נסי שוב בעוד דקה או שתיים.');

  const data = JSON.parse(body);
  const cand = data.candidates && data.candidates[0];
  const text = cand && cand.content && cand.content.parts
    ? cand.content.parts.map(function (p) { return p.text || ''; }).join('')
    : '';
  if (!text) throw new Error('Gemini לא החזיר תשובה. נסי שוב.');
  return JSON.parse(text);
}

function clean_(t) {
  const due = String(t.dueDate || '').trim();
  return {
    title: String(t.title || '').trim().slice(0, 200),
    notes: String(t.notes || '').trim(),
    dueDate: /^\d{4}-\d{2}-\d{2}$/.test(due) ? due : '',
    importance: t.importance === 'high' ? 'high' : 'normal'
  };
}

/* ================= גיליון: משימות ותחומים ================= */

const SHEETS = {
  tasks: ['id', 'title', 'notes', 'area', 'dueDate', 'importance', 'status', 'createdAt', 'updatedAt', 'completedAt', 'source'],
  areas: ['name', 'color', 'order'],
  whatsapp: ['id', 'name', 'phone', 'message', 'sendAt', 'status', 'calendarEventId', 'createdAt']
};
const DEFAULT_AREAS = [
  { name: 'החלפת רכזת מרחב חדרה–אור עקיבא', color: '#2E6B5E' },
  { name: 'ריכוז מבוא לתכנות', color: '#3D5A99' },
  { name: 'סדר וארגון מסמכים', color: '#A0672B' },
  { name: 'שיחות והודעות', color: '#8A4A7A' }
];
const AREA_PALETTE = ['#2E6B5E', '#3D5A99', '#A0672B', '#8A4A7A', '#5C7A2E', '#B04A4A', '#4A7F99', '#6B5E2E'];

/** להרצה ידנית פעם אחת מהעורך: יוצרת את הגיליון ומאשרת את ההרשאות */
function setup() {
  const ss = withLock_(db_);
  Logger.log('הגיליון מוכן: ' + ss.getUrl());
  return ss.getUrl();
}

function withLock_(fn) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) throw apiErr_('BUSY', 'השרת עסוק כרגע. נסי שוב בעוד כמה שניות.');
  try { return fn(); } finally { lock.releaseLock(); }
}

/** מחזירה את קובץ הנתונים, ויוצרת אותו בהרצה הראשונה. לקרוא רק מתוך withLock_ */
function db_() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('SHEET_ID');
  if (id) return SpreadsheetApp.openById(id);

  const ss = SpreadsheetApp.create('משימות בקול – נתונים');
  Object.keys(SHEETS).forEach(function (name, i) {
    const sh = i === 0 ? ss.getSheets()[0].setName(name) : ss.insertSheet(name);
    const head = SHEETS[name];
    sh.getRange(1, 1, sh.getMaxRows(), head.length).setNumberFormat('@'); // הכול טקסט, בלי המרת תאריכים
    sh.getRange(1, 1, 1, head.length).setValues([head]).setFontWeight('bold');
    sh.setFrozenRows(1);
    sh.setRightToLeft(true);
  });
  appendRows_(ss.getSheetByName('areas'), DEFAULT_AREAS.map(function (a, i) { return [a.name, a.color, String(i + 1)]; }));
  props.setProperty('SHEET_ID', ss.getId());
  return ss;
}

function sheet_(name) {
  const sh = db_().getSheetByName(name);
  if (!sh) throw new Error('בקובץ הנתונים חסר הגיליון "' + name + '".');
  return sh;
}

function readTable_(sh) {
  const values = sh.getDataRange().getDisplayValues();
  const head = values.shift() || [];
  return values.filter(function (r) { return r[0]; }).map(function (r) {
    const o = {};
    head.forEach(function (h, i) { o[h] = r[i]; });
    return o;
  });
}

function appendRows_(sh, rows) {
  if (!rows.length) return;
  const range = sh.getRange(sh.getLastRow() + 1, 1, rows.length, rows[0].length);
  range.setNumberFormat('@');
  range.setValues(rows);
}

function newId_() {
  return 't_' + Utilities.getUuid().replace(/-/g, '').slice(0, 8);
}

/** ההערות שנשלחות ל-To Do: הטקסט, ואחריו שורות התחום והמזהה (כך Power Automate ימצא את המשימה) */
function notesWithMeta_(notes, area, id) {
  const lines = [];
  if (area) lines.push('תחום: ' + area);
  lines.push('מזהה: ' + id);
  return (notes ? notes + '\n\n' : '') + lines.join('\n');
}

function listTasks() {
  return withLock_(function () {
    return readTable_(sheet_('tasks')).filter(function (t) { return t.status !== 'deleted'; });
  });
}

function listAreas() {
  return withLock_(function () {
    return readTable_(sheet_('areas'))
      .map(function (a) { return { name: a.name, color: a.color, order: Number(a.order) || 0 }; })
      .sort(function (a, b) { return a.order - b.order; });
  });
}

/** מחליפה את כל רשימת התחומים. areas = [{name, color}] לפי הסדר */
function saveAreas(areas) {
  const seen = {};
  const rows = [];
  (areas || []).forEach(function (a) {
    const name = String((a && a.name) || '').trim().slice(0, 60);
    if (!name || seen[name]) return;
    seen[name] = true;
    const color = /^#[0-9a-f]{6}$/i.test(a.color) ? a.color : AREA_PALETTE[rows.length % AREA_PALETTE.length];
    rows.push([name, color, String(rows.length + 1)]);
  });
  withLock_(function () {
    const sh = sheet_('areas');
    if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, 3).clearContent();
    appendRows_(sh, rows);
  });
  return listAreas();
}

/**
 * שומרת משימות בגיליון ושולחת מייל יצירה לכל אחת (באותו פורמט שהזרימה הקיימת מכירה).
 * מחזירה {created, failed, error}: אם המכסה נגמרה באמצע, מה שנשלח נשמר ומה שלא – חוזר לדף.
 */
function createTasks(tasks, source) {
  const settings = getSettings_();
  const src = ['quick', 'chat', 'manual'].indexOf(source) >= 0 ? source : 'manual';
  const now = new Date().toISOString();
  const created = [], failed = [], rows = [];
  let error = '';

  (tasks || []).forEach(function (raw) {
    const t = clean_(raw);
    if (!t.title) return;
    t.area = String((raw && raw.area) || '').trim().slice(0, 60);
    if (error) { failed.push(t); return; }
    t.id = newId_();
    try {
      MailApp.sendEmail({
        to: settings.to,
        subject: '[משימה] ' + t.title,
        body: JSON.stringify({ title: t.title, notes: notesWithMeta_(t.notes, t.area, t.id), dueDate: t.dueDate, importance: t.importance })
      });
    } catch (e) {
      error = /limit|quota|too many/i.test(String(e.message))
        ? 'נגמרה מכסת המיילים היומית של Gmail. אפשר לנסות שוב מחר.'
        : 'שליחת המייל נכשלה: ' + e.message;
      delete t.id;
      failed.push(t);
      return;
    }
    t.status = 'open'; t.createdAt = now; t.updatedAt = now; t.completedAt = ''; t.source = src;
    rows.push(SHEETS.tasks.map(function (k) { return t[k] || ''; }));
    created.push(t);
  });

  if (rows.length) withLock_(function () { appendRows_(sheet_('tasks'), rows); });
  return { created: created, failed: failed, error: error };
}

/** ישן: שולח מייל אחד לכל משימה, בלי שמירה. נשאר זמין עד שכל הדפים יעברו ל-createTasks */
function sendTasks(tasks) {
  const settings = getSettings_();
  let count = 0;
  (tasks || []).map(clean_).forEach(function (task) {
    if (!task.title) return;
    MailApp.sendEmail({
      to: settings.to,
      subject: '[משימה] ' + task.title,
      body: JSON.stringify(task)
    });
    count++;
  });
  return count;
}
