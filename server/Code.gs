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
    importance: { type: 'STRING', enum: ['normal', 'high'] },
    area: { type: 'STRING' },          // שם תחום מהרשימה, תחום חדש מוצע, או ריק
    duplicateOf: { type: 'STRING' }    // כותרת של משימה פתוחה קיימת שדומה לזו, או ריק
  },
  required: ['title', 'notes', 'dueDate', 'importance', 'area', 'duplicateOf']
};

/** משימה מ-Gemini: השדות הרגילים, ובנוסף תחום וכפילות */
function cleanAI_(t) {
  const c = clean_(t);
  c.area = String((t && t.area) || '').trim().slice(0, 60);
  c.duplicateOf = String((t && t.duplicateOf) || '').trim().slice(0, 200);
  return c;
}

/** מה העוזר צריך לדעת: התחומים והמשימות הפתוחות (לשיוך ולזיהוי כפילויות) */
function aiContext_() {
  let areas = [], open = [];
  try {
    areas = listAreas().map(function (a) { return a.name; });
    open = listTasks().filter(function (t) { return t.status === 'open'; })
      .sort(function (a, b) { return (a.dueDate || '9999').localeCompare(b.dueDate || '9999'); })
      .slice(0, 150);
  } catch (e) {
    console.warn('אין גישה לגיליון בשביל ההקשר: ' + e.message); // ממשיכים בלי הקשר
  }
  return [
    'התחומים שלי: ' + (areas.length ? areas.map(function (n) { return '"' + n + '"'; }).join(', ') : '(אין עדיין)') + '.',
    'שיוך לתחום (area): בחר את שם התחום **בדיוק כפי שהוא כתוב ברשימה**. ' +
      'משימה של להתקשר למישהו או לשלוח לו הודעה שייכת ל"שיחות והודעות" אם התחום קיים. ' +
      'הצע תחום חדש רק אם המשימה בבירור לא מתאימה לאף תחום קיים. אם לא ברור – השאר ריק.',
    open.length
      ? 'המשימות הפתוחות שכבר יש לי:\n' + open.map(function (t) {
        return '- ' + t.title + (t.area ? ' [' + t.area + ']' : '') + (t.dueDate ? ' (יעד ' + t.dueDate + ')' : '');
      }).join('\n')
      : 'אין לי כרגע משימות פתוחות.',
    'כפילות (duplicateOf): אם משימה חדשה דומה מאוד למשימה פתוחה קיימת, כתוב את הכותרת של הקיימת בדיוק. אחרת השאר ריק.'
  ].join('\n');
}

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
      bootstrap: bootstrap,
      createTasks: createTasks, listTasks: listTasks, listAreas: listAreas, saveAreas: saveAreas,
      updateTask: updateTask, completeTask: completeTask, reopenTask: reopenTask, deleteTask: deleteTask,
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
      ? 'מצורפת הקלטה קולית בעברית שבה אני מכתיבה משימות לעבודה. ייתכן שהיא ארוכה ומבולגנת. תמלל אותה במדויק לשדה transcript.'
      : 'להלן טקסט שבו כתבתי משימות לעבודה.',
    'חלץ את כל המשימות. ייתכן שיש משימה אחת או הרבה, מעורבבות.',
    'משימה גדולה שכוללת כמה פעולות שונות – פרק לכמה משימות נפרדות וקונקרטיות.',
    'לכל משימה:',
    '- title: כותרת קצרה וברורה, עד 10 מילים, שמתחילה בפועל (למשל "להתקשר לרונית בנוגע לשיבוץ").',
    '- notes: פרטים נוספים שנאמרו, או מחרוזת ריקה.',
    '- dueDate: תאריך יעד בפורמט YYYY-MM-DD אם נאמר מועד, אחרת מחרוזת ריקה.',
    '- importance: "high" אם נאמר שהמשימה דחופה או חשובה, אחרת "normal".',
    '- area ו-duplicateOf לפי ההנחיות למטה.',
    'unclear: רשימה קצרה של דברים שלא ברורים ושכדאי לשאול עליהם (משימה מעורפלת, למי הכוונה, מה המועד כשנאמר "בקרוב"). אם הכול ברור – רשימה ריקה.',
    aiContext_(),
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
        properties: {
          transcript: { type: 'STRING' },
          tasks: { type: 'ARRAY', items: TASK_SCHEMA },
          unclear: { type: 'ARRAY', items: { type: 'STRING' } }
        },
        required: ['transcript', 'tasks', 'unclear']
      }
    }
  }, 'minimal');
  return {
    transcript: out.transcript || '',
    tasks: (out.tasks || []).map(cleanAI_).filter(function (t) { return t.title; }),
    unclear: (out.unclear || []).map(String).filter(Boolean).slice(0, 10)
  };
}

/* ================= מצב שיחה ================= */

function chatSystem_() {
  return [
    'אתה עוזר אישי לניהול משימות בעבודה. המשתמשת כותבת או מקליטה בעברית מה עובר לה בראש, לפעמים בצורה לא מסודרת.',
    'לפעמים זו הקלטה ארוכה של כמה דקות, עם הרבה משימות מעורבבות.',
    'התפקיד שלך:',
    '- לעזור לה להבין מה בדיוק צריך לעשות, ולפרק דברים גדולים או מעורפלים למשימות קונקרטיות ונפרדות (כל צעד – משימה משלו).',
    '- לשייך כל משימה לתחום, ולארגן לפי זמן: היום, השבוע, אחר כך, בלי תאריך.',
    '- כשחסר מידע חשוב, לשאול: מה תאריך היעד, כמה זה דחוף, האם זה תלוי במישהו אחר, מה הצעד הראשון.',
    '- לשאול שאלה אחת או שתיים בכל פעם, בקצרה ובבהירות. להתחיל ממה שהכי חשוב או הכי קרוב בזמן. לא להרצות ולא לתת עצות שלא ביקשה.',
    '- אם משימה דומה למשימה שכבר יש לה, לומר זאת בקצרה ("זה דומה למשימה שכבר יש לך: ...") ולשאול אם להוסיף בכל זאת.',
    '- אחרי הקלטה ארוכה, לפתוח בסיכום קצר מאוד: כמה משימות זיהית ובאילו תחומים, ואז את השאלה הראשונה.',
    '- אם היא מתלבטת או מרגישה עומס, לעזור לה לסדר עדיפויות בשאלות פשוטות.',
    '- לענות בעברית, בפנייה בלשון נקבה, בטון חם וענייני ובמשפטים קצרים, בלי הטפות.',
    'בכל תשובה מלא את השדות:',
    '- userText: אם ההודעה האחרונה שלה הייתה הקלטה, התמלול המדויק שלה. אם הייתה טקסט, החזר אותו כמו שהוא.',
    '- reply: מה שאתה עונה לה. אל תעתיק לתוכו את כל רשימת המשימות – היא מוצגת לה בנפרד.',
    '- tasks: הרשימה המלאה והעדכנית של כל המשימות שגובשו עד עכשיו בשיחה, לא רק החדשות. ' +
      'title קצר שמתחיל בפועל; notes עם פרטים; dueDate בפורמט YYYY-MM-DD או ריק; importance "high" או "normal"; area; duplicateOf. ' +
      'משימה שהיא ביקשה לא להוסיף – הסר מהרשימה.',
    '- done: true רק כשכל המשימות ברורות ואין לך עוד שאלות. אז אמור לה בקצרה שהרשימה מוכנה לשמירה.',
    'אל תמציא מועדים או פרטים שלא נאמרו. אם אמרה שאין תאריך יעד, השאר ריק ואל תשאל שוב.',
    aiContext_(),
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
  }, 'low');

  return {
    userText: out.userText || text || '',
    reply: out.reply || '',
    tasks: (out.tasks || []).map(cleanAI_).filter(function (t) { return t.title; }),
    done: !!out.done
  };
}

/* ================= משותף ================= */

/**
 * כמה "לחשוב" לפני התשובה. חשיבה ארוכה היא הסיבה העיקרית לאיטיות.
 * דגמי Gemini 3 מקבלים thinkingLevel, ודגמי 2.5 מקבלים thinkingBudget. לא ידוע מראש לאיזה דגם הכינוי מפנה,
 * לכן מנסים לפי הסדר, ושומרים ב-Script Properties את הצורה שעבדה לכל דגם.
 */
const THINKING_VARIANTS = [
  function (level) { return { thinkingLevel: level }; },
  function (level) { return { thinkingBudget: level === 'minimal' ? 0 : 512 }; },
  function () { return null; } // בלי הגדרה – ברירת המחדל של הדגם
];

function withThinking_(request, variant, level) {
  const tc = level ? THINKING_VARIANTS[variant](level) : null;
  if (!tc) return JSON.stringify(request);
  const gc = Object.assign({}, request.generationConfig, { thinkingConfig: tc });
  return JSON.stringify(Object.assign({}, request, { generationConfig: gc }));
}

/** קוראת ל-Gemini. בעומס (429/5xx) מנסה שוב, ואז עוברת לדגם הגיבוי. level = 'minimal' | 'low' */
function geminiJson_(request, level) {
  const settings = getSettings_();
  const props = PropertiesService.getScriptProperties();
  const models = [MODEL].concat(FALLBACK_MODELS);
  let body = '';
  let ok = false;

  for (let m = 0; m < models.length && !ok; m++) {
    const key = 'THINKING_' + models[m];
    const saved = Number(props.getProperty(key) || 0);
    let variant = saved;
    let attempt = 0;
    while (attempt < 2) {
      const res = UrlFetchApp.fetch(
        'https://generativelanguage.googleapis.com/v1beta/models/' + models[m] + ':generateContent', {
          method: 'post',
          contentType: 'application/json',
          headers: { 'x-goog-api-key': settings.key },
          payload: withThinking_(request, variant, level),
          muteHttpExceptions: true
        });
      const code = res.getResponseCode();
      body = res.getContentText();
      if (code === 200) {
        ok = true;
        if (level && variant !== saved) props.setProperty(key, String(variant));
        break;
      }
      // הדגם לא מכיר את צורת הגדרת החשיבה – עוברים לצורה הבאה (לא נחשב כניסיון)
      if (code === 400 && level && /thinking/i.test(body) && variant < THINKING_VARIANTS.length - 1) {
        variant++;
        continue;
      }
      if (code === 404 && m > 0) break; // דגם הגיבוי לא קיים – מדלגים
      if (code !== 429 && code < 500) {
        throw new Error('Gemini החזיר שגיאה (' + code + '): ' + body.slice(0, 300));
      }
      attempt++;
      if (attempt === 1) Utilities.sleep(2000);
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
  whatsapp: ['id', 'name', 'phone', 'message', 'sendAt', 'status', 'calendarEventId', 'createdAt'],
  // עדכונים שממתינים למייל המאוחד ל-To Do (שלב 4). התוכן נלקח מהמצב העדכני של המשימה בזמן השליחה
  outbox: ['taskId', 'action', 'at', 'sentAt']
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
    initSheet_(i === 0 ? ss.getSheets()[0].setName(name) : ss.insertSheet(name), SHEETS[name]);
  });
  appendRows_(ss.getSheetByName('areas'), DEFAULT_AREAS.map(function (a, i) { return [a.name, a.color, String(i + 1)]; }));
  props.setProperty('SHEET_ID', ss.getId());
  return ss;
}

function initSheet_(sh, head) {
  sh.getRange(1, 1, sh.getMaxRows(), head.length).setNumberFormat('@'); // הכול טקסט, בלי המרת תאריכים
  sh.getRange(1, 1, 1, head.length).setValues([head]).setFontWeight('bold');
  sh.setFrozenRows(1);
  sh.setRightToLeft(true);
  return sh;
}

/** גיליון לפי שם. גיליון שנוסף בגרסה מאוחרת יותר נוצר כאן אוטומטית */
function sheet_(name) {
  const ss = db_();
  const sh = ss.getSheetByName(name);
  if (sh) return sh;
  if (!SHEETS[name]) throw new Error('בקובץ הנתונים חסר הגיליון "' + name + '".');
  return initSheet_(ss.insertSheet(name), SHEETS[name]);
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
  const free = sh.getMaxRows() - sh.getLastRow();
  if (free < rows.length) sh.insertRowsAfter(sh.getMaxRows(), rows.length - free + 100);
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

// קריאות לא נועלות: הן רק קוראות, ונעילה הייתה מעכבת אותן בזמן שה-trigger רץ

function listTasks() {
  return readTable_(sheet_('tasks')).filter(function (t) { return t.status !== 'deleted'; });
}

function listAreas() {
  return readTable_(sheet_('areas'))
    .map(function (a) { return { name: a.name, color: a.color, order: Number(a.order) || 0 }; })
    .sort(function (a, b) { return a.order - b.order; });
}

/** כל מה שהדף צריך בפתיחה, בבקשה אחת */
function bootstrap() {
  return { areas: listAreas(), tasks: listTasks() };
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

  let areasOut = null;
  if (rows.length) withLock_(function () {
    appendRows_(sheet_('tasks'), rows);
    // תחום חדש שהעוזר הציע ונשאר במשימה שנשמרה – מתווסף לרשימת התחומים
    const sh = sheet_('areas');
    const known = readTable_(sh).map(function (a) { return a.name; });
    const fresh = [];
    created.forEach(function (t) {
      if (t.area && known.indexOf(t.area) < 0 && fresh.indexOf(t.area) < 0) fresh.push(t.area);
    });
    if (fresh.length) {
      appendRows_(sh, fresh.map(function (name, i) {
        const n = known.length + i;
        return [name, AREA_PALETTE[n % AREA_PALETTE.length], String(n + 1)];
      }));
      areasOut = true;
    }
  });
  return { created: created, failed: failed, error: error, areas: areasOut ? listAreas() : null };
}

/* ================= שינוי משימות קיימות ================= */

/** שדות שמותר לערוך מהאפליקציה */
function updateTask(id, patch) {
  return mutateTask_(id, 'update', function (t) {
    const p = patch || {};
    const c = clean_({
      title: 'title' in p ? p.title : t.title,
      notes: 'notes' in p ? p.notes : t.notes,
      dueDate: 'dueDate' in p ? p.dueDate : t.dueDate,
      importance: 'importance' in p ? p.importance : t.importance
    });
    if (!c.title) throw apiErr_('INPUT', 'לכל משימה צריכה להיות כותרת.');
    t.title = c.title; t.notes = c.notes; t.dueDate = c.dueDate; t.importance = c.importance;
    if ('area' in p) t.area = String(p.area || '').trim().slice(0, 60);
  });
}

function completeTask(id) {
  return mutateTask_(id, 'complete', function (t, now) { t.status = 'done'; t.completedAt = now; });
}

function reopenTask(id) {
  return mutateTask_(id, 'reopen', function (t) { t.status = 'open'; t.completedAt = ''; });
}

function deleteTask(id) {
  return mutateTask_(id, 'delete', function (t) { t.status = 'deleted'; });
}

/** מוצאת משימה לפי מזהה, משנה אותה, ומוסיפה רשומה לתור העדכונים ל-To Do */
function mutateTask_(id, action, change) {
  return withLock_(function () {
    const sh = sheet_('tasks');
    const head = SHEETS.tasks;
    const ids = sh.getRange(1, 1, sh.getLastRow(), 1).getDisplayValues();
    let rowNum = 0;
    for (let i = 1; i < ids.length; i++) if (ids[i][0] === id) { rowNum = i + 1; break; }
    if (!rowNum) throw apiErr_('NOT_FOUND', 'המשימה לא נמצאה בגיליון. רענני את הרשימה.');

    const range = sh.getRange(rowNum, 1, 1, head.length);
    const t = {};
    range.getDisplayValues()[0].forEach(function (v, i) { t[head[i]] = v; });
    if (t.status === 'deleted') throw apiErr_('NOT_FOUND', 'המשימה כבר נמחקה. רענני את הרשימה.');

    const now = new Date().toISOString();
    change(t, now);
    t.updatedAt = now;
    range.setNumberFormat('@');
    range.setValues([head.map(function (k) { return t[k] || ''; })]);
    appendRows_(sheet_('outbox'), [[id, action, now, '']]);
    return t;
  });
}

/* ================= סנכרון ל-To Do (מייל מאוחד) ================= */

const FLUSH_DELAY_MIN = 5;    // לא שולחים עדכון למשימה שנוצרה ממש עכשיו – Power Automate עוד לא יצר אותה ב-To Do
const QUIET_SEC = 60;         // שולחים אחרי דקה בלי שינויים חדשים, כדי שסדרת לחיצות תצא במייל אחד
const MAX_WAIT_MIN = 4;       // אבל לא מחכים יותר מזה מהשינוי הראשון
const MAX_UPDATES_PER_MAIL = 100;
const TODO_STATUS = { open: 'notStarted', done: 'completed', deleted: 'notStarted' };

/**
 * רץ מ-trigger כל דקה: אוסף את כל העדכונים שממתינים ב-outbox ושולח מייל אחד.
 * לכל משימה נשלח רק המצב העדכני שלה, גם אם היו כמה שינויים.
 */
function flushOutbox() {
  const settings = getSettings_();
  withLock_(function () {
    const ob = sheet_('outbox');
    const rows = ob.getDataRange().getDisplayValues();
    const pending = {};   // taskId -> {rows: [מספרי שורות], actions: {}}
    const order = [];
    let first = Infinity, last = 0;
    for (let i = 1; i < rows.length; i++) {
      const id = rows[i][0];
      if (!id || rows[i][3]) continue;
      if (!pending[id]) { pending[id] = { rows: [], actions: {} }; order.push(id); }
      pending[id].rows.push(i);
      pending[id].actions[rows[i][1]] = true;
      const at = new Date(rows[i][2]).getTime() || 0; // זמן לא תקין נחשב ישן
      first = Math.min(first, at);
      last = Math.max(last, at);
    }
    if (!order.length) return;
    const quiet = Date.now() - last >= QUIET_SEC * 1000;
    const tooLong = Date.now() - first >= MAX_WAIT_MIN * 60000;
    if (!quiet && !tooLong) return; // עדיין לוחצים – מחכים לסבב הבא

    const byId = {};
    readTable_(sheet_('tasks')).forEach(function (t) { byId[t.id] = t; });
    const cutoff = new Date(Date.now() - FLUSH_DELAY_MIN * 60000).toISOString();
    const now = new Date().toISOString();
    const updates = [], done = [];

    order.forEach(function (id) {
      const t = byId[id];
      if (!t) { done.push.apply(done, pending[id].rows); return; } // המשימה לא קיימת – אין מה לשלוח
      if (t.createdAt > cutoff || updates.length >= MAX_UPDATES_PER_MAIL) return; // יישלח בסבב הבא
      const action = t.status === 'deleted' ? 'delete'
        : t.status === 'done' ? 'complete'
        : pending[id].actions.reopen ? 'reopen' : 'update';
      const c = clean_(t);
      updates.push({
        id: id, action: action, title: c.title,
        notes: notesWithMeta_(c.notes, t.area, id),
        dueDate: c.dueDate, importance: c.importance,
        status: TODO_STATUS[t.status] || 'notStarted'
      });
      done.push.apply(done, pending[id].rows);
    });

    if (updates.length) {
      MailApp.sendEmail({
        to: settings.to,
        subject: '[עדכון-משימה] ' + updates.length,
        body: JSON.stringify({ updates: updates }, null, 1) // שורה לכל שדה – בלי שורות ארוכות מדי במייל
      });
    }
    if (done.length) {
      const col = rows.map(function (r) { return [r[3]]; });
      done.forEach(function (i) { col[i][0] = now; });
      col.shift();
      ob.getRange(2, 4, col.length, 1).setNumberFormat('@').setValues(col);
    }
  });
}

/* ================= עריכה ישירה בגיליון ================= */

const TRACKED_FIELDS = ['title', 'notes', 'area', 'dueDate', 'importance', 'status'];
const STATUS_ACTION = { done: 'complete', open: 'reopen', deleted: 'delete' };

/**
 * trigger מסוג onEdit (מותקן ע"י setupSync). רץ כשעורכים ידנית את הגיליון tasks, גם מאפליקציית Sheets.
 * כתיבה של הסקריפט עצמו לא מפעילה אותו.
 */
function onSheetEdit(e) {
  if (!e || !e.range) return;
  const sh = e.range.getSheet();
  if (sh.getName() !== 'tasks') return;
  const head = SHEETS.tasks;
  const r0 = Math.max(2, e.range.getRow()), r1 = e.range.getLastRow();
  if (r1 < r0) return;
  const edited = head.slice(e.range.getColumn() - 1, e.range.getLastColumn());
  if (!edited.some(function (c) { return TRACKED_FIELDS.indexOf(c) >= 0; })) return;
  const statusEdited = edited.indexOf('status') >= 0;
  const settings = getSettings_();

  withLock_(function () {
    const range = sh.getRange(r0, 1, r1 - r0 + 1, head.length);
    const values = range.getDisplayValues();
    const now = new Date().toISOString();
    const queue = [];

    const out = values.map(function (row) {
      const t = {};
      head.forEach(function (k, i) { t[k] = row[i]; });
      if (!String(t.title).trim()) return row; // שורה ריקה, או שהכותרת נמחקה – לא נוגעים

      if (!t.id) {
        // שורה חדשה שהוקלדה בגיליון → משימה חדשה ב-To Do
        const c = clean_(t);
        const id = newId_();
        try {
          MailApp.sendEmail({
            to: settings.to,
            subject: '[משימה] ' + c.title,
            body: JSON.stringify({ title: c.title, notes: notesWithMeta_(c.notes, t.area, id), dueDate: c.dueDate, importance: c.importance })
          });
        } catch (err) {
          console.error('מייל יצירה לשורה חדשה נכשל: ' + err.message);
          return row; // נשאר בלי מזהה; ייווצר בעריכה הבאה של השורה
        }
        Object.assign(t, c);
        t.id = id; t.createdAt = now; t.updatedAt = now; t.source = 'manual';
        t.status = t.status === 'done' ? 'done' : 'open';
        t.completedAt = t.status === 'done' ? now : '';
        if (t.status === 'done') queue.push([id, 'complete', now, '']);
      } else {
        if (t.status === 'deleted' && !statusEdited) return row; // עריכה של משימה שכבר נמחקה – לא מסונכרנת
        t.updatedAt = now;
        let action = 'update';
        if (statusEdited) {
          action = STATUS_ACTION[t.status] || 'update';
          if (t.status === 'done' && !t.completedAt) t.completedAt = now;
          if (t.status === 'open') t.completedAt = '';
        }
        queue.push([t.id, action, now, '']);
      }
      return head.map(function (k) { return t[k] || ''; });
    });

    range.setNumberFormat('@');
    range.setValues(out);
    appendRows_(sheet_('outbox'), queue);
  });
}

/**
 * להרצה ידנית פעם אחת מהעורך, אחרי שהזרימה החדשה ב-Power Automate מוכנה:
 * מתקינה את ה-triggers ומוסיפה לגיליון רשימות נפתחות והגנות.
 */
function setupSync() {
  const ss = withLock_(db_);
  ScriptApp.getProjectTriggers().forEach(function (tr) {
    if (['flushOutbox', 'onSheetEdit'].indexOf(tr.getHandlerFunction()) >= 0) ScriptApp.deleteTrigger(tr);
  });
  ScriptApp.newTrigger('flushOutbox').timeBased().everyMinutes(1).create();
  ScriptApp.newTrigger('onSheetEdit').forSpreadsheet(ss).onEdit().create();

  const sh = sheet_('tasks');
  const col = function (name) { return SHEETS.tasks.indexOf(name) + 1; };
  const rows = sh.getMaxRows() - 1;
  const list = function (values) {
    return SpreadsheetApp.newDataValidation().requireValueInList(values, true).setAllowInvalid(false).build();
  };
  sh.getRange(2, col('status'), rows, 1).setDataValidation(list(['open', 'done', 'deleted']));
  sh.getRange(2, col('importance'), rows, 1).setDataValidation(list(['normal', 'high']));
  sh.getRange(2, col('area'), rows, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInRange(sheet_('areas').getRange('A2:A200'), true)
      .setAllowInvalid(true).setHelpText('בחרי תחום מהרשימה, או השאירי ריק.').build());
  sh.getRange(2, col('dueDate'), rows, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireFormulaSatisfied('=OR(E2="",REGEXMATCH(E2,"^\\d{4}-\\d{2}-\\d{2}$"))')
      .setAllowInvalid(false).setHelpText('תאריך בפורמט 2026-10-20, או ריק.').build());

  sh.getProtections(SpreadsheetApp.ProtectionType.RANGE).forEach(function (p) { p.remove(); });
  ['id', 'createdAt', 'updatedAt', 'completedAt', 'source'].forEach(function (name) {
    sh.getRange(1, col(name), sh.getMaxRows(), 1).protect()
      .setDescription('מתמלא אוטומטית – לא לערוך').setWarningOnly(true);
  });
  Logger.log('הסנכרון הופעל. עדכונים יישלחו ל-To Do תוך דקה עד ארבע דקות.');
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
