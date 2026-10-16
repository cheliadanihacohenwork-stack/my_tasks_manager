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
    const actions = { processAudio: processAudio, processText: processText, chatTurn: chatTurn, sendTasks: sendTasks, ping: function () { return 'ok'; } };
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

function geminiJson_(request) {
  const settings = getSettings_();
  const url = 'https://generativelanguage.googleapis.com/v1beta/models/' + MODEL + ':generateContent';
  const res = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    headers: { 'x-goog-api-key': settings.key },
    payload: JSON.stringify(request),
    muteHttpExceptions: true
  });

  const code = res.getResponseCode();
  const body = res.getContentText();
  if (code !== 200) {
    throw new Error('Gemini החזיר שגיאה (' + code + '): ' + body.slice(0, 300));
  }

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

/** שולח מייל אחד לכל משימה. Power Automate הופך אותו למשימה ב-To Do */
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
