# משימות בקול

אפליקציה אישית לניהול משימות: מקליטים או כותבים בעברית, Gemini מחלץ משימות, והן נשלחות במייל ל-Outlook של העבודה. משם Power Automate מכניס אותן ל-Microsoft To Do.

כתובת הדף: https://cheliadanihacohenwork-stack.github.io/my_tasks_manager/

## מבנה

```
index.html               הדף (GitHub Pages מפרסם מהשורש של main)
server/Code.gs           השרת (Google Apps Script)
server/appsscript.json   המניפסט של Apps Script
server/.clasp.json       קישור לפרויקט ה-Apps Script (scriptId)
scripts/deploy.ps1       סקריפט הפרסום
docs/power-automate.md   תיעוד הזרימות ב-Power Automate
deploy.config.json       מזהה ה-deployment (מקומי בלבד, לא במאגר)
```

## פרסום

דרישות: Git, Node.js, `npm install -g @google/clasp`, ו-`clasp login` עם חשבון ה-Gmail האישי.
בפעם הראשונה מעתיקים את `deploy.config.example.json` ל-`deploy.config.json` וממלאים את `deploymentId`.

```
git add -A
git commit -m "תיאור השינוי"
npm run deploy
```

הפקודה `npm run deploy` מבצעת `clasp push`, ואחר כך מעדכנת את **אותו deployment** (כתובת ה-`/exec` לא משתנה), ובסוף מריצה `git push` כדי לעדכן את הדף.
יש גם `npm run deploy:server` לשרת בלבד ו-`npm run deploy:web` לדף בלבד.

## Script Properties

ב-Apps Script: **Project Settings** ← **Script Properties**

| שם | תוכן |
|---|---|
| `GEMINI_API_KEY` | מפתח מ-Google AI Studio |
| `OUTLOOK_EMAIL` | כתובת ה-Outlook של העבודה |
| `APP_TOKEN` | קוד הגישה שמקלידים בדף |

| `SHEET_ID` | נוצר אוטומטית כשמריצים את `setup` – מזהה קובץ הנתונים ב-Google Sheets |

אסור לשמור את הערכים האלה במאגר, כי הוא ציבורי.

## קובץ הנתונים

הקובץ "משימות בקול – נתונים" ב-Google Drive של החשבון האישי הוא מקור האמת: גיליונות `tasks`, `areas`, `whatsapp`.
בהתקנה חדשה מריצים פעם אחת את הפונקציה `setup` מהעורך של Apps Script. היא יוצרת את הקובץ ומאשרת את ההרשאות.

## סנכרון ל-To Do

- מייל יצירה (`[משימה]`) נשלח מיד לכל משימה חדשה. את המייל קולטת זרימה 1 ב-Power Automate.
- עריכה, השלמה, החזרה ומחיקה, גם מהאפליקציה וגם בעריכה ישירה של הגיליון `tasks`, נכנסות לגיליון `outbox`. ה-trigger `flushOutbox` רץ כל 5 דקות ושולח מייל מאוחד אחד (`[עדכון-משימה]`), ואת המייל הזה קולטת זרימה 2.
- שורה חדשה שמקלידים בגיליון, עם כותרת ובלי `id`, הופכת למשימה חדשה.
- מחיקת שורה מהגיליון לא מסתנכרנת. כדי למחוק משנים את `status` ל-`deleted`.
- הפעלה: אחרי שזרימה 2 מוכנה, מריצים פעם אחת את `setupSync` מהעורך. הפונקציה מתקינה את ה-triggers, מוסיפה רשימות נפתחות ומגינה על העמודות האוטומטיות. מותר להריץ אותה שוב.
- המדריך ל-Power Automate נמצא ב-[docs/power-automate.md](docs/power-automate.md).
