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

אסור לשמור את הערכים האלה במאגר, כי הוא ציבורי.
