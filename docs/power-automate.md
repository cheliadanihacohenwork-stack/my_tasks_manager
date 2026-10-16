# זרימות Power Automate

הזרימות רצות בחשבון העבודה (cyber.org.il).

## זרימה 1: יצירת משימה (קיימת, עובדת – לא לשנות)

1. **When a new email arrives (V3)**: Subject Filter `[משימה]`, From = כתובת ה-Gmail האישית.
2. **Html to text** על גוף המייל.
3. **Parse JSON** עם השדות `title`, `notes`, `dueDate`, `importance` (כולם string).
4. **Add a to-do (V3)** של **Microsoft To-Do (Business)**:
   - בשדה Due Date יש ביטוי שמחזיר null כשהתאריך ריק.
   - Importance: ברירת מחדל normal.
5. **Move email (V2)** לתיקייה "משימות שנקלטו".

פורמט המייל: נושא `[משימה] <כותרת>`, וגוף שהוא JSON בלבד, למשל:

```json
{"title":"להתקשר לרונית","notes":"","dueDate":"2026-10-20","importance":"normal"}
```

## זרימה 2: עדכון משימות (חדשה ונפרדת)

הזרימה מקבלת מייל אחד עם כמה עדכונים, ומעדכנת או מוחקת כל משימה ב-To Do לפי המזהה שבהערות.
כל הפעולות כאן נמצאות במחברים **Standard**, ואין צורך ב-premium.

פורמט המייל ששולח השרת:
- נושא: `[עדכון-משימה] 3`
- גוף:

```json
{
 "updates": [
  {
   "id": "t_ab12cd34",
   "action": "complete",
   "title": "להתקשר לרונית",
   "notes": "פרטים\n\nתחום: שיחות והודעות\nמזהה: t_ab12cd34",
   "dueDate": "2026-10-20",
   "importance": "normal",
   "status": "completed"
  }
 ]
}
```

`action` יכול להיות `update`, `complete`, `reopen` או `delete`. `status` הוא הערך שנכנס ל-To Do: `notStarted` או `completed`.
לכן הזרימה צריכה רק שני מסלולים: **מחיקה**, או **עדכון** של כל השדות כולל הסטטוס.

### לפני שמתחילים

פותחים את הזרימה הקיימת (זרימה 1), לוחצים על הפעולה **Add a to-do (V3)** ובודקים מה נבחר בשדה **To-do List**. בזרימה החדשה בוחרים **בדיוק את אותה רשימה** בכל מקום.

### שלב א – יצירת הזרימה והטריגר

1. נכנסים ל-https://make.powerautomate.com עם חשבון העבודה.
2. בתפריט הצד לוחצים **+ Create**, ואחר כך **Automated cloud flow**.
3. ב-**Flow name** כותבים `משימות בקול – עדכונים`.
4. בחיפוש הטריגר כותבים `When a new email arrives (V3)`, בוחרים את הטריגר של **Office 365 Outlook** ולוחצים **Create**.
5. לוחצים על הטריגר וממלאים:
   - **Folder:** `Inbox`
   - **Include Attachments:** `No`
   - לוחצים **Show all** (או **Advanced parameters**) ומוסיפים:
     - **From:** כתובת ה-Gmail האישית
     - **Subject Filter:** `[עדכון-משימה]`

### שלב ב – קריאת המייל

6. מתחת לטריגר לוחצים **+**, ואחר כך **Add an action**. מחפשים `Html to text` (מחבר **Content Conversion**).
   - **Content:** מ-Dynamic content (סמל הברק) בוחרים **Body** של הטריגר.
7. **+**, ואחר כך **Add an action**. מחפשים `Parse JSON` (מחבר **Data Operation**).
   - **Content:** מ-Dynamic content בוחרים **The plain text content** של Html to text.
   - **Schema:** מדביקים:

```json
{
  "type": "object",
  "properties": {
    "updates": {
      "type": "array",
      "items": {
        "type": "object",
        "properties": {
          "id": { "type": "string" },
          "action": { "type": "string" },
          "title": { "type": "string" },
          "notes": { "type": "string" },
          "dueDate": { "type": "string" },
          "importance": { "type": "string" },
          "status": { "type": "string" }
        }
      }
    }
  }
}
```

### שלב ג – טעינת המשימות מ-To Do (פעם אחת)

8. **+**, ואחר כך **Add an action**. מחפשים `List to-do's by folder (V2)` (מחבר **Microsoft To-Do (Business)**).
   - **To-do List:** אותה רשימה כמו בזרימה 1.
   - **Show all** (או **Advanced parameters**), ואז **Top Count:** `999`. **חשוב:** ברירת המחדל היא 10 משימות בלבד.
9. **לשנות את שם הפעולה:** לוחצים על **...** של הפעולה, אחר כך **Rename** (או לוחצים על השם בראש החלונית), וקוראים לה `ListTodos`. הביטויים בהמשך משתמשים בשם הזה.

### שלב ד – לולאה על העדכונים

10. **+**, ואחר כך **Add an action**. מחפשים `Apply to each` (בעורך החדש היא נקראת **For each**, מחבר **Control**).
    - **Select an output from previous steps:** מ-Dynamic content בוחרים **updates** של Parse JSON.
    - משנים את שם הלולאה ל-`ForEachUpdate`.
11. **בתוך הלולאה** לוחצים **+**, ואחר כך **Add an action**. מחפשים `Filter array` (מחבר **Data Operation**).
    - משנים את השם ל-`FindTask`.
    - **From:** לוחצים על סמל ה-fx (Expression), מדביקים `body('ListTodos')` ולוחצים **Add**.
    - **Filter Query** (שלוש תיבות):
      - בתיבה השמאלית (ערך): Expression `coalesce(item()?['body']?['content'], '')`
      - באמצעית: `contains`
      - בימנית: Expression `items('ForEachUpdate')?['id']`
    - אם יש כפתור **Edit in advanced mode**, אפשר במקום זה להדביק בו:
      `@contains(coalesce(item()?['body']?['content'], ''), items('ForEachUpdate')?['id'])`
12. בתוך הלולאה, מתחת ל-FindTask: **+**, ואחר כך **Add an action**, ואז **Condition** (מחבר **Control**).
    - שמאל: Expression `length(body('FindTask'))`
    - אופרטור: `is greater than`
    - ימין: `0`
    - (אם המשימה לא נמצאה ב-To Do, לא עושים כלום.)

### שלב ה – מחיקה או עדכון

13. בענף **True** של התנאי מוסיפים עוד **Condition**:
    - שמאל: Expression `items('ForEachUpdate')?['action']`
    - אופרטור: `is equal to`
    - ימין: `delete`
14. בענף **True** של התנאי הפנימי: **Add an action**, ואז `Delete to-do (V2)` (Microsoft To-Do (Business)).
    - **To-do List:** אותה רשימה.
    - **To-do task:** Expression `first(body('FindTask'))?['id']`
15. בענף **False** של התנאי הפנימי: **Add an action**, ואז `Update to-do (V2)` (Microsoft To-Do (Business)).
    - **To-do List:** אותה רשימה.
    - **To-do task:** Expression `first(body('FindTask'))?['id']`
    - **Show all** כדי לראות את כל השדות, ואז:
    - **Title:** Expression `items('ForEachUpdate')?['title']`
    - **Content:** Expression `replace(items('ForEachUpdate')?['notes'], decodeUriComponent('%0A'), '<br>')`
    - **Importance:** Expression `items('ForEachUpdate')?['importance']` (אם זו רשימה נפתחת, בוחרים **Enter custom value** ואז מכניסים את הביטוי)
    - **Status:** Expression `items('ForEachUpdate')?['status']` (גם כאן **Enter custom value** אם צריך)
    - **Due Date:** Expression `if(empty(items('ForEachUpdate')?['dueDate']), null, concat(items('ForEachUpdate')?['dueDate'], 'T00:00:00'))`

### שלב ו – העברת המייל לתיקייה

16. **מחוץ ללולאה** (מתחת ל-ForEachUpdate, ברמה הראשית) לוחצים **+**, ואחר כך **Add an action**, ואז `Move email (V2)` (Office 365 Outlook).
    - **Message Id:** מ-Dynamic content בוחרים **Message Id** של הטריגר.
    - **Folder:** `משימות שנקלטו`.
17. לוחצים **Save** למעלה. אם מופיעות שגיאות, הן יסומנו באדום על הפעולה הבעייתית.

### בדיקה

1. ב-Apps Script מריצים פעם אחת את `setupSync` (ראה README). עד שהיא רצה, לא נשלחים מיילי עדכון.
2. באפליקציה מסמנים משימה כהושלמה.
3. אחרי 2–5 דקות מגיע ל-Outlook מייל `[עדכון-משימה] 1`, והזרימה מעבירה אותו לתיקייה. משימה שנוצרה בחמש הדקות האחרונות ממתינה לסבב הבא, כדי שזרימה 1 תספיק ליצור אותה.
4. ב-To Do המשימה מסומנת כהושלמה.
5. אם משהו נכשל: ב-Power Automate פותחים את הזרימה, ואז **28-day run history**, ולוחצים על הריצה האחרונה. הפעולה שנכשלה מסומנת באדום, ואפשר לראות בה את הקלט והפלט.

### מגבלות ידועות

- **מחיקת תאריך יעד** באפליקציה כנראה לא תמחק את התאריך ב-To Do, כי שליחת ערך ריק ל-Due Date לא מוחקת אותו.
- הזרימה מוצאת משימות רק מתוך 999 המשימות הראשונות ברשימה. אם הרשימה תגדל מעבר לזה, צריך לנקות אותה מדי פעם ממשימות שהושלמו.
- משימות שנוצרו לפני שלב 1 (בלי מזהה בהערות) לא יתעדכנו.
