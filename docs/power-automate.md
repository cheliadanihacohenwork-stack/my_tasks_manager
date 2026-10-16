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

## זרימה 2: עדכון משימות

עוד לא נבנתה (שלב 4).
