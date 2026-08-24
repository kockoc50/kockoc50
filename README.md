# Claude for Word

חיבור של Claude AI ל-Microsoft Word באמצעות Office Add-in. מאפשר לשכתב, לסכם, לתרגם ולשפר טקסט ישירות בתוך המסמך.

## תכונות

- **פעולות מהירות על טקסט מסומן** - שיפור ניסוח, סיכום, תרגום (עברית/אנגלית), קיצור, הרחבה, תיקון שגיאות, סגנון רשמי.
- **שיחה מרובת פניות** - שיחה עם זיכרון על פניות קודמות, עם כפתור "שיחה חדשה".
- **Streaming** - תוצאה מופיעה בזמן אמת תוך כדי יצירה (ניתן לכבות בהגדרות).
- **כפתורי Ribbon ישירים** - "Improve", "Fix Errors", "Summarize" פועלים ישירות על הטקסט המסומן בלי לפתוח taskpane.
- **הזרקת תוצאה ישירות למסמך** - החלפת הטקסט המסומן או הוספה בסוף המסמך.
- **בחירת מודל** - Opus 4.6 / Sonnet 4.6 / Haiku 4.5.
- **Prompt caching** - הוראות המערכת נשמרות ב-cache להוזלת עלות.

## דרישות

- Microsoft Word (Desktop / Microsoft 365 / Online).
- Node.js 18+ (לטעינה בפיתוח).
- מפתח API של Anthropic: https://console.anthropic.com/

## התקנה והפעלה (פיתוח)

```bash
npm install
npm start
```

הפקודה `npm start` תיצור תעודות SSL מקומיות, תפעיל שרת על `https://localhost:3000`, ותטען את ה-Add-in לתוך Word. בפעם הראשונה ייתכן שתתבקש לאשר את התעודה.

### טעינה ידנית (sideload)

אם אוטומציה אינה עובדת, ניתן לטעון את ה-Add-in ידנית:

1. **Word Desktop (Windows/Mac):**
   `Insert → My Add-ins → Upload My Add-in → Browse` ובחר את `manifest.xml`.
2. **Word Online:**
   `Insert → Office Add-ins → Upload My Add-in` ובחר את `manifest.xml`.

## שימוש

1. פתח מסמך Word.
2. לחץ על **Open Claude** בלשונית Home (או פתח את ה-taskpane).
3. בהגדרות (⚙️) הזן את מפתח ה-API של Anthropic ובחר מודל. ההגדרות נשמרות מקומית בדפדפן.
4. סמן טקסט במסמך ולחץ על אחת הפעולות המהירות, או כתוב הוראה חופשית.
5. השתמש בכפתורים **החלף טקסט מסומן** / **הוסף בסוף המסמך** כדי להזריק את התוצאה למסמך.

## מבנה הפרויקט

```
manifest.xml              # מניפסט ה-Add-in של Office
src/taskpane/
  taskpane.html           # ממשק ה-taskpane
  taskpane.css            # עיצוב
  taskpane.js             # לוגיקה + קריאה ל-Claude API
package.json              # תלויות ו-scripts
```

## אבטחה

ב-Add-in הזה המפתח נשמר ב-`localStorage` של ה-taskpane, והקריאות ל-Anthropic API יוצאות ישירות מהדפדפן (עם הכותרת `anthropic-dangerous-direct-browser-access`). זה מתאים לשימוש אישי/פיתוח.

**לשימוש בארגון או בהפצה:** יש להעביר את הקריאות דרך פרוקסי בצד שרת שמחזיק את המפתח, ולהחליף את הקריאה ב-`taskpane.js` לכתובת הפרוקסי.

## מודלים

ברירת המחדל היא `claude-sonnet-4-6`. ניתן לבחור גם:
- `claude-opus-4-6` - החזק ביותר
- `claude-haiku-4-5-20251001` - המהיר ביותר

## פרויקטים נוספים במאגר

- [`shift-management/`](./shift-management) - אפליקציית ניהול משמרות לנהגי משאיות (שלוש משמרות יומיות), עצמאית ונפרדת מה-Add-in של Word.
