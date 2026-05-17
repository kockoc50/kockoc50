// ============================================================
// WhatsApp Tasks → Google Sheets + Claude AI
// Google Apps Script – מדביקים ב-script.google.com
// ============================================================

const ANTHROPIC_API_KEY = "הדביקי-כאן-את-המפתח-שלך"; // ← שני בלבד
const SHEET_NAME = "משימות";

// ── Web App entry point ──────────────────────────────────────

function doGet(e) {
  if (e.parameter.task) {
    return handleTaskSubmit(e.parameter.task);
  }
  return HtmlService.createHtmlOutput(getFormHtml())
    .setTitle("הוסף משימה")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function doPost(e) {
  const body = JSON.parse(e.postData.contents || "{}");
  return handleTaskSubmit(body.task || "");
}

function handleTaskSubmit(taskText) {
  if (!taskText.trim()) {
    return HtmlService.createHtmlOutput(resultHtml("⚠️ לא הוזן טקסט", false));
  }
  try {
    const parsed = callClaude(taskText);
    addRowToSheet(parsed, taskText);
    return HtmlService.createHtmlOutput(
      resultHtml("✅ המשימה נוספה בהצלחה!<br><b>" + parsed.task + "</b>", true)
    );
  } catch (err) {
    return HtmlService.createHtmlOutput(resultHtml("❌ שגיאה: " + err.message, false));
  }
}

// ── Claude API ───────────────────────────────────────────────

function callClaude(text) {
  const today = new Date().toLocaleDateString("he-IL");
  const payload = {
    model: "claude-sonnet-4-6",
    max_tokens: 512,
    system: "אתה עוזר שמחלץ פרטי משימות מטקסט בעברית. החזר JSON בלבד, ללא טקסט נוסף. היום: " + today,
    messages: [{
      role: "user",
      content: `נתח את המשימה הבאה וחלץ ממנה:
"${text}"

החזר JSON בפורמט הבא בדיוק:
{
  "task": "כותרת קצרה וברורה",
  "priority": "גבוהה|בינונית|נמוכה",
  "due_date": "DD/MM/YYYY או null",
  "category": "עבודה|אישי|קניות|בית|בריאות|פיננסי|לימודים|אחר",
  "notes": "פרטים נוספים או null"
}`
    }]
  };

  const response = UrlFetchApp.fetch("https://api.anthropic.com/v1/messages", {
    method: "post",
    contentType: "application/json",
    headers: {
      "x-api-key": ANTHROPIC_API_KEY,
      "anthropic-version": "2023-06-01"
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  const data = JSON.parse(response.getContentText());
  if (response.getResponseCode() !== 200) {
    throw new Error("Claude API: " + (data.error?.message || response.getResponseCode()));
  }

  const raw = data.content[0].text.trim().replace(/^```(?:json)?\s*/i, "").replace(/\s*```$/, "");
  return JSON.parse(raw);
}

// ── Google Sheets ────────────────────────────────────────────

function ensureSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (sheet) return sheet;

  sheet = ss.insertSheet(SHEET_NAME);
  sheet.setRightToLeft(true);

  const headers = ["#", "תאריך קבלה", "משימה", "עדיפות", "תאריך יעד", "קטגוריה", "סטטוס", "הודעה מקורית", "הערות"];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setBackground("#2B579A").setFontColor("#FFFFFF").setFontWeight("bold").setHorizontalAlignment("center");
  sheet.setFrozenRows(1);

  sheet.setColumnWidth(1, 40);
  sheet.setColumnWidth(2, 110);
  sheet.setColumnWidth(3, 300);
  sheet.setColumnWidth(4, 80);
  sheet.setColumnWidth(5, 110);
  sheet.setColumnWidth(6, 100);
  sheet.setColumnWidth(7, 90);
  sheet.setColumnWidth(8, 300);
  sheet.setColumnWidth(9, 200);

  return sheet;
}

function addRowToSheet(task, originalText) {
  const sheet = ensureSheet();
  const lastRow = sheet.getLastRow();
  const rowNum = lastRow; // header = 1, so lastRow = next task #

  const values = [
    rowNum,
    new Date().toLocaleDateString("he-IL"),
    task.task || originalText,
    task.priority || "בינונית",
    task.due_date || "",
    task.category || "אחר",
    "פתוח",
    originalText,
    task.notes || ""
  ];

  sheet.appendRow(values);

  const newRow = sheet.getLastRow();
  const priorityColors = { "גבוהה": "#FF4444", "בינונית": "#FFA500", "נמוכה": "#44BB44" };
  const color = priorityColors[task.priority] || "#CCCCCC";

  sheet.getRange(newRow, 4).setBackground(color).setFontColor("#FFFFFF").setFontWeight("bold").setHorizontalAlignment("center");
  sheet.getRange(newRow, 7).setBackground("#DDFFDD").setHorizontalAlignment("center");

  const isEven = (newRow % 2 === 0);
  if (isEven) {
    const rowRange = sheet.getRange(newRow, 1, 1, 9);
    rowRange.setBackground("#F0F4FF");
  }
}

// ── HTML ─────────────────────────────────────────────────────

function getFormHtml() {
  return `<!DOCTYPE html>
<html lang="he" dir="rtl">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1, maximum-scale=1">
<title>הוסף משימה</title>
<style>
  * { box-sizing: border-box; margin: 0; padding: 0; }
  body {
    font-family: -apple-system, 'Segoe UI', Arial, sans-serif;
    background: #f0f4ff;
    min-height: 100vh;
    display: flex;
    align-items: center;
    justify-content: center;
    padding: 20px;
  }
  .card {
    background: white;
    border-radius: 16px;
    padding: 28px 24px;
    width: 100%;
    max-width: 420px;
    box-shadow: 0 4px 24px rgba(0,0,0,0.10);
  }
  h1 { font-size: 22px; color: #2B579A; margin-bottom: 6px; }
  p.sub { color: #888; font-size: 14px; margin-bottom: 22px; }
  textarea {
    width: 100%;
    min-height: 110px;
    border: 2px solid #dde3f0;
    border-radius: 10px;
    padding: 12px 14px;
    font-size: 16px;
    font-family: inherit;
    resize: vertical;
    direction: rtl;
    outline: none;
    transition: border 0.2s;
  }
  textarea:focus { border-color: #2B579A; }
  button {
    margin-top: 14px;
    width: 100%;
    background: #2B579A;
    color: white;
    border: none;
    border-radius: 10px;
    padding: 14px;
    font-size: 17px;
    font-weight: bold;
    cursor: pointer;
    transition: background 0.2s;
  }
  button:active { background: #1a3d72; }
  .hint { margin-top: 14px; color: #aaa; font-size: 12px; text-align: center; }
</style>
</head>
<body>
<div class="card">
  <h1>📋 הוסף משימה</h1>
  <p class="sub">כתבי בחופשיות – Claude ינתח ויסדר אוטומטית</p>
  <form method="GET" action="">
    <textarea name="task" placeholder="לדוגמה: להגיש דוח עד יום שלישי – דחוף!&#10;לקנות חלב ולחם&#10;לקבוע תור לרופא" autofocus></textarea>
    <button type="submit">➕ הוסף משימה</button>
  </form>
  <p class="hint">המשימה תיכנס לגיליון Google Sheets שלך</p>
</div>
</body>
</html>`;
}

function resultHtml(msg, success) {
  const color = success ? "#2B579A" : "#cc3333";
  const bg = success ? "#e8f0fe" : "#fff0f0";
  return `<!DOCTYPE html>
<html lang="he" dir="rtl">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<style>
  body { font-family: -apple-system, Arial, sans-serif; background: #f0f4ff; display: flex; align-items: center; justify-content: center; min-height: 100vh; padding: 20px; }
  .card { background: white; border-radius: 16px; padding: 30px 24px; max-width: 400px; width: 100%; text-align: center; box-shadow: 0 4px 24px rgba(0,0,0,0.10); }
  .msg { font-size: 18px; color: ${color}; background: ${bg}; border-radius: 10px; padding: 18px; margin-bottom: 18px; }
  a { display: block; background: #2B579A; color: white; text-decoration: none; border-radius: 10px; padding: 13px; font-size: 16px; font-weight: bold; }
</style>
</head>
<body>
<div class="card">
  <div class="msg">${msg}</div>
  <a href="javascript:history.back()">← הוסף משימה נוספת</a>
</div>
</body>
</html>`;
}
