require("dotenv").config();
const { Client, LocalAuth } = require("whatsapp-web.js");
const qrcode = require("qrcode-terminal");
const Anthropic = require("@anthropic-ai/sdk");
const ExcelJS = require("exceljs");
const path = require("path");

const EXCEL_PATH = process.env.EXCEL_PATH || path.join(__dirname, "tasks.xlsx");
const MODEL = "claude-sonnet-4-6";

const PRIORITY_COLORS = {
  גבוהה: "FFFF4444",
  בינונית: "FFFFA500",
  נמוכה: "FF44BB44",
};

const HEADER_COLOR = "FF2B579A";

const anthropic = new Anthropic({ apiKey: process.env.ANTHROPIC_API_KEY });

const client = new Client({
  authStrategy: new LocalAuth({ dataPath: path.join(__dirname, ".wwebjs_auth") }),
  puppeteer: {
    headless: true,
    args: ["--no-sandbox", "--disable-setuid-sandbox", "--disable-dev-shm-usage"],
  },
});

client.on("qr", (qr) => {
  console.log("\n📱 סרוק את קוד ה-QR עם הוואטסאפ שלך (הגדרות → מכשירים מקושרים → קשר מכשיר):\n");
  qrcode.generate(qr, { small: true });
  console.log("");
});

client.on("authenticated", () => {
  console.log("🔑 אימות הצליח – טוען...");
});

client.on("auth_failure", (msg) => {
  console.error("❌ כשל אימות:", msg);
  console.error('מחק את תיקיית .wwebjs_auth והפעל מחדש.');
  process.exit(1);
});

client.on("ready", async () => {
  console.log("✅ מחובר לוואטסאפ!");
  await ensureExcelFile();
  const me = client.info.wid.user;
  console.log(`👤 מחובר כ: ${me}`);
  console.log(`📊 קובץ משימות: ${EXCEL_PATH}`);
  console.log("\n🎯 מאזין להודעות ל\"הודעות שמורות\"...\n");
  console.log("──────────────────────────────────────────");
});

client.on("disconnected", (reason) => {
  console.log("⚠️  ניתוק:", reason);
});

// Fires for every message including ones you create (fromMe=true)
client.on("message_create", async (message) => {
  if (!message.fromMe) return;
  if (!client.info) return;

  const myId = client.info.wid._serialized;
  if (message.to !== myId) return; // Only "Saved Messages"

  const text = message.body.trim();
  if (!text) return;

  console.log(`\n📩 הודעה חדשה: "${text}"`);
  console.log("🤖 Claude מנתח...");

  try {
    const task = await parseTaskWithClaude(text);
    const rowNum = await addTaskToExcel(task, text);
    console.log(`✅ שורה ${rowNum} נוספה: [${task.priority}] "${task.task}"`);
    if (task.due_date) console.log(`   📅 תאריך יעד: ${task.due_date}`);
    if (task.category) console.log(`   🏷  קטגוריה: ${task.category}`);
    console.log("──────────────────────────────────────────");
  } catch (err) {
    console.error("❌ שגיאה בעיבוד ההודעה:", err.message);
  }
});

async function parseTaskWithClaude(messageText) {
  const today = new Date().toLocaleDateString("he-IL");

  const response = await anthropic.messages.create({
    model: MODEL,
    max_tokens: 512,
    system: [
      {
        type: "text",
        text: `אתה עוזר שמנתח הודעות בעברית ומחלץ מהן פרטי משימות.
תחזיר JSON תקין בלבד – ללא שום טקסט לפני או אחרי.
היום: ${today}.
אם לא ניתן לקבוע ערך מסוים, השתמש ב-null.`,
        cache_control: { type: "ephemeral" },
      },
    ],
    messages: [
      {
        role: "user",
        content: `נתח את ההודעה הבאה וחלץ ממנה פרטי משימה:
"${messageText}"

החזר JSON בפורמט הבא בדיוק:
{
  "task": "כותרת קצרה וברורה של המשימה",
  "priority": "גבוהה|בינונית|נמוכה",
  "due_date": "DD/MM/YYYY או null",
  "category": "אחת מ: עבודה | אישי | קניות | בית | בריאות | פיננסי | לימודים | אחר",
  "notes": "פרטים שלא נכנסו לכותרת, או null"
}`,
      },
    ],
  });

  const raw = response.content[0].text.trim();

  // Strip markdown code fences if Claude wrapped the JSON
  const jsonText = raw.replace(/^```(?:json)?\s*/i, "").replace(/\s*```$/, "");

  try {
    return JSON.parse(jsonText);
  } catch {
    throw new Error("Claude החזיר JSON לא תקין: " + raw.slice(0, 120));
  }
}

async function ensureExcelFile() {
  const workbook = new ExcelJS.Workbook();

  try {
    await workbook.xlsx.readFile(EXCEL_PATH);
    console.log("📊 קובץ אקסל קיים נמצא.");
    return;
  } catch {
    // File doesn't exist yet – create it
  }

  const sheet = workbook.addWorksheet("משימות", {
    views: [{ rightToLeft: true }],
    properties: { tabColor: { argb: HEADER_COLOR } },
  });

  sheet.columns = [
    { header: "#", key: "num", width: 5 },
    { header: "תאריך קבלה", key: "received", width: 14 },
    { header: "משימה", key: "task", width: 42 },
    { header: "עדיפות", key: "priority", width: 10 },
    { header: "תאריך יעד", key: "due_date", width: 13 },
    { header: "קטגוריה", key: "category", width: 13 },
    { header: "סטטוס", key: "status", width: 12 },
    { header: "הודעה מקורית", key: "original", width: 50 },
    { header: "הערות", key: "notes", width: 35 },
  ];

  const headerRow = sheet.getRow(1);
  headerRow.eachCell((cell) => {
    cell.font = { bold: true, color: { argb: "FFFFFFFF" }, size: 11 };
    cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: HEADER_COLOR } };
    cell.alignment = { horizontal: "center", vertical: "middle" };
    cell.border = {
      bottom: { style: "medium", color: { argb: "FF1A3A6B" } },
    };
  });
  headerRow.height = 22;

  await workbook.xlsx.writeFile(EXCEL_PATH);
  console.log("📊 קובץ אקסל חדש נוצר:", EXCEL_PATH);
}

async function addTaskToExcel(task, originalMessage) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(EXCEL_PATH);

  const sheet = workbook.getWorksheet("משימות");
  const newNum = sheet.rowCount; // header = row 1, so rowCount == next task number

  const row = sheet.addRow({
    num: newNum,
    received: new Date().toLocaleDateString("he-IL"),
    task: task.task || originalMessage,
    priority: task.priority || "בינונית",
    due_date: task.due_date || "",
    category: task.category || "אחר",
    status: "פתוח",
    original: originalMessage,
    notes: task.notes || "",
  });

  // Alternating row background
  const isEven = newNum % 2 === 0;
  const rowBg = isEven ? "FFF0F4FF" : "FFFFFFFF";

  row.eachCell((cell) => {
    cell.alignment = { wrapText: true, vertical: "top" };
    cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: rowBg } };
  });

  // Color the priority cell
  const priorityCell = row.getCell("priority");
  const priorityColor = PRIORITY_COLORS[task.priority] || "FFCCCCCC";
  priorityCell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: priorityColor } };
  priorityCell.font = { bold: true, color: { argb: "FFFFFFFF" } };
  priorityCell.alignment = { horizontal: "center", vertical: "top" };

  // Status cell default styling
  const statusCell = row.getCell("status");
  statusCell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFDDFFDD" } };
  statusCell.alignment = { horizontal: "center", vertical: "top" };

  row.height = 18;

  await workbook.xlsx.writeFile(EXCEL_PATH);
  return newNum;
}

// ── Start ─────────────────────────────────────────────────────────────────────

if (!process.env.ANTHROPIC_API_KEY) {
  console.error("❌ חסר מפתח ANTHROPIC_API_KEY");
  console.error("   צור קובץ .env עם: ANTHROPIC_API_KEY=sk-ant-...");
  process.exit(1);
}

console.log("🚀 מאתחל חיבור לוואטסאפ...");
client.initialize();
