/* global Office, Word */

const STORAGE_KEY = "claude-word-settings";
const DEFAULT_MODEL = "claude-sonnet-4-6";
const ANTHROPIC_API_URL = "https://api.anthropic.com/v1/messages";
const ANTHROPIC_VERSION = "2023-06-01";

const QUICK_ACTIONS = {
    improve: {
        system: "You are an expert editor. Rewrite the user's text to improve clarity, flow, and word choice while preserving meaning, tone, and language. Output ONLY the rewritten text, no explanations.",
        user: (text) => `Rewrite and improve this text:\n\n${text}`,
    },
    summarize: {
        system: "You are an expert summarizer. Produce a concise summary in the same language as the input. Output ONLY the summary, no preamble.",
        user: (text) => `Summarize:\n\n${text}`,
    },
    "translate-en": {
        system: "You are a professional translator. Translate to natural, fluent English. Output ONLY the translation.",
        user: (text) => `Translate to English:\n\n${text}`,
    },
    "translate-he": {
        system: "You are a professional translator. Translate to natural, fluent Hebrew. Output ONLY the translation.",
        user: (text) => `Translate to Hebrew:\n\n${text}`,
    },
    shorten: {
        system: "You are an expert editor. Shorten the text while preserving its core meaning and language. Output ONLY the shortened text.",
        user: (text) => `Shorten:\n\n${text}`,
    },
    expand: {
        system: "You are an expert writer. Expand the text with additional relevant detail and examples, in the same language and tone. Output ONLY the expanded text.",
        user: (text) => `Expand this text with more detail:\n\n${text}`,
    },
    fix: {
        system: "You are a proofreader. Fix spelling, grammar, and punctuation errors while preserving the original language, meaning, and tone. Output ONLY the corrected text.",
        user: (text) => `Fix errors in:\n\n${text}`,
    },
    formal: {
        system: "You are an expert editor. Rewrite the text in a formal, professional register while preserving meaning and language. Output ONLY the rewritten text.",
        user: (text) => `Rewrite in a formal tone:\n\n${text}`,
    },
};

let currentResult = "";

Office.onReady((info) => {
    if (info.host === Office.HostType.Word) {
        loadSettings();
        bindEvents();
    }
});

function bindEvents() {
    document.getElementById("settings-toggle").addEventListener("click", () => {
        document.getElementById("settings-panel").classList.toggle("hidden");
    });

    document.getElementById("save-settings").addEventListener("click", saveSettings);

    document.querySelectorAll(".action-btn").forEach((btn) => {
        btn.addEventListener("click", () => runQuickAction(btn.dataset.action));
    });

    document.getElementById("run-custom").addEventListener("click", runCustomPrompt);

    document.getElementById("insert-result").addEventListener("click", replaceSelection);
    document.getElementById("append-result").addEventListener("click", appendToDocument);
    document.getElementById("copy-result").addEventListener("click", copyResult);
}

function loadSettings() {
    try {
        const saved = JSON.parse(localStorage.getItem(STORAGE_KEY) || "{}");
        if (saved.apiKey) document.getElementById("api-key").value = saved.apiKey;
        if (saved.model) document.getElementById("model").value = saved.model;
        if (!saved.apiKey) {
            document.getElementById("settings-panel").classList.remove("hidden");
            setStatus("נא להזין מפתח API של Anthropic בהגדרות", "info");
        }
    } catch (e) {
        console.error("loadSettings failed", e);
    }
}

function saveSettings() {
    const apiKey = document.getElementById("api-key").value.trim();
    const model = document.getElementById("model").value;
    localStorage.setItem(STORAGE_KEY, JSON.stringify({ apiKey, model }));
    setStatus("ההגדרות נשמרו", "success");
    document.getElementById("settings-panel").classList.add("hidden");
}

function getSettings() {
    try {
        return JSON.parse(localStorage.getItem(STORAGE_KEY) || "{}");
    } catch (e) {
        return {};
    }
}

async function getSelectedText() {
    return Word.run(async (context) => {
        const selection = context.document.getSelection();
        selection.load("text");
        await context.sync();
        return selection.text || "";
    });
}

async function runQuickAction(action) {
    const preset = QUICK_ACTIONS[action];
    if (!preset) return;
    try {
        const text = await getSelectedText();
        if (!text.trim()) {
            setStatus("יש לסמן טקסט במסמך לפני הפעולה", "error");
            return;
        }
        await callClaude(preset.system, preset.user(text));
    } catch (err) {
        setStatus("שגיאה: " + err.message, "error");
    }
}

async function runCustomPrompt() {
    const prompt = document.getElementById("custom-prompt").value.trim();
    if (!prompt) {
        setStatus("יש להזין הוראה", "error");
        return;
    }
    const useSelection = document.getElementById("use-selection").checked;
    let userMessage = prompt;
    if (useSelection) {
        try {
            const text = await getSelectedText();
            if (text.trim()) {
                userMessage = `${prompt}\n\n---\nטקסט מסומן:\n${text}`;
            }
        } catch (e) {
            /* ignore */
        }
    }
    const system =
        "You are Claude, a helpful writing assistant embedded inside Microsoft Word. Respond in the same language as the user's request. When the user asks for text to insert into their document, output ONLY that text with no preamble or explanation.";
    await callClaude(system, userMessage);
}

async function callClaude(system, userMessage) {
    const { apiKey, model } = getSettings();
    if (!apiKey) {
        setStatus("חסר מפתח API - פתח את ההגדרות", "error");
        document.getElementById("settings-panel").classList.remove("hidden");
        return;
    }

    const output = document.getElementById("output");
    output.textContent = "";
    currentResult = "";
    document.getElementById("result-actions").classList.add("hidden");
    setButtonsDisabled(true);
    setStatus("שולח בקשה ל-Claude...", "info");

    try {
        const response = await fetch(ANTHROPIC_API_URL, {
            method: "POST",
            headers: {
                "Content-Type": "application/json",
                "x-api-key": apiKey,
                "anthropic-version": ANTHROPIC_VERSION,
                "anthropic-dangerous-direct-browser-access": "true",
            },
            body: JSON.stringify({
                model: model || DEFAULT_MODEL,
                max_tokens: 4096,
                system: [
                    {
                        type: "text",
                        text: system,
                        cache_control: { type: "ephemeral" },
                    },
                ],
                messages: [{ role: "user", content: userMessage }],
            }),
        });

        if (!response.ok) {
            const errText = await response.text();
            throw new Error(`HTTP ${response.status}: ${errText}`);
        }

        const data = await response.json();
        const text = (data.content || [])
            .filter((block) => block.type === "text")
            .map((block) => block.text)
            .join("");

        currentResult = text;
        output.textContent = text;
        document.getElementById("result-actions").classList.remove("hidden");

        const usage = data.usage || {};
        const cacheInfo =
            usage.cache_read_input_tokens > 0
                ? ` (קריאת cache: ${usage.cache_read_input_tokens} טוקנים)`
                : "";
        setStatus(
            `הושלם. קלט: ${usage.input_tokens || 0}, פלט: ${usage.output_tokens || 0}${cacheInfo}`,
            "success"
        );
    } catch (err) {
        console.error(err);
        setStatus("שגיאה: " + err.message, "error");
    } finally {
        setButtonsDisabled(false);
    }
}

async function replaceSelection() {
    if (!currentResult) return;
    try {
        await Word.run(async (context) => {
            const selection = context.document.getSelection();
            selection.insertText(currentResult, Word.InsertLocation.replace);
            await context.sync();
        });
        setStatus("הטקסט הוחלף במסמך", "success");
    } catch (err) {
        setStatus("שגיאה בהחלפה: " + err.message, "error");
    }
}

async function appendToDocument() {
    if (!currentResult) return;
    try {
        await Word.run(async (context) => {
            const body = context.document.body;
            body.insertParagraph(currentResult, Word.InsertLocation.end);
            await context.sync();
        });
        setStatus("הטקסט נוסף בסוף המסמך", "success");
    } catch (err) {
        setStatus("שגיאה בהוספה: " + err.message, "error");
    }
}

async function copyResult() {
    if (!currentResult) return;
    try {
        await navigator.clipboard.writeText(currentResult);
        setStatus("הועתק ללוח", "success");
    } catch (err) {
        setStatus("העתקה נכשלה: " + err.message, "error");
    }
}

function setStatus(msg, type) {
    const el = document.getElementById("status");
    el.textContent = msg;
    el.className = "status " + (type || "");
}

function setButtonsDisabled(disabled) {
    document.querySelectorAll("button").forEach((b) => {
        if (b.id !== "settings-toggle") b.disabled = disabled;
    });
}
