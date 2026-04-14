/* global Office, Word */

const STORAGE_KEY = "claude-word-settings";
const DEFAULT_MODEL = "claude-sonnet-4-6";
const ANTHROPIC_API_URL = "https://api.anthropic.com/v1/messages";
const ANTHROPIC_VERSION = "2023-06-01";

const CHAT_SYSTEM =
    "You are Claude, a helpful writing assistant embedded inside Microsoft Word. Respond in the same language as the user's most recent message. When the user asks for text to insert into their document, output ONLY that text with no preamble or explanation. Otherwise, answer conversationally and concisely.";

const QUICK_ACTIONS = {
    improve: {
        system:
            "You are an expert editor. Rewrite the user's text to improve clarity, flow, and word choice while preserving meaning, tone, and language. Output ONLY the rewritten text, no explanations.",
        user: (text) => `Rewrite and improve this text:\n\n${text}`,
    },
    summarize: {
        system:
            "You are an expert summarizer. Produce a concise summary in the same language as the input. Output ONLY the summary, no preamble.",
        user: (text) => `Summarize:\n\n${text}`,
    },
    "translate-en": {
        system:
            "You are a professional translator. Translate to natural, fluent English. Output ONLY the translation.",
        user: (text) => `Translate to English:\n\n${text}`,
    },
    "translate-he": {
        system:
            "You are a professional translator. Translate to natural, fluent Hebrew. Output ONLY the translation.",
        user: (text) => `Translate to Hebrew:\n\n${text}`,
    },
    shorten: {
        system:
            "You are an expert editor. Shorten the text while preserving its core meaning and language. Output ONLY the shortened text.",
        user: (text) => `Shorten:\n\n${text}`,
    },
    expand: {
        system:
            "You are an expert writer. Expand the text with additional relevant detail and examples, in the same language and tone. Output ONLY the expanded text.",
        user: (text) => `Expand this text with more detail:\n\n${text}`,
    },
    fix: {
        system:
            "You are a proofreader. Fix spelling, grammar, and punctuation errors while preserving the original language, meaning, and tone. Output ONLY the corrected text.",
        user: (text) => `Fix errors in:\n\n${text}`,
    },
    formal: {
        system:
            "You are an expert editor. Rewrite the text in a formal, professional register while preserving meaning and language. Output ONLY the rewritten text.",
        user: (text) => `Rewrite in a formal tone:\n\n${text}`,
    },
};

let currentResult = "";
let conversation = []; // array of { role, content }

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
    document.getElementById("new-chat").addEventListener("click", newConversation);
    document.getElementById("save-settings").addEventListener("click", saveSettings);

    document.querySelectorAll(".action-btn").forEach((btn) => {
        btn.addEventListener("click", () => runQuickAction(btn.dataset.action));
    });

    document.getElementById("run-custom").addEventListener("click", runChatTurn);
    document.getElementById("custom-prompt").addEventListener("keydown", (e) => {
        if (e.key === "Enter" && (e.ctrlKey || e.metaKey)) {
            e.preventDefault();
            runChatTurn();
        }
    });

    document.getElementById("insert-result").addEventListener("click", replaceSelection);
    document.getElementById("append-result").addEventListener("click", appendToDocument);
    document.getElementById("copy-result").addEventListener("click", copyResult);
}

function loadSettings() {
    try {
        const saved = JSON.parse(localStorage.getItem(STORAGE_KEY) || "{}");
        if (saved.apiKey) document.getElementById("api-key").value = saved.apiKey;
        if (saved.model) document.getElementById("model").value = saved.model;
        if (typeof saved.streaming === "boolean") {
            document.getElementById("streaming-enabled").checked = saved.streaming;
        }
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
    const streaming = document.getElementById("streaming-enabled").checked;
    localStorage.setItem(STORAGE_KEY, JSON.stringify({ apiKey, model, streaming }));
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

function newConversation() {
    conversation = [];
    document.getElementById("chat-log").innerHTML = "";
    setStatus("שיחה חדשה נפתחה", "info");
}

function appendChatMessage(role, text) {
    const log = document.getElementById("chat-log");
    const msg = document.createElement("div");
    msg.className = "chat-message " + role;
    const roleLabel = document.createElement("div");
    roleLabel.className = "role";
    roleLabel.textContent = role === "user" ? "אתה" : "Claude";
    const body = document.createElement("div");
    body.className = "body";
    body.textContent = text;
    msg.appendChild(roleLabel);
    msg.appendChild(body);
    log.appendChild(msg);
    log.scrollTop = log.scrollHeight;
    return body;
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
        await callClaudeOneShot(preset.system, preset.user(text));
    } catch (err) {
        setStatus("שגיאה: " + err.message, "error");
    }
}

async function runChatTurn() {
    const input = document.getElementById("custom-prompt");
    const prompt = input.value.trim();
    if (!prompt) {
        setStatus("יש להזין הודעה", "error");
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

    conversation.push({ role: "user", content: userMessage });
    appendChatMessage("user", prompt);
    input.value = "";

    const assistantBody = appendChatMessage("assistant", "");

    try {
        const text = await callClaudeChat(CHAT_SYSTEM, conversation, assistantBody);
        conversation.push({ role: "assistant", content: text });
        currentResult = text;
        document.getElementById("output").textContent = text;
        document.getElementById("last-result-group").classList.remove("hidden");
    } catch (err) {
        assistantBody.textContent = "שגיאה: " + err.message;
        setStatus("שגיאה: " + err.message, "error");
        // roll back user message so retry makes sense
        conversation.pop();
    }
}

async function callClaudeOneShot(system, userMessage) {
    const bodyEl = { textContent: "" }; // dummy sink
    const output = document.getElementById("output");
    output.textContent = "";
    output.classList.remove("streaming");
    currentResult = "";
    document.getElementById("last-result-group").classList.remove("hidden");
    const sink = {
        set textContent(v) {
            output.textContent = v;
        },
        get textContent() {
            return output.textContent;
        },
    };
    try {
        setButtonsDisabled(true);
        setStatus("שולח בקשה ל-Claude...", "info");
        const text = await callClaudeAPI(system, [{ role: "user", content: userMessage }], sink);
        currentResult = text;
        output.textContent = text;
    } finally {
        setButtonsDisabled(false);
        output.classList.remove("streaming");
    }
}

async function callClaudeChat(system, messages, targetEl) {
    try {
        setButtonsDisabled(true);
        setStatus("שולח בקשה ל-Claude...", "info");
        const text = await callClaudeAPI(system, messages, targetEl);
        return text;
    } finally {
        setButtonsDisabled(false);
    }
}

async function callClaudeAPI(system, messages, targetEl) {
    const { apiKey, model, streaming } = getSettings();
    if (!apiKey) {
        setStatus("חסר מפתח API - פתח את ההגדרות", "error");
        document.getElementById("settings-panel").classList.remove("hidden");
        throw new Error("Missing API key");
    }

    const useStreaming = streaming !== false;
    const body = {
        model: model || DEFAULT_MODEL,
        max_tokens: 4096,
        system: [
            {
                type: "text",
                text: system,
                cache_control: { type: "ephemeral" },
            },
        ],
        messages,
        stream: useStreaming,
    };

    const response = await fetch(ANTHROPIC_API_URL, {
        method: "POST",
        headers: {
            "Content-Type": "application/json",
            "x-api-key": apiKey,
            "anthropic-version": ANTHROPIC_VERSION,
            "anthropic-dangerous-direct-browser-access": "true",
        },
        body: JSON.stringify(body),
    });

    if (!response.ok) {
        const errText = await response.text();
        throw new Error(`HTTP ${response.status}: ${errText}`);
    }

    if (!useStreaming) {
        const data = await response.json();
        const text = (data.content || [])
            .filter((b) => b.type === "text")
            .map((b) => b.text)
            .join("");
        targetEl.textContent = text;
        const usage = data.usage || {};
        setStatus(usageString(usage), "success");
        return text;
    }

    // SSE streaming
    const output = document.getElementById("output");
    output.classList.add("streaming");
    const reader = response.body.getReader();
    const decoder = new TextDecoder();
    let buffer = "";
    let fullText = "";
    let usage = {};

    while (true) {
        const { value, done } = await reader.read();
        if (done) break;
        buffer += decoder.decode(value, { stream: true });

        const events = buffer.split("\n\n");
        buffer = events.pop() || "";

        for (const evt of events) {
            const lines = evt.split("\n");
            let dataLine = "";
            for (const l of lines) {
                if (l.startsWith("data: ")) dataLine = l.slice(6);
            }
            if (!dataLine) continue;
            try {
                const parsed = JSON.parse(dataLine);
                if (parsed.type === "content_block_delta" && parsed.delta?.type === "text_delta") {
                    fullText += parsed.delta.text;
                    targetEl.textContent = fullText;
                    const log = document.getElementById("chat-log");
                    log.scrollTop = log.scrollHeight;
                } else if (parsed.type === "message_start" && parsed.message?.usage) {
                    usage = { ...usage, ...parsed.message.usage };
                } else if (parsed.type === "message_delta" && parsed.usage) {
                    usage = { ...usage, ...parsed.usage };
                }
            } catch (e) {
                // ignore malformed SSE chunks
            }
        }
    }

    output.classList.remove("streaming");
    setStatus(usageString(usage), "success");
    return fullText;
}

function usageString(usage) {
    const cacheRead = usage.cache_read_input_tokens || 0;
    const cacheCreate = usage.cache_creation_input_tokens || 0;
    const parts = [
        `קלט: ${usage.input_tokens || 0}`,
        `פלט: ${usage.output_tokens || 0}`,
    ];
    if (cacheRead) parts.push(`cache-read: ${cacheRead}`);
    if (cacheCreate) parts.push(`cache-write: ${cacheCreate}`);
    return "הושלם. " + parts.join(", ");
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
        if (b.id !== "settings-toggle" && b.id !== "new-chat") b.disabled = disabled;
    });
}
