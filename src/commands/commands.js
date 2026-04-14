/* global Office, Word */

const STORAGE_KEY = "claude-word-settings";
const DEFAULT_MODEL = "claude-sonnet-4-6";
const ANTHROPIC_API_URL = "https://api.anthropic.com/v1/messages";
const ANTHROPIC_VERSION = "2023-06-01";

Office.onReady(() => {
    // Commands ready.
});

function getSettings() {
    try {
        return JSON.parse(localStorage.getItem(STORAGE_KEY) || "{}");
    } catch (e) {
        return {};
    }
}

async function callClaude(system, userMessage) {
    const { apiKey, model } = getSettings();
    if (!apiKey) {
        throw new Error("Missing API key. Open the Claude taskpane and set it in Settings.");
    }
    const res = await fetch(ANTHROPIC_API_URL, {
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
            system: [{ type: "text", text: system, cache_control: { type: "ephemeral" } }],
            messages: [{ role: "user", content: userMessage }],
        }),
    });
    if (!res.ok) throw new Error(`HTTP ${res.status}: ${await res.text()}`);
    const data = await res.json();
    return (data.content || [])
        .filter((b) => b.type === "text")
        .map((b) => b.text)
        .join("");
}

async function transformSelection(systemPrompt, userPromptFn, event) {
    try {
        await Word.run(async (context) => {
            const selection = context.document.getSelection();
            selection.load("text");
            await context.sync();
            const text = (selection.text || "").trim();
            if (!text) {
                throw new Error("No text selected.");
            }
            const result = await callClaude(systemPrompt, userPromptFn(text));
            selection.insertText(result, Word.InsertLocation.replace);
            await context.sync();
        });
    } catch (err) {
        console.error(err);
        // Notification via dialog (Office dialogs require user gesture; log instead).
    } finally {
        event.completed();
    }
}

function quickImprove(event) {
    transformSelection(
        "You are an expert editor. Rewrite the user's text to improve clarity, flow, and word choice while preserving meaning, tone, and language. Output ONLY the rewritten text, no explanations.",
        (text) => `Rewrite and improve this text:\n\n${text}`,
        event
    );
}

function quickFix(event) {
    transformSelection(
        "You are a proofreader. Fix spelling, grammar, and punctuation errors while preserving language, meaning, and tone. Output ONLY the corrected text.",
        (text) => `Fix errors in:\n\n${text}`,
        event
    );
}

function quickSummarize(event) {
    transformSelection(
        "You are an expert summarizer. Produce a concise summary in the same language as the input. Output ONLY the summary.",
        (text) => `Summarize:\n\n${text}`,
        event
    );
}

// Register the functions for the manifest.
Office.actions.associate("quickImprove", quickImprove);
Office.actions.associate("quickFix", quickFix);
Office.actions.associate("quickSummarize", quickSummarize);
