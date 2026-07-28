(function (global) {
  "use strict";
  const G = global.Guiden, C = G.core;

  function plainFallback(content) {
    return String(content || "").trim().replace(/^```(?:html)?\s*/i, "").replace(/\s*```$/, "");
  }

  function parseAnswer(content) {
    try {
      const data = C.extractJson(content), answer = data.answer || data.content || data.text;
      if (typeof answer !== "string" || !answer.trim()) throw new Error("Guiden returned an invalid answer.");
      return { answer, sources: C.safeSources(data.sources) };
    } catch (error) {
      const fallback = plainFallback(content);
      if (!fallback || /^[\[{]/.test(fallback)) throw error;
      console.warn("[Guiden] Writer answer used plain-text fallback", { parseError: error.message, responsePreview: fallback.slice(0, 1000) });
      return { answer: fallback, sources: [] };
    }
  }

  function responseText(value) {
    return String(value || "").replace(/<[^>]*(?:>|$)/g, " ").replace(/\s+/g, " ").trim();
  }

  function refusalMessage(value) {
    const text = responseText(value);
    if (!text) return "";
    const explicit = /^(?:i(?: am|[’']m) sorry[,;:]?\s*(?:but\s+)?)?(?:(?:i|guiden|this action|the request|the requested action)\s+)?(?:cannot|can[’']t|am unable to|is unable to|unable to)\s+(?:complete|draft|write|generate|perform|provide)\b/i.test(text)
      || /^this action cannot be completed\b/i.test(text);
    const missing = /\bno\s+(?:source text|source material|context|information|details)\b/i.test(text)
      && /\b(?:provide|supply|specify|select|add)\b/i.test(text);
    return explicit || missing ? text : "";
  }

  function replacementResult(html, sources) {
    const refusal = refusalMessage(html);
    return refusal
      ? { needsInput: true, message: refusal, sources: sources || [] }
      : { replacementHtml: html, sources: sources || [] };
  }

  function parseReplacement(content) {
    const data = C.extractJson(content);
    const sources = C.safeSources(data.sources);
    const status = String(data.status || "").toLowerCase();
    if (status === "needs_input") {
      if (typeof data.message !== "string" || !data.message.trim()) throw new Error("Guiden needs more information but did not explain what is missing.");
      return { needsInput: true, message: data.message.trim(), sources };
    }
    if (status && status !== "ready") throw new Error("Guiden returned an unsupported response status.");
    const html = data.replacementHtml || data.replacement_html;
    if (typeof html !== "string" || !html.trim()) throw new Error("Guiden did not return validated replacement content. Nothing can be applied.");
    return replacementResult(html, sources);
  }

  function optionsText(options) {
    return "Tone: " + options.tone + ". Length: " + options.length + ". Structure: " + options.structure + ".";
  }

  function messages(action, snapshot, instruction, options) {
    const readOnly = !action.mutates;
    const draft = action.id === "writer.draft";
    const system = readOnly
      ? "You are Guiden inside Euro-Office Writer. Answer using the supplied document text and any Guiden retrieval context. Return JSON only: {\"answer\":\"...\"}. Do not create citations or URLs; trusted sources are attached by the server."
      : "You are Guiden inside Euro-Office Writer. "
        + (draft
          ? "Drafting at an empty cursor is normal: create new content from the instruction and any supplied Guiden retrieval context. Do not require existing document text. "
          : "Preserve facts, names, numbers, meaning, and existing links. ")
        + "For a valid result return JSON only: {\"status\":\"ready\",\"replacementHtml\":\"safe document HTML\"}. "
        + "If required facts or instructions are genuinely missing, return {\"status\":\"needs_input\",\"message\":\"a specific explanation of what is needed\"}; never put a refusal or explanation in replacementHtml. "
        + "For company-specific facts, rely on supplied retrieval context and do not invent details. "
        + "Allowed HTML: h1,h2,h3,p,strong,em,a,ul,ol,li,blockquote,table,thead,tbody,tr,th,td,br. "
        + "Use only hyperlinks already present in source HTML or supplied retrieval context; never invent or repair a URL.";
    const context = "Action: " + action.label + "\nTarget: " + snapshot.targetLabel + "\n" + optionsText(options) + "\n\nSource text:\n" + snapshot.text + (snapshot.sourceHtml ? "\n\nSource HTML (preserve its safe hyperlinks):\n" + snapshot.sourceHtml : "");
    return C.groundedMessages(system, context, instruction);
  }

  function register(id, label, description, mutates, defaultInstruction, placeholder) {
    const action = {
      id,
      editor: "word",
      label,
      description,
      mutates,
      supportedTargets: ["selection", "full_document", "cursor"],
      defaultInstruction,
      placeholder,
      buildRequest(snapshot, instruction, options) { return messages(action, snapshot, instruction || defaultInstruction, options); },
      parseResponse: mutates ? parseReplacement : parseAnswer,
      buildPreview(result, snapshot) {
        if (!mutates) return { kind: "answer", answer: result.answer, sources: result.sources };
        if (result.needsInput) return { kind: "needs-input", message: result.message, sources: result.sources };
        if (id === "writer.draft" && snapshot.target === "cursor") return { kind: "writer-new", afterHtml: result.replacementHtml, sources: result.sources };
        return { kind: "writer-diff", before: snapshot.text, afterHtml: result.replacementHtml, sources: result.sources, warning: snapshot.target === "full_document" && snapshot.text.trim() ? "This will replace the whole document." : "" };
      },
      revalidateTarget: mutates ? snapshot => G.host.revalidate(snapshot) : async () => ({ ok: true }),
      async apply(result, snapshot) {
        if (result.needsInput || !result.replacementHtml) throw new Error("There is no validated document change to apply.");
        return G.host.applyWriter(G.ui.sanitizeHtml(result.replacementHtml), snapshot);
      }
    };
    G.registerAction(action);
  }

  register("writer.ask", "Ask", "Ask about the selected text or document", false, "Answer my question using this content.", "What would you like to know about this content?");
  register("writer.improve", "Improve Writing", "Make writing clearer and more polished", true, "Improve clarity, flow, and readability without changing the meaning.", "What should Guiden improve?");
  register("writer.fix", "Fix Spelling & Grammar", "Correct mistakes without changing meaning", true, "Fix spelling, grammar, punctuation, and awkward phrasing.", "Add any preferences, or leave blank.");
  register("writer.summarise", "Summarise", "Create a concise summary", true, "Summarise the important points clearly.", "What should the summary focus on?");
  register("writer.rewrite", "Rewrite or Translate", "Change tone, wording, or language", true, "Rewrite as instructed while preserving facts and intent.", "For example: translate to French, or rewrite for customers.");
  register("writer.draft", "Draft New Content", "Create content at the cursor or replace selection", true, "Draft new content from the instruction.", "Describe what you want Guiden to draft.");
  register("writer.convert", "Convert Format", "Turn content into an email, report, briefing, bullets, or minutes", true, "Convert the content to the selected structure.", "Add audience or purpose if useful.");
})(window);
