(function (global) {
  "use strict";
  const G = global.Guiden,
    C = G.core;

  function plainFallback(content) {
    return String(content || "")
      .trim()
      .replace(/^```(?:html)?\s*/i, "")
      .replace(/\s*```$/, "");
  }

  function parseAnswer(content) {
    try {
      const data = C.extractJson(content),
        answer = data.answer || data.content || data.text;
      if (typeof answer !== "string" || !answer.trim()) throw new Error("Guiden returned an invalid answer.");
      return { answer, sources: C.safeSources(data.sources) };
    } catch (error) {
      const fallback = plainFallback(content);
      if (!fallback || /^[\[{]/.test(fallback)) throw error;
      console.warn("[Guiden] Writer answer used plain-text fallback", {
        parseError: error.message,
        responsePreview: fallback.slice(0, 1000)
      });
      return { answer: fallback, sources: [] };
    }
  }

  function responseText(value) {
    return String(value || "")
      .replace(/<[^>]*(?:>|$)/g, " ")
      .replace(/\s+/g, " ")
      .trim();
  }

  function refusalMessage(value) {
    const text = responseText(value);
    if (!text) return "";
    const explicit =
      /^(?:i(?: am|[’']m) sorry[,;:]?\s*(?:but\s+)?)?(?:(?:i|guiden|this action|the request|the requested action)\s+)?(?:cannot|can[’']t|am unable to|is unable to|unable to)\s+(?:complete|draft|write|generate|perform|provide)\b/i.test(
        text
      ) || /^this action cannot be completed\b/i.test(text);
    const missing =
      /\bno\s+(?:source text|source material|context|information|details)\b/i.test(text) &&
      /\b(?:provide|supply|specify|select|add)\b/i.test(text);
    return explicit || missing ? text : "";
  }

  function replacementResult(html, sources) {
    const refusal = refusalMessage(html);
    return refusal
      ? { needsInput: true, message: refusal, sources: sources || [] }
      : { replacementHtml: html, sources: sources || [] };
  }

  // Guiden sometimes returns plain text (a translation, say) with no paragraph tags. Each line becomes a <p>, so the
  // edit keeps the paragraph breaks and tracked changes can still match paragraph for paragraph.
  const BLOCK_TAG = /<\s*(?:p|h[1-6]|ul|ol|li|blockquote|table|div)\b/i;
  function ensureBlocks(html) {
    const value = String(html || "").trim();
    if (!value || BLOCK_TAG.test(value)) return value;
    return value
      .split(/\s*(?:<br\s*\/?>\s*){2,}|\s*\r?\n\s*/i)
      .map(part => part.trim())
      .filter(Boolean)
      .map(part => "<p>" + part + "</p>")
      .join("");
  }

  function parseReplacement(content) {
    const data = C.extractJson(content);
    const sources = C.safeSources(data.sources);
    const status = String(data.status || "").toLowerCase();
    if (status === "needs_input") {
      if (typeof data.message !== "string" || !data.message.trim())
        throw new Error("Guiden needs more information but did not explain what is missing.");
      return { needsInput: true, message: data.message.trim(), sources };
    }
    if (status && status !== "ready") throw new Error("Guiden returned an unsupported response status.");
    const html = data.replacementHtml || data.replacement_html;
    if (typeof html !== "string" || !html.trim())
      throw new Error("Guiden did not return validated replacement content. Nothing can be applied.");
    // A [placeholder] must stay plain text: Guiden once linked [Media contact] to an unrelated email address.
    const unlinked = html.replace(/<a\b[^>]*>\s*(\[[^\]<]{1,60}\])\s*<\/a>/gi, "$1");
    return replacementResult(ensureBlocks(unlinked), sources);
  }

  function optionsText(options) {
    return "Tone: " + options.tone + ". Length: " + options.length + ". Structure: " + options.structure + ".";
  }

  function messages(action, snapshot, instruction, options) {
    const readOnly = !action.mutates;
    const draft = action.id === "writer.draft";
    const system = readOnly
      ? 'You are Guiden inside Euro-Office Writer. Answer using the supplied document text and any Guiden retrieval context. Return JSON only: {"answer":"..."}. Do not create citations or URLs; trusted sources are attached by the server.'
      : "You are Guiden inside Euro-Office Writer. " +
        (draft
          ? "Drafting at an empty cursor is normal: create new content from the instruction and any supplied Guiden retrieval context. Do not require existing document text. " +
            "When the rest of the document is supplied, the new content must fit its subject, terms, and style; do not repeat or rewrite it. "
          : "Preserve the facts, names, numbers, meaning, tense, how strong each claim is, and existing links, and change the wording, tone, and sentence structure as much as the instruction asks. " +
            "Keep one paragraph for each source paragraph unless the instruction asks for a different structure. " +
            "A collective noun with a plural verb is correct Irish English; leave it as it is. ") +
        'For a valid result return JSON only: {"status":"ready","replacementHtml":"safe document HTML"}. ' +
        'If facts the request depends on are genuinely missing, return {"status":"needs_input","message":"a specific explanation of what is needed"}; never put a refusal or explanation in replacementHtml. ' +
        "Use only facts from the instruction, the document, and supplied retrieval context, and use each fact only for what it is. " +
        "Never make up quotes, people, dates, places (including the city in a press release dateline), prices, figures, contact details, sign-in or access details, or facts about real companies, products, or laws. " +
        "Facts in retrieval context about one product or service do not apply to another unless the context says so. " +
        "Where the format calls for a detail you were not given, write a short placeholder in square brackets instead: a press release dateline starts with [City] unless the user named a place, a quote is [Spokesperson quote] with no quoted words, and a contact is [Media contact]. " +
        "Use dates and terms exactly as given; a launch on a given date is not available immediately. " +
        "Allowed HTML: h1,h2,h3,p,strong,em,a,ul,ol,li,blockquote,table,thead,tbody,tr,th,td,br. Put every paragraph in its own <p>, even when there is only one. " +
        "Use only hyperlinks already present in source HTML or supplied retrieval context; never invent or repair a URL.";
    const context =
      "Action: " +
      action.label +
      "\nTarget: " +
      snapshot.targetLabel +
      "\n" +
      optionsText(options) +
      "\n\nSource text:\n" +
      snapshot.text +
      (snapshot.sourceHtml ? "\n\nSource HTML (preserve its safe hyperlinks):\n" + snapshot.sourceHtml : "") +
      (draft && snapshot.contextText
        ? "\n\nThe rest of the document, for context only (write new content that fits it; do not repeat it):\n" +
          snapshot.contextText
        : "");
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
      buildRequest(snapshot, instruction, options) {
        return messages(action, snapshot, instruction || defaultInstruction, options);
      },
      parseResponse: mutates ? parseReplacement : parseAnswer,
      buildPreview(result, snapshot) {
        if (!mutates) return { kind: "answer", answer: result.answer, sources: result.sources };
        if (result.needsInput) return { kind: "needs-input", message: result.message, sources: result.sources };
        if (id === "writer.draft" && snapshot.target === "cursor")
          return { kind: "writer-new", afterHtml: result.replacementHtml, sources: result.sources };
        return {
          kind: "writer-diff",
          before: snapshot.text,
          afterHtml: result.replacementHtml,
          sources: result.sources,
          warning:
            snapshot.target === "full_document" && snapshot.text.trim() ? "This will replace the whole document." : ""
        };
      },
      revalidateTarget: mutates ? snapshot => G.host.revalidate(snapshot) : async () => ({ ok: true }),
      async apply(result, snapshot, applyOptions) {
        if (result.needsInput || !result.replacementHtml)
          throw new Error("There is no validated document change to apply.");
        return G.host.applyWriter(G.ui.sanitizeHtml(result.replacementHtml), snapshot, applyOptions);
      }
    };
    G.registerAction(action);
  }

  register(
    "writer.ask",
    "Ask",
    "Ask about the selected text or document",
    false,
    "Answer my question using this content.",
    "What would you like to know about this content?"
  );
  register(
    "writer.improve",
    "Improve Writing",
    "Make writing clearer and more polished",
    true,
    "Improve clarity, flow, and readability without changing the meaning.",
    "What should Guiden improve?"
  );
  register(
    "writer.fix",
    "Fix Spelling & Grammar",
    "Correct mistakes without changing meaning",
    true,
    "Fix spelling, grammar, punctuation, and awkward phrasing.",
    "Add any preferences, or leave blank."
  );
  register(
    "writer.summarise",
    "Summarise",
    "Create a concise summary",
    true,
    "Summarise the important points clearly.",
    "What should the summary focus on?"
  );
  register(
    "writer.rewrite",
    "Rewrite or Translate",
    "Change tone, wording, or language",
    true,
    "Rewrite as instructed while preserving facts and intent.",
    "For example: translate to French, or rewrite for customers."
  );
  register(
    "writer.draft",
    "Draft New Content",
    "Create content at the cursor or replace selection",
    true,
    "Draft new content from the instruction.",
    "Describe what you want Guiden to draft."
  );
  // Review: comments in the editor's own comment pane, each anchored to a phrase quoted from the text. Quotes Guiden
  // did not copy exactly from the text are dropped before the preview, so every comment shown can be placed.
  const REVIEW_LIMIT = 25;
  function parseReview(content, snapshot) {
    const data = C.extractJson(content);
    if (String(data.status || "").toLowerCase() === "needs_input" && typeof data.message === "string")
      return { needsInput: true, message: data.message.trim(), sources: [] };
    if (!Array.isArray(data.comments)) throw new Error("Guiden did not return a list of comments.");
    const text = String((snapshot && snapshot.text) || ""),
      seen = new Set(),
      comments = [];
    let unmatched = 0;
    data.comments.forEach(item => {
      const quote = String((item && item.quote) || "").trim(),
        comment = String((item && item.comment) || "").trim();
      if (!quote || !comment || seen.has(quote + "\n" + comment)) return;
      seen.add(quote + "\n" + comment);
      if (text.indexOf(quote) < 0) {
        unmatched += 1;
        return;
      }
      // A quote inside another quote would stack two comments on the same words, so the points are joined.
      const overlap = comments.find(c => c.quote !== quote && (c.quote.includes(quote) || quote.includes(c.quote)));
      if (overlap) overlap.comment += " " + comment;
      else if (comments.length < REVIEW_LIMIT) comments.push({ quote, comment });
    });
    if (!comments.length && unmatched)
      throw new Error("Guiden's comments did not quote the document exactly, so none could be placed. Try again.");
    return { comments, unmatched, sources: [] };
  }
  G.registerAction({
    id: "writer.review",
    editor: "word",
    label: "Review with Comments",
    description: "Leave review comments without changing the text",
    mutates: true,
    supportedTargets: ["selection", "full_document"],
    defaultInstruction:
      "Review this as a careful editor. Comment on errors, unclear wording, missing information, inconsistencies and tone.",
    placeholder: "What should the review focus on? For example: legal risks, or clarity for customers.",
    buildRequest(snapshot, instruction) {
      const system =
        "You are Guiden inside Euro-Office Writer, reviewing a document the way a careful colleague would. " +
        'Return JSON only: {"status":"ready","comments":[{"quote":"...","comment":"..."}]}. ' +
        "Each quote is a short phrase of 2 to 12 words copied exactly, character for character, from the source text, " +
        "long enough to be unique. Each comment says what the problem is and how to fix it, in one or two sentences. " +
        "Comment only where something is actually wrong, unclear, or missing; leave correct text alone, and leave out weak comments. " +
        "Never comment on a collective noun with a plural verb, such as 'the team have': it is correct Irish English, not an agreement error. " +
        "Name grammar accurately: only call something passive when it is in the passive voice. " +
        "Comment on each phrase once and never quote overlapping phrases; put every point about the same words in one comment. " +
        "Give at most 15 comments, the most important first, and an empty list when nothing needs a comment. " +
        "Do not rewrite the document.";
      const context = "Target: " + snapshot.targetLabel + "\n\nSource text:\n" + snapshot.text;
      return C.groundedMessages(system, context, instruction || this.defaultInstruction);
    },
    parseResponse: parseReview,
    buildPreview(result) {
      if (result.needsInput) return { kind: "needs-input", message: result.message, sources: [] };
      return {
        kind: "comments",
        comments: result.comments,
        sources: [],
        warning: result.unmatched
          ? result.unmatched +
            (result.unmatched === 1 ? " comment was" : " comments were") +
            " left out because the text it quoted is not in the document."
          : ""
      };
    },
    revalidateTarget: snapshot => G.host.revalidate(snapshot),
    async apply(result, snapshot) {
      if (result.needsInput || !result.comments || !result.comments.length)
        throw new Error("There are no comments to add.");
      return G.host.addComments(result.comments, snapshot);
    }
  });
  register(
    "writer.convert",
    "Convert Format",
    "Turn content into an email, report, briefing, bullets, or minutes",
    true,
    "Convert the content to the selected structure.",
    "Add audience or purpose if useful."
  );
})(window);
