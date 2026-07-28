(function (global) {
  "use strict";
  const G = global.Guiden, C = G.core;
  const emptyMetadata = () => ({ available: false, retrievalUsed: false, sources: [] });
  const api = { url: "https://inference.cloudcix.com/v1/guiden/chat/completions", model: "Mistral-Medium-3.5", maxTokens: 15000, lastMetadata: emptyMetadata() };
  function detailedError(message, details) { const error = new Error(message); error.guidenDetails = details || {}; return error; }
  function errorMessage(value) { return value && typeof value === "object" ? value.message || value.error || JSON.stringify(value) : String(value || "Unknown API error"); }
  function normalizeMetadata(value) { return value && typeof value === "object" ? { available: true, retrievalUsed: Boolean(value.retrieval_used), sources: C.safeSources(value.sources) } : emptyMetadata(); }
  function chunk(payload) { const choice = payload && payload.choices && payload.choices[0]; return choice && ((choice.delta && choice.delta.content) || (choice.message && choice.message.content) || choice.text) || ""; }
  async function stream(response, onUpdate) {
    const reader = response.body.getReader(), decoder = new TextDecoder("utf-8"); let buffer = "", content = "", eventCount = 0, metadata = emptyMetadata(), finished = false;
    function event(block) {
      const data = block.split(/\r?\n/).filter(line => line.startsWith("data:")).map(line => line.slice(5).trim()).join("\n");
      if (!data) return; if (data === "[DONE]") { finished = true; return; }
      let payload;
      try { payload = JSON.parse(data); } catch (parseError) { console.warn("[Guiden] Ignored malformed streaming event", { error: parseError.message, event: data.slice(0, 1000) }); return; }
      eventCount += 1;
      if (payload.error) throw detailedError("CloudCIX streaming error: " + errorMessage(payload.error), { phase: "stream", payload: payload.error, eventCount });
      if (payload.guiden) metadata = normalizeMetadata(payload.guiden);
      const value = chunk(payload); if (value) { content += value; if (onUpdate) onUpdate(content); }
    }
    while (true) { const part = await reader.read(); if (part.done) break; buffer += decoder.decode(part.value, { stream: true }); const blocks = buffer.split(/\r?\n\r?\n/); buffer = blocks.pop() || ""; for (const block of blocks) event(block); }
    buffer += decoder.decode(); if (buffer.trim()) event(buffer);
    if (!finished) throw detailedError("CloudCIX response ended before completion. Nothing can be applied.", { phase: "stream", eventCount, incomplete: true });
    if (!content.trim()) throw detailedError("CloudCIX returned an empty streaming response.", { phase: "stream", eventCount });
    return { content, metadata };
  }
  api.complete = async function (key, messages, signal, onUpdate) {
    api.lastMetadata = emptyMetadata(); let response;
    try { response = await fetch(api.url, { method: "POST", headers: { "Content-Type": "application/json", Authorization: "Bearer " + key }, body: JSON.stringify({ model: api.model, stream: true, max_tokens: api.maxTokens, response_format: { type: "json_object" }, messages }), signal }); }
    catch (error) { if (error.name === "AbortError") throw error; throw detailedError("Unable to reach CloudCIX: " + error.message, { phase: "fetch", cause: error.message }); }
    if (!response.ok) { const text = await response.text(); let message = text; try { const data = JSON.parse(text); message = data.error && (data.error.message || data.error) || text; } catch (_) {} throw detailedError("CloudCIX HTTP " + response.status + ": " + errorMessage(message || response.statusText), { phase: "http", status: response.status, statusText: response.statusText, response: text.slice(0, 4000) }); }
    const type = response.headers.get("Content-Type") || ""; let content = "";
    if (response.body && type.includes("text/event-stream")) { const streamed = await stream(response, onUpdate); content = streamed.content; api.lastMetadata = streamed.metadata; }
    else { let data; try { data = await response.json(); } catch (error) { throw detailedError("CloudCIX returned invalid response JSON.", { phase: "response", contentType: type, cause: error.message }); } if (data.error) throw detailedError("CloudCIX error: " + errorMessage(data.error), { phase: "response", payload: data.error }); content = chunk(data); api.lastMetadata = normalizeMetadata(data.guiden); }
    if (!String(content).trim()) throw detailedError("CloudCIX returned an empty response.", { phase: "response", contentType: type });
    return content;
  };
  api.test = async function (key, signal) { const response = await fetch(api.url, { method: "POST", headers: { "Content-Type": "application/json", Authorization: "Bearer " + key }, body: JSON.stringify({ model: api.model, stream: false, max_tokens: 4, messages: [{ role: "user", content: "Reply with OK." }] }), signal }); if (!response.ok) throw detailedError("Connection failed (HTTP " + response.status + ").", { phase: "test", status: response.status }); return true; };
  G.api = api;
})(window);
