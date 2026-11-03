(function (global) {
  "use strict";
  const G = global.Guiden,
    C = G.core;
  const jsonOnly = "Return JSON only. Do not use Markdown or code fences.";
  const keep = "Keep the meaning, facts, numbers, and tense: a plan that says 'we will' still says what will happen. ";
  function source(snapshot, options) {
    const titles = snapshot.deckTitles || [],
      outline =
        snapshot.actionId === "slides.new" && titles.length
          ? "\n\nSlides in this presentation (titles):\n" +
            titles.map((t, i) => i + 1 + ". " + (t || "(no title)")).join("\n") +
            "\nNew slides go after slide " +
            (snapshot.slideIndex + 1) +
            "."
          : "";
    return (
      "Target: " +
      snapshot.targetLabel +
      "\nTone: " +
      options.tone +
      "\nLength: " +
      options.length +
      "\n\nIndexed slide content:\n" +
      snapshot.text +
      outline
    );
  }
  function request(system, snapshot, instruction, options) {
    return C.groundedMessages(system, source(snapshot, options), instruction);
  }
  function answer(content) {
    const d = C.extractJson(content);
    if (typeof d.answer !== "string") throw new Error("Guiden returned an invalid slide answer.");
    return { answer: d.answer, sources: C.safeSources(d.sources) };
  }
  function selectedReplacement(content) {
    const d = C.extractJson(content);
    if (typeof d.replacementText !== "string") throw new Error("Guiden did not return replacement slide text.");
    return { replacementText: d.replacementText };
  }
  function shapeUpdates(content, snapshot) {
    const d = C.extractJson(content);
    return { narrative: String(d.narrative || ""), updates: C.validateShapeUpdates(d.updates, snapshot) };
  }
  function register(action) {
    G.registerAction(Object.assign({ editor: "slide", supportedTargets: ["selected_text", "current_slide"] }, action));
  }
  register({
    id: "slides.ask",
    label: "Ask",
    description: "Ask about the current slide",
    mutates: false,
    defaultInstruction: "Answer my question about this slide.",
    placeholder: "What would you like to know about this slide?",
    buildRequest(snapshot, instruction, options) {
      return request(
        'You are Guiden inside Euro-Office Slides. Answer without changing the presentation. Return {"answer":"...","sources":[]}. ' +
          jsonOnly,
        snapshot,
        instruction,
        options
      );
    },
    parseResponse: answer,
    buildPreview(result) {
      return { kind: "answer", answer: result.answer, sources: result.sources };
    },
    revalidateTarget: async () => ({ ok: true }),
    apply: async () => {
      throw new Error("This action is read-only.");
    }
  });
  function edit(id, label, description, defaultInstruction, selectedOnly) {
    register({
      id,
      label,
      description,
      mutates: true,
      supportedTargets: selectedOnly ? ["selected_text"] : ["selected_text", "current_slide"],
      defaultInstruction,
      placeholder: "Add any preferences, or leave blank.",
      buildRequest(snapshot, instruction, options) {
        const selected = snapshot.target === "selected_text",
          system = selected
            ? "Improve only the selected slide text. " +
              keep +
              'Do not make it longer unless asked. Return {"replacementText":"..."}. ' +
              jsonOnly
            : "Update only identified text areas. Preserve unmentioned shapes and all non-text objects. " +
              keep +
              'Return {"narrative":"...","updates":[{"shapeIndex":0,"before":"exact original","after":"new text"}]}. shapeIndex must match an indexed area. ' +
              jsonOnly;
        return request(system, snapshot, instruction || defaultInstruction, options);
      },
      parseResponse(content, snapshot) {
        if (selectedOnly && snapshot.target !== "selected_text")
          throw new Error("Select text on the slide before using this action.");
        return snapshot.target === "selected_text" ? selectedReplacement(content) : shapeUpdates(content, snapshot);
      },
      buildPreview(result, snapshot) {
        return snapshot.target === "selected_text"
          ? { kind: "text-diff", before: snapshot.selectedText, after: result.replacementText }
          : { kind: "shape-updates", narrative: result.narrative, updates: result.updates };
      },
      revalidateTarget: snapshot => G.host.revalidate(snapshot),
      apply(result, snapshot) {
        return snapshot.target === "selected_text"
          ? G.host.applySelectedSlideText(result.replacementText, snapshot)
          : G.host.applyShapeUpdates(result.updates, snapshot);
      }
    });
  }
  edit(
    "slides.improve",
    "Improve selected text",
    "Polish only selected slide text",
    "Improve clarity and impact while keeping it slide-ready.",
    true
  );
  edit(
    "slides.shorten",
    "Shorten current slide",
    "Reduce text while preserving key points",
    "Shorten the slide substantially and keep the important message.",
    false
  );
  edit(
    "slides.rewrite",
    "Rewrite current slide",
    "Improve identified text areas safely",
    "Rewrite the current slide for clarity and impact.",
    false
  );
  const MAX_NEW_SLIDES = 8,
    LAYOUTS = ["title", "section", "content", "two"];
  // Slide text, not prose: no bullet characters or numbering at the start, as the layout draws its own bullets.
  function slideLine(value, limit) {
    return String(value == null ? "" : value)
      .replace(/\s+/g, " ")
      .replace(/^\s*(?:[•●▪◦‣*–—-]|\d{1,2}[.)])\s+/, "")
      .trim()
      .slice(0, limit || 200);
  }
  function slideLines(list, max) {
    return (Array.isArray(list) ? list : [])
      .map(x => slideLine(x))
      .filter(Boolean)
      .slice(0, max);
  }
  // Checks the model's slides and puts them in the shape host.createSlides expects. Accepts the older single-slide
  // shape ({title, bullets}) too.
  function newSlides(d) {
    const list = Array.isArray(d.slides) ? d.slides : d.title ? [d] : [];
    if (!list.length) throw new Error("Guiden did not return a valid slide.");
    return list.slice(0, MAX_NEW_SLIDES).map(raw => {
      const item = raw || {},
        title = slideLine(item.title, 120);
      let layout = LAYOUTS.indexOf(item.layout) >= 0 ? item.layout : "content";
      if (!title) throw new Error("Guiden returned a slide without a title.");
      const slide = { layout, title, subtitle: slideLine(item.subtitle, 200), bullets: [], left: null, right: null };
      if (layout === "two") {
        const side = part => ({
          heading: slideLine(part && part.heading, 60),
          bullets: slideLines(part && part.bullets, 6)
        });
        slide.left = side(item.left);
        slide.right = side(item.right);
        if (!slide.left.bullets.length || !slide.right.bullets.length) {
          // Not really two columns: keep whichever bullets came back as one list.
          slide.layout = layout = "content";
          slide.bullets = slide.left.bullets.concat(slide.right.bullets, slideLines(item.bullets, 6)).slice(0, 6);
          slide.left = slide.right = null;
        }
      } else if (layout === "content") slide.bullets = slideLines(item.bullets, 6);
      if (layout === "content" && !slide.bullets.length) {
        if (slide.subtitle) slide.layout = "section";
        else throw new Error("Guiden returned a slide without any bullets.");
      }
      return slide;
    });
  }
  register({
    id: "slides.new",
    label: "Create new slides",
    description: "Add slides in this deck's own design",
    mutates: true,
    defaultInstruction: "Create one useful new slide.",
    placeholder: "Describe the slide or slides you want to add.",
    buildRequest(snapshot, instruction, options) {
      return request(
        "Create new slides to add after the current slide. They are drawn with the presentation's own layouts, fonts and colours, so return only the words. " +
          "Make one slide unless the request asks for more, and never more than " +
          MAX_NEW_SLIDES +
          ". " +
          'Choose a layout for each slide: "title" opens a talk or deck (title and a one-line subtitle); "section" starts a part of the talk (title and an optional one-line subtitle); ' +
          '"content" is a title and 3 to 6 bullets; "two" compares or pairs two things (title, and two lists of 2 to 5 bullets with optional short headings). ' +
          "An introduction or opening slide for a talk is usually a title slide. " +
          "Write slide text, not prose: titles under 8 words, subtitles under 15 words, bullets under 12 words, no full stop at the end of a bullet, no bullet characters or numbering, no Markdown. " +
          "Use only facts from the request, the indexed slide content and any supplied retrieval context. Do not invent figures, dates, names or claims. " +
          'Return {"slides":[{"layout":"content","title":"...","subtitle":"","bullets":["..."]},{"layout":"two","title":"...","left":{"heading":"...","bullets":["..."]},"right":{"heading":"...","bullets":["..."]}}]}. ' +
          jsonOnly,
        snapshot,
        instruction,
        options
      );
    },
    parseResponse(content) {
      return { slides: newSlides(C.extractJson(content)) };
    },
    buildPreview(result) {
      return { kind: "new-slides", slides: result.slides };
    },
    revalidateTarget: snapshot => G.host.revalidate(snapshot),
    apply: result => G.host.createSlides(result.slides)
  });
  // Ready-made designs. Each sets the deck's theme colours and fonts (so every slide, layout and chart follows it), the
  // background, and a few accents at the slide edges. They are fixed rather than invented by the model, so text always
  // contrasts with its background, and every font is one Euro-Office ships.
  const DESIGNS = [
    {
      id: "harbour",
      name: "Harbour",
      description: "clean blue and white; technical, corporate and product decks",
      fonts: { heading: "Open Sans Semibold", body: "Open Sans" },
      colors: {
        dk1: "#1B2733",
        lt1: "#FFFFFF",
        dk2: "#0B3D63",
        lt2: "#E9F1F7",
        accent1: "#1F7AE0",
        accent2: "#14A3B8",
        accent3: "#2BB673",
        accent4: "#F2A900",
        accent5: "#7A5AF8",
        accent6: "#E8590C",
        hlink: "#1F7AE0",
        folHlink: "#7A5AF8"
      }
    },
    {
      id: "midnight",
      name: "Midnight",
      description: "dark navy with bright accents; keynotes, launches and tech talks",
      fonts: { heading: "Ubuntu", body: "Ubuntu" },
      colors: {
        dk1: "#EEF2F7",
        lt1: "#0F1B2B",
        dk2: "#9FB3C8",
        lt2: "#1B2A3F",
        accent1: "#4FB3FF",
        accent2: "#7BE0C3",
        accent3: "#FFC857",
        accent4: "#FF7A7A",
        accent5: "#B69CFF",
        accent6: "#8FD14F",
        hlink: "#7FC8FF",
        folHlink: "#C9B6FF"
      }
    },
    {
      id: "paper",
      name: "Paper",
      description: "warm cream with serif headings; reports, research, history and the humanities",
      fonts: { heading: "Caladea", body: "Carlito" },
      colors: {
        dk1: "#2B2A27",
        lt1: "#FBF8F2",
        dk2: "#5B3A29",
        lt2: "#F1E9DC",
        accent1: "#B5522B",
        accent2: "#3F7D6E",
        accent3: "#D9A441",
        accent4: "#6B5B95",
        accent5: "#8C6D46",
        accent6: "#2F5C8A",
        hlink: "#2F5C8A",
        folHlink: "#6B5B95"
      }
    },
    {
      id: "forest",
      name: "Forest",
      description: "calm greens; sustainability, health, nature and public sector",
      fonts: { heading: "Open Sans Semibold", body: "Open Sans" },
      colors: {
        dk1: "#1E2B24",
        lt1: "#F7FAF7",
        dk2: "#1F4D3A",
        lt2: "#E3EFE6",
        accent1: "#2E7D5B",
        accent2: "#88B04B",
        accent3: "#F4A259",
        accent4: "#5B8E7D",
        accent5: "#BC4B51",
        accent6: "#3D5A80",
        hlink: "#2E7D5B",
        folHlink: "#5B8E7D"
      }
    },
    {
      id: "mono",
      name: "Mono",
      description: "black and white with a red accent; minimal, bold and direct",
      fonts: { heading: "Liberation Sans", body: "Liberation Sans" },
      colors: {
        dk1: "#111111",
        lt1: "#FFFFFF",
        dk2: "#333333",
        lt2: "#F2F2F2",
        accent1: "#E63946",
        accent2: "#1D3557",
        accent3: "#457B9D",
        accent4: "#A8A8A8",
        accent5: "#F4A261",
        accent6: "#2A9D8F",
        hlink: "#1D3557",
        folHlink: "#457B9D"
      }
    },
    {
      id: "sunrise",
      name: "Sunrise",
      description: "warm orange and plum; marketing, education, events and community",
      fonts: { heading: "Ubuntu", body: "Open Sans" },
      colors: {
        dk1: "#2D1E2F",
        lt1: "#FFFDF9",
        dk2: "#5E2750",
        lt2: "#FCEFE6",
        accent1: "#F26B38",
        accent2: "#7B4B94",
        accent3: "#F2A541",
        accent4: "#E94F64",
        accent5: "#2BAE9E",
        accent6: "#3C6E9E",
        hlink: "#C2410C",
        folHlink: "#7B4B94"
      }
    }
  ];
  G.designs = DESIGNS;
  function designById(id) {
    return DESIGNS.filter(d => d.id === String(id || "").toLowerCase())[0] || null;
  }
  register({
    id: "slides.design",
    label: "Design deck",
    description: "Give the whole deck a ready-made look",
    mutates: true,
    defaultInstruction: "Choose the design that best suits this presentation.",
    placeholder: "Describe the look you want, or leave this for Guiden to choose.",
    buildRequest(snapshot, instruction, options) {
      return request(
        "Choose a design for the whole presentation from this list. Each sets the theme colours, fonts and background of every slide.\n" +
          DESIGNS.map(d => "- " + d.id + ": " + d.description).join("\n") +
          "\nFollow the request when it names a colour, mood or design; otherwise choose what suits the subject and audience. " +
          'Return {"design":"one id from the list","reason":"one short sentence saying why it suits this deck"}. ' +
          jsonOnly,
        snapshot,
        instruction,
        options
      );
    },
    parseResponse(content) {
      const d = C.extractJson(content),
        design = designById(d.design);
      if (!design) throw new Error("Guiden did not choose one of the designs.");
      return { design: design.id, reason: slideLine(d.reason, 240) };
    },
    buildPreview(result) {
      return { kind: "design", design: result.design, reason: result.reason, designs: DESIGNS };
    },
    revalidateTarget: async () => ({ ok: true }),
    apply: result => G.host.applyDesign(designById(result.design))
  });
  register({
    id: "slides.outline",
    label: "Create slide outline",
    description: "Preview a presentation structure",
    mutates: false,
    defaultInstruction: "Create a concise slide outline from this content.",
    placeholder: "What should the outline cover?",
    buildRequest(snapshot, instruction, options) {
      return request(
        'Create a presentation outline. Return {"answer":"short overview","slides":[{"title":"...","bullets":["..."]}],"sources":[]}. Return no more than ' +
          options.slideCount +
          " slides. " +
          jsonOnly,
        snapshot,
        instruction,
        options
      );
    },
    parseResponse(content) {
      const d = C.extractJson(content);
      if (!Array.isArray(d.slides)) throw new Error("Guiden did not return a slide outline.");
      return { answer: String(d.answer || "Slide outline"), slides: d.slides, sources: C.safeSources(d.sources) };
    },
    buildPreview(result) {
      return { kind: "outline", answer: result.answer, slides: result.slides, sources: result.sources };
    },
    revalidateTarget: async () => ({ ok: true }),
    apply: async () => {
      throw new Error("The outline is read-only. Choose Create new slide to add a slide.");
    }
  });
  register({
    id: "slides.notes",
    label: "Generate speaker notes",
    description: "Add notes without changing slide content",
    mutates: true,
    defaultInstruction: "Create useful speaker notes for this slide.",
    placeholder: "What should the speaker notes emphasise?",
    buildRequest(snapshot, instruction, options) {
      return request(
        "Create speaker notes for the current slide: what the presenter should say about each point, in plain sentences they can read aloud. " +
          "No Markdown, HTML, or headings. Follow the requested length. " +
          "Talk through only what the slide says. Do not add deadlines, dates, figures, quality claims, promises, or outcomes that are not on the slide or in supplied retrieval context. " +
          'Return {"notes":"plain text notes"}. ' +
          jsonOnly,
        snapshot,
        instruction,
        options
      );
    },
    parseResponse(content) {
      const d = C.extractJson(content);
      if (typeof d.notes !== "string" || !d.notes.trim()) throw new Error("Guiden did not return speaker notes.");
      return { notes: d.notes.trim() };
    },
    buildPreview(result) {
      return { kind: "notes", notes: result.notes };
    },
    revalidateTarget: snapshot => G.host.revalidate(snapshot),
    apply: (result, snapshot) => G.host.addNotes(result.notes, snapshot)
  });
})(window);
