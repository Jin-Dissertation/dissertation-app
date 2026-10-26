/* Open-module controller. Deliberately independent of participant authentication,
 * event logging, feedback submission, and D1 research storage. Only the public
 * curriculum JSON and approved static training media are fetched. */
(() => {
  "use strict";
  const KEY = "aqg-open-training-v2-card";
  const ANSWERS_KEY = "aqg-open-training-v2-answers";
  const MEDIA = "https://dissertation-study-api.professor-jin.workers.dev/v1/training/media/2026-04-27-strict-content-a/videos/";
  const SWAP = {
    "05a-guided-practice-participant-v2.mp4": "05b-guided-practice-nonparticipant-v2.mp4",
    "06a-wrapping-up-participant-v2.mp4": "06b-wrapping-up-nonparticipant-v2.mp4"
  };
  const $ = id => document.getElementById(id);
  const elem = (tag, cls, text) => {
    const el = document.createElement(tag);
    if (cls) el.className = cls;
    if (text !== undefined) el.textContent = text;
    return el;
  };
  const add = (parent, node) => { parent.appendChild(node); return node; };
  const options = value => String(value || "").split(/\s*\|\|\s*/).map(x => x.trim()).filter(Boolean);
  const shuffle = input => {
    const a = [...input];
    for (let i = a.length - 1; i > 0; --i) {
      const j = Math.floor(Math.random() * (i + 1));
      [a[i], a[j]] = [a[j], a[i]];
    }
    return a;
  };
  const readAnswers = () => {
    try { return JSON.parse(localStorage.getItem(ANSWERS_KEY) || "{}"); }
    catch (_) { return {}; }
  };
  let rows = [], current = 0, answers = readAnswers(), done = false;
  const save = () => {
    localStorage.setItem(KEY, String(current));
    localStorage.setItem(ANSWERS_KEY, JSON.stringify(answers));
  };
  const listen = (el, event, fn) => { el.addEventListener(event, fn); return el; };
  const button = (text, cls, fn) => listen(elem("button", cls, text), "click", fn);
  const normalize = s => String(s || "").trim().toLocaleLowerCase();
  const sectionIds = () => [...new Set(rows.map(x => x.section_id))];
  const parts = text => String(text || "").split(/\n\s*\n/).map(x => x.trim()).filter(Boolean);

  function inline(parent, text) {
    const tokens = String(text).split(/(https?:\/\/[^\s]+|\*\*[^*]+\*\*|\*[^*]+\*)/g);
    for (const piece of tokens) {
      if (/^https?:\/\//.test(piece)) {
        const a = add(parent, elem("a", "", piece)); a.href = piece;
        a.target = "_blank"; a.rel = "noopener noreferrer";
      } else if (/^\*\*[^*]+\*\*$/.test(piece)) {
        add(parent, elem("strong", "", piece.slice(2, -2)));
      } else if (/^\*[^*]+\*$/.test(piece)) {
        add(parent, elem("em", "", piece.slice(1, -1)));
      } else add(parent, document.createTextNode(piece));
    }
  }
  function formatted(parent, content) {
    const chunks = parts(content);
    for (const chunk of chunks) {
      const lines = chunk.split(/\n/).map(x => x.trim()).filter(Boolean);
      if (!lines.length) continue;
      const ordered = lines.every(x => /^\d+\.\s/.test(x));
      const bulleted = lines.every(x => /^--\s/.test(x));
      if (ordered || bulleted) {
        const list = add(parent, elem(ordered ? "ol" : "ul"));
        for (const line of lines) {
          const li = add(list, elem("li"));
          inline(li, ordered ? line.replace(/^\d+\.\s*/, "") : line.replace(/^--\s*/, ""));
        }
      } else {
        const p = add(parent, elem("p"));
        lines.forEach((line, index) => {
          if (index) add(p, document.createTextNode(" "));
          inline(p, line);
        });
      }
    }
  }
  function adapt(row) {
    const r = { ...row };
    if (r.header === "Welcome to this module") {
      r.body = String(r.body || "").replace(/\s*You can leave this module at any time; when you open it again, it will resume where you left off\.\s*$/, "");
    }
    if (r.header === "Begin your practice session") {
      r.body = "Open the app in training mode by clicking the link below:\n\n" +
        "https://jin-dissertation.github.io/dissertation-app/np/npindex.html?mode=training\n\n" +
        "Keep your AI chatbot open in another tab or window. Spend about 15 minutes trying to create questions on a topic you teach. These are practice goals, not graded requirements:\n\n" +
        "1. Try to accept at least 5 questions.\n" +
        "2. Try each main question-refinement prompt at least once.\n" +
        "3. Try each audit prompt at least once.\n\n" +
        "The blue training tips in the app will provide helpful tips. About 15 minutes after you begin, the app will alert you and ask you to return to this module. When you're ready, come back here to wrap things up.";
    }
    return r;
  }
  function next(index) {
    if (index < current) return;
    if (index === rows.length - 1) {
      done = true;
      localStorage.removeItem(KEY);
      localStorage.removeItem(ANSWERS_KEY);
      $("sections-container").hidden = true;
      $("completion").style.display = "block";
      $("overall-completed").textContent = "Completed: " + sectionIds().length;
      $("progress-fill").style.width = "100%";
      $("section-counter").textContent = "Training complete";
      window.scrollTo(0, 0);
      return;
    }
    current = index + 1;
    save();
    render();
    const block = document.getElementById("section-" + current);
    if (block) block.scrollIntoView({ behavior:"smooth", block:"start" });
  }
  function labelCard(i, type) {
    const block = elem("div", "section-block revealed"); block.id = "section-" + i;
    const head = add(block, elem("div", "section-label"));
    add(head, elem("span", "section-num", "Card " + (i + 1)));
    add(head, elem("div", "section-rule"));
    add(head, elem("span", "section-type-badge", type));
    return block;
  }
  function continueRow(block, i, visible = true) {
    const wrap = add(block, elem("div", "continue-row" + (visible ? " show" : "")));
    const b = button(i === rows.length - 1 ? "Finish training" : "Continue", "action-btn continue-btn", () => next(i));
    b.type = "button"; add(wrap, b);
    return wrap;
  }
  function image(parent, row) {
    const mobile = window.matchMedia("(max-width: 800px)").matches;
    const url = (mobile && row.mobile_image_url) || row.media_url || "";
    if (!url.startsWith("images/")) return;
    const img = add(parent, elem("img", "open-image"));
    img.src = url; img.alt = row.header || "Training illustration"; img.loading = "lazy";
    listen(img, "click", () => {
      $("open-image-large").src = img.src;
      $("open-image-zoom").classList.add("show");
    });
  }
  function textCard(row, i) {
    const isGuide = String(row.settings || "").includes("beginnerHelp");
    const block = labelCard(i, isGuide ? "Guide" : "Reading");
    if (isGuide) {
      add(block, window.AQGBeginnerHelp.renderCard(row));
    } else {
      const card = add(block, elem("div", "text-card"));
      if (row.header) add(card, elem("h2", "content-header", row.header));
      const body = add(card, elem("div", "content-body open-text-body"));
      formatted(body, row.body);
      image(body, row);
    }
    continueRow(block, i); return block;
  }
  function videoCard(row, i) {
    const block = labelCard(i, "Video");
    const card = add(block, elem("div", "video-card"));
    const wrap = add(card, elem("div", "video-wrapper"));
    const filename = String(row.video_name || "").replace(/^videos\//, "").split("/").pop();
    const video = add(wrap, elem("video"));
    video.controls = true; video.playsInline = true; video.preload = "metadata";
    const actual = SWAP[filename] || filename;
    const source = add(video, elem("source")); source.src = MEDIA + encodeURIComponent(actual); source.type = "video/mp4";
    const track = add(video, elem("track")); track.kind = "subtitles"; track.srclang = "en"; track.label = "English";
    track.src = MEDIA + encodeURIComponent(actual.replace(/\.mp4$/i, ".vtt"));
    listen(track, "error", () => track.remove());
    const footer = add(card, elem("div", "video-footer"));
    add(footer, elem("span", "video-footer-label", "Watch the video above, then click when you're done."));
    add(footer, button("Done watching", "action-btn", () => next(i)));
    return block;
  }
  function choiceCard(row, i) {
    const multi = row.type === "sata";
    const block = labelCard(i, "Knowledge Check");
    const card = add(block, elem("div", "quiz-card"));
    const header = add(card, elem("div", "quiz-header"));
    add(header, elem("p", "quiz-header-label", multi ? "Select All That Apply" : "Knowledge Check"));
    const question = add(card, elem("div", "question-block active"));
    add(question, elem("p", "question-stem", row.stem));
    if (multi) add(question, elem("p", "quiz-header-sub", "Select all that apply."));
    const correct = multi ? options(row.correct_blob) : [row.correct_blob];
    const state = answers[i] || { order:shuffle([...correct, ...options(row.distractors_blob)]), selected:[], checked:false };
    answers[i] = state;
    if (!Array.isArray(state.order) || !state.order.length) state.order = shuffle([...correct,...options(row.distractors_blob)]);
    if (!Array.isArray(state.selected)) state.selected = [];
    const list = add(question, elem("ul", "options-list"));
    const optionButtons = [];
    const select = value => {
      if (state.checked) return;
      if (multi) {
        state.selected = state.selected.includes(value) ? state.selected.filter(x => x !== value) : [...state.selected, value];
      } else state.selected = [value];
      save();
      if (multi) optionButtons.forEach(([b,v]) => b.classList.toggle("selected", state.selected.includes(v)));
      else check();
    };
    state.order.forEach((value, order) => {
      const li = add(list, elem("li"));
      const b = button("", "option-btn", () => select(value));
      add(b, elem("span", "option-letter", "ABCDEF"[order] || String(order + 1)));
      add(b, elem("span", "option-text", value));
      add(li, b); optionButtons.push([b, value]);
    });
    if (!multi) {
      const li = add(list, elem("li"));
      const b = button("", "option-btn idk-option", () => select("__idk__"));
      add(b, elem("span", "option-letter", "?"));
      add(b, elem("span", "option-text", "I don't know"));
      add(li, b); optionButtons.push([b, "__idk__"]);
    }
    const fb = add(question, elem("div", "feedback-box"));
    const actions = add(question, elem("div", "feedback-actions"));
    if (multi && !state.checked) add(actions, button("Check Answer", "feedback-next-btn show", check));
    const rowContinue = continueRow(block, i, !!state.checked);
    function check() {
      if (state.checked) return;
      state.checked = true;
      const picked = [...state.selected].sort(), expected = [...correct].sort();
      state.isCorrect = picked.length === expected.length && picked.every((x,k) => x === expected[k]);
      save();
      show();
      rowContinue.classList.add("show");
    }
    function show() {
      optionButtons.forEach(([b,v]) => {
        b.disabled = !!state.checked;
        b.classList.remove("selected", "correct", "incorrect", "missed-correct", "dimmed");
        if (!state.checked) {
          b.classList.toggle("selected", state.selected.includes(v));
        } else if (state.selected.includes(v) && correct.includes(v)) b.classList.add("correct");
        else if (state.selected.includes(v)) b.classList.add("incorrect");
        else if (correct.includes(v)) b.classList.add("missed-correct");
        else b.classList.add("dimmed");
      });
      if (state.checked) {
        fb.replaceChildren();
        fb.className = "feedback-box show " + (state.isCorrect ? "correct-fb" : "incorrect-fb");
        add(fb, elem("span", "feedback-label", state.isCorrect ? "✓ Correct" : "✗ Not quite"));
        const message = state.selected.includes("__idk__") ? "Here's the explanation: " : "";
        add(fb, elem("span", "feedback-text", message + (row.feedback || "")));
        actions.replaceChildren();
      }
    }
    show(); return block;
  }
  function matchingCard(row, i) {
    const block = labelCard(i, "Knowledge Check");
    const card = add(block, elem("div", "categorize-card"));
    add(card, elem("p", "question-stem", row.stem));
    const mobile = window.matchMedia("(max-width: 800px)").matches;
    add(card, elem("p", "categorize-helper", mobile ?
      "Tap an item to select it, then tap a matching option to place it. Leave unmatched distractors in the item pool." :
      "Drag each item to its correct match. Leave unmatched distractors in the item pool."));
    const pairs = String(row.correct_blob || "").split("\n").map(x => x.split("::")).filter(x => x.length === 2);
    const answersByItem = new Map(pairs.filter(x => x[0].trim() !== "N/A").map(([left,right]) => [right,left]));
    const names = shuffle(pairs.map(x => x[1]));
    const categories = pairs.filter(x => x[0].trim() !== "N/A").map(x => x[0]);
    const state = answers[i] || { order:names, placements:{}, checked:false };
    answers[i] = state;
    if (!Array.isArray(state.order)) state.order = names;
    if (!state.placements) state.placements = {};
    const layout = add(card, elem("div", "categorize-layout"));
    const pool = add(layout, elem("div", "categorize-pool")); pool.dataset.category = "pool";
    add(pool, elem("p", "categorize-pool-title", "Items to match"));
    const poolItems = add(pool, elem("div", "categorize-items"));
    const hint = add(pool, elem("p", "categorize-mobile-hint", "No item selected."));
    if (!mobile) hint.hidden = true;
    const zones = add(layout, elem("div", "categorize-zones"));
    const targets = [pool], items = [];
    categories.forEach(cat => {
      const zone = add(zones, elem("div", "categorize-zone")); zone.dataset.category = cat;
      add(zone, elem("p", "categorize-zone-title", cat));
      add(zone, elem("div", "categorize-items")); targets.push(zone);
    });
    let selected = null, dragged = null;
    const locate = key => targets.find(t => t.dataset.category === key) || pool;
    const placement = item => item.closest("[data-category]")?.dataset.category || "pool";
    const move = (item, target) => {
      if (state.checked) return;
      add(target.querySelector(".categorize-items"), item);
      state.placements[item.dataset.item] = placement(item); save();
    };
    state.order.forEach(name => {
      const el = elem("button", "categorize-item", name); el.type = "button"; el.dataset.item = name;
      add(poolItems, el); items.push(el);
      if (mobile) listen(el, "click", e => {
        e.stopPropagation(); if (state.checked) return;
        if (selected) selected.classList.remove("selected");
        selected = selected === el ? null : el;
        if (selected) selected.classList.add("selected");
        hint.textContent = selected ? "Selected: " + selected.dataset.item : "No item selected.";
      });
      else {
        el.draggable = !state.checked;
        listen(el, "dragstart", e => { dragged = el; e.dataTransfer?.setData("text/plain", name); });
        listen(el, "dragend", () => { dragged = null; targets.forEach(t => t.classList.remove("categorize-drop-target-active")); });
      }
    });
    targets.forEach(target => {
      if (mobile) listen(target, "click", e => {
        if (e.target.closest(".categorize-item") || !selected || state.checked) return;
        move(selected, target);
        selected.classList.remove("selected"); selected = null;
        hint.textContent = "No item selected.";
      });
      else {
        listen(target, "dragover", e => { e.preventDefault(); target.classList.add("categorize-drop-target-active"); });
        listen(target, "dragleave", () => target.classList.remove("categorize-drop-target-active"));
        listen(target, "drop", e => {
          e.preventDefault(); target.classList.remove("categorize-drop-target-active");
          if (dragged) move(dragged, target);
        });
      }
    });
    // Restore browser-only placements without re-sending them anywhere.
    items.forEach(item => {
      const target = locate(state.placements[item.dataset.item] || "pool");
      add(target.querySelector(".categorize-items"), item);
    });
    const controls = add(card, elem("div", "categorize-controls"));
    const reset = button("Reset placements", "action-btn secondary", () => {
      if (state.checked) return;
      state.placements = {}; items.forEach(item => {
        item.classList.remove("selected");
        add(poolItems, item);
      }); selected = null; hint.textContent = "No item selected."; save();
    });
    const checkButton = button("Check Answer", "action-btn", () => {
      if (state.checked) return;
      state.checked = true;
      state.isCorrect = items.every(item => normalize(placement(item)) === normalize(answersByItem.get(item.dataset.item) || "pool"));
      save(); show();
    });
    add(controls, reset); add(controls, checkButton);
    const feedback = add(card, elem("div", "feedback-box"));
    const continuation = continueRow(block, i, !!state.checked);
    function show() {
      if (!state.checked) return;
      reset.disabled = true; checkButton.disabled = true;
      items.forEach(item => {
        item.draggable = false;
        item.classList.add(normalize(placement(item)) === normalize(answersByItem.get(item.dataset.item) || "pool") ? "correct" : "incorrect");
      });
      feedback.className = "feedback-box show " + (state.isCorrect ? "correct-fb" : "incorrect-fb");
      feedback.replaceChildren();
      add(feedback, elem("span", "feedback-label", state.isCorrect ? "Correct ✓" : "Incorrect ✗"));
      add(feedback, elem("span", "feedback-text", row.feedback || ""));
      continuation.classList.add("show");
    }
    show(); return block;
  }
  function makeCard(row, i) {
    if (row.type === "video") return videoCard(row, i);
    if (row.type === "sba" || row.type === "sata") return choiceCard(row, i);
    if (row.type === "matching" || row.type === "categorization") return matchingCard(row, i);
    return textCard(row, i);
  }
  function render() {
    done = false;
    $("completion").style.display = "none";
    $("sections-container").hidden = false;
    const container = $("sections-container");
    container.replaceChildren();
    const ids = sectionIds();
    $("overall-total").textContent = "Sections: " + ids.length;
    const currentSection = rows[current].section_id;
    const activeSection = Math.max(0, ids.indexOf(currentSection));
    $("overall-current").textContent = "Current: " + (activeSection + 1) + " / " + ids.length;
    $("overall-completed").textContent = "Completed: " + ids.filter(id => {
      const last = rows.findLastIndex(x => x.section_id === id);
      return last < current;
    }).length;
    $("progress-fill").style.width = (100 * (current + 1) / rows.length) + "%";
    $("section-counter").textContent = "Card " + (current + 1) + " of " + rows.length;
    ids.forEach((id, sectionIndex) => {
      const cardIds = rows.map((x,i) => x.section_id === id ? i : -1).filter(i => i >= 0);
      const first = cardIds[0], last = cardIds[cardIds.length - 1];
      const locked = first > current;
      const acc = add(container, elem("section", "section-accordion" +
        (locked ? " locked collapsed" : sectionIndex === activeSection ? "" : " collapsed")));
      acc.dataset.sectionIdx = String(sectionIndex);
      const toggle = add(acc, elem("button", "section-toggle")); toggle.type = "button"; toggle.disabled = locked;
      const wrap = add(toggle, elem("span", "section-title-wrap"));
      add(wrap, elem("strong", "", "Section " + (sectionIndex + 1) + ": " + (rows[first].section_title || "Training")));
      add(wrap, elem("span", "section-status", last < current ? "Completed" : locked ? "Not started" : "In progress"));
      add(toggle, elem("span", "section-chevron", "▾"));
      listen(toggle, "click", () => acc.classList.toggle("collapsed"));
      const content = add(acc, elem("div", "section-content"));
      cardIds.filter(i => i <= current).forEach(i => add(content, makeCard(adapt(rows[i]), i)));
    });
    $("last-updated").textContent = document.lastModified || "Current";
  }
  listen($("open-image-close"), "click", () => $("open-image-zoom").classList.remove("show"));
  listen($("open-image-zoom"), "click", e => { if (e.target.id === "open-image-zoom") e.currentTarget.classList.remove("show"); });
  listen(document, "keydown", e => { if (e.key === "Escape") $("open-image-zoom").classList.remove("show"); });
  listen($("btn-save-exit"), "click", () => { save(); window.location.href = "../np/npindex.html"; });
  listen($("btn-report-issue"), "click", () => { window.location.href = "mailto:tjin@crimson.ua.edu?subject=Nonparticipant%20training%20issue"; });
  fetch("module-content-v2.json", { cache:"no-store" })
    .then(response => { if (!response.ok) throw Error("HTTP " + response.status); return response.json(); })
    .then(data => {
      if (!Array.isArray(data.rows) || data.rows.length !== 24) throw Error("Unexpected curriculum length");
      rows = data.rows;
      const stored = Number(localStorage.getItem(KEY));
      current = Number.isInteger(stored) && stored >= 0 && stored < rows.length ? stored : 0;
      render();
    })
    .catch(error => {
      $("open-error").classList.add("show");
      $("open-error").textContent = "Unable to load the training module. " + error.message;
    });
})();
