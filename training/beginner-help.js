/* Optional learning support, not a scored assessment or research response.
 * Shared by participant and nonparticipant modules. No network or storage.
 */
(() => {
  'use strict';
  let activeOverlay = null;
  const e = (tag, cls, text) => {
    const node = document.createElement(tag);
    if (cls) node.className = cls;
    if (text !== undefined) node.textContent = text;
    return node;
  };
  const add = (target, node) => { target.appendChild(node); return node; };
  const paragraph = (parent, text) => add(parent, e('p', '', text));
  function close() {
    if (!activeOverlay) return;
    const { overlay, focusBefore, keydown } = activeOverlay;
    document.removeEventListener('keydown', keydown);
    overlay.remove();
    document.body.style.overflow = activeOverlay.overflow;
    activeOverlay = null;
    if (focusBefore?.isConnected) focusBefore.focus();
  }
  function heading(parent, text) { return add(parent, e('h3', '', text)); }
  function chatbotContent(parent) {
    const section = add(parent, e('section', 'aqg-basics-section'));
    heading(section, 'What is an AI chatbot?');
    paragraph(section, 'Artificial intelligence (AI) is a broad term for computer systems that do things we normally associate with human intelligence. AI can take many forms, but much of the recent excitement is about AI chatbots, such as ChatGPT, Gemini, Copilot, and Claude. A chatbot works a lot like a text-message conversation: you send it a message, and it sends a reply that can sound as though a person wrote it. But you are communicating with software, not another person.');
    heading(section, 'How do I use an AI chatbot?');
    paragraph(section, "Think of it as texting someone who is always ready to try to respond right away. Open the chatbot, type a message, and press Send. If you ask a question, it will try to answer. If you give instructions, it will try to follow them. If you share an observation, it will try to react to what you said. You can send another message to keep the conversation going.");
    paragraph(section, 'Each message you send to the chatbot is called a prompt.');
    paragraph(section, 'An AI chatbot can make mistakes. Check clinical facts and never paste protected patient, student, confidential, or restricted information into a chatbot.');
    paragraph(section, 'For this training, use the AI chatbot provided or approved by your institution. If you need an account, follow your institution’s instructions.');
  }
  function clipboardContent(parent) {
    const section = add(parent, e('section', 'aqg-basics-section'));
    heading(section, 'What do copying, pasting, and the clipboard mean?');
    paragraph(section, 'The clipboard is a temporary holding place for text you copy. After copying, you can paste the same text into another text box, including one on a different website.');
    const flow = add(section, e('div', 'aqg-basics-flow'));
    add(flow, e('span', '', 'Copy prompt in the AQG app'));
    add(flow, e('b', '', '→'));
    add(flow, e('span', '', 'Clipboard'));
    add(flow, e('b', '', '→'));
    add(flow, e('span', '', 'Paste in chatbot → Send'));
    paragraph(section, 'On Windows or Chromebook, highlight text and press Ctrl+C to copy, then Ctrl+V to paste. On a Mac, use Command+C and Command+V. You can also right-click to choose Copy or Paste.');
    paragraph(section, 'On a phone or tablet, press and hold text to select and copy it. Press and hold the message box to bring up Paste.');
    const hint = add(section, e('div', 'aqg-basics-hint'));
    hint.textContent = 'In the AQG app, you do NOT have to highlight a prompt: clicking Copy Starter Prompt (or another prompt button) copies its text automatically. Switch to the AI chatbot, paste the text into its message box, and press Send. The two websites do not send messages to one another automatically.';
  }
  function open() {
    close();
    const focusBefore = document.activeElement;
    const overlay = e('div', 'aqg-basics-overlay');
    const drawer = add(overlay, e('aside', 'aqg-basics-drawer'));
    drawer.setAttribute('role', 'dialog');
    drawer.setAttribute('aria-modal', 'true');
    drawer.setAttribute('aria-label', 'Quick start guide');
    const top = add(drawer, e('div', 'aqg-basics-top'));
    add(top, e('h2', '', 'Quick start guide'));
    const x = add(top, e('button', 'aqg-basics-dismiss', '×'));
    x.type='button'; x.setAttribute('aria-label', 'Close quick start guide'); x.addEventListener('click', close);
    chatbotContent(drawer);
    clipboardContent(drawer);
    const footer = add(drawer, e('div', 'aqg-basics-footer'));
    const done = add(footer, e('button', 'aqg-basics-btn', 'Close guide'));
    done.type='button'; done.addEventListener('click', close);
    overlay.addEventListener('click', event => { if(event.target === overlay) close(); });
    const keydown = event => {
      if (event.key === 'Escape') close();
      if (event.key === 'Tab') {
        const controls = [...drawer.querySelectorAll('a[href], button:not([disabled])')];
        if (!controls.length) return;
        const first=controls[0], last=controls[controls.length-1];
        if (event.shiftKey && document.activeElement === first) { event.preventDefault(); last.focus(); }
        else if (!event.shiftKey && document.activeElement === last) { event.preventDefault(); first.focus(); }
      }
    };
    document.addEventListener('keydown', keydown);
    const overflow = document.body.style.overflow;
    document.body.style.overflow = 'hidden';
    document.body.appendChild(overlay);
    activeOverlay = { overlay, focusBefore, keydown, overflow };
    x.focus();
  }
  function renderCard() {
    const card = e('div', 'aqg-basics-card');
    add(card, e('h2', '', 'Quick start guide'));
    add(card, e('p', 'basics-sub',
      "If you aren't confident about what I mean when I say AI chatbot, or what it means to copy and paste something, here is a quick start guide that will give you the basic information you need to get started."));
    const actions = add(card, e('div', 'basics-actions'));
    const guide = add(actions, e('button', 'aqg-basics-btn secondary', 'Quick start guide'));
    guide.type = 'button';
    guide.addEventListener('click', () => open());
    return card;
  }
  window.AQGBeginnerHelp = Object.freeze({ renderCard, open, close });
})();
