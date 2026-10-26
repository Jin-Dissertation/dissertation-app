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
    richLine(section, 'p', [
      'Artificial intelligence (AI) is a broad term for any technology that performs tasks that ordinarily require human intelligence. However, much of the recent excitement about AI is specifically about ',
      {term:'AI chatbots'}, ', such as ChatGPT, Gemini, Copilot, and Claude.'
    ]);
    paragraph(section, 'These chatbots are designed to mimic a human text message conversation: you send a message to the AI chatbot, and it then responds in a way that resembles a human.');

    heading(section, 'How do I use an AI chatbot?');
    paragraph(section, 'When speaking to an AI chatbot, it might help to imagine that you are texting someone who is ready to do their best to respond to whatever you send them.');
    richLine(section, 'p', [
      'For example, if you ', {term:'ask a question'}, ', the chatbot will try to answer it. If you ',
      {term:'give it instructions'}, ', it will try to follow them. If you ',
      {term:'share an observation'}, ', it will react to it.'
    ]);
    paragraph(section, 'You can continue the conversation by sending additional messages, just as you would when conversing with a person by text messages.');
  }
  // Emphasize interface commands and shortcut keys without using HTML strings.
  function richLine(parent, tag, parts) {
    const node = add(parent, e(tag));
    for (const part of parts) {
      if (typeof part === 'string') add(node, document.createTextNode(part));
      else add(node, e(part.key ? 'code' : 'strong', '', part.key || part.term));
    }
    return node;
  }
  // Clipboard exercises are optional and local to this drawer. Nothing is sent or stored.
  function practiceTextArea(parent, description, initialText = null) {
    const label = add(parent, e('label', 'aqg-basics-practice-label'));
    add(label, e('span', '', description));
    const field = add(label, e('textarea', 'aqg-basics-practice-input'));
    field.rows = 2;
    field.spellcheck = false;
    if (initialText !== null) {
      field.value = initialText;
      field.readOnly = true;
      field.setAttribute('aria-readonly', 'true');
    } else {
      field.placeholder = 'Paste the text here';
    }
    return field;
  }
  function practiceStatus(parent) {
    const status = add(parent, e('p', 'aqg-basics-practice-status'));
    status.setAttribute('role', 'status');
    status.setAttribute('aria-live', 'polite');
    return status;
  }
  function clipboardPractice(parent) {
    const practice = add(parent, e('div', 'aqg-basics-practice'));
    const manual = add(practice, e('div', 'aqg-basics-practice-step'));
    add(manual, e('h4', '', 'Try it: Copy and Paste'));
    paragraph(manual, 'Select the text in the first box, Copy it using your device, and Paste it into the second box.');
    const firstMessage = 'Practice makes progress. I copied this message myself.';
    practiceTextArea(manual, 'Text to copy (read-only)', firstMessage);
    const manualDestination = practiceTextArea(manual, 'Paste the text here');
    const manualStatus = practiceStatus(manual);
    manualDestination.addEventListener('input', () => {
      const success = manualDestination.value === firstMessage;
      manualStatus.textContent = success ? 'Success! You copied and pasted the text correctly.' : '';
      manualStatus.classList.toggle('is-success', success);
    });

    const automatic = add(practice, e('div', 'aqg-basics-practice-step'));
    add(automatic, e('h4', '', 'Try it: Copy with a button'));
    paragraph(automatic, 'Websites can make it quick and easy to copy a block of text with a single click or tap. To try this, push the button below, then paste the copied message into the box.');
    add(automatic, e('p', 'aqg-basics-step-label', 'First, push this button'));
    const secondMessage = 'This text was copied to the clipboard the moment you pushed the button.';
    const copyButton = add(automatic, e('button', 'aqg-basics-btn secondary', 'Copy sample text'));
    copyButton.type = 'button';
    const copyStatus = practiceStatus(automatic);
    const automaticDestination = practiceTextArea(automatic, 'Second, paste the text in the box below');
    const automaticStatus = practiceStatus(automatic);
    let copiedByButton = false;
    const updateAutomatic = () => {
      const success = copiedByButton && automaticDestination.value === secondMessage;
      automaticStatus.textContent = success ? 'Success! The button copied the message to your Clipboard.' : '';
      automaticStatus.classList.toggle('is-success', success);
    };
    automaticDestination.addEventListener('input', updateAutomatic);
    copyButton.addEventListener('click', async () => {
      try {
        if (!navigator.clipboard?.writeText) throw new Error('Clipboard access unavailable');
        await navigator.clipboard.writeText(secondMessage);
        copiedByButton = true;
        copyStatus.textContent = 'Copied! Now Paste the message into the box below.';
      } catch {
        copiedByButton = false;
        copyStatus.textContent = 'Automatic copying is unavailable in this browser. You can still continue the training.';
      }
      updateAutomatic();
    });
  }
  function clipboardContent(parent) {
    const section = add(parent, e('section', 'aqg-basics-section'));
    heading(section, 'What do "copying," "pasting," and "clipboard" mean?');
    richLine(section, 'p', [
      'When using a computer or phone, you can highlight a portion of text and perform the ',
      {term:'Copy'}, " function. Doing so places the selected text into a temporary holding area in the computer's memory called the ",
      {term:'clipboard'}, '.'
    ]);
    paragraph(section, 'If you then copy something else, the previous contents of the clipboard are immediately forgotten and replaced with the new text.');
    richLine(section, 'p', [
      'You can then point to another location and use the ', {term:'Paste'},
      ' function to insert whatever is currently on the clipboard into the selected location.'
    ]);
    paragraph(section, "Here's how to perform these functions on different devices:");
    const devices = add(section, e('ul', 'aqg-basics-device-list'));
    richLine(devices, 'li', [
      {term:'Windows or Chromebook'}, ': Highlight the text and press ', {key:'Ctrl + C'},
      ' to copy it. Click where you want the text to go and press ', {key:'Ctrl + V'}, ' to paste it.'
    ]);
    richLine(devices, 'li', [
      {term:'Mac'}, ': Highlight the text and press ', {key:'⌘ + C'},
      ' to copy it. Click where you want the text to go and press ', {key:'⌘ + V'}, ' to paste it.'
    ]);
    richLine(devices, 'li', [
      {term:'Phone or tablet'}, ': Long-press the text to select it, adjust the selection if necessary, and tap ',
      {term:'Copy'}, '. Then long-press where you want to insert the text and tap ', {term:'Paste'}, '.'
    ]);

    const flow = add(section, e('div', 'aqg-basics-flow'));
    richLine(flow, 'span', [{term:'Copy'}, ' a prompt in the AQG app']);
    add(flow, e('b', '', '→'));
    richLine(flow, 'span', [{term:'Clipboard'}]);
    add(flow, e('b', '', '→'));
    richLine(flow, 'span', [{term:'Paste'}, ' in chatbot → Send']);

    clipboardPractice(section);
    paragraph(section, 'You now have the basic knowledge you need to continue with this training module.');
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
  function renderCard(item = {}) {
    const card = e('div', 'aqg-basics-card');
    add(card, e('h2', '', item.header || 'Using AI Chatbots'));
    // Display this non-graded reading and its list before offering the guide.
    // Build DOM nodes rather than injecting HTML from editable curriculum text.
    const body = String(item.body || '');
    for (const chunk of body.split(/\n\s*\n/).map(part => part.trim()).filter(Boolean)) {
      const lines = chunk.split(/\n/).map(line => line.trim()).filter(Boolean);
      if (lines.every(line => line.startsWith('-- '))) {
        const list = add(card, e('ul', 'aqg-basics-intro-list'));
        for (const line of lines) add(list, e('li', '', line.slice(3)));
      } else {
        add(card, e('p', 'basics-sub', lines.join(' ')));
      }
    }
    const actions = add(card, e('div', 'basics-actions'));
    const guide = add(actions, e('button', 'aqg-basics-btn secondary', 'Quick Start Guide'));
    guide.type = 'button';
    guide.addEventListener('click', () => open());
    return card;
  }
  window.AQGBeginnerHelp = Object.freeze({ renderCard, open, close });
})();
