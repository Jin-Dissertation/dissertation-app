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
    paragraph(automatic, 'Websites such as the AQG app can copy a predetermined message with one click or tap. Select the button below, then Paste the copied message into the box.');
    const secondMessage = 'This text was copied to the clipboard the moment you pushed the button.';
    const copyButton = add(automatic, e('button', 'aqg-basics-btn secondary', 'Copy sample text'));
    copyButton.type = 'button';
    const copyStatus = practiceStatus(automatic);
    const automaticDestination = practiceTextArea(automatic, 'Paste the text copied by the button here');
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
    heading(section, 'What are Copy, Paste, and the Clipboard?');
    richLine(section, 'p', [
      'On a computer or phone, you can highlight a portion of text and choose ',
      {term:'Copy'}, '. This places that text in a temporary holding area called the ',
      {term:'Clipboard'}, '. If you ', {term:'Copy'}, ' something else, the new text replaces what was there before. ',
      'You can then go to another document, app, or website and choose ', {term:'Paste'},
      '. This inserts whatever text is currently on the ', {term:'Clipboard'},
      ' wherever your cursor is.'
    ]);
    const flow = add(section, e('div', 'aqg-basics-flow'));
    richLine(flow, 'span', [{term:'Copy'}, ' a prompt in the AQG app']);
    add(flow, e('b', '', '→'));
    richLine(flow, 'span', [{term:'Clipboard'}]);
    add(flow, e('b', '', '→'));
    richLine(flow, 'span', [{term:'Paste'}, ' in chatbot → Send']);
    paragraph(section, 'Choose the instructions for your device:');
    const devices = add(section, e('ul', 'aqg-basics-device-list'));
    richLine(devices, 'li', [
      {term:'Windows / Chromebook'}, ': Select text and press ', {key:'Ctrl+C'},
      ' to ', {term:'Copy'}, ', then ', {key:'Ctrl+V'}, ' to ', {term:'Paste'},
      '. You can also right-click and choose ', {term:'Copy'}, ' or ', {term:'Paste'}, '.'
    ]);
    richLine(devices, 'li', [
      {term:'Mac'}, ': Select text and press ', {key:'Command+C'}, ' to ', {term:'Copy'},
      ', then ', {key:'Command+V'}, ' to ', {term:'Paste'},
      '. You can also right-click and choose ', {term:'Copy'}, ' or ', {term:'Paste'}, '.'
    ]);
    richLine(devices, 'li', [
      {term:'Phone / Tablet'}, ': Press and hold to select text, then tap ', {term:'Copy'},
      '. Press and hold where you want to insert the text, then tap ', {term:'Paste'}, '.'
    ]);
    clipboardPractice(section);
    const hint = add(section, e('div', 'aqg-basics-hint'));
    richLine(hint, 'span', [
      'In the AQG app, you do not need to highlight a prompt. Selecting ',
      {term:'Copy Starter Prompt'}, ' (or another prompt button) copies its text automatically. ',
      'Switch to your AI chatbot, ', {term:'Paste'},
      ' the text into its message box, and press Send. The two websites do not send messages to one another automatically.'
    ]);
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
