class GoogleSheetsAIAssistant {
  constructor() {
    this.sessionId = this.generateSessionId();
    this.chatHistory = [];
    this.currentSheetMetadata = {};
    this.isInitialized = false;
    this.retryCount = 0;
    this.waitForSheetsLoad();
  }

  generateSessionId() {
    return 'session_' + Math.random().toString(36).substr(2, 9);
  }

  waitForSheetsLoad() {
    const checkInterval = setInterval(() => {
      const sheetsLoaded = document.querySelector('.grid3-wrapper') ||
        document.querySelector('#t-main-container') ||
        document.querySelector('.docs-texteventtarget-iframe');

      if (sheetsLoaded && !this.isInitialized) {
        clearInterval(checkInterval);
        setTimeout(() => {
          this.initializeUI();
          this.captureSheetMetadata();
          this.isInitialized = true;
        }, 1500);
      }

      if (++this.retryCount > 60) {
        clearInterval(checkInterval);
        console.warn('Google Sheets AI Assistant: Timeout waiting for sheets to load');
      }
    }, 500);
  }

  initializeUI() {
    if (document.getElementById('ai-chat-container')) return;

    const chatContainer = document.createElement('div');
    chatContainer.id = 'ai-chat-container';
    chatContainer.innerHTML = `
      <div id="ai-chat-header">
        <h3>🤖 AI Sheets Assistant</h3>
        <button id="ai-chat-toggle">−</button>
      </div>
      <div id="ai-chat-messages"></div>
      <div id="ai-chat-suggestions">
        <button class="suggestion-btn">Sum column A</button>
        <button class="suggestion-btn">Create VLOOKUP formula</button>
        <button class="suggestion-btn">Add conditional formatting</button>
        <button class="suggestion-btn">Generate pivot table</button>
      </div>
      <div id="ai-chat-input-container">
        <div id="ai-chat-input-wrapper">
          <textarea id="ai-chat-input" placeholder="Ask the AI to help with your sheet..."></textarea>
          <button id="ai-chat-send">Send</button>
        </div>
        <div id="ai-chat-status"></div>
      </div>
    `;

    document.body.appendChild(chatContainer);
    this.attachEventListeners();
    this.addWelcomeMessage();
  }

  attachEventListeners() {
    const sendBtn = document.getElementById('ai-chat-send');
    const chatInput = document.getElementById('ai-chat-input');
    const toggleBtn = document.getElementById('ai-chat-toggle');
    const suggestions = document.querySelectorAll('.suggestion-btn');

    sendBtn.addEventListener('click', () => this.sendMessage());

    chatInput.addEventListener('keypress', (e) => {
      if (e.key === 'Enter' && !e.shiftKey) {
        e.preventDefault();
        this.sendMessage();
      }
    });

    toggleBtn.addEventListener('click', () => this.toggleChat());

    suggestions.forEach(btn => {
      btn.addEventListener('click', () => {
        chatInput.value = btn.textContent;
        this.sendMessage();
      });
    });

    chatInput.addEventListener('input', () => {
      chatInput.style.height = 'auto';
      chatInput.style.height = Math.min(chatInput.scrollHeight, 120) + 'px';
    });
  }

  addWelcomeMessage() {
    this.addMessage('assistant', `👋 **Welcome to AI Sheets Assistant!**\n\nI can help you with:\n-  **Excel formulas** (VLOOKUP, SUMIF, INDEX/MATCH, etc.)\n-  **Data analysis** and calculations\n-  **Chart creation** and pivot tables  \n-  **VBA/Apps Script** code generation\n-  **Data formatting** and conditional rules\n\n**Try asking:** *"Create a formula to sum all values in column A"* or *"Generate a VLOOKUP to find product prices"*\n\nWhat would you like to do with your spreadsheet?`);
  }

  async sendMessage() {
    const input = document.getElementById('ai-chat-input');
    const message = input.value.trim();
    if (!message) return;

    this.addMessage('user', message);
    input.value = '';
    input.style.height = 'auto';

    this.setStatus('🤔 Thinking...');

    try {
      await this.captureSheetMetadata();
      const response = await this.queryAI(message);
      if (response.success) {
        this.addMessage('assistant', response.response.response);
        await this.executeAIResponse(response.response.response, message);
      } else {
        this.addMessage('assistant', 'Sorry, I encountered an error. Please try again.');
        console.error('AI Query Error:', response.error);
      }
    } catch (error) {
      this.addMessage('assistant', 'Connection error. Please check your internet and try again.');
      console.error('Send Message Error:', error);
    }

    this.setStatus('');
  }

  addMessage(role, content) {
    const messagesContainer = document.getElementById('ai-chat-messages');
    const messageDiv = document.createElement('div');
    messageDiv.className = `message ${role}-message`;

    if (role === 'assistant') {
      content = this.formatAssistantMessage(content);
    }

    messageDiv.innerHTML = `
      <div class="message-content">${content}</div>
      <div class="message-time">${new Date().toLocaleTimeString()}</div>
    `;

    messagesContainer.appendChild(messageDiv);
    messagesContainer.scrollTop = messagesContainer.scrollHeight;

    this.chatHistory.push({ role, content, timestamp: Date.now() });
  }

  formatAssistantMessage(content) {
    content = content.replace(/```([\s\S]*?)```/g, '<div class="code-block"><pre>$1</pre></div>');
    content = content.replace(/`([^`]+)`/g, '<span class="inline-code">$1</span>');
    content = content.replace(/=([\w\(\),\s\+\-\*\/\:\$"]+)/g, '<span class="formula">=$1</span>');
    content = content.replace(/\*\*(.*?)\*\*/g, '<strong>$1</strong>');
    content = content.replace(/\n/g, '<br/>');
    return content;
  }

  queryAI(prompt) {
    return new Promise((resolve) => {
      chrome.runtime.sendMessage({
        type: "PROMPT_TO_SERVER",
        payload: {
          session_id: this.sessionId,
          prompt: prompt,
          metadata: this.currentSheetMetadata,
          model: "qwen-2.5-coder-14b"
        }
      }, (response) => resolve(response));
    });
  }

  async captureSheetMetadata() {
    try {
      const sheetTitle = document.querySelector('#docs-title-input-label-inner')?.textContent || 'Unknown Sheet';
      const activeSheet = document.querySelector('.grid3-name-box input')?.value || 'Sheet1';

      const headers = [];
      const headerCells = document.querySelectorAll('.grid3-hh-cell');
      headerCells.forEach((cell) => {
        const text = cell.textContent?.trim();
        if (text) headers.push(text);
      });

      this.currentSheetMetadata = {
        sheetTitle,
        activeSheet,
        headers: headers.slice(0, 26),
        timestamp: Date.now(),
      };
    } catch (e) {
      console.warn('Failed to capture metadata', e);
    }
  }

  async executeAIResponse(aiResponse, userPrompt) {
    try {
      this.addActionButtons(aiResponse, userPrompt);
    } catch (e) {
      console.warn('Failed to execute AI response', e);
    }
  }

  addActionButtons(aiResponse, userPrompt) {
    const actionsDiv = document.createElement('div');
    actionsDiv.className = 'action-buttons';

    const buttons = [];

    if (aiResponse.includes('=')) {
      buttons.push({ text: '📝 Insert Formula', action: () => this.showFormulaInstructions(aiResponse) });
    }
    if (aiResponse.match(/[A-Z]+\d+/)) {
      buttons.push({ text: '🎯 Highlight Cells', action: () => this.showCellSelectionTip() });
    }
    if (userPrompt.toLowerCase().includes('script') || userPrompt.toLowerCase().includes('macro')) {
      buttons.push({ text: '⚡ View Script', action: () => this.showScriptCode(aiResponse) });
    }

    buttons.forEach(btn => {
      const buttonEl = document.createElement('button');
      buttonEl.textContent = btn.text;
      buttonEl.className = 'action-btn';
      buttonEl.addEventListener('click', btn.action);
      actionsDiv.appendChild(buttonEl);
    });

    if (buttons.length > 0) {
      const messagesContainer = document.getElementById('ai-chat-messages');
      messagesContainer.appendChild(actionsDiv);
      messagesContainer.scrollTop = messagesContainer.scrollHeight;
    }
  }

  showFormulaInstructions(response) {
    const formulas = response.match(/=[\w\(\),\s\+\-\*\/\:\$"]+/g) || [];
    let instructions = "📝 **To insert these formulas:**\n\n";
    formulas.forEach((formula, index) => {
      instructions += `${index + 1}. Click on the target cell\n`;
      instructions += `2. Type: \`${formula}\`\n`;
      instructions += `3. Press Enter\n\n`;
    });
    this.addMessage('system', instructions);
  }

  showCellSelectionTip() {
    this.addMessage('system', '🎯 **Cell Selection Tip:**\n\n1. Click and drag to select a range\n2. Hold Ctrl to select multiple cells\n3. Use Shift+Click to select a range\n\nThe AI suggestions above reference specific cells you should work with.');
  }

  showScriptCode(response) {
    const codeBlocks = response.match(/```([\s\S]*?)```/g) || [];
    if (codeBlocks.length > 0) {
      this.addMessage('system', `⚡ **Script Code Found:**\n\n${codeBlocks.join('\n\n')}\n\n**To use:** Go to Extensions → Apps Script in Google Sheets and paste this code.`);
    }
  }

  toggleChat() {
    const container = document.getElementById('ai-chat-container');
    const toggle = document.getElementById('ai-chat-toggle');
    if (container.classList.contains('minimized')) {
      container.classList.remove('minimized');
      toggle.textContent = '−';
    } else {
      container.classList.add('minimized');
      toggle.textContent = '+';
    }
  }

  setStatus(message) {
    const el = document.getElementById('ai-chat-status');
    if (el) el.textContent = message;
  }
}

if (document.readyState === 'loading') {
  document.addEventListener('DOMContentLoaded', () => {
    setTimeout(() => new GoogleSheetsAIAssistant(), 2000);
  });
} else {
  setTimeout(() => new GoogleSheetsAIAssistant(), 2000);
}


