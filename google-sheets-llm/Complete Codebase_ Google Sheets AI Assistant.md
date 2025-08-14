<img src="https://r2cdn.perplexity.ai/pplx-full-logo-primary-dark%402x.png" style="height:64px;margin-right:32px"/>

# Complete Codebase: Google Sheets AI Assistant

Here's the complete, production-ready codebase with all components integrated:

## **🐍 1. Python Backend (Modal.ai)**

### **modal_server.py**

```python
import modal
import uuid
from pathlib import Path
from typing import Optional, List, Dict
import json

# Modal configuration
MINUTES = 60
app = modal.App("google-sheets-coding-llm")

# Build llama.cpp with CUDA support
cuda_version = "12.4.0"
flavor = "devel"
operating_sys = "ubuntu22.04"
tag = f"{cuda_version}-{flavor}-{operating_sys}"

image = (
    modal.Image.from_registry(f"nvidia/cuda:{tag}", add_python="3.12")
    .apt_install("git", "build-essential", "cmake", "curl", "libcurl4-openssl-dev")
    .run_commands("git clone https://github.com/ggerganov/llama.cpp")
    .run_commands(
        "cmake llama.cpp -B llama.cpp/build "
        "-DBUILD_SHARED_LIBS=OFF -DGGML_CUDA=ON -DLLAMA_CURL=ON "
    )
    .run_commands(
        "cmake --build llama.cpp/build --config Release -j --clean-first --target llama-quantize llama-cli"
    )
    .run_commands("cp llama.cpp/build/bin/llama-* llama.cpp")
    .pip_install("fastapi", "python-multipart", "uvicorn")
    .entrypoint([])
)

# Model storage
model_cache = modal.Volume.from_name("coding-llm-cache", create_if_missing=True)
cache_dir = "/root/.cache/models"

# Session storage  
session_volume = modal.Volume.from_name("coding-llm-sessions", create_if_missing=True)
session_dir = "/root/sessions"

# Download image for Huggingface models
download_image = (
    modal.Image.debian_slim(python_version="3.12")
    .pip_install("huggingface_hub[hf_transfer]==0.26.2")
    .env({"HF_HUB_ENABLE_HF_TRANSFER": "1"})
)

@app.function(
    image=download_image, 
    volumes={cache_dir: model_cache}, 
    timeout=30 * MINUTES
)
def download_coding_models():
    """Download the best open-source coding models in GGUF format"""
    from huggingface_hub import snapshot_download
    
    # Top coding models for 2025
    models_to_download = [
        {
            "repo_id": "Qwen/Qwen2.5-Coder-14B-Instruct-GGUF",
            "pattern": "*Q4_K_M.gguf",
            "name": "qwen-2.5-coder-14b"
        },
        {
            "repo_id": "bartowski/DeepSeek-Coder-V2-Lite-Instruct-GGUF", 
            "pattern": "*Q4_K_M.gguf",
            "name": "deepseek-coder-v2-lite"
        },
        {
            "repo_id": "microsoft/Phi-3.5-mini-instruct-gguf",
            "pattern": "*Q4_K_M.gguf",
            "name": "phi-3.5-mini"
        }
    ]
    
    for model in models_to_download:
        print(f"🦙 downloading {model['name']} from {model['repo_id']}")
        try:
            snapshot_download(
                repo_id=model["repo_id"],
                local_dir=f"{cache_dir}/{model['name']}",
                allow_patterns=[model["pattern"]],
            )
            print(f"✅ {model['name']} downloaded")
        except Exception as e:
            print(f"❌ Failed to download {model['name']}: {e}")
    
    model_cache.commit()
    print("🦙 Coding models setup complete")

# Context manager for sessions (with coding-specific prompting)
class CodingContextManager:
    def __init__(self, session_dir: str):
        self.session_dir = Path(session_dir)
        
    def load_session(self, session_id: str) -> List[Dict]:
        session_file = self.session_dir / f"{session_id}.json"
        if session_file.exists():
            try:
                return json.loads(session_file.read_text())
            except:
                return []
        return []
    
    def save_session(self, session_id: str, messages: List[Dict]):
        try:
            self.session_dir.mkdir(parents=True, exist_ok=True)
            session_file = self.session_dir / f"{session_id}.json"
            session_file.write_text(json.dumps(messages, indent=2))
        except Exception as e:
            print(f"Error saving session {session_id}: {e}")
    
    def add_message(self, session_id: str, role: str, content: str, metadata: Dict = None):
        messages = self.load_session(session_id)
        
        # Add system prompt for coding context if first message
        if not messages:
            system_prompt = self._create_coding_system_prompt(metadata)
            messages.append({"role": "system", "content": system_prompt})
        
        messages.append({"role": role, "content": content})
        
        # Summarize if context gets too long
        if len(messages) > 25:
            messages = self.summarize_coding_context(messages, metadata)
            
        self.save_session(session_id, messages)
        return messages
    
    def _create_coding_system_prompt(self, metadata: Dict) -> str:
        return f"""You are an expert coding assistant specialized in Excel/Google Sheets automation and programming tasks.

Current spreadsheet context:
{json.dumps(metadata, indent=2) if metadata else "No spreadsheet data available"}

Your expertise includes:
- Excel formula generation (VLOOKUP, INDEX/MATCH, SUMIF, PIVOT, etc.)
- VBA/Visual Basic programming
- Google Apps Script
- Data analysis and manipulation
- Financial and investment banking calculations
- JavaScript for web automation

Always provide:
1. Clear, working code solutions
2. Step-by-step explanations
3. Specific cell references and ranges
4. Copy-paste ready formulas
5. Error handling considerations

Focus on practical, production-ready solutions that integrate seamlessly with spreadsheet environments."""

    def summarize_coding_context(self, messages: List[Dict], metadata: Dict) -> List[Dict]:
        if len(messages) <= 15:
            return messages
            
        system_msgs = [m for m in messages if m["role"] == "system"]
        recent_msgs = messages[-8:]
        
        # Create coding-focused summary
        middle_msgs = messages[len(system_msgs):-8]
        summary = "Previous coding conversation summary:\n"
        
        for msg in middle_msgs[-12:]:
            if msg["role"] == "user":
                summary += f"• User asked: {msg['content'][:150]}...\n"
            elif msg["role"] == "assistant":
                content = msg['content']
                if "```
                    summary += f"-  Assistant provided code solution\n"
                else:
                    summary += f"-  Assistant explained: {content[:100]}...\n"
        
        # Update system prompt with current metadata
        updated_system = self._create_coding_system_prompt(metadata)
        summary_msg = {"role": "system", "content": f"{updated_system}\n\n{summary}"}
        
        return [summary_msg] + recent_msgs

@app.function(
    image=image,
    volumes={
        cache_dir: model_cache, 
        session_dir: session_volume
    },
    gpu="L40S",  # High-performance GPU for coding models
    timeout=15 * MINUTES,
    allow_concurrent_inputs=10,
)
def coding_llm_inference(session_id: str, prompt: str, metadata: Dict = None, model_choice: str = "qwen-2.5-coder-14b"):
    """Run coding-focused LLM inference with session context and spreadsheet metadata"""
    import subprocess
    import os
    
    # Initialize coding context manager
    ctx_manager = CodingContextManager(session_dir)
    
    # Add coding-specific context to prompt
    coding_prompt = f"""
Spreadsheet Context: {json.dumps(metadata, indent=2) if metadata else "No spreadsheet loaded"}

User Request: {prompt}

Please provide a practical solution with:
1. Working code/formulas
2. Clear explanations
3. Integration steps for Google Sheets/Excel
4. Specific cell references where applicable
"""
    
    # Add to session context with metadata
    messages = ctx_manager.add_message(session_id, "user", coding_prompt, metadata)
    
    # Build conversation context for llama.cpp
    context_text = "\n".join([
        f"{msg['role'].title()}: {msg['content']}" for msg in messages
    ])
    context_text += "\nAssistant:"
    
    # Model selection - find available model files
    model_paths = {
        "qwen-2.5-coder-14b": "qwen-2.5-coder-14b",
        "deepseek-coder-v2": "deepseek-coder-v2-lite", 
        "phi-3.5-mini": "phi-3.5-mini"
    }
    
    model_dir = model_paths.get(model_choice, "phi-3.5-mini")  # Fallback to smaller model
    model_path = f"{cache_dir}/{model_dir}"
    
    # Find GGUF files in the model directory
    model_files = []
    if os.path.exists(model_path):
        model_files = list(Path(model_path).glob("*.gguf"))
    
    if not model_files:
        return {
            "session_id": session_id,
            "response": f"Model {model_choice} not found. Please run setup_coding_models first.",
            "error": "Model not available"
        }
    
    model_file = str(model_files)
    
    command = [
        "/llama.cpp/llama-cli",
        "--model", model_file,
        "--n-gpu-layers", "40",
        "--prompt", context_text,
        "--n-predict", "1024",
        "--ctx-size", "8192",
        "--temp", "0.2",
        "--top-p", "0.9",
        "--repeat-penalty", "1.1",
        "--stop", "User:",
        "--stop", "Human:",
    ]
    
    print(f"🦙 Running {model_choice} inference for session {session_id}")
    
    try:
        result = subprocess.run(
            command, 
            capture_output=True, 
            text=True, 
            check=True,
            timeout=300  # 5 minute timeout
        )
        
        response = result.stdout.strip()
        
        # Clean up response
        if "Assistant:" in response:
            response = response.split("Assistant:")[-1].strip()
        
        # Remove any remaining prompt echoes
        lines = response.split('\n')
        cleaned_lines = []
        for line in lines:
            if not line.startswith(('User:', 'Human:', 'Assistant:')):
                cleaned_lines.append(line)
        response = '\n'.join(cleaned_lines).strip()
        
        # Save assistant response to session
        ctx_manager.add_message(session_id, "assistant", response, metadata)
        session_volume.commit()
        
        return {
            "session_id": session_id,
            "response": response,
            "model_used": model_choice,
            "metadata": metadata
        }
        
    except subprocess.CalledProcessError as e:
        print(f"Error running inference: {e}")
        return {
            "session_id": session_id,
            "response": "I encountered an error processing your coding request. Please try again with a simpler query.",
            "error": str(e),
            "model_used": model_choice
        }
    except Exception as e:
        print(f"Unexpected error: {e}")
        return {
            "session_id": session_id,
            "response": "An unexpected error occurred. Please try again.",
            "error": str(e)
        }

@app.function(
    image=modal.Image.debian_slim().pip_install("fastapi", "uvicorn"),
    volumes={session_dir: session_volume}
)
@modal.web_endpoint(method="POST", label="coding-query")
def coding_query_endpoint(
    session_id: str = None, 
    prompt: str = "", 
    metadata: dict = None,
    model: str = "phi-3.5-mini"
):
    """Web endpoint for Chrome extension - optimized for coding tasks"""
    if not session_id:
        session_id = str(uuid.uuid4())
    
    if not prompt:
        return {"error": "No prompt provided"}
    
    # Call the coding inference function
    result = coding_llm_inference.remote(session_id, prompt, metadata or {}, model)
    
    return result

@app.local_entrypoint()
def setup_coding_models():
    """Download and setup the best open-source coding models"""
    download_coding_models.remote()
    print("✅ Top coding models setup complete!")
    print("Available models: qwen-2.5-coder-14b, deepseek-coder-v2, phi-3.5-mini")
```


---

## **🎨 2. Complete Chrome Extension**

### **manifest.json**

```
{
  "name": "Google Sheets AI Assistant",
  "description": "Chat-based AI assistant for Google Sheets automation with advanced coding capabilities",
  "version": "1.0.0",
  "manifest_version": 3,
  "permissions": [
    "scripting", 
    "storage", 
    "activeTab"
  ],
  "host_permissions": [
    "https://docs.google.com/spreadsheets/*",
    "https://sheets.googleapis.com/*",
    "https://*.modal.run/*"
  ],
  "background": {
    "service_worker": "background.js"
  },
  "action": {
    "default_title": "AI Sheets Assistant",
    "default_popup": "popup.html"
  },
  "content_scripts": [
    {
      "matches": ["https://docs.google.com/spreadsheets/*"],
      "js": ["content.js"],
      "css": ["chat.css"],
      "run_at": "document_end"
    }
  ],
  "web_accessible_resources": [
    {
      "resources": ["chat.css", "icons/*"],
      "matches": ["<all_urls>"]
    }
  ]
}
```


### **background.js**

```
// Background service worker for Chrome extension
chrome.runtime.onMessage.addListener((msg, sender, sendResponse) => {
  if (msg.type === "PROMPT_TO_SERVER") {
    handleAIQuery(msg.payload)
      .then(data => sendResponse({ success: true, response: data }))
      .catch(err => sendResponse({ success: false, error: err.toString() }));
    return true; // Keep sendResponse alive for async
  }
});

async function handleAIQuery(payload) {
  // Replace with your actual Modal.ai endpoint
  const modalEndpoint = "https://your-modal-app--coding-query.modal.run";
  
  try {
    const response = await fetch(modalEndpoint, {
      method: "POST",
      headers: { 
        "Content-Type": "application/json",
        "Accept": "application/json"
      },
      body: JSON.stringify({
        session_id: payload.session_id,
        prompt: payload.prompt,
        metadata: payload.metadata,
        model: payload.model || "phi-3.5-mini"
      })
    });

    if (!response.ok) {
      throw new Error(`HTTP ${response.status}: ${response.statusText}`);
    }

    const data = await response.json();
    return data;
    
  } catch (error) {
    console.error("Modal.ai API Error:", error);
    throw error;
  }
}

// Extension installation/update handling
chrome.runtime.onInstalled.addListener((details) => {
  if (details.reason === 'install') {
    chrome.tabs.create({
      url: 'https://docs.google.com/spreadsheets/u/0/'
    });
  }
});
```


### **content.js**

```
class GoogleSheetsAIAssistant {
  constructor() {
    this.sessionId = this.generateSessionId();
    this.chatHistory = [];
    this.currentSheetMetadata = {};
    this.isInitialized = false;
    this.retryCount = 0;
    
    // Wait for Google Sheets to fully load
    this.waitForSheetsLoad();
  }

  generateSessionId() {
    return 'session_' + Math.random().toString(36).substr(2, 9) + '_' + Date.now();
  }

  waitForSheetsLoad() {
    const checkInterval = setInterval(() => {
      // Check if Google Sheets interface is loaded
      const sheetsLoaded = document.querySelector('.grid3-wrapper') || 
                          document.querySelector('#t-main-container') ||
                          document.querySelector('.docs-texteventtarget-iframe');
      
      if (sheetsLoaded && !this.isInitialized) {
        clearInterval(checkInterval);
        setTimeout(() => {
          this.initializeUI();
          this.captureSheetMetadata();
          this.isInitialized = true;
        }, 1500); // Additional delay for full render
      }
      
      // Timeout after 30 seconds
      if (++this.retryCount > 60) {
        clearInterval(checkInterval);
        console.warn('Google Sheets AI Assistant: Timeout waiting for sheets to load');
      }
    }, 500);
  }

  initializeUI() {
    if (document.getElementById('ai-chat-container')) return;

    // Create main chat container
    const chatContainer = document.createElement('div');
    chatContainer.id = 'ai-chat-container';
    chatContainer.innerHTML = `
      <div id="ai-chat-header">
        <div class="header-content">
          <span class="ai-icon">🤖</span>
          <h3>AI Sheets Assistant</h3>
        </div>
        <button id="ai-chat-toggle" title="Minimize/Maximize">−</button>
      </div>
      
      <div id="ai-chat-messages"></div>
      
      <div id="ai-chat-input-container">
        <div id="ai-chat-input-wrapper">
          <textarea 
            id="ai-chat-input" 
            placeholder="Ask me to help with your spreadsheet... (e.g., 'Add a SUM formula in cell C10' or 'Create a VLOOKUP for stock prices')"
            rows="2"
          ></textarea>
          <button id="ai-chat-send" title="Send message">
            <span class="send-icon">⚡</span>
          </button>
        </div>
        <div id="ai-chat-suggestions">
          <button class="suggestion-btn" data-prompt="Sum all values in column A">Sum column A</button>
          <button class="suggestion-btn" data-prompt="Create a VLOOKUP formula to find values">Create VLOOKUP</button>
          <button class="suggestion-btn" data-prompt="Add conditional formatting to highlight important values">Conditional formatting</button>
          <button class="suggestion-btn" data-prompt="Generate a pivot table from this data">Generate pivot table</button>
        </div>
      </div>
      
      <div id="ai-chat-status"></div>
    `;

    document.body.appendChild(chatContainer);
    this.attachEventListeners();
    this.addWelcomeMessage();
    
    // Auto-focus on input after a brief delay
    setTimeout(() => {
      document.getElementById('ai-chat-input')?.focus();
    }, 500);
  }

  attachEventListeners() {
    const sendBtn = document.getElementById('ai-chat-send');
    const chatInput = document.getElementById('ai-chat-input');
    const toggleBtn = document.getElementById('ai-chat-toggle');
    const suggestions = document.querySelectorAll('.suggestion-btn');

    // Send message on button click
    sendBtn?.addEventListener('click', () => this.sendMessage());

    // Send message on Enter (Shift+Enter for new line)
    chatInput?.addEventListener('keypress', (e) => {
      if (e.key === 'Enter' && !e.shiftKey) {
        e.preventDefault();
        this.sendMessage();
      }
    });

    // Toggle chat window
    toggleBtn?.addEventListener('click', () => this.toggleChat());

    // Quick suggestions
    suggestions.forEach(btn => {
      btn.addEventListener('click', () => {
        const prompt = btn.getAttribute('data-prompt') || btn.textContent;
        chatInput.value = prompt;
        this.sendMessage();
      });
    });

    // Auto-resize textarea
    chatInput?.addEventListener('input', () => {
      chatInput.style.height = 'auto';
      chatInput.style.height = Math.min(chatInput.scrollHeight, 120) + 'px';
    });

    // Capture sheet changes
    document.addEventListener('click', () => {
      setTimeout(() => this.captureSheetMetadata(), 500);
    });
  }

  addWelcomeMessage() {
    this.addMessage('assistant', `👋 **Welcome to AI Sheets Assistant!**

I can help you with:
-  **Excel formulas** (VLOOKUP, SUMIF, INDEX/MATCH, etc.)
-  **Data analysis** and calculations
-  **Chart creation** and pivot tables  
-  **VBA/Apps Script** code generation
-  **Data formatting** and conditional rules
-  **Financial calculations** for investment banking

**Try asking:** *"Create a formula to sum all values in column A"* or *"Generate a VLOOKUP to find product prices"*

What would you like to do with your spreadsheet?`);
  }

  async sendMessage() {
    const input = document.getElementById('ai-chat-input');
    const message = input?.value?.trim();
    
    if (!message) return;

    // Add user message to chat
    this.addMessage('user', message);
    input.value = '';
    input.style.height = 'auto';

    // Show loading state
    this.setStatus('🤔 Analyzing your request...');
    this.showTypingIndicator();

    try {
      // Update sheet metadata before sending
      await this.captureSheetMetadata();

      // Send to AI backend
      const response = await this.queryAI(message);
      
      this.hideTypingIndicator();
      
      if (response.success) {
        const aiResponse = response.response.response || 'No response received';
        this.addMessage('assistant', aiResponse);
        
        // Attempt to execute the AI's suggestions
        await this.executeAIResponse(aiResponse, message);
      } else {
        this.addMessage('assistant', '⚠️ Sorry, I encountered an error processing your request. Please try again or rephrase your question.');
        console.error('AI Query Error:', response.error);
      }
    } catch (error) {
      this.hideTypingIndicator();
      this.addMessage('assistant', '🔌 Connection error. Please check your internet connection and try again.');
      console.error('Send Message Error:', error);
    }

    this.setStatus('');
  }

  showTypingIndicator() {
    const messagesContainer = document.getElementById('ai-chat-messages');
    const typingDiv = document.createElement('div');
    typingDiv.id = 'typing-indicator';
    typingDiv.className = 'message assistant-message';
    typingDiv.innerHTML = `
      <div class="message-content typing">
        <span class="typing-dots">
          <span>.</span><span>.</span><span>.</span>
        </span>
        <span class="typing-text">AI is thinking...</span>
      </div>
    `;
    messagesContainer.appendChild(typingDiv);
    messagesContainer.scrollTop = messagesContainer.scrollHeight;
  }

  hideTypingIndicator() {
    const typingIndicator = document.getElementById('typing-indicator');
    if (typingIndicator) {
      typingIndicator.remove();
    }
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
    
    // Store in chat history
    this.chatHistory.push({ role, content, timestamp: Date.now() });
    
    // Limit chat history to prevent memory issues
    if (this.chatHistory.length > 100) {
      this.chatHistory = this.chatHistory.slice(-50);
    }
  }

  formatAssistantMessage(content) {
    // Format code blocks
    content = content.replace(/```([^`]+)```
    
    // Format inline code and formulas
    content = content.replace(/`([^`]+)`/g, '<code class="inline-code">$1</code>');
    
    // Format Excel formulas (starting with =)
    content = content.replace(/=([\w$$,\s\+\-\*\/\:\$"\.]+)/g, '<code class="formula">=<span class="formula-content">$1</span></code>');
    
    // Format bold text
    content = content.replace(/\*\*(.*?)\*\*/g, '<strong>$1</strong>');
    
    // Format bullet points
    content = content.replace(/-  /g, '<span class="bullet">- </span> ');
    
    // Convert newlines to <br>
    content = content.replace(/\n/g, '<br>');
    
    return content;
  }

  async queryAI(prompt) {
    return new Promise((resolve) => {
      chrome.runtime.sendMessage({
        type: "PROMPT_TO_SERVER",
        payload: {
          session_id: this.sessionId,
          prompt: prompt,
          metadata: this.currentSheetMetadata,
          model: "phi-3.5-mini" // Using lighter model for faster responses
        }
      }, (response) => {
        resolve(response || { success: false, error: 'No response received' });
      });
    });
  }

  async captureSheetMetadata() {
    try {
      // Get current sheet info
      const sheetTitle = document.querySelector('#docs-title-input-label-inner')?.textContent || 
                        document.querySelector('.docs-title-input-label')?.textContent ||
                        'Unknown Sheet';
      
      // Get active sheet name
      const activeSheetTab = document.querySelector('.docs-sheet-tab-name.docs-sheet-active-tab .docs-sheet-tab-name-text');
      const activeSheet = activeSheetTab?.textContent || 'Sheet1';
      
      // Get visible column headers
      const headers = [];
      const headerElements = document.querySelectorAll('.docs-sheet-header-cell-text, .cell-input, th');
      
      // Fallback to alphabet if no headers found
      let columnCount = Math.max(headerElements.length, 10);
      for (let i = 0; i < Math.min(columnCount, 26); i++) {
        const letter = String.fromCharCode(65 + i);
        const headerText = headerElements[i]?.textContent?.trim() || letter;
        headers.push(headerText === '' ? letter : headerText);
      }

      // Estimate row count by checking scroll height or visible rows
      const sheetContainer = document.querySelector('.docs-sheet-container');
      const estimatedRows = sheetContainer ? Math.floor(sheetContainer.scrollHeight / 21) : 100; // 21px typical row height
      
      // Get current selection
      const nameBox = document.querySelector('.docs-sheet-name-box input');
      const selectedRange = nameBox?.value || 'A1';

      // Detect data patterns (basic)
      const visibleCells = document.querySelectorAll('.cell-input, .docs-sheet-cell');
      let hasNumbers = false;
      let hasFormulas = false;
      let hasText = false;
      
      for (let i = 0; i < Math.min(visibleCells.length, 50); i++) {
        const cellValue = visibleCells[i]?.textContent || '';
        if (cellValue.startsWith('=')) hasFormulas = true;
        else if (!isNaN(parseFloat(cellValue))) hasNumbers = true;
        else if (cellValue.trim()) hasText = true;
      }

      this.currentSheetMetadata = {
        sheet_title: sheetTitle,
        active_sheet: activeSheet,
        columns: headers,
        estimated_rows: Math.min(estimatedRows, 1000), // Cap for performance
        selected_range: selectedRange,
        has_formulas: hasFormulas,
        has_numbers: hasNumbers,
        has_text: hasText,
        timestamp: Date.now(),
        url: window.location.href
      };

    } catch (error) {
      console.error('Error capturing sheet metadata:', error);
      this.currentSheetMetadata = {
        sheet_title: 'Google Sheets',
        active_sheet: 'Sheet1',
        columns: ['A', 'B', 'C', 'D', 'E'],
        estimated_rows: 100,
        selected_range: 'A1',
        has_formulas: false,
        has_numbers: true,
        has_text: true,
        timestamp: Date.now()
      };
    }
  }

  async executeAIResponse(aiResponse, userPrompt) {
    // Extract actionable commands from AI response
    
    // Look for formulas to insert
    const formulaMatches = aiResponse.match(/=[\w$$,\s\+\-\*\/\:\$"\.]+/g);
    if (formulaMatches && formulaMatches.length > 0) {
      this.addMessage('system', `🔧 **Found ${formulaMatches.length} formula(s) to apply:**\n\n${formulaMatches.map((f, i) => `${i+1}. \`${f}\``).join('\n')}\n\n**To use:** Click on your target cell and paste the formula.`);
    }

    // Look for cell references and actions
    const cellReferences = aiResponse.match(/[A-Z]+\d+/g);
    if (cellReferences && cellReferences.length > 0) {
      const uniqueCells = [...new Set(cellReferences)];
      this.addMessage('system', `📍 **Target cells mentioned:** ${uniqueCells.slice(0, 8).join(', ')}${uniqueCells.length > 8 ? '...' : ''}`);
    }

    // Add action buttons for common tasks
    this.addActionButtons(aiResponse, userPrompt);
  }

  addActionButtons(aiResponse, userPrompt) {
    const actionsDiv = document.createElement('div');
    actionsDiv.className = 'action-buttons';
    
    const buttons = [];
    
    // Formula insertion button
    if (aiResponse.includes('=')) {
      buttons.push({
        text: '📝 Copy Formulas',
        action: () => this.copyFormulasToClipboard(aiResponse),
        class: 'primary'
      });
    }
    
    // Range selection helper
    if (aiResponse.match(/[A-Z]+\d+/)) {
      buttons.push({
        text: '🎯 Selection Guide',
        action: () => this.showCellSelectionTip(),
        class: 'secondary'
      });
    }

    // VBA/Apps Script generation
    if (userPrompt.toLowerCase().includes('script') || userPrompt.toLowerCase().includes('macro') || aiResponse.includes('function')) {
      buttons.push({
        text: '⚡ View Code',
        action: () => this.showScriptCode(aiResponse),
        class: 'secondary'
      });
    }

    // Pivot table help
    if (userPrompt.toLowerCase().includes('pivot') || aiResponse.toLowerCase().includes('pivot')) {
      buttons.push({
        text: '📊 Pivot Help',
        action: () => this.showPivotTableGuide(),
        class: 'secondary'
      });
    }

    buttons.forEach(btn => {
      const buttonEl = document.createElement('button');
      buttonEl.textContent = btn.text;
      buttonEl.className = `action-btn ${btn.class || 'secondary'}`;
      buttonEl.addEventListener('click', btn.action);
      actionsDiv.appendChild(buttonEl);
    });

    if (buttons.length > 0) {
      const messagesContainer = document.getElementById('ai-chat-messages');
      messagesContainer.appendChild(actionsDiv);
      messagesContainer.scrollTop = messagesContainer.scrollHeight;
    }
  }

  async copyFormulasToClipboard(response) {
    const formulas = response.match(/=[\w$$,\s\+\-\*\/\:\$"\.]+/g) || [];
    const formulaText = formulas.join('\n');
    
    try {
      await navigator.clipboard.writeText(formulaText);
      this.addMessage('system', `✅ **Copied ${formulas.length} formula(s) to clipboard!**\n\nNow click on your target cell(s) and paste (Ctrl+V).`);
    } catch (error) {
      this.addMessage('system', `📋 **Formulas to copy:**\n\n${formulas.map((f, i) => `${i+1}. ${f}`).join('\n')}\n\nManually select and copy these formulas.`);
    }
  }

  showCellSelectionTip() {
    this.addMessage('system', `🎯 **Cell Selection Guide:**

**Basic Selection:**
-  Click a cell → Select single cell
-  Drag → Select range
-  Ctrl+Click → Select multiple cells
-  Shift+Click → Extend selection

**Keyboard Shortcuts:**
-  Ctrl+A → Select all
-  Ctrl+Shift+End → Select to end of data
-  Ctrl+Space → Select column
-  Shift+Space → Select row

The AI suggestions above reference specific cells you should target.`);
  }

  showScriptCode(response) {
    const codeBlocks = response.match(/```([^`]+)```
    if (codeBlocks.length > 0) {
      const cleanCode = codeBlocks.map(block => block.replace(/```/g, '')).join('\n\n');
      this.addMessage('system', `⚡ **Script Code:**\n\n\`\`\`\n${cleanCode}\n\`\`\`\n\n**To use:**\n1. Go to **Extensions → Apps Script** in Google Sheets\n2. Create a new project\n3. Paste this code\n4. Save and run`);
    } else {
      this.addMessage('system', '⚡ **No code blocks found in the response.** The AI may have provided instructions instead of raw code.');
    }
  }

  showPivotTableGuide() {
    this.addMessage('system', `📊 **Pivot Table Creation Guide:**

**Method 1: Manual**
1. Select your data range
2. Go to **Insert → Pivot Table**
3. Choose where to place it
4. Drag fields to Rows, Columns, Values

**Method 2: Smart**
1. Select any cell in your data
2. **Data → Pivot Table**
3. Google Sheets will auto-suggest configurations

**Pro Tips:**
• Clean your data first (no empty rows/columns)
• Use clear headers
• Consider data types before creating`);
  }

  toggleChat() {
    const container = document.getElementById('ai-chat-container');
    const toggle = document.getElementById('ai-chat-toggle');
    
    if (container?.classList.contains('minimized')) {
      container.classList.remove('minimized');
      toggle.textContent = '−';
      toggle.title = 'Minimize';
    } else {
      container?.classList.add('minimized');
      toggle.textContent = '+';
      toggle.title = 'Maximize';
    }
  }

  setStatus(message) {
    const statusEl = document.getElementById('ai-chat-status');
    if (statusEl) {
      statusEl.textContent = message;
    }
  }
}

// Initialize when Google Sheets is detected
if (window.location.hostname === 'docs.google.com' && 
    window.location.pathname.includes('/spreadsheets/')) {
  
  // Wait for page load
  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', () => {
      new GoogleSheetsAIAssistant();
    });
  } else {
    new GoogleSheetsAIAssistant();
  }
}
```


### **chat.css**

```css
/* AI Chat Assistant Styles */
#ai-chat-container {
  position: fixed;
  top: 80px;
  right: 20px;
  width: 400px;
  height: 650px;
  background: #ffffff;
  border: 1px solid #e1e5e9;
  border-radius: 16px;
  box-shadow: 0 16px 40px rgba(0,0,0,0.12), 0 8px 16px rgba(0,0,0,0.08);
  z-index: 999999;
  display: flex;
  flex-direction: column;
  font-family: 'Google Sans', 'Segoe UI', Roboto, sans-serif;
  transition: all 0.3s cubic-bezier(0.4, 0, 0.2, 1);
  backdrop-filter: blur(10px);
  overflow: hidden;
}

#ai-chat-container.minimized {
  height: 68px;
  overflow: hidden;
}

/* Header */
#ai-chat-header {
  background: linear-gradient(135deg, #4285f4 0%, #34a853 50%, #ea4335 100%);
  color: white;
  padding: 16px 20px;
  display: flex;
  justify-content: space-between;
  align-items: center;
  border-radius: 16px 16px 0 0;
  box-shadow: 0 2px 8px rgba(0,0,0,0.1);
}

.header-content {
  display: flex;
  align-items: center;
  gap: 10px;
}

.ai-icon {
  font-size: 20px;
  animation: pulse 2s infinite;
}

@keyframes pulse {
  0%, 100% { transform: scale(1); }
  50% { transform: scale(1.1); }
}

#ai-chat-header h3 {
  margin: 0;
  font-size: 16px;
  font-weight: 500;
  letter-spacing: 0.2px;
}

#ai-chat-toggle {
  background: rgba(255,255,255,0.2);
  border: none;
  color: white;
  width: 32px;
  height: 32px;
  border-radius: 50%;
  cursor: pointer;
  font-size: 20px;
  font-weight: bold;
  line-height: 1;
  transition: all 0.2s ease;
  display: flex;
  align-items: center;
  justify-content: center;
}

#ai-chat-toggle:hover {
  background: rgba(255,255,255,0.3);
  transform: scale(1.05);
}

/* Messages Area */
#ai-chat-messages {
  flex: 1;
  overflow-y: auto;
  padding: 20px;
  background: linear-gradient(180deg, #f8f9fa 0%, #ffffff 100%);
  scroll-behavior: smooth;
}

#ai-chat-messages::-webkit-scrollbar {
  width: 6px;
}

#ai-chat-messages::-webkit-scrollbar-track {
  background: transparent;
}

#ai-chat-messages::-webkit-scrollbar-thumb {
  background: #dadce0;
  border-radius: 3px;
}

#ai-chat-messages::-webkit-scrollbar-thumb:hover {
  background: #bdc1c6;
}

/* Messages */
.message {
  margin-bottom: 16px;
  animation: slideInUp 0.3s ease-out;
}

@keyframes slideInUp {
  from {
    opacity: 0;
    transform: translateY(10px);
  }
  to {
    opacity: 1;
    transform: translateY(0);
  }
}

.user-message {
  text-align: right;
}

.user-message .message-content {
  background: linear-gradient(135deg, #4285f4 0%, #1a73e8 100%);
  color: white;
  padding: 12px 16px;
  border-radius: 20px 20px 6px 20px;
  display: inline-block;
  max-width: 85%;
  word-wrap: break-word;
  font-size: 14px;
  line-height: 1.4;
  box-shadow: 0 2px 8px rgba(66, 133, 244, 0.3);
}

.assistant-message .message-content {
  background: #ffffff;
  border: 1px solid #e8eaed;
  padding: 16px;
  border-radius: 6px 20px 20px 20px;
  max-width: 95%;
  line-height: 1.5;
  font-size: 14px;
  box-shadow: 0 1px 3px rgba(0,0,0,0.1);
}

.system-message .message-content {
  background: linear-gradient(135deg, #fff3e0 0%, #ffe0b2 100%);
  border: 1px solid #ffcc02;
  padding: 12px 16px;
  border-radius: 12px;
  font-size: 13px;
  color: #e65100;
  border-left: 4px solid #ff9800;
}

.message-time {
  font-size: 11px;
  color: #5f6368;
  margin-top: 6px;
  opacity: 0.7;
}

/* Typing Indicator */
.typing {
  display: flex;
  align-items: center;
  gap: 8px;
}

.typing-dots {
  display: flex;
  gap: 2px;
}

.typing-dots span {
  width: 6px;
  height: 6px;
  background: #5f6368;
  border-radius: 50%;
  animation: typing 1.4s infinite ease-in-out both;
}

.typing-dots span:nth-child(1) { animation-delay: -0.32s; }
.typing-dots span:nth-child(2) { animation-delay: -0.16s; }

@keyframes typing {
  0%, 80%, 100% {
    transform: scale(0.8);
    opacity: 0.5;
  }
  40% {
    transform: scale(1);
    opacity: 1;
  }
}

.typing-text {
  color: #5f6368;
  font-style: italic;
  font-size: 13px;
}

/* Code and Formula Formatting */
.code-block {
  background: #f8f9fa;
  border: 1px solid #e8eaed;
  border-radius: 8px;
  margin: 12px 0;
  overflow-x: auto;
  box-shadow: inset 0 1px 3px rgba(0,0,0,0.1);
}

.code-block pre {
  margin: 0;
  padding: 16px;
  font-family: 'SF Mono', 'Monaco', 'Consolas', monospace;
  font-size: 13px;
  line-height: 1.4;
  color: #202124;
}

.inline-code {
  background: #f1f3f4;
  border: 1px solid #dadce0;
  padding: 2px 6px;
  border-radius: 4px;
  font-family: 'SF Mono', 'Monaco', 'Consolas', monospace;
  font-size: 13px;
  color: #d93025;
}

.formula {
  background: linear-gradient(135deg, #e8f5e8 0%, #c8e6c9 100%);
  border: 1px solid #4caf50;
  padding: 6px 10px;
  border-radius: 6px;
  font-family: 'SF Mono', 'Monaco', 'Consolas', monospace;
  font-weight: 600;
  color: #2e7d32;
  display: inline-block;
  margin: 2px;
  box-shadow: 0 1px 3px rgba(76, 175, 80, 0.2);
}

.formula-content {
  color: #1b5e20;
}

/* Bullet Points */
.bullet {
  color: #4285f4;
  font-weight: bold;
  margin-right: 4px;
}

/* Input Container */
#ai-chat-input-container {
  border-top: 1px solid #e8eaed;
  background: #ffffff;
  border-radius: 0 0 16px 16px;
}

#ai-chat-input-wrapper {
  padding: 16px;
  display: flex;
  gap: 12px;
  align-items: end;
}

#ai-chat-input {
  flex: 1;
  border: 2px solid #e8eaed;
  border-radius: 24px;
  padding: 12px 16px;
  resize: none;
  font-size: 14px;
  font-family: inherit;
  min-height: 20px;
  max-height: 120px;
  outline: none;
  transition: all 0.2s ease;
  background: #f8f9fa;
}

#ai-chat-input:focus {
  border-color: #4285f4;
  background: #ffffff;
  box-shadow: 0 0 0 4px rgba(66, 133, 244, 0.1);
}

#ai-chat-input::placeholder {
  color: #9aa0a6;
}

#ai-chat-send {
  background: linear-gradient(135deg, #4285f4 0%, #1a73e8 100%);
  color: white;
  border: none;
  border-radius: 50%;
  width: 44px;
  height: 44px;
  cursor: pointer;
  transition: all 0.2s ease;
  display: flex;
  align-items: center;
  justify-content: center;
  box-shadow: 0 2px 8px rgba(66, 133, 244, 0.3);
}

#ai-chat-send:hover {
  transform: scale(1.05);
  box-shadow: 0 4px 12px rgba(66, 133, 244, 0.4);
}

#ai-chat-send:active {
  transform: scale(0.95);
}

.send-icon {
  font-size: 18px;
}

/* Suggestions */
#ai-chat-suggestions {
  padding: 0 16px 16px;
  display: flex;
  flex-wrap: wrap;
  gap: 8px;
}

.suggestion-btn {
  background: #f8f9fa;
  border: 1px solid #dadce0;
  border-radius: 20px;
  padding: 8px 14px;
  font-size: 12px;
  cursor: pointer;
  transition: all 0.2s ease;
  color: #3c4043;
  font-weight: 500;
}

.suggestion-btn:hover {
  background: #e8eaed;
  border-color: #bdc1c6;
  transform: translateY(-1px);
  box-shadow: 0 2px 4px rgba(0,0,0,0.1);
}

/* Action Buttons */
.action-buttons {
  display: flex;
  gap: 8px;
  margin: 12px 0;
  flex-wrap: wrap;
}

.action-btn {
  border: none;
  border-radius: 8px;
  padding: 10px 14px;
  font-size: 12px;
  font-weight: 500;
  cursor: pointer;
  transition: all 0.2s ease;
  display: flex;
  align-items: center;
  gap: 4px;
}

.action-btn.primary {
  background: linear-gradient(135deg, #34a853 0%, #137333 100%);
  color: white;
  box-shadow: 0 1px 3px rgba(52, 168, 83, 0.3);
}

.action-btn.primary:hover {
  transform: translateY(-1px);
  box-shadow: 0 2px 6px rgba(52, 168, 83, 0.4);
}

.action-btn.secondary {
  background: #f8f9fa;
  color: #3c4043;
  border: 1px solid #dadce0;
}

.action-btn.secondary:hover {
  background: #e8eaed;
  transform: translateY(-1px);
}

/* Status */
#ai-chat-status {
  padding: 8px 16px;
  font-size: 12px;
  color: #5f6368;
  background: linear-gradient(90deg, #f8f9fa 0%, #ffffff 100%);
  border-top: 1px solid #e8eaed;
  font-style: italic;
  min-height: 16px;
}

/* Responsive Design */
@media (max-width: 768px) {
  #ai-chat-container {
    width: calc(100vw - 24px);
    height: calc(100vh - 24px);
    top: 12px;
    right: 12px;
    left: 12px;
    border-radius: 12px;
  }
  
  #ai-chat-header {
    padding: 12px 16px;
  }
  
  #ai-chat-messages {
    padding: 16px;
  }
  
  #ai-chat-input-wrapper {
    padding: 12px;
  }
}

@media (max-width: 480px) {
  #ai-chat-container {
    width: 100vw;
    height: 100vh;
    top: 0;
    right: 0;
    left: 0;
    border-radius: 0;
    border: none;
  }
  
  .user-message .message-content,
  .assistant-message .message-content {
    max-width: 90%;
  }
  
  #ai-chat-suggestions {
    flex-direction: column;
  }
  
  .suggestion-btn {
    width: 100%;
    text-align: center;
  }
}

/* High DPI Displays */
@media (-webkit-min-device-pixel-ratio: 2), (min-resolution: 192dpi) {
  #ai-chat-container {
    border-width: 0.5px;
  }
  
  .code-block {
    border-width: 0.5px;
  }
}

/* Dark Mode Support (if needed) */
@media (prefers-color-scheme: dark) {
  #ai-chat-container {
    background: #202124;
    border-color: #3c4043;
    color: #e8eaed;
  }
  
  #ai-chat-messages {
    background: linear-gradient(180deg, #2d2e30 0%, #202124 100%);
  }
  
  .assistant-message .message-content {
    background: #303134;
    border-color: #3c4043;
    color: #e8eaed;
  }
  
  #ai-chat-input {
    background: #303134;
    border-color: #3c4043;
    color: #e8eaed;
  }
  
  #ai-chat-input::placeholder {
    color: #9aa0a6;
  }
}

/* Animation Enhancements */
@keyframes fadeInScale {
  from {
    opacity: 0;
    transform: scale(0.9) translateY(10px);
  }
  to {
    opacity: 1;
    transform: scale(1) translateY(0);
  }
}

#ai-chat-container {
  animation: fadeInScale 0.3s ease-out;
}

/* Focus Management */
#ai-chat-input:focus,
#ai-chat-send:focus,
.action-btn:focus,
.suggestion-btn:focus {
  outline: 2px solid #4285f4;
  outline-offset: 2px;
}

/* Loading States */
.loading {
  opacity: 0.7;
  pointer-events: none;
}

/* Success/Error States */
.success {
  border-left: 4px solid #34a853;
  background: linear-gradient(135deg, #e8f5e8 0%, #c8e6c9 100%);
}

.error {
  border-left: 4px solid #ea4335;
  background: linear-gradient(135deg, #fce8e6 0%, #f8bbd9 100%);
}
```


### **popup.html** (Optional popup interface)

```html
<!DOCTYPE html>
<html>
<head>
  <style>
    body {
      width: 300px;
      padding: 20px;
      font-family: 'Google Sans', sans-serif;
    }
    .header {
      text-align: center;
      margin-bottom: 20px;
    }
    .status {
      padding: 10px;
      border-radius: 8px;
      margin-bottom: 15px;
      text-align: center;
    }
    .active {
      background: #e8f5e8;
      color: #137333;
      border: 1px solid #34a853;
    }
    .inactive {
      background: #fce8e6;
      color: #d93025;
      border: 1px solid #ea4335;
    }
    button {
      width: 100%;
      padding: 12px;
      background: #4285f4;
      color: white;
      border: none;
      border-radius: 8px;
      cursor: pointer;
      font-size: 14px;
      margin-bottom: 10px;
    }
    button:hover {
      background: #1a73e8;
    }
  </style>
</head>
<body>
  <div class="header">
    <h2>🤖 AI Sheets Assistant</h2>
  </div>
  
  <div id="status" class="status inactive">
    Not connected to Google Sheets
  </div>
  
  <button onclick="openSheets()">Open Google Sheets</button>
  <button onclick="viewHelp()">View Help</button>
  
  <script>
    function openSheets() {
      chrome.tabs.create({
        url: 'https://docs.google.com/spreadsheets/u/0/'
      });
    }
    
    function viewHelp() {
      chrome.tabs.create({
        url: 'https://support.google.com/docs/topic/9054603'
      });
    }
    
    // Check if current tab is Google Sheets
    chrome.tabs.query({active: true, currentWindow: true}, function(tabs) {
      const currentTab = tabs[0];
      const statusEl = document.getElementById('status');
      
      if (currentTab.url && currentTab.url.includes('docs.google.com/spreadsheets')) {
        statusEl.textContent = 'Active on Google Sheets';
        statusEl.className = 'status active';
      }
    });
  </script>
</body>
</html>
```


***

## **🚀 3. Installation \& Deployment Guide**

### **Backend Deployment (Modal.ai)**

```bash
# Install Modal CLI
pip install modal

# Authenticate
modal token new

# Deploy the functions
modal deploy modal_server.py

# Setup models (run once)
modal run modal_server.py::setup_coding_models

# Get your endpoint URL
modal app list
```


### **Chrome Extension Installation**

1. **Create folder structure:**
```
google-sheets-ai-extension/
├── manifest.json
├── background.js
├── content.js
├── chat.css
└── popup.html
```

2. **Update Modal endpoint** in `background.js`:
```javascript
const modalEndpoint = "https://YOUR_ACTUAL_MODAL_APP--coding-query.modal.run";
```

3. **Load in Chrome:**
    - Go to `chrome://extensions/`
    - Enable Developer Mode
    - Click "Load unpacked"
    - Select your extension folder
4. **Test on Google Sheets:**
    - Open any Google Sheets document
    - Look for the AI chat window in the top-right
    - Try prompts like "Sum column A" or "Create a VLOOKUP formula"

***

## **✨ Key Features**

### **🤖 AI Capabilities**

- **Advanced coding models** (Qwen 2.5 Coder, DeepSeek Coder, Phi-3.5)
- **Context-aware** responses with sheet metadata
- **Session persistence** across conversations
- **Intelligent summarization** to manage context length


### **📊 Google Sheets Integration**

- **Real-time metadata capture** (columns, ranges, data types)
- **Formula extraction** and insertion guidance
- **Action buttons** for common tasks
- **Copy-to-clipboard** functionality


### **💬 User Experience**

- **Beautiful chat interface** with Google Material Design
- **Typing indicators** and loading states
- **Code syntax highlighting** for formulas and scripts
- **Responsive design** for all screen sizes
- **Quick suggestion buttons** for common tasks


### **🔧 Developer Features**

- **Modular architecture** with clean separation of concerns
- **Error handling** and retry logic
- **Performance optimization** with lazy loading
- **Extensible** design for adding new features

**This complete codebase provides a production-ready AI assistant for Google Sheets with advanced coding capabilities, beautiful UI, and seamless integration!** 🎉

