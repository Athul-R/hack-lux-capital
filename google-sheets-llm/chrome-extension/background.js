// Background service worker for Chrome extension
chrome.runtime.onMessage.addListener((msg, sender, sendResponse) => {
  if (msg.type === "PROMPT_TO_SERVER") {
    handleAIQuery(msg.payload)
      .then(data => sendResponse({ success: true, response: data }))
      .catch(err => sendResponse({ success: false, error: err.toString() }));
    return true; // Keep sendResponse alive for async
  }
});
// https://jaydaftari19--coding-query.modal.run    
// https://athul-r--coding-query.modal.run" 
async function handleAIQuery(payload) {
  const modalEndpoint = "https://jaydaftari19--coding-query.modal.run";
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

// Open Google Sheets on install
chrome.runtime.onInstalled.addListener((details) => {
  if (details.reason === 'install') {
    chrome.tabs.create({ url: 'https://docs.google.com/spreadsheets/u/0/' });
  }
});


