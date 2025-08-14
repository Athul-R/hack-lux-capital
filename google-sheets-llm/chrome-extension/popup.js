// Popup script for Chrome extension
document.addEventListener('DOMContentLoaded', function() {
  // Button event listeners
  document.getElementById('open-sheets-btn').addEventListener('click', openSheets);
  document.getElementById('view-help-btn').addEventListener('click', viewHelp);
  
  // Check connection status on load
  checkConnectionStatus();
});

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

function checkConnectionStatus() {
  // Check if current tab is Google Sheets
  chrome.tabs.query({active: true, currentWindow: true}, function(tabs) {
    const currentTab = tabs[0];
    const statusEl = document.getElementById('status');
    
    if (chrome.runtime.lastError) {
      console.error('Error querying tabs:', chrome.runtime.lastError);
      statusEl.textContent = 'Permission error - please reload extension';
      statusEl.className = 'status inactive';
      return;
    }
    
    if (currentTab && currentTab.url && currentTab.url.includes('docs.google.com/spreadsheets')) {
      statusEl.textContent = 'Connected to Google Sheets';
      statusEl.className = 'status active';
      
      // Try to communicate with content script to verify it's working
      chrome.tabs.sendMessage(currentTab.id, {type: 'PING'}, function(response) {
        if (chrome.runtime.lastError) {
          statusEl.textContent = 'Google Sheets detected - AI loading...';
          statusEl.className = 'status active';
        } else {
          statusEl.textContent = 'AI Assistant Active';
          statusEl.className = 'status active';
        }
      });
    } else {
      statusEl.textContent = 'Not on Google Sheets';
      statusEl.className = 'status inactive';
    }
  });
}
