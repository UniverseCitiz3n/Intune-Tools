// Background script for the Intune Tools Extension

// Listen for messages from content scripts
chrome.runtime.onMessage.addListener((request, sender, sendResponse) => {
  if (request.action === 'runScript') {
    // Handle any background processing if needed
    sendResponse({ success: true, message: 'Script executed in background' });
    return true; // Required for async response
  }
});

// Extension installation or update handler
chrome.runtime.onInstalled.addListener((details) => {
  if (details.reason === 'install') {
    console.log('Intune Tools Extension installed');
  } else if (details.reason === 'update') {
    console.log('Intune Tools Extension updated');
  }
});
