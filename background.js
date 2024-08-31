
// Create context menu
chrome.runtime.onInstalled.addListener(() => {
  chrome.contextMenus.create({
    id: "exportGrades",
    title: "成績一覧をcsvで書き出す",
    contexts: ["page"]
  });
});

// Execute script and send message when context menu item is clicked
chrome.contextMenus.onClicked.addListener((info, tab) => {
  if (info.menuItemId === "exportGrades") {
    chrome.scripting.executeScript({
      target: { tabId: tab.id },
      files: ['content.js']
    }, () => {
      chrome.tabs.sendMessage(tab.id, { action: "exportCSV" });
    });
  }
});
