function processUnreadEmails() {
  // Get the active spreadsheet and sheet
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  
  // Get all unread emails
  const threads = GmailApp.search('is:unread category:primary');
  sheet.clear();
  sheet.appendRow(['date','sender','subject','body']);
  
  // Process each email thread
  threads.forEach(thread => {
    const messages = thread.getMessages();
    
    messages.forEach(message => {
      // Extract email data
      const subject = message.getSubject();
      const sender = message.getFrom();
      const date = message.getDate();
      const body = message.getPlainBody();
      
      // Add data to sheet
      sheet.appendRow([date,sender,subject,body]);
    });
  });
}

// Create a menu item to run the script
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Gmail Processing')
    .addItem('Process Unread Emails', 'processUnreadEmails')
    .addToUi();
}

// Test function to debug email processing
function testEmailProcessing() {
  const threads = GmailApp.search('is:unread category:primary');
  
  if (threads.length > 0) {
    const firstThread = threads[0];
    Logger.log('First thread ID: ' + firstThread.getId());
    
    const messages = firstThread.getMessages();
    if (messages.length > 0) {
      const firstMessage = messages[0];
      Logger.log('Message subject: ' + firstMessage.getSubject());
      Logger.log('Message body: ' + firstMessage.getPlainBody());
    }
  } else {
    Logger.log('No unread threads found');
  }
} 