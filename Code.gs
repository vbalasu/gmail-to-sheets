function processUnreadEmails() {
  // Get the active spreadsheet and sheet
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  
  // Get Gmail API with read-only scope
  const gmail = Gmail.Users.Messages;
  
  // Get all unread emails
  const threads = Gmail.Users.Threads.list('me', {
    q: 'is:unread category:primary'
  });
  
  // Process each email thread
  if (threads.threads) {
    threads.threads.forEach(thread => {
      const threadDetails = Gmail.Users.Threads.get('me', thread.id);
      
      threadDetails.messages.forEach(message => {
        // Get full message details
        const messageDetails = Gmail.Users.Messages.get('me', message.id);
        
        // Extract email data
        const headers = messageDetails.payload.headers;
        const subject = headers.find(h => h.name === 'Subject')?.value || '';
        const sender = headers.find(h => h.name === 'From')?.value || '';
        const date = new Date(parseInt(messageDetails.internalDate));
        
        // Get message body
        let body = '';
        if (messageDetails.payload.parts) {
          const textPart = messageDetails.payload.parts.find(part => part.mimeType === 'text/plain');
          if (textPart) {
            body = Utilities.newBlob(Utilities.base64Decode(textPart.body.data)).getDataAsString();
          }
        } else if (messageDetails.payload.body.data) {
          body = Utilities.newBlob(Utilities.base64Decode(messageDetails.payload.body.data)).getDataAsString();
        }
        
        // Add data to sheet
        sheet.appendRow([
          date,
          sender,
          subject,
          body
        ]);
      });
    });
  }
}

// Create a menu item to run the script
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Gmail Processing')
    .addItem('Process Unread Emails', 'processUnreadEmails')
    .addToUi();
} 