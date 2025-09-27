
/** Purpose: Centralize utility functions to avoid code repetition.
 * Best Practices:
 * Keep functions generic (not tied to Gmail or Sheets yet).
 * Name functions clearly (parseX, getY, formatZ).
 * Avoid hardcoding Sheet IDs; pass them as parameters.
 */
var CoreFunctions = (function(){

  
// Logging helper
function logMessage(message) {
  Logger.log(`[${new Date()}] ${message}`);
}

// Email sanitization
function sanitizeText(text) {
  return text.replace(/\s+/g, ' ').trim();
}

// Date formatting
function formatDate(date) {
  return Utilities.formatDate(date, Session.getScriptTimeZone(), "yyyy-MM-dd");
}

// ✅ Extract message Id from Gmail URL (Inbox OR Label)
function extractMessageIdFromUrl(url) {
  if (!url) return null;
  // Matches both:
  // https://mail.google.com/mail/u/0/#inbox/1991a2e8d843a1c5
  // https://mail.google.com/mail/u/0/#label/Others/1991a4d022697e8c
  var match = url.match(/#(?:inbox|label\/[^/]+)\/([0-9a-f]+)/i);
  return match ? match[1] : null;
}


// Get Template by name from Gmail
function getGmailTemplateByName(templateName) {
  // Get all drafts from Gmail
  var drafts = GmailApp.getDrafts(); 
  // Find the draft with the specified name
  for (var i = 0; i < drafts.length; i++) {
    var draft = drafts[i];
    var subject = draft.getMessage().getSubject();
    if (subject.toLowerCase().trim() === templateName.toLowerCase().trim()) {
      return draft.getMessage();
    }
  }  
  // If no draft with the specified name is found, log a message to the logger
  Logger.log("Template not found");
  return null;
}

// Column Index getter from Letter 
function columnLetterToNumber(column) {
  let result = 0;
  for (let i = 0; i < column.length; i++) {
    result *= 26;
    result += column.toUpperCase().charCodeAt(i) - 'A'.charCodeAt(0) + 1;
  }
  return result;
}

// get the last non-empty row along a given column
function getLastRowFromColumn(sheet, columnLetter) {
  // Get all values from the column
  var values = sheet.getRange(columnLetter + ":" + columnLetter).getValues()
                    .map(String)  // convert to string
                    .flat();      // flatten 2D array
  // Loop backward to find the last non-empty cell
  for (var i = values.length - 1; i >= 0; i--) {
    if (values[i].trim() !== "") {
      return i + 1; // Row number (1-based index)
    }
  }
  // If column is empty, return 1 (first row)
  return 1;
}

// Function to get the key of a given value in an object
function getKeyByValue(object, value) {
  for (var key in object) {
    if (object.hasOwnProperty(key) && object[key] === value) {
      return key; // Return the key if the value matches
    }
  }
  return null; // Return null if the value is not found in the object
}

// Set Label in user properties service
function setLabelInUserProperties(userLabel, userEmail) {
  if (!userEmail || !userLabel) {
    Logger.log('Email address and label are required.');
    return false;
  }
  var userProperties = PropertiesService.getUserProperties();
  userProperties.setProperty(userEmail, userLabel);
  // Return info instead of alert 
  return {
    email: userEmail,
    label: userLabel,
    message: 'Emails from ' + userEmail + ' are assigned to auto-label ' + userLabel // caller function can use it to Display message in UI
  };
}

/**
 * Clean email body to include only readable text
 */
function parseEmailBody(rawBody) {
  if (!rawBody) return "";

  let text = rawBody;

  // Remove <style> and <script> content
  text = text.replace(/<style[\s\S]*?<\/style>/gi, '')
             .replace(/<script[\s\S]*?<\/script>/gi, '');

  // Replace <br>, <p>, <div>, <li>, headings, <tr> with line breaks
  text = text.replace(/<(br|p|div|li|tr|h[1-6])[^>]*>/gi, '\n');

  // Remove all remaining HTML tags
  text = text.replace(/<[^>]+>/g, '');

  // Decode common HTML entities
  const htmlEntities = {
    '&nbsp;': ' ',
    '&amp;': '&',
    '&lt;': '<',
    '&gt;': '>',
    '&quot;': '"',
    '&#39;': "'"
  };
  text = text.replace(/&[a-z#0-9]+;/gi, entity => htmlEntities[entity] || '');

  // Remove long tracking URLs
  text = text.replace(/https?:\/\/[^\s]+(\?upn|trackingId|&trk|&lipi|&midToken|&midSig|&otpToken)=[^\s]+/gi, '[link]');

  // Normalize whitespace
  text = text.replace(/(\r\n|\r|\n){2,}/g, '\n')
             .replace(/[ \t]{2,}/g, ' ')
             .trim();

  // Remove URLs in parentheses and standalone URLs
  text = text.replace(/\(https?:\/\/[^\s)]+\)/g, '');
  text = text.replace(/https?:\/\/\S+/g, '');

  // Collapse multiple blank lines
  text = text.replace(/\n\s*\n+/g, '\n\n');

  // Trim each line & remove empties
  let lines = text.split('\n').map(l => l.trim()).filter(l => l.length > 0);

  // Remove duplicates while preserving order
  let seen = {};
  let cleanedLines = [];
  for (let i = 0; i < lines.length; i++) {
    if (!seen[lines[i]]) {
      cleanedLines.push(lines[i]);
      seen[lines[i]] = true;
    }
  }
  text = cleanedLines.join('\n');

  // 🟢 NEW: Keep only up to the first `>>` block (cut thread)
  let cutoff = text.search(/\n>>/);  // find the first ">>"
  if (cutoff !== -1) {
    text = text.substring(0, cutoff).trim();
  }

  // Limit output to 10,000 characters
  if (text.length > 2000) {
    text = text.slice(0, 2000) + "\n\n[Content truncated]";
  }

  return text;
}



return {
    logMessage: logMessage,
    sanitizeText: sanitizeText,
    formatDate: formatDate,
    extractMessageIdFromUrl: extractMessageIdFromUrl,
    getGmailTemplateByName: getGmailTemplateByName,
    columnLetterToNumber: columnLetterToNumber,
    getLastRowFromColumn: getLastRowFromColumn,
    getKeyByValue: getKeyByValue,
    setLabelInUserProperties: setLabelInUserProperties,
    parseEmailBody: parseEmailBody
  };

})();




// ++++++++++++++++++++++++++++++++++ end of code ++++++++++++++++++++++++++++++++++++++++++++++++



