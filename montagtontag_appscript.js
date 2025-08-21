/**
 * TELEGRAM SCHEDULED MESSAGES BOT
 * 
 * This script reads a Google Spreadsheet and sends scheduled Telegram messages
 * based on date/time and repeat patterns.
 * 
 * Spreadsheet columns:
 * A: DateTime - When to send the message
 * B: Message - The message content to send
 * C: Repeat - How often to repeat (daily, weekly, every X days, etc.)
 * D: Enabled - Whether the message is active (✅, true, yes)
 * E: Last Sent - Timestamp of when message was last sent (auto-updated)
 */

// =============================================================================
// CONFIGURATION
// =============================================================================

// Spreadsheet column indices (0-based)
const COL_DATETIME = 0;   // Column A: Date/Time to send
const COL_MESSAGE = 1;    // Column B: Message content
const COL_REPEAT = 2;     // Column C: Repeat pattern
const COL_ENABLED = 3;    // Column D: Enabled status
const COL_LAST_SENT = 4;  // Column E: Last sent timestamp

// Telegram API credentials from PropertiesService (secure storage)
// Setup: Go to Project Settings > Script Properties and add:
// - TELEGRAM_TOKEN = your_bot_token
// - CHAT_ID = your_production_chat_id  
// - DEBUG_CHAT_ID = your_debug_chat_id
const TOKEN = PropertiesService.getScriptProperties().getProperty('TELEGRAM_TOKEN');
const CHAT_ID = PropertiesService.getScriptProperties().getProperty('CHAT_ID');
const DEBUG_CHAT_ID = PropertiesService.getScriptProperties().getProperty('DEBUG_CHAT_ID');

// =============================================================================
// DEBUG FUNCTIONS
// =============================================================================

/**
 * Debug function to test the scheduling logic
 * Sends detailed analysis of each enabled message to the debug chat
 * Shows what would happen based on current time and scheduling rules
 */
function debugTest() {
  console.log('=== DEBUG TEST STARTED ===');
  console.log('Current time:', new Date());
  
  // Test PropertiesService
  console.log('Available script properties:', PropertiesService.getScriptProperties().getProperties());
  console.log('TOKEN from properties:', PropertiesService.getScriptProperties().getProperty('TELEGRAM_TOKEN') ? 'Found' : 'Not found');
  console.log('CHAT_ID from properties:', PropertiesService.getScriptProperties().getProperty('CHAT_ID') ? 'Found' : 'Not found');
  
  // Test actual values being used
  console.log('Using TOKEN:', TOKEN ? `${TOKEN.substring(0, 10)}...` : 'null');
  console.log('Using CHAT_ID:', CHAT_ID);
  
  /**
   * Helper function to send debug messages to Telegram
   * @param {string} message - The message to send
   * @param {boolean} isSystemMessage - Whether this is a system status message
   * @returns {boolean} - True if sent successfully
   */
  function sendDebugMessage(message, isSystemMessage = false) {
    const url = `https://api.telegram.org/bot${TOKEN}/sendMessage`;
    const payload = {
      chat_id: DEBUG_CHAT_ID,
      text: message
      // Removed parse_mode to use plain text and avoid Markdown parsing issues
    };
    
    try {
      const response = UrlFetchApp.fetch(url, {
        method: 'post',
        contentType: 'application/json',
        payload: JSON.stringify(payload)
      });
      
      const responseData = JSON.parse(response.getContentText());
      
      if (responseData.ok) {
        console.log(isSystemMessage ? '✅ System debug message sent' : '✅ Message sent successfully');
        return true;
      } else {
        console.error('❌ Telegram API error:', responseData.description);
        return false;
      }
      
    } catch (e) {
      console.error('❌ Network/Script error:', e.message);
      return false;
    }
  }
  
  // Send initial system status message
  const systemMessage = `🧪 DEBUG TEST STARTED - ${new Date().toLocaleString()}\n\n` +
                        `✅ Script is running\n` +
                        `🔑 TOKEN: ${TOKEN ? 'Found' : 'Missing'}\n` +
                        `💬 CHAT_ID: ${CHAT_ID}\n` +
                        `📊 Properties working: ${PropertiesService.getScriptProperties().getProperty('TELEGRAM_TOKEN') ? 'Yes' : 'No'}\n\n` +
                        `📋 Now analyzing all enabled messages from spreadsheet...`;
  
  console.log('Sending system debug message...');
  sendDebugMessage(systemMessage, true);
  
  // Brief delay before processing spreadsheet to avoid rate limiting
  Utilities.sleep(1000);
  
  // Process spreadsheet and analyze each enabled message
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const data = sheet.getDataRange().getValues();
    
    console.log('Processing spreadsheet with', data.length - 1, 'rows');
    
    let enabledCount = 0;
    let sentCount = 0;
    
    // Process each row (skip header row at index 0)
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      
      // Skip rows without a datetime value
      if (!row[COL_DATETIME]) continue;
      
      // Extract row data
      const start = new Date(row[COL_DATETIME]);
      const msg = row[COL_MESSAGE];
      const repeatRaw = row[COL_REPEAT] || '';
      const enabled = row[COL_ENABLED];
      
      // Check if message is enabled (supports multiple formats)
      const isEnabled = enabled === true || 
                       enabled === '✅' || 
                       (typeof enabled === 'string' && (
                         enabled.toLowerCase().trim() === 'yes' ||
                         enabled.toLowerCase().trim() === 'true' ||
                         enabled.includes('✅')
                       ));
      
      if (msg && isEnabled) {
        enabledCount++;
        
        // Analyze scheduling logic (same as main function)
        const now = new Date();
        now.setSeconds(0, 0);
        
        // Parse Last Sent, handling error messages
        let lastSent = null;
        if (row[COL_LAST_SENT]) {
          const lastSentValue = row[COL_LAST_SENT];
          if (typeof lastSentValue === 'string' && lastSentValue.startsWith('ERROR:')) {
            lastSent = null; // Treat errors as no previous send
          } else {
            try {
              lastSent = new Date(lastSentValue);
            } catch (e) {
              lastSent = null;
            }
          }
        }
        
        // Duplicate prevention check
        let duplicateCheck = 'PASS';
        if (lastSent) {
          const lastSentMinute = new Date(lastSent.getFullYear(), lastSent.getMonth(), lastSent.getDate(), lastSent.getHours(), lastSent.getMinutes());
          const currentMinute = new Date(now.getFullYear(), now.getMonth(), now.getDate(), now.getHours(), now.getMinutes());
          
          if (currentMinute.getTime() === lastSentMinute.getTime()) {
            duplicateCheck = 'BLOCKED - Already sent this minute';
          }
        }
        
        // Time and date matching logic
        const exactDateMatch = start.getFullYear() === now.getFullYear() &&
                               start.getMonth() === now.getMonth() &&
                               start.getDate() === now.getDate();
        
        // More forgiving time match - allows for trigger delays within the same minute
        const timeMatch = now.getHours() === start.getHours() &&
                          now.getMinutes() === start.getMinutes();
        
        const day = now.getDay(); // 0=Sunday, 1=Monday, etc.
        const mDay = now.getDate();
        const month = now.getMonth();
        const diffDays = Math.floor((now - start) / (1000 * 60 * 60 * 24));
        
        const repeat = repeatRaw.toLowerCase().trim();
        const isOneTime = repeat === '' || repeat === "don't" || repeat === "don't repeat";
        
        // Determine if message should send
        const shouldSend =
          (isOneTime && exactDateMatch && timeMatch) || // Removed !lastSent check - ignore previous sends for one-time messages
          (repeat === 'every minute' && (!lastSent || (now.getTime() - lastSent.getTime()) >= 60000)) ||
          (repeat === 'daily' && timeMatch) ||
          (repeat === 'weekly' && timeMatch && day === start.getDay()) ||
          (repeat === 'monthly' && timeMatch && mDay === start.getDate()) ||
          (repeat === 'yearly' && timeMatch && mDay === start.getDate() && month === start.getMonth()) ||
          (repeat === 'every 2 days' && timeMatch && diffDays >= 0 && diffDays % 2 === 0) ||
          (repeat === 'every 3 days' && timeMatch && diffDays >= 0 && diffDays % 3 === 0) ||
          (repeat === 'every 7 days' && timeMatch && diffDays >= 0 && diffDays % 7 === 0) ||
          (repeat === 'every 14 days' && timeMatch && diffDays >= 0 && diffDays % 14 === 0) ||
          (repeat === 'monday' && timeMatch && day === 1) ||
          (repeat === 'tuesday' && timeMatch && day === 2) ||
          (repeat === 'wednesday' && timeMatch && day === 3) ||
          (repeat === 'thursday' && timeMatch && day === 4) ||
          (repeat === 'friday' && timeMatch && day === 5) ||
          (repeat === 'saturday' && timeMatch && day === 6) ||
          (repeat === 'sunday' && timeMatch && day === 0);
        
        // Create detailed analysis
        const dayNames = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
        const currentDayName = dayNames[day];
        const scheduledDayName = dayNames[start.getDay()];
        
        let scheduleAnalysis = '';
        const scheduledTime = `${start.getHours().toString().padStart(2, '0')}:${start.getMinutes().toString().padStart(2, '0')}`;
        const currentTime = `${now.getHours().toString().padStart(2, '0')}:${now.getMinutes().toString().padStart(2, '0')}`;
        
        if (isOneTime) {
          const scheduledDate = start.toDateString();
          const currentDate = now.toDateString();
          
          if (exactDateMatch && timeMatch) {
            scheduleAnalysis = `One-time message scheduled for today at ${scheduledTime} - Ready to send! ${lastSent ? '(Will resend despite previous send)' : ''}`;
          } else if (!exactDateMatch) {
            scheduleAnalysis = `One-time message scheduled for ${scheduledDate}, today is ${currentDate} - Date mismatch ❌`;
          } else if (!timeMatch) {
            scheduleAnalysis = `One-time message scheduled for ${scheduledTime} - Current time is ${currentTime} ❌`;
          } else {
            scheduleAnalysis = `One-time message: Date match ${exactDateMatch ? '✅' : '❌'}, Time match ${timeMatch ? '✅' : '❌'} (Last Sent ignored)`;
          }
        } else if (repeat === 'every minute') {
          const timeSinceLastSent = lastSent ? Math.floor((now.getTime() - lastSent.getTime()) / 60000) : null;
          if (!lastSent) {
            scheduleAnalysis = `Every minute - Never sent before, ready to send`;
          } else if (timeSinceLastSent >= 1) {
            scheduleAnalysis = `Every minute - Last sent ${timeSinceLastSent} minutes ago, ready to send`;
          } else {
            scheduleAnalysis = `Every minute - Sent less than 1 minute ago, waiting`;
          }
        } else if (repeat === 'daily') {
          scheduleAnalysis = `Daily at ${scheduledTime} - Current time is ${currentTime} ${timeMatch ? '✅' : '❌'}`;
        } else if (repeat === 'weekly') {
          const dayMatch = day === start.getDay();
          scheduleAnalysis = `Weekly on ${scheduledDayName}s at ${scheduledTime} - Today is ${currentDayName} ${dayMatch ? '✅' : '❌'}, time is ${currentTime} ${timeMatch ? '✅' : '❌'}`;
        } else if (repeat.includes('monday') || repeat.includes('tuesday') || repeat.includes('wednesday') || repeat.includes('thursday') || repeat.includes('friday') || repeat.includes('saturday') || repeat.includes('sunday')) {
          const targetDay = repeat.charAt(0).toUpperCase() + repeat.slice(1);
          const dayMatch = currentDayName === targetDay;
          scheduleAnalysis = `Every ${targetDay} at ${scheduledTime} - Today is ${currentDayName} ${dayMatch ? '✅' : '❌'}, time is ${currentTime} ${timeMatch ? '✅' : '❌'}`;
        } else if (repeat.includes('days')) {
          const dayInterval = repeat.match(/every (\d+) days/);
          if (dayInterval) {
            const interval = parseInt(dayInterval[1]);
            const remainder = diffDays % interval;
            const dayMatch = remainder === 0 && diffDays >= 0;
            scheduleAnalysis = `Every ${interval} days at ${scheduledTime} - ${diffDays} days since start ${dayMatch ? '✅' : '❌'}, time is ${currentTime} ${timeMatch ? '✅' : '❌'}`;
          }
        } else if (repeat === 'monthly') {
          const dayMatch = mDay === start.getDate();
          scheduleAnalysis = `Monthly on day ${start.getDate()} at ${scheduledTime} - Today is day ${mDay} ${dayMatch ? '✅' : '❌'}, time is ${currentTime} ${timeMatch ? '✅' : '❌'}`;
        } else if (repeat === 'yearly') {
          const dayMatch = mDay === start.getDate() && month === start.getMonth();
          scheduleAnalysis = `Yearly on ${start.toDateString().slice(4, 10)} at ${scheduledTime} - Today matches ${dayMatch ? '✅' : '❌'}, time is ${currentTime} ${timeMatch ? '✅' : '❌'}`;
        } else {
          scheduleAnalysis = `"${repeat}" pattern - Time check: ${scheduledTime} vs ${currentTime} ${timeMatch ? '✅' : '❌'}`;
        }
        
        const debugMsg = `ROW ${i + 1} - ${repeatRaw || 'One-time'}\n\n` +
                        `SCHEDULE CHECK:\n` +
                        `${scheduleAnalysis}\n\n` +
                        `DUPLICATE CHECK:\n` +
                        `${duplicateCheck === 'PASS' ? 'No recent send detected' : duplicateCheck}\n\n` +
                        `RESULT: ${shouldSend ? '✅ WILL SEND' : '❌ WILL NOT SEND'}\n\n` +
                        `Message: "${msg.substring(0, 100)}${msg.length > 100 ? '...' : ''}"`;
        
        console.log(`Sending analysis for row ${i + 1}`);
        
        if (sendDebugMessage(debugMsg)) {
          sentCount++;
        }
        
        // Small delay between messages to avoid rate limiting
        Utilities.sleep(800);
      }
    }
    
    // Send summary message
    const summaryMessage = `📊 DEBUG TEST COMPLETE\n\n` +
                          `📋 Total rows processed: ${data.length - 1}\n` +
                          `✅ Enabled messages found: ${enabledCount}\n` +
                          `📤 Messages sent: ${sentCount}\n` +
                          `🕐 Test completed at: ${new Date().toLocaleString()}`;
    
    sendDebugMessage(summaryMessage, true);
    
    console.log('=== DEBUG TEST COMPLETED ===');
    return `Debug test completed. Sent ${sentCount} of ${enabledCount} enabled messages.`;
    
  } catch (e) {
    const errorMessage = `❌ DEBUG TEST ERROR\n\nError reading spreadsheet: ${e.message}`;
    sendDebugMessage(errorMessage, true);
    console.error('Debug test error:', e.message);
    return `Error: ${e.message}`;
  }
}

// =============================================================================
// MAIN SCHEDULER FUNCTION
// =============================================================================

/**
 * Main function that processes the spreadsheet and sends scheduled messages
 * This function should be triggered by Google Apps Script time-based triggers
 * 
 * Process:
 * 1. Read all rows from the active spreadsheet
 * 2. For each enabled message, check if it should send based on:
 *    - Current date/time vs scheduled date/time
 *    - Repeat pattern (daily, weekly, one-time, etc.)
 *    - Duplicate prevention (except for one-time messages)
 * 3. Send qualifying messages via Telegram API
 * 4. Update Last Sent column with timestamp or error message
 */
function sendScheduledMessages() {
  try {
    console.log('Script started at:', new Date());
    
    // Random delay to reduce collision risk when multiple triggers run simultaneously
    Utilities.sleep(Math.floor(Math.random() * 500));
    
    // Get spreadsheet data and normalize current time
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const data = sheet.getDataRange().getValues();
    const now = new Date();
    now.setSeconds(0, 0); // Ignore seconds for more forgiving time matching
    
    console.log('Current time:', now);
    console.log('Processing', data.length - 1, 'rows');
    
    // Process each row (skip header row at index 0)
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      
      // Skip rows without a datetime value
      if (!row[COL_DATETIME]) continue;
      
      // Extract and normalize row data
      const start = new Date(row[COL_DATETIME]);
      start.setSeconds(0, 0); // Ignore seconds for time matching
      const msg = row[COL_MESSAGE];
      const repeatRaw = row[COL_REPEAT] || '';
      const enabled = row[COL_ENABLED];
      
      // Parse Last Sent column, handling error messages gracefully
      let lastSent = null;
      if (row[COL_LAST_SENT]) {
        const lastSentValue = row[COL_LAST_SENT];
        if (typeof lastSentValue === 'string' && lastSentValue.startsWith('ERROR:')) {
          lastSent = null; // Treat errors as no previous send
        } else {
          try {
            lastSent = new Date(lastSentValue);
          } catch (e) {
            console.warn(`Invalid date in row ${i + 1}, treating as no previous send:`, lastSentValue);
            lastSent = null;
          }
        }
      }
      
      console.log(`Row ${i + 1}: DateTime=${start}, Repeat=${repeatRaw}, Enabled=${enabled}`);
      
      // Check if message is enabled (supports multiple formats)
      const isEnabled = enabled === true || 
                       enabled === '✅' || 
                       (typeof enabled === 'string' && (
                         enabled.toLowerCase().trim() === 'yes' ||
                         enabled.toLowerCase().trim() === 'true' ||
                         enabled.includes('✅')
                       ));
      
            // Skip if no message or not enabled
      if (!msg || !isEnabled) {
        console.log(`Skipping row ${i + 1}: msg=${!!msg}, enabled=${isEnabled}`);
        continue;
      }
      
      // Prevent duplicate sends within the same minute (for repeating messages)
      if (lastSent) {
        const lastSentMinute = new Date(lastSent.getFullYear(), lastSent.getMonth(), lastSent.getDate(), lastSent.getHours(), lastSent.getMinutes());
        const currentMinute = new Date(now.getFullYear(), now.getMonth(), now.getDate(), now.getHours(), now.getMinutes());
        
        if (currentMinute.getTime() === lastSentMinute.getTime()) {
          console.log(`Skipping row ${i + 1}: Already sent this minute`);
        continue;
        }
      }
      
      // Calculate date and time matching conditions
      const exactDateMatch = start.getFullYear() === now.getFullYear() &&
                             start.getMonth() === now.getMonth() &&
                             start.getDate() === now.getDate();
      
      // Time matching (ignores seconds for forgiving trigger timing)
      const timeMatch = now.getHours() === start.getHours() &&
                        now.getMinutes() === start.getMinutes();
      
      // Calculate time-based variables for repeat logic
      const day = now.getDay(); // 0=Sunday, 1=Monday, etc.
      const mDay = now.getDate();
      const month = now.getMonth();
      const diffDays = Math.floor((now - start) / (1000 * 60 * 60 * 24));
      
      // Parse repeat pattern
      const repeat = repeatRaw.toLowerCase().trim();
      const isOneTime = repeat === '' || repeat === "don't" || repeat === "don't repeat";
      
      console.log(`Row ${i + 1}: timeMatch=${timeMatch}, repeat=${repeat}, diffDays=${diffDays}`);
      console.log(`Row ${i + 1}: exactDateMatch=${exactDateMatch}, lastSent=${lastSent}`);
      
      // Determine if message should be sent based on scheduling rules
      const shouldSend =
        // One-time messages: send if date and time match (ignore previous sends)
        (isOneTime && exactDateMatch && timeMatch) ||
        // Every minute: send if never sent or at least 1 minute has passed
        (repeat === 'every minute' && (!lastSent || (now.getTime() - lastSent.getTime()) >= 60000)) ||
        // Daily: send if time matches
        (repeat === 'daily' && timeMatch) ||
        // Weekly: send if time and weekday match
        (repeat === 'weekly' && timeMatch && day === start.getDay()) ||
        // Monthly: send if time and day of month match
        (repeat === 'monthly' && timeMatch && mDay === start.getDate()) ||
        // Yearly: send if time, day, and month match
        (repeat === 'yearly' && timeMatch && mDay === start.getDate() && month === start.getMonth()) ||
        // Interval-based repeats: every X days
        (repeat === 'every 2 days' && timeMatch && diffDays >= 0 && diffDays % 2 === 0) ||
        (repeat === 'every 3 days' && timeMatch && diffDays >= 0 && diffDays % 3 === 0) ||
        (repeat === 'every 7 days' && timeMatch && diffDays >= 0 && diffDays % 7 === 0) ||
        (repeat === 'every 14 days' && timeMatch && diffDays >= 0 && diffDays % 14 === 0) ||
        // Specific weekdays: send if time and weekday match
        (repeat === 'monday' && timeMatch && day === 1) ||
        (repeat === 'tuesday' && timeMatch && day === 2) ||
        (repeat === 'wednesday' && timeMatch && day === 3) ||
        (repeat === 'thursday' && timeMatch && day === 4) ||
        (repeat === 'friday' && timeMatch && day === 5) ||
        (repeat === 'saturday' && timeMatch && day === 6) ||
        (repeat === 'sunday' && timeMatch && day === 0);
      
      console.log(`Row ${i + 1}: shouldSend=${shouldSend}`);
      
      // Send message if scheduling conditions are met
      if (shouldSend) {
        console.log(`Sending message for row ${i + 1}:`, msg.substring(0, 50) + '...');
        
        // Prepare Telegram API request
        const url = `https://api.telegram.org/bot${TOKEN}/sendMessage`;
        const payload = {
          chat_id: CHAT_ID,
          text: msg,
          parse_mode: 'Markdown'
        };
        
        try {
          // Send message via Telegram API
          const response = UrlFetchApp.fetch(url, {
            method: 'post',
            contentType: 'application/json',
            payload: JSON.stringify(payload)
          });
          
          const responseData = JSON.parse(response.getContentText());
          console.log('Telegram API response:', responseData);
          
          if (responseData.ok) {
            // Update Last Sent column with current timestamp
            sheet.getRange(i + 1, COL_LAST_SENT + 1).setValue(now);
            console.log(`Successfully sent message for row ${i + 1}`);
          } else {
            throw new Error(responseData.description || 'Unknown Telegram API error');
          }
          
        } catch (e) {
          console.error(`Error sending message for row ${i + 1}:`, e.message);
          // Store error message in Last Sent column (prefixed to avoid date parsing issues)
          const errorMsg = `ERROR: ${e.message}`;
          sheet.getRange(i + 1, COL_LAST_SENT + 1).setValue(errorMsg);
        }
      }
    }
    
    console.log('Script completed at:', new Date());
    
  } catch (err) {
    // Log any unexpected errors that occur during script execution
    console.error("Script error:", err.message);
    console.error("Stack trace:", err.stack);
    // Note: Error is logged but not re-thrown to prevent trigger failures
  }
}