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
// - AG_CHAT_ID = chat_id for the AG helper poll
const TOKEN = PropertiesService.getScriptProperties().getProperty('TELEGRAM_TOKEN');
const CHAT_ID = PropertiesService.getScriptProperties().getProperty('CHAT_ID');
const DEBUG_CHAT_ID = PropertiesService.getScriptProperties().getProperty('DEBUG_CHAT_ID');
const AG_CHAT_ID = PropertiesService.getScriptProperties().getProperty('AG_CHAT_ID');

// =============================================================================
// SECURITY CONFIGURATION
// =============================================================================

// Domain whitelist - only URLs from these domains will be allowed
// Leave empty array [] to allow all domains (not recommended)
// Add trusted domains like: ['telegram.org', 'github.com', 'example.com']
const ALLOWED_DOMAINS = [];

// Domain blacklist - URLs from these domains will be blocked
// Add suspicious or known spam domains here
const BLOCKED_DOMAINS = [
  'bit.ly',
  'tinyurl.com',
  't.co',
  'goo.gl',
  'short.link',
  'rebrand.ly',
  'cutt.ly',
  't.me'  // Telegram invite links - often used for spam
];

// Enable strict mode: if true, messages with URLs not in whitelist will be blocked
// If false, only blacklisted domains will be blocked
const STRICT_MODE = false;

// =============================================================================
// RATE LIMITING CONFIGURATION
// =============================================================================

// Maximum number of messages that can be sent per hour
const MAX_MESSAGES_PER_HOUR = 3

// Maximum number of messages that can be sent per day
const MAX_MESSAGES_PER_DAY = 5

// =============================================================================
// CONTENT VALIDATION CONFIGURATION
// =============================================================================

// Maximum message length (Telegram limit is 4096 characters)
const MAX_MESSAGE_LENGTH = 4096;

// Maximum number of URLs allowed per message
const MAX_URLS_PER_MESSAGE = 10;

// Block message if it contains more than this many URLs (stricter check)
const MAX_SUSPICIOUS_URLS = 1;

// Debug test: only analyze this many upcoming messages (by schedule time)
const DEBUG_TEST_MAX_MESSAGES = 5;

// Suspicious keywords that may indicate spam (English and German)
const SUSPICIOUS_KEYWORDS = [
  // English spam keywords
  'click here',
  'limited time',
  'act now',
  'urgent',
  'free money',
  'guaranteed',
  'no risk',
  'limited offer',
  'exclusive deal',
  'act fast',
  'don\'t miss out',
  'one time only',
  // German spam keywords (from actual spam received)
  'wichtiges update',
  'nutzen sie die gelegenheit',
  'solange sie noch besteht',
  'finanzielles wachstum',
  'verpassen sie sie nicht',
  'bevor der link abläuft',
  'entscheidender wendepunkt',
  'finanzielle freiheit',
  'passives einkommen',
  'jetzt handeln',
  'begrenzte plätze',
  'exklusives angebot',
  'schnell reich',
  'geld verdienen'
];

// Threshold for excessive capitalization (0.0 to 1.0, where 1.0 = 100% caps)
const EXCESSIVE_CAPS_THRESHOLD = 0.5; // 50% caps = suspicious

// Maximum consecutive punctuation marks before flagging as suspicious
const EXCESSIVE_PUNCTUATION_THRESHOLD = 3; // 3+ consecutive punctuation marks

// =============================================================================
// REQUIRED KEYWORD VALIDATION CONFIGURATION
// =============================================================================

// Enable required keyword validation: if true, messages must contain at least one required keyword
const REQUIRE_KEYWORD_VALIDATION = true;

// Required keywords/phrases that legitimate messages should contain
// At least ONE keyword must be present in the message (case-insensitive matching)
// Keywords support partial matches (e.g., "Montag" matches "MontagTontag")
const REQUIRED_KEYWORDS = [
  'test',  // For testing purposes
  'montag',
  'montagtontag',
  'tontag',
  'glasiertag',
  'luma.com',
  'melde dich',
  'zeitslot',
  'aufbau',
  'abbau',
  'stornieren',
  'fertigen werke',
  'helfende hände',
  'töpfern',
  'glasieren'
];

// =============================================================================
// SECURITY & VALIDATION FUNCTIONS
// =============================================================================

/**
 * Check rate limiting to prevent too many messages in a short time
 * Uses CacheService to track message counts per hour and per day.
 * Note: CacheService is best-effort and values may reset if the cache is cleared.
 * NOTE: This only CHECKS the limit, does NOT increment counters
 * @returns {Object} - Rate limit check result with isAllowed and reason
 */
function checkRateLimit() {
  const now = new Date();
  
  // Use script cache for transient rate limit counters
  const cache = CacheService.getScriptCache();
  const hourKey = `rate_hour_${now.getFullYear()}_${now.getMonth()}_${now.getDate()}_${now.getHours()}`;
  const dayKey = `rate_day_${now.getFullYear()}_${now.getMonth()}_${now.getDate()}`;
  
  // Get current counts
  const hourCount = parseInt(cache.get(hourKey) || '0');
  const dayCount = parseInt(cache.get(dayKey) || '0');
  
  // Check hourly limit
  if (hourCount >= MAX_MESSAGES_PER_HOUR) {
    return {
      isAllowed: false,
      reason: `Rate limit exceeded: ${hourCount} messages sent this hour (limit: ${MAX_MESSAGES_PER_HOUR})`,
      hourCount: hourCount,
      dayCount: dayCount
    };
  }
  
  // Check daily limit
  if (dayCount >= MAX_MESSAGES_PER_DAY) {
    return {
      isAllowed: false,
      reason: `Rate limit exceeded: ${dayCount} messages sent today (limit: ${MAX_MESSAGES_PER_DAY})`,
      hourCount: hourCount,
      dayCount: dayCount
    };
  }
  
  return {
    isAllowed: true,
    reason: `Rate limit OK (${hourCount}/${MAX_MESSAGES_PER_HOUR} per hour, ${dayCount}/${MAX_MESSAGES_PER_DAY} per day)`,
    hourCount: hourCount,
    dayCount: dayCount
  };
}

/**
 * Increment rate limit counters after a message is successfully sent
 * Call this ONLY after a message has been successfully sent
 */
function incrementRateLimit() {
  const now = new Date();
  
  // Use script cache for transient rate limit counters
  const cache = CacheService.getScriptCache();
  const hourKey = `rate_hour_${now.getFullYear()}_${now.getMonth()}_${now.getDate()}_${now.getHours()}`;
  const dayKey = `rate_day_${now.getFullYear()}_${now.getMonth()}_${now.getDate()}`;
  
  // Get current counts and increment
  const hourCount = parseInt(cache.get(hourKey) || '0');
  const dayCount = parseInt(cache.get(dayKey) || '0');

  // Cache lifetimes: 1 hour for hourly counter, up to 6 hours for daily counter (Apps Script max)
  const newHourCount = hourCount + 1;
  const newDayCount = dayCount + 1;
  cache.put(hourKey, newHourCount.toString(), 60 * 60);      // 1 hour
  cache.put(dayKey, newDayCount.toString(), 6 * 60 * 60);    // up to 6 hours
  
  console.log(`Rate limit counters (cache) incremented: ${newHourCount}/${MAX_MESSAGES_PER_HOUR} per hour, ${newDayCount}/${MAX_MESSAGES_PER_DAY} per day`);
}

/**
 * Reset rate limit counters - useful after fixing bugs or for testing
 * Run this function manually from Apps Script to clear the counters
 */
function resetRateLimit() {
  // Clear current cache-based counters
  const now = new Date();
  const cache = CacheService.getScriptCache();
  const hourKey = `rate_hour_${now.getFullYear()}_${now.getMonth()}_${now.getDate()}_${now.getHours()}`;
  const dayKey = `rate_day_${now.getFullYear()}_${now.getMonth()}_${now.getDate()}`;
  
  cache.remove(hourKey);
  cache.remove(dayKey);

  // Clean up any legacy PropertiesService-based rate limit keys
  const props = PropertiesService.getScriptProperties();
  const allProps = props.getProperties();
  for (const key in allProps) {
    if (key.startsWith('msg_count_hour_') || key.startsWith('msg_count_day_')) {
      props.deleteProperty(key);
      console.log(`Deleted legacy rate limit key: ${key}`);
    }
  }
  
  console.log('Rate limit counters have been reset (cache cleared, legacy properties removed)');
  return 'Rate limit counters reset successfully (cache + legacy properties)';
}

/**
 * Mask token for safe logging - shows only first 4 and last 4 characters
 * @param {string} token - The token to mask
 * @returns {string} - Masked token (e.g., "1234...xyz" or "***MASKED***" if too short)
 */
function maskToken(token) {
  if (!token || typeof token !== 'string') {
    return 'null';
  }
  
  if (token.length <= 8) {
    return '***MASKED***';
  }
  
  const firstPart = token.substring(0, 4);
  const lastPart = token.substring(token.length - 4);
  return `${firstPart}...${lastPart}`;
}

/**
 * Sanitize API response data to remove any token information
 * @param {Object} responseData - The API response object
 * @returns {Object} - Sanitized response object
 */
function sanitizeApiResponse(responseData) {
  if (!responseData || typeof responseData !== 'object') {
    return responseData;
  }
  
  // Create a deep copy to avoid modifying the original
  const sanitized = JSON.parse(JSON.stringify(responseData));
  
  // Recursively sanitize the object
  function sanitizeValue(obj) {
    if (obj === null || typeof obj !== 'object') {
      return obj;
    }
    
    if (Array.isArray(obj)) {
      return obj.map(item => sanitizeValue(item));
    }
  
    const sanitizedObj = {};
    for (const key in obj) {
      const value = obj[key];
      const lowerKey = key.toLowerCase();
      
      // Mask token-related fields
      if (lowerKey.includes('token') || lowerKey.includes('api_key') || lowerKey.includes('secret')) {
        sanitizedObj[key] = maskToken(String(value));
      } else if (typeof value === 'string' && value.includes(TOKEN)) {
        // Mask token if it appears anywhere in a string
        sanitizedObj[key] = value.replace(new RegExp(TOKEN, 'g'), maskToken(TOKEN));
      } else if (typeof value === 'object') {
        sanitizedObj[key] = sanitizeValue(value);
      } else {
        sanitizedObj[key] = value;
      }
    }
    
    return sanitizedObj;
  }
  
  return sanitizeValue(sanitized);
}

/**
 * Sanitize error messages to remove token information
 * @param {string} errorMessage - The error message to sanitize
 * @returns {string} - Sanitized error message
 */
function sanitizeErrorMessage(errorMessage) {
  if (!errorMessage || typeof errorMessage !== 'string') {
    return errorMessage;
  }
  
  // Replace token in error messages
  if (TOKEN && errorMessage.includes(TOKEN)) {
    return errorMessage.replace(new RegExp(TOKEN, 'g'), maskToken(TOKEN));
  }
  
  // Also check for token patterns in URLs
  const urlPattern = /https?:\/\/api\.telegram\.org\/bot[^\/\s]+/gi;
  return errorMessage.replace(urlPattern, (match) => {
    const tokenMatch = match.match(/\/bot([^\/]+)/);
    if (tokenMatch && tokenMatch[1]) {
      return match.replace(tokenMatch[1], maskToken(tokenMatch[1]));
    }
    return match;
  });
}

/**
 * Check if message contains at least one required keyword
 * @param {string} text - The message text to check
 * @returns {Object} - Object with hasRequired boolean and foundKeywords array
 */
function hasRequiredKeywords(text) {
  if (!text || typeof text !== 'string') {
    return {
      hasRequired: false,
      foundKeywords: []
    };
  }
  
  // Convert message to lowercase for case-insensitive matching
  const lowerText = text.toLowerCase();
  const foundKeywords = [];
  
  // Check each required keyword
  for (const keyword of REQUIRED_KEYWORDS) {
    // Check if keyword appears in the message (supports partial matches)
    if (lowerText.includes(keyword.toLowerCase())) {
      foundKeywords.push(keyword);
    }
  }
  
  return {
    hasRequired: foundKeywords.length > 0,
    foundKeywords: foundKeywords
  };
}

/**
 * Extract all URLs from a message text
 * @param {string} text - The message text to scan
 * @returns {Array<string>} - Array of found URLs
 */
function extractUrls(text) {
  if (!text || typeof text !== 'string') return [];
  
  // Regex pattern to match URLs (http, https, www, and common TLDs)
  const urlPattern = /(https?:\/\/[^\s]+|www\.[^\s]+|[a-zA-Z0-9-]+\.[a-zA-Z]{2,}[^\s]*)/gi;
  const matches = text.match(urlPattern) || [];
  
  // Clean and normalize URLs
  return matches.map(url => {
    // Remove trailing punctuation that might not be part of URL
    url = url.replace(/[.,;:!?]+$/, '');
    // Add protocol if missing
    if (!url.startsWith('http://') && !url.startsWith('https://')) {
      url = 'https://' + url;
    }
    return url.toLowerCase();
  });
}

/**
 * Check if message has excessive capitalization (spam indicator)
 * @param {string} text - The message text to check
 * @returns {boolean} - True if excessive caps detected
 */
function hasExcessiveCaps(text) {
  if (!text || typeof text !== 'string') return false;
  
  // Remove URLs and whitespace for accurate calculation
  const cleanText = text.replace(/https?:\/\/[^\s]+/gi, '').replace(/\s+/g, '');
  if (cleanText.length === 0) return false;
  
  // Count uppercase letters
  const capsCount = (cleanText.match(/[A-Z]/g) || []).length;
  const capsRatio = capsCount / cleanText.length;
  
  return capsRatio >= EXCESSIVE_CAPS_THRESHOLD;
}

/**
 * Check if message has excessive punctuation (spam indicator)
 * @param {string} text - The message text to check
 * @returns {boolean} - True if excessive punctuation detected
 */
function hasExcessivePunctuation(text) {
  if (!text || typeof text !== 'string') return false;
  
  // Check for consecutive punctuation marks (!!!, ???, etc.)
  const consecutivePunctPattern = /[!?.]{3,}/g;
  if (consecutivePunctPattern.test(text)) return true;
  
  // Check for excessive punctuation overall
  const punctCount = (text.match(/[!?.]/g) || []).length;
  const punctRatio = punctCount / text.length;
  
  // Flag if more than 10% of characters are punctuation
  return punctRatio > 0.1;
}

/**
 * Check if message contains suspicious patterns
 * @param {string} text - The message text to check
 * @returns {Object} - Object with hasSuspiciousPatterns boolean and patterns array
 */
function hasSuspiciousPatterns(text) {
  if (!text || typeof text !== 'string') {
    return { hasSuspiciousPatterns: false, patterns: [] };
  }
  
  const lowerText = text.toLowerCase();
  const patterns = [];
  
  // Check for suspicious keywords
  for (const keyword of SUSPICIOUS_KEYWORDS) {
    if (lowerText.includes(keyword)) {
      patterns.push(`Suspicious keyword: "${keyword}"`);
    }
  }
  
  // Check for excessive numbers (potential phone numbers, etc.)
  const numberPattern = /\d{10,}/g;
  if (numberPattern.test(text)) {
    patterns.push('Excessive numbers detected');
  }
  
  // Check for suspicious character repetition (e.g., "aaaaaa", "!!!!!!")
  const repeatPattern = /(.)\1{4,}/g;
  if (repeatPattern.test(text)) {
    patterns.push('Suspicious character repetition');
  }
  
  return {
    hasSuspiciousPatterns: patterns.length > 0,
    patterns: patterns
  };
}

/**
 * Extract domain from a URL
 * @param {string} url - The URL to parse
 * @returns {string|null} - The domain name or null if invalid
 */
function extractDomain(url) {
  try {
    // Remove protocol and path, extract domain
    let domain = url.replace(/^https?:\/\//, '').replace(/^www\./, '');
    domain = domain.split('/')[0].split('?')[0].split('#')[0];
    
    // Remove port if present
    domain = domain.split(':')[0];
    
    return domain.toLowerCase();
  } catch (e) {
    return null;
  }
}

/**
 * Detect if message content was recently sent (duplicate detection)
 * Checks the Last Sent column for messages sent in the last 24 hours
 * @param {string} message - The message content to check
 * @param {Array} sheetData - All rows from the spreadsheet
 * @param {number} currentRowIndex - Current row index (0-based, excluding header)
 * @returns {Object} - Duplicate detection result with isDuplicate and reason
 */
function detectDuplicateContent(message, sheetData, currentRowIndex) {
  if (!message || !sheetData || currentRowIndex < 0) {
    return { isDuplicate: false, reason: '' };
  }
  
  const now = new Date();
  const twentyFourHoursAgo = new Date(now.getTime() - 24 * 60 * 60 * 1000);
  
  // Normalize message for comparison (remove extra whitespace, lowercase)
  const normalizedMessage = message.trim().toLowerCase().replace(/\s+/g, ' ');
  
  // Check other rows for duplicate content
  for (let i = 1; i < sheetData.length; i++) {
    // Skip current row
    if (i === currentRowIndex + 1) continue;
    
    const row = sheetData[i];
    const otherMessage = row[COL_MESSAGE];
    const lastSent = row[COL_LAST_SENT];
    
    if (!otherMessage) continue;
    
    // Normalize other message
    const normalizedOther = otherMessage.trim().toLowerCase().replace(/\s+/g, ' ');
    
    // Check if messages are identical
    if (normalizedMessage === normalizedOther && lastSent) {
      try {
        const lastSentDate = new Date(lastSent);
        // Only flag as duplicate if sent within last 24 hours
        if (lastSentDate >= twentyFourHoursAgo && !lastSent.toString().startsWith('ERROR:') && !lastSent.toString().startsWith('BLOCKED:')) {
          return {
            isDuplicate: true,
            reason: `Duplicate message detected: Same content was sent at ${lastSentDate.toISOString()}`
          };
        }
      } catch (e) {
        // Invalid date, skip
        continue;
      }
    }
  }
  
  return { isDuplicate: false, reason: '' };
}

/**
 * Validate message content for suspicious URLs and spam patterns
 * @param {string} message - The message to validate
 * @returns {Object} - Validation result with isValid, reason, and foundUrls
 */
function validateMessage(message) {
  if (!message || typeof message !== 'string') {
    return {
      isValid: false,
      reason: 'Empty or invalid message',
      foundUrls: [],
      blockedDomains: [],
      validationDetails: []
    };
  }
  
  const validationDetails = [];
  
  // Check message length
  if (message.length > MAX_MESSAGE_LENGTH) {
    return {
      isValid: false,
      reason: `Message too long: ${message.length} characters (limit: ${MAX_MESSAGE_LENGTH})`,
      foundUrls: [],
      blockedDomains: [],
      validationDetails: [`Length check failed: ${message.length} > ${MAX_MESSAGE_LENGTH}`]
    };
  }
  
  // Check for required keywords (if validation is enabled)
  if (REQUIRE_KEYWORD_VALIDATION) {
    const keywordCheck = hasRequiredKeywords(message);
    if (!keywordCheck.hasRequired) {
      validationDetails.push(`Missing required keywords`);
      return {
        isValid: false,
        reason: `Message does not contain any required keywords. Expected keywords include: ${REQUIRED_KEYWORDS.slice(0, 5).join(', ')}${REQUIRED_KEYWORDS.length > 5 ? '...' : ''}`,
        foundUrls: [],
        blockedDomains: [],
        validationDetails: validationDetails,
        missingKeywords: true
      };
    } else {
      validationDetails.push(`Found required keywords: ${keywordCheck.foundKeywords.join(', ')}`);
    }
  }
  
  // Extract URLs
  const urls = extractUrls(message);
  
  // Check URL count
  if (urls.length > MAX_URLS_PER_MESSAGE) {
    validationDetails.push(`Too many URLs: ${urls.length} (limit: ${MAX_URLS_PER_MESSAGE})`);
    return {
      isValid: false,
      reason: `Too many URLs detected: ${urls.length} (limit: ${MAX_URLS_PER_MESSAGE})`,
      foundUrls: urls,
      blockedDomains: [],
      validationDetails: validationDetails
    };
  }
  
  // Check for suspicious URL count (stricter check)
  if (urls.length > MAX_SUSPICIOUS_URLS) {
    validationDetails.push(`Suspicious number of URLs: ${urls.length}`);
  }
  
  // If no URLs, still check for other spam patterns
  if (urls.length === 0) {
    // Check for excessive caps
    if (hasExcessiveCaps(message)) {
      validationDetails.push('Excessive capitalization detected');
      return {
        isValid: false,
        reason: 'Excessive capitalization detected (potential spam)',
        foundUrls: [],
        blockedDomains: [],
        validationDetails: validationDetails
      };
    }
    
    // Check for excessive punctuation
    if (hasExcessivePunctuation(message)) {
      validationDetails.push('Excessive punctuation detected');
      return {
        isValid: false,
        reason: 'Excessive punctuation detected (potential spam)',
        foundUrls: [],
        blockedDomains: [],
        validationDetails: validationDetails
      };
    }
    
    // Check for suspicious patterns
    const patternCheck = hasSuspiciousPatterns(message);
    if (patternCheck.hasSuspiciousPatterns) {
      validationDetails.push(...patternCheck.patterns);
      return {
        isValid: false,
        reason: `Suspicious patterns detected: ${patternCheck.patterns.join('; ')}`,
        foundUrls: [],
        blockedDomains: [],
        validationDetails: validationDetails
      };
    }
    
    return {
      isValid: true,
      reason: 'No URLs found, content validated',
      foundUrls: [],
      blockedDomains: [],
      validationDetails: []
    };
  }
  
  // Validate URLs
  const domains = urls.map(extractDomain).filter(d => d !== null);
  const blockedDomains = [];
  
  // Check against blacklist
  for (const domain of domains) {
    for (const blockedDomain of BLOCKED_DOMAINS) {
      if (domain === blockedDomain || domain.endsWith('.' + blockedDomain)) {
        blockedDomains.push(domain);
      }
    }
  }
  
  if (blockedDomains.length > 0) {
    validationDetails.push(`Blocked domains: ${blockedDomains.join(', ')}`);
    return {
      isValid: false,
      reason: `Blocked domain(s) detected: ${blockedDomains.join(', ')}`,
      foundUrls: urls,
      blockedDomains: blockedDomains,
      validationDetails: validationDetails
    };
  }
  
  // Check against whitelist if strict mode is enabled
  if (STRICT_MODE && ALLOWED_DOMAINS.length > 0) {
    const allowedDomains = [];
    for (const domain of domains) {
      let isAllowed = false;
      for (const allowedDomain of ALLOWED_DOMAINS) {
        if (domain === allowedDomain || domain.endsWith('.' + allowedDomain)) {
          isAllowed = true;
          allowedDomains.push(domain);
          break;
        }
      }
      if (!isAllowed) {
        validationDetails.push(`Domain not in whitelist: ${domain}`);
        return {
          isValid: false,
          reason: `Domain not in whitelist: ${domain}`,
          foundUrls: urls,
          blockedDomains: [],
          allowedDomains: allowedDomains,
          validationDetails: validationDetails
        };
      }
    }
  }
  
  // Additional spam checks for messages with URLs
  // Check for excessive caps
  if (hasExcessiveCaps(message)) {
    validationDetails.push('Excessive capitalization detected');
  }
  
  // Check for excessive punctuation
  if (hasExcessivePunctuation(message)) {
    validationDetails.push('Excessive punctuation detected');
  }
  
  // Check for suspicious patterns
  const patternCheck = hasSuspiciousPatterns(message);
  if (patternCheck.hasSuspiciousPatterns) {
    validationDetails.push(...patternCheck.patterns);
    // If multiple suspicious patterns, block the message
    if (patternCheck.patterns.length >= 2) {
      return {
        isValid: false,
        reason: `Multiple suspicious patterns detected: ${patternCheck.patterns.join('; ')}`,
        foundUrls: urls,
        blockedDomains: [],
        validationDetails: validationDetails
      };
    }
  }
  
  return {
    isValid: true,
    reason: `URLs validated: ${domains.join(', ')}`,
    foundUrls: urls,
    blockedDomains: [],
    allowedDomains: domains,
    validationDetails: validationDetails
  };
}

/**
 * Send the helper poll to the AG chat.
 * This can be run directly from the Apps Script editor for testing.
 * @returns {boolean} - True if the poll was sent successfully
 */
function sendHelperPoll() {
  if (!TOKEN) {
    console.error('Cannot send helper poll: TELEGRAM_TOKEN not configured');
    return false;
  }
  if (!AG_CHAT_ID) {
    console.error('Cannot send helper poll: AG_CHAT_ID not configured in Script Properties');
    return false;
  }

  const pollUrl = `https://api.telegram.org/bot${TOKEN}/sendPoll`;
  const pollPayload = {
    chat_id: AG_CHAT_ID,
    question: 'Wer kann beim nächsten Tontag helfen?',
    options: ['Aufbau', 'Abbau', 'Ich kann nicht'],
    allows_multiple_answers: true,
    is_anonymous: false
  };

  try {
    const pollResponse = UrlFetchApp.fetch(pollUrl, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(pollPayload)
    });
    const pollResponseData = JSON.parse(pollResponse.getContentText());
    const sanitizedPollResponse = sanitizeApiResponse(pollResponseData);
    console.log('Telegram Poll API response:', sanitizedPollResponse);

    if (pollResponseData.ok) {
      console.log('Helper poll sent successfully');
      return true;
    } else {
      console.error('Failed to send helper poll:', sanitizeErrorMessage(pollResponseData.description || 'Unknown Telegram Poll API error'));
      return false;
    }
  } catch (e) {
    console.error('Error sending helper poll:', sanitizeErrorMessage(e.message));
    return false;
  }
}

/**
 * Notify admin via debug chat when messages are blocked
 * @param {number} rowNumber - The row number (1-based, spreadsheet row)
 * @param {string} message - Full message content that was blocked
 * @param {string} blockReason - Reason for blocking
 * @param {Object} validationResult - Validation result object
 * @param {string} additionalInfo - Any additional information to include
 * @returns {boolean} - True if notification sent successfully
 */
function notifyAdmin(rowNumber, message, blockReason, validationResult = null, additionalInfo = '') {
  if (!DEBUG_CHAT_ID || !TOKEN) {
    console.warn('Cannot send admin notification: DEBUG_CHAT_ID or TOKEN not configured');
    return false;
  }
  
  try {
    const timestamp = new Date().toISOString();
    let notification = `🚫 BLOCKED MESSAGE ALERT\n\n`;
    notification += `Timestamp: ${timestamp}\n`;
    notification += `Row Number: ${rowNumber}\n`;
    notification += `Block Reason: ${blockReason}\n\n`;
    
    if (validationResult) {
      if (validationResult.foundUrls && validationResult.foundUrls.length > 0) {
        notification += `Found URLs: ${validationResult.foundUrls.length}\n`;
        notification += `URLs: ${validationResult.foundUrls.join(', ')}\n\n`;
      }
      if (validationResult.blockedDomains && validationResult.blockedDomains.length > 0) {
        notification += `Blocked Domains: ${validationResult.blockedDomains.join(', ')}\n\n`;
      }
      if (validationResult.validationDetails && validationResult.validationDetails.length > 0) {
        notification += `Validation Details:\n`;
        validationResult.validationDetails.forEach(detail => {
          notification += `- ${detail}\n`;
        });
        notification += `\n`;
      }
    }
    
    if (additionalInfo) {
      notification += `Additional Info: ${additionalInfo}\n\n`;
    }
    
    notification += `Full Message Content:\n"${message}"\n\n`;
    notification += `Message Length: ${message.length} characters`;
    
    const url = `https://api.telegram.org/bot${TOKEN}/sendMessage`;
    const payload = {
      chat_id: DEBUG_CHAT_ID,
      text: notification
    };
    
    const response = UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(payload)
    });
    
    const responseData = JSON.parse(response.getContentText());
    if (responseData.ok) {
      console.log('Admin notification sent successfully');
      return true;
    } else {
      console.error('Failed to send admin notification:', responseData.description);
      return false;
    }
  } catch (e) {
    console.error('Error sending admin notification:', sanitizeErrorMessage(e.message));
    return false;
  }
}

/**
 * Enhanced logging function for message sends
 * Logs full message content, row number, repeat pattern, timestamp, and validation status
 * @param {number} rowNumber - The row number (1-based, spreadsheet row)
 * @param {string} message - Full message content
 * @param {string} repeatPattern - Repeat pattern
 * @param {Date} timestamp - When the message is being sent
 * @param {Object} validationResult - Result from validateMessage()
 * @param {boolean} wasBlocked - Whether message was blocked
 * @param {string} blockReason - Reason for blocking (if applicable)
 */
function logMessageSend(rowNumber, message, repeatPattern, timestamp, validationResult, wasBlocked = false, blockReason = '') {
  // Handle undefined or invalid timestamp
  let timestampStr;
  if (timestamp && timestamp instanceof Date) {
    timestampStr = timestamp.toISOString();
  } else if (timestamp) {
    // Try to convert to Date if it's not already
    try {
      timestampStr = new Date(timestamp).toISOString();
    } catch (e) {
      timestampStr = new Date().toISOString(); // Fallback to current time
    }
  } else {
    timestampStr = new Date().toISOString(); // Fallback to current time if undefined
  }
  
  const logEntry = {
    timestamp: timestampStr,
    rowNumber: rowNumber,
    repeatPattern: repeatPattern || 'One-time',
    messageLength: message ? message.length : 0,
    fullMessage: message || '[EMPTY]',
    validationStatus: validationResult && validationResult.isValid ? 'PASSED' : 'FAILED',
    validationReason: validationResult && validationResult.reason ? validationResult.reason : 'Unknown',
    foundUrls: validationResult && validationResult.foundUrls ? validationResult.foundUrls : [],
    blockedDomains: validationResult && validationResult.blockedDomains ? validationResult.blockedDomains : [],
    wasBlocked: wasBlocked,
    blockReason: blockReason
  };
  
  // Log to console with full details
  console.log('=== MESSAGE SEND LOG ===');
  console.log(`Timestamp: ${logEntry.timestamp}`);
  console.log(`Row Number: ${logEntry.rowNumber}`);
  console.log(`Repeat Pattern: ${logEntry.repeatPattern}`);
  console.log(`Message Length: ${logEntry.messageLength} characters`);
  console.log(`Full Message Content: ${logEntry.fullMessage}`);
  console.log(`Validation Status: ${logEntry.validationStatus}`);
  console.log(`Validation Reason: ${logEntry.validationReason}`);
  if (logEntry.foundUrls.length > 0) {
    console.log(`Found URLs: ${logEntry.foundUrls.join(', ')}`);
  }
  if (logEntry.blockedDomains.length > 0) {
    console.log(`Blocked Domains: ${logEntry.blockedDomains.join(', ')}`);
  }
  if (logEntry.wasBlocked) {
    console.log(`⚠️ MESSAGE BLOCKED: ${logEntry.blockReason}`);
  } else {
    console.log('✅ Message approved for sending');
  }
  console.log('========================');
  
  return logEntry;
}

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
  
  // Test actual values being used (token is masked for security)
  console.log('Using TOKEN:', maskToken(TOKEN));
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
        // Sanitize error description before logging
        const sanitizedDescription = sanitizeErrorMessage(responseData.description || 'Unknown error');
        console.error('❌ Telegram API error:', sanitizedDescription);
        return false;
      }
      
    } catch (e) {
      // Sanitize error message before logging
      const sanitizedError = sanitizeErrorMessage(e.message);
      console.error('❌ Network/Script error:', sanitizedError);
      return false;
    }
  }
  
  // Send initial system status message
  const systemMessage = `🧪 DEBUG TEST STARTED - ${new Date().toLocaleString()}\n\n` +
                        `✅ Script is running\n` +
                        `🔑 TOKEN: ${TOKEN ? 'Found' : 'Missing'}\n` +
                        `💬 CHAT_ID: ${CHAT_ID}\n` +
                        `📊 Properties working: ${PropertiesService.getScriptProperties().getProperty('TELEGRAM_TOKEN') ? 'Yes' : 'No'}\n\n` +
                        `📋 Now analyzing next ${DEBUG_TEST_MAX_MESSAGES} upcoming enabled messages...`;
  
  console.log('Sending system debug message...');
  sendDebugMessage(systemMessage, true);
  
  // Brief delay before processing spreadsheet to avoid rate limiting
  Utilities.sleep(1000);
  
  // Process spreadsheet and analyze only the next N upcoming enabled messages
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const data = sheet.getDataRange().getValues();
    
    console.log('Processing spreadsheet with', data.length - 1, 'rows');
    
    // Collect all enabled rows with datetime, then sort by schedule time (soonest first)
    const enabledRows = [];
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (!row[COL_DATETIME] || !row[COL_MESSAGE]) continue;
      const enabled = row[COL_ENABLED];
      const isEnabled = enabled === true ||
                       enabled === '✅' ||
                       (typeof enabled === 'string' && (
                         enabled.toLowerCase().trim() === 'yes' ||
                         enabled.toLowerCase().trim() === 'true' ||
                         enabled.includes('✅')
                       ));
      if (isEnabled) {
        enabledRows.push({ i: i, row: row, start: new Date(row[COL_DATETIME]) });
      }
    }
    enabledRows.sort((a, b) => a.start.getTime() - b.start.getTime());
    const rowsToTest = enabledRows.slice(0, DEBUG_TEST_MAX_MESSAGES);
    
    console.log('Total enabled:', enabledRows.length, '- Testing upcoming', rowsToTest.length, 'messages');
    
    const totalEnabled = enabledRows.length;
    let sentCount = 0;
    let blockedCount = 0;
    let validationFailedCount = 0;
    
    // Process only the upcoming N messages
    for (const { i, row } of rowsToTest) {
      const start = new Date(row[COL_DATETIME]);
      const msg = row[COL_MESSAGE];
      const repeatRaw = row[COL_REPEAT] || '';
      
      if (msg) {
        
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
        
        // Validate message content
        const validationResult = validateMessage(msg);
        
        // Track validation failures
        if (!validationResult.isValid) {
          validationFailedCount++;
        }
        
        // Track blocked messages
        if (shouldSend && !validationResult.isValid) {
          blockedCount++;
        }
        
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
        
        // Build validation status message
        let validationStatus = '';
        if (validationResult.foundUrls.length > 0) {
          validationStatus = `\nVALIDATION CHECK:\n`;
          validationStatus += `Status: ${validationResult.isValid ? '✅ PASSED' : '🚫 BLOCKED'}\n`;
          validationStatus += `Reason: ${validationResult.reason}\n`;
          validationStatus += `Found URLs: ${validationResult.foundUrls.length}\n`;
          if (validationResult.foundUrls.length > 0) {
            validationStatus += `URLs: ${validationResult.foundUrls.join(', ')}\n`;
          }
          if (validationResult.blockedDomains.length > 0) {
            validationStatus += `⚠️ Blocked Domains: ${validationResult.blockedDomains.join(', ')}\n`;
          }
        } else {
          validationStatus = `\nVALIDATION CHECK:\n✅ No URLs found - No validation needed\n`;
        }
        
        // Determine final send status considering validation
        const willActuallySend = shouldSend && validationResult.isValid;
        const sendStatus = willActuallySend ? '✅ WILL SEND' : 
                          (!shouldSend ? '❌ WILL NOT SEND (Schedule)' : 
                          '🚫 WILL BE BLOCKED (Validation)');
        
        const debugMsg = `ROW ${i + 1} - ${repeatRaw || 'One-time'}\n\n` +
                        `SCHEDULE CHECK:\n` +
                        `${scheduleAnalysis}\n\n` +
                        `DUPLICATE CHECK:\n` +
                        `${duplicateCheck === 'PASS' ? 'No recent send detected' : duplicateCheck}\n` +
                        `${validationStatus}` +
                        `RESULT: ${sendStatus}\n\n` +
                        `FULL MESSAGE CONTENT:\n"${msg}"\n\n` +
                        `Message Length: ${msg.length} characters`;
        
        console.log(`Sending analysis for row ${i + 1}`);
        console.log(`Validation result:`, validationResult);
        
        if (sendDebugMessage(debugMsg)) {
          sentCount++;
        }
        
        // Small delay between messages to avoid rate limiting
        Utilities.sleep(800);
      }
    }
    
    // Send summary message
    const summaryMessage = `📊 DEBUG TEST COMPLETE\n\n` +
                          `📋 Total rows in sheet: ${data.length - 1}\n` +
                          `✅ Enabled messages (total): ${totalEnabled}\n` +
                          `📌 Upcoming messages tested: ${rowsToTest.length}\n` +
                          `📤 Analysis messages sent: ${sentCount}\n` +
                          `🚫 Messages that would be blocked: ${blockedCount}\n` +
                          `⚠️ Validation failures: ${validationFailedCount}\n` +
                          `🕐 Test completed at: ${new Date().toLocaleString()}`;
    
    sendDebugMessage(summaryMessage, true);
    
    console.log('=== DEBUG TEST COMPLETED ===');
    return `Debug test completed. Sent ${sentCount} of ${rowsToTest.length} upcoming messages (${totalEnabled} enabled total).`;
    
  } catch (e) {
    // Sanitize error message before logging and sending
    const sanitizedError = sanitizeErrorMessage(e.message);
    const errorMessage = `❌ DEBUG TEST ERROR\n\nError reading spreadsheet: ${sanitizedError}`;
    sendDebugMessage(errorMessage, true);
    console.error('Debug test error:', sanitizedError);
    return `Error: ${sanitizedError}`;
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
        // Check for duplicate content
        const duplicateCheck = detectDuplicateContent(msg, data, i);
        if (duplicateCheck.isDuplicate) {
          const blockReason = `BLOCKED: ${duplicateCheck.reason}`;
          console.error(`🚫 BLOCKED MESSAGE - Row ${i + 1}: ${blockReason}`);
          console.error(`Full message content: ${msg}`);
          
          // Store block reason in Last Sent column
          const blockMsg = blockReason;
          sheet.getRange(i + 1, COL_LAST_SENT + 1).setValue(blockMsg);
          
          // Notify admin about duplicate
          notifyAdmin(i + 1, msg, blockReason, null, duplicateCheck.reason);
          
          continue; // Skip sending this message
        }
        
        // Validate message content before sending
        const validationResult = validateMessage(msg);
        
        // Log message details before sending (or blocking)
        logMessageSend(
          i + 1,
          msg,
          repeatRaw,
          now,
          validationResult,
          !validationResult.isValid,
          !validationResult.isValid ? validationResult.reason : ''
        );
        
        // Block message if validation failed
        if (!validationResult.isValid) {
          const blockReason = `BLOCKED: ${validationResult.reason}`;
          console.error(`🚫 BLOCKED MESSAGE - Row ${i + 1}: ${blockReason}`);
          console.error(`Full message content: ${msg}`);
          if (validationResult.blockedDomains && validationResult.blockedDomains.length > 0) {
            console.error(`Blocked domains: ${validationResult.blockedDomains.join(', ')}`);
          }
          
          // Store block reason in Last Sent column
          const blockMsg = `BLOCKED: ${validationResult.reason}`;
          sheet.getRange(i + 1, COL_LAST_SENT + 1).setValue(blockMsg);
          
          // Notify admin about blocked message
          notifyAdmin(i + 1, msg, blockReason, validationResult, '');
          
          continue; // Skip sending this message
        }
        
        // Check rate limit before sending
        const rateLimitCheck = checkRateLimit();
        if (!rateLimitCheck.isAllowed) {
          console.error(`🚫 RATE LIMIT - Row ${i + 1}: ${rateLimitCheck.reason}`);
          const blockMsg = `RATE_LIMITED: ${rateLimitCheck.reason}`;
          sheet.getRange(i + 1, COL_LAST_SENT + 1).setValue(blockMsg);
          
          // Notify admin about rate limit with the message that couldn't be sent
          notifyAdmin(i + 1, msg, `RATE_LIMITED: ${rateLimitCheck.reason}`, null, `Hour: ${rateLimitCheck.hourCount}/${MAX_MESSAGES_PER_HOUR}, Day: ${rateLimitCheck.dayCount}/${MAX_MESSAGES_PER_DAY}`);
          
          continue; // Skip this message due to rate limit
        }
        
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
          // Sanitize response before logging to mask any token information
          const sanitizedResponse = sanitizeApiResponse(responseData);
          console.log('Telegram API response:', sanitizedResponse);
          
          if (responseData.ok) {
            // Update Last Sent column with current timestamp
            sheet.getRange(i + 1, COL_LAST_SENT + 1).setValue(now);
            console.log(`Successfully sent message for row ${i + 1}`);
            
            // Increment rate limit counter only after successful send
            incrementRateLimit();

            // If this message contained URLs, also send helper poll to AG chat
            if (validationResult && validationResult.foundUrls && validationResult.foundUrls.length > 0) {
              const pollSent = sendHelperPoll();
              if (!pollSent) {
                console.warn('Helper poll could not be sent (see logs above).');
              }
            }
          } else {
            throw new Error(responseData.description || 'Unknown Telegram API error');
          }
          
        } catch (e) {
          // Sanitize error message before logging to mask any token information
          const sanitizedError = sanitizeErrorMessage(e.message);
          console.error(`Error sending message for row ${i + 1}:`, sanitizedError);
          // Store sanitized error message in Last Sent column (prefixed to avoid date parsing issues)
          const errorMsg = `ERROR: ${sanitizedError}`;
          sheet.getRange(i + 1, COL_LAST_SENT + 1).setValue(errorMsg);
          
          // Notify admin about the error with full message content
          notifyAdmin(i + 1, msg, `ERROR: ${sanitizedError}`, null, `Row ${i + 1} failed to send`);
        }
      }
    }
    
    console.log('Script completed at:', new Date());
    
  } catch (err) {
    // Log any unexpected errors that occur during script execution
    // Sanitize error messages to mask any token information
    const sanitizedMessage = sanitizeErrorMessage(err.message);
    const sanitizedStack = sanitizeErrorMessage(err.stack || '');
    console.error("Script error:", sanitizedMessage);
    console.error("Stack trace:", sanitizedStack);
    // Note: Error is logged but not re-thrown to prevent trigger failures
  }
}