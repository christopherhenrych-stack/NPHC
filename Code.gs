/**
 * ============================================================================
 * EXECUTIVE DASHBOARD - COMPLETE GOOGLE APPS SCRIPT
 * ============================================================================
 * VP of Value Brands - Executive Command Center
 * 
 * FEATURES:
 * - Auto-creates sheet tabs if they don't exist
 * - Syncs executive's Google Calendar to Meetings sheet
 * - Serves dashboard with real-time data
 * - Creates agenda documents
 * - Supports recurring meeting management
 * 
 * Version: 2.0 (Complete with Calendar Sync)
 * Last Updated: February 1, 2026
 * ============================================================================
 */

// ============================================================================
// CONFIGURATION - UPDATE THESE
// ============================================================================

const CONFIG = {
  // Main Sheet IDs
  DELIVERABLES_SHEET_ID: '1ZmC-04S_OdhuoJs-XIBiOujj5rtsJTZNOiH26_0LTgk',
  
  // If meetings are in same workbook, use same ID:
  MEETINGS_SHEET_ID: '1ZmC-04S_OdhuoJs-XIBiOujj5rtsJTZNOiH26_0LTgk',
  
  // Tab/Sheet Names (will be auto-created if missing)
  MEETINGS_TAB_NAME: 'Meetings',
  DELIVERABLES_TAB_NAME: 'Document Repository',
  ARCHIVE_TAB_NAME: 'Meeting Archive', // NEW - for completed meetings
  
  // Calendar Settings for Auto-Sync
  EXECUTIVE_CALENDAR_ID: 'michael.sarcone@verizon.com',
  SYNC_DAYS_AHEAD: 90,      // Pull meetings 90 days forward
  SYNC_DAYS_BEHIND: 0,      // Do not pull past meetings
  
  // Row ranges for Document Repository sheet
  // 
  // ⭐ IMPORTANT: These ranges are FLEXIBLE!
  // The code uses DYNAMIC detection and stops at the first empty row.
  // You can add/remove rows freely - just update START_ROW if sections move.
  // END_ROW values are safety limits, not hard requirements.
  // 
  // ✅ You CAN change section names in the sheet without breaking code
  // ✅ You CAN add more rows without updating END_ROW
  // ✅ You CAN move sections - just update START_ROW values below
  
  EXEC1_START_ROW: 3,          // Mike's/Sarcone's Deliverables (Skipping Title)
  EXEC1_END_ROW: 13,
  
  EXEC2_START_ROW: 27,         // DK's Governance Deliverables (Skipping Title)
  EXEC2_END_ROW: 37,
  
  ADMIN_START_ROW: 47,         // Admin & Settings Links (Skipping Title)
  ADMIN_END_ROW: 52,
  
  KPI_START_ROW: 57,           // KPIs & Performance (Skipping Title)
  KPI_END_ROW: 68,
  
  EXEC_UPDATES_START_ROW: 75,  // Executive Updates Archive (Skipping Title)
  EXEC_UPDATES_END_ROW: 120,
  
  DOCUMENTS_START_ROW: 161,    // Document Repository Section (Skipping Title)
  DOCUMENTS_END_ROW: 202,
  
  COMPLETED_START_ROW: 216,    // Completed Items (Skipping Title)
  COMPLETED_END_ROW: 219,
  
  // Alert Banner Settings
  ALERT_SHEET_NAME: 'Alerts',
  ALERT_START_ROW: 2,
  
  // Column indices for Meetings sheet (0-indexed)
  MEETING_COLUMNS: {
    COMPLETE: 0,         // A
    APPROVED: 1,         // B
    HOT_TOPIC: 2,        // C
    MEETING_NAME: 3,     // D
    DATE: 4,             // E
    TIME: 5,             // F
    DURATION: 6,         // G
    FREQUENCY: 7,        // H
    CATEGORY: 8,         // I
    DESCRIPTION: 9,      // J
    ATTENDEES: 10,       // K
    DELIVERABLE_LINK: 11,// L
    PREP_REQUIRED: 12,   // M
    COLOR_ID: 13,        // N (Event color or hot topic level)
    NOTES: 14,           // O
    EVENT_ID: 15         // P
  }
};

// ============================================================================
// MAIN FUNCTIONS
// ============================================================================

/**
 * Setup function - Initializes sheets (run this first!)
 */
function setup() {
  ensureSheetsExist();
  Logger.log('✅ Setup complete!');
  Logger.log('📊 Sheets verified/created');
  Logger.log('');
  Logger.log('🚀 Next steps:');
  Logger.log('1. Refresh your Google Sheet to see the menu');
  Logger.log('2. Use menu: 📊 Executive Dashboard > 🔄 Sync Calendar Now');
  Logger.log('3. Deploy as Web App (Deploy > New Deployment)');
}

/**
 * Create custom menu - called automatically when sheet opens
 */
function onOpen() {
  createCustomMenu();
}

/**
 * Create custom menu in Google Sheets
 */
function createCustomMenu() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('📊 Executive Dashboard')
    .addItem('🔄 Sync Calendar Now', 'syncCalendarToMeetingsSheet')
    .addItem('📅 Setup Daily Auto-Sync', 'setupDailySyncTrigger')
    .addItem('🗑️ Remove Auto-Sync', 'removeSyncTriggers')
    .addSeparator()
    .addItem('📦 Archive Completed Meetings', 'archiveCompletedMeetings')
    .addSeparator()
    .addItem('🏗️ Initialize Sheets', 'ensureSheetsExist')
    .addItem('🧪 Test Data Retrieval', 'testDataRetrieval')
    .addToUi();
}

/**
 * Ensure all required sheets exist, create if missing
 */
function ensureSheetsExist() {
  const ss = SpreadsheetApp.openById(CONFIG.DELIVERABLES_SHEET_ID);
  
  let sheetsCreated = [];
  
  // Create Meetings sheet if missing
  let meetingsSheet = ss.getSheetByName(CONFIG.MEETINGS_TAB_NAME);
  if (!meetingsSheet) {
    Logger.log('📝 Creating Meetings sheet...');
    meetingsSheet = ss.insertSheet(CONFIG.MEETINGS_TAB_NAME);
    initializeMeetingsSheet(meetingsSheet);
    sheetsCreated.push(CONFIG.MEETINGS_TAB_NAME);
  } else {
    Logger.log('✅ Meetings sheet already exists');
  }
  
  // Create Deliverables sheet if missing
  let deliverablesSheet = ss.getSheetByName(CONFIG.DELIVERABLES_TAB_NAME);
  if (!deliverablesSheet) {
    Logger.log('📝 Creating Team Updates sheet...');
    deliverablesSheet = ss.insertSheet(CONFIG.DELIVERABLES_TAB_NAME);
    initializeDeliverablesSheet(deliverablesSheet);
    sheetsCreated.push(CONFIG.DELIVERABLES_TAB_NAME);
  } else {
    Logger.log('✅ Team Updates sheet already exists');
  }
  
  // Show results
  if (sheetsCreated.length > 0) {
    Logger.log('✅ Created sheets: ' + sheetsCreated.join(', '));
    // Try to show UI alert if possible, but don't fail if not
    try {
      SpreadsheetApp.getUi().alert('✅ Sheets Created!\n\nNew sheets:\n• ' + sheetsCreated.join('\n• '));
    } catch (e) {
      Logger.log('(Running from script editor - UI alerts disabled)');
    }
  } else {
    Logger.log('✅ All required sheets already exist');
    try {
      SpreadsheetApp.getUi().alert('✅ All sheets verified!\n\nExisting sheets:\n• ' + CONFIG.MEETINGS_TAB_NAME + '\n• ' + CONFIG.DELIVERABLES_TAB_NAME);
    } catch (e) {
      Logger.log('(Running from script editor - UI alerts disabled)');
    }
  }
}

/**
 * Initialize Meetings sheet with headers
 */
function initializeMeetingsSheet(sheet) {
  const headers = [
    'COMPLETE', 'APPROVED', 'HOT TOPIC', 'MEETING NAME', 'DATE', 'TIME',
    'DURATION', 'FREQUENCY', 'CATEGORY', 'DESCRIPTION', 'ATTENDEES',
    'DELIVERABLE LINK', 'PREP REQUIRED', 'COLOR ID', 'NOTES', 'EVENT ID'
  ];
  
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length)
    .setFontWeight('bold')
    .setBackground('#1e293b')
    .setFontColor('#ffffff')
    .setHorizontalAlignment('center');
  
  // Set column widths
  sheet.setColumnWidth(1, 80);  // Complete
  sheet.setColumnWidth(2, 80);  // Approved
  sheet.setColumnWidth(3, 90);  // Hot Topic
  sheet.setColumnWidth(4, 250); // Meeting Name
  sheet.setColumnWidth(5, 100); // Date
  sheet.setColumnWidth(6, 80);  // Time
  sheet.setColumnWidth(7, 90);  // Duration
  sheet.setColumnWidth(8, 100); // Frequency
  sheet.setColumnWidth(9, 120); // Category
  sheet.setColumnWidth(10, 300);// Description
  sheet.setColumnWidth(11, 200);// Attendees
  sheet.setColumnWidth(12, 200);// Deliverable Link
  sheet.setColumnWidth(13, 110);// Prep Required
  sheet.setColumnWidth(14, 50); // Empty
  sheet.setColumnWidth(15, 250);// Notes
  sheet.setColumnWidth(16, 200);// Event ID
  
  // Format checkboxes
  sheet.getRange('A2:C1000').insertCheckboxes();
  sheet.getRange('M2:M1000').insertCheckboxes();
  
  // Freeze header row
  sheet.setFrozenRows(1);
  
  Logger.log('✅ Meetings sheet initialized with headers and formatting');
}

/**
 * Initialize Deliverables sheet with section headers
 */
function initializeDeliverablesSheet(sheet) {
  const sections = [
    { row: 1, title: 'Mike\'s/Sarcone\'s Upcoming Deliverables & Hot Topics' },
    { row: 21, title: 'DK\'s Governance Upcoming Deliverables & Hot Topics' },
    { row: 36, title: 'Admin & Settings - Quick Links' },
    { row: 47, title: 'KPIs & Performance Dashboards' },
    { row: 61, title: 'Executive Updates - Completed Hot Topics Archive' },
    { row: 161, title: 'Document Repository' },
    { row: 213, title: 'DK\'s Completed Governance Items' }
  ];
  
  const columnHeaders = ['Description', 'ETA', 'Owner', 'Link to Materials', 'Days Left', '', 'Comment'];
  
  sections.forEach(section => {
    // Section title
    sheet.getRange(section.row, 1, 1, 7).merge()
      .setValue(section.title)
      .setFontWeight('bold')
      .setFontSize(12)
      .setBackground('#f3f4f6')
      .setHorizontalAlignment('left');
    
    // Column headers
    sheet.getRange(section.row + 1, 1, 1, columnHeaders.length)
      .setValues([columnHeaders])
      .setFontWeight('bold')
      .setBackground('#e5e7eb')
      .setHorizontalAlignment('center');
  });
  
  // Set column widths
  sheet.setColumnWidth(1, 300); // Description
  sheet.setColumnWidth(2, 100); // ETA
  sheet.setColumnWidth(3, 120); // Owner
  sheet.setColumnWidth(4, 250); // Link
  sheet.setColumnWidth(5, 100); // Days Left
  sheet.setColumnWidth(6, 50);  // Empty
  sheet.setColumnWidth(7, 250); // Comment
  
  Logger.log('✅ Deliverables sheet initialized with section structure');
}

// ============================================================================
// CALENDAR SYNC FUNCTIONS
// ============================================================================

/**
 * Sync executive's calendar to Meetings sheet
 */
function syncCalendarToMeetingsSheet() {
  Logger.log('🔄 Starting calendar sync...');
  
  const ss = SpreadsheetApp.openById(CONFIG.MEETINGS_SHEET_ID);
  const sheet = ss.getSheetByName(CONFIG.MEETINGS_TAB_NAME);
  
  if (!sheet) {
    Logger.log('❌ Meetings sheet not found! Run ensureSheetsExist() first.');
    try {
      SpreadsheetApp.getUi().alert('❌ Meetings sheet not found!\n\nRun "Initialize Sheets" first.');
    } catch (e) {
      Logger.log('(Run ensureSheetsExist() function to create the sheet)');
    }
    return;
  }
  
  // Calculate date range
  const now = new Date();
  const startDate = new Date(now);
  startDate.setHours(0, 0, 0, 0);
  const endDate = new Date(now.getTime() + (CONFIG.SYNC_DAYS_AHEAD * 24 * 60 * 60 * 1000));
  
  Logger.log(`📅 Syncing from ${startDate.toDateString()} to ${endDate.toDateString()}`);
  
  // Get calendar
  let calendar;
  try {
    if (CONFIG.EXECUTIVE_CALENDAR_ID === 'primary') {
      calendar = CalendarApp.getDefaultCalendar();
    } else {
      calendar = CalendarApp.getCalendarById(CONFIG.EXECUTIVE_CALENDAR_ID);
    }
  } catch (e) {
    Logger.log('❌ Error accessing calendar: ' + e);
    try {
      SpreadsheetApp.getUi().alert('❌ Cannot access calendar!\n\n' + e);
    } catch (uiError) {
      Logger.log('(Run from Google Sheets to see UI alerts)');
    }
    return;
  }
  
  // Get events
  const events = calendar.getEvents(startDate, endDate);
  Logger.log(`📊 Found ${events.length} calendar events`);
  
  // Get existing data to preserve manual edits
  const existingData = getExistingMeetingData(sheet);
  
  // Process events
  const newRows = [];
  for (const event of events) {
    const title = event.getTitle();
    
    // Skip all-day events
    if (event.isAllDayEvent()) {
      continue;
    }
    
    // Skip noisy meetings (DNS, lunch, 1:1, no attendees, no virtual link)
    if (shouldFilterMeeting(event, title)) {
      continue;
    }
    
    // Get event details
    const startTime = event.getStartTime();
    const endTime = event.getEndTime();
    const duration = Math.round((endTime - startTime) / (1000 * 60)); // minutes
    
    // Skip past meetings (including earlier today)
    if (startTime < now) continue;
    
    // Skip 0-minute meetings
    if (duration <= 0) continue;
    
    // Get description
    const description = event.getDescription() || '';
    const location = event.getLocation() || '';
    
    const eventId = event.getId();
    const existingRow = existingData[eventId];
    
    // Get attendees (limit to first 5)
    const guests = event.getGuestList();
    const attendees = guests.slice(0, 5)
      .map(g => g.getName() || g.getEmail().split('@')[0])
      .join(', ');
    
    // Clean description - remove HTML formatting
    const cleanDescription = cleanHtmlDescription(description);
    
    // Determine meeting type and frequency
    const meetingType = determineMeetingType(title);
    const frequency = determineFrequency(event, title);
    
    // Determine hot topic based on color and keywords
    const hotTopicMeta = getHotTopicMeta(event, title, cleanDescription);
    const isAutoHotTopic = hotTopicMeta.isHotTopic === true;
    const hotTopicValue = isAutoHotTopic || (existingRow ? existingRow[CONFIG.MEETING_COLUMNS.HOT_TOPIC] === true : false);
    
    // Build row - preserve checkbox states if meeting exists
    const row = [
      existingRow ? existingRow[CONFIG.MEETING_COLUMNS.COMPLETE] : false,        // A: Complete
      existingRow ? existingRow[CONFIG.MEETING_COLUMNS.APPROVED] : false,        // B: Approved
      hotTopicValue,                                                             // C: Hot Topic
      title,                                                                       // D: Meeting Name
      startTime,                                                                   // E: Date
      Utilities.formatDate(startTime, Session.getScriptTimeZone(), 'HH:mm'),     // F: Time
      duration + ' min',                                                           // G: Duration
      frequency,                                                                   // H: Frequency
      meetingType,                                                                 // I: Category
      cleanDescription,                                                            // J: Description
      attendees,                                                                   // K: Attendees
      existingRow ? existingRow[CONFIG.MEETING_COLUMNS.DELIVERABLE_LINK] : '',   // L: Deliverable Link
      existingRow ? existingRow[CONFIG.MEETING_COLUMNS.PREP_REQUIRED] : false,   // M: Prep Required
      hotTopicMeta.metaValue || '',                                               // N: Color ID / Hot Topic Level
      existingRow ? existingRow[CONFIG.MEETING_COLUMNS.NOTES] : '',              // O: Notes
      eventId                                                                      // P: Event ID
    ];
    
    newRows.push(row);
  }
  
  // Clear existing data (except header)
  if (sheet.getLastRow() > 1) {
    sheet.getRange(2, 1, sheet.getLastRow() - 1, 16).clearContent();
  }
  
  // Write new data
  if (newRows.length > 0) {
    sheet.getRange(2, 1, newRows.length, 16).setValues(newRows);
  }

  // Ensure checkbox formatting for A, B, C, and M columns
  const lastRow = Math.max(sheet.getLastRow(), 2);
  sheet.getRange(2, 1, lastRow - 1, 1).insertCheckboxes(); // A
  sheet.getRange(2, 2, lastRow - 1, 1).insertCheckboxes(); // B
  sheet.getRange(2, 3, lastRow - 1, 1).insertCheckboxes(); // C
  sheet.getRange(2, 13, lastRow - 1, 1).insertCheckboxes(); // M
  
  Logger.log(`✅ Sync complete! ${newRows.length} meetings synced`);
  try {
    SpreadsheetApp.getUi().alert(`✅ Calendar Sync Complete!\n\n${newRows.length} meetings synced\nFrom: ${startDate.toDateString()}\nTo: ${endDate.toDateString()}`);
  } catch (e) {
    Logger.log('(Sync complete - view results in the sheet)');
  }
}

/**
 * Get existing meeting data to preserve manual edits
 */
function getExistingMeetingData(sheet) {
  const data = {};
  if (sheet.getLastRow() < 2) return data;
  
  const range = sheet.getRange(2, 1, sheet.getLastRow() - 1, 16);
  const values = range.getValues();
  
  values.forEach(row => {
    const eventId = row[CONFIG.MEETING_COLUMNS.EVENT_ID];
    if (eventId) {
      data[eventId] = row;
    }
  });
  
  return data;
}

/**
 * Determine meeting type from title
 */
function determineMeetingType(title) {
  const titleLower = title.toLowerCase();
  
  if (titleLower.includes('1:1') || titleLower.includes('one on one')) return '1:1';
  if (titleLower.includes('staff') || titleLower.includes('team')) return 'Staff Meeting';
  if (titleLower.includes('board') || titleLower.includes('executive')) return 'Executive';
  if (titleLower.includes('standup') || titleLower.includes('daily')) return 'Standup';
  if (titleLower.includes('review') || titleLower.includes('retrospective')) return 'Review';
  if (titleLower.includes('planning')) return 'Planning';
  if (titleLower.includes('sync')) return 'Sync';
  if (titleLower.includes('deep work') || titleLower.includes('focus')) return 'Deep Work';
  
  return 'Meeting';
}

/**
 * Determine frequency from event
 */
function determineFrequency(event, title) {
  const titleLower = title.toLowerCase();
  
  if (titleLower.includes('daily')) return 'Daily';
  if (titleLower.includes('weekly')) return 'Weekly';
  if (titleLower.includes('biweekly') || titleLower.includes('bi-weekly')) return 'Bi-weekly';
  if (titleLower.includes('monthly')) return 'Monthly';
  if (titleLower.includes('quarterly')) return 'Quarterly';
  
  // Check if it's a recurring event
  // Note: EventRecurrence is not directly accessible in Apps Script
  // This is a simplified check
  return 'One-time';
}

/**
 * Check if event is non-meeting (OOO, holiday, DNS, lunch, 1:1, etc.)
 */
function isNonMeetingEvent(title) {
  const titleLower = title.toLowerCase();
  const nonMeetingKeywords = [
    'ooo', 'out of office', 'vacation', 'pto', 'holiday',
    'off', 'personal', 'dentist', 'doctor', 'appointment',
    'dns', 'do not schedule', 'lunch', 'break', 'blocked', 'hold',
    '1:1', '1-1', 'one on one', 'one-on-one', '1 on 1',
    'focus time', 'no meetings'
  ];
  
  return nonMeetingKeywords.some(keyword => titleLower.includes(keyword));
}

/**
 * Check if meeting should be filtered out (no attendees or no virtual link)
 */
function shouldFilterMeeting(event, title) {
  // First check title-based filters (DNS, Lunch, OOO, etc.)
  if (isNonMeetingEvent(title)) {
    return true;
  }
  
  // Rule for Attendees and Virtual Links has been removed.
  // All other events will now be synced to the sheet.
  return false;
}

function getEventColorIdSafe(event) {
  try {
    const color = event.getColor();
    if (color === CalendarApp.EventColor.ORANGE) return '6';
    if (color === CalendarApp.EventColor.RED) return '11';
    if (color === CalendarApp.EventColor.YELLOW) return '5';
    if (color === CalendarApp.EventColor.GREEN) return '10';
    if (color === CalendarApp.EventColor.BLUE) return '9';
    if (color === CalendarApp.EventColor.GRAY) return '8';
    if (color === CalendarApp.EventColor.PURPLE) return '7';
    if (color) return String(color);
  } catch (e) {
    Logger.log('Color lookup failed: ' + e);
  }
  return '';
}

function getHotTopicLevelFromText(title, description) {
  const text = ((title || '') + ' ' + (description || '')).toLowerCase();
  if (text.includes('urgent') || text.includes('critical')) return 'urgent';
  if (text.includes('must attend') || text.includes('must-attend')) return 'must-attend';
  return '';
}

function getHotTopicMeta(event, title, description) {
  const colorId = getEventColorIdSafe(event);
  let level = '';
  if (colorId === '11') level = 'urgent';
  if (colorId === '6') level = 'must-attend';
  
  if (!level) {
    level = getHotTopicLevelFromText(title, description);
  }
  
  const isHotTopic = level === 'urgent' || level === 'must-attend';
  const metaValue = colorId ? colorId : (level || '');
  return { isHotTopic: isHotTopic, level: level, metaValue: metaValue };
}

/**
 * Clean HTML formatting from description
 */
function cleanHtmlDescription(html) {
  if (!html) return '';
  
  // Remove HTML tags but keep the text content
  let cleaned = html
    // Remove <br>, <br/>, <br /> tags with newlines
    .replace(/<br\s*\/?>/gi, '\n')
    // Remove HTML tags
    .replace(/<[^>]+>/g, '')
    // Decode common HTML entities
    .replace(/&nbsp;/g, ' ')
    .replace(/&amp;/g, '&')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&#39;/g, "'")
    // Remove multiple spaces
    .replace(/\s+/g, ' ')
    // Remove multiple newlines
    .replace(/\n\s*\n/g, '\n')
    // Trim
    .trim();
  
  // Limit length to avoid cell overflow (max 500 chars)
  if (cleaned.length > 500) {
    cleaned = cleaned.substring(0, 497) + '...';
  }
  
  return cleaned;
}

/**
 * Setup daily automatic calendar sync
 */
function setupDailySyncTrigger() {
  // Remove existing triggers first
  removeSyncTriggers();
  
  // Create new daily trigger at 6 AM
  ScriptApp.newTrigger('syncCalendarToMeetingsSheet')
    .timeBased()
    .atHour(6)
    .everyDays(1)
    .create();
  
  Logger.log('✅ Daily sync trigger created (6 AM daily)');
  try {
    SpreadsheetApp.getUi().alert('✅ Automatic Daily Sync Enabled!\n\nCalendar will sync every day at 6 AM.\n\nTo disable, run "Remove Auto-Sync" from the menu.');
  } catch (e) {
    Logger.log('(Trigger created successfully - check Triggers in Apps Script)');
  }
}

/**
 * Remove all sync triggers
 */
function removeSyncTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  let removed = 0;
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'syncCalendarToMeetingsSheet') {
      ScriptApp.deleteTrigger(trigger);
      removed++;
    }
  });
  Logger.log(`✅ Removed ${removed} sync trigger(s)`);
  try {
    if (removed > 0) {
      SpreadsheetApp.getUi().alert(`✅ Auto-sync disabled!\n\n${removed} trigger(s) removed.`);
    } else {
      SpreadsheetApp.getUi().alert('ℹ️ No auto-sync triggers found.');
    }
  } catch (e) {
    Logger.log('(Triggers removed - view in Apps Script > Triggers)');
  }
}

// ============================================================================
// WEB APP FUNCTIONS
// ============================================================================

/**
 * Serves the HTML dashboard
 */
function doGet(e) {
  const template = HtmlService.createTemplateFromFile('index');
  return template.evaluate()
    .setTitle('Value Governance Dashboard')
    .setFaviconUrl('https://ssl.gstatic.com/docs/spreadsheets/favicon3.ico')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
}

/**
 * Include HTML files
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ============================================================================
// DATA RETRIEVAL FUNCTIONS
// ============================================================================

/**
 * Get all dashboard data in one call
 */
function getAllDashboardData() {
  try {
    Logger.log('Starting getAllDashboardData...');
    
    const result = {
      exec1Deliverables: [],
      exec2Deliverables: [],
      meetings: [],
      recurringMeetings: [],
      adminLinks: [],
      kpiLinks: [],
      executiveUpdates: [],
      documentRepository: [],
      documentRepoLinks: [],
      completedItems: []
    };
    
    // Get each section with individual error handling
    try {
      result.exec1Deliverables = getExec1Deliverables();
      Logger.log('✅ Exec1: ' + result.exec1Deliverables.length + ' items');
    } catch (e) {
      Logger.log('❌ Error getting Exec1: ' + e);
    }
    
    try {
      result.exec2Deliverables = getExec2Deliverables();
      Logger.log('✅ Exec2: ' + result.exec2Deliverables.length + ' items');
    } catch (e) {
      Logger.log('❌ Error getting Exec2: ' + e);
    }
    
    try {
      result.meetings = getMeetings();
      Logger.log('✅ Meetings: ' + result.meetings.length + ' items');
    } catch (e) {
      Logger.log('❌ Error getting Meetings: ' + e);
      Logger.log('Full error: ' + e.stack);
    }
    
    try {
      result.recurringMeetings = getRecurringMeetings();
      Logger.log('✅ Recurring: ' + result.recurringMeetings.length + ' items');
    } catch (e) {
      Logger.log('❌ Error getting Recurring: ' + e);
    }
    
    try {
      result.adminLinks = getAdminLinks();
      Logger.log('✅ Admin: ' + result.adminLinks.length + ' items');
    } catch (e) {
      Logger.log('❌ Error getting Admin: ' + e);
    }
    
    try {
      result.kpiLinks = getKPILinks();
      Logger.log('✅ KPI: ' + result.kpiLinks.length + ' items');
    } catch (e) {
      Logger.log('❌ Error getting KPI: ' + e);
    }
    
    try {
      result.executiveUpdates = getExecutiveUpdates();
      Logger.log('✅ Exec Updates: ' + result.executiveUpdates.length + ' items');
    } catch (e) {
      Logger.log('❌ Error getting Exec Updates: ' + e);
    }
    
    try {
      result.documentRepository = getDocumentRepository();
      Logger.log('✅ Documents: ' + result.documentRepository.length + ' items');
    } catch (e) {
      Logger.log('❌ Error getting Documents: ' + e);
    }
    
    try {
      result.documentRepoLinks = getDocumentRepoLinks();
      Logger.log('✅ Doc Repo Links: ' + result.documentRepoLinks.length + ' items');
    } catch (e) {
      Logger.log('❌ Error getting Doc Repo Links: ' + e);
    }
    
    try {
      result.completedItems = getCompletedItems();
      Logger.log('✅ Completed: ' + result.completedItems.length + ' items');
    } catch (e) {
      Logger.log('❌ Error getting Completed: ' + e);
    }
    
    Logger.log('Returning data...');
    
    // *** CRITICAL FIX: Force JSON serialization to strip out problematic Date objects ***
    const jsonSafeData = JSON.parse(JSON.stringify(result));
    return jsonSafeData;
    
  } catch (error) {
    Logger.log('❌❌❌ CRITICAL ERROR in getAllDashboardData: ' + error);
    Logger.log('Stack: ' + error.stack);
    return { 
      error: error.toString(),
      exec1Deliverables: [],
      exec2Deliverables: [],
      meetings: [],
      recurringMeetings: [],
      adminLinks: [],
      kpiLinks: [],
      executiveUpdates: [],
      documentRepository: [],
      documentRepoLinks: [],
      completedItems: []
    };
  }
}

/**
 * Simple test function to verify backend connection
 */
function testConnection() {
  return {
    status: 'success',
    message: 'Backend is working!',
    timestamp: new Date().toISOString(),
    sheetId: CONFIG.DELIVERABLES_SHEET_ID
  };
}

/**
 * Test if getAllDashboardData can be serialized
 */
function testSerializability() {
  try {
    const data = getAllDashboardData();
    const jsonString = JSON.stringify(data);
    
    return {
      success: true,
      dataSize: jsonString.length,
      itemCounts: {
        exec1: data.exec1Deliverables.length,
        exec2: data.exec2Deliverables.length,
        meetings: data.meetings.length,
        recurring: data.recurringMeetings.length,
        admin: data.adminLinks.length,
        kpi: data.kpiLinks.length,
        updates: data.executiveUpdates.length,
        documents: data.documentRepository.length,
        completed: data.completedItems.length
      }
    };
  } catch (e) {
    return {
      success: false,
      error: e.toString(),
      stack: e.stack
    };
  }
}

/**
 * Archive completed meetings to separate sheet
 */
function archiveCompletedMeetings() {
  const sheet = SpreadsheetApp.openById(CONFIG.MEETINGS_SHEET_ID).getSheetByName('Meetings');
  let archive = SpreadsheetApp.openById(CONFIG.MEETINGS_SHEET_ID).getSheetByName('Meeting Archive');
  
  if (!archive) {
    archive = SpreadsheetApp.openById(CONFIG.MEETINGS_SHEET_ID).insertSheet('Meeting Archive');
    archive.getRange(1, 1, 1, 16).setValues([sheet.getRange(1, 1, 1, 16).getValues()[0]]);
    archive.getRange(1, 1, 1, 16).setFontWeight('bold').setBackground('#1e293b').setFontColor('#ffffff');
  }
  
  const data = sheet.getDataRange().getValues();
  const toArchive = [];
  const toDelete = [];
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === true) { // COMPLETE checkbox
      toArchive.push(data[i]);
      toDelete.push(i + 1);
    }
  }
  
  if (toArchive.length > 0) {
    const lastRow = archive.getLastRow();
    archive.getRange(lastRow + 1, 1, toArchive.length, 16).setValues(toArchive);
    
    for (let i = toDelete.length - 1; i >= 0; i--) {
      sheet.deleteRow(toDelete[i]);
    }
    
    Logger.log('✅ Archived ' + toArchive.length + ' completed meetings');
    try {
      SpreadsheetApp.getUi().alert('✅ Archived ' + toArchive.length + ' completed meetings');
    } catch (e) {
      Logger.log('Alert not available in this context');
    }
  } else {
    Logger.log('No completed meetings to archive');
    try {
      SpreadsheetApp.getUi().alert('No completed meetings to archive');
    } catch (e) {
      Logger.log('Alert not available in this context');
    }
  }
}

/**
 * Get Exec 1 Deliverables (DYNAMIC - stops at first empty row)
 */
function getExec1Deliverables() {
  try {
    const sheet = getSheet(CONFIG.DELIVERABLES_SHEET_ID, CONFIG.DELIVERABLES_TAB_NAME);
    const startRow = CONFIG.EXEC1_START_ROW;
    const maxRows = CONFIG.EXEC1_END_ROW - startRow + 1;
    
    const range = sheet.getRange(startRow, 1, maxRows, 7);
    const values = range.getValues();
    const richTextValues = range.getRichTextValues();
    
    const deliverables = [];
    const today = new Date();
    
    for (let i = 0; i < values.length; i++) {
      const row = values[i];
      if (!row[0] || row[0].toString().trim() === '') break;
      
      const etaCell = row[1];
      const etaStr = etaCell ? String(etaCell).trim() : '';
      const isWeeklyEta = etaStr && etaStr.toLowerCase().includes('weekly');
      const eta = !isWeeklyEta && etaCell ? new Date(etaCell) : null;
      const daysRemaining = isWeeklyEta ? null : (eta ? Math.ceil((eta - today) / (1000 * 60 * 60 * 24)) : null);
      
      // Extract URL from Column D hyperlink
      let linkedMaterials = '';
      const richText = richTextValues[i][3];
      if (richText) {
        linkedMaterials = richText.getLinkUrl() || '';
      }
      if (!linkedMaterials && row[3]) {
        linkedMaterials = String(row[3]);
      }
      
      deliverables.push({
        id: 'exec1-' + i,
        description: row[0] || '',
        eta: isWeeklyEta ? 'Weekly' : (eta ? Utilities.formatDate(eta, Session.getScriptTimeZone(), 'MMM dd') : ''),
        etaFull: isWeeklyEta ? '' : (eta ? Utilities.formatDate(eta, Session.getScriptTimeZone(), 'yyyy-MM-dd') : ''),
        owner: row[2] || '',
        linkedMaterials: linkedMaterials,
        daysRemaining: daysRemaining,
        daysRemainingText: isWeeklyEta ? 'N/A' : (daysRemaining !== null ? Math.abs(daysRemaining) + 'd' : ''),
        comment: row[6] || '',
        isHotTopic: !isWeeklyEta && daysRemaining !== null && daysRemaining <= 7,
        isOverdue: daysRemaining !== null && daysRemaining < 0,
        urgency: getUrgencyLevel(daysRemaining)
      });
    }
    
    return deliverables;
  } catch (error) {
    Logger.log('Error getting Exec 1 deliverables: ' + error);
    return [];
  }
}

/**
 * Get Exec 2 Deliverables (DYNAMIC)
 */
function getExec2Deliverables() {
  try {
    const sheet = getSheet(CONFIG.DELIVERABLES_SHEET_ID, CONFIG.DELIVERABLES_TAB_NAME);
    const startRow = CONFIG.EXEC2_START_ROW;
    const maxRows = CONFIG.EXEC2_END_ROW - startRow + 1;
    
    const range = sheet.getRange(startRow, 1, maxRows, 7);
    const values = range.getValues();
    const richTextValues = range.getRichTextValues();
    
    const deliverables = [];
    const today = new Date();
    
    for (let i = 0; i < values.length; i++) {
      const row = values[i];
      if (!row[0] || row[0].toString().trim() === '') break;
      
      const etaCell = row[1];
      const etaStr = etaCell ? String(etaCell).trim() : '';
      const isWeeklyEta = etaStr && etaStr.toLowerCase().includes('weekly');
      const eta = !isWeeklyEta && etaCell ? new Date(etaCell) : null;
      const daysRemaining = isWeeklyEta ? null : (eta ? Math.ceil((eta - today) / (1000 * 60 * 60 * 24)) : null);
      
      // Extract URL from Column D hyperlink
      let linkedMaterials = '';
      const richText = richTextValues[i][3];
      if (richText) {
        linkedMaterials = richText.getLinkUrl() || '';
      }
      if (!linkedMaterials && row[3]) {
        linkedMaterials = String(row[3]);
      }
      
      deliverables.push({
        id: 'exec2-' + i,
        description: row[0] || '',
        eta: isWeeklyEta ? 'Weekly' : (eta ? Utilities.formatDate(eta, Session.getScriptTimeZone(), 'MMM dd') : ''),
        etaFull: isWeeklyEta ? '' : (eta ? Utilities.formatDate(eta, Session.getScriptTimeZone(), 'yyyy-MM-dd') : ''),
        owner: row[2] || '',
        linkedMaterials: linkedMaterials,
        daysRemaining: daysRemaining,
        daysRemainingText: isWeeklyEta ? 'N/A' : (daysRemaining !== null ? Math.abs(daysRemaining) + 'd' : ''),
        comment: row[6] || '',
        isHotTopic: !isWeeklyEta && daysRemaining !== null && daysRemaining <= 7,
        isOverdue: daysRemaining !== null && daysRemaining < 0,
        urgency: getUrgencyLevel(daysRemaining)
      });
    }
    
    return deliverables;
  } catch (error) {
    Logger.log('Error getting Exec 2 deliverables: ' + error);
    return [];
  }
}

/**
 * Get meetings for 2-week schedule view
 */
function normalizeMeetingTime(rawTime) {
  if (!rawTime) {
    return { time12: '', time24: '' };
  }
  
  if (rawTime instanceof Date) {
    return {
      time12: Utilities.formatDate(rawTime, Session.getScriptTimeZone(), 'h:mm a'),
      time24: Utilities.formatDate(rawTime, Session.getScriptTimeZone(), 'HH:mm')
    };
  }
  
  const str = String(rawTime).trim();
  const ampmMatch = str.match(/^(\d{1,2})(?::(\d{2}))?\s*(AM|PM)$/i);
  if (ampmMatch) {
    let hour = parseInt(ampmMatch[1], 10);
    const minute = (ampmMatch[2] || '00').padStart(2, '0');
    const ampm = ampmMatch[3].toUpperCase();
    if (ampm === 'PM' && hour !== 12) hour += 12;
    if (ampm === 'AM' && hour === 12) hour = 0;
    const time24 = String(hour).padStart(2, '0') + ':' + minute;
    const hour12 = hour % 12 || 12;
    const time12 = hour12 + ':' + minute + ' ' + ampm;
    return { time12: time12, time24: time24 };
  }
  
  const militaryMatch = str.match(/^(\d{1,2}):(\d{2})/);
  if (militaryMatch) {
    let hour = parseInt(militaryMatch[1], 10);
    const minute = militaryMatch[2];
    const ampm = hour >= 12 ? 'PM' : 'AM';
    const hour12 = hour % 12 || 12;
    const time12 = hour12 + ':' + minute + ' ' + ampm;
    const time24 = String(hour).padStart(2, '0') + ':' + minute;
    return { time12: time12, time24: time24 };
  }
  
  return { time12: str, time24: '' };
}

/**
 * Get meetings for 2-week schedule view
 */
function getMeetings() {
  try {
    const sheet = getSheet(CONFIG.MEETINGS_SHEET_ID, CONFIG.MEETINGS_TAB_NAME);
    if (sheet.getLastRow() < 2) return [];
    
    const dataRange = sheet.getDataRange();
    const values = dataRange.getValues();
    
    const today = new Date();
    today.setHours(0, 0, 0, 0);
    // Show meetings from start of previous month to end of next year for full calendar support
    const rangeStart = new Date(today.getFullYear(), today.getMonth() - 1, 1);
    const rangeEnd = new Date(today.getFullYear() + 1, 11, 31);
    
    const meetings = [];
    
    for (let i = 1; i < values.length; i++) {
      const row = values[i];
      if (!row[CONFIG.MEETING_COLUMNS.MEETING_NAME]) continue;
      
      // ONLY show approved meetings
      if (row[CONFIG.MEETING_COLUMNS.APPROVED] !== true) continue;
      
      const meetingDate = new Date(row[CONFIG.MEETING_COLUMNS.DATE]);
      
      // Skip weekends (0 = Sunday, 6 = Saturday)
      const dayOfWeek = meetingDate.getDay();
      if (dayOfWeek === 0 || dayOfWeek === 6) continue;

      if (meetingDate >= rangeStart && meetingDate <= rangeEnd) {
        const normalizedTime = normalizeMeetingTime(row[CONFIG.MEETING_COLUMNS.TIME]);
        const colorMeta = row[CONFIG.MEETING_COLUMNS.COLOR_ID] ? String(row[CONFIG.MEETING_COLUMNS.COLOR_ID]).trim() : '';
        const keywordLevel = getHotTopicLevelFromText(row[CONFIG.MEETING_COLUMNS.MEETING_NAME], row[CONFIG.MEETING_COLUMNS.DESCRIPTION]);
        let hotTopicLevel = '';
        if (colorMeta === '11') hotTopicLevel = 'urgent';
        if (colorMeta === '6') hotTopicLevel = 'must-attend';
        if (!hotTopicLevel && (colorMeta === 'urgent' || colorMeta === 'must-attend')) hotTopicLevel = colorMeta;
        if (!hotTopicLevel && keywordLevel) hotTopicLevel = keywordLevel;
        
        const isHotTopic = row[CONFIG.MEETING_COLUMNS.HOT_TOPIC] === true || !!hotTopicLevel;
        
        meetings.push({
          id: row[CONFIG.MEETING_COLUMNS.EVENT_ID] || 'meeting-' + i,
          title: row[CONFIG.MEETING_COLUMNS.MEETING_NAME] || '',
          date: Utilities.formatDate(meetingDate, Session.getScriptTimeZone(), 'yyyy-MM-dd'),
          dateFormatted: Utilities.formatDate(meetingDate, Session.getScriptTimeZone(), 'MMM dd'),
          time: normalizedTime.time12,
          time24: normalizedTime.time24,
          duration: String(row[CONFIG.MEETING_COLUMNS.DURATION] || ''),
          category: row[CONFIG.MEETING_COLUMNS.CATEGORY] || 'meeting',
          description: row[CONFIG.MEETING_COLUMNS.DESCRIPTION] || '',
          attendees: row[CONFIG.MEETING_COLUMNS.ATTENDEES] || '',
          isHotTopic: isHotTopic,
          hotTopicLevel: hotTopicLevel,
          colorId: colorMeta,
          prepRequired: row[CONFIG.MEETING_COLUMNS.PREP_REQUIRED] === true,
          deliverableLink: row[CONFIG.MEETING_COLUMNS.DELIVERABLE_LINK] || '',
          notes: row[CONFIG.MEETING_COLUMNS.NOTES] || '',
          dayOfWeek: meetingDate.getDay()
        });
      }
    }
    
    // Sort by date and time
    meetings.sort((a, b) => {
      const dateCompare = a.date.localeCompare(b.date);
      if (dateCompare !== 0) return dateCompare;
      
      // Compare times as strings (format: "HH:mm")
      const timeA = a.time24 || '00:00';
      const timeB = b.time24 || '00:00';
      return timeA.localeCompare(timeB);
    });
    
    return meetings;
  } catch (error) {
    Logger.log('Error getting meetings: ' + error);
    return [];
  }
}

/**
 * Get recurring meetings for agenda builder
 */
function getRecurringMeetings() {
  try {
    const sheet = getSheet(CONFIG.MEETINGS_SHEET_ID, CONFIG.MEETINGS_TAB_NAME);
    if (sheet.getLastRow() < 2) return [];
    
    const dataRange = sheet.getDataRange();
    const values = dataRange.getValues();
    
    const recurringMeetings = [];
    const seen = new Set();
    
    for (let i = 1; i < values.length; i++) {
      const row = values[i];
      const frequency = row[CONFIG.MEETING_COLUMNS.FREQUENCY] ? row[CONFIG.MEETING_COLUMNS.FREQUENCY].toString().toLowerCase() : '';
      const meetingName = row[CONFIG.MEETING_COLUMNS.MEETING_NAME];
      
      if (meetingName && frequency && 
          (frequency.includes('weekly') || frequency.includes('monthly') || 
           frequency.includes('daily') || frequency.includes('bi-weekly'))) {
        
        const key = meetingName.toString().trim();
        if (!seen.has(key)) {
          seen.add(key);
          recurringMeetings.push({
            id: row[CONFIG.MEETING_COLUMNS.EVENT_ID] || 'recurring-' + i,
            name: meetingName,
            frequency: row[CONFIG.MEETING_COLUMNS.FREQUENCY]
          });
        }
      }
    }
    
    return recurringMeetings;
  } catch (error) {
    Logger.log('Error getting recurring meetings: ' + error);
    return [];
  }
}

/**
 * Get Admin Links (DYNAMIC)
 */
function getAdminLinks() {
  try {
    const sheet = getSheet(CONFIG.DELIVERABLES_SHEET_ID, CONFIG.DELIVERABLES_TAB_NAME);
    const startRow = CONFIG.ADMIN_START_ROW;
    const maxRows = CONFIG.ADMIN_END_ROW - startRow + 1;
    
    const range = sheet.getRange(startRow, 1, maxRows, 7);
    const values = range.getValues();
    const richTextValues = range.getRichTextValues();
    
    const links = [];
    for (let i = 0; i < values.length; i++) {
      const row = values[i];
      if (!row[0] || row[0].toString().trim() === '') break;
      
      // Extract URL from Column D hyperlink
      let url = '';
      const richText = richTextValues[i][3];
      if (richText) {
        url = richText.getLinkUrl() || '';
      }
      if (!url && row[3]) {
        url = String(row[3]);
      }
      
      links.push({
        id: 'admin-' + i,
        name: String(row[0] || ''),
        url: url,
        category: 'admin'
      });
    }
    
    return links;
  } catch (error) {
    Logger.log('Error getting admin links: ' + error);
    return [];
  }
}

/**
 * Get KPI Links (DYNAMIC)
 */
function getKPILinks() {
  try {
    const sheet = getSheet(CONFIG.DELIVERABLES_SHEET_ID, CONFIG.DELIVERABLES_TAB_NAME);
    const startRow = CONFIG.KPI_START_ROW;
    const maxRows = CONFIG.KPI_END_ROW - startRow + 1;
    
    const range = sheet.getRange(startRow, 1, maxRows, 7);
    const values = range.getValues();
    const richTextValues = range.getRichTextValues();
    
    const links = [];
    for (let i = 0; i < values.length; i++) {
      const row = values[i];
      if (!row[0] || row[0].toString().trim() === '') break;
      
      // Extract URL from Column D hyperlink
      let url = '';
      const richText = richTextValues[i][3];
      if (richText) {
        url = richText.getLinkUrl() || '';
      }
      if (!url && row[3]) {
        url = String(row[3]);
      }
      
      links.push({
        id: 'kpi-' + i,
        name: String(row[0] || ''),
        url: url,
        category: 'kpi'
      });
    }
    
    return links;
  } catch (error) {
    Logger.log('Error getting KPI links: ' + error);
    return [];
  }
}

/**
 * Get Executive Updates (DYNAMIC)
 */
function getExecutiveUpdates() {
  try {
    const sheet = getSheet(CONFIG.DELIVERABLES_SHEET_ID, CONFIG.DELIVERABLES_TAB_NAME);
    const startRow = CONFIG.EXEC_UPDATES_START_ROW;
    const maxRows = CONFIG.EXEC_UPDATES_END_ROW - startRow + 1;
    
    const range = sheet.getRange(startRow, 1, maxRows, 7);
    const values = range.getValues();
    const richTextValues = range.getRichTextValues();
    
    const updates = [];
    for (let i = 0; i < values.length; i++) {
      const row = values[i];
      if (!row[0] || row[0].toString().trim() === '') break;
      
      // Extract URL from Column D hyperlink
      let url = '';
      const richText = richTextValues[i][3];
      if (richText) {
        url = richText.getLinkUrl() || '';
      }
      if (!url && row[3]) {
        url = String(row[3]);
      }
      
      updates.push({
        id: 'update-' + i,
        name: String(row[0] || ''),
        url: url,
        dateAdded: row[1] ? Utilities.formatDate(new Date(row[1]), Session.getScriptTimeZone(), 'MMM dd, yyyy') : '',
        category: 'executive-update'
      });
    }
    
    return updates;
  } catch (error) {
    Logger.log('Error getting executive updates: ' + error);
    return [];
  }
}

/**
 * Get Document Repository items (DYNAMIC) - NEW SECTION!
 */
function getDocumentRepository() {
  try {
    const sheet = getSheet(CONFIG.DELIVERABLES_SHEET_ID, CONFIG.DELIVERABLES_TAB_NAME);
    const startRow = CONFIG.DOCUMENTS_START_ROW;
    const maxRows = CONFIG.DOCUMENTS_END_ROW - startRow + 1;
    
    const range = sheet.getRange(startRow, 1, maxRows, 7);
    const values = range.getValues();
    
    const documents = [];
    for (let i = 0; i < values.length; i++) {
      const row = values[i];
      if (!row[0] || row[0].toString().trim() === '') break;
      
      // Safely handle date formatting
      let dateAdded = '';
      try {
        if (row[1] && row[1] instanceof Date) {
          dateAdded = Utilities.formatDate(row[1], Session.getScriptTimeZone(), 'MMM dd, yyyy');
        } else if (row[1]) {
          dateAdded = row[1].toString();
        }
      } catch (e) {
        Logger.log('Warning: Invalid date in Document Repository row ' + (startRow + i));
      }
      
      documents.push({
        id: 'document-' + i,
        name: String(row[0] || ''),
        url: String(row[3] || ''),
        dateAdded: dateAdded,
        owner: String(row[2] || ''),
        category: 'document'
      });
    }
    
    return documents;
  } catch (error) {
    Logger.log('Error getting document repository: ' + error);
    return [];
  }
}

/**
 * Get Completed Items (function getCompletedItems() {
  try {
    const sheet = getSheet(CONFIG.DELIVERABLES_SHEET_ID, CONFIG.DELIVERABLES_TAB_NAME);
    const startRow = CONFIG.COMPLETED_START_ROW;
    const maxRows = CONFIG.COMPLETED_END_ROW - startRow + 1;
    
    const range = sheet.getRange(startRow, 1, maxRows, 7);
    const values = range.getValues();
    
    const completed = [];
    for (let i = 0; i < values.length; i++) {
      const row = values[i];
      if (!row[0] || row[0].toString().trim() === '') break;
      
      completed.push({
        id: 'completed-' + i,
        description: row[0],
        eta: row[1] instanceof Date ? Utilities.formatDate(row[1], Session.getScriptTimeZone(), 'MMM dd') : row[1],
        owner: row[2],
        link: row[3],
        daysLeft: row[4],
        comment: row[6]
      });
    }
    return completed;
  } catch (error) {
    Logger.log('Error getting completed items: ' + error);
    return [];
  }
}

/**
 * Get active alerts for the scrolling banner
 */
function getAlerts() {
  try {
    const sheet = getSheet(CONFIG.DELIVERABLES_SHEET_ID, CONFIG.ALERT_SHEET_NAME);
    if (!sheet) return [];
    
    const data = sheet.getDataRange().getValues();
    const alerts = [];
    
    // Column A: Active (Checkbox), Column B: Alert Text
    for (let i = CONFIG.ALERT_START_ROW - 1; i < data.length; i++) {
      if (data[i][0] === true && data[i][1]) {
        alerts.push(data[i][1]);
      }
    }
    return alerts;
  } catch (error) {
    Logger.log('Error getting alerts: ' + error);
    return [];
  }
}
// ===========================================
// AGENDA CREATION FUNCTIONS
// ============================================================================

/**
 * Create agenda document
 */
function createAgenda(meetingData) {
  try {
    const { meetingName, meetingDate, topics, meetingLink } = meetingData;
    
    const doc = DocumentApp.create(`${meetingName} - Agenda - ${meetingDate}`);
    const body = doc.getBody();
    
    // Title
    const title = body.appendParagraph(meetingName);
    title.setHeading(DocumentApp.ParagraphHeading.HEADING1);
    title.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    
    // Date
    const dateP = body.appendParagraph(meetingDate);
    dateP.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    body.appendParagraph('');
    
    // Meeting link
    if (meetingLink) {
      const linkP = body.appendParagraph('Meeting Link: ');
      linkP.appendText(meetingLink).setLinkUrl(meetingLink);
      body.appendParagraph('');
    }
    
    // Agenda topics
    const topicsHeading = body.appendParagraph('Agenda Topics');
    topicsHeading.setHeading(DocumentApp.ParagraphHeading.HEADING2);
    
    topics.forEach((topic, index) => {
      if (topic && topic.trim() !== '') {
        const topicP = body.appendParagraph(`${index + 1}. ${topic}`);
        topicP.setIndentFirstLine(36);
        body.appendParagraph('   Notes:');
        body.appendParagraph('');
      }
    });
    
    // Action items
    body.appendParagraph('');
    const actionHeading = body.appendParagraph('Action Items');
    actionHeading.setHeading(DocumentApp.ParagraphHeading.HEADING2);
    body.appendParagraph('• ');
    body.appendParagraph('• ');
    body.appendParagraph('• ');
    
    // Next steps
    body.appendParagraph('');
    const nextStepsHeading = body.appendParagraph('Next Steps');
    nextStepsHeading.setHeading(DocumentApp.ParagraphHeading.HEADING2);
    body.appendParagraph('');
    
    doc.saveAndClose();
    
    return {
      success: true,
      documentUrl: doc.getUrl(),
      documentId: doc.getId(),
      documentName: doc.getName()
    };
    
  } catch (error) {
    Logger.log('Error creating agenda: ' + error);
    return {
      success: false,
      error: error.toString()
    };
  }
}

/**
 * Attach agenda document link to a meeting row
 */
function attachAgendaToMeeting(data) {
  try {
    const { meetingId, documentUrl, documentName, agendaSummary } = data;
    if (!meetingId || !documentUrl) {
      return { success: false, error: 'Missing meetingId or documentUrl' };
    }
    
    const sheet = getSheet(CONFIG.MEETINGS_SHEET_ID, CONFIG.MEETINGS_TAB_NAME);
    if (!sheet) {
      return { success: false, error: 'Meetings sheet not found' };
    }
    
    const values = sheet.getDataRange().getValues();
    let targetRow = -1;
    for (let i = 1; i < values.length; i++) {
      if (String(values[i][CONFIG.MEETING_COLUMNS.EVENT_ID] || '') === String(meetingId)) {
        targetRow = i + 1;
        break;
      }
    }
    
    if (targetRow === -1) {
      return { success: false, error: 'Meeting not found' };
    }
    
    const linkCell = sheet.getRange(targetRow, CONFIG.MEETING_COLUMNS.DELIVERABLE_LINK + 1);
    linkCell.setFormula('=HYPERLINK(\"' + documentUrl + '\", \"' + (documentName || 'Agenda') + '\")');

    if (agendaSummary && String(agendaSummary).trim()) {
      const notesCell = sheet.getRange(targetRow, CONFIG.MEETING_COLUMNS.NOTES + 1);
      const existing = notesCell.getValue() ? String(notesCell.getValue()).trim() : '';
      const agendaBlock = 'Agenda Topics:\n' + agendaSummary;
      const merged = existing ? (existing + '\n\n' + agendaBlock) : agendaBlock;
      notesCell.setValue(merged);
    }
    
    return { success: true };
  } catch (error) {
    Logger.log('Error attaching agenda: ' + error);
    return { success: false, error: error.toString() };
  }
}

/**
 * Update meeting fields from the dashboard modal
 */
function updateMeeting(meetingId, updates) {
  try {
    if (!meetingId) return { success: false, error: 'Missing meetingId' };
    
    const sheet = getSheet(CONFIG.MEETINGS_SHEET_ID, CONFIG.MEETINGS_TAB_NAME);
    if (!sheet) return { success: false, error: 'Meetings sheet not found' };
    
    const values = sheet.getDataRange().getValues();
    let targetRow = -1;
    for (let i = 1; i < values.length; i++) {
      if (String(values[i][CONFIG.MEETING_COLUMNS.EVENT_ID] || '') === String(meetingId)) {
        targetRow = i + 1;
        break;
      }
    }
    
    if (targetRow === -1) return { success: false, error: 'Meeting not found' };
    
    const baseNotes = (updates.notes || '').split('Agenda Topics:')[0].trim();
    const agendaTopics = (updates.agendaTopics || '').trim();
    const agendaBlock = agendaTopics ? ('Agenda Topics:\n' + agendaTopics) : '';
    const finalNotes = baseNotes && agendaBlock ? (baseNotes + '\n\n' + agendaBlock) : (baseNotes || agendaBlock);
    
    const notesCell = sheet.getRange(targetRow, CONFIG.MEETING_COLUMNS.NOTES + 1);
    notesCell.setValue(finalNotes);
    
    const linkCell = sheet.getRange(targetRow, CONFIG.MEETING_COLUMNS.DELIVERABLE_LINK + 1);
    if (updates.deliverableLink && String(updates.deliverableLink).trim()) {
      linkCell.setFormula('=HYPERLINK(\"' + updates.deliverableLink + '\", \"Agenda\")');
    } else {
      linkCell.setValue('');
    }
    
    return { success: true };
  } catch (error) {
    Logger.log('Error updating meeting: ' + error);
    return { success: false, error: error.toString() };
  }
}

// ============================================================================
// HELPER FUNCTIONS
// ============================================================================

/**
 * Get sheet by ID and tab name
 */
function getSheet(sheetId, tabName) {
  const ss = SpreadsheetApp.openById(sheetId);
  return ss.getSheetByName(tabName);
}

/**
 * Calculate urgency level
 */
function getUrgencyLevel(daysRemaining) {
  if (daysRemaining === null) return 'none';
  if (daysRemaining < 0) return 'overdue';
  if (daysRemaining <= 5) return 'critical';
  if (daysRemaining <= 10) return 'high';
  if (daysRemaining <= 20) return 'medium';
  return 'low';
}

/**
 * Test function
 */
function testDataRetrieval() {
  Logger.log('Testing data retrieval...');
  Logger.log('Exec 1 Deliverables:', JSON.stringify(getExec1Deliverables(), null, 2));
  Logger.log('Meetings:', JSON.stringify(getMeetings(), null, 2));
  Logger.log('Recurring Meetings:', JSON.stringify(getRecurringMeetings(), null, 2));
}

// ============================================================================
// V8 NEW FUNCTIONS
// ============================================================================

/**
 * Get Document Repository Links (separate from main document list)
 */
function getDocumentRepoLinks() {
  try {
    const sheet = getSheet(CONFIG.DELIVERABLES_SHEET_ID, CONFIG.DELIVERABLES_TAB_NAME);
    const startRow = CONFIG.DOCUMENTS_START_ROW;
    const maxRows = CONFIG.DOCUMENTS_END_ROW - startRow + 1;
    
    const range = sheet.getRange(startRow, 1, maxRows, 7);
    const values = range.getValues();
    const richTextValues = range.getRichTextValues();
    
    const links = [];
    for (let i = 0; i < values.length; i++) {
      const row = values[i];
      if (!row[0] || row[0].toString().trim() === '') break;
      
      // Extract URL from Column D hyperlink
      let url = '';
      const richText = richTextValues[i][3];
      if (richText) {
        url = richText.getLinkUrl() || '';
      }
      if (!url && row[3]) {
        url = String(row[3]);
      }
      
      links.push({
        id: 'doc-repo-' + i,
        name: String(row[0] || ''),
        description: String(row[0] || ''),
        url: url,  // Column D
        link: url,  // Column D
        owner: String(row[2] || ''),
        comment: String(row[4] || ''),
        category: 'document-repo'
      });
    }
    
    return links;
  } catch (error) {
    Logger.log('Error getting document repo links: ' + error);
    return [];
  }
}

/**
 * Complete an item and move it to destination
 */
function completeItem(itemId, destination) {
  try {
    Logger.log('Completing item: ' + itemId + ' to ' + destination);
    
    const ss = SpreadsheetApp.openById(CONFIG.DELIVERABLES_SHEET_ID);
    const sheet = ss.getSheetByName(CONFIG.DELIVERABLES_TAB_NAME);
    
    if (!sheet) {
      return { success: false, error: 'Sheet not found: ' + CONFIG.DELIVERABLES_TAB_NAME };
    }
    
    // Parse item ID to extract section and index (e.g., "exec1-5" or "document-3")
    const idParts = itemId.split('-');
    const itemType = idParts[0];
    const itemIndex = parseInt(idParts[1]);
    
    // Determine source section based on item type
    let startRow, endRow;
    if (itemType === 'exec1') {
      startRow = CONFIG.EXEC1_START_ROW;
      endRow = CONFIG.EXEC1_END_ROW;
    } else if (itemType === 'exec2') {
      startRow = CONFIG.EXEC2_START_ROW;
      endRow = CONFIG.EXEC2_END_ROW;
    } else if (itemType === 'document') {
      startRow = CONFIG.DOCUMENTS_START_ROW;
      endRow = CONFIG.DOCUMENTS_END_ROW;
    } else {
      return { success: false, error: 'Unknown item type: ' + itemType };
    }
    
    // Calculate actual row number
    const sourceRow = startRow + itemIndex;
    
    if (sourceRow > endRow) {
      return { success: false, error: 'Item index out of range' };
    }
    
    // Get the item data (7 columns standard)
    const sourceRange = sheet.getRange(sourceRow, 1, 1, 7);
    const itemData = sourceRange.getValues()[0];
    
    // Determine destination section (typically "completed" section)
    let destinationStartRow;
    if (destination === 'completed') {
      destinationStartRow = CONFIG.COMPLETED_START_ROW;
    } else {
      return { success: false, error: 'Unknown destination: ' + destination };
    }
    
    // Find first empty row in destination
    let destinationRow = destinationStartRow;
    const range = sheet.getRange(destinationStartRow, 1, CONFIG.COMPLETED_END_ROW - destinationStartRow + 1, 1);
    const values = range.getValues();
    
    for (let i = 0; i < values.length; i++) {
      if (!values[i][0] || values[i][0].toString().trim() === '') {
        destinationRow = destinationStartRow + i;
        break;
      }
    }
    
    // Move item to destination
    sheet.getRange(destinationRow, 1, 1, 7).setValues([itemData]);
    
    // Clear source row
    sourceRange.clearContent();
    
    Logger.log('✅ Item ' + itemId + ' moved to ' + destination + ' at row ' + destinationRow);
    return { 
      success: true, 
      message: 'Item completed and moved to ' + destination,
      destinationRow: destinationRow
    };
    
  } catch (error) {
    Logger.log('Error completing item: ' + error);
    return { success: false, error: error.toString() };
  }
}

/**
 * Add new item to deliverables
 */
function addNewItem(itemData) {
  try {
    Logger.log('Adding new item to: ' + itemData.section);
    
    const sheet = getSheet(CONFIG.DELIVERABLES_SHEET_ID, CONFIG.DELIVERABLES_TAB_NAME);
    
    // Determine target section
    let targetStartRow;
    let targetEndRow;
    if (itemData.section === 'exec1') {
      targetStartRow = CONFIG.EXEC1_START_ROW;
      targetEndRow = CONFIG.EXEC1_END_ROW;
    } else if (itemData.section === 'exec2') {
      targetStartRow = CONFIG.EXEC2_START_ROW;
      targetEndRow = CONFIG.EXEC2_END_ROW;
    } else {
      throw new Error('Unknown section: ' + itemData.section);
    }
    
    // Find first empty row in section
    let targetRow = targetStartRow;
    while (targetRow <= targetEndRow && sheet.getRange(targetRow, 1).getValue() !== '') {
      targetRow++;
    }
    
    if (targetRow > targetEndRow) {
      throw new Error('No empty rows available in section: ' + itemData.section);
    }
    
    // Prepare row data (adjust columns based on your sheet structure)
    const rowData = [
      itemData.description,
      itemData.eta || '',
      itemData.owner || '',
      itemData.link || ''
    ];
    
    // Insert data into columns A-D
    sheet.getRange(targetRow, 1, 1, rowData.length).setValues([rowData]);
    // Comments go in column G
    sheet.getRange(targetRow, 7).setValue(itemData.comments || '');
    
    return { success: true, message: 'Item added successfully' };
    
  } catch (error) {
    Logger.log('Error adding item: ' + error);
    throw error;
  }
}
// ============================================================================
// DELIVERABLE CRUD FUNCTIONS
// ============================================================================

/**
 * Update a deliverable item
 * @param {Object} data - Deliverable data including execType, itemIndex, and all fields
 * @return {Object} Success/error response
 */
function updateDeliverable(data) {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.DELIVERABLES_SHEET_ID);
    const sheet = ss.getSheetByName(CONFIG.DELIVERABLES_TAB_NAME);
    
    if (!sheet) {
      return { success: false, error: 'Sheet not found: ' + CONFIG.DELIVERABLES_TAB_NAME };
    }
    
    let execType = data.execType;
    let itemIndex = data.itemIndex;
    
    if ((!execType || itemIndex === undefined) && data.id) {
      const parts = String(data.id).split('-');
      if (parts.length === 2) {
        execType = parts[0];
        itemIndex = parseInt(parts[1], 10);
      }
    }
    
    if (execType !== 'exec1' && execType !== 'exec2') {
      return { success: false, error: 'Invalid execType: ' + execType };
    }
    
    if (isNaN(itemIndex)) {
      return { success: false, error: 'Invalid item index' };
    }
    
    const startRow = execType === 'exec1' ? CONFIG.EXEC1_START_ROW : CONFIG.EXEC2_START_ROW;
    const endRow = execType === 'exec1' ? CONFIG.EXEC1_END_ROW : CONFIG.EXEC2_END_ROW;
    const actualRow = startRow + itemIndex;
    
    if (actualRow > endRow) {
      return { success: false, error: 'Item index out of range' };
    }
    
    // Update columns A-D and G (comment). Preserve calculated columns.
    sheet.getRange(actualRow, 1, 1, 4).setValues([[
      data.description || '',
      data.eta || '',
      data.owner || '',
      data.link || ''
    ]]);
    sheet.getRange(actualRow, 7).setValue(data.comments || '');
    
    if (data.link) {
      const linkCell = sheet.getRange(actualRow, 4);
      linkCell.setFormula('=HYPERLINK("' + data.link + '", "Link to File")');
    }
    
    Logger.log('✅ Updated deliverable: ' + data.description);
    return { success: true, message: 'Deliverable updated successfully' };
    
  } catch (error) {
    Logger.log('❌ Error updating deliverable: ' + error.message);
    return { success: false, error: error.message };
  }
}

/**
 * Delete a deliverable item
 * @param {Object} data - Contains execType and itemIndex
 * @return {Object} Success/error response
 */
function deleteDeliverable(data) {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.DELIVERABLES_SHEET_ID);
    const sheet = ss.getSheetByName(CONFIG.DELIVERABLES_TAB_NAME);
    
    if (!sheet) {
      return { success: false, error: 'Sheet not found: ' + CONFIG.DELIVERABLES_TAB_NAME };
    }
    
    // Determine the row range based on execType
    let startRow, endRow;
    if (data.execType === 'exec1') {
      startRow = CONFIG.EXEC1_START_ROW;
      endRow = CONFIG.EXEC1_END_ROW;
    } else if (data.execType === 'exec2') {
      startRow = CONFIG.EXEC2_START_ROW;
      endRow = CONFIG.EXEC2_END_ROW;
    } else {
      return { success: false, error: 'Invalid execType: ' + data.execType };
    }
    
    // Calculate actual row number
    const actualRow = startRow + data.itemIndex;
    
    if (actualRow > endRow) {
      return { success: false, error: 'Item index out of range' };
    }
    
    // Clear the row (don't delete to preserve row numbers)
    const range = sheet.getRange(actualRow, 1, 1, 7);
    range.clearContent();
    
    Logger.log('✅ Deleted deliverable at row: ' + actualRow);
    return { success: true, message: 'Deliverable deleted successfully' };
    
  } catch (error) {
    Logger.log('❌ Error deleting deliverable: ' + error.message);
    return { success: false, error: error.message };
  }
}

/**
 * Add a new deliverable item
 * @param {Object} data - Deliverable data including execType and all fields
 * @return {Object} Success/error response
 */
function addDeliverable(data) {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.DELIVERABLES_SHEET_ID);
    const sheet = ss.getSheetByName(CONFIG.DELIVERABLES_TAB_NAME);
    
    if (!sheet) {
      return { success: false, error: 'Sheet not found: ' + CONFIG.DELIVERABLES_TAB_NAME };
    }
    
    // Determine the row range based on execType
    let startRow, endRow;
    if (data.execType === 'exec1') {
      startRow = CONFIG.EXEC1_START_ROW;
      endRow = CONFIG.EXEC1_END_ROW;
    } else if (data.execType === 'exec2') {
      startRow = CONFIG.EXEC2_START_ROW;
      endRow = CONFIG.EXEC2_END_ROW;
    } else {
      return { success: false, error: 'Invalid execType: ' + data.execType };
    }
    
    // Find the first empty row in the range
    const range = sheet.getRange(startRow, 1, endRow - startRow + 1, 1);
    const values = range.getValues();
    let emptyRow = null;
    
    for (let i = 0; i < values.length; i++) {
      if (!values[i][0] || values[i][0].toString().trim() === '') {
        emptyRow = startRow + i;
        break;
      }
    }
    
    if (!emptyRow) {
      return { success: false, error: 'No empty rows available in ' + data.execType + ' section' };
    }
    
    // Add the new item
    const newRange = sheet.getRange(emptyRow, 1, 1, 4);
    const newValues = [[
      data.description || '',
      data.eta || '',
      data.owner || '',
      data.link || ''
    ]];
    
    newRange.setValues(newValues);
    sheet.getRange(emptyRow, 7).setValue(data.comments || '');
    
    // Set hyperlink in Column D if provided
    if (data.link) {
      const linkCell = sheet.getRange(emptyRow, 4);
      linkCell.setFormula('=HYPERLINK("' + data.link + '", "Link to File")');
    }
    
    Logger.log('✅ Added new deliverable: ' + data.description + ' at row ' + emptyRow);
    return { success: true, message: 'Deliverable added successfully', row: emptyRow };
    
  } catch (error) {
    Logger.log('❌ Error adding deliverable: ' + error.message);
    return { success: false, error: error.message };
  }
}

/**
 * Generate PDF of 2-week schedule
 * @param {Object} options - Export options including selected meetings, week title, etc.
 * @return {Object} - Contains success status and download URL or error
 */
function generateWeeklySchedulePDF(options) {
  try {
    const { selectedMeetingIds, weekTitle, includePrepMeetings, agendaItems, notes } = options;
    
    // Get all meetings
    const allMeetings = getMeetings();
    
    // Filter to selected meetings only
    let meetings = allMeetings;
    if (selectedMeetingIds && selectedMeetingIds.length > 0) {
      meetings = allMeetings.filter(m => selectedMeetingIds.includes(m.id));
    }
    
    // Create a new Google Doc for the PDF
    const doc = DocumentApp.create('Two-Week Schedule - ' + (weekTitle || new Date().toLocaleDateString()));
    const body = doc.getBody();
    
    // Set page to landscape
    body.setPageWidth(792).setPageHeight(612); // Letter landscape
    body.setMarginTop(36);
    body.setMarginBottom(36);
    body.setMarginLeft(36);
    body.setMarginRight(36);
    
    // Compute current two-week range (Mon-Fri + next week)
    const scheduleToday = new Date();
    scheduleToday.setHours(0, 0, 0, 0);
    const scheduleDay = scheduleToday.getDay();
    const scheduleMondayOffset = scheduleDay === 0 ? -6 : 1 - scheduleDay;
    const monday = new Date(scheduleToday);
    monday.setDate(scheduleToday.getDate() + scheduleMondayOffset);
    const secondWeekMonday = new Date(monday);
    secondWeekMonday.setDate(monday.getDate() + 7);
    const rangeEnd = new Date(secondWeekMonday);
    rangeEnd.setDate(secondWeekMonday.getDate() + 4);
    const rangeLabel = Utilities.formatDate(monday, Session.getScriptTimeZone(), 'MMMM d') +
      '-' + Utilities.formatDate(rangeEnd, Session.getScriptTimeZone(), 'd, yyyy');
    const generatedLabel = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'MMM d, yyyy');

    // Header row (title left, generated right)
    const headerTable = body.appendTable([['', '']]);
    headerTable.setBorderWidth(0);
    const headerRow = headerTable.getRow(0);
    const titleCell = headerRow.getCell(0);
    const generatedCell = headerRow.getCell(1);
    titleCell.setWidth(420);
    generatedCell.setWidth(200);
    
    const titlePara = titleCell.appendParagraph('Executive Calendar Report');
    titlePara.setFontSize(16);
    titlePara.setBold(true);
    titlePara.setForegroundColor('#cc0000');
    
    const dateRangePara = titleCell.appendParagraph(rangeLabel);
    dateRangePara.setFontSize(9);
    dateRangePara.setForegroundColor('#555555');
    
    const generatedPara = generatedCell.appendParagraph('Generated: ' + generatedLabel);
    generatedPara.setFontSize(8);
    generatedPara.setAlignment(DocumentApp.HorizontalAlignment.RIGHT);
    generatedPara.setForegroundColor('#555555');
    
    // Red divider
    body.appendHorizontalRule();
    body.appendParagraph('').setSpacingAfter(6);

    // Executive summary box
    const summaryTable = body.appendTable([['', '']]);
    summaryTable.setBorderWidth(0);
    const summaryRow = summaryTable.getRow(0);
    const barCell = summaryRow.getCell(0);
    const summaryCell = summaryRow.getCell(1);
    barCell.setBackgroundColor('#cc0000');
    barCell.setWidth(6);
    summaryCell.setBackgroundColor('#f8fafc');
    summaryCell.setWidth(600);
    
    const summaryHeader = summaryCell.appendParagraph('EXECUTIVE SUMMARY');
    summaryHeader.setFontSize(9);
    summaryHeader.setBold(true);
    summaryHeader.setForegroundColor('#cc0000');
    summaryHeader.setSpacingAfter(6);
    
    const totalMeetings = meetings.length;
    const hotCount = meetings.filter(m => m.isHotTopic).length;
    const prepCount = meetings.filter(m => m.prepRequired).length;
    const conflictCount = meetings.filter(m => m.hasConflict).length;
    const summaryText = notes && notes.trim()
      ? notes.trim()
      : `Two-week outlook includes ${totalMeetings} meetings, ${hotCount} hot topics, ${prepCount} prep-required sessions, and ${conflictCount} conflicts.`;
    const summaryPara = summaryCell.appendParagraph(summaryText);
    summaryPara.setFontSize(9);
    summaryPara.setForegroundColor('#111111');
    summaryPara.setSpacingAfter(8);
    
    // Group meetings by date
    const meetingsByDate = {};
    meetings.forEach(m => {
      if (!meetingsByDate[m.date]) {
        meetingsByDate[m.date] = [];
      }
      meetingsByDate[m.date].push(m);
    });
    
    // Detect conflicts (meetings at same time on same day)
    for (const date of Object.keys(meetingsByDate)) {
      const dayMeetings = meetingsByDate[date];
      const timeSlots = {};
      
      dayMeetings.forEach(m => {
        const timeKey = m.time24 || '';
        if (!timeSlots[timeKey]) {
          timeSlots[timeKey] = [];
        }
        timeSlots[timeKey].push(m);
      });
      
      for (const timeKey of Object.keys(timeSlots)) {
        if (timeSlots[timeKey].length > 1) {
          timeSlots[timeKey].forEach(m => {
            m.hasConflict = true;
          });
        }
      }
    }
    
    const dayNames = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday'];

    // Key action items (deliverables due this week for Mike Sarcone)
    const actionHeader = body.appendParagraph("☑ Sarcone's Upcoming Deliverables");
    actionHeader.setFontSize(10);
    actionHeader.setBold(true);
    actionHeader.setForegroundColor('#cc0000');
    actionHeader.setSpacingAfter(6);
    
    const exec1Deliverables = getExec1Deliverables();
    const today = new Date();
    today.setHours(0, 0, 0, 0);
    const weekDay = today.getDay();
    const mondayOffset = weekDay === 0 ? -6 : 1 - weekDay;
    const weekStart = new Date(today);
    weekStart.setDate(today.getDate() + mondayOffset);
    const weekEnd = new Date(weekStart);
    weekEnd.setDate(weekStart.getDate() + 4);
    
    const weekDeliverables = exec1Deliverables.filter(d => {
      if (!d.etaFull) return false;
      const due = new Date(d.etaFull + 'T00:00:00');
      return due >= weekStart && due <= weekEnd;
    });
    
    if (weekDeliverables.length === 0) {
      const none = body.appendParagraph('No deliverables due this week.');
      none.setFontSize(9);
      none.setForegroundColor('#666666');
      none.setSpacingAfter(10);
    } else {
      weekDeliverables.sort((a, b) => a.etaFull.localeCompare(b.etaFull));
      weekDeliverables.forEach(d => {
        const dueLabel = d.eta || d.etaFull;
        const ownerLabel = d.owner ? ` (${d.owner})` : '';
        const overdueLabel = d.isOverdue ? ' — OVERDUE' : '';
        const itemText = `${dueLabel} — ${d.description}${ownerLabel}${overdueLabel}`;
        const item = body.appendListItem(itemText);
        item.setFontSize(9);
        if (d.linkedMaterials) {
          try {
            item.editAsText().setLinkUrl(0, itemText.length - 1, d.linkedMaterials);
          } catch (e) {
            Logger.log('Link set failed for deliverable: ' + e);
          }
        }
      });
      body.appendParagraph('').setSpacingAfter(10);
    }
    
    function appendWeekTable(monday, label) {
      if (label) {
        const labelPara = body.appendParagraph(label);
        labelPara.setFontSize(9);
        labelPara.setBold(true);
        labelPara.setSpacingAfter(6);
      }
      
      const table = body.appendTable();
      table.setBorderWidth(0);
      
      const headerRow = table.appendTableRow();
      const contentRow = table.appendTableRow();
      
      for (let i = 0; i < 5; i++) {
        const dayDate = new Date(monday);
        dayDate.setDate(monday.getDate() + i);
        const dateShort = Utilities.formatDate(dayDate, Session.getScriptTimeZone(), 'MMM d');
        const dateStr = Utilities.formatDate(dayDate, Session.getScriptTimeZone(), 'yyyy-MM-dd');
        const dayName = dayNames[i];
        
        const headerCell = headerRow.appendTableCell(`${dayName}\n${dateShort}`);
        headerCell.setBackgroundColor('#000000');
        headerCell.getChild(0).asParagraph().setForegroundColor('#ffffff');
        headerCell.getChild(0).asParagraph().setFontSize(8);
        headerCell.getChild(0).asParagraph().setBold(true);
        headerCell.setWidth(144);
        
        const bodyCell = contentRow.appendTableCell();
        bodyCell.setVerticalAlignment(DocumentApp.VerticalAlignment.TOP);
        
        const dayMeetings = meetingsByDate[dateStr] || [];
        dayMeetings.sort((a, b) => (a.time24 || '').localeCompare(b.time24 || ''));
        
        if (dayMeetings.length === 0) {
          const empty = bodyCell.appendParagraph('No meetings');
          empty.setFontSize(8);
          empty.setForegroundColor('#999999');
        } else {
          dayMeetings.forEach((m, idx) => {
            const timePara = bodyCell.appendParagraph(m.time || 'TBD');
            timePara.setBold(true);
            timePara.setFontSize(8);
            
            const titlePara = bodyCell.appendParagraph((includePrepMeetings && m.prepRequired ? '[PREP] ' : '') + m.title);
            titlePara.setFontSize(8);
            
            if (m.hotTopicLevel === 'urgent') {
              titlePara.setForegroundColor('#cc0000');
            } else if (m.hotTopicLevel === 'must-attend') {
              titlePara.setForegroundColor('#ea580c');
            } else if (m.isHotTopic) {
              titlePara.setForegroundColor('#444444');
            }
            
            if (m.hasConflict) {
              const conflictPara = bodyCell.appendParagraph('⚠ CONFLICT');
              conflictPara.setForegroundColor('#ff6600');
              conflictPara.setFontSize(7);
              conflictPara.setBold(true);
            }
            
            if (idx < dayMeetings.length - 1) {
              bodyCell.appendParagraph('');
            }
          });
        }
      }
    }
    
    // Two-week calendar view
    const calendarHeader = body.appendParagraph('TWO-WEEK CALENDAR VIEW');
    calendarHeader.setFontSize(10);
    calendarHeader.setBold(true);
    calendarHeader.setForegroundColor('#cc0000');
    calendarHeader.setSpacingAfter(6);
    
    appendWeekTable(monday, 'Week 1');
    body.appendParagraph('');
    appendWeekTable(secondWeekMonday, 'Week 2');
    
    // Add Verizon footer
    body.appendParagraph('');
    const footer = body.appendParagraph('verizon');
    footer.setFontFamily('Arial');
    footer.setBold(true);
    footer.setFontSize(14);
    footer.setForegroundColor('#cc0000');
    
    // Save and get PDF
    doc.saveAndClose();
    
    // Convert to PDF
    const docFile = DriveApp.getFileById(doc.getId());
    const pdfBlob = docFile.getAs('application/pdf');
    pdfBlob.setName('Two_Week_Schedule_' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd') + '.pdf');
    
    // Save PDF to Drive
    const pdfFile = DriveApp.createFile(pdfBlob);
    
    // Delete the temporary Doc
    docFile.setTrashed(true);
    
    return {
      success: true,
      pdfUrl: pdfFile.getUrl(),
      pdfId: pdfFile.getId(),
      downloadUrl: 'https://drive.google.com/uc?export=download&id=' + pdfFile.getId()
    };
    
  } catch (error) {
    Logger.log('Error generating PDF: ' + error);
    return {
      success: false,
      error: error.toString()
    };
  }
}

/**
 * Get meetings data formatted for PDF export selection
 */
function getMeetingsForPDFExport() {
  try {
    const meetings = getMeetings();
    
    // Group by date and detect conflicts
    const meetingsByDate = {};
    meetings.forEach(m => {
      if (!meetingsByDate[m.date]) {
        meetingsByDate[m.date] = [];
      }
      meetingsByDate[m.date].push(m);
    });
    
    // Detect conflicts
    for (const date of Object.keys(meetingsByDate)) {
      const dayMeetings = meetingsByDate[date];
      const timeSlots = {};
      
      dayMeetings.forEach(m => {
        const timeKey = m.time24;
        if (!timeSlots[timeKey]) {
          timeSlots[timeKey] = [];
        }
        timeSlots[timeKey].push(m);
      });
      
      // Mark conflicts
      for (const timeKey of Object.keys(timeSlots)) {
        if (timeSlots[timeKey].length > 1) {
          timeSlots[timeKey].forEach(m => {
            m.hasConflict = true;
            m.conflictCount = timeSlots[timeKey].length;
          });
        }
      }
    }
    
    return {
      success: true,
      meetings: meetings,
      meetingsByDate: meetingsByDate
    };
    
  } catch (error) {
    Logger.log('Error getting meetings for PDF: ' + error);
    return {
      success: false,
      error: error.toString()
    };
  }
}
