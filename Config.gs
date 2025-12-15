/**
 * CENTRAL CONFIGURATION FILE
 * All community settings and sheet IDs in one place
 * UPDATED: Type column completely removed - no more Member/Guest distinction
 */

// ==================== MASTER SHEET ====================
const MASTER_SHEET_ID = '1Z7JrY-7US55iUnNmyXB3m5hKFGoWpR0bGMW1XXRgWnQ';

// ==================== COMMUNITY SHEETS ====================
const COMMUNITY_SHEETS = {
  'Belvedere Family Church': '1U9VS9kEH7DGD_gbwqo1bzVTfgHgL-WS3mwQlysoNGJE',
  'Boston Family Church': '10GAXtMKcys3iTnqvL76wFYrh24wdjgCWEbF-9XJzeEI',
  'Connecticut Family Church': '1FQm-agxjxztFgHyY5_HHFEDFsari6MtmbVRs44nmnlc',
  'Elizabeth Family Church': '1BS_VKC54LXzMt8Td4ubhpwQozID7VuSkJS96-Hy1yk0',
  'Manhattan Family Church': '1kL_S0OpOlJqDwQBEa3yIk0XXxKogI7Dg8--yEZy_7bk',
  'New Jersey Family Church': '1G_ErSZ2ItZETLEG1Ipy7HeO0tzgLwHZxFohTUUDAIHg',
  'Philadelphia Family Church': '1vlThBRqh3f9WsvxVqbf-Z-xsZtAuwOtU88reVDeUw0A',
  'Worcester Family Church': '1D1AfcLkg15dYQI4CFnUwQqiueZw-TcOMUcaZGe2_4t4'
};

// ==================== TAB NAMES ====================
const TAB_NAMES = {
  REGISTRATIONS: 'Registrations',
  MONTHLY: 'Current Month',
  YEARLY: 'Yearly Summary',
  SUNDAY: 'Sunday Breakdown'
};

// Master sheet specific tabs
const MASTER_TAB_NAMES = {
  BELVEDERE: 'Belvedere Family Church',
  BOSTON: 'Boston Family Church',
  CONNECTICUT: 'Connecticut Family Church',
  ELIZABETH: 'Elizabeth Family Church',
  MANHATTAN: 'Manhattan Family Church',
  NEW_JERSEY: 'New Jersey Family Church',
  PHILADELPHIA: 'Philadelphia Family Church',
  WORCESTER: 'Worcester Family Church',
  OVERALL_SUMMARY: 'Overall Summary',
  COMMUNITY_COMPARISON: 'Community Comparison',
  WEEKLY_TRENDS: 'Weekly Trends'
};

// ==================== COLUMN STRUCTURE ====================
// Standard columns - TYPE REMOVED
const COLUMNS = {
  TIMESTAMP: 'Timestamp',
  COMMUNITY: 'Community',
  FIRST_NAME: 'First Name',
  LAST_NAME: 'Last Name',
  EMAIL: 'Email',
  SESSION: 'Session',
  SUNDAY_DATE: 'Sunday Date'
};

// Column indices (0-based) - TYPE REMOVED
const COL_INDEX = {
  TIMESTAMP: 0,
  COMMUNITY: 1,
  FIRST_NAME: 2,
  LAST_NAME: 3,
  EMAIL: 4,
  SESSION: 5,
  SUNDAY_DATE: 6
};

// ==================== SETTINGS ====================
const SETTINGS = {
  LIVE_STREAM_LINK: 'https://www.youtube.com/live/xU2alDtuXkw',
  SERVICE_END_HOUR: 14, // 2:00 PM
  
  // Colors
  COLORS: {
    PRIMARY: '#667eea',
    SECONDARY: '#764ba2',
    LIGHT_BG: '#f3f4ff',
    WHITE: '#ffffff',
    GRAY: '#cccccc',
    SUCCESS: '#28a745',
    DANGER: '#dc3545'
  }
};

// ==================== COMMUNITY LIVESTREAM LINKS ====================
const COMMUNITY_LIVESTREAM_LINKS = {
  'Belvedere Family Church': 'https://www.youtube.com/live/xU2alDtuXkw',
  'Boston Family Church': 'https://www.youtube.com/live/xU2alDtuXkw',
  'Connecticut Family Church': 'https://www.youtube.com/live/xU2alDtuXkw',
  'Elizabeth Family Church': 'https://www.youtube.com/live/xU2alDtuXkw',
  'Manhattan Family Church': 'https://www.youtube.com/live/xU2alDtuXkw',
  'New Jersey Family Church': 'https://www.youtube.com/live/xU2alDtuXkw',
  'Philadelphia Family Church': 'https://www.youtube.com/live/xU2alDtuXkw',
  'Worcester Family Church': 'https://www.youtube.com/live/xU2alDtuXkw'
};

/**
 * Get all community names
 */
function getCommunityNames() {
  return Object.keys(COMMUNITY_SHEETS);
}

/**
 * Get sheet ID for a community
 */
function getCommunitySheetId(communityName) {
  return COMMUNITY_SHEETS[communityName];
}

/**
 * Validate community name
 */
function isValidCommunity(communityName) {
  return COMMUNITY_SHEETS.hasOwnProperty(communityName);
}

/**
 * Get livestream link for a community (fast lookup, no spreadsheet access)
 */
function getCommunityLiveStreamLinkFast(communityName) {
  return COMMUNITY_LIVESTREAM_LINKS[communityName] || SETTINGS.LIVE_STREAM_LINK;
}
