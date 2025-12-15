/**
 * COMMUNITIES CONFIGURATION
 * Consolidated configuration for all 8 communities
 * Replaces individual community files (Belvedere.gs, Boston.gs, etc.)
 */

// ==================== COMMUNITY DEFINITIONS ====================

const COMMUNITIES = {
  BELVEDERE: {
    name: 'Belvedere Family Church',
    sheetId: '1U9VS9kEH7DGD_gbwqo1bzVTfgHgL-WS3mwQlysoNGJE'
  },
  BOSTON: {
    name: 'Boston Family Church',
    sheetId: '10GAXtMKcys3iTnqvL76wFYrh24wdjgCWEbF-9XJzeEI'
  },
  CONNECTICUT: {
    name: 'Connecticut Family Church',
    sheetId: '1FQm-agxjxztFgHyY5_HHFEDFsari6MtmbVRs44nmnlc'
  },
  ELIZABETH: {
    name: 'Elizabeth Family Church',
    sheetId: '1BS_VKC54LXzMt8Td4ubhpwQozID7VuSkJS96-Hy1yk0'
  },
  MANHATTAN: {
    name: 'Manhattan Family Church',
    sheetId: '1kL_S0OpOlJqDwQBEa3yIk0XXxKogI7Dg8--yEZy_7bk'
  },
  NEW_JERSEY: {
    name: 'New Jersey Family Church',
    sheetId: '1G_ErSZ2ItZETLEG1Ipy7HeO0tzgLwHZxFohTUUDAIHg'
  },
  PHILADELPHIA: {
    name: 'Philadelphia Family Church',
    sheetId: '1vlThBRqh3f9WsvxVqbf-Z-xsZtAuwOtU88reVDeUw0A'
  },
  WORCESTER: {
    name: 'Worcester Family Church',
    sheetId: '1D1AfcLkg15dYQI4CFnUwQqiueZw-TcOMUcaZGe2_4t4'
  }
};

// ==================== SETUP FUNCTIONS ====================

/**
 * Setup individual community by name
 * Central function that handles setup for any community
 */
function setupCommunityByName(communityName) {
  try {
    // Validate community exists
    if (!isValidCommunity(communityName)) {
      throw new Error('Invalid community: ' + communityName);
    }
    
    // Run the generic setup
    var result = setupCommunity(communityName);
    
    if (result.success) {
      // Also setup the Service Settings tab
      var sheetId = getCommunitySheetId(communityName);
      var spreadsheet = SpreadsheetApp.openById(sheetId);
      setupServiceSettingsTab(spreadsheet, communityName);
      
      Logger.log('✓ Service Settings tab created for ' + communityName);
    }
    
    return result;
    
  } catch (error) {
    Logger.log('Error in setupCommunityByName: ' + error.toString());
    return {
      success: false,
      message: error.toString()
    };
  }
}

// ==================== INDIVIDUAL SETUP FUNCTIONS ====================
// These functions exist for backward compatibility and menu access

function setupBelvedere() {
  setupCommunityByName(COMMUNITIES.BELVEDERE.name);
}

function setupBoston() {
  setupCommunityByName(COMMUNITIES.BOSTON.name);
}

function setupConnecticut() {
  setupCommunityByName(COMMUNITIES.CONNECTICUT.name);
}

function setupElizabeth() {
  setupCommunityByName(COMMUNITIES.ELIZABETH.name);
}

function setupManhattan() {
  setupCommunityByName(COMMUNITIES.MANHATTAN.name);
}

function setupNewJersey() {
  setupCommunityByName(COMMUNITIES.NEW_JERSEY.name);
}

function setupPhiladelphia() {
  setupCommunityByName(COMMUNITIES.PHILADELPHIA.name);
}

function setupWorcester() {
  setupCommunityByName(COMMUNITIES.WORCESTER.name);
}

// ==================== HELPER FUNCTIONS ====================

/**
 * Get all community configuration objects
 */
function getAllCommunities() {
  return Object.values(COMMUNITIES);
}

/**
 * Get community configuration by name
 */
function getCommunityConfig(communityName) {
  var communities = getAllCommunities();
  for (var i = 0; i < communities.length; i++) {
    if (communities[i].name === communityName) {
      return communities[i];
    }
  }
  return null;
}
