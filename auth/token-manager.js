/**
 * Token management for Microsoft Graph API authentication
 */
const fs = require('fs');
const config = require('../config');
const { refreshAccessToken, needsRefresh } = require('./auto-refresh');

// Global variable to store tokens
let cachedTokens = null;

// Debounce concurrent refresh requests
let refreshPromise = null;

/**
 * Loads authentication tokens from the token file
 * @returns {object|null} - The loaded tokens or null if not available
 */
function loadTokenCache() {
  try {
    const tokenPath = config.AUTH_CONFIG.tokenStorePath;
    
    // Quick existence check without stat
    if (!fs.existsSync(tokenPath)) {
      console.error('[TOKEN-MANAGER] Token file does not exist');
      return null;
    }
    
    const tokenData = fs.readFileSync(tokenPath, 'utf8');
    
    try {
      const tokens = JSON.parse(tokenData);
      
      // Check for access token presence
      if (!tokens.access_token) {
        console.error('[TOKEN-MANAGER] No access_token found in tokens');
        return null;
      }
      
      // Update the cache
      cachedTokens = tokens;
      return tokens;
    } catch (parseError) {
      console.error('[TOKEN-MANAGER] Error parsing token JSON:', parseError.message);
      return null;
    }
  } catch (error) {
    console.error('[TOKEN-MANAGER] Error loading token cache:', error.message);
    return null;
  }
}

/**
 * Saves authentication tokens to the token file
 * @param {object} tokens - The tokens to save
 * @returns {boolean} - Whether the save was successful
 */
function saveTokenCache(tokens) {
  try {
    const tokenPath = config.AUTH_CONFIG.tokenStorePath;
    console.error(`Saving tokens to: ${tokenPath}`);
    
    fs.writeFileSync(tokenPath, JSON.stringify(tokens, null, 2));
    console.error('Tokens saved successfully');
    
    // Update the cache
    cachedTokens = tokens;
    return true;
  } catch (error) {
    console.error('Error saving token cache:', error);
    return false;
  }
}

/**
 * Gets the current access token, loading from cache if necessary
 * @param {boolean} autoRefresh - Whether to automatically refresh expired tokens
 * @returns {Promise<string>|string|null} - The access token or null if not available
 */
async function getAccessToken(autoRefresh = true) {
  // First check cache
  if (cachedTokens && cachedTokens.access_token) {
    if (autoRefresh && needsRefresh(cachedTokens)) {
      console.error('[TOKEN-MANAGER] Cached token needs refresh');
      
      // Debounce concurrent refresh requests
      if (!refreshPromise) {
        refreshPromise = refreshAccessToken()
          .then(newTokens => {
            cachedTokens = newTokens;
            refreshPromise = null;
            return newTokens;
          })
          .catch(error => {
            console.error('[TOKEN-MANAGER] Auto-refresh failed:', error.message);
            refreshPromise = null;
            throw error;
          });
      }
      
      try {
        const newTokens = await refreshPromise;
        return newTokens.access_token;
      } catch (error) {
        // Return existing token if refresh fails (might still work briefly)
        return cachedTokens.access_token;
      }
    }
    return cachedTokens.access_token;
  }
  
  // Load from file
  const tokens = loadTokenCache();
  if (!tokens || !tokens.access_token) {
    return null;
  }
  
  // Check if refresh needed
  if (autoRefresh && needsRefresh(tokens)) {
    console.error('[TOKEN-MANAGER] Token needs refresh');
    
    // Debounce concurrent refresh requests
    if (!refreshPromise) {
      refreshPromise = refreshAccessToken()
        .then(newTokens => {
          cachedTokens = newTokens;
          refreshPromise = null;
          return newTokens;
        })
        .catch(error => {
          console.error('[TOKEN-MANAGER] Auto-refresh failed:', error.message);
          refreshPromise = null;
          throw error;
        });
    }
    
    try {
      const newTokens = await refreshPromise;
      return newTokens.access_token;
    } catch (error) {
      // Return existing token if refresh fails
      return tokens.access_token;
    }
  }
  
  return tokens.access_token;
}

/**
 * Creates a test access token for use in test mode
 * @returns {object} - The test tokens
 */
function createTestTokens() {
  const testTokens = {
    access_token: "test_access_token_" + Date.now(),
    refresh_token: "test_refresh_token_" + Date.now(),
    expires_at: Date.now() + (3600 * 1000) // 1 hour
  };
  
  saveTokenCache(testTokens);
  return testTokens;
}

module.exports = {
  loadTokenCache,
  saveTokenCache,
  getAccessToken,
  createTestTokens
};
