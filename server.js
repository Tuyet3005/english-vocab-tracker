// Load environment variables from .env.local if it exists
require('dotenv').config({ path: '.env.local' });

// Express server for Vocab Tracker
const express = require('express');
const fs = require('fs');
const path = require('path');
const { get, put, list } = require('@tigrisdata/storage');
const graphHelper = require('./lib/graphHelper');
const dataTransformer = require('./lib/dataTransformer');

const app = express();
const PORT = process.env.PORT || 3001;

// File keys for Tigris storage
const STATE_KEY = 'server-state.json';
const CACHE_KEY = 'sheet-data-cache.json';
const STATS_CACHE_KEY = 'sheet-stats-cache.json';
const METADATA_CACHE_KEY = 'sheet-metadata-cache.json';

// IMPORTANT: This data is now stored in TigrisData cloud storage.
// NEVER expose authentication tokens and credentials to the frontend.
// The tokens, clientSecret, and authentication state must remain server-side only.
// Local file paths (kept for backward compatibility in some functions)
const STATE_FILE = path.join(__dirname, 'server-state.json');
const CACHE_FILE = path.join(__dirname, 'sheet-data-cache.json');
const STATS_CACHE_FILE = path.join(__dirname, 'sheet-stats-cache.json');
const METADATA_CACHE_FILE = path.join(__dirname, 'sheet-metadata-cache.json');

// Middleware
app.use(express.json());
// Prevent caching of HTML to ensure fresh content
app.use((req, res, next) => {
  res.setHeader('Cache-Control', 'no-cache, no-store, must-revalidate');
  res.setHeader('Pragma', 'no-cache');
  res.setHeader('Expires', '0');
  next();
});

// Custom route to serve index.html with cached data pre-embedded (BEFORE static middleware)
app.get('/', async (req, res) => {
  try {
    let htmlContent = fs.readFileSync(path.join(__dirname, 'public', 'index.html'), 'utf8');
    
    // Get cached data to embed
    const cachedData = await loadCache();
    if (cachedData && cachedData.data) {
      console.log('Embedding cached data into HTML for instant display');
      // Embed the cached data as a script tag before the closing head tag
      const embeddedScript = `
      <script>
        window.CACHED_DATA = ${JSON.stringify(cachedData.data)};
        window.CACHE_TIMESTAMP = "${cachedData.timestamp}";
        console.log('✅ Embedded cached data loaded successfully');
        console.log('📊 Data contains', window.CACHED_DATA.worksheets ? window.CACHED_DATA.worksheets.length : 0, 'worksheets');
      </script>
      `;
      htmlContent = htmlContent.replace('</head>', embeddedScript + '</head>');
    } else {
      console.log('❌ No cached data available to embed');
      // Add debug script even without data
      const debugScript = `
      <script>
        console.log('❌ No embedded cached data available');
      </script>
      `;
      htmlContent = htmlContent.replace('</head>', debugScript + '</head>');
    }
    
    res.setHeader('Content-Type', 'text/html');
    res.send(htmlContent);
  } catch (err) {
    console.error('Error serving index with cached data:', err);
    res.sendFile(path.join(__dirname, 'public', 'index.html'));
  }
});

// Serve other static files (excluding index.html which is handled above)
app.use(express.static(path.join(__dirname, 'public'), {
  index: false // Don't serve index.html automatically
}));

// App settings
const settings = {
  clientId: '298b0410-cb65-491f-8b6c-4ba5bc612d2a',
  clientSecret: 'QA38Q~5lMVt17bIvDFkRwOB3Jex2DsJ.XUzo2cLv',
  tenantId: 'common',
  graphUserScopes: ['user.read', 'files.read']
};

// Load or initialize server state from Tigris
async function loadState() {
  try {
    const data = await get(STATE_KEY, 'string');
    const parsed = JSON.parse(data.data);
    if (parsed && data.data != '{}') return parsed;
  } catch (err) {
    if (err.code === 'NoSuchKey' || err.message?.includes('not found')) {
      console.log('No state found in Tigris storage, using default state');
    } else {
      console.error('Error loading state from Tigris:', err.message);
    }
  }

  return {
    sheetUrl: 'https://studenthcmusedu-my.sharepoint.com/:x:/g/personal/20120422_student_hcmus_edu_vn/IQAB2D7k19hCSZ1YTQzM7xzuAbC5ZP5VTOe0khG8r7N9o_c?e=HvJTg2',
    sheetName: '',
    token: null,
    expiresOn: null,
    deviceCode: null,
    userCode: null,
    verificationUri: null
  };
}

async function saveState(state) {
  try {
    await put(STATE_KEY, JSON.stringify(state, null, 2));
    console.log('State saved to Tigris storage');
  } catch (err) {
    console.error('Error saving state to Tigris:', err.message);
  }
}

// Initialize server state (will be loaded asynchronously)
let serverState = {
  sheetUrl: 'https://studenthcmusedu-my.sharepoint.com/:x:/g/personal/20120422_student_hcmus_edu_vn/IQAB2D7k19hCSZ1YTQzM7xzuAbC5ZP5VTOe0khG8r7N9o_c?e=HvJTg2',
  sheetName: '',
  token: null,
  expiresOn: null,
  deviceCode: null,
  userCode: null,
  verificationUri: null
};

// Load state from TigrisData on server start
(async () => {
  try {
    const loadedState = await loadState();
    if (loadedState) {
      serverState = loadedState;
      console.log('Server state loaded from TigrisData storage');
    }
  } catch (error) {
    console.error('Failed to load server state:', error.message);
  }
})();

// Cache management with Tigris
async function loadCache() {
  try {
    const data = await get(CACHE_KEY, 'string');
    return JSON.parse(data.data);
  } catch (err) {
    if (err.code === 'NoSuchKey' || err.message?.includes('not found')) {
      console.log('No cache found in Tigris storage');
    } else {
      console.error('Error loading cache from Tigris:', err.message);
    }
  }
  return null;
}

async function saveCache(sheetUrl, data, sheetName = '') {
  try {
    const cache = await loadCache() || {};
    
    // Initialize the sheetUrl object if it doesn't exist
    if (!cache[sheetUrl]) {
      cache[sheetUrl] = {};
    }
    
    // Store data for the specific sheet
    cache[sheetUrl][sheetName || 'default'] = {
      data,
      timestamp: new Date().toISOString()
    };
    
    await put(CACHE_KEY, JSON.stringify(cache, null, 2));
    console.log('Cache saved to Tigris for URL:', sheetUrl, sheetName ? `(Sheet: ${sheetName})` : '(Default sheet)');
  } catch (err) {
    console.error('Error saving cache to Tigris:', err.message);
  }
}

async function getCachedData(sheetUrl, sheetName = '') {
  const cache = await loadCache();
  if (cache && cache[sheetUrl] && cache[sheetUrl][sheetName || 'default']) {
    const cachedSheet = cache[sheetUrl][sheetName || 'default'];
    console.log('Using cached data from:', cachedSheet.timestamp);
    return {
      ...cachedSheet.data,
      _cached: true,
      _cachedAt: cachedSheet.timestamp
    };
  }
  return null;
}

// Clear cache for specific URL
async function clearCacheForUrl(sheetUrl) {
  try {
    const cache = await loadCache() || {};
    if (cache[sheetUrl]) {
      delete cache[sheetUrl];
      await put(CACHE_KEY, JSON.stringify(cache, null, 2));
      console.log('Cleared cache for URL:', sheetUrl);
    }
    
    // Also clear stats cache
    const statsCache = await loadStatsCache() || {};
    if (statsCache[sheetUrl]) {
      delete statsCache[sheetUrl];
      await put(STATS_CACHE_KEY, JSON.stringify(statsCache, null, 2));
      console.log('Cleared stats cache for URL:', sheetUrl);
    }
    // Also clear metadata cache
    const metadataCache = await loadMetadataCache() || {};
    if (metadataCache[sheetUrl]) {
      delete metadataCache[sheetUrl];
      await put(METADATA_CACHE_KEY, JSON.stringify(metadataCache, null, 2));
      console.log('Cleared metadata cache for URL:', sheetUrl);
    }
  } catch (err) {
    console.error('Error clearing cache:', err.message);
  }
}

// Metadata cache management with Tigris
async function loadMetadataCache() {
  try {
    const data = await get(METADATA_CACHE_KEY, 'string');
    return JSON.parse(data.data);
  } catch (err) {
    if (err.code === 'NoSuchKey' || err.message?.includes('not found')) {
      console.log('No metadata cache found in Tigris storage');
    } else {
      console.error('Error loading metadata cache from Tigris:', err.message);
    }
  }
  return null;
}

async function saveMetadataCache(sheetUrl, metadata) {
  try {
    const cache = await loadMetadataCache() || {};
    
    cache[sheetUrl] = {
      metadata,
      timestamp: new Date().toISOString()
    };
    
    await put(METADATA_CACHE_KEY, JSON.stringify(cache, null, 2));
    console.log('Metadata cache saved to Tigris for URL:', sheetUrl);
  } catch (err) {
    console.error('Error saving metadata cache to Tigris:', err.message);
  }
}

async function getCachedMetadata(sheetUrl) {
  const cache = await loadMetadataCache();
  if (cache && cache[sheetUrl]) {
    const cachedMeta = cache[sheetUrl];
    // Check if cache is less than 1 hour old
    const cacheTime = new Date(cachedMeta.timestamp);
    const now = new Date();
    const hoursDiff = (now - cacheTime) / (1000 * 60 * 60);
    
    if (hoursDiff < 1) {
      console.log('Using cached metadata from:', cachedMeta.timestamp);
      return cachedMeta.metadata;
    } else {
      console.log('Metadata cache expired, will fetch fresh');
    }
  }
  return null;
}

// Statistics cache management with Tigris
async function loadStatsCache() {
  try {
    const data = await get(STATS_CACHE_KEY, 'string');
    return JSON.parse(data.data);
  } catch (err) {
    if (err.code === 'NoSuchKey' || err.message?.includes('not found')) {
      console.log('No stats cache found in Tigris storage');
    } else {
      console.error('Error loading stats cache from Tigris:', err.message);
    }
  }
  return null;
}

async function saveStatsCache(sheetUrl, data, sheetName = '') {
  try {
    // Extract statistics from worksheets and topics
    const worksheetStats = [];
    if (data.worksheets && Array.isArray(data.worksheets)) {
      data.worksheets.forEach(worksheet => {
        const topicStats = [];
        if (worksheet.topics && Array.isArray(worksheet.topics)) {
          worksheet.topics.forEach(topic => {
            topicStats.push({
              name: topic.name,
              statistics: topic.statistics || null
            });
          });
        }
        worksheetStats.push({
          name: worksheet.name,
          statistics: worksheet.statistics || null,
          topicStats: topicStats
        });
      });
    }

    const statsCache = await loadStatsCache() || {};
    
    // Initialize the sheetUrl object if it doesn't exist
    if (!statsCache[sheetUrl]) {
      statsCache[sheetUrl] = {};
    }
    
    // Store stats for the specific sheet
    statsCache[sheetUrl][sheetName || 'default'] = {
      worksheetStats,
      timestamp: new Date().toISOString()
    };
    
    await put(STATS_CACHE_KEY, JSON.stringify(statsCache, null, 2));
    console.log('Statistics cache saved to Tigris for URL:', sheetUrl, sheetName ? `(Sheet: ${sheetName})` : '(Default sheet)');
  } catch (err) {
    console.error('Error saving stats cache to Tigris:', err.message);
  }
}

function getCachedStats(sheetUrl, sheetName = '') {
  const cache = loadStatsCache();
  if (cache && cache[sheetUrl] && cache[sheetUrl][sheetName || 'default']) {
    const cachedSheet = cache[sheetUrl][sheetName || 'default'];
    console.log('Using cached statistics from:', cachedSheet.timestamp);
    return {
      sheetUrl,
      sheetName,
      worksheetStats: cachedSheet.worksheetStats,
      timestamp: cachedSheet.timestamp
    };
  }
  return null;
}

// Check if authenticated and token is valid
function isAuthenticated() {
  if (!serverState.token || !serverState.expiresOn) {
    return false;
  }

  const expiresOn = new Date(serverState.expiresOn);
  const now = new Date();
  const bufferMinutes = 5;

  return expiresOn > new Date(now.getTime() + bufferMinutes * 60 * 1000);
}

// Check if token needs refresh (within 10 minutes of expiry)
function shouldRefreshToken() {
  if (!serverState.token || !serverState.expiresOn) {
    return false;
  }

  const expiresOn = new Date(serverState.expiresOn);
  const now = new Date();
  const refreshBufferMinutes = 10;

  // Refresh if token expires in less than 10 minutes
  return expiresOn <= new Date(now.getTime() + refreshBufferMinutes * 60 * 1000);
}

// Refresh authentication token
async function refreshAuthToken() {
  if (!serverState.token) {
    console.log('No token to refresh');
    return false;
  }

  try {
    console.log('Attempting to refresh authentication token...');
    
    // Try to get a new token by making a simple API call
    // The credential provider will automatically refresh the token
    const user = await graphHelper.getUserAsync();
    
    // Get the refreshed token from the helper
    const cachedToken = graphHelper.getCachedToken();
    
    if (cachedToken && cachedToken.token && cachedToken.token !== serverState.token) {
      serverState.token = cachedToken.token;
      serverState.expiresOn = cachedToken.expiresOn;
      await saveState(serverState);
      console.log('Token refreshed successfully. New expiry:', new Date(cachedToken.expiresOn).toLocaleString());
      return true;
    }
    
    console.log('Token is still valid, no refresh needed');
    return true;
  } catch (err) {
    console.error('Failed to refresh token:', err.message);
    return false;
  }
}

// Auto-refresh token periodically
let tokenRefreshTimer = null;

function startTokenRefreshTimer() {
  // Don't start if already running
  if (tokenRefreshTimer) {
    console.log('Token auto-refresh already running');
    return;
  }

  // Check every 5 minutes
  const checkIntervalMinutes = 5;
  
  tokenRefreshTimer = setInterval(async () => {
    if (shouldRefreshToken()) {
      console.log('Token expiring soon, refreshing...');
      await refreshAuthToken();
    }
  }, checkIntervalMinutes * 60 * 1000);
  
  console.log(`Token auto-refresh enabled (checking every ${checkIntervalMinutes} minutes)`);
}

function stopTokenRefreshTimer() {
  if (tokenRefreshTimer) {
    clearInterval(tokenRefreshTimer);
    tokenRefreshTimer = null;
    console.log('Token auto-refresh stopped');
  }
}

// Initialize Graph client and ensure token sync
let deviceCodeInfo = null;

function initializeGraph() {
  graphHelper.initializeGraphForUserAuth(
    settings,
    serverState.token ? { token: serverState.token, expiresOn: serverState.expiresOn } : null,
    (info) => {
      // Store device code info for frontend
      deviceCodeInfo = info;
      serverState.deviceCode = info.deviceCode;
      serverState.userCode = info.userCode;
      serverState.verificationUri = info.verificationUri || 'https://microsoft.com/devicelogin';
      saveState(serverState);
    }
  );
  
  // Debug token state
  console.log('GraphHelper initialized. Token available:', !!serverState.token);
  if (serverState.token) {
    console.log('Token expires at:', new Date(serverState.expiresOn).toLocaleString());
  }
}

initializeGraph();

// Start automatic token refresh if already authenticated
if (isAuthenticated()) {
  startTokenRefreshTimer();
  const expiresOn = new Date(serverState.expiresOn);
  console.log('Already authenticated. Token expires at:', expiresOn.toLocaleString());
}

// Routes

// Get authentication status (non-sensitive data only)
app.get('/api/auth/status', (req, res) => {
  res.json({
    authenticated: isAuthenticated(),
    userCode: serverState.userCode,
    verificationUri: serverState.verificationUri
  });
});

// Start authentication process
app.post('/api/auth/start', async (req, res) => {
  try {
    // Stop token refresh timer
    stopTokenRefreshTimer();
    
    // Reset state
    deviceCodeInfo = null;
    serverState.token = null;
    serverState.expiresOn = null;
    serverState.deviceCode = null;
    serverState.userCode = null;
    serverState.verificationUri = null;
    saveState(serverState);

    // Reinitialize to trigger device code flow
    initializeGraph();

    // Trigger the device code flow by attempting to get user info
    // This will invoke the device code callback
    // We don't await it because we just want to trigger the callback
    graphHelper.getUserAsync().catch(() => {
      // Ignore errors - we're just triggering the device code prompt
      console.log('Device code flow initiated');
    });

    // Wait for device code to be generated
    await new Promise(resolve => setTimeout(resolve, 2000));

    res.json({
      success: true,
      userCode: serverState.userCode,
      verificationUri: serverState.verificationUri
    });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// Poll for authentication completion
app.get('/api/auth/poll', async (req, res) => {
  try {
    // Check if token has been updated in the graph helper
    // (without triggering a new authentication attempt)
    const cachedToken = graphHelper.getCachedToken();

    if (cachedToken && cachedToken.token && !isAuthenticated()) {
      // Token was acquired, update our state
      serverState.token = cachedToken.token;
      serverState.expiresOn = cachedToken.expiresOn;
      serverState.deviceCode = null;
      serverState.userCode = null;
      saveState(serverState);
      
      // Start auto-refresh timer
      startTokenRefreshTimer();
      console.log('Authentication completed, token auto-refresh started');
    }

    res.json({ authenticated: isAuthenticated() });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// Get current sheet URL (non-sensitive)
app.get('/api/sheet/url', (req, res) => {
  res.json({ 
    sheetUrl: serverState.sheetUrl,
    sheetName: serverState.sheetName || ''
  });
});

// Update sheet URL
app.post('/api/sheet/url', async (req, res) => {
  const { sheetUrl, sheetName } = req.body;

  if (!sheetUrl) {
    return res.status(400).json({ error: 'Sheet URL is required' });
  }

  // Clear cache if URL is different
  const urlChanged = serverState.sheetUrl !== sheetUrl;
  if (urlChanged && serverState.sheetUrl) {
    console.log('Sheet URL changed, clearing cache for old URL');
    await clearCacheForUrl(serverState.sheetUrl);
  }

  serverState.sheetUrl = sheetUrl;
  serverState.sheetName = sheetName || '';
  await saveState(serverState);

  res.json({ 
    success: true, 
    sheetUrl: serverState.sheetUrl,
    sheetName: serverState.sheetName,
    cacheCleared: urlChanged
  });
});

// Get worksheets list
app.get('/api/sheet/worksheets', async (req, res) => {
  try {
    if (!serverState.sheetUrl) {
      return res.status(400).json({ error: 'No sheet URL configured' });
    }

    // Check cache first
    const cachedMetadata = await getCachedMetadata(serverState.sheetUrl);
    if (cachedMetadata) {
      console.log('Returning cached worksheet metadata');
      return res.json({
        fileName: cachedMetadata.fileName,
        worksheets: cachedMetadata.worksheets,
        sheetUrl: serverState.sheetUrl,
        _cached: true,
        _cachedAt: cachedMetadata.timestamp
      });
    }

    if (!isAuthenticated()) {
      return res.status(401).json({ 
        error: 'Authentication required to fetch worksheets list'
      });
    }

    console.log('Fetching fresh worksheets list for URL:', serverState.sheetUrl);
    const worksheetList = await graphHelper.getWorksheetListAsync(serverState.sheetUrl, serverState.token);
    
    // Cache the metadata
    await saveMetadataCache(serverState.sheetUrl, worksheetList);
    
    res.json({
      fileName: worksheetList.fileName,
      worksheets: worksheetList.worksheets,
      sheetUrl: serverState.sheetUrl,
      _cached: false,
      _fetchedAt: new Date().toISOString()
    });
  } catch (err) {
    console.error('Error fetching worksheets:', err.message);
    res.status(500).json({ error: err.message });
  }
});

// Get sheet data
app.get('/api/sheet/data', async (req, res) => {
  const forceRefresh = req.query.refresh === 'true';
  const sheetNameParam = req.query.sheetName || serverState.sheetName || '';
  
  console.log(`📊 API Request - sheetNameParam: "${sheetNameParam}", forceRefresh: ${forceRefresh}`);
  
  // Parse multiple sheet names (comma-separated)
  const sheetNames = sheetNameParam ? sheetNameParam.split(',').map(name => name.trim()).filter(name => name) : [''];
  console.log('📋 Parsed sheet names:', sheetNames);
  
  // Handle empty sheet name case
  if (sheetNames.length === 1 && sheetNames[0] === '') {
    console.log('📋 Using default sheet (empty name)');
  }

  try {
    // Save the sheet names to state if different (save as comma-separated string)
    const sheetNamesString = sheetNames.join(', ');
    if (sheetNamesString !== serverState.sheetName) {
      serverState.sheetName = sheetNamesString;
      await saveState(serverState);
      console.log('Saved sheet names to state:', sheetNamesString);
    }

    // Check cached data for all requested sheets
    const cachedSheets = {};
    const missingSheets = [];
    
    console.log(`🔍 Checking cache for sheets:`, sheetNames);
    for (const sheetName of sheetNames) {
      const cachedData = await getCachedData(serverState.sheetUrl, sheetName);
      if (cachedData && !forceRefresh) {
        cachedSheets[sheetName || 'default'] = cachedData;
        console.log(`✅ Found cached data for sheet: "${sheetName || 'default'}"`);
      } else {
        missingSheets.push(sheetName);
        console.log(`❌ Missing cached data for sheet: "${sheetName || 'default'}"`);
      }
    }
    
    console.log(`📊 Cache summary - Cached: ${Object.keys(cachedSheets).length}, Missing: ${missingSheets.length}`);
    
    // If all sheets are cached and no force refresh, return combined cached data
    if (missingSheets.length === 0 && !forceRefresh) {
      console.log('Returning all cached data instantly');
      const combinedData = combineSheetData(cachedSheets);
      return res.json(combinedData);
    }

    // For refresh requests or missing data, require authentication
    if ((forceRefresh || missingSheets.length > 0) && !isAuthenticated()) {
      console.log(`🔐 Authentication required - Authenticated: ${isAuthenticated()}, Missing sheets: ${missingSheets}`);
      
      // If we have some cached data for the EXACT sheets requested, return it with a message
      if (Object.keys(cachedSheets).length > 0 && missingSheets.length === 0) {
        // All requested sheets are cached, return them
        console.log('Returning cached data for requested sheets (auth required for refresh)');
        const combinedData = combineSheetData(cachedSheets);
        return res.json({
          ...combinedData,
          _message: 'Authentication required to refresh. Showing cached data.'
        });
      }
      
      // If some sheets are missing and we need authentication
      if (missingSheets.length > 0) {
        console.log('Missing sheets and no authentication:', missingSheets);
        return res.status(401).json({ 
          error: `Authentication required to fetch sheet(s): ${missingSheets.join(', ')}`,
          cached: false,
          missingSheets: missingSheets
        });
      }
      
      // Force refresh requires auth
      return res.status(401).json({ 
        error: 'Authentication required to refresh data.',
        cached: false
      });
    }

    // Fetch fresh data for missing sheets or specific sheet if force refresh  
    const sheetsToFetch = forceRefresh ? [sheetNames[0]] : missingSheets; // Only refresh first sheet on force refresh
    console.log('Fetching fresh data from OneDrive for sheets:', sheetsToFetch);
    
    const fetchedSheets = {};
    for (const sheetName of sheetsToFetch) {
      try {
        console.log(`Fetching data for sheet: "${sheetName || 'default'}"`);
        const data = await graphHelper.readExcelFileAsync(serverState.sheetUrl, sheetName, serverState.token);
        const structuredData = dataTransformer.transformVocabData(data);
        
        // Save to cache
        await saveCache(serverState.sheetUrl, structuredData, sheetName);
        await saveStatsCache(serverState.sheetUrl, structuredData, sheetName);
        
        fetchedSheets[sheetName || 'default'] = {
          ...structuredData,
          _cached: false,
          _fetchedAt: new Date().toISOString()
        };
        console.log(`Data cached successfully for sheet: "${sheetName || 'default'}"`);
      } catch (sheetError) {
        console.error(`Error fetching sheet "${sheetName}":`, sheetError.message);
        fetchedSheets[sheetName || 'default'] = {
          error: `Failed to fetch sheet "${sheetName}": ${sheetError.message}`,
          worksheets: []
        };
      }
    }
    
    // Combine fetched data with existing cached data
    const allSheets = { ...cachedSheets, ...fetchedSheets };
    const combinedData = combineSheetData(allSheets);
    
    res.json(combinedData);
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// Helper function to combine data from multiple sheets
function combineSheetData(sheetsData) {
  const combined = {
    worksheets: [],
    _cached: true,
    _fetchedAt: new Date().toISOString(),
    _sheets: Object.keys(sheetsData)
  };
  
  // Check if any sheet was freshly fetched
  let hasAnyFresh = false;
  let oldestCache = null;
  
  for (const [sheetName, data] of Object.entries(sheetsData)) {
    if (data.error) {
      // Add error information
      combined.worksheets.push({
        name: `Error - ${sheetName}`,
        topics: [],
        error: data.error
      });
      continue;
    }
    
    if (!data._cached) {
      hasAnyFresh = true;
    }
    
    if (data._cachedAt) {
      if (!oldestCache || new Date(data._cachedAt) < new Date(oldestCache)) {
        oldestCache = data._cachedAt;
      }
    }
    
    if (data.worksheets && Array.isArray(data.worksheets)) {
      // Add sheet name prefix to worksheet names to distinguish them
      const prefixedWorksheets = data.worksheets.map(worksheet => ({
        ...worksheet,
        name: sheetName ? `[${sheetName}] ${worksheet.name}` : worksheet.name,
        _sheetSource: sheetName || 'default'
      }));
      combined.worksheets.push(...prefixedWorksheets);
    }
  }
  
  // Update metadata
  combined._cached = !hasAnyFresh;
  if (oldestCache && combined._cached) {
    combined._cachedAt = oldestCache;
  }
  
  return combined;
}

// Start server
app.listen(PORT, () => {
  console.log(`Vocab Tracker server running at http://localhost:${PORT}`);
  console.log(`Authenticated: ${isAuthenticated()}`);
});