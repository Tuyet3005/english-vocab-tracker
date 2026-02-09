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

// IMPORTANT: This data is now stored in TigrisData cloud storage.
// NEVER expose authentication tokens and credentials to the frontend.
// The tokens, clientSecret, and authentication state must remain server-side only.
// Local file paths (kept for backward compatibility in some functions)
const STATE_FILE = path.join(__dirname, 'server-state.json');
const CACHE_FILE = path.join(__dirname, 'sheet-data-cache.json');
const STATS_CACHE_FILE = path.join(__dirname, 'sheet-stats-cache.json');

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
    const cacheData = {
      sheetUrl,
      sheetName,
      data,
      timestamp: new Date().toISOString()
    };
    await put(CACHE_KEY, JSON.stringify(cacheData, null, 2));
    console.log('Cache saved to Tigris for URL:', sheetUrl, sheetName ? `(Sheet: ${sheetName})` : '');
  } catch (err) {
    console.error('Error saving cache to Tigris:', err.message);
  }
}

async function getCachedData(sheetUrl, sheetName = '') {
  const cache = await loadCache();
  if (cache && cache.sheetUrl === sheetUrl && cache.sheetName === sheetName) {
    console.log('Using cached data from:', cache.timestamp);
    return {
      ...cache.data,
      _cached: true,
      _cachedAt: cache.timestamp
    };
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

    const statsCache = {
      sheetUrl,
      sheetName,
      worksheetStats,
      timestamp: new Date().toISOString()
    };
    await put(STATS_CACHE_KEY, JSON.stringify(statsCache, null, 2));
    console.log('Statistics cache saved to Tigris for URL:', sheetUrl, sheetName ? `(Sheet: ${sheetName})` : '');
  } catch (err) {
    console.error('Error saving stats cache to Tigris:', err.message);
  }
}

function getCachedStats(sheetUrl, sheetName = '') {
  const cache = loadStatsCache();
  if (cache && cache.sheetUrl === sheetUrl && cache.sheetName === sheetName) {
    console.log('Using cached statistics from:', cache.timestamp);
    return cache;
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

// Initialize Graph client
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

  serverState.sheetUrl = sheetUrl;
  serverState.sheetName = sheetName || '';
  await saveState(serverState);

  res.json({ 
    success: true, 
    sheetUrl: serverState.sheetUrl,
    sheetName: serverState.sheetName
  });
});

// Get sheet data
app.get('/api/sheet/data', async (req, res) => {
  const forceRefresh = req.query.refresh === 'true';
  const sheetName = req.query.sheetName || serverState.sheetName || '';

  try {
    // Save the sheet name to state if it's different
    if (sheetName !== serverState.sheetName) {
      serverState.sheetName = sheetName;
      await saveState(serverState);
      console.log('Saved sheet name to state:', sheetName);
    }

    // Always serve cached data when available, regardless of refresh flag or authentication
    const cachedData = await getCachedData(serverState.sheetUrl, sheetName);
    if (cachedData && !forceRefresh) {
      console.log('Returning cached data instantly');
      return res.json(cachedData);
    }

    // For refresh requests, require authentication
    if (forceRefresh && !isAuthenticated()) {
      // If force refresh but no auth, return cached data with a message if available
      if (cachedData) {
        console.log('Returning cached data (refresh requires auth)');
        return res.json({
          ...cachedData,
          _message: 'Authentication required to refresh. Showing cached data.'
        });
      }
      return res.status(401).json({ 
        error: 'Authentication required to refresh data.',
        cached: false
      });
    }

    // For non-refresh requests without exact cached data, try to get any cached data
    if (!forceRefresh && !cachedData) {
      // Try to get any cached data (even with different sheet name)
      const allCachedData = loadCache();
      if (allCachedData && allCachedData.data) {
        console.log('Returning any available cached data instantly');
        return res.json({
          ...allCachedData.data,
          _cached: true,
          _cachedAt: allCachedData.timestamp,
          _message: 'Showing available cached data'
        });
      }
    }

    // If we reach here, we need to fetch fresh data and require authentication
    if (!isAuthenticated()) {
      return res.status(401).json({ 
        error: 'Not authenticated. Please authenticate to fetch data.',
        cached: false
      });
    }

    // Fetch fresh data from OneDrive
    console.log('Fetching fresh data from OneDrive...');
    const data = await graphHelper.readExcelFileAsync(serverState.sheetUrl);
    
    // Filter by sheet name if specified
    let filteredData = data;
    if (sheetName && data.worksheets) {
      filteredData = {
        ...data,
        worksheets: data.worksheets.filter(ws => 
          ws.name.toLowerCase().includes(sheetName.toLowerCase())
        )
      };
      
      if (filteredData.worksheets.length === 0) {
        return res.status(404).json({ 
          error: `Sheet "${sheetName}" not found. Available sheets: ${data.worksheets.map(ws => ws.name).join(', ')}` 
        });
      }
    }

    // Transform data into structured format
    const structuredData = dataTransformer.transformVocabData(filteredData);
    
    // Save to cache (includes both data and statistics)
    await saveCache(serverState.sheetUrl, structuredData, sheetName);
    
    // Also save statistics separately for faster access on subsequent requests
    await saveStatsCache(serverState.sheetUrl, structuredData, sheetName);
    console.log('Data and statistics cached successfully');
    
    res.json({
      ...structuredData,
      _cached: false,
      _fetchedAt: new Date().toISOString()
    });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// Start server
app.listen(PORT, () => {
  console.log(`Vocab Tracker server running at http://localhost:${PORT}`);
  console.log(`Authenticated: ${isAuthenticated()}`);
});