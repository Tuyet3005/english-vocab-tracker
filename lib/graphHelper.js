// Microsoft Graph Helper
require('isomorphic-fetch');
const azure = require('@azure/identity');
const graph = require('@microsoft/microsoft-graph-client');
const authProviders = require('@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials');

let _settings = undefined;
let _deviceCodeCredential = undefined;
let _userClient = undefined;
let _cachedToken = null;

// Initialize Graph with token cache
function initializeGraphForUserAuth(settings, cachedToken, deviceCodePrompt) {
  if (!settings) {
    throw new Error('Settings cannot be undefined');
  }

  _settings = settings;
  _cachedToken = cachedToken;

  _deviceCodeCredential = new azure.DeviceCodeCredential({
    clientId: settings.clientId,
    tenantId: settings.tenantId,
    userPromptCallback: deviceCodePrompt
  });

  // Custom credential wrapper that uses cached token
  const cachedCredential = {
    getToken: async (scopes) => {
      // Check if cached token is still valid
      if (_cachedToken && _cachedToken.expiresOn) {
        const expiresOn = new Date(_cachedToken.expiresOn);
        if (expiresOn > new Date(Date.now() + 5 * 60 * 1000)) {
          return {
            token: _cachedToken.token,
            expiresOnTimestamp: new Date(_cachedToken.expiresOn).getTime()
          };
        }
      }

      // Get new token
      const response = await _deviceCodeCredential.getToken(scopes);
      _cachedToken = {
        token: response.token,
        expiresOn: response.expiresOnTimestamp
      };
      return response;
    }
  };

  const authProvider = new authProviders.TokenCredentialAuthenticationProvider(
    cachedCredential, {
      scopes: settings.graphUserScopes
    });

  _userClient = graph.Client.initWithMiddleware({
    authProvider: authProvider
  });
}

// Get current cached token
function getCachedToken() {
  return _cachedToken;
}

// Get user information
async function getUserAsync() {
  if (!_userClient) {
    throw new Error('Graph has not been initialized for user auth');
  }

  return _userClient.api('/me')
    .select(['displayName', 'mail', 'userPrincipalName'])
    .get();
}

// Read Excel file from SharePoint/OneDrive sharing link
async function readExcelFileAsync(sharingUrl) {
  if (!_userClient) {
    throw new Error('Graph has not been initialized for user auth');
  }

  // Convert sharing URL to sharing token
  const base64Value = Buffer.from(sharingUrl).toString('base64');
  const encodedUrl = 'u!' + base64Value.replace(/=+$/, '').replace(/\//g, '_').replace(/\+/g, '-');

  // Get the shared item metadata
  const sharedItem = await _userClient.api(`/shares/${encodedUrl}/driveItem`).get();

  // Extract drive ID and item ID
  const driveId = sharedItem.parentReference.driveId;
  const itemId = sharedItem.id;

  // Get worksheets
  const worksheets = await _userClient.api(`/drives/${driveId}/items/${itemId}/workbook/worksheets`).get();

  const result = {
    fileName: sharedItem.name,
    fileSize: sharedItem.size,
    worksheets: []
  };

  // Read data from each worksheet
  for (const worksheet of worksheets.value) {
    try {
      const range = await _userClient.api(`/drives/${driveId}/items/${itemId}/workbook/worksheets/${worksheet.id}/usedRange`).get();

      result.worksheets.push({
        name: worksheet.name,
        range: range.address,
        rowCount: range.rowCount,
        columnCount: range.columnCount,
        values: range.values
      });
    } catch (err) {
      result.worksheets.push({
        name: worksheet.name,
        error: err.message
      });
    }
  }

  return result;
}

module.exports = {
  initializeGraphForUserAuth,
  getCachedToken,
  getUserAsync,
  readExcelFileAsync
};
