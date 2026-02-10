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

// Helper function to get bearer token for API calls
async function getBearerToken(serverToken = null) {
  // If server token is provided, use it (for server-side calls)
  if (serverToken) {
    return serverToken;
  }
  
  if (!_cachedToken || !_cachedToken.token) {
    throw new Error('No cached token available');
  }
  
  // Check if token is still valid
  if (_cachedToken.expiresOn) {
    const expiresOn = new Date(_cachedToken.expiresOn);
    if (expiresOn <= new Date(Date.now() + 5 * 60 * 1000)) {
      throw new Error('Token expired or expires soon');
    }
  }
  
  return _cachedToken.token;
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
async function readExcelFileAsync(sharingUrl, sheetName = '', serverToken = null) {
  // Convert sharing URL to sharing token
  const base64Value = Buffer.from(sharingUrl).toString('base64');
  const encodedUrl = 'u!' + base64Value.replace(/==$/, '').replaceAll('/', '_').replaceAll('+', '-');

  const token = await getBearerToken(serverToken);
  const headers = {
    'Authorization': `Bearer ${token}`,
    'Content-Type': 'application/json'
  };

  // Get the shared item metadata
  console.log("Fetching shared item metadata...");
  const sharedItemResponse = await fetch(`https://graph.microsoft.com/v1.0/shares/${encodedUrl}/driveItem`, {
    headers
  });
  
  if (!sharedItemResponse.ok) {
    throw new Error(`Failed to fetch shared item: ${sharedItemResponse.statusText}`);
  }
  
  const sharedItem = await sharedItemResponse.json();
  console.log("Shared item:", sharedItem.name, `(${sharedItem.id})`);

  // Extract drive ID and item ID
  const driveId = sharedItem.parentReference.driveId;
  const itemId = sharedItem.id;

  const result = {
    fileName: sharedItem.name,
    fileSize: sharedItem.size,
    worksheets: []
  };

  if (sheetName) {
    // Get specific worksheet by name
    try {
      const worksheetsResponse = await fetch(`https://graph.microsoft.com/v1.0/drives/${driveId}/items/${itemId}/workbook/worksheets`, {
        headers
      });
      
      if (!worksheetsResponse.ok) {
        throw new Error(`Failed to fetch worksheets: ${worksheetsResponse.statusText}`);
      }
      
      const worksheetsData = await worksheetsResponse.json();
      console.log("Found worksheets:", worksheetsData.value.map(ws => ws.name));
      
      const targetWorksheet = worksheetsData.value.find(ws => 
        ws.name.toLowerCase().includes(sheetName.toLowerCase())
      );
      
      if (!targetWorksheet) {
        const availableSheets = worksheetsData.value.map(ws => ws.name).join(', ');
        throw new Error(`Sheet "${sheetName}" not found. Available sheets: ${availableSheets}`);
      }

      // Read data from the specific worksheet
      try {
        const rangeResponse = await fetch(`https://graph.microsoft.com/v1.0/drives/${driveId}/items/${itemId}/workbook/worksheets/${targetWorksheet.id}/usedRange?$select=address,rowCount,columnCount,values&valuesOnly=true`, {
          headers
        });
        
        if (!rangeResponse.ok) {
          throw new Error(`Failed to fetch range data: ${rangeResponse.statusText}`);
        }
        
        const range = await rangeResponse.json();

        result.worksheets.push({
          name: targetWorksheet.name,
          range: range.address,
          rowCount: range.rowCount,
          columnCount: range.columnCount,
          values: range.values
        });
      } catch (err) {
        result.worksheets.push({
          name: targetWorksheet.name,
          error: err.message
        });
      }
    } catch (err) {
      throw new Error(`Failed to fetch worksheet "${sheetName}": ${err.message}`);
    }
  } else {
    // Get all worksheets (original behavior)
    const worksheetsResponse = await fetch(`https://graph.microsoft.com/v1.0/drives/${driveId}/items/${itemId}/workbook/worksheets`, {
      headers
    });
    
    if (!worksheetsResponse.ok) {
      throw new Error(`Failed to fetch worksheets: ${worksheetsResponse.statusText}`);
    }
    
    const worksheetsData = await worksheetsResponse.json();
    console.log("Found worksheets:", worksheetsData.value.map(ws => ws.name));

    // Read data from each worksheet
    for (const worksheet of worksheetsData.value) {
      try {
        const rangeResponse = await fetch(`https://graph.microsoft.com/v1.0/drives/${driveId}/items/${itemId}/workbook/worksheets/${worksheet.id}/usedRange?$select=address,rowCount,columnCount,values`, {
          headers
        });
        
        if (!rangeResponse.ok) {
          throw new Error(`Failed to fetch range data: ${rangeResponse.statusText}`);
        }
        
        const range = await rangeResponse.json();

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
  }

  return result;
}

// Get list of worksheet names only (lightweight)
async function getWorksheetListAsync(sharingUrl, serverToken = null) {
  // Convert sharing URL to sharing token
  const base64Value = Buffer.from(sharingUrl).toString('base64');
  const encodedUrl = 'u!' + base64Value.replace(/=+$/g, '').replaceAll('/', '_').replaceAll('+', '-');

  const token = await getBearerToken(serverToken);
  const headers = {
    'Authorization': `Bearer ${token}`,
    'Content-Type': 'application/json'
  };

  // Get the shared item metadata
  console.log('Fetching worksheet list...');
  const sharedItemResponse = await fetch(`https://graph.microsoft.com/v1.0/shares/${encodedUrl}/driveItem`, {
    headers
  });
  
  if (!sharedItemResponse.ok) {
    throw new Error(`Failed to fetch shared item: ${sharedItemResponse.statusText}`);
  }
  
  const sharedItem = await sharedItemResponse.json();
  
  // Extract drive ID and item ID
  const driveId = sharedItem.parentReference.driveId;
  const itemId = sharedItem.id;

  // Get worksheets metadata only
  const worksheetsResponse = await fetch(`https://graph.microsoft.com/v1.0/drives/${driveId}/items/${itemId}/workbook/worksheets?$select=id,name,position`, {
    headers
  });
  
  if (!worksheetsResponse.ok) {
    throw new Error(`Failed to fetch worksheets: ${worksheetsResponse.statusText}`);
  }
  
  const worksheetsData = await worksheetsResponse.json();
  console.log('Found worksheets:', worksheetsData.value.map(ws => ws.name));
  
  return {
    fileName: sharedItem.name,
    worksheets: worksheetsData.value.map(ws => ({
      id: ws.id,
      name: ws.name,
      position: ws.position
    })).sort((a, b) => a.position - b.position)
  };
}

module.exports = {
  initializeGraphForUserAuth,
  getCachedToken,
  getBearerToken,
  getUserAsync,
  readExcelFileAsync,
  getWorksheetListAsync
};
