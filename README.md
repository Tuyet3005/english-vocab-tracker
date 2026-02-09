# 🌼Vocab tracker 🌼 --- Tuyết

A web-based vocabulary tracker that integrates with Microsoft Graph to read Excel files from SharePoint/OneDrive.

## Features

- Microsoft Graph authentication with device code flow
- Token caching for persistent authentication
- Read Excel files from SharePoint/OneDrive sharing links
- Simple web interface to view and manage vocabulary data

## Setup

1. Install dependencies:
   ```bash
   npm install
   ```

2. Run the development server with auto-reload:
   ```bash
   npm run dev
   ```

3. Or run in production mode:
   ```bash
   npm start
   ```

4. Open your browser to `http://localhost:3000`

## Authentication Flow

1. When you first visit the app, you'll be redirected to the authentication page
2. Click the button to open Microsoft's device login page
3. Enter the provided code
4. Once authenticated, you'll be redirected back to the home page
5. The authentication token is cached and will persist across server restarts

## Configuration

The app settings (Client ID, scopes, etc.) are configured in `server.js`.

The server state (including authentication tokens and sheet URL) is stored in `server-state.json` - this file is automatically created with default values if it doesn't exist.

**IMPORTANT:** Never commit `server-state.json` to version control as it contains sensitive authentication data. This file is already in `.gitignore`.

## API Endpoints

- `GET /api/auth/status` - Check authentication status
- `POST /api/auth/start` - Start authentication process
- `GET /api/auth/poll` - Poll for authentication completion
- `GET /api/sheet/url` - Get current sheet URL
- `POST /api/sheet/url` - Update sheet URL
- `GET /api/sheet/data` - Get sheet data (requires authentication)

## Security Notes

- Authentication tokens and credentials are stored server-side only
- The frontend never receives sensitive authentication data
- Only non-sensitive data (like sheet URL) is exposed to the frontend
