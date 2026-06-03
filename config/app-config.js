/**
 * Asterwix Unified Web Portal Configuration
 * Save as config/app-config.js. Ensure this file is gitignored in production.
 */

window.ASTERWIX_CONFIG = {
  // Set to true to bypass MSAL/Graph and run entirely with simulated local data
  // When false, the portal will attempt to authenticate with Azure AD and call SharePoint Graph API
  MOCK_MODE_ENABLED: true,

  // Azure AD (Entra ID) Configuration
  auth: {
    clientId: 'YOUR_CLIENT_ID_HERE', // Placeholder
    authority: 'https://login.microsoftonline.com/YOUR_TENANT_ID_HERE', // Placeholder
    redirectUri: 'http://localhost:3000/index.html', // Development local URL or index.html
    productionRedirectUri: 'https://hub.asterwix.com/' // Production URL
  },

  // SharePoint Online Site Configuration
  sharePoint: {
    hostname: 'asterwix.sharepoint.com',
    siteName: 'AsterwixHub',
    siteId: '' // Will be resolved dynamically if empty
  }
};
