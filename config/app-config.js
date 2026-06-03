/**
 * Asterwix Unified Web Portal — Configuration File
 * =====================================================================
 * IMPORTANT: Ensure this file is listed in .gitignore for security.
 * =====================================================================
 */

window.ASTERWIX_CONFIG = {

  // Live Microsoft 365 authentication is now ENABLED
  MOCK_MODE_ENABLED: false,

  // Azure AD (Entra ID) — Asterwix Portal App Registration
  auth: {
    clientId: 'bf5114ed-db8d-4137-b2a3-2d4ea4cd25e8',
    
    // Explicitly target the asterwix007 tenant to avoid login cross-tenant mismatch
    authority: 'https://login.microsoftonline.com/asterwix007.onmicrosoft.com',
    
    redirectUri: window.location.origin + window.location.pathname,
    productionRedirectUri: 'https://hub.asterwix.com/'
  },

  // SharePoint Online Site Configuration
  sharePoint: {
    hostname: 'asterwix007.sharepoint.com',
    siteName: 'AsterwixBIMHub',
    siteId: '' // Resolved dynamically if left empty
  }
};
