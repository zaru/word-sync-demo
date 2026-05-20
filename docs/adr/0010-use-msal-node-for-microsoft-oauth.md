# Use MSAL Node for Microsoft OAuth

The demo will implement Microsoft OAuth with MSAL Node instead of Auth.js, using the `common` authority and storing each editor's MSAL token cache in SQLite rather than handling refresh tokens directly. This adds custom session and token-cache code, but gives direct control over personal plus work or school Microsoft account support, `Files.ReadWrite.AppFolder`, `offline_access`, and Graph token refresh during long **Word編集セッション**.
