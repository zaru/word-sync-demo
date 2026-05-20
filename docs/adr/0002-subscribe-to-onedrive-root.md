# Subscribe to OneDrive root for update notifications

Status: superseded by ADR-0005

The demo supports both personal and work or school OneDrive accounts, so update subscriptions will be created at the editor's OneDrive root and filtered by active **OneDrive作業コピー** identifiers in the web app. This avoids relying on file- or folder-level subscription behavior that can differ by account type, at the cost of receiving broader notifications than the app ultimately imports.
