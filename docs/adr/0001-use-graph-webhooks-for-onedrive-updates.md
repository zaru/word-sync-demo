# Use Graph webhooks for OneDrive updates

Status: superseded by ADR-0005

The demo requires OneDrive changes to be pushed back to the web app rather than detected by polling. We will use Microsoft Graph change notifications for **OneDrive更新通知**, accepting the extra setup cost of a public HTTPS callback and subscription lifecycle because the demo must validate push-based synchronization.
