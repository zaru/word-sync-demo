# Import on explicit session finish

The demo will not subscribe to OneDrive change notifications; instead, the editor uses an explicit finish action to trigger **終了取り込み** from the latest **OneDrive作業コピー**. This removes the need for public webhook infrastructure and broad root subscriptions, accepting that Web-side updates happen when the editor finishes rather than immediately on each Word save.
