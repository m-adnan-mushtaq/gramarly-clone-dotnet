# Security

## Permissions
- **VSTO Add-in**: Runs with Full Trust on the local machine. Requires installation certificate or Trust Center settings to allow local add-ins.
- **Overlay**: Standard user application.

## Data Privacy
- The `SuggestionServer` is local-only for this demo.
- **Warning**: In a real production scenario, sending document text to a remote server requires strict encryption (WSS) and user consent.
- No data is persisted in this demo.

## IPC Security
- Communication between Add-in and Overlay should be secured (e.g., Named Pipes with ACLs) to prevent malicious local processes from injecting fake overlays.
