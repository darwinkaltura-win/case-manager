# SF Case Manager

Internal tool for the Kaltura Customer Service team to view and manage Salesforce cases, auto-assign JPMC restore requests, and send scheduled email reports.

## Requirements

- **Node.js** — uses the Salesforce CLI's bundled Node:
  `C:\Program Files\sf\client\bin\node.exe`
  
  If you don't have Salesforce CLI, install Node.js from https://nodejs.org

- **Salesforce CLI** (`sf`) — must be authenticated to the `kaltura` org:
  ```
  sf org login web --alias kaltura
  ```

## Installation

1. Clone the repo:
   ```
   git clone <repo-url>
   cd sf-case-manager
   ```

2. Authenticate Salesforce CLI to the kaltura org (if not already):
   ```
   sf org login web --alias kaltura
   ```

3. Double-click **`launch_sf_report.bat`** to start the server.

4. The browser will open automatically at **http://localhost:3737**

## Usage

| Tab | Description |
|-----|-------------|
| Dashboard | Team case metrics — open, escalated, black flags |
| Report | Detailed case handling report with filters |
| Case Handling | Agent activity and response tracking |
| JPMC | JP Morgan Chase restore request queue + auto-assign |

### JPMC Auto-Assign
- Check the **Auto-Assign** checkbox to automatically distribute new unassigned restore request cases across selected team members
- Use the assign pool checkboxes to control who receives cases

### Email Reports
- Configure recipient and schedule under the email icon
- Requires PowerShell and Windows Task Scheduler (handled automatically)

## Files

| File | Description |
|------|-------------|
| `sf_report_server.js` | Main Node.js server (no npm install needed) |
| `launch_sf_report.bat` | Double-click launcher |
| `send_sf_report.ps1` | PowerShell script for scheduled email reports |
| `parse_har.py` | Utility script for HAR file parsing |

## Notes

- No npm dependencies — uses only Node.js built-in modules
- Data is pulled live from Salesforce on each page load
- CSV exports are excluded from this repo (generated at runtime)
