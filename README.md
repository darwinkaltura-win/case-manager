# SF Case Manager

A local web app for Salesforce customer service teams to view and manage cases, auto-assign JPMC video restore requests, send scheduled email reports, and manage your team roster — all without any npm dependencies.

## Requirements

- **Salesforce CLI** (`sf`) authenticated to your org:
  ```
  sf org login web --alias kaltura
  ```
  Download: https://developer.salesforce.com/tools/salesforcecli

- **Node.js** — the Salesforce CLI ships with its own Node at:
  `C:\Program Files\sf\client\bin\node.exe`
  
  Alternatively install Node.js from https://nodejs.org (v16+)

## Quick Start

1. Clone the repo:
   ```
   git clone https://github.com/darwinkaltura-win/case-manager.git
   cd case-manager
   ```

2. Authenticate Salesforce CLI to your org (if not already):
   ```
   sf org login web --alias kaltura
   ```
   > Change the `SF_ORG` constant in `sf_report_server.js` to match your org alias.

3. Double-click **`launch_sf_report.bat`** (or run `node sf_report_server.js`)

4. Browser opens automatically at **http://localhost:3737**

## First-Time Setup — Team List

Go to the **Settings** tab and search for your team members by name or Salesforce User ID. Add them and click **Save**. This populates the assignee pool for JPMC case assignment and the main case report.

## Tabs

| Tab | Description |
|-----|-------------|
| Open Cases Report | Full team case load by agent — open, escalated, black-flag |
| Case Handling | Agent response activity for the last 35 days |
| Dashboard | Pinnable charts — build your own view |
| JPMC Restore Request | JP Morgan Chase video restore queue with auto-assign |
| Settings | Manage team member list with live Salesforce user search |

## JPMC Auto-Assign

1. Open the **JPMC Restore Request** tab
2. Add team members to the **Assignee Pool** card on the right
3. Click **Assign All** to evenly distribute unassigned cases
4. Enable **Auto-Assign** to distribute automatically on each refresh

## Settings — Team List

- Search Salesforce users by **name** or **18-character User ID**
- Results are pulled live from Salesforce
- Salesforce User ID is shown beside each name for reference
- Saved to `settings.json` (local only, not committed to git)

## Email Reports

- Click the envelope icon on the Open Cases Report tab
- Schedule daily/weekly reports via Windows Task Scheduler (configured automatically)
- Requires PowerShell

## Files

| File | Description |
|------|-------------|
| `sf_report_server.js` | Main server — all HTML/CSS/JS embedded, no build step |
| `launch_sf_report.bat` | Double-click to start |
| `send_sf_report.ps1` | PowerShell email sender |
| `parse_har.py` | Utility for HAR file parsing |
| `settings.json` | Auto-created on first save — stores your team list (gitignored) |

## Customising for Your Org

Edit these constants at the top of `sf_report_server.js`:

```js
const SF_ORG    = 'kaltura';   // your sf org alias
const PORT      = 3737;        // local port
const TEAM_NAMES = [...];      // default team list (overridden by Settings tab)
```

## Notes

- **Zero dependencies** — only Node.js built-in modules (`http`, `fs`, `child_process`)
- All data is fetched live from Salesforce on demand
- `settings.json` and CSV exports are gitignored — nothing sensitive is committed
