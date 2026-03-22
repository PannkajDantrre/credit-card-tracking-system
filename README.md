# Personal Credit Card Management Web App

A beginner-friendly, single-user credit card dashboard built with:

- HTML, CSS, Vanilla JavaScript
- Google Sheets as the database
- Google Apps Script as the API
- GitHub Pages compatible static hosting

## Files

- `index.html` - Main app UI
- `styles.css` - Responsive styling
- `app.js` - Frontend logic, calculations, sync, and rendering
- `google-apps-script/Code.gs` - Backend API for Google Sheets
- `SHEET_STRUCTURE.md` - Sample sheet layout

## Features Included

- Add transactions with auto reward calculation
- Card management with workbook-style fields like statement date, anniversary date, renewal target, shared group, milestones, and annual benefits
- Best card suggestion engine
- Dashboard with monthly/yearly totals, utilization, reward summary, and category breakdown
- Milestone tracker
- Payment and due alerts
- Anniversary spend tracker
- Reward redemption and expiry tracking
- Advanced reward rule support for category + merchant-based rates
- Local browser storage fallback for testing before Google Sheets setup

## How To Deploy

### 1. Prepare Google Sheets

1. Create a new Google Spreadsheet.
2. Open `Extensions` -> `Apps Script`.
3. Replace the default file with the code from `google-apps-script/Code.gs`.
4. Save the project.
5. In Apps Script, run the `setupSheets` function once.
6. Grant permissions when prompted.

### 2. Deploy the Apps Script Web App

1. In Apps Script, click `Deploy` -> `New deployment`.
2. Choose type `Web app`.
3. Set `Execute as` to `Me`.
4. Set access to `Anyone`.
5. Deploy and copy the Web App URL.

You can test the backend by opening:

```text
YOUR_WEB_APP_URL?action=fetchAll
```

It should return JSON with `cards`, `transactions`, and `rewardActivities`.

### 3. Connect the Frontend

1. Open `index.html` locally or after hosting it.
2. Go to the `Settings` section.
3. Paste your Apps Script Web App URL.
4. Click `Save Settings`.
5. Use `Sync From Google Sheets` to pull remote data anytime.

### 4. Host on GitHub Pages

1. Create a GitHub repository.
2. Upload all files in this project.
3. Push to the `main` branch.
4. In GitHub, open `Settings` -> `Pages`.
5. Under `Build and deployment`, choose:
   - Source: `Deploy from a branch`
   - Branch: `main`
   - Folder: `/ (root)`
6. Save.
7. GitHub will generate your site URL.

## How Reward Rules Work

Enter category-wise reward rates in card setup using JSON.

Example:

```json
{
  "Fuel": 1,
  "Grocery": 2,
  "Travel": 4,
  "Shopping": 5,
  "Others": 1
}
```

For a `₹5,000` shopping transaction with `Shopping: 5`, the app calculates:

```text
5000 x 5 / 100 = 250 reward points
```

## Beginner Notes

- The app saves data locally in your browser even before you configure Google Sheets.
- Once you add the Apps Script URL, new records are also pushed to Google Sheets.
- `Sync From Google Sheets` downloads all sheet data back into the app.
- Reward rules, advanced reward rules, milestones, and annual benefits are kept as JSON so the logic stays flexible without paid services.
- The data model now aligns more closely with Excel-style fields such as statement date, payment status, unbilled amount, bill link, and anniversary tracking.

## Optional Improvements

- Add edit/delete buttons for cards and transactions
- Add charts using a free client-side chart library
- Add card images or bank logos
- Add filters by month, year, or card
