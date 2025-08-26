# Playwright Excel Online Automation

This project automates **Excel Online** using [Playwright Test](https://playwright.dev/) and verifies that the **`TODAY()` formula** correctly displays the current date in a workbook.

## 📂 Project Structure

```
.
├── config/                 # Stores configuration like credentials & workbook URL
├── functions/
│   ├── login.ts            # Handles Microsoft Office login
│   ├── excel.ts            # Excel workbook navigation & cell interactions
│   ├── screenshot.ts       # Screenshot & OCR (Tesseract.js) for cell validation
├── tests/
│   └── excel-online.test.ts # Main Playwright test for =TODAY()
├── package.json
└── README.md
```

## ⚙️ Requirements

- [Node.js](https://nodejs.org/) (>= 18.x recommended)
- [npm](https://www.npmjs.com/) or [yarn](https://yarnpkg.com/)
- A valid **Microsoft Office 365 account** (with access to Excel Online)

## 📦 Installation

Clone the repository and install dependencies:

```bash
git clone https://github.com/your-repo/playwright-excel-automation.git
cd playwright-excel-automation
npm install
```

Install Playwright browsers:

```bash
npx playwright install
```

## 🔑 Configuration

Create a `config/config.ts` file with your Office 365 credentials and workbook URL:

```ts
export const config = {
  credentials: {
    email: "your-email@domain.com",
    password: "your-password"
  },
  urls: {
    workbook: "https://www.office.com/launch/excel?auth=2"
  }
};
```

⚠️ Do **NOT** commit real credentials into version control. Instead, use environment variables or `.env` files.

## 🧪 Running Tests

Run the Playwright test suite:

```bash
npx playwright test
```

Run in headed mode (see the browser):

```bash
npx playwright test --headed
```

Run a single test file:

```bash
npx playwright test tests/excel-online.test.ts
```

## 🔍 How It Works

1. **Login**  
   - Navigates to Office 365 Excel Online.  
   - Signs in with provided credentials.  
   - Opens a new blank workbook.  

2. **Workbook Operations**  
   - Navigates to the workbook iframe.  
   - Selects cell `A1`.  
   - Types the formula `=TODAY()` and presses Enter.  

3. **Validation**  
   - Takes a screenshot of cell `A1`.  
   - Uses `tesseract.js` OCR to extract the date text.  
   - Compares extracted text with today’s date (`dd/MM/yyyy` format).  

4. **Cleanup**  
   - Undo the changes (`Ctrl+Z`).  
   - Close the browser tab.  

## 📸 Example Output

- **Screenshot:** `cell_A1.png` → Captured image of cell A1 after formula execution.  
- **Logs:** Console output shows progress (login, cell enabled, formula entered, screenshot captured).  

## 🛠️ Troubleshooting

- **Cannot find module `@playwright/test`**  
  → Run `npm install -D @playwright/test`  
- **Tesseract OCR slow or inaccurate**  
  → Try limiting the capture area (already clipped to `75x18` pixels).  
- **Excel Online not loading**  
  → Check your network, VPN, or Microsoft account access.  
