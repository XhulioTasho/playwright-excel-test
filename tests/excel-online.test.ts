import { test, expect } from '@playwright/test';
// Importing helper functions for login and Excel operations
import { loginToOffice } from './functions/login';
import { goToWorkbook, enableCellA1, enterFormula } from './functions/excel';
import { getScreenshot, getText } from './functions/screenshot';
import { format } from 'date-fns';

// Grouping related tests under a common test suite name
test.describe('Excel Online TEST FORMULA TODAY()', () => {

  // Before each test: log in and navigate to the Excel workbook
  test.beforeEach(async ({ page }) => {
    await loginToOffice(page);               // Log in to Office 365
    await goToWorkbook(page);          // Open the Excel workbook
  });

  // After each test: undo the changes and close the browser tab
  test.afterEach(async ({ page }) => {
    await page.keyboard.press('Control+Z');  // Undo any cell changes made during the test
    console.log("Undo the changes");

    await page.close();                      // Close the browser tab
    console.log("Close the browser");
  });

  // Main test: verify that the TODAY() function works correctly in Excel Online
  test('check TODAY() function', async ({ page }) => {
    const frame = page.frameLocator('iframe');  // Locate the iframe that contains the Excel app

    await enableCellA1(page, frame);            // Select cell A1 to input the formula
    await enterFormula(page, "=TODAY()");       // Enter the TODAY() formula in cell A1
    await getScreenshot(page, frame);           // Take a screenshot for visual validation (optional)

    const cellText = await getText();           // Extract the value displayed in cell A1
    const today = format(new Date(), 'dd/MM/yyyy');  // Format today's date for comparison

    // Validate that the cell contains today's date
    expect(cellText).toContain(today);
  });
});
