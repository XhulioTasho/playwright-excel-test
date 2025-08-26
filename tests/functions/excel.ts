import { Page, FrameLocator } from '@playwright/test';
import { config } from '../config';

// transition to Excel Book
export async function goToWorkbook(page: Page): Promise<void> {
  await page.goto(config.urls.workbook);
  
  // Wait for Excel Online to fully load
  const frame = page.frameLocator('iframe');
  try {
    await frame.locator('#Sheet0_0_0_1 div canvas').waitFor({ timeout: 10000 });
    console.log("Excel Online is loaded");
  } catch (error) {
    throw new Error("Excel Online is not loaded");
  }
}

// click on A1 using cell coordinates
export async function enableCellA1(page: Page, frame: FrameLocator) {
  const canvas = frame.locator('#Sheet0_0_0_1 canvas');
  await canvas.click({ position: { x: 10, y: 10 } });
  
  console.log("Cell A1 is enabled");
}

// enter the formula into an active cell
export async function enterFormula(page: Page, formula: string) {
  await page.keyboard.type(formula, { delay: 500 });
  await page.keyboard.press('Enter');
  await page.waitForTimeout(2000);
  
  console.log("Formula should be entered");
}
