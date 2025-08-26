import { Page } from '@playwright/test';
import { config } from '../config';

export async function loginToOffice(page: Page): Promise<void> {
  // 1. Sign in
  await page.goto('https://www.office.com/launch/excel');
  await page.click('text=Sign in');

  await page.fill('input[name="loginfmt"]', config.credentials.email);
  await page.click('input[type="submit"]');
  await page.click('text=Work or school account');

  await page.fill('input[name="passwd"]', config.credentials.password);
  await page.click('input[type="submit"]');

  await page.click('text=No');

  // 2. Open new workbook
  await page.click('text=Blank workbook');
}