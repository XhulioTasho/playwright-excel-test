import * as dotenv from 'dotenv';
dotenv.config();

export interface TestConfig {
  credentials: {
    email: string;
    password: string;
  };
  urls: {
    login: string;
    excelOnline: string;
    workbook: string;
  };
  timeouts: {
    navigation: number;
    element: number;
  };
}

export const config: TestConfig = {
  credentials: {
    email: process.env.EXCEL_EMAIL || "example@email.com",
    password: process.env.EXCEL_PASSWORD || "password",
  },
  urls: {
    login: "https://www.office.com/login",
    excelOnline: "https://www.office.com/launch/excel",
    workbook: "https://1drv.ms/x/s!AgY7gIampfG6cmGCC225Aa6ORdU?e=JkoNU2",
  },
  timeouts: {
    navigation: 30000,
    element: 5000,
  },
};