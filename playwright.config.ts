import path from "node:path";

import { defineConfig } from "@playwright/test";

export default defineConfig({
  testDir: path.resolve(__dirname, "tests/e2e"),
  timeout: 60_000,
  use: {
    baseURL: "http://127.0.0.1:3010",
    trace: "retain-on-failure",
  },
  webServer: {
    command: "npm run dev -- --port 3010",
    url: "http://127.0.0.1:3010",
    reuseExistingServer: !process.env.CI,
    timeout: 120_000,
  },
});
