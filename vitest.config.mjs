import { defineConfig } from 'vitest/config';

// Unit tests: Office.js (test/fake-office.ts) and Converter Studio's API (test/fake-api.ts) are
// stood in for; the pane's tests run in jsdom.
export default defineConfig({
  test: {
    include: ['test/**/*.test.ts'],
    environment: 'node',
  },
});
