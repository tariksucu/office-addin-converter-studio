import { defineConfig } from 'eslint/config';
import officeAddins from 'eslint-plugin-office-addins';
import globals from 'globals';
import tseslint from 'typescript-eslint';

// Microsoft's lint rules for Office add-ins (eslint-plugin-office-addins, its "recommended"
// config: Office.js load/sync rules, no Office.initialize, prettier), with typescript-eslint's
// type-checked rules on the TypeScript.
export default defineConfig(
  {
    // tools/manifest.d.mts: type declarations only (checked by tsc).
    ignores: ['dist/**', 'node_modules/**', 'coverage/**', 'test-results/**', 'tools/*.d.mts'],
  },
  ...officeAddins.configs.recommended,
  {
    files: ['**/*.ts'],
    extends: [tseslint.configs.recommendedTypeChecked],
    languageOptions: {
      parserOptions: { projectService: true, tsconfigRootDir: import.meta.dirname },
    },
  },
  {
    files: ['src/**/*.ts', 'src/**/*.js', 'test/**/*.ts', 'e2e/**/*.ts', 'dev/**/*.ts'],
    languageOptions: { globals: { ...globals.browser } },
  },
  {
    // The ES5 page for webviews too old for the pane.
    files: ['src/legacy.js'],
    languageOptions: { ecmaVersion: 5, sourceType: 'script' },
  },
  {
    files: ['tools/**/*.mjs', 'dev/**/*.mjs', 'e2e/**/*.mjs', '*.mjs'],
    languageOptions: { globals: { ...globals.node }, ecmaVersion: 2024 },
  },
  {
    // The browser run's checks are partly functions run inside the page (page.evaluate).
    files: ['e2e/**/*.mjs'],
    languageOptions: { globals: { ...globals.node, ...globals.browser } },
  },
  {
    files: ['test/**/*.ts'],
    rules: {
      // Fakes stand in for Office's typed objects.
      '@typescript-eslint/no-unsafe-assignment': 'off',
      '@typescript-eslint/no-unsafe-member-access': 'off',
      '@typescript-eslint/no-unsafe-call': 'off',
      '@typescript-eslint/no-unsafe-argument': 'off',
      '@typescript-eslint/unbound-method': 'off',
    },
  },
);
