const js = require('@eslint/js');
const globals = require('globals');

// Globals that the browser scripts share (loaded by index.html in this order)
const appGlobals = {
  XLSX: 'readonly',
  Diff: 'readonly',
  I18n: 'readonly',
  ComparatorEngine: 'readonly',
  SampleGenerator: 'readonly',
  StorageManager: 'readonly',
  Importers: 'readonly',
  Exporters: 'readonly'
};

module.exports = [
  { ignores: ['lib/**', 'node_modules/**'] },
  js.configs.recommended,
  {
    // Universal modules: run in the browser and in Node.js (tests)
    files: ['js/core/**/*.js', 'js/io/**/*.js'],
    languageOptions: {
      ecmaVersion: 2022,
      sourceType: 'script',
      globals: Object.assign({}, globals.browser, globals.commonjs, { define: 'readonly', Buffer: 'readonly' }, appGlobals)
    }
  },
  {
    files: ['js/app/**/*.js'],
    languageOptions: {
      ecmaVersion: 2022,
      sourceType: 'script',
      globals: Object.assign({}, globals.browser, appGlobals)
    }
  },
  {
    files: ['test/**/*.js', 'sample/**/*.js', 'serve.js', 'eslint.config.js'],
    languageOptions: {
      ecmaVersion: 2022,
      sourceType: 'commonjs',
      globals: globals.node
    }
  },
  {
    rules: {
      'no-unused-vars': ['error', { args: 'none', caughtErrors: 'none' }],
      eqeqeq: ['error', 'smart'],
      'no-var': 'error',
      'prefer-const': 'error'
    }
  }
];
