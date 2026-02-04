require('@rushstack/eslint-patch/modern-module-resolution');

module.exports = {
  root: true, // This prevents looking up parent directories
  env: {
    browser: true,
    es6: true,
    node: true
  },
  plugins: ['@typescript-eslint', 'react-hooks'],
  parserOptions: {
    ecmaVersion: 2018,
    sourceType: 'module'
  },
  rules: {
    // Minimal rules to avoid any plugin conflicts
    'no-unused-vars': 'off',
    'no-undef': 'off'
  },
  overrides: [
    {
      files: ['*.ts', '*.tsx'],
      parser: '@typescript-eslint/parser',
      parserOptions: {
        tsconfigRootDir: __dirname,
        project: './tsconfig.json'
      },
      rules: {}
    }
  ]
};
