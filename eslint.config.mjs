import js from '@eslint/js';

export default [
  {
    files: ['src/*.gs'],
    languageOptions: {
      ecmaVersion: 2022,
      sourceType: 'script'
    },
    rules: {
      ...js.configs.recommended.rules,
      // Apps Script combines all .gs files into one global script at runtime.
      'no-undef': 'off',
      'no-unused-vars': 'off'
    }
  }
];
