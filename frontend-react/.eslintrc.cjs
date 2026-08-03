module.exports = {
  root: true,
  extends: ['eslint:recommended'],
  ignorePatterns: ['*.css', '*.scss', '*.sass'],
  rules: {
    'no-unused-vars': 'warn',
  },
  overrides: [
    {
      files: ['*.css', '*.scss', '*.sass'],
      rules: {
        'all': 'off',
      },
    },
  ],
}
