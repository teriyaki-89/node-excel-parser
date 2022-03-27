module.exports = {
    parser: '@typescript-eslint/parser',
    parserOptions: {
        project: './tsconfig.json',
        sourceType: 'module'
    },
    plugins: ['@typescript-eslint/eslint-plugin'],
    extends: [
        'plugin:@typescript-eslint/recommended',
        //'plugin:prettier/recommended'
    ],
    root: true,
    env: {
        node: true,
        jest: true
    },
    //ignorePatterns: ['.eslintrc.js'],
    rules: {
        // indent: ['error', 2, { "SwitchCase": 1 }],\
        'linebreak-style': ['error', 'windows'],
        'quotes': ['error', 'double'],
        'semi': [2, 'always'],
        'no-trailing-spaces': 'error',
        'object-curly-spacing': ['error', 'always'],
        'brace-style': ['error', '1tbs'],
        'comma-dangle': ['error', 'never'],
        'array-bracket-newline': ['error', { multiline: true, minItems: 8 }],
        yoda: 'error',
        'no-whitespace-before-property': 'error',
        'space-unary-ops': [
            2, {
                'words': true,
                'nonwords': false
            }
        ],
        'no-multiple-empty-lines': ['error', { 'max': 1, 'maxEOF': 0 }]
    }
}