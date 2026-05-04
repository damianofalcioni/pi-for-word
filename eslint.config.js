import js from "@eslint/js";
import importPlugin from "eslint-plugin-import";
import sonarjs from "eslint-plugin-sonarjs";
import globals from "globals";

export default [
  {
    ignores: [
      "**/node_modules/**",
      "dist/**",
      "build/**",
      "coverage/**",
      "public/**"
    ]
  },
  js.configs.recommended,

  {
    files: ["scripts/**/*.{js,mjs,cjs}", "eslint.config.js"],
    languageOptions: {
      ecmaVersion: "latest",
      sourceType: "module",
      globals: {
        ...globals.nodeBuiltin
      }
    }
  },

  {
    files: ["src/**/*.{js,mjs,cjs}"],
    plugins: {
      sonarjs,
      import: importPlugin
    },
    settings: {
      "import/resolver": {
        node: true
      }
    },
    languageOptions: {
      ecmaVersion: "latest",
      sourceType: "module",
      globals: {
        ...globals.browser,
        Office: false,
        Word: false
      }
    },
    rules: {
      // Complexity: main AI-agent refactoring signal
      "complexity": ["warn", 8],
      "sonarjs/cognitive-complexity": ["warn", 10],
      "max-depth": ["warn", 3],
      "max-lines-per-function": ["warn", {
        max: 60,
        skipBlankLines: true,
        skipComments: true
      }],
      "max-statements": ["warn", 20],
      "max-params": ["warn", 4],

      // Control-flow clarity
      "no-else-return": "warn",
      "no-nested-ternary": "warn",
      "consistent-return": "warn",
      "no-unreachable": "error",

      // Explicitness
      "eqeqeq": ["warn", "always"],
      "no-implicit-coercion": "warn",
      "no-unused-vars": ["warn", {
        argsIgnorePattern: "^_",
        varsIgnorePattern: "^_"
      }],
      "no-use-before-define": ["warn", {
        functions: false,
        classes: true,
        variables: true
      }],
      "no-shadow": "warn",

      // Lower hidden state / side effects
      "prefer-const": "warn",
      "no-var": "error",
      "no-param-reassign": "warn",

      // Module structure
      "max-lines": ["warn", {
        max: 400,
        skipBlankLines: true,
        skipComments: true
      }],
      "max-classes-per-file": ["warn", 1],
      "import/no-cycle": ["warn", {
        maxDepth: 5,
        ignoreExternal: true
      }],

      // Duplication / suspicious code
      "sonarjs/no-identical-functions": "warn",
      "sonarjs/no-duplicated-branches": "warn",
      "sonarjs/no-small-switch": "warn"
    }
  }
];
