import officeAddins from "eslint-plugin-office-addins";
import globals from "globals";
import js from "@eslint/js";
import tsParser from "@typescript-eslint/parser";
import tseslint from "@typescript-eslint/eslint-plugin";
import importPlugin from "eslint-plugin-import";
import prettierConfig from "eslint-config-prettier";

/** @type {import('eslint').Linter.Config[]} */
export default [
  {
    ignores: [
      "**/node_modules/**",
      "**/dist/**",
      "**/assets/**",
      "**/public/**",
      "**/manifest-configurations/**",
      "webpack.config.js",
      "eslint.config.mjs",
    ],
  },
  js.configs.recommended,
  ...officeAddins.configs.recommended,
  {
    files: ["**/*.ts", "**/*.tsx"],
    languageOptions: {
      parser: tsParser,
      parserOptions: {
        project: "./tsconfig.eslint.json", // Required for type-aware rules
        ecmaVersion: "latest",
        sourceType: "module",
      },
      globals: {
        ...globals.browser,
        ...globals.node,
        Office: "readonly",
        HeadersInit: "readonly",
        BodyInit: "readonly",
        Headers: "readonly",
      },
    },
    plugins: {
      "@typescript-eslint": tseslint,
      import: importPlugin,
    },
    rules: {
      // Disable base ESLint rules that conflict with TypeScript
      "no-undef": "off",
      "no-unused-vars": "off",

      // TypeScript-specific rules
      "@typescript-eslint/no-explicit-any": "error",
      "@typescript-eslint/explicit-function-return-type": [
        "warn",
        {
          allowExpressions: true,
          allowTypedFunctionExpressions: true,
        },
      ],
      "@typescript-eslint/no-unused-vars": [
        "error",
        {
          argsIgnorePattern: "^_",
          varsIgnorePattern: "^_",
        },
      ],

      // Async/Promise safety (type-aware)
      "@typescript-eslint/no-floating-promises": "error",
      "@typescript-eslint/await-thenable": "error",
      "@typescript-eslint/no-misused-promises": [
        "error",
        {
          checksVoidReturn: false, // Often needed for event handlers
        },
      ],
      "@typescript-eslint/promise-function-async": "warn",

      // Type safety
      "@typescript-eslint/strict-boolean-expressions": [
        "warn",
        {
          allowNullableObject: true,
        },
      ],
      "@typescript-eslint/no-unnecessary-condition": "warn",
      "@typescript-eslint/no-unnecessary-type-assertion": "warn",

      // Code quality
      "@typescript-eslint/no-redundant-type-constituents": "warn",
      "@typescript-eslint/prefer-nullish-coalescing": "warn",
      "@typescript-eslint/prefer-optional-chain": "warn",
      "@typescript-eslint/consistent-type-imports": [
        "warn",
        {
          prefer: "type-imports",
        },
      ],

      // Office Add-in specific: prevent common mistakes
      "no-console": ["warn", { allow: ["warn", "error"] }],
      eqeqeq: ["error", "always"],

      // Disable rules that conflict with Prettier
      ...prettierConfig.rules,

      // Imports
      "import/order": [
        "error",
        {
          groups: [
            "builtin",
            "external",
            "internal",
            ["parent", "sibling"],
            "index",
            "object",
            "type",
          ],
          "newlines-between": "always",
          alphabetize: {
            order: "asc",
            caseInsensitive: true,
          },
        },
      ],
      "import/no-duplicates": "error",
    },
  },
];
