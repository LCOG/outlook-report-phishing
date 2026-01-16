import { defineConfig } from "eslint/config";
import officeAddins from "eslint-plugin-office-addins";
import globals from "globals";
import path from "node:path";
import { fileURLToPath } from "node:url";
import js from "@eslint/js";
import { FlatCompat } from "@eslint/eslintrc";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const compat = new FlatCompat({
    baseDirectory: __dirname,
    recommendedConfig: js.configs.recommended,
    allConfig: js.configs.all
});

export default defineConfig([{
    extends: compat.extends("plugin:office-addins/recommended"),

    plugins: {
        "office-addins": officeAddins,
    },

    languageOptions: {
        globals: {
            ...globals.browser,
            ...globals.node,
            Office: "readonly",
            HeadersInit: "readonly",
            BodyInit: "readonly",
            Headers: "readonly",
        },
    },

    rules: {
        "no-undef": "off",
    },
}]);