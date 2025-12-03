import js from "@eslint/js";
import globals from "globals";

export default [
    {
        ignores: ["node_modules", "**/app.bundle.js", "**/*.map"]
    },
    js.configs.recommended,
    {
        languageOptions: {
            ecmaVersion: "latest",
            sourceType: "module",
            globals: {
                ...globals.browser,
                ...globals.node,
                Excel: "readonly",
                Office: "readonly",
                OfficeRuntime: "readonly",
                PrairieForge: "writable"
            }
        },
        rules: {
            "no-console": "off",
            "no-unused-vars": [
                "warn",
                {
                    argsIgnorePattern: "^_",
                    varsIgnorePattern: "^_",
                    caughtErrors: "none"
                }
            ]
        }
    }
];
