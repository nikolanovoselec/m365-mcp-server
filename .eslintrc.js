module.exports = {
  parser: "@typescript-eslint/parser",
  extends: ["eslint:recommended"],
  plugins: ["@typescript-eslint"],
  env: {
    node: true,
    es6: true,
    worker: true,
  },
  parserOptions: {
    ecmaVersion: 2022,
    sourceType: "module",
  },
  globals: {
    DurableObjectNamespace: "readonly",
    KVNamespace: "readonly",
    ExecutionContext: "readonly",
    DurableObjectState: "readonly",
    RequestInit: "readonly",
  },
  rules: {
    "@typescript-eslint/no-unused-vars": [
      "warn",
      { 
        argsIgnorePattern: "^_",
        varsIgnorePattern: "^_"
      }
    ],
    "@typescript-eslint/no-explicit-any": "off",
    "no-unused-vars": "off",
    "no-undef": "error",
    "no-console": "off",
  },
};
