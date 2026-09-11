# Markdown to DOCX (`markdown-to-docx`)

> **Effortlessly convert Markdown (`.md`) files into professional Microsoft Word (`.docx`) documents with a single click.** Tailored for technical writers and teams who need publication-ready documents featuring company letterheads, executive table styling, GitHub-style alerts, and watermark protection.

---

## Features

- **1-Click Conversion**: Right-click any `.md` file in the Explorer menu or click the Word icon in the editor title bar.
- **Custom Letterhead Support**: Embed company headers (`header.png`) and footers (`footer.png`) automatically.
- **Executive Typography & Tables**: Formatted headings, colored table headers, alternating row shading, and cell padding.
- **GitHub Callout Alerts**: Renders `> [!IMPORTANT]`, `> [!NOTE]`, and `> [!WARNING]` callout boxes with colored left borders.
- **Code Block Containers**: Styled monospaced code blocks with soft grey background shading.
- **Zero External CLI Dependencies**: Runs 100% locally inside VS Code using built-in TypeScript/Node.js document builders.

---

## How to Use

### 1. Convert via File or Editor Context Menu
Right-click any `.md` file in the File Explorer **or right-click anywhere inside an open Markdown document**, and select **"Convert Markdown to DOCX"**.

### 2. Convert via Editor Title Bar
Open any Markdown document and click the **Word icon** in the top-right editor action bar.

### 3. Convert via Command Palette
Press `Cmd+Shift+P` (or `Ctrl+Shift+P`), type `Markdown to DOCX: Convert Current Document`, and press `Enter`.

The generated `.docx` file is saved in the exact same directory as your `.md` file!

---

## Extension Settings

Customize your document styling in VS Code Settings (`markdownToDocx.*`):

| Setting | Type | Default | Description |
| :--- | :--- | :--- | :--- |
| `markdownToDocx.headerImage` | `string` | `""` | Path to header image (`header.png`) for company letterhead. |
| `markdownToDocx.footerImage` | `string` | `""` | Path to footer image (`footer.png`) for company letterhead. |
| `markdownToDocx.watermarkText` | `string` | `"CONFIDENTIAL & PROPRIETARY"` | Background watermark text. |
| `markdownToDocx.accentColor` | `string` | `"B48C3C"` | Hex color code for table headers and callout borders. |
| `markdownToDocx.openAfterConversion` | `boolean` | `true` | Automatically open converted `.docx` file after generation. |

---

## Marketplace Publishing Guide

To package this extension into a `.vsix` file for publishing to the VS Code Marketplace:

```bash
# Package extension into .vsix file
npm run package
```

To publish to the VS Code Marketplace:

```bash
npx @vscode/vsce publish
```

---

## License

MIT License &copy; 2026 PPHLX.
