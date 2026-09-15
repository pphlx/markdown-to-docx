# Change Log

All notable changes to the "Markdown to DOCX" extension will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [1.0.1] - 2026-09-15

### Added
- **LaTeX Math Equation Support**:
  - Full native Microsoft Word Equation (OMML - Office Math Markup Language) rendering via Word's built-in Equation Tools.
  - **Display Math Blocks**: Centered equations with professional vertical spacing for `$$...$$`, `\[...\]`, and `\begin{equation}...\end{equation}` blocks.
  - **Math Code Blocks**: Support for ```` ```math ```` and ```` ```latex ```` markdown code blocks.
  - **Inline Equations**: Native inline equation rendering (`$...$` and `$$...$$`) within paragraphs, bullet points, and table cells.
  - Native handling of fractions (`\frac{...}{...}`), superscripts, subscripts, Greek symbols, matrices, and normal text (`\text{...}`).
  - Unicode character and currency symbol support within equations (e.g., `₹`, `$`, `€`).
- **Documentation**: Added `CHANGELOG.md` documenting release history.

### Changed
- Improved inline formatting tokenizer to intelligently differentiate between inline math formulas and currency amounts (e.g., `$50`).
- Priority given to code spans to protect code variables inside backticks from accidental math translation.

---

## [1.0.0] - 2026-09-11

### Added
- **Initial Release of Markdown to DOCX Converter**:
  - Convert any `.md` file to a beautifully styled `.docx` Microsoft Word document with 1-click.
  - **Context Menus**: Right-click support in both the File Explorer tree (`explorer/context`) and inside opened Markdown documents (`editor/context`).
  - **Editor Title Action**: Dedicated Word icon button in the editor title bar.
  - **Letterhead Support**: Configurable top header (`header.png`) and bottom footer (`footer.png`) image banners.
  - **Watermark Support**: Customizable diagonal background watermark (e.g., "CONFIDENTIAL & PROPRIETARY").
  - **Executive Tables**: Beautiful styled tables with colored header rows, subtle borders, and alternating zebra stripes.
  - **GitHub-style Callout Alerts**: Supported `[!NOTE]`, `[!TIP]`, `[!IMPORTANT]`, `[!WARNING]`, and `[!CAUTION]` boxes with accent border highlights.
  - **Mermaid Diagrams**: Automatic rendering of Mermaid diagrams into high-resolution embedded images.
  - **Rich Typography**: Proper heading hierarchies, list indentation, strikethrough, bold, italic, and HTML entity support (`&nbsp;`, `<br>`).
  - **Zero External CLI Dependencies**: 100% local conversion without requiring Pandoc or external software.
