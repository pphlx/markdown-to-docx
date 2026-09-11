import * as fs from 'fs';
import * as path from 'path';
import { marked } from 'marked';
import {
  Document,
  Paragraph,
  TextRun,
  Table,
  TableRow,
  TableCell,
  HeadingLevel,
  Header,
  Footer,
  ImageRun,
  BorderStyle,
  WidthType,
  AlignmentType,
  ShadingType,
  Packer,
} from 'docx';
import {
  ConverterOptions,
  DEFAULT_ACCENT_COLOR,
  DEFAULT_DARK_COLOR,
  DEFAULT_LIGHT_BG,
  DEFAULT_TEXT_COLOR,
  TABLE_BORDER_STYLE,
  CELL_MARGINS,
} from './styles';

export async function convertMarkdownToDocx(
  markdownContent: string,
  options: ConverterOptions = {}
): Promise<Buffer> {
  const accentColor = (options.accentColor || DEFAULT_ACCENT_COLOR).replace('#', '');
  const tokens = marked.lexer(markdownContent);
  const children: (Paragraph | Table)[] = [];

  for (let i = 0; i < tokens.length; i++) {
    const token = tokens[i];

    if (token.type === 'heading') {
      const headingLevel = getHeadingLevel(token.depth);
      const runs = parseInlineFormatting(token.text);
      children.push(
        new Paragraph({
          heading: headingLevel,
          children: runs,
          spacing: {
            before: token.depth === 1 ? 300 : 200,
            after: 100,
          },
          keepNext: true,
        })
      );
    } else if (token.type === 'paragraph') {
      const text = token.text;

      // Handle GitHub Alert Blocks (> [!IMPORTANT])
      if (text.startsWith('[!') && text.includes(']')) {
        const alertTypeMatch = text.match(/^\[!([A-Z]+)\]\s*(.*)/s);
        if (alertTypeMatch) {
          const type = alertTypeMatch[1];
          const contentText = alertTypeMatch[2];
          children.push(createAlertBox(type, contentText, accentColor));
          continue;
        }
      }

      const runs = parseInlineFormatting(text);
      children.push(
        new Paragraph({
          children: runs,
          spacing: { after: 120 },
        })
      );
    } else if (token.type === 'blockquote') {
      const bText = token.text;
      if (bText.includes('[!')) {
        const alertTypeMatch = bText.match(/\[!([A-Z]+)\]\s*(.*)/s);
        if (alertTypeMatch) {
          children.push(createAlertBox(alertTypeMatch[1], alertTypeMatch[2], accentColor));
          continue;
        }
      }
      const runs = parseInlineFormatting(bText);
      children.push(
        new Paragraph({
          children: runs,
          indent: { left: 360 },
          spacing: { before: 120, after: 120 },
        })
      );
    } else if (token.type === 'table') {
      children.push(createDocxTable(token, accentColor));
    } else if (token.type === 'code') {
      if (token.lang === 'mermaid') {
        const mermaidImgParagraph = await renderMermaidDiagram(token.text);
        if (mermaidImgParagraph) {
          children.push(mermaidImgParagraph);
          continue;
        }
      }
      children.push(createCodeBlockContainer(token.text));
    } else if (token.type === 'list') {
      token.items.forEach((item: any, idx: number) => {
        const itemRuns = parseInlineFormatting(item.text);
        const bulletText = token.ordered
          ? `${(typeof token.start === 'number' ? token.start : 1) + idx}. `
          : '• ';
        children.push(
          new Paragraph({
            children: [
              new TextRun({ text: bulletText, bold: true, color: accentColor }),
              ...itemRuns,
            ],
            indent: { left: 360 },
            spacing: { after: 60 },
          })
        );
      });
    } else if (token.type === 'html') {
      const htmlText = token.text.trim();
      if (/^(?:<br(?:\s+[^>]*)?\/?>\s*)+$/i.test(htmlText)) {
        const brMatches = htmlText.match(/<br(?:\s+[^>]*)?\/?>/gi);
        const count = brMatches ? brMatches.length : 1;
        const breakRuns: TextRun[] = [];
        for (let b = 0; b < count; b++) {
          breakRuns.push(new TextRun({ break: 1 }));
        }
        children.push(
          new Paragraph({
            children: breakRuns,
            spacing: { after: 120 },
          })
        );
      } else if (/^<hr\s*\/?>/i.test(htmlText)) {
        children.push(
          new Paragraph({
            border: {
              bottom: { style: BorderStyle.SINGLE, size: 6, color: 'CBD5E1' },
            },
            spacing: { before: 180, after: 180 },
          })
        );
      } else {
        const stripped = htmlText.replace(/<\/?(?:p|div|span|section|article)[^>]*>/gi, '').trim();
        if (stripped) {
          const runs = parseInlineFormatting(stripped);
          children.push(
            new Paragraph({
              children: runs,
              spacing: { after: 120 },
            })
          );
        }
      }
    } else if (token.type === 'hr') {
      children.push(
        new Paragraph({
          border: {
            bottom: { style: BorderStyle.SINGLE, size: 6, color: 'CBD5E1' },
          },
          spacing: { before: 180, after: 180 },
        })
      );
    }
  }

  // Configure Header & Footer Elements
  let headerElement: Header | undefined;
  let footerElement: Footer | undefined;

  if (options.headerImage && fs.existsSync(options.headerImage)) {
    const headerImgBuffer = fs.readFileSync(options.headerImage);
    headerElement = new Header({
      children: [
        new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [
            new ImageRun({
              data: headerImgBuffer,
              transformation: { width: 595, height: 60 },
            }),
          ],
        }),
      ],
    });
  }

  if (options.footerImage && fs.existsSync(options.footerImage)) {
    const footerImgBuffer = fs.readFileSync(options.footerImage);
    footerElement = new Footer({
      children: [
        new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [
            new ImageRun({
              data: footerImgBuffer,
              transformation: { width: 595, height: 50 },
            }),
          ],
        }),
      ],
    });
  }

  const doc = new Document({
    styles: {
      default: {
        document: {
          run: {
            font: 'Helvetica',
            size: 20, // 10pt
            color: DEFAULT_TEXT_COLOR,
          },
        },
      },
    },
    sections: [
      {
        properties: {
          page: {
            margin: {
              top: headerElement ? 1400 : 1440,
              bottom: footerElement ? 1200 : 1440,
              left: 1080,
              right: 1080,
            },
          },
        },
        headers: headerElement ? { default: headerElement } : undefined,
        footers: footerElement ? { default: footerElement } : undefined,
        children: children,
      },
    ],
  });

  return await Packer.toBuffer(doc);
}

function getHeadingLevel(depth: number): (typeof HeadingLevel)[keyof typeof HeadingLevel] {
  switch (depth) {
    case 1:
      return HeadingLevel.HEADING_1;
    case 2:
      return HeadingLevel.HEADING_2;
    case 3:
      return HeadingLevel.HEADING_3;
    case 4:
      return HeadingLevel.HEADING_4;
    case 5:
      return HeadingLevel.HEADING_5;
    default:
      return HeadingLevel.HEADING_6;
  }
}

interface InlineStyleOptions {
  bold?: boolean;
  italics?: boolean;
  strike?: boolean;
  font?: string;
  size?: number;
  color?: string;
}

function decodeHtmlEntities(str: string): string {
  const entityMap: Record<string, string> = {
    nbsp: '\u00A0',
    ensp: '\u2002',
    emsp: '\u2003',
    thinsp: '\u2009',
    amp: '&',
    lt: '<',
    gt: '>',
    quot: '"',
    apos: "'",
    copy: '©',
    reg: '®',
    trade: '™',
    bull: '•',
    ndash: '–',
    mdash: '—',
  };

  let decoded = str.replace(/&(?:nbsp;?|#160;?|#xa0;?)/gi, '\u00A0');

  decoded = decoded.replace(/&(?:#(\d+);?|#x([0-9a-fA-F]+);?|([a-zA-Z]+);)/g, (match, dec, hex, name) => {
    if (dec) {
      const code = parseInt(dec, 10);
      return String.fromCharCode(code);
    }
    if (hex) {
      const code = parseInt(hex, 16);
      return String.fromCharCode(code);
    }
    const lower = name.toLowerCase();
    if (entityMap[lower] !== undefined) {
      return entityMap[lower];
    }
    return match;
  });

  return decoded;
}

function addTextWithLineBreaks(
  rawText: string,
  runs: TextRun[],
  styleOptions: InlineStyleOptions = {}
): void {
  const brRegex = /\s*<br(?:\s+[^>]*)?\/?>\s*/gi;
  const segments = rawText.split(brRegex);

  for (let i = 0; i < segments.length; i++) {
    if (i > 0) {
      runs.push(new TextRun({ break: 1 }));
    }
    const decoded = decodeHtmlEntities(segments[i]);
    if (decoded.length > 0) {
      runs.push(new TextRun({ ...styleOptions, text: decoded }));
    }
  }
}

function parseInlineFormatting(
  text: string,
  baseStyle: InlineStyleOptions = {}
): TextRun[] {
  const runs: TextRun[] = [];
  const inlineRegex = /(\*\*\*[\s\S]*?\*\*\*|___[\s\S]*?___|\*\*[\s\S]*?\*\*|__[\s\S]*?__|~~[\s\S]*?~~|`[^`]*?`|\*[^*\n]+?\*|_[^_\n]+?_)/g;
  const parts = text.split(inlineRegex);

  for (const part of parts) {
    if (!part) continue;

    if (
      (part.startsWith('***') && part.endsWith('***')) ||
      (part.startsWith('___') && part.endsWith('___'))
    ) {
      const inner = part.slice(3, -3);
      addTextWithLineBreaks(inner, runs, {
        ...baseStyle,
        bold: true,
        italics: true,
      });
    } else if (
      (part.startsWith('**') && part.endsWith('**')) ||
      (part.startsWith('__') && part.endsWith('__'))
    ) {
      const inner = part.slice(2, -2);
      addTextWithLineBreaks(inner, runs, {
        ...baseStyle,
        bold: true,
      });
    } else if (
      (part.startsWith('*') && part.endsWith('*')) ||
      (part.startsWith('_') && part.endsWith('_'))
    ) {
      const inner = part.slice(1, -1);
      addTextWithLineBreaks(inner, runs, {
        ...baseStyle,
        italics: true,
      });
    } else if (part.startsWith('~~') && part.endsWith('~~')) {
      const inner = part.slice(2, -2);
      addTextWithLineBreaks(inner, runs, {
        ...baseStyle,
        strike: true,
      });
    } else if (part.startsWith('`') && part.endsWith('`')) {
      const inner = part.slice(1, -1);
      runs.push(
        new TextRun({
          ...baseStyle,
          text: inner,
          font: 'Courier New',
          size: 18,
          color: DEFAULT_DARK_COLOR,
        })
      );
    } else {
      addTextWithLineBreaks(part, runs, baseStyle);
    }
  }

  return runs.length > 0 ? runs : [new TextRun({ ...baseStyle, text: '' })];
}

function createDocxTable(tableToken: any, accentColor: string): Table {
  const rows: TableRow[] = [];

  // Header Row
  const headerCells = tableToken.header.map((col: any) => {
    const runs = parseInlineFormatting(col.text, {
      bold: true,
      color: 'FFFFFF',
      size: 18,
    });
    return new TableCell({
      children: [
        new Paragraph({
          children: runs,
        }),
      ],
      shading: {
        fill: '000000',
        type: ShadingType.CLEAR,
      },
      margins: CELL_MARGINS,
      borders: {
        top: TABLE_BORDER_STYLE,
        bottom: TABLE_BORDER_STYLE,
        left: TABLE_BORDER_STYLE,
        right: TABLE_BORDER_STYLE,
      },
    });
  });

  rows.push(new TableRow({ children: headerCells }));

  // Data Rows
  tableToken.rows.forEach((row: any[], rIdx: number) => {
    const cells = row.map((cell: any) => {
      const runs = parseInlineFormatting(cell.text);
      return new TableCell({
        children: [new Paragraph({ children: runs })],
        shading: {
          fill: rIdx % 2 === 0 ? 'FFFFFF' : DEFAULT_LIGHT_BG,
          type: ShadingType.CLEAR,
        },
        margins: CELL_MARGINS,
        borders: {
          top: TABLE_BORDER_STYLE,
          bottom: TABLE_BORDER_STYLE,
          left: TABLE_BORDER_STYLE,
          right: TABLE_BORDER_STYLE,
        },
      });
    });
    rows.push(new TableRow({ children: cells }));
  });

  return new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    rows: rows,
  });
}

function createAlertBox(type: string, content: string, accentColor: string): Paragraph {
  const label = type === 'IMPORTANT' ? 'IMPORTANT OPERATIONAL NOTICE' : type;
  const runs = parseInlineFormatting(content);

  return new Paragraph({
    children: [
      new TextRun({ text: `${label}: `, bold: true, color: '78350F' }),
      ...runs,
    ],
    border: {
      left: { style: BorderStyle.SINGLE, size: 24, color: accentColor },
    },
    shading: {
      fill: 'FFFBE6',
      type: ShadingType.CLEAR,
    },
    indent: { left: 180 },
    spacing: { before: 140, after: 140 },
  });
}

function createCodeBlockContainer(codeText: string): Paragraph {
  const lines = codeText.split('\n');
  const runs: TextRun[] = [];

  lines.forEach((line, idx) => {
    runs.push(new TextRun({ text: line, font: 'Courier New', size: 18, color: '0F172A' }));
    if (idx < lines.length - 1) {
      runs.push(new TextRun({ break: 1 }));
    }
  });

  return new Paragraph({
    children: runs,
    shading: {
      fill: 'F1F5F9',
      type: ShadingType.CLEAR,
    },
    border: {
      top: { style: BorderStyle.SINGLE, size: 4, color: 'CBD5E1' },
      bottom: { style: BorderStyle.SINGLE, size: 4, color: 'CBD5E1' },
      left: { style: BorderStyle.SINGLE, size: 4, color: 'CBD5E1' },
      right: { style: BorderStyle.SINGLE, size: 4, color: 'CBD5E1' },
    },
    indent: { left: 180 },
    spacing: { before: 120, after: 120 },
  });
}

async function renderMermaidDiagram(mermaidCode: string): Promise<Paragraph | null> {
  try {
    const cleanCode = mermaidCode.trim();
    const encoded = Buffer.from(cleanCode).toString('base64');
    const url = `https://mermaid.ink/img/${encoded}?bgColor=FFFFFF`;

    const res = await fetch(url);
    if (res.ok) {
      const arrayBuffer = await res.arrayBuffer();
      const imageBuffer = Buffer.from(arrayBuffer);

      return new Paragraph({
        alignment: AlignmentType.CENTER,
        children: [
          new ImageRun({
            data: imageBuffer,
            transformation: { width: 520, height: 280 },
          }),
        ],
        spacing: { before: 180, after: 180 },
      });
    }
  } catch (err) {
    console.warn('Could not render Mermaid diagram to image:', err);
  }
  return null;
}
