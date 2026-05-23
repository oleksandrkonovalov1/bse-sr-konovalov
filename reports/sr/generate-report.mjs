#!/usr/bin/env node

/**
 * Self-study report generator
 * Discipline: Основи програмної інженерії
 * Topic: Розгортання застосунку на сервері або у хмарі
 */

import {
  Document, Packer, Paragraph, TextRun, Header, ImageRun,
  AlignmentType, PageNumber, HeadingLevel,
  PageBreak, BorderStyle, ShadingType,
  Table, TableRow, TableCell, WidthType, TableBorders,
} from "docx";
import { writeFileSync, readFileSync } from "fs";
import { dirname, join } from "path";
import { fileURLToPath } from "url";

const __dirname = dirname(fileURLToPath(import.meta.url));

const MM_TO_DXA = 56.693;
const PT_TO_HALF_PT = 2;

const FONT = "Times New Roman";
const FONT_CODE = "Courier New";
const BODY_SIZE = 14 * PT_TO_HALF_PT;
const CODE_SIZE = 9 * PT_TO_HALF_PT;
const TITLE_SIZE = 14 * PT_TO_HALF_PT;

const margins = {
  top:    Math.round(20 * MM_TO_DXA),
  bottom: Math.round(20 * MM_TO_DXA),
  left:   Math.round(30 * MM_TO_DXA),
  right:  Math.round(15 * MM_TO_DXA),
};

const LINE_SPACING_15 = 360;

// ── Typography helpers ────────────────────────────────────────────

function titleRun(text, opts = {}) {
  return new TextRun({
    text,
    font: FONT,
    size: TITLE_SIZE,
    bold: opts.bold ?? false,
    ...opts,
  });
}

function bodyRun(text, opts = {}) {
  return new TextRun({
    text,
    font: FONT,
    size: BODY_SIZE,
    ...opts,
  });
}

function centeredParagraph(runs, spacing = {}) {
  return new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { after: 0, line: LINE_SPACING_15, lineRule: "auto", ...spacing },
    children: Array.isArray(runs) ? runs : [runs],
  });
}

function emptyLine() {
  return centeredParagraph(titleRun(""));
}

function sectionHeading(number, title, { pageBreakBefore: pbBefore = false } = {}) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_1,
    spacing: { before: 240, after: 120, line: LINE_SPACING_15, lineRule: "auto" },
    keepNext: true,
    pageBreakBefore: pbBefore,
    children: [
      new TextRun({
        text: `${number} ${title}`.toUpperCase(),
        font: FONT,
        size: BODY_SIZE,
        bold: true,
      }),
    ],
  });
}

function subsectionHeading(number, title) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_2,
    spacing: { before: 120, after: 60, line: LINE_SPACING_15, lineRule: "auto" },
    indent: { firstLine: Math.round(12.5 * MM_TO_DXA) },
    keepNext: true,
    children: [
      new TextRun({
        text: `${number} ${title}`,
        font: FONT,
        size: BODY_SIZE,
        bold: true,
      }),
    ],
  });
}

function bodyParagraph(text) {
  return new Paragraph({
    spacing: { after: 0, line: LINE_SPACING_15, lineRule: "auto" },
    indent: { firstLine: Math.round(12.5 * MM_TO_DXA) },
    alignment: AlignmentType.JUSTIFIED,
    children: [bodyRun(text)],
  });
}

function codeParagraph(text) {
  return new Paragraph({
    spacing: { after: 0, line: 240, lineRule: "auto" },
    indent: { left: 283 },
    shading: { fill: "F2F2F2", type: ShadingType.CLEAR },
    border: { left: { style: BorderStyle.SINGLE, size: 12, color: "4472C4", space: 4 } },
    children: [
      new TextRun({
        text,
        font: FONT_CODE,
        size: CODE_SIZE,
      }),
    ],
  });
}

function listingCaption(number, title) {
  return new Paragraph({
    spacing: { before: 120, after: 60, line: LINE_SPACING_15, lineRule: "auto" },
    indent: { firstLine: Math.round(12.5 * MM_TO_DXA) },
    keepNext: true,
    children: [
      bodyRun(`Лістинг ${number} — ${title}`),
    ],
  });
}

function tableCaption(number, title) {
  return new Paragraph({
    spacing: { before: 120, after: 60, line: LINE_SPACING_15, lineRule: "auto" },
    indent: { firstLine: Math.round(12.5 * MM_TO_DXA) },
    children: [
      bodyRun(`Таблиця ${number} — ${title}`),
    ],
  });
}

function codeBlock(lines) {
  return lines.map(line => codeParagraph(line));
}

function bulletParagraph(text) {
  return new Paragraph({
    spacing: { after: 0, line: LINE_SPACING_15, lineRule: "auto" },
    indent: { firstLine: Math.round(12.5 * MM_TO_DXA) },
    alignment: AlignmentType.JUSTIFIED,
    children: [bodyRun(`– ${text}`)],
  });
}

function numberedItem(num, text) {
  return new Paragraph({
    spacing: { after: 0, line: LINE_SPACING_15, lineRule: "auto" },
    indent: { firstLine: Math.round(12.5 * MM_TO_DXA) },
    alignment: AlignmentType.JUSTIFIED,
    children: [bodyRun(`${num}. ${text}`)],
  });
}

function figureCaption(number, title) {
  return new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 60, after: 120, line: LINE_SPACING_15, lineRule: "auto" },
    children: [bodyRun(`Рисунок ${number} — ${title}`)],
  });
}

function imageParagraph(filePath, widthPx, heightPx) {
  const imgData = readFileSync(join(__dirname, filePath));
  const maxWidthEmu = 15.5 * 914400 / 2.54; // ~15.5cm in EMU (page width minus margins)
  const scale = Math.min(1, maxWidthEmu / (widthPx * 9525));
  return new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 120, after: 0, line: LINE_SPACING_15, lineRule: "auto" },
    children: [
      new ImageRun({
        data: imgData,
        type: "png",
        transformation: {
          width: Math.round(widthPx * scale),
          height: Math.round(heightPx * scale),
        },
      }),
    ],
  });
}

function appendixHeading(letter, title, { pageBreakBefore: pbBefore = true } = {}) {
  return [
    new Paragraph({
      heading: HeadingLevel.HEADING_1,
      alignment: AlignmentType.CENTER,
      spacing: { before: 240, after: 0, line: LINE_SPACING_15, lineRule: "auto" },
      keepNext: true,
      pageBreakBefore: pbBefore,
      children: [
        new TextRun({
          text: `ДОДАТОК ${letter}`,
          font: FONT,
          size: BODY_SIZE,
          bold: true,
        }),
      ],
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 0, after: 120, line: LINE_SPACING_15, lineRule: "auto" },
      keepNext: true,
      children: [
        new TextRun({
          text: title,
          font: FONT,
          size: BODY_SIZE,
          bold: true,
        }),
      ],
    }),
  ];
}

// ── Table helper ──────────────────────────────────────────────────

const ALL_BORDERS = {
  top: { style: BorderStyle.SINGLE, size: 1, color: "000000" },
  bottom: { style: BorderStyle.SINGLE, size: 1, color: "000000" },
  left: { style: BorderStyle.SINGLE, size: 1, color: "000000" },
  right: { style: BorderStyle.SINGLE, size: 1, color: "000000" },
};

function makeTable(headers, rows, columnWidths) {
  const headerCells = headers.map((h, i) =>
    new TableCell({
      borders: ALL_BORDERS,
      width: { size: columnWidths[i], type: WidthType.PERCENTAGE },
      children: [
        new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { after: 0, line: LINE_SPACING_15, lineRule: "auto" },
          children: [bodyRun(h, { bold: true })],
        }),
      ],
    })
  );

  const dataRows = rows.map(row =>
    new TableRow({
      children: row.map((cell, i) =>
        new TableCell({
          borders: ALL_BORDERS,
          width: { size: columnWidths[i], type: WidthType.PERCENTAGE },
          children: [
            new Paragraph({
              alignment: AlignmentType.LEFT,
              spacing: { after: 0, line: LINE_SPACING_15, lineRule: "auto" },
              indent: { left: 57 },
              children: [bodyRun(cell)],
            }),
          ],
        })
      ),
    })
  );

  return new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    rows: [
      new TableRow({ children: headerCells }),
      ...dataRows,
    ],
  });
}

// ── TITLE PAGE ────────────────────────────────────────────────────

const year = new Date().getFullYear();

const titlePageParagraphs = [
  centeredParagraph(titleRun("Міністерство освіти і науки України")),
  centeredParagraph(titleRun("Харківський національний університет радіоелектроніки")),
  emptyLine(),
  centeredParagraph(titleRun("Кафедра програмної інженерії")),
  emptyLine(),
  emptyLine(),
  emptyLine(),
  emptyLine(),
  centeredParagraph(titleRun("ЗВІТ", { bold: true })),
  centeredParagraph(titleRun("з самостійної роботи")),
  centeredParagraph(titleRun("з дисципліни «Основи програмної інженерії»")),
  centeredParagraph(titleRun("на тему: «Розгортання застосунку на сервері або у хмарі»")),
  emptyLine(),
  emptyLine(),
  new Paragraph({
    alignment: AlignmentType.RIGHT,
    spacing: { after: 0, line: LINE_SPACING_15, lineRule: "auto" },
    children: [titleRun("Виконав: ст. гр. ПЗПІ-25-6")],
  }),
  new Paragraph({
    alignment: AlignmentType.RIGHT,
    spacing: { after: 0, line: LINE_SPACING_15, lineRule: "auto" },
    children: [titleRun("Коновалов О. О.")],
  }),
  emptyLine(),
  new Paragraph({
    alignment: AlignmentType.RIGHT,
    spacing: { after: 0, line: LINE_SPACING_15, lineRule: "auto" },
    children: [titleRun("Перевірив: Гребенюк В. О.")],
  }),
];

for (let i = 0; i < 6; i++) titlePageParagraphs.push(emptyLine());

titlePageParagraphs.push(
  centeredParagraph(titleRun(`Харків — ${year}`)),
);

// ── SOURCE CODE (actual content from project files) ───────────────

const serverTsLines = [
  `import express from 'express';`,
  `import swaggerUi from 'swagger-ui-express';`,
  `import { readFileSync } from 'fs';`,
  `import { fileURLToPath } from 'url';`,
  `import { dirname, join } from 'path';`,
  `import { SearchRequest } from './search-request.js';`,
  `import { CitationStyle, StyleName } from './citation-style.js';`,
  `import { CitationGenerator } from './citation-generator.js';`,
  `import { CrossrefProvider } from './crossref-provider.js';`,
  `import { OpenLibraryProvider } from './open-library-provider.js';`,
  `import { WebScraperProvider } from './web-scraper-provider.js';`,
  `import { IdentifierType } from './types.js';`,
  `import { MetadataProvider } from './metadata-provider.js';`,
  ``,
  `const __dirname = dirname(fileURLToPath(import.meta.url));`,
  `const swaggerDocument = JSON.parse(readFileSync(join(__dirname, 'openapi.json'), 'utf-8'));`,
  ``,
  `const app = express();`,
  `app.use(express.json());`,
  `app.use((_req, res, next) => {`,
  `  res.header('Access-Control-Allow-Origin', '*');`,
  `  res.header('Access-Control-Allow-Headers', 'Content-Type');`,
  `  res.header('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');`,
  `  next();`,
  `});`,
  ``,
  `app.use('/docs', swaggerUi.serve, swaggerUi.setup(swaggerDocument));`,
  `app.get('/', (_req, res) => { res.redirect('/docs'); });`,
  ``,
  `const providers = new Map<IdentifierType, MetadataProvider>([`,
  `  [IdentifierType.DOI, new CrossrefProvider()],`,
  `  [IdentifierType.ISBN, new OpenLibraryProvider()],`,
  `  [IdentifierType.URL, new WebScraperProvider()],`,
  `]);`,
  `const generator = new CitationGenerator(providers);`,
  ``,
  `app.get('/health', (_req, res) => {`,
  `  res.json({ status: 'ok', timestamp: new Date().toISOString() });`,
  `});`,
  ``,
  `app.get('/styles', (_req, res) => {`,
  `  res.json({ styles: ['APA', 'MLA', 'Chicago'] });`,
  `});`,
  ``,
  `app.post('/citations/generate', async (req, res) => {`,
  `  const { identifier, style } = req.body;`,
  ``,
  `  if (!identifier || !style) {`,
  `    res.status(400).json({ error: 'identifier and style are required' });`,
  `    return;`,
  `  }`,
  ``,
  `  const validStyles: StyleName[] = ['APA', 'MLA', 'Chicago'];`,
  `  if (!validStyles.includes(style)) {`,
  `    res.status(400).json({ error: \`Invalid style. Use one of: \${validStyles.join(', ')}\` });`,
  `    return;`,
  `  }`,
  ``,
  `  try {`,
  `    const request = new SearchRequest(identifier);`,
  `    const citationStyle = new CitationStyle(style);`,
  `    const citation = await generator.generate(request, citationStyle);`,
  ``,
  `    res.json({`,
  `      citation: citation.getFormattedText(),`,
  `      metadata: citation.rawMetadata,`,
  `      style: citation.getStyle().name,`,
  `    });`,
  `  } catch (err) {`,
  `    const message = err instanceof Error ? err.message : 'Unknown error';`,
  `    res.status(422).json({ error: message });`,
  `  }`,
  `});`,
  ``,
  `const PORT = parseInt(process.env.PORT || '3000', 10);`,
  `app.listen(PORT, () => {`,
  `  console.log(\`Server running on port \${PORT}\`);`,
  `});`,
  ``,
  `export default app;`,
];

const serverTsHandlerLines = serverTsLines.slice(43, 72);

const dockerfileLines = [
  `FROM node:22-alpine AS builder`,
  `WORKDIR /app`,
  `COPY package*.json ./`,
  `RUN npm ci`,
  `COPY tsconfig.build.json ./`,
  `COPY src ./src`,
  `RUN npm run build`,
  ``,
  `FROM node:22-alpine`,
  `WORKDIR /app`,
  `COPY package*.json ./`,
  `RUN npm ci --omit=dev`,
  `COPY --from=builder /app/dist ./dist`,
  `ENV PORT=3000`,
  `EXPOSE 3000`,
  `CMD ["node", "dist/server.js"]`,
];

const dockerComposeLines = [
  `services:`,
  `  api:`,
  `    build: .`,
  `    ports:`,
  `      - "3000:3000"`,
  `    environment:`,
  `      - PORT=3000`,
  `    restart: unless-stopped`,
];

const deployYmlLines = [
  `name: CI/CD`,
  ``,
  `on:`,
  `  push:`,
  `    branches: [main]`,
  ``,
  `jobs:`,
  `  test:`,
  `    runs-on: ubuntu-latest`,
  `    steps:`,
  `      - uses: actions/checkout@v4`,
  `      - uses: actions/setup-node@v4`,
  `        with:`,
  `          node-version: 22`,
  `      - run: npm ci`,
  `      - run: npm test`,
  ``,
  `  deploy:`,
  `    needs: test`,
  `    runs-on: ubuntu-latest`,
  `    if: github.ref == 'refs/heads/main'`,
  `    steps:`,
  `      - name: Deploy to VPS`,
  `        uses: appleboy/ssh-action@v1`,
  `        with:`,
  `          host: \${{ secrets.VPS_HOST }}`,
  `          username: \${{ secrets.VPS_USER }}`,
  `          key: \${{ secrets.VPS_SSH_KEY }}`,
  `          script: |`,
  `            cd ~/bse-sr-konovalov`,
  `            git pull origin main`,
  `            sudo docker-compose up -d --build`,
];

// ── BODY SECTIONS ─────────────────────────────────────────────────

const bodyParagraphs = [
  new Paragraph({ children: [new PageBreak()] }),

  // 1 МЕТА РОБОТИ
  sectionHeading("1", "Мета роботи"),
  bodyParagraph("Набути практичних навичок розгортання веб-застосунків на віддаленому сервері. Перетворити програмний модуль системи генерації бібліографічних цитат (з ЛР 03–04) у REST API та забезпечити його безперервну доставку через CI/CD pipeline."),

  // 2 ЗАВДАННЯ
  sectionHeading("2", "Завдання"),
  numberedItem(1, "Взяти відрефакторений модуль з ЛР 03–04 (Citation Generator, TypeScript/Vitest)."),
  numberedItem(2, "Створити REST API обгортку на Express із ендпоінтами: GET /health, GET /styles, POST /citations/generate."),
  numberedItem(3, "Підготувати Dockerfile для контейнеризації."),
  numberedItem(4, "Розгорнути застосунок на VPS (Hetzner Cloud) через Docker Compose."),
  numberedItem(5, "Налаштувати CI/CD pipeline (GitHub Actions) для автоматичного деплою при push у main."),
  numberedItem(6, "Додати Swagger UI документацію (/docs)."),
  numberedItem(7, "Верифікувати працездатність через публічний URL."),

  // 3 ХІД РОБОТИ
  sectionHeading("3", "Хід роботи"),

  // 3.1
  subsectionHeading("3.1", "Створення REST API"),
  bodyParagraph("Для обгортки модуля використано фреймворк Express 5. Сервер приймає запити на порті 3000 (конфігурується через змінну оточення PORT)."),
  bodyParagraph("Реалізовано три ендпоінти:"),
  bulletParagraph("GET /health — перевірка працездатності, повертає JSON з полями status та timestamp;"),
  bulletParagraph("GET /styles — список підтримуваних стилів цитування (APA, MLA, Chicago);"),
  bulletParagraph("POST /citations/generate — генерація цитати за ідентифікатором (DOI, ISBN або URL) та обраним стилем."),

  listingCaption("3.1", "Ключовий фрагмент server.ts — обробник POST /citations/generate"),
  ...codeBlock(serverTsHandlerLines),

  // 3.2
  subsectionHeading("3.2", "Контейнеризація (Docker)"),
  bodyParagraph("Створено multi-stage Dockerfile: перший етап (builder) компілює TypeScript у JavaScript, другий — копіює лише скомпільований код та production-залежності."),

  listingCaption("3.2", "Dockerfile"),
  ...codeBlock(dockerfileLines),

  bodyParagraph("Для оркестрації контейнера використано Docker Compose."),

  listingCaption("3.3", "docker-compose.yml"),
  ...codeBlock(dockerComposeLines),

  // 3.3
  subsectionHeading("3.3", "Розгортання на VPS"),
  bodyParagraph("Застосунок розгорнуто на VPS Hetzner Cloud (Ubuntu 20.04, 4GB RAM). Для безпеки створено окремого користувача deploy з обмеженими правами sudo (дозвіл лише на docker-compose)."),
  bodyParagraph("Кроки розгортання:"),
  numberedItem(1, "Встановлено Docker та Docker Compose на сервері."),
  numberedItem(2, "Склоновано репозиторій у /home/deploy/bse-sr-konovalov."),
  numberedItem(3, "Запущено контейнер через docker-compose up -d --build."),
  numberedItem(4, "Верифіковано працездатність через публічний IP."),
  bodyParagraph("Публічний URL: http://78.47.249.61:3000"),

  // 3.4
  subsectionHeading("3.4", "CI/CD Pipeline"),
  bodyParagraph("Налаштовано GitHub Actions workflow з двома етапами: test (запуск тестів) та deploy (SSH-підключення до VPS, оновлення коду, перезбірка контейнера)."),

  listingCaption("3.4", "Workflow файл .github/workflows/deploy.yml"),
  ...codeBlock(deployYmlLines),

  bodyParagraph("Для безпечної аутентифікації створено виділений SSH-ключ (ED25519), збережений у GitHub Encrypted Secrets. Ключ надає доступ лише користувачу deploy."),

  // 3.5
  subsectionHeading("3.5", "API документація (Swagger UI)"),
  bodyParagraph("Додано інтерактивну документацію на основі OpenAPI 3.0 специфікації, доступну за адресою /docs. Swagger UI дозволяє тестувати всі ендпоінти безпосередньо у браузері."),

  // 4 РЕЗУЛЬТАТИ
  sectionHeading("4", "Результати", { pageBreakBefore: true }),
  bodyParagraph("Застосунок успішно розгорнуто та доступно за адресою http://78.47.249.61:3000."),
  bodyParagraph("Результати тестування ендпоінтів:"),

  tableCaption("4.1", "Результати тестування ендпоінтів"),
  makeTable(
    ["Ендпоінт", "Метод", "Статус", "Відповідь"],
    [
      ["/health", "GET", "200", '{"status":"ok","timestamp":"2026-05-23T10:05:18.234Z"}'],
      ["/styles", "GET", "200", '{"styles":["APA","MLA","Chicago"]}'],
      ["/citations/generate", "POST", "200", "Генерація цитати у форматі APA за DOI"],
      ["/docs", "GET", "200", "Swagger UI інтерфейс"],
    ],
    [20, 10, 10, 60],
  ),

  bodyParagraph("Конфігурація розгортання:"),

  tableCaption("4.2", "Конфігурація розгортання"),
  makeTable(
    ["Параметр", "Значення"],
    [
      ["Платформа", "Hetzner Cloud VPS"],
      ["ОС", "Ubuntu 20.04 LTS"],
      ["Runtime", "Docker (node:22-alpine)"],
      ["Порт", "3000"],
      ["CI/CD", "GitHub Actions"],
      ["Репозиторій", "github.com/oleksandrkonovalov1/bse-sr-konovalov"],
    ],
    [35, 65],
  ),

  bodyParagraph("Результати CI/CD: усі запуски pipeline завершились успішно. Тести (63/63) проходять перед кожним деплоєм."),

  bodyParagraph("Скріншоти роботи застосунку:"),

  imageParagraph("screenshots/swagger-ui.png", 1510, 810),
  figureCaption("4.1", "Swagger UI — інтерфейс документації API"),

  imageParagraph("screenshots/health-response.png", 1554, 608),
  figureCaption("4.2", "Відповідь ендпоінту /health (код 200)"),

  imageParagraph("screenshots/citations-response.png", 1554, 780),
  figureCaption("4.3", "Відповідь ендпоінту POST /citations/generate (код 200)"),

  imageParagraph("screenshots/hetzner-dashboard.png", 1510, 630),
  figureCaption("4.4", "Dashboard Hetzner Cloud — сервер у статусі Running"),

  // 5 ВИСНОВКИ
  sectionHeading("5", "Висновки", { pageBreakBefore: true }),
  bodyParagraph("У ході виконання самостійної роботи набуто практичних навичок розгортання веб-застосунків. Програмний модуль генерації бібліографічних цитат успішно перетворено на REST API та розгорнуто на віддаленому сервері через Docker. Налаштований CI/CD pipeline забезпечує автоматичну доставку змін при кожному push у main-гілку, що є стандартною практикою у сучасній програмній інженерії."),

  // ДОДАТОК А
  ...appendixHeading("А", "Вихідний код server.ts"),
  ...codeBlock(serverTsLines),

  // ДОДАТОК Б
  ...appendixHeading("Б", "Dockerfile"),
  ...codeBlock(dockerfileLines),

  // ДОДАТОК В
  ...appendixHeading("В", "docker-compose.yml"),
  ...codeBlock(dockerComposeLines),

  // ДОДАТОК Г
  ...appendixHeading("Г", "CI/CD Pipeline (deploy.yml)"),
  ...codeBlock(deployYmlLines),
];

// ── DOCUMENT ──────────────────────────────────────────────────────

const doc = new Document({
  styles: {
    default: {
      document: {
        run: {
          font: FONT,
          size: BODY_SIZE,
          language: { value: "uk-UA" },
        },
        paragraph: {
          spacing: { after: 0, line: LINE_SPACING_15, lineRule: "auto" },
        },
      },
    },
    paragraphStyles: [
      {
        id: "Heading1",
        name: "Heading 1",
        basedOn: "Normal",
        next: "Normal",
        quickFormat: true,
        run: { size: BODY_SIZE, bold: true, font: FONT },
        paragraph: {
          spacing: { before: 240, after: 120, line: LINE_SPACING_15, lineRule: "auto" },
          outlineLevel: 0,
        },
      },
      {
        id: "Heading2",
        name: "Heading 2",
        basedOn: "Normal",
        next: "Normal",
        quickFormat: true,
        run: { size: BODY_SIZE, bold: true, font: FONT },
        paragraph: {
          spacing: { before: 120, after: 60, line: LINE_SPACING_15, lineRule: "auto" },
          outlineLevel: 1,
        },
      },
    ],
  },
  sections: [
    {
      properties: {
        page: {
          size: { width: 11906, height: 16838 },
          margin: { ...margins, header: 708, footer: 708 },
        },
        titlePage: true,
      },
      headers: {
        default: new Header({
          children: [
            new Paragraph({
              alignment: AlignmentType.RIGHT,
              children: [
                new TextRun({ children: [PageNumber.CURRENT], font: FONT, size: BODY_SIZE }),
              ],
            }),
          ],
        }),
      },
      children: [...titlePageParagraphs, ...bodyParagraphs],
    },
  ],
});

const OUTPUT = "reports/sr/Звіт_СР_Коновалов_ПЗПІ-25-6.docx";
const buffer = await Packer.toBuffer(doc);
writeFileSync(OUTPUT, buffer);
console.log(`Created: ${OUTPUT}`);
