import express from 'express';
import swaggerUi from 'swagger-ui-express';
import { readFileSync } from 'fs';
import { fileURLToPath } from 'url';
import { dirname, join } from 'path';
import { SearchRequest } from './search-request.js';
import { CitationStyle, StyleName } from './citation-style.js';
import { CitationGenerator } from './citation-generator.js';
import { CrossrefProvider } from './crossref-provider.js';
import { OpenLibraryProvider } from './open-library-provider.js';
import { WebScraperProvider } from './web-scraper-provider.js';
import { IdentifierType } from './types.js';
import { MetadataProvider } from './metadata-provider.js';

const __dirname = dirname(fileURLToPath(import.meta.url));
const swaggerDocument = JSON.parse(readFileSync(join(__dirname, 'openapi.json'), 'utf-8'));

const app = express();
app.use(express.json());
app.use((_req, res, next) => {
  res.header('Access-Control-Allow-Origin', '*');
  res.header('Access-Control-Allow-Headers', 'Content-Type');
  res.header('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
  next();
});

app.use('/docs', swaggerUi.serve, swaggerUi.setup(swaggerDocument));
app.get('/', (_req, res) => { res.redirect('/docs'); });

const providers = new Map<IdentifierType, MetadataProvider>([
  [IdentifierType.DOI, new CrossrefProvider()],
  [IdentifierType.ISBN, new OpenLibraryProvider()],
  [IdentifierType.URL, new WebScraperProvider()],
]);
const generator = new CitationGenerator(providers);

app.get('/health', (_req, res) => {
  res.json({ status: 'ok', timestamp: new Date().toISOString() });
});

app.get('/styles', (_req, res) => {
  res.json({ styles: ['APA', 'MLA', 'Chicago'] });
});

app.post('/citations/generate', async (req, res) => {
  const { identifier, style } = req.body;

  if (!identifier || !style) {
    res.status(400).json({ error: 'identifier and style are required' });
    return;
  }

  const validStyles: StyleName[] = ['APA', 'MLA', 'Chicago'];
  if (!validStyles.includes(style)) {
    res.status(400).json({ error: `Invalid style. Use one of: ${validStyles.join(', ')}` });
    return;
  }

  try {
    const request = new SearchRequest(identifier);
    const citationStyle = new CitationStyle(style);
    const citation = await generator.generate(request, citationStyle);

    res.json({
      citation: citation.getFormattedText(),
      metadata: citation.rawMetadata,
      style: citation.getStyle().name,
    });
  } catch (err) {
    const message = err instanceof Error ? err.message : 'Unknown error';
    res.status(422).json({ error: message });
  }
});

const PORT = parseInt(process.env.PORT || '3000', 10);
app.listen(PORT, () => {
  console.log(`Server running on port ${PORT}`);
});

export default app;
