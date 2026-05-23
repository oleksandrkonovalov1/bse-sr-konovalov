import express from 'express';
import { SearchRequest } from './search-request.js';
import { CitationStyle, StyleName } from './citation-style.js';
import { CitationGenerator } from './citation-generator.js';
import { CrossrefProvider } from './crossref-provider.js';
import { OpenLibraryProvider } from './open-library-provider.js';
import { WebScraperProvider } from './web-scraper-provider.js';
import { IdentifierType } from './types.js';
import { MetadataProvider } from './metadata-provider.js';

const app = express();
app.use(express.json());

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
