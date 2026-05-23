# Citation Generator REST API

REST API for generating academic citations from DOI, ISBN, or URL identifiers.

## Endpoints

- `GET /health` — Health check
- `GET /styles` — List available citation styles
- `POST /citations/generate` — Generate a citation

## Usage

```bash
# Health check
curl http://localhost:3000/health

# List styles
curl http://localhost:3000/styles

# Generate citation
curl -X POST http://localhost:3000/citations/generate \
  -H "Content-Type: application/json" \
  -d '{"identifier": "10.1145/3544548.3581388", "style": "APA"}'
```

## Development

```bash
npm install
npm run dev    # Development server
npm test       # Run tests
npm run build  # Build for production
npm start      # Start production server
```
