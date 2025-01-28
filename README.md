# Azure AI for Google Sheets

A Google Sheets add-on that integrates Azure AI capabilities directly into your spreadsheets, allowing for powerful text analysis and AI-driven content generation.

## Features

- Direct integration with Azure OpenAI services
- Custom function for text analysis (AZURE_ANALYZE_TEXT)
- Built-in rate limiting and retry mechanisms
- Automatic error handling and recovery
- User-friendly configuration interface

## Setup

1. Open your Google Sheet
2. Navigate to the "Azure AI" menu in the top menu bar
3. Click "Configure Azure Settings"
4. Enter your Azure OpenAI API credentials:
   - Endpoint URL
   - API Key

## Usage

Use the `AZURE_ANALYZE_TEXT` function in any cell with three parameters:

```
=AZURE_ANALYZE_TEXT(text, systemPrompt, temperature)
```

- `text`: The content you want to analyze (required)
- `systemPrompt`: Instructions for the AI (optional)
- `temperature`: Controls response creativity (optional, default: 0.7)

## Rate Limiting

The add-on includes built-in rate limiting:
- Maximum 120 requests per minute
- Automatic retry with exponential backoff
- Request spacing of 500ms minimum

## Error Handling

- Automatic retry for failed requests
- Loading status indicators
- Batch processing for error recovery

## Development

This project is built using Google Apps Script and integrates with Azure OpenAI services. The main components:

- `Code.gs`: Core functionality and API integration
- `ConfigDialog.html`: Configuration interface
- `index.html`: Landing page
- `privacy.html`: Privacy policy

## License

MIT

## Created By

Aura AI Taskforce