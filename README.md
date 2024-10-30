# GPT Extension for Google Docs

A powerful Google Docs extension that integrates OpenAI's GPT models to enhance your document editing experience. The extension provides various text processing capabilities with outputs in Russian and English translations.

## Features

### Text Processing
- **Summarize Text (in Russian)**: Create concise summaries of selected text
- **Improve Writing (in Russian)**: Enhance text clarity and professionalism
- **Fix Grammar (in Russian)**: Correct grammatical errors
- **Make Formal/Casual (in Russian)**: Adjust text tone
- **Translate to English**: Convert text from any language to English

### User Interface
- **Convenient Sidebar**: Easy-to-use interface for text processing
- **Custom Menu**: Quick access to all features from the Google Docs menu
- **Settings Panel**: Customize API and model parameters

### Settings Configuration
- Base URL for OpenAI API
- Model selection (GPT-3.5-turbo, GPT-4, GPT-4-turbo-preview, Claude-3-Sonnet, or custom)
- Temperature adjustment (0-1)
- Max tokens limit (150+)

## Installation

1. Open Google Apps Script editor
2. Create new project
3. Copy the following files into your project:
   - `Code.gs`
   - `Sidebar.html`
   - `Settings.html`
4. Set up your OpenAI API key:
   - Go to Project Settings
   - Click on "Script Properties"
   - Add new property:
     - Name: `OPENAI_API_KEY`
     - Value: Your OpenAI API key
5. Save and deploy as Google Docs add-on

## Usage

1. Open a Google Doc
2. Find "GPT Extension" in the menu bar
3. Choose one of the following options:
   - Use the sidebar for interactive text processing
   - Use direct menu options for quick actions
   - Configure settings for customization

### Sidebar Usage
1. Click "Show Sidebar" from the GPT Extension menu
2. Enter or paste your text
3. Choose desired operation
4. Review the result
5. Click "Insert to Document" to add the processed text

### Settings Configuration
1. Click "Settings" in the GPT Extension menu
2. Configure:
   - API Base URL
   - Model selection
   - Temperature
   - Max tokens
3. Click Save to apply changes

## Project Structure

```
├── Code.gs                  # Main script file with core functionality
├── Sidebar.html            # UI for prompts and responses
├── Settings.html           # Settings panel interface
└── appsscript.json         # Project configuration (auto-generated)
```

## Technical Details

- Built using Google Apps Script
- Integrates with OpenAI API
- Supports multiple GPT models
- Handles multiple languages with focus on Russian output
- Implements error handling and loading states
- Uses Script Properties for secure API key storage

## Requirements

- Google Workspace account
- OpenAI API key
- Google Chrome browser (recommended)

## Security Note

The extension stores your OpenAI API key in Google Apps Script's secure Script Properties. Never share your API key or include it directly in the code.

## Contributing

Feel free to submit issues and enhancement requests!

## License

[MIT License](LICENSE) 