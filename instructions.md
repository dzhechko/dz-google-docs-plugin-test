# Project Overview
GPT extension for Google Docs using Google Apps Script

# Core Functionalities
- Connect Google Apps Script to OpenAI API
- Add a Custom Menu "GPT Extension" in Google Docs
- Set Up the Sidebar UI that can provide a user-friendly interface where users enter prompts and view results
- There should be a settings panel, so one can edit 
the following parameters:
-- base url of openai compatible model
-- model itself (either chose from the short list of 4 most popular openai models or enter manually the name of the model)
-- temperature (from 0 to 1, of possible use progress bar or something like this)
-- max tokens (from 150 till infinity)

# Documentation
## Connect Google Apps Script to OpenAI API
```
// Your OpenAI API Key - store this in Script Properties for security
const OPENAI_API_KEY = PropertiesService.getScriptProperties().getProperty('OPENAI_API_KEY');
const OPENAI_API_URL = 'https://api.openai.com/v1/chat/completions';

function callOpenAI(prompt) {
  const headers = {
    'Authorization': 'Bearer ' + OPENAI_API_KEY,
    'Content-Type': 'application/json'
  };
  
  const payload = {
    'model': 'gpt-3.5-turbo',
    'messages': [
      {
        'role': 'user',
        'content': prompt
      }
    ],
    'temperature': 0.7
  };
  
  const options = {
    'method': 'post',
    'headers': headers,
    'payload': JSON.stringify(payload),
    'muteHttpExceptions': true
  };
  
  try {
    const response = UrlFetchApp.fetch(OPENAI_API_URL, options);
    const json = JSON.parse(response.getContentText());
    return json.choices[0].message.content;
  } catch (error) {
    Logger.log('Error: ' + error);
    return 'Error: ' + error;
  }
}

// Example function to test the API
function testOpenAI() {
  const prompt = "What is the capital of France?";
  const response = callOpenAI(prompt);
  Logger.log(response);
}
```

## Add a Custom Menu "GPT Extension" in Google Docs
// Create menu when the document opens
function onOpen() {
  DocumentApp.getUi()
    .createMenu('PU GPT Extension')
    .addItem('Summarize Selection', 'summarizeSelection')
    .addItem('Improve Writing', 'improveWriting')
    .addItem('Translate to Thai', 'translateToThai')
    .addSeparator()
    .addSubMenu(DocumentApp.getUi().createMenu('Format')
      .addItem('Fix Grammar', 'fixGrammar')
      .addItem('Make Formal', 'makeFormal')
      .addItem('Make Casual', 'makeCasual'))
    .addToUi();
}

// Helper function to get selected text
function getSelectedText() {
  const doc = DocumentApp.getActiveDocument();
  const selection = doc.getSelection();
  
  if (selection) {
    const elements = selection.getSelectedElements();
    return elements.map(element => element.getElement().asText().getText()).join(' ');
  } else {
    throw new Error('Please select some text first.');
  }
}

// Helper function to replace selected text
function replaceSelectedText(newText) {
  const doc = DocumentApp.getActiveDocument();
  const selection = doc.getSelection();
  
  if (selection) {
    const elements = selection.getSelectedElements();
    elements.forEach(element => {
      if (element.getElement().editAsText) {
        const text = element.getElement().editAsText();
        const startOffset = element.getStartOffset();
        const endOffset = element.getEndOffsetInclusive();
        
        if (startOffset !== null && endOffset !== null) {
          text.deleteText(startOffset, endOffset);
          text.insertText(startOffset, newText);
        }
      }
    });
  }
}

// Example menu functions
function summarizeSelection() {
  try {
    const selectedText = getSelectedText();
    const prompt = `Please summarize the following text concisely:\n\n${selectedText}`;
    const summary = callOpenAI(prompt); // Using the previous OpenAI function
    
    const ui = DocumentApp.getUi();
    const response = ui.alert('Summary', summary, ui.ButtonSet.OK_CANCEL);
    
    if (response === ui.Button.OK) {
      replaceSelectedText(summary);
    }
  } catch (error) {
    DocumentApp.getUi().alert('Error: ' + error.toString());
  }
}

function improveWriting() {
  try {
    const selectedText = getSelectedText();
    const prompt = `Please improve the following text, making it more professional and clear while maintaining its meaning:\n\n${selectedText}`;
    const improved = callOpenAI(prompt);
    
    const ui = DocumentApp.getUi();
    const response = ui.alert('Improved Text', improved, ui.ButtonSet.OK_CANCEL);
    
    if (response === ui.Button.OK) {
      replaceSelectedText(improved);
    }
  } catch (error) {
    DocumentApp.getUi().alert('Error: ' + error.toString());
  }
}

function translateToThai() {
  try {
    const selectedText = getSelectedText();
    const prompt = `Please translate the following text to Thai:\n\n${selectedText}`;
    const translated = callOpenAI(prompt);
    
    const ui = DocumentApp.getUi();
    const response = ui.alert('Thai Translation', translated, ui.ButtonSet.OK_CANCEL);
    
    if (response === ui.Button.OK) {
      replaceSelectedText(translated);
    }
  } catch (error) {
    DocumentApp.getUi().alert('Error: ' + error.toString());
  }
}

function fixGrammar() {
  try {
    const selectedText = getSelectedText();
    const prompt = `Please fix any grammar issues in the following text:\n\n${selectedText}`;
    const fixed = callOpenAI(prompt);
    
    const ui = DocumentApp.getUi();
    const response = ui.alert('Grammar Fixed', fixed, ui.ButtonSet.OK_CANCEL);
    
    if (response === ui.Button.OK) {
      replaceSelectedText(fixed);
    }
  } catch (error) {
    DocumentApp.getUi().alert('Error: ' + error.toString());
  }
}

## Set Up the Sidebar UI that can provide a user-friendly interface where users enter prompts and view results

Code.gs
```
function showSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('Sidebar')
    .setTitle('GPT Assistant')
    .setWidth(300);
  DocumentApp.getUi().showSidebar(html);
}

// Add this to your onOpen() menu
// ... existing menu code ...
.addItem('Show GPT Assistant', 'showSidebar')
// ... rest of menu code ...
```

Sidebar.html
```
<!DOCTYPE html>
<html>
<head>
    <base target="_top">
    <style>
        body {
            font-family: Arial, sans-serif;
            padding: 10px;
        }
        #promptInput {
            width: 100%;
            height: 100px;
            margin-bottom: 10px;
            padding: 8px;
            border: 1px solid #ddd;
            border-radius: 4px;
            resize: vertical;
        }
        #response {
            width: 100%;
            min-height: 150px;
            margin-top: 10px;
            padding: 8px;
            border: 1px solid #ddd;
            border-radius: 4px;
            background-color: #f9f9f9;
            white-space: pre-wrap;
        }
        .button-container {
            display: flex;
            gap: 8px;
            margin-bottom: 10px;
        }
        button {
            background-color: #4285f4;
            color: white;
            padding: 8px 16px;
            border: none;
            border-radius: 4px;
            cursor: pointer;
        }
        button:hover {
            background-color: #357abd;
        }
        #loading {
            display: none;
            color: #666;
            text-align: center;
            margin: 10px 0;
        }
    </style>
</head>
<body>
    <textarea id="promptInput" placeholder="Enter your prompt here..."></textarea>
    <div class="button-container">
        <button onclick="sendPrompt()">Send</button>
        <button onclick="clearAll()">Clear</button>
    </div>
    <div id="loading">Processing...</div>
    <div id="response"></div>

    <script>
        function sendPrompt() {
            const prompt = document.getElementById('promptInput').value;
            if (!prompt.trim()) return;

            document.getElementById('loading').style.display = 'block';
            
            google.script.run
                .withSuccessHandler(function(result) {
                    document.getElementById('loading').style.display = 'none';
                    document.getElementById('response').textContent = result;
                })
                .withFailureHandler(function(error) {
                    document.getElementById('loading').style.display = 'none';
                    document.getElementById('response').textContent = 'Error: ' + error;
                })
                .callOpenAI(prompt);
        }

        function clearAll() {
            document.getElementById('promptInput').value = '';
            document.getElementById('response').textContent = '';
        }
    </script>
</body>
</html>
```

## Settings panel
<!DOCTYPE html>
<html>
<head>
    <base target="_top">
    <style>
        body {
            font-family: Arial, sans-serif;
            padding: 20px;
        }
        .form-group {
            margin-bottom: 15px;
        }
        label {
            display: block;
            margin-bottom: 5px;
            font-weight: bold;
        }
        input[type="text"], input[type="number"], select {
            width: 100%;
            padding: 8px;
            margin-bottom: 10px;
            border: 1px solid #ddd;
            border-radius: 4px;
        }
        .slider-container {
            display: flex;
            align-items: center;
            gap: 10px;
        }
        .slider {
            flex-grow: 1;
        }
        .value-display {
            min-width: 40px;
        }
        button {
            background-color: #4285f4;
            color: white;
            padding: 10px 20px;
            border: none;
            border-radius: 4px;
            cursor: pointer;
        }
        button:hover {
            background-color: #357abd;
        }
    </style>
</head>
<body>
    <div class="form-group">
        <label for="baseUrl">API Base URL:</label>
        <input type="text" id="baseUrl" placeholder="https://api.openai.com/v1">
    </div>

    <div class="form-group">
        <label for="model">Model:</label>
        <select id="model">
            <option value="gpt-3.5-turbo">GPT-3.5 Turbo</option>
            <option value="gpt-4">GPT-4</option>
            <option value="gpt-4-turbo-preview">GPT-4 Turbo</option>
            <option value="claude-3-sonnet">Claude-3-Sonnet</option>
            <option value="custom">Custom Model</option>
        </select>
        <input type="text" id="customModel" placeholder="Enter custom model name" style="display: none;">
    </div>

    <div class="form-group">
        <label for="temperature">Temperature:</label>
        <div class="slider-container">
            <input type="range" id="temperature" class="slider" min="0" max="1" step="0.1" value="0.7">
            <span id="temperatureValue" class="value-display">0.7</span>
        </div>
    </div>

    <div class="form-group">
        <label for="maxTokens">Max Tokens:</label>
        <input type="number" id="maxTokens" min="150" value="2000">
    </div>

    <button onclick="saveSettings()">Save Settings</button>

    <script>
        // Initialize values from stored settings
        function loadSettings() {
            google.script.run.withSuccessHandler(function(settings) {
                document.getElementById('baseUrl').value = settings.baseUrl || 'https://api.openai.com/v1';
                document.getElementById('model').value = settings.model || 'gpt-3.5-turbo';
                document.getElementById('temperature').value = settings.temperature || 0.7;
                document.getElementById('temperatureValue').textContent = settings.temperature || 0.7;
                document.getElementById('maxTokens').value = settings.maxTokens || 2000;
                
                if (settings.model === 'custom') {
                    document.getElementById('customModel').style.display = 'block';
                    document.getElementById('customModel').value = settings.customModel || '';
                }
            }).getSettings();
        }

        // Handle model selection change
        document.getElementById('model').addEventListener('change', function(e) {
            const customModelInput = document.getElementById('customModel');
            customModelInput.style.display = e.target.value === 'custom' ? 'block' : 'none';
        });

        // Update temperature display
        document.getElementById('temperature').addEventListener('input', function(e) {
            document.getElementById('temperatureValue').textContent = e.target.value;
        });

        // Save settings
        function saveSettings() {
            const settings = {
                baseUrl: document.getElementById('baseUrl').value,
                model: document.getElementById('model').value,
                customModel: document.getElementById('customModel').value,
                temperature: parseFloat(document.getElementById('temperature').value),
                maxTokens: parseInt(document.getElementById('maxTokens').value)
            };

            google.script.run
                .withSuccessHandler(function() {
                    google.script.host.close();
                })
                .withFailureHandler(function(error) {
                    alert('Error saving settings: ' + error);
                })
                .saveSettings(settings);
        }

        // Load settings when page loads
        window.onload = loadSettings;
    </script>
</body>
</html>

# Project File Structure

├── Code.gs                  # Main script file with core functionality
├── Sidebar.html            # UI for prompts and responses + settings panel
└── appsscript.json         # Project configuration (auto-generated)