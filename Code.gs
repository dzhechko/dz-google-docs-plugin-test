// Your OpenAI API Key - store this in Script Properties for security
const OPENAI_API_KEY = PropertiesService.getScriptProperties().getProperty('OPENAI_API_KEY');

// Get settings from Script Properties or use defaults
function getSettings() {
  const scriptProperties = PropertiesService.getScriptProperties();
  return {
    baseUrl: scriptProperties.getProperty('baseUrl') || 'https://api.openai.com/v1',
    model: scriptProperties.getProperty('model') || 'gpt-3.5-turbo',
    temperature: parseFloat(scriptProperties.getProperty('temperature')) || 0.7,
    maxTokens: parseInt(scriptProperties.getProperty('maxTokens')) || 2000,
    customModel: scriptProperties.getProperty('customModel') || ''
  };
}

// Save settings to Script Properties
function saveSettings(settings) {
  const scriptProperties = PropertiesService.getScriptProperties();
  scriptProperties.setProperties({
    'baseUrl': settings.baseUrl,
    'model': settings.model,
    'temperature': settings.temperature.toString(),
    'maxTokens': settings.maxTokens.toString(),
    'customModel': settings.customModel
  });
  return true;
}

// Main function to call OpenAI API
function callOpenAI(prompt) {
  const settings = getSettings();
  const apiUrl = `${settings.baseUrl}/chat/completions`;
  
  const headers = {
    'Authorization': 'Bearer ' + OPENAI_API_KEY,
    'Content-Type': 'application/json'
  };
  
  const payload = {
    'model': settings.model === 'custom' ? settings.customModel : settings.model,
    'messages': [
      {
        'role': 'user',
        'content': prompt
      }
    ],
    'temperature': settings.temperature,
    'max_tokens': settings.maxTokens
  };
  
  const options = {
    'method': 'post',
    'headers': headers,
    'payload': JSON.stringify(payload),
    'muteHttpExceptions': true
  };
  
  try {
    const response = UrlFetchApp.fetch(apiUrl, options);
    const json = JSON.parse(response.getContentText());
    
    if (json.error) {
      throw new Error(json.error.message || 'Unknown API Error');
    }
    
    return json.choices[0].message.content;
  } catch (error) {
    Logger.log('Error: ' + error);
    throw new Error('API Error: ' + error.message);
  }
}

// Test function to verify API connection
function testOpenAI() {
  try {
    const prompt = "What is the capital of France?";
    const response = callOpenAI(prompt);
    Logger.log('Test Response: ' + response);
    return response;
  } catch (error) {
    Logger.log('Test Error: ' + error);
    throw error;
  }
}

// Create menu when the document opens
function onOpen() {
  DocumentApp.getUi()
    .createMenu('GPT Помощник')
    .addItem('Показать панель', 'showSidebar')
    .addSeparator()
    .addItem('Сделать краткое содержание', 'summarizeSelection')
    .addItem('Улучшить текст', 'improveWriting')
    .addItem('Перевести на английский', 'translateToEnglish')
    .addSeparator()
    .addSubMenu(DocumentApp.getUi().createMenu('Форматирование')
      .addItem('Исправить грамматику', 'fixGrammar')
      .addItem('Формальный стиль', 'makeFormal')
      .addItem('Разговорный стиль', 'makeCasual'))
    .addSeparator()
    .addItem('Настройки', 'showSettings')
    .addToUi();
}

// Helper function to get selected text
function getSelectedText() {
  const doc = DocumentApp.getActiveDocument();
  const selection = doc.getSelection();
  
  if (selection) {
    const elements = selection.getSelectedElements();
    return elements.map(element => {
      const elementText = element.getElement().asText().getText();
      const startOffset = element.getStartOffset();
      const endOffset = element.getEndOffsetInclusive();
      
      if (startOffset !== null && endOffset !== null) {
        return elementText.substring(startOffset, endOffset + 1);
      }
      return elementText;
    }).join(' ');
  }
  
  // If no selection, try to get cursor position and current paragraph
  const cursor = doc.getCursor();
  if (cursor) {
    const element = cursor.getElement();
    return element.asText().getText();
  }
  
  throw new Error('Please select some text first.');
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
    return;
  }
  
  // If no selection, insert at cursor position
  const cursor = doc.getCursor();
  if (cursor) {
    const element = cursor.getElement();
    const position = cursor.getOffset();
    element.asText().insertText(position, newText);
    return;
  }
  
  throw new Error('No selection or cursor position found.');
}

// Add this new helper function
function insertTextBelow(newText, title) {
  const doc = DocumentApp.getActiveDocument();
  const selection = doc.getSelection();
  let targetElement;
  
  if (selection) {
    // Get the last selected element
    const elements = selection.getSelectedElements();
    const lastElement = elements[elements.length - 1];
    targetElement = lastElement.getElement();
  } else {
    // If no selection, get cursor position
    const cursor = doc.getCursor();
    if (cursor) {
      targetElement = cursor.getElement();
    } else {
      throw new Error('No selection or cursor position found.');
    }
  }
  
  // Find the parent paragraph
  while (targetElement.getType() !== DocumentApp.ElementType.PARAGRAPH) {
    targetElement = targetElement.getParent();
  }
  
  // Get the parent container (body or table cell)
  const container = targetElement.getParent();
  const index = container.getChildIndex(targetElement);
  
  // Insert title and new text in bold
  const titleParagraph = container.insertParagraph(index + 1, '');
  titleParagraph.appendText('\n' + title).setBold(true);
  
  // Insert the content
  const contentParagraph = container.insertParagraph(index + 2, newText);
  contentParagraph.appendText('\n'); // Add extra line break
  
  return true;
}

// Add language detection helper function
function detectLanguage(text) {
  try {
    const language = LanguageApp.detect(text);
    return language;
  } catch (error) {
    Logger.log('Language detection error: ' + error);
    return 'en'; // Default to English if detection fails
  }
}

// Add localized prompts helper function
function getPromptInLanguage(action, text) {
  // Instruct GPT to respond in Russian regardless of input language
  const systemPrompt = `Please process the following text and respond in Russian. `;
  
  const prompts = {
    'summarize': systemPrompt + `Summarize this text concisely:\n\n${text}`,
    'improve': systemPrompt + `Improve this text, making it more professional and clear while maintaining its meaning:\n\n${text}`,
    'grammar': systemPrompt + `Fix any grammar issues in this text:\n\n${text}`,
    'formal': systemPrompt + `Rewrite this text in a formal, professional tone:\n\n${text}`,
    'casual': systemPrompt + `Rewrite this text in a casual, friendly tone:\n\n${text}`
  };

  return prompts[action];
}

// Update the menu action functions
function summarizeSelection() {
  try {
    const selectedText = getSelectedText();
    const prompt = getPromptInLanguage('summarize', selectedText);
    const summary = callOpenAI(prompt);
    insertTextBelow(summary, 'Summary:');
  } catch (error) {
    DocumentApp.getUi().alert('Error: ' + error.toString());
  }
}

function improveWriting() {
  try {
    const selectedText = getSelectedText();
    const prompt = getPromptInLanguage('improve', selectedText);
    const improved = callOpenAI(prompt);
    insertTextBelow(improved, 'Improved Version:');
  } catch (error) {
    DocumentApp.getUi().alert('Error: ' + error.toString());
  }
}

function fixGrammar() {
  try {
    const selectedText = getSelectedText();
    const prompt = getPromptInLanguage('grammar', selectedText);
    const fixed = callOpenAI(prompt);
    insertTextBelow(fixed, 'Grammar Fixed Version:');
  } catch (error) {
    DocumentApp.getUi().alert('Error: ' + error.toString());
  }
}

function makeFormal() {
  try {
    const selectedText = getSelectedText();
    const prompt = getPromptInLanguage('formal', selectedText);
    const formal = callOpenAI(prompt);
    insertTextBelow(formal, 'Formal Version:');
  } catch (error) {
    DocumentApp.getUi().alert('Error: ' + error.toString());
  }
}

function makeCasual() {
  try {
    const selectedText = getSelectedText();
    const prompt = getPromptInLanguage('casual', selectedText);
    const casual = callOpenAI(prompt);
    insertTextBelow(casual, 'Casual Version:');
  } catch (error) {
    DocumentApp.getUi().alert('Error: ' + error.toString());
  }
}

// Function to show settings dialog (we'll implement this later with the HTML)
function showSettings() {
  const html = HtmlService.createHtmlOutputFromFile('Settings')
    .setWidth(400)
    .setHeight(500);
  DocumentApp.getUi().showModalDialog(html, 'GPT Extension Settings');
}

// Function to show sidebar (we'll implement this later with the HTML)
function showSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('Sidebar')
    .setTitle('GPT Extension');
  DocumentApp.getUi().showSidebar(html);
}

// Replace translateToThai with translateToEnglish
function translateToEnglish() {
  try {
    const selectedText = getSelectedText();
    const prompt = `Please translate the following text to English, maintaining the original meaning and tone:\n\n${selectedText}`;
    const translated = callOpenAI(prompt);
    insertTextBelow(translated, 'English Translation:');
  } catch (error) {
    DocumentApp.getUi().alert('Error: ' + error.toString());
  }
}

// Add this function to handle sidebar requests
function processTextWithGPT(text, action) {
  switch(action) {
    case 'summarize':
      return callOpenAI(getPromptInLanguage('summarize', text));
    case 'improve':
      return callOpenAI(getPromptInLanguage('improve', text));
    case 'grammar':
      return callOpenAI(getPromptInLanguage('grammar', text));
    case 'translate':
      return callOpenAI(`Please translate the following text to English, maintaining the original meaning and tone:\n\n${text}`);
    case 'formal':
      return callOpenAI(getPromptInLanguage('formal', text));
    case 'casual':
      return callOpenAI(getPromptInLanguage('casual', text));
    default:
      throw new Error('Invalid action specified');
  }
} 