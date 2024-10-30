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