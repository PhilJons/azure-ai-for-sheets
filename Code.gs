// Initialize global rate limits and execution tracking
const RATE_LIMITS = {
  requests: {
    maxPerMinute: 120,     // Target 120 RPM (well below the 180 limit for safety)
    used: 0,
    lastReset: Date.now(),
    lastRequest: Date.now()
  },
  backoff: {
    initialDelay: 2000,    // Start with 2 second delay
    maxDelay: 32000,       // Max 32 second delay
    attempts: 0
  }
};

// Initialize execution tracking
let executionStartTime = Date.now();
const MAX_EXECUTION_TIME = 270000; // 4.5 minutes

// Document properties for persistence
const documentProperties = PropertiesService.getDocumentProperties();

/**
 * Makes Azure API call with retry logic and rate limiting
 */
function callAzureWithRetry(text, systemPrompt, temperature, config, retryAttempt = 0) {
  const maxRetries = 5;
  const baseDelay = 2000;

  // Check execution time
  if (Date.now() - executionStartTime > MAX_EXECUTION_TIME) {
    createContinuationTrigger();
    return "⌛ Loading...";
  }

  try {
    checkAndResetLimits();
    
    const now = Date.now();
    const minSpacing = 500;
    const timeSinceLastRequest = now - RATE_LIMITS.requests.lastRequest;
    
    if (timeSinceLastRequest < minSpacing) {
      Utilities.sleep(minSpacing - timeSinceLastRequest);
    }
    
    if (RATE_LIMITS.backoff.attempts > 0) {
      const backoffTime = Math.min(
        RATE_LIMITS.backoff.initialDelay * Math.pow(2, RATE_LIMITS.backoff.attempts - 1),
        RATE_LIMITS.backoff.maxDelay
      );
      Utilities.sleep(backoffTime);
    }

    RATE_LIMITS.requests.lastRequest = Date.now();
    
    const response = UrlFetchApp.fetch(`${config.endpoint.replace(/\/$/, '')}`, {
      method: 'post',
      headers: {
        'Content-Type': 'application/json',
        'api-key': config.apiKey,
        'Cache-Control': 'no-cache'
      },
      muteHttpExceptions: true,
      payload: JSON.stringify({
        messages: [
          ...(systemPrompt ? [{role: 'system', content: systemPrompt}] : []),
          {role: 'user', content: text}
        ],
        temperature: temperature,
        max_tokens: 800
      })
    });

    const responseCode = response.getResponseCode();
    
    if (responseCode === 429) {
      RATE_LIMITS.backoff.attempts++;
      if (retryAttempt < maxRetries) {
        const retryDelay = Math.min(baseDelay * Math.pow(2, retryAttempt), 32000);
        Utilities.sleep(retryDelay);
        return callAzureWithRetry(text, systemPrompt, temperature, config, retryAttempt + 1);
      }
      return "⌛ Loading... (Rate Limited)";
    }

    if (responseCode >= 500) {
      if (retryAttempt < maxRetries) {
        const retryDelay = Math.min(baseDelay * Math.pow(2, retryAttempt), 32000);
        Utilities.sleep(retryDelay);
        return callAzureWithRetry(text, systemPrompt, temperature, config, retryAttempt + 1);
      }
    }

    if (responseCode !== 200) {
      throw new Error(`HTTP ${responseCode}: ${response.getContentText()}`);
    }

    RATE_LIMITS.requests.used++;
    RATE_LIMITS.backoff.attempts = Math.max(0, RATE_LIMITS.backoff.attempts - 1);
    
    const result = JSON.parse(response.getContentText());
    if (result && result.choices && result.choices.length > 0) {
      return result.choices[0].message.content;
    }
    throw new Error('Invalid response format from Azure OpenAI');
    
  } catch (error) {
    if (retryAttempt < maxRetries) {
      const retryDelay = Math.min(baseDelay * Math.pow(2, retryAttempt), 32000);
      Utilities.sleep(retryDelay);
      return callAzureWithRetry(text, systemPrompt, temperature, config, retryAttempt + 1);
    }
    throw error;
  }
}

/**
 * Custom function to analyze text using Azure AI with loading status
 */
function AZURE_ANALYZE_TEXT(text, systemPrompt = '', temperature = 0.7) {
  // Reset execution timer if it's expired
  if (Date.now() - executionStartTime > MAX_EXECUTION_TIME) {
    executionStartTime = Date.now();
  }

  const config = loadAzureConfig();
  if (!config.endpoint || !config.apiKey) {
    throw new Error('Please configure Azure AI settings first');
  }

  try {
    const response = callAzureWithRetry(text, systemPrompt, temperature, config);
    return response;
  } catch (error) {
    return "⌛ Loading...";
  }
}

/**
 * Reset rate limit counters if a minute has passed
 */
function checkAndResetLimits() {
  const now = Date.now();
  if (now - RATE_LIMITS.requests.lastReset >= 60000) {
    RATE_LIMITS.requests.used = 0;
    RATE_LIMITS.requests.lastReset = now;
  }
}

/**
 * Creates a continuation trigger if not already exists
 */
function createContinuationTrigger() {
  const triggerId = documentProperties.getProperty("timeOutTriggerId");
  if (!triggerId) {
    const trigger = ScriptApp.newTrigger("continueProcessing")
      .timeBased()
      .everyMinutes(1)
      .create();
    documentProperties.setProperty("timeOutTriggerId", trigger.getUniqueId());
  }
}

/**
 * Continues processing from where we left off
 */
function continueProcessing() {
  executionStartTime = Date.now();  // Reset execution timer
  const triggerId = documentProperties.getProperty("timeOutTriggerId");
  
  if (triggerId) {
    // Clear trigger if all processing is complete
    const sheet = SpreadsheetApp.getActiveSheet();
    const range = sheet.getDataRange();
    const values = range.getValues();
    const formulas = range.getFormulas();
    let hasErrors = false;

    // Check if any cells still need processing
    for (let i = 0; i < values.length; i++) {
      for (let j = 0; j < values[i].length; j++) {
        if (values[i][j] === '#ERROR!' || values[i][j] === '#ERROR' || 
            values[i][j] === '⌛ Loading...' || 
            (formulas[i][j].includes('AZURE_ANALYZE_TEXT') && values[i][j] === '')) {
          hasErrors = true;
          break;
        }
      }
      if (hasErrors) break;
    }

    if (!hasErrors) {
      // All done, clean up
      ScriptApp.getProjectTriggers()
        .filter(trigger => trigger.getUniqueId() === triggerId)
        .forEach(trigger => ScriptApp.deleteTrigger(trigger));
      documentProperties.deleteProperty("timeOutTriggerId");
    }
  }
}

// Keep existing menu and configuration functions
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Azure AI')
    .addItem('Configure Azure Settings', 'showConfigDialog')
    .addItem('Reload Error Cells', 'reloadErrorCells')
    .addItem('About', 'showAboutDialog')
    .addToUi();
}

function showConfigDialog() {
  const html = HtmlService.createHtmlOutputFromFile('ConfigDialog')
    .setWidth(400)
    .setHeight(300);
  SpreadsheetApp.getUi().showModalDialog(html, 'Azure AI Configuration');
}

function saveAzureConfig(endpoint, apiKey) {
  const userProperties = PropertiesService.getUserProperties();
  userProperties.setProperties({
    'azureEndpoint': endpoint,
    'azureApiKey': apiKey
  });
  return true;
}

function loadAzureConfig() {
  const userProperties = PropertiesService.getUserProperties();
  return {
    endpoint: userProperties.getProperty('azureEndpoint'),
    apiKey: userProperties.getProperty('azureApiKey')
  };
}

function showAboutDialog() {
  const ui = SpreadsheetApp.getUi();
  ui.alert(
    'About Azure AI Add-on',
    'This add-on provides integration with Azure AI services for text analysis.\n\n' +
    'Version: 1.0\n' +
    'Created with AI by: Aura AI Taskforce',
    ui.ButtonSet.OK
  );
}