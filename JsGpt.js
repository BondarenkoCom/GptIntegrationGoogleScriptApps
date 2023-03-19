var spreadsheetId = '---';


function readGPTSettings() {
  //var spreadsheetId = '1duF8iQ7nPQ8Ej7NjW3SssIxYIIjaW5_HPxKSQNCR8wc'; // Replace with your spreadsheet ID

  // Access the sheet by its ID and the active sheet
  var sheet = SpreadsheetApp.openById(spreadsheetId).getActiveSheet();

  var gptSettings = {};

  for (var rangeSettingsCount = 1; rangeSettingsCount <= 10; rangeSettingsCount++) {
    var settingLetterRange = 'A' + rangeSettingsCount;
    var settingValueRange = 'B' + rangeSettingsCount;

    var settingsName = sheet.getRange(settingLetterRange).getValue();
    var settingsValue = sheet.getRange(settingValueRange).getValue();

    if (settingsName) {
      switch (settingsName) {
        case 'API':
          gptSettings.API = settingsValue;
          break;
        case 'FINETUNEID':
          gptSettings.FinetuneID = settingsValue;
          break;
        case 'MODELID':
          gptSettings.ModelID = settingsValue;
          break;
        case 'max temp':
          gptSettings.MaxTemp = parseFloat(settingsValue);
          break;
        case 'max tokens':
          gptSettings.MaxTokens = parseInt(settingsValue, 10);
          break;
        case 'Top P':
          gptSettings.TopP = parseFloat(settingsValue);
          break;
        case 'Best Of':
          gptSettings.BestOf = parseInt(settingsValue, 10);
          break;
        case 'engine':
          gptSettings.Engine = settingsValue;
          break;
        case 'model':
          gptSettings.Model = settingsValue;
          break;
        case 'OpenApIKey':
          gptSettings.OpenApiKey = settingsValue;
          break;
      }
    }
  }
  Logger.log(gptSettings);
  return gptSettings;
}

function initGPTAPI() {
  var gptSettings = readGPTSettings();

  var apiKey = gptSettings.FinetuneID;
  Logger.log('This api key ' + apiKey);

  var endpoint = gptSettings.API;
  Logger.log('This endPoint ' + endpoint);

  var modelAI = gptSettings.Model;
  Logger.log('This model AI - ' + modelAI);

  var maxTem = gptSettings.MaxTemp;
  Logger.log('This max temp - ' + maxTem);

  var maxTokens = gptSettings.MaxTokens;
  Logger.log('This max temp - ' + maxTokens);

  var getEngine = gptSettings.Engine;
  Logger.log('This max temp - ' + getEngine);
   
   var topP = gptSettings.TopP;
  Logger.log('This max temp - ' + getEngine);


  var contentFromDocs = getNumberOfFilesToUpload(); // Replace this with the appropriate function from your script
  Logger.log('This content for GPT ' + contentFromDocs);

  var promptGPT = getPropmtForGPT(); // Replace this with the appropriate function from your script
  Logger.log('This command for GPT - ' + promptGPT);

  var promptMessage = promptGPT + " -- " + contentFromDocs;
  Logger.log('This Prompt + content for GPT - ' + promptMessage);


   const options = {
    "max_temp": maxTem,
    "max_tokens": maxTokens,
    "top_p": topP,
    "engine": getEngine,
    "model": modelAI
  };

   const requestOptions = {
    "method": "POST",
    "headers": {
      "Authorization": `Bearer ${apiKey}`,
      "Content-Type": "application/json"
    },
    "payload": JSON.stringify({
      "model": options.model,
      "messages": [{"role": "system", "content": promptMessage}, {"role": "user", "content": promptMessage}],
      "temperature": options.max_temp,
      "max_tokens": options.max_tokens,
      "top_p": options.top_p
    })
  };

   try {
    const response = UrlFetchApp.fetch(endpoint, requestOptions);
    const jsonResponse = JSON.parse(response.getContentText());
    
    if (jsonResponse && jsonResponse.choices && jsonResponse.choices.length > 0) {
      const answer = jsonResponse.choices[0].message.content;
      Logger.log(`GPT Response: ${answer}`);
      pushValueToCell("C15", answer); 
    } else {
      Logger.log("No response from GPT.");
    }
  } catch (error) {
    Logger.log(`Error: ${error}`);
  }
}

function getNumberOfFilesToUpload() {
  var sheet = SpreadsheetApp.openById(spreadsheetId);
  var range = "E1";
  var responseFileCount = sheet.getRange(range).getValue();
  var numbersOfFiles = parseInt(responseFileCount, 10);
  var txtFileForGPT;

  for (var rangeCount = 1; rangeCount <= numbersOfFiles; rangeCount++) {
    try {
      var currentRange = "E" + rangeCount;
      var url = sheet.getRange(currentRange).getValue();
      //Logger.log('This url from sheets ' + url);

      if (url && url.includes("https:")) {
        try {
          txtFileForGPT = readDocsByLink(url); // Replace this with the appropriate function from your script
          
          Utilities.sleep(30000);

          //Logger.log('This text from Docs' + txtFileForGPT);
          return txtFileForGPT.toString();
        } catch (error) {
          console.warn("Invalid Google Docs URL - " + error.message);
        }
      }
    } catch (error) {
      console.error("Error in E ranges - " + error.message);
    }
  }
  return null;
}

function readDocsByLink(urlFromSheetCell) {
  var googleDocsUrl = urlFromSheetCell;
  //Logger.log('Url from cells - ' + googleDocsUrl)
  
  var content = downloadTextFileFromGoogleDrive(googleDocsUrl);
  //Logger.log('Content  - ' + content)

  return content;
}

function downloadTextFileFromGoogleDrive(googleDocsUrl) {

  //Logger.log('Url in method downloadTextFileFromGoogleDrive - ' + googleDocsUrl);

  var docIdRegex = /\/document\/d\/([\w-]+)/;
  var docIdMatch = googleDocsUrl.match(docIdRegex);

  if (!docIdMatch || !docIdMatch[1]) {
    Logger.log('Invalid Google Docs URL');
    return null;
  }

  var docId = docIdMatch[1];

  // Get the content of the document
  var doc = Docs.Documents.get(docId);
  var docContent = doc.body.content;

  // Extract the text from the content
  var contentText = '';
  docContent.forEach(function(element) {
    if (element.paragraph) {
      element.paragraph.elements.forEach(function(element) {
        if (element.textRun) {
          contentText += element.textRun.content;
        }
      });
    }
  });

  //Logger.log(contentText);
  return contentText;
}

function getPropmtForGPT() {
  var sheet = SpreadsheetApp.openById(spreadsheetId);
  var promptRange, outputRange;
  var promptText = '';

  for (var rangePromptCount = 14; rangePromptCount <= 100; rangePromptCount++) {
    try {
      promptRange = 'B' + rangePromptCount;
      outputRange = 'C' + rangePromptCount;

      var promptCell = sheet.getRange(promptRange).getValue();

      if (promptCell === '') {
        console.log('Stop loop');
        break;
      } else {
        promptText = promptCell.toString();
       
        if (promptText != null && !promptText.includes('PROMPT')) {
           Logger.log('This prompt from method getPropmtForGPT - ' + promptText);
          return promptText;
        }
      }
    } catch (ex) {
      console.log(ex.message);
      continue;
    }
  }
  return null;
}

function pushValueToCell(cellName, content) {
  //var spreadsheetId = 'your_spreadsheet_id'; // Replace with your spreadsheet ID

  var values = [
    [content]
  ];

  try {
    var sheet = SpreadsheetApp.openById(spreadsheetId);
    var range = sheet.getRange(cellName);
    
    range.setValues(values);
  } catch (error) {
    Logger.log('Error while sending data to Google Sheets: ' + error.message);
  }
}

