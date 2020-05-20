function doGet(request) {
  var bearer = request.parameter.access_token;
  if(!bearer)
    return ContentService.createTextOutput('Missing request parameter: access_token');
  
  var openid = JSON.parse(get('https://accounts.google.com/.well-known/openid-configuration'));
  var active = JSON.parse(get(openid.userinfo_endpoint, bearer));
  if(active.email != Session.getEffectiveUser().getEmail() || !active.email_verified)
    return ContentService.createTextOutput('Unauthorized user: ' + active.email);
  
  var tracker = request.parameter.tracker;
  if(!tracker)
    return ContentService.createTextOutput('Missing request parameter: tracker');
  
  var config = JSON.parse(lookup('trackers', tracker));
  if(!config)
    return ContentService.createTextOutput('Invalid request parameter: \'tracker=' + tracker + '\'');
  
  var sheet = SpreadsheetApp.openById(tracker).getSheets()[0];
  if(!sheet)
    return ContentService.createTextOutput('Unauthorized to access spreadsheet: ' + tracker);
  
  var record = request.parameter.record;
  if(!record)
    return ContentService.createTextOutput('Missing request parameter: record');
  
  var entry = sheet.getRange(record, 1, 1, sheet.getLastColumn()).getValues()[0];
  if(!entry)
    return ContentService.createTextOutput('Invalid reqest parameter: \'record=' + record + '\'');
  
  var expires = config.expiryColumn;
  if(typeof expires == 'number') {
    var unexpired = entry[expires];
    if(typeof unexpired != 'boolean' || !unexpired)
      return ContentService.createTextOutput('Link has expired');
  }
  
  var filename = entry[config.filenameColumn];
  var folder = DriveApp.getFolderById(config.folder);
  var file = folder.getFilesByName(filename);
  if(!file.hasNext())
    return ContentService.createTextOutput('No file with the name: \'' + filename + '\'');
  return ContentService.createTextOutput(file.next().getBlob().getDataAsString());
}

function lookup(table, key, fallback = null) {
  if(!key)
    return fallback;
  
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(table);
  if(sheet.getLastRow() <= 2)
    return fallback;
  
  var cell = sheet.getRange(2, 1, sheet.getLastRow() - 1)
                  .createTextFinder(key)
                  .matchEntireCell(true)
                  .findNext();
  if(!cell)
    return fallback;
  return sheet.getRange(cell.getRow(), 2).getValue();
}

function get(resource, authorization) {
  if(authorization)
    authorization = {
      headers: {
        Authorization: 'Bearer ' + authorization,
      },
    };
  return UrlFetchApp.fetch(resource, authorization).getContentText();
}
