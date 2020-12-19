// This program is free software: you can redistribute it and/or modify
// it under the terms of the GNU General Public License as published by
// the Free Software Foundation, either version 3 of the License, or
// (at your option) any later version.
//
// This program is distributed in the hope that it will be useful,
// but WITHOUT ANY WARRANTY; without even the implied warranty of
// MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
// GNU General Public License for more details.
//
// You should have received a copy of the GNU General Public License
// along with this program.  If not, see <https://www.gnu.org/licenses/>.
//
// Copyright (C) 2020, Sol Boucher
// Copyright (C) 2020, The Vanguard Campaign Corps Mods (vanguardcampaign.org)

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
  var files = folder.getFilesByName(filename);
  var file;
  var contents;
  if(!files.hasNext()
     || (file = files.next()).getName() != filename
     || (contents = file.getBlob().getDataAsString()).startsWith('%PDF'))
    return ContentService.createTextOutput('No file with the name: \'' + filename + '\'');
  
  return ContentService.createTextOutput(contents);
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
