// @OnlyCurrentDoc
// Session.getEffectiveUser().getEmail()

const REQUESTS_SHEET = 'Requests';
const TYPE_COLUMN = 1;
const FILENAME_COLUMN = 2;
const CLAIM_COLUMN = 3;
const DONE_COLUMN = 4;

const VANID = 'contact[external_id]';
const TAGS = [
  'Angry/Refused',
  'Engel Supporter',
  'Going to Attend',
  'Maybe',
  'Moved',
  'Petition',
  'Spanish',
  'Supporter/Not going to attend',
  'Trump supporter',
  
  // The last tag is special: if set, all others will be cleared!
  'Wrong Number',
];

function doEdit(ed) {
  var range = ed.range;
  if(range.getSheet().getName() != REQUESTS_SHEET
     || range.getNumRows() != 1 || range.getNumColumns() != 1)
    return;
  else if(range.getColumn() == DONE_COLUMN) {
    var claim = range.getSheet().getRange(range.getRow(), CLAIM_COLUMN);
    if(claim.getValue() == '')
      claim.setValue('??');
    return;
  } else if(range.getColumn() != CLAIM_COLUMN)
    return;
  else if(!ed.value) {
    range.insertCheckboxes();
    return;
  } else if(!range.isChecked())
    return;
  
  var type = range.getSheet().getRange(range.getRow(), TYPE_COLUMN).getValue();
  if(type != 'Virtual town hall') {
    range.uncheck();
    SpreadsheetApp.getUi().alert('Unsupported request type: \'' + type + '\'');
    return;
  }
  
  var href = PropertiesService.getScriptProperties().getProperty('self') + '?record=' + range.getRow();
  var link = HtmlService.createHtmlOutput('<a href="' + href + '" target="_blank">' + href + '</a>');
  SpreadsheetApp.getUi().showModalDialog(link, 'Click to download');
}

function doGet(ter) {
  var doc = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = doc.getSheetByName(REQUESTS_SHEET);
  var row = ter.parameter.record;
  var filename = sheet.getRange(row, FILENAME_COLUMN).getValue();
  var csv = get(PropertiesService.getScriptProperties().getProperty('downloader') + '?'
                + 'access_token=' + ScriptApp.getOAuthToken() + '&'
                + 'tracker=' + doc.getId() + '&'
                + 'record=' + row);
  sheet.getRange(row, CLAIM_COLUMN).removeCheckboxes();
  if(csv.indexOf(',') == -1)
    return ContentService.createTextOutput(csv);
  
  csv = csv.replace(/\r/g, '');
  
  var vanidx;
  csv = csv.replace(/^.*$/mg, function(line) {
    if(!vanidx) {
      vanidx = line.split(',').indexOf(VANID);
      return 'vanId,' + line;
    }
    
    return line.split(',')[vanidx] + ',' + line;
  });
  
  var date = ',' + filename.match(/[0-9]{4}-[0-9]{2}-[0-9]{2}/);
  var skipped = false;
  csv = csv.replace(/tags$/m, 'date,tag[' + TAGS.join('],tag[') + ']');
  csv = csv.replace(/,([^,\n]+|"[^"]+")?$/mg, function(match, stripped) {
    if(!skipped) {
      skipped = true;
      return match;
    } else if(!stripped)
      return date + ','.repeat(TAGS.length);
    
    var repl = date;
    var tags = stripped.replace(/"/g, '').split(',');
    if(tags.includes(TAGS[TAGS.length - 1]))
      return repl + ','.repeat(TAGS.length) + 'true';
    
    for(var tag of TAGS) {
      repl += ',';
      if(tags.includes(tag))
        repl += 'true';
    }
    return repl;
  });
  
  return ContentService.createTextOutput(csv)
                       .setMimeType(ContentService.MimeType.CSV)
                       .downloadAsFile(filename.replace(/\(/g, '').replace(/\)/g, '').replace(/!/g, ''));
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
