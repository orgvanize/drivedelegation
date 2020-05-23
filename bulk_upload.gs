// @OnlyCurrentDoc
// Session.getEffectiveUser().getEmail()

const REQUESTS_SHEET = 'Requests';
const TYPE_COLUMN = 1;
const FILENAME_COLUMN = 2;
const CLAIM_COLUMN = 3;
const DONE_COLUMN = 4;

const WHITELIST = {
  'contact[cell]': true,
  'contact[optOut]': true,
  
  // Type 'Absentee ballot survey':
  'question[Do they see the link?]': true,
  'question[Done submitting?]': true,
  'question[Have they requested an absentee ballot yet?]': true,
  'question[Want additional help?]': true,
  
  // Type 'Virtual town hall':
  'question[Will this person attend?]': true,
  'tags': true,
};
const TAGS = [
  'Absentee - Will get to it later',
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
  
  var href = PropertiesService.getScriptProperties().getProperty('self') + '?record=' + range.getRow();
  var link = HtmlService.createHtmlOutput('<a href="' + href + '">' + href + '</a>');
  SpreadsheetApp.getUi().showModalDialog(link, 'Click to download');
}

function doGet(ter) {
  var doc = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = doc.getSheetByName(REQUESTS_SHEET);
  var row = ter.parameter.record;
  var filename = sheet.getRange(row, FILENAME_COLUMN).getValue();
  var type = sheet.getRange(row, TYPE_COLUMN).getValue();
  var vanid = lookup('Request types', type);
  if(!vanid)
    return HtmlService.createHtmlOutput('Unsupported request type: \'' + type + '\'')
                      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  
  var csv = get(PropertiesService.getScriptProperties().getProperty('downloader') + '?'
                + 'access_token=' + ScriptApp.getOAuthToken() + '&'
                + 'tracker=' + doc.getId() + '&'
                + 'record=' + row);
  sheet.getRange(row, CLAIM_COLUMN).removeCheckboxes();
  if(csv.indexOf(',') == -1)
    return HtmlService.createHtmlOutput(csv).setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  
  // Start with a good 'ol dos2unix.
  csv = csv.replace(/\r/g, '');
  
  // And for my next trick, I'll disappear all strings except the last one.
  // This assumes that the tags column is the very last one!
  csv = csv.replace(/"([^"]+)",/g, function(match, quoted) {
    return quoted.replace(/,/g, ';') + ',';
  });
  
  // Compile a list of the columns we'll be preserving. Note that we always "keep" vanId.
  var preserve = csv.match(/^.*$/m)[0].split(',');
  for(var idx = 0; idx < preserve.length; ++idx)
    if(WHITELIST[preserve[idx]])
      preserve[idx] = idx;
  preserve.filter(function(each) {
    return typeof each == 'number';
  });
  
  // Time to actually remove the unwanted columns, then add the vanId one.
  var vanidx;
  csv = csv.replace(/^[^"\n]+/mg, function(line) {
    var fields = line.split(',');
    line = fields.filter(function(match, idx) {
      return preserve.includes(idx);
    }).join(',');
    
    if(!vanidx) {
      vanidx = fields.indexOf(vanid);
      return 'vanId,' + line;
    }
    
    return fields[vanidx] + ',' + line;
  });
  
  // Add a date column parsed from the filename and explode tags into multiple columns.
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
