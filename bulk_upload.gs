// @OnlyCurrentDoc
// Session.getEffectiveUser().getEmail()

const USE_STAGING = true;

const REQUESTS_SHEET = 'Requests';
const TYPES_SHEET = 'Request types';

const TYPE_COLUMN = 1;
const FILENAME_COLUMN = 2;
const CLAIM_COLUMN = 3;
const DONE_COLUMN = 4;

const WHITELIST = {
  'contact[cell]': true,
  'contact[optOut]': true,
  
  // Type 'Absentee ballot survey':
  'question[Do they see the link?]': true,
  'question[Finished?]': true,
  'question[Have they requested an absentee ballot yet?]': true,
  'question[Want additional help?]': true,
  'question[Will they actually vote in-person?]': true,
  
  // Type 'HubDialer voter ID':
  'Absentee Ballot': true,
  'Date': true,
  'Phone Number': true,
  'Support': true,
  'Status': true,
  
  // Type 'Persuasion Campaign':
  'question[RO Ask_1]': true,
  'question[RO Ask_2]': true,
  'question[RO Ask_3]': true,
  'question[RO Ask_4]': true,
  'question[RO Ask_5]': true,
  'question[VBM Issues?]': true,
  'question[Vote by mail?]': true,
  'question[Will they vote for Jamaal?]': true,
  
  // Type 'Third Outreach':
  'question[Will you reach out to 3-5 voters?]': true,
  
  // Type 'Undecided phonebank':
  'First Call Date': true,
  'Pledge': true,
  
  // Type 'Virtual town hall':
  'question[Will this person attend?]': true,
  'tags': true,
  
  // Type 'Westchester absentee ballot request':
  'DT_ADDED': transform(addColumnBoolean, 'DT_ADDED', 'DT_ADDED,ADDED'),
  'DT_REQUEST': transform(addColumnBoolean, 'DT_REQUEST', 'DT_REQUEST,REQUEST'),
  'DT_RETURN': transform(addColumnBoolean, 'DT_RETURN', 'DT_RETURN,RETURN'),
  'DT_MAILED': transform(addColumnBoolean, 'DT_MAILED', 'DT_MAILED,MAILED'),
  
  // Type 'WFP voter ID':
  'question[In District?]': true,
  'question[VBM Issues?_1]': true,
  'question[Vote by mail?_1]': true,
  'question[Voter Disposition]': true,
  
  // Type 'WFP voter ID consolidated':
  'cell': true,
  'date_updated': transform(truncateColumnSpace, 'date_updated'),
  'question_response': true,
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
  
  var self = 'self';
  var target = '';
  if(USE_STAGING) {
    self += '_staging';
    target = ' target="_blank"';
  }
  
  var href = PropertiesService.getScriptProperties().getProperty(self) + '?record=' + range.getRow();
  var link = HtmlService.createHtmlOutput('<a href="' + href + '"' + target + '>' + href + '</a>');
  SpreadsheetApp.getUi().showModalDialog(link, 'Click to download');
}

function doGet(ter) {
  var doc = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = doc.getSheetByName(REQUESTS_SHEET);
  var row = ter.parameter.record;
  var filename = sheet.getRange(row, FILENAME_COLUMN).getValue();
  var type = sheet.getRange(row, TYPE_COLUMN).getValue();
  var vanid = lookup(TYPES_SHEET, type);
  if(!vanid)
    return HtmlService.createHtmlOutput('Unsupported request type: \'' + type + '\'')
                      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  
  var csv = get(PropertiesService.getScriptProperties().getProperty('downloader') + '?'
                + 'access_token=' + ScriptApp.getOAuthToken() + '&'
                + 'tracker=' + doc.getId() + '&'
                + 'record=' + row);
  sheet.getRange(row, CLAIM_COLUMN).removeCheckboxes();
  
  if(filename.endsWith('.txt'))
    csv = csv.replace(/^([A-Z]*[0-9]+)[^\n]+[A-Z](\d{4})?(\d{2})?(\d{2})?[^\n]*$/mg, '$1,$2-$3-$4');
  if(csv.indexOf(',') == -1)
    return HtmlService.createHtmlOutput(csv)
                      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  if(filename.endsWith('.txt'))
    return serve(filename + '.csv', 'countyId,date\n' + csv);
  
  // Start with a good 'ol dos2unix.
  csv = csv.replace(/\r/g, '');
  
  // Add a dummy column on non-Spoke files to force the comma replacement
  // logic to apply to strings in the last column
  // TODO: In the future, we should not treat the last column specially
  if(!csv.match(/^[^\n]+,tags\n/)) {
    csv = csv.replace(/\n/g, ",\n");
  }
  
  // Remove double-double quotes since they compromise the comma replacement logic
  csv = csv.replace(/""/g, '');
  
  // And for my next trick, I'll disappear all strings except the last one.
  // This assumes that the tags column is the very last one!
  csv = csv.replace(/"([^"]+)",/g, function(match, quoted) {
    return quoted.replace(/,|\n/g, ';') + ',';
  });
  
  // Compile a list of the columns we'll be preserving. Note that we always "keep" vanId.
  var preserve = csv.match(/^.*$/m)[0].split(',');
  for(var idx = 0; idx < preserve.length; ++idx) {
    var action = WHITELIST[preserve[idx]];
    if(action)
      preserve[idx] = [idx, action];
  }
  preserve = Object.fromEntries(preserve.filter(Array.isArray));
  
  // Time to actually remove the unwanted columns, then add the vanId one.
  var vanidx = -1;
  var unkeyed = false;
  var datadump = {};
  csv = csv.replace(/^[^"\n]+/mg, function(line) {
    var fields = line.split(',');
    var line = [];
    for(var idx = 0; idx < fields.length; ++idx) {
      var transform = preserve[String(idx)];
      if(typeof transform == 'function')
        line.push(transform(fields[idx]));
      else if(transform)
        line.push(fields[idx]);
    }
    line = line.join(',');
    
    if(vanidx == -1) {
      vanidx = fields.indexOf(vanid);
      return 'vanId,' + line;
    }
    
    var id = fields[vanidx];
    if(id)
      id = id.replace(/^.+(\d{10})$/, '$1');
    else if(fields[0]) {
      unkeyed = true;
      datadump = { fields: fields, vanidx: vanidx, id: id };
    }
    return id + ',' + line;
  });
  
  if(unkeyed)
    return HtmlService.createHtmlOutput('Data file missing primary key column: \'' + vanid + '\'' + JSON.stringify(datadump))
                      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  
  var date = filename.match(/[0-9]{4}[-_][0-9]{2}[-_][0-9]{2}/);
  if(date)
    date = date[0].replace(/_/g, '-');
  if(!csv.match(/^[^\n]+,tags\n/)) {
    if(date) {
      var labelled = false;
      csv = csv.replace(/\n/g, function() {
        if(!labelled) {
          labelled = true;
          return ',date\n';
        }
        return ',' + date + '\n';
      });
    }
    return serve(filename, csv);
  }
  
  // Add a date column parsed from the filename and explode tags into multiple columns.
  var skipped = false;
  csv = csv.replace(/tags$/m, 'date,tag[' + TAGS.join('],tag[') + ']');
  csv = csv.replace(/,([^,\n]+|"[^"]+")?$/mg, function(match, stripped) {
    if(!skipped) {
      skipped = true;
      return match;
    } else if(!stripped)
      return ',' + date + ','.repeat(TAGS.length);
    
    var repl = ',' + date;
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
  
  return serve(filename, csv);
}

function transform(transform, oldlabel, newlabel = oldlabel) {
  return function(value) {
    if(value == oldlabel)
      return newlabel;
    return transform(value);
  }
}

function addColumnBoolean(value) {
  return value + ',' + Boolean(value);
}

function truncateColumnSpace(value) {
  return value.split(' ')[0];
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

function serve(filename, contents, mime = ContentService.MimeType.CSV) {
  return ContentService.createTextOutput(contents)
                       .setMimeType(mime)
                       .downloadAsFile(filename.replace(/\(|\)|!|#/g, ''));
}
