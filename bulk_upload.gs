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
// Copyright (C) 2020, Josh Cain
// Copyright (C) 2020, Sol Boucher
// Copyright (C) 2020, The Vanguard Campaign Corps Mods (vanguardcampaign.org)

// @OnlyCurrentDoc
// Session.getEffectiveUser().getEmail()

const USE_STAGING = true;

const REQUESTS_SHEET = 'Requests';
const TYPES_SHEET = 'Request types';

const TYPE_COLUMN = 1;
const FILENAME_COLUMN = 2;
const CLAIM_COLUMN = 3;
const DONE_COLUMN = 4;

const ALLOWLIST = {
  'Date Called': true,
  'Voter Phone': true,
  
  // Type 'Thrutalk Spanish Voter ID Script Results':
  // TODO 8/5/20: Per Jonah, wait on revised Spanish script
  //'english_support': true,
  //'english volunteer ask': true,
  
  // Type 'Thrutalk Voter ID Script Results':
  'starting_question': true,
  'support': true,
  'undecided-support': true,
  'VBM Ask': true,
  'volunteer ask': true,
  'Already Voted Support': true,
};
const TAGS = [];

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
  
  // Terminate the file with a newline if there isn't one already.
  if(!csv.endsWith('\n'))
    csv += '\n';
  
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
    var action = ALLOWLIST[preserve[idx]];
    if(action)
      preserve[idx] = [idx, action];
  }
  preserve = Object.fromEntries(preserve.filter(Array.isArray));
  
  // Time to actually remove the unwanted columns, then add the vanId one.
  var vanidx = -1;
  var unkeyed = 0;
  var records = 0;
  var datadump = '';
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
      id = '0';
      ++unkeyed;
      datadump += '<br>' + JSON.stringify({ fields: fields, vanidx: vanidx, id: id });
    }
    ++records;
    return id + ',' + line;
  });
  
  if(unkeyed == records)
    return HtmlService.createHtmlOutput('Data file missing primary key column: \'' + vanid + '\':' + datadump)
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
