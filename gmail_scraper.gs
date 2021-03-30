const SPREADSHEET_ID = '1JODYgPEmkpu7GsSKa_9VSdaweTXILYqOld1xNR5eCTE';
const SHEET_NAME = 'Requests';
const DRIVE_FOLDER_ID = '1MkcuhcRemv1Pp1fdu3PeAqJ7MB3i2GUi';
const NOTIFICATION_EMAIL = 'josh.cain@gmail.com';

// I created this from the morsebox@boucher-johnson.net account but it looks like scripts are shared between all accounts
// That's probably fine; I ran `createEmailTrigger` from this account 
function scrapeEmail() {
  try { 
    var urls = [];
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(SHEET_NAME);
  
    var label = GmailApp.getUserLabelByName("VANguard Bulk Upload");
    var threads = label.getThreads();
    
    for (let currentThread of threads) {
      if (!currentThread) {
        Logger.log('No valid message found');
        continue;
      }

      for (let msg of currentThread.getMessages()) {
        Logger.log('Parsing ' + msg.getSubject());
        
        var regExp = /(https:\/\/\S+script_results\S+\.csv)/;
        var match = regExp.exec(msg.getPlainBody()); 
        
        if (!match) {
          Logger.log('No link found in mail');
        } else {
          var url = match[1];
          var content = UrlFetchApp.fetch(url).getContentText();
          var fileName = url.replace('https://dialer-reports.s3.amazonaws.com/', '');
          var folder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
          folder.createFile(fileName, content, 'text/csv');
          
          let type = 'Thrutalk Voter ID Script Results';
          if (url.indexOf('spanish') > -1) {
            type = 'Thrutalk Spanish Voter ID Script Results';
          }
          else if (url.indexOf('special') > -1) {
            //type = 'Thrutalk Special Projects Script Results';
            type = '(Not yet ready!)';
          }
          
          sheet.appendRow([type, fileName, false, false]);
          sheet.getRange(sheet.getLastRow(), 3, 1, 2).insertCheckboxes();
          
          Logger.log('Added URL ' + url);
          urls.push(url);
        }
      }
  
      const newLabel = GmailApp.createLabel('VANguard Completed Bulk Upload');
      
      currentThread.addLabel(newLabel);
      currentThread.removeLabel(label);
    }

    GmailApp.sendEmail(NOTIFICATION_EMAIL, 'VANguard bulk upload tracker', urls.join("\n")); 
  } catch (err) {
    GmailApp.sendEmail(NOTIFICATION_EMAIL, 'VANguard bulk upload extractor error: ' + err.name, err.message + "\n\n" + err.stack);
    return;
  }
}

function createEmailTrigger() {
  ScriptApp.newTrigger('scrapeEmail').timeBased().everyDays(1).atHour(19).create();
}