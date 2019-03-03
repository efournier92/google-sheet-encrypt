var GLOBALID = "EncryptedSheet";

function clearDb() {
  var prop = PropertiesService.getUserProperties();
  if (prop.getProperty("sheetencrypted-state-" + GLOBALID) != null) {
    prop.deleteProperty("sheetencrypted-state-" + GLOBALID);
  }
  if (prop.getProperty("sheetencrypted-password-" + GLOBALID) != null) {
    prop.deleteProperty("sheetencrypted-password-" + GLOBALID);
  }
  if (prop.getProperty("sheetencrypted-id-" + GLOBALID) != null) {
    prop.deleteProperty("sheetencrypted-id-" + GLOBALID);
  }
}

function showChangePasswordForm() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  spreadsheet.show(HtmlService.createHtmlOutputFromFile('password-change'));
}

function changePassword(obj) {
  Logger.log(obj.oldpassword);

  var prop = PropertiesService.getUserProperties();
  if (prop.getProperty("sheetencrypted-password-" + GLOBALID) != null) {
    if (prop.getProperty("sheetencrypted-password-" + GLOBALID) != obj.oldpassword) {
      return ({ 'status': 'notmatching' });
    }
  }

  prop.setProperty("sheetencrypted-password-" + GLOBALID, obj.newpassword);

  return ({ 'status': 'done' });
}

function checkstate1() {
  var prop = PropertiesService.getUserProperties();
  Logger.log("State - " + prop.getProperty("sheetencrypted-state-" + GLOBALID));
  Logger.log("Id - " + prop.getProperty("sheetencrypted-id-" + GLOBALID));
  Logger.log("Password - " + prop.getProperty("sheetencrypted-password-" + GLOBALID));

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var id = ss.getActiveSheet().getSheetId();
  Logger.log(DriveApp.getFileById(SpreadsheetApp.getActiveSpreadsheet().getId()).getUrl() + "&gid=" + id);
  SpreadsheetApp.getActiveSpreadsheet().getSheets()[0].getRange('C1').setValue(ScriptApp.getService().getUrl());
}

function EncodeFromSheet() {
  Logger.log("Starting EncodeFromSheet");
  var prop = PropertiesService.getUserProperties();
  var encrypted = prop.getProperty("sheetencrypted-state-" + GLOBALID);
  if (encrypted == 2) {
    Browser.msgBox('ATTENTION', 'The sheet is already encrypted!!', Browser.Buttons.OK);
    return;
  }
  Logger.log("Sheet is un-encrypted. Proceeding.");

  var password = '';
  if (prop.getProperty("sheetencrypted-password-" + GLOBALID) == null) {
    Logger.log("Got null password, asking for one");
    password = Browser.inputBox("Create a new password.", Browser.Buttons.OK_CANCEL);
    if (password == 'cancel') {
      return;
    }
    prop.setProperty("sheetencrypted-password-" + GLOBALID, password);
    prop.setProperty("sheetencrypted-id-" + GLOBALID, SpreadsheetApp.getActiveSpreadsheet().getId());
    prop.setProperty("sheetencrypted-url-" + GLOBALID, DriveApp.getFileById(SpreadsheetApp.getActiveSpreadsheet().getId()).getUrl());
    Logger.log("Going to encoding after getting password");
    EnCodeSheet(false);
    getWebAppUrl();
  }
  else {
    Logger.log("Got password. Encrypting");
    SpreadsheetApp.getActiveSpreadsheet().show(HtmlService.createHtmlOutputFromFile('password-encrypt'));
  }
}
function encodeForRequest(obj) {
  var prop = PropertiesService.getUserProperties();
  Logger.log("|" + obj.password + "|" + prop.getProperty("sheetencrypted-password-" + GLOBALID) + "|");
  if (prop.getProperty("sheetencrypted-password-" + GLOBALID) != obj.password) {
    Logger.log("Passwords not matching. Return false");
    return ({ 'status': 'failed' });
  }
  else {
    EnCodeSheet(false);
    getWebAppUrl();
    return ({ 'status': 'success' });
  }
}
function EnCodeSheet(id) {
  var prop = PropertiesService.getUserProperties();
  Logger.log(id);

  var activesheet;
  if (id == false) {
    activesheet = SpreadsheetApp.getActiveSpreadsheet();
    activesheet.setActiveSelection("A1:A1");
  }
  else {
    activesheet = SpreadsheetApp.openById(prop.getProperty("sheetencrypted-id-" + GLOBALID));
  }
  if (prop.getProperty("sheetencrypted-state-" + GLOBALID) == 2) {
    return;
  }

  for (var k = 0; k < activesheet.getSheets().length; k++) {
    var ss = activesheet.getSheets()[k];
    var range = ss.getDataRange();
    var vals = range.getValues();
    //var actvals=[];


    for (var i = 2; i < vals.length; i++) {
      for (var j = 0; j < vals[i].length; j++) {
        if (vals[i][j] != "") {
          if (ss.getRange(i + 1, j + 1, 1, 1).getFormula() == "") {
            vals[i][j] = encrypt(vals[i][j], 1);
            ss.getRange(i + 1, j + 1, 1, 1).setValue(vals[i][j]);
          }
        }
      }
    }
  }
  prop.setProperty("sheetencrypted-state-" + GLOBALID, 2);
}

function DecodeFromSheet() {
  var prop = PropertiesService.getUserProperties();
  if (prop.getProperty("sheetencrypted-state-" + GLOBALID) == 1) {
    Browser.msgBox('ATTENTION', 'The sheet is already in normal state!!', Browser.Buttons.OK);
    return;
  }

  if (prop.getProperty("sheetencrypted-password-" + GLOBALID) == null) {
    Browser.msgBox("You have not encoded the file yet!!!!");
    return;
  }
  else {
    SpreadsheetApp.getActiveSpreadsheet().show(HtmlService.createHtmlOutputFromFile('inputpassworddecrypt')); getWebAppUrl();
  }
}
function decodeForRequest(obj) {
  var prop = PropertiesService.getUserProperties();

  Logger.log("Starting decodeForRequest - " + obj.password);
  if (prop.getProperty("sheetencrypted-password-" + GLOBALID) != obj.password) {
    Logger.log("Login failed");
    return ({ 'status': 'failed' });
  }
  else {
    Logger.log("Login success");
    DeCodeSheet(false);
    getWebAppUrl();
    return ({ 'status': 'success' });
  }
}

// 1 - sheet is in normal state.
// 2 - sheet is encrypted.
function DeCodeSheet(id) {
  Logger.log("From DecodeSheet");
  var prop = PropertiesService.getUserProperties();
  var activesheet;
  if (id == false) {
    activesheet = SpreadsheetApp.getActiveSpreadsheet();
    activesheet.setActiveSelection("A1:A1");
  }
  else {
    activesheet = SpreadsheetApp.openById(prop.getProperty("sheetencrypted-id-" + GLOBALID));
  }

  if (prop.getProperty("sheetencrypted-state-" + GLOBALID) == 1) {
    Logger.log("Already decoded");
    return;
  }

  for (var k = 0; k < activesheet.getSheets().length; k++) {
    var ss = activesheet.getSheets()[k];
    var range = ss.getDataRange();
    var vals = range.getValues();

    for (var i = 2; i < vals.length; i++) {
      for (var j = 0; j < vals[i].length; j++) {
        if (vals[i][j] != "") {
          if (ss.getRange(i + 1, j + 1, 1, 1).getFormula() == "") {
            vals[i][j] = decrypt(vals[i][j], 1);
            ss.getRange(i + 1, j + 1, 1, 1).setValue(vals[i][j]);
          }
        }
      }
    }
  }
  prop.setProperty("sheetencrypted-state-" + GLOBALID, 1);
}


function encrypt(text, key) {
  var endResult = "";
  key = key * 7;
  Logger.log(typeof (text));
  if (typeof (text) == "number") {
    text = text.toString();
  }
  if (typeof (text) != "string") {
    Logger.log("Got invalid " + typeof (text) + " " + text);
    return text;
  }
  var aa = text.split('');

  var a; var b;
  for (var j = 0; j < aa.length; j++) {
    a = text.charCodeAt(j);
    if (j == 0 && String.fromCharCode(a) == 6) {
      //= at start of cell will convert value to formula.
      endResult += String.fromCharCode(a);
      continue;
    }
    for (var i = 0; i < key; i++) {
      if (!(a >= 123 || a < 31)) {
        if (a + 1 != 123) {
          a += 1;
        }
        else {
          a = 32;
        }
      }
    }
    endResult += String.fromCharCode(a);
  }
  return endResult;
}

function decrypt(text, key) {
  var endResult = "";
  key = key * 7;
  Logger.log(typeof (text));
  if (typeof (text) == "number") {
    text = text.toString();
  }
  if (typeof (text) != "string") {
    Logger.log("Got invalid " + typeof (text) + " " + text);
    return text;
  }
  var aa = text.split('');

  var a;
  for (var j = 0; j < aa.length; j++) {
    a = text.charCodeAt(j);
    if (j == 0 && String.fromCharCode(a) == 6) {
      //= at start of cell will convert value to formula.
      endResult += String.fromCharCode(a);
      continue;
    }
    for (var i = 0; i < key; i++) {
      if (!(a >= 123 || a < 31)) {
        if (a - 1 != 31) {
          a -= 1;
        }
        else {
          a = 122;
        }
      }
      else {
        break;
      }
    }
    endResult += String.fromCharCode(a);
  }
  return endResult;
}

function getHtml(msg, butt) {
  html = '<html>' +
    '<head>' +
    '</head>' +
    '<body>' +
    '<div style="width:100%; text-align:center; font-family:Georgia;">' +
    '<h2 style="font-size:40px;"><i>Input You password.</i></h2>' +
    '<form type="submit" action="' + ScriptApp.getService().getUrl() + '" method="post" style="font-size:22px;">' +
    '<label>' + msg + '</label>' +
    '<input type="password" name="password" value="" style="padding:5px; width:300px;" />' +
    '<input type="submit" name="submit" value="' + butt + '" style="padding:5px;" />' +
    '</form>' +
    '</div>' +
    '</body>' +
    '</html>';
  return html;
}

function doGet() {
  var prop = PropertiesService.getUserProperties();
  var password = '';
  var html = '';
  if (prop.getProperty("sheetencrypted-password-" + GLOBALID) == null) {
    html = '<html><body>You have not set any password</body></html>';
  }
  else {
    var butt;
    if (prop.getProperty("sheetencrypted-state-" + GLOBALID) == 1) {
      butt = 'Encrypt';
    }
    else {
      butt = 'decrypt';
    }
    html = getHtml('', butt);
  }
  return HtmlService.createHtmlOutput(html)
}

function doPost(e) {
  var prop = PropertiesService.getUserProperties();
  var html = '';
  if (prop.getProperty("sheetencrypted-password-" + GLOBALID) == null) {
    html = '<html><body>You have not set any password</body></html>';
  }
  else {
    var butt;
    if (prop.getProperty("sheetencrypted-state-" + GLOBALID) == 1) {
      butt = 'Encrypt';
    }
    else {
      butt = 'Decrypt';
    }

    var docurl = prop.getProperty("sheetencrypted-url-" + GLOBALID);

    if (e.parameter.password != prop.getProperty("sheetencrypted-password-" + GLOBALID)) {
      html = getHtml('<span style="color:red;">Incorrect password. Please retry!!!</span><br/>', butt);
      return HtmlService.createHtmlOutput(html);
    }
    else {
      if (e.parameter.submit == 'Encrypt') {
        EnCodeSheet(true);
        html = getHtml('<span style="color:green;">Encoded Successfully!! <a href="' + docurl + '">Click here to go back.</a></span><br/>', 'Decrypt');
      }
      else {
        DeCodeSheet(true);
        html = getHtml('<span style="color:green;">Decoded Successfully!!  <a href="' + docurl + '">Click here to go back.</a></span><br/>', 'Encrypt');
      }
      return HtmlService.createHtmlOutput(html);
    }
  }
}


function onOpen() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var menuEntries = [{ name: "Initialize", functionName: "Initialize" },
    null,
  { name: "Encrypt File", functionName: "EncodeFromSheet" },
  { name: "Decrypt File", functionName: "DecodeFromSheet" },
    null,
  { name: "Change Password", functionName: "showChangePasswordForm" },
  { name: "Get Webapp URL", functionName: "getWebAppUrl" }];
  ss.addMenu("Encrypt", menuEntries);
}

function getWebAppUrl() {
  SpreadsheetApp.getActiveSpreadsheet().getSheets()[0].getRange('C1').setValue('=HYPERLINK("' + ScriptApp.getService().getUrl() + '", "http://script.google.com/...")');
}

function onInstall() {
  onOpen();
}

function Initialize() {
  return;
}
