/* refer http://labnol.org/?p=21060 */

var c = {
  pos: {
    user: "B2",
    pass: "B3",
    email: "B4",
    sendSMS: "B5",
    interval: "B6",
    debug: "B7"
  },
  api: {
    base: "http://bbs.seu.edu.cn/api",
    token: "/token.json",
    notifications: "/notifications.json"
  }
};

function onOpen() {  
 
  var sheet = SpreadsheetApp.getActiveSpreadsheet();

  var menu = [
    {name: "start", functionName: "init"},
    {name: "notifications", functionName: "notifications"},
    {name: "stop", functionName: "removeJobs"}
  ];  
  
  sheet.addMenu("SBBS API", menu);    
}

function init() {
    
  removeJobs(true);

  var sheet = SpreadsheetApp.getActiveSpreadsheet();

  var item = {
    user: getPosVal(sheet, c.pos.user),
    pass: getPosVal(sheet, c.pos.pass),
    token: "",
    email: getPosVal(sheet, c.pos.email),
    sendSMS: getPosVal(sheet, c.pos.sendSMS).toLowerCase(),
    interval: getPosVal(sheet, c.pos.interval),
    debug: getPosVal(sheet, c.pos.debug).toLowerCase()
  };
  ScriptProperties.setProperties(item);
  
  token();
    
  ScriptApp.newTrigger("init")
    .forSpreadsheet(sheet)
    .onEdit()
    .create();
  
  ScriptApp.newTrigger("notifications")
    .timeBased()
    .everyMinutes(item.interval)
    .create();
  
  sheet.toast("Initialized", "SBBS API", -1);
}

function token() {
  var item = ScriptProperties.getProperties();    
  var posts = {
    user: item.user,
    pass: item.pass
  };
  var ret = fetch(c.api.token, {}, posts, "post");
  debug(ret);

  if (ret.code == 200 && ret.content.success) {
    item.token = ret.content.token;
    ScriptProperties.setProperties(item);
    return true;
  }

  return false;
}

function notifications() {  
  var item = ScriptProperties.getProperties();    

  var ret = fetch(c.api.notifications);
  debug(ret);

  if (ret.code == 200 && ret.content.success) {
    var mails = ret.content.mails || [];
    var ats = ret.content.ats || [];
    var replies = ret.content.replies || [];

    var msg = "";

    for (var i = 0; i < mails.length; i++) {
	var j = mails[i];
	var time = new Date(j.time * 1000);
	msg += "[信-" + j.sender + "][" + j.id + "][" + time + "]" + j.title + "\n";
    }
    for (var i = 0; i < ats.length; i++) {
	var j = ats[i];
	var time = new Date(j.time * 1000);
	msg += "[AT-" + j.user + "][" + j.board + "][" + j.id + "][" + time + "]" + j.title + "\n";
    }
    for (var i = 0; i < replies.length; i++) {
	var j = replies[i];
	var time = new Date(j.time * 1000);
	msg += "[RE-" + j.user + "][" + j.board + "][" + j.id + "][" + time + "]" + j.title + "\n";
    }

    if (msg.length)
	alertUser(msg);

    return true;
  }

  return false;
}

function debug(info) {
  var item = ScriptProperties.getProperties();
  if (item.debug == "yes")
      log(info);
}

function log(info) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var item = ScriptProperties.getProperties();
  
  var row   = sheet.getLastRow() + 1;    
  var time  = new Date();  
  
  sheet.getRange(row,1).setValue(time);
  sheet.getRange(row,2).setValue(info);
}
    
function alertUser(msg) {
  var item = ScriptProperties.getProperties();
  var alert = msg;
  
  MailApp.sendEmail(item.email, "SBBS有新通知", alert);

  var time  = new Date();  
  if (item.sendSMS == "yes") {
    time = new Date(time.getTime() + 15000);
    CalendarApp.createEvent(alert, time, time).addSmsReminder(0); 
  }
}

function removeJobs(quiet) {    
  var triggers = ScriptApp.getScriptTriggers();

  for (i = 0; i < triggers.length; i++) {
    ScriptApp.deleteTrigger(triggers[i]);
  }
  
  var sheet = SpreadsheetApp.getActiveSpreadsheet();

  sheet.toast("Stoped", "SBBS API", -1);
}

function getPosVal(sheet, pos) {
  return sheet.getSheets()[0].getRange(pos).getValue();
}

function fetch(url, gets, posts, method) {
  var gets = gets || {};
  var posts = posts || {};
  var method = method || "get";

  var item = ScriptProperties.getProperties();    
  var code = -1;
  var content = {};

  url = c.api.base + url + "?";

  if (item.token.length)
    gets.token = item.token;

  var keys = Object.keys(gets);
  for (var i = 0; i <  keys.length; i++) {
    var k = keys[i];
    url = url + k + "=" + gets[k] + "&";
  }

  try {
      var options = {
	method: method,
	payload: posts
      };
      var response = UrlFetchApp.fetch(url, options);
      code = response.getResponseCode();    
      content = Utilities.jsonParse(response.getContentText());
  } catch(error) {
  }

    return {code: code, content: content};
}

/* vim:set shiftwidth=2: */ 
