//
// Create and email weekly (or daily) reports for various task lists: summarize the past and upcoming weeks.
// Set a trigger or two to run these on Friday or Monday.
//
// Recognizes special formatting tags used by Jorte.
// 
//
// These are top-level calls that you would normallyt fire by regular triggers
//  each line sends a report for a single google tasks tak list to the indicated email address
//

// weekly reports -- all at once, or individually
//    in my standard usage I call "send_all()" on a weekly trigger: every Friday at 4PM
//

var OFFICE = "my_address@big_co.com";
var HOME = "my_home_adress@gmail.com";

var WORK_LIST = "Office";
var PROJECT1 = "NewGarage";

// daily report -- in my usage, I set up FIVE weekly triggers, MTWTF at 9AM -- this report
//     reminds my team &  I what needs to be done each morning

function daily_office() {
  _daily_report_(WORK_LIST,OFFICE);
}

// weekly reports ////////

function send_all() {
  _weekly_report_(WORK_LIST,OFFICE);
  _weekly_report_(PROJECT1,HOME);
  _weekly_report_("Default List",HOME);
}

// single-list calls (mostly for debugging)
function send_vmOnly() {
  _weekly_report_(PROJECT1,HOME);
}
function send_cOnly() {
  _weekly_report_(WORK_LIST,OFFICE);
}
function send_defOnly() {
  _weekly_report_("Default List",HOME);
}

//////////////////////////////////////////////////////////////////
//////////////////////////////////////////////////////////////////
//////////////////////////////////////////////////////////////////

// accepts positive or negative Days counts
function _days_away_(Days) {
  var d = new Date();
  var n = d.valueOf() + 1000*3600*24*Days;
  return new Date(n);
}

////////////////////////////////////

// look for a named list
function _list_by_name_(ListName) {
  var allLists = Tasks.Tasklists.list().getItems();
  var id;
  for (var i in allLists) {
    if (ListName == allLists[i].getTitle()) {
      id = allLists[i].getId();
    }
  }
  return id;
}

// debugging function /////////////////////////
function _describe_task_(T) {
  Logger.log(T.getTitle());
  Logger.log('Due '+T.getDue());
}

// advertise //////////////////////
function _footer_text_() {
  var ft = '<tr align="center"><td style="background-color:#dddddd;margin-left:auto;margin-right:auto;padding:5px;font-size:smaller;font-style:italic;" colspan="2">';
  ft += 'Robots love to <a href="https://github.com/joker-b/Apps-Script-Task-Status-Reports" target="_blank">make more robots!</a> ';
  ft += 'Check out my GoogleApps projects.';
  ft += '</td></tr>\n';
  return ft;
}

//////////////////////////////////////////////////////////

// sort by name (and time?)
// X and Y are Items!!!!
function _sort_items(X,Y) {
  var yt, xt, comp;
  yt = Y.getTitle().toLowerCase();
  xt = X.getTitle().toLowerCase();
  comp = (xt>yt)?1:((yt>xt)?-1:0);
  return (comp);
}

//
// look at the indicated item list and write both text and HTML fragments into the indicated message
//
function _add_to_message_(Msg,Items,Label) {
  if (Items == undefined) {
    return Msg;
  }
  if (Items.length == 0) {
    return Msg;
  }
  var iSort = Items.sort(_sort_items);
  Msg['body'] += Label+":\n";
  Msg['htmlBody'] += "<tr><td style=\"background-color:#dddddd;font-weight:bolder;padding:5px;margin-left:auto;margin-right:auto;\" colspan=\"2\">"+Label+"</td></tr>\n";
  var n, i, t, c, desc, hasCompleted, hasPending, spc, cl;
  n = 0;
  for (i=0; i<iSort.length; i+=1) {
    c = iSort[i].getCompleted();
    if (c != undefined) {
      n += 1;
      break;
    }
  }
  hasCompleted = (n>0);
  hasPending = ((iSort.length-n)>0);
  Msg['htmlBody'] += "<tr style=\"background-color:#eaeaea;font-style:italic;\"><th>Pending</th><th>Completed</th></tr>\n";
  Msg['htmlBody'] += "<tr>";
  Msg['htmlBody'] += "<td valign=\"top\" style=\"padding:8px;max-width:300px;background-color:#f8f8f8;\"><dl>";
  if (hasPending) {
    Msg['body'] += "Pending:\n";
    spc = "";
    for (i=0; i<iSort.length; i+=1) {
      c = iSort[i].getCompleted();
      if (c == undefined) {
        t = iSort[i].getTitle();
        Msg['body'] += "* "+t+"\n";
        if (t.match(/\[!\]/)) {
          t = "<span style=\"color:#bb0000;\">"+t+"</span>";
        }
        Msg['htmlBody'] += "<dt>"+t+"</dt>\n";
        desc = iSort[i].getNotes();
        if ((desc != undefined) && (desc > "")) {
            Msg['htmlBody'] += "<dd style=\"color:#999999;font-style:italic;\">"+desc+"</dd>\n";
        }
        spc = "<BR />";
      }
    }
  } else {
    Msg['htmlBody'] += '<dd style="font-style:italic;">(none)</dd>\n';
  }
  Msg['htmlBody'] += "</dl></td>";
  Msg['htmlBody'] += "<td valign=\"top\" STYLE=\"color:#aaaaaa;text-decoration:line-through; padding: 8px;max-width: 300px;background-color:#f8f8f8;\"><dl>";
  if (hasCompleted) {
    n = 0;
    Msg['body'] += "Completed:\n";
    for (i=0; i<iSort.length; i+=1) {
      c = iSort[i].getCompleted();
      if (c != undefined) {
        t = iSort[i].getTitle();
        Msg['body'] += "* "+t+"\n";
        cl="#555555";
        if (t.match(/\[!\]/)) {
          cl = "#bb0000";
        }
        t = "<span style=\"color:"+cl+";\">"+t+"</span>";
        Msg['htmlBody'] += "<dt>"+t+"</dt>\n";
        desc = iSort[i].getNotes();
        if ((desc != undefined) && (desc > "")) {
            Msg['htmlBody'] += "<dd style=\"color:#999999;font-style:italic;\">"+desc+"</dd>\n";
        }
        n += 1;
      }
    }
  } else {
    Msg['htmlBody'] += '<dd style="font-style:italic;">(none)</dd>\n';
  }
  Msg['htmlBody'] += "</dl></td>"
  Msg['htmlBody'] += "</tr>";
  return Msg;
}


/////////////////////////

//
// Send a report about the indicated taks list to the desintation email address
//
function _weekly_report_(ListName,Destination)
{
  var listId = _list_by_name_(ListName);
  if (!listId) {
    Logger.log("No such list");
    return;
  }
  var now = new Date();
  var later = _days_away_(7);
  var before = _days_away_(-7);
  var i, n, c, t, tasks, items;
  var message = {
    to: Destination,
    subject: "Weekly Task Status: "+ListName,
    name: "Happy Tasks Robot",
    htmlBody: '<div STYLE="background-color:rgba(1,.9,.4,.9);"><p>Task statuses for the past and upcoming week.</p>\n'+
    '<table style="border-spacing:6px;">\n',
    body:  "Task report for "+ListName+" during the past and upcoming week:\n"
  };
  //Logger.log(before.toISOString());
  // past week
  tasks = Tasks.Tasks.list(listId, {'showCompleted':true, 'dueMin':before.toISOString(), 'dueMax': now.toISOString()});
  items = tasks.getItems();
  message = _add_to_message_(message,items,"Past Week");
  // upcoming week
  tasks = Tasks.Tasks.list(listId, {'showCompleted':true, 'dueMin':now.toISOString(), 'dueMax': later.toISOString()});
  items = tasks.getItems();
  message = _add_to_message_(message,items,"Upcoming Week");
  message['htmlBody'] += _footer_text_();
  message['htmlBody'] += "</table>\n";
  message['htmlBody'] += "</div>";
  MailApp.sendEmail(message);
  Logger.log('Sent email:\n'+message['body']);
  Logger.log('Sent email:\n'+message['htmlBody']);
}

//
// Send a report about the indicated taks list to the desintation email address
//
function _daily_report_(ListName,Destination)
{
  var listId = _list_by_name_(ListName);
  if (!listId) {
    Logger.log("No such list");
    return;
  }
  var now = new Date();
  var lookBack2 = (now.getDay() == 1) ? -4 : -2;
  var later = _days_away_(1);
  var before = _days_away_(-1);
  var before2 = _days_away_(lookBack2);
  var i, n, c, t, tasks, items;
  var message = {
    to: Destination,
    subject: "Daily Task Status: "+ListName,
    name: "Happy Tasks Robot",
    htmlBody: "<div STYLE=\"background-color:rgba(1,.9,.4,.9);\"><table>\n",
    body:  "Task report for "+now+":\n"
  };
  //Logger.log(before.toISOString());
  // past week
  tasks = Tasks.Tasks.list(listId, {'showCompleted':true, 'dueMin':before.toISOString(), 'dueMax': now.toISOString()});
  items = tasks.getItems();
  message = _add_to_message_(message,items,"Today's Scheduled Tasks");
  tasks = Tasks.Tasks.list(listId, {'showCompleted':true, 'dueMin':before2.toISOString(), 'dueMax': before.toISOString()});
  items = tasks.getItems();
  message = _add_to_message_(message,items,"Yesterday's Schedule");
  message['htmlBody'] += _footer_text_();
  message['htmlBody'] += "</table>\n";
  message['htmlBody'] += "</div>";
  //
  MailApp.sendEmail(message);
  Logger.log('Sent email:\n'+message['body']);
  Logger.log('Sent email:\n'+message['htmlBody']);
}

//////////////////////////////// eof //

function _add_tester_(ListName) {
  var id = _list_by_name_(ListName);
  if (!id) {
    Logger.log("Tasklist not found"); 
  } else {
    var newTask = Tasks.newTask()
        .setTitle("make task list sniffer");
    var inserted = Tasks.Tasks.insert(newTask, id);  
    Logger.log("Task added");
  }
  return id;
}

function _add_TL_(SomeName) {
   var newTaskList = Tasks.newTaskList()
      .setTitle(SomeName);
  
  var created = Tasks.Tasklists.insert(newTaskList);
  return created;
}