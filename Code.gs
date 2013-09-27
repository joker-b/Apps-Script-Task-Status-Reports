//
// Create and email weekly reports for various task lists: summarize the past and upcoming weeks.
// Set a trigger or two to run these on Friday or Monday.
//
// Recognizes special tags used by Jorte.
// 

//
// These are top-level calls that you would normally fire by clock triggers -- I have
//    "send_all()" triggered to run every Friday afternoon.
//  Each line sends a report for a single google tasks tak list to the indicated email address
//
function send_all() {
  fetch_tasks("Caustic","somename@gmail.com");
  fetch_tasks("vuMondo","anotherAddr@anywhere.com");
  fetch_tasks("Default List","my.name+reports@gmail.com");
}

// single-list calls (mostly for debugging)
function send_vmOnly() {
  fetch_tasks("vuMondo","anotherAddr@anywhere.com");
}

function send_defOnly() {
  fetch_tasks("Default List","my.name+reports@gmail.com");
}

//////////

function week_ahead() {  
  var d = new Date();
  var n = d.valueOf() + 1000*3600*24*7;
  return new Date(n);
}

function week_ago() {  
  var d = new Date();
  var n = d.valueOf() - 1000*3600*24*7;
  return new Date(n);
}

//
// seek the named task list
//
function list_by_name(ListName) {
  var allLists = Tasks.Tasklists.list().getItems();
  var id;
  for (var i in allLists) {
    if (ListName == allLists[i].getTitle()) {
      id = allLists[i].getId();
    }
  }
  return id;
}

//////////////////////////////////////////////////////////

//
// look at the indicated item list and write both text and HTML fragments into the indicated message
//
function add_to_message(Msg,Items,Label) {
  if (Items == undefined) {
    return Msg;
  }
  if (Items.length == 0) {
    return Msg;
  }
  Msg['body'] += Label+":\n";
  Msg['htmlBody'] += "<tr><td style=\"background-color:#dddddd; padding: 5px;\" colspan=\"2\"><strong>"+Label+"</strong></td></tr>\n";
  var n, i, t, c, hasCompleted, hasPending, spc, cl;
  n = 0;
  for (i=0; i<Items.length; i+=1) {
    c = Items[i].getCompleted();
    if (c != undefined) {
      n += 1;
      break;
    }
  }
  hasCompleted = (n>0);
  hasPending = ((Items.length-n)>0);
  Msg['htmlBody'] += "<tr><th><i>Completed</i></th><th><i>Pending</i></th></tr>\n";
  Msg['htmlBody'] += "<tr><td STYLE=\"color:#aaaaaa;text-decoration:line-through; padding: 3px;\">";
  if (hasCompleted) {
    n = 0;
    Msg['body'] += "Completed:\n";
    spc = "";
    for (i=0; i<Items.length; i+=1) {
      t = Items[i].getTitle();
      c = Items[i].getCompleted();
      if (c != undefined) {
        Msg['body'] += "* "+t+"\n";
        cl="#000000";
        if (t.match(/\[!\]/)) {
          cl = "#bb0000";
        }
        t = "<span style=\"color:"+cl+";\">"+t+"</span>";
        Msg['htmlBody'] += spc+t+"\n";
        n += 1;
        spc = "<BR />";
      }
    }
  }
  Msg['htmlBody'] += "</td><td style=\"padding: 3px;\">";
  if (hasPending) {
    Msg['body'] += "Pending:\n";
    spc = "";
    for (i=0; i<Items.length; i+=1) {
      t = Items[i].getTitle();
      c = Items[i].getCompleted();
      if (c == undefined) {
        Msg['body'] += "* "+t+"\n";
        if (t.match(/\[!\]/)) {
          t = "<span style=\"color:#bb0000;\">"+t+"</span>";
        }
        Msg['htmlBody'] += spc+t+"\n";
        spc = "<BR />";
      }
    }
  }
  Msg['htmlBody'] += "</td></tr>";
  return Msg;
}

//
// Send a report about the indicated taks list to the desintation email address
//
function fetch_tasks(ListName,Destination)
{
  var listId = list_by_name(ListName);
  if (!listId) {
    Logger.log("No such list");
    return;
  }
  var now = new Date();
  var later = week_ahead();
  var before = week_ago();
  var i, n, c, t, tasks, items;
  var message = {
    to: Destination,
    subject: "Weekly Task Report: "+ListName,
    name: "Happy Tasks Robot",
    htmlBody: "<div STYLE=\"background-color:rgba(1,.9,.4,.9);\"><p>Task statuses for the past and upcoming week.</p>\n<table>\n",
    body:  "Task report for "+ListName+" during the past and upcoming week:\n"
  };
  //Logger.log(before.toISOString());
  // past week
  tasks = Tasks.Tasks.list(listId, {'showCompleted':true, 'dueMin':before.toISOString(), 'dueMax': now.toISOString()});
  items = tasks.getItems();
  message = add_to_message(message,items,"Past Week");
  // upcoming week
  tasks = Tasks.Tasks.list(listId, {'showCompleted':true, 'dueMin':now.toISOString(), 'dueMax': later.toISOString()});
  items = tasks.getItems();
  message = add_to_message(message,items,"Upcoming Week");
  message['htmlBody'] += "</table></div>";
  MailApp.sendEmail(message);
  Logger.log('Sent email:\n'+message['body']);
  Logger.log('Sent email:\n'+message['htmlBody']);
}

//////////////////////////////// eof //
