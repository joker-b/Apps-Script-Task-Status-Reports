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

function send_all() {
  weekly_report("Caustic","some.dude@bigcorp.com");
  weekly_report("vuMondo","some.dude+reports@gmail.com");
  weekly_report("House","some.dude+reports@gmail.com");
  weekly_report("Default List","some.dude+reports@gmail.com");
}

// single-list calls (mostly for debugging)
function send_vmOnly() {
  weekly_report("vuMondo","some.dude+reports@gmail.com");
}
function send_cOnly() {
  weekly_report("Caustic","some.dude@bigcorp.com");
}
function send_defOnly() {
  weekly_report("Default List","some.dude+reports@gmail.com");
}

// daily report -- in my usage, I set up FIVE weekly triggers, MTWTF at 9AM -- this report
//     reminds my team &  I what needs to be done each morning

function daily_caustic() {
  daily_report("Caustic","some.dude@bigcorp.com");
}

//////////////////////////////////////////////////////////////////
//////////////////////////////////////////////////////////////////
//////////////////////////////////////////////////////////////////

// accepts positive or negative Days counts
function days_away(Days) {
  var d = new Date();
  var n = d.valueOf() + 1000*3600*24*Days;
  return new Date(n);
}

////////////////////////////////////

// look for a named list
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

// debugging function
function describe_task(T) {
  Logger.log(T.getTitle());
  Logger.log('Due '+T.getDue());
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
  var n, i, t, c, desc, hasCompleted, hasPending, spc, cl;
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
  Msg['htmlBody'] += "<tr><th><i>Pending</i></th><th><i>Completed</i></th></tr>\n";
  Msg['htmlBody'] += "<tr>";
  Msg['htmlBody'] += "<td valign=\"top\" style=\"padding: 8px;max-width=300px;\"><dl>";
  if (hasPending) {
    Msg['body'] += "Pending:\n";
    spc = "";
    for (i=0; i<Items.length; i+=1) {
      c = Items[i].getCompleted();
      if (c == undefined) {
        t = Items[i].getTitle();
        Msg['body'] += "* "+t+"\n";
        if (t.match(/\[!\]/)) {
          t = "<span style=\"color:#bb0000;\">"+t+"</span>";
        }
        Msg['htmlBody'] += "<dt>"+t+"</dt>\n";
        desc = Items[i].getNotes();
        if ((desc != undefined) && (desc > "")) {
            Msg['htmlBody'] += "<dd style=\"color:#999999;\"><i>"+desc+"</i></dd>\n";
        }
        spc = "<BR />";
      }
    }
  } else {
    Msg['htmlBody'] += "<dd><i>(none)</i></dd>\n";
  }
  Msg['htmlBody'] += "</dl></td>";
  Msg['htmlBody'] += "<td valign=\"top\" STYLE=\"color:#aaaaaa;text-decoration:line-through; padding: 8px;max-width: 300px;\"><dl>";
  if (hasCompleted) {
    n = 0;
    Msg['body'] += "Completed:\n";
    for (i=0; i<Items.length; i+=1) {
      c = Items[i].getCompleted();
      if (c != undefined) {
        t = Items[i].getTitle();
        Msg['body'] += "* "+t+"\n";
        cl="#555555";
        if (t.match(/\[!\]/)) {
          cl = "#bb0000";
        }
        t = "<span style=\"color:"+cl+";\">"+t+"</span>";
        Msg['htmlBody'] += "<dt>"+t+"</dt>\n";
        desc = Items[i].getNotes();
        if ((desc != undefined) && (desc > "")) {
            Msg['htmlBody'] += "<dd style=\"color:#999999;\"><i>"+desc+"</i></dd>\n";
        }
        n += 1;
      }
    }
  } else {
    Msg['htmlBody'] += "<dd><i>(none)</i></dd>\n";
  }
  Msg['htmlBody'] += "</dl></td>"
  Msg['htmlBody'] += "</tr>";
  return Msg;
}


/////////////////////////

//
// Send a report about the indicated taks list to the desintation email address
//
function weekly_report(ListName,Destination)
{
  var listId = list_by_name(ListName);
  if (!listId) {
    Logger.log("No such list");
    return;
  }
  var now = new Date();
  var later = days_away(7);
  var before = days_away(-7);
  var i, n, c, t, tasks, items;
  var message = {
    to: Destination,
    subject: "Weekly Task Status: "+ListName,
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

//
// Send a report about the indicated taks list to the desintation email address
//
function daily_report(ListName,Destination)
{
  var listId = list_by_name(ListName);
  if (!listId) {
    Logger.log("No such list");
    return;
  }
  var now = new Date();
  var later = days_away(1);
  var before = days_away(-1);
  var before2 = days_away(-2);
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
  message = add_to_message(message,items,"Today's Scheduled Tasks");
  tasks = Tasks.Tasks.list(listId, {'showCompleted':true, 'dueMin':before2.toISOString(), 'dueMax': before.toISOString()});
  items = tasks.getItems();
  message = add_to_message(message,items,"Yesterday's Schedule");
  message['htmlBody'] += "</table></div>";
  //
  MailApp.sendEmail(message);
  Logger.log('Sent email:\n'+message['body']);
  Logger.log('Sent email:\n'+message['htmlBody']);
}

//////////////////////////////// eof //

function add_tester(ListName) {
  var id = list_by_name(ListName);
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

function add_tl(SomeName) {
   var newTaskList = Tasks.newTaskList()
      .setTitle(SomeName);
  
  var created = Tasks.Tasklists.insert(newTaskList);
  return created;
}
