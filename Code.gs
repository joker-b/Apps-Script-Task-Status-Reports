//
// Create and email weekly (or daily) reports for various task lists: summarize the past and upcoming weeks.
// Set a trigger or two to run these on Friday or Monday.
//
// Recognizes special formatting tags used by Jorte.
//
// Nov 2013 extra: archive old, completed tasks
//
// These are top-level calls that you would normally fire by regular triggers
//  each line sends a report for a single google tasks tak list to the indicated email address
//

// weekly reports -- all at once, or individually
//    in my standard usage I call "send_all()" on a weekly trigger: every Friday at 4PM
//

var OFFICE = "some_addr+work@gmail.com";
var HOME = "some_other_addr@gmail.com";

var WORK_LIST = "Classes";
var PROJECT1 = "vuMondo";

var testData = {
  'listID': '<some list id>',  // Default List
};

// daily report -- in my usage, I set up FIVE weekly triggers, MTWTF at 9AM -- this report
//     reminds my team &  I what needs to be done each morning

function daily_office() {
  _daily_report_(WORK_LIST, OFFICE);
}

// weekly reports ////////

function send_all() {
  _weekly_report_("Default List",HOME);
  _weekly_report_("vuMondo",HOME);
  _weekly_report_("myTipsy",HOME);
}

function send_photo() {
  _weekly_report_("Pix make",HOME);
  _weekly_report_("pix printing",HOME);
  _weekly_report_("Pix view",HOME);
  _weekly_report_("pix website",HOME);
}

function send_all_weekend() {
  _weekly_report_("Sonoma",HOME);
  _weekly_report_("Bike",HOME);
  _weekly_report_("Classes",HOME);
}

function occasions() {
  Lists = ["Pix make", "pix printing", "Pix view", "pix website", "Books", "Bike",
    "Sonoma", "Classes", "vuMondo", "myTipsy", "Default List"];
    _occasional_report_(Lists,HOME);
}

// single-list calls (mostly for debugging)
function send_vmOnly() {
  _weekly_report_(PROJECT1, HOME);
}
function send_cOnly() {
  _weekly_report_(WORK_LIST, OFFICE);
}
function send_defOnly() {
  _weekly_report_("Default List", HOME);
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

function _shuffle(array) {
  var currentIndex = array.length, temporaryValue, randomIndex;
  // While there remain elements to shuffle...
  while (0 !== currentIndex) {
    // Pick a remaining element...
    randomIndex = Math.floor(Math.random() * currentIndex);
    currentIndex -= 1;
    // And swap it with the current element.
    temporaryValue = array[currentIndex];
    array[currentIndex] = array[randomIndex];
    array[randomIndex] = temporaryValue;
  }
  return array;
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
  Msg['htmlBody'] += "<tr><td style=\"font-weight:bolder;padding:5px;margin-left:auto;margin-right:auto;\" colspan=\"2\">"+Label+"</td></tr>\n";
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
  Msg['htmlBody'] += "<tr style=\"font-style:italic;\"><th>Pending</th><th>Completed</th></tr>\n";
  Msg['htmlBody'] += "<tr>";
  Msg['htmlBody'] += "<td valign=\"top\" style=\"padding:8px;max-width:300px;\"><dl>";
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
  Msg['htmlBody'] += "<td valign=\"top\" STYLE=\"color:#aaaaaa;text-decoration:line-through; padding: 8px;max-width: 300px;\"><dl>";
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


///

//
// look at the indicated item list and write both text and HTML fragments into the indicated message
//
function _add_to_archive_message_(Msg, Items, Label) {
  if (Items == undefined) {
    return Msg;
  }
  if (Items.length == 0) {
    return Msg;
  }
  var iSort = Items; // .sort(_sort_items);
  Msg['body'] += Label+" ("+Items.length+" items):\n";
  Msg['htmlBody'] += "<h3>"+Label+" ("+Items.length+" items)</h3><DL>\n";
  var n, i, t, c, desc, hasCompleted, hasPending, spc, cl, due;
  n = 0;
  for (i=0; i<iSort.length; i+=1) {
    c = iSort[i].getCompleted();
    if (c != undefined) {
      n += 1;
      break;
    }
  }
  for (i=0; i<iSort.length; i+=1) {
    c = iSort[i].getCompleted();
    if (c != undefined) {
      t = iSort[i].getTitle();
      due = iSort[i].getDue();
      if ((due != undefined) && (due > "")) {
          t += (' ('+due.slice(0,10)+')');
      }
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
      // iSort[i].setDeleted(true).setHidden(true);
    }
  }
  Msg['htmlBody'] += "</DL>";
  return Msg;
}

function debug_completed()
{
  Logger.log("Looking at tasklists and tasks");
  var now = new Date().toISOString();
  var recent = _days_away_(-0.25/24).toISOString();
  var before = _days_away_(-14).toISOString();
  Logger.log('completedMin '+before);
  let allLists = Tasks.Tasklists.list().getItems();
  for (var i in allLists) {
    Logger.log("list item "+i);
    let title = Logger.log(allLists[i].getTitle());
    Logger.log("List '"+title+"'");
    let listID = allLists[i].getId();
    Logger.log("List ID '"+listID+"'");
    let tasks = Tasks.Tasks.list(listID, {'showCompleted':true, 'showDeleted':false, 'showHidden':true, 'completedMin':before, 'completedMax':now, 'maxResults': 100});
    items = tasks.getItems();
    Logger.log("List has "+items.length+" items");
    break;
  }
  Logger.log("Adios");
};

function _remove_completed_items(ListID, Items)
{
  if (Items == undefined) {
    return;
  }
  if (Items.length == 0) {
    return;
  }
  Logger.log('Reviewing '+Items.length+" items for removal");
  Logger.log("  from list "+ListID);
  var nItems = 0;
  for (var i in Items) {
    let item = Items[i];
    let t = item.getTitle();
    let comp = item.getCompleted();
    if (comp != undefined) {
      let taskID = item.getId();
      // Logger.log("Want to remove Item '"+t+"' : "+comp);
      // we use "remove" because "delete" is a reserved word
      Tasks.Tasks.remove(ListID, taskID);
      nItems = nItems + 1;
    //} else {
    //  Logger.log("ignoring incomplete item '"+t+"'");
    }
  }
  Logger.log("Removed "+nItems+" items.");
}

function archive_completed()
{
  var i, j, c, n, title, id, allLists, tasks, items;
  var delay = 32;
  var now = new Date();
  var recent = _days_away_(-1).toISOString();
  var before = _days_away_(-delay).toISOString();
  Logger.log("from "+before);
  var message = {
    to: HOME,
    subject: "Completed Task Report",
    name: "Happy Tasksniffer Robot",
    htmlBody: '<div><p>These tasks are already done!</p>\n',
    body:  "Completed Tasks:\n"
  };
  var allLists = Tasks.Tasklists.list().getItems();
  for (i in allLists) {
    title = allLists[i].getTitle();
    id = allLists[i].getId();
    tasks = Tasks.Tasks.list(id, {'showCompleted':true, 'showDeleted':false, 'showHidden':true, 'completedMin':before, 'completedMax':recent, 'maxResults': 100});
    items = tasks.getItems();
    _add_to_archive_message_(message, items, title);
    _remove_completed_items(id, items);
    // Logger.log("trying to clear "+title);
    // Tasks.Tasks.clear(id);
  }
  message['htmlBody'] += "</div>";
  MailApp.sendEmail(message);
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
  var i, n, c, t, ptasks, pitems, ntasks, nitems;
  var message = {
    to: Destination,
    subject: "Weekly Task Status: "+ListName,
    name: "Happy Tasks Robot",
    htmlBody: '<div STYLE="background-color:#EEEE88;"><p>Task statuses for the past and upcoming week.</p>\n'+
    '<table style="border-spacing:6px;">\n',
    body:  "Task report for "+ListName+" during the past and upcoming week:\n"
  };
  //Logger.log(before.toISOString());
  // past week
  ptasks = Tasks.Tasks.list(listId, {'showCompleted':true, 'dueMin':before.toISOString(), 'dueMax': now.toISOString()});
  pitems = ptasks.getItems();
  // upcoming week
  ntasks = Tasks.Tasks.list(listId, {'showCompleted':true, 'dueMin':now.toISOString(), 'dueMax': later.toISOString()});
  nitems = ntasks.getItems();
  if (pitems && nitems && (nitems.length>0 || pitems.length>0)) {
    message = _add_to_message_(message,pitems,"Past Week");
    message = _add_to_message_(message,nitems,"Upcoming Week");
  } else {
    message.subject = "Task Status: "+ListName
    var tasks = Tasks.Tasks.list(listId, {'showCompleted':true, 'maxResults': 12 });
    var items = tasks.getItems();
    message = _add_to_message_(message,items,"Top Twelve Undated");
  }
  // TODO: if no items, look at the top of old items....
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
    htmlBody: "<div STYLE=\"background-color:#88ffff;\"><table>\n",
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

//
// look at the indicated item list and write both text and HTML fragments into the indicated message
//
function _add_to_list_message_(Msg,Items,Label) {
  Msg['body'] += Label+":\n";
  Msg['htmlBody'] += "<dt><b>"+Label+"</b></dt>\n";
  if ((Items == undefined)|| (Items.length == 0)) {
    Msg['body'] += "** Nothing\n";
    Msg['htmlBody'] += "<dd><i>Nothing</i></dd>\n";
    return Msg;
  }
  var iSort = Items.sort(_sort_items);
  var i, t, c, desc;
  for (i=0; i<iSort.length; i+=1) {
      t = iSort[i].getTitle();
      Msg['body'] += "&nbsp;&nbsp;&nbsp;"+t+"\n";
      if (t.match(/\[!\]/)) {
        t = "<span style=\"color:#bb0000;\">"+t+"</span>";
      }
      Msg['htmlBody'] += "<dd>"+t+"";
      desc = iSort[i].getNotes();
      if ((desc != undefined) && (desc > "")) {
          Msg['htmlBody'] += ": <i>"+desc+"</i>";
      }
      Msg['htmlBody'] += "</dd>\n";
  }
  return Msg;
}

//
// Send a report about the indicated taks list to the desintation email address
//
function _occasional_report_(ListNames,Destination)
{
  var now = new Date();
  var later = _days_away_(3);
  var before = _days_away_(-3);
  var i, n, c, t, ntasks, nitems, listName;
  var message = {
    to: Destination,
    subject: "Occasional Tasks for "+now.toLocaleString('en-US',{ weekday: 'long'}),
    name: "Occasional Tasks Robot",
    htmlBody: '<div STYLE="background-color:#aaeeaa;"><h3>Some Random Items:</h3>\n'+
    '<dl>\n',
    body:  "Some Random Items:\n"
  };
  //Logger.log(before.toISOString());
  // past week
  
  for (var ii=0; ii < ListNames.length; ii+=1) {
    listName = ListNames[ii];
    var listId = _list_by_name_(listName);
    if (!listId) {
      Logger.log("No such list: "+listName);
      continue;
    }
    ntasks = Tasks.Tasks.list(listId, {'showCompleted':false, 'dueMin':before.toISOString(), 'dueMax': later.toISOString()});
    nitems = ntasks.getItems();
    if (nitems && (nitems.length>0)) {
      nitems = _shuffle(nitems);
      message = _add_to_list_message_(message,nitems.slice(0,5),'"'+listName+'"');
    } else {
      var tasks = Tasks.Tasks.list(listId, {'showCompleted':false});
      var items = tasks.getItems();
      items = _shuffle(items);
      message = _add_to_list_message_(message,items.slice(0,5),'"'+listName+'" (Undated)');
    }
  }
  // TODO: if no items, look at the top of old items....
  message['htmlBody'] += _footer_text_();
  message['htmlBody'] += "</dl>\n";
  message['htmlBody'] += "</div>";
  MailApp.sendEmail(message);
  Logger.log('Sent email:\n'+message['body']);
  Logger.log('Sent email:\n'+message['htmlBody']);
}

//////////////////////////////// eof almost //

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
