// This code is meant to be attached to a Google Spreadsheet as a Google Script. Once configured, this code parses out space requests from OrgSync Event Forms data, and then passes along each space request (via email) to the department that mangages the requested space.
// Administrators may then click on the link in the emails they receive (which takes them to the OrgSync event request page), and and write the approved space assignment in the event's Location parameter.
// This code also watches each future event for changes in time, and emails administrators when an event changes date/time (to alert them to update the space reservation). 
// Our institution's situation required us to separate requests for classroom spaces from all other space requests, and list them on a separate tab in our Google Sheet. This is reflected in the code below.

// By Jeremy Hissong, NYU Shanghai
// jeremy.hissong@nyu.edu

function main() {
  var fromOrgSync = getInfoFromOrgSync();
  processSheet(fromOrgSync);
}

function sendEmail(emailData) {
  var msg;
  var plainText = '';
  var to;
  var toGreeting;
  var purposeOfEmail;
  var seeBelow = 'For further details, see information below.<br /><br />';
  if (emailData.to=='registrar') {to = 'NYU Shanghai Classroom Reservations Staff <shanghai.classroom.reservations-group@nyu.edu>';toGreeting='Registrar Staff';}
  else if (emailData.to=='facilities') {to = 'NYU Shanghai Space Reservations Staff <shanghai.space.reservations-group@nyu.edu>';toGreeting='Campus & Facilities Staff';}
  else if (emailData.to=='athletics') {to = 'NYU Shanghai Athletics <shanghai.athletics@nyu.edu>';toGreeting='Athletics Staff';}
  var emailSig = '<u><font color="#cccccc">&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</font></u>'
    + '<u><font color="#cccccc">&nbsp;&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;</font></u></div></div>'
    + '<div style="color: #000000; font-family: arial, sans-serif; font-size: 13px; font-weight: normal; background-color: #ffffff;"><br />'
    + '<div class="im" style="color: #000000; line-height: 145%;"><img class="CToWUd" style="float: left; padding-right: 12pt; padding-top: 2pt;" src="https://ci6.googleusercontent.com/proxy/hNFJDHcfyVFJKwqFUYAolRBwEYWnW0k8YfOgDSAqI0N2MJrt3cl8VEcDvNxwBdP36XxGyz81FBP1QAaWNWchIQJbq5JRbISSe6I6uATphN77GAyxUMFW6No=s0-d-e1-ft#https://dl.dropboxusercontent.com/u/83124818/nyu_shanghai_logo.jpeg" alt="" />'
    + '<strong>Office of Program Management</strong><br />NYU Shanghai&nbsp;&nbsp;|&nbsp;&nbsp;上海纽约大学<br />'
    + '1555 Century Avenue &nbsp;|&nbsp;&nbsp;世纪大道1555号<br />'
    + 'Pudong New Area, Shanghai, China&nbsp;&nbsp;|&nbsp;&nbsp;中国上海浦东新区<br />'
    + 'Email (邮箱):&nbsp;<a href="mailto:shanghai.program.management@nyu.edu">shanghai.program.management@nyu.edu</a></div>';
  
  var type = emailData.type;
  var name = emailData.name;
  while (name.slice(-1) == ' ') {name = name.slice(0,-1);} // removes extra spaces at end of event name string
  var subject = type+' ['+emailData.id+']: '+name;
  var msg = '<div style="color: #000000; font-family: arial, sans-serif; font-size: 13px; font-weight: normal; background-color: #ffffff;"><div class="im" style="color: #000000;">';
  var greeting = 'Dear '+toGreeting+',<br /><br />';
  msg += greeting;
  var eventBlob = '';
  var assigned = emailData.assigned;
  if (assigned == 'No location has been entered for this event.') {assigned='None';}
  
  if (emailData.type=='New Space Request') {
    purposeOfEmail = 'A new space request has been created: ' + '<a href="'+emailData.url+'">'+name+'</a>'+'<br /><br />';
    var startDate = new Date(emailData.start+28800000);
    var amPm = 'AM';
    var hours = startDate.getHours();
    if (hours>=12) {amPm = 'PM';if(hours>12){hours = (hours-12);}}
    else if (parseInt(hours)==0) {hours=12;}
    var mins = ("0" + (startDate.getMinutes())).slice(-2);
    var startTime = hours + ':' + mins + ' ' + amPm; 
    var startMonth = startDate.getMonth();
    if (startMonth == 0) {startMonth = 'January';}
    if (startMonth == 1) {startMonth = 'February';}
    if (startMonth == 2) {startMonth = 'March';}
    if (startMonth == 3) {startMonth = 'April';}
    if (startMonth == 4) {startMonth = 'May';}
    if (startMonth == 5) {startMonth = 'June';}
    if (startMonth == 6) {startMonth = 'July';}
    if (startMonth == 7) {startMonth = 'August';}
    if (startMonth == 8) {startMonth = 'September';}
    if (startMonth == 9) {startMonth = 'October';}
    if (startMonth == 10) {startMonth = 'November';}
    if (startMonth == 11) {startMonth = 'December';}
    var startDayOfWeek = startDate.getDay();
    if (startDayOfWeek == 0) {startDayOfWeek = 'Sunday';}
    if (startDayOfWeek == 1) {startDayOfWeek = 'Monday';}
    if (startDayOfWeek == 2) {startDayOfWeek = 'Tuesday';}
    if (startDayOfWeek == 3) {startDayOfWeek = 'Wednesday';}
    if (startDayOfWeek == 4) {startDayOfWeek = 'Thursday';}
    if (startDayOfWeek == 5) {startDayOfWeek = 'Friday';}
    if (startDayOfWeek == 6) {startDayOfWeek = 'Saturday';}
    var startD = startDate.getDate();
    var startY = startDate.getFullYear();
    startDate = startDayOfWeek + ", " + startMonth + " " + startD + ", " + startY;
      
    var endDate = new Date(emailData.end+28800000);
    var hours = endDate.getHours();
    if (hours>=12) {amPm = 'PM';if (hours>12) {hours = (hours-12);}}
    else {if (parseInt(hours)==0) {hours=12;}}
    var mins = ("0" + (endDate.getMinutes())).slice(-2);
    var endTime = hours + ':' + mins + ' ' + amPm; 
    var endMonth = endDate.getMonth();
    if (endMonth == 0) {endMonth = 'January';}
    if (endMonth == 1) {endMonth = 'February';}
    if (endMonth == 2) {endMonth = 'March';}
    if (endMonth == 3) {endMonth = 'April';}
    if (endMonth == 4) {endMonth = 'May';}
    if (endMonth == 5) {endMonth = 'June';}
    if (endMonth == 6) {endMonth = 'July';}
    if (endMonth == 7) {endMonth = 'August';}
    if (endMonth == 8) {endMonth = 'September';}
    if (endMonth == 9) {endMonth = 'October';}
    if (endMonth == 10) {endMonth = 'November';}
    if (endMonth == 11) {endMonth = 'December';}
    var endDayOfWeek = endDate.getDay();
    if (endDayOfWeek == 0) {endDayOfWeek = 'Sunday';}
    if (endDayOfWeek == 1) {endDayOfWeek = 'Monday';}
    if (endDayOfWeek == 2) {endDayOfWeek = 'Tuesday';}
    if (endDayOfWeek == 3) {endDayOfWeek = 'Wednesday';}
    if (endDayOfWeek == 4) {endDayOfWeek = 'Thursday';}
    if (endDayOfWeek == 5) {endDayOfWeek = 'Friday';}
    if (endDayOfWeek == 6) {endDayOfWeek = 'Saturday';}
    var endD = endDate.getDate();
    var endY = endDate.getFullYear();
    endDate = endDayOfWeek + ", " + endMonth + " " + endD + ", " + endY;
    var datesCheck = emailData.dates;
    if ((emailData.isAllDay)==true) {startTime='All Day';endTime='All Day';}
    if ((parseInt(datesCheck.length))>10) {
      eventBlob += ('<b>Dates:</b> '+datesCheck+'<br /><br />'+'<b>Start Time:</b> '+startTime+'<br />'+'<b>End Time:</b> '+endTime+'<br /><br />'+'SETUP time: <i>'+emailData.setupTime+'</i><br />'+'CLEANUP time: <i>'+emailData.cleanupTime+'</i><br /><br />'+'<b>SPACE REQUESTED:</b> '+emailData.requested+'<br /><br /><b>SPACE ASSIGNED:</b> '+assigned);
      plainText += 'Dates: '+datesCheck+'\nStart Time: '+startTime+'\nEnd Time: '+endTime+'\n\nSPACE REQUESTED: '+emailData.requested+'\n\nLink:'+emailData.url;
    }
    else {
      if (startDate===endDate) {
        eventBlob += ('<b>Date:</b> '+startDate+'<br /><br />'+'<b>Start Time:</b> '+startTime+'<br />'+'<b>End Time:</b> '+endTime+'<br /><br />'+'SETUP time: <i>'+emailData.setupTime+'</i><br />'+'CLEANUP time: <i>'+emailData.cleanupTime+'</i><br /><br />'+'<b>SPACE REQUESTED:</b> '+emailData.requested+'<br /><br /><b>SPACE ASSIGNED:</b> '+assigned);
        plainText += 'Date: '+startDate+'\nStart Time: '+startTime+'\nEnd Time: '+endTime+'\n\nSPACE REQUESTED: '+emailData.requested+'\n\nLink:'+emailData.url;
      }
      else {
        eventBlob += ('<b>Start Date:</b> '+startDate+'<br />'+'<b>Start Time:</b> '+startTime+'<br /><br />'+'<b>End Date:</b> '+endDate+'<br />'+'<b>End Time:</b> '+endTime+'<br /><br />'+'SETUP time: <i>'+emailData.setupTime+'</i><br />'+'CLEANUP time: <i>'+emailData.cleanupTime+'</i><br /><br />'+'<b>SPACE REQUESTED:</b> '+emailData.requested+'<br /><br /><b>SPACE ASSIGNED:</b> '+assigned);
        plainText += 'Start Date: '+startDate+'\nStart Time: '+startTime+'\nEnd Date: '+endDate+'\nEnd Time: '+endTime+'\n\nSPACE REQUESTED: '+emailData.requested+'\n\nLink:'+emailData.url;
      }
    }
    
    if (emailData.specialRequirements != 'None') {
      eventBlob += ('<br /><br /><b>SPECIAL REQUIREMENTS:</b> <i>'+emailData.specialRequirements+'</i>');
    }
    
    msg += purposeOfEmail;
    msg += eventBlob+'<br /><br />';
    msg += emailSig;
    GmailApp.sendEmail(to, subject, plainText, {
     from: 'shanghai.space.requests@nyu.edu',
     name: 'NYU Shanghai Space Requests',
     htmlBody: msg
    });
    return;
  } // end of stuff that happens when a new event is created
  
  else if (emailData.type==='Start Time Changed' || emailData.type==='End Time Changed' || emailData.type==='Rescheduled') {
    if (emailData.type==='Start Time Changed') {
      purposeOfEmail = 'The start time for <a href="'+emailData.url+'">'+name+'</a> has changed. '+seeBelow;
      plainText += 'Dear '+toGreeting+',\n\nThe start time for '+name+' has changed. For further details, see information below.\n\n';
    }
    else if (emailData.type==='End Time Changed') {
      purposeOfEmail = 'The end time for <a href="'+emailData.url+'">'+name+'</a> has changed. '+seeBelow;
      plainText += 'Dear '+toGreeting+',\n\nThe end time for '+name+' has changed. For further details, see information below.\n\n';
    }
    else if (emailData.type==='Rescheduled') {
      purposeOfEmail = '<a href="'+emailData.url+'">'+name+'</a> has been rescheduled. '+seeBelow;
      plainText += 'Dear '+toGreeting+',\n\n'+name+' has been rescheduled. For further details, see information below.\n\n';
    }
    // OLD START ------------------------------------------
    var oldStart = new Date(emailData.oldStart+28800000);
    var amPm = 'AM';
    var hours = oldStart.getHours();
    if (hours>=12) {amPm = 'PM';if(hours>12){hours = (hours-12);}}
    else if (parseInt(hours)==0) {hours=12;}
    var mins = ("0" + (oldStart.getMinutes())).slice(-2);
    var oldStartTime = hours + ':' + mins + ' ' + amPm; 
    var oldStartMonth = oldStart.getMonth();
    if (oldStartMonth == 0) {oldStartMonth = 'January';}
    if (oldStartMonth == 1) {oldStartMonth = 'February';}
    if (oldStartMonth == 2) {oldStartMonth = 'March';}
    if (oldStartMonth == 3) {oldStartMonth = 'April';}
    if (oldStartMonth == 4) {oldStartMonth = 'May';}
    if (oldStartMonth == 5) {oldStartMonth = 'June';}
    if (oldStartMonth == 6) {oldStartMonth = 'July';}
    if (oldStartMonth == 7) {oldStartMonth = 'August';}
    if (oldStartMonth == 8) {oldStartMonth = 'September';}
    if (oldStartMonth == 9) {oldStartMonth = 'October';}
    if (oldStartMonth == 10) {oldStartMonth = 'November';}
    if (oldStartMonth == 11) {oldStartMonth = 'December';}
    var oldStartDayOfWeek = oldStart.getDay();
    if (oldStartDayOfWeek == 0) {oldStartDayOfWeek = 'Sunday';}
    if (oldStartDayOfWeek == 1) {oldStartDayOfWeek = 'Monday';}
    if (oldStartDayOfWeek == 2) {oldStartDayOfWeek = 'Tuesday';}
    if (oldStartDayOfWeek == 3) {oldStartDayOfWeek = 'Wednesday';}
    if (oldStartDayOfWeek == 4) {oldStartDayOfWeek = 'Thursday';}
    if (oldStartDayOfWeek == 5) {oldStartDayOfWeek = 'Friday';}
    if (oldStartDayOfWeek == 6) {oldStartDayOfWeek = 'Saturday';}
    var oldStartD = oldStart.getDate();
    var oldStartY = oldStart.getFullYear();
    oldStart = oldStartDayOfWeek + ", " + oldStartMonth + " " + oldStartD + ", " + oldStartY;
    // NEW START ------------------------------------------
    var newStart = new Date(emailData.newStart+28800000);
    var amPm = 'AM';
    var hours = newStart.getHours();
    if (hours>=12) {amPm = 'PM';if (hours>12) {hours = (hours-12);}}
    else {if (parseInt(hours)==0) {hours=12;}}
    var mins = ("0" + (newStart.getMinutes())).slice(-2);
    var newStartTime = hours + ':' + mins + ' ' + amPm; 
    var newStartMonth = newStart.getMonth();
    if (newStartMonth == 0) {newStartMonth = 'January';}
    if (newStartMonth == 1) {newStartMonth = 'February';}
    if (newStartMonth == 2) {newStartMonth = 'March';}
    if (newStartMonth == 3) {newStartMonth = 'April';}
    if (newStartMonth == 4) {newStartMonth = 'May';}
    if (newStartMonth == 5) {newStartMonth = 'June';}
    if (newStartMonth == 6) {newStartMonth = 'July';}
    if (newStartMonth == 7) {newStartMonth = 'August';}
    if (newStartMonth == 8) {newStartMonth = 'September';}
    if (newStartMonth == 9) {newStartMonth = 'October';}
    if (newStartMonth == 10) {newStartMonth = 'November';}
    if (newStartMonth == 11) {newStartMonth = 'December';}
    var newStartDayOfWeek = newStart.getDay();
    if (newStartDayOfWeek == 0) {newStartDayOfWeek = 'Sunday';}
    if (newStartDayOfWeek == 1) {newStartDayOfWeek = 'Monday';}
    if (newStartDayOfWeek == 2) {newStartDayOfWeek = 'Tuesday';}
    if (newStartDayOfWeek == 3) {newStartDayOfWeek = 'Wednesday';}
    if (newStartDayOfWeek == 4) {newStartDayOfWeek = 'Thursday';}
    if (newStartDayOfWeek == 5) {newStartDayOfWeek = 'Friday';}
    if (newStartDayOfWeek == 6) {newStartDayOfWeek = 'Saturday';}
    var newStartD = newStart.getDate();
    var newStartY = newStart.getFullYear();
    newStart = newStartDayOfWeek + ", " + newStartMonth + " " + newStartD + ", " + newStartY;
    // OLD END ------------------------------------------
    var oldEnd = new Date(emailData.oldEnd+28800000);
    var amPm = 'AM';
    var hours = oldEnd.getHours();
    if (hours>=12) {amPm = 'PM';if(hours>12){hours = (hours-12);}}
    else if (parseInt(hours)==0) {hours=12;}
    var mins = ("0" + (oldEnd.getMinutes())).slice(-2);
    var oldEndTime = hours + ':' + mins + ' ' + amPm; 
    var oldEndMonth = oldEnd.getMonth();
    if (oldEndMonth == 0) {oldEndMonth = 'January';}
    if (oldEndMonth == 1) {oldEndMonth = 'February';}
    if (oldEndMonth == 2) {oldEndMonth = 'March';}
    if (oldEndMonth == 3) {oldEndMonth = 'April';}
    if (oldEndMonth == 4) {oldEndMonth = 'May';}
    if (oldEndMonth == 5) {oldEndMonth = 'June';}
    if (oldEndMonth == 6) {oldEndMonth = 'July';}
    if (oldEndMonth == 7) {oldEndMonth = 'August';}
    if (oldEndMonth == 8) {oldEndMonth = 'September';}
    if (oldEndMonth == 9) {oldEndMonth = 'October';}
    if (oldEndMonth == 10) {oldEndMonth = 'November';}
    if (oldEndMonth == 11) {oldEndMonth = 'December';}
    var oldEndDayOfWeek = oldEnd.getDay();
    if (oldEndDayOfWeek == 0) {oldEndDayOfWeek = 'Sunday';}
    if (oldEndDayOfWeek == 1) {oldEndDayOfWeek = 'Monday';}
    if (oldEndDayOfWeek == 2) {oldEndDayOfWeek = 'Tuesday';}
    if (oldEndDayOfWeek == 3) {oldEndDayOfWeek = 'Wednesday';}
    if (oldEndDayOfWeek == 4) {oldEndDayOfWeek = 'Thursday';}
    if (oldEndDayOfWeek == 5) {oldEndDayOfWeek = 'Friday';}
    if (oldEndDayOfWeek == 6) {oldEndDayOfWeek = 'Saturday';}
    var oldEndD = oldEnd.getDate();
    var oldEndY = oldEnd.getFullYear();
    oldEnd = oldEndDayOfWeek + ", " + oldEndMonth + " " + oldEndD + ", " + oldEndY;
    // NEW END ------------------------------------------
    var newEnd = new Date(emailData.newEnd+28800000);
    var hours = newEnd.getHours();
    var amPM = 'AM';
    if (hours>=12) {amPm = 'PM';if (hours>12) {hours = (hours-12);}}
    else {if (parseInt(hours)==0) {hours=12;}}
    var mins = ("0" + (newEnd.getMinutes())).slice(-2);
    var newEndTime = hours + ':' + mins + ' ' + amPm; 
    var newEndMonth = newEnd.getMonth();
    if (newEndMonth == 0) {newEndMonth = 'January';}
    if (newEndMonth == 1) {newEndMonth = 'February';}
    if (newEndMonth == 2) {newEndMonth = 'March';}
    if (newEndMonth == 3) {newEndMonth = 'April';}
    if (newEndMonth == 4) {newEndMonth = 'May';}
    if (newEndMonth == 5) {newEndMonth = 'June';}
    if (newEndMonth == 6) {newEndMonth = 'July';}
    if (newEndMonth == 7) {newEndMonth = 'August';}
    if (newEndMonth == 8) {newEndMonth = 'September';}
    if (newEndMonth == 9) {newEndMonth = 'October';}
    if (newEndMonth == 10) {newEndMonth = 'November';}
    if (newEndMonth == 11) {newEndMonth = 'December';}
    var newEndDayOfWeek = newEnd.getDay();
    if (newEndDayOfWeek == 0) {newEndDayOfWeek = 'Sunday';}
    if (newEndDayOfWeek == 1) {newEndDayOfWeek = 'Monday';}
    if (newEndDayOfWeek == 2) {newEndDayOfWeek = 'Tuesday';}
    if (newEndDayOfWeek == 3) {newEndDayOfWeek = 'Wednesday';}
    if (newEndDayOfWeek == 4) {newEndDayOfWeek = 'Thursday';}
    if (newEndDayOfWeek == 5) {newEndDayOfWeek = 'Friday';}
    if (newEndDayOfWeek == 6) {newEndDayOfWeek = 'Saturday';}
    var newEndD = newEnd.getDate();
    var newEndY = newEnd.getFullYear();
    newEnd = newEndDayOfWeek + ", " + newEndMonth + " " + newEndD + ", " + newEndY;
    // -------------------------------------------------- 
    if ((emailData.isAllDay)==true) {newStartTime='All Day';newEndTime='All Day';}
    var datesCheck = emailData.dates;
    
    if (emailData.isAllDay==true) {
      if ((parseInt(datesCheck.length))>10) {
        eventBlob += ('<b>Dates:</b> '+datesCheck+'<br />');
      }
      else {
        if (newStart==newEnd && oldStart==oldEnd) {
          eventBlob += '<b>Date:</b> ';
          if (oldStart===newStart) {eventBlob+= newStart+'<br />';}
          else {eventBlob += '</s>'+oldStart+'</s> '+newStart+'<br /><br />';}
        }
        else { // one of the start dates is different than one of the end dates
          eventBlob += '<b>Start Date:</b> ';
          if (oldStart===newStart) {eventBlob += newStart;}
          else {eventBlob += ('<s>'+oldStart+'</s> '+newStart);}
          eventBlob += '<br /><b>End Date:</b> ';
          if (oldEnd===newEnd) {eventBlob += newEnd}
          else{eventBlob += '<s>'+oldEnd+'</s> '+newEnd;}
        }
      }
      eventBlob += ('<br /><br /><b>Time:</b> '+'All Day'+'<br />');
      eventBlob += ('<br /><br /><b>SPACE REQUESTED:</b> '+emailData.requested+'<br /><br /><b>SPACE ASSIGNED:</b> '+emailData.assigned);
    }
    else { // it's not an all-day event
      if ((parseInt(datesCheck.length))>10) {
        eventBlob += ('<b>Dates:</b> '+datesCheck+'<br /><br />'+'<b>Start Time:</b> ');
        if (oldStartTime===newStartTime) {eventBlob+=newStartTime;}
        else {eventBlob += '<s>'+oldStartTime+'</s> '+newStartTime;}
        eventBlob += '<br /><b>End Time:</b> ';
        if (oldEndTime===newEndTime) {eventBlob+=newEndTime;}
        else {eventBlob += '<s>'+oldEndTime+'</s> '+newEndTime;}
        eventBlob += ('<br /><br /><b>SPACE REQUESTED:</b> '+emailData.requested+'<br /><br /><b>SPACE ASSIGNED:</b> '+emailData.assigned);
        plainText += 'Dates: '+datesCheck+'\nStart Time: '+newStartTime+'\nEnd Time: '+newEndTime+'\n\nSPACE REQUESTED: '+emailData.requested+'\n\nSPACE ASSIGNED:'+emailData.assigned+'\n\nLink:'+emailData.url;
      }
      else if (newStart===newEnd && oldStart===oldEnd) { // if both old and new start/end on the same date (but this date might have changed)
        eventBlob += '<b>Date:</b> ';
        if (oldStart===newStart) {eventBlob+= newStart;}
        else {eventBlob += '<s>'+oldStart+'</s> '+newStart;}
        eventBlob += '<br /><br /><b>Start Time:</b> ';
        if (oldStartTime===newStartTime) {eventBlob+=newStartTime;}
        else {eventBlob += '<s>'+oldStartTime+'</s> '+newStartTime;}
        eventBlob += '<br /><b>End Time:</b> ';
        if (oldEndTime===newEndTime) {eventBlob+=newEndTime;}
        else {eventBlob += '<s>'+oldEndTime+'</s> '+newEndTime;}
        eventBlob += ('<br /><br /><b>SPACE REQUESTED:</b> '+emailData.requested+'<br /><br /><b>SPACE ASSIGNED:</b> '+emailData.assigned);
        plainText += 'Start Date: '+newStart+'\nStart Time: '+newStartTime+'\nEnd Date: '+newEnd+'\nEnd Time: '+newEndTime+'\n\nSPACE REQUESTED: '+emailData.requested+'\n\nSpace Assigned: '+emailData.assigned+'\n\nLink:'+emailData.url; 
      }
      else if (newStart!==newEnd || oldStart!==oldEnd) { // the start date and end date are different in either the old or new version
        eventBlob += '<b>Start Date:</b> ';
        if (oldStart===newStart) {eventBlob += newStart;}
        else {eventBlob += ('<s>'+oldStart+'</s> '+newStart);}
        eventBlob += '<br /><b>Start Time:</b> ';
        if (oldStartTime===newStartTime) {eventBlob += newStartTime;}
        else{eventBlob += '<s>'+oldStartTime+'</s> '+newStartTime;}
        eventBlob += '<br /><br /><b>End Date:</b> ';
        if (oldEnd===newEnd) {eventBlob += newEnd}
        else{eventBlob += '<s>'+oldEnd+'</s> '+newEnd;}
        eventBlob += '<br /><b>End Time:</b> '; 
        if (oldEndTime===newEndTime) {eventBlob += newEndTime;}
        else {eventBlob += '<s>'+oldEndTime+'</s> '+newEndTime;}
        eventBlob += ('<br /><br /><b>SPACE REQUESTED:</b> '+emailData.requested+'<br /><br /><b>SPACE ASSIGNED:</b> '+emailData.assigned);
        plainText += 'Start Date: '+newStart+'\nStart Time: '+newStartTime+'\nEnd Date: '+newEnd+'\nEnd Time: '+newEndTime+'\n\nSPACE REQUESTED: '+emailData.requested+'\n\nSpace Assigned: '+emailData.assigned+'\n\nLink:'+emailData.url;
      }
    }
    msg += purposeOfEmail;
    msg += eventBlob+'<br /><br />';
    msg += emailSig;
    GmailApp.sendEmail(to, subject, plainText, {
     from: 'shanghai.space.requests@nyu.edu',
     name: 'NYU Shanghai Space Requests',
     htmlBody: msg
    });
    return; 
  } // end of stuff that happens when Rescheduled or Start/End Time Changed
  
  else if (emailData.type=='Cancelled') {
    
    name = (name.charAt(0).toUpperCase() + name.slice(1));
    purposeOfEmail = '<b>'+name+'</b> has been cancelled. This event will no longer take place.<br /><br />';
    plainText += 'Dear '+toGreeting+',\n\n'+name+' has been cancelled. For further details, see information below.\n\n';
    
    var startDate = new Date(emailData.start+28800000);
    var amPm = 'AM';
    var hours = startDate.getHours();
    if (hours>=12) {amPm = 'PM';if(hours>12){hours = (hours-12);}}
    else if (parseInt(hours)==0) {hours=12;}
    var mins = ("0" + (startDate.getMinutes())).slice(-2);
    var startTime = hours + ':' + mins + ' ' + amPm; 
    var startMonth = startDate.getMonth();
    if (startMonth == 0) {startMonth = 'January';}
    if (startMonth == 1) {startMonth = 'February';}
    if (startMonth == 2) {startMonth = 'March';}
    if (startMonth == 3) {startMonth = 'April';}
    if (startMonth == 4) {startMonth = 'May';}
    if (startMonth == 5) {startMonth = 'June';}
    if (startMonth == 6) {startMonth = 'July';}
    if (startMonth == 7) {startMonth = 'August';}
    if (startMonth == 8) {startMonth = 'September';}
    if (startMonth == 9) {startMonth = 'October';}
    if (startMonth == 10) {startMonth = 'November';}
    if (startMonth == 11) {startMonth = 'December';}
    var startDayOfWeek = startDate.getDay();
    if (startDayOfWeek == 0) {startDayOfWeek = 'Sunday';}
    if (startDayOfWeek == 1) {startDayOfWeek = 'Monday';}
    if (startDayOfWeek == 2) {startDayOfWeek = 'Tuesday';}
    if (startDayOfWeek == 3) {startDayOfWeek = 'Wednesday';}
    if (startDayOfWeek == 4) {startDayOfWeek = 'Thursday';}
    if (startDayOfWeek == 5) {startDayOfWeek = 'Friday';}
    if (startDayOfWeek == 6) {startDayOfWeek = 'Saturday';}
    var startD = startDate.getDate();
    var startY = startDate.getFullYear();
    startDate = startDayOfWeek + ", " + startMonth + " " + startD + ", " + startY;
    
    var endDate = new Date(emailData.end+28800000);
    var amPm = 'AM';
    var hours = endDate.getHours();
    if (hours>=12) {amPm = 'PM';if(hours>12){hours = (hours-12);}}
    else if (parseInt(hours)==0) {hours=12;}
    var mins = ("0" + (endDate.getMinutes())).slice(-2);
    var endTime = hours + ':' + mins + ' ' + amPm; 
    var endMonth = endDate.getMonth();
    if (endMonth == 0) {endMonth = 'January';}
    if (endMonth == 1) {endMonth = 'February';}
    if (endMonth == 2) {endMonth = 'March';}
    if (endMonth == 3) {endMonth = 'April';}
    if (endMonth == 4) {endMonth = 'May';}
    if (endMonth == 5) {endMonth = 'June';}
    if (endMonth == 6) {endMonth = 'July';}
    if (endMonth == 7) {endMonth = 'August';}
    if (endMonth == 8) {endMonth = 'September';}
    if (endMonth == 9) {endMonth = 'October';}
    if (endMonth == 10) {endMonth = 'November';}
    if (endMonth == 11) {endMonth = 'December';}
    var endDayOfWeek = endDate.getDay();
    if (endDayOfWeek == 0) {endDayOfWeek = 'Sunday';}
    if (endDayOfWeek == 1) {endDayOfWeek = 'Monday';}
    if (endDayOfWeek == 2) {endDayOfWeek = 'Tuesday';}
    if (endDayOfWeek == 3) {endDayOfWeek = 'Wednesday';}
    if (endDayOfWeek == 4) {endDayOfWeek = 'Thursday';}
    if (endDayOfWeek == 5) {endDayOfWeek = 'Friday';}
    if (endDayOfWeek == 6) {endDayOfWeek = 'Saturday';}
    var endD = endDate.getDate();
    var endY = endDate.getFullYear();
    endDate = endDayOfWeek + ", " + endMonth + " " + endD + ", " + endY;
    
    var datesCheck = emailData.dates;
    if ((parseInt(datesCheck.length))>10) {
      eventBlob += '<b>Dates:</b> '+datesCheck;
      eventBlob += '<br /><br /><b>Start Time:</b> '+startTime+'<br />';
      eventBlob += '<b>End Time:</b> '+endTime+'<br />';
    }
    else if (emailData.isAllDay==true) {
      if ((parseInt(datesCheck.length))>10) {
        eventBlob += ('<b>Dates:</b> '+datesCheck+'<br />');
        plainText += 'Dates: '+datesCheck;
      }
      else {
        if (startDate===endDate) {
          eventBlob += ('<b>Date:</b> '+startDate+'<br />');
          plainText += 'Date: '+startDate;
        }
        else { // the start date and end date are different
          eventBlob += '<b>Start Date:</b> '+startDate+'<br / >';
          eventBlob += '<b>End Date:</b> '+endDate+'<br / >';
          plainText += 'Start Date: '+startDate+'\nEnd Date: '+endDate;
        }
      }
      eventBlob += ('<br /><b>Time:</b> '+'All Day');
      plainText += '\nStart Time: '+'All Day';
    }
    
    else { // if it's not an all day event
      if (startDate==endDate) {
        eventBlob += '<b>Date:</b> '+startDate+'<br /><br />';
        eventBlob += '<b>Start Time:</b> '+startTime+'<br />';
        eventBlob += '<b>End Time:</b> '+endTime+'<br />';
        plainText += 'Date: '+startDate+'\nStart Time: '+startTime+'\nEnd Time: '+endTime;
      }
      else {
        eventBlob += '<b>Start Date:</b> '+startDate+'<br / >';
        eventBlob += '<b>Start Time:</b> '+startTime+'<br /><br />';
        eventBlob += '<b>End Date:</b> '+endDate+'<br / >';
        eventBlob += '<b>End Time:</b> '+endTime+'<br / >';
        plainText += 'Start Date: '+startDate+'\nStart Time: '+startTime+'\nEnd Date: '+endDate+'\nEnd Time: '+endTime;
      }
    }
    // here
    eventBlob += '<br /><b>SPACE REQUESTED:</b> '+emailData.requested;
    eventBlob += '<br /><b>SPACE ASSIGNED:</b> '+emailData.assigned;
    plainText += '\n\nSPACE REQUESTED: '+emailData.requested+'\n\nSpace Assigned: '+emailData.assigned+'\n\nLink:'+emailData.url;
    
    msg += purposeOfEmail;
    msg += eventBlob+'<br /><br />';
    msg += emailSig;
    GmailApp.sendEmail(to, subject, plainText, {
     from: 'shanghai.space.requests@nyu.edu',
     name: 'NYU Shanghai Space Requests',
     htmlBody: msg
    });
    return; 
  } // end of cancellation email code
} // end of sendEmail


function processSheet (data) {
  var d = new Date();
  var msecNow = d.getTime();
  msecNow = parseInt(msecNow)-28800000;
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetClassrooms = ss.getSheetByName("Classroom Requests");
  var sheet15 = ss.getSheetByName("Other Requests");
  var sheetAssigned = ss.getSheetByName("Location Assigned");
  var lastRowClassrooms = sheetClassrooms.getLastRow();
  var lastColClassrooms = sheetClassrooms.getLastColumn();
  var lastRow15 = sheet15.getLastRow();
  var lastCol15 = sheet15.getLastColumn();
  var lastRowAssigned = sheetAssigned.getLastRow();
  var lastColAssigned = sheetAssigned.getLastColumn();
  if (lastRowClassrooms > 4) { sheetClassrooms.getRange("M5:M").setValue("No"); }
  if (lastRow15 > 4) { sheet15.getRange("M5:M").setValue("No"); }
  if (lastRowAssigned > 4) { sheetAssigned.getRange("M5:M").setValue("No"); }
  for (i in data) {
    var whichSheet;
    var otherSheet;
    var whichRow;
    var found = false;
    var needToEmail = false;
    var to;
    var suspiciousLocation = false;
    var locationString = '' + data[i].location;
    if ( (locationString.search("TBD") != -1) || (locationString.search("tbd") != -1) || (locationString.search("Tbd") != -1) || (locationString.search("be determin") != -1) || (locationString.search("o be confirm") != -1) || (locationString.search("o Be Confirm") != -1) || (locationString.search("o be Confirm") != -1) || (locationString.search("ssign") != -1) || (locationString.search("TBA") != -1) || (locationString.search("tba") != -1) || (locationString.search("Tba") != -1) || (locationString.search("nnounce") != -1) || (locationString.search("TBC") != -1) || (locationString.search("tbc") != -1) || (locationString.search("Tbc") != -1) ) { // looking for a bunch of suspicious words in the location, which might indicate that the room reservation hasn't actually been approved by the registrar yet
      suspiciousLocation = true;
    }
    whichSheet = sheetClassrooms;
    var theseIDs = new Array();
    if (lastRowClassrooms > 4) {theseIDs = whichSheet.getRange(5,11,(lastRowClassrooms-4)).getValues();}
    var numRemoved = 0;
    for (x in theseIDs) {
      var thisIdFromSht = theseIDs[x][0];
      if (data[i].eventID == thisIdFromSht) {
        found = true;
        whichRow = parseInt(x)+parseInt(5)-parseInt(numRemoved);
        if (data[i].endCode < msecNow) { // if the event has already passed, delete it
          whichSheet.deleteRow(whichRow);
          lastRowClassrooms = lastRowClassrooms-1;
          numRemoved = numRemoved+1;
          break;
        }
        var numColumns = lastColClassrooms;
        var range = whichSheet.getRange(whichRow,1,1,numColumns);
        var rowPull = range.getValues();
        rowPull[0][0] = data[i].name;
        rowPull[0][2] = data[i].location;
        rowPull[0][3] = data[i].dates;
        rowPull[0][4] = data[i].startTime;
        rowPull[0][5] = data[i].endTime;
        rowPull[0][6] = data[i].startCode;
        rowPull[0][7] = data[i].endCode;
        rowPull[0][12] = 'Yes';
        range.setValues(rowPull);
        if ( (data[i].location != '') && (data[i].location != 'No location has been entered for this event.') && (suspiciousLocation != true) ) { // if there is a legit location in the data
          otherSheet = sheetAssigned;
          var numRows = lastRowAssigned;
          otherSheet.insertRowsAfter(numRows, 1);
          range.moveTo(otherSheet.getRange(numRows + 1, 1));
          whichSheet.deleteRow(whichRow);
          lastRowClassrooms = lastRowClassrooms-1;
          lastRowAssigned = lastRowAssigned+1;
          numRemoved = numRemoved+1;
        }
        break;
      }
    }
    if (found == false) {
      whichSheet = sheet15;
      var theseIDs = new Array();
      if (lastRow15 > 4) {theseIDs = whichSheet.getRange(5,11,(lastRow15-4)).getValues();}
      var numRemoved = 0;
      for (y in theseIDs) {
        var thisIdFromSht = theseIDs[y][0];
        if (data[i].eventID == thisIdFromSht) {
          found = true;
          whichRow = parseInt(y)+parseInt(5)-parseInt(numRemoved);
          if (data[i].endCode < msecNow) { // if the event has already passed, delete it
            whichSheet.deleteRow(whichRow);
            lastRow15 = lastRow15-1;
            numRemoved = numRemoved+1;
            break;
          }
          var numColumns = lastCol15;
          var range = whichSheet.getRange(whichRow,1,1,numColumns);
          var rowPull = range.getValues();
          rowPull[0][0] = data[i].name;
          rowPull[0][2] = data[i].location;
          rowPull[0][3] = data[i].dates;
          rowPull[0][4] = data[i].startTime;
          rowPull[0][5] = data[i].endTime;
          rowPull[0][6] = data[i].startCode;
          rowPull[0][7] = data[i].endCode;
          rowPull[0][12] = 'Yes';
          range.setValues(rowPull);
          if ( (data[i].location != '') && (data[i].location != 'No location has been entered for this event.') && (suspiciousLocation != true) ) { // if there is a legit location in the data
            otherSheet = sheetAssigned;
            var numRows = lastRowAssigned;
            otherSheet.insertRowsAfter(numRows, 1);
            range.moveTo(otherSheet.getRange(numRows + 1, 1));
            whichSheet.deleteRow(whichRow);
            lastRow15 = lastRow15-1;
            lastRowAssigned = lastRowAssigned+1;
            numRemoved = numRemoved+1;
          }
          break;
        }
      }
    }
    if (found == false) { // go through assigned sheet
      whichSheet = sheetAssigned;
      var theseIDs = new Array();
      if (lastRowAssigned > 4) {theseIDs = whichSheet.getRange(5,11,(lastRowAssigned-4)).getValues();}
      var numRemoved = 0;
      for (z in theseIDs) {
        var thisIdFromSht = theseIDs[z][0];
        if (data[i].eventID == thisIdFromSht) { // go through Assigned sheet looking for this event. if found, do stuff depending on some other stuff. when finished doing stuff, break (stop looking for this event in another sheet)
          found = true;
          whichRow = parseInt(z)+parseInt(5)-parseInt(numRemoved);
          if (data[i].endCode < msecNow) { // delete old event
            whichSheet.deleteRow(whichRow);
            lastRowAssigned = lastRowAssigned-1;
            numRemoved = numRemoved+1;
            break;
          }
          var numColumns = lastColAssigned;
          var range = whichSheet.getRange(whichRow,1,1,numColumns);
          var rowPull = range.getValues();
          rowPull[0][0] = data[i].name;
          rowPull[0][2] = data[i].location;
          rowPull[0][3] = data[i].dates;
          rowPull[0][4] = data[i].startTime;
          rowPull[0][5] = data[i].endTime;
          rowPull[0][6] = data[i].startCode;
          rowPull[0][7] = data[i].endCode;
          rowPull[0][12] = 'Yes';
          range.setValues(rowPull);
          if (data[i].location == '' || data[i].location == 'No location has been entered for this event.' || data[i].location == '' || suspiciousLocation == true) { // check if location in data is empty or suspicious; if it is, move back to another sheet
            var typeRequested = rowPull[0][1];
            var indicator = 0;
            var numRows;
            if ( (typeRequested.search("15") != -1) || (typeRequested.search("Art Gallery") != -1) || (typeRequested.search("Cafe") != -1) ) {
              otherSheet = sheet15;
              indicator = 1;
              numRows=lastRow15;
            }
            else {
              if ( (typeRequested.search("Classroom") != -1) || (typeRequested.search("Auditorium") != -1) || (typeRequested.search("Video Conference") != -1) || (typeRequested.search("Computer Lab") != -1) || (typeRequested.search("Geography Building") != -1) ) {
                otherSheet = sheetClassrooms;
                indicator = 2;
                numRows=lastRowClassrooms;
              }
            }
            if (indicator !=0) {
              otherSheet.insertRowsAfter(numRows, 1);
              range.moveTo(otherSheet.getRange(numRows + 1, 1));
              whichSheet.deleteRow(whichRow);
              lastRowAssigned = lastRowAssigned-1;
              numRemoved = numRemoved+1;
            }
          }
          break;
        }
      }
    }
    if ( (found ==  false) && (data[i].approved == true) && (data[i].endCode >= msecNow) ) { // if it wasn't found in any sheet, and the end hasn't already passed, and it's an approved event
      var string = '';
      var url = "https://api.orgsync.com/api/v2/events/" + data[i].eventID + ".json?key=PutApiKeyHere"; // API call to get info about what kind of room is needed
      var response = UrlFetchApp.fetch(url);
      var json = response.getContentText();
      var event = JSON.parse(json);
      var typeRequested = 'N/A';
      var setupTime = 'None';
      var cleanupTime = 'None';
      var specialRequirements = 'None';
      whichSheet = sheetAssigned;
      indicator = 3;
      for (q in event.request_responses) {
        if (event.request_responses[q].element.name == "Space reservation needs:") {
          if ( (event.request_responses[q].data.name == "Pudong Campus Building - Classroom / Auditorium") || (event.request_responses[q].data.name == "Pudong Campus Building - Classroom") || (event.request_responses[q].data.name == 'ECNU Geography Building - Classroom') ) {
            needToEmail = true;
            to = 'registrar';
            for (r in event.request_responses) {
              if ( (event.request_responses[r].data) && (event.request_responses[r].data.name) && (event.request_responses[r].element.name) && ((event.request_responses[r].element.name=='Please indicate the type of space you would like to request:') || (event.request_responses[r].element.name=='Please indicate the type of classroom you would like to request:')) ) {
                typeRequested = event.request_responses[r].data.name;
                continue;
              }
              if ( (event.request_responses[r].data) && (event.request_responses[r].element.name) && (event.request_responses[r].element.name=='Please estimate the amount of SETUP time you will need to prepare the space (if any):') && (event.request_responses[r].data != '') ) {
                setupTime = event.request_responses[r].data;
                continue;
              }
              if ( (event.request_responses[r].data) && (event.request_responses[r].element.name) && (event.request_responses[r].element.name=='Please estimate the amount of CLEANUP time you will require (if any):') && (event.request_responses[r].data != '') ) {
                cleanupTime = event.request_responses[r].data;
                continue;
              }
              if ( (event.request_responses[r].data) && (event.request_responses[r].element.name) && (event.request_responses[r].element.name=='Special requirements:') && (event.request_responses[r].data != '') && (event.request_responses[r].data != null) && (event.request_responses[r].data != undefined) ) {
                specialRequirements = event.request_responses[r].data;
                continue;
              }
            }
            if ( (data[i].location == '') || (data[i].location == 'No location has been entered for this event.') || (suspiciousLocation==true) ) {
              whichSheet = sheetClassrooms;
              indicator = 1;
            }
          }
          else if (event.request_responses[q].data.name == "Pudong Campus Building - Auditorium") {
            needToEmail = true;
            to = 'registrar';
            typeRequested = 'Auditorium (300 Person) - Fixed Seating';
            for (r in event.request_responses) {
              if ( (event.request_responses[r].data) && (event.request_responses[r].element.name) && (event.request_responses[r].element.name=='Please estimate the amount of SETUP time you will need to prepare the space (if any):') && (event.request_responses[r].data != '') ) {
                setupTime = event.request_responses[r].data;
                continue;
              }
              if ( (event.request_responses[r].data) && (event.request_responses[r].element.name) && (event.request_responses[r].element.name=='Please estimate the amount of CLEANUP time you will require (if any):') && (event.request_responses[r].data != '') ) {
                cleanupTime = event.request_responses[r].data;
                continue;
              }
              if ( (event.request_responses[r].data) && (event.request_responses[r].element.name) && (event.request_responses[r].element.name=='Special requirements:') && (event.request_responses[r].data != '') && (event.request_responses[r].data != null) && (event.request_responses[r].data != undefined) ) {
                specialRequirements = event.request_responses[r].data;
                continue;
              }
            }
            if ( (data[i].location == '') || (data[i].location == 'No location has been entered for this event.') || (suspiciousLocation==true) ) {
              whichSheet = sheetClassrooms;
              indicator = 1;
            }
          }
          else {
            var requested = (event.request_responses[q].data.name);
            if ( (requested == "Pudong Campus Building - 15th Floor Colloquium Space") || (requested == "Pudong Campus Building - Art Gallery") || (requested == "Pudong Campus Building - B1 Cafeteria") || (requested == "Pudong Campus Building - 2nd Floor Cafe") ) {
              typeRequested = ''+(requested.substring(25)); // the substring thing removes 'Pudong Campus Building - ' from the front of the string
              to = 'facilities';
              needToEmail = true;
              for (r in event.request_responses) {
                if ( (event.request_responses[r].data) && (event.request_responses[r].element.name) && (event.request_responses[r].element.name=='Please estimate the amount of SETUP time you will need to prepare the space (if any):') && (event.request_responses[r].data != '') ) {
                  setupTime = event.request_responses[r].data;
                  continue;
                }
                if ( (event.request_responses[r].data) && (event.request_responses[r].element.name) && (event.request_responses[r].element.name=='Please estimate the amount of CLEANUP time you will require (if any):') && (event.request_responses[r].data != '') ) {
                  cleanupTime = event.request_responses[r].data;
                  continue;
                }
                if ( (event.request_responses[r].data) && (event.request_responses[r].element.name) && (event.request_responses[r].element.name=='Special requirements:') && (event.request_responses[r].data != '') && (event.request_responses[r].data != null) && (event.request_responses[r].data != undefined) ) {
                  specialRequirements = event.request_responses[r].data;
                  continue;
                }
              }
              if ( (data[i].location == '') || (data[i].location == 'No location has been entered for this event.') || (suspiciousLocation==true) ) {
                whichSheet = sheet15;
                indicator = 2;
              }
            }
            else if (requested == "Pudong Campus Building - 8th Floor Dance Studio") { // This is intentionally not the same as "8th Floor Dance Room" (which is on the Student Clubs & Organizations form), so that Clubs requests don't go to Athletics (they are handled by the club advisros instead)
              typeRequested = ''+(requested.substring(25));
              to = 'athletics';
              needToEmail = true;
              for (r in event.request_responses) {
                if ( (event.request_responses[r].data) && (event.request_responses[r].element.name) && (event.request_responses[r].element.name=='Please estimate the amount of SETUP time you will need to prepare the space (if any):') && (event.request_responses[r].data != '') ) {
                  setupTime = event.request_responses[r].data;
                  continue;
                }
                if ( (event.request_responses[r].data) && (event.request_responses[r].element.name) && (event.request_responses[r].element.name=='Please estimate the amount of CLEANUP time you will require (if any):') && (event.request_responses[r].data != '') ) {
                  cleanupTime = event.request_responses[r].data;
                  continue;
                }
                if ( (event.request_responses[r].data) && (event.request_responses[r].element.name) && (event.request_responses[r].element.name=='Special requirements:') && (event.request_responses[r].data != '') && (event.request_responses[r].data != null) && (event.request_responses[r].data != undefined) ) {
                  specialRequirements = event.request_responses[r].data;
                  continue;
                }
              }
            }
          }
        }     
      }
      
      if (indicator == 1) {lastRowClassrooms=lastRowClassrooms+1;}
      else {if (indicator == 2) {lastRow15=lastRow15+1;}
      else {if (indicator == 3) {lastRowAssigned=lastRowAssigned+1;}}}
      whichSheet.appendRow([data[i].name, typeRequested, data[i].location, data[i].dates, data[i].startTime, data[i].endTime, data[i].startCode, data[i].endCode, data[i].startCode, data[i].endCode, data[i].eventID, data[i].orgID, 'Yes', data[i].url, '']);
      var isAllDay = false;
      if (data[i].startTime==='All Day') {isAllDay=true;}
      if (needToEmail==true) {
        var type = 'New Space Request';
        var eventData = {
          name:data[i].name,
          id:data[i].eventID,
          requested:typeRequested,
          assigned:data[i].location,
          url:data[i].url,
          type:type,
          to:to,
          isAllDay:isAllDay,
          dates:data[i].dates,
          start:data[i].startCode,
          end:data[i].endCode,
          setupTime:setupTime,
          cleanupTime:cleanupTime,
          specialRequirements:specialRequirements
        };
        sendEmail(eventData);
      }
 
    } // end of stuff that happens if it wasn't found on any sheet
  }
  // stuff in this level happens after it has finished going through the OrgSync data
  
  checkTimeChanges(ss, sheetClassrooms, sheet15, sheetAssigned, lastRowClassrooms, lastColClassrooms, lastRow15, lastCol15, lastRowAssigned, lastColAssigned);
  
//  Some parts of deleteStuff (the function called below) are turned off because the OrgSync API pull does not include any pending (un-approved) events, and if an event is revised and isn't immediately re-approved, it looks like it has been deleted.
  deleteStuff(ss, sheetClassrooms, sheet15, sheetAssigned, lastRowClassrooms, lastColClassrooms, lastRow15, lastCol15, lastRowAssigned, lastColAssigned);


  var dateFormat = 'yyyy-MM-dd';
  var timeFormat = 'h:mm AM/PM';
//  var numSheets = ss.getNumSheets();
//  this for loop causes some bugs on Google's end
//  for (m=0; m<(numSheets-1); m++) {
//    Logger.log(numSheets + '');
//    ss.getSheets()[m].getRange('C2:C').setNumberFormat(dateFormat);
//    ss.getSheets()[m].getRange('D2:D').setNumberFormat(timeFormat);
//    ss.getSheets()[m].getRange('E2:E').setNumberFormat(timeFormat);
//    ss.getSheets()[m].getRange('A5:O').sort(3);
//  }
 
  sheetClassrooms.getRange('D2:D').setNumberFormat('yyyy-MM-dd');
  sheetClassrooms.getRange('E2:E').setNumberFormat('h:mm AM/PM');
  sheetClassrooms.getRange('F2:F').setNumberFormat('h:mm AM/PM');
  if (lastRowClassrooms>4) {
    var allClassroomStuff = sheetClassrooms.getRange('A5:O');
    allClassroomStuff.sort([4,5]);
    allClassroomStuff.setWrap(false);
  }
  
  sheet15.getRange('D2:D').setNumberFormat('yyyy-MM-dd');
  sheet15.getRange('E2:E').setNumberFormat('h:mm AM/PM');
  sheet15.getRange('F2:F').setNumberFormat('h:mm AM/PM');
  if (lastRow15>4) {
    var all15Stuff = sheet15.getRange('A5:O');
    all15Stuff.sort([4,5]);
    all15Stuff.setWrap(false);
  }
  
  sheetAssigned.getRange('D2:D').setNumberFormat('yyyy-MM-dd');
  sheetAssigned.getRange('E2:E').setNumberFormat('h:mm AM/PM');
  sheetAssigned.getRange('F2:F').setNumberFormat('h:mm AM/PM');
  if (lastRowAssigned>4) {
    var allAssignedStuff = sheetAssigned.getRange('A5:O');
    allAssignedStuff.sort([4,5]);
    allAssignedStuff.setWrap(false);
  }
}


function deleteStuff(ss, sheetClassrooms, sheet15, sheetAssigned, lastRowClassrooms, lastColClassrooms, lastRow15, lastCol15, lastRowAssigned, lastColAssigned) {
//  var ss = SpreadsheetApp.getActiveSpreadsheet();
//  var sheetClassrooms = ss.getSheetByName("Classroom Requests");
//  var sheet15 = ss.getSheetByName("Other Requests");
//  var sheetAssigned = ss.getSheetByName("Location Assigned");
//  var lastRowClassrooms = sheetClassrooms.getLastRow();
//  var lastColClassrooms = sheetClassrooms.getLastColumn();
//  var lastRow15 = sheet15.getLastRow();
//  var lastCol15 = sheet15.getLastColumn();
//  var lastRowAssigned = sheetAssigned.getLastRow();
//  var lastColAssigned = sheetAssigned.getLastColumn();
  
  if (lastRowClassrooms > 4) {var classroomsFound = sheetClassrooms.getRange(5, 13, (lastRowClassrooms-4), 1).getValues();}
  if (lastRow15 > 4) {var fifteenFound = sheet15.getRange(5,13,(lastRow15-4),1).getValues();}
  if (lastRowAssigned > 4) {var assignedFound = sheetAssigned.getRange(5, 13, (lastRowAssigned-4), 1).getValues();}
  
  var date = new Date();
  var msec = date.getTime(); // this is the current time in UTC (needed because if an event is deleted because it passed, then no email is necessary)
  msec = parseInt(msec)-28800000;
  var numRemovedClassrooms = 0;
  if (classroomsFound) {
    for (i in classroomsFound) {
      if (classroomsFound[i] == 'No') {
        var rowToDelete = parseInt(i)+parseInt(5)-parseInt(numRemovedClassrooms);
        var numColumns = lastColClassrooms;
        var whichSheet = sheetClassrooms;
        var range = whichSheet.getRange(rowToDelete,1,1,numColumns);
        var rowPull = range.getValues();
        if (rowPull[0][7] < msec) {
          sheetClassrooms.deleteRow(rowToDelete);
          numRemovedClassrooms = numRemovedClassrooms + 1;
          continue;
        }
//        This is the code that checks for deleted events and sends an email. Re-enable this if OrgSync starts providing "Pending" events in the JSON data from the API.
//        if (rowPull[0][7] > msec) { // if the event has not already ended
//          var isAllDay = false;
//          if (rowPull[0][4] == 'All Day') {isAllDay = true;}
//          var to = 'registrar';
//          var emailData = {
//            name:rowPull[0][0],
//            id:rowPull[0][10],
//            requested:rowPull[0][1],
//            assigned:rowPull[0][2],
//            start:rowPull[0][6],
//            end:rowPull[0][7],
//            to:to,
//            isAllDay:isAllDay,
//            dates:rowPull[0][3],
//            type:'Cancelled'
//          }
//          sendEmail(emailData);
//        }
//        sheetClassrooms.deleteRow(rowToDelete);
//        numRemovedClassrooms = numRemovedClassrooms + 1;
      }
    }
  }
  var numRemoved15 = 0;
  if (fifteenFound) {
    for (i in fifteenFound) {
      if (fifteenFound[i] == 'No') {
        var rowToDelete = parseInt(i)+parseInt(5)-parseInt(numRemoved15);
        var numColumns = lastCol15;
        var whichSheet = sheet15;
        var range = whichSheet.getRange(rowToDelete,1,1,numColumns);
        var rowPull = range.getValues();
        if (rowPull[0][7] < msec) {
          sheet15.deleteRow(rowToDelete);
          numRemoved15 = numRemoved15 + 1;
          continue;
        }
//        if (rowPull[0][7] > msec) { // if the event has not already ended
//          var isAllDay = false;
//          if (rowPull[0][4] == 'All Day') {isAllDay = true;}
//          var to = 'facilities';
//          var emailData = {
//            name:rowPull[0][0],
//            id:rowPull[0][10],
//            requested:rowPull[0][1],
//            assigned:rowPull[0][2],
//            start:rowPull[0][6],
//            end:rowPull[0][7],
//            to:to,
//            isAllDay:isAllDay,
//            dates:rowPull[0][3],
//            type:'Cancelled'
//          }
//          sendEmail(emailData);
//        }
//        sheet15.deleteRow(rowToDelete);
//        numRemoved15 = numRemoved15 + 1;
      }
    }
  }
  var numRemovedAssigned = 0;
  if (assignedFound) {
    for (i in assignedFound) {
      if (assignedFound[i] == 'No') {
        var rowToDelete = parseInt(i)+parseInt(5)-parseInt(numRemovedAssigned);
        
        var numColumns = lastColAssigned;
        var whichSheet = sheetAssigned;
        var range = whichSheet.getRange(rowToDelete,1,1,numColumns);
        var rowPull = range.getValues();
        if (rowPull[0][7] < msec) {
          sheetAssigned.deleteRow(rowToDelete);
          numRemovedAssigned = numRemovedAssigned + 1;
          continue;
        }
//        if (rowPull[0][7] > msec) { // if the event has not already ended
//          var to = 'N/A';
//          if ( (rowPull[0][1].search("Classroom") != -1) || (rowPull[0][1].search("Auditorium") != -1) || (rowPull[0][1].search("Video Conference") != -1) || (rowPull[0][1].search("Computer Lab") != -1) || (rowPull[0][1].search("Other Space") != -1) || (rowPull[0][1].search("ECNU") != -1) ) {
//            to = 'registrar';
//          }
//          else if ( (rowPull[0][1].search("15") != -1) || (rowPull[0][1].search("Art Gallery") != -1) || (rowPull[0][1].search("Cafe") != -1) ) {
//            to = 'facilities';
//          }
//          else if ( (rowPull[0][1].search("Dance Studio") != -1) ) {
//            to = 'facilities';
//          }
//          if (to != 'N/A') {
//            var isAllDay = false;
//            if (rowPull[0][4] == 'All Day') {isAllDay = true;}
//            var emailData = {
//              name:rowPull[0][0],
//              id:rowPull[0][10],
//              requested:rowPull[0][1],
//              assigned:rowPull[0][2],
//              start:rowPull[0][6],
//              end:rowPull[0][7],
//              to:to,
//              isAllDay:isAllDay,
//              dates:rowPull[0][3],
//              type:'Cancelled'
//            }
//            sendEmail(emailData);
//          }
//        }
//        sheetAssigned.deleteRow(rowToDelete);
//        numRemovedAssigned = numRemovedAssigned + 1;
      }
    }
  }
}


function getInfoFromOrgSync() {
  var eventIDs = [];
  var response = UrlFetchApp.fetch("https://api.orgsync.com/api/v2/orgs.json?key=PutApiKeyHere"); // Fetch list of all portals
  var json = response.getContentText();
  var formData = JSON.parse(json);
  var portalIDs = [];
  for (i in formData) {
    if (formData[i].is_disabled == false) { // This was revised this to only ignore disabled portals, and to *** not ignore Res Life events *** since they might ask for classrooms
      portalIDs.push(formData[i].id);
    }
  }
  var infoForSheet = [];
  for (i in portalIDs) {
    var url = "https://api.orgsync.com/api/v2/orgs/" + portalIDs[i] + "/events.json?key=PutApiKeyHere";
    var response = UrlFetchApp.fetch(url);
    var json = response.getContentText();
    var formData = JSON.parse(json);
    
//     if (portalIDs[i] == 85022) {
//       Logger.log('Techno portal: '+ json);
//     }
    
    for (j in formData) { // For each event
      var startTime;
      var endTime;
      var numOccurrences = 0;
      var dates = '';
      var endCode;
      var startUTCDate; //  = new Date((formData[j].occurrences[0].starts_at || "").replace(/-/g,"/").replace(/[TZ]/g," ")) // formatting junk so the date constructor can understand the OrgSync data
      var offsetMsecDate; //  = (startUTCDate.getTime() + 28800000); this gets the milliseconds for the date in China local time
      var startDate; // = new Date(offsetMsecDate)
      var startCode; // = startUTCDate.getTime(); // this is in milliseconds
      for (k in formData[j].occurrences) { // For each occurence of that event
        numOccurrences += 1;
        startUTCDate = new Date((formData[j].occurrences[k].starts_at || "").replace(/-/g,"/").replace(/[TZ]/g," "));
        offsetMsecDate = (startUTCDate.getTime() + 28800000);
        startDate = new Date(offsetMsecDate);
        if (numOccurrences == 1) {
          startCode = startUTCDate.getTime();
          startTime = startDate.getHours() + ':' + startDate.getUTCMinutes();
        }
        var startD = startDate.getDate();
        startD = ("0" + startD).slice(-2);
        var startM = startDate.getMonth() + 1;
        startM = ("0" + startM).slice(-2);
        var startY = startDate.getFullYear();
        startDate = startY + "-" + startM + "-" + startD;
        if (k>0) { dates = dates + ', ' + startDate;}
        else { dates = dates + startDate;}
        var endUTCDate = new Date((formData[j].occurrences[k].ends_at || "").replace(/-/g,"/").replace(/[TZ]/g," "));
        var offsetMsecEnd = (endUTCDate.getTime() + 28800000);
        var endDate = new Date(offsetMsecEnd);
        endTime = endDate.getHours() + ':' + endDate.getMinutes();
        endCode = endUTCDate.getTime();
      }
      if (formData[j].occurrences[0].is_all_day) {startTime = 'All Day';endTime = 'All Day';}
      var newInfoForSheet = [{
        eventID:formData[j].id,
        orgID:formData[j].org_id,
        url:('https://orgsync.com/' + formData[j].org_id + '/events/' + formData[j].id + '/request#conversation'),
        name:formData[j].name,
        startTime:startTime,
        startCode:startCode,
        endTime:endTime,
        endCode:endCode,
        location:formData[j].location,
        dates:dates,
        approved:formData[j].is_approved
      }];
      infoForSheet = infoForSheet.concat(newInfoForSheet);
    
    }
  } // finished going through portals and making giant array of JSON objects for sheet
  return infoForSheet;
}



function checkTimeChanges(ss, sheetClassrooms, sheet15, sheetAssigned, lastRowClassrooms, lastColClassrooms, lastRow15, lastCol15, lastRowAssigned, lastColAssigned) {
//  var ss = SpreadsheetApp.getActiveSpreadsheet();
//  var sheetClassrooms = ss.getSheetByName("Classroom Requests");
//  var sheet15 = ss.getSheetByName("Other Requests");
//  var sheetAssigned = ss.getSheetByName("Location Assigned");
//  var lastRowClassrooms = sheetClassrooms.getLastRow();
//  var lastColClassrooms = sheetClassrooms.getLastColumn();
//  var lastRow15 = sheet15.getLastRow();
//  var lastCol15 = sheet15.getLastColumn();
//  var lastRowAssigned = sheetAssigned.getLastRow();
//  var lastColAssigned = sheetAssigned.getLastColumn();
  
  var classroomTimeCodes = [[]];
  var fifteenTimeCodes = [[]];
  var assignedTimeCodes = [[]];
  
  if (lastRowClassrooms>4) {var classroomTimeCodes = sheetClassrooms.getRange(5, 7, (lastRowClassrooms-4), 4).getValues();}
  if (lastRow15>4) {var fifteenTimeCodes = sheet15.getRange(5, 7, (lastRow15-4), 4).getValues();}
  if (lastRowAssigned>4) {var assignedTimeCodes = sheetAssigned.getRange(5, 7, (lastRowAssigned-4), 4).getValues();}
  
  for (i in classroomTimeCodes) {
    var startChanged = false;
    var endChanged = false;
    if (classroomTimeCodes[i][0] != classroomTimeCodes[i][2]) {
      startChanged = true;
    }
    if (classroomTimeCodes[i][1] != classroomTimeCodes[i][3]) {
      endChanged = true;
    }
    if (startChanged==true || endChanged==true) { // if an email needs to be sent (about something on the classrooms sheet)
      var type;
      if (startChanged==true && endChanged==false) {
        type = 'Start Time Changed';
      }
      else {
        if (startChanged==false && endChanged==true) {
          type = 'End Time Changed';
        }
        else {
          if (startChanged==true && endChanged==true) {
            type = 'Rescheduled';
          }
        }
      }
      var eventPull = sheetClassrooms.getRange(parseInt(i)+5,1,1,14).getValues();
      var isAllDay = false;
      if (eventPull[0][4]==='All Day') {isAllDay=true;}
      var eventData = {
        name:eventPull[0][0],
        requested:eventPull[0][1],
        assigned:eventPull[0][2],
        url:eventPull[0][13],
        id:eventPull[0][10],
        type:type,
        to:'registrar',
        dates:eventPull[0][3],
        isAllDay:isAllDay,
        oldStart:classroomTimeCodes[i][2],
        oldEnd:classroomTimeCodes[i][3],
        newStart:classroomTimeCodes[i][0],
        newEnd:classroomTimeCodes[i][1],
      };
      sendEmail(eventData);
      sheetClassrooms.getRange(parseInt(i)+5,9,1,2).setValues([[classroomTimeCodes[i][0], classroomTimeCodes[i][1]]]);
      
    }
  } // done with emailing any changes from classrooms sheet
  
  for (i in fifteenTimeCodes) {
    var startChanged = false;
    var endChanged = false;
    if (fifteenTimeCodes[i][0] != fifteenTimeCodes[i][2]) {
      startChanged = true;
    }
    if (fifteenTimeCodes[i][1] != fifteenTimeCodes[i][3]) {
      endChanged = true;
    }
    if (startChanged==true || endChanged==true) { // if an email needs to be sent (about something on the 15th Floor sheet)
      var type;
      if (startChanged==true && endChanged==false) {
        type = 'Start Time Changed';
      }
      else {
        if (startChanged==false && endChanged==true) {
          type = 'End Time Changed';
        }
        else {
          if (startChanged==true && endChanged==true) {
            type = 'Rescheduled';
          }
        }
      }
      var eventPull = sheet15.getRange(parseInt(i)+5,1,1,14).getValues();
      var isAllDay = false;
      if (eventPull[0][4]==='All Day') {isAllDay=true;}
      var eventData = {
        name:eventPull[0][0],
        requested:eventPull[0][1],
        assigned:eventPull[0][2],
        url:eventPull[0][13],
        id:eventPull[0][10],
        type:type,
        to:'facilities',
        dates:eventPull[0][3],
        isAllDay:isAllDay,
        oldStart:fifteenTimeCodes[i][2],
        oldEnd:fifteenTimeCodes[i][3],
        newStart:fifteenTimeCodes[i][0],
        newEnd:fifteenTimeCodes[i][1]
      };
      sendEmail(eventData);
      sheet15.getRange(parseInt(i)+5,9,1,2).setValues([[fifteenTimeCodes[i][0], fifteenTimeCodes[i][1]]]);
    }
  }
  
  for (i in assignedTimeCodes) {
    var startChanged = false;
    var endChanged = false;
    if (assignedTimeCodes[i][0] != assignedTimeCodes[i][2]) {startChanged = true;}
    if (assignedTimeCodes[i][1] != assignedTimeCodes[i][3]) {endChanged = true;}
    if (startChanged==true || endChanged==true) { // if an email needs to be sent (about something on the 15th Floor sheet)
      var type;
      if (startChanged==true && endChanged==false) {
        type = 'Start Time Changed';
      }
      else {
        if (startChanged==false && endChanged==true) {
          type = 'End Time Changed';
        }
        else {
          if (startChanged==true && endChanged==true) {
            type = 'Rescheduled';
          }
        }
      }
      var eventPull = sheetAssigned.getRange(parseInt(i)+5,1,1,14).getValues();
      var isClassroomRequest = false;
      var locationString = ''+eventPull[0][1];
      var to = 'N/A'
      if ( (locationString.search("Classroom") != -1) || (locationString.search("Auditorium") != -1) || (locationString.search("Video Conference") != -1) || (locationString.search("Computer Lab") != -1) || (locationString.search("Other Space") != -1) || (locationString.search("ECNU Geography Building") != -1) ) { // if it was a classroom request
        to = 'registrar'; }
      else if ( (locationString.search("15") != -1) || (locationString.search("Art Gallery") != -1) || (locationString.search("Cafe") != -1) ) { to = 'facilities';}
      if (to != 'N/A') {
        var isAllDay = false;
        if (eventPull[0][4]==='All Day') {isAllDay=true;}
        var eventData = {
          name:eventPull[0][0],
          requested:eventPull[0][1],
          assigned:eventPull[0][2],
          url:eventPull[0][13],
          id:eventPull[0][10],
          type:type,
          to:to,
          dates:eventPull[0][3],
          isAllDay:isAllDay,
          oldStart:assignedTimeCodes[i][2],
          oldEnd:assignedTimeCodes[i][3],
          newStart:assignedTimeCodes[i][0],
          newEnd:assignedTimeCodes[i][1],
        };
        sendEmail(eventData);
        sheetAssigned.getRange(parseInt(i)+parseInt(5),9,1,2).setValues([[assignedTimeCodes[i][0], assignedTimeCodes[i][1]]]);
      }
    }
  }
 
}