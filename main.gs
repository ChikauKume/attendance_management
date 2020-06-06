// post message on slack.
function postSlack(text) {
  let postUrl = 'https://hooks.slack.com/XXXXXXXXXXX'; // insert URL of incoming token
  let jsonData =
      {
        "text" : text
      };
  let payload = JSON.stringify(jsonData);
  
  let options =
      {
        "method" : "post",
        "contentType" : "application/json",
        "payload" : '{"text":"' + text + '"}'
      };
  UrlFetchApp.fetch(postUrl, options);
}

// get message from slack
function doPost(e) {
  let token = 'XXXXXXXXXX'; // insert outgoing token
  if (token == e.parameter.token){
    let datetime = new Date();
    let date = (('0' + (datetime.getMonth() + 1)).slice(-2) + '/' + ('0' + datetime.getDate()).slice(-2))
    let time = (('0' + datetime.getHours()).slice(-2) + ':' + ('0' + datetime.getMinutes()).slice(-2));
    let user_name = e.parameter.user_name;
    let trigger_word = e.parameter.trigger_word;
    let text = check_input(e.parameter.text);
    let dayOfWeek = datetime.getDay();
    let day = datetime.getDate();
    let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('sheet');
    let key = show_message(text);

    if(user_name == "slackbot"){
      exit();
    }

    if(key == 0){
      record(sheet,date,time,user_name,trigger_word,text,dayOfWeek);
    }
    else{
      postSlack("please choose selection as follow");
      postSlack('・clock-in: enter you start working');
      postSlack('・clock-out: enter you finish working');
      postSlack('・overtime-in: enter you start overtime working');
      postSlack('・overtime-out: enter you finish overtime working');
      postSlack('・confirmed:<anything>: you can insert a remark');
    }
    
  }
  else {
    postSlack("token is invalid");
    return false;
  }
}

// show message on slack after input 
function show_message(text){
  switch(text){
    case 'clock-in':
      postSlack("clock-in data inserted");
      return 0;
      break;
    case 'clock-out':
      postSlack("clock-out data inserted");
      return 0;
      break;
    case 'overtime-in':
      postSlack("overtime-in data inserted");
      return 0;
      break;
    case 'overtime-out':
      postSlack("overtime-out data inserted");
      return 0;
      break;
    default:
      if (text.indexOf('confirmed:') != -1) {
        postSlack("confirmed data inserted");
        return 0;
      }
      else{
        return 1;
      }
      break;
  }
}

// write down attendance data on sheet
function record(sheet,date,time,user_name,trigger_word,text,key){
  let day = check_day_of_week(key);
  let array = [[day,date,time]];

  switch(text){
    case "clock-in":
      clock_in(sheet,array,time,date);
      break;
    case "clock-out":
      clock_out(sheet,date,time);
      break;
    case "overtime-in":
      overtime_in(sheet,array,time,date);
      break;
    case "overtime-out":
      overtime_out(sheet,array,time,date);
      break;
    default:
      if (text.indexOf('confirmed:') != -1) {
        add_confirmed(sheet,date,text);
      }
      else{
        postSlack("insert data is invalid");
      }
      break;
  }
  return;
}

// the flow of clock-in
function clock_in(sheet,array,time,date) {
  const column = sheet.getRange('C8:C22').getValues(); 
  const lastrow = column.filter(String).length+8;  
  let val = sheet.getRange(lastrow-1,2).getValue();
  if (val == date){
    sheet.getRange(lastrow-1,3).setValue(time);
  }
  else {
    sheet.getRange(lastrow,1,1,3).setValues(array);
  } 
  return;
}

// the flow of clock-out
function clock_out(sheet,date,time) {
  const column = sheet.getRange('D8:D22').getValues(); 
  const lastrow = column.filter(String).length+8;  
  let val = sheet.getRange(lastrow-1,2).getValue();
  if (val == date){
    sheet.getRange(lastrow-1,4).setValue(time);
  }
  else {
    sheet.getRange(lastrow,4).setValue(time);
  } 
  return;
}

// the flow of overtime-in
function overtime_in(sheet,array,time,date) {
  const column = sheet.getRange('C8:C22').getValues(); 
  const lastrow = column.filter(String).length+8;  
  let val = sheet.getRange(lastrow-1,2).getValue();
  if (val == date){
    sheet.getRange(lastrow-1,6).setValue(time);
  }
  return;
}

// the flow of overtime-out
function overtime_out(sheet,array,time,date){
  const column = sheet.getRange('C8:C22').getValues(); 
  const lastrow = column.filter(String).length+8;  
  let val = sheet.getRange(lastrow-1,2).getValue();
  if (val == date){
    sheet.getRange(lastrow-1,7).setValue(time);
  }
  return;
}

// the flow of adding remarks
function add_confirmed(sheet,date,text) {
  var txt = text.slice(10);
  const column = sheet.getRange('C8:C22').getValues(); 
  const lastrow = column.filter(String).length+8;  
  let val = sheet.getRange(lastrow-1,2).getValue();
  if (val == date){
    sheet.getRange(lastrow-1,9).setValue(txt);
  }
  return;
}


// identify the day of week
function check_day_of_week(key) {
  switch (key) {
    case 0:
      day = "Sun";
      break;
    case 1:
      day = "Mon";
      break;
    case 2:
      day = "Tue";
      break;
    case 3:
      day = "Wed";
      break;
    case 4:
      day = "Thu";
      break;
    case 5:
       day = "Fri";
      break;
    case 6:
       day = "Sat";
  }
  return day;
}

// check input data from slack
function check_input(text){
    switch (text) {
      case "Clock-In":
        text = "clock-in";
        break;
      case "Clock-Out":
        text = "clock-out";
        break;
      case "Overtime-In":
        text="overtime-in";
        break;
      case "Overtime-Out":
        text="overtime-out"
        break;
    }
  return text;
}