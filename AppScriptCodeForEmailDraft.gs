function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Draft Email')
    .addItem('AVT Update Email', 'draftGMail')
    .addToUi();
}
function myFunction() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Sheet2");
  var status = sheet.getRange("F5").getValue();
  if (status == "Compliant") {
    sheet.hideRows(6);
  }
  else {
    sheet.showRows(6);
    var cDataCell = sheet.getRange("F6");
    cDataCell.setValue("Complaint");
    cDataCell.setBackground('#92d050')
  }
}

function copyEmailDetails() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var sheet3 = ss.getSheetByName("Sheet3");
  var appname = sheet3.getRange("B3").getValue();

  var sheet2 = ss.getSheetByName("Sheet2");
  var dstrng_to = sheet2.getRange("B9");
  var dstrng_cc = sheet2.getRange("B10");

  var srcrng_to_rl = sheet3.getRange("H7");
  var srcrng_cc_rl = sheet3.getRange("H8");
  var srcrng_to_ob = sheet3.getRange("H10");
  var srcrng_cc_ob = sheet3.getRange("H11");
  var srcrng_to_cm = sheet3.getRange("H13");
  var srcrng_cc_cm = sheet3.getRange("H14");
  var srcrng_to_sc = sheet3.getRange("H16");
  var srcrng_cc_sc = sheet3.getRange("H17");
  var srcrng_to_pp = sheet3.getRange("H19");
  var srcrng_cc_pp = sheet3.getRange("H20");
  var srcrng_to_dp = sheet3.getRange("H22");
  var srcrng_cc_dp = sheet3.getRange("H23");

  switch (appname) {
    case "Risklab":
      srcrng_to_rl.copyTo(dstrng_to);
      srcrng_cc_rl.copyTo(dstrng_cc);
      break;
    case "Online Boarding":
      srcrng_to_ob.copyTo(dstrng_to);
      srcrng_cc_ob.copyTo(dstrng_cc);
      break;
    case "Client Manager":
      srcrng_to_cm.copyTo(dstrng_to);
      srcrng_cc_cm.copyTo(dstrng_cc);
      break;
    case "SalesComp":
      srcrng_to_sc.copyTo(dstrng_to);
      srcrng_cc_sc.copyTo(dstrng_cc);
      break;
    case "Pricing Portal":
      srcrng_to_pp.copyTo(dstrng_to);
      srcrng_cc_pp.copyTo(dstrng_cc);
      break;
    case "Devportal":
      srcrng_to_dp.copyTo(dstrng_to);
      srcrng_cc_dp.copyTo(dstrng_cc);
      break;
  }
  //draftGMail();
}

function draftGMail() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var emailsheet = ss.getSheetByName("Sheet2");
  var sent_to = emailsheet.getRange("B9").getValue();
  var sent_cc = emailsheet.getRange("B10").getValue();
  var app_name = emailsheet.getRange("C5").getValue();
  var app_env = emailsheet.getRange("D5").getValue();
  var disclaimer = emailsheet.getRange("B3:G3").getValue().replace(/\n/g, '<br>');
  var ticketNo = emailsheet.getRange("B5").getValue();
  var appName = emailsheet.getRange("C5").getValue();
  var env = emailsheet.getRange("D5").getValue();
  var server = emailsheet.getRange("E5").getValue();
  var ncserver = emailsheet.getRange("E6").getValue();
  var cStatus = emailsheet.getRange("F5").getValue();
  var cStatusComplaint = emailsheet.getRange("F6").getValue();
  var cStatusColor = emailsheet.getRange("F5").getBackground();
  var cStatusColorComplaint = emailsheet.getRange("F6").getBackground();
  
  var appStatus = emailsheet.getRange("G5").getValue();
  var appStatusColor = emailsheet.getRange("G5").getBackground();
  var todaydate = Utilities.formatDate(new Date(), "GMT+1", "dd/MM/yyyy")
  var body = '<p style="font-family: Arial"> Hi, Team </p>';
  body += '<div> Please find the below Patching Status: </div>';
  body += `<figure class="table" style="width:91.85%;">`;
  body += `<table class="ck-table-resized" style="border-color:hsl(0, 0%, 0%);border-style:solid;">`;
  body += `<colgroup>`;
  body += `<col style="width:18.12%;">`;
  body += `<col style="width:9.7%;">`;
  body += `<col style="width:12.16%;">`;
  body += `<col style="width:25.01%;">`;
  body += `<col style="width:17.13%;">`;
  body += `<col style="width:17.88%;">`;
  body += `</colgroup>`;
  body += `<thead>`;
  body += `<tr>`;
  body += `<th style="border-color:hsl(0, 0%, 0%);border-style:solid;text-align: left;vertical-align: baseline;" colspan="6">` + disclaimer + `</th>`;
  body += `</tr>`;
  body += `<tr>`;
  body += `<th style="background-color:hsl(210, 75%, 60%);border-color:hsl(0, 0%, 0%);border-style:solid;text-align:center;">`;
  body += `<p style="text-align:center;"><span style="color:#263238;font-size:12px;"><strong>MASTER CHANGE TICKET</strong></span></p>`;
  body += `</th>`;
  body += `<th style="background-color:hsl(210, 75%, 60%);border-color:hsl(0, 0%, 0%);border-style:solid;text-align:center;">`;
  body += `<p style="text-align:center;"><span style="color:#263238;font-size:12px;"><strong>APPLICATION</strong></span></p>`;
  body += `</th>`;
  body += `<th style="background-color:hsl(210, 75%, 60%);border-color:hsl(0, 0%, 0%);border-style:solid;text-align:center;">`;
  body += `<p style="text-align:center;"><span style="color:#263238;font-size:12px;"><strong>ENVIRONMENT</strong></span></p>`;
  body += `</th>`;
  body += `<th style="background-color:hsl(210, 75%, 60%);border-color:hsl(0, 0%, 0%);border-style:solid;text-align:center;">`;
  body += `<p style="text-align:center;"><span style="color:#263238;font-size:12px;"><strong>SERVERS</strong></span></p>`;
  body += `</th>`;
  body += `<th style="background-color:hsl(210, 75%, 60%);border-color:hsl(0, 0%, 0%);border-style:solid;text-align:center;">`;
  body += `<p style="text-align:center;"><span style="color:#263238;font-size:12px;"><strong>COMPLIANCE STATUS</strong></span></p>`;
  body += `</th>`;
  body += `<th style="background-color:hsl(210, 75%, 60%);border-color:hsl(0, 0%, 0%);border-style:solid;text-align:center;">`;
  body += `<p style="text-align:center;"><span style="color:#263238;font-size:12px;"><strong>APPLICATION STATUS</strong></span></p>`;
  body += `</th>`;
  body += `</tr>`;
  body += `</thead>`;
  body += `<tbody>`;


  if (cStatus == 'Compliant') {
    body += `<tr>`;
    body += `<td style="border-color:hsl(0, 0%, 0%);border-style:solid;text-align: center;">` + ticketNo + `</td>`;
    body += `<td style="border-color:hsl(0, 0%, 0%);border-style:solid;text-align: center;">` + appName + `</td>`;
    body += `<td style="border-color:hsl(0, 0%, 0%);border-style:solid;text-align: center;">` + env + `</td>`;
    body += `<td style="border-color:hsl(0, 0%, 0%);border-style:solid;text-align: center;">` + server + `</td>`;
    body += `<td style="border-color:hsl(0, 0%, 0%);border-style:solid;text-align: center;background-color:` + cStatusColor + `;">` + cStatus + `</td>`;
    body += `<td style="border-color:hsl(0, 0%, 0%);border-style:solid;text-align: center;background-color:` + appStatusColor + `;">` + appStatus + `</td>`;
    body += `</tr>`;
  }
  else {
    body += `<tr>`;
    body += `<td style="border-color:hsl(0, 0%, 0%);border-style:solid;text-align: center;" rowspan = 2>` + ticketNo + `</td>`;
    body += `<td style="border-color:hsl(0, 0%, 0%);border-style:solid;text-align: center;"rowspan = 2>` + appName + `</td>`;
    body += `<td style="border-color:hsl(0, 0%, 0%);border-style:solid;text-align: center;"rowspan = 2>` + env + `</td>`;
    body += `<td style="border-color:hsl(0, 0%, 0%);border-style:solid;text-align: center;">` + server + `</td>`;
    body += `<td style="border-color:hsl(0, 0%, 0%);border-style:solid;text-align: center;background-color:` + cStatusColor + `;">` + cStatus + `</td>`;
    body += `<td rowspan="2" style="border-color:hsl(0, 0%, 0%);border-style:solid;text-align: center;background-color:` + appStatusColor + `">` + appStatus + `</td>`;
    body += `</tr>`;
    body += `<tr>`;
    body += `<td style="border-color:hsl(0, 0%, 0%);border-style:solid;text-align: center;"></td>`;
    body += `<td style="border-color:hsl(0, 0%, 0%);border-style:solid;text-align: center;background-color:` + cStatusColorComplaint + `;">` + cStatusComplaint + `</td>`;
    body += `</tr>`;
  }

  body += `</tbody>`;
  body += `</table>`;
  body += `</figure>`;
  body += '<p style="font-family: Arial"><b><u> Validation Status:</u></b> </p>';
  if (appStatus == 'PASS') {
    body += '<div>No Issues were found and the application is working fine as expected.</div>';
  }
  else {
    body += '<div><b><u>Steps To Reproduce,</u></b></div>';
    body += '<div> </div>'
    body += '<div><b><u>ScreenShots:,</u></b></div>';
  }


  GmailApp.createDraft(sent_to, "AVT Update:" + todaydate + "|" + app_name + "|" + app_env, body, { cc: sent_cc, htmlBody: body });

  //MailApp.sendEmail({
  //to: sent_to,
  //subject: "AVT Update:" + todaydate + "|" + app_name + "|" + app_env,
  //htmlBody: body
  //});

}
