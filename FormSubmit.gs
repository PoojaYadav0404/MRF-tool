function onForm(e){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var responseSheet = ss.getSheetByName("Form responses 1");
  var tempSheet = ss.getSheetByName("Mail temp");
  var range = e.range
  var row = range.getRow();
  var NewReqNo = tempSheet.getRange("A2").getValue() ;
  var ReqCode = "R-" + NewReqNo ;
  responseSheet.getRange(row, 1).setValue(ReqCode);
  var data = range.getValues();

  //JD according to Department and post
  
  var to = tempSheet.getRange(2,4).getValue();
  var sub = tempSheet.getRange(3,4).getValue() + "// " + responseSheet.getRange(row, 1).getValue() ;
  var cc = data[0][3]
  
  var row = range.getRow();
  var body = "Dear sir," +  "<p>New Manpower requisition application submitted for approval. Details as follow:" + ".<p>Request Id: " + responseSheet.getRange(row, 1).getValue() + "<br>Requested by: " +  data[0][1] + "<br>Company Name: " + data[0][4]  + "<br>Department: " + data[0][5]  + "<br>Job Title: " + data[0][6] + "<br>Attached JD: " + responseSheet.getRange(row, 17).getValue() + "<br>Response Sheet Link: " + "https://docs.google.com/spreadsheets/d/1AKiAd-GQhdPiSgxlAJc3jRz87fQgeF73oL97-H79AlM/edit?resourcekey#gid=2024975217" + "<p>Thanks"

  GmailApp.sendEmail(to, sub,"", {htmlBody: body, cc:cc});
  tempSheet.getRange("A2").setValue(NewReqNo + 1);
  
  //if (Session.getActiveUser() == "dme@bajatoparts.com" )
  
}

function onOpen(){
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Form Ranger Dropdowns").hideSheet();
}

function onEdit(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var activeSheet = ss.getActiveSheet();
  var row = activeSheet.getActiveCell().getRow();
  var col = activeSheet.getActiveCell().getColumn();

  if (activeSheet.getName()==="Form responses 1" && row>1 && col==17){
    activeSheet.getRange("Q2:Q").clearContent();
    activeSheet.getRange("Q1").setFormula(`=ArrayFormula({"JD Link"; if(H2:H="","", VLOOKUP(H2:H,'Form Ranger Dropdowns'!J:K,2,false))})`)
    var ui = SpreadsheetApp.getUi();
    ui.alert("Warning!!", "Don't write anything in this column. This column will fetch details automatically. If not then please go to 'Form Ranger Dropdowns' sheet and enter JD link in J column.", ui.ButtonSet.OK);
  }






}












