var ResponseSheet = "Form responses 1"

function firstStepApproval(e) {
  var SS = SpreadsheetApp.getActiveSpreadsheet();
  var Sheet = SS.getSheetByName("Form responses 1");
  var MailSheet = SS.getSheetByName("Mail temp");
  var ActiveCell = e.range;
  var Row = ActiveCell.getRow();
  var Col = ActiveCell.getColumn();
  var SSName = ActiveCell.getSheet().getName();
  var MailId = Session.getActiveUser().getEmail();

  
  if (MailId == "dme@bajatoparts.com" || MailId == "director@bajatoparts.com" || MailId == "rajatbajaj@bajato.com"){
    if (SSName == ResponseSheet && Col == 18 && Row > 1 && Sheet.getRange(Row,18).getValue()!= ""){
      var Ui = SpreadsheetApp.getUi();
      var ButtonPressed = Ui.alert("Approval Status!", "Are you sure you want to continue?", Ui.ButtonSet.YES_NO_CANCEL);

      if (Sheet.getActiveCell().getValue() == "Approved"  && ButtonPressed == Ui.Button.YES){
        
        var ReqId = Sheet.getRange(Row, 1).getValue();
        var ReqName = Sheet.getRange(Row, 3).getValue();
        var Company = Sheet.getRange(Row, 6).getValue();
        var Dept = Sheet.getRange(Row, 7).getValue();
        var JobTitle = Sheet.getRange(Row, 8).getValue();
        var Remark = Sheet.getRange(Row, 19).getValue();
        var JD = Sheet.getRange(Row, 17).getValue();
        var ReqStatus = Sheet.getRange(Row, 18).getValue();
        
        var to = MailSheet.getRange(2,7).getValue();
        var sub = "New Vacancy//" + JobTitle + "//" + ReqId ;
        var cc = Sheet.getRange(Row, 5).getValue() + ", " + MailSheet.getRange(2,4).getValue();

        var body = "Dear HR," +  "<p>New Manpower requisition application request is #" + "<b>" + ReqStatus + "</b>" + ". Do the necessary changes in the attached JD if needed and share the details of Request Id #" + "<b>" + ReqId + "</b>" + " with the consultant." + "<p>Few more details for your reference:" + "<br>Requested by: " +  ReqName + "<br>Company Name: " + Company  + "<br>Department: " + Dept  + "<br>Job Title: " + JobTitle + "<br>Remark: " + Remark + "<br>JD Link: " + JD + "<br>Response Sheet Link: " + "https://docs.google.com/spreadsheets/d/1AKiAd-GQhdPiSgxlAJc3jRz87fQgeF73oL97-H79AlM/edit?resourcekey#gid=2024975217" + "<p>Thanks,"

        GmailApp.sendEmail(to, sub,"", {htmlBody: body, cc:cc});
        Sheet.getRange(Row, 20).setValue(new Date());

      }else if ((Sheet.getActiveCell().getValue() == "Not approved" ||Sheet.getActiveCell().getValue() == "On Hold") && ButtonPressed == Ui.Button.YES){
        
        var ReqId = Sheet.getRange(Row, 1).getValue();
        var ReqName = Sheet.getRange(Row, 3).getValue();
        var Company = Sheet.getRange(Row, 6).getValue();
        var Dept = Sheet.getRange(Row, 7).getValue();
        var JobTitle = Sheet.getRange(Row, 8).getValue();
        var Remark = Sheet.getRange(Row, 19).getValue();
        var JD = Sheet.getRange(Row, 17).getValue();
        var ReqStatus = Sheet.getRange(Row, 18).getValue();

        var to = Sheet.getRange(Row, 5).getValue();
        var cc = MailSheet.getRange(2,7).getValue() + ", " + MailSheet.getRange(2,4).getValue();;
        var sub = "MRF request status//" + ReqId;

        var body = "Dear Team," +  "<p>New Manpower requisition application request is #" + "<b>" + ReqStatus + "</b>" + ". Details as follow:" + "<p>Request Id: " + ReqId + "<br>Requested by: " +  ReqName + "<br>Company Name: " + Company  + "<br>Department: " + Dept  + "<br>Job Title: " + JobTitle + "<br>Remark: " + Remark  + "<br>Response Sheet Link: " + "https://docs.google.com/spreadsheets/d/1AKiAd-GQhdPiSgxlAJc3jRz87fQgeF73oL97-H79AlM/edit?resourcekey#gid=2024975217" + "<p>Thanks,"

        GmailApp.sendEmail(to, sub,"", {htmlBody: body, cc:cc});
        Sheet.getRange(Row, 20).setValue(new Date());

      }else if (ButtonPressed == Ui.Button.NO){
        SS.getActiveSheet().getActiveCell().clearContent();
      }
    }
  } else {var Ui = SpreadsheetApp.getUi();
            Ui.alert("Alert!!", "Invalid User Id.", Ui.ButtonSet.OK)
  }
}


function secondStepApproval(e) {
  var SS = SpreadsheetApp.getActiveSpreadsheet();
  var Sheet = SS.getSheetByName("Form responses 1");
  var MailSheet = SS.getSheetByName("Mail temp"); 
  var PdfFolder = DriveApp.getFolderById("11KwklgVXzDXgAoepp4yO2z_CCfNKwryEIZ3Gu6sNfA7lh08JIV3mBnC2ORcYM268754sD3hd")
  var ActiveCell = e.range;
  var Row = ActiveCell.getRow();
  var Col = ActiveCell.getColumn();
  var AppStatus = Sheet.getRange(Row, 18).getValue();
  var SSName = ActiveCell.getSheet().getName();
  var MailId = Session.getActiveUser().getEmail();

  if (MailId == "dme@bajatoparts.com"){
    if (SSName == ResponseSheet && Col == 21 && Row > 1 && AppStatus == "Approved" && Sheet.getRange(Row, 17).getValue()!= "" && Sheet.getRange(Row, 21).getValue()!= "" ){
      var Ui = SpreadsheetApp.getUi();
      var ButtonPressed = Ui.alert("Sharing details with Consultant", "Are you sure you want to continue?", Ui.ButtonSet.YES_NO_CANCEL);

      if (Sheet.getActiveCell().getValue() == "Share" && ButtonPressed == Ui.Button.YES){
        
        var Company = Sheet.getRange(Row, 6).getValue();
        var Dept = Sheet.getRange(Row, 7).getValue();
        var JobTitle = Sheet.getRange(Row, 8).getValue();
        var ExpDFJ = Sheet.getRange(Row, 13).getValue();
        var Remark = Sheet.getRange(Row, 22).getValue();
        var JD = Sheet.getRange(Row, 17).getValue();
        var ActivateJdFile = DocumentApp.openByUrl(JD);
        var BlobFile = ActivateJdFile.getAs(MimeType.PDF);
        var PdfFile = PdfFolder.createFile(BlobFile).setName(Utilities.formatDate(new Date(), "GMT+05:30", "dd/MM/yyyy"));

        var to = MailSheet.getRange(2,10).getValue();
        var sub = "Bajato Parts and System Pvt. Ltd.//" + JobTitle; 
        var cc = MailSheet.getRange(2,4).getValue() + "," + MailSheet.getRange(2,7).getValue();

        var body = "Dear,<p>Hope You are doing well ! <p>We have new vacancy in our organisation. Please find the attached JD and details below for your reference." + "<p>Company Name: " + Company  + "<br>Department: " + Dept  + "<br>Job Title: " + JobTitle + "<br>Expected date of joining: " + ExpDFJ  + "<br>Remark: " + Remark  + "<p>If you require any further information, feel free to contact us." + "<p>Best Regards," + "<br>Bajato Parts and System Pvt. Ltd."

        GmailApp.sendEmail(to, sub, "", {
          htmlBody : body,
          cc: (cc),
          attachments: [PdfFile],
          name: 'HR BajatoParts'
      });
        Sheet.getRange(Row, 23).setValue("Details shared " + Utilities.formatDate(new Date(), "GMT+05:30", "dd/MM/yyyy"));
        PdfFile.setTrashed(true);

      }else if ((Sheet.getActiveCell().getValue() == "Share" || Sheet.getActiveCell().getValue() == "on Hold" ) && ButtonPressed == Ui.Button.NO){
        SS.getActiveSheet().getActiveCell().clearContent();
      }else if (Sheet.getActiveCell().getValue() == "on Hold" && ButtonPressed == Ui.Button.YES){
        Sheet.getRange(Row, 23).setValue(Utilities.formatDate(new Date(), "GMT+05:30", "dd/MM/yyyy"));
      }
    } else if(SSName == ResponseSheet && Col == 18 && Row > 1 && Sheet.getRange(Row, 17).getValue()== ""){
      var Ui = SpreadsheetApp.getUi();
      Ui.alert("Alert!!", "Please enter JD link in Column Q.", Ui.ButtonSet.OK)

    }
  }else {var Ui = SpreadsheetApp.getUi();
            Ui.alert("Alert!!", "Invalid User Id.", Ui.ButtonSet.OK)
  }
}


function finalStepStatus(e) {
  var SS = SpreadsheetApp.getActiveSpreadsheet();
  var Sheet = SS.getSheetByName("Form responses 1");
  var MailSheet = SS.getSheetByName("Mail temp");
  //var ActiveCell = SS.getActiveCell();
  var ActiveCell = e.range;
  var Row = ActiveCell.getRow();
  var Col = ActiveCell.getColumn();
  var SSName = ActiveCell.getSheet().getName();
  var MailId = Session.getActiveUser().getEmail();
  var AppStatus = Sheet.getRange(Row, 18).getValue();

  
  if (MailId == "dme@bajatoparts.com" ){
    if (SSName == ResponseSheet && Col == 24 && Row > 1 && AppStatus == "Approved" && Sheet.getRange(Row, 17).getValue()!= "" && Sheet.getRange(Row, 24).getValue()!= ""){
      var Ui = SpreadsheetApp.getUi();
      var ButtonPressed = Ui.alert("Final Status Update!", "Are you sure you want to continue?", Ui.ButtonSet.YES_NO_CANCEL);
      

      if ((Sheet.getActiveCell().getValue() == "Cancel" ||Sheet.getActiveCell().getValue() == "Closed") && ButtonPressed == Ui.Button.YES){
        
        var ReqId = Sheet.getRange(Row, 1).getValue();
        var ReqName = Sheet.getRange(Row, 3).getValue();
        var Company = Sheet.getRange(Row, 6).getValue();
        var Dept = Sheet.getRange(Row, 7).getValue();
        var JobTitle = Sheet.getRange(Row, 8).getValue();
        var Remark = Sheet.getRange(Row, 25).getValue();
        var ReqStatus = Sheet.getRange(Row, 24).getValue();

        var to = Sheet.getRange(Row, 5).getValue() + ", " + MailSheet.getRange(2,4).getValue();
        var cc = MailSheet.getRange(2,7).getValue() ;
        var sub = "MRF//" + ReqId + "//Final Status";

        var body = "Dear Team," +  "<p>Manpower requisition application request #" + "<b>" + ReqId + "</b>" + " is #" + "<b>" + ReqStatus + "</b>" + ". Details as follow:" + "<p>Requested by: " +  ReqName + "<br>Company Name: " + Company  + "<br>Department: " + Dept  + "<br>Job Title: " + JobTitle + "<br>Remark: " + Remark +  "<br>Response Sheet Link: " + "https://docs.google.com/spreadsheets/d/1AKiAd-GQhdPiSgxlAJc3jRz87fQgeF73oL97-H79AlM/edit?resourcekey#gid=2024975217" + "<p>For more updates and details connect with HR Team." + "<p>Best Regards," + "<br>HR Team"

        GmailApp.sendEmail(to, sub,"", {htmlBody: body, cc:cc});
        Sheet.getRange(Row, 26).setValue(new Date());

      }else if (ButtonPressed == Ui.Button.NO){
        SS.getActiveSheet().getActiveCell().clearContent();
      }
    } else if(SSName == ResponseSheet && Col == 18 && Row > 1 && Sheet.getRange(Row, 17).getValue()== ""){
      var Ui = SpreadsheetApp.getUi();
      Ui.alert("Alert!!", "Please enter JD link in Column Q.", Ui.ButtonSet.OK)
    }
  } else {var Ui = SpreadsheetApp.getUi();
            Ui.alert("Alert!!", "Invalid User Id.", Ui.ButtonSet.OK)
  }
}



