// THIS IS THE CORRECT ONE

function getOrgNickname(formalName) { // returns nickname of org
  const orgDict = {
    '':'',  
    'I am not representing any UPD CoE Organization.': '',
    'I have no other affiliation with any other UPD CoE Organization.': '',
    'Institute of Electrical and Electronics Engineering UP Diliman Student Branch (IEEE UPD SB)': 'IEEE UPD SB',
    'Institute of Integrated Electrical Engineers - Council of Student Chapters UP Diliman (IIEE-CSC-UPD)': 'IIEE-CSC-UPD',
    'Philippine Institute of Civil Engineers - American Concrete Institute Philippines - University of the Philippines Diliman Student Chapter (PICE-ACIP-UPDSC)': 'PICE-ACIP-UPDSC',
    'Philippine Society of Mechanical Engineers - University of the Philippines Student Unit (PSME-UPSU)': 'PSME-UPSU',
    'Society of Manufacturing Engineers - UP Diliman (SME-UPD)': 'SME-UPD',
    'University of the Philippines Chemical Engineering Society, Inc. (UP KEM)': 'UP KEM',
    'University of the Philippines Circle of Industrial Engineering Majors (UP CIEM)': 'UP CIEM',
    'UP Academic League of Chemical Engineering Students, Inc. (UP ALCHEMES)': 'UP ALCHEMES',
    'UP Aggregates Incorporated': 'UP Aggregates Incorporated',
    'UP Association for Computing Machinery (ACM) - Diliman Student Chapter Inc.': 'ACM',
    'UP Association of of Civil Engineering Students (UP ACES)': 'UP ACES',
    'UP Association of Computer Science Majors (UP CURSOR)': 'UP CURSOR',
    'UP Circle of Engineering Students (UP CREST)': 'UP CREST',
    'UP Circuit': 'UP Circuit',
    'UP Engineering Radio Guild (UP ERG)': 'UP ERG',
    'UP Engineering Society (UP Engg Soc)': 'UP Engg Soc',
    'UP Gears and Pinions (UP GPs)': 'UP GPs',
    'UP Geodetic Engineering Club (GEC)': 'GEC',
    'UP Industrial Engineering Club (UP IE Club)': 'UP IE Club',
    'UP Materials Science Society (UP MSS)': 'UP MSS',
    'UP Mining, Metallurgical, and Materials Engineering Association, Inc. (UP 49ers)': 'UP 49ers',
    'UP Society of Computer Scientists (UP SoComSci)': 'UP SoComSci',
    'UP Society of Geodetic Engineering Majors (UP GEOP)': 'UP GEOP',
    'UP Center for Student Innovations (UP CSI)': 'UP CSI',
  }

  return orgDict[formalName];
}


function getData(e) {

  var namedResponses = e.namedValues;
  
  var studentData = {
    "surname": namedResponses["Surname"][0],
    "firstName": namedResponses["First Name"][0],
    "middleInitial": namedResponses["Middle Initial"][0],
    "nickname": namedResponses["Nickname"][0],
    "emailAddress": namedResponses["Email Address"][0],
    "studentNumber": namedResponses["Student Number"][0],
    "contactNumber": namedResponses["Contact Number (09XXXXXXXXX)"][0],
    "degreeProgram": namedResponses["Degree Program"][0],
    "yearLevel": namedResponses["Year Standing "][0],
    "isFromUPD": namedResponses["Are you from UP Diliman?"][0],
    "universityName": namedResponses["If you are not from UP Diliman, please select your College/University below."][0],
    "expectedGraduation": namedResponses["Expected Semester of Graduation "][0],
    "isCAPESMember": namedResponses["Are you a UP CAPES Member?"][0],
    "facebookProfileLink": namedResponses["Facebook Profile Link"][0],
    "instagramProfileLink": namedResponses["Instagram Profile Link"][0],
    "twitterProfileLink": namedResponses["X (Twitter) Profile Link"][0],
    "affiliatedOrganization1": getOrgNickname(namedResponses["What is your first choice for representing any UPD College of Engineering Organization this year?"][0]),
    "affiliatedOrganization2": getOrgNickname(namedResponses["If you're representing two or more organizations, what is your second choice for representing any UPD College of Engineering Organization this year?"][0]),
    "affiliatedOrganization3": getOrgNickname(namedResponses["What is your third choice for representing any UPD College of Engineering Organization this year?"][0]),
    "talentsNetworkSubscription": namedResponses["Do you wish to subscribe to the UP CAPES Talent Network?"][0],
    "capesUsername": "",
    "capesCardQRFileID": "",
  };

  if (studentData["isFromUPD"] == "Yes") {
    studentData['universityName'] = '';
  } else {
    studentData['universityName'] = studentData["universityName"];
  }

  const spreadsheet = SpreadsheetApp.openByUrl('https://docs.google.com/spreadsheets/d/18diQrWtpIDOHz0gOjMbRtSa0InRPpzhCysWJRcxl0xQ/edit?usp=sharing');
  const capesCardsSheet = spreadsheet.getSheetByName('Copy of CAPES Card 2425'); // Trial
  
  const capesCardsData = capesCardsSheet.getDataRange().getValues();
  const capesCardsSheetLastRow = capesCardsSheet.getRange("A:A").getValues().filter(String).length + 1;

  // const capesCardsSheetLastRow = capesCardsSheet.getLastRow()+1;

  // [PRE-PROCESSING] Erases whitespace characters in column names
  // Needed because ' Surname' is different to 'Surname' and will cause an error
  for (let i = 0; i < capesCardsData[0].length; i += 1) {
    capesCardsData[0][i] = String(capesCardsData[0][i]).trim();
  }

  // Gets column numbers of needed data in 'Copy of CAPES Cards 2425' sheet
  // NOTE: Key names must be the SAME as what's used in responsesSheetColumns. 'capesUsername' is the only exception.
  var capesCardsSheetColumns = {
    "surname": capesCardsData[0].indexOf("Surname"),
    "firstName": capesCardsData[0].indexOf("First Name"),
    "middleInitial": capesCardsData[0].indexOf("Middle Initial"),
    "nickname": capesCardsData[0].indexOf("Nickname"),
    "emailAddress": capesCardsData[0].indexOf("Email Address"),
    "studentNumber": capesCardsData[0].indexOf("Student Number"),
    "contactNumber": capesCardsData[0].indexOf("Contact Number"),
    "degreeProgram": capesCardsData[0].indexOf("Degree Program"),
    "yearLevel": capesCardsData[0].indexOf("Year Level"),
    "universityName": capesCardsData[0].indexOf("College/University"),
    "expectedGraduation": capesCardsData[0].indexOf("Expected Graduation"),
    "isCAPESMember": capesCardsData[0].indexOf("CAPES Member"),
    "facebookProfileLink": capesCardsData[0].indexOf("Facebook Profile Link"),
    "instagramProfileLink": capesCardsData[0].indexOf("Instagram Profile Link"),
    "twitterProfileLink": capesCardsData[0].indexOf("X (Twitter) Profile Link"),
    "affiliatedOrganization1": capesCardsData[0].indexOf("Organization 1"),
    "affiliatedOrganization2": capesCardsData[0].indexOf("Organization 2"),
    "affiliatedOrganization3": capesCardsData[0].indexOf("Organization 3"),
    "capesUsername": capesCardsData[0].indexOf("CAPES Username"),
    "updateStatus": capesCardsData[0].indexOf("Update Status")
  }

  // // CAPES Cards Username
  var cell = capesCardsSheet.getRange(capesCardsSheetLastRow, capesCardsSheetColumns["capesUsername"] + 1);
  var capesUsername = '';

  var firstName = String(studentData["firstName"]).trim();
  var firstNameChar = firstName.charAt(0).toLowerCase();

  var middleInitial = String(studentData["middleInitial"]).trim();
  var middleInitialChar = middleInitial.charAt(0).toLowerCase();

  var lastName = String(studentData["surname"]).trim().replace(/\s+/g, '');

  capesUsername = firstNameChar + middleInitialChar + lastName.toLowerCase();
  fullNameCheck = String(studentData["surname"]).trim().toLowerCase() + "," + firstName + "," + middleInitialChar;

  var studentNumberCheck = String(studentData["studentNumber"]).trim();
  var emailAddressCheck = String(studentData["emailAddress"]).trim().toLowerCase();


  // Checks if there are people w/ the same capesUsername
  usernamesList = capesCardsData.map(row => row[capesCardsSheetColumns["capesUsername"]]);
  fullNameList = capesCardsData.map(row => {
    var surname = row[capesCardsSheetColumns["surname"]] || "";
    var firstName = row[capesCardsSheetColumns["firstName"]] || "";
    var middleInitial = row[capesCardsSheetColumns["middleInitial"]] || "";
    return (surname + "," + firstName + "," + middleInitial.charAt(0)).toLowerCase().trim();
  });
  studentNumberList = capesCardsData.map(row => String(row[capesCardsSheetColumns["studentNumber"]]));
  emailAddressList = capesCardsData.map(row => row[capesCardsSheetColumns["emailAddress"]]);

  var numOfOccurencesOfUsername = 0;
  for(var i=0;i<usernamesList.length;i++){
      if(usernamesList[i].includes(capesUsername))
        numOfOccurencesOfUsername++;
  }

  var numOfOccurencesOfFullName = 0;
  for(var i=0;i<fullNameList.length;i++){
      if(fullNameList[i].includes(fullNameCheck))
        numOfOccurencesOfFullName++;
  }

  var numOfOccurencesOfStudentNumber = 0;
  for(var i=0;i<studentNumberList.length;i++){
      if(studentNumberList[i].includes(studentNumberCheck))
        numOfOccurencesOfStudentNumber++;
  }

  var numOfOccurencesOfEmailAddress = 0;
  for(var i=0;i<emailAddressList.length;i++){
      if(emailAddressList[i].includes(emailAddressCheck))
        numOfOccurencesOfEmailAddress++;
  }

  // Adds number at the end for new username if another already has it
  if (numOfOccurencesOfFullName > 0 || numOfOccurencesOfStudentNumber > 0 || numOfOccurencesOfEmailAddress > 0) {
    var duplicateRowIndex = -1;
    for (var i = 1; i < capesCardsData.length; i++) {
      var rowSurname = String(capesCardsData[i][capesCardsSheetColumns["surname"]]).trim().toLowerCase();
      var rowFirstName = String(capesCardsData[i][capesCardsSheetColumns["firstName"]]).trim().toLowerCase();
      var rowMiddleInitial = String(capesCardsData[i][capesCardsSheetColumns["middleInitial"]]).trim().toLowerCase();
      var rowFullName = rowSurname + "," + rowFirstName + "," + rowMiddleInitial.charAt(0);

      var rowStudentNumber = String(capesCardsData[i][capesCardsSheetColumns["studentNumber"]]).trim();
      var rowEmail = String(capesCardsData[i][capesCardsSheetColumns["emailAddress"]]).trim().toLowerCase();

      if (rowFullName === fullNameCheck || rowStudentNumber === studentNumberCheck || rowEmail === emailAddressCheck) {
        duplicateRowIndex = i;
        break;
      }
    }

    if (duplicateRowIndex !== -1) {
      var duplicateSheetRow = duplicateRowIndex + 1;  // sheet row number
      var existingUsername = capesCardsData[duplicateRowIndex][capesCardsSheetColumns["capesUsername"]];
      var existingStudentNumber = capesCardsData[duplicateRowIndex][capesCardsSheetColumns["studentNumber"]];
      Logger.log("Duplicate found in row " + duplicateSheetRow + " with username " + existingUsername);

      capesUsername = existingUsername;
      studentData["capesUsername"] = existingUsername;

      studentData["studentNumber"] = existingStudentNumber;

      for (key of Object.keys(capesCardsSheetColumns)) {

        var cell = capesCardsSheet.getRange(duplicateSheetRow, capesCardsSheetColumns[key] + 1);
        
        if (key == "universityName") {
          if (studentData["isFromUPD"] == "Yes") {
            cell.setValue('University of the Philippines Diliman');
          } else {
            cell.setValue(studentData["universityName"]);
          }  
        } else if (key == 'capesUsername'){
            cell.setValue(studentData["capesUsername"]);
        } else if (key == 'affiliatedOrganization1') {
            cell.setValue(studentData["affiliatedOrganization1"]);
        } else if (key == 'affiliatedOrganization2') {
        cell.setValue(studentData["affiliatedOrganization2"]);
          } else if (key == 'affiliatedOrganization3') {
        cell.setValue(studentData["affiliatedOrganization3"]);
        } else if (key == 'updateStatus') {
        cell.setValue("UPDATED");
        } else if (key == 'capesUsername') {
        cell.setValue(studentData['capesUsername']);
        } else if (key == 'studentNumber') {
        cell.setValue(studentData['studentNumber']);
        } else {
            cell.setValue(studentData[key]);
        }
      }
      return [studentData, true];

    }
    }
  else {
    if (numOfOccurencesOfUsername > 0) {
      capesUsername = capesUsername + (numOfOccurencesOfUsername + 1).toString();
    }
  }

  Logger.log(capesUsername);

  studentData["capesUsername"] = capesUsername;
  // Sets CAPES Username in appropriate cell in CAPES Cards Sheet
  cell.setValue(capesUsername);

  // [CREATING QR]
  //var imageData = UrlFetchApp.fetch('chart.googleapis.com/chart', { 'method' : 'post', 'payload' : { 'cht': 'qr', 'chl': capesID, 'chs': '300x300' }}).getBlob();
  var imageData = UrlFetchApp.fetch('https://quickchart.io/chart?cht=qr&chs=150x150&chl=' + capesUsername).getBlob();
  let fileId = DriveApp.createFile(imageData).setName(capesUsername + '_QR.png').getId();
  DriveApp.getFileById(fileId).moveTo(DriveApp.getFolderById('1LrPWLf8DONPV4qs-z7SgTW4VSgIJPMUN'));
  
  studentData["capesCardQRFileID"] = fileId;

  // Puts data values from 'Form Responses 1' sheet to 'CAPES Cards' sheet
  for (key of Object.keys(capesCardsSheetColumns)) {

    var cell = capesCardsSheet.getRange(capesCardsSheetLastRow, capesCardsSheetColumns[key] + 1);
    
    if (key == "universityName") {
      if (studentData["isFromUPD"] == "Yes") {
        cell.setValue('University of the Philippines Diliman');
      } else {
        cell.setValue(studentData["universityName"]);
      }  
    } else if (key == 'capesUsername'){
        cell.setValue(studentData["capesUsername"]);
    } else if (key == 'affiliatedOrganization1') {
        cell.setValue(studentData["affiliatedOrganization1"]);
    } else if (key == 'affiliatedOrganization2') {
    cell.setValue(studentData["affiliatedOrganization2"]);
      } else if (key == 'affiliatedOrganization3') {
    cell.setValue(studentData["affiliatedOrganization3"]);
    } else if (key == 'updateStatus') {
    } else {
        cell.setValue(studentData[key]);
    }
  }

  return [studentData, false];
}


function replaceTextToImage(body, searchText, fileId) {
  var width = 250; // Please set this.
  var blob = DriveApp.getFileById(fileId).getBlob();
  var r = body.findText(searchText).getElement();
  r.asText().setText("");
  var img = r.getParent().asParagraph().insertInlineImage(0, blob);
  var w = img.getWidth();
  var h = img.getHeight();
  img.setWidth(width);
  img.setHeight(width * h / w);
}

function generateID(studentData) {
  var templateId = '1BgAa9aL9uhNKv5K8edZaCTjzasp4h_-rBTIod1JXUz8'; // template Doc ID
  var folderId = '1vCbkyj8gTJexXTEwWYvehrT0d-bcqxBt'; // target folder ID
  
  var dir = DriveApp.getFolderById(folderId);

  // Create copy of the template
  var copy = DriveApp.getFileById(templateId).makeCopy('CAPES_ID_' + studentData["capesUsername"]);
  var docId = copy.getId();

  Logger.log("Created Doc ID: " + docId);
  Logger.log("Folder ID: " + dir.getId());

  // Move copy to the correct folder
  DriveApp.getFileById(docId).moveTo(dir);

  // Open document for editing
  var doc = DocumentApp.openById(docId);
  var body = doc.getBody();

  // Replace placeholders
  // Combine name properly
  var fullName = studentData["firstName"] + " " + studentData["middleInitial"] + " " + studentData["surname"];
  var academicYear = "2025-2026";

  body.replaceText("<<academicyear>>", academicYear);
  body.replaceText("<<name>>", fullName);
  body.replaceText("<<ylevel>>", studentData["yearLevel"]);
  body.replaceText("<<course>>", studentData["degreeProgram"]);
  body.replaceText("<<username>>", studentData["capesUsername"]);
  body.replaceText("<<university>>", studentData["universityName"]);

  replaceTextToImage(body, '<<CAPES QR>>', studentData["capesCardQRFileID"]);

  doc.saveAndClose();

  // Create PDF version
  var pdfBlob = copy.getAs('application/pdf');
  var pdfCopy = dir.createFile(pdfBlob);
  pdfCopy.setName('CAPES_ID_' + studentData["capesUsername"] + '.pdf');

  Logger.log("Created PDF ID: " + pdfCopy.getId());

  return { docId: docId, pdfId: pdfCopy.getId() };
}

function sendEmailOldQr(studentData) {
  Logger.log(studentData)
  var folderId = '1vCbkyj8gTJexXTEwWYvehrT0d-bcqxBt'; // Folder containing old CAPES IDs
  var folder = DriveApp.getFolderById(folderId);
  var files = folder.getFilesByName('CAPES_ID_' + studentData["capesUsername"] + '.pdf');

  if (files.hasNext()) {
    var file = files.next();
    var pdfId = file.getId();
    var url_pdf = "https://drive.google.com/file/d/" + pdfId + "/view?usp=sharing";
    var linkword = '<a href="' + url_pdf + '">CAPES Card</a>';

    var emailSubject = 'UP CAPES 2526 CAPES ID - ' + studentData["capesUsername"];

    var msg = '<p>Thank you for requesting a CAPES Card! Upon checking our records, we found that a CAPES Card ID already exists under your given information.'
            + ' Here is the link to the PDF file containing your ' + linkword + '.'
            + ' Please review the information and confirm that it is correct. If there are any errors, kindly contact the CAPES Card team through capescards@upcapes.org.</p>'
            + '<p>Please save a copy of your CAPES Card. Scanning of your CAPES QR will be required for entry to all future CAPES events.'
            + ' Thank you, and we hope to see you in our upcoming events!</p>';

    var emailBody = 
      '<p>Good day, ' + studentData["nickname"] + '!</p><p>' + msg
      + '</p><p>Regards,</p><p>UP CAPES Cards Team</p>';

    var aliases = GmailApp.getAliases();

    if (aliases.includes("capescards@upcapes.org")) {
      var options = {
        htmlBody: emailBody,
        from: "capescards@upcapes.org",
        name: "UP CAPES Cards Team"
      };
      GmailApp.sendEmail(studentData["emailAddress"], emailSubject, emailBody, options); 
    } else {
      var message = {
        to: studentData["emailAddress"],
        subject: emailSubject,
        htmlBody: emailBody,
        name: "UP CAPES Cards Team"
      };
      MailApp.sendEmail(message);
    }

    Logger.log("Sent OLD QR email to " + studentData["emailAddress"]);
  } else {
    Logger.log("No old CAPES Card found for username: " + studentData["capesUsername"]);
  }
}


function sendEmailNewQr(studentData, pdfId) {
  Logger.log("pdfID: " + pdfId);
  var pdfIdStr = String(pdfId);
  Logger.log(pdfIdStr);
  var url_pdf = "https://drive.google.com/file/d/" + pdfIdStr + "/view?usp=sharing";
  var linkword = '<a href="' + url_pdf + '">CAPES Card</a>';

  var emailSubject = 'UP CAPES 2526 CAPES ID - ' + studentData["capesUsername"];

  var msg = '<p>Thank you for requesting for a CAPES Card! Attached to this email is a pdf file containing your ' 
          + linkword 
          + ' with your unique CAPES QR and CAPES username. Please review all the information and confirm that it is correct. If there are any errors, please contact the CAPES Card team through capescards@upcapes.org.</p>'
          + '<p>Please save a copy of your CAPES Card. Scanning of your CAPES QR will be required for entry to all future CAPES events.</p>'
          + '<p>Thank you, and we hope to see you in our upcoming events!</p>';

  var emailBody = 
    '<p>Good day, ' + studentData["nickname"] + '!</p><p>' + msg
    + '</p><p>Regards,</p><p>UP CAPES Cards Team</p>';

  var aliases = GmailApp.getAliases();

  // Uses email alias if possible
  if (aliases.includes("capescards@upcapes.org")) {
    var options = {
      htmlBody: emailBody,
      from: "capescards@upcapes.org",
      name: "UP CAPES Cards Team"
    };
    GmailApp.sendEmail(studentData["emailAddress"], emailSubject, emailBody, options); 
  } else {
    var message = {
      to: studentData["emailAddress"],
      subject: emailSubject,
      htmlBody: emailBody,
      name: "UP CAPES Cards Team"
    };
    MailApp.sendEmail(message);
  }

  console.log("Sent email to " + studentData["emailAddress"]);
}


function onFormSubmit(e) {

  const lock = LockService.getScriptLock();
  lock.tryLock(200000);
  if (lock.hasLock()){

    var studentData = getData(e);
    Logger.log("This is the studentData" + studentData);

    if (studentData[1] == true) {
      sendEmailOldQr(studentData[0]);
    }
    else {
      var ids = generateID(studentData[0]);   // generateID returns an object
      var pdfId = ids.pdfId;               // get only the pdfId string
      // var pdfId = generateID(studentData[0]);

      sendEmailNewQr(studentData[0], pdfId);
    }

    lock.releaseLock();
  }
}
