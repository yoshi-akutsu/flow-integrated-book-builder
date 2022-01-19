function doGet(e) {
  let paramsArray = [];
  let spreadsheetId = e.parameter.spreadsheeturl.split("/")[5];
  let spreadsheet = SpreadsheetApp.openById(spreadsheetId);
  let sheet = spreadsheet.getSheetByName("Schools");
  let schoolInfo = {};
  for (let item in e.parameter) {
    paramsArray.push([item, e.parameter[item]]);
  }
  for (let j = 0; j < e.parameter.numschools; j++) {
    let schoolArray = [];
    let programArray = [];
    for (let i = 0; i < paramsArray.length; i++) {
      if (paramsArray[i][0].split(".")[0] == "school" + (j + 1)) {
        if (
          paramsArray[i][0].split(".")[1] == "name" &&
          paramsArray[i][0].split(".")[0] == "school" + (j + 1)
        ) {
          schoolInfo["schoolName" + (j + 1)] = paramsArray[i][1];
        }
        if (paramsArray[i][0].split(".")[1].includes("reason")) {
          schoolInfo[paramsArray[i][0].split(".")[1] + (j + 1)] =
            paramsArray[i][1];
        }
        if (
          paramsArray[i][0].split(".")[1].includes("program") ||
          paramsArray[i][0].split(".")[1].includes("link")
        ) {
          programArray.push(paramsArray[i]);
          if (paramsArray[i][0].split(".")[1].includes("numprograms")) {
            schoolArray.push(paramsArray[i]);
          }
        } else {
          schoolArray.push(paramsArray[i]);
        }
      }
    }
    let schoolObject = {};
    for (let i = 0; i < schoolArray.length; i++) {
      schoolObject[schoolArray[i][0]] = schoolArray[i][1];
    }
    printSchool(schoolObject, sheet, j + 1, programArray);
  }
  let contactInfo = {};
  contactInfo.studentName = e.parameter.studentname;
  contactInfo.studentEmail = e.parameter.studentemail;
  contactInfo.parent1Email = e.parameter.parent1email;
  contactInfo.parent2Email = e.parameter.parent2email;
  //draftEmail(schoolInfo, contactInfo);

  return HtmlService.createHtmlOutput(
    "Schools and program links added to the book"
  );
}

let advisorData = {};
let userEmail = Session.getActiveUser().getEmail();

switch (userEmail.split("@")[0]) {
  case "yoshi":
    advisorData.fullName = "Yoshi Akutsu";
    advisorData.position = "Lead Advisor";
    advisorData.phoneExtension = "709";
    advisorData.advisorEmail = "yoshi@incollegeplanning.com";
    break;
  case "juleanna":
    advisorData.fullName = "Juleanna Smith";
    advisorData.position = "Lead Advisor";
    advisorData.phoneExtension = "710";
    advisorData.advisorEmail = "juleanna@incollegeplanning.com";
    break;
  case "emma":
    advisorData.fullName = "Emma Mote";
    advisorData.position = "Lead Advisor";
    advisorData.phoneExtension = "708";
    advisorData.advisorEmail = "emma@incollegeplanning.com";
    break;
  case "sara":
    advisorData.fullName = "Sara Kapaj";
    advisorData.position = "Lead Advisor";
    advisorData.phoneExtension = "711";
    advisorData.advisorEmail = "sara@incollegeplanning.com";
    break;
  case "hannah":
    advisorData.fullName = "Hannah Laubach";
    advisorData.position = "College Planning Advisor";
    advisorData.phoneExtension = "712";
    advisorData.advisorEmail = "hannah@incollegeplanning.com";
    break;
  case "sian":
    advisorData.fullName = "Siân Lewis";
    advisorData.position = "College Planning Advisor";
    advisorData.phoneExtension = "723";
    advisorData.advisorEmail = "sian@incollegeplanning.com";
    break;
  case "eric":
    advisorData.fullName = "Eric Martinez";
    advisorData.position = "College Planning Advisor";
    advisorData.phoneExtension = "718";
    advisorData.advisorEmail = "eric@incollegeplanning.com";
    break;
  case "alecea":
    advisorData.fullName = "Alecea Howell";
    advisorData.position = "College Planning Advisor";
    advisorData.phoneExtension = "721";
    advisorData.advisorEmail = "alecea@incollegeplanning.com";
    break;
  case "samantha":
    advisorData.fullName = "Sam Rubinoski";
    advisorData.position = "College Planning Advisor";
    advisorData.phoneExtension = "720";
    advisorData.advisorEmail = "samantha@incollegeplanning.com";
    break;
  case "reilly":
    advisorData.fullName = "Reilly Grealis";
    advisorData.position = "College Planning Advisor";
    advisorData.phoneExtension = "719";
    advisorData.advisorEmail = "reilly@incollegeplanning.com";
    break;
  case "alex":
    advisorData.fullName = "Alex Horn";
    advisorData.position = "College Planning Advisor";
    advisorData.phoneExtension = "711";
    advisorData.advisorEmail = "alex@incollegeplanning.com";
    break;
  case "drake":
    advisorData.fullName = "Drake Hankins";
    advisorData.position = "College Planning Advisor";
    advisorData.phoneExtension = "702";
    advisorData.advisorEmail = "drake@incollegeplanning.com";
    break;
  case "sarah":
    advisorData.fullName = "Sarah Cook";
    advisorData.position = "College Planning Advisor";
    advisorData.phoneExtension = "700";
    advisorData.advisorEmail = "sarah@incollegeplanning.com";
    break;
  case "lydia":
    advisorData.fullName = "Lydia Crannell";
    advisorData.position = "College Planning Advisor";
    advisorData.phoneExtension = "713";
    advisorData.advisorEmail = "lydia@incollegeplanning.com";
    break;
  default:
    break;
}

function getFromEmail() {
  let aliases = GmailApp.getAliases();
  for (let i = 0; i < aliases.length; i++) {
    if (aliases[i].includes("incoll")) {
      return aliases[i];
    }
  }
}

function draftEmail(schoolObject, contactObject) {
  Logger.log(schoolObject);
  Logger.log(contactObject);
  let template = HtmlService.createTemplateFromFile("email");
  let fromEmail = getFromEmail();
  template.advisor = advisorData;
  template.schoolInfo = schoolObject;
  template.name = contactObject.studentName;

  let message = template.evaluate().getContent();
  let emailTitle = "Your Next Book";
  let emails = contactObject.studentEmail;
  if (contactObject.parent1Email.length > 0) {
    emails += "," + contactObject.parent1Email;
  }
  if (contactObject.parent2Email.length > 0) {
    emails += "," + contactObject.parent2Email;
  }
  GmailApp.createDraft(emails, emailTitle, message, {
    htmlBody: message,
    from: fromEmail,
  });
}

function printSchool(object, sheet, iteration, programs) {
  let school = "school" + iteration;
  let rowToColor = sheet.getDataRange().getNumRows() + 2;
  sheet.getRange(rowToColor, 1, 1, 14).setBackgroundColor("#b6dde8");
  sheet.getRange(1, 14, 1000, 1).setBackgroundColor("#b6dde8");
  let topleft = sheet.getDataRange().getNumRows() + 3;

  // School Name
  sheet
    .getRange("A" + topleft)
    .setValue(object[school + ".name"])
    .setFontColor("#366092")
    .setFontWeight("bold");
  // Address
  sheet.getRange("A" + (topleft + 2)).setValue("Address");
  sheet.getRange("B" + (topleft + 2)).setValue(object[school + ".address"]);
  // Website
  sheet.getRange("A" + (topleft + 3)).setValue("Website");
  sheet.getRange("B" + (topleft + 3)).setValue(object[school + ".website"]);
  // Private/Public
  sheet.getRange("A" + (topleft + 4)).setValue("Private/Public");
  sheet
    .getRange("B" + (topleft + 4))
    .setValue(object[school + ".publicprivate"]);
  // Enrollment
  sheet.getRange("A" + (topleft + 5)).setValue("Enrollment");
  sheet.getRange("B" + (topleft + 5)).setValue(object[school + ".enrollment"]);
  // Early Action Deadline
  sheet.getRange("A" + (topleft + 6)).setValue("Early Action Deadline");
  sheet.getRange("B" + (topleft + 6)).setValue(object[school + ".earlyaction"]);
  // Regular Decision Deadline
  sheet.getRange("A" + (topleft + 7)).setValue("Regular Decision Deadline");
  sheet
    .getRange("B" + (topleft + 7))
    .setValue(object[school + ".regulardecision"]);
  // Accepts Common App
  sheet.getRange("A" + (topleft + 8)).setValue("Accepts Common App");
  sheet.getRange("B" + (topleft + 8)).setValue(object[school + ".commonapp"]);

  // Key Stats
  sheet
    .getRange("D" + (topleft + 2))
    .setValue("Key Stats")
    .setFontWeight("bold");
  // U.S. News College Ranking
  sheet.getRange("D" + (topleft + 3)).setValue("U.S. News College Ranking");
  sheet.getRange("E" + (topleft + 3)).setValue(object[school + ".usnews"]);
  // ACT Score Avg
  sheet.getRange("D" + (topleft + 4)).setValue("ACT Score Avg");
  sheet.getRange("E" + (topleft + 4)).setValue(object[school + ".act"]);
  // SAT Score Avg
  sheet.getRange("D" + (topleft + 5)).setValue("SAT Score Avg");
  sheet.getRange("E" + (topleft + 5)).setValue(object[school + ".sat"]);
  // Avg High School GPA
  sheet.getRange("D" + (topleft + 6)).setValue("Avg High School GPA");
  sheet.getRange("E" + (topleft + 6)).setValue(object[school + ".gpa"]);
  // App Fee
  sheet.getRange("D" + (topleft + 7)).setValue("App Fee");
  sheet
    .getRange("E" + (topleft + 7))
    .setValue(object[school + ".fee"])
    .setNumberFormat("$0");
  // Freshman Retention
  sheet.getRange("D" + (topleft + 8)).setValue("Freshman Retention");
  sheet
    .getRange("E" + (topleft + 8))
    .setValue(object[school + ".retention"] / 100)
    .setNumberFormat("0%");

  // Cost
  sheet
    .getRange("G" + (topleft + 2))
    .setValue("Cost")
    .setFontWeight("bold");
  // Tuition (In State)
  sheet.getRange("G" + (topleft + 3)).setValue("Tuition (In State)");
  sheet.getRange("H" + (topleft + 3)).setValue(object[school + ".instate"]);
  // Tuition (Out of State)
  sheet.getRange("G" + (topleft + 4)).setValue("Tuition (Out of State)");
  sheet.getRange("H" + (topleft + 4)).setValue(object[school + ".outofstate"]);
  // Room and Board
  sheet.getRange("G" + (topleft + 5)).setValue("Room and Board");
  sheet.getRange("H" + (topleft + 5)).setValue(object[school + ".roomboard"]);
  // Books (avg)
  sheet.getRange("G" + (topleft + 6)).setValue("Books (avg)");
  sheet.getRange("H" + (topleft + 6)).setValue(object[school + ".books"]);
  // Average Debt
  sheet.getRange("G" + (topleft + 7)).setValue("Average Debt");
  sheet.getRange("H" + (topleft + 7)).setValue(object[school + ".avgdebt"]);
  // Proportion who Borrowed
  sheet.getRange("G" + (topleft + 8)).setValue("Proportion who Borrowed");
  sheet
    .getRange("H" + (topleft + 8))
    .setValue(object[school + ".borrowed"] / 100)
    .setNumberFormat("0%");

  // Financial Aid Deadline
  sheet
    .getRange("J" + (topleft + 2))
    .setValue("Financial Aid Deadline")
    .setFontWeight("bold");
  sheet.getRange("K" + (topleft + 2)).setValue(object[school + ".fadeadline"]);
  // % of undergrads applying for aid
  sheet
    .getRange("J" + (topleft + 3))
    .setValue("% of undergrads applying for aid");
  sheet
    .getRange("K" + (topleft + 3))
    .setValue(object[school + ".propapplying"] / 100)
    .setNumberFormat("0%");
  // % of undergrads who received aid
  sheet
    .getRange("J" + (topleft + 4))
    .setValue("% of undergrads who received aid");
  sheet
    .getRange("K" + (topleft + 4))
    .setValue(object[school + ".propreceiving"] / 100)
    .setNumberFormat("0%");
  // % of need met in full
  sheet.getRange("J" + (topleft + 5)).setValue("% of need met in full");
  sheet
    .getRange("K" + (topleft + 5))
    .setValue(object[school + ".propfull"] / 100)
    .setNumberFormat("0%");

  // Endowment
  sheet
    .getRange("L" + (topleft + 2))
    .setValue("Endowment")
    .setFontWeight("bold");
  sheet
    .getRange("M" + (topleft + 2))
    .setValue(object[school + ".endowment"])
    .setNumberFormat('$[<999950]0.0,"K";$[<999950000]0.0,,"M";$0.0,,,"B"');
  // Avg Financial Aid Package
  sheet.getRange("L" + (topleft + 3)).setValue("Avg Financial Aid Package");
  sheet.getRange("M" + (topleft + 3)).setValue(object[school + ".avgpackage"]);
  // Avg non-need Gift Aid
  sheet.getRange("L" + (topleft + 4)).setValue("Avg Non-need Gift Aid");
  sheet.getRange("M" + (topleft + 4)).setValue(object[school + ".avgnonneed"]);
  // % of non-need Gift Aid
  sheet.getRange("L" + (topleft + 5)).setValue("% of Non-need Gift Aid");
  sheet
    .getRange("M" + (topleft + 5))
    .setValue(object[school + ".propnonneed"] / 100)
    .setNumberFormat("0%");
  // Avg Need Based Loan
  sheet.getRange("L" + (topleft + 6)).setValue("Avg Need Based Loan");
  sheet.getRange("M" + (topleft + 6)).setValue(object[school + ".needloan"]);

  let numPrograms = object[school + ".numprograms"];

  for (let i = 0; i < numPrograms; i++) {
    topleft = sheet.getDataRange().getNumRows() + 1;
    let programName = "";
    let programLink = "";
    let outcomesLink = "";
    let catalogLink = "";
    let secondaryLink = "";
    for (let j = 0; j < programs.length; j++) {
      if (programs[j][0].split(".")[1] == "programname" + (i + 1)) {
        programName = programs[j][1];
      }
      if (programs[j][0].split(".")[1] == "programlink" + (i + 1)) {
        programLink = programs[j][1];
      }
      if (programs[j][0].split(".")[1] == "outcomeslink" + (i + 1)) {
        outcomesLink = programs[j][1];
      }
      if (programs[j][0].split(".")[1] == "cataloglink" + (i + 1)) {
        catalogLink = programs[j][1];
      }
      if (programs[j][0].split(".")[1] == "secondarylink" + (i + 1)) {
        secondaryLink = programs[j][1];
      }
    }
    sheet
      .getRange("A" + topleft)
      .setValue(programName)
      .setFontColor("#e36c09")
      .setFontWeight("bold");
    sheet.getRange("B" + topleft).setValue(programLink);
    if (secondaryLink != "") {
      sheet
        .getRange("A" + (topleft + 1))
        .setValue(programName + " cont.")
        .setFontColor("#e36c09")
        .setFontWeight("bold");
      sheet.getRange("B" + (topleft + 1)).setValue(secondaryLink);
      topleft += 1;
    }
    if (catalogLink != "") {
      sheet
        .getRange("A" + (topleft + 1))
        .setValue(programName + " Catalog")
        .setFontColor("#e36c09")
        .setFontWeight("bold");
      sheet.getRange("B" + (topleft + 1)).setValue(catalogLink);
      topleft += 1;
    }
    if (outcomesLink != "") {
      sheet
        .getRange("A" + (topleft + 1))
        .setValue(programName + " Outcomes")
        .setFontColor("#e36c09")
        .setFontWeight("bold");
      sheet.getRange("B" + (topleft + 1)).setValue(outcomesLink);
    }
  }
}
