'use strict';

var appConfig = {
  scoreCellRef: 'B17',// Required
  rosterSheetName: 'Roster',
  templateSheetName: 'Grading_Template',
  studentScoreSummaryConfig: {
    sheetName: 'Student_Scores',
    headers: ['Student Ref Name', 'Score', 'Special Notes'],
  },
};

/**
Entry for application: generates new sheets from
a grading template, as well as a summary of student scores.
*/
function generateSheets() {
  // Get reference to spreadsheet app
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  if (appConfig.scoreCellRef === ''){
    ss.toast("Please add the cell reference for a student's score based on your template!",
             'Misconfiguration Error: ');
    return;
  }

  var studentNames = generateStudentSheetNames(ss);

  createStudentAssessmentSheets(ss, studentNames);

  var assessmentSummarySheet = createAssessmentSummarySheet(ss);

  interpolateStudentScores(studentNames, assessmentSummarySheet);

  ss.toast("Script ran to completion!", 'Notice: ');
}

/**
Generates an array of student names to be fed into the sheetGenerator

@param: {SpreadSheet} ss
@returns: {Array} array of concatenated student names
*/
function generateStudentSheetNames(ss){

  var rosterSheet = ss.getSheetByName(appConfig.rosterSheetName);

  // Get the size of the roster from the spreadsheet
  var rosterSize = rosterSheet.getLastRow();

  // Create a list of student names to feed into the insertSheet function
  var studentNames = [];
  for(var i=2; i <= rosterSize; i++){ // Skip header row
    var studentFullName = (
      rosterSheet.getRange(i, 1).getDisplayValue()
      + rosterSheet.getRange(i, 2).getDisplayValue()
    );
    studentNames.push(studentFullName);
  }

  return studentNames;
}

/**
Inserts new sheets into workbook named after student names,
and based on the grading template supplied.

@param: {Spreadsheet} ss,
@param: {Array} studentNames
*/
function createStudentAssessmentSheets(ss, studentNames){
  var gradingTemplate = ss.getSheetByName(appConfig.templateSheetName);
  gradingTemplate.activate(); // Focus sheet to insert templates after

  // Create new student-named sheet from grading template
  for(var i = 0; i < studentNames.length; i++){
      ss.insertSheet(studentNames[i], {template: gradingTemplate});
  }
}


/**
Inserts new sheet into workbook to summarize student scores.

@param: {Spreadsheet} ss
@returns: {Sheet} assessmentSummarySheet
*/
function createAssessmentSummarySheet(ss){
  // Create new student score summary sheet and add headers
  ss.insertSheet(appConfig.studentScoreSummaryConfig.sheetName);
  var assessmentSummarySheet = ss.getSheetByName(
    appConfig
    .studentScoreSummaryConfig
    .sheetName
  );
  assessmentSummarySheet.activate();
  assessmentSummarySheet.appendRow(appConfig.studentScoreSummaryConfig.headers);

  return assessmentSummarySheet;
}


/**
Adds the cell references for each student's score to the summary sheet.

@param: {Array} studentNames
@param: {Sheet} studentScoreSheet
*/
function interpolateStudentScores(studentNames, studentScoreSheet){
  /*
  Iterate through student reference names and add cell reference to the
  appropriate student's score
  */
  for(var i = 0; i < studentNames.length; i++){
    var studentRefName = studentNames[i];
    studentScoreSheet.appendRow([
      studentRefName,
      '=' + studentRefName + '!' + appConfig.scoreCellRef
    ]);
  }
}
