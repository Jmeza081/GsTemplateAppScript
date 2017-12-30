# GsTemplateAppScript

A google sheet script for generating named sheets from a grading template

_Note:_ I will more than likely try to get this code published in order to make the script a lot more user-friendly.

## Installation And Usage

1. Begin by copying the specified google sheet template into your own google account or drive. I recommend saving the example spreadsheet as a template so that it can continually be reused once this setup has been completed. You can import the existing excel sheet by doing the following: `File > import > upload > create new spreadsheet`
2. Navigate to `Tools > Script Editor` to open the embedded code editor for your spread sheet.
3. Create a new `project` that we will use to paste our code into.
4. Take the contents of the `src/generateFromTemplate.js` and embed them in the scriptEditor.
5. Under the provided toolbar, select the entry point for the script by selecting `generateSheets` from the functions dropdown.
6. Before attempting to run the script see the [Config](#config). When you are ready to run the script, you will be asked to accept permissions so that the script can access and modify your spreadsheet (Once this is published as an extension, this whole process can be by-passed).

## Functions

<dl>
  <dt>
  <a href="#generateSheets">generateSheets()</a>
</dt>
  <dd>
  <p>Entry for application: generates new sheets from
a grading template, as well as a summary of student scores.</p>
</dd>
  <dt>
  <a href="#generateStudentSheetNames">generateStudentSheetNames()</a>
</dt>
  <dd>
  <p>Generates an array of student names to be fed into the sheetGenerator</p>
</dd>
  <dt>
  <a href="#createStudentAssessmentSheets">createStudentAssessmentSheets()</a>
</dt>
  <dd>
  <p>Inserts new sheets into workbook named after student names,
and based on the grading template supplied.</p>
</dd>
  <dt>
  <a href="#createAssessmentSummarySheet">createAssessmentSummarySheet()</a>
</dt>
  <dd>
  <p>Inserts new sheet into workbook to summarize student scores.</p>
</dd>
  <dt>
  <a href="#interpolateStudentScores">interpolateStudentScores()</a>
</dt>
  <dd>
  <p>Adds the cell references for each student's score to the summary sheet.</p>
</dd>
</dl>

[]()

## generateSheets()

Entry for application: generates new sheets from a grading template, as well as a summary of student scores.

**Kind**: global function []()

## generateStudentSheetNames()

Generates an array of student names to be fed into the sheetGenerator

**Kind**: global function **Param:**: `SpreadSheet` ss **Returns:**: `Array` array of concatenated student names []()

## createStudentAssessmentSheets()

Inserts new sheets into workbook named after student names, and based on the grading template supplied.

**Kind**: global function **Param:**: `Spreadsheet` ss, **Param:**: `Array` studentNames []()

## createAssessmentSummarySheet()

Inserts new sheet into workbook to summarize student scores.

**Kind**: global function **Param:**: `Spreadsheet` ss **Returns:**: `Sheet` assessmentSummarySheet []()

## interpolateStudentScores()

Adds the cell references for each student's score to the summary sheet.

**Kind**: global function **Param:**: `Array` studentNames **Param:**: `Sheet` studentScoreSheet

## Config

This script can be configured to use a specific name for the template sheet used in generating new sheets, as well as the output summary sheet. See the full configuration options below:

```
var appConfig = {
  scoreCellRef: 'B17', // Required
  rosterSheetName: 'Roster',
  templateSheetName: 'Grading_Template',
  studentScoreSummaryConfig: {
    sheetName: 'Student_Scores',
    headers: ['Student Ref Name', 'Score', 'Special Notes'],
  },
};
```

**IMPORTANT:** it's important to correctly configure the cell reference of the template's score value. This cell reference is used by the script to correctly interpolate a students score based on the student's sheet.

For example, if our template sheet has the student's assignment score embedded at cell `B17`, we must add this to the `scoreCellRef` section of the config in order for the script to correctly generate the summary sheet. Not doing so will throw an error.
