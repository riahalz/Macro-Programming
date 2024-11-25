/* Version Control - Log and display Version details each time user modifies cell in sheet */

function versionUpdate(e) {
    var sheet = SpreadsheetApp.getActiveSpreadsheet();
    var controlSheet = sheet.getSheetByName("Version History");
    var activeUser = Session.getActiveUser().getEmail();
    var date = new Date();
    var version = "V" + (controlSheet.getLastRow());
    var description = "Edited cell: " + e.range.getA1Notation();
    controlSheet.appendRow([date, version, description, activeUser, ""]);
}
