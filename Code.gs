var data=[];
var spreadsheetID=SpreadsheetApp.getActiveSpreadsheet().getId();
var sheetName='headcount';
var domain=Session.getActiveUser().getEmail().split('@')[1];

function getMembershipCount() {
  var pageToken, page,membership,groupName;
  do {
    page = AdminDirectory.Groups.list({
      domain : domain
    });
    var groups = page.groups;
    for (var i = 0; i< groups.length; i++){
      var group = groups[i];
      membership = group.directMembersCount;
      groupName = group.email;
      data.push([groupName,membership]);
    }
    pageToken = page.nextPageToken;
  }while (pageToken);
  
  var maxRows = SpreadsheetApp.openById(spreadsheetID).getSheetByName(sheetName).getLastRow();
  if (maxRows>1){
    SpreadsheetApp.openById(spreadsheetID).getSheetByName(sheetName).deleteRows(2, maxRows*1-1);
  }
  SpreadsheetApp.openById(spreadsheetID).getSheetByName(sheetName).getRange(2, 1, data.length, data[0].length).setValues(data);
  SpreadsheetApp.openById(spreadsheetID).getSheetByName(sheetName).getRange('E1').setValue(new Date());
}
