var connectionName = 'quality-228008:asia-east1:wallace-quality';
var user = 'wallace';
var userPwd = 'datong180#';
var db = 'quality';

var dbUrl = 'jdbc:google:mysql://' + connectionName + '/' + db;

function updateAll() {
  addProofreading()
  addQuality()
}

function findUserId() {
  //Returns the employee_id associated with the current spreadsheet's id
  var id = SpreadsheetApp.getActiveSpreadsheet().getId()
  var conn = Jdbc.getCloudSqlConnection(dbUrl, user, userPwd)
  var stmt = conn.prepareStatement("SELECT Employee_id FROM employees WHERE Employee_dashboard_id = ?")
  stmt.setString(1, id)
  var results = stmt.executeQuery()
  var numCols = results.getMetaData().getColumnCount();
  var data = []
  while (results.next()) {
    var row = []
    for (var col = 0; col < numCols; col++) {
      row.push(results.getString(col + 1));
    }
    data.push(row)
  }
  return data[0][0]
}

function addColumn(name, user_id, sheet, column_id, start_row, team) {
  //Adds a single column of data from the GCS mysql database
  //column_id is the 
  var conn = Jdbc.getCloudSqlConnection(dbUrl, user, userPwd)
  var stmt = conn.createStatement()
  if (sheet.getSheetName() == "Proofreading Feedback") {
    var results = stmt.executeQuery("SELECT GROUP_CONCAT(JSON_ARRAY(" 
                                    + name 
                                    + ") ORDER BY t1.Proofreading_deadline DESC, t1.Proofreader_id)" 
                                    + "FROM cases t1 " 
                                    + "INNER JOIN Word_count_limit_updates t2 ON t1.Proofreader_id = t2.Proofreader_id "
                                    + "INNER JOIN employees t3 ON t1.Proofreader_id = t3.Employee_id "
                                    + "INNER JOIN employees t4 ON t1.Editor_id = t4.Employee_id "
                                    + "WHERE Proofreading_deadline IS NOT NULL "
                                    + "AND t1.Proofreading_deadline > Date "
                                    + "AND (t1.Proofreading_deadline < End_date OR End_date IS NULL) "
                                    + "AND t1.Editor_id = " + user_id)
  }
  else if (sheet.getSheetName() == "Quality History"){
    
    var editor = ""
    if (!team) {
      editor =  "AND t1.Editor_id = " + user_id
    }
    Logger.log(editor)
    var results = stmt.executeQuery("SELECT GROUP_CONCAT(JSON_ARRAY(Field_name) ORDER BY Month)"
                                    + "FROM (SELECT DATE_FORMAT(t1.Proofreading_deadline, '%Y-%m') AS Month," 
                                    + name + " AS Field_name "
                                    + "FROM cases t1 INNER JOIN Word_count_limit_updates t2 ON t1.Proofreader_id = t2.Proofreader_id "
                                    + "WHERE t1.Proofreading_deadline IS NOT NULL " + editor
                                    + " GROUP BY Month ORDER BY Month) AS Grouped")
  }
  var numCols = results.getMetaData().getColumnCount();
  var data = []
  while (results.next()) {
    var row = []
    for (var col = 0; col < numCols; col++) {
      row.push(results.getString(col + 1));
    }
    data.push(row)
  }
  var json = JSON.parse('{"field": [' + data + ']}')
  sheet.getRange(start_row, column_id, json.field.length, 1).setValues(json.field)
}

function addProofreading() {
  //Column names must match mysql columns
  var column_names = [
                      "t1.Proofreading_deadline",
                      "CONCAT('O', SUBSTR(t1.Case_id, 8, 6))",
                      "t3.Employee_f_name",
                      "t1.Word_count",
                      "(20000 / t2.New_word_count_limit) * (t1.Word_count/t1.Time)",
                      "t1.C",
                      "t1.G",
                      "t1.M",
                      "t1.A",
                      "t1.L",
                      "t1.P",
                      "t1.Comments",
                      "t1.Editing_distance / t1.Word_count",
                      "t1.Proofreading_distance / t1.Word_count"
                      ]
  
  var proofreading_sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Proofreading Feedback")
  proofreading_sheet.getRange(4, 1, proofreading_sheet.getMaxRows() - 4, 14).clearContent()
  var user_id = findUserId()
  for (var i = 0; i < column_names.length; i++) {
    addColumn(column_names[i], user_id, proofreading_sheet, i + 1, 4, team=false)
  }
}

function addQuality() {
  
  var column_names_individual = [
                      "DATE_FORMAT(MAX(t1.Proofreading_deadline), '%b %Y')",
                      "AVG((20000 / t2.New_word_count_limit) * (t1.Word_count/t1.Time))",
                      "SUM(((t2.New_word_count_limit * t1.Time)/20000) - (t1.Word_count/41.67))/60",
                      "SUM(t1.Word_count)",
                      "AVG((t1.C + t1.G + t1.M + t1.A + t1.L + t1.P)/6)",
                      "DATE_FORMAT(MAX(t1.Proofreading_deadline), '%b %Y')",
                      "SUM(Internal_complaint)"
                      ]
  var column_names_team = [
                      "DATE_FORMAT(MAX(t1.Proofreading_deadline), '%b %Y')",
                      "AVG((20000 / t2.New_word_count_limit) * (t1.Word_count/t1.Time))",
                      "SUM(((t2.New_word_count_limit * t1.Time)/20000) - (t1.Word_count/41.67))/60",
                      "SUM(t1.Word_count)",
                      "AVG((t1.C + t1.G + t1.M + t1.A + t1.L + t1.P)/6)"
                      ]
  var column_ids_individual = [9, 10, 12, 14, 16, 37, 38]
  var column_ids_team = [1, 2, 3, 4, 5]
  var quality_sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Quality History")
  quality_sheet.getRange(4, 9, quality_sheet.getMaxRows() - 4, 2).clearContent()
  quality_sheet.getRange(4, 12, quality_sheet.getMaxRows() - 4, 1).clearContent()
  quality_sheet.getRange(4, 14, quality_sheet.getMaxRows() - 4, 1).clearContent()
  quality_sheet.getRange(4, 16, quality_sheet.getMaxRows() - 4, 1).clearContent()
  quality_sheet.getRange(4, 2, quality_sheet.getMaxRows() - 4, 5).clearContent()
  quality_sheet.getRange(4, 37, quality_sheet.getMaxRows() - 4, 2).clearContent()
  //Add individual quality data
  var user_id = findUserId()
  for (var i = 0; i < column_names_individual.length; i++) {
    addColumn(column_names_individual[i], user_id, quality_sheet, column_ids_individual[i], 4, team=false)
  }
  //Add team quality data
  for (var i = 0; i < column_names_team.length; i++) {
    addColumn(column_names_team[i], user_id, quality_sheet, column_ids_team[i], 4, team=true)
  }
}