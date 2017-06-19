var dbUrl = 'jdbc:mysql://rdsl5o.mysql.rds.aliyuncs.com:3306/uskk_database';
var dbUser = 'gzb';
var dbPwd = 'Blac02';

var SMS_SENT = "SMS_SENT";
var SMS_READY = '';
var toNumber = "6263412345";
var toNumberCS1 = "6263412345";
var toNumberCSATT = "8009019878";
var toEmailCS1 = "he@hotmail.com";
var bodyMessage = "testv33322r";
var SID = "AC5946b7a";
var Token = "5cb312ec28781";
var twilioNumber = "90150123456"
var EMAIL_SENT = "EMAIL_SENT";
var EMAIL_READY = "1";

function runall(){
  fetch_uskkdata(); //复制主库新记录
  copyGtoM();  //同步已有记录
}

function copyGtoM() {   //工作表activation表格同步到主数据库mysql，第一行“处理终结”=不会再同步，“需要同步”在不变除了sim卡号码，工作表的数据覆盖主数据库数据。
  var conn = Jdbc.getConnection(dbUrl, dbUser, dbPwd);
  var stmt = conn.createStatement();
  stmt.setMaxRows(1000);
  var start = new Date();


  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Activation');
  var cell = sheet.getRange('C1');
  var numRows = sheet.getLastRow();
  var row = numRows;

  var data = sheet.getDataRange().getValues();
  var startID = SearchStartDate();
  //for (var i = 1; i < data.length; i++) { 
  for (var i = startID; i < data.length; i++) {
   //for (var i = data.length; i > 0; i--) {
   var simnumber = data[i][8].toString().replace(/[^0-9]/ig, "");
   Logger.log(simnumber);
   
   
   if(data[i][0] != '处理终结' && data[i][0] != '格式错误' && simnumber.length > 18) {  //如果第一行是处理终结字样则不处理,sim卡为空也不处理,或者少于19位数字不处理
     Logger.log(i);
   
   if (data[i][9] != '') {   //测试是否为日期，不是则设置为2000年，需要加提示吗？
     var fd01 = data[i][9];
     // var fd01= convert2jsDate(data[i][9]);
      fd01 = Utilities.formatDate(fd01, "GMT", "yyyy-MM-dd");
     
    }else {fd01 = "2099-01-01"}
    
     if (data[i][10] != '') {   //测试是否为日期，不是则设置为2000年，需要加提示吗？
     var fd02 = data[i][10];
      //var fd02= convert2jsDate(data[i][10]);
      fd02 = Utilities.formatDate(fd02, "GMT", "yyyy-MM-dd");
     
    }else {fd02 = "2099-01-01"}
   

   
  var rs = stmt.executeQuery('select * from activation where replace(replace(replace(replace(replace(replace(replace(ICCID," ",""),"－",""),"（",""),"-",""),"(",""),")",""),"）","") like "%'+simnumber+'%" ORDER BY id DESC limit 1');
   
  if (rs.next()) {   //如果主数据库有记录
    var mysqlid = rs.getString(17);
    var phonenumber = data[i][5].toString().replace(/[^0-9]/ig, ""); 
    var rsColname = rs.getString(1);
    var rsColpukCode= rs.getString(6);
    
    if (rsColname == '客服修改'){   //如果主数据库第一行为‘客服修改’，则同步整行
      for (var icol = 1; icol < 26; icol++) { 
        sheet.getRange(i+1,icol+3).setValue(rs.getString(icol+1));
        }
      //sheet.getRange(i+1,1).setValue("已经同步到M1");
      var sql =  "UPDATE uskk_database.activation SET   name = '已经同步到S1'  WHERE uskk_database.activation.id = '" + mysqlid + "'";
      //var sql =  "UPDATE uskk_database.activation SET uskk_database.activation.MSISDN = '" + phonenumber + "', name = '已经同步到S1' , verplan='" + data[i][3] + "',  Status='" + data[i][4] + "',   MSISDN='" + data[i][5] + "',  taobaoID='" + data[i][6] + "', pukCode='" + data[i][7] + "', ICCID='" + data[i][8] + "', ACTDate='" + fd01 + "', ReturnDate='" + fd02 + "', Plan='" + data[i][11] + "', Contact1='" + data[i][12] + "', IMEI='" + data[i][13] + "', Email1='" + data[i][14] + "', Memo='" + data[i][15] + "' , PlanID='" + data[i][17] + "', C1='" + data[i][20] + "' , C2='" + data[i][21] + "', C3='" + data[i][22] + "', C4='" + data[i][23] + "', C5='" + data[i][24] + "', C6='" + data[i][25] + "', C7='" + data[i][9] + "', C8='" + data[i][10] + "'  WHERE uskk_database.activation.id = '" + mysqlid + "'";
      Logger.log(sql);
      var count = stmt.executeUpdate(sql,1);
      sheet.getRange(i+1,1).setValue("已经同步到M1");
       sql = mysql_real_escape_string(sql);
		  var sqlLog = "INSERT INTO `uskk_database`.`edit_log` (`sql_log`,`user`) VALUES ('" + sql + "', '主库标客服修改更新主库')";
			
		  count = stmt.executeUpdate(sqlLog,1);
       
       }
    
    Logger.log(phonenumber);
    
    if(data[i][0] == '需要同步') {   //第一行填“需要同步”则无条件整条记录同步复制覆盖主数据库
      //var mysqlid = rs.getString(17);
      var sql =  "UPDATE uskk_database.activation SET uskk_database.activation.MSISDN = '" + phonenumber + "', name = '已经同步到S2' , verplan='" + data[i][3] + "',  Status='" + data[i][4] + "',   MSISDN='" + data[i][5] + "',  taobaoID='" + data[i][6] + "', pukCode='" + data[i][7] + "', ICCID='" + data[i][8] + "', ACTDate='" + fd01  + "', ReturnDate='" +  fd02 + "', Plan='" + data[i][11] + "', Contact1='" + data[i][12] + "', IMEI='" + data[i][13] + "', Email1='" + data[i][14] + "', Memo='" + data[i][15] + "' , PlanID='" + data[i][17] + "', C1='" + data[i][20] + "' , C2='" + data[i][21] + "', C3='" + data[i][22] + "', C4='" + data[i][23] + "', C5='" + data[i][24] + "', C6='" + data[i][25] + "', C7='" + data[i][26] + "', C8='" + data[i][27] + "'  WHERE uskk_database.activation.id = '" + mysqlid + "'";
      //var sql =  "UPDATE uskk_database.activation SET uskk_database.activation.MSISDN = '" + phonenumber + "', name = '已经同步' , verplan='" + data[i][3] + "',  Status='" + data[i][4] + "',   MSISDN='" + data[i][5] + "',  taobaoID='" + data[i][6] + "', pukCode='" + data[i][7] + "', ICCID='" + data[i][8] + "', ACTDate='" + fd01 + "', ReturnDate='" + fd02 + "', Plan='" + data[i][11] + "', Contact1='" + data[i][12] + "', IMEI='" + data[i][13] + "', Email1='" + data[i][14] + "', Memo='" + data[i][15] + "' , PlanID='" + data[i][17] + "'  WHERE uskk_database.activation.id = '" + mysqlid + "'";
      //var sql =  "UPDATE `uskk_database`.`activation` SET `name`='已经同步S修改', `verplan`='.$iverplan.', `Status`='.$iStatus.', `MSISDN`='$iMSISDN', `taobaoID`='$itaobaoID', `pukCode`='$ipukCode', `ICCID`='$iICCID', `ACTDate`='$iACTDate', `ReturnDate`='$iReturnDate', `Plan`='$iPlan', `Contact1`='$iContact1', `IMEI`='$iIMEI', `Email1`='$iEmail1', `Memo`='$iMemo', `SubmitTime`='$iSubmitTime', `PlanID`='$iPlanID', `id`='$iid', `lastchange`='$ilastchange', `C1`='$iC1', `C2`='$iC2', `C3`='$iC3', `C4`='$iC4', `C5`='$iC5', `C6`='$iC6', `C7`='$iC7', `C8`='$iC8' WHERE (`id`='$iid')";
     Logger.log(sql);
      var count = stmt.executeUpdate(sql,1);
      sheet.getRange(i+1,1).setValue("已经同步到M2");
      
             sql = mysql_real_escape_string(sql);
		  var sqlLog = "INSERT INTO `uskk_database`.`edit_log` (`sql_log`,`user`) VALUES ('" + sql + "', 'google库标需要同步')";
			
		  count = stmt.executeUpdate(sqlLog,1);
      
      }
    
    if (phonenumber.length > 9 && phonenumber.length <12){ //开通完有号码了
      //var mysqlid = rs.getString(17);
      var sql =  "UPDATE uskk_database.activation SET uskk_database.activation.MSISDN = '" + phonenumber + "', name = '处理终结' WHERE uskk_database.activation.id = '" + mysqlid + "'";
      var count = stmt.executeUpdate(sql,1);
      Logger.log("nextture");
      sheet.getRange(i+1,1).setValue("处理终结");
       sql = mysql_real_escape_string(sql);
		  var sqlLog = "INSERT INTO `uskk_database`.`edit_log` (`sql_log`,`user`) VALUES ('" + sql + "', 'google库有号码标终结')";
			
		  count = stmt.executeUpdate(sqlLog,1);
      
      
      SpreadsheetApp.flush();

    }else {   //有sim卡无号码的处理
      Logger.log("有记录但无号码")
      if(rsColname == '客人更新') {
        var sql =  "UPDATE uskk_database.activation SET name = '已经同步s3' WHERE uskk_database.activation.id = '" + mysqlid + "'";
        var count = stmt.executeUpdate(sql,1);
        sheet.getRange(i+1,1).setValue("客人更新3");
        sheet.getRange(i+1,8).setValue(rsColpukCode);
        sql = mysql_real_escape_string(sql);
		  var sqlLog = "INSERT INTO `uskk_database`.`edit_log` (`sql_log`,`user`) VALUES ('" + sql + "', '主库标客人更新')";
			
		  count = stmt.executeUpdate(sqlLog,1);
         
        }
    }
  }else {   //无记录则插入
    Logger.log("nextfalse");
    Logger.log(data[i][8]);
      var sql =  "INSERT INTO `uskk_database`.`activation` (`name`, `verplan`, `Status`, `MSISDN`, `taobaoID`, `pukCode`, `ICCID`, `ACTDate`, `ReturnDate`, `Plan`, `Contact1`, `IMEI`, `Email1`, `Memo`, `SubmitTime`, `PlanID`, `id`) VALUES ('已经同步','" + data[i][3] + "','" + data[i][4] + "', '" + data[i][5] + "', '" + data[i][6] + "', '" + data[i][7] + "', '" + data[i][8] + "','" + fd01 + "','" + fd02 + "','" + data[i][11] + "', '" + data[i][12] + "', '" + data[i][13] + "', '" + data[i][14] + "', '" + data[i][15] + "', '" + data[i][16] + "', '" + data[i][17] + "', default)";
      Logger.log(sql);
    var count = stmt.executeUpdate(sql,1);
    sheet.getRange(i+1,1).setValue("已经同步S新无记录4");
    var rsid = stmt.getGeneratedKeys(); 
    while(rsid.next()) { 
    sheet.getRange(i+1,19).setValue(rsid.getString(1));
    }
      sql = mysql_real_escape_string(sql);
		  var sqlLog = "INSERT INTO `uskk_database`.`edit_log` (`sql_log`,`user`) VALUES ('" + sql + "', 'google无记录插入主库')";
			
		  count = stmt.executeUpdate(sqlLog,1);
      //SpreadsheetApp.flush();
  //}
 

    
    }
    }
 }
  if (count >0){rs.close()};
  stmt.close();
  conn.close();
  var end = new Date();
  Logger.log('Time elapsed: ' + (end.getTime() - start.getTime()));
  SpreadsheetApp.flush();
}



function fetch_uskkdata() {   //获取主库新数据
  var conn = Jdbc.getConnection(dbUrl, dbUser, dbPwd);
  var stmt = conn.createStatement();
  stmt.setMaxRows(200);
  var start = new Date();
  //var rs = stmt.executeQuery("select * from activation  where ACTDate >= CURDATE() - INTERVAL 90 DAY and (name IS not '已经同步' or name = '处理终结') and CHAR_LENGTH(ICCID)  > 18 ORDER BY id DESC limit 200");
  var rs = stmt.executeQuery("select * from activation  where lastchange >= CURDATE() - INTERVAL 30 DAY and (name not like '%已经同步%' and name not like '%处理终结%') and CHAR_LENGTH(ICCID)  > 18 ORDER BY id DESC limit 200");
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Activation');
  var cell = sheet.getRange('C1');
  var numRows = sheet.getLastRow();
  var row = numRows;
  var mysqlid;
  

  while (rs.next()) {
    Logger.log(rs.getString(7));
    if (SearchColumn(rs.getString(7)) > 1){
      continue;
    }else{
      cell.offset(row, -1).setValue(rs.getString(19));
      for (var col = 0; col < rs.getMetaData().getColumnCount(); col++) {
        cell.offset(row, col).setValue(rs.getString(col + 1));
        mysqlid = rs.getString(17);
    }
    cell.offset(row, -2).setValue("已经同步");
     var sql2 =  "UPDATE uskk_database.activation SET name = '已经同步' WHERE uskk_database.activation.id = '" +mysqlid+ "'";
     var stmt2 = conn.createStatement();
        var count = stmt2.executeUpdate(sql2,1);
    Logger.log(count);
    Logger.log(sql2);
      sql2 = mysql_real_escape_string(sql2);
		  var sqlLog = "INSERT INTO `uskk_database`.`edit_log` (`sql_log`,`user`) VALUES ('" + sql2 + "', '从主库新增到副库')";
          var stmt3 = conn.createStatement();
		  count = stmt3.executeUpdate(sqlLog,1);
      row++;
        stmt3.close();
        stmt2.close();
    }
  
  }
  rs.close();
  //stmt3.close();
  //stmt2.close();
  stmt.close();
  conn.close();
  var end = new Date();
  Logger.log('Time elapsed: ' + (end.getTime() - start.getTime()));
  
}



function SearchColumn(searchString) {  //查找activation的id是否有，并返回行号。
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Activation');
  //var values = sh.getDataRange().getValues();
  var numRows = sh.getLastRow();
  var values = sh.getSheetValues(1, 9, numRows,1);
  
  for(var i=1, iLen=values.length; i<iLen; i++) {
    var iccid1 = values[i][0].toString().replace(/[^0-9]/ig, ""); 
    var iccid2 = searchString.replace(/[^0-9]/ig, ""); 
    if(iccid1 == iccid2) {
      return i+1;
    }
  }     
}


function SearchStartDate() {  //返回今日开通日期开始的行号。-两天
  var searchString       = new Date();
  searchString.setDate(searchString.getDate()-2); //2天前
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Activation');
  //var values = sh.getDataRange().getValues();
  var numRows = sh.getLastRow();
  var values = sh.getSheetValues(1, 10, numRows,1);
  
  for(var i=1, iLen=values.length; i<iLen; i++) {
    var date1 = values[i][0]; 
    var date2 = searchString; 

    date1 = Utilities.formatDate(date1, "GMT", "yyyy-MM-dd");
    date2 = Utilities.formatDate(date2, "GMT", "yyyy-MM-dd");
    if(date1 == date2) {
       Logger.log(i);
      return i+1;
     
    }
  }     
}



function SearchColumn_old(searchString) {  //查找activation的id是否有，并返回行号。
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Activation');
  var values = sh.getDataRange().getValues();

  for(var i=1, iLen=values.length; i<iLen; i++) {
    var iccid1 = values[i][8].toString().replace(/[^0-9]/ig, ""); 
    var iccid2 = searchString.replace(/[^0-9]/ig, ""); 
    if(iccid1 == iccid2) {
      return i+1;
    }
  }     
}


function removeDuplicates() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var data = sheet.getDataRange().getValues();
  var newData = new Array();
  for(i in data){
    var row = data[i];
    var duplicate = false;
    for(j in newData){
      if(row.join() == newData[j].join()){
        duplicate = true;
      }
    }
    if(!duplicate){
      newData.push(row);
    }
  }
  sheet.clearContents();
  sheet.getRange(1, 1, newData.length, newData[0].length).setValues(newData);
}


function sortActivation() {    //给开卡表自动排序
  var sortFirst = 10; //index of column to be sorted by; 1 = column A, 2 = column B, etc.
  var sortFirstAsc = true; //Set to false to sort descending
  var sortSecond = 21;
  var sortSecondAsc = false;
  var sortThird = 12;
  var sortThirdAsc = false;

  var headerRows = 1; 
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Activation');
  var data = sheet.getDataRange().getValues();
  var range = sheet.getRange(headerRows+1, 1, sheet.getMaxRows()-headerRows, sheet.getLastColumn());
  range.sort([{column: sortFirst, ascending: sortFirstAsc}, {column: sortSecond, ascending: sortSecondAsc}, {column: sortThird, ascending: sortThirdAsc}]);
}


function sortActivationShort(){
  sortActivationShort1();
  SpreadsheetApp.flush();
  Utilities.sleep(3500);
  sortActivationShort2();
  SpreadsheetApp.flush();
}
  

function sortActivationShort1() {    //给Tmobile短期卡表自动排序(按照日期）
  var sortFirst = 4; //index of column to be sorted by; 1 = column A, 2 = column B, etc.
  var sortFirstAsc = true; //Set to false to sort descending
  var sortSecond = 5;
  var sortSecondAsc = true;
  var sortThird = 6;
  var sortThirdAsc = false;
  var sortFourth = 3;
  var sortFourthAsc = true;


  var headerRows = 1; 
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Tmobile短期卡表');
  var data = sheet.getDataRange().getValues();
  
  var range = sheet.getRange(headerRows+1, 1, sheet.getMaxRows()-headerRows, sheet.getLastColumn());  //先整体排序
  range.sort([{column: sortFourth, ascending: sortFourthAsc}, {column: sortFirst, ascending: sortFirstAsc}, {column: sortThird, ascending: sortThirdAsc}]);

}

function sortActivationShort2() {    //给Tmobile短期卡表自动排序（按照当天以前的长度）
  var sortFirst = 4; //index of column to be sorted by; 1 = column A, 2 = column B, etc.
  var sortFirstAsc = true; //Set to false to sort descending
  var sortSecond = 5;
  var sortSecondAsc = true;
  var sortThird = 6;
  var sortThirdAsc = false;
  var sortFourth = 3;
  var sortFourthAsc = true;
  var todayRow = SearchStartDateShort();

  var headerRows = 1; 
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Tmobile短期卡表');
  var data = sheet.getDataRange().getValues();
  
  var range2 = sheet.getRange(headerRows+1, 1, todayRow, sheet.getLastColumn()); //再按照当天排序
  range2.sort([{column: sortFirst, ascending: sortFirstAsc}, {column: sortSecond, ascending: sortSecondAsc}, {column: sortThird, ascending: sortThirdAsc}]);
  
}


function convert2jsDateOLD( value ) {   //转换为日期格式
  var jsDate = new Date();  // default to now
  if (value) {
    // If we were given a date object, use it as-is
    if (typeof value === 'date') {
      jsDate = value;
    }
    else {
      if (typeof value === 'number') {
        // Assume this is spreadsheet "serial number" date
        var daysSince01Jan1900 = value;
        var daysSince01Jan1970 = daysSince01Jan1900 - 25569 // 25569 = days TO Unix Time Reference
        var msSince01Jan1970 = daysSince01Jan1970 * 24 * 60 * 60 * 1000; // Convert to numeric unix time
        var timezoneOffsetInMs = jsDate.getTimezoneOffset() * 60 * 1000;
        jsDate = new Date( msSince01Jan1970 + timezoneOffsetInMs );
      }
      else if (typeof value === 'string') {
        // Hope the string is formatted as a date string
        jsDate = new Date( value );
      }
    }
  }
  return jsDate;
}


function convert2jsDate( value ) {   //转换为日期格式
  if (value) {
    var jsDate = new Date();  // default to now
    // If we were given a date object, use it as-is
    if (typeof value === 'date') {
      jsDate = value;
    }
    else {
      if (typeof value === 'number') {
        // Assume this is spreadsheet "serial number" date
        var daysSince01Jan1900 = value;
        var daysSince01Jan1970 = daysSince01Jan1900 - 25569 // 25569 = days TO Unix Time Reference
        var msSince01Jan1970 = daysSince01Jan1970 * 24 * 60 * 60 * 1000; // Convert to numeric unix time
        var timezoneOffsetInMs = jsDate.getTimezoneOffset() * 60 * 1000;
        jsDate = new Date( msSince01Jan1970 + timezoneOffsetInMs );
      }
      else if (typeof value === 'string') {
        // Hope the string is formatted as a date string
        jsDate = new Date( value );
      }
    }
    return jsDate;
  }
}


function onEdit1() {   //activation表更改后则在表格显示需要同步。
 var s = SpreadsheetApp.getActiveSheet();
 if( s.getName() == "Activation" ) { //checks that we're on the correct sheet
   var r = s.getActiveCell();
   var rRow = r.getRowIndex();
   var rCol = r.getColumn();
   Logger.log(rCol);
   var firstCell = r.offset(0,1-rCol);
   firstCell.setValue('需要同步');
 };
}


function mysql_real_escape_string (str) {
    return str.replace(/[\0\x08\x09\x1a\n\r"'\\\%]/g, function (char) {
        switch (char) {
            case "\0":
                return "\\0";
            case "\x08":
                return "\\b";
            case "\x09":
                return "\\t";
            case "\x1a":
                return "\\z";
            case "\n":
                return "\\n";
            case "\r":
                return "\\r";
            case "\"":
            case "'":
            case "\\":
            case "%":
                return "\\"+char; // prepends a backslash to backslash, percent,
                                  // and double/single quotes
        }
    });
}


function sendSms2() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('sendSMS');
  //var sheet = SpreadsheetApp.getActiveSheet();
  var startRow = 2;  // First row of data to process
  var numRows = sheet.getLastRow();   // Number of rows to process
  // Fetch the range of cells A2:B3
  var dataRange = sheet.getRange(startRow, 1, numRows, 16)
  // Fetch values for each row in the Range.
  var data = dataRange.getValues();
  for (var i = 0; i < data.length; ++i) {
    var row = data[i];
    //var emailAddress = row[14];  // First column
    var emailAddress = toEmailCS1;
    //var formattedDate = Utilities.formatDate(row[0], "GMT", "yyyy-MM-dd");  // change date format
     //Logger.log(formattedDate);
    var smsSent = row[14];     // Third column
    if (smsSent != SMS_SENT && row[5] != '') {  // Prevents sending duplicates
      var message = row[4] + row[12]+ row[13];
      sheet.getRange(startRow + i, 15).setValue(SMS_SENT);
      toNumber = row[5];
      bodyMessage = message;
      sendSms(toNumber,bodyMessage, SID, Token, twilioNumber);
      Utilities.sleep(1100);
      SpreadsheetApp.flush();
      // Make sure the cell is updated right away in case thescript is interrupted
     
    }
  }
}


function checksiminsheet(){
  
  //var documentProperties = PropertiesService.getDocumentProperties();
 // var irow_ID = documentProperties.getProperty('row_ID');
  //if (irow_ID == ''){
  //  var documentProperties = PropertiesService.getDocumentProperties();
  //  documentProperties.setProperty('row_ID', '10000');
  //}
 // Logger.log(irow_ID);
  
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Activation');
  var cell = sheet.getRange('C1');
  var numRows = sheet.getLastRow();
  var row = numRows;
  
  var data = sheet.getDataRange().getValues();   //data 开卡表
  
  
  var sheet2 = ss.getSheetByName('Tmobile短期卡表');
  var cell2 = sheet2.getRange('A1');
  var numRows2 = sheet2.getLastRow();
  var row2 = numRows2;

  var data2 = sheet2.getDataRange().getValues();  //data2 短卡表
  
  var lasttime_row = data[0][28];
  //sheet.getRange(1,29).setValue(lasttime_row);
  
  
  
for (var i = lasttime_row; i < data.length; i++) { 
  //Logger.log(i);
var simnumber = data[i][8].toString().replace(/[^0-9]/ig, "");
  sheet.getRange(1,29).setValue(i);
  if (i == (numRows-2)) {
    sheet.getRange(1,29).setValue('2');
  }
  
  for (var x = 1; x < data2.length; x++) { 
    for (var y = 1; y < 45; y++) { 
      if (data2[x][y] != ''&& typeof data2[x][y] === 'string') {
       var simnumber2 = data2[x][y]; 

       if (simnumber2.length > 18 && simnumber2.length <21){
         simnumber2=simnumber2.toString().replace(/[^0-9]/ig, "")
         if (simnumber == simnumber2){
           sheet.getRange(i+1,2).setBackgroundRGB(255,20,147);
           Logger.log(simnumber2);
          
         }
       }
      }
    }
  }
  
   
}
  
  
  
  
   
   
   // if(data[i][0] != '处理终结' && data[i][0] != '格式错误' && simnumber.length > 18) {  //如果第一行是处理终结字样则不处理,sim卡为空也不处理,或者少于19位数字不处理
  
  
}



function sendSms(toNumber,bodyMessage, SID, Token, twilioNumber) {
 
      
   
    var url = "https://api.twilio.com/2010-04-01/Accounts/" +            // URL used to enter correct Twilio acct
      SID + "/SMS/Messages.json";                                           
    var options = {                                                      // Specify type of message             
      method: "post",                                                    // Post rather than Get since we are sending
      headers: {                                                         
        Authorization: "Basic " + 
        Utilities.base64Encode(SID + ":" + Token)
      },
      payload: {                                                         // SMS details
        From: twilioNumber,
        To: toNumber,
        Body: bodyMessage
      }
    };
    var response = UrlFetchApp.fetch(url, options);                   // Invokes the action
    Logger.log(response);
}


function Makecall(toNumber,bodyMessage, SID, Token, twilioNumber) {
//toNumber = row[13];
var toNumber = "62634771234";
var bodyMessage = "1111";
var SID = "AC5946b7c38e3a";
var Token = "5cb31281";
var twilioNumber = "9015099996"
      
    var url = 'https://api.twilio.com/2010-04-01/Accounts/'+SID+'/Calls.json';
  
    var options = {                                                      // Specify type of message             
      method: "post",                                                    // Post rather than Get since we are sending
      headers: {                                                         
        Authorization: "Basic " + 
        Utilities.base64Encode(SID + ":" + Token)
      },
      payload: {                                                         // SMS details
        From: twilioNumber,
        To: toNumber,
        "Url": url,
        Body: bodyMessage
      }
    };
    var response = UrlFetchApp.fetch(url, options);                   // Invokes the action
    Logger.log(response);
}


  
  
  
  
  













