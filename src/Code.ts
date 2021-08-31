class Column {
  id: Number
  name: String
  type: String
  validation: String
  constructor(id, name, type, validation){
    this.id = id;
    this.name = name;
    this.type = type;
    this.validation = validation;
  }

  getValue(value){
    if (this.type=="number" || this.type=="bool"){
      if (value==null){
        return "";
      } else {
        return `${value.replace(/,/g, '').replace(/"/g, '""')}`;
      }
    } 
    else if (this.type=="date"){
      if (value!=null && value!=""){
        var dt = new Date("Dec 30, 1899 00:00:00");
        dt.setDate(dt.getDate() + Number(value));
        var dateStr = `"${dt.getFullYear()}-${dt.getMonth() + 1}-${dt.getDate()}"`;
        return dateStr;
      } else {
        if(value==null){
          return `""`;
        } else {
          return `"${value.replace(/"/g, '""')}"`;
        }
      }
    }
    else {
      if (value==null){
        return `""`;
      } else {
        return `"${value.replace(/"/g, '""')}"`;
      }
    }
  }

  isEmpty(value){
    if(value==""){
      return this.name + "は必須です。";
    } else {
      return "";
    }
  }
}

const regExp_requireNotNull = new RegExp("requireNotNull");
const regExp_requireStringSize = new RegExp("requireStringSize\\(([0-9]+)\\)");
const regExp_requiredNumericRange = new RegExp("requiredNumericRange\\(([0-9 ]+),([0-9 ]+)\\)");


class Columns {
  columns: Array<Column>
  constructor(columnSheet){

    var data = columnSheet.getDataRange().getDisplayValues();
    if (data.length > 1) {
      this.columns = [];
      data.forEach(row => {
        if(!isNaN(row[0])){
          var c = new Column(
            Number(row[0]),
            row[1],
            row[2],
            row[3]
          )
          //console.log(`${c.id} ${c.type} ${c.name}`)

          this.columns.push(c);
        }
      })
    }
  }

  //emptyチェック
  requireNotNull(value, name) {
    if(value==""){
      return name + "は必須です。\n";
    } else {
      return "";
    }
  }

  requireStringSize(value, size, name) {
    var result = ""
    if(value.length > Number(size)){
      result = name + "は" + size + "文字以内です。(" + value.length + ")\n";
    } else {
      result = "";
    }
    return result;
  }

  numericRange(value, lower, higher, name) {
    var result = "";
    if (value==""){
      result = "";
    } else if(value < lower || value > higher){
      result = name + "は範囲が" + lower + "～" + higher + "です。(" + value +")\n";
    } else {
      result = "";
    }
    return result;
  }

  validate(row, rowOffset){
    var result = ""
    for (var i = 0; i < row.length;i++){
      var c = this.getColumn(i)
      if (c!=null){
        var vs = c.validation.split(":");
        vs.forEach( v =>{
          if (v!=""){
            var r1 = String(v).match(regExp_requireNotNull);
            var r2 = String(v).match(regExp_requireStringSize);
            var r3 = String(v).match(regExp_requiredNumericRange);
            if(r1!=null){
              result = result + this.requireNotNull(`${row[i]}`, c.name);
            }
            if(r2!=null){
              result = result + this.requireStringSize(`${row[i]}`, r2[1] ,c.name);
            }
            if(r3!=null){
              result = result + this.numericRange(`${row[i]}`, r3[1], r3[2] ,c.name);
            }

            console.log(`${i} ${c.name} ${v} ${row[i]} : ${r1} ${r2} ${r3} >>> ${result}`);
          }
        })
      }
    }
    return result;

  }

  getColumn(id){
    var result = null;
    for (var i =0; i<this.columns.length;i++){
      if(this.columns[i].id==id){
        result = this.columns[i];
        break;
      }
    }
    return result;
  }

  getLine(row, columnOffset){
    var cols = [];
    for (var i = 0 ; i < this.columns.length; i++){
      var c = this.columns[i]
      //console.log(`${c.id} ${c.type} ${c.name}`)
      var location = c.id + columnOffset + 1;
      if(location>row.length){
        cols.push(c.getValue(null));
      } else {
        cols.push(c.getValue(row[location]));
      }
    }
    return cols;
  }

  isRecordEmpty(row) {
    var result = true
    row.forEach( col=> {
        result = result && (col==null)
        result = result && (col=="")
    })
    return result
  }

  convertRangeToCsvFile(sheet, rowOffset, columnOffset, _minRow: number, _maxRow: number) {
  try {
    let minRow = _minRow - 1
    let maxRow = _maxRow - 1
    var data = sheet.getDataRange().getDisplayValues();
    if (data.length > 1) {
      var rows = [];
      var rowIndex = 0;
      data.forEach(row => {

        var cols = this.getLine(row, columnOffset);

        //行スキップ
        if (rowIndex < minRow) {
          //範囲外スキップ
        } else if (maxRow < rowIndex) {
          //範囲外スキップ
        } else if (rowIndex > rowOffset){
          //先頭の必須が空の行は追加しない
          var headCol = cols[0];
          if(!this.isRecordEmpty(cols)) {
            console.log(`add ${rowIndex} ${cols.length} ${headCol}`);
            var line = cols.join(',');
            console.log(line);
            rows.push(line);
          } else {
          //範囲外スキップ
          }
        }
        rowIndex = rowIndex + 1;
      });

      return rows.join('\r\n');
    }
  } catch(err) {
    Logger.log(err);
    Browser.msgBox(err);
  }
}

}

var rowOffset = 3; //この行までヘッダ（0スタート）
var columnOffset = 2; //この列までヘッダ（0スタート）
var saveFolder = ""
var columnsSheet = "columns"
var omakeHtml = ""
var prefix = "" //ファイル名prefix

function adCheckTest(){
  Utilities.sleep(1000);
  var returnValue = "";

  //settings
  var ss = SpreadsheetApp.getActiveSpreadsheet(); 
  readConfig();
  console.log(`rowOffset: ${rowOffset}`);
  console.log(`columnOffset : ${columnOffset}`);

  //column読み込み
  var sheetColumn = ss.getSheetByName(columnsSheet);
  var cols = new Columns(sheetColumn)
  var value = ss.getActiveRange().getValues()[0]
  var result = cols.validate(value, 0);
  console.log(`result: ${result}`);
  return result;
}


/**
 * シート上の間違い指摘
 * @param value 求人票の1レコード分(D~BM)
 */
function record_check(value) {
//  Utilities.sleep(1000);
  var returnValue = "";

  //settings
  var ss = SpreadsheetApp.getActiveSpreadsheet(); 
  readConfig();
  console.log(`rowOffset: ${rowOffset}`);
  console.log(`columnOffset : ${columnOffset}`);

  //column読み込み
  var sheetColumn = ss.getSheetByName(columnsSheet);
  var cols = new Columns(sheetColumn)

  returnValue = cols.validate(value[0], 0);

  return returnValue;
}

//settingをLoad
function readConfig(){
  const ss = SpreadsheetApp.getActiveSpreadsheet(); 
  //「settings」から情報取得
  var settingSheetName = "settings"
  var settingSheet = ss.getSheetByName(settingSheetName)

  rowOffset = settingSheet.getRange("D2").getValue()
  columnOffset = settingSheet.getRange("D3").getValue()
  saveFolder = settingSheet.getRange("D4").getValue()
  columnsSheet = settingSheet.getRange("D5").getValue()
  omakeHtml = settingSheet.getRange("D6").getValue()
  prefix = settingSheet.getRange("D7").getValue()
}

function getOmakeHtml(){
  readConfig();
  return omakeHtml;
}

function getDateStr(dt){
  return `${dt.getFullYear()}-${dt.getMonth() + 1}-${dt.getDate()}_${dt.getHours()}_${dt.getMinutes()}_${dt.getSeconds()}`;
}

function onOpen(){
  SpreadsheetApp.getUi()
                .createMenu('extra')
                .addItem('CSVユーティリティ', 'dialog')
                .addToUi();
}

//ダウンロードダイアログ用
function dialog() {
  var html = HtmlService.createHtmlOutputFromFile('download');
  SpreadsheetApp.getUi().showSidebar(html);//showSidebar
}

function saveAsCSV():string {
  return saveAsCSVRange("1-65536")
}

//spreadsheetを保存
function saveAsCSVRange(rowRange):string {
  let ranges = rowRange.split('-')
  var rMin = 1
  var rMax = 65536
  if (ranges.length>=2){
    rMin = Number(ranges[0])
    rMax = Number(ranges[1])
  } else {
    rMin = Number(ranges[0])
    rMax = rMin
  }
  //settings
  var ss = SpreadsheetApp.getActiveSpreadsheet(); 
  readConfig();
  console.log(`rowOffset: ${rowOffset}`);
  console.log(`columnOffset : ${columnOffset}`);

  //column読み込み
  var sheetColumn = ss.getSheetByName(columnsSheet);
  var cols = new Columns(sheetColumn)

  //sheet
  var sheet = ss.getActiveSheet();

  // create a folder from the name of the spreadsheet
  var folder = DriveApp.getFolderById(saveFolder); 
  // append ".csv" extension to the sheet name
  var fileName = prefix + getDateStr(new Date()) + ".csv";
  // convert all available sheet data to csv format
  var csvFile = cols.convertRangeToCsvFile(sheet,rowOffset, columnOffset, rMin, rMax);
  // create a file in the Docs List with the given name and the csv data
  var file = folder.createFile(fileName, csvFile);
  //File downlaod
  var url = file.getUrl();
  console.log(`${file.getUrl()}`);
  return url;
}
