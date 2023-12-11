const RootSs =              SpreadsheetApp.getActiveSpreadsheet().getActiveSheet(); // 管理統合シート
const MozerSs =             SpreadsheetApp.openById('1FidBw6sJEgwo24e85yDB9OGAXOZIKlbXEchkgRB5SPw'); // unity
const SlideTemplate_Base =  DriveApp.getFileById('1BRccrWp-uW1Fe0vBkNqOCP1sXKkqRXmtrHseZJsmXv8');    // MOZER_Base制作用テンプレート_2020.2.3
const SlideFolder_Base =    DriveApp.getFolderById('1Spt4y7BUduPmF_mY67-9wPuLDE5W4l83');             // 作業スライド_Base
const SlideTemplate_Tips =  DriveApp.getFileById('12Tn9OlhQamydJ3q3oUiGDrne1hpjt0ZMkRV2VPAtKTk');    // MOZER_Tips制作用テンプレート_2020.2.3
const SlideFolder_Tips =    DriveApp.getFolderById('1aHP8tm70OGsKDWIUUWXt62aidSa-g0q4');             // 作業スライド_Tips
const Design_ClearSs =      SpreadsheetApp.openById('1rotD13DWPAWtHFD5o5d-bdjBjMPmPdr3PKXYFviL6ns'); // 【Camp】MOZER教材設計（清書）
const ReviewSs =            SpreadsheetApp.openById('1wl-nEUZN63XG-f8lr5pBTN9E1Rz6NO73n6DzKH4VVHU'); // 【Camp】MOZER教材レビュー記録（レビューイ用）
const Design_HoursSs =      SpreadsheetApp.openById('1AfUdIYxuvoyhTxorBfArQSXZ5jxTcR4GlF1ak816ixs'); // 【Camp】MOZER教材設計工数
const Product_HoursSs =     SpreadsheetApp.openById('1_R_FXqlo4zBNVmK_DoaeLu5hVoIoWYngANJxxpCM8oY'); // 【Camp】MOZER教材制作工数
const StrCntSs =            SpreadsheetApp.openById('1tMZVjM33_iCSOfUDtfUbC04pKeIT9n8bjrTAusAyc1o'); // 【Camp】Unity教材文字列数えますくん

const COL_PROJECTTYPE =           4;
const COL_DESIGNER =              8;
const COL_PRODUCER =              9;
const COL_MAINTITLE =             14;
const COL_SUBTITLE =              15;
const COL_SKILL =                 16;
const COL_PROJECTNAME =           31;
const COL_MOZERSHEET_URL =        32;
const COL_IMAGESLIDE_URL =        34;
const COL_DESIGNSHEET_URL =       35;
const COL_DESIGNHOURSSHEET_URL =  36;
const COL_PRODUCTHOURSSHEET_URL = 37;
const COL_REVIEWSHEET_URL =       38;


function onOpen() {
  SpreadsheetApp.getUi().createMenu('教材関連資料')
    .addItem('新規作成', 'showInputForm')
    .addItem('削除',"deleteProject")
    .addItem('名前変更','changeFileName')
    .addToUi();
}

function changeFileName(){
  var projectName = Browser.inputBox("現在の教材の名前を入力してください");
  var newprojectName = Browser.inputBox("教材の新しい名前を入力してください");
  var rowData = fetchRowWithKey(projectName, COL_PROJECTNAME);
  
  if(!projectName || projectName == "cancel") return;
  if(!newprojectName || newprojectName == "cancel") return;
  if(!rowData) return;

  RootSs.getRange(rowData[0], COL_PROJECTNAME).setValue(newprojectName);
  
  showLoadingDialog();

  var sheets = [
        MozerSs,
        Design_ClearSs,
        ReviewSs,
        Design_HoursSs,
        Product_HoursSs,
        StrCntSs
  ];

  sheets.forEach((value) => {
    var targetSheet = value.getSheetByName(projectName);
    if(targetSheet != null && targetSheet != undefined){
      showLoadingDialog("名前変更中" + value.getName());
      targetSheet.setName(newprojectName);
      if(value.getSheetByName("一覧") !=null){
        for(var i = 2;i <= value.getSheetByName('一覧').getLastRow();i++){
          if(value.getSheetByName('一覧').getRange(i,2).getValue() == projectName){
            value.getSheetByName('一覧').getRange(i,2).setValue(newprojectName);
          }
        }
      }
      
    }
  });

  showLoadingDialog("スライドを探しています");
  var files = SlideFolder_Base.getFilesByName(projectName);
  while(files.hasNext()){
    var target = files.next();
    if(target.getName() == projectName){
      target.setName(newprojectName);
    }
  }
  var files = SlideFolder_Tips.getFilesByName(projectName);
  while(files.hasNext()){
    var target = files.next();
    if(target.getName() == projectName){
      target.setName(newprojectName);
    }
  }
  Browser.msgBox("名前の変更が完了しました")

}

/**
 * データ入力フォームを表示する
 */
function showInputForm(){
  var html = HtmlService.createTemplateFromFile('DataForm_Start');
  var htmlOutput = html.evaluate().setWidth(500).setHeight(800);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput,"データ入力フォーム 1/2");
}

/**
 * スタートフォームの入力を受け取る
 */
function inputTxType(form){
  var rowNum = form.rowNum;
  var projectType = form.projectType;
  if(rowNum == "" || projectType == "") return;

  if(isExsitDataRow(rowNum)){
    Browser.msgBox("指定行にデータが存在します！\\n行を変更して再度実行してください");
    return;
  }

  // 次のフォーム表示
  var html;
  var title = "";
  if(projectType == "Tips"){ // Tips作成フォーム表示
    html = HtmlService.createTemplateFromFile('DataForm_Tips');
    title = "データ入力フォーム【Tips】 2/2"
  }else if(projectType == "Base_Root"){ // Base_Root作成フォーム表示
    html = HtmlService.createTemplateFromFile('DataForm_Base');
    title = "データ入力フォーム【Base01番】 2/2"
    html.type = "first"
  }else if(projectType == "Base_Section"){ // "Base_Section作成フォーム表示"
    html = HtmlService.createTemplateFromFile('DataForm_Base');
    title = "データ入力フォーム【Base02番以降】 2/2"
    html.type = "other"
  }
  html.rowNum = rowNum;
  var htmlOutput = html.evaluate().setWidth(600).setHeight(800);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput,title);
}

/**
 * Tips作成フォームの入力を受け取る
 */
function inputTipsInfo(form){
  var rowNum = Number.parseInt(form.rowNum);
  var projectName = String(form.projectName); // 必須
  var mainTitle = String(form.mainTitle);
  var subTitle = String(form.subTitle); // 必須
  var designer = String(form.designer); // 必須
  var producer = String(form.producer); // 必須
  var skillText = String(form.skillText);

  // データチェック部
  if(projectName == "" || subTitle == "" || designer == "" || producer == "") return;
  if(Browser.msgBox(
    "以下の内容で間違いないですか？\\n"+
    "挿入行番号：" + rowNum + "\\n" +
    "管理上の表記名（タブ / ファイル名）：" + projectName + "\\n" +
    "主題教材名（任意）：" + mainTitle + "\\n" +
    "副題教材名：" + subTitle + "\\n" +
    "設計担当者名：" + designer + "\\n" +
    "制作担当者名：" + producer + "\\n" +
    "スキル内容：" + skillText
    ,Browser.Buttons.OK_CANCEL) == "cancel") return;

    // おまけ
    showLoadingDialog();

    var existFiles = getExistShName(projectName);
    if(existFiles != ""){
      Browser.msgBox(
        "以下のスプレッドシートに同名のシートが存在しています\\n" +
        "シートを削除するかプロジェクト名を変更して再実行してください\\n" +
        existFiles );
      return;
    }
  if(true){ // シート生成部***********************************************************************************************************************
    var rootSh = RootSs;
    rootSh.getRange(rowNum, COL_PROJECTTYPE ).setValue("Tips");
    rootSh.getRange(rowNum, COL_DESIGNER    ).setValue(designer);
    rootSh.getRange(rowNum, COL_PRODUCER    ).setValue(producer);
    rootSh.getRange(rowNum, COL_MAINTITLE   ).setValue(mainTitle);
    rootSh.getRange(rowNum, COL_SUBTITLE    ).setValue(subTitle);
    rootSh.getRange(rowNum, COL_SKILL       ).setValue(skillText);
    rootSh.getRange(rowNum, COL_PROJECTNAME ).setValue(projectName);

    showLoadingDialog("unityシートを作成しています...");
    // mozerシート:unity
    var mozerSheet = MozerSs.getSheetByName("template_3D").copyTo(MozerSs);
    mozerSheet.setName(projectName);
    rootSh.getRange(rowNum, COL_MOZERSHEET_URL).setValue('=hyperlink("'+ MozerSs.getUrl() + '#gid=' + mozerSheet.getSheetId() + '")')
    .setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);

    // // 設計書シート:【Camp】MOZER教材設計（清書）
    createSheetFromSs(projectName,Design_ClearSs,true,rowNum, COL_DESIGNSHEET_URL);

    // 設計工数シート:【Camp】MOZER教材設計工数
    createSheetFromSs(projectName,Design_HoursSs,true,rowNum, COL_DESIGNHOURSSHEET_URL);

    // 制作工数シート:【Camp】MOZER教材制作工数
    createSheetFromSs(projectName,Product_HoursSs,true,rowNum, COL_PRODUCTHOURSSHEET_URL);

    // レビューシート：【Camp】MOZER教材レビュー記録（レビューイ用）
    createSheetFromSs(projectName,ReviewSs,true,rowNum, COL_REVIEWSHEET_URL);

    // Unity教材数えますくん：【Camp】Unity教材文字列数えますくん
    createSheetFromSs(projectName,StrCntSs,true);
    
    // スライド作成処理*********************************************************************
    rootSh.getRange(rowNum, COL_IMAGESLIDE_URL).setValue(SlideTemplate_Tips.makeCopy(SlideFolder_Tips).setName(projectName).getUrl())
    .setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);

  }

  Browser.msgBox("実行が終了しました");

}

/**
 * Base作成フォームの入力を受け取る
 */
function inputBaseInfo(form){
  var rowNum = Number.parseInt(form.rowNum);
  var projectName = form.projectName; // 必須
  var mainTitle = form.mainTitle; // 必須
  var subTitle = form.subTitle; // 必須
  var designer = form.designer; // 必須
  var producer = form.producer; // 必須
  var skillText = form.skillText; // 任意

  // データチェック部
  if(projectName == "" || subTitle == "" || designer == "" || producer == "" || mainTitle == "") return;
  if(Browser.msgBox(
    "以下の内容で間違いありませんか？\\n"+
    "挿入行番号：" + rowNum + "\\n" +
    "管理上の表記名（タブ / ファイル名）：" + projectName + "\\n" +
    "主題教材名（任意）：" + mainTitle + "\\n" +
    "副題教材名：" + subTitle + "\\n" +
    "設計担当者名：" + designer + "\\n" +
    "制作担当者名：" + producer + "\\n" +
    "スキル内容：" + skillText
    ,Browser.Buttons.OK_CANCEL) == "cancel") return;

    // おまけ
    showLoadingDialog("");

    var existFiles = getExistShName(projectName);
    if(existFiles != ""){
      Browser.msgBox(
        "以下のスプレッドシートに同名のシートが存在しています\\n"+
      + "シートを削除するかプロジェクト名を変更して再実行してください\\n"
      + existFiles );
      return;
    }

   if(true){ // シート生成部***********************************************************************************************************************
    var rootSh = SpreadsheetApp.getActiveSheet();
    rootSh.getRange(rowNum, COL_PROJECTTYPE ).setValue("Base");
    rootSh.getRange(rowNum, COL_DESIGNER    ).setValue(designer);
    rootSh.getRange(rowNum, COL_PRODUCER    ).setValue(producer);
    rootSh.getRange(rowNum, COL_MAINTITLE   ).setValue(mainTitle);
    rootSh.getRange(rowNum, COL_SUBTITLE    ).setValue(subTitle);
    rootSh.getRange(rowNum, COL_SKILL       ).setValue(skillText);
    rootSh.getRange(rowNum, COL_PROJECTNAME ).setValue(projectName);

    // mozerシート:unity
    showLoadingDialog("unityシートを作成しています...");
    var mozerSheet = MozerSs.getSheetByName("template_3D").copyTo(MozerSs);
    mozerSheet.setName(projectName);
    rootSh.getRange(rowNum, COL_MOZERSHEET_URL).setValue('=hyperlink("'+ MozerSs.getUrl() + '#gid=' + mozerSheet.getSheetId() + '")')
    .setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);

    // // 設計書シート:【Camp】MOZER教材設計（清書）
    createSheetFromSs(projectName,Design_ClearSs,true,rowNum, COL_DESIGNSHEET_URL);

    if(form.type == "first"){ // 01番のみ作成
      // 設計工数シート:【Camp】MOZER教材設計工数
      createSheetFromSs(projectName,Design_HoursSs,true,rowNum, COL_DESIGNHOURSSHEET_URL);

      // 制作工数シート:【Camp】MOZER教材制作工数
      createSheetFromSs(projectName,Product_HoursSs,true,rowNum, COL_PRODUCTHOURSSHEET_URL);
    }

    // レビューシート：【Camp】MOZER教材レビュー記録（レビューイ用）
    createSheetFromSs(projectName,ReviewSs,true,rowNum, COL_REVIEWSHEET_URL);

    // Unity教材数えますくん：【Camp】Unity教材文字列数えますくん
    createSheetFromSs(projectName,StrCntSs,true);
    
    // スライド作成処理*********************************************************************
    showLoadingDialog("スライドを作成しています...");
    rootSh.getRange(rowNum, COL_IMAGESLIDE_URL).setValue(SlideTemplate_Base.makeCopy(SlideFolder_Base).setName(projectName).getUrl())
    .setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);

  }

  Browser.msgBox("実行が終了しました");

}

/**
 * 対象のスプレッドシートのテンプレートからシートを作成する
 * @param {string} projectName プロジェクト名
 * @param {SpreadsheetApp.Spreadsheet} ss シートを作成したいスプレッドシート
 * @param {boolean} isAppendSheet 一覧シートに記入するか
 * @param {number} rowNum URLを記入したい定義シートの行番号
 * @param {number} clmNum URLを記入したい定義シートの列番号
 * @return {SpreadsheetApp.Sheet} 作成したシート
 */
function createSheetFromSs(projectName,ss,isAppendSheet,rowNum = 0,clmNum = 0){
    showLoadingDialog(ss.getName()+"\n新規シートを作成しています");
    var template = ss.getSheetByName('テンプレート').copyTo(ss);
    template.setName(projectName).setTabColor(null);
    if(isAppendSheet){
      var itiranSheet = ss.getSheetByName("一覧");
      itiranSheet.getRange(itiranSheet.getLastRow()+1,2).setValue(projectName)
      .offset(0,1).setValue('=hyperlink("#gid=' + template.getSheetId() + '",B' + itiranSheet.getLastRow() + ')');
    }
    if(rowNum!=0 && clmNum!=0){
      var teigiSheet = SpreadsheetApp.getActiveSheet();
      teigiSheet.getRange(rowNum,clmNum).setValue(ss.getUrl() + '#gid=' + template.getSheetId())
      .setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
    }
    return template;
}

function showLoadingDialog(message){
  var html = HtmlService.createTemplateFromFile('Dialog_Loading');
  html.text = message;
  SpreadsheetApp.getUi().showModalDialog(html.evaluate().setHeight(400),"処理中です...");
}

function showErrorDialog(title,text){
  var html = HtmlService.createTemplateFromFile('Dialog_Error');
  html.title = title;
  html.text = text;
  SpreadsheetApp.getUi().showModalDialog(html.evaluate().setHeight(400),"処理中に例外が発生しました");
}

// 指定したシートの内, col列 で value（値）に一致する一行の全ての情報を読み込むための関数
function fetchRowWithKey(value, col){
  //  col列 の値が送られてきた value と一致する行を抜き出す
  var rowData = RootSs.getDataRange().getValues();
  for (var i = 0; i < rowData.length; i++) {
    if (rowData[i][col-1] == value) {
      rowData[i].unshift(i+1);
      return rowData[i];
    }
  }
  return null;
}

/**
 * 対象の行のデータの存在をチェック
 * 返り値：true = データが存在する
 * 返り値：false = データが存在しない
 */
function isExsitDataRow(row){
  var tmp = SpreadsheetApp.getActiveSheet().getRange(row+1, 1).getValue();
  SpreadsheetApp.getActiveSheet().getRange(row+1, 1).setFormula("=counta(AF" + row + ":AL" + row + ")");
  if(SpreadsheetApp.getActiveSheet().getRange(row+1, 1).getValue() != 0){
    SpreadsheetApp.getActiveSheet().getRange(row+1, 1).setValue(tmp);
    return true;
  }
  SpreadsheetApp.getActiveSheet().getRange(row+1, 1).setValue(tmp);
  return false;
}

/**
 * 同名のシートが存在するかチェック
 * @return {string} returnText 同名のシートが存在するスプレッドシート名
 */
function getExistShName(projectName){
  var returnText = "";
  var sheets = [
        MozerSs,
        Design_ClearSs,
        ReviewSs,
        Design_HoursSs,
        Product_HoursSs,
        StrCntSs
  ];
  sheets.forEach((value) => {
      var targetSheet = value.getSheetByName(projectName);
      if(targetSheet != null && targetSheet != undefined){
        returnText += value.getName() + "\\n";
      }
    });

  return returnText;
}

function deleteProject(){
  var projectName = Browser.inputBox("プロジェクト名を入力してください");
  if(!projectName || projectName == "cancel") return;
  var rowData = fetchRowWithKey(projectName, COL_PROJECTNAME);
  if(!rowData) return;

  showLoadingDialog();

  var sheets = [
        MozerSs,
        Design_ClearSs,
        ReviewSs,
        Design_HoursSs,
        Product_HoursSs,
        StrCntSs
  ];

  sheets.forEach((value) => {
    var targetSheet = value.getSheetByName(projectName);
    if(targetSheet != null && targetSheet != undefined){
      showLoadingDialog("削除中" + value.getName());
      value.deleteSheet(targetSheet);
      if(value.getSheetByName("一覧") !=null){
        for(var i = 2;i <= value.getSheetByName('一覧').getLastRow();i++){
          if(value.getSheetByName('一覧').getRange(i,2).getValue() == projectName){
            value.getSheetByName('一覧').getRange(i,2).clearContent()
            .offset(0,1).clearContent();
          }
        }
      }
      
    }
  });
  showLoadingDialog("スライドを探しています");
  var files = SlideFolder_Base.getFilesByName(projectName);
  while(files.hasNext()){
    var target = files.next();
    if(target.getName() == projectName){
      target.setTrashed(true);
    }
  }
  var files = SlideFolder_Tips.getFilesByName(projectName);
  while(files.hasNext()){
    var target = files.next();
    if(target.getName() == projectName){
      target.setTrashed(true);
    }
  }
  RootSs.getRange(rowData[0],32,1,7).clear();
  Browser.msgBox("削除が完了しました")
}
