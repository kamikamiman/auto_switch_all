/*
  [ プログラム説明 ]

① プロジェクトトリガーで startTrigger() を実行する。 → 8:25 に AutoSwitch() が実行される。
   当日の在席状態を自動で変更する。
  
② プロジェクトトリガーで endTrigger() を実行する。  → 17:30 に AutoSwitch() が実行される。
   当日・翌日の在席状態を自動で変更する。

・ 当日・翌日の情報を取得し、それ以降の予定は取得しない。
・ 予定表の記入方法
  外出 : 外出の文字が含まれるように記載する。
  出張 : 出張の文字が含まれるように記載する。
  休み : 休み・有給・振休・代休・RHのいずれかの文字が含まれるように記載する。
  
  

*/

  // スプレットシートを取得（(ここにスプレットシートのアドレスを記入)）
  const ssGet = SpreadsheetApp.openById('1WY8sAykoyiu1bbGglSWuSmGpLyQwrwXTAYwZzvK0oR4'); // サービス作業予定表
  const ssSet = SpreadsheetApp.openById('1Itid9HCrW0wy_ATM4lqBDzkQs64DpemrWL4THmOsEIg'); // 在席リスト

  // 対象者を格納する配列( 切替対象者, 出社時, 退社時, 予定欄 )
  const targets = []; // 切替対象者
  const starts  = []; // 出社時
  const ends    = []; // 退社時
  const details = []; // 予定欄


/******************************************************/
/***   指定したメンバーの予定を取得し、在席リストに書込む       ***/
/******************************************************/
function AutoSwitchTest() {
  
  const membersObj = ReadDataTest(); // 当日、翌日の予定
  
  WriteDataTest(...membersObj);      // 取得した予定を在席リストに書込

}



/**************************************************/
/***   指定した時間にスクリプトを実行するトリガー設定       ***/
/**************************************************/

// プロジェクトトリガーで実行(出社時)
function startTriggerTest() {
  
  const time = new Date();
  time.setHours(8);
  time.setMinutes(25);
  ScriptApp.newTrigger('AutoSwitchTest').timeBased().at(time).create();

}


// プロジェクトトリガーで実行(退社時)
function endTriggerTest(){

  const time = new Date();
  time.setHours(17);
  time.setMinutes(30);
  ScriptApp.newTrigger('AutoSwitchTest').timeBased().at(time).create();
 
}








