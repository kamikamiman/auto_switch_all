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

  // ここに対象者を記入
  const targets    = [ "上倉健太", "後藤　勉" ]; // 自動切換
  const starts     = [ "上倉健太", "後藤　勉" ]; // 始業切替
  const ends       = [ "上倉健太", "後藤　勉" ]; // 終業切替
  const mtgStarts  = [ "上倉健太", "後藤　勉" ]; // 会議開始切替
  const mtgEnds    = [ "上倉健太", "後藤　勉" ]; // 会議終了切替
  const details    = [ "上倉健太", "後藤　勉" ]; // 詳細書込


/******************************************************/
/***   指定したメンバーの予定を取得し、在席リストに書込む       ***/
/******************************************************/
function AutoSwitchTest() {
  
  const membersObj = ReadDataTest(); // 当日、翌日の予定
  
  WhiteDataTest(...membersObj);      // 取得した予定を在席リストに書込

}



/**************************************************/
/***   指定した時間にスクリプトを実行するトリガー設定       ***/
/**************************************************/

function startTriggerTest() {
  
  const time = new Date();
  time.setHours(8);
  time.setMinutes(25);
  ScriptApp.newTrigger('AutoSwitchTest').timeBased().at(time).create();

}

// プロジェクトトリガーで実行
function mtgStartTriggerTest(){
  
  const time = new Date();
  time.setHours(10);
  time.setMinutes(01);
  ScriptApp.newTrigger('AutoSwitchTest').timeBased().at(time).create();

}

// プロジェクトトリガーで実行
function mtgEndTriggerTest(){
  
  const time = new Date();
  time.setHours(11);
  time.setMinutes(01);
  ScriptApp.newTrigger('AutoSwitchTest').timeBased().at(time).create();

}

// プロジェクトトリガーで実行
function endTriggerTest(){

  const time = new Date();
  time.setHours(17);
  time.setMinutes(30);
  ScriptApp.newTrigger('AutoSwitchTest').timeBased().at(time).create();

  
}
