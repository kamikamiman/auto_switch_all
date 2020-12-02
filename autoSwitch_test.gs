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


 [ AutoSwitch() 内の関数の説明 ] 
   ・ ReadData(member) で指定した member の当日 ・ 翌日 の予定を取得する。
     * return で totalContents を返す。
   ・ WhiteData(member, totalContents) で指定した member の在席状態を変更する。
     * totalContents は、member の 当日 ・ 翌日 の予定が配列で格納されている。


*/

/******************************************************/
/***   指定したメンバーの予定を取得し、在席リストに書込む   ***/
/******************************************************/
function AutoSwitch() {
  
  const member = ["上倉健太", "後藤　勉"];        // メンバーを指定
  const membersObj = ReadDataTest(member);     // 当日、翌日の予定
  WhiteDataTest(...membersObj);                // 取得した予定を在席リストに書込

}



/**************************************************/
/***   指定した時間にスクリプトを実行するトリガー設定   ***/
/**************************************************/




//function AutoSwitchTest() {
//  
//    const time = new Date();
//  time.setHours(8);
//  time.setMinutes(25);
//  ScriptApp.newTrigger('AutoSwitch').timeBased().at(time).create();
//
//}
//
//// プロジェクトトリガーで実行
//function mtgStartTrigger(){
//  
//  const time = new Date();
//  time.setHours(10);
//  time.setMinutes(01);
//  ScriptApp.newTrigger('AutoSwitch').timeBased().at(time).create();
//
//}
//
//// プロジェクトトリガーで実行
//function mtgEndTrigger(){
//  
//  const time = new Date();
//  time.setHours(11);
//  time.setMinutes(01);
//  ScriptApp.newTrigger('AutoSwitch').timeBased().at(time).create();
//
//}
//
//// プロジェクトトリガーで実行
//function endTrigger(){
//
//  const time = new Date();
//  time.setHours(17);
//  time.setMinutes(30);
//  ScriptApp.newTrigger('AutoSwitch').timeBased().at(time).create();
//
//  
//}
