  // 対象者を格納する配列( 自動切替のシートから取得 関数readDataより ) * 新たに切替を追加する場合はここに追加！！
  const targets = [];    // 切替対象者
  const starts  = [];    // 出社時
  const ends    = [];    // 退社時
  const details = [];    // 予定欄
  const attendSels = []; // 在席時の表示選択(東館選択時に格納, 通常は在席表示)

//  =============  以下に記入された以外の文字の場合は、設定が反映されませんので注意!  =============  //

  // 休日への切替判定を設定 (サービス予定表に設定文字の記入があるか)
  const offDays = [ '休み', '有給', '有休', '振休', '振替', '代休', 'RH', 'ＲＨ', 'フレ休み', 'フレ休', '完フレ', '完ﾌﾚ', '完全フレ', '完全ﾌﾚ', '完全フレックス', '完全ﾌﾚｯｸｽ' ];

  // 完全フレックス判定 ( 通常フレックスとの判別 [フレ]を含む場合、予定詳細にフレックス時間まで書き込まれてしまう対策 )
  const fullFlex = [ 'フレ休み', 'フレ休', '完フレ', '完ﾌﾚ', '完全フレ', '完全ﾌﾚ', '完全フレックス', '完全ﾌﾚｯｸｽ' ];
  
  // フレックスパターンを設定
  const flexPatterns = [ 'フレ', 'ﾌﾚ' ];

  // 在席リストの状態・予定詳細内容の書込先 (在席リストの列のセルが追加・削除された場合、ここの列番号も変更も必要です!) 
  let posiL =  5; // memLeftの書込先の列番号
  let posiC = 10; // memCenterの書込先の列番号
  let posiR = 15; // memRightの書込み先の列番号

/******************************************************/
/***   指定したメンバーの予定を取得し、在席リストに書込む   ***/
/******************************************************/
function AutoSwitchTest() {
  
  // 在席状態・予定詳細を取得
  const object = ReadDataTest();

  // 取得した情報を分割
  const membersObj  = object[0];
  const nightRowCol = object[1];
  const satRowCol   = object[2];
  
  // 取得した情報から在席状態・予定を書込
  WriteDataTest(nightRowCol, satRowCol, ...membersObj);
}

/**************************************************/
/***   指定した時間にスクリプトを実行するトリガー設定   ***/
/**************************************************/
// プロジェクトトリガーで実行
function StartTrigger() {

  // 自動切替の設定時間を設定
  let setTimes = [ "0:00",  "0:30",  "1:00",  "1:30",  "2:00",  "2:30",  "3:00",  "3:30",  "4:00",
                   "4:30",  "5:00",  "5:30",  "6:00",  "6:30",  "7:00",  "7:30",  "8:00",  "8:25",
                   "9:00",  "9:30", "10:00", "10:30", "11:00", "11:30", "12:00", "12:30", "13:00", 
                  "13:30", "14:00", "14:30", "15:00", "15:30", "16:00", "16:30", "17:00", "17:15", 
                  "18:00", "18:30", "19:00", "19:30", "20:00", "20:30", "21:00", "21:30", "22:00", 
                  "22:30", "23:00", "23:30" ];
    
  /* 関数[delTrigger]を実行する時間帯を設定 ( プロジェクトトリガー(AutoSwitch)を削除 )         //
  // 使用済みのプロジェクトトリガー(AutoSwitch)を削除、新たにプロジェクトトリガー(AutoSwitch)を作成。 //
  // setTimesで指定した時間帯に実行。                                                  */
  
  let delTimes = [ "1*00", "3*00", "6*00", "10*30", "12*30", "15*00", "18*00", "19*00", "21*30", "23*00" ];
  
  // フレックス実行時間とトリガー削除時間を配列に格納 ( 要素の順番に注意。 不要なトリガーを削除 >>> 新たにトリガーを作成 )
  const setDelTimes = [ delTimes, setTimes ];
  
  
  // 変数の初期化
  let i = 0;
  let j = 0;
  
  // 現在の時間とフレックス時間を比較
  setDelTimes.forEach( el => {
    el.forEach( el2 => {
                      
      console.log(`el2: ${el2}`);
      // 設定時間を取得
      let setTime;
      const colonJudge    = el2.indexOf(":") !== -1;
      const asteriskJudge = el2.indexOf("*") !== -1;
      if ( colonJudge )    setTime = el2.split(":");   // 時間を分割 [:]の場合
      if ( asteriskJudge ) setTime = el2.split("*");   // 時間を分割 [*]の場合
      setTime[0] = Number(setTime[0]);  // 時間
      setTime[1] = Number(setTime[1]);  // 分
  
      // 現在の時間を取得
      const time = new Date();                                       // 現在時間を取得
      const HHmm = Utilities.formatDate(date, 'Asia/Tokyo', 'H:m');  // 現在の時間・分をフォーマット
      const Hm = HHmm.split(":");                                    // 時間を分割
      const HH = Number(Hm[0]);                                      // 時間
      const mm = Number(Hm[1]);                                      // 分
      let delJudge;
      let setJudge;
  
      // 現在時刻と設定時刻が一致した時に実行
      if ( HH == setTime[0] ) {
        
          console.log("AAA");
          console.log(`setTime[0]: ${setTime[0]}`);
          console.log(`HH: ${HH + 1}`);
          console.log(`setTime[1]: ${setTime[1]}`);
          console.log(setTime[0] == HH + 1 && setTime[1] == 0);
      
        // 現在の分より設定の分が大きい時に実行
        if ( mm <= setTime[1] ) {
          console.log("BBB");
          console.log(`実行時間1 ${setTime[0]}:${setTime[1]}`);
        
          // 関数の実行時間を設定
          time.setHours(setTime[0]);
          time.setMinutes(setTime[1]);
        
          // 設定時間に指定した関数を実行
          delJudge = delTimes.indexOf(el2);
          console.log(`el2: ${el2}, delJudge: ${delJudge}`);
          if ( delJudge !== -1 && j === 0 ) {
            DeleteTrigger();
            j++;
            console.log("delTrigger実行(1)");
          }
          
          setJudge = setTimes.indexOf(el2);
          console.log(`el2: ${el2}, setJudge: ${setJudge}`);
          if ( setJudge !== -1 && i === 0 ) {
            ScriptApp.newTrigger('AutoSwitchTest').timeBased().at(time).create();
            i++;
            console.log("AutoSwitchTest実行(1)");
          }        
        }

      // 現在時刻より設定時刻が大きい場合に実行
      } else if ( ( setTime[0] == HH + 1 && setTime[1] < 30 ) || ( HH == 23 && setTime[0] == 23 && 30 <= setTime[1] ) ) {
        console.log("CCC");
        console.log(`実行時間2 ${setTime[0]}:${setTime[1]}`);
        // 関数の実行時間を設定
        time.setHours(setTime[0]);
        time.setMinutes(setTime[1]);
      
        // 設定時間に指定した関数を実行
        delJudge = delTimes.indexOf(el2);
        console.log(`el2: ${el2}, delJudge: ${delJudge}`);
        if ( delJudge !== -1 && j === 0 ) {
          DeleteTrigger();
          j++;
          console.log("delTrigger実行(2)");          
        }
        
        setJudge = setTimes.indexOf(el2);
        console.log(`el2: ${el2}, setJudge: ${setJudge}`);
        if ( setJudge !== -1 && i === 0 ) {
          ScriptApp.newTrigger('AutoSwitchTest').timeBased().at(time).create();
          i++;
          console.log("AutoSwitchTest実行(2)");
        }
      }
  
    });
  
  });
}

// プロジェクトトリガーで実行(不要なプロジェクトトリガーの削除用)
function DeleteTrigger() {
  const triggers = ScriptApp.getProjectTriggers();  // 現在設定されているトリガーを取得
  
  // 現在設定されているトリガーから指定したトリガーを全て削除
  for ( const trigger of triggers ){
    if ( trigger.getHandlerFunction() == "AutoSwitchTest" ) {
      ScriptApp.deleteTrigger(trigger);
    }
  }
  
}

// -------------------------------- // 
//    サービス作業予定表のシート情報     //
// -------------------------------- //
// スプレットシートを取得
const ssGet = SpreadsheetApp.openById('1WY8sAykoyiu1bbGglSWuSmGpLyQwrwXTAYwZzvK0oR4'); // サービス作業予定表 (テスト用ID)
    
// 本日の月のシート情報を取得
const date   = new Date();                                             // 日付を取得
const nowDay = Utilities.formatDate(date, 'Asia/Tokyo', 'M/d');        // 本日の日付のフォーマット
  let period = 69;                                                     // 第〇〇期
const nowMonth = Utilities.formatDate(date, 'Asia/Tokyo', 'M');        // 本日の月を取得
const schedule = ssGet.getSheetByName('${period}期${nowMonth}月'
                                      .replace('${period}', period)
                                      .replace('${nowMonth}', nowMonth));
  
// 翌月の月のシート情報を取得
const nextDate = new Date(date.setMonth(date.getMonth()+1));
const nextMonth = Utilities.formatDate(nextDate, 'Asia/Tokyo', 'M');
let nextSchedule;                // 翌月
const nextPeriod = period + 1;   // 来期

// 翌月が４月以外に実行
if ( nextMonth != 4 ) {
  nextSchedule  = ssGet.getSheetByName('${period}期${nextMonth}月'
                                    .replace('${period}', period)
                                    .replace('${nextMonth}', nextMonth));
// 翌月が４月なら実行
} else {
  nextSchedule  = ssGet.getSheetByName('${period}期${nextMonth}月'
                                    .replace('${period}', nextPeriod)
                                    .replace('${nextMonth}', nextMonth));
}
    

// サービス作業予定表の情報を取得 ( 昼休み当番・土曜当番・２４Hサービス・日付 )
const firstRow = 2;                                                    // セル選択開始行
const lastCol  = schedule.getLastColumn();                             // セル選択終了列
const _scheduleData = schedule.getRange(firstRow, 2, 5, lastCol -1);   // 情報の取得範囲
const scheduleData = _scheduleData.getValues();                        // 情報を取得(配列)
const getCellBackgrounds = _scheduleData.getBackgrounds();             // セルの背景色を取得(配列)

// 各情報を分割して取得
const lunchDuty    = scheduleData[0];      // 昼当番
const saturdayDuty = scheduleData[1];      // 土曜当番
const nightDuty    = scheduleData[2];      // ２４Ｈ当番
const nightArea    = scheduleData[3];      // ２４Ｈ管轄
const days         = scheduleData[4];      // 日付

// 各セルの背景色を分割して取得
const lunchBG     = getCellBackgrounds[0]; // 昼当番
const saturdayBG  = getCellBackgrounds[1]; // 土曜当番
const nightBG     = getCellBackgrounds[2]; // ２４Ｈ当番
const nightAreaBG = getCellBackgrounds[3]; // ２４Ｈ管轄
const daysBG      = getCellBackgrounds[4]; // 日付

  
// 本日の日付のセルの列番号を取得
let nowDayNum = 2;                                                     // 列番号(初期値)
let dayNum, dayOfWeek;                                                 // 本日の日付の列番号
days.forEach( getDay => {
   const day = Utilities.formatDate(getDay, 'Asia/Tokyo', 'M/d');
   if( day === nowDay ) {
     dayNum = nowDayNum - 1;                                           // 日付の番号
     arrDayNum = dayNum - 1;                                           // 日付の番号(配列用)
     dayOfWeek = new Date(getDay).getDay();                            // 曜日番号
   }

   nowDayNum += 1;
});

// 予定表のメンバーを取得(今月)
const monLastRow = schedule.getRange('A:A').getLastRow();                                               // 最終行番号
const monLastCol = schedule.getLastColumn() -1;                                                         // 最終列番号
const members = schedule.getRange(1, 1, monLastRow, 1).getValues().flat();                              // メンバー情報
const memSched = schedule.getRange(1, 1, monLastRow, monLastCol +1).getValues();                        // メンバー情報(予定も含む)
const memColor = schedule.getRange(1, 1, monLastRow, monLastCol +1).getBackgrounds();                   // セルの背景色

// 予定表のメンバーを取得(来月)
const nextMonLastRow = nextSchedule.getRange('A:A').getLastRow();                                       // 最終行番号
const nextMonLastCol = nextSchedule.getLastColumn() -1;                                                 // 最終列番号
const nextMembers    = nextSchedule.getRange(1, 1, nextMonLastRow, 1).getValues().flat();               // メンバー情報
const nextMemSched   = nextSchedule.getRange(1, 1, nextMonLastRow, nextMonLastCol +1).getValues();      // メンバー情報(予定も含む)
const nextMemColor   = nextSchedule.getRange(1, 1, nextMonLastRow, nextMonLastCol +1).getBackgrounds(); // セルの背景色

// -------------------------------- // 
//       在席リストのシート情報         //
// -------------------------------- //

const ssSet = SpreadsheetApp.openById('1Itid9HCrW0wy_ATM4lqBDzkQs64DpemrWL4THmOsEIg');                      // ここに在席リストのIDを記入(テスト用ID)
const attendList = ssSet.getSheetByName('当日在席(69期)');                                                    // 在席リスト
const offDayList = ssSet.getSheetByName('69期サービス土日休み');                                               // サービス土日休み
const offLastRow = offDayList.getRange(1, 1).getNextDataCell(SpreadsheetApp.Direction.DOWN).getRow();       // サービス土日休みの最終行番号
const isoOffDayList = ssSet.getSheetByName('isowa休日');                                                     // isowa休日
const isoOffLastRow = isoOffDayList.getRange(1, 1).getNextDataCell(SpreadsheetApp.Direction.DOWN).getRow(); // isowa休日の最終行番号
const isowaOffDays = isoOffDayList.getRange(1,  1, isoOffLastRow, 1).getValues().flat();                    // ISOWAの休日情報(GW等)を取得

// メンバーの休日パターンを取得(シート２)  ( シート左列のメンバー + シート2記入メンバー ) >>> 土日休み
const normalHolMems = offDayList.getRange(1, 1, offLastRow, 1).getValues().flat();                          // 土日休みのメンバーを取得(サービス)
