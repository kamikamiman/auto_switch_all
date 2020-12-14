// ============================================================================================================ //
//       【関数】 サービス予定表から予定を取得                                                                           //
// ============================================================================================================ //

function ReadDataTest() {
  
  // ------------------------------------------------------------------------------------------------------------ //
  //       サービス予定表のシート情報                                                                                   //
  // ------------------------------------------------------------------------------------------------------------ //  
    
  // 本日の月のシート情報を取得
  const date   = new Date();                                             // 日付を取得
  const nowDay = Utilities.formatDate(date, 'Asia/Tokyo', 'M/d');        // 本日の日付のフォーマット
  const period = 69;                                                     // 第〇〇期
  const nowMonth = Utilities.formatDate(date, 'Asia/Tokyo', 'M');        // 本日の月を取得
  const schedule = ssGet.getSheetByName('${period}期${nowMonth}月'
                                      .replace('${period}', period)
                                      .replace('${nowMonth}', nowMonth));
  
  // 翌月の月のシート情報を取得
  const nextDate = new Date();                                           // 日付を取得
  nextDate.setDate(nextDate.getDate() + 1);                              // 明日の日付をセット
  const nextMonth = Utilities.formatDate(nextDate, 'Asia/Tokyo', 'M');   // 明日の月を取得
  const nextSchedule  = ssGet.getSheetByName('${period}期${nextMonth}月'
                                      .replace('${period}', period)
                                      .replace('${nextMonth}', nextMonth));
  
  // 日付を取得するセル範囲を指定
  const firstRow = 6;                                                    // セル選択開始行
  const lastCol  = schedule.getLastColumn();                             // セル選択終了列
  const _days = schedule.getRange(firstRow, 2, 1, lastCol -1);           // 日付情報の取得範囲
  const days = _days.getValues().flat();                                 // 日付情報を取得(配列)
  
  // 本日の日付のセルの列番号を取得
  let nowDayNum = 2;                                                     // 列番号(初期値)
  let dayNum;                                                            // 本日の日付の列番号
  days.forEach( getDay => {
     const day = Utilities.formatDate(getDay, 'Asia/Tokyo', 'M/d');
     if( day === nowDay ) dayNum = nowDayNum;
     nowDayNum += 1;
  });


  // 予定表のメンバーを取得(今月)
  const monLastCol = schedule.getRange('A:A').getLastRow();                              // 最終列番号
  const members = schedule.getRange(1, 1, monLastCol, 1).getValues().flat();             // メンバー情報

  // 予定表のメンバーを取得(来月)
  const nextMonLastCol = nextSchedule.getRange('A:A').getLastRow();                      // 最終列番号
  const nextMembers = nextSchedule.getRange(1, 1, nextMonLastCol, 1).getValues().flat(); // メンバー情報


  // ------------------------------------------------------------------------------------------------------------ //
  //       在席リストのシート(切替設定)情報                                                                              //
  // ------------------------------------------------------------------------------------------------------------ //  


  // 対象者を取得
  const SwitchSet  = ssSet.getSheetByName('切替設定');                                                 // シート(切替設定)
  const lastRow = SwitchSet.getRange(1, 2).getNextDataCell(SpreadsheetApp.Direction.DOWN).getRow();  // シートの最終行番号
  const infoRanges = SwitchSet.getRange(2, 2, lastRow-1, 4).getValues();                             // 切替設定情報を取得

  // 対象者を配列に追加
  infoRanges.forEach( infoRange => {
    targets.push(infoRange[0]);                               // 切替対象者を取得
    if ( infoRange[1] === "有効" ) starts.push(infoRange[0]);  // 出社時
    if ( infoRange[2] === "有効" ) ends.push(infoRange[0]);    // 退社時
    if ( infoRange[3] === "有効" ) details.push(infoRange[0]); // 予定欄
  });
  

//  console.log(targets, starts, ends, details);


  // ------------------------------------------------------------------------------------------------------------ //
  //       クラスでオブジェクトを作成                                                                                    //
  // ------------------------------------------------------------------------------------------------------------ //

  class MemberObj {
    
    constructor(name, rowNum, nextRowNum) {
      
      this.name = name;
      this.rowNum = rowNum;
      this.nextRowNum = nextRowNum;
      
    }
    
    
    // 本日の予定 ・ セルの背景色 を取得 [ メソッド ]
    getContents() {
      
      const selection = schedule.getRange(this.rowNum, dayNum, 1, 1); // 取得範囲
      const contents = selection.getValue();                          // 本日の予定
      const color = selection.getBackground();                        // セルの背景色
      
      this.contents = contents;
      this.color = color;
      
    };
    
    
    // 翌日の予定 ・ セルの背景色 を取得 [ メソッド ]
    getNextContents() {
      
      let nextContents;
      let selection     = schedule.getRange(this.rowNum, dayNum + 1, 1, 1);
      let nextSelection = nextSchedule.getRange(this.nextRowNum, 2, 1, 1);
      let nextColor;
      
      // 本日が月末以外だった場合
      if ( dayNum !== lastCol ) {
        nextContents = selection.getValue();       // 本日が月末以外、翌日の予定
        nextColor = selection.getBackground();         // セルの背景色
        
      // 本日が月末だった場合
      } else {
        nextContents = nextSelection.getValue();   // 本日が月末、翌月初日の予定
        nextColor = nextSelection.getBackground();     // セルの背景色
      }
      
      this.nextContents = nextContents;
      this.nextColor = nextColor;
      
    };
    
    
    // 自動切替設定の情報 を取得 [ メソッド ]
    getSwitchSet() {
    
      // シート(自動切替)の名前と一致した場合、オブジェクトに切替設定を追加
      infoRanges.forEach( infoRange => {
        if ( infoRange[0] === this.name ) {
          this.swStart  = infoRange[1];  // 出社時(有効・無効)
          this.swEnd    = infoRange[2];  // 退社時(有効・無効)
          this.swDetail = infoRange[3];  // 予定欄(有効・無効)
        }
      });   
    
    }
    
    
    
  };

  // ------------------------------------------------------------------------------------------------------------ //
  //      配列 [ memberObj ] にメンバー毎のオブジェクトを追加                                                              //
  // ------------------------------------------------------------------------------------------------------------ //
  
  // メンバーオブジェクトを格納する配列
  let membersObj = [];


  targets.forEach( (target) => {
    
    // サービス予定表よりメンバーの行番号を取得 ( 本日 ・ 翌日 )
    const rowNum  = members.indexOf(target) + 1;           // 行番号 （本日）
    const nextRowNum  = nextMembers.indexOf(target) + 1;   // 行番号 （翌日）

    // オブジェクト作成(予定表に名前があるメンバーのみ実行)
    if ( rowNum > 0 ) {
      const obj = new MemberObj(target, rowNum, nextRowNum); // オブジェクト{obj}作成
      obj.getContents();                                     // {obj} に本日の予定を追加
      obj.getNextContents();                                 // {obj} に翌日の予定を追加
      obj.getSwitchSet();                                    // {obj} に切替設定を追加

      // {obj}を[memberObj]に追加
      membersObj.push(obj);
    }

//  console.log(obj); // ログ確認用(メンバー毎のオブジェクト)

  });
  
//  console.log(membersObj); // ログ確認用(メンバー毎のオブジェクト)

  return membersObj;       // 配列を返す 
  
}
