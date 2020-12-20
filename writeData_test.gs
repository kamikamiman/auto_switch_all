/* 機能追加予定

   ・ 翌日以降が連続して休日・出張・外出・在宅の場合、予定詳細の項目に開始日〜終了日を書込
     今月・翌月予定の配列を状態変換した配列にする。(休日・外出・出張・在宅等の表示に変換)  >>> 完了
     休日の予定期間の記入が数日の場合に書き込みがうまくいかない。                    >>> 完了
     複数の予定期間を外出・出張・在宅で確認。
     
     ・予定期間の月・日の不要な[0]を削除する。 例） 01/08 >> 1/8  >>> 完了
     ・外出時の書込がおかしい。(1日のみの場合)                  >>> 完了 
     
 
     ・その日付までは予定欄を更新しないようにする。 >>>このままで問題ない？
     ・writeData >>> readData にプログラムを移行させられる箇所を確認。



・ フレックスに対応する。





  本番前の注意点
  ・ 予定表と在席リストの漢字が違いエラーとなる事例あり。(福崎さん等)
  ・ 予定表に名前がないが在席リストに存在する人は予定表に追加する。
  ・ 平日休みの方の休日のセル背景色のパターンを統一したい。>>> 亀岡さんに確認




*/


// ============================================================================================================ //
//       【関数】 取得した予定を在席リストに書込                                                                          //
// ============================================================================================================ //

function WriteDataTest(...membersObj) {
  
   // スプレットシートを取得（データ書込み用）
  const attendList = ssSet.getSheetByName('当日在席(69期)');                                                // シート1
  const lastRow    = attendList.getRange('C:C').getLastRow();                                             // シート1の最終行番号
  const offDayList = ssSet.getSheetByName('69期サービス土日休み');                                            // シート2
  const offLastRow = offDayList.getRange(1, 1).getNextDataCell(SpreadsheetApp.Direction.DOWN).getRow();   // シート2の最終行番号
  const isoOffDayList = ssSet.getSheetByName('isowa休日');                                                     // シート3
  const isoOffLastRow = isoOffDayList.getRange(1, 1).getNextDataCell(SpreadsheetApp.Direction.DOWN).getRow(); // シート3の最終行番号

  
  // 現在の時間（△時）を取得
  const date = new Date();
  const nowTime  = Utilities.formatDate(date, 'Asia/Tokyo', 'H');         // 現在の時間
  const dayOfNum = date.getDay();                                         // 曜日番号
  date.setDate(date.getDate() + 1);                                       // 明日の日付をセット
  const nextDate = Utilities.formatDate(date, 'Asia/Tokyo', 'M/d');       // 明日の日付を取得
  
  // 在席リストのメンバー情報を取得
  const membersL   = attendList.getRange(1,  3, lastRow, 4).getValues();  // メンバー情報(在席リスト左)
  const membersC   = attendList.getRange(1,  8, lastRow, 4).getValues();  // メンバー情報(在席リスト中)
  const membersR   = attendList.getRange(1, 13, lastRow, 4).getValues();  // メンバー情報(在席リスト右)  
  const members = [ membersL, membersC, membersR ];                       // メンバー情報(全員分)

  members.forEach( member => member = member.filter( value => (value[0].length > 0))); // メンバー情報から空の配列を削除
  
  
  // 在席リストのメンバー名前を取得
  const membersNameL = [];
  const membersNameC = [];
  const membersNameR = [];
  const membersNames = [ membersNameL, membersNameC, membersNameR ];
    
  membersL.forEach( memberL => membersNameL.push(memberL[0]));
  membersC.forEach( memberC => membersNameC.push(memberC[0]));
  membersR.forEach( memberR => membersNameR.push(memberR[0]));
  
  
  // 休日パターンを取得
  const normalHolMems = offDayList.getRange(1, 1, offLastRow, 1).getValues().flat();    // 土日休みのメンバーを取得(サービス)
  const normalHolMembers = membersNameL.concat(normalHolMems);                          // 土日休みのメンバー(全員分・空白含む)



  // 休日 への切替判定を設定 (サービス予定表に設定文字の記入があるか)
  const offDays = [ '休み', '有給', '有休', '振休', '振替', '代休', 'RH', 'ＲＨ', '完フレ', '完ﾌﾚ', '完全フレックス', '完全ﾌﾚｯｸｽ' ];  // 休日パターンを設定
  
  // ISOWAの休日を取得(年末年始・GW等)
  const isowaOffDays = isoOffDayList.getRange(1,  1, isoOffLastRow, 1).getValues().flat(); // ISOWAの休日情報を取得
  
  
  // 宿直パターンを設定(24時間サービス)
  const nightShtPatterns = [ '24', '24h', '24H', '２４', '２４ｈ', '２４Ｈ' ];
  
  // フレックスパターンを設定
  const flexPatterns = [ 'フレックス', 'ﾌﾚｯｸｽ', 'フレ', 'ﾌﾚ' ];
  
  // 曜日判定
  const saturday = dayOfNum === 6;   // 土曜日 判定
  const sunday   = dayOfNum === 0;   // 日曜日 判定
  
  // 直近の予定・状態 
  let setContents;  // 直近の状態
  let detail;       // 直近の予定
  
  // メンバー情報を記入位置別で３つの配列に分ける。
  let memLeft   = [];
  let memCenter = [];
  let memRight  = [];
  
  // 配列の書込先の列番号
  let posiL =  5; // memLeftの書込先の列番号
  let posiC = 10; // memCenterの書込先の列番号
  let posiR = 15; // memRightの書込み先の列番号
  
  
  
  
// ===  配列 [ memberObj ] のメンバー情報から在席リストの書込情報を取得  ===  //

  membersObj.forEach( el => {
    
    // 配列[ target ] に記入したメンバーの情報
    const name            = String(el.name);          // メンバーの名前
    const monContents     = el.monContents;           // 今月の予定
    const nextMonContents = el.nextMonContents;       // 翌月の予定
    let   contents        = String(el.contents);      // 当日の予定
    let   nextContents    = String(el.nextContents);  // 翌日の予定
    const monColor        = el.monColor;              // 今月のセル背景色
    const color           = el.color;                 // 当日のセル背景色
    const nextMonColor    = el.nextMonColor;          // 翌月のセル背景色
    const nextColor       = el.nextColor;             // 翌日のセル背景色
    const swStart         = el.swStart;               // 出社時の切替設定
    const swEnd           = el.swEnd;                 // 退社時の切替設定
    const swDetail        = el.swDetail;              // 予定欄の切替設定
                     
                     
    // 変数の宣言
    let rowNum;            // メンバーの行番号
    let colNum;            // 列番号(メンバーの在席状態の書込先)
    let detailNum;         // メンバー予定詳細の列番号
    let position;          // メンバー記入の位置  
    let normalHol;         // 通常休み判定
  
    // 関数実行
    GetRowColNum();      // 行 ・ 列番号 ・ 記入位置 を取得
  
    // 2ヶ月分の予定(配列)を取得
    const twoMonthStatus = GetSchedule();  
  
    // 直近の在席リストの状態と予定
    setContents = attendList.getRange(rowNum, colNum, 1, 1).getValue();   // 在席リストの状態
    detail = attendList.getRange(rowNum, detailNum, 1, 1).getValue();     // 予定詳細の状態

    // 宿直判定(当日)
    const nightSft = nightShtPatterns.some( nightShtPattern => contents.indexOf(nightShtPattern) !== -1 ); // 宿直

    // フレックス判定(当日・翌日)
    const flex     = flexPatterns.some( flexPattern => contents.indexOf(flexPattern) !== -1 );      // フレックス(当日)
    const nextFlex = flexPatterns.some( flexPattern => nextContents.indexOf(flexPattern) !== -1 );  // フレックス(翌日)

    // 現在の状態判定
    const goOutNow = setContents === '外出中';  // 外出中
  
    // 状態判定(当日)
    const offDayJudge = twoMonthStatus[0] === '休み' || twoMonthStatus[0] === '公休';    // 休み
    const attend      = twoMonthStatus[0] === '在席';　  // 在席
    const satDuty     = twoMonthStatus[0] === '当番';    // 当番
    const goOut       = twoMonthStatus[0] === '外出';    // 外出
    const trip        = twoMonthStatus[0] === '出張';    // 出張
    const atHome      = twoMonthStatus[0] === '在宅';    // 在宅
    const work        = twoMonthStatus[0] === '出勤';    // 出勤
    
    // 状態判定(翌日)
    const nextOffDayJudge = twoMonthStatus[1] === '休み'; // 休み
    const nextAttend      = twoMonthStatus[1] === '在席'; // 在席
    const nextGoOut       = twoMonthStatus[1] === '外出'; // 外出
    const nextTrip        = twoMonthStatus[1] === '出張'; // 出張
    const nextWork        = twoMonthStatus[1] === '出勤'; // 出勤  
    const nextPaidHol     = offDays.some( paid => nextContents.indexOf(paid) !== -1 ); // 公休

  
    // 詳細項目の文言から 外出 ・ 出張 を削除
    if ( goOut ) contents = contents.replace('外出', '');              
    if ( trip ) contents = contents.replace('出張', '');              
    if ( nextGoOut ) nextContents = contents.replace('外出', '');      
    if ( nextTrip ) nextContents = nextContents.replace('出張', '');
  

    // 予定詳細欄に描く予定の期間を記入(例：12/1〜4 休み 等)
    const schedPeriod = AddSchedPeriod();

      // ログ確認用(メンバーの情報)
    console.log(name, rowNum, colNum, detailNum, position, monContents, nextMonContents, contents, nextContents, setContents, detail, swStart, swEnd, swDetail);
//    console.log(twoMonthStatus);

  
  
//  ===  在席リストの状態・詳細項目を書込　関数を実行  === //
  
    // プロジェクト実行時間の設定
    const startTimer = nowTime == 17 && !flex;
    const endTimer = nowTime == 21 && !flex; 
    const flexTimer = contents === flex;
  
    // 出社時に当日の在席状態を書込
    starts.forEach( start => {
      if ( start === name && startTimer ) StartWrite();
    });

    // 帰宅時に翌日の在席状態を書込
    ends.forEach( end => {
      if ( end === name && endTimer ) EndWrite(schedPeriod);
    });
      
    // 期間限定で発動(iサーチ打ち合わせ開始)
    mtgStarts.forEach( mtgStart => {
      if ( mtgStart === name && nowTime == 1 ) MtgStartWrite();
    });

    // 期間限定で発動(iサーチ打ち合わせ終了)
    mtgEnds.forEach( mtgEnd => {
      if ( mtgEnd === name && nowTime == 2 ) MtgEndWrite();
    });


    SetArray();  // メンバー情報を配列に追加


    /* ========================================================================= /
    /  ===  在席リストの状態・詳細項目を書込　関数                                    === /
    /  ======================================================================== */
    function SetStatus(select, value) {
      
      // 状態を書込
      if ( select === "contents" ) setContents = value;
      
      // 詳細項目を書込
      if ( select === "detail" ) detail = value;

    }; 


    /* ========================================================================= /
    /  ===  行 ・ 列番号 ・ 記入位置 を取得　関数                                   === /
    /  ======================================================================== */

    function GetRowColNum() {
    
      // メンバーの行番号と記入位置を取得
      membersNames.forEach( member => {
                      
        if ( member.indexOf(name) !== -1 ) {
          rowNum = member.indexOf(name) + 1; // 行番号を取得
        
          if ( member === membersNameL ) {
            colNum =  posiL;
            position = "L";
          }
        
          if ( member === membersNameC ) {
            colNum = posiC;
            position = "C";
          }
          
          if ( member === membersNameR ) {
            colNum = posiR;
            position = "R";
          }
        
        }
      
      });

      // 列番号(メンバーの予定詳細の書込先)
      detailNum = colNum + 1;
          
    }
    
    
    
    /* ========================================================================= /
    /  ===  休日パターン別の休日判定　関数                                         === /
    /  ======================================================================== */
    
    function HolidayJudge( colors, dateNum, conts, offJudge ) {
    
      // 変数の初期化
      let holJudge = false;
      let holColor = false;
      
      // 通常の土日休みパターンの場合、 「true」 を返す
      normalHol = normalHolMembers.includes(name);
       
      // 曜日判定
      const saturday = dateNum === 6;   // 土曜日 判定
      const sunday   = dateNum === 0;   // 日曜日 判定
      
      // 当番判定
      const satDuty  = saturday && conts.indexOf('当番') !== -1; // 当番(当日)
      
      // 休日パターンのセルの背景色
      holColor = colors === "#d9d9d9" || colors === "#efefef" || colors === "#cccccc";
      
      // セル背景色が灰色 又は ISOWA休日 の場合は休日と判定
      if ( holColor || offJudge ) {
        holJudge = true;
      } else {
        holJudge = false;        
      }

      // 通常休みパターンで土曜(当番以外)・日曜日の場合は休日と判定
      if ( normalHol ) {
        if ( sunday || saturday && !satDuty ) holJudge = true;
      }
      
      return holJudge;

    }





    /* ========================================================================= /
    /  ===  始業時に実行   関数                                                 === /
    /  ======================================================================== */
    function StartWrite() {
      
      // 状態を書込
      if ( attend || satDuty || work ) SetStatus("contents", '在席'); // 「在席」を書込
      if ( offDayJudge ) SetStatus("contents", '休み');               // 「休み」 を書込
      if ( goOut ) SetStatus("contents", '外出中');                   // 「外出」を書込
      if ( trip )  SetStatus("contents", '出張中');                   // 「出張」を書込
      if ( atHome ) SetStatus("contents", '在宅');                    // 「在宅」を書込
      
      // 詳細予定を書込
      if ( goOut || trip || nightSft ) {                             // 外出 ・ 出張 ・ 宿直 の場合
        details.forEach( detail => {
          if ( detail === name ) SetStatus("detail", contents);
        });
      }

    }



    /* ========================================================================= /
    /  ===  終業時に実行   関数                                                 === /
    /  ======================================================================== */
    function EndWrite(schedPeriod) {

      // 翌日が 出張以外 ・ 外出中 でなければ実行
      if ( !nextTrip && !goOutNow ) {
        SetStatus("contents", '帰宅');                        // 「帰宅」 を書込
        if ( nextOffDayJudge ) SetStatus("contents", '休み'); // 「休み」を書込
      };
      
      
      details.forEach( detail => SetStatus("detail", schedPeriod));
    
      // 予定詳細欄に書込
      // 翌日が 休み(公休日以外) の場合
//      if ( nextPaidHol ) details.forEach( detail => SetStatus("detail", schedPeriod));
//      
//      // 翌日が 出張 の場合
//      if ( nextTrip ) details.forEach( detail => SetStatus("detail", schedPeriod));
//      
//      // 翌日が 外出 の場合
//      if ( goOut ) details.forEach( detail => SetStatus("detail", schedPeriod));
//      
//      // 翌日が 在宅 の場合
//      if ( atHome ) details.forEach( detail => SetStatus("detail", schedPeriod));
//      
//
//      // 翌日が 休日 ・ 外出 ・ 出張 以外 かつ 当日の状態が「外出中」でない場合は予定を削除
//      if ( nextAttend && setContents !== "外出中" ) details.forEach( detail => SetStatus("detail", ''));

    }


    /* ========================================================================= /
    /  ===  会議開始時に実行   関数                                              === /
    /  ======================================================================== */
    function MtgStartWrite() {
      const goto = attendList.getRange(19, 5, 1, 1).getValue();
      const kamikura = attendList.getRange(28, 5, 1, 1).getValue();     
      const run =  goto !== "休み" && kamikura !== "休み" && goto != "外出中" &&　kamikura != "外出中" && goto !== "出張中" && kamikura != "出張中";
      
      if ( run ) {
        SetStatus("contents", '会議中');
        if ( detail !== '' ) SetStatus("detail", '10 ~ 11時');
      }
    }


    /* ========================================================================= /
    /  ===  会議終了時に実行   関数                                              === /
    /  ======================================================================== */
    function MtgEndWrite() {
      if ( setContents === '会議中' ) {
        SetStatus("contents", '在席');
        if ( detail !== '' ) SetStatus("detail", '');
      }
    }



    /* ========================================================================= /
    /  ===  メンバー情報を記入位置別で３つの配列に分ける。  関数                        === /
    /  ======================================================================== */
    
    function SetArray() {
    
      // メンバー記入位置によって３つの配列に分類
      if ( position === "L" ) memLeft.push([name, rowNum, setContents, detail]);
      if ( position === "C" ) memCenter.push([name, rowNum, setContents, detail]);
      if ( position === "R" ) memRight.push([name, rowNum, setContents, detail]);  
      
    }



    /* ========================================================================= /
    /  ===  予定詳細欄に予定の期間を記入 関数                                      === /
    /  ======================================================================== */       
      
    function AddSchedPeriod() {
        
      // 変数の初期化
      let dayStatus     = '';
      let latestSched   = '';
      let periodStart   = '';
      let periodEnd     = '';
      let writePeriod   = '';
      let twoMonthPlus  = [];
      let twoMonthSched = [];
      let stateArr      = [];
      let periodArr     = [];
      let periodContinue = false;
      let iStart = 0;
      let iEnd   = 0;
      
      i = 1;
      
      // 予定期間を取得(開始日・終了日)
      twoMonthPlus = twoMonthStatus;
      twoMonthPlus.shift();          // 配列[twoMonthStatus]から当日予定を削除
      twoMonthSched = monContents.concat(nextMonContents);
      twoMonthSched.shift();
      console.log(twoMonthSched)

      twoMonthPlus.forEach( el => {
                
        // 状態の格納条件
        const slctStatus = ( (el === '公休' || el === '休み' || el === '外出' || el === '出張' || el === '在宅') && dayStatus === '' );
      
        // 状態を変数に格納
        if ( slctStatus ) dayStatus = el;
      
        // 予定期間の開始日を取得
        if ( el === dayStatus && periodStart === '' ) {
         
          // 日付を取得                     
          const date = new Date();                                   // 当日の日付
          date.setDate(date.getDate() + i);                          // 当日 + i の日付
          day = Utilities.formatDate(date, 'Asia/Tokyo', 'M/d');   // 日付のフォーマット
          
          // 開始日を書込
          periodStart = day;
          
          // 配列[twoMonthStatus]のi番目より前の要素を削除
          latestSched = twoMonthSched[i-1];
          stateArr = twoMonthPlus;
          stateArr = stateArr.slice(i-1);
          iStart = i;
          
        }
      
      
        // 予定期間の終了日を取得
        if ( el !== dayStatus && periodStart !== '' && periodEnd === '') {
          
          // 日付を取得
          const date = new Date(); 
          date.setDate(date.getDate() + i-1);                        // 当日 + (i-1) の日付
          day = Utilities.formatDate(date, 'Asia/Tokyo', 'M/d');   // 日付のフォーマット
          
          // 終了日を書込
          if ( dayStatus === '公休' &&　el === '休み' || dayStatus === '休み' &&　el === '公休' ) {
            periodContinue = true;
          } else  {
            periodContinue = false;
          }
          
          // 終了日を書込
          if ( !periodContinue ) { 
            
            // 終了日を書込
            periodEnd = day;
          
            // 配列[twoMonthStatus]のi番目より後の要素を削除
            iEnd = i - iStart;
            stateArr.splice(iEnd);
            
          }
          
        }
      
        // 日付を1日進める
        i++;
        
      });


      // 変数
      let startDate = periodStart.split('/');             // 開始月・日の分割
      let endDate   = periodEnd.split('/');               // 終了月・日の分割 
      const startEnd  = startDate[0] === endDate[0];      // 開始・終了月の一致判定


      // 配列が全て公休日の場合にtrueを返す
      const hol       = '公休';
      const holOnly   = stateArr.every( period => hol.indexOf(period) !== -1 );
//      console.log(holOnly, startDate, endDate);


      // 予定期間の書込内容 (公休日のみの場合は書込なし)
      if ( !holOnly ) {
        
        // 予定期間が １日 だった場合
        if ( stateArr.length === 1 ) {
          writePeriod = `${periodStart}`;
          
        // 予定期間が 2日以上 + 同じ月の場合
        } else if ( startEnd ) {
          writePeriod = `${periodStart}〜${endDate[1]}`;
          
        // 予定期間が 2日以上 + 月をまたぐ場合 
        } else {
          writePeriod = `${periodStart}〜${periodEnd}`;
        }
        
        // 予定内容が 休み or 公休 だった場合
        if ( dayStatus === '休み' || dayStatus === '公休' ) {
          writePeriod += ' 休み';
        }         
        
        // 予定内容が 出張 だった場合
        if ( dayStatus === '出張' ) {
          latestSched = latestSched.replace('出張', '');
          writePeriod += ` ${latestSched}`;
        }
          
        // 予定内容が 外出 だった場合
        if ( dayStatus === '外出' ) {
          latestSched = latestSched.replace('外出', ''); 
          writePeriod += ` ${latestSched}`;
        }
        
        // 予定内容が 在宅 だった場合
        if ( dayStatus === '在宅' ) writePeriod += ' 在宅';
        
      }

      // ログ確認用
//      console.log(dayStatus, writePeriod, twoMonthPlus, stateArr, iStart, iEnd, twoMonthSched);

    return writePeriod;
 
    }














    /* ========================================================================= /
    /  ===  翌日以降の予定を取得  関数                                           === /
    /  ======================================================================== */
  
    function GetSchedule() {

      let status;                   // 状態
      let holidayJudges     = [];   // 今月の休日判定(土日休み・平日休み)
      let nextHolidayJudges = [];   // 翌月の休日判定(土日休み・平日休み)
      let offDayJudges      = [];   // 今月の休日判定
      let nextOffDayJudges  = [];   // 翌月の休日判定
      let isoOffJudge       = [];   // ISOWAの休日判定(今月)
      let nextIsoOffJudge   = [];   // ISOWAの休日判定(翌月)      
      let offHolJudges      = [];   // 状態(今月分・休日のみ)
      let nextOffHolJudges  = [];   // 状態(来月分・休日のみ)
      let monStatus         = [];   // 状態(今月分・予定/休日除く)
      let nextMonStatus     = [];   // 状態(翌月分・予定/休日除く)
      let monthStatus       = [];   // 状態(今月分・統合)
      let nextMonthStatus   = [];   // 状態(翌月分・統合)
      let goOut   = false;          // 外出判定
      let trip    = false;          // 出張判定　　　　　　　
      let atHome  = false;          // 在宅判定
      let satDuty = false;          // 当番判定
      let work    = false;          // 出勤判定
      

      // [関数の引数] 月のメンバー情報( セル背景色, 予定[作業予定表], 予定[状態], 休日判定[セル背景色], ISOWA休日判定, 休日判定[記入] )
      const schedMonth   = [ monColor, monContents, monStatus, offDayJudges, isoOffJudge, "thisMon", holidayJudges ];                            // 今月のメンバーの情報
      const nextSchedMon = [ nextMonColor, nextMonContents, nextMonStatus, nextOffDayJudges, nextIsoOffJudge, "nextMon", nextHolidayJudges ];    // 翌月のメンバーの情報
      
      // 関数を実行
      StatusJudge(...schedMonth);   // 今月分の在席状態を取得
      StatusJudge(...nextSchedMon); // 来月分の在籍状態を取得
      
      
      // [関数の引数] 休日判定の配列( 休日判定[記入], 休日判定[セル背景色・土日], 統合先 )
      const offHolMonth = [ offDayJudges , holidayJudges, offHolJudges ];                 // 今月の休日判定の配列
      const nextOffHolMonth = [ nextOffDayJudges, nextHolidayJudges, nextOffHolJudges ];  // 翌月の休日判定の配列
      
      // 関数を実行
      OffHolIntegral(...offHolMonth);      // 今月分(休日判定[2種]の配列を新たな配列に統合)
      OffHolIntegral(...nextOffHolMonth);  // 翌月分(休日判定[2種]の配列を新たな配列に統合)
      
      // [関数の引数] 月の予定・休日の配列( 月の予定, 休日判定[統合], 統合先 )
      const thisMonth = [ monStatus, offHolJudges , monthStatus ];
      const nextMonth = [ nextMonStatus, nextOffHolJudges , nextMonthStatus ];
      
      // 関数を実行
      ArrayIntegral(...thisMonth);  // 今月分(予定・休日の配列を新たな配列に統合)
      ArrayIntegral(...nextMonth);  // 翌月分(予定・休日の配列を新たな配列に統合)
      
      // 今月・翌月の予定を統合
      const twoMonStatus = monthStatus.concat(nextMonthStatus);
            
      return twoMonStatus;
      
      
      /* ========================================================================= /
      /  ===  1ヶ月分の在席状態を取得  関数                                         === /
      /  ======================================================================== */  

      function StatusJudge(colorArr, contsArr, statusArr, offDayArr, isoJudge, func, holArr) {
      
        let i = 0;　// 変数の初期化
        colorArr.forEach( el => {
                  
          // 変数の定義
          let day, dateNum;
          
                    
          // 今月の曜日・予定を取得
          if ( func === "thisMon" ) {
            const date = new Date();                                       // 当日の日付
            date.setDate(date.getDate() + i);                              // 当日 + i の日付
            day = Utilities.formatDate(date, 'Asia/Tokyo', 'MM/dd');       // 日付のフォーマット
            dateNum = date.getDay();                                       // 当日 + i の曜日
          }

          // 翌月の曜日・予定を取得        
          if ( func === "nextMon" ) {
            const date = new Date();
            let _day = new Date(date.getFullYear(), date.getMonth()+1, 1); // 翌月の初日
            _day.setDate(_day.getDate() + i);                              // 初日 + i の日付
            day = Utilities.formatDate(_day, 'Asia/Tokyo', 'MM/dd');       // 日付のフォーマット
            dateNum = _day.getDay();                                       // 初日 + i の曜日をセット
          }
        
          // 当日 + i の予定を取得
          const conts = String(contsArr[i]);
        
          // 初期設定
          status = "在席";
      
          // 外出判定
          goOut = conts.indexOf('外出') !== -1;
          if ( goOut ) status = "外出";
      
          // 出張判定
          trip = conts.indexOf('出張') !== -1;
          if ( trip ) status = "出張";

          // 在宅判定
          atHome = conts.indexOf('在宅') !== -1;
          if ( atHome ) status = "在宅";
        
          // 当番判定
          satDuty = conts.indexOf('当番') !== -1;
          if ( satDuty ) status = "当番";
        
          // 出勤判定
          work = conts.indexOf('出勤')　!== -1;
          if ( work ) status = "出勤";

          // 配列[statusArr] に状態を格納
          statusArr.push(status);
      
          // 配列[offDayArr] に休日判定を格納
          offDayArr.push(offDays.some( offDay => conts.indexOf(offDay) !== -1 ));
      
          // ISOWA休日判定
          isoOffJudge = isowaOffDays.some( isowaOffDay => {
                        isowaOffDay = Utilities.formatDate(isowaOffDay, 'Asia/Tokyo', 'MM/dd');
                        isowaOffDay = String(isowaOffDay);
                        day = String(day);
                        return day.indexOf(isowaOffDay) != -1; 
                      });
      
          // 休日判定を実行
          const holiday = HolidayJudge(el, dateNum, conts, isoOffJudge);

          // 結果を配列に格納
          holArr.push(holiday);

          // 日付を1日進める
          i++; 
      
        });
      
      }

      
      
      
      /* ========================================================================= /
      /  ===  2種類の休日判定の配列を統合 関数                                      === /
      /  ======================================================================== */  
      
      function OffHolIntegral(offDayJudges, holidayJudges, offHolJudges) {
        i = 0; // 変数の初期化
        offDayJudges.forEach( el => {                  
          let offHol = "出勤";                                                  // 変数の初期化
//          const offHolJudge = !( el === false && holidayJudges[i] === false ); // 休日判定(true/false)
          const offJudge = !( el === false );                                   // 休日判定(true/false)
          const holJudge = !( holidayJudges[i] === false );                     // 公休判定(true/false)
          if ( offJudge ) offHol = "休み";                                      // 休日判定がtrueの場合、休みを格納
          if ( holJudge ) offHol = "公休";
          offHolJudges.push(offHol);                                           // 配列[offHolJudges]に追加
          i++;                                                                 // 配列[holidayJudges]内の参照先を次の要素に切替
        });
      }


      /* ========================================================================= /
      /  ===  2種類の休日判定の配列を統合 関数                                      === /
      /  ======================================================================== */ 

      function ArrayIntegral(monStatus, offHolJudges , statusArr) {

        // 変数の初期化
        i = 0;
        
        monStatus.forEach( el => {
          
          // 休日予定
          const offHol = offHolJudges[i];
                          
          // 状態判定
          const attend  = el === '在席';              // 在席判定
          const offDay  = offHol === '休み';          // 休日判定
          const holDay  = offHol === '公休';          // 公休判定
          const goOut   = el.indexOf('外出') !== -1;  // 外出判定
          const trip    = el.indexOf('出張') !== -1;  // 出張判定
          const atHome  = el.indexOf('在宅') !== -1;  // 在宅判定
          const satDuty = el.indexOf('当番') !== -1;  // 当番判定
          const work    = el.indexOf('出勤') !== -1;  // 出勤判定
        
          // 配列[statusArr]に追加
          if ( goOut ) {
            statusArr.push("外出");
          } else if ( trip ) {
            statusArr.push("出張");
          } else if ( atHome ) {
            statusArr.push("在宅");
          } else if ( satDuty ) {
            statusArr.push("当番");
          } else if ( work ) {
            statusArr.push("出勤");
          } else if ( offDay ) {
            statusArr.push("休み");
          } else if ( holDay ) {
            statusArr.push("公休");
          } else if ( attend ) {
            statusArr.push("在席");
          } else {
            statusArr.push("在席");
          }
        
          // 配列[offHolJudges]内の参照先を次の要素に切替
          i++;
                            
        });

      }
      
      


      
      
    }

  });

// ------------------------------------------------------------------------------------------------------------ //


  setInfo(); // 在席リストに情報を書込

  /* ========================================================================= /
  /  ===  在席リストに情報を書込   関数                                          === /
  /  ======================================================================== */
  function setInfo() {

　  // メンバー情報を格納する配列を定義
    const memInfoL = [ memLeft, posiL ];          // 在席リスト左側の列(メンバー情報)
    const memInfoC = [ memCenter, posiC ];        // 在席リスト中側の列(メンバー情報) 
    const memInfoR = [ memRight,posiR ];          // 在席リスト右側の列(メンバー情報)
    let mems = [ memInfoL, memInfoC, memInfoR ];  // 在席リストの書込み情報元
  

    // 在席リストの列毎に書込情報を取得しスプレットシートに順番に書込
    mems.forEach( mem => {
    
      // データの初期化
      let nums        = [];    // 行番号の存在チェック用配列
      let memNum      = [];    // メンバーの行番号の配列
      let _blankArray = [];    // 存在しない行番号の配列
      let blankArray  = [];    // 存在しない行番号の空白を含んだ配列(mem[0]に書込用)
      let memsArray   = [];    // メンバー(空白を含んだ)情報を格納する配列
      let blankContents;       // 自動切替リストに登録されていない要素


      // 配列の要素を行番号の昇順に並べ替え
      mem[0].sort((a, b) => a[1] - b[1]);       

      minNum = mem[0].slice(0).flat()[1];       // 最小値(書込行)
      maxNum = mem[0].slice(-1).flat()[1];      // 最大値(書込行)


      // 配列[nums]に最小値(minNum)〜最大値(maxNumの)数値を格納
      for ( let i = minNum; i <= maxNum; i++) nums.push( i );

      // 配列[memNum]にメンバーの行番号を格納
      mem[0].forEach( el => memNum.push(el[1]) );

      // 配列[memNum]と配列[nums]を比較。重複しない番号を新たな配列に格納
      _blankArray = nums.filter( el => memNum.indexOf(el) == -1 );
    
      // 配列[_blankArray]の行の情報を取得(空情報)
      _blankArray.forEach( el => {
        blankContents = attendList.getRange( el, mem[1]-2, 1, 4 ).getValues().flat();
        blankContents[1] = el;
        blankArray.push(blankContents);
    
      });

      // 配列[blankArray]を配列[mem[0]]に格納
      memsArray = mem[0].concat(blankArray);
    
      // 配列を行番号の昇順に並べ替え
      memsArray.sort((a, b) => a[1] - b[1]);

      // ログ確認用(書込情報)
//      console.log(memsArray);

      // メンバー名・行番号を削除
      memsArray.forEach( el => el.splice( 0, 2 ) );  

      // メンバー状態をスプレットシートに書込
      const row = maxNum - 2;
      attendList.getRange(3, mem[1], row, 2).setValues(memsArray);
 
    });
    
  }

}
