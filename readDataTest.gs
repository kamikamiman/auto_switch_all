// ============================================================================================================ //
//       【関数】 サービス予定表・在席リストから予定を取得                                                              //
// ============================================================================================================ //

function ReadDataTest() {

  // ---------------------------------------------------------------------------------------------------------- // 
  //       在席リストのシート(切替設定)情報                                                                            //
  // ---------------------------------------------------------------------------------------------------------- // 

  // 対象者を取得
  const SwitchSet  = ssSet.getSheetByName('切替設定');                                                  // シート(切替設定)
  const swLastRow = SwitchSet.getRange(1, 2).getNextDataCell(SpreadsheetApp.Direction.DOWN).getRow();  // シートの最終行番号
  const infoRanges = SwitchSet.getRange(2, 2, swLastRow-1, 5).getValues();                             // 切替設定情報を取得

  // 対象者を配列に追加
  infoRanges.forEach( infoRange => {
    targets.push(infoRange[0]);                                  // 切替対象者を取得
    if ( infoRange[1] === "有効" ) starts.push(infoRange[0]);     // 出社時
    if ( infoRange[2] === "有効" ) ends.push(infoRange[0]);       // 退社時
    if ( infoRange[3] === "有効" ) details.push(infoRange[0]);    // 予定詳細内容
    if ( infoRange[4] === "東館" ) attendSels.push(infoRange[0]); // 在席時の表示選択
  });

  // ログ確認用(確定)
  console.log("targets(切替対象者):" + targets);
  console.log("starts(切替 出社時):" + starts);
  console.log("ends(切替 帰社時):" + ends);
  console.log("details(切替 予定記入):" + details);
  console.log("attendSels(東館選択):" + attendSels);

  // ----------------------------------------------------------------------------------------------------------- // 
  //       在席リストのメンバー情報                                                                                  //
  // ----------------------------------------------------------------------------------------------------------- //   
      
  // 在席リストの最終行番号
  const lastRow    = attendList.getRange('C:C').getLastRow();

  // 在席リストのメンバー情報を取得
  const membersL   = attendList.getRange(1,  3, lastRow, 4).getValues();  // メンバー情報(在席リスト左)
  const membersC   = attendList.getRange(1,  8, lastRow, 4).getValues();  // メンバー情報(在席リスト中)
  const membersR   = attendList.getRange(1, 13, lastRow, 4).getValues();  // メンバー情報(在席リスト右)  

  // 在席リストのメンバー名を取得
  const membersNameL = [];   // 左側の列メンバー
  const membersNameC = [];   // 中央の列メンバー
  const membersNameR = [];   // 右側の列メンバー

  membersL.forEach( memberL => membersNameL.push(memberL[0]));   // 左側の列メンバーを取得
  membersC.forEach( memberC => membersNameC.push(memberC[0]));   // 中央の列メンバーを取得
  membersR.forEach( memberR => membersNameR.push(memberR[0]));   // 右側の列メンバーを取得

  const membersNames = [ membersNameL, membersNameC, membersNameR ]; // メンバー全員



  // ----------------------------------------------------------------------------------------------------------- // 
  //       在席リストの夜勤担当者の行・列番号を取得                                                                     //
  // ----------------------------------------------------------------------------------------------------------- //   
  
  // 変数の初期化
  const dutyData1 = [];                             // 夜勤担当者の行・列番号(1人目)
  const dutyData2 = [];                             // 夜勤担当者の行・列番号(2人目)
  const nightDutysData = [ dutyData1, dutyData2 ];  // 夜勤担当者の行・列番号(2人分)

  const nightDutyNames = GetNightDutys();  // 夜勤担当者2名を取得(フルネーム・配列)
  i = 0;

  // 夜勤担当者と在席リストメンバーリストを比較
  nightDutyNames.forEach( member => {

    // 変数の定義
    let rowNum,colNum;                // 夜勤当番者の行・列番号(在席リスト)
    let setRowJudge, setColJudge;     // 行・列番号の存在判定

    // メンバーの存在チェック(1人目)
    CheckMembers();

    // メンバーの存在判定がtrueで実行(2人目)
    if ( setRowJudge && setColJudge ) CheckMembers();


    // [関数] 夜勤担当者の名前が在席リストメンバーがあるかチェック
    function CheckMembers() {

      // 初期化
      rowNum = 0;  // 行番号
      colNum = 0;  // 列番号

      // 夜勤担当者と在席リストメンバーが一致した時の列番号を取得
      if ( membersNameL.indexOf(nightDutyNames[i]) != -1 ) {
        rowNum = membersNameL.indexOf(nightDutyNames[i]) + 1;
        colNum = posiL - 2;
      }

      if ( membersNameC.indexOf(nightDutyNames[i]) != -1 ) {
        rowNum = membersNameC.indexOf(nightDutyNames[i]) + 1;        
        colNum = posiC - 2;
      }

      if ( membersNameR.indexOf(nightDutyNames[i]) != -1 ) {
        rowNum = membersNameR.indexOf(nightDutyNames[i]) + 1;        
        colNum = posiR - 2;
      }

      console.log("membersNameL:" + membersNameL);
      console.log("membersNameC:" + membersNameC);
      console.log("membersNameR:" + membersNameR);
      console.log("nightDutyNames[i]:" + nightDutyNames[i]);

      // 配列への書込み条件
      setRowJudge = rowNum !== 0;
      setColJudge = colNum !== undefined;
      
      // ログ確認用(確定)
      console.log("member(夜勤担当者):" + member);
      console.log("nightDutyNames[i](在席リストのメンバー):" + nightDutyNames[i] );
      console.log("rowNum(在席リストの行番号):" + rowNum);
      console.log("colNum(在席リストの列番号):" + colNum);      
      console.log("setRowJudge(実行判定1):" + setRowJudge);
      console.log("setColJudge(実行判定2):" + setColJudge);

      if ( setRowJudge && setColJudge ) {
        nightDutysData[i].push(rowNum, colNum);                     // 配列へ行・列番号を書込
        i++;                                                        // 次のメンバーをチェック
      }
    };
  });




  // ----------------------------------------------------------------------------------------------------------- // 
  //       土曜当番の担当者を取得                                                                                    //
  // ----------------------------------------------------------------------------------------------------------- //

  // 変数の初期化
  const satData1 = [];                             // 土曜当番者の行・列番号(1人目)
  const satData2 = [];                             // 土曜当番者の行・列番号(2人目)
  const satDutysData = [ satData1, satData2 ];     // 土曜当番者の行・列番号(2人分)

  const satDutyNames = GetSatDutys();  // 土曜当番者2名を取得(フルネーム・配列)

  i = 0;

// 土曜当番者と在席リストメンバーリストを比較
  membersNames.forEach( member => {

    // 変数の定義
    let rowNum,colNum;                // 土曜当番者の行・列番号(在席リスト)
    let setRowJudge, setColJudge;     // 行・列番号の存在判定

    // メンバーの存在チェック(1人目)
    CheckMembers();

    // メンバーの存在判定がtrueで実行(2人目)
    if ( setRowJudge && setColJudge ) CheckMembers();


    // [関数] 土曜当番者の名前が在席リストメンバーがあるかチェック
    function CheckMembers() {

      // 初期化
      rowNum = 0;  // 行番号
      colNum = 0;  // 列番号
　
      // 土曜当番者と在席リストのメンバーが一致した時の行番号を取得
      rowNum = member.indexOf(satDutyNames[i]) + 1;

      // 土曜当番者と在席リストメンバーが一致した時の列番号を取得
      if ( membersNameL.indexOf(satDutyNames[i]) != -1 ) colNum = posiL - 2;
      if ( membersNameC.indexOf(satDutyNames[i]) != -1 ) colNum = posiC - 2;
      if ( membersNameR.indexOf(satDutyNames[i]) != -1 ) colNum = posiR - 2;

      // 配列への書込み条件
      setRowJudge = rowNum !== 0;
      setColJudge = colNum !== undefined;

      if ( setRowJudge && setColJudge ) {
        satDutysData[i].push(rowNum, colNum);                     // 配列へ行・列番号を書込
        i++;                                                        // 次のメンバーをチェック
      }
    };
  });





  // ----------------------------------------------------------------------------------------------------------- // 
  //       当日から翌月までの日付・曜日・ISOWA休日判定を取得                                                             //
  // ----------------------------------------------------------------------------------------------------------- //

  const thisMonArr   = [];                    // 今月の配列(日付・曜日・ISOWA休日判定)
  const nextMonArr   = [];                    // 翌月の配列(日付・曜日・ISOWA休日判定)
  const isoOffFormat = [];                  // ISOWA休日(書式設定後)
  const daysOfMonth   = monLastCol - dayNum;  // 今月の残り日数
  const daysOfNextMon = nextMonLastCol -1;     // 翌月の日数

  const date = new Date();                                           // 当日の日付
  let dateFormat = Utilities.formatDate(date, 'Asia/Tokyo', 'M/d');  // 日付の書式をフォーマット

  // ISOWA休日の書式設定
  isowaOffDays.forEach( el => {
                          el = Utilities.formatDate(el, 'Asia/Tokyo', 'M/d');
                          el = String(el);
                          isoOffFormat.push(el);
  });


  // -------------------------------------------------------------------------- // 
  //     クラスでオブジェクトを作成(今月・翌月分の日付・曜日・ISOWA休日判定情報を取得)      //
  // -------------------------------------------------------------------------- // 

  class MonthObj {
    
    constructor(setMonth, daysLeft) { 
      this.setMonth = setMonth;
      this.daysLeft = daysLeft;
    };

    /* ========================================================================= /
    /  ===  日付・曜日・ISOWA休日情報を取得 [ メソッド ]                           === /
    /  ======================================================================== */      

    GetDate() {

      for ( let h = 0; h <= this.daysLeft; h++ ) {

        const data = [];
        const month = dateFormat;   // 日付を取得
        const day = date.getDay();  // 曜日を取得

        // isowa休日判定
        const isoOffJudge = isoOffFormat.some( el => {
                              const strMonth = String(month);
                              const isoOff = strMonth === el; 
                              return isoOff;
                            });

        // 配列に情報を追加
        data.push(month);           // 日付を配列[data]に追加
        data.push(day);             // 曜日を配列[data]に追加
        data.push(isoOffJudge);     // isowa休日判定を配列[data]に追加 
        this.setMonth.push(data);   // 日付・曜日・isowa休日判定を配列[dateArr]に追加

        // 日付を1日進める。
        date.setDate(date.getDate() + 1);                              // 当日 + i の日付
        dateFormat = Utilities.formatDate(date, 'Asia/Tokyo', 'M/d');  // 日付のフォーマット
      }

    };

  };


  // オブジェクト作成(今月の残り日数分)
  const thisMonObj = new MonthObj(thisMonArr, daysOfMonth);   // 今月の日付・曜日・ISOWA休日判定を取得
  thisMonObj.GetDate();                                       // {thisMonObj} に今月の予定を追加

  // オブジェクト作成(翌月の日数分)
  const nextMonObj = new MonthObj(nextMonArr, daysOfNextMon); // 来月の日付・曜日・ISOWA休日判定を取得
  nextMonObj.GetDate();                                       // {nextMonObj} に翌月の予定を追加

  // 今月と翌月の配列を結合(2ヶ月分)
  const twoMonArr = thisMonArr.concat(nextMonArr);




  // -------------------------------------------------------------------------- // 
  //       クラスでオブジェクトを作成                                                //
  // -------------------------------------------------------------------------- // 

  class MemberObj {
    
    constructor(name, row, nextRow) {
      
      this.name = name;
      this.row = row;
      this.nextRow = nextRow;
      
    }
  
    
    /* ========================================================================= /
    /  ===  今月 ・ 翌月 の予定 ・ セルの背景色 を取得 [ メソッド ]                  === /
    /  ======================================================================== */      
 
    GetContents() {

      // 今月の予定を取得
      const monContents = memSched[this.row - 1];

      for( let i = 0; i < dayNum; i++ ) monContents.shift();
      
      const monColor = memColor[this.row - 1];
      for( let i = 0; i < dayNum; i++ ) monColor.shift();      
      const contents = monContents[0];                                    // 本日の予定
      const color = monColor[0];                                          // 本日のセル背景色
      
      // 翌月の予定を取得
      const nextMonContents = nextMemSched[this.nextRow - 1];
      nextMonContents.shift();
      
      const nextMonColor = nextMemColor[this.nextRow - 1];
      nextMonColor.shift();
      
      let nextContents, nextColor;
      
      
      // 翌日の予定・セルの背景色を取得
      if ( dayNum !== lastCol ) {
        // 本日が月末以外の場合
        nextContents = monContents[1];         // 翌日の予定
        nextColor = monColor[1];               // 翌日のセルの背景色       
      
      } else {
        // 本日が月末の場合
        nextContents = nextMonContents[0];     // 翌日の予定
        nextColor = nextMonColor[0];           // 翌日のセルの背景色
      }

      
      // オブジェクトに追加
      this.monContents = monContents;          // 今月の予定
      this.monColor = monColor;                // 今月のセル背景色
      this.contents = contents;                // 本日の予定
      this.color = color;                      // 本日のセル背景色
      this.nextMonContents = nextMonContents;  // 翌月の予定
      this.nextMonColor = nextMonColor;        // 翌月のセルの背景色
      this.nextContents = nextContents;        // 翌日の予定
      this.nextColor = nextColor;              // 翌日のセルの背景色
      
    };
    
  


    /* ========================================================================= /
    /  ===  自動切替設定の情報 を取得 [ メソッド ]                                  === /
    /  ======================================================================== */  

    GetSwitchSet() {
    
      // シート(自動切替)の名前と一致した場合、オブジェクトに切替設定を追加
      infoRanges.forEach( infoRange => {
        if ( infoRange[0] === this.name ) {
          this.swStart  = infoRange[1];  // 出社時(有効・無効)
          this.swEnd    = infoRange[2];  // 退社時(有効・無効)
          this.swDetail = infoRange[3];  // 予定欄(有効・無効)
        }
      });   
    
    }
    
    


    /* ========================================================================= /
    /  ===  在席リストのメンバーの 行 ・ 列番号 ・ 記入位置 を取得　[ メソッド ]       === /
    /  ======================================================================== */      

    GetRowColNum() {
      
      // 変数の宣言
      let rowNum;            // メンバーの行番号
      let colNum;            // 列番号(メンバーの在席状態の書込先)
      let detailNum;         // メンバー予定詳細の列番号
      let position;          // メンバー記入の位置  
    
      // メンバーの行番号と記入位置を取得
      membersNames.forEach( member => {
                      
        if ( member.indexOf(this.name) !== -1 ) {
          rowNum = member.indexOf(this.name) + 1; // 行番号を取得
        
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

          // 列番号(メンバーの予定詳細の書込先)
          detailNum = colNum + 1;

        }
          
      });      

    
      // オブジェクトに追加
      this.rowNum    = rowNum;     // 在席状態 ・ 予定詳細 記入行
      this.colNum    = colNum;     // 在席状態 記入列
      this.position  = position;   // 記入位置 ( 左 ・ 中央 ・ 右 )
      this.detailNum = detailNum;  // 予定詳細 記入列
    
    }
    



    /* ========================================================================= /
    /  ===  直近での在席リストのメンバー在席状態と予定　[ メソッド ]                  === /
    /  ======================================================================== */
    
    GetNowAttendance() {

      // 初期設定
      let setContents = '帰宅';
      let detail = '登録名が存在しません。';
      
      // 在席リストにメンバーの名前が存在する場合は実行
      if ( this.rowNum !== null && typeof this.rowNum !== 'undefined' ) {
        // @ts-ignore
        setContents = attendList.getRange(this.rowNum, this.colNum, 1, 1).getValue();   // 在席リストの状態
        // @ts-ignore
        detail = attendList.getRange(this.rowNum, this.detailNum, 1, 1).getValue();     // 予定詳細の状態
      }

        // オブジェクトに直近での状態と予定詳細を追加
        this.setContents = setContents;
        this.detail = detail;
      
    }
    
    


    /* ========================================================================= /
    /  ===  2か月分の在席状態を取得 [ メソッド ]                                    === /
    /  ======================================================================== */
  
    GetSchedule() {

      let holidayJudges     = [];   // 今月の休日判定(土日休み・平日休み)
      let nextHolidayJudges = [];   // 翌月の休日判定(土日休み・平日休み)
      let offDayJudges      = [];   // 今月の休日判定
      let nextOffDayJudges  = [];   // 翌月の休日判定
      let offHolJudges      = [];   // 状態(今月分・休日のみ)
      let nextOffHolJudges  = [];   // 状態(来月分・休日のみ)
      let monStatus         = [];   // 状態(今月分・予定/休日除く)
      let nextMonStatus     = [];   // 状態(翌月分・予定/休日除く)
      let monthStatus       = [];   // 状態(今月分・統合)
      let nextMonthStatus   = [];   // 状態(翌月分・統合)

      

      // [関数の引数] 月のメンバー情報( セル背景色, 予定[作業予定表], 予定[状態], 休日判定[セル背景色], ISOWA休日判定, 休日判定[記入] )
      const schedMonth   = [ this.monColor, this.monContents, monStatus, offDayJudges, "thisMon", holidayJudges ];                        // 今月のメンバーの情報
      const nextSchedMon = [ this.nextMonColor, this.nextMonContents, nextMonStatus, nextOffDayJudges, "nextMon", nextHolidayJudges ];    // 翌月のメンバーの情報
      
      // 関数を実行
      this.StatusJudge(...schedMonth);   // 今月分の在席状態を取得
      this.StatusJudge(...nextSchedMon); // 来月分の在籍状態を取得
      
      
      // [関数の引数] 休日判定の配列( 休日判定[記入], 休日判定[セル背景色・土日], 統合先 )
      const offHolMonth = [ offDayJudges , holidayJudges, offHolJudges ];                 // 今月の休日判定の配列
      const nextOffHolMonth = [ nextOffDayJudges, nextHolidayJudges, nextOffHolJudges ];  // 翌月の休日判定の配列
      
      // 関数を実行
      this.OffHolIntegral(...offHolMonth);      // 今月分(休日判定[2種]の配列を新たな配列に統合)
      this.OffHolIntegral(...nextOffHolMonth);  // 翌月分(休日判定[2種]の配列を新たな配列に統合)
      
      // [関数の引数] 月の予定・休日の配列( 月の予定, 休日判定[統合], 統合先 )
      const thisMonth = [ monStatus, offHolJudges , monthStatus ];
      const nextMonth = [ nextMonStatus, nextOffHolJudges , nextMonthStatus ];
      
      // 関数を実行
      this.ArrayIntegral(...thisMonth);  // 今月分(予定・休日の配列を新たな配列に統合)
      this.ArrayIntegral(...nextMonth);  // 翌月分(予定・休日の配列を新たな配列に統合)
      
      // 今月・翌月の予定を統合
      const twoMonStatus = monthStatus.concat(nextMonthStatus);
            
      // 2ヶ月分の状態をオブジェクトに追加
      this.twoMonStatus = twoMonStatus;
      
      // 直近予定からフレックストリガー実行判定の時間範囲をオブジェクトに追加
      this.FlexTimeRange();
      
    }
    
    
    
    

    /* ========================================================================= /
    /  ===  1ヶ月分の在席状態を取得  [ メソッド ]                                 === /
    /  ======================================================================== */  

    StatusJudge(colorArr, contsArr, statusArr, offDayArr, func, holArr) {
      // 変数の初期化
      let i = 0;    // 日付判定に使用
      let t = 0;    // 当日判定に使用
      let n = 0;    // 出発時間確定判定(1以上で確定)
      
      let isoOffJudge;   // ISOWAの休日判定

      colorArr.forEach( el => {
                  
        // 変数の定義
        let day, dateNum, timeAndConts;
        let goOut    = false;          // 外出判定
        let trip     = false;          // 出張判定　　　　　　　
        let atHome   = false;          // 在宅判定
        let satDuty  = false;          // 当番判定
        let work     = false;          // 出勤判定
        let flex     = false;          // フレックス判定
        let flexFull = false;          // 完全フレックス判定
        let timeDiff1 = false;         // 時差判定
        let timeDiff2 = false;         // 時短判定
        let notGoTimeJudge = true;     // 外出・出張時の出発時間指定無し判定
                                        
        // 今月の曜日・予定を取得
        if ( func === "thisMon" ) {
          day = thisMonArr[i][0];
          dateNum = thisMonArr[i][1];
          isoOffJudge = thisMonArr[i][2];
        }

        // 翌月の曜日・予定を取得        
        if ( func === "nextMon" ) {
          day = nextMonArr[i][0];
          dateNum = nextMonArr[i][1];
          isoOffJudge = nextMonArr[i][2];
        }
        
        // 当日 + i の予定を取得
        let conts = String(contsArr[i]);
        
        // 初期設定
        let status = "在席";
      
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
      
       
        // フレックス・完全フレックス判定
        flex = flexPatterns.some( el => conts.indexOf(el) !== -1 );
        flexFull = fullFlex.some( el => conts.indexOf(el) !== -1 );
 

        // 時差出勤判定
        timeDiff1 = conts.indexOf("時差") !== -1;

        // 時短出勤判定
        timeDiff2 = conts.indexOf("時短") !== -1;


        if ( (timeDiff1 || timeDiff2 || flex) && !flexFull ) {

          // 20210209_0940_フレックスの記載方法が違っていた場合はキャンセルする処理
          timeAndConts = this.FlexTimeAndConts(conts, func);

          if ( !timeAndConts[2] && !timeDiff1 && !timeDiff2 ) status += "ﾌﾚ";
          if ( timeDiff1 ) status += "時差";
          if ( timeDiff2 ) status += "時短";

          
          // 直近のフレックスを予定取得時に実行
          if ( func === "thisMon" && t === 0 && !timeAndConts[2] ) {
            const getTimes = timeAndConts[0].split("〜");  // フレックスの開始終了時間を分割
            this.strFlexTime = getTimes[0];               // {obj} に直近のフレックス開始時間を追加
            this.endFlexTime = getTimes[1];               // {obj} に直近のフレックス終了時間を追加
            t++;  // 以降の予定のフレックス時間は取得しない。
          }
          
          status += `${timeAndConts[0]} ${timeAndConts[1]}`;
          
        }

        // 20210218_追加
        if ( (trip || goOut) && n === 0 ) {
          this.FlexTimeAndConts(conts, func);

          let decimalPoint0, decimalPoint5;
          let goOutTime = this.goOutTime;

          console.log("this.numJudge:" + this.numJudge);
   
          if ( this.numJudge ) {
            console.log("goOutTime:" + goOutTime);
            decimalPoint0  = goOutTime.indexOf('.0');        // 小数点の位置
            decimalPoint5  = goOutTime.indexOf('.5');        // 小数点の位置
            notGoTimeJudge = goOutTime == '';                // 外出時間設定なし判定

            if ( decimalPoint0 !== -1 ) {
              goOutTime = goOutTime.replace('.0', ':00');    // 小数点を含む場合、[.0]>>>[:00]へ変換
            } else if ( decimalPoint5 !== -1 ) {
              goOutTime = goOutTime.replace('.5', ':30');    // 小数点を含む場合、[.5]>>>[:30]へ変換
            } else if ( !notGoTimeJudge ) {
              goOutTime = goOutTime.replace(goOutTime, `${goOutTime}:00`);  // 小数点を含まない場合、[:00]を追加
            } else {
              goOutTime = undefined;
            }

            console.log("goOutTime:" + goOutTime);
            this.strGoTime = goOutTime;
          }
          n++;  // 出発時間確定(上書きを禁止するため)
        }

        this.notGoTimeJudge = notGoTimeJudge;

        // 配列[statusArr] に状態を格納
        statusArr.push(status);
      
        // 配列[offDayArr] に休日判定を格納
        offDayArr.push(offDays.some( offDay => conts.indexOf(offDay) !== -1 ));
            
        // 休日判定を実行
        const holiday = this.HolidayJudge(el, dateNum, conts, isoOffJudge);

        // 結果を配列に格納
        holArr.push(holiday);

        // 日付を1日進める
        i++; 
      
      })
      
    };
    
    
    
    
    
    /* ========================================================================= /
    /  ===  2種類の休日判定の配列を統合 [ メソッド ]                                 === /
    /  ======================================================================== */  
      
    OffHolIntegral(offDayJudges, holidayJudges, offHolJudges) {
      
      let i = 0; // 変数の初期化
      
      offDayJudges.forEach( el => {                  
        let offHol = "出勤";                                                  // 変数の初期化
        const offJudge = !( el === false );                                  // 休日判定(true/false)
        const holJudge = !( holidayJudges[i] === false );                    // 公休判定(true/false)
        if ( offJudge ) offHol = "休み";                                      // 休日判定がtrueの場合、休みを格納
        if ( holJudge ) offHol = "公休";
        offHolJudges.push(offHol);                                           // 配列[offHolJudges]に追加
        i++;                                                                 // 配列[holidayJudges]内の参照先を次の要素に切替
      });

    }

    
    
    
    /* ========================================================================= /
    /  ===  2種類の休日判定の配列を統合 [メソッド]                                   === /
    /  ======================================================================== */ 

    ArrayIntegral(monStatus, offHolJudges , statusArr) {

      let i = 0; // 変数の初期化
        
      monStatus.forEach( el => {
          
        // 休日予定
        const offHol = offHolJudges[i];
                          
        // 状態判定
        let   status;                                  // 状態書込先 
        const attend  = el === '在席';                  // 在席判定
        const offDay  = offHol === '休み';              // 休日判定
        const holDay  = offHol === '公休';              // 公休判定
        const goOut   = el.indexOf('外出') !== -1;      // 外出判定
        const trip    = el.indexOf('出張') !== -1;      // 出張判定
        const atHome  = el.indexOf('在宅') !== -1;      // 在宅判定
        const satDuty = el.indexOf('当番') !== -1;      // 当番判定
        const work    = el.indexOf('出勤') !== -1;      // 出勤判定
        const flex    = el.indexOf('ﾌﾚ') !== -1;        // フレックス判定
        const fullFlexJudge = el.indexOf(fullFlex) !== -1; // 完全フレックス判定
        const timeDiff1 = el.indexOf('時差') !== -1;
        const timeDiff2 = el.indexOf('時短') !== -1;


        // 状態判定によって変数[status]に格納
        if ( goOut ) {
          status = "外出";
        } else if ( trip ) {
          status = "出張";
        } else if ( atHome ) {
          status = "在宅";
        } else if ( satDuty ) {
          status = "当番";
        } else if ( work ) {
          status = "出勤";
        } else if ( offDay || fullFlexJudge ) {
          status = "休み";
        } else if ( holDay ) {
          status = "公休";
        } else if ( attend ) {
          status = "在席";
        } else {
          status = "在席";
        }
      
        // フレックス判定がtrueの場合
      if( flex || timeDiff1 || timeDiff2 ) status = el;

        // 配列[statusArr]に追加      
        statusArr.push(status);

        i++;  // 配列[offHolJudges]内の参照先を次の要素に切替
                            
      });

    }
    
    
    
    
    
    
    /* ========================================================================= /
    /  ===  休日パターン別の休日判定　[ メソッド ]                                   === /
    /  ======================================================================== */
    
    HolidayJudge( colors, dateNum, conts, offJudge ) {
    
      // 変数の初期化
      let holJudge = false;
      let holColor = false;
      let normalHol;         // 通常休み判定
      
//      // メンバーの休日パターンを取得(シート２)  ( シート左列のメンバー + シート2記入メンバー ) >>> 土日休み
      const normalHolMembers = membersNameL.concat(normalHolMems);    // 土日休みのメンバー(全員分・空白含む)
      
      // 通常の土日休みパターンの場合、 「true」 を返す
      normalHol = normalHolMembers.includes(this.name);
      
      // 曜日判定
      const saturday = dateNum === 6;   // 土曜日 判定
      const sunday   = dateNum === 0;   // 日曜日 判定
      
      // 当番判定
      const satDuty  = saturday && conts.indexOf('当番') !== -1; // 当番(当日)
      
      // 休日パターンのセルの背景色
      holColor = colors === "#d9d9d9" || colors === "#efefef" || colors === "#cccccc";
      
      // セル背景色が灰色 又は ISOWA休日 の場合は休日と判定
      if ( holColor || offJudge ) holJudge = true;

      // 通常休みパターンで土曜(当番以外)・日曜日の場合は休日と判定
      if ( normalHol ) {
        if ( sunday || saturday && !satDuty ) holJudge = true;
      }
      
      return holJudge;

    };


  
  
    /* ========================================================================= /
    /  ===  フレックスの時間を取得　[ メソッド ]                                  === /
    /  ======================================================================== */
    
    FlexTimeAndConts(conts, funcMonth) {
      
      // 変数の初期化
      let flexTime = '';
      let flexConts = '';
      let blankJudge = false;
      let numJudge = false;
      let compAndTime;
      let comparison;
      let goOutTime;
      let i = 0;

      // 予定の内容を[フレ]で前後分割して取得(フレックスは[ﾌﾚ][フレ]に対応)
      let _flexTime;

      const halfJudge = conts.indexOf('ﾌﾚ') !== -1;
      const fullJudge = conts.indexOf('フレ') !== -1;
      const timeDiff1 = conts.indexOf('時差') !== -1;
      const timeDiff2 = conts.indexOf('時短') !== -1;
      const goOut = conts.indexOf('外出') !== -1;      // 20200218_追加
      const trip  = conts.indexOf('出張') !== -1;      // 20200218_追加

      if ( halfJudge ) {
        _flexTime = conts.split('ﾌﾚ');
      } else if ( fullJudge ) {
        _flexTime  = conts.split('フレ');
      } else {
         console.log("フレックス：不明");
      }

      if ( timeDiff1 ) _flexTime = conts.split('時差');
      if ( timeDiff2 ) _flexTime = conts.split('時短');
      if ( goOut ) _flexTime = conts.split('外出');    // 20200218_追加
      if ( trip ) _flexTime  = conts.split('出張');    // 20200218_追加

      let above = _flexTime[0];             // 前半の文字列
      let below = _flexTime[1];             // 後半の文字列

      console.log("conts:" + conts);
      console.log("_flexTime:" + _flexTime);
      console.log("above:" + above);
      console.log("below:" + below);


      // フレックス記入方法が違っていた場合 20210209_0940_追加
      const cancelFlex = below.indexOf("ックス") !== -1 || below.indexOf("ｯｸｽ") !== -1;
      if ( cancelFlex ) below = '';

      // 出張・外出時に状態切替の時間を取得 20200218_追加
      if ( goOut || trip ) {
        const okNumber = [  1, 1.5, 2, 2.5, 3, 3.5, 4, 4.5, 5, 5.5, 6, 6.5, 7, 7.5,
                            8, 8.5, 9, 9.5, 10, 10.5, 11, 11.5, 12, 12.5, 13, 13.5,
                           14, 14.5, 15, 15.5, 16, 16.5, 17, 17.5, 18, 18.5, 19,
                           19.5, 20, 20.5, 21, 21.5, 22, 22.5, 23, 23.5, 24 ];

        blankJudge = above.search(/\s/) !== -1;  // 外出・出張の前の文字列に空白を含むかの判定

        // 空白を含む場合,訪問先と出発時間を取得
        if ( blankJudge ) {
          compAndTime = above.split(/\s/);                       // 訪問先と出発時間を分割
          comparison = compAndTime[0];                           // 訪問先
          goOutTime = compAndTime[1];                            // 出発時間
          numJudge = okNumber.indexOf(Number(goOutTime)) !== -1; // 出発時間判定(数字：ture) 

          // 出発時間判定がtrueの場合、変数[above]に時間を格納
          if ( numJudge ) {
            above = goOutTime;
          } else {
            comparison = '';
            above = '';
          }

          console.log("numJudge:" + numJudge);
          console.log(Number(goOutTime));
          console.log("Number(goOutTime):" + Number(goOutTime));
          console.log("above:" + above);
          console.log("comparison:" + comparison);
        }       
      }

      // オブジェクトに追加_20210218_追加
      console.log("funcMonth:" + funcMonth);
      if ( funcMonth === "thisMon" ) {
        this.numJudge = numJudge;
        this.goOutTime = goOutTime;
      }


      const flexAboBel = [ above, below ];    // 前半・後半の文字列を格納[配列]
            
      // 空文字判定(フレックス時間)
      const emptyStrAbo = above === "";       // 開始時間
      const emptyStrBel = below === "";       // 終了時間
      
      
      flexAboBel.forEach( el => {
                         
        let converted, reversedArray, reveresdString;
        let blankNum;
        let times = "";
        let getConts = "";
                         
        // 実行中の要素判定
        const aboveNow = el === above;
        const belowNow = el === below;
                         
        // 文字列を反対に並べ替える[前半の文字列のみ]
        if ( aboveNow ) {
          converted = el.split('');
          reversedArray = converted.reverse();
          reveresdString = reversedArray.join('');
        }
      
        // 文字列の一番目に空白があるか判定
        if ( aboveNow ) blankNum = reveresdString.search(/\s/); // 前半の文字列
        if ( belowNow ) blankNum = el.search(/\s/);             // 後半の文字列
      
        // 文字列から時間と予定をそれぞれ取得
        // 空白がない時
        if ( blankNum === -1 ) {
          times = el;
          
        // 空白が１文字目だった時 
        } else if ( blankNum === 0 ) {
          getConts = el;
          
        // 空白が1文字目以外の時(分割文字が前半の場合)
        } else if ( aboveNow ) {
          times = el.slice(-blankNum);
          getConts = el.slice(0, -blankNum);
          
        // 空白が1文字目以外の時(分割文字が後半の場合)
        } else if ( belowNow ) {
          times = el.slice(0, blankNum);
          getConts = el.slice(blankNum);          
        }
      
        if ( blankNum !== 0 ) times = this.Hankaku(times);      // 全角数字>>半角数字に変換
        if ( blankNum === 0 ) times = times.trim();             // スペースを削除
        const decimalPoint0 = times.indexOf('.0');              // 小数点の位置
        const decimalPoint5 = times.indexOf('.5');              // 小数点の位置

        if ( decimalPoint0 !== -1 ) {
          times = times.replace('.0', ':00');    // 小数点を含む場合、[.0]>>>[:00]へ変換
        } else if ( decimalPoint5 !== -1 ) {
          times = times.replace('.5', ':30');    // 小数点を含む場合、[.5]>>>[:30]へ変換
        } else {
          // 開始時間・終了時間が存在する場合に実行
          if ( aboveNow && !emptyStrAbo && blankNum || belowNow && !emptyStrBel && blankNum ) {
            times = times.replace(times, `${times}:00`);  // 小数点を含まない場合、timesに[:00]を追加
          }
        }
      
        // 時間を変数[flexTime]に格納     
        flexTime += times;
        if ( i === 0 ) flexTime += '〜';
      
        // 予定があった場合に情報を取得
        flexConts += getConts;
      
        // 予定が外出・出張時に不要な文字を削除
        const goOut = flexConts.indexOf('外出') !== -1;           // 外出判定
        if ( goOut ) flexConts = flexConts.replace('外出', '');   // 文字列[外出]を削除
        const trip = flexConts.indexOf('出張') !== -1;            // 出張判定      
        if ( trip ) flexConts = flexConts.replace('出張', '');    // 文字列[出張]を削除
        flexConts = flexConts.trim();                            // スペースを削除
        const atHome = flexConts.indexOf('在宅') !== -1;           // 在宅判定
        if ( atHome ) flexConts = flexConts.replace('在宅', '');   // 文字列[在宅]を削除
        i++;
      
      });

      let flexTimeConts = [ flexTime, flexConts, cancelFlex ]; // フレックス時間, 予定, キャンセル判定 を配列に格納
      
      return flexTimeConts;  // 配列を返す
      
    };


    /* ========================================================================= /
    /  ===  英数字を全角から半角に変換を取得　[ メソッド ]                             === /
    /  ======================================================================== */

    Hankaku(str) {
      return str.replace(/[０-９]/g, function(s) {
        return String.fromCharCode(s.charCodeAt(0) - 0xFEE0);
        });
    };




    /* ========================================================================= /
    /  ===  フレックスの実行時間範囲を取得　[ メソッド ]                               === /
    /  ======================================================================== */

    FlexTimeRange() {
      
      // 配列
      const strEndFlexTime = [ this.strFlexTime, this.endFlexTime ]; // 開始・終了時間を配列に格納
      const strGoOutTime = [ this.strGoTime ];                       // 外出・出張時の出発時間を配列に格納
      let strRange = [];                                             // 開始の範囲を初期化
      let endRange = [];                                             // 終了の範囲を初期化
      const npFlexTime = [ strRange, endRange ];                     // 開始・終了時間の許容範囲を格納
      let notExecStrEnd = [ false, false ];                          // 出社・帰宅トリガー実行判定
      
      // // 変数の初期化
      // let i = 0;         // 配列[pnFlexTime]の要素指定に使用

      // フレックスの場合に実行。
      if ( this.strFlexTime !== undefined || this.endFlexTime !== undefined ) notExecStrEnd = getFlexTime(strEndFlexTime);
      if ( this.strGoTime !== undefined ) getFlexTime(strGoOutTime);
      
      // ① 開始時間・終了時間の範囲をそれぞれ取得。
      // ② フレックス開始時間が、通常の出社時間より遅い場合の判定結果を取得。
      // ③ フレックス終了時間が、通常の帰宅時間より遅い場合の判定結果を取得
      function getFlexTime(strEndTime) {

        let setTime = strEndTime;

        let noExecStr = false;
        let noExecEnd = false;
        const noExec = [];
        
        console.log("setTime:" + setTime);
        // strEndFlexTime.forEach( el => {

        // 変数の初期化
        let i = 0;         // 配列[pnFlexTime]の要素指定に使用

        setTime.forEach( el => {

          // 分割 ( 時間 ・ 分 )
          // @ts-ignore
          console.log("el:" + el);
          const hourMinutes = el.split(':');
          let hours = Number(hourMinutes[0]);                // 時間
          const minutes = Number(hourMinutes[1]);            // 分

          // [-]範囲の初期値      
          let nHours   = 0;
          let nMinutes = 0;
      
          // [+]範囲の初期値
          let pHours   = 0;
          let pMinutes = 0;

          // [:00]分だった場合
          if ( minutes === 0 ) {
            nHours = hours - 1;                     // 時間を[hours - 1]
            if ( hours === 0 ) nHours = 23;         // 時間が0時だったら、23時に変更
            pHours = hours;
            nMinutes = 50;                          // 分を[50]
            pMinutes = 10;                          // 分を[10]
        
          // [:30]分だった場合
          } else if ( minutes === 30 ) {
            nHours   = hours;
            pHours   = hours;
            nMinutes = 20;                          // 分を[20]
            pMinutes = 40;                          // 分を[40]
          }

          console.log("時間範囲:" + nHours, nMinutes, pHours, pMinutes);
       
          // ① 配列に開始時間・終了時間の範囲を格納
          npFlexTime[i].push(nHours, nMinutes, pHours, pMinutes);   // 配列に追加

          // ② フレックス開始時間が8時より遅い場合の判定結果を格納
          if ( el === setTime[0] && hours >=  9 ) noExecStr = true;

          // ③ フレックス終了時間が18時より遅い場合の判定結果を格納
          if ( el === setTime[1] &&　hours >= 18 ) noExecEnd = true;
      
          i++; // 配列の追加先を変更

        });

        noExec.push(noExecStr, noExecEnd);

        return noExec;
      }
  
      // オブジェクトに追加
      this.strRange = strRange;
      this.endRange = endRange;
      this.notExecStr = notExecStrEnd[0];
      this.notExecEnd = notExecStrEnd[1];
  
    }
    
    
    
    /* ========================================================================= /
    /  ===  予定詳細の期間を記入 関数                                            === /
    /  ======================================================================== */       
    
    GetSchedPeriod() {
      
      // 変数の宣言
      let schedPeriod;       // メンバーの直近の予定詳細期間(当日〜2ヶ月)
      let nextSchedPeriod;   // メンバーの直近の予定詳細期間(翌日〜2ヶ月)
      
      // 予定詳細内容と期間を取得(例：12/1〜4 休み 等)
      details.forEach( el => {
        if ( el === this.name ) {
          schedPeriod     = this.AddSchedPeriod(0); // 当日を含む直近予定
          nextSchedPeriod = this.AddSchedPeriod(1); // 翌日以降の直近予定
        
          // オブジェクトに追加
          this.schedPeriod     = schedPeriod;
          this.nextSchedPeriod = nextSchedPeriod;
        }
      });
    
    };
    
    
    AddSchedPeriod(period) {

      // 変数の初期化
      let dayStatus      = '';
      let dayStatusFlex  = '';
      let latestSched    = '';
      let periodStart    = '';
      let periodEnd      = '';
      let writePeriod    = '';
      let twoMonthPlus      = [];
      let twoMonthSched     = [];
      let nextTwoMonthPlus  = [];
      let nextTwoMonthSched = [];
      let stateArr          = [];
      let todayStatus, todayContents;
      let twoMonthsP, twoMonthsS;
      let iStart = 0;
      let iEnd   = 0;
      let i = 0;
      let j = 0;
      let k = 0;
      let day;
      let blankNum;
      let getDestination = false;
      
      
      // 予定期間を取得(2ヶ月分の状態と予定)
      twoMonthPlus  = this.twoMonStatus;                             // 2ヶ月分の状態[配列]
      twoMonthSched = this.monContents.concat(this.nextMonContents); // 2ヶ月分の予定[配列]

      // ログ確認用(確定)
      console.log("twoMonthSched:" + twoMonthSched);
      
      // 取得期間が翌日〜で指定された場合、当日予定を帰宅に変更
      if ( period === 1 ) {
        k = 0; // 変数の初期化
        nextTwoMonthPlus = twoMonthPlus.map( el => {
          if ( k === 0 ) el = '帰宅';
          k++;
          return el;
        });

        k = 0; // 変数の初期化
        nextTwoMonthSched = twoMonthSched.map( el => {
          if ( k === 0 ) el = '帰宅';
          k++;
          return el;
        });
      }

      if ( period === 0 ) {
        twoMonthsP = twoMonthPlus;
        twoMonthsS = twoMonthSched;
      }

      if ( period === 1 ) {
        twoMonthsP = nextTwoMonthPlus;
        twoMonthsS = nextTwoMonthSched;
      }


      // 変数の初期化 (20210209)
      let _comparisons, comparisons, destination;
      let _destinationArr = [];
      let destinationArr  = [];


      twoMonthsP.forEach( el => {
                           
        // 条件判定
        const offDayJudge = el.indexOf('休み') !== -1;
        const goOut       = el.indexOf('外出') !== -1;
        const trip        = el.indexOf('出張') !== -1;
        const atHome      = el.indexOf('在宅') !== -1;
        const flex        = el.indexOf('ﾌﾚ') !== -1;
        const attend      = el.indexOf('在席') !== -1;
        const satDuty     = el.indexOf('当番') !== -1;
        const timeDiff1   = el.indexOf('時差') !== -1;
        const timeDiff2   = el.indexOf('時短') !== -1;

        // 外出 or 出張の場合、配列に行き先を格納(20210214)
        if ( goOut || trip ) {
          if ( goOut ) _comparisons = twoMonthsS[i].split("外出");
          if ( trip ) _comparisons = twoMonthsS[i].split("出張");
          comparisons = _comparisons[0].trim();

          let blankJudge = comparisons.search(/\s/) !== -1;  // 外出・出張の前の文字列に空白を含むかの判定

          // 空白を含む場合,訪問先と出発時間を取得し、訪問先のみ格納
          if ( blankJudge ) {
            const compAndTime = comparisons.split(/\s/);           // 訪問先と出発時間を分割
            comparisons = compAndTime[0];                          // 訪問先
          }

          if ( (dayStatus === el) && !getDestination ) _destinationArr.push(comparisons);
          console.log("_destinationArr:" + _destinationArr)
        }
                
        // 在席状態の格納条件
        const slctStatus = ( offDayJudge || goOut || trip || atHome || flex || timeDiff1 || timeDiff2 || satDuty ) && dayStatus === '' ;        
      
        // 予定が休み・外出・出張・在宅・フレックスの場合、[daystatus]に在席状態を格納
        if ( slctStatus ) {

          dayStatus = el;                             // 予定を格納
          if ( goOut || trip ) {
            _destinationArr.push(comparisons);          // 行き先を配列に格納
          }

          // フレックスの場合、状態とフレックス時間を分割して格納
          if ( (flex || timeDiff1 || timeDiff2) && !(goOut || trip) ) {
            // 在席以外の予定の場合
            if ( !attend ) {
              dayStatusFlex = dayStatus.slice(2);     // フレックス時間
              dayStatus     = dayStatus.slice(0, 2);  // 状態

            }
            // 在席の予定の場合
            if ( attend ) {
              dayStatusFlex = dayStatus.slice(2);     // フレックス時間
              dayStatus     = dayStatusFlex;          // 状態
            }
          }
        }

      
        // フレックスの場合に実行
        if ( (flex || timeDiff1 || timeDiff2) && !(goOut || trip) ) {
          if ( !attend ) el = el.slice(0, 2); // 在席以外の場合
          if ( attend ) el = el.slice(2);     // 在席の場合   
        }
         
        // 予定期間の開始日を取得
        if ( el === dayStatus && periodStart === '' ) {
                   
          // 日付を取得                     
          day = twoMonArr[i][0];
          
          // 開始日を書込
          periodStart = day;
          
          // 状態と予定を格納
          latestSched = twoMonthsS[i];
          stateArr = twoMonthsP;          
          j = i;
          iStart = j;
          
        }
            
        // 予定期間の終了日を取得
        if ( el !== dayStatus && periodStart !== '' && periodEnd === '') {
          console.log("予定終了");
          day = twoMonArr[i-1][0];
          periodEnd = day;                                        // 終了日を書込
          j = i;
          iEnd = j;
          getDestination = true;
          stateArr = stateArr.slice(iStart, iEnd);                // 配列[twoMonStatus]のi番目より後の要素を削除

          console.log("name:" + this.name);
          console.log("iEnd:" + iEnd);
          console.log("j:" + j);
          console.log("i:" + i);
          console.log("el:" + el);
          console.log("dayStatus:" + dayStatus);
          console.log("periodStart:" + periodStart);
          console.log("periodEnd:" + periodEnd);
          console.log("getDestination:" + getDestination);
          
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

      // 状態判定
      const offDayJudge = dayStatus.indexOf('公休') !== -1 || dayStatus.indexOf('休み') !== -1;
      const goOut       = dayStatus.indexOf('外出') !== -1;
      const trip        = dayStatus.indexOf('出張') !== -1;
      const atHome      = dayStatus.indexOf('在宅') !== -1;
      const flex        = dayStatusFlex.indexOf('ﾌﾚ') !== -1;
      // @ts-ignore
      const fullFlexJudge = dayStatus.indexOf(fullFlex) !== -1;
      const satDuty     = dayStatus.indexOf('当番') !== -1;
      const timeDiff1 = dayStatus.indexOf('時差') !== -1;
      const timeDiff2 = dayStatus.indexOf('時短') !== -1;

      // 予定期間の書込内容

      // 予定期間が １日 だった場合
      if ( stateArr.length <= 1 ) {
        writePeriod = `${periodStart}`;
          
      // 予定期間が 2日以上 + 同じ月の場合
      } else if ( startEnd ) {
        writePeriod = `${periodStart}〜${endDate[1]}`;
          
      // 予定期間が 2日以上 + 月をまたぐ場合 
      } else {
        writePeriod = `${periodStart}〜${periodEnd}`;
      }
        
      // 予定内容が 休み or 公休 だった場合
      if ( offDayJudge ) writePeriod += ' 休み';
        
      // 予定内容が 外出・出張 だった場合
      if ( (goOut || trip) && !flex && !timeDiff1 && !timeDiff2 ) {
        
        // 配列内の重複する行き先を削除
        destinationArr = _destinationArr.filter((value, index, self) => self.indexOf(value) === index);

        // 変数[destination]に行き先を追加
        destinationArr.forEach( el => {
          if ( destination === undefined ) {
            destination = el;
          } else {
            destination += `, ${el}`;
          }
        })

        // ログ確認用(確定)
        console.log("destinationArr:" + destinationArr);
        console.log("destination:" + destination);

        writePeriod += ` ${destination}`;

      }
                
      // 予定内容が 在宅 だった場合
      if ( atHome ) writePeriod += ' 在宅';
        
      // 予定内容が フレックス だった場合
      if ( (flex || timeDiff1 || timeDiff2) && !(goOut || trip) && !fullFlexJudge ) writePeriod += ` ${dayStatusFlex}`;

      // 時短の場合
      if ( timeDiff2 ) writePeriod = writePeriod.replace('時短','');

      // 予定内容が 当番 だった場合
      if ( satDuty ) writePeriod += ' 当番';


      return writePeriod;

    };

  };

  // ------------------------------------------------------------------------------------------------------------ //
  //      配列 [ memberObj ] にメンバー毎のオブジェクトを追加                                                            //
  // ------------------------------------------------------------------------------------------------------------ //
  
  // メンバーオブジェクトを格納する配列
  let membersObj = [];

  // 自動切替のシートに登録したメンバーの情報を取得 (オブジェクトを作成)
  targets.forEach( (target) => {
    
    // サービス予定表よりメンバーの行番号を取得 ( 本日 ・ 翌日 )
    const row  = members.indexOf(target) + 1;             // 行番号 （本日）
    const nextRow  = nextMembers.indexOf(target) + 1;     // 行番号 （翌日）

    // ログ確認用(確定)
    console.log("row(サービス作業予定表の行番号):" + row);

    // オブジェクト作成(予定表に名前があるメンバーのみ実行)
    if ( row > 0 && nextRow > 0) {
      const obj = new MemberObj(target, row, nextRow);       // オブジェクト{obj}作成
      obj.GetContents();                                     // {obj} に本日の予定を追加
      obj.GetSwitchSet();                                    // {obj} に切替設定を追加
      obj.GetSchedule();                                     // {obj} に2ヶ月分の在席状態を追加
      obj.GetRowColNum();                                    // {obj} に行・列番号・記入位置を追加
      obj.GetNowAttendance();                                // {obj} に在席状態と予定詳細を追加
      obj.GetSchedPeriod();                                  // {obj} に当日・翌日の予定(期間)を追加

      // ログ確認用(確定)
      console.log("メンバー:" + obj.name);

      // {obj}を[memberObj]に追加
      membersObj.push(obj);
    }

  });
  
  // ログ確認用(メンバー毎のオブジェクト)
  // console.log(membersObj);

  const object = [ membersObj, nightDutysData, satDutysData ];

  // 配列を返す >>> whiteData関数に情報を渡す
  return object; 
  
};
