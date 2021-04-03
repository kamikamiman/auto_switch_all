// @ts-nocheck
/*
□ 機能追加予定
  ・ 休みの表示をそれぞれ分ける。                        >>> 自動・手動の方がいるのであまり意味がないので保留。(ソフトは作成済み)
  ・ ２４時間・土曜当番の表示色を変更。                   >>> 完了
  ・ ２４担当者の更新時間を追加。16時に追加               >>> 完了
  ・ 時差出勤・時短出勤に対応。 時差 時短                 >>> 完了
  ・ テストスイッチを取り付ける。                        >>> 完了
  ・ 出張・外出時の状態切替時の時間を変更機能の追加         >>> 完了
  ・ 半休(午前・午後)に対応                             >>> 完了


□ 注意点
  ・ 予定表と在席リストの漢字が違いエラーとなる事例あり。(福崎さん等)
  ・ 予定表に名前がないが在席リストに存在する人は予定表に追加する。
  ・ 平日休みの方の休日のセル背景色のパターンを統一したい。
  ・ startTriggerの実行タイミングに注意(タイミングを間違えると出社・帰社時のスクリプトが働かない可能性がある)

□ 確認事項
  ・ フレックス時に予定に誤ってフレックスと記載されていた場合の処理  >>> キャンセルする様に修正
  ・ 4月からサービス作業予定表が変更になるため、修正が必要。昼当番・２４時間・土曜当番


□ バグ修正・追加事項・動作確認
　・併用予定
　　・切り替わり確認(最低限) >>> 完了
　　・'/'なしの場合にエラーとなる。 >>> 完了
　　・'/'の記載方法のパターンを追加。 全角もOKとした。 >>> 完了
　　・在宅、半休の併用 >>> 完了
　　・在宅・時短・時差の併用 >>> 完了
　　・外出(出張)・フレックスの記入順番 >>> 完了
　　・予定の変更点を表示するログが一部ない。>>> 完了
　　・フレ／外出のパターンで外出に状態が切り替わらない。 >>> 完了　

*/


// ============================================================================================================ //
//       【関数】 取得した予定を在席リストに書込    更新日 (20210310)                                                  //
// ============================================================================================================ //

// 20210209_1435_公休日とそれ以外を判別するために状態を色分けする。
function WriteDataTest(getRowCol, getSatNames, ...membersObj) {
  
  // 現在の時間（△時）を取得
  const date = new Date();
  const nowHours = Utilities.formatDate(date, 'Asia/Tokyo', 'H');         // 現在の時間
  const nowMinutes = Utilities.formatDate(date, 'Asia/Tokyo', 'm');       // 現在の分
  const dayOfNum = date.getDay();                                         // 曜日番号
  date.setDate(date.getDate() + 1);                                       // 明日の日付をセット
  const nextDate = Utilities.formatDate(date, 'Asia/Tokyo', 'M/d');       // 明日の日付を取得

  // 出社時・退社時のプロジェクト実行判定
  const startTimer = ( nowHours ==  8 && 22 <= nowMinutes && nowMinutes <= 28 ); // 出社時
  const endTimer   = ( nowHours == 17 && 12 <= nowMinutes && nowMinutes <= 18 ); // 退社時

  // 24時間担当の再確認のプロジェクト実行判定
  const nightTimer = ( nowHours ==  15 && 57 <= nowMinutes && nowMinutes <= 03 ); // 出社時

  // 曜日判定
  const saturday = (dayOfNum === 6);   // 土曜日 判定
  const sunday   = (dayOfNum === 0);   // 日曜日 判定
  
  // * メンバー情報を記入位置別で３つの配列に分ける。
  let memLeft   = [];
  let memCenter = [];
  let memRight  = [];

  // メンバー情報を記入位置別で３つの配列に分ける。(変更前の情報)
  let memLeftLog   = [];
  let memCenterLog = [];
  let memRightLog  = [];

  // 変更点取得用の配列(ログ確認用)
  let rogLeftList   = [];
  let rogCenterList = [];
  let rogRightList  = [];
  let logLists = [];

  
  
  // 配列 [ memberObj ] のメンバー情報 (1人ずつ情報を取得)
  membersObj.forEach( el => {
    
    // 配列[ target ] に記入したメンバーの情報
    const name            = String(el.name);          // メンバーの名前
    let   contents        = String(el.contents);      // 当日の予定
    let   nextContents    = String(el.nextContents);  // 翌日の予定
    const rowNum          = el.rowNum;                // 在席状態記入の行番号
    const position        = el.position;              // 記入位置(左・中央・右)
    let   twoMonStatus    = el.twoMonStatus;          // 2ヶ月分の在席状態
    let   setContents     = el.setContents;           // 在席状態
    let   detail          = el.detail;                // 予定詳細内容
    const strRange        = el.strRange;              // フレックス開始範囲
    const endRange        = el.endRange;              // フレックス終了範囲
    const goStrRange      = el.goStrRange;            // 外出・出張開始範囲_20210314_追加
    const goEndRange      = el.goEndRange;            // 外出・出張終了範囲_20210314_追加
    const schedPeriod     = el.schedPeriod;           // 当日〜直近の予定
    const nextSchedPeriod = el.nextSchedPeriod;       // 翌日〜直近の予定
    const notExecStr      = el.notExecStr;            // フレックスの開始時間が 8時以降の判定( 8時以降で true)
    const notExecEnd      = el.notExecEnd;            // フレックスの終了時間が18時以降の判定(18時以降で true)
    const notGoTimeJudge  = el.notGoTimeJudge;        // 外出・出張時の出発時間指定なし判定_未使用?
    const numJudge        = el.numJudge;              // 外出・出張時の出発時間指定判定 

    // 現在の状態判定
    const goOutNow = setContents === '外出中';
    const tripNow  = setContents === '出張中';
  
    // 当日の状態判定
    const offDayJudge   = twoMonStatus[0].indexOf('休み') !== -1
    const pubHolJudge   = twoMonStatus[0].indexOf('公休') !== -1;    
    const attend        = twoMonStatus[0].indexOf('在席') !== -1;
    const satDuty       = twoMonStatus[0].indexOf('当番') !== -1;
    const goOut         = twoMonStatus[0].indexOf('外出') !== -1;
    const trip          = twoMonStatus[0].indexOf('出張') !== -1;
    const atHome        = twoMonStatus[0].indexOf('在宅') !== -1;
    const work          = twoMonStatus[0].indexOf('出勤') !== -1;
    const flex          = twoMonStatus[0].indexOf('ﾌﾚ') !== -1;
    const timeDiff1     = twoMonStatus[0].indexOf('時差') !== -1;
    const timeDiff2     = twoMonStatus[0].indexOf('時短') !== -1;
    const halfBeforeHol = twoMonStatus[0].indexOf('午前半休') !== -1;
    const halfAfterHol  = twoMonStatus[0].indexOf('午後半休') !== -1;

    // 状態判定(翌日)
    const nextOffDayJudge = twoMonStatus[1].indexOf('休み') !== -1      // 休み
    const nextPubHolJudge = twoMonStatus[1].indexOf('公休') !== -1;     // 公休
    const nextGoOut       = twoMonStatus[1].indexOf('外出') !== -1;     // 外出
    const nextTrip        = twoMonStatus[1].indexOf('出張') !== -1;     // 出張
    const nextTimeDiff1   = twoMonStatus[1].indexOf('時差') !== -1;     // 時差
    const nextTimeDiff2   = twoMonStatus[1].indexOf('時短') !== -1;     // 時短
    const nextHalfBefHol  = twoMonStatus[0].indexOf('午前半休') !== -1;  // 午前半休
    const nextHalfAftHol  = twoMonStatus[0].indexOf('午後半休') !== -1;  // 午後半休
  
    // 予定詳細内容の文言から 外出 ・ 出張 を削除
    if ( goOut ) contents = contents.replace('外出', '');              
    if ( nextGoOut ) nextContents = contents.replace('外出', '');      
    if ( trip ) contents = contents.replace('出張', '');              
    if ( nextTrip ) nextContents = nextContents.replace('出張', '');


    // 出社・帰社時の状態切替の実行時間範囲_20210314_更新
    const timeRanges   = [ strRange, endRange ];     // フレックス・時短・時差・半休
    const goTimeRanges = [ goStrRange, goEndRange];  // 外出・出張

    // 開始・終了実行判定
    let strEndFlexTimers = [];
    let strEndGoTimers = [];

    if ( flex || timeDiff1 || timeDiff2 || halfBeforeHol || halfAfterHol ) strEndFlexTimers = GetTimeRnage(timeRanges);   // フレックス・時短・時差・半休
    if ( goOut || trip ) strEndGoTimers = GetTimeRnage(goTimeRanges); // 外出・出張


    // 【関数】開始・終了時の実行時間範囲を取得
    function GetTimeRnage(getTimeRanges) {

      console.log("【GetTimeRange実行！】");

      console.log(`getTimeRanges: ${getTimeRanges}`);


      // 出社・帰社時の実行判定
      let strTimeJudge = false;
      let endTimeJudge = false;
      const strEndTimeJudges = []; // 最終的な strTimeJudge, endTimeJudgeの真偽値を格納

     // 変数の初期化
      i = 0;   // 0:開始範囲, 1:終了範囲

      // 開始範囲・終了範囲を順番に取り出し、情報を取得
      getTimeRanges.forEach( el => {

        // ログ確認用
        console.log(`el: ${el}`);
        console.log(`el[0]: ${el[0]}`);
        console.log(`el[1]: ${el[1]}`);
        console.log(`el[2]: ${el[2]}`);
        console.log(`el[3]: ${el[3]}`);

        // 実行時間が **:00 の場合
        if ( el[1] > el[3] ) {
          if ( i === 0 ) strTimeJudge = ( nowHours == el[0] + 1 ) && ( nowHours == el[2] ) && ( nowMinutes <= el[3] );  // 開始判定
          if ( i === 1 ) endTimeJudge = ( nowHours == el[0] + 1 ) && ( nowHours == el[2] ) && ( nowMinutes <= el[3] );  // 終了判定

        // 実行時間が **:00 以外の場合
        } else { 
          if ( i === 0 ) strTimeJudge = ( nowHours == el[0] ) && ( nowHours == el[2] ) && ( nowMinutes >= el[1] ) && ( nowMinutes <= el[3] );  // 開始判定
          if ( i === 1 ) endTimeJudge = ( nowHours == el[0] ) && ( nowHours == el[2] ) && ( nowMinutes >= el[1] ) && ( nowMinutes <= el[3] );  // 終了判定
        }

      i++;  // 取得情報を開始範囲 >>> 終了範囲へ切替

      });

      // 配列に開始・終了判定を格納
      strEndTimeJudges.push(strTimeJudge, endTimeJudge);

      return strEndTimeJudges;

    }; // 関数GetTimeRnage終了


    // メンバー記入位置( 左 ・ 中央 ・ 右 )によって３つの配列に分類し、在席状態と予定詳細を格納(ログ確認用)
    if ( position === "L" ) memLeftLog.push([name, rowNum, setContents, detail]);
    if ( position === "C" ) memCenterLog.push([name, rowNum, setContents, detail]);
    if ( position === "R" ) memRightLog.push([name, rowNum, setContents, detail]);


    // ログ確認用(確定)
    console.log("name(名前):" + name);
    console.log("schedPeriod(当日の直近予定):" + schedPeriod);
    console.log("nextSchedPeriod(翌日の直近予定):" + nextSchedPeriod);
    console.log("twoMonStatus(2ヶ月分の在席状態):" + twoMonStatus);
    console.log("numJudge(出発時間指定判定):" + numJudge);
    console.log("strRange(フレックス開始範囲):" + strRange );
    console.log("endRange(フレックス終了範囲):" + endRange );
    console.log("goStrRange(フレックス開始範囲):" + goStrRange );
    console.log("goEndRange(フレックス終了範囲):" + goEndRange );
    console.log(`strEndFlexTimers: ${strEndFlexTimers}`);
    console.log(`strEndGoTimers: ${strEndGoTimers}`);


    // console.log("午前半休判定:" + halfBeforeHol);
    // console.log("午後半休判定:" + halfAfterHol);


// -----   ここから在席状態と予定を書込関数実行   --------------------------------- //

  // 関数の実行判定
    const flexEtcJudge   = flex || timeDiff1 || timeDiff2;

    // 在席状態の切替条件
    const startJudge     = !(timeDiff1 || timeDiff2 || halfBeforeHol) && (startTimer || startButton);       // 出社時 8:25 ( 時短・時差・午前半休 以外 )
    const endJudge       = !(timeDiff1 || timeDiff2 || goOutNow || tripNow) && (endTimer || endButton);     // 帰宅時 17:15 ( 時短・時差・外出中・出張中 以外 )
    const flexStartJudge = (halfBeforeHol || flexEtcJudge || numJudge) && strEndFlexTimers[0];              // フレックス・時短・時差 出社時 ( 指定した時間 )
    const flexEndJudge   = (halfAfterHol || flexEtcJudge) && strEndFlexTimers[1];                           // フレックス・時短・時差 帰宅時 ( 指定した時間 )
    const goStartJudge   = ( goOut || trip ) && strEndGoTimers[0];                                          // 外出・出張時
    const allOffJudge    = offDayJudge || pubHolJudge || nextOffDayJudge || nextPubHolJudge;                // 休日 (当日・翌日)

    // 予定記入の切替条件
    const strDtailJudge  = startTimer || startButton;                                                       // 出社時 8:25


    // 在席状態と当日の予定を書込( 出社時 or フレックス開始時 )
    starts.forEach( start => {
      if ( start === name ) {

        // ログ確認用
        console.log("name:" + name);
        console.log("startJudge):" + startJudge);
        console.log("flexStartJudge:" + flexStartJudge);
        console.log("goStartJudge:" + goStartJudge);

        // 出社時に実行
        if ( startJudge || flexStartJudge || goStartJudge ) StartWrite();
      }
    });

    // 在席状態と翌日の予定を書込( 帰宅時 or フレックス終了時 )
    ends.forEach( end => {
      if ( end === name ) {

        // ログ確認用
        console.log("name:" + name);
        console.log("endJudge:" + endJudge);
        console.log("flexEndJudge:" + flexEndJudge);

        // 帰宅時に実行
        if ( endJudge || flexEndJudge) EndWrite();        
      }
    });

    // 詳細予定を書込
    details.forEach( el => {
      if (el === name ) {

        // ログ確認用
        console.log("name:" + name);
        console.log("strDtailJudge:" + startJudge);
        console.log("flexStartJudge:" + flexStartJudge);
        console.log("goStartJudge:" + goStartJudge);
        console.log("endJudge:" + endJudge);
        console.log("flexEndJudge:" + flexEndJudge);

        // 出社時に実行
        if ( strDtailJudge || flexStartJudge || goStartJudge ) DetailWrite(schedPeriod);

        // 帰社時に実行
        if ( endJudge || flexEndJudge ) DetailWrite(nextSchedPeriod);
      }
    });

// --------------------------------------------------------------------------- //



    // メンバー記入位置( 左 ・ 中央 ・ 右 )によって３つの配列に分類し、在席状態と予定詳細を格納
    if ( position === "L" ) memLeft.push([name, rowNum, setContents, detail]);
    if ( position === "C" ) memCenter.push([name, rowNum, setContents, detail]);
    if ( position === "R" ) memRight.push([name, rowNum, setContents, detail]);


    // メンバー状態の変更点を取得(ログ確認用)
    const leftLog   = [ memLeft, memLeftLog ];
    const centerLog = [ memCenter, memCenterLog ];
    const rightLog  = [ memRight, memRightLog];
    const allLog = [ leftLog, centerLog, rightLog ];
    logLists = [ rogLeftList, rogCenterList, rogRightList ];

    i = 0; // 変数の初期化

    // 状態の変更点を３列分取得
    allLog.forEach( el => {
      logLists[i] = [...el[0], ...el[1]].filter( value => 
                    !el[0].some(ary => ary.every((a, b) => a === value[b])) || 
                    !el[1].some(ary => ary.every((a, b) => a === value[b]))
                  );
      i++;  // 書込先を変更
    })


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
    /  ===  出社時 or フレックス開始時に実行  関数                                  === /
    /  ======================================================================== */
    function StartWrite() {
      
      const attendSel = attendSels.indexOf(name) !== -1; // 在席時の表示を東館するかの判定

      // データの書込(状態・予定記入)
      console.log("name:" + name);
      console.log("attend:" + attend);
      console.log("goOut:" + goOut);
      console.log("trip:" + trip);
      console.log("atHome:" + atHome);
      console.log("flex:" + flex);
      console.log("timeDiff1:" + timeDiff1);
      console.log("timeDiff2:" + timeDiff2);
      console.log("halfBeforeHol:" + halfBeforeHol);
      console.log("halfAfterHol:" + halfAfterHol);
      console.log("notExecStr:" + notExecStr);
      console.log("strEndGoTimers[0]:" + strEndGoTimers[0]);
      console.log("numJudge:" + numJudge);

      // [フレックス] >>> 状態を[フレックス]に変更
      if ( flex && notExecStr && !strEndFlexTimers[0] && !strEndGoTimers[0] ) {
        SetStatus("contents", 'フレックス');
        console.log('フレックス');

      // [在席] and [東館選択] >>> 状態を[東館]に変更 
      } else if ( attend && attendSel ) {
        SetStatus("contents", '東館');
        console.log('東館');

      // [在席] or [当番] or [出勤] >>> 状態を[在席]に変更
      } else if ( attend || satDuty || work ) {
        SetStatus("contents", '在席');
        console.log('在席1');

      // [外出] or [出張] and [出発時間指定あり] >>> 状態を[在席]に変更 ( 外出中・出張中は除く )
      } else if ( (goOut || trip) && numJudge && !strEndGoTimers[0] && !goOutNow && !tripNow ) {
        SetStatus("contents", '在席');
        console.log('在席2');

      // [外出] >>> 状態を[外出]に変更
      } else if ( goOut && (!numJudge || (numJudge && strEndGoTimers[0])) ) {
        SetStatus("contents", '外出中');
        console.log('外出中');
        
      // [出張] >>> 状態を[出張]に変更
      } else if ( trip && (!numJudge || (numJudge && strEndGoTimers[0])) ) {
        SetStatus("contents", '出張中');
        console.log('出張中');
        
      // [在宅] >>> 状態を[在宅]に変更
      } else if ( atHome ) {
        SetStatus("contents", '在宅');
        console.log('在宅');
        
      // [休み] >>> 状態を[休み]に変更
      } else if ( offDayJudge || pubHolJudge ) {
        SetStatus("contents", '休み');
        console.log('休み');

      // 上記以外の場合に表示  
      } else {
        console.log("在席状態：不明");
      }
      
    };


    /* ========================================================================= /
    /  ===  帰宅 or フレックス終了時に実行  関数                                    === /
    /  ======================================================================== */
    function EndWrite() {

      SetStatus("contents", '帰宅');      // データの書込(状態・予定記入)
      console.log('帰宅');

      if ( !flexEtcJudge && allOffJudge ) {
        SetStatus("contents", '休み');
        console.log('休み');
      }
    };

    /* ========================================================================= /
    /  ===  出社 or 帰宅 or フレックス開始・終了時に実行  関数   　                 === /
    /  ======================================================================== */
    function DetailWrite(period) {
      SetStatus("detail", period);
    };


  });

  // ログ確認用(確定)
  console.log("左側　変更前:" + memLeftLog);
  console.log("左側　変更後:" + memLeft);
  console.log("中央　変更前:" + memCenterLog);
  console.log("中央　変更後:" + memCenter);
  console.log("右側　変更前:" + memRightLog);
  console.log("右側　変更後:" + memRight);

  console.log("変更点:" + logLists);



// ------------------------------------------------------------------------------------------------------------ //


  SetInfo(); // 在席リストに情報を書込
  if ( !saturday && !sunday && (startTimer || startButton || nightTimer) ) SetDuty(getRowCol, "nightDuty");   // 夜勤担当者の名前のセルを塗りつぶす
  if ( saturday && (startTimer || startButton) ) SetDuty(getSatNames, "satDuty");                             // 土曜当番の名前のセルを塗りつぶす
  if ( saturday && (endTimer || endButton) ) ResetBackground();                                               // メンバーの名前のセルの背景色をリセット

  // ログ確認用(確定)
  console.log("getRowCol  (列・行番号　夜勤):" + getRowCol);
  console.log("getSatNames(列・行番号　土曜):" + getSatNames);


  /* ========================================================================= /
  /  ===  在席リストに情報を書込   関数                                         === /
  /  ======================================================================== */
  function SetInfo() {

    console.log("SetInfo実行！");
    
    // 在席リストの書込対象
    let mems = [];

　  // メンバー情報を格納する配列を定義
    const memInfoL = [ memLeft, posiL ];          // 在席リスト左側の列(メンバー情報)
    const memInfoC = [ memCenter, posiC ];        // 在席リスト中側の列(メンバー情報) 
    const memInfoR = [ memRight,posiR ];          // 在席リスト右側の列(メンバー情報)

    
    // 各列にメンバーが１人以上いた場合、配列[mems]に格納
    if ( memInfoL[0].length !== 0 ) mems.push(memInfoL);
    if ( memInfoC[0].length !== 0 ) mems.push(memInfoC);
    if ( memInfoR[0].length !== 0 ) mems.push(memInfoR);

    // 在席リストの左側・中央・右側の順にスプレットシートに在席状態と予定詳細内容を書込
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

      // 配列[nums]に最小値(minNum)〜最大値(maxNumの)数値を格納
      minNum = mem[0].slice(0).flat()[1];       // 最小値(書込行)
      maxNum = mem[0].slice(-1).flat()[1];      // 最大値(書込行)
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

      // ログ確認用(確定)
      console.log("書込内容");
      console.log(memsArray);


      // メンバー名・行番号を削除
      memsArray.forEach( el => el.splice(0, 2) );

      // メンバー状態をスプレットシートに書込
      let row = maxNum - 2;
      attendList.getRange(3, mem[1], row, 2).setValues(memsArray);
 
    });
    
  }



  /* ========================================================================= /
  /  ===  当番の名前のセルを塗りつぶす(土曜当番・２４時間) 関数                    === /
  /  ======================================================================== */
  function SetDuty(rowCol, select) {

    console.log("SetDuty実行!");

    // セルの背景色の初期化
    ResetBackground();

    // 担当者の在席リストの行・列番号(配列)
    const rowCol1 = rowCol[0];
    const rowCol2 = rowCol[1];

    // 担当者1人目の行・列番号
    const row1 = rowCol1[0];  // 行番号
    const col1 = rowCol1[1];  // 列番号

    // 担当者2人目の行・列番号
    const row2 = rowCol2[0];  // 行番号
    const col2 = rowCol2[1];  // 列番号


    // 担当者のセルの背景色を塗りつぶす(当番ありの場合)
    let setColor;
    if ( select === "nightDuty" ) setColor = "#FBB9EE";  // 24時間サービスの場合
    if ( select === "satDuty" ) setColor = "#FBB9EE";    // 土曜当番の場合

    const notDutyDay =  row1 === undefined || row2 === undefined;  // 当番無し判定

    if ( !notDutyDay ) {
    attendList.getRange(row1, col1, 1, 1).setBackground(setColor);
    attendList.getRange(row2, col2, 1, 1).setBackground(setColor);
    }

  };


  /* ========================================================================= /
  /  ===  セルの色をリセット 関数                                              === /
  /  ======================================================================== */
  function ResetBackground() {

    console.log("ResetBackground実行!");

    // セルの背景色の初期化
    const colL = posiL - 2;
    const colC = posiC - 2;
    const colR = posiR - 2;
    const setCols = [ colL, colC, colR ];
    const lastRow = attendList.getLastRow();
    setCols.forEach( setCol => attendList.getRange(3, setCol, lastRow, 1).setBackground("#ffffff"));


  };

};