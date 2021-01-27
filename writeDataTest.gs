// @ts-nocheck
/*
□ 機能追加予定
  ・ 土曜当番を予定詳細に追加する。 >>> 完了
  ・ 在席リストと切替設定の名前が一致していない場合にエラーにせずにスキップする。 >>> 完了
     切替設定と予定表の名前は一致している事。
     一致しない場合は、当日の予定が外出中・出張中の場合も予定が書き変わってしまう状態になる。
  ・ 切替設定と予定表の名前が一致していない場合、予定表に名前がない場合。
     >>> 在席リストに書き込みされない。 エラーにならないためこれでOK！

  ・ フレックス時間が18時以降の場合は、帰宅時のトリガーは実行しない様にする。  >>> 完了
  ・ 実行時間に時間が掛かり過ぎる。(約30〜50秒) >>> 20〜30秒 >>> 6〜7秒    >>> 完了
  ・ 24時間の場合、名前のセルを塗りつぶす。>> 8:30に切替を行う。



□ 本番前の注意点
  ・ 予定表と在席リストの漢字が違いエラーとなる事例あり。(福崎さん等)
  ・ 予定表に名前がないが在席リストに存在する人は予定表に追加する。
  ・ 平日休みの方の休日のセル背景色のパターンを統一したい。>>> 亀岡さんに確認
  ・ 在席リストの状態表示の文字列とプログラム内部の文字列が一致していることを確認
  ・ startTriggerの実行タイミングに注意

□ 自動切替のルール(取説)を作成

□ 確認事項
  ・ 自動切替の動作確認作業。
  
  　[状態パターン]
    在席
    在宅
    出張
    外出
    9フレ
    フレ15
    7フレ
    フレ18
    帰宅
    休み
    
    [予定パターン]
    公休日のみ
    休日のみ
    公休日・休日の混在
    出張               
    外出               
    在宅
    フレ
    連続フレ 同時間
    連続フレ 時間違い
  

□ バグ修正
   ・ 

*/


// ============================================================================================================ //
//       【関数】 取得した予定を在席リストに書込                                                                          //
// ============================================================================================================ //

function WriteDataTest(...membersObj) {
  
  // 現在の時間（△時）を取得
  const date = new Date();
  const nowHours = Utilities.formatDate(date, 'Asia/Tokyo', 'H');         // 現在の時間
  const nowMinutes = Utilities.formatDate(date, 'Asia/Tokyo', 'm');       // 現在の分
  const dayOfNum = date.getDay();                                         // 曜日番号
  date.setDate(date.getDate() + 1);                                       // 明日の日付をセット
  const nextDate = Utilities.formatDate(date, 'Asia/Tokyo', 'M/d');       // 明日の日付を取得
  
  // 曜日判定
  const saturday = dayOfNum === 6;   // 土曜日 判定
  const sunday   = dayOfNum === 0;   // 日曜日 判定
  
  // * メンバー情報を記入位置別で３つの配列に分ける。
  let memLeft   = [];
  let memCenter = [];
  let memRight  = [];
  
  
  // 配列 [ memberObj ] のメンバー情報 (1人ずつ情報を取得)
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
    const rowNum          = el.rowNum;                // 在席状態記入の行番号
    const colNum          = el.colNum;                // 在席状態記入の列番号
    const position        = el.position;              // 記入位置(左・中央・右)
    const detailNum       = el.detailNum;             // 予定詳細記入の列番号
    let   twoMonStatus    = el.twoMonStatus;          // 2ヶ月分の在席状態
    let   setContents     = el.setContents;           // 在席状態
    let   detail          = el.detail;                // 予定詳細内容
    const strFlexTime     = el.strFlexTime;           // フレックス開始時間
    const endFlexTime     = el.endFlexTime;           // フレックス終了時間
    const strRange        = el.strRange;              // フレックス開始範囲
    const endRange        = el.endRange;              // フレックス終了範囲
    const schedPeriod     = el.schedPeriod;           // 当日〜直近の予定
    const nextSchedPeriod = el.nextSchedPeriod;       // 翌日〜直近の予定
    const notExecStr      = el.notExecStr;            // フレックスの開始時間が 8時以降の判定( 8時以降で true)
    const notExecEnd      = el.notExecEnd;            // フレックスの終了時間が18時以降の判定(18時以降で true)
 
    console.log("名前:" + name);
    console.log("当日の直近予定:" + schedPeriod);
    console.log("翌日の直近予定:" + nextSchedPeriod);
  
  
    // フレックス開始範囲
    const nStrHours   = strRange[0];  // 開始 時間[-]範囲
    const nStrMinutes = strRange[1];  // 開始 　分[-]範囲
    const pStrHours   = strRange[2];  // 開始 時間[+]範囲
    const pStrMinutes = strRange[3];  // 開始 　分[+]範囲

    // フレックス終了範囲
    const nEndHours   = endRange[0];  // 終了 時間[-]範囲
    const nEndMinutes = endRange[1];  // 終了 　分[-]範囲
    const pEndHours   = endRange[2];  // 終了 時間[+]範囲
    const pEndMinutes = endRange[3];  // 終了 　分[+]範囲

    // 現在の状態判定
    const goOutNow = setContents === '外出中';
    const tripNow  = setContents === '出張中';
  
    // 当日の状態判定
    const offDayJudge = twoMonStatus[0].indexOf('休み') !== -1 
                     || twoMonStatus[0].indexOf('公休') !== -1;    
    const attend  = twoMonStatus[0].indexOf('在席') !== -1;
    const satDuty = twoMonStatus[0].indexOf('当番') !== -1;
    const goOut   = twoMonStatus[0].indexOf('外出') !== -1;
    const trip    = twoMonStatus[0].indexOf('出張') !== -1;
    const atHome  = twoMonStatus[0].indexOf('在宅') !== -1;
    const work    = twoMonStatus[0].indexOf('出勤') !== -1;
    const flex    = twoMonStatus[0].indexOf('ﾌﾚ') !== -1;

    // 出社トリガー実行条件 ( 当日がフレックス and 開始時間が9時以降 )
    const notStrJudge = notExecStr && flex;
    
    // 帰宅トリガー実行条件 ( 当日がフレックス and 終了時間が18時以降 )
    const notEndJudge = notExecEnd && flex;


    // 当日の宿直判定
    const nightSft = nightShtPatterns.some( nightShtPattern => contents.indexOf(nightShtPattern) !== -1 );
    
    // 状態判定(翌日)
    const nextOffDayJudge = twoMonStatus[1].indexOf('休み') !== -1 || twoMonStatus[1].indexOf('公休') !== -1;  // 休み
    const nextAttend      = twoMonStatus[1].indexOf('在席') !== -1;     // 在席
    const nextGoOut       = twoMonStatus[1].indexOf('外出') !== -1;     // 外出
    const nextTrip        = twoMonStatus[1].indexOf('出張') !== -1;     // 出張
    const nextWork        = twoMonStatus[1].indexOf('出勤') !== -1;     // 出勤  
    const nextFlex        = twoMonStatus[1].indexOf('ﾌﾚ') !== -1;       // フレックス
    const nextPaidHol     = offDays.some( paid => nextContents.indexOf(paid) !== -1 ); // 公休
  
    // 予定詳細内容の文言から 外出 ・ 出張 を削除
    if ( goOut ) contents = contents.replace('外出', '');              
    if ( trip ) contents = contents.replace('出張', '');              
    if ( nextGoOut ) nextContents = contents.replace('外出', '');      
    if ( nextTrip ) nextContents = nextContents.replace('出張', '');

  

    // 在席リストに書込む最終の状態・詳細項目を変数に格納
  
    // 出社時・退社時のプロジェクト実行判定
    const startTimer = ( nowHours ==  8 && 20 <= nowMinutes && nowMinutes <= 30 ); // 出社時
    const endTimer   = ( nowHours == 17 && 13 <= nowMinutes && nowMinutes <= 35 ); // 退社時

    console.log(nowHours, nowMinutes, startTimer, endTimer);

    // フレックス開始時のプロジェクト実行判定
    let strFlexTimer = false;
    
    // 実行時間が **:00 の場合
    if ( nStrMinutes > pStrMinutes ) {
      strFlexTimer = ( nowHours == nStrHours + 1 ) && ( nowHours == pStrHours ) && ( nowMinutes <= pStrMinutes );
    // 実行時間が **:00 以外の場合
    } else { 
      strFlexTimer = ( nowHours == nStrHours ) && ( nowHours == pStrHours ) && ( nowMinutes <= pStrMinutes );
    };
          
    // フレックス終了時のプロジェクト実行判定
    let endFlexTimer = false;

    // 実行時間が **:00 の場合
    if ( nEndMinutes > pEndMinutes ) {
      endFlexTimer = ( nowHours == nEndHours + 1 ) && ( nowHours == pEndHours ) && ( nowMinutes <= pEndMinutes );
    // 実行時間が **:00 以外の場合
    } else {
      endFlexTimer = ( nowHours == nEndHours ) && ( nowHours == pEndHours ) && ( nowMinutes <= pEndMinutes );
    };


    // 在席状態と当日の予定を書込( 出社時 or フレックス開始時 )
    starts.forEach( start => {
      if ( start === name ) {
        if ( startTimer ) StartWrite(schedPeriod);  // 通常出社時に実行
        if ( flex && strFlexTimer ) StartWrite(schedPeriod);       // フレックス開始時に実行
      }
    });

    // 在席状態と翌日の予定を書込( 帰宅時 or フレックス終了時 )
    ends.forEach( end => {
      if ( end === name ) {
        if ( endTimer && !notExecEnd ) EndWrite(nextSchedPeriod);  // 通常帰宅時に実行
        if ( flex && endFlexTimer ) EndWrite(nextSchedPeriod);     // フレックス終了時に実行
      }
    });

    // メンバー記入位置( 左 ・ 中央 ・ 右 )によって３つの配列に分類し、在席状態と予定詳細を格納
    if ( position === "L" ) memLeft.push([name, rowNum, setContents, detail]);
    if ( position === "C" ) memCenter.push([name, rowNum, setContents, detail]);
    if ( position === "R" ) memRight.push([name, rowNum, setContents, detail]);

    // console.log(memLeft, memCenter, memRight);


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
    function StartWrite(period) {
      
      const attendSel = attendSels.indexOf(name) !== -1; // 在席時の表示を東館するかの判定
//      console.log(attendSel);
      
      //  == 状態を書込 == //
      
      // [フレックス] >>> 状態を[フレックス]に変更
      if ( notStrJudge && !strFlexTimer ) {
        SetStatus("contents", 'フレックス');
        console.log("フレックス");

      // [在席] and [東館選択] >>> 状態を[東館]に変更 
      } else if ( attend && attendSel ) {
        SetStatus("contents", '東館');
        console.log("東館");

      // [在席] or [当番] or [出勤] >>> 状態を[在席]に変更
      } else if ( attend || satDuty || work ) {
        SetStatus("contents", '在席');
        console.log("在席");
        
      // [外出] >>> 状態を[外出]に変更
      } else if ( goOut ) {
        SetStatus("contents", '外出中');
        
      // [出張] >>> 状態を[出張]に変更
      } else if ( trip ) {
        SetStatus("contents", '出張中');
        
      // [在宅] >>> 状態を[在宅]に変更
      } else if ( atHome ) {
        SetStatus("contents", '在宅');
        
      // [休み] >>> 状態を[休み]に変更
      } else if ( offDayJudge ) {
        SetStatus("contents", '休み');
        console.log("休み");
      } else {
        console.log("不明");
      }
  
      // 詳細予定を書込 //
      details.forEach( el => {
        if ( el === name ) SetStatus("detail", period);
      });
      
    }


    /* ========================================================================= /
    /  ===  帰宅 or フレックス終了時に実行  関数                                    === /
    /  ======================================================================== */
    function EndWrite(period) {
      
      // == 状態を書込 == //
      console.log("name:" + name);
      console.log("setContents:" + setContents);
      console.log("detail:" + detail);
      console.log("notExecStr:" + notExecStr);
      console.log("notExecEnd:" + notExecEnd);
      
      // 現在の状態が[外出] or [出張]以外 >>> 状態を[帰宅]に変更 
      if ( !goOutNow && !tripNow ) {
        SetStatus("contents", '帰宅');
        
        // 当日 or 翌日 が[休日] >>> 状態を[休み]に変更
        if ( offDayJudge || nextOffDayJudge ) SetStatus("contents", '休み');

        // == 詳細予定を書込 == //
        details.forEach( el => SetStatus("detail", period));
      }
      

      
    }

  });



// ------------------------------------------------------------------------------------------------------------ //


  setInfo(); // 在席リストに情報を書込
  

  /* ========================================================================= /
  /  ===  在席リストに情報を書込   関数                                          === /
  /  ======================================================================== */
  function setInfo() {
    
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

      // ログ確認用(書込情報)
    //  console.log(memsArray);

      // メンバー名・行番号を削除
      memsArray.forEach( el => el.splice( 0, 2 ) );  

      // メンバー状態をスプレットシートに書込
      let row = maxNum - 2;
      attendList.getRange(3, mem[1], row, 2).setValues(memsArray);
 
    });
    
  }

}
