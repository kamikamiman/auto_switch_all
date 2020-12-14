/* 機能追加予定
   ・ 在宅の設定を追加する。                >>> 完了
   ・ サービスのメンバーを追加する。           >>> 完了
   ・ 帰宅時に外出中の場合、詳細が消えてしまう。 >>> 完了
   ・ 公休日が土日以外のメンバーも対応する。    >>> 完了
   ・ スプレットシートから有効無効の切替を行える様にする。 >>> 完了    
   ・ 在席リストの詳細項目の外出・出張の文言を削除。 >>> 完了
   ・ ２４h当番で完全フレックス使用する場合に対応。  >>> 完了
   ・ 個別書込ではなく、全体書込で１回で書込みを行う。(実行時間短縮)
   ・ 休日の変更が正常に動作しない。holJudgeの判定がおかしい。 >>> 完了
   ・ 切替設定に登録してあるメンバー以外の記載があった場合、空白として更新されてしまう対処。  >>> 完了
   ・ 切替設定に名前があるが、予定表に名前がない場合にエラーとなる対処。  >>> 完了


   ・ 2ヶ月分の予定を取得する。

・ 休日・出張が連続する場合に詳細欄に予定を〇〇〜△△と記載できる様にする。
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
  
  // 現在の時間（△時）を取得
  const date = new Date();
  const nowTime  = Utilities.formatDate(date, 'Asia/Tokyo', 'H');   // 現在の時間
  const dayOfNum = date.getDay();                                   // 曜日番号
  date.setDate(date.getDate() + 1);                                 // 明日の日付をセット
  const nextDate = Utilities.formatDate(date, 'Asia/Tokyo', 'M/d'); // 明日の日付を取得
  
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
  
//  console.log(membersNameL, membersNameC, membersNameR);
  
  
  // 休日パターンを取得
  const normalHolMems = offDayList.getRange(1, 1, offLastRow, 1).getValues().flat();    // 土日休みのメンバーを取得(サービス)
  const normalHolMembers = membersNameL.concat(normalHolMems);                              // 土日休みのメンバー(全員分・空白含む)
//  console.log(normalHolMembers);


  // 在席 ・ 休日 への切替判定を設定 (サービス予定表に設定文字の記入があるか)
  const offDays = [ '休み', '有給', '振休', '振替', '代休', 'RH', 'ＲＨ', '完フレ', '完ﾌﾚ', '完全フレックス', '完全ﾌﾚｯｸｽ' ];  // 休日パターンを設定
  
  // 宿直パターンを設定(24時間サービス)
  const nightShtPatterns = [ '24', '24h', '24H', '２４', '２４ｈ', '２４Ｈ' ];
  
  // フレックスパターンを設定
  const flexPatterns = [ 'フレックス', 'ﾌﾚｯｸｽ', 'フレ', 'ﾌﾚ' ];
  
  // 曜日判定
  const friday   = dayOfNum === 5;   // 金曜日 判定
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
  
  
  // ------------------------------------------------------------------------------------------------------------ //
  //      配列 [ memberObj ] のメンバー情報から在席リストの書込情報を取得                                                    //
  // ------------------------------------------------------------------------------------------------------------ //
  
  membersObj.forEach( el => {
    
    // 配列[ target ] に記入したメンバーの情報
    const name         = String(el.name);          // メンバーの名前
    let   contents     = String(el.contents);      // 当日の予定
    let   nextContents = String(el.nextContents);  // 翌日の予定
    const color        = el.color;                 // 当日のセル背景色
    const nextColor    = el.nextColor;             // 翌日のセル背景色
    const swStart      = el.swStart;               // 出社時の切替設定
    const swEnd        = el.swEnd;                 // 退社時の切替設定
    const swDetail     = el.swDetail;              // 予定欄の切替設定
                     
                     
    // スプレットの記載情報を取得
    let rowNum;            // メンバーの行番号
    let colNum;            // 列番号(メンバーの在席状態の書込先)
    let detailNum;         // メンバー予定詳細の列番号
    let position;          // メンバー記入の位置  
  
    // 休日パターン ・ 休日判定
    let normalHol;         // 通常休み判定
    let holJudge;          // 当日の休日判定
    let nextHolJudge;      // 翌日の休日判定
  
    // 関数実行
    GetRowColNum();        // 行 ・ 列番号 ・ 記入位置 を取得
    HolidayJudge();        // 休日パターン別の休日判定
  
    // 直近の在席リストの状態と予定
    setContents = attendList.getRange(rowNum, colNum, 1, 1).getValue();   // 在席リストの状態
    detail = attendList.getRange(rowNum, detailNum, 1, 1).getValue();     // 予定詳細の状態
    
  
    // ログ確認用(メンバーの情報)
//    console.log(name, rowNum, colNum, detailNum, position, contents, nextContents, setContents, detail, swStart, swEnd, swDetail);

  
  
    // 休日 判定
    const offDayJudge     = offDays.some( offDay => contents.indexOf(offDay) !== -1 );     // 休日判定(当日)
    const nextOffDayJudge = offDays.some( offDay => nextContents.indexOf(offDay) !== -1 ); // 休日判定(翌日)

    // 宿直 判定
    const nightSft = nightShtPatterns.some( nightShtPattern => contents.indexOf(nightShtPattern) !== -1 );  // 当日の宿直判定

    // フレックス 判定
    const flex     = flexPatterns.some( flexPattern => contents.indexOf(flexPattern) !== -1 );      // フレックス判定(当日)
    const nextFlex = flexPatterns.some( flexPattern => nextContents.indexOf(flexPattern) !== -1 );  // フレックス判定(翌日)

  
    // 当番 ・ 外出 ・ 出張 ・ 在宅 判定
    const satDuty  = saturday && contents.indexOf('当番') !== -1;       // 当番(当日)
    const goOut    = contents.indexOf('外出') !== -1;                   // 外出(当日)
    const trip     = contents.indexOf('出張') !== -1;                   // 出張(当日)
    const atHome   = contents.indexOf('在宅') !== -1;                   // 在宅(当日)
      
    const nextGoOut = nextContents.indexOf('外出') !== -1;             // 外出(翌日)
    const nextTrip  = nextContents.indexOf('出張') !== -1;             // 出張(翌日)
  
    // 在席 判定
    const attend = !offDayJudge && !holJudge && !satDuty && !goOut && !trip && !atHome; // 在席判定(当日)
    const nextAttend = !nextOffDayJudge && !holJudge && !nextGoOut && !nextTrip;        // 在席判定(翌日)
  
  
    // 詳細項目の文言から 外出 ・ 出張 を削除
    if ( goOut ) contents = contents.replace('外出', '');              
    if ( trip ) contents = contents.replace('出張', '');              
    if ( nextGoOut ) nextContents = contents.replace('外出', '');      
    if ( nextTrip ) nextContents = nextContents.replace('出張', '');
  
  
  

    /* ========================================================================= /
    /  ===  在席リストの状態・詳細項目を書込　関数を実行                                === /
    /  ======================================================================== */  
  
    // プロジェクト実行時間の設定
    const startTimer = nowTime == 8 && !flex;
    const endTimer = nowTime == 23 && !flex; 
    const flexTimer = contents === flex;
  
    // 出社時に当日の在席状態を書込
    starts.forEach( start => {
      if ( start === name && startTimer ) StartWrite();
    });

    // 帰宅時に翌日の在席状態を書込
    ends.forEach( end => {
      if ( end === name && endTimer ) EndWrite();
    });
      
    // 期間限定で発動(iサーチ打ち合わせ開始)
    mtgStarts.forEach( mtgStart => {
      if ( mtgStart === name && nowTime == 10 ) MtgStartWrite();
    });

    // 期間限定で発動(iサーチ打ち合わせ終了)
    mtgEnds.forEach( mtgEnd => {
      if ( mtgEnd === name && nowTime == 11 ) MtgEndWrite();
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
    
    function HolidayJudge() {
    
      // 通常の土日休みパターンの場合、 「true」 を返す
      normalHol = normalHolMembers.includes(name);
    
      // 休日パターンのセルの背景色
      holColor = color === "#d9d9d9" || color === "#efefef" || color === "#cccccc";
      nextHolColor = nextColor === "#d9d9d9" || nextColor === "#efefef" || nextColor === "#cccccc";
//      console.log(normalHol, holColor);
      
      // 平日パターンでセル背景色が灰色の場合は休日と判定
//      if ( !normalHol ) {
        if ( holColor ) {
          holJudge = true;
        } else {
          holJudge = false;        
        }

        if ( nextHolColor ) {
          nextHolJudge = true;
        } else {
          nextHolJudge = false;        
        }
//      }

      // 通常休みパターンで土曜(当番以外)・日曜日の場合は休日と判定
      if ( normalHol ) {
        if ( saturday || sunday && !satDuty ) {
          holJudge = true;
        } else {
          holJudge = false;
        }

        if ( friday || saturday && !satDuty ) {
          nextHolJudge = true;
        } else {
          nextHolJudge = false;
        }
      }

      // ログ確認用
      console.log(name, normalHol, holJudge, nextHolJudge, color, nextColor);

    }


    /* ========================================================================= /
    /  ===  始業時に実行   関数                                                 === /
    /  ======================================================================== */
    function StartWrite() {
      
      // 休日 ・ 外出 ・ 出張 以外の場合、 「在席」を書込
      if ( attend ) SetStatus("contents", '在席');
  
      // 当日の休日判定がtrue ・ 日曜日 ・ 土曜日(当番でない) の場合、「休み」 を書込
      if ( offDayJudge || holJudge ) SetStatus("contents", '休み');
      
      // 宿直の場合、予定表の内容を書込
      if ( nightSft ) SetStatus("detail", contents);
      
      // 外出の場合、 「外出」を書込
      if ( goOut ) SetStatus("contents", '外出中');

      // 出張の場合、「出張」を書込
      if ( trip ) SetStatus("contents", '出張中');
 
      // 外出 ・ 出張 の場合、予定表の内容を書込
      if ( goOut || trip ) {
        details.forEach( detail => {
          if ( detail === name ) SetStatus("detail", contents);
        });
      }

      // 在宅の場合、「在宅」を書込
      if ( atHome ) SetStatus("contents", '在宅');




    }



    /* ========================================================================= /
    /  ===  終業時に実行   関数                                                 === /
    /  ======================================================================== */
    function EndWrite() {

      // 翌日が 出張以外 ・ 外出中 でなければ実行
      if ( nextContents === !nextTrip || setContents !== '外出中' ) {
        
        // 「帰宅」 を書込
        SetStatus("contents", '帰宅');
        
        // 休日の場合、「休み」を書込
        if ( nextOffDayJudge || nextHolJudge ) SetStatus("contents", '休み');
        
      };
    
      // 翌日が 休み の場合、[ 日付 + 休み ] を詳細項目に書込
      const dateHol = `${nextDate} 休み`; // 記入文字
      
      // 公休日以外の休日判定で実行
      if ( nextOffDayJudge ) {
        details.forEach( detail => {
          SetStatus("detail", dateHol);
        });
      }
      
      // 翌日が 出張 の場合、状態を「出張中」、詳細項目に書込
      if ( nextTrip ) {
        SetStatus("contents", '出張中');
        details.forEach( detail => {
          SetStatus("detail", nextContents);
        });
      }

      // 翌日が休日 ・ 外出 ・ 出張 以外 かつ 当日の状態が「外出中」でない場合は予定を削除
        if ( nextAttend && setContents !== "外出中" ) {
          details.forEach( detail => {
            SetStatus("detail", '');
          });
        }

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
    let nums = [];           // 行番号の存在チェック用配列
    let memNum = [];         // メンバーの行番号の配列
    let _blankArray = [];    // 存在しない行番号の配列
    let blankArray = [];     // 存在しない行番号の空白を含んだ配列(mem[0]に書込用)
    let memsArray = [];      // メンバー(空白を含んだ)情報を格納する配列
    let blankContents;


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
      console.log(blankContents);
    
    });

//    // 配列[blankArray]に空情報を格納
//    _blankArray.forEach( el => blankArray.push(['', el, '', '']) );

    // 配列[blankArray]を配列[mem[0]]に格納
    memsArray = mem[0].concat(blankArray);
    
    // 配列を行番号の昇順に並べ替え
    memsArray.sort((a, b) => a[1] - b[1]);

    // ログ確認用(書込情報)
    console.log(memsArray);

    // メンバー名・行番号を削除
    memsArray.forEach( el => el.splice( 0, 2 ) );  

    // メンバー状態をスプレットシートに書込
    const row = maxNum - 2;
    attendList.getRange(3, mem[1], row, 2).setValues(memsArray);
 
  });


    
    
    
    
    
  }




  // 在席リストの情報を更新
//  attendList.getRange(3,  5, lastRowL, 2).setValues(memLeft);  // メンバー情報(在席リスト左)
//  attendList.getRange(3, 10, lastRowC, 2).setValues(memCenter);  // メンバー情報(在席リスト中)
//  attendList.getRange(3, 15, lastRowR, 2).setValues(memRight);  // メンバー情報(在席リスト右)  

}
