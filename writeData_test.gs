/* 機能追加予定
   ・ 在宅の設定を追加する。                >>> 完了
   ・ サービスのメンバーを追加する。           >>> 完了
   ・ 帰宅時に外出中の場合、詳細が消えてしまう。 >>> 完了
   ・ 公休日が土日以外のメンバーも対応する。    >>> 完了
   
   ・ スプレットシートから有効無効の切替を行える様にする。
   ・ 休日・出張が連続する場合に詳細欄に予定を〇〇〜△△と記載できる様にする。
   ・ フレックスに対応する。
   ・ 

*/


// ============================================================================================================ //
//       【関数】 取得した予定を在席リストに書込                                                                          //
// ============================================================================================================ //

function WriteDataTest(...membersObj) {
  
  // 現在の時間（△時）を取得
  const date = new Date();
  const nowTime  = Utilities.formatDate(date, 'Asia/Tokyo', 'H');   // 現在の時間
  const dayOfNum = date.getDay();                                   // 曜日番号
  date.setDate(date.getDate() + 1);                                 // 明日の日付をセット
  const nextDate = Utilities.formatDate(date, 'Asia/Tokyo', 'M/d'); // 明日の日付を取得
  
 
  // スプレットシートを取得（データ書込み用）
  const ssSet      = SpreadsheetApp.openById('1Itid9HCrW0wy_ATM4lqBDzkQs64DpemrWL4THmOsEIg');             // 在席リスト(ここにスプレットシートのアドレス記入)
  const attendList = ssSet.getSheetByName('当日在席(69期)');                                                // シート1
  const lastRow    = attendList.getRange('C:C').getLastRow();                                             // シート1の最終行番号
  const offDayList = ssSet.getSheetByName('69期サービス土日休み');                                            // シート2
  const offLastRow = offDayList.getRange(1, 1).getNextDataCell(SpreadsheetApp.Direction.DOWN).getRow();   // シート2の最終行番号
  
  // 在席リストのメンバー情報を取得
  const membersL   = attendList.getRange(1,  3, lastRow, 1).getValues().flat();  // メンバー情報(在席リスト左)
  const membersC   = attendList.getRange(1,  8, lastRow, 1).getValues().flat();  // メンバー情報(在席リスト中)
  const membersR   = attendList.getRange(1, 13, lastRow, 1).getValues().flat();  // メンバー情報(在席リスト右)
  const members = [ membersL, membersC, membersR ];                              // メンバー情報(全員分)
  
  // 休日パターンを取得
  const normalHolMems = offDayList.getRange(1, 1, offLastRow, 1).getValues().flat();    // 土日休みのメンバーを取得(サービス)
  const normalHolMembers = membersL.concat(normalHolMems);                              // 土日休みのメンバー(全員分・空白含む)


  // 在席 ・ 休日 への切替判定を設定 (サービス予定表に設定文字の記入があるか)
  const offDays = [ '休み', '有給', '振休', '代休', 'RH', 'ＲＨ' ];  // 休日パターンを設定
  
  // 曜日判定
  const friday   = dayOfNum === 5;   // 金曜日 判定
  const saturday = dayOfNum === 6;   // 土曜日 判定
  const sunday   = dayOfNum === 0;   // 日曜日 判定
  
  // 直近の予定・状態 
  let setContents;  // 直近の状態
  let detail;       // 直近の予定
  
  
  
  
  // ------------------------------------------------------------------------------------------------------------ //
  //      配列 [ memberObj ] のメンバー情報から在席リストに予定を書き込んでいく                                                //
  // ------------------------------------------------------------------------------------------------------------ //
  
  membersObj.forEach( el => {
    
    // 配列[ target ] に記入したメンバーの情報
    const name = el.name;  // メンバーの名前
    const contents = el.contents;           // 当日の予定
    const nextContents = el.nextContents;   // 翌日の予定
    const color = el.color;                 // 当日のセル背景色
    const nextColor = el.nextColor;         // 翌日のセル背景色
                     
                     
    // スプレットの記載情報を取得
    let rowNum;            // メンバーの行番号
    let colNum;            // 列番号(メンバーの在席状態の書込先)
    let detailNum;         // メンバー予定詳細の列番号
    let position;          // メンバー記入の位置

    GetRowColNum();        // 行 ・ 列番号 ・ 記入位置 を取得
  
  
    // 休日パターン ・ 休日判定
    let normalHol;    // 通常休み判定
    let holJudge;     // 当日の休日判定
    let nextHolJudge; // 翌日の休日判定
  
    HolidayJudge();   // 休日パターン別の休日判定
  
 
    setContents = attendList.getRange(rowNum, colNum, 1, 1).getValue();   // 在席リストの状態
  
    // ログ確認用(メンバーの情報)
//    console.log(name, rowNum, colNum, detailNum, position, contents, nextContents, setContents);

  
  
    // 休日 判定
    const offDayJudge = offDays.some( offDay => contents.indexOf(offDay) !== -1 );         // 当日の休日判定
    const nextOffDayJudge = offDays.some( offDay => nextContents.indexOf(offDay) !== -1 ); // 翌日の休日判定

  
  
    // 当番 ・ 外出 ・ 出張 ・ 在宅 判定
    const satDuty = saturday && contents.indexOf('当番') !== -1;       // 当番(当日)
    const goOut   = contents.indexOf('外出') !== -1;                   // 外出(当日)
    const trip    = contents.indexOf('出張') !== -1;                   // 出張(当日)
    const atHome  = contents.indexOf('在宅') !== -1;                   // 在宅(当日)
      
    const nextGoOut = nextContents.indexOf('外出') !== -1;             // 外出(翌日)
    const nextTrip  = nextContents.indexOf('出張') !== -1;             // 出張(翌日)
  
    const attend = !offDayJudge && !holJudge && !satDuty && !goOut && !trip && !atHome; // 在席判定(当日)
    const nextAttend = !nextOffDayJudge && !holJudge && !nextGoOut && !nextTrip;        // 在席判定(翌日)
    

 
    // 出社時に当日の在席状態を書込
    starts.forEach( start => {
      if ( start === name && nowTime < 9 ) StartWrite();
    });

    // 帰宅時に翌日の在席状態を書込
    ends.forEach( end => {
      if ( end === name && nowTime > 16 ) EndWrite();
    });
      
    // 期間限定で発動(iサーチ打ち合わせ開始)
    mtgStarts.forEach( mtgStart => {
      if ( mtgStart === name && nowTime == 10 ) MtgStartWrite();
    });

    // 期間限定で発動(iサーチ打ち合わせ終了)
    mtgEnds.forEach( mtgEnd => {
      if ( mtgEnd === name && nowTime == 11 ) MtgEndWrite();
    });


/* ========================================================================= /
/  ===  在席リストの状態・詳細項目を書込　関数                                    === /
/  ======================================================================== */
    function SetStatus(select, value) {
      
      // 状態を書込
      if ( select === "contents" ) setContents = attendList.getRange(rowNum, colNum, 1, 1).setValue(value);  
      
      // 詳細項目を書込
      if ( select === "detail" ) detail = attendList.getRange(rowNum, detailNum, 1, 1).setValue(value);
      
    }; 


/* ========================================================================= /
/  ===  行 ・ 列番号 ・ 記入位置 を取得　関数                                   === /
/  ======================================================================== */

    function GetRowColNum() {
    
      // メンバーの行番号と記入位置を取得
      members.forEach( member => {
                      
        if ( member.indexOf(name) !== -1 ) {
          rowNum = member.indexOf(name) + 1; // 行番号を取得
        
          if ( member === membersL ) {
            colNum =  5;
            position = "L";
          }
        
          if ( member === membersC ) {
            colNum = 10;
            position = "C";
          }
          
          if ( member === membersR ) {
            colNum = 15;
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
    
      // 平日パターンでセル背景色が灰色の場合は休日と判定
      if ( !normalHol ) {
        if ( color === "#d9d9d9" ) {
          holJudge = true;
        } else {
          holJudge = false;        
        }

        if ( nextColor === "#d9d9d9" ) {
          nextHolJudge = true;
        } else {
          nextHolJudge = false;        
        }
      }

      // 通常休みパターンで土曜(当番以外)・日曜日の場合は休日と判定
      if ( normalHol ) {
        if ( saturday && sunday && !satDuty ) {
          holJudge = true;
        } else {
          holJudge = false;
        }

        if ( friday && saturday && !satDuty ) {
          nextHolJudge = true;
        } else {
          nextHolJudge = false;
        }
      }

      // ログ確認用
//      console.log(name, normalHol, holJudge, nextHolJudge, color, nextColor);

    }


/* ========================================================================= /
/  ===  始業時に実行する関数                                                 === /
/  ======================================================================== */
    function StartWrite() {
      
      // 休日 ・ 外出 ・ 出張 以外の場合、 「在席」を書込
      if ( attend ) SetStatus("contents", '在席');
  
      // 当日の休日判定がtrue ・ 日曜日 ・ 土曜日(当番でない) の場合、「休み」 を書込
      if ( offDayJudge || holJudge ) SetStatus("contents", '休み');
      
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
/  ===  終業時に実行する関数                                                 === /
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
/  ===  会議開始時に実行する関数                                              === /
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
/  ===  会議終了時に実行する関数                                              === /
/  ======================================================================== */
    function MtgEndWrite() {
      if ( setContents === '会議中' ) {
        SetStatus("contents", '在席');
        if ( detail !== '' ) SetStatus("detail", '');
      }
    }


  });


}
