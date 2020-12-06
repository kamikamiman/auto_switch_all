/*


*/

// ============================================================================================================ //
//       【関数】 取得した予定を在席リストに書込                                                                          //
// ============================================================================================================ //

function WhiteDataTest(...membersObj) {
  
  // 現在の時間（△時）を取得
  const date = new Date();
  const nowTime  = Utilities.formatDate(date, 'Asia/Tokyo', 'H');   // 現在の時間
  const dayOfNum = date.getDay();                                   // 曜日番号
  date.setDate(date.getDate() + 1);                                 // 明日の日付をセット
  const nextDate = Utilities.formatDate(date, 'Asia/Tokyo', 'M/d'); // 明日の日付を取得
  
 
  // スプレットシートを取得（データ書込み用）
  const ssSet      = SpreadsheetApp.openById('1Itid9HCrW0wy_ATM4lqBDzkQs64DpemrWL4THmOsEIg'); // 在席リスト(ここにスプレットシートのアドレス記入)
  const attendList = ssSet.getSheetByName('当日在席(69期)');                                    // シート名よりシート情報  
  const lastRow    = attendList.getRange('C:C').getLastRow();                                 // 最終行番号
  const members    = attendList.getRange(1, 3, lastRow, 1).getValues().flat();                // メンバー情報
  

  // 「在席」への切替判定
  const attends = [ '', '24H', '当番' ];  // 在席判定
  
  // 直近の予定・状態 
  let setContents;                       // 直近の状態
  let detail;                            // 直近の予定
  
  
  
  
  // ------------------------------------------------------------------------------------------------------------ //
  //      配列 [ memberObj ] のメンバー情報から在席リストに予定を書き込んでいく                                                //
  // ------------------------------------------------------------------------------------------------------------ //
  
  membersObj.forEach( el => {
    
    // 配列[ target ] に記入したメンバーの情報
    const name = el.name;                                                              // メンバーの名前
    const rowNum = members.indexOf(name) + 1;                                          // メンバーの行番号
    const contents = el.contents;                                                      // 当日の予定
    const nextContents = el.nextContents;                                              // 翌日の予定
    setContents = attendList.getRange(rowNum, 5, 1, 1).getValue();                     // 在席リストの状態
  
    console.log(name, rowNum, contents, nextContents, setContents);                    // ログ確認用(メンバーの情報)

  
  
    // 休日 ・ 土曜 ・ 日曜 判定
    const offDays = [ '休み', '有給', '振休', '代休', 'RH', 'ＲＨ' ];                       // 休日パターンを設定
    const offDayJudge = offDays.some( offDay => contents.indexOf(offDay) !== -1 );     // 当日の休日判定
    const nextOffDayJudge = offDays.some( offDay => contents.indexOf(offDay) !== -1 ); // 翌日の休日判定
    const saturday = dayOfNum === 6;                                                   // 土曜日判定
    const sunday   = dayOfNum === 0;                                                   // 日曜日判定

  
    // 当番 ・ 外出 ・ 出張 判定
    const satDuty = contents.indexOf('当番') !== -1;       // 当番(当日)
    const goOut   = contents.indexOf('外出') !== -1;       // 外出(当日)
    const trip    = contents.indexOf('出張') !== -1;       // 出張(当日)
  
    const nextGoOut = nextContents.indexOf('外出') !== -1; // 外出(翌日)
    const nextTrip  = nextContents.indexOf('出張') !== -1; // 出張(翌日)
  

 
    // 出社時に当日の在席状態を書込
    starts.forEach( start => {
      if ( start === name && nowTime < 9 ) StartWrite();
    });

    // 帰宅時に翌日の在席状態を書込
    ends.forEach( end => {
      if ( end === name && nowTime < 15 ) EndWrite();
    });
      
    // 期間限定で発動(iサーチ打ち合わせ開始)
    mtgStarts.forEach( mtgStart => {
      if ( mtgStart === name && nowTime == 21 ) MtgStartWrite();
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
      if ( select === "contents" ) setContents = attendList.getRange(rowNum, 5, 1, 1).setValue(value);  
      
      // 詳細項目を書込
      if ( select === "detail" ) detail = attendList.getRange(rowNum, 6, 1, 1).setValue(value);
      
    }; 


/* ========================================================================= /
/  ===  始業時に実行する関数                                                 === /
/  ======================================================================== */
    function StartWrite() {

      // 予定無し ・ 24H ・ 当番 の場合、 「在席」を書込
      attends.forEach( attend => {
        if ( contents === attend ) {
          SetStatus("contents", '在席');
          if ( rowNum === 19 || rowNum === 28 ) cancel.push('可能');
        }
      });
  
      // 当日の休日判定がtrue ・ 日曜日 ・ 土曜日(当番でない) の場合、「休み」 を書込
      if ( offDayJudge || sunday || (saturday && !satDuty) ) SetStatus("contents", '休み');
      
      // 外出の場合、 「外出」を書込
      if ( goOut ) SetStatus("contents", '外出中');

      // 出張の場合、「出張」を書込
      if ( trip ) SetStatus("contents", '出張中');
 
      // 本日の予定が 外出 ・ 出張 の場合、予定表の内容を書込
      if ( goOut || trip ) {
        details.forEach( detail => {
          if ( detail === name ) SetStatus("detail", contents);
        });
      }
    }



/* ========================================================================= /
/  ===  終業時に実行する関数                                                 === /
/  ======================================================================== */
    function EndWrite() {

      // 翌日の予定が 出張以外 ・ 外出中 でなければ、「帰宅」 を書込
      if ( nextContents === !nextTrip || setContents !== '外出中' ) {
        SetStatus("contents", '帰宅');
        
        // 翌日の休日判定がtrueの場合、「休み」を書込
        if ( nextOffDayJudge ) SetStatus("contents", '休み');
        
      };
    
      // 翌日の予定が 休み の場合、[ 日付 + 休み ] を詳細項目に書込
      const dateHol = `${nextDate} 休み`;
      if ( nextOffDayJudge ) {
        details.forEach( detail => {
          SetStatus("detail", dateHol);
        });
      }
      
      // 翌日の予定が 出張 の場合、詳細項目に書込
      if ( nextTrip ) {
        SetStatus("contents", '出張中');
        details.forEach( detail => {
          SetStatus("detail", nextContents);
        });
      }

      // 翌日の予定が予定無し ・ 24H ・ 当番 の場合、予定を削除
      attends.forEach( attend => {
        if ( attend === nextContents ) {
          details.forEach( detail => {
            SetStatus("detail", '');
          });
        }
      });

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
