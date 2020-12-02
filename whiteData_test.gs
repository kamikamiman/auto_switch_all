function WhiteDataTest(...membersObj) {
  
    // 現在の時間（△時）を取得
  const date = new Date();
  const nowTime  = Utilities.formatDate(date, 'Asia/Tokyo', 'H');   // 現在の時間
  const dayOfNum = date.getDay();                                   // 曜日番号
  date.setDate(date.getDate() + 1);                                 // 明日の日付をセット
  const nextDate = Utilities.formatDate(date, 'Asia/Tokyo', 'M/d'); // 明日の日付を取得
  
 
  // スプレットシートを取得（データ書込み用）
  const ssSet      = SpreadsheetApp.openById('1Itid9HCrW0wy_ATM4lqBDzkQs64DpemrWL4THmOsEIg'); // 在席リスト
  const attendList = ssSet.getSheetByName('当日在席(69期)');                     // シート名よりシート情報  
  const lastRow    = attendList.getRange('C:C').getLastRow();                  // 最終列番号
  const members    = attendList.getRange(1, 3, lastRow, 1).getValues().flat(); // メンバー情報
  
  

  
  // 「在席」への切替判定
  const attend  = [ '', '24H', '当番' ];        // 在席判定
  
 
  let detail; // 直近の予定
  let setContents; // 直近の状態
  
  membersObj.forEach( el => {
    const rowNum       = members.indexOf(el.name) + 1; // メンバーの行番号
    const contents     = el.contents;                  // 当日の予定
    const nextContents = el.nextContents;              // 翌日の予定
    
    setContents = attendList.getRange(rowNum, 5, 1, 1).getValue(); // 在席リストの状態
  
    console.log(rowNum, contents, nextContents, setContents);
  
    // 当日の休日パターン
    const holiday1 = contents.indexOf('休み') !== -1;
    const holiday2 = contents.indexOf('有給') !== -1;
    const holiday3 = contents.indexOf('振休') !== -1;
    const holiday4 = contents.indexOf('代休') !== -1;
    const holiday5 = contents.indexOf('RH')  !== -1;
  
    // 翌日の休日パターン
    const nexthol1 = nextContents.indexOf('休み') !== -1;
    const nexthol2 = nextContents.indexOf('有給') !== -1;
    const nexthol3 = nextContents.indexOf('振休') !== -1;
    const nexthol4 = nextContents.indexOf('代休') !== -1;
    const nexthol5 = nextContents.indexOf('RH')  !== -1;

  
    // 土曜当番 ・ 外出 ・ 出張のパターン
    const satDuty = contents.indexOf('当番') !== -1;              // 当日土曜当番
    const goOut        = contents.indexOf('外出') !== -1;         // 当日外出
    const businessTrip = contents.indexOf('出張') !== -1;         // 当日出張
    const nextGoOut        = nextContents.indexOf('外出') !== -1; // 翌日外出
    const nextBusinessTrip = nextContents.indexOf('出張') !== -1; // 翌日出張
  
    // 土曜 ・ 日曜 ・ 休日判定
    const holidayJudg = holiday1 || holiday2 || holiday3 || holiday4 || holiday5; // 当日休日判定(trueで休日)
    const nextholJudg = nexthol1 || nexthol2 || nexthol3 || nexthol4 || nexthol5; // 翌日休日判定(trueで休日)
    const saturday = dayOfNum === 6;  // 土曜日判定
    const sunday   = dayOfNum === 0;  // 日曜日判定

 
    // 出社時に当日の在席状態を書込
    if ( nowTime < 6 ) {
      // 予定無し ・ 24H ・ 当番 の場合、 「在席」を書込
      attend.forEach( el2 => {
        if ( contents === el2 ) setContents = attendList.getRange(rowNum, 5, 1, 1).setValue('在席');
      });
  
      // 当日の休日判定がtrue ・ 日曜日 ・ 土曜日(当番でない) の場合、「休み」 を書込
      if ( holidayJudg || sunday ) setContents = attendList.getRange(rowNum, 5, 1, 1).setValue('休み');
      if ( saturday && !satDuty )  setContents = attendList.getRange(rowNum, 5, 1, 1).setValue('休み');
    
      // 外出の場合、 「外出」を書込
      if ( goOut )        setContents = attendList.getRange(rowNum, 5, 1, 1).setValue('外出中');
    
      // 出張の場合、「出張」を書込
      if ( businessTrip ) setContents = attendList.getRange(rowNum, 5, 1, 1).setValue('出張中');
  
      // 本日の予定が 外出 ・ 出張 の場合、予定表の内容を書込
      if ( goOut || businessTrip ) detail = attendList.getRange(rowNum, 6, 1, 1).setValue(contents);
    }


    // 帰宅時に翌日の在席状態を書込
    if ( nowTime > 16 ) {
      // 翌日の予定が出張以外で、在席リストの状態が外出中でなければ、「帰宅」 を書込
      if ( nextContents === !nextBusinessTrip || setContents !== '外出中' ) {
        setContents = attendList.getRange(rowNum, 5, 1, 1).setValue('帰宅');
      };
    
      // 翌日の予定が 休み の場合、[ 日付 + 休み ] を詳細項目に書込
      const dateHol = `${nextDate} 休み`;
      if ( nextholJudg ) detail = attendList.getRange(rowNum, 6, 1, 1).setValue(dateHol);
     

      // 翌日の予定が 出張 の場合、詳細項目に書込
      if ( nextBusinessTrip ) {
        setContents = attendList.getRange(rowNum, 5, 1, 1).setValue('出張中');
        detail = attendList.getRange(rowNum, 6, 1, 1).setValue(nextContents);
      }
      // 翌日の休日判定がtrueの場合、「休み」を書込
      if ( nextholJudg ) setContents = attendList.getRange(rowNum, 5, 1, 1).setValue('休み');

      // 翌日の予定が予定無し ・ 24H ・ 当番 の場合、予定を削除
      attend.forEach( el => {
        if ( el === nextContents ) detail = attendList.getRange(rowNum, 6, 1, 1).setValue('');
      });  
    };
  

    // 期間限定で発動(iサーチ打ち合わせ開始)
    if ( nowTime == 10 ) {
      if ( setContents === '在席' ) {
        setContents = attendList.getRange(rowNum, 5, 1, 1).setValue('会議中');
        if ( detail !== '' ) detail = attendList.getRange(rowNum, 6, 1, 1).setValue('10 ~ 11時')
      }
    }


    // 期間限定で発動(iサーチ打ち合わせ終了)
    if ( nowTime == 11 ) {
      if ( setContents === '会議中' ) {
        setContents = attendList.getRange(rowNum, 5, 1, 1).setValue('在席');
        if ( detail !== '' ) detail = attendList.getRange(rowNum, 6, 1, 1).setValue('')
      }
    }

  });

}
