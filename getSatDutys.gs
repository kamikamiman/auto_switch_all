/* ========================================================================= /
/  ===  在席リストの土曜当番者の行・列番号を取得 [ 関数 ]                       === /
/  ======================================================================== */      


function GetSatDutys() {
  
  console.log("GetSatDutys実行！")

  // メンバー（土曜・24h）から対象者の情報を取得
  const sheetAdress = ssGet.getSheetByName('メンバー（土曜・24h）'); // スプレットシート情報を取得
  const lastRow = sheetAdress.getLastRow() - 1;
  const names = sheetAdress.getRange(2, 2, lastRow, 3).getValues();

  // スプレットシートの行と列を反転させる。
  const _ = Underscore.load();
  const namesTrans = _.zip.apply(_, names);

  // 土曜当番者を取得
  let saturdayDuty1 = saturdayDuty[arrDayNum];   // 土曜当番者(1人目)
  let saturdayDuty2 = nightDuty[arrDayNum];      // 土曜当番者(2人目)

  // 土曜当番者の配列番号を取得
  nameNum1 = namesTrans[0].indexOf(saturdayDuty1);
  nameNum2 = namesTrans[0].indexOf(saturdayDuty2);

  // 土曜当番者のフルネームを取得
  saturdayName1 = namesTrans[2][nameNum1];
  saturdayName2 = namesTrans[2][nameNum2];

  // 土曜当番者
  const satDuty = [ saturdayName1, saturdayName2 ];

  return satDuty;

};
