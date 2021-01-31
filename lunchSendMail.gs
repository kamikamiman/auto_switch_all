// 昼当番の担当者に通知メールを送信
function LunchSendMail() {

  // スプレットシートの情報を取得
  const ssSheetAdress = SpreadsheetApp.openById('1HIP359dJRclqwRV-H1KUo9Qluddv4SdWttLsr02DU18');
  const sheetAdress = ssSheetAdress.getSheetByName('メールアドレス（昼休み当番）');                     // シートを取得
  const lastRow = sheetAdress.getLastRow();                                                        // シートの最終行を取得
  const staffs = sheetAdress.getRange(2, 1, lastRow, 2).getValues();                               // 担当者、メールアドレスを取得

  // アンダースコアを使用して行と列を反転させる。
  const _ = Underscore.load();
  const staffsTrans = _.zip.apply(_, staffs);

  const staffNames    = staffsTrans[0];                                                            // 昼当番担当者名　一覧
  const staffAdresses = staffsTrans[1];                                                            // 昼当番担当者アドレス　一覧
  const lunchDutyName = lunchDuty[arrDayNum];                                                      // 昼当番担当者(サービス作業予定表)

  // 昼当番担当者のメールアドレスを取得する。 
  staffNames.forEach( staffName => {
    if(staffName == lunchDutyName){
      const staffNumber = staffNames.indexOf(staffName);
      const staffAdress = staffAdresses[staffNumber];

      // ログ確認用
      console.log("担当者：" + staffAdress);
        
      // メールの送信先
      const to = staffAdress;
      // メールのタイトル
      const subject = '本日の昼当番お願いします。';
      // メールの本文
      const body = '\
  ${staffName}さん\n\n\
  お仕事お疲れ様です。\n\
  本日、昼休みの電話当番です。\n\n\n\
  よろしくお願いします。'.replace('${staffName}', staffName)

      // オプション設定(送信元の名前)
      const options = { name: 'ISOWA_support'};

      // GmailApp.sendEmail(
      // to,
      // subject,
      // body,
      // options
      // );
    }
  });

  // ログ確認用
  console.log("メンバー(サービス作業予定表)：" + lunchDuty);
  console.log("シート列番号(当日)：" + arrDayNum);
  console.log("担当者 + アドレス：" + staffsTrans);
  console.log("担当者:" + lunchDutyName);
  console.log("メンバー(昼休み当番用)：" + staffNames);
  console.log("メンバーアドレス(昼休み用)：" + staffAdresses);


}
