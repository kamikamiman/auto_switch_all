// 前月の当番管轄が取得されていない。


// 24時間サービスの転送設定がされていない場合に通知メールを送信
function TransferSetMail() {

  // 前月の月を取得する。
  let beforeMonth;

  // 今月が1月の場合、翌月は12月
  if ( nowMonth != 1 ) { 
    beforeMonth = nowMonth -1;
  } else {
    beforeMonth = 12;
  }

  // 今月が4月の場合、〇〇期 - 1とする。
  if ( nowMonth == 4 ) period = period -1;

  // 前月のシート情報を取得
  const beforeSchedule = ssGet.getSheetByName('${period}期${beforeMonth}月'
                                      .replace('${period}', period)
                                      .replace('${beforeMonth}', beforeMonth));

  // 前月のシートの最終列を取得
  const lastColumn = beforeSchedule.getLastColumn();

  const getCellBackground1 = saturdayBG2[arrDayNum];        // 土曜日当番２の当日のセル背景色
  const getCellBackground2 = nightBG[arrDayNum];            // ２４Ｈメンバーの当日のセル背景色


  // メール送信判定
  const offDay         = (dayOfWeek === 0 || dayOfWeek === 6);                                    // 土曜・日曜でtrue
  const sameName       = (nightArea[arrDayNum] === nightArea[arrDayNum-1]);                       // 前日と担当管轄一致でtrue
  const notNightShift  = (nightDuty[arrDayNum] == '');                                            // 24時間なしの場合でtrue
  const switchCheck    = (getCellBackground1 === '#ffff00' || getCellBackground2 === '#ffff00');  // 転送切替完了でtrue


  // [平日] and [前日と管轄が違う] and [24Hサービス行の当日セルが空白] and [背景色が黄色でない] で実行
  if ( !offDay && !sameName && !notNightShift && !switchCheck ) {

    console.log("24時間当番の転送設定の確認が取れない為、メールを送信しました！");

    // 夜勤当番者に電話が繋がるかの確認依頼メールを送信する。
    const to = 'technical-support@isowa.co.jp';       // メールの送信先
    const subject = '24時間当番の電話転送設定の確認依頼';  // メールのタイトル

    // メールの本文
    const body = '\
テクニカルサポートの皆様\n\n\
お仕事お疲れ様です。\n\
本日、24時間当番の電話転送設定の確認が取れていません。\n\
お手間ですが、転送設定の確認をお願いします。\n\n\n\
よろしくお願いします。';
          
    // オプション( 送信元・bcc )
    const options = { name: 'ISOWA_support',
                      bcc: 'k.kamikura@isowa.co.jp'};
        
    //メール送信実行
    // GmailApp.sendEmail(
    //   to,
    //   subject,
    //   body,
    //   options
    // )

 };

// ログ確認用(確定)
 console.log("今月:" + nowMonth)
 console.log("前月:" + beforeMonth);
 console.log("前月の最終列:" + lastColumn);
 console.log("本日の日付:" + nowDay);
 console.log("夜勤当番の管轄情報:" + nightArea);
 console.log("当日の当番管轄：" + nightArea[arrDayNum]);
 console.log("前日の当番管轄：" + nightArea[arrDayNum-1]);
 console.log("当日の列番号：" + (dayNum + 1));
 console.log("土曜日当番２のセル背景色：" + saturdayBG2);
 console.log("２４Ｈの当日のセル背景色1：" + getCellBackground1);
 console.log("２４Ｈの当日のセル背景色2：" + getCellBackground2);
 console.log("当日の曜日:" + dayOfWeek);
 console.log("平日判定:" + !offDay);
 console.log("管轄判定:" + !sameName);
 console.log("宿直判定:" + !notNightShift);
 console.log("切替判定:" + !switchCheck);


};