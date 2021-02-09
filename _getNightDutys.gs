/* ========================================================================= /
/  ===  在席リストの夜勤担当者の行・列番号を取得 [ 関数 ]                       === /
/  ======================================================================== */      

function _GetNightDutys() {

  console.log("GetNightDutys実行！");

  // 変数の初期化
  let nightAddress1, nightAddress2;           // 夜勤当番者のアドレス
  let nightName1, nightName2;                 // 夜勤当番者のフルネーム

  // 夜勤当番者を取得
  let nightDuty1 = nightDutys1[arrDayNum];   // 夜勤当番者1人目
  let nightDuty2 = nightDutys2[arrDayNum];   // 夜勤当番者2人目

  // アドレス取得の条件判定
  const dutysJudge = nightDuty1 != '';       // 当番1名判定
  
  // スプレットシートの行と列を反転させる。
  const _ = Underscore.load();
  const namesTrans = _.zip.apply(_, names);
 
  // サービス作業予定表の24H記入欄に空白の場合 >>> 予定表から夜勤担当者を取得
  if ( nightDuty1 == "" ) {
    GetNightContetns();

    // 夜勤当番者の配列番号を取得
    nameNum1 = namesTrans[2].indexOf(nightName1);
    nameNum2 = namesTrans[2].indexOf(nightName2);
  } else {

    // 夜勤当番者の配列番号を取得
    nameNum1 = namesTrans[0].indexOf(nightDuty1);
    nameNum2 = namesTrans[0].indexOf(nightDuty2);

    // 夜勤当番者のフルネームを取得
    nightName1 = namesTrans[2][nameNum1];
    nightName2 = namesTrans[2][nameNum2];
  }

  // 夜勤当番者のアドレスを取得
  nightAddress1 = namesTrans[1][nameNum1];
  nightAddress2 = namesTrans[1][nameNum2];

  // 時間を取得
  const date = new Date();
  const nowHours = Utilities.formatDate(date, 'Asia/Tokyo', 'H');

  // [関数]メール送信の実行条件
  const sendTimer = ( nowHours == "16" );

  // 実行時間が16時台だったら実行
  if ( sendTimer ) NightSendMail();
    
  // 夜勤担当者
  const nightShiftDuty = [ nightName1, nightName2 ];

  // ログ確認用
  console.log("namesTrans:" + namesTrans[0]);
  console.log("nameNum1:" + nameNum1);
  console.log("nameNum2:" + nameNum2);
  console.log("nightName1:" + nightName1);
  console.log("nightName2:" + nightName2);
  console.log("nightAddress1:" + nightAddress1);
  console.log("nightAddress2:" + nightAddress2);
  console.log("nightShiftDuty:" + nightShiftDuty);


  return nightShiftDuty;



  /* ========================================================================= /
  /  ===  当日の夜勤担当者の記入がない場合にメンバーの当日の予定から判定 [ 関数 ]     === /
  /  ======================================================================== */      

  function GetNightContetns() {

    console.log("GetNightContetns実行！");
      
    // サービス作業予定表からメンバー全員分の当日の予定を取得
    const thisContents = schedule.getRange(9, dayNum+1, monLastRow, 1).getValues().flat();

    // 変数の初期化
    const nightDutysArr = [];                              // 夜勤予定の行を格納
    i = 8;                                                 // 変数の初期化(予定取得開始の行番号)

    // 当日の予定の中に夜勤予定があるかチェック
    thisContents.forEach( el => {
      el = String(el);                                     // 予定を文字列Stringに変換(searchが数値NumberがNGのため)
      const nightJudge = el.search(24) !== -1;             // 文字列に[24]を含んでいればtrue判定
      if ( nightJudge ) nightDutysArr.push(members[i-1]);  // 配列に夜勤担当者を追加
      i++;                                                 // 行番号を加算
    });

    nightName1 = nightDutysArr[0];
    nightName2 = nightDutysArr[1];

  };


  /* ========================================================================= /
  /  ===  セルの背景色が赤色でなければ通知メールを送信 [ 関数 ]                    === /
  /  ======================================================================== */      

  function NightSendMail() {

    console.log("NightSendMail実行！");


    // セル背景色(赤色)判定
    const red1 = nightBG1[arrDayNum] === '#ff0000';
    const red2 = nightBG2[arrDayNum] === '#ff0000';

    // セルの背景色が赤色でなければ実行
    if ( !red1 || !red2 ) {

      // メールの送信先
      // let to = [ nightAddress1,nightAddress2 ];
      let to = [ "k.kamikura@isowa.co.jp", "kamikurakenta@gmail.com" ];


      // メールのタイトル
      const subject = '24時間当番の電話確認をお願いします。';

      // メールの本文
      const body = '\
${nightName1}さん, ${nightName2}さん\n\n\
お仕事お疲れ様です。\n\
本日、24時間当番お願いします。\n\
まだ電話確認が取れていません。\n\
確認をお願いします。'
.replace('${nightName1}', nightName1)
.replace('${nightName2}', nightName2);

      // オプション(送信元・bcc)      
      const options = { name: 'ISOWA_support',
                        bcc: 'k.kamikura@isowa.co.jp'
                      };

      GmailApp.sendEmail(
        to,
        subject,
        body,
        options
      )
    }
  };
};