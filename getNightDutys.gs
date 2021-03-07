/* ========================================================================= /
/  ===  在席リストの夜勤担当者の行・列番号を取得 [ 関数 ]                       === /
/  ======================================================================== */      

function GetNightDutys() {

  console.log("GetNightDutys実行！");

  // 変数の初期化
  let nightDuty1, nightDuty2;                 // 夜勤当番者の名前
  let nightAddress1, nightAddress2;           // 夜勤当番者のアドレス
  let nightName1, nightName2;                 // 夜勤当番者のフルネーム

  // 夜勤当番者を取得
  let nightDutys    = nightDuty[arrDayNum];   // 夜勤当番者2名
  let nightDutysFri = nightArea[arrDayNum];   // 金曜当番者2名

  // アドレス取得の条件判定_20210210
  const friday   = dayOfWeek === 5;                                // 金曜日でtrue
  const saturday = dayOfWeek === 6;                                // 土曜日でtrue
  const sunday   = dayOfWeek === 0;                                // 日曜日でtrue

  const dutysJudge = nightDutys.indexOf('/') != -1;                // 当番2名判定
  const dutysFriJudge = nightDutysFri.indexOf('/') != -1;          // 当番2名判定(金曜日)

  // 金曜日以外 and セルに[/]を含む場合に実行
  if ( !friday && dutysJudge ) {
    nightDuty1 = nightDutys.split('/')[0];        // 1人目
    nightDuty2 = nightDutys.split('/')[1];        // 2人目

  // 金曜日 and セルに[/]を含む場合に実行
  } else if ( friday && dutysFriJudge ) {
    nightDuty1 = nightDutysFri.split('/')[0];        // 1人目
    nightDuty2 = nightDutysFri.split('/')[1];        // 2人目
  }
  
  // スプレットシートの行と列を反転させる。
  const _ = Underscore.load();
  const namesTrans = _.zip.apply(_, names);
 
  // サービス作業予定表の24H記入欄に空白の場合 >>> 予定表から夜勤担当者を取得
  if ( nightDutys == "" ) {
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
  const nowMinutes = Utilities.formatDate(date, 'Asia/Tokyo', 'm');


  // [関数]メール送信の実行条件
  const sendTimer = ( nowHours == "7" && nowMinutes >= "25" );

  // 実行時間が16時台だったら実行
  // NightSendMail();
    
  // 夜勤担当者
  const nightShiftDuty = [ nightName1, nightName2 ];

  // ログ確認用(確定)
  console.log("namesTrans(夜勤担当リスト名):" + namesTrans[0]);
  console.log("nameNum1(リストの配列番号1):" + nameNum1);
  console.log("nameNum2(リストの配列番号2):" + nameNum2);
  console.log("nightName1(夜勤担当名1):" + nightName1);
  console.log("nightName2(夜勤担当名2):" + nightName2);
  console.log("nightAddress1(アドレス1):" + nightAddress1);
  console.log("nightAddress2(アドレス2):" + nightAddress2);


  return nightShiftDuty;



  /* ========================================================================= /
  /  ===  当日の夜勤担当者の記入がない場合にメンバーの当日の予定から判定 [ 関数 ]     === /
  /  ======================================================================== */      

  function GetNightContetns() {

    console.log("GetNightContetns実行！");
      
    // 変数の初期化
    const nightDutysArr = [];                              // 夜勤予定の行を格納
    i = 8;                                                 // 変数の初期化(予定取得開始の行番号)

    // サービス作業予定表からメンバー全員分の当日の予定を取得
    const thisContents = schedule.getRange(i, dayNum+1, monLastRow, 1).getValues().flat();


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
    const red = nightBG[arrDayNum] === '#ff0000';
    const redFri = nightAreaBG[arrDayNum] === '#ff0000';

    // ログ確認用
    console.log("red:" + red);
    console.log("redFri:" + redFri);
    console.log("!red && !redFri:" + (!red && !redFri));

    // セルの背景色が赤色でなければ実行
    if ( !red && !redFri ) {

      console.log("メールを送信しました！");

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