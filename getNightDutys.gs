// ・メール送信の関数がない。
// ・修正時、スプレットシートURLを変更する。
// ・24時間のメンバー名を在席リストの名前とリンクさせる。
//   当日の在席リストのメンバー情報を取得。                                  >>> 完了
//   ①24時間担当者の在席リストの名前のセルを色で塗りつぶす。                   >>> 完了
//   ②情報に24, 24hがあれば、在席リストのそのメンバーのセルを色で塗りつぶす。    >>> 完了
//   実行する時間帯は、当日の朝8:25                                        >>> 完了
//   このスクリプトを在席リストに移す。                                     >>> 完了
//   プログラムの実行条件・整理                                            >>> 完了


function GetNightDutys() {

  // メールアドレス一覧から送信対象者の情報を取得
  const sheetAdress = ssGet.getSheetByName('メールアドレス（24h）'); // スプレットシート情報を取得
  const lastRow = sheetAdress.getLastRow() - 1;
  const names = sheetAdress.getRange(2, 2, lastRow, 3).getValues();

  // 変数の初期化
  let nightDuty1, nightDuty2;                 // 夜勤当番者の名前
  let nightAddress1, nightAddress2;           // 夜勤当番者のアドレス
  let nightName1, nightName2;                 // 夜勤当番者のフルネーム

  // 夜勤当番者を取得
  let nightDutys    = nightDuty[arrDayNum];   // 夜勤当番者2名
  let nightDutysFri = nightArea[arrDayNum];   // 金曜当番者2名

  // アドレス取得の条件判定
  const friday = dayOfWeek === 5;                                  // 金曜日でtrue
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

  // 夜勤当番者の配列番号を取得
  nameNum1 = namesTrans[0].indexOf(nightDuty1);
  nameNum2 = namesTrans[0].indexOf(nightDuty2);

  // 夜勤当番者のアドレスを取得
  nightAddress1 = namesTrans[1][nameNum1];
  nightAddress2 = namesTrans[1][nameNum2];

  // 夜勤当番者のフルネームを取得
  nightName1 = namesTrans[2][nameNum1];
  nightName2 = namesTrans[2][nameNum2];


  // サービス作業予定表の24H記入欄に空白の場合 >>> 予定表から夜勤担当者を取得
  if ( nightDutys == "" ) GetNightContetns();

  // 時間を取得
  const date = new Date();
  const nowHours = Utilities.formatDate(date, 'Asia/Tokyo', 'H');

  // [関数]メール送信の実行条件
  const sendTimer = ( nowHours == "16" );

  // 実行時間が16時台だったら実行
  if ( sendTimer ) NightSendMail();

  // 夜勤担当者
  const nightShiftDuty = [ nightName1, nightName2 ];

  console.log("nightShiftDuty:" + nightShiftDuty);

  // ログ確認用
  console.log("最終行：" + lastRow);
  console.log("送信対象者：" + names);
  console.log("２４Ｈ当番(2名)：" + nightDutys);
  console.log("２４Ｈ当番・金曜日(2名)：" + nightDutysFri);
  console.log("当番2名判定：" + dutysJudge);
  console.log("当番2名判定(金曜日)：" + dutysFriJudge);
  console.log("24時間当番(1人目)：" + nightDuty1);
  console.log("24時間当番(2人目)：" + nightDuty2);
  console.log("当番フルネーム(1人目)：" + nightName1);
  console.log("当番フルネーム(2人目)：" + nightName2);
  console.log("送信者一覧:" + namesTrans[0]);
  console.log("メールアドレス一覧:" + namesTrans[1]);
  console.log("担当者1人目：" + nightAddress1);
  console.log("担当者2人目：" + nightAddress2);
  console.log("arrDayNum:" + arrDayNum);
  console.log(nightBG[arrDayNum]);
  console.log(nightAreaBG[arrDayNum]);
  console.log("nightBG:" + nightBG); 
  console.log("nightAreaBG:" + nightAreaBG);


  return nightShiftDuty;



  /* ========================================================================= /
  /  ===  当日の夜勤担当者の記入がない場合にメンバーの当日の予定から判定 [ 関数 ]     === /
  /  ======================================================================== */      

  function GetNightContetns() {

    console.log("GetNightContetns実行！");
      
    // サービス作業予定表からメンバー全員分の当日の予定を取得
    const thisContents = schedule.getRange(8, dayNum+1, monLastRow, 1).getValues().flat();

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

    // ログ確認用
    console.log("nightDutysArr:" + nightDutysArr);

  };


  /* ========================================================================= /
  /  ===  セルの背景色が赤色でなければ通知メールを送信 [ 関数 ]                    === /
  /  ======================================================================== */      

  function NightSendMail() {

    // セル背景色(赤色)判定
    const red = nightBG[arrDayNum] === '#ff0000';
    const redFri = nightAreaBG[arrDayNum] === '#ff0000';

    // セルの背景色が赤色でなければ実行
    if ( !red || !redFri ) {

      // メールの送信先
      // let to = [ nightAddress1,nightAddress2 ];
      let to = [ "k.kamikura@isowa.co.jp", "kamikurakenta@gmail.com" ];


      // メールのタイトル
      const subject = '24時間当番の電話確認をお願いします。';

      // メールの本文
      const body = '\
${nightDuty1}さん,${nightDuty2}さん\n\n\
お仕事お疲れ様です。\n\
本日、24時間当番お願いします。\n\
まだ電話確認が取れていません。\n\
確認をお願いします。'
.replace('${nightDuty1}', nightDuty1)
.replace('${nightDuty2}', nightDuty2);

      // オプション(送信元・bcc)      
      const options = { name: 'ISOWA_support',
                        bcc: 'k.kamikura@isowa.co.jp'
                      };

      GmailApp.sendEmail(
        to,
        subject,
        body,
        options
      );

    // ログ確認用
    // console.log("宛先:" + to);
    // console.log("タイトル:" + subject);
    // console.log("本文:" + body);
    // console.log("オプション:" + options);
    // console.log("赤色:" + red); 
    // console.log("赤色(金曜):" + redFri); 


    }
  };
};