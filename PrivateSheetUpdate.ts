import RPA from 'ts-rpa';
import { WebElement } from 'selenium-webdriver';
const moment = require('moment');

// SlackのトークンとチャンネルID
const Slack_Token = process.env.AbemaTV_RPAError_Token;
const Slack_Channel = process.env.AbemaTV_RPAError_Channel;
const Slack_Text = [`【Youtube 非公開シート追加】シートに転記完了しました`];

// スプレッドシートIDとシート名を記載
// const mySSID = process.env.My_SheetID;
const SSID = process.env.Senden_Youtube_SheetID;
const SSName2 = process.env.Senden_Youtube_SheetName2;

let Today;
// 本日より1週間前の日付
let AWeekAgo;
// 前月の25日
let TwentyFive;

// エラー発生時のテキストを格納
const ErrorText = [];

async function Start() {
  if (ErrorText.length == 0) {
    // デバッグログを最小限(INFOのみ)にする ※[DEBUG]が非表示になる
    RPA.Logger.level = 'INFO';
    await RPA.Google.authorize({
      // accessToken: process.env.GOOGLE_ACCESS_TOKEN,
      refreshToken: process.env.GOOGLE_REFRESH_TOKEN,
      tokenType: 'Bearer',
      expiryDate: parseInt(process.env.GOOGLE_EXPIRY_DATE, 10)
    });
    // B列(本日の日付)を取得
    Today = await RPA.Google.Spreadsheet.getValues({
      spreadsheetId: `${SSID}`,
      range: `${SSName2}!B1:B1`
    });
    await RPA.Logger.info('本日の日付　　 　　 → ' + Today[0][0]);
    // 1週間前の日付をフォーマット変更して取得
    AWeekAgo = moment()
      .subtract(7, 'days')
      .format('YYYY/MM/DD');
    await RPA.Logger.info('1週間前の日付 　　　→ ' + AWeekAgo);
    // 前月の25日をフォーマット変更して取得
    TwentyFive = moment()
      .subtract(1, 'months')
      .format('YYYY/MM/25');
    await RPA.Logger.info('前月の25日　　　　　→ ' + TwentyFive);
    await Work();
  }
  // エラー発生時の処理
  if (ErrorText.length >= 1) {
    // const DOM = await RPA.WebBrowser.driver.getPageSource();
    // await RPA.Logger.info(DOM);
    await RPA.SystemLogger.error(ErrorText);
    Slack_Text[0] = `【Youtube 非公開シート追加】でエラーが発生しました！\n${ErrorText}`;
    await RPA.WebBrowser.takeScreenshot();
  }
  await RPA.Logger.info('作業を終了します');
  await SlackPost(Slack_Text[0]);
  await RPA.WebBrowser.quit();
  await RPA.sleep(1000);
  await process.exit();
}

Start();

async function Work() {
  try {
    // Youtube Studioにログイン
    await YoutubeLogin();
    // 動画の日付を指定
    await VideoDate();
    // データを取得
    await GetData();
    // シートに取得したデータを記載
    await SetData();
  } catch (error) {
    ErrorText[0] = error;
    await Start();
  }
}

async function SlackPost(Text) {
  await RPA.Slack.chat.postMessage({
    token: Slack_Token,
    channel: Slack_Channel,
    text: `${Text}`
  });
}

// テストは海外版、本番は日本版であることに注意！
async function YoutubeLogin() {
  await RPA.Logger.info('Youtube Studioにログインします');
  await RPA.WebBrowser.get(process.env.Youtube_Studio_Url);
  await RPA.sleep(2000);
  const LoginID = await RPA.WebBrowser.wait(
    RPA.WebBrowser.Until.elementLocated({ id: 'Email' }),
    8000
  );
  await RPA.WebBrowser.sendKeys(LoginID, [`${process.env.Youtube_Login_ID}`]);
  const NextButton1 = await RPA.WebBrowser.findElementById('next');
  await RPA.WebBrowser.mouseClick(NextButton1);
  await RPA.sleep(5000);
  const GoogleLoginPW_1 = await RPA.WebBrowser.findElementsById('Passwd');
  const GoogleLoginPW_2 = await RPA.WebBrowser.findElementsById('password');
  await RPA.Logger.info(`GoogleLoginPW_1 ` + GoogleLoginPW_1.length);
  await RPA.Logger.info(`GoogleLoginPW_2 ` + GoogleLoginPW_2.length);
  if (GoogleLoginPW_1.length == 1) {
    await RPA.WebBrowser.sendKeys(GoogleLoginPW_1[0], [
      `${process.env.Youtube_Login_PW}`
    ]);
    const NextButton2 = await RPA.WebBrowser.findElementById('signIn');
    await RPA.WebBrowser.mouseClick(NextButton2);
  }
  if (GoogleLoginPW_2.length == 1) {
    const GoogleLoginPW = await RPA.WebBrowser.findElementById('password');
    await RPA.WebBrowser.sendKeys(GoogleLoginPW, [
      `${process.env.Youtube_Login_PW}`
    ]);
    const NextButton2 = await RPA.WebBrowser.findElementById('submit');
    await RPA.WebBrowser.mouseClick(NextButton2);
  }
  while (0 == 0) {
    await RPA.sleep(5000);
    const Filter = await RPA.WebBrowser.findElementsByClassName(
      'style-scope ytcp-table-footer'
    );
    if (Filter.length >= 1) {
      await RPA.Logger.info('＊＊＊ログイン成功しました＊＊＊');
      break;
    }
  }
}

async function VideoDate() {
  // 逆三角形のアイコンをクリック
  const Icon = await RPA.WebBrowser.wait(
    RPA.WebBrowser.Until.elementLocated({ id: 'filter-icon' }),
    5000
  );
  await RPA.WebBrowser.mouseClick(Icon);
  await RPA.sleep(1000);

  await RPA.WebBrowser.takeScreenshot();

  // テスト用
  // const Videodate = await RPA.WebBrowser.wait(
  //   RPA.WebBrowser.Until.elementLocated({ id: 'text-item-6' }),
  //   5000
  // );

  //「動画の日付」をクリック
  // 本番用
  const Videodate = await RPA.WebBrowser.wait(
    RPA.WebBrowser.Until.elementLocated({ id: 'text-item-4' }),
    5000
  );

  await RPA.WebBrowser.mouseClick(Videodate);
  await RPA.sleep(1000);

  await RPA.WebBrowser.takeScreenshot();

  // テスト用
  // const Click = await RPA.WebBrowser.wait(
  //   RPA.WebBrowser.Until.elementLocated({
  //     xpath: '//*[@id="items"]/div[181]/div[3]/span[2]' // ← 要素は適宜変更
  //   }),
  //   5000
  // );
  // await RPA.WebBrowser.mouseClick(Click);
  // await RPA.sleep(1000);
  // const Click2 = await RPA.WebBrowser.wait(
  //   RPA.WebBrowser.Until.elementLocated({
  //     xpath: '//*[@id="items"]/div[181]/div[3]/span[2]' // ← 要素は適宜変更
  //   }),
  //   5000
  // );
  // await RPA.WebBrowser.mouseClick(Click2);

  // 本番用
  // そのまま入力するとエラーが起きるため、日付を分割して入力する
  var split = await AWeekAgo.split(/\//);
  // 開始・終了期間に同日を入力
  const StartDate: WebElement = await RPA.WebBrowser.driver.executeScript(
    `return document.getElementsByTagName('input')[2]`
  );
  await RPA.WebBrowser.sendKeys(StartDate, [split[0] + '/' + split[1] + '/']);
  await RPA.sleep(500);
  await RPA.WebBrowser.sendKeys(StartDate, [RPA.WebBrowser.Key.BACK_SPACE]);
  await RPA.WebBrowser.sendKeys(StartDate, [RPA.WebBrowser.Key.BACK_SPACE]);
  await RPA.sleep(500);
  await RPA.WebBrowser.sendKeys(StartDate, [split[2]]);
  await RPA.sleep(500);
  const EndDate: WebElement = await RPA.WebBrowser.driver.executeScript(
    `return document.getElementsByTagName('input')[3]`
  );
  await RPA.WebBrowser.sendKeys(EndDate, [split[0] + '/' + split[1] + '/']);
  await RPA.sleep(500);
  // 1週間前の日付と前月の25日が一致した場合はツールの仕様が変わるため以下フローで対応
  if (AWeekAgo == TwentyFive) {
    await RPA.Logger.info(`前月の25日のため、日にちは削除しませんでした`);
  } else {
    // 日にちを一旦削除してから再度日にちを入力
    await RPA.WebBrowser.sendKeys(EndDate, [RPA.WebBrowser.Key.BACK_SPACE]);
    await RPA.WebBrowser.sendKeys(EndDate, [RPA.WebBrowser.Key.BACK_SPACE]);
    await RPA.sleep(500);
    await RPA.Logger.info(`日にちを削除しました`);
  }
  await RPA.WebBrowser.sendKeys(EndDate, [split[2]]);
  // スペースを入れると「適用」ボタンを押すことができる
  await RPA.WebBrowser.sendKeys(EndDate, [RPA.WebBrowser.Key.SPACE]);

  await RPA.WebBrowser.takeScreenshot();

  await RPA.sleep(1000);
  //「適用」をクリック
  const ApplyButton = await RPA.WebBrowser.findElementById('apply-button');
  await RPA.WebBrowser.mouseClick(ApplyButton);
  await RPA.sleep(3000);
}

// 日付検索でヒットした合計数
let Total;
// 1ページ目でデータを取得できたか判定するフラグ
let FisrtGetDataFlag = 'true';
// 取得したタイトルを格納する
const TitleList = [];
// 取得したステータスを格納する
const StatusList = [];
// 取得したアップロード日を格納する
const DateList = [];
// 取得したアップロード日の下のステータスを格納する
const TextList = [];
// ループを抜けるためのフラグ
let LoopFlag = 'true';
//「処理を中止しました」がタイトルになっている番組を格納する
const CancelProcessTextList = [];
async function GetData() {
  const Range = await RPA.WebBrowser.wait(
    RPA.WebBrowser.Until.elementLocated({
      className: 'page-description style-scope ytcp-table-footer'
    }),
    5000
  );
  const RangeText = await Range.getText();

  // テスト用
  // var text = await RangeText.split(/[\–\/\s]+/);

  // 本番用
  var text = await RangeText.split(/[\～\/\s]+/);

  await RPA.Logger.info('ヒットした合計数 → ' + text[3]);
  Total = Number(text[3]);
  // 検索にヒットしたタイトルの合計数が 31 未満の場合
  if (Total < 31) {
    await RPA.Logger.info('合計数が 31 未満ですので何もしません');
  } else {
    // 31 以上の場合は表示件数を 50 にする
    await RPA.Logger.info('1ページあたりの行数を 50 に増やします');
    //「1ページあたりの行数」の数字をクリック
    const PageSize = await RPA.WebBrowser.findElementById('page-size');
    await RPA.WebBrowser.mouseClick(PageSize);
    await RPA.sleep(1000);
    //「50」をクリック
    const PageSize2 = await RPA.WebBrowser.findElementById('text-item-2');
    await RPA.WebBrowser.mouseClick(PageSize2);
    await RPA.sleep(1000);
  }
  var count = 1;
  while (LoopFlag == 'true') {
    // 2ページ目以降は最初にヒットした件数から常に 50 を引いた数だけ回す
    if (FisrtGetDataFlag == 'false') {
      await RPA.Logger.info(`${count} 週目...`);
      Total = Number(text[3]) - 50;
      await RPA.Logger.info('残数　　　　　　 → ' + Total);
    }
    // カーソルが被るとテキストが取得できないため、画面上部にカーソルを動かす
    const UploadVideo = await RPA.WebBrowser.findElementById(
      'video-list-uploads-tab'
    );
    await RPA.WebBrowser.mouseMove(UploadVideo);
    await RPA.sleep(1000);
    await RPA.Logger.info('エラー文言の有無を確認します');
    for (let i = 0; i <= Total - 1; i++) {
      try {
        //「処理を中止しました」と書かれている場合はその文字を取得
        const CancelProcess: WebElement = await RPA.WebBrowser.driver.executeScript(
          `return document.getElementsByClassName('style-scope ytcp-video-row failure-status')[${i}]`
        );
        const CancelProcessText = await CancelProcess.getText();

        // テスト用
        // if (CancelProcessText == 'Processing abandoned') {

        // 本番用
        if (CancelProcessText == '処理を中止しました') {
          CancelProcessTextList.push(CancelProcessText);
        }
      } catch {}
    }
    if (CancelProcessTextList.length < 1) {
      await RPA.Logger.info('エラー文言はありませんでした');
    } else {
      await RPA.Logger.info(
        `エラー文言が ${CancelProcessTextList.length} つあります`
      );
    }
    // 常に「処理を中止しました」というタイトルがあった分だけ検索にヒットした件数から引いていく
    for (let i = 0; i <= Total - 1 - CancelProcessTextList.length; i++) {
      // タイトルを取得
      const Title: WebElement = await RPA.WebBrowser.driver.executeScript(
        `return document.getElementsByClassName('style-scope ytcp-video-list-cell-video video-title-wrapper')[${i}].children[0]`
      );
      const TitleText = await Title.getText();
      TitleList.push(TitleText);
      // ステータスを取得
      const Status: WebElement = await RPA.WebBrowser.driver.executeScript(
        `return document.getElementsByClassName('style-scope ytcp-video-row cell-body tablecell-visibility')[${i}].children[0].children[0].children[1]`
      );
      const StatusText = await Status.getText();
      StatusList.push(StatusText);
      // 日付を取得
      const date: WebElement = await RPA.WebBrowser.driver.executeScript(
        `return document.getElementsByClassName('style-scope ytcp-video-row cell-body tablecell-date sortable column-sorted')[${i}]`
      );
      const JudgeText = await date.getText();
      var split = await JudgeText.split('\n');
      DateList.push(split[0]);
      TextList.push(split[1]);
      // 1ページ目で取得が完了しなければ次のページに進む
      if (i == 49) {
        FisrtGetDataFlag = 'false';
        await RPA.Logger.info('次のページに進みます');
        const NextPage = await RPA.WebBrowser.findElementById('navigate-after');
        await NextPage.click();
        await RPA.sleep(3000);
        break;
      }
    }
    // 検索にヒットした件数が残り 50 以下になったらループを抜ける
    if (Total < 51) {
      LoopFlag = 'false';
      await RPA.Logger.info('データ取得を終了します');
      await RPA.Logger.info(TitleList);
      await RPA.Logger.info(StatusList);
      await RPA.Logger.info(DateList);
      await RPA.Logger.info(TextList);
      await RPA.Logger.info(`取得した合計数　 → ${TextList.length}`);
    }
    count += 1;
  }
}

async function SetData() {
  // シートの最終行の次の行を取得
  const WorkRow3 = await RPA.Google.Spreadsheet.getValues({
    spreadsheetId: `${SSID}`,

    // テスト用
    // range: `${SSName2}!B1:B30000`

    // 本番用
    range: `${SSName2}!E1:E30000`
  });
  const LastRow = WorkRow3.length + 1;
  await RPA.Logger.info(`この行から記載します　　　 → ${LastRow}`);
  // 取得したデータの数だけ回す
  for (let i in TextList) {
    await RPA.Logger.info(LastRow + Number(i), '行目');

    // テスト用
    // if (TextList[i] == 'Published') {

    // 本番用
    //「公開日」と書かれていたタイトルのみ日付を記載する
    if (TextList[i] == '公開日') {
      await RPA.Logger.info('日付を記載します');
      // B列に公開日を記載
      await RPA.Google.Spreadsheet.setValues({
        spreadsheetId: `${SSID}`,
        range: `${SSName2}!B${LastRow + Number(i)}:B${LastRow + Number(i)}`,
        values: [[DateList[i]]],
        // 数字の前の ' が入力されなくなる
        parseValues: true
      });
    } else {
      await RPA.Logger.info('日付は記載しません');
    }
    // D列にタイトル、E列にステータスを記載
    await RPA.Google.Spreadsheet.setValues({
      spreadsheetId: `${SSID}`,
      range: `${SSName2}!D${LastRow + Number(i)}:E${LastRow + Number(i)}`,
      values: [[TitleList[i], StatusList[i]]]
    });
    await RPA.sleep(1000);
  }
}
