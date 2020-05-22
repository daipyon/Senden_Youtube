import RPA from 'ts-rpa';
import { WebElement } from 'selenium-webdriver';

// SlackのトークンとチャンネルID
const Slack_Token = process.env.AbemaTV_RPAError_Token;
const Slack_Channel = process.env.AbemaTV_RPAError_Channel;
const Slack_Text = [`【Youtube 非公開シート追加】シートに転記完了しました`];

// スプレッドシートIDとシート名を記載
// const mySSID = process.env.My_SheetID;
const SSID = process.env.Senden_Youtube_SheetID;
const SSName2 = process.env.Senden_Youtube_SheetName2;

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
    const Today = await RPA.Google.Spreadsheet.getValues({
      spreadsheetId: `${SSID}`,
      range: `${SSName2}!B1:B1`
    });
    // 1週間前の日付に変換
    var date = await new Date(Today[0][0] + ' 09:00:00');
    await date.setDate(date.getDate() - 7);
    var year = await date.getFullYear();
    var month = (await date.getMonth()) + 1;
    var day = await date.getDate();
    var month2 = await ('00' + month).slice(-2);
    var day2 = await ('00' + day).slice(-2);
    const Yesterday = `${year}` + '/' + `${month2}` + '/' + `${day2}`;
    await RPA.Logger.info('1週間前の日付 → ' + Yesterday);
    await Work(Yesterday);
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

async function Work(Yesterday) {
  try {
    // Youtube Studioにログイン
    await YoutubeLogin();
    // 動画の日付を指定
    await VideoDate(Yesterday);
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

async function VideoDate(Yesterday) {
  const Icon = await RPA.WebBrowser.wait(
    RPA.WebBrowser.Until.elementLocated({ id: 'filter-icon' }),
    5000
  );
  await RPA.WebBrowser.mouseClick(Icon);
  await RPA.sleep(1000);

  // テスト用
  // const Videodate = await RPA.WebBrowser.wait(
  //   RPA.WebBrowser.Until.elementLocated({ id: 'text-item-6' }),
  //   5000
  // );

  // 本番用
  const Videodate = await RPA.WebBrowser.wait(
    RPA.WebBrowser.Until.elementLocated({ id: 'text-item-4' }),
    5000
  );

  await RPA.WebBrowser.mouseClick(Videodate);
  await RPA.sleep(1000);

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
  var split = await Yesterday.split(/\//);
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

  // 1週間前の日付を取得するため月が変わると入力でエラーになる。以下2行については対策が必要。
  // 5/1 コメントアウト → 5/8 コメントイン
  await RPA.WebBrowser.sendKeys(EndDate, [RPA.WebBrowser.Key.BACK_SPACE]);
  await RPA.WebBrowser.sendKeys(EndDate, [RPA.WebBrowser.Key.BACK_SPACE]);
  await RPA.sleep(500);

  await RPA.WebBrowser.sendKeys(EndDate, [split[2]]);
  // スペースを入れると「適用」ボタンを押すことができる
  await RPA.WebBrowser.sendKeys(EndDate, [RPA.WebBrowser.Key.SPACE]);

  // テスト
  await RPA.WebBrowser.takeScreenshot();

  await RPA.sleep(1000);
  const ApplyButton = await RPA.WebBrowser.findElementById('apply-button');
  await RPA.WebBrowser.mouseClick(ApplyButton);
  await RPA.sleep(3000);
}

let Total = [];
const TitleList = [];
const StatusList = [];
const DateList = [];
const TextList = [];
const LoopFlag = ['true'];
const FisrtGetDataFlag = ['true'];
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
  Total[0] = Number(text[3]);
  if (Total[0] < 31) {
    await RPA.Logger.info('合計数が 31 未満ですので何もしません');
  } else {
    await RPA.Logger.info('1ページあたりの行数を 50 に増やします');
    const PageSize = await RPA.WebBrowser.findElementById('page-size');
    await RPA.WebBrowser.mouseClick(PageSize);
    await RPA.sleep(1000);
    const PageSize2 = await RPA.WebBrowser.findElementById('text-item-2');
    await RPA.WebBrowser.mouseClick(PageSize2);
    await RPA.sleep(1000);
  }
  var count = 1;
  while (0 == 0) {
    if (FisrtGetDataFlag[0] == 'false') {
      await RPA.Logger.info(count + ' 週目...');
      Total[0] = Number(text[3]) - 50;
      await RPA.Logger.info('残数　　　　　　 → ' + Total[0]);
    }
    // カーソルが被るとテキストが取得できないため、画面上部にカーソルを動かす
    const UploadVideo = await RPA.WebBrowser.findElementById(
      'video-list-uploads-tab'
    );
    await RPA.WebBrowser.mouseMove(UploadVideo);
    await RPA.sleep(1000);
    await RPA.Logger.info('エラー文言の有無を確認します');
    for (let i = 0; i <= Total[0] - 1; i++) {
      try {
        // エラー文言が出てきた場合の処理
        const CancelProcess: WebElement = await RPA.WebBrowser.driver.executeScript(
          `return document.getElementsByClassName('style-scope ytcp-video-row failure-status')[${i}]`
        );
        const CancelProcessText = await CancelProcess.getText();

        // テスト用
        // if (CancelProcessText == 'Processing abandoned') {

        // 本番用
        if (CancelProcessText == '処理を中止しました') {
          await RPA.Logger.info(`${i} 行目 エラー文言があります`);
          CancelProcessTextList.push(CancelProcessText);
        }
      } catch {
        await RPA.Logger.info('エラー文言はありませんでした');
      }
    }
    for (let i = 0; i <= Total[0] - 1 - CancelProcessTextList.length; i++) {
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
      if (i == 50) {
        FisrtGetDataFlag[0] = 'false';
        await RPA.Logger.info('次のページに進みます');
        const NextPage = await RPA.WebBrowser.findElementById('navigate-after');
        await RPA.WebBrowser.mouseClick(NextPage);
        await RPA.sleep(3000);
        break;
      }
    }
    if (Total[0] < 51) {
      LoopFlag[0] = 'false';
      await RPA.Logger.info('データ取得を終了します');
      await RPA.Logger.info(TitleList);
      await RPA.Logger.info(StatusList);
      await RPA.Logger.info(DateList);
      await RPA.Logger.info(TextList);
      await RPA.Logger.info('取得した合計数　 → ' + TextList.length);
      break;
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
  await RPA.Logger.info('この行から記載します　　　 → ', LastRow);
  for (let i in TextList) {
    await RPA.Logger.info(LastRow + Number(i), '行目');

    // テスト用
    // if (TextList[i] == 'Published') {

    // 本番用
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
    // D列にタイトルを記載
    await RPA.Google.Spreadsheet.setValues({
      spreadsheetId: `${SSID}`,
      range: `${SSName2}!D${LastRow + Number(i)}:D${LastRow + Number(i)}`,
      values: [[TitleList[i]]]
    });
    // E列にステータスを記載
    await RPA.Google.Spreadsheet.setValues({
      spreadsheetId: `${SSID}`,
      range: `${SSName2}!E${LastRow + Number(i)}:E${LastRow + Number(i)}`,
      values: [[StatusList[i]]]
    });
  }
}
