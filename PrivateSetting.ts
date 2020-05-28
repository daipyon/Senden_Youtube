import RPA from 'ts-rpa';
import { By, WebElement } from 'selenium-webdriver';

// SlackのトークンとチャンネルID
const Slack_Token = process.env.AbemaTV_RPAError_Token;
const Slack_Channel = process.env.AbemaTV_RPAError_Channel;
const Slack_Text = [`【Youtube 非公開設定】設定完了しました`];
const Slack_Text2 = [`【Youtube 非公開設定】シートにエラー行があります`];

// スプレッドシートIDとシート名を記載
// const SSID = process.env.My_SheetID;
const SSID = process.env.Senden_Youtube_SheetID;
const SSID2 = process.env.Senden_Youtube_SheetID2;
const SSName = ['2019', '2020'];
const SSName2 = process.env.Senden_Youtube_SheetName3;

// 作業するスプレッドシートから読み込む行数を記載
const StartRow = 3;
const StartRow2 = 1;
const LastRow = 20000;

// 作業対象行とデータを取得
let WorkData;
let Row;

// 現在のシート名
let CurrentSSName;

// 動画のIDを保持する変数
const VideoID = [];
// 動画URL
const Video_URL = process.env.Youtube_Video_Url;
// 一般URL
const General_URL = process.env.Youtube_General_Url;

// エラー発生時のテキストを格納
const ErrorText = [];

async function Start() {
  // デバッグログを最小限(INFOのみ)にする ※[DEBUG]が非表示になる
  RPA.Logger.level = 'INFO';
  await RPA.Google.authorize({
    // accessToken: process.env.GOOGLE_ACCESS_TOKEN,
    refreshToken: process.env.GOOGLE_REFRESH_TOKEN,
    tokenType: 'Bearer',
    expiryDate: parseInt(process.env.GOOGLE_EXPIRY_DATE, 10)
  });
  for (let i in SSName) {
    CurrentSSName = SSName[i];
    if (Number(i) == 1) {
      await RPA.Logger.info(`${CurrentSSName} 年のシートに移動します`);
    }
    // B列(本日の日付)を取得
    const Today = await RPA.Google.Spreadsheet.getValues({
      spreadsheetId: `${SSID}`,
      range: `${SSName[i]}!B1:B1`
    });
    const SheetData = await RPA.Google.Spreadsheet.getValues({
      spreadsheetId: `${SSID}`,
      range: `${SSName[i]}!A${StartRow}:J${LastRow}`
    });
    for (let n in SheetData) {
      Row = Number(n) + 3;
      if (Today[0][0] == SheetData[n][2] && SheetData[n][5] == 'エラー') {
        await RPA.Logger.info(
          `${Row} 行目のステータスが"エラー"ですのでスキップします`
        );
      } else if (
        Today[0][0] == SheetData[n][2] &&
        SheetData[n][5] != '非公開完了'
      ) {
        await RPA.Logger.info('本日の日付 　　　　　　　→ ', Today[0][0]);
        await RPA.Logger.info('本日の日付と一致です 　　→ ', SheetData[n][2]);
        await RPA.Logger.info('この行の作業を開始します → ', Row);
        // シートから作業対象行のデータを取得
        WorkData = await RPA.Google.Spreadsheet.getValues({
          spreadsheetId: `${SSID}`,
          range: `${SSName[i]}!A${Row}:J${Row}`
        });
        await RPA.Logger.info(WorkData);
        // F列に"確認中"と記載
        await RPA.Google.Spreadsheet.setValues({
          spreadsheetId: `${SSID}`,
          range: `${SSName[i]}!F${Row}:F${Row}`,
          values: [['確認中']]
        });
        await RPA.Logger.info(
          `${Row} 行目のステータスを"確認中"に変更しました`
        );
        await Work();
      } else if (
        Today[0][0] == SheetData[n][2] &&
        SheetData[n][5] == '非公開完了'
      ) {
        await RPA.Logger.info(
          `${Row} 行目のステータスが"非公開完了"ですのでスキップします`
        );
      }
    }
  }
  await RPA.Logger.info('作業を終了します');
  if (ErrorText.length == 0) {
    await SlackPost(Slack_Text[0]);
  }
  if (ErrorText.length >= 1) {
    await SlackPost(Slack_Text2[0]);
  }
  await RPA.WebBrowser.quit();
  await RPA.sleep(1000);
  await process.exit();
}

Start();

async function SlackPost(Text) {
  await RPA.Slack.chat.postMessage({
    token: Slack_Token,
    channel: Slack_Channel,
    text: `${Text}`
  });
}

let FirstLoginFlag = 'true';
async function Work() {
  try {
    // Youtube Studioにログイン
    if (FirstLoginFlag == 'true') {
      await YoutubeLogin();
    }
    if (FirstLoginFlag == 'false') {
      await ContentsPage();
    }
    // 一度ログインしたら、以降はログインページをスキップ
    FirstLoginFlag = 'false';
    // タイトルを入力
    await TitleInput();
    // 非公開設定を開始
    await PrivateSetting();
    // 非公開シートに取得したデータを入力
    await SetDataRow();
  } catch (error) {
    ErrorText[0] = error;
    await RPA.SystemLogger.error(ErrorText);
    Slack_Text[0] = `【Youtube 非公開設定】${CurrentSSName} 年シート ${Row} 行目でエラーが発生しました！\n${ErrorText}`;
    await RPA.WebBrowser.takeScreenshot();
    await SlackPost(Slack_Text[0]);
    // F列に"エラー"と記載
    await RPA.Google.Spreadsheet.setValues({
      spreadsheetId: `${SSID}`,
      range: `${CurrentSSName}!F${Row}:F${Row}`,
      values: [['エラー']]
    });
    await Start();
  }
}

// テストは海外版、本番は日本版であることに注意！
async function YoutubeLogin() {
  await RPA.Logger.info('Youtube Studioにログインします');
  await RPA.WebBrowser.get(process.env.Youtube_Studio_Url);
  await RPA.sleep(2000);

  // ヘッドレスモードオフ（テスト）用
  // const LoginID = await RPA.WebBrowser.wait(
  //   RPA.WebBrowser.Until.elementLocated({ id: 'identifierId' }),
  //   8000
  // );
  // await RPA.WebBrowser.sendKeys(LoginID, [`${process.env.Youtube_Login_ID}`]);
  // const NextButton1 = await RPA.WebBrowser.findElementById('identifierNext');
  // await RPA.WebBrowser.mouseClick(NextButton1);
  // await RPA.sleep(5000);
  // const LoginPW = await RPA.WebBrowser.wait(
  //   RPA.WebBrowser.Until.elementLocated({ name: 'password' }),
  //   8000
  // );
  // await RPA.WebBrowser.sendKeys(LoginPW, [`${process.env.Youtube_Login_PW}`]);
  // const NextButton2 = await RPA.WebBrowser.findElementById('passwordNext');
  // await RPA.WebBrowser.mouseClick(NextButton2);

  // 本番・ヘッドレスモードオン（テスト）用
  const LoginID = await RPA.WebBrowser.wait(
    RPA.WebBrowser.Until.elementLocated({ id: 'Email' }),
    8000
  );
  await RPA.WebBrowser.sendKeys(LoginID, [process.env.Youtube_Login_ID]);
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
      ErrorText.length = 0;
      break;
    }
  }
}

async function ContentsPage() {
  await RPA.Logger.info('コンテンツ管理ページに戻ります');
  const DeliteIcon = await RPA.WebBrowser.findElementById('delete-icon');
  await RPA.WebBrowser.mouseClick(DeliteIcon);
  await RPA.sleep(3000);
}

let LoadingFlag = 'false';
async function TitleInput() {
  // 5/28 仕様変更
  while (LoadingFlag == 'false') {
    const Element = await RPA.WebBrowser.wait(
      RPA.WebBrowser.Until.elementLocated({
        className: 'style-scope ytcp-video-list-cell-video remove-default-style'
      }),
      15000
    );
    const ElementText = await Element.getText();
    if (ElementText.length < 1) {
      await RPA.Logger.info('動画の一覧が出ないため、ブラウザを更新します');
      await RPA.WebBrowser.refresh();
      await RPA.sleep(1000);
    } else {
      LoadingFlag = 'true';
      await RPA.Logger.info('＊＊＊動画の一覧が出ました＊＊＊');
    }
  }
  LoadingFlag = 'false';

  // アイコンをクリック
  const Icon = await RPA.WebBrowser.findElementById('text-input');
  await RPA.WebBrowser.mouseClick(Icon);
  await RPA.sleep(2000);

  // 「タイトル」をクリック
  // ヘッドレスモードオン（テスト）用
  // const Title = await RPA.WebBrowser.findElementById('text-item-5');

  // 本番用・ヘッドレスモードオフ（テスト）用
  const Title = await RPA.WebBrowser.findElementById('text-item-0');

  await RPA.WebBrowser.mouseClick(Title);
  await RPA.sleep(1000);

  // D列のタイトルを入力
  const InputTitle: WebElement = await RPA.WebBrowser.driver.executeScript(
    `return document.getElementsByTagName('input')[2]`
  );
  await RPA.WebBrowser.sendKeys(InputTitle, [WorkData[0][3]]);
  await RPA.sleep(1000);
  // 「適用」をクリック
  const Application = await RPA.WebBrowser.findElementById('apply-button');
  await RPA.WebBrowser.mouseClick(Application);
  await RPA.sleep(3000);
  // カーソルが被るとテキストが取得できないため、画面上部にカーソルを動かす
  const UploadVideo = await RPA.WebBrowser.wait(
    RPA.WebBrowser.Until.elementLocated({
      id: 'video-list-uploads-tab'
    }),
    5000
  );
  await RPA.WebBrowser.mouseMove(UploadVideo);
  await RPA.sleep(1000);
}

async function PrivateSetting() {
  // タイトルによっては検索に複数ヒットする場合があるため完全一致判定を行う
  for (let i = 0; i <= 29; i++) {
    const VideoTitle: WebElement = await RPA.WebBrowser.driver.executeScript(
      `return document.getElementsByClassName('style-scope ytcp-video-list-cell-video video-title-wrapper')[${i}].children[0]`
    );
    const VideoTitleText = await VideoTitle.getText();
    await RPA.Logger.info(VideoTitleText);
    if (VideoTitleText != WorkData[0][3]) {
      await RPA.Logger.info(
        `タイトルが不一致のため、URLの取得をスキップします`
      );
    } else {
      await RPA.Logger.info(VideoTitleText + ' → 一致しました');
      // 画像のURLを取得
      const Image: WebElement = await RPA.WebBrowser.driver.executeScript(
        `return document.getElementsByClassName('style-scope ytcp-video-row cell-body tablecell-video floating-column last-floating-column')[${i}].children[0].children[0].children[0]`
      );
      const ImageUrl = await Image.getAttribute('href');
      const ImageUrlSplit = await ImageUrl.split('/');
      await RPA.Logger.info(ImageUrlSplit);
      // 動画のIDを取得
      VideoID[0] = ImageUrlSplit[4];
      await RPA.Logger.info(VideoID);
      // 「公開設定」をクリック
      const PublishingSettings: WebElement = await RPA.WebBrowser.driver.executeScript(
        `return document.getElementsByClassName('style-scope ytcp-video-row cell-body tablecell-visibility')[${i}].children[0].children[0]`
      );
      await RPA.WebBrowser.mouseClick(PublishingSettings);
      await RPA.sleep(1000);
      break;
    }
  }
  // 「非公開」にチェック
  const Private = await RPA.WebBrowser.driver.findElement(By.name('PRIVATE'));
  await RPA.WebBrowser.mouseClick(Private);
  await RPA.sleep(1000);
  // 「保存」をクリック

  // テスト用
  const Cancel = await RPA.WebBrowser.findElementById('cancel-button');
  await RPA.WebBrowser.mouseClick(Cancel);

  // 本番用
  const Save = await RPA.WebBrowser.findElementById('save-button');
  // await RPA.WebBrowser.mouseClick(Save);

  await RPA.sleep(5000);
  // F列に"非公開完了"と記載
  await RPA.Google.Spreadsheet.setValues({
    spreadsheetId: `${SSID}`,
    range: `${CurrentSSName}!F${Row}:F${Row}`,
    values: [['非公開完了']]
  });
}

async function SetDataRow() {
  // 非公開シートの最終行の次の行を取得
  const WorkRow3 = await RPA.Google.Spreadsheet.getValues({
    spreadsheetId: `${SSID2}`,
    range: `${SSName2}!B${StartRow2}:B${LastRow}`
  });
  const LastRow2 = WorkRow3.length + 1;
  await RPA.Logger.info('この行に転記します　　　 → ', LastRow2);
  // B列に非公開指定日、C列に"RPA"と記載
  await RPA.Google.Spreadsheet.setValues({
    spreadsheetId: `${SSID2}`,
    range: `${SSName2}!B${LastRow2}:C${LastRow2}`,
    values: [[WorkData[0][2], 'RPA']],
    // 数字の前の ' が入力されなくなる
    parseValues: true
  });
  // E列に動画URLを記載
  await RPA.Google.Spreadsheet.setValues({
    spreadsheetId: `${SSID2}`,
    range: `${SSName2}!E${LastRow2}:E${LastRow2}`,
    values: [[`${Video_URL}=${VideoID[0]}`]]
  });
  // G列に一般URLを記載
  await RPA.Google.Spreadsheet.setValues({
    spreadsheetId: `${SSID2}`,
    range: `${SSName2}!G${LastRow2}:G${LastRow2}`,
    values: [[`${General_URL}/${VideoID[0]}`]]
  });
  // I列にタイトル、J列に公開日を記載
  await RPA.Google.Spreadsheet.setValues({
    spreadsheetId: `${SSID2}`,
    range: `${SSName2}!I${LastRow2}:J${LastRow2}`,
    values: [[WorkData[0][3], `(${WorkData[0][1]})`]]
  });
  await RPA.Logger.info(`${SSName2} シートに転記が完了しました`);
}
