import RPA from 'ts-rpa';
import { By, WebElement } from 'selenium-webdriver';

// SlackのトークンとチャンネルID
const Slack_Token = process.env.AbemaTV_RPAError_Token;
const Slack_Channel = process.env.AbemaTV_RPAError_Channel;
const Slack_Text = [`【Youtube 非公開設定】設定完了しました`];
const Slack_Text2 = [`【Youtube 非公開設定】シートにエラー行があります`];

// スプレッドシートIDとシート名を記載
// const mySSID = process.env.My_SheetID;
const SSID = process.env.Senden_Youtube_SheetID;
const SSID2 = process.env.Senden_Youtube_SheetID2;
const SSName = process.env.Senden_Youtube_SheetName;
const SSName2 = process.env.Senden_Youtube_SheetName2;
const SSName3 = process.env.Senden_Youtube_SheetName3;

// 作業対象行とデータを取得
const WorkData = [];
const Row = [];

// 動画のIDを保持する変数
const VideoID = [];
// 動画URL
const Video_URL = process.env.Youtube_Video_Url;
// 一般URL
const General_URL = process.env.Youtube_General_Url;

// エラー発生時のテキストを格納
const ErrorText = [];

async function Start(WorkData, Row) {
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
    range: `${SSName}!B1:B1`
  });
  // C列(非公開指定日)を取得
  const WorkRow = await RPA.Google.Spreadsheet.getValues({
    spreadsheetId: `${SSID}`,
    range: `${SSName}!C3:C10000`
  });
  // F列(RPAステータス)を取得
  const WorkStatus = await RPA.Google.Spreadsheet.getValues({
    spreadsheetId: `${SSID}`,
    range: `${SSName}!F3:F10000`
  });
  const WorkRow2 = await RPA.Google.Spreadsheet.getValues({
    spreadsheetId: `${SSID}`,
    range: `${SSName2}!C3:C10000`
  });
  const WorkStatus2 = await RPA.Google.Spreadsheet.getValues({
    spreadsheetId: `${SSID}`,
    range: `${SSName2}!F3:F10000`
  });
  // 2019年のシート
  for (var i in WorkRow) {
    if (Today[0][0] == WorkRow[i][0] && WorkStatus[i][0] == 'エラー') {
      Row[0] = Number(i) + 3;
      await RPA.Logger.info(
        `${Row[0]} 行目のステータスが"エラー"ですのでスキップします`
      );
    } else if (
      Today[0][0] == WorkRow[i][0] &&
      WorkStatus[i][0] != '非公開完了'
    ) {
      await RPA.Logger.info('本日の日付 　　　　　　　→ ', Today[0][0]);
      await RPA.Logger.info('本日の日付と一致です 　　→ ', WorkRow[i][0]);
      Row[0] = Number(i) + 3;
      await RPA.Logger.info('この行の作業を開始します → ', Row[0]);
      // シートから作業対象行のデータを取得
      WorkData[0] = await RPA.Google.Spreadsheet.getValues({
        spreadsheetId: `${SSID}`,
        range: `${SSName}!A${Row[0]}:F${Row[0]}`
      });
      await RPA.Logger.info(WorkData[0]);
      // F列に"確認中"と記載
      await RPA.Google.Spreadsheet.setValues({
        spreadsheetId: `${SSID}`,
        range: `${SSName}!F${Row[0]}:F${Row[0]}`,
        values: [['確認中']]
      });
      await RPA.Logger.info(
        `${Row[0]} 行目のステータスを"確認中"に変更しました`
      );
      await Work();
    } else if (
      Today[0][0] == WorkRow[i][0] &&
      WorkStatus[i][0] == '非公開完了'
    ) {
      Row[0] = Number(i) + 3;
      await RPA.Logger.info(
        `${Row[0]} 行目のステータスが"非公開完了"ですのでスキップします`
      );
    }
  }
  if (Today[0][0] != WorkRow[i][0]) {
    SheetFlag[0] = '2020';
    await RPA.Logger.info(`${SheetFlag[0]} 年のシートに移動します`);
    // 2020年のシート
    for (var i in WorkRow2) {
      if (Today[0][0] == WorkRow2[i][0] && WorkStatus2[i][0] == 'エラー') {
        Row[0] = Number(i) + 3;
        await RPA.Logger.info(
          `${Row[0]} 行目のステータスが"エラー"ですのでスキップします`
        );
      } else if (
        Today[0][0] == WorkRow2[i][0] &&
        WorkStatus2[i][0] != '非公開完了'
      ) {
        await RPA.Logger.info('本日の日付 　　　　　　　→ ', Today[0][0]);
        await RPA.Logger.info('本日の日付と一致です 　　→ ', WorkRow2[i][0]);
        Row[0] = Number(i) + 3;
        await RPA.Logger.info('この行の作業を開始します → ', Row[0]);
        // シートから作業対象行のデータを取得
        WorkData[0] = await RPA.Google.Spreadsheet.getValues({
          spreadsheetId: `${SSID}`,
          range: `${SSName2}!A${Row[0]}:F${Row[0]}`
        });
        await RPA.Logger.info(WorkData[0]);
        // F列に"確認中"と記載
        await RPA.Google.Spreadsheet.setValues({
          spreadsheetId: `${SSID}`,
          range: `${SSName2}!F${Row[0]}:F${Row[0]}`,
          values: [['確認中']]
        });
        await RPA.Logger.info(
          `${Row[0]} 行目のステータスを"確認中"に変更しました`
        );
        await Work();
      } else if (
        Today[0][0] == WorkRow2[i][0] &&
        WorkStatus2[i][0] == '非公開完了'
      ) {
        Row[0] = Number(i) + 3;
        await RPA.Logger.info(
          `${Row[0]} 行目のステータスが"非公開完了"ですのでスキップします`
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

Start(WorkData, Row);

async function SlackPost(Text) {
  await RPA.Slack.chat.postMessage({
    token: Slack_Token,
    channel: Slack_Channel,
    text: `${Text}`
  });
}

const FirstLoginFlag = ['true'];
const SheetFlag = ['2019'];
async function Work() {
  // Youtube Studioにログイン
  if (FirstLoginFlag[0] == 'true') {
    await YoutubeLogin();
  }
  if (FirstLoginFlag[0] == 'false') {
    await YoutubeLogin2();
  }
  // 一度ログインしたら、以降はログインページをスキップ
  FirstLoginFlag[0] = 'false';
  // 非公開設定を開始
  await PrivateSetting(WorkData, Row);
}

// テストは海外版、本番は日本版であることに注意！
async function YoutubeLogin() {
  try {
    // throw new Error();
    await RPA.Logger.info('Youtube Studioにログインします');
    await RPA.WebBrowser.get(process.env.Youtube_Studio_Url);
    await RPA.sleep(2000);
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
  } catch (error) {
    await CatchError(error);
  }
}

async function YoutubeLogin2() {
  await RPA.Logger.info('コンテンツ管理ページに直接遷移します');
  await RPA.WebBrowser.get(process.env.Youtube_Studio_Url);
  await RPA.sleep(3000);
}

async function PrivateSetting(WorkData, Row) {
  try {
    // アイコンをクリック
    const Icon = await RPA.WebBrowser.wait(
      RPA.WebBrowser.Until.elementLocated({ id: 'filter-icon' }),
      5000
    );
    await RPA.WebBrowser.mouseClick(Icon);
    await RPA.sleep(1000);
    // 「タイトル」をクリック

    // テスト用
    // const Title = await RPA.WebBrowser.findElementById('text-item-5');
    // const Title = await RPA.WebBrowser.wait(
    //   RPA.WebBrowser.Until.elementLocated({ id: 'text-item-5' }),
    //   5000
    // );

    // 本番用
    // const Title = await RPA.WebBrowser.findElementById('text-item-0');
    const Title = await RPA.WebBrowser.wait(
      RPA.WebBrowser.Until.elementLocated({ id: 'text-item-0' }),
      15000
    );

    await RPA.WebBrowser.mouseClick(Title);
    await RPA.sleep(1000);
    // D列のタイトルを入力
    const InputTitle: WebElement = await RPA.WebBrowser.driver.executeScript(
      `return document.getElementsByTagName('input')[2]`
    );
    await RPA.WebBrowser.sendKeys(InputTitle, [WorkData[0][0][3]]);
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
    // タイトルによっては検索に複数ヒットする場合があるため完全一致判定を行う
    for (let i = 0; i <= 29; i++) {
      const VideoTitle: WebElement = await RPA.WebBrowser.driver.executeScript(
        `return document.getElementsByClassName('style-scope ytcp-video-list-cell-video video-title-wrapper')[${Number(
          i
        )}].children[0]`
      );
      const VideoTitleText = await VideoTitle.getText();
      await RPA.Logger.info(VideoTitleText);
      if (VideoTitleText != WorkData[0][0][3]) {
        await RPA.Logger.info(
          `タイトルが不一致のため、URLの取得をスキップします`
        );
      } else {
        await RPA.Logger.info(VideoTitleText + ' → 一致しました');
        // 画像のURLを取得
        const Image: WebElement = await RPA.WebBrowser.driver.executeScript(
          `return document.getElementsByClassName('style-scope ytcp-video-row cell-body tablecell-video floating-column last-floating-column')[${Number(
            i
          )}].children[0].children[0].children[0]`
        );
        const ImageUrl = await Image.getAttribute('href');
        const ImageUrlSplit = await ImageUrl.split('/');
        await RPA.Logger.info(ImageUrlSplit);
        // 動画のIDを取得
        VideoID[0] = ImageUrlSplit[4];
        await RPA.Logger.info(VideoID);
        // 「公開設定」をクリック
        const PublishingSettings: WebElement = await RPA.WebBrowser.driver.executeScript(
          `return document.getElementsByClassName('style-scope ytcp-video-row cell-body tablecell-visibility')[${Number(
            i
          )}].children[0].children[0]`
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
    const Save = await RPA.WebBrowser.findElementById('save-button');
    await RPA.WebBrowser.mouseClick(Save);
    await RPA.sleep(5000);
    // F列に"非公開完了"と記載
    if (SheetFlag[0] == '2019') {
      await RPA.Google.Spreadsheet.setValues({
        spreadsheetId: `${SSID}`,
        range: `${SSName}!F${Row[0]}:F${Row[0]}`,
        values: [['非公開完了']]
      });
    }
    if (SheetFlag[0] == '2020') {
      await RPA.Google.Spreadsheet.setValues({
        spreadsheetId: `${SSID}`,
        range: `${SSName2}!F${Row[0]}:F${Row[0]}`,
        values: [['非公開完了']]
      });
    }
    // 非公開シートに取得したデータを入力
    await SetDataRow(WorkData, Video_URL, VideoID, General_URL);
  } catch (error) {
    await CatchError(error);
  }
}

async function SetDataRow(WorkData, Video_URL, VideoID, General_URL) {
  // 非公開シートの最終行の次の行を取得
  const WorkRow3 = await RPA.Google.Spreadsheet.getValues({
    spreadsheetId: `${SSID2}`,
    range: `${SSName3}!B1:B30000`
  });
  const LastRow = WorkRow3.length + 1;
  await RPA.Logger.info('この行に転記します　　　 → ', LastRow);
  // B列に非公開指定日を記載
  await RPA.Google.Spreadsheet.setValues({
    spreadsheetId: `${SSID2}`,
    range: `${SSName3}!B${LastRow}:B${LastRow}`,
    values: [[WorkData[0][0][2]]],
    // 数字の前の ' が入力されなくなる
    parseValues: true
  });
  // C列に"RPA"と記載
  await RPA.Google.Spreadsheet.setValues({
    spreadsheetId: `${SSID2}`,
    range: `${SSName3}!C${LastRow}:C${LastRow}`,
    values: [['RPA']]
  });
  // E列に動画URLを記載
  await RPA.Google.Spreadsheet.setValues({
    spreadsheetId: `${SSID2}`,
    range: `${SSName3}!E${LastRow}:E${LastRow}`,
    values: [[`${Video_URL}=${VideoID[0]}`]]
  });
  // G列に一般URLを記載
  await RPA.Google.Spreadsheet.setValues({
    spreadsheetId: `${SSID2}`,
    range: `${SSName3}!G${LastRow}:G${LastRow}`,
    values: [[`${General_URL}/${VideoID[0]}`]]
  });
  // I列にタイトルを記載
  await RPA.Google.Spreadsheet.setValues({
    spreadsheetId: `${SSID2}`,
    range: `${SSName3}!I${LastRow}:I${LastRow}`,
    values: [[WorkData[0][0][3]]]
  });
  // J列に公開日を記載
  await RPA.Google.Spreadsheet.setValues({
    spreadsheetId: `${SSID2}`,
    range: `${SSName3}!J${LastRow}:J${LastRow}`,
    values: [[`(${WorkData[0][0][1]})`]]
  });
  await RPA.Logger.info('「非公開」シートにも転記が完了しました');
}

async function CatchError(error) {
  ErrorText[0] = error;
  await RPA.SystemLogger.error(ErrorText);
  Slack_Text[0] = `【Youtube 非公開設定】${SheetFlag[0]} 年シート ${Row[0]} 行目でエラーが発生しました！\n${ErrorText}`;
  await RPA.WebBrowser.takeScreenshot();
  await SlackPost(Slack_Text[0]);
  // F列に"エラー"と記載
  if (SheetFlag[0] == '2019') {
    await RPA.Google.Spreadsheet.setValues({
      spreadsheetId: `${SSID}`,
      range: `${SSName}!F${Row[0]}:F${Row[0]}`,
      values: [['エラー']]
    });
  }
  if (SheetFlag[0] == '2020') {
    await RPA.Google.Spreadsheet.setValues({
      spreadsheetId: `${SSID}`,
      range: `${SSName2}!F${Row[0]}:F${Row[0]}`,
      values: [['エラー']]
    });
  }
  await Start(WorkData, Row);
}
