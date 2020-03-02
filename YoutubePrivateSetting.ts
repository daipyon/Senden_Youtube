import RPA from 'ts-rpa';
import { WebDriver, By, FileDetector, Key } from 'selenium-webdriver';
import { rootCertificates } from 'tls';
import { worker } from 'cluster';
import { cachedDataVersionTag } from 'v8';
import { start } from 'repl';
import { Command } from 'selenium-webdriver/lib/command';
import { Driver } from 'selenium-webdriver/safari';
import { Dataset } from '@google-cloud/bigquery';
// デバッグログを最小限(INFOのみ)にする ※[DEBUG]が非表示になる
RPA.Logger.level = 'INFO';

// スプレッドシートIDとシート名を記載
const SSID = process.env.Senden_Youtube_SheetID;
const SSID2 = process.env.Senden_Youtube_SheetID2;
const SSName = process.env.Senden_Youtube_SheetName;
const SSName2 = process.env.Senden_Youtube_SheetName2;
const SSName3 = process.env.Senden_Youtube_SheetName3;
// Youtubeアカウント
const Youtube_ID = process.env.Youtube_Login_ID;
const Youtube_PW = process.env.Youtube_Login_PW;
// YoutubeStudioのURL
const Youtube_Studio = process.env.Youtube_Studio_Url;
// 動画URL
const Video_URL = process.env.Youtube_Video_Url;
// 一般URL
const General_URL = process.env.Youtube_General_Url;
// SlackのトークンとチャンネルID
const Slack_Token = process.env.AbemaTV_RPAError_Token;
const Slack_Channel = process.env.AbemaTV_RPAError_Channel;

// 作業対象行とデータを取得
const WorkData = [];
const Row = [];
// 動画のIDを保持する変数
const VideoID = [];
const FirstLoginFlag = ['true'];
async function WorkStart() {
  await RPA.Google.authorize({
    // accessToken: process.env.GOOGLE_ACCESS_TOKEN,
    refreshToken: process.env.GOOGLE_REFRESH_TOKEN,
    tokenType: 'Bearer',
    expiryDate: parseInt(process.env.GOOGLE_EXPIRY_DATE, 10)
  });
  // AAAMSにログイン
  if (FirstLoginFlag[0] == 'true') {
    await AbemaLogin();
  }
  if (FirstLoginFlag[0] == 'false') {
    await AbemaLogin2();
  }
  // 一度ログインしたら、以降はログインページをスキップ
  FirstLoginFlag[0] = 'false';

  // 非公開設定を開始
  await PrivateSetting(WorkData, Row);

  // 非公開シートに取得したデータを入力
  await SetDataRow(WorkData, Video_URL, VideoID, General_URL);
}

async function Start(WorkData, Row) {
  try {
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
      range: `${SSName}!F3:F10000`
    });
    // 2019年のシート
    for (var i in WorkRow) {
      if (Today[0][0] == WorkRow[i][0] && WorkStatus[i][0] != '非公開完了') {
        RPA.Logger.info('本日の日付 　　　　　　　→ ', Today[0][0]);
        RPA.Logger.info('本日の日付と一致です 　　→ ', WorkRow[i][0]);
        Row[0] = Number(i) + 3;
        RPA.Logger.info('この行の作業を開始します → ', Row[0]);
        // シートから作業対象行のデータを取得
        WorkData[0] = await RPA.Google.Spreadsheet.getValues({
          spreadsheetId: `${SSID}`,
          range: `${SSName}!A${Row[0]}:F${Row[0]}`
        });
        RPA.Logger.info(WorkData[0]);
        // F列に"確認中"と記載
        await RPA.Google.Spreadsheet.setValues({
          spreadsheetId: `${SSID}`,
          range: `${SSName}!F${Row[0]}:F${Row[0]}`,
          values: [['確認中']]
        });
        RPA.Logger.info(`${Row[0]} 行目のステータスを"確認中"に変更しました`);
        await WorkStart();
      } else if (
        Today[0][0] == WorkRow[i][0] &&
        WorkStatus[i][0] == '非公開完了'
      ) {
        Row[0] = Number(i) + 3;
        RPA.Logger.info(
          `${Row[0]} 行目のステータスが"非公開完了"ですのでスキップします`
        );
      }
    }
    if (Today[0][0] != WorkRow[i][0]) {
      RPA.Logger.info(`${SSName2} 年のシートに移動します`);
      // 2020年のシート
      for (var i in WorkRow2) {
        if (
          Today[0][0] == WorkRow2[i][0] &&
          WorkStatus2[i][0] != '非公開完了'
        ) {
          RPA.Logger.info('本日の日付 　　　　　　　→ ', Today[0][0]);
          RPA.Logger.info('本日の日付と一致です 　　→ ', WorkRow2[i][0]);
          Row[0] = Number(i) + 3;
          RPA.Logger.info('この行の作業を開始します → ', Row[0]);
          // シートから作業対象行のデータを取得
          WorkData[0] = await RPA.Google.Spreadsheet.getValues({
            spreadsheetId: `${SSID}`,
            range: `${SSName2}!A${Row[0]}:F${Row[0]}`
          });
          RPA.Logger.info(WorkData[0]);
          // F列に"確認中"と記載
          await RPA.Google.Spreadsheet.setValues({
            spreadsheetId: `${SSID}`,
            range: `${SSName2}!F${Row[0]}:F${Row[0]}`,
            values: [['確認中']]
          });
          RPA.Logger.info(`${Row[0]} 行目のステータスを"確認中"に変更しました`);
          await WorkStart();
        } else if (
          Today[0][0] == WorkRow2[i][0] &&
          WorkStatus2[i][0] == '非公開完了'
        ) {
          Row[0] = Number(i) + 3;
          RPA.Logger.info(
            `${Row[0]} 行目のステータスが"非公開完了"ですのでスキップします`
          );
        }
      }
    }
  } catch (error) {
    RPA.Logger.info('エラーが発生しました！');
    await RPA.WebBrowser.takeScreenshot();
    // Slackにも通知
    await RPA.Slack.chat.postMessage({
      token: Slack_Token,
      channel: Slack_Channel,
      text: '【宣伝_Youtube 非公開設定】でエラーが発生しました！'
    });
  }
  RPA.Logger.info('作業を終了します');
  await RPA.WebBrowser.quit();
}

Start(WorkData, Row);

async function AbemaLogin() {
  // Abemaログイン
  await RPA.WebBrowser.get(Youtube_Studio);
  await RPA.sleep(2000);
  const LoginID = await RPA.WebBrowser.wait(
    RPA.WebBrowser.Until.elementLocated({ id: 'Email' }),
    8000
  );
  await RPA.WebBrowser.sendKeys(LoginID, [Youtube_ID]);
  const NextButton1 = await RPA.WebBrowser.findElementById('next');
  await RPA.WebBrowser.mouseClick(NextButton1);
  await RPA.sleep(5000);
  const GoogleLoginPW_1 = await RPA.WebBrowser.findElementsById('Passwd');
  const GoogleLoginPW_2 = await RPA.WebBrowser.findElementsById('password');
  await RPA.Logger.info(`GoogleLoginPW_1 ` + GoogleLoginPW_1.length);
  await RPA.Logger.info(`GoogleLoginPW_2 ` + GoogleLoginPW_2.length);
  if (GoogleLoginPW_1.length == 1) {
    await RPA.WebBrowser.sendKeys(GoogleLoginPW_1[0], [`${Youtube_PW}`]);
    const NextButton2 = await RPA.WebBrowser.findElementById('signIn');
    await RPA.WebBrowser.mouseClick(NextButton2);
  }
  if (GoogleLoginPW_2.length == 1) {
    const GoogleLoginPW = await RPA.WebBrowser.findElementById('password');
    await RPA.WebBrowser.sendKeys(GoogleLoginPW, [`${Youtube_PW}`]);
    const NextButton2 = await RPA.WebBrowser.findElementById('submit');
    await RPA.WebBrowser.mouseClick(NextButton2);
  }
  await RPA.sleep(8000);
}

async function AbemaLogin2() {
  RPA.Logger.info('コンテンツ管理ページに直接遷移します');
  await RPA.WebBrowser.get(Youtube_Studio);
  await RPA.sleep(3000);
}

async function PrivateSetting(WorkData, Row) {
  // アイコンをクリック
  const Icon = await RPA.WebBrowser.findElementByXPath(
    '//*[@id="filter-icon"]'
  );
  await RPA.WebBrowser.mouseClick(Icon);
  await RPA.sleep(1000);
  // 「タイトル」をクリック
  const Title = await RPA.WebBrowser.findElementByXPath(
    '//*[@id="text-item-0"]'
  );
  await RPA.WebBrowser.mouseClick(Title);
  await RPA.sleep(1000);
  // D列のタイトルを入力
  const InputTitle = await RPA.WebBrowser.findElementByXPath(
    '//*[@id="input-1"]/input'
  );
  await RPA.WebBrowser.sendKeys(InputTitle, [WorkData[0][0][3]]);
  await RPA.sleep(1000);
  // 「適用」をクリック
  const Application = await RPA.WebBrowser.findElementByXPath(
    '//*[@id="apply-button"]/div'
  );
  await RPA.WebBrowser.mouseClick(Application);
  await RPA.sleep(3000);
  // 画像のURLを取得
  const ImageUrl = await RPA.WebBrowser.driver
    .findElement(By.css('#img-with-fallback'))
    .getAttribute('src');
  const ImageUrlSplit = ImageUrl.split('/');
  RPA.Logger.info(ImageUrlSplit);
  // 動画のIDを取得
  VideoID[0] = ImageUrlSplit[4];
  RPA.Logger.info(VideoID);
  // 「公開設定」をクリック
  const PublishingSettings = await RPA.WebBrowser.findElementByXPath(
    '//*[@id="row-container"]/div[3]/div/div'
  );
  await RPA.WebBrowser.mouseClick(PublishingSettings);
  await RPA.sleep(1000);
  // 「非公開」にチェック
  const Private = await RPA.WebBrowser.findElementByXPath(
    '//*[@id="privacy-radios"]/paper-radio-button[3]'
  );
  await RPA.WebBrowser.mouseClick(Private);
  await RPA.sleep(1000);
  // 「保存」をクリック
  const Save = await RPA.WebBrowser.findElementByXPath(
    '//*[@id="save-button"]/div'
  );
  await RPA.WebBrowser.mouseClick(Save);
  await RPA.sleep(5000);
  // F列に"非公開完了"と記載
  await RPA.Google.Spreadsheet.setValues({
    spreadsheetId: `${SSID}`,
    range: `${SSName}!F${Row[0]}:F${Row[0]}`,
    values: [['非公開完了']]
  });
}

async function SetDataRow(WorkData, Video_URL, VideoID, General_URL) {
  // 非公開シートの最終行の次の行を取得
  const WorkRow3 = await RPA.Google.Spreadsheet.getValues({
    spreadsheetId: `${SSID2}`,
    range: `${SSName3}!B1:B30000`
  });
  const LastRow = WorkRow3.length + 1;
  RPA.Logger.info('この行に転記します　　　 → ', LastRow);
  // B列に非公開指定日を記載
  await RPA.Google.Spreadsheet.setValues({
    spreadsheetId: `${SSID2}`,
    range: `${SSName3}!B${LastRow}:B${LastRow}`,
    values: [[WorkData[0][0][2]]]
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
  RPA.Logger.info('非公開シートにも転記が完了しました');
}
