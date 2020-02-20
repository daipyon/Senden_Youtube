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
const SSName = process.env.Senden_Youtube_SheetName;
const SSName2 = process.env.Senden_Youtube_SheetName2;
// Youtubeアカウント
const Youtube_ID = process.env.Youtube_Login_ID;
const Youtube_PW = process.env.Youtube_Login_PW;
// YoutubeStudioのURL
const Youtube_Studio = process.env.Youtube_Studio_Url;

// 作業対象行とデータを取得
const WorkData = [];
const Row = [];
const FirstLoginFlag = ['true'];
async function WorkStart() {
  await RPA.Google.authorize({
    accessToken: process.env.GOOGLE_ACCESS_TOKEN,
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
}

async function Start(WorkData, Row) {
  await RPA.Google.authorize({
    //accessToken: process.env.GOOGLE_ACCESS_TOKEN,
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
  const WorkRow2 = await RPA.Google.Spreadsheet.getValues({
    spreadsheetId: `${SSID}`,
    range: `${SSName2}!C2:C10000`
  });
  // 2019年のシート
  for (var i in WorkRow) {
    if (Today[0][0] == WorkRow[i][0]) {
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
    }
  }
  if (Today[0][0] != WorkRow[i][0]) {
    RPA.Logger.info(`${SSName2} 年のシートに移動します`);
    // 2020年のシート
    for (var i in WorkRow2) {
      if (Today[0][0] == WorkRow2[i][0]) {
        RPA.Logger.info('本日の日付 　　　　　　　→ ', Today[0][0]);
        RPA.Logger.info('本日の日付と一致です 　　→ ', WorkRow2[i][0]);
        Row[0] = Number(i) + 2;
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
      }
    }
  }
  RPA.Logger.info('作業を終了します');
  await RPA.WebBrowser.quit();
}

Start(WorkData, Row);

async function AbemaLogin() {
  // Abemaログイン
  await RPA.WebBrowser.get(Youtube_Studio);
  await RPA.sleep(1000);
  try {
    const GoogleLoginID = await RPA.WebBrowser.wait(
      RPA.WebBrowser.Until.elementLocated({ xpath: '//*[@id="identifierId"]' }),
      5000
    );
    await RPA.WebBrowser.sendKeys(GoogleLoginID, [`${Youtube_ID}`]);
    await RPA.sleep(1000);
  } catch {
    return;
  }
  const NextButton1 = await RPA.WebBrowser.findElementByXPath(
    '//*[@id="identifierNext"]'
  );
  await RPA.WebBrowser.mouseClick(NextButton1);
  await RPA.sleep(1000);
  const GoogleLoginPW = await RPA.WebBrowser.wait(
    RPA.WebBrowser.Until.elementLocated({
      xpath: '//*[@id="password"]/div[1]/div/div[1]/input'
    }),
    5000
  );
  await RPA.WebBrowser.sendKeys(GoogleLoginPW, [`${Youtube_PW}`]);
  const NextButton2 = await RPA.WebBrowser.findElementByXPath(
    '//*[@id="passwordNext"]'
  );
  await RPA.WebBrowser.mouseClick(NextButton2);
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
  RPA.Logger.info(WorkData[0][0][3]);
  const InputTitle = await RPA.WebBrowser.findElementByXPath(
    '//*[@id="value-input"]'
  );
  await RPA.WebBrowser.sendKeys(InputTitle, [WorkData[0][0][3]]);
  await RPA.sleep(1000);
  // 「適用」をクリック
  const Application = await RPA.WebBrowser.findElementByXPath(
    '//*[@id="apply-button"]/div'
  );
  await RPA.WebBrowser.mouseClick(Application);
  await RPA.sleep(3000);
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
