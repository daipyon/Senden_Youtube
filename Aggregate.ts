import RPA from 'ts-rpa';
import { WebElement } from 'selenium-webdriver';
const moment = require('moment');

// SlackのトークンとチャンネルID
const Slack_Token = process.env.AbemaTV_RPAError_Token;
const Slack_Channel = process.env.AbemaTV_RPAError_Channel;
const Slack_Text = [`【Youtube 集計】集計完了しました`];

// RPAトリガーシートのID
// const mySSID = process.env.My_SheetID2;
// const mySSID2 = [];
const SSID = process.env.Senden_Youtube_SheetID3;
const SSID2 = [];
// シート名
const SSName4 = ['News', '公式', 'バラエティ', 'Mリーグ'];

// 作業するスプレッドシートから読み込む行数を記載
const StartRow = 4;
const LastRow = 100;
const StartRow2 = 2;
const LastRow2 = 200;

// 現在のシートIDとシート名
const CurrentSSID = [];
const CurrentSSName = [];

// 作業対象行とデータを取得
const WorkData = [];
const Row = [];
const Row2 = [];
const Row3 = [];

// 日付
let date = [];

// 番組名のリスト
let ProgramList = [];
// 取得した番組名を格納
let ProgramName = [];
// タイトルのリスト
let TitleList = [];
let TitleList2 = [];

// エラー発生時のテキストを格納
const ErrorText = [];

// 今月分からスタートする場合はここをコメントアウト
let MonthFlag = '前月';
// 前月分からスタートする場合はここをコメントアウト
// let MonthFlag = '今月';

let JudgeFlag = '';
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
    // シートからスプレッドシートIDを取得
    const LastMonthSheetID = await RPA.Google.Spreadsheet.getValues({
      spreadsheetId: `${SSID}`,
      // テスト用
      // range: `RPAトリガーシート!B10:B10`

      // 本番用
      range: `シート1!B10:B10`
    });
    const ThisMonthSheetID = await RPA.Google.Spreadsheet.getValues({
      spreadsheetId: `${SSID}`,
      // テスト用
      // range: `RPAトリガーシート!B13:B13`

      // 本番用
      range: `シート1!B13:B13`
    });
    // 今月分からスタートする場合は上をコメントアウト
    await SSID2.push(LastMonthSheetID[0][0]);
    await SSID2.push(ThisMonthSheetID[0][0]);
    await RPA.Logger.info(SSID2);
    await RPA.Logger.info(`【${MonthFlag}】分のシートです`);
    if (MonthFlag == '前月') {
      await RPA.Logger.info('視聴回数のみ取得を開始します');
    }
    if (MonthFlag == '今月') {
      await RPA.Logger.info(
        'タイトル、アップロード日、視聴回数の取得を開始します'
      );
    }
    // 番組名をリストに格納
    for (let j in SSID2) {
      for (let n in SSName4) {
        CurrentSSID[0] = SSID2[j];
        CurrentSSName[0] = SSName4[n];
        const Program = await RPA.Google.Spreadsheet.getValues({
          spreadsheetId: `${SSID2[j]}`,
          range: `${SSName4[n]}!B${StartRow}:B${LastRow}`
        });
        for (let i in Program) {
          if (Program[i][0] != undefined) {
            await ProgramList.push(Program[i][0]);
          } else {
            break;
          }
        }
        // M列(本日の日付)を取得
        const Today = await RPA.Google.Spreadsheet.getValues({
          spreadsheetId: `${SSID2[j]}`,
          range: `${SSName4[n]}!M1:M1`
        });
        // 本日の日付をフォーマット変更して取得
        const NewToday = moment().format('YYYY/MM/DD');
        await RPA.Logger.info('本日の日付　　　　　→ ' + NewToday);
        // 前々月の日付をフォーマット変更して取得
        const TwoMonthsBefore = moment()
          .add(-2, 'month')
          .format('YYYY/MM');
        await RPA.Logger.info('前々月　　　　　　　→ ' + TwoMonthsBefore);
        date = await RPA.Google.Spreadsheet.getValues({
          spreadsheetId: `${SSID2[j]}`,
          range: `${SSName4[n]}!Q3:CA${LastRow2}`
        });
        const ProgramName2 = await RPA.Google.Spreadsheet.getValues({
          spreadsheetId: `${SSID2[j]}`,
          range: `${SSName4[n]}!L${StartRow2}:L${LastRow2}`
        });
        for (let v in ProgramList) {
          for (let k in ProgramName2) {
            if (ProgramName2[k][0] == ProgramList[v]) {
              // 格納したタイトルリストを空にする
              TitleList = [];
              Row[0] = Number(k) + StartRow;
              await RPA.Logger.info(`${Row[0]} 行目から取得を開始します`);
              if (MonthFlag == '前月') {
                // テスト④
                // 空白行を確認
                for (let i = 0; i <= LastRow2; i++) {
                  const NoTitle = await RPA.Google.Spreadsheet.getValues({
                    spreadsheetId: `${SSID2[j]}`,
                    range: `${SSName4[n]}!Q${Row[0] + i}:Q${Row[0] + i}`
                  });
                  if (NoTitle != undefined) {
                    Row2[0] = Row[0] + i + 1;
                    // await RPA.Logger.info(
                    //   `${Row[0]} 行目にタイトルがありますが、${Row2[0]} 行目にタイトルがあるか確認します`
                    // );
                    break;
                    // } else {
                    // await RPA.Logger.info(`${Row[0] + i} 行目　空白です`);
                    // JudgeFlag = 'false';
                    // break;
                  }
                }
                // if (JudgeFlag == '') {
                //   await RPA.Logger.info(`パターン 1 です`);
                // 次の行が空白かどうか確認
                for (let i = 0; i <= LastRow2; i++) {
                  const NoTitle2 = await RPA.Google.Spreadsheet.getValues({
                    spreadsheetId: `${SSID2[j]}`,
                    // range: `${SSName4[n]}!Q${Row2[0] + i}:Q${Row2[0] + i}`
                    range: `${SSName4[n]}!Q${Row[0] + i}:Q${Row[0] + i}`
                  });
                  // if (NoTitle2 == undefined) {
                  // await RPA.Logger.info(`${Row2[0] + i} 行目　空白です`);
                  // JudgeFlag = 'true';
                  // break;
                  if (NoTitle2 != undefined) {
                    Row3[0] = Row[0] + i;
                    await RPA.Logger.info(
                      `${Row3[0]} 行目にタイトルがあります`
                    );
                    break;
                  }
                }
                //   if (JudgeFlag == 'true') {
                //     await RPA.Logger.info(`パターン 3 です`);
                //     // タイトルのある行を取得
                //     for (let i = 0; i <= LastRow2; i++) {
                //       const JudgeTitle = await RPA.Google.Spreadsheet.getValues(
                //         {
                //           spreadsheetId: `${SSID2[j]}`,
                //           range: `${SSName4[n]}!Q${Row[0] + i}:Q${Row[0] + i}`
                //         }
                //       );
                //       if (JudgeTitle == undefined) {
                //         await RPA.Logger.info(`${Row[0] + i} 行目　空白です`);
                //       } else {
                //         Row3[0] = Row[0] + i;
                //         await RPA.Logger.info(
                //           `${Row3[0]} 行目にタイトルがあります`
                //         );
                //         break;
                //       }
                //     }
                //   }
                // } else {
                //   await RPA.Logger.info(`パターン 2 です`);
                //   // タイトルのある行を取得
                //   for (let i = 0; i <= LastRow2; i++) {
                //     const JudgeTitle = await RPA.Google.Spreadsheet.getValues({
                //       spreadsheetId: `${SSID2[j]}`,
                //       range: `${SSName4[n]}!Q${Row[0] + i}:Q${Row[0] + i}`
                //     });
                //     if (JudgeTitle == undefined) {
                //       await RPA.Logger.info(`${Row[0] + i} 行目　空白です`);
                //     } else {
                //       Row3[0] = Row[0] + i;
                //       await RPA.Logger.info(
                //         `${Row3[0]} 行目にタイトルがあります`
                //       );
                //       break;
                //     }
                //   }
                // }
                // テスト①
                if (ProgramName2[k][0] == 'アベプラ') {
                  await RPA.Logger.info(
                    `【アベプラ】のため、${Row3[0]} 行目から取得を開始します`
                    // `【アベプラ】のため、${Row[0]} 行目から取得を開始します`
                  );
                  WorkData[0] = await RPA.Google.Spreadsheet.getValues({
                    spreadsheetId: `${SSID2[j]}`,
                    range: `${SSName4[n]}!Q${Row3[0]}:Q${LastRow2}`
                    // range: `${SSName4[n]}!Q${Row[0]}:Q${LastRow2}`
                  });
                } else {
                  WorkData[0] = await RPA.Google.Spreadsheet.getValues({
                    spreadsheetId: `${SSID2[j]}`,
                    range: `${SSName4[n]}!Q${Row3[0]}:Q${LastRow2}`
                    // range: `${SSName4[n]}!Q${Row[0]}:Q${LastRow2}`
                  });
                }
                for (let i in WorkData[0]) {
                  if (WorkData[0][i][0] != undefined) {
                    await TitleList.push(WorkData[0][i][0]);
                  }
                  if (
                    WorkData[0][i][0] == 'タイトル' ||
                    WorkData[0][i][0] == 'ここまで'
                  ) {
                    // テスト④
                    break;
                    if (WorkData[0][i][0] == 'タイトル') {
                      break;
                    }
                    if (WorkData[0][i][0] == 'ここまで') {
                      break;
                    }
                  }
                }
              } else {
                // 保留
                // 先に上部のタイトルを取得
                // for (let i = 0; i <= LastRow2; i++) {
                //   const JudgeTitle = await RPA.Google.Spreadsheet.getValues({
                //     spreadsheetId: `${SSID2[j]}`,
                //     range: `${SSName4[n]}!Q${Row[0] + i}:Q${Row[0] + i}`
                //   });
                //   if (JudgeTitle != undefined) {
                //     await TitleList2.push(JudgeTitle[0][0]);
                //   } else {
                //     Row2[0] = Number(i) + Row[0];
                //     break;
                //   }
                // }
                // if (TitleList2.length < 1) {
                //   await RPA.Logger.info('上部のタイトル　　  →', TitleList2);
                // }
                // await RPA.Logger.info(`${Row2[0]} 行目から空白があります`);
                // 空白行を確認
                for (let i = 0; i <= LastRow2; i++) {
                  const NoTitle = await RPA.Google.Spreadsheet.getValues({
                    spreadsheetId: `${SSID2[j]}`,
                    range: `${SSName4[n]}!Q${Row[0] + i}:Q${Row[0] + i}`
                    // range: `${SSName4[n]}!Q${Row2[0] + i}:Q${Row2[0] + i}`
                  });
                  if (NoTitle != undefined) {
                    Row3[0] = Row[0] + i;
                    break;
                  }
                  // else {
                  // }
                }
                await RPA.Logger.info(`${Row3[0]} 行目にタイトルがあります`);
                WorkData[0] = await RPA.Google.Spreadsheet.getValues({
                  spreadsheetId: `${SSID2[j]}`,
                  range: `${SSName4[n]}!Q${Row3[0]}:Q${LastRow2}`
                });
                for (let i in WorkData[0]) {
                  if (WorkData[0][i][0] != undefined) {
                    await TitleList.push(WorkData[0][i][0]);
                  }
                  if (
                    WorkData[0][i][0] == 'タイトル' ||
                    WorkData[0][i][0] == 'ここまで'
                  ) {
                    // テスト④
                    break;
                    if (WorkData[0][i][0] == 'タイトル') {
                      break;
                    }
                    if (WorkData[0][i][0] == 'ここまで') {
                      break;
                    }
                  }
                }
              }
              if (ProgramName[0] < 1) {
                await ProgramName.push(ProgramList[v]);
              } else {
                // 格納した番組名の先頭を削除
                await ProgramName.shift();
                await ProgramName.push(ProgramList[v]);
              }
              // タイトルリストの末尾を削除
              await TitleList.pop();
              await Work(
                // テスト①
                // ProgramName,
                // TitleList,
                Today,
                NewToday,
                TwoMonthsBefore
                // date,
                // WorkData
              );
              break;
            }
          }
        }
        // 格納した番組名リストを空にする
        ProgramList = [];
        await RPA.Logger.info('番組名リストをクリアします');
        await RPA.Logger.info(ProgramList);
      }
      if (Number(j) == 0) {
        // 今月分からスタートする場合はここをコメントアウト
        MonthFlag = '今月';
        await RPA.Logger.info(`【${MonthFlag}】分のシートに移動します`);
        await RPA.Logger.info(
          'タイトル、アップロード日、視聴回数の取得を開始します'
        );
        // 前月分からスタートする場合はここをコメントアウト
        // await RPA.Logger.info('作業を終了します');
      }
      if (Number(j) == 1) {
        await RPA.Logger.info('作業を終了します');
      }
    }
  }

  // エラー発生時の処理
  if (ErrorText.length >= 1) {
    // const DOM = await RPA.WebBrowser.driver.getPageSource();
    // await RPA.Logger.info(DOM);
    await RPA.SystemLogger.error(ErrorText);
    Slack_Text[0] = `【Youtube 集計】でエラーが発生しました！\n${ErrorText}`;
    await RPA.WebBrowser.takeScreenshot();
  }
  await SlackPost(Slack_Text[0]);
  await RPA.WebBrowser.quit();
  await RPA.sleep(1000);
  await process.exit();
}

Start();

let FirstLoginFlag = 'true';
async function Work(
  // テスト①
  // ProgramName,
  // TitleList,
  Today,
  NewToday,
  TwoMonthsBefore
  // date,
  // WorkData
) {
  try {
    // Youtube Studioにログイン
    if (FirstLoginFlag == 'true') {
      await YoutubeLogin();
    }
    if (FirstLoginFlag == 'false') {
      await YoutubeLogin2();
    }
    // 一度ログインしたら、以降はログインページをスキップ
    FirstLoginFlag = 'false';
    // 番組名を入力
    await SetProgram(/*ProgramName*/);
    // 番組名からタイトルを検索
    await GetData(/*TitleList, */ NewToday, TwoMonthsBefore);
    // シートに取得したデータを記載
    await SetData(Today, /*date, WorkData, ProgramName, Row, */ NewToday);
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

async function YoutubeLogin2() {
  while (0 == 0) {
    await RPA.sleep(5000);
    const Filter = await RPA.WebBrowser.findElementsByClassName(
      'style-scope ytcp-table-footer'
    );
    if (Filter.length >= 1) {
      await RPA.Logger.info('＊＊＊ログインをスキップしました＊＊＊');
      break;
    }
  }
}

let NumberofViewsList = [];
async function SetProgram(/*ProgramName*/) {
  await RPA.Logger.info(`取得した番組名　    → 【${ProgramName[0]}】`);
  // テスト④
  if (ProgramName[0] == 'スポット') {
    await RPA.Logger.info('番組名が【スポット】です');
    // 5/18 追加修正①
    await Spot(TitleList, NumberofViewsList);
  } else {
    const Icon = await RPA.WebBrowser.wait(
      RPA.WebBrowser.Until.elementLocated({ id: 'text-input' }),
      5000
    );
    await RPA.WebBrowser.mouseClick(Icon);
    await RPA.sleep(2000);

    // 「タイトル」をクリック
    // ヘッドレスモードオン（テスト）用
    // const Title = await RPA.WebBrowser.findElementById('text-item-5');

    // 本番用・ヘッドレスモードオフ（テスト）用
    const Title = await RPA.WebBrowser.findElementById('text-item-0');

    await RPA.WebBrowser.mouseClick(Title);
    await RPA.sleep(1000);

    // 格納した番組名を入力
    const InputTitle: WebElement = await RPA.WebBrowser.driver.executeScript(
      `return document.getElementsByTagName('input')[2]`
    );
    // if (ProgramName[0] == 'GENE高') {
    //   await RPA.WebBrowser.sendKeys(InputTitle, ['ジェネ高']);
    // } else if (ProgramName[0] == 'ラストアイドル') {
    //   await RPA.WebBrowser.sendKeys(InputTitle, ['ラスアイ']);
    // } else {
    await RPA.WebBrowser.sendKeys(InputTitle, [ProgramName[0]]);
    // }
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
}

let LoopFlag = 'true';
let Total = [];
let YoutubeTitleList = [];
let MatchTitleList = [];
let UpdateTitleList = [];
let NumberofViewsList2 = [];
async function GetData(/*TitleList, */ NewToday, TwoMonthsBefore) {
  // タイトルがない場合はスキップ
  try {
    const NoContent = await RPA.WebBrowser.wait(
      RPA.WebBrowser.Until.elementLocated({
        className:
          'no-content-header no-content-header-with-filter style-scope ytcp-video-section-content'
      }),
      10000
    );
    const NoContentText = await NoContent.getText();
    if (NoContentText == '一致する動画がありません。もう一度お試しください。') {
      await RPA.Logger.info('一致する動画がないためデータ取得をスキップします');
    }
  } catch {
    // 一覧が出るまで待機
    await RPA.WebBrowser.wait(
      RPA.WebBrowser.Until.elementLocated({
        className: 'style-scope ytcp-video-list-cell-video video-title-wrapper'
      }),
      10000
    );
    await RPA.Logger.info('タイトルリスト数　　→', TitleList.length, '個');
    if (TitleList.length < 1) {
      await RPA.Logger.info(
        'シートにタイトルの記載がないためデータ取得をスキップします'
      );
    } else {
      await RPA.Logger.info('データ取得を開始します');
      // アップロード日が前々月のタイトルを取得した時点で処理を抜ける
      while (0 == 0) {
        const Range = await RPA.WebBrowser.wait(
          RPA.WebBrowser.Until.elementLocated({
            className: 'page-description style-scope ytcp-table-footer'
          }),
          10000
        );
        const RangeText = await Range.getText();
        var text = await RangeText.split(/[\～\/\s]+/);
        // きちんと遷移できているか確認
        await RPA.Logger.info('現在のページ範囲　  →', RangeText);
        Total[0] = Number(text[3]);

        // 検索にヒットした動画が30件以下の場合
        if (Total[0] < 30) {
          for (let i = 0; i <= Total[0] - 1; i++) {
            // 以下だとなぜかページの上から6番目のタイトルが取得不可
            // const Title: WebElement = await RPA.WebBrowser.driver.executeScript(
            //   `return document.getElementsByClassName('style-scope ytcp-video-list-cell-video video-title-wrapper')[${i}].children[0]`
            // );
            // const TitleText = await Title.getText();
            // タイトルを取得
            const Title: WebElement = await RPA.WebBrowser.driver.executeScript(
              `return document.getElementsByClassName('style-scope ytcp-video-row cell-body tablecell-video floating-column last-floating-column')[${i}].children[0].children[0].children[0].getAttribute('aria-label')`
            );
            await YoutubeTitleList.push(Title);
            // 日付を取得
            const UpdateDate: WebElement = await RPA.WebBrowser.driver.executeScript(
              `return document.getElementsByClassName('style-scope ytcp-video-row cell-body tablecell-date sortable column-sorted')[${i}]`
            );
            const UpdateDateText = await UpdateDate.getText();
            var split = await UpdateDateText.split('\n');
            if (split[0] == NewToday && MonthFlag == '今月') {
              await RPA.Logger.info('公開日が一致したタイトルを取得します');
              const Title2: WebElement = await RPA.WebBrowser.driver.executeScript(
                `return document.getElementsByClassName('style-scope ytcp-video-list-cell-video video-title-wrapper')[${i}].children[0]`
              );
              const TitleText = await Title2.getText();
              // 視聴回数を取得（文字列だと数字の前に ' が付きシートの関数が効かないため Number型に変更する）
              const NumberofView: WebElement = await RPA.WebBrowser.driver.executeScript(
                `return document.getElementsByClassName('style-scope ytcp-video-row cell-body tablecell-views sortable right-align')[${i}].innerText`
              );
              const NumberofViewText = await String(NumberofView);
              const NumberofViews = await NumberofViewText.replace(/,/g, '');
              if (TitleText.indexOf('CMS用') > -1) {
                await RPA.Logger.info(TitleText);
                await RPA.Logger.info(
                  '"CMS用"の文字が含まれているため取得をスキップします'
                );
              } else {
                await RPA.Logger.info(TitleText);
                await UpdateTitleList.push(TitleText);
                await RPA.Logger.info('視聴回数　　        →', NumberofViews);
                // 以下注意！
                await NumberofViewsList2.push(Number(NumberofViews));
              }
            }
          }
        } else {
          for (let i = 0; i <= 29; i++) {
            // 一覧が出るまで待機
            await RPA.WebBrowser.wait(
              RPA.WebBrowser.Until.elementLocated({
                className:
                  'style-scope ytcp-video-list-cell-video video-title-wrapper'
              }),
              15000
            );
            // タイトルを取得
            const Title: WebElement = await RPA.WebBrowser.driver.executeScript(
              `return document.getElementsByClassName('style-scope ytcp-video-row cell-body tablecell-video floating-column last-floating-column')[${i}].children[0].children[0].children[0].getAttribute('aria-label')`
            );
            await YoutubeTitleList.push(Title);
            // 日付を取得
            const UpdateDate: WebElement = await RPA.WebBrowser.driver.executeScript(
              `return document.getElementsByClassName('style-scope ytcp-video-row cell-body tablecell-date sortable column-sorted')[${i}]`
            );
            const UpdateDateText = await UpdateDate.getText();
            var split = await UpdateDateText.split('\n');
            if (split[0] == NewToday && MonthFlag == '今月') {
              await RPA.Logger.info('公開日が一致したタイトルを取得します');
              const Title2: WebElement = await RPA.WebBrowser.driver.executeScript(
                `return document.getElementsByClassName('style-scope ytcp-video-list-cell-video video-title-wrapper')[${i}].children[0]`
              );
              const TitleText = await Title2.getText();
              const NumberofView: WebElement = await RPA.WebBrowser.driver.executeScript(
                `return document.getElementsByClassName('style-scope ytcp-video-row cell-body tablecell-views sortable right-align')[${i}].innerText`
              );
              const NumberofViewText = await String(NumberofView);
              const NumberofViews = await NumberofViewText.replace(/,/g, '');
              if (TitleText.indexOf('CMS用') > -1) {
                await RPA.Logger.info(TitleText);
                await RPA.Logger.info(
                  '"CMS用"の文字が含まれているため取得をスキップします'
                );
              } else {
                await RPA.Logger.info(TitleText);
                await UpdateTitleList.push(TitleText);
                await RPA.Logger.info('視聴回数　　        →', NumberofViews);
                // 以下注意！
                await NumberofViewsList2.push(Number(NumberofViews));
              }
            }
            if (split[0].includes(TwoMonthsBefore) == true) {
              await RPA.Logger.info('前々月のアップロード日があります');
              LoopFlag = 'false';
              break;
            }
          }
        }
        await RPA.Logger.info('シートのリスト　　　→', TitleList);
        await RPA.Logger.info('ツールのリスト　　　→', YoutubeTitleList);
        await RPA.Logger.info('アップロードのタイトル →', UpdateTitleList);
        // 5/11 シート内でタイトルが重複している場合、ツール上では最初にヒットしたものを取得するため注意！（ツール上で区別できる要素があればエラー回避可能）
        // 　　  シートにタイトルの記載があっても、ツール上に動画自体がない場合は取得できないため注意！（途中に「ここまで」の文字を入れておけばエラー回避可能）
        for (let i in TitleList) {
          for (let n in YoutubeTitleList) {
            if (YoutubeTitleList[n].indexOf(TitleList[i]) == 0) {
              await RPA.Logger.info('シートのタイトル　　→', TitleList[i]);
              await RPA.Logger.info(
                'ツールのタイトル　　→',
                YoutubeTitleList[n]
              );
              await RPA.Logger.info(`${Number(n) + 1} 番目　一致しました`);
              const NumberofView: WebElement = await RPA.WebBrowser.driver.executeScript(
                `return document.getElementsByClassName('style-scope ytcp-video-row cell-body tablecell-views sortable right-align')[${n}].innerText`
              );
              const NumberofViewText = await String(NumberofView);
              const NumberofViews = await NumberofViewText.replace(/,/g, '');
              await RPA.Logger.info('視聴回数　　        →', NumberofViews);
              await NumberofViewsList.push(Number(NumberofViews));
              // ツールと一致したタイトルをリストに再格納
              await MatchTitleList.push(YoutubeTitleList[n]);
              break;
            }
          }
        }
        await RPA.Logger.info('一致したタイトル　　→', MatchTitleList);
        if (NumberofViewsList.length != TitleList.length) {
          if (LoopFlag == 'false') {
            if (TitleList.length != MatchTitleList.length) {
              // 視聴回数を取得できなかったタイトルを元のリストから除外
              for (let i = 0; i <= TitleList.length - 1; i++) {
                if (MatchTitleList.includes(TitleList[i]) == false) {
                  await RPA.Logger.info(
                    `タイトルリストの ${i} 番目のタイトルを削除します`
                  );
                  await TitleList.splice(i, 1);
                }
              }
              await RPA.Logger.info('シートのリスト　　　→', TitleList);
              await RPA.Logger.info('データ取得を終了します');
              await RPA.Logger.info('タイトルリストをクリアします');
              YoutubeTitleList = [];
              break;
            } else {
              await RPA.Logger.info('不一致のタイトルはありませんでした');
              await RPA.Logger.info('データ取得を終了します');
              await RPA.Logger.info('タイトルリストをクリアします');
              YoutubeTitleList = [];
              break;
            }
          } else {
            await RPA.Logger.info(
              '取得が完了していないため次のページに進みます'
            );
            await RPA.Logger.info('タイトルリストをクリアします');
            YoutubeTitleList = [];
            const NextPage = await RPA.WebBrowser.findElementById(
              'navigate-after'
            );
            // 5/13 以下だどなぜかJenkinsでは処理が重複する → ジャバスクのクリックで回避
            // await RPA.WebBrowser.mouseClick(NextPage);
            await NextPage.click();
          }
        } else {
          await RPA.Logger.info('データ取得を終了します');
          await RPA.Logger.info('タイトルリストをクリアします');
          YoutubeTitleList = [];
          await RPA.Logger.info(
            '視聴回数リスト数　　→',
            NumberofViewsList.length
          );
          if (MonthFlag == '今月') {
            await RPA.Logger.info(
              'アップロードされたタイトルの視聴回数の数　→',
              NumberofViewsList2.length
            );
          }
          break;
        }
      }
    }
  }
  // 5/11 ページ数を 50 にすると謎のエラーが出る → 30件に固定で回避
  const DeliteIcon = await RPA.WebBrowser.findElementById('delete-icon');
  await RPA.WebBrowser.mouseClick(DeliteIcon);
  await RPA.sleep(1000);
}

// 番組名が【スポット】の場合の処理
async function Spot(TitleList, NumberofViewsList) {
  await RPA.Logger.info('タイトルごとに視聴回数を取得します');
  await RPA.Logger.info('タイトルリスト数　　→', TitleList.length, '個');
  for (let i in TitleList) {
    // 一覧が出るまで待機
    await RPA.WebBrowser.wait(
      RPA.WebBrowser.Until.elementLocated({
        className: 'style-scope ytcp-video-list-cell-video video-title-wrapper'
      }),
      10000
    );
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

    // 格納した番組名を入力
    const InputTitle: WebElement = await RPA.WebBrowser.driver.executeScript(
      `return document.getElementsByTagName('input')[2]`
    );
    await RPA.WebBrowser.sendKeys(InputTitle, [TitleList[i]]);
    await RPA.sleep(1000);
    // 「適用」をクリック
    const Application = await RPA.WebBrowser.findElementById('apply-button');
    await RPA.WebBrowser.mouseClick(Application);
    await RPA.sleep(3000);
    // 結果が出るまで待機
    await RPA.WebBrowser.wait(
      RPA.WebBrowser.Until.elementLocated({
        className: 'style-scope ytcp-video-list-cell-video video-title-wrapper'
      }),
      10000
    );
    const NumberofView: WebElement = await RPA.WebBrowser.driver.executeScript(
      `return document.getElementsByClassName('style-scope ytcp-video-row cell-body tablecell-views sortable right-align')[0].innerText`
    );
    const NumberofViewText = await String(NumberofView);
    const NumberofViews = await NumberofViewText.replace(/,/g, '');
    await RPA.Logger.info('視聴回数　　        →', NumberofViews);
    await NumberofViewsList.push(Number(NumberofViews));
    const DeliteIcon = await RPA.WebBrowser.findElementById('delete-icon');
    await RPA.WebBrowser.mouseClick(DeliteIcon);
    await RPA.sleep(3000);
  }
  await RPA.Logger.info('【スポット】のデータ取得を終了します');
  await RPA.Logger.info('視聴回数リスト数　　→', NumberofViewsList.length);
}

const Column = [];
const ColumnList = [
  'Q',
  'R',
  'S',
  'T',
  'U',
  'V',
  'W',
  'X',
  'Y',
  'Z',
  'AA',
  'AB',
  'AC',
  'AD',
  'AE',
  'AF',
  'AG',
  'AH',
  'AI',
  'AJ',
  'AK',
  'AL',
  'AM',
  'AN',
  'AO',
  'AP',
  'AQ',
  'AR',
  'AS',
  'AT',
  'AU',
  'AV',
  'AW',
  'AX',
  'AY',
  'AZ',
  'BA',
  'BB',
  'BC',
  'BD',
  'BE',
  'BF',
  'BG',
  'BH',
  'BI',
  'BJ',
  'BK',
  'BL',
  'BM',
  'BN',
  'BO',
  'BP',
  'BQ',
  'BR',
  'BS',
  'BT',
  'BU',
  'BV',
  'BW',
  'BX',
  'BY',
  'BZ',
  'CA'
];
const Row4 = [];
const Row5 = [];
async function SetData(Today, /*date, WorkData, ProgramName, Row, */ NewToday) {
  if (TitleList.length < 1) {
    await RPA.Logger.info('視聴回数リストが 0 のため記載をスキップします');
  } else {
    await RPA.Logger.info('現在のシートID　　　→', CurrentSSID[0]);
    await RPA.Logger.info('現在のシート名　　　→', CurrentSSName[0]);
    // 記載する開始列を取得
    for (let i in date[0]) {
      if (Today[0][0] == date[0][i]) {
        Column[0] = ColumnList[i];
        await RPA.Logger.info(`${Column[0]} 列目に記載を開始します`);
        break;
      }
    }
    // テスト④
    await RPA.Logger.info(`${Row3[0]} 行目から記載します`);
    // 記載する開始行を取得
    // for (let i in WorkData[0]) {
    //   if (TitleList[0] == WorkData[0][i][0]) {
    //     Row4[0] = Number(i) + Row[0];
    //     await RPA.Logger.info(`${Row4[0]} 行目から記載します`);
    //     break;
    //   }
    // }
    // if (TitleList.length < 1) {
    //   await RPA.Logger.info('視聴回数リストが 0 のため記載をスキップします');
    // } else {
    if (MonthFlag == '前月') {
      // テスト①
      if (ProgramName[0] == 'アベプラ') {
        // const NewRow = Row4[0] + 7;
        // const NewRow = Row[0] + 7;
        // await RPA.Logger.info(`${Row3[0]} 行目から記載します`);
        // 再度シートのタイトルを取得
        const SheetTitle2 = await RPA.Google.Spreadsheet.getValues({
          spreadsheetId: `${CurrentSSID[0]}`,
          range: `${CurrentSSName[0]}!Q${Row3[0]}:Q${Row3[0] +
            // range: `${CurrentSSName[0]}!Q${Row3[0] + 1}:Q${Row3[0] +
            //   1 +
            TitleList.length -
            1}`
        });
        for (let i = 0; i <= SheetTitle2.length - 1; i++) {
          if (TitleList[i] == SheetTitle2[i][0]) {
            await RPA.Google.Spreadsheet.setValues({
              spreadsheetId: `${CurrentSSID[0]}`,
              range: `${CurrentSSName[0]}!${Column[0]}${Row3[0] + i}:${
                Column[0]
              }${Row3[0] + i}`,
              values: [[NumberofViewsList[i]]]
            });
          }
        }
      } else {
        // テスト①
        // 先に上部のタイトルを取得
        for (let i = 0; i <= LastRow2; i++) {
          const JudgeTitle = await RPA.Google.Spreadsheet.getValues({
            spreadsheetId: `${CurrentSSID[0]}`,
            range: `${CurrentSSName[0]}!Q${Row[0] + i}:Q${Row[0] + i}`
          });
          // テスト④
          // if (JudgeTitle == undefined) {
          //   Row5[0] = Number(i) + Row[0];
          // } else {
          if (JudgeTitle != undefined) {
            Row5[0] = Number(i) + Row[0];
            await RPA.Logger.info(
              `${Row5[0]} 行目 タイトルですのでここから記載を開始します`
            );
            break;
          }
        }
        // 再度番組名ごとにタイトルを取得
        const SheetTitle2 = await RPA.Google.Spreadsheet.getValues({
          spreadsheetId: `${CurrentSSID[0]}`,
          range: `${CurrentSSName[0]}!Q${Row5[0]}:Q${Row5[0] +
            // range: `${CurrentSSName[0]}!Q${Row3[0]}:Q${Row3[0] +
            TitleList.length -
            1}`
        });
        // 一致したタイトルのみ視聴回数を記載
        for (let i = 0; i <= SheetTitle2.length - 1; i++) {
          if (TitleList[i] == SheetTitle2[i][0]) {
            await RPA.Google.Spreadsheet.setValues({
              spreadsheetId: `${CurrentSSID[0]}`,
              range: `${CurrentSSName[0]}!${Column[0]}${Row5[0] + i}:${
                Column[0]
              }${Row5[0] + i}`,
              // range: `${CurrentSSName[0]}!${Column[0]}${Row3[0] + i}:${
              //   Column[0]
              // }${Row3[0] + i}`,
              values: [[NumberofViewsList[i]]]
            });
          }
        }
      }
    } else {
      if (UpdateTitleList.length < 1) {
        await RPA.Logger.info(
          'アップロードされたタイトルが 0 のため記載をスキップします'
        );
      } else {
        await RPA.Logger.info('アップロードされたタイトルを記載します');
        // 上に向かって記載するためデータも合わせて逆順にする
        await UpdateTitleList.reverse();
        await NumberofViewsList2.reverse();
        // 5/12 RPAでは追加で行数を増やすことができないので事前に何行か増やしていただく必要がある
        // アップロードされたタイトルとその視聴回数を記載
        for (let i = 0; i <= UpdateTitleList.length - 1; i++) {
          await RPA.Google.Spreadsheet.setValues({
            spreadsheetId: `${CurrentSSID[0]}`,
            range: `${CurrentSSName[0]}!Q${Row3[0] - 1 - i}:R${Row3[0] -
              1 -
              i}`,
            values: [[UpdateTitleList[i], NewToday]]
          });
          await RPA.Google.Spreadsheet.setValues({
            spreadsheetId: `${CurrentSSID[0]}`,
            range: `${CurrentSSName[0]}!${Column[0]}${Row3[0] - 1 - i}:${
              Column[0]
            }${Row3[0] - 1 - i}`,
            values: [[NumberofViewsList2[i]]]
          });
          await RPA.Logger.info(`${Row3[0] - 1 - i} 行目に記載しました`);
        }
      }
      // 再度番組名ごとにタイトルを取得
      const SheetTitle3 = await RPA.Google.Spreadsheet.getValues({
        spreadsheetId: `${CurrentSSID[0]}`,
        range: `${CurrentSSName[0]}!Q${Row3[0]}:Q${Row3[0] +
          TitleList.length -
          1}`
      });
      // 一致したタイトルのみ視聴回数を記載
      // for (let i = 0; i <= SheetTitle3.length - 1; i++) {
      //   for (let n = 0; n <= TitleList.length - 1; n++) {
      for (let i = 0; i <= TitleList.length - 1; i++) {
        for (let n = 0; n <= SheetTitle3.length - 1; n++) {
          // if (SheetTitle3[i][0] == TitleList[n]) {
          if (TitleList[i] == SheetTitle3[n][0]) {
            await RPA.Logger.info('リストのタイトル　　→', TitleList[i]);
            await RPA.Logger.info('シートのタイトル　　→', SheetTitle3[n][0]);
            await RPA.Logger.info(
              '視聴回数　　　　　　→',
              NumberofViewsList[i]
            );
            await RPA.Google.Spreadsheet.setValues({
              spreadsheetId: `${CurrentSSID[0]}`,
              range: `${CurrentSSName[0]}!${Column[0]}${Row3[0] + n}:${
                Column[0]
              }${Row3[0] + n}`,
              values: [[NumberofViewsList[i]]]
            });
          }
        }
      }
    }
    await RPA.Logger.info(`【${ProgramName[0]}】の記載が完了しました`);
    await RPA.Logger.info('視聴回数リストをクリアします');
    NumberofViewsList = [];
    await RPA.Logger.info('一致したタイトルをクリアします');
    MatchTitleList = [];
    if (MonthFlag == '今月') {
      await RPA.Logger.info(
        'アップロードされたタイトルと視聴回数をクリアします'
      );
      UpdateTitleList = [];
      NumberofViewsList2 = [];
    }
    // 保留
    // if (TitleList2.length > 0) {
    //   await RPA.Logger.info('上部のタイトルをクリアします');
    //   TitleList2 = [];
    // }
  }
}
