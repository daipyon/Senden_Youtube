import RPA from 'ts-rpa';
import { WebElement } from 'selenium-webdriver';

// SlackのトークンとチャンネルID
const Slack_Token = process.env.AbemaTV_RPAError_Token;
const Slack_Channel = process.env.AbemaTV_RPAError_Channel;
const Slack_Text = [`【Youtube 集計】集計完了しました`];

// RPAトリガーシートのID
const mySSID = process.env.My_SheetID2;
const mySSID2 = [];
// const SSID = process.env.Senden_Youtube_SheetID3;
// const SSID2 = [];
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
const Row3 = [];
const Row4 = [];

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
      spreadsheetId: `${mySSID}`,
      range: `RPAトリガーシート!B10:B10`
    });
    const ThisMonthSheetID = await RPA.Google.Spreadsheet.getValues({
      spreadsheetId: `${mySSID}`,
      range: `RPAトリガーシート!B13:B13`
    });
    // 今月分からスタートする場合は上をコメントアウト
    mySSID2.push(LastMonthSheetID[0][0]);
    mySSID2.push(ThisMonthSheetID[0][0]);
    await RPA.Logger.info(mySSID2);
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
    for (let j in mySSID2) {
      for (let n in SSName4) {
        CurrentSSID[0] = mySSID2[j];
        CurrentSSName[0] = SSName4[n];
        const Program = await RPA.Google.Spreadsheet.getValues({
          spreadsheetId: `${mySSID2[j]}`,
          range: `${SSName4[n]}!B${StartRow}:B${LastRow}`
        });
        for (let i in Program) {
          if (Program[i][0] != undefined) {
            ProgramList.push(Program[i][0]);
          } else {
            break;
          }
        }
        // M列(本日の日付)を取得
        const Today = await RPA.Google.Spreadsheet.getValues({
          spreadsheetId: `${mySSID2[j]}`,
          range: `${SSName4[n]}!M1:M1`
        });
        // 本日の日付のフォーマットを変更
        var today = await new Date();
        var year = await today.getFullYear();
        var date2 = await new Date(Today[0][0] + ' 09:00:00');
        var month = (await date2.getMonth()) + 1;
        var day = await date2.getDate();
        var month2 = await ('00' + month).slice(-2);
        var day2 = await ('00' + day).slice(-2);
        const NewToday = `${year}` + '/' + `${month2}` + '/' + `${day2}`;
        await RPA.Logger.info('本日の日付　　　　　→ ' + NewToday);
        if (MonthFlag == '前月') {
          date = await RPA.Google.Spreadsheet.getValues({
            spreadsheetId: `${mySSID2[j]}`,
            range: `${SSName4[n]}!Q3:AV${LastRow2}`
          });
        }
        if (MonthFlag == '今月') {
          date = await RPA.Google.Spreadsheet.getValues({
            spreadsheetId: `${mySSID2[j]}`,
            range: `${SSName4[n]}!Q3:CA${LastRow2}`
          });
        }
        const ProgramName2 = await RPA.Google.Spreadsheet.getValues({
          spreadsheetId: `${mySSID2[j]}`,
          range: `${SSName4[n]}!L${StartRow2}:L${LastRow2}`
        });
        // 3月、4月 → GENE高・ラストアイドル・MUSIC BOMBのタイトル不一致のため揃えた（入力するときはジェネ高、ラスアイ）
        for (let v in ProgramList) {
          for (let k in ProgramName2) {
            if (ProgramName2[k][0] == ProgramList[v]) {
              // 格納したタイトルリストを空にする
              TitleList = [];
              Row[0] = Number(k) + StartRow;
              await RPA.Logger.info(`${Row[0]} 行目から取得を開始します`);
              if (MonthFlag == '前月') {
                WorkData[0] = await RPA.Google.Spreadsheet.getValues({
                  spreadsheetId: `${mySSID2[j]}`,
                  range: `${SSName4[n]}!Q${Row[0]}:Q${LastRow2}`
                });
                for (let i in WorkData[0]) {
                  if (WorkData[0][i][0] != undefined) {
                    TitleList.push(WorkData[0][i][0]);
                  }
                  if (
                    WorkData[0][i][0] == 'タイトル' ||
                    WorkData[0][i][0] == 'ここまで'
                  ) {
                    if (WorkData[0][i][0] == 'タイトル') {
                      break;
                    }
                    if (WorkData[0][i][0] == 'ここまで') {
                      break;
                    }
                  }
                }
              } else {
                // 先に上部のタイトルを取得
                for (let i = 0; i <= LastRow2; i++) {
                  const JudgeTitle = await RPA.Google.Spreadsheet.getValues({
                    spreadsheetId: `${mySSID2[j]}`,
                    range: `${SSName4[n]}!Q${Row[0] + i}:Q${Row[0] + i}`
                  });
                  if (JudgeTitle != undefined) {
                    TitleList2.push(JudgeTitle[0][0]);
                  } else {
                    Row3[0] = Number(i) + Row[0];
                    break;
                  }
                }
                if (TitleList2.length < 0) {
                  await RPA.Logger.info('上部のタイトル　　  →', TitleList2);
                }
                await RPA.Logger.info(`${Row3[0]} 行目に空白があります`);
                // 空白行を確認
                for (let i = 0; i <= LastRow2; i++) {
                  const NoTitle = await RPA.Google.Spreadsheet.getValues({
                    spreadsheetId: `${mySSID2[j]}`,
                    range: `${SSName4[n]}!Q${Row3[0] + i}:Q${Row3[0] + i}`
                  });
                  if (NoTitle == undefined) {
                  } else {
                    Row4[0] = Number(i) + Row3[0];
                    break;
                  }
                }
                await RPA.Logger.info(`${Row4[0]} 行目にタイトルがあります`);
                WorkData[0] = await RPA.Google.Spreadsheet.getValues({
                  spreadsheetId: `${mySSID2[j]}`,
                  range: `${SSName4[n]}!Q${Row4[0]}:Q${LastRow2}`
                });
                for (let i in WorkData[0]) {
                  if (WorkData[0][i][0] != undefined) {
                    TitleList.push(WorkData[0][i][0]);
                  }
                  if (
                    WorkData[0][i][0] == 'タイトル' ||
                    WorkData[0][i][0] == 'ここまで'
                  ) {
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
                ProgramName.push(ProgramList[v]);
              } else {
                // 格納した番組名の先頭を削除
                await ProgramName.shift();
                ProgramName.push(ProgramList[v]);
              }
              // タイトルリストの末尾を削除
              TitleList.pop();
              await Work(
                ProgramName,
                TitleList,
                Today,
                NewToday,
                date,
                WorkData
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
async function Work(ProgramName, TitleList, Today, NewToday, date, WorkData) {
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
    await SetProgram(ProgramName);
    // 番組名からタイトルを検索
    await GetData(TitleList, NewToday);
    // シートに取得したデータを記載
    await SetData(Today, date, WorkData, ProgramName, Row, NewToday);
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
    await RPA.sleep(2000);
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
    await RPA.sleep(2000);
    const Filter = await RPA.WebBrowser.findElementsByClassName(
      'style-scope ytcp-table-footer'
    );
    if (Filter.length >= 1) {
      await RPA.Logger.info('＊＊＊ログインをスキップしました＊＊＊');
      break;
    }
  }
}

async function SetProgram(ProgramName) {
  await RPA.Logger.info(`取得した番組名　    →【${ProgramName[0]}】`);
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
  if (ProgramName[0] == 'GENE高') {
    await RPA.WebBrowser.sendKeys(InputTitle, ['ジェネ高']);
  } else if (ProgramName[0] == 'ラストアイドル') {
    await RPA.WebBrowser.sendKeys(InputTitle, ['ラスアイ']);
  } else {
    await RPA.WebBrowser.sendKeys(InputTitle, [ProgramName[0]]);
  }
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

let Total = [];
let YoutubeTitleList = [];
let UpdateTitleList = [];
let NumberofViewsList = [];
let NumberofViewsList2 = [];
async function GetData(TitleList, NewToday) {
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
      while (0 == 0) {
        const Range = await RPA.WebBrowser.wait(
          RPA.WebBrowser.Until.elementLocated({
            className: 'page-description style-scope ytcp-table-footer'
          }),
          10000
        );
        const RangeText = await Range.getText();

        // 本番用・ヘッドレスモードオフ（テスト）用
        var text = await RangeText.split(/[\～\/\s]+/);

        await RPA.Logger.info('ヒットした合計数　　→', text[3]);
        Total[0] = Number(text[3]);
        // 検索にヒットした動画が30件以下の場合
        if (Total[0] < 30) {
          for (let i = 0; i <= Total[0] - 1; i++) {
            // これだとなぜか上から6番目のタイトルが取得できない
            // const Title: WebElement = await RPA.WebBrowser.driver.executeScript(
            //   `return document.getElementsByClassName('style-scope ytcp-video-list-cell-video video-title-wrapper')[${i}].children[0]`
            // );
            // const TitleText = await Title.getText();
            // タイトルを取得
            const Title: WebElement = await RPA.WebBrowser.driver.executeScript(
              `return document.getElementsByClassName('style-scope ytcp-video-row cell-body tablecell-video floating-column last-floating-column')[${i}].children[0].children[0].children[0].getAttribute('aria-label')`
            );
            YoutubeTitleList.push(Title);
            // 日付を取得
            const UpdateDate: WebElement = await RPA.WebBrowser.driver.executeScript(
              `return document.getElementsByClassName('style-scope ytcp-video-row cell-body tablecell-date sortable column-sorted')[${i}]`
            );
            const UpdateDateText = await UpdateDate.getText();
            var split = await UpdateDateText.split('\n');
            if (split[0] == NewToday) {
              await RPA.Logger.info('公開日が一致したタイトルを取得します');
              const Title2: WebElement = await RPA.WebBrowser.driver.executeScript(
                `return document.getElementsByClassName('style-scope ytcp-video-list-cell-video video-title-wrapper')[${i}].children[0]`
              );
              const TitleText = await Title2.getText();
              const NumberofViews: WebElement = await RPA.WebBrowser.driver.executeScript(
                `return document.getElementsByClassName('style-scope ytcp-video-row cell-body tablecell-views sortable right-align')[${i}].innerText`
              );
              if (TitleText.indexOf('_CMS用') > -1) {
                await RPA.Logger.info(
                  '"_CMS用"の文字が含まれているため取得をスキップします'
                );
              } else {
                UpdateTitleList.push(TitleText);
                await RPA.Logger.info('視聴回数　　        →', NumberofViews);
                NumberofViewsList2.push(NumberofViews);
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
            const Title: WebElement = await RPA.WebBrowser.driver.executeScript(
              `return document.getElementsByClassName('style-scope ytcp-video-row cell-body tablecell-video floating-column last-floating-column')[${i}].children[0].children[0].children[0].getAttribute('aria-label')`
            );
            YoutubeTitleList.push(Title);
            // 日付を取得
            const UpdateDate: WebElement = await RPA.WebBrowser.driver.executeScript(
              `return document.getElementsByClassName('style-scope ytcp-video-row cell-body tablecell-date sortable column-sorted')[${i}]`
            );
            const UpdateDateText = await UpdateDate.getText();
            var split = await UpdateDateText.split('\n');
            if (split[0] == NewToday) {
              await RPA.Logger.info('公開日が一致したタイトルを取得します');
              const Title2: WebElement = await RPA.WebBrowser.driver.executeScript(
                `return document.getElementsByClassName('style-scope ytcp-video-list-cell-video video-title-wrapper')[${i}].children[0]`
              );
              const TitleText = await Title2.getText();
              const NumberofViews: WebElement = await RPA.WebBrowser.driver.executeScript(
                `return document.getElementsByClassName('style-scope ytcp-video-row cell-body tablecell-views sortable right-align')[${i}].innerText`
              );
              if (TitleText.indexOf('_CMS用') > -1) {
                await RPA.Logger.info(TitleText);
                await RPA.Logger.info(
                  '"_CMS用"の文字が含まれているため取得をスキップします'
                );
              } else {
                await RPA.Logger.info(TitleText);
                UpdateTitleList.push(TitleText);
                await RPA.Logger.info('視聴回数　　        →', NumberofViews);
                NumberofViewsList2.push(NumberofViews);
              }
            }
          }
        }
        await RPA.Logger.info('シートのリスト　　　→', TitleList);
        await RPA.Logger.info('ツールのリスト　　　→', YoutubeTitleList);
        await RPA.Logger.info('アップロードのタイトル →', UpdateTitleList);
        await RPA.Logger.info(
          '視聴回数リスト数　　→',
          NumberofViewsList.length
        );
        // 5/11 そもそもシート内で同じタイトルがある場合、ツール上では最初にヒットしたものを取得するため注意
        // 　　  シートにタイトルの記載があっても、ツール上に動画自体がない場合は取得できないため注意
        for (let i in TitleList) {
          for (let n in YoutubeTitleList) {
            if (YoutubeTitleList[n].indexOf(TitleList[i]) == 0) {
              await RPA.Logger.info('シートのタイトル　　→', TitleList[i]);
              await RPA.Logger.info(
                'ツールのタイトル　　→',
                YoutubeTitleList[n]
              );
              await RPA.Logger.info(Number(n) + 1, '番目　一致しました');
              // 視聴回数を取得
              const NumberofViews: WebElement = await RPA.WebBrowser.driver.executeScript(
                `return document.getElementsByClassName('style-scope ytcp-video-row cell-body tablecell-views sortable right-align')[${n}].innerText`
              );
              await RPA.Logger.info('視聴回数　　        →', NumberofViews);
              NumberofViewsList.push(NumberofViews);
              break;
            }
          }
        }
        if (NumberofViewsList.length != TitleList.length) {
          await RPA.Logger.info('取得が完了していないため次のページに進みます');
          await RPA.Logger.info('タイトルリストをクリアします');
          YoutubeTitleList = [];
          const NextPage = await RPA.WebBrowser.findElementById(
            'navigate-after'
          );
          await RPA.WebBrowser.mouseClick(NextPage);
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
  // 5/11 ページ数を 50 にすると謎のエラーが出る
  const DeliteIcon = await RPA.WebBrowser.findElementById('delete-icon');
  await RPA.WebBrowser.mouseClick(DeliteIcon);
  await RPA.sleep(1000);
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
const Row2 = [];
async function SetData(Today, date, WorkData, ProgramName, Row, NewToday) {
  await RPA.Logger.info('現在のシートID　　　→', CurrentSSID[0]);
  await RPA.Logger.info('現在のシート名　　　→', CurrentSSName[0]);
  // 記載する開始列を取得
  for (let i in date[0]) {
    if (Today[0][0] == date[0][i]) {
      Column[0] = ColumnList[i];
      await RPA.Logger.info(Column[0], '列目に記載を開始します');
      break;
    }
  }
  // 記載する開始行を取得
  for (let i in WorkData[0]) {
    if (TitleList[0] == WorkData[0][i][0]) {
      Row2[0] = Number(i) + Row[0];
      await RPA.Logger.info(Row2[0], '行目から記載します');
      break;
    }
  }
  if (TitleList.length < 1) {
    await RPA.Logger.info('視聴回数リストが 0 のため記載をスキップします');
  } else {
    if (MonthFlag == '前月') {
      // 再度シートのタイトルを取得
      const SheetTitle2 = await RPA.Google.Spreadsheet.getValues({
        spreadsheetId: `${CurrentSSID[0]}`,
        range: `${CurrentSSName[0]}!Q${Row2[0]}:Q${Row2[0] +
          TitleList.length -
          1}`
      });
      for (let i = 0; i <= SheetTitle2.length - 1; i++) {
        if (TitleList[i] == SheetTitle2[i][0]) {
          await RPA.Google.Spreadsheet.setValues({
            spreadsheetId: `${CurrentSSID[0]}`,
            range: `${CurrentSSName[0]}!${Column[0]}${Row2[0] + i}:${
              Column[0]
            }${Row2[0] + i}`,
            values: [[NumberofViewsList[i]]]
          });
        }
      }
    } else {
      if (UpdateTitleList.length < 0) {
        await RPA.Logger.info(
          'アップロードされたタイトルが 0 のため記載をスキップします'
        );
      } else {
        await UpdateTitleList.reverse();
        await NumberofViewsList2.reverse();
        // 5/12 RPAでは追加で行数を増やすことができないので事前に何行か増やしていただく必要がある
        // 　　　各シートごとにタイトルの最後に「ここまで」の文字が必要
        // アップロードされたタイトルを記載
        for (let i = 0; i <= UpdateTitleList.length - 1; i++) {
          await RPA.Google.Spreadsheet.setValues({
            spreadsheetId: `${CurrentSSID[0]}`,
            range: `${CurrentSSName[0]}!Q${Row4[0] - 1 - i}:R${Row4[0] -
              1 -
              i}`,
            values: [[UpdateTitleList[i], NewToday]]
          });
          await RPA.Google.Spreadsheet.setValues({
            spreadsheetId: `${CurrentSSID[0]}`,
            range: `${CurrentSSName[0]}!${Column[0]}${Row4[0] - 1 - i}:${
              Column[0]
            }${Row4[0] - 1 - i}`,
            values: [[NumberofViewsList2[i]]]
          });
          await RPA.Logger.info(Row4[0] - 1 - i, '行目に記載しました');
        }
      }
      // 再度シートのタイトルを取得
      const SheetTitle3 = await RPA.Google.Spreadsheet.getValues({
        spreadsheetId: `${CurrentSSID[0]}`,
        range: `${CurrentSSName[0]}!Q${Row4[0]}:Q${Row4[0] +
          TitleList.length -
          1}`
      });
      for (let i = 0; i <= SheetTitle3.length - 1; i++) {
        if (TitleList[i] == SheetTitle3[i][0]) {
          await RPA.Google.Spreadsheet.setValues({
            spreadsheetId: `${CurrentSSID[0]}`,
            range: `${CurrentSSName[0]}!${Column[0]}${Row4[0] + i}:${
              Column[0]
            }${Row4[0] + i}`,
            values: [[NumberofViewsList[i]]]
          });
        }
      }
    }
    await RPA.Logger.info(`【${ProgramName[0]}】の記載が完了しました`);
    await RPA.Logger.info('視聴回数リストをクリアします');
    NumberofViewsList = [];
    if (MonthFlag == '今月') {
      await RPA.Logger.info('アップロードされたタイトルをクリアします');
    }
    UpdateTitleList = [];
    if (TitleList2.length < 0) {
      await RPA.Logger.info('上部のタイトルをクリアします');
    }
    TitleList2 = [];
  }
}
