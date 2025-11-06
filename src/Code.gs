/**
 * Pubmed論文自動要約アプリケーション (Gemini版)
 *
 * 機能:
 * - PubmedのAPIを使用し、特定の検索用語で週に１回検索を行う
 * - 過去１週間にpublishされた論文のAbstractをGemini 2.5 Flash APIで要約
 * - 結果をスプレッドシートに記録し、メールで送信
 */

// グローバル変数
const SHEET_NAMES = {
  SETTINGS: 'Settings',
  SEARCH_TERMS: 'SearchTerms',
  RESULTS: 'Results'
};

const SETTING_KEYS = {
  EMAIL: 'email',
  GEMINI_API_KEY: 'geminiApiKey',
  MAX_RESULTS_DEFAULT: 'maxResultsDefault'
};

/**
 * Webアプリとして公開時に呼び出される関数
 */
function doGet() {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('Pubmed論文自動要約アプリ')
    .setFaviconUrl('https://www.ncbi.nlm.nih.gov/favicon.ico');
}

/**
 * 外部ファイル(HTML/CSS/JS)を読み込む
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * スタイルシートのURLを取得
 */
function getStylesheetUrl() {
  return ScriptApp.getService().getUrl() + '?resource=style.css';
}

/**
 * JavaScriptのURLを取得
 */
function getScriptUrl() {
  return ScriptApp.getService().getUrl() + '?resource=script.js';
}

/**
 * スプレッドシートを初期化する
 */
function initializeSpreadsheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 設定シートの作成
  let settingsSheet = ss.getSheetByName(SHEET_NAMES.SETTINGS);
  if (!settingsSheet) {
    settingsSheet = ss.insertSheet(SHEET_NAMES.SETTINGS);
    settingsSheet.appendRow(['Key', 'Value']);
    settingsSheet.appendRow([SETTING_KEYS.EMAIL, '']);
    settingsSheet.appendRow([SETTING_KEYS.GEMINI_API_KEY, '']);
    settingsSheet.appendRow([SETTING_KEYS.MAX_RESULTS_DEFAULT, '10']);
    settingsSheet.getRange('A1:B1').setFontWeight('bold');
  }
  
  // 検索用語シートの作成
  let searchTermsSheet = ss.getSheetByName(SHEET_NAMES.SEARCH_TERMS);
  if (!searchTermsSheet) {
    searchTermsSheet = ss.insertSheet(SHEET_NAMES.SEARCH_TERMS);
    searchTermsSheet.appendRow(['SearchTerm', 'MaxResults']);
    searchTermsSheet.getRange('A1:B1').setFontWeight('bold');
  }
  
  // 結果シートの作成
  let resultsSheet = ss.getSheetByName(SHEET_NAMES.RESULTS);
  if (!resultsSheet) {
    resultsSheet = ss.insertSheet(SHEET_NAMES.RESULTS);
    resultsSheet.appendRow([
      'PubmedID', 
      'Title', 
      'FirstAuthor', 
      'PublicationDate', 
      'SearchTerm', 
      'Abstract', 
      'Summary', 
      'SearchDate'
    ]);
    resultsSheet.getRange('A1:H1').setFontWeight('bold');
  }
}

/**
 * 設定を取得する
 */
function getSettings() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let settingsSheet = ss.getSheetByName(SHEET_NAMES.SETTINGS);
  
  if (!settingsSheet) {
    initializeSpreadsheet();
    settingsSheet = ss.getSheetByName(SHEET_NAMES.SETTINGS);
  }
  
  const data = settingsSheet.getDataRange().getValues();
  const settings = {};
  
  // ヘッダー行をスキップ
  for (let i = 1; i < data.length; i++) {
    settings[data[i][0]] = data[i][1];
  }
  
  return settings;
}

/**
 * 設定を保存する
 */
function saveSettings(settings) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let settingsSheet = ss.getSheetByName(SHEET_NAMES.SETTINGS);
  
  if (!settingsSheet) {
    initializeSpreadsheet();
    settingsSheet = ss.getSheetByName(SHEET_NAMES.SETTINGS);
  }
  
  // キーごとに値を更新
  for (const key in settings) {
    const rowNum = findRowByValue(settingsSheet, key, 0);
    if (rowNum > 0) {
      settingsSheet.getRange(rowNum, 2).setValue(settings[key]);
    } else {
      settingsSheet.appendRow([key, settings[key]]);
    }
  }
  
  return getSettings();
}

/**
 * 検索用語一覧を取得する - バグ修正済み
 */
function getSearchTerms() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let searchTermsSheet = ss.getSheetByName(SHEET_NAMES.SEARCH_TERMS);
  
  if (!searchTermsSheet) {
    initializeSpreadsheet();
    searchTermsSheet = ss.getSheetByName(SHEET_NAMES.SEARCH_TERMS);
    return [];
  }
  
  const data = searchTermsSheet.getDataRange().getValues();
  const terms = [];
  
  // ヘッダー行をスキップ
  for (let i = 1; i < data.length; i++) {
    if (data[i][0]) { // 空の行をスキップ
      terms.push({
        term: data[i][0],
        maxResults: data[i][1] || 10
      });
    }
  }
  
  // デバッグログを追加
  Logger.log("Retrieved search terms: " + JSON.stringify(terms));
  
  return terms;
}

/**
 * 検索用語を保存する
 */
function saveSearchTerm(term, maxResults) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let searchTermsSheet = ss.getSheetByName(SHEET_NAMES.SEARCH_TERMS);
  
  if (!searchTermsSheet) {
    initializeSpreadsheet();
    searchTermsSheet = ss.getSheetByName(SHEET_NAMES.SEARCH_TERMS);
  }
  
  // 既存の検索用語を確認
  const data = searchTermsSheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === term) {
      // 既存の用語の場合は更新
      searchTermsSheet.getRange(i + 1, 2).setValue(maxResults);
      return getSearchTerms();
    }
  }
  
  // 新規追加
  searchTermsSheet.appendRow([term, maxResults]);
  return getSearchTerms();
}

/**
 * 検索用語を更新する
 */
function updateSearchTerm(index, term, maxResults) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let searchTermsSheet = ss.getSheetByName(SHEET_NAMES.SEARCH_TERMS);
  
  if (!searchTermsSheet) {
    initializeSpreadsheet();
    searchTermsSheet = ss.getSheetByName(SHEET_NAMES.SEARCH_TERMS);
    return getSearchTerms();
  }
  
  const rowNum = index + 2; // ヘッダー + 0ベースインデックス
  searchTermsSheet.getRange(rowNum, 1).setValue(term);
  searchTermsSheet.getRange(rowNum, 2).setValue(maxResults);
  
  return getSearchTerms();
}

/**
 * 検索用語を削除する
 */
function deleteSearchTerm(index) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let searchTermsSheet = ss.getSheetByName(SHEET_NAMES.SEARCH_TERMS);
  
  if (!searchTermsSheet) {
    return getSearchTerms();
  }
  
  const rowNum = index + 2; // ヘッダー + 0ベースインデックス
  searchTermsSheet.deleteRow(rowNum);
  
  return getSearchTerms();
}

/**
 * シート内の特定の値を持つ行を検索する
 */
function findRowByValue(sheet, value, columnIndex) {
  const data = sheet.getDataRange().getValues();
  
  for (let i = 0; i < data.length; i++) {
    if (data[i][columnIndex] === value) {
      return i + 1; // 1ベースのインデックスを返す
    }
  }
  
  return -1; // 見つからない場合
}

/**
 * PubmedAPIを使って論文を検索する
 */
function searchPubmed(term, maxResults) {
  const url = "https://eutils.ncbi.nlm.nih.gov/entrez/eutils/esearch.fcgi";
  const params = {
    db: "pubmed",
    term: term,
    retmode: "json",
    retmax: maxResults,
    reldate: 7,  // 過去7日間
    datetype: "pdat"  // 出版日
  };
  
  const options = {
    method: "get",
    muteHttpExceptions: true
  };
  
  const response = UrlFetchApp.fetch(url + "?" + formatParams(params), options);
  
  if (response.getResponseCode() !== 200) {
    Logger.log("Pubmed検索APIエラー: " + response.getContentText());
    return [];
  }
  
  const data = JSON.parse(response.getContentText());
  return data.esearchresult.idlist || [];
}

/**
 * パラメータをURLクエリ文字列に変換
 */
function formatParams(params) {
  return Object.keys(params)
    .map(key => key + "=" + encodeURIComponent(params[key]))
    .join("&");
}

/**
 * PubmedIDリストから論文情報とアブストラクトを取得
 */
function getArticleDetails(idList) {
  if (idList.length === 0) return [];
  
  const url = "https://eutils.ncbi.nlm.nih.gov/entrez/eutils/efetch.fcgi";
  const params = {
    db: "pubmed",
    id: idList.join(","),
    retmode: "xml"
  };
  
  const options = {
    method: "get",
    muteHttpExceptions: true
  };
  
  const response = UrlFetchApp.fetch(url + "?" + formatParams(params), options);
  
  if (response.getResponseCode() !== 200) {
    Logger.log("Pubmed詳細取得APIエラー: " + response.getContentText());
    return [];
  }
  
  const xml = response.getContentText();
  return parseArticleXml(xml, idList);
}

/**
 * XMLから論文情報を解析する
 */
function parseArticleXml(xml, idList) {
  try {
    const document = XmlService.parse(xml);
    const root = document.getRootElement();
    const articles = [];
    
    // PubmedArticleSet内のPubmedArticle要素を取得
    const articleElements = root.getChildren("PubmedArticle");
    
    for (let i = 0; i < articleElements.length; i++) {
      const articleElement = articleElements[i];
      
      try {
        // MedlineCitationを取得
        const medlineCitation = articleElement.getChild("MedlineCitation");
        if (!medlineCitation) continue;
        
        // PMIDを取得
        const pmidElement = medlineCitation.getChild("PMID");
        const pmid = pmidElement ? pmidElement.getText() : "";
        
        // 記事情報を取得
        const article = medlineCitation.getChild("Article");
        if (!article) continue;
        
        // タイトルを取得
        const titleElement = article.getChild("ArticleTitle");
        const title = titleElement ? titleElement.getText() : "";
        
        // 著者リストを取得
        const authorList = article.getChild("AuthorList");
        let firstAuthor = "";
        
        if (authorList) {
          const authors = authorList.getChildren("Author");
          if (authors.length > 0) {
            const author = authors[0];
            const lastName = author.getChild("LastName");
            const foreName = author.getChild("ForeName");
            
            if (lastName && foreName) {
              firstAuthor = lastName.getText() + " " + foreName.getText();
            } else if (lastName) {
              firstAuthor = lastName.getText();
            }
          }
        }
        
        // アブストラクトを取得
        const abstractElement = article.getChild("Abstract");
        let abstract = "";
        
        if (abstractElement) {
          const abstractTexts = abstractElement.getChildren("AbstractText");
          abstract = abstractTexts.map(text => text.getText()).join(" ");
        }
        
        // 出版日を取得
        const pubDate = getPublicationDate(article);
        
        articles.push({
          pmid: pmid,
          title: title,
          firstAuthor: firstAuthor,
          abstract: abstract,
          publicationDate: pubDate
        });
      } catch (e) {
        Logger.log("論文情報の解析エラー: " + e.toString());
      }
    }
    
    return articles;
  } catch (e) {
    Logger.log("XMLパースエラー: " + e.toString());
    return [];
  }
}

/**
 * 論文の出版日を取得
 */
function getPublicationDate(article) {
  try {
    const journalElement = article.getChild("Journal");
    if (!journalElement) return "";

    const journalIssue = journalElement.getChild("JournalIssue");
    if (!journalIssue) return "";

    const pubDateElement = journalIssue.getChild("PubDate");
    if (!pubDateElement) return "";

    const yearElement = pubDateElement.getChild("Year");
    const monthElement = pubDateElement.getChild("Month");
    const dayElement = pubDateElement.getChild("Day");

    const year = yearElement ? yearElement.getText() : "";
    const month = monthElement ? monthElement.getText() : "";
    const day = dayElement ? dayElement.getText() : "";

    if (year && month && day) {
      return year + "-" + month + "-" + day;
    } else if (year && month) {
      return year + "-" + month;
    } else {
      return year;
    }
  } catch (e) {
    Logger.log("出版日パースエラー: " + e.toString());
    return "";
  }
}

/**
 * Gemini 2.5 Flash APIを使用してアブストラクトを要約
 */
function summarizeWithGemini(abstract) {
  if (!abstract || abstract.trim() === "") {
    return "アブストラクトがありません";
  }

  // 設定からAPIキーを取得
  const settings = getSettings();
  const apiKey = settings[SETTING_KEYS.GEMINI_API_KEY];

  if (!apiKey || apiKey.trim() === "") {
    return "Gemini APIキーが設定されていません";
  }

  const url = "https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=" + apiKey;
  const payload = {
    contents: [
      {
        parts: [
          {
            text: "以下の医学論文のアブストラクトを100単語以内で簡潔に日本語で要約してください：\n\n" + abstract
          }
        ]
      }
    ],
    generationConfig: {
      temperature: 0.7,
      maxOutputTokens: 1000
    }
  };

  const options = {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  try {
    const response = UrlFetchApp.fetch(url, options);

    if (response.getResponseCode() !== 200) {
      Logger.log("Gemini API エラー: " + response.getContentText());
      return "要約APIエラー";
    }

    const data = JSON.parse(response.getContentText());

    if (data.candidates && data.candidates.length > 0 &&
        data.candidates[0].content && data.candidates[0].content.parts &&
        data.candidates[0].content.parts.length > 0) {
      return data.candidates[0].content.parts[0].text || "要約できませんでした";
    } else {
      Logger.log("Gemini API レスポンス形式エラー: " + JSON.stringify(data));
      return "要約できませんでした";
    }
  } catch (e) {
    Logger.log("Gemini API例外: " + e.toString());
    return "要約処理エラー";
  }
}

/**
 * 結果をスプレッドシートに記録
 */
function recordToSheet(article, searchTerm, summary) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let resultsSheet = ss.getSheetByName(SHEET_NAMES.RESULTS);
  
  if (!resultsSheet) {
    initializeSpreadsheet();
    resultsSheet = ss.getSheetByName(SHEET_NAMES.RESULTS);
  }
  
  // ヘッダー行を除いた行数を確認
  const lastRow = resultsSheet.getLastRow();
  const dataRows = lastRow - 1; // ヘッダー行を除く
  
  Logger.log(`スプレッドシートの行数: ${lastRow}, データ行数: ${dataRows}`);
  Logger.log(`検証対象論文ID: ${article.pmid}`);
  
  // 既存IDを確認
  if (dataRows > 0) {
    const existingIdsRange = resultsSheet.getRange(2, 1, dataRows, 1);
    const existingIds = existingIdsRange.getValues();
    
    Logger.log(`既存ID数: ${existingIds.length}`);
    
    // 文字列として比較
    const pmidStr = String(article.pmid).trim();
    
    for (let i = 0; i < existingIds.length; i++) {
      const existingPmid = String(existingIds[i][0]).trim();
      Logger.log(`比較: "${existingPmid}" vs "${pmidStr}"`);
      
      if (existingPmid === pmidStr) {
        Logger.log(`論文ID ${pmidStr} は既に記録済みのためスキップします (行 ${i+2})`);
        return false;
      }
    }
  }
  
  // 新規レコードを追加
  resultsSheet.appendRow([
    article.pmid,
    article.title,
    article.firstAuthor,
    article.publicationDate,
    searchTerm,
    article.abstract,
    summary,
    new Date().toISOString()
  ]);
  
  Logger.log(`新規論文 ${article.pmid} を記録しました`);
  return true;
}

/**
 * 検索結果をメールで送信
 */
function sendEmail(results, searchTerm) {
  if (results.length === 0) return;
  
  const settings = getSettings();
  const email = settings[SETTING_KEYS.EMAIL];
  
  if (!email || email.trim() === "") {
    Logger.log("送信先メールアドレスが設定されていません");
    return;
  }
  
  // プレーンテキスト用の本文
  let plainBody = `Pubmed検索「${searchTerm}」の新着論文（${results.length}件）\n\n`;
  
  results.forEach((result, index) => {
    plainBody += `${index + 1}. ${result.title}\n`;
    plainBody += `   著者: ${result.firstAuthor}\n`;
    plainBody += `   PubmedID: ${result.pmid}\n`;
    plainBody += `   要約: ${result.summary}\n\n`;
  });
  
  // HTML形式の本文 - Pubmed風のスタイル
  let htmlBody = `
  <div style="font-family: 'Helvetica Neue', Arial, sans-serif; max-width: 800px; margin: 0 auto;">
    <div style="background-color: #336699; color: white; padding: 15px; border-radius: 5px 5px 0 0;">
      <h2 style="margin: 0; font-size: 20px;">Pubmed検索「${searchTerm}」の新着論文（${results.length}件）</h2>
    </div>
    <div style="background-color: #f8f8f8; padding: 20px; border: 1px solid #ddd; border-top: none; border-radius: 0 0 5px 5px;">
  `;
  
  results.forEach((result, index) => {
    htmlBody += `
      <div style="background-color: white; margin-bottom: 15px; padding: 15px; border-radius: 5px; border-left: 4px solid #336699; box-shadow: 0 1px 3px rgba(0,0,0,0.1);">
        <p style="font-size: 17px; font-weight: bold; color: #336699; margin-top: 0;">${index + 1}. ${result.title}</p>
        <p style="margin: 8px 0; color: #555;"><strong>著者:</strong> ${result.firstAuthor}</p>
        <p style="margin: 8px 0; color: #555;"><strong>PubmedID:</strong> <a href="https://pubmed.ncbi.nlm.nih.gov/${result.pmid}/" style="color: #336699; text-decoration: none;">${result.pmid}</a></p>
        <div style="margin-top: 12px; padding-top: 12px; border-top: 1px solid #eee;">
          <p style="color: #333;"><strong>要約:</strong> ${result.summary}</p>
        </div>
      </div>
    `;
  });
  
  htmlBody += `
    </div>
    <div style="text-align: center; padding: 10px; color: #666; font-size: 12px;">
      <p>この通知メールは Pubmed論文自動要約アプリにより送信されました</p>
    </div>
  </div>
  `;
  
  // メール送信（HTML形式）
  try {
    MailApp.sendEmail({
      to: email,
      subject: `Pubmed新着論文通知 - ${searchTerm} (${results.length}件)`,
      htmlBody: htmlBody,
      body: plainBody  // HTML未対応のメールクライアント用
    });
    Logger.log("メール送信完了: " + email);
  } catch (e) {
    Logger.log("メール送信エラー: " + e.toString());
  }
}

/**
 * 検索実行のメイン関数
 */
function runSearch() {
  const searchTerms = getSearchTerms();
  const settings = getSettings();
  let allResults = [];
  
  // 検索用語がない場合のエラーハンドリング
  if (!searchTerms || searchTerms.length === 0) {
    Logger.log("検索用語が設定されていません");
    return {
      success: false,
      count: 0,
      message: "検索用語が設定されていません。「検索設定」タブで検索用語を追加してください。"
    };
  }
  
  // APIキーのチェック
  if (!settings.geminiApiKey || settings.geminiApiKey.trim() === "") {
    Logger.log("Gemini APIキーが設定されていません");
    return {
      success: false,
      count: 0,
      message: "Gemini APIキーが設定されていません。「API設定」タブでAPIキーを設定してください。"
    };
  }
  
  // メールアドレスのチェック
  if (!settings.email || settings.email.trim() === "") {
    Logger.log("メールアドレスが設定されていません");
    return {
      success: false,
      count: 0,
      message: "メールアドレスが設定されていません。「メール設定」タブでメールアドレスを設定してください。"
    };
  }
  
  // 各検索用語を処理
  for (const searchTerm of searchTerms) {
    try {
      // 検索用語のバリデーション
      if (!searchTerm.term || searchTerm.term.trim() === "") {
        Logger.log("検索用語が空です");
        continue;
      }
      
      // Pubmed検索実行
      const idList = searchPubmed(searchTerm.term, searchTerm.maxResults);
      Logger.log(`「${searchTerm.term}」の検索結果: ${idList.length}件`);
      
      if (idList.length === 0) continue;
      
      // 論文詳細を取得
      const articles = getArticleDetails(idList);
      Logger.log(`詳細取得結果: ${articles.length}件`);
      
      // 各論文を処理
      const termResults = [];
      
      for (const article of articles) {
        // アブストラクト要約
        const summary = summarizeWithGemini(article.abstract);

        // スプレッドシートに記録
        const isNew = recordToSheet(article, searchTerm.term, summary);
        
        if (isNew) {
          article.summary = summary;
          termResults.push(article);
        }
      }
      
      // 新規論文があればメール送信
      if (termResults.length > 0) {
        sendEmail(termResults, searchTerm.term);
        allResults = allResults.concat(termResults);
      }
    } catch (e) {
      Logger.log(`「${searchTerm.term}」の処理中にエラーが発生: ${e.toString()}`);
    }
    
    // API制限を避けるため少し待機
    Utilities.sleep(1000);
  }
  
  return {
    success: true,
    count: allResults.length,
    message: `${allResults.length}件の新規論文を処理しました`
  };
}

/**
 * 週次実行トリガーを設定
 */
function setWeeklyTrigger(dayOfWeek, hour) {
  try {
    // 既存のトリガーをすべて削除
    deleteTriggers();
    
    // 数値に変換（文字列で渡された場合に備えて）
    dayOfWeek = parseInt(dayOfWeek, 10);
    hour = parseInt(hour, 10);
    
    // ScriptApp.WeekDay列挙型に変換
    let weekDay;
    switch(dayOfWeek) {
      case 1: weekDay = ScriptApp.WeekDay.MONDAY; break;
      case 2: weekDay = ScriptApp.WeekDay.TUESDAY; break;
      case 3: weekDay = ScriptApp.WeekDay.WEDNESDAY; break;
      case 4: weekDay = ScriptApp.WeekDay.THURSDAY; break;
      case 5: weekDay = ScriptApp.WeekDay.FRIDAY; break;
      case 6: weekDay = ScriptApp.WeekDay.SATURDAY; break;
      case 7: weekDay = ScriptApp.WeekDay.SUNDAY; break;
      default: weekDay = ScriptApp.WeekDay.MONDAY; // デフォルト
    }
    
    // 新しいトリガーを設定
    ScriptApp.newTrigger('runSearch')
      .timeBased()
      .onWeekDay(weekDay)
      .atHour(hour)
      .create();
      
    return "トリガーを設定しました: 毎週" + getDayName(dayOfWeek) + " " + hour + "時";
  } catch (e) {
    Logger.log("トリガー設定エラー: " + e.toString());
    return "エラー: " + e.toString();
  }
}

/**
 * トリガーをすべて削除
 */
function deleteTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  for (const trigger of triggers) {
    if (trigger.getHandlerFunction() === 'runSearch') {
      ScriptApp.deleteTrigger(trigger);
    }
  }
}

/**
 * 曜日の数値から名前を取得
 */
function getDayName(dayOfWeek) {
  const days = ['', '月曜日', '火曜日', '水曜日', '木曜日', '金曜日', '土曜日', '日曜日'];
  return days[dayOfWeek] || '';
}

/**
 * 現在のトリガー設定を取得
 */
function getCurrentTrigger() {
  try {
    const triggers = ScriptApp.getProjectTriggers();
    for (const trigger of triggers) {
      if (trigger.getHandlerFunction() === 'runSearch') {
        const triggerSource = trigger.getTriggerSource();
        if (triggerSource === ScriptApp.TriggerSource.CLOCK) {
          // WeekDay列挙型から数値に変換
          let dayOfWeek;
          switch(trigger.getWeekDay()) {
            case ScriptApp.WeekDay.MONDAY: dayOfWeek = 1; break;
            case ScriptApp.WeekDay.TUESDAY: dayOfWeek = 2; break;
            case ScriptApp.WeekDay.WEDNESDAY: dayOfWeek = 3; break;
            case ScriptApp.WeekDay.THURSDAY: dayOfWeek = 4; break;
            case ScriptApp.WeekDay.FRIDAY: dayOfWeek = 5; break;
            case ScriptApp.WeekDay.SATURDAY: dayOfWeek = 6; break;
            case ScriptApp.WeekDay.SUNDAY: dayOfWeek = 7; break;
            default: dayOfWeek = 1;
          }
          
          return {
            exists: true,
            dayOfWeek: dayOfWeek,
            hour: trigger.getHour()
          };
        }
      }
    }
    
    return { exists: false };
  } catch (e) {
    Logger.log("トリガー取得エラー: " + e.toString());
    return { exists: false, error: e.toString() };
  }
}

/**
 * Gemini APIを使ってPubMed検索式を生成
 */
function generatePubmedQuery(userRequest) {
  if (!userRequest || userRequest.trim() === "") {
    return {
      success: false,
      message: "検索内容が入力されていません"
    };
  }

  // 設定からAPIキーを取得
  const settings = getSettings();
  const apiKey = settings[SETTING_KEYS.GEMINI_API_KEY];

  if (!apiKey || apiKey.trim() === "") {
    return {
      success: false,
      message: "Gemini APIキーが設定されていません"
    };
  }

  const url = "https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=" + apiKey;

  const prompt = `あなたはPubMed検索のエキスパートです。以下のユーザーの要望に基づいて、最適なPubMed検索式を生成してください。

【重要な指示】
1. MeSH用語（Medical Subject Headings）を積極的に使用してください
2. 必要に応じてフリーワードも組み合わせてください
3. AND、OR、NOTなどのブール演算子を適切に使用してください
4. [MeSH Terms]、[Title/Abstract]、[Author]などのフィールドタグを適切に使用してください
5. 検索式のみを出力してください（説明文は不要です）
6. 検索式は1行で出力してください

ユーザーの要望:
${userRequest}

PubMed検索式:`;

  const payload = {
    contents: [
      {
        parts: [
          {
            text: prompt
          }
        ]
      }
    ],
    generationConfig: {
      temperature: 0.3,
      maxOutputTokens: 500
    }
  };

  const options = {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  try {
    const response = UrlFetchApp.fetch(url, options);

    if (response.getResponseCode() !== 200) {
      Logger.log("Gemini API エラー: " + response.getContentText());
      return {
        success: false,
        message: "Gemini APIエラー"
      };
    }

    const data = JSON.parse(response.getContentText());

    if (data.candidates && data.candidates.length > 0 &&
        data.candidates[0].content && data.candidates[0].content.parts &&
        data.candidates[0].content.parts.length > 0) {
      const generatedQuery = data.candidates[0].content.parts[0].text.trim();
      return {
        success: true,
        query: generatedQuery
      };
    } else {
      Logger.log("Gemini API レスポンス形式エラー: " + JSON.stringify(data));
      return {
        success: false,
        message: "検索式を生成できませんでした"
      };
    }
  } catch (e) {
    Logger.log("Gemini API例外: " + e.toString());
    return {
      success: false,
      message: "検索式生成エラー: " + e.toString()
    };
  }
}