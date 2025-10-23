/**
* プロジェクト①：入力アプリ用GAS - サーバーサイドコード (Ver.2.0)
* 役割：index.htmlからの要求に応じて、データの取得や保存を行う。
* 戦略：データ永続化をAppSheet API に一本化し、DriveApp / SpreadsheetApp を廃止する。
*/

/**
* AppSheet API との通信をカプセル化するサービスモジュール。
* @namespace
*/
const AppSheetApiService = (function() {
const SCRIPT_PROPS = PropertiesService.getScriptProperties();
const APP_ID = SCRIPT_PROPS.getProperty('APPSHEET_APP_ID');
const ACCESS_KEY = SCRIPT_PROPS.getProperty('APPSHEET_ACCESS_KEY');
const API_BASE_URL = `https://api.appsheet.com/api/v2/apps/${APP_ID}/tables/`;

if (!APP_ID || !ACCESS_KEY) {
throw new Error('スクリプトプロパティ「APPSHEET_APP_ID」または「APPSHEET_ACCESS_KEY」が設定されていません。');
}

/**
* AppSheet API にリクエストを送信するプライベートなコア関数。
* @private
* @param {string} tableName - 対象のテーブル名
* @param {Object} payload - API リクエストのボディ
* @returns {Object|Array} API からのレスポンスボディをパースした JSON オブジェクト
* @throws {Error} API リクエストが失敗した場合
*/
function _request(tableName, payload) {
const url = `${API_BASE_URL}${tableName}/Action`;
const options = {
method: 'post',
contentType: 'application/json',
headers: {
'applicationAccessKey': ACCESS_KEY
},
payload: JSON.stringify(payload),
muteHttpExceptions: true // HTTP エラー時もレスポンスを取得するため
};

   const response = UrlFetchApp.fetch(url, options);
   const responseCode = response.getResponseCode();
   const responseBody = response.getContentText();

   if (responseCode === 200) {
     return JSON.parse(responseBody);
   } else {
     Logger.log(`AppSheet API Error: Status ${responseCode}, Body: ${responseBody}`);
     throw new Error(`AppSheet API request failed for table ${tableName} with status ${responseCode}.`);
   }
}

return {
/**
* UrlFetchApp.fetchAll で使用するためのリクエストオブジェクトを生成する。
* @param {string} tableName - 対象のテーブル名
* @param {Object} payload - API リクエストのボディ
* @returns {Object} UrlFetchApp.fetch に渡すためのリクエストオブジェクト
*/
createFetchRequest: function(tableName, payload) {
const url = `${API_BASE_URL}${tableName}/Action`;
return {
url: url,
method: 'post',
contentType: 'application/json',
headers: { 'applicationAccessKey': ACCESS_KEY },
payload: JSON.stringify(payload),
muteHttpExceptions: true
};
},

   /**
    * 条件に一致するレコードを検索する (Find Action)。
    * @param {string} tableName - 検索対象のテーブル名
    * @param {string} selector - AppSheet の Selector 式 (例: "Filter(TableName,...)")
    * @returns {Array<Object>} 検索結果のレコード配列
    */
   findRecords: function(tableName, selector) {
     const payload = {
       "Action": "Find",
       "Properties": {},
       "Rows": [],
       "Selector": selector
     };
     return _request(tableName, payload);
   },

   /**
    * テーブルに新しいレコードを追加する (Add Action)。
    * @param {string} tableName - 追加対象のテーブル名
    * @param {Array<Object>} rowsArray - 追加するレコードのオブジェクト配列
    * @returns {Array<Object>} 追加されたレコード
    */
   addRecords: function(tableName, rowsArray) {
     const payload = {
       "Action": "Add",
       "Properties": { "Locale": "ja-JP" },
       "Rows": rowsArray
     };
     return _request(tableName, payload);
   },

   /**
    * 既存のレコードを削除する (Delete Action)。
    * @param {string} tableName - 削除対象のテーブル名
    * @param {Array<Object>} rowsArray - 削除するレコードの主キーを含むオブジェクト配列
    * @returns {Object} API からのレスポンス
    */
   deleteRecords: function(tableName, rowsArray) {
     const payload = {
       "Action": "Delete",
       "Properties": {},
       "Rows": rowsArray
     };
     return _request(tableName, payload);
   },

   /**
    * 既存のレコードを更新する (Edit Action)。
    * @param {string} tableName - 更新対象のテーブル名
    * @param {Array<Object>} rowsArray - 更新するレコードのオブジェクト配列（主キーを含む）
    * @returns {Object} API からのレスponse
    */
   editRecords: function(tableName, rowsArray) {
     const payload = {
       "Action": "Edit",
       "Properties": { "Locale": "ja-JP" },
       "Rows": rowsArray
     };
     return _request(tableName, payload);
   }
};
})();

/**
* Web アプリケーションの GET リクエストを処理し、メインの HTML ページを返す。
* @param {Object} e - イベントオブジェクト
* @returns {HtmlOutput} レンダリングされる HTML オブジェクト
*/
function doGet(e) {
  // createTemplateFromFile を使用してテンプレートとして読み込み、evaluate() で評価する
  const template = HtmlService.createTemplateFromFile('index.html');
  return template.evaluate()
    .setTitle('店舗データ管理アプリケーション') // アプリケーションのタイトルを適切に設定
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * HTML テンプレートに他のHTMLファイルをインクルードするためのヘルパー関数。
 * @param {string} filename - インクルードするHTMLファイル名（.html拡張子は不要）
 * @returns {string} ファイルの内容
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}


/**
* AppSheet の'Stores'テーブルから全店舗のリストを取得し、フロントエンド用に整形して返す。
* @returns {Array} 店舗名のリスト。例: [{ storeName: "店舗 A", companyName: "会社 X" },...]
* @throws {Error} API からのデータ取得に失敗した場合
*/
function getStoresList() {
    try {
        const response = AppSheetApiService.findRecords('Stores', 'Filter(Stores, TRUE)');

        // フロントエンドのフィルター が必要とする全カラムを
        // camelCase にマッピングして返す
        const storesList = response.map(record => ({
            storeName: record.StoreName,
            companyName: record.CompanyName,
            teamName: record.TeamName,       // TeamName を teamName にマッピング
            interviewer: record.Interviewer  // Interviewer を interviewer にマッピング
        }));

        return storesList;
    } catch (error) {
        Logger.log(`getStoresListでエラーが発生しました: ${error.toString()}`);
        throw new Error(`店舗リストの取得に失敗しました。詳細: ${error.message}`);
    }
}

/**
* 指定された店舗名に基づいて、関連する全てのデータを AppSheet から取得する。
* 1. 店舗名から StoreID を解決
* 2. StoreID を使い、関連データを並列取得
* 3. データを単一の JSON オブジェクトに集約して返す
* @param {string} storeName - 取得対象の店舗名
* @returns {Object} 店舗に関連する全データを含むオブジェクト。例: { storeDetails: {...}, outsourcingCosts: [...] }
* @throws {Error} 店舗が見つからない、またはデータ取得に失敗した場合
*/
function getStoreDataByStoreName(storeName) {
try {
// --- ステージ 1: StoreID の解決 ---
const storeFilter = `Filter(Stores, ([StoreName] = "${storeName}"))`;
const stores = AppSheetApiService.findRecords('Stores', storeFilter);

 if (!stores || stores.length === 0) {
throw new Error(`店舗「${storeName}」が見つかりません。`);
}
if (stores.length > 1) {
throw new Error(`店舗名「${storeName}」が重複しています。データを確認してください。`);
}
const storeDetails = stores;
const storeId = storeDetails[0].StoreID; // Assuming StoreID is the primary key in Stores table

   // --- ステージ 2: 関連データの並列取得 ---
   const relatedTables = ['OutsourcingCosts', 'RecruitmentMedia', 'OvertimeSubjects', 'OrganizationCharts'];

   // UrlFetchApp.fetchAll のために、各テーブルへのリクエストオブジェクトの配列を生成
   const requests = relatedTables.map(tableName => {
     // Assuming primary key for related tables is 'ID' and foreign key is 'StoreID'
     const filter = `Filter(${tableName}, ([StoreID] = "${storeId}"))`;
     return AppSheetApiService.createFetchRequest(tableName, {
       "Action": "Find",
       "Properties": {},
       "Rows": [],
       "Selector": filter
     });
   });

   // 全てのリクエストを並列実行
   const responses = UrlFetchApp.fetchAll(requests);

   // --- データ集約 ---
   const aggregatedData = {
     storeDetails: storeDetails[0], // Assuming storeDetails is an array with one element
   };

   responses.forEach((response, index) => {
     const tableName = relatedTables[index];
     if (response.getResponseCode() !== 200) {
       throw new Error(`${tableName}テーブルのデータ取得に失敗しました。Status: ${response.getResponseCode()}`);
     }
     // AppSheet API の Find アクションは常に配列を返すため、そのまま格納する
     aggregatedData[tableName.charAt(0).toLowerCase() + tableName.slice(1)] = JSON.parse(response.getContentText());
   });

   return aggregatedData;

} catch (error) {
Logger.log(`getStoreDataByStoreNameでエラーが発生しました: ${error.toString()}`);
throw new Error(`店舗データの取得に失敗しました。詳細: ${error.message}`);
}
}

/**
* フロントエンドから受け取ったデータオブジェクトを AppSheet に保存する。
* 「DELETE then INSERT」パターンによる擬似トランザクションを実行する。
* @param {Object} dataObject - 保存するデータ。例: { storeName: "...", storeDetails: {...}, outsourcingCosts: [...] }
* @returns {Object} 成功時は { success: true }、失敗時は { success: false, error: "..." }
*/
function saveStoreData(dataObject) {
try {
const { storeName, storeDetails, outsourcingCosts, recruitmentMedia, overtimeSubjects, organizationCharts } = dataObject;

 // --- StoreID の解決 ---
const storeFilter = `Filter(Stores, ([StoreName] = "${storeName}"))`;
const stores = AppSheetApiService.findRecords('Stores', storeFilter);
if (!stores || stores.length === 0) {
throw new Error(`保存対象の店舗「${storeName}」が見つかりません。`);
}
const storeId = stores[0].StoreID; // Assuming StoreID is the primary key in Stores table

   // 削除と挿入の対象となる子テーブルを定義
   const childDataMap = {
     'OutsourcingCosts': outsourcingCosts,
     'RecruitmentMedia': recruitmentMedia,
     'OvertimeSubjects': overtimeSubjects,
     'OrganizationCharts': organizationCharts
   };

   // --- DELETE フェーズ ---
   for (const tableName in childDataMap) {
     const filter = `Filter(${tableName}, ([StoreID] = "${storeId}"))`; // Filter by foreign key StoreID
     const existingRecords = AppSheetApiService.findRecords(tableName, filter);

     if (existingRecords && existingRecords.length > 0) {
       // AppSheet の Delete アクションは主キーの配列を要求する
       const primaryKeyColumn = 'ID'; // 各テーブルの主キーカラム名に合わせて変更
       const keysToDelete = existingRecords.map(record => ({ [primaryKeyColumn]: record[primaryKeyColumn] }));
       AppSheetApiService.deleteRecords(tableName, keysToDelete);
     }
   }

   // --- INSERT フェーズ ---
   // 1. 親テーブル(Stores)を更新
   // editRecords は主キーを含むレコードの配列を期待する
   if (storeDetails) { // Ensure storeDetails exists
     storeDetails.StoreID = storeId; // 主キーを明示
     AppSheetApiService.editRecords('Stores', [storeDetails]);
   }


   // 2. 子テーブルに新しいデータを挿入
   for (const tableName in childDataMap) {
     const newRows = childDataMap[tableName];
     if (newRows && newRows.length > 0) {
       // 全ての子レコードに外部キーである StoreID を付与
       const rowsWithFk = newRows.map(row => ({...row, StoreID: storeId }));
       AppSheetApiService.addRecords(tableName, rowsWithFk);
     }
   }

   return { success: true, message: 'データの保存が完了しました。' };

} catch (error) {
Logger.log(`saveStoreDataでエラーが発生しました: ${error.toString()}`);
// フロントエンドにはユーザーフレンドリーなメッセージを返す
return { success: false, error: `データの保存に失敗しました。管理者にお問い合わせください。詳細: ${error.message}` };
}
}


/**
* スプレッドシートを開いた時にメニューを追加する (必要に応じて項目を修正)
*/
function onOpen() {
 // AppSheet への移行に伴い、メニュー項目は不要になる可能性があります。
 // 必要に応じて、新しい機能へのメニュー項目を追加してください。
 // SpreadsheetApp.getUi().createMenu('便利ツール').addToUi();
}
