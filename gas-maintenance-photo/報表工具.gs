/**
 * 佑發 · 高公局保養報表 — PDF 匯出器
 * =================================================================
 * ★安裝位置★ 試算表的「內嵌指令碼」：
 *   開啟試算表 → 擴充功能 → Apps Script → 新增檔案，貼上本檔內容
 *   （不要貼到「高公局空調報表自動化」那個網頁 App 專案）
 *
 * ★本檔刻意不含 onOpen()★
 *   你的 選單按鈕.gs 已經有 onOpen() 了，兩個同名函式會互相覆蓋。
 *   請改為在既有的 onOpen() 選單鏈中加入一行：
 *
 *     function onOpen() {
 *       var ui = SpreadsheetApp.getUi();
 *       ui.createMenu('高公局報表工具')
 *           .addItem('➕ 建立新月份報表', 'showRolloverPrompt')
 *           .addItem('📥 匯出當月 PDF', 'exportCurrentMonthToPDF')   // ← 加這一行
 *           .addToUi();
 *     }
 *
 * 功能：直式 A4、符合頁寬、無格線、不印分頁名稱，蓋電子印章不跑位。
 */

// ===== 匯出範圍設定 =====
// ★關鍵★ 報表內容只在 A~G 欄；Z 欄是照片上傳系統用的定位鍵（約 1992 格），
// 不限制範圍的話，那些定位字串會被印進 PDF，而且 fitw 會把 A~Z 壓進一頁寬 →
// 整張報表縮到看不清楚。所以一律鎖定 A~G。
const PDF_FIRST_COL = 'A';
const PDF_LAST_COL  = 'G';

// 模板分頁名稱（匯出時要擋下，避免誤匯出空母版）
const PDF_TEMPLATE_TAB = '模板';

/**
 * 一鍵匯出「目前正在觀看的分頁」為直式 A4 PDF
 * （選單項目對應的函式名稱，onOpen 裡填 'exportCurrentMonthToPDF'）
 */
function exportCurrentMonthToPDF() {
  const ui = SpreadsheetApp.getUi();
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getActiveSheet();          // 使用者目前正在看的分頁
    const tabName = sh.getName();

    // ① 防呆：擋下模板分頁（空母版匯出沒有意義，還會誤導行政人員）
    if (tabName === PDF_TEMPLATE_TAB) {
      ui.alert('無法匯出模板',
               '你目前在「' + PDF_TEMPLATE_TAB + '」分頁。\n\n' +
               '請先切換至當月分頁（例如 202608）再匯出。',
               ui.ButtonSet.OK);
      return;
    }

    // ② 防呆：分頁沒有任何內容時不匯出
    const lastRow = sh.getLastRow();
    if (lastRow < 1) {
      ui.alert('這個分頁是空的，沒有可匯出的內容。');
      return;
    }

    // ③ 組裝匯出網址與參數（全部鎖定，行政人員不必自己調整列印設定）
    const gid = sh.getSheetId();
    const range = PDF_FIRST_COL + '1:' + PDF_LAST_COL + lastRow;   // 只匯出 A~G
    const params = {
      exportFormat: 'pdf',
      format:       'pdf',
      size:         'A4',      // A4 紙張
      portrait:     'true',    // 直式
      fitw:         'true',    // 欄位縮放符合頁寬
      gridlines:    'false',   // 不印網格線
      printtitle:   'false',   // 不印試算表檔名
      sheetnames:   'false',   // 不印分頁名稱
      pagenumbers:  'false',   // 不印頁碼（報表本身已有頁首）
      fzr:          'false',   // 不重複凍結列
      gid:          gid,       // 只匯出這一個分頁
      range:        range,     // ★只匯出 A~G，排除 Z 欄定位鍵★
      // 邊界留 0.4 吋：紙張利用率高，同時保留裝訂與蓋章空間
      top_margin:    '0.4',
      bottom_margin: '0.4',
      left_margin:   '0.4',
      right_margin:  '0.4'
    };

    let url = ss.getUrl().replace(/\/edit.*$/, '') + '/export?';
    url += Object.keys(params)
      .map(function (k) { return k + '=' + encodeURIComponent(params[k]); })
      .join('&');

    console.log('[exportPDF] 分頁=' + tabName + '，gid=' + gid + '，範圍=' + range);

    // ④ 產生下載對話框
    //    採「瀏覽器直接下載」而非後端抓檔：
    //    大型報表（7000+ 列含照片）用 UrlFetchApp 會撞到 50MB 限制與執行逾時，
    //    交給瀏覽器下載則沒有大小與時間限制，且使用者已登入 Google 帳號，權限自然通過。
    const tpl = HtmlService.createTemplateFromFile('下載對話框');
    tpl.url = url;
    tpl.tabName = tabName;
    tpl.title = pdfFriendlyTitle_(tabName);

    const html = tpl.evaluate().setWidth(420).setHeight(260);
    ui.showModalDialog(html, '匯出 PDF 報表');

  } catch (err) {
    console.error('[exportPDF] 失敗：' + err + (err && err.stack ? '\n' + err.stack : ''));
    ui.alert('匯出失敗', '錯誤訊息：' + err.message, ui.ButtonSet.OK);
  }
}

/**
 * 分頁名稱轉成易讀標題：202608 → 115年8月
 * 非六位數年月的分頁（自訂名稱）則原樣顯示
 * （函式名加 pdf 前綴，避免與 選單按鈕.gs 既有函式重名）
 */
function pdfFriendlyTitle_(tabName) {
  if (/^\d{6}$/.test(tabName)) {
    const y = parseInt(tabName.substring(0, 4), 10) - 1911;
    const m = parseInt(tabName.substring(4, 6), 10);
    return y + '年' + m + '月';
  }
  return tabName;
}
