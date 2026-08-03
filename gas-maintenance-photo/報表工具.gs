/**
 * 佑發 · 高公局保養報表工具 — PDF 匯出器
 * =================================================================
 * ★★ 這個檔案要放在「試算表的內嵌指令碼」裡，不是網頁 App 的專案！★★
 *   開啟方式：在試算表中點  擴充功能 → Apps Script
 *   原因：onOpen() 這類簡單觸發器只有「內嵌指令碼」會在開檔時自動執行，
 *        獨立的網頁 App 專案不會觸發，選單不會出現。
 *
 * 功能：選單「佑發報表工具 → 📥 匯出當月 PDF」
 *   直式 A4、符合頁寬、無格線、不印分頁名稱，蓋電子印章不跑位。
 */

// ===== 匯出範圍設定 =====
// ★關鍵★ 報表內容只在 A~G 欄；Z 欄是系統用的定位鍵（1992 格），
// 不限制範圍的話，定位鍵會被印進 PDF，而且 fitw 會把 A~Z 壓進一頁寬 →
// 整張報表縮到看不清楚。所以一律鎖定 A~G。
const PDF_FIRST_COL = 'A';
const PDF_LAST_COL  = 'G';

// 模板分頁名稱（匯出時要擋下，避免誤匯出空母版）
const PDF_TEMPLATE_TAB = '模板';

/**
 * 開啟試算表時自動建立自訂選單
 * ⚠ 若你的內嵌指令碼「已經有」另一個 onOpen()，
 *   請不要重複貼這一段，改為把 addPdfMenu_() 那一行加進既有的 onOpen() 裡。
 */
function onOpen() {
  addPdfMenu_();
}

/** 建立選單（獨立成函式，方便與既有的 onOpen 合併） */
function addPdfMenu_() {
  SpreadsheetApp.getUi()
    .createMenu('佑發報表工具')
    .addItem('📥 匯出當月 PDF', 'exportCurrentMonthToPDF')
    .addToUi();
}

/**
 * 一鍵匯出「目前正在觀看的分頁」為直式 A4 PDF
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
               '請先切換至當月分頁（例如 202607）再匯出。',
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
    tpl.title = friendlyTitle_(tabName);

    const html = tpl.evaluate().setWidth(420).setHeight(260);
    ui.showModalDialog(html, '匯出 PDF 報表');

  } catch (err) {
    console.error('[exportPDF] 失敗：' + err + (err && err.stack ? '\n' + err.stack : ''));
    ui.alert('匯出失敗', '錯誤訊息：' + err.message, ui.ButtonSet.OK);
  }
}

/**
 * 分頁名稱轉成易讀標題：202607 → 115年7月
 * 非六位數年月的分頁（自訂名稱）則原樣顯示
 */
function friendlyTitle_(tabName) {
  if (/^\d{6}$/.test(tabName)) {
    const y = parseInt(tabName.substring(0, 4), 10) - 1911;
    const m = parseInt(tabName.substring(4, 6), 10);
    return y + '年' + m + '月';
  }
  return tabName;
}
