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

/**
 * ★診斷用★ 量測「一頁報表」在 A4 上到底塞不塞得下。
 *
 * 為什麼需要這個？
 *   fitw=true 只讓「寬度」符合頁寬，高度完全不管，而且縮放是等比例的
 *   —— 寬度縮多少、高度就跟著縮多少。所以真正決定會不會爆頁的是：
 *       一頁的列高總和 × 縮放比例  vs  A4 可列印高度
 *   只要超出一點點，最後一列（廠商用印簽名欄）就被擠到下一頁，
 *   而且誤差會逐頁累積，越後面跑版越嚴重。
 *   在 Sheets 畫面上看不出來，因為畫面沒有紙張邊界。
 *
 * 完全唯讀：不改任何列高、欄寬、內容或列印設定。
 */
function diagPdfPageFit() {
  const ui = SpreadsheetApp.getUi();
  try {
    const sh = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const tabName = sh.getName();
    const lastRow = sh.getLastRow();
    if (lastRow < 1) { ui.alert('這個分頁是空的。'); return; }

    // ① 找出每一頁的起點（與 normalizeLayout_ 相同：A 欄的頁首大標題）
    const aVals = sh.getRange(1, 1, lastRow, 1).getValues();
    const pages = [];
    aVals.forEach(function (r, i) {
      if (String(r[0] || '').indexOf('保養檢查照片') >= 0) pages.push(i + 1);
    });
    if (pages.length === 0) { ui.alert('找不到頁首標題，無法定位頁面。'); return; }

    // ② A~G 總欄寬 → 算出 fitw 的縮放比例
    let contentW = 0;
    for (let c = 1; c <= 7; c++) contentW += sh.getColumnWidth(c);

    // A4 直式 @96dpi = 794 × 1123 px；邊界 0.4 吋 = 38.4 px
    const A4_W = 794, A4_H = 1123, MARGIN = 0.4 * 96;
    const printW = A4_W - MARGIN * 2;
    const printH = A4_H - MARGIN * 2;
    const scale  = printW / contentW;          // fitw 的等比例縮放

    // ③ 逐頁量列高總和（抽樣前 12 頁，避免 getRowHeight 呼叫過多而逾時）
    const SAMPLE = Math.min(12, pages.length);
    const rows = [];
    for (let p = 0; p < SAMPLE; p++) {
      const top  = pages[p];
      const span = (p + 1 < pages.length) ? (pages[p + 1] - top) : (lastRow - top + 1);
      let h = 0;
      for (let k = 0; k < span; k++) h += sh.getRowHeight(top + k);
      rows.push({ page: p + 1, top: top, span: span, h: h, scaled: h * scale });
    }

    const out = [];
    out.push('══════ A4 落版診斷（唯讀）══════');
    out.push('分頁：' + tabName + '　共 ' + pages.length + ' 頁');
    out.push('');
    out.push('A~G 總欄寬：' + contentW + ' px');
    out.push('A4 可列印區（邊界 0.4 吋）：寬 ' + printW.toFixed(1) + ' × 高 ' + printH.toFixed(1) + ' px');
    out.push('fitw 等比例縮放：' + (scale * 100).toFixed(1) + '%');
    out.push('');
    out.push('→ 一頁的列高總和若超過 ' + (printH / scale).toFixed(0) + ' px，就會爆頁');
    out.push('');
    out.push('頁  起始列  列數  列高總和  縮放後  狀態');

    let over = 0;
    rows.forEach(function (r) {
      const diff = r.scaled - printH;
      const st = (diff > 0)
        ? '✗ 超出 ' + diff.toFixed(0) + ' px'
        : '✓ 剩 ' + (-diff).toFixed(0) + ' px';
      if (diff > 0) over++;
      out.push(String(r.page).padEnd(4) + String(r.top).padEnd(8) + String(r.span).padEnd(6) +
               String(r.h).padEnd(10) + r.scaled.toFixed(0).padEnd(8) + st);
    });

    out.push('');
    out.push('══════════════════════════════');
    if (over === 0) {
      out.push('✓ 抽樣的 ' + SAMPLE + ' 頁都塞得下，跑版可能另有原因');
    } else {
      const worst = Math.max.apply(null, rows.map(function (r) { return r.scaled - printH; }));
      out.push('✗ ' + over + ' / ' + SAMPLE + ' 頁塞不下，最多超出 ' + worst.toFixed(0) + ' px');
      out.push('');
      out.push('可行的調整方向（擇一即可，數字供參考）：');
      const needMargin = (MARGIN * 2 - worst) / 2 / 96;
      if (needMargin > 0.05) {
        out.push('  · 上下邊界改為 ' + needMargin.toFixed(2) + ' 吋（只改匯出參數，不動表格）');
      } else {
        out.push('  · 只縮邊界不夠，還需要減少列高');
      }
      out.push('  · 或每頁列高總共減少 ' + (worst / scale).toFixed(0) + ' px');
    }

    const msg = out.join('\n');
    console.log(msg);
    const safe = msg.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
    ui.showModalDialog(
      HtmlService.createHtmlOutput(
        '<pre style="font:12px/1.5 monospace;white-space:pre-wrap">' + safe + '</pre>')
        .setWidth(640).setHeight(520),
      'A4 落版診斷');

  } catch (err) {
    console.error('[diagPdfPageFit] ' + err);
    ui.alert('診斷失敗：' + err.message);
  }
}
