/**
 * 佑發 · 高公局保養報表 — 分頁健檢
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
 *           .addItem('📥 匯出當月 PDF', 'exportCurrentMonthToPDF')
 *           .addItem('🔍 分頁健檢', 'checkMonthTabs')                // ← 加這一行
 *           .addToUi();
 *     }
 *
 * 這個工具在檢查什麼？
 *   月份分頁是「複製當下」的模板快照。之後模板改了檢查項目文字，
 *   已經存在的月份分頁不會跟著變，而且重建 Z 欄也救不回來
 *   —— 重建只能重寫既有文字，無法無中生有。
 *   這種分頁上傳時會直接跳「找不到定位鍵」，師傅在現場當場卡住。
 *   本工具就是提早把這種「過期分頁」抓出來。
 *
 * 為什麼拿「模板」當基準，而不是拿網頁版的 MENU_TREE？
 *   因為模板與選單本來就維持同步（由網頁版的 auditMenuVsMaster() 把關），
 *   而月份分頁本來就是模板的複製品。實測比對過，用模板當基準抓到的問題
 *   與用 MENU_TREE 當基準「逐鍵完全相同」。
 *   這樣綁定專案就不必複製一份 33KB 的 MENU_TREE，也沒有兩邊不同步的風險。
 *
 * ★完全唯讀★ 不寫入任何一格、不改任何格式，跑幾次都安全。
 */

// 模板分頁名稱（健檢的比對基準）
const CHK_TEMPLATE_TAB = '模板';

// 每個機房最多列出幾筆明細，避免對話框爆掉
const CHK_SAMPLE_LIMIT = 10;

/**
 * 一鍵健檢所有月份分頁
 * （選單項目對應的函式名稱，onOpen 裡填 'checkMonthTabs'）
 */
function checkMonthTabs() {
  const ui = SpreadsheetApp.getUi();
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // ① 以模板為基準
    const tpl = ss.getSheetByName(CHK_TEMPLATE_TAB);
    if (!tpl) {
      ui.alert('找不到「' + CHK_TEMPLATE_TAB + '」分頁，無法比對。');
      return;
    }
    const base = chkSlots_(tpl);
    const baseTotal = chkCountSlots_(base);

    const out = [];
    out.push('══════ 分頁健檢（唯讀）══════');
    out.push('基準：「' + CHK_TEMPLATE_TAB + '」分頁，共 ' + baseTotal + ' 個照片格');

    // ② 只掃年月分頁；手動改名的備份（例如 202608備份）自動排除
    const months = ss.getSheets().filter(function (sh) {
      return /^\d{6}$/.test(sh.getName());
    });

    if (months.length === 0) {
      out.push('');
      out.push('目前沒有任何月份分頁（形如 202609）。');
      out.push('師傅第一次上傳時，系統會自動從「' + CHK_TEMPLATE_TAB + '」複製一份出來。');
      out.push('');
      out.push('※ 手動改名的備份分頁（例如「202608備份」）不在檢查範圍，');
      out.push('　 任何程式也都不會寫入它，可以放心保留。');
      chkShow_(ui, out.join('\n'));
      return;
    }

    let bad = 0;
    months.forEach(function (sh) {
      const name = sh.getName();
      const slots = chkSlots_(sh);
      const photos = chkCountPhotos_(sh);

      // ③ 雙向比對
      const missing = Object.keys(base).filter(function (k) { return !slots[k]; });
      const orphan  = Object.keys(slots).filter(function (k) { return !base[k];  });

      out.push('');
      out.push('──────────────────────────────');
      out.push('【' + name + '】照片格 ' + chkCountSlots_(slots) + ' 個　已放照片 ' + photos + ' 張');

      if (missing.length === 0 && orphan.length === 0) {
        out.push('  ✓ 與目前的模板完全一致，可正常使用');
        return;
      }

      bad++;
      if (missing.length) {
        out.push('  ✗ 舊模板殘留：' + missing.length + ' 項師傅點了會跳「找不到定位鍵」');
        // 依機房歸類，現場才看得懂是哪一區出問題
        const byRoom = {};
        missing.forEach(function (k) {
          const r = k.split('__')[0];
          byRoom[r] = (byRoom[r] || 0) + 1;
        });
        Object.keys(byRoom).forEach(function (r) {
          out.push('      ' + r + '：' + byRoom[r] + ' 項');
        });
        missing.slice(0, CHK_SAMPLE_LIMIT).forEach(function (k) {
          out.push('      · ' + k.replace('__', ' / '));
        });
        if (missing.length > CHK_SAMPLE_LIMIT) {
          out.push('      · …其餘 ' + (missing.length - CHK_SAMPLE_LIMIT) + ' 項省略');
        }
      }
      if (orphan.length) {
        out.push('  ✗ 這個分頁有、模板已經沒有：' + orphan.length + ' 項（該格永遠空白）');
        orphan.slice(0, CHK_SAMPLE_LIMIT).forEach(function (k) {
          out.push('      · ' + k.replace('__', ' / ') + '（第 ' + slots[k].join(', ') + ' 列）');
        });
        if (orphan.length > CHK_SAMPLE_LIMIT) {
          out.push('      · …其餘 ' + (orphan.length - CHK_SAMPLE_LIMIT) + ' 項省略');
        }
      }
      out.push(photos === 0
        ? '  → 這個分頁還沒有照片，可以直接刪除，下次上傳會自動用新模板重建'
        : '  → ⚠ 已經有 ' + photos + ' 張照片，刪掉會一起不見，請先確認要不要保留');
    });

    out.push('');
    out.push('══════════════════════════════');
    out.push('共檢查 ' + months.length + ' 個月份分頁：' +
             (bad === 0 ? '✓ 全部正常' : '✗ ' + bad + ' 個需要處理'));
    out.push('');
    out.push('※ 手動改名的備份分頁（例如「202608備份」）不在檢查範圍，');
    out.push('　 任何程式也都不會寫入它，可以放心保留。');
    out.push('※ 模板本身與手機選單是否同步，是另一個檢查，');
    out.push('　 在網頁版專案執行 auditMenuVsMaster()。');

    chkShow_(ui, out.join('\n'));

  } catch (err) {
    console.error('[checkMonthTabs] ' + err);
    ui.alert('健檢失敗：' + err.message);
  }
}

/**
 * 用與網頁版 rebuildZColumnFor_ 完全相同的規則，
 * 從 F 欄推導出這個分頁真實的照片格清單。
 * 回傳 { 定位鍵: [照片格起始列, ...] }
 */
function chkSlots_(sh) {
  const lastRow = sh.getLastRow();
  const slots = {};
  if (lastRow < 1) return slots;

  const F = sh.getRange(1, 6, lastRow, 1).getValues().map(function (r) {   // F 欄一次讀入
    return String(r[0] || '').trim();
  });

  let room = '', photoStartRow = -1;
  for (let i = 0; i < F.length; i++) {
    const f = F[i];
    if (f.indexOf('保養檢查項目') === 0) photoStartRow = i + 2;   // 下一列 = 照片格起始列
    if (f.indexOf('機房名稱:') === 0) room = f.replace('機房名稱:', '').trim();
    if (f.indexOf('照片內容說明:') === 0 && room && photoStartRow > 0) {
      let desc = f.replace('照片內容說明:', '').trim();
      // 說明文字折到下一列時接上（排除撞到其他標題列）
      const next = String(F[i + 1] || '').trim();
      if (next && next.indexOf('施工') !== 0 && next.indexOf('完工') !== 0 &&
          next.indexOf('機房') !== 0 && next.indexOf('保養') !== 0 &&
          next.indexOf('照片') !== 0) {
        desc += next;
      }
      const k = room + '__' + desc;
      if (!slots[k]) slots[k] = [];
      slots[k].push(photoStartRow);
    }
  }
  return slots;
}

/** 把 { 鍵: [列,...] } 加總成實際格數（同名多槽位要各算一格） */
function chkCountSlots_(slots) {
  let n = 0;
  Object.keys(slots).forEach(function (k) { n += slots[k].length; });
  return n;
}

/**
 * 數這個分頁已經放了幾張照片 —— 用來判斷能不能安全刪掉重建。
 * 照片可能是內嵌圖（讀出來會是物件）或 =IMAGE() 公式，兩種都要算。
 */
function chkCountPhotos_(sh) {
  const lastRow = sh.getLastRow();
  if (lastRow < 1) return 0;
  const vals = sh.getRange(1, 1, lastRow, 1).getValues();     // A 欄一次讀入
  const fmls = sh.getRange(1, 1, lastRow, 1).getFormulas();
  let n = 0;
  for (let r = 0; r < lastRow; r++) {
    if (typeof vals[r][0] === 'object' && vals[r][0] !== null) n++;
    else if (String(fmls[r][0] || '').indexOf('=IMAGE') === 0) n++;
  }
  return n;
}

/** 把報告顯示成可捲動的對話框 */
function chkShow_(ui, text) {
  const safe = String(text)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  const html = HtmlService.createHtmlOutput(
      '<div style="font:13px/1.6 -apple-system,\'Noto Sans TC\',sans-serif;padding:4px 8px">' +
      '<pre style="white-space:pre-wrap;word-break:break-all;margin:0">' + safe + '</pre>' +
      '</div>')
    .setWidth(680).setHeight(520);
  ui.showModalDialog(html, '分頁健檢結果');
  console.log(text);   // 同時寫進執行紀錄，方便複製貼給別人看
}
