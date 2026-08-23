/**
 * 高公局保養拍照 — v4.3 自我修復版後端 (Google Apps Script)
 * =================================================================
 * 沿用 v3 的核心設計（完全不變）：
 *   1. Z 欄 TextFinder 動態定位（不用 8000 行 CELL_MAP）
 *   2. CellImage 入格 + =IMAGE() 公式後備
 *   3. 圖片存 Drive 後設為公開可讀
 *
 * v4 修正與優化（不動任何試算表格式）：
 *   1. LockService 防併發：兩位師傅同時上傳、當月分頁又還沒建立時，
 *      v3 會兩邊同時複製模板 → 第二個 setName 直接報錯或產生重複分頁
 *   2. buildZColumn 改批次寫入：v3 逐格 setValue 664 次（分鐘級、易逾時），
 *      v4 一次 setValues 寫回（秒級）
 *   3. 欄寬列高改用 setColumnWidths / setRowHeights 批次版：
 *      每次上傳從 14 次 API 呼叫降為 2 次
 *   4. 參數防呆：dataUrl / room / item 缺漏時回明確錯誤，不再直接 throw
 *   5. 模板舊年月文字改為常數 TEMPLATE_OLD_YM_TEXT（改模板只需改一處）
 *   6. 關鍵步驟加 console.log（執行記錄可直接追蹤卡在哪一步）
 *
 * v4.5（A 方案：拍攝日為準）：
 *   1. 前端從照片 EXIF 讀出拍攝日（p.shotYmd）；施工/完工日期、月份分頁、
 *      Drive 資料夾全部以拍攝日推導，補傳自動歸回正確月份
 *   2. 讀不到 EXIF → 退回上傳日（單次取時間，跨年瞬間不撕裂）
 *   3. 日期掃描加邊界夾限，槽位在試算表最底部也能正確寫入日期
 *
 * v4.3 自我修復（不需要手動執行任何工具函式）：
 *   1. 上傳時偵測到定位鍵重複 → 自動重建該分頁 Z 欄後重新定位
 *   2. 從模板建立新月份分頁時 → 自動重建新分頁 Z 欄（模板髒了也不影響新月份）
 *   3. 同名多槽位（泰控空氣濾清器等）→ 依序填入第一個空格，
 *      「照片」佔位文字視為空格；Drive 備份檔名加 _槽位N 避免互相覆蓋
 *
 * 試算表母版準備（與 v3 相同）：
 *   在每個照片槽位對應的 Z 欄格子，填入定位字串
 *   格式：機房__AHU編號+項目說明
 *   例：汐止機房__AHU-01濾網及蒸發器檢查及清潔
 *   （可用 buildZColumn() 一次寫入所有槽位）
 */

// ===== 基本設定（要改 ID / 模板文字只需改這一區）=====
const MASTER_SHEET_ID  = '10EbiceMS2pYfw0RPPA4ti6fe1Xux1qMX38z9hnLq8xw';
const BACKUP_FOLDER_ID = '1jFDjYbkZofewjb9J-uKx0fYRv6YO6OKf';
const SHEET_TAB        = 'Sheet2';

const TEMPLATE_TAB_NAME    = '模板';     // 模板分頁名稱
const TEMPLATE_OLD_YM_TEXT = '115年5月'; // 模板頁首的預設年月文字（模板改版時改這裡）

// ★★ 照片容器統一規格 ★★
const BOX_COLS   = 4;     // A~D 共 4 欄
const BOX_ROWS   = 10;    // 跨 10 列
const COL_WIDTH  = 101;   // 每欄 101px → 4 × 101 = 404px
const ROW_HEIGHT = 33;    // 每列 33px  → 10 × 33 = 330px
const Z_COL      = 26;    // Z 欄（定位字串所在欄）

function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('佑發 · 高公局保養拍照')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// MENU_TREE 是寫死的常數，直接回傳已是最快，不需要快取或讀表
function getMenu() { return MENU_TREE; }

/**
 * 師傅上傳一張 → Drive 備份 + 試算表寫入 CellImage
 */
function uploadPhoto(p) {
  const t0 = Date.now(); // 計時，方便從執行記錄看每張耗時
  try {
    // 0. 參數防呆：前端傳來的 payload 缺欄位時，回明確錯誤而不是直接炸掉
    if (!p || !p.room || !p.item || !p.dataUrl) {
      return { ok: false, error: '參數不完整（room / item / dataUrl 缺一不可）' };
    }
    const key   = p.room + '__' + p.item;
    const parts = String(p.dataUrl).split(',');
    if (parts.length < 2 || !parts[1]) {
      return { ok: false, error: '照片資料格式錯誤（dataUrl 非 base64）' };
    }

    // ★v4.6 破圖防護★ base64 解碼獨立攔截
    //   舊手機過熱／訊號中斷時，dataUrl 可能只傳一半。
    //   若剛好斷在 4 的倍數，base64 仍能「合法解碼」成半張 JPEG →
    //   會一路存進 Drive 並寫入儲存格，變成報表上的破圖。必須在此擋下。
    let bytes;
    try {
      bytes = Utilities.base64Decode(parts[1]);
    } catch (e) {
      console.warn('[uploadPhoto] base64 解碼失敗：' + e.message);
      return { ok: false, error: '照片資料損毀（解碼失敗），請重新拍攝或重新選擇照片' };
    }

    // JPEG 完整性三重檢查（GAS 的位元組是有號值：0xFF=-1、0xD8=-40、0xD9=-39）
    const badReason = validateJpegBytes_(bytes);
    if (badReason) {
      console.warn('[uploadPhoto] 照片不完整：' + badReason + '（' + bytes.length + ' bytes）');
      return { ok: false, error: '照片傳輸不完整（' + badReason + '），請在訊號較好的地方重新上傳' };
    }

    const blob = Utilities.newBlob(bytes, 'image/jpeg', key + '.jpg');
    console.log('[uploadPhoto] 開始：' + key + '，照片 ' + Math.round(bytes.length / 1024) + ' KB');

    // ★v4.5 A 方案★ 基準日：優先用前端從照片 EXIF 讀出的「拍攝日」(p.shotYmd,
    // 格式 YYYY-MM-DD)；讀不到或格式不合法 → 退回伺服器當下日期。
    // 之後「月份分頁、施工/完工日期」全部由同一個基準日推導，
    // 補傳照片會自動歸回拍攝當月，跨年瞬間也不會出現分頁與日期不一致。
    let baseY, baseM, baseD, dateSource;
    const sm = /^(\d{4})-(\d{2})-(\d{2})$/.exec(String(p.shotYmd || ''));
    if (sm && +sm[1] >= 2020 && +sm[1] <= 2100 &&
        +sm[2] >= 1 && +sm[2] <= 12 && +sm[3] >= 1 && +sm[3] <= 31) {
      baseY = +sm[1]; baseM = +sm[2]; baseD = +sm[3]; dateSource = '拍攝日';
    } else {
      // 伺服器日期只取一次再拆解，年月日永遠來自同一瞬間（跨年不撕裂）
      const nowStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd').split('-');
      baseY = +nowStr[0]; baseM = +nowStr[1]; baseD = +nowStr[2]; dateSource = '上傳日';
    }
    const pad2 = function (n) { return ('0' + n).slice(-2); };
    const rocDate = (baseY - 1911) + '/' + pad2(baseM) + '/' + pad2(baseD);   // 例：115/08/01
    console.log('[uploadPhoto] 基準日：' + rocDate + '（' + dateSource + '）');

    // 1.（v4.2 調整順序）Drive 備份移到「定位完成之後」執行：
    //    多槽位項目需要用槽位編號命名備份檔，避免互相覆蓋

    // 2. 開啟母版試算表
    let ss;
    try {
      ss = SpreadsheetApp.openById(MASTER_SHEET_ID);
    } catch (e) {
      console.error('[uploadPhoto] 開啟試算表失敗：' + e.message);
      return { ok: false, error: '無法開啟試算表:' + e.message };
    }

    // 檢查前端傳來的 tabName；如果空白、或是選到「模板」，
    // ★v4.5★ 就導向「基準日」的年月分頁（拍攝日優先 → 補傳自動歸回正確月份）
    let targetTab = p.tabName;
    if (!targetTab || targetTab === TEMPLATE_TAB_NAME) {
      targetTab = '' + baseY + pad2(baseM);
    }

    // 取得當月分頁；不存在就從模板建立（內含 LockService 防併發，見下方函式）
    const sh = getOrCreateMonthSheet_(ss, targetTab);
    console.log('[uploadPhoto] 目標分頁：' + targetTab);

    // 3. Z 欄 TextFinder 動態定位
    let matches = sh.getRange('Z:Z')
      .createTextFinder(key)
      .matchEntireCell(true)
      .findAll();

    // 3-1. ★v4.3 自我修復★
    //     同一鍵出現多筆時，最常見原因是「模板列位移後殘留的舊鍵」。
    //     偵測到就自動重建本分頁的 Z 欄、再查一次——不需要任何人工步驟。
    //     每分頁 6 小時內只自動重建一次；重建後仍多筆 = 真正的同名多槽位，
    //     交給步驟 4 依序填入。
    if (matches.length > 1) {
      const cache = CacheService.getScriptCache();
      const flag  = 'zfix_' + targetTab;
      if (!cache.get(flag)) {
        const lock = LockService.getScriptLock();
        lock.waitLock(60000);   // v4.6：月初分頁建立含 Z 欄重建較久，等候放寬到 60 秒
        try {
          if (!cache.get(flag)) {
            console.log('[uploadPhoto] 偵測到重複定位鍵，自動重建 Z 欄：' + targetTab);
            rebuildZColumnFor_(sh);
            cache.put(flag, '1', 21600);   // 6 小時內不重複重建
          }
        } finally {
          lock.releaseLock();
        }
        matches = sh.getRange('Z:Z')
          .createTextFinder(key)
          .matchEntireCell(true)
          .findAll();
      }
    }

    if (matches.length === 0) {
      console.warn('[uploadPhoto] Z 欄找不到定位鍵：' + key);
      return { ok: false, error: 'Z 欄找不到定位鍵：' + key };
    }

    // 4. 逐一解析槽位的照片格起始列
    //    （Z 鍵的下一列 + 合併儲存格校正安全網，勿移除；詳見 v4 註解）
    //    多槽位（泰控空氣濾清器等「同名多格」項目）→ 依序填入第一個空格：
    //      有照片 = 儲存格值是 CellImage 物件、或帶有公式（=IMAGE）
    //      空槽位 = 空白「或任何純文字」——模板的「照片」佔位字也算空格
    const targetCol = 1;   // A 欄
    let targetRow = -1, slotNo = 0, slotTotal = 0;
    const seenRows = {};
    for (let m = 0; m < matches.length; m++) {
      let row = matches[m].getRow() + 1;
      try {
        const merged = sh.getRange(row, targetCol).getMergedRanges();
        if (merged.length > 0) row = merged[0].getRow();
      } catch (e) {
        console.warn('[uploadPhoto] 合併儲存格校正略過：' + e.message);
      }
      if (seenRows[row]) continue;   // 校正後指向同一格的殘鍵只算一個槽位（雙保險）
      seenRows[row] = true;
      slotTotal++;
      if (targetRow > 0) continue;   // 槽位已選定，僅繼續統計總數
      if (matches.length === 1) { targetRow = row; slotNo = 1; continue; }
      const cell = sh.getRange(row, targetCol);
      const v = cell.getValue();
      const hasPhoto = (v !== null && typeof v === 'object') || cell.getFormula() !== '';
      if (!hasPhoto) { targetRow = row; slotNo = slotTotal; }
    }
    if (targetRow < 0) {
      console.warn('[uploadPhoto] 槽位全滿：' + key + '（共 ' + slotTotal + ' 格）');
      return { ok: false, error: '「' + key + '」的 ' + slotTotal + ' 個槽位都已有照片；要更換請先到試算表清空要重拍的那一格再上傳' };
    }

    // 4-2. 備份原檔到 Drive（設為公開可讀，CellImage 才讀得到）
    //      多槽位時檔名加「_槽位N」：避免互相覆蓋、造成先前槽位圖片斷鏈
    const suffix = (slotTotal > 1) ? '_槽位' + slotNo : '';
    const fileId = backupAndGetId_(p, blob, suffix);
    console.log('[uploadPhoto] Drive 備份完成 fileId=' + fileId);

    // 5. 強制統一欄寬列高
    //    ★批次版 setColumnWidths / setRowHeights：一次呼叫取代迴圈 14 次呼叫★
    //    ⚠ 範圍嚴格限制在照片這 4 欄 × 10 列，不觸及施工日期、機房名稱等其他列
    sh.setColumnWidths(targetCol, BOX_COLS, COL_WIDTH);
    sh.setRowHeights(targetRow, BOX_ROWS, ROW_HEIGHT);

    // 6. 清除該位置殘留的舊「浮動圖片」（相容舊版 insertImage 產生的物件）
    try {
      sh.getImages().forEach(function (im) {
        if (im.getAnchorCell().getRow() === targetRow &&
            im.getAnchorCell().getColumn() === targetCol) {
          im.remove();
        }
      });
    } catch (e) {
      // 清舊圖失敗不影響主流程，記錄即可
      console.warn('[uploadPhoto] 清除舊浮動圖片失敗（不影響上傳）：' + e.message);
    }

    // 7. ★CellImage：圖片真正成為儲存格「內容」，非浮動物件★
    //    前端已將照片裁成 404:330，比例吻合容器，故會 100% 填滿無白邊
    const targetCell = sh.getRange(targetRow, targetCol);
    const imageUrl   = 'https://lh3.googleusercontent.com/d/' + fileId;

    let method = 'CellImage';
    try {
      const cellImage = SpreadsheetApp.newCellImage()
        .setSourceUrl(imageUrl)
        .setAltTextTitle(key)
        .build();
      targetCell.setValue(cellImage);
    } catch (e) {
      // 後備方案：CellImage 不可用時改用 =IMAGE() 公式（模式 1：等比填入儲存格）
      console.warn('[uploadPhoto] CellImage 失敗，改用 IMAGE 公式：' + e.message);
      method = 'IMAGE公式';
      targetCell.setFormula('=IMAGE("' + imageUrl + '",1)');
    }

    // 8. ★v4.4 新增★ 自動填寫施工/完工日期（民國年格式，如 115/07/27）
    //    - 只在照片成功入格後執行；獨立 try...catch，日期失敗不影響照片歸檔
    //    - 只掃描「本槽位」targetRow 起算 BOX_ROWS（10）列的 F 欄，
    //      絕不整頁取代，不會誤傷其他還沒保養的設備日期
    let dateNote = '';
    try {
      // rocDate 已在函式開頭由「基準日」（拍攝日優先）算好，這裡直接使用

      // ★v4.5 邊界夾限★ 槽位落在試算表最底部時，讀取範圍不能超過總列數，
      // 否則 getRange 會拋 out of range（最後一台設備的日期會靜默漏寫）
      const scanRows = Math.min(BOX_ROWS, sh.getMaxRows() - targetRow + 1);
      if (scanRows < 1) throw new Error('槽位超出試算表範圍（targetRow=' + targetRow + '）');

      // 一次 getValues 讀入本槽位範圍的 F 欄（僅 1 次讀取 API）
      const fVals = sh.getRange(targetRow, 6, scanRows, 1).getValues();
      const hit = [];   // 記錄需要改寫的列位移（相對 targetRow）
      for (let r = 0; r < fVals.length; r++) {
        const t = String(fVals[r][0] || '').trim();
        // 比對不強制要冒號：母版第 6446 列有「完工日期115/05/08」漏冒號的 typo，
        // 寫回時一律用標準「施工日期:/完工日期:」格式，順手矯正
        if (t.startsWith('施工日期'))      { fVals[r][0] = '施工日期:' + rocDate; hit.push(r); }
        else if (t.startsWith('完工日期')) { fVals[r][0] = '完工日期:' + rocDate; hit.push(r); }
      }

      if (hit.length > 0) {
        // 施工/完工在模板中相鄰 → 取連續子範圍一次 setValues 寫回（僅 1 次寫入 API）；
        // 萬一格式特殊不相鄰則退回逐格寫。兩種路徑都只碰「日期格」本身，
        // 不整段回寫 F 欄，避免覆蓋機房名稱/照片說明等其他文字
        const first = hit[0], last = hit[hit.length - 1];
        if (last - first + 1 === hit.length) {
          const sub = [];
          for (let r = first; r <= last; r++) sub.push(fVals[r]);
          sh.getRange(targetRow + first, 6, sub.length, 1).setValues(sub);
        } else {
          hit.forEach(function (r) {
            sh.getRange(targetRow + r, 6).setValue(fVals[r][0]);
          });
        }
        dateNote = ' | 日期' + rocDate + '(' + dateSource + ')';
        console.log('[uploadPhoto] 施工/完工日期已更新：' + rocDate +
                    '（' + dateSource + '，' + hit.length + ' 格）');
      } else {
        console.warn('[uploadPhoto] 槽位範圍內未找到施工/完工日期欄（不影響照片歸檔）');
      }
    } catch (e) {
      console.warn('[uploadPhoto] 日期更新失敗（照片已入格，不影響）：' + e.message);
    }

    console.log('[uploadPhoto] 完成：' + key + ' → ' + targetTab + '!A' + targetRow +
                '，耗時 ' + (Date.now() - t0) + ' ms');
    return {
      ok: true,
      name: key + '.jpg',
      where: targetTab + '!A' + targetRow +
             (slotTotal > 1 ? ' | 槽位' + slotNo + '/' + slotTotal : '') +
             ' | 容器404x330 | ' + method + ' | 已入格' + dateNote
    };

  } catch (err) {
    console.error('[uploadPhoto] 未預期錯誤：' + err + (err && err.stack ? '\n' + err.stack : ''));
    // ★v4.6★ 把系統原文錯誤翻成現場看得懂的說法（師傅看不懂 Lock timeout / Timed out）
    const raw = String(err);
    let friendly = raw;
    if (/lock|timed out|timeout/i.test(raw)) {
      friendly = '系統忙碌中（其他人正在上傳），請等 10 秒後再按一次上傳';
    } else if (/quota|limit/i.test(raw)) {
      friendly = '今日用量已達 Google 上限，請明天再試或聯絡管理者';
    } else if (/permission|access|authoriz/i.test(raw)) {
      friendly = '權限不足，請聯絡管理者確認試算表與 Drive 資料夾的存取權';
    }
    return { ok: false, error: friendly };
  }
}

/**
 * 取得當月分頁；不存在就從「模板」複製建立並替換頁首年月
 * ★v4 新增 LockService★
 *   情境：月初第一天，兩位師傅幾乎同時上傳，當月分頁都還沒建立。
 *   v3 會兩邊同時 copyTo → 第二個 setName 因同名而報錯（或產生「模板 (副本)」殘留）。
 *   加 ScriptLock 後：一人建立、另一人等待後直接拿到現成分頁。
 */
function getOrCreateMonthSheet_(ss, targetTab) {
  let sh = ss.getSheetByName(targetTab);
  if (sh) return sh;   // 快路徑：分頁已存在（絕大多數情況），不用進鎖

  const lock = LockService.getScriptLock();
  lock.waitLock(60000);   // v4.6：月初分頁建立含 Z 欄重建較久，等候放寬到 60 秒
  try {
    // 拿到鎖後再查一次：可能等待期間已被另一個請求建立
    sh = ss.getSheetByName(targetTab);
    if (sh) return sh;

    const templateSheet = ss.getSheetByName(TEMPLATE_TAB_NAME);
    if (!templateSheet) {
      throw new Error('找不到「' + TEMPLATE_TAB_NAME + '」分頁，無法自動建立當月分頁！');
    }
    sh = templateSheet.copyTo(ss).setName(targetTab);
    console.log('[getOrCreateMonthSheet_] 已從模板建立分頁：' + targetTab);

    // ✅ 智慧型修改標題：將模板的舊月份（例如 115年5月）替換為新月份
    if (targetTab.length === 6) {
      const yearROC  = parseInt(targetTab.substring(0, 4), 10) - 1911;
      const monthNum = parseInt(targetTab.substring(4, 6), 10);
      sh.createTextFinder(TEMPLATE_OLD_YM_TEXT)
        .replaceAllWith(yearROC + '年' + monthNum + '月');
    }

    // ★v4.3★ 新分頁建立後立即重建 Z 欄：
    //   就算模板 Z 欄殘留舊鍵，每個新月份分頁一出生就是乾淨的
    rebuildZColumnFor_(sh);

    // 確保寫入落盤後才放鎖，避免其他請求拿到「半成品」分頁
    SpreadsheetApp.flush();
    return sh;
  } finally {
    lock.releaseLock();
  }
}

/** 備份到 Drive、設公開可讀、回傳檔案 ID（suffix：多槽位時的「_槽位N」） */
function backupAndGetId_(p, blob, suffix) {
  const root = DriveApp.getFolderById(BACKUP_FOLDER_ID);
  const ym   = getOrCreate_(root, p.ym);
  const room = getOrCreate_(ym, p.room);
  const name = p.room + '__' + p.item + (suffix || '') + '.jpg';

  // 覆蓋舊檔（同槽位重拍時，舊照移到垃圾桶）
  const dup = room.getFilesByName(name);
  if (dup.hasNext()) dup.next().setTrashed(true);

  const file = room.createFile(blob.copyBlob().setName(name));

  // 設為公開可讀（CellImage / IMAGE 公式必須能存取）
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  return file.getId();
}

/**
 * 取得或建立資料夾
 * ★v4.3.1★ 建立前加鎖＋二次檢查：月初兩人同毫秒上傳時，
 * 舊寫法會各建一個同名資料夾（Drive 允許同名），備份檔因此散落兩處。
 * 資料夾已存在時走快路徑，完全不進鎖、不影響日常速度。
 */
function getOrCreate_(parent, name) {
  let it = parent.getFoldersByName(name);
  if (it.hasNext()) return it.next();   // 快路徑：已存在

  const lock = LockService.getScriptLock();
  lock.waitLock(60000);   // v4.6：月初分頁建立含 Z 欄重建較久，等候放寬到 60 秒
  try {
    it = parent.getFoldersByName(name); // 拿到鎖後再查一次
    return it.hasNext() ? it.next() : parent.createFolder(name);
  } finally {
    lock.releaseLock();
  }
}

/** 每月開新報表：複製母版、清空 A 欄 IMAGE 公式（保留文字格線與 Z 欄定位） */
function newMonthCopy(ymLabel) {
  try {
    const src  = DriveApp.getFileById(MASTER_SHEET_ID);
    const copy = src.makeCopy(ymLabel + '_高公局保養報表');
    const ss   = SpreadsheetApp.openById(copy.getId());
    const sh   = ss.getSheetByName(SHEET_TAB) || ss.getSheets()[0];
    // 清空 A 欄所有 IMAGE 公式（批次一次清除，保留其他格線與 Z 欄）
    const lastRow = sh.getLastRow();
    sh.getRange(1, 1, lastRow, 1).clearContent();
    console.log('[newMonthCopy] 完成：' + ymLabel);
    return { id: copy.getId(), url: ss.getUrl() };
  } catch (err) {
    console.error('[newMonthCopy] 失敗：' + err);
    throw err;   // 手動執行的工具函式，直接拋出讓執行記錄看得到
  }
}

/**
 * 執行一次：把所有槽位定位字串寫入 Z 欄
 * ★v4 改為批次寫入★
 *   v3 在迴圈內逐格 setValue（664 次 API 呼叫，分鐘級、大表可能逾時）。
 *   v4 先把整條 Z 欄 getValues() 讀進陣列 → 在記憶體中填入定位鍵 →
 *   最後 setValues() 一次寫回（2 次 API 呼叫，秒級完成）。
 *   Z 欄原有的其他內容會原樣保留（先讀後寫，只覆蓋要寫鍵的那幾格）。
 */
function buildZColumn() {
  try {
    const ss = SpreadsheetApp.openById(MASTER_SHEET_ID);
    const sh = ss.getSheetByName(SHEET_TAB) || ss.getSheets()[0];
    const lastRow = sh.getLastRow();
    const fVals = sh.getRange(1, 6, lastRow, 1).getValues();      // F 欄：一次讀入
    const zVals = sh.getRange(1, Z_COL, lastRow, 1).getValues();  // Z 欄：先讀既有內容

    let written = 0;
    let currentRoom = '';
    let photoStartRow = -1;

    for (let i = 0; i < fVals.length; i++) {
      const fVal = String(fVals[i][0] || '').trim();
      const row = i + 1;

      // 保養檢查項目的下一列 = 照片格起始列
      if (fVal.startsWith('保養檢查項目')) {
        photoStartRow = row + 1;
      }
      // 抓機房名稱
      if (fVal.startsWith('機房名稱:')) {
        currentRoom = fVal.replace('機房名稱:', '').trim();
      }
      // 抓照片說明 → 組 key → 填進 Z 欄陣列（照片格起始列）
      if (fVal.startsWith('照片內容說明:') && currentRoom && photoStartRow > 0) {
        let desc = fVal.replace('照片內容說明:', '').trim();
        // 說明文字太長被折到下一列時，把下一列接上（排除撞到其他標題列的情況）
        const nextVal = String(fVals[i + 1] ? fVals[i + 1][0] : '').trim();
        if (nextVal &&
            !nextVal.startsWith('施工') && !nextVal.startsWith('完工') &&
            !nextVal.startsWith('機房') && !nextVal.startsWith('保養') &&
            !nextVal.startsWith('照片')) {
          desc += nextVal;
        }
        const key = currentRoom + '__' + desc;
        if (photoStartRow <= lastRow) {
          zVals[photoStartRow - 1][0] = key;   // 寫進記憶體陣列（0-based）
        } else {
          // 極端狀況：照片格起始列超出目前資料範圍，退回單格寫入
          sh.getRange(photoStartRow, Z_COL).setValue(key);
        }
        written++;
      }
    }

    // ★一次寫回整條 Z 欄（批次寫入的關鍵）★
    sh.getRange(1, Z_COL, lastRow, 1).setValues(zVals);
    SpreadsheetApp.flush();

    console.log('[buildZColumn] 完成，共寫入 ' + written + ' 個（寫在照片格起始列）');
    return '完成，共寫入 ' + written + ' 個定位字串';
  } catch (err) {
    console.error('[buildZColumn] 失敗：' + err);
    throw err;
  }
}

/**
 * ★一鍵重建 Z 欄定位鍵（解決「定位鍵重複」）★
 * 背景：模板列位置調整過後再跑 buildZColumn，舊鍵會殘留在位移前的
 * 格子（buildZColumn 設計上保留 Z 欄既有內容），新舊鍵並存 → 重複。
 * 本函式「先清空整條 Z 欄、再重建」，保證每個槽位只有一個鍵。
 * 執行方式：在 GAS 編輯器上方函式選單選 rebuildZColumnEverywhere → 執行。
 * 會自動處理「模板」與所有六位數年月分頁（202607、202608…），其他分頁不碰。
 */
function rebuildZColumnEverywhere() {
  const ss = SpreadsheetApp.openById(MASTER_SHEET_ID);
  const results = [];
  ss.getSheets().forEach(function (sh) {
    const name = sh.getName();
    if (name !== TEMPLATE_TAB_NAME && !/^\d{6}$/.test(name)) return; // 只碰模板與年月分頁
    results.push(name + '：' + rebuildZColumnFor_(sh));
  });
  const msg = results.join('\n');
  console.log('[rebuildZColumnEverywhere]\n' + msg);
  return msg;
}

/** 對單一分頁清空並重建 Z 欄（掃描邏輯與 buildZColumn 完全相同） */
function rebuildZColumnFor_(sh) {
  const lastRow = sh.getLastRow();
  const fVals = sh.getRange(1, 6, lastRow, 1).getValues();   // F 欄一次讀入
  // ★與 buildZColumn 唯一的差異：從「全空」陣列開始 → 殘留舊鍵一律清除
  const zVals = [];
  for (let r = 0; r < lastRow; r++) zVals.push(['']);

  let written = 0, currentRoom = '', photoStartRow = -1;
  for (let i = 0; i < fVals.length; i++) {
    const fVal = String(fVals[i][0] || '').trim();
    if (fVal.startsWith('保養檢查項目')) photoStartRow = i + 2;   // 下一列 = 照片格起始列
    if (fVal.startsWith('機房名稱:')) currentRoom = fVal.replace('機房名稱:', '').trim();
    if (fVal.startsWith('照片內容說明:') && currentRoom && photoStartRow > 0) {
      let desc = fVal.replace('照片內容說明:', '').trim();
      // 說明文字折到下一列時接上（排除撞到其他標題列）
      const nextVal = String(fVals[i + 1] ? fVals[i + 1][0] : '').trim();
      if (nextVal && !nextVal.startsWith('施工') && !nextVal.startsWith('完工') &&
          !nextVal.startsWith('機房') && !nextVal.startsWith('保養') && !nextVal.startsWith('照片')) {
        desc += nextVal;
      }
      if (photoStartRow <= lastRow) {
        zVals[photoStartRow - 1][0] = currentRoom + '__' + desc;
        written++;
      }
    }
  }
  sh.getRange(1, Z_COL, lastRow, 1).setValues(zVals);   // 一次寫回
  SpreadsheetApp.flush();
  return '重建完成，共寫入 ' + written + ' 個定位鍵';
}

/**
 * ★v4.6★ JPEG 完整性檢查：回傳 null = 正常，回傳字串 = 不合格原因
 * 三道防線（實測可攔下「傳一半就斷」「極短殘骸」「非 JPEG」）：
 *   ① 檔案過小：正常照片經前端壓縮後至少數十 KB，3KB 以下必為殘骸
 *   ② 檔頭 FFD8（SOI）：不是 JPEG 就擋
 *   ③ 檔尾 FFD9（EOI）：★截斷偵測的關鍵★ 傳一半的檔案結尾必定不是 FFD9
 * 已知限制：頭尾都在、僅中段損毀的檔案無法由位元組判斷（機率極低，
 *   且此類檔案 Google 圖床通常仍能顯示大部分內容）。
 */
function validateJpegBytes_(bytes) {
  if (!bytes || bytes.length < 3000) {
    return '檔案過小 ' + (bytes ? bytes.length : 0) + ' bytes';
  }
  if (bytes[0] !== -1 || bytes[1] !== -40) {          // 0xFF 0xD8
    return '非 JPEG 檔頭';
  }
  if (bytes[bytes.length - 2] !== -1 || bytes[bytes.length - 1] !== -39) {   // 0xFF 0xD9
    return '檔尾遺失，照片只傳了一部分';
  }
  return null;
}

/**
 * ★一鍵清空「模板」分頁的照片，並還原標題與日期★
 * 用途：把某個月份分頁升級成新模板後，用這個把該月的照片、日期、
 *      標題年月洗乾淨，避免內容被複製到未來每一個月。
 * 執行方式：GAS 編輯器上方函式下拉選單選 clearTemplatePhotos → 按執行
 *          （這是手動工具，不需要重新部署）
 */
function clearTemplatePhotos() {
  const ss = SpreadsheetApp.openById(MASTER_SHEET_ID);
  const sh = ss.getSheetByName(TEMPLATE_TAB_NAME);
  if (!sh) throw new Error('找不到「' + TEMPLATE_TAB_NAME + '」分頁');

  const log = [];

  // ① 移除浮動圖片（舊版 insertImage 殘留物）
  let imgN = 0;
  try {
    sh.getImages().forEach(function (im) { im.remove(); imgN++; });
  } catch (e) { log.push('浮動圖片略過：' + e.message); }
  log.push('移除浮動圖片 ' + imgN + ' 張');

  // ② 只清「A~D 合併且高度 ≥5 列」的照片格，補回「照片」佔位字
  //    ⚠ 絕不整欄 clearContent：A 欄同時放著 229 個頁首大標題，整欄清會全毀
  const boxes = [];
  sh.getDataRange().getMergedRanges().forEach(function (mr) {
    if (mr.getColumn() === 1 && mr.getLastColumn() === 4 && mr.getNumRows() >= 5) {
      boxes.push(mr.getA1Notation());
    }
  });
  if (boxes.length > 0) {
    const rl = sh.getRangeList(boxes);
    rl.clearContent();                    // 清掉 CellImage 與 =IMAGE 公式，保留框線背景
    rl.setValue('照片')
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle')
      .setFontWeight('bold')
      .setFontSize(28);
  }
  log.push('清空照片格 ' + boxes.length + ' 個並補回「照片」');

  // ③ 日期還原成 115/05/07
  //    選單按鈕.gs 用 regex「115/05/\d{2}」批次換日期，模板若留著其他月份會失效
  const d1 = sh.createTextFinder('施工日期.*').useRegularExpression(true)
               .replaceAllWith('施工日期:115/05/07');
  const d2 = sh.createTextFinder('完工日期.*').useRegularExpression(true)
               .replaceAllWith('完工日期:115/05/07');
  log.push('重置日期：施工 ' + d1 + ' 格、完工 ' + d2 + ' 格');

  // ④ 標題年月還原成 TEMPLATE_OLD_YM_TEXT，後端建立新分頁時才找得到並替換
  //    （把常見的月份都掃一遍，模板來自哪個月都能還原）
  let tN = 0;
  for (let m = 1; m <= 12; m++) {
    const s = '115年' + m + '月';
    if (s === TEMPLATE_OLD_YM_TEXT) continue;
    tN += sh.createTextFinder(s).replaceAllWith(TEMPLATE_OLD_YM_TEXT);
  }
  log.push('標題年月還原為 ' + TEMPLATE_OLD_YM_TEXT + '（' + tN + ' 處）');

  // ⑤ 重建 Z 欄定位鍵
  log.push('Z 欄：' + rebuildZColumnFor_(sh));

  SpreadsheetApp.flush();
  const msg = '【模板清理完成】\n' + log.join('\n');
  console.log(msg);
  return msg;
}

/** 診斷用：在 Apps Script 執行這個，看看試算表能不能開 */
function diagTest() {
  try {
    const ss = SpreadsheetApp.openById(MASTER_SHEET_ID);
    const sh = ss.getSheetByName(SHEET_TAB) || ss.getSheets()[0];
    const name = ss.getName();
    return '✓ 試算表可以存取：' + name + '，分頁：' + sh.getName();
  } catch (e) {
    return '✗ 錯誤：' + e.message;
  }
}

// ===== 選單資料（每月固定，與原版完全相同、原封不動）=====
const MENU_TREE = {"汐止機房": {"AHU-01": [{"label": "濾網及蒸發器檢查及清潔", "desc": "AHU-01濾網及蒸發器檢查及清潔"}, {"label": "壓縮機運轉電流值測試", "desc": "AHU-01壓縮機運轉電流值測試"}, {"label": "送風風扇軸檢查", "desc": "AHU-01送風風扇軸檢查"}, {"label": "蒸發器洩水管檢查及清潔", "desc": "AHU-01蒸發器洩水管檢查及清潔"}, {"label": "設備外觀照", "desc": "AHU-01設備外觀照"}], "AHU-02": [{"label": "濾網及蒸發器檢查及清潔", "desc": "AHU-02濾網及蒸發器檢查及清潔"}, {"label": "壓縮機運轉電流值測試", "desc": "AHU-02壓縮機運轉電流值測試"}, {"label": "送風風扇軸檢查", "desc": "AHU-02送風風扇軸檢查"}, {"label": "蒸發器洩水管檢查及清潔", "desc": "AHU-02蒸發器洩水管檢查及清潔"}, {"label": "設備外觀照", "desc": "AHU-02設備外觀照"}], "ACP-01": [{"label": "電壓檢查", "desc": "ACP-01電壓檢查"}, {"label": "各馬達電流檢查", "desc": "ACP-01各馬達電流檢查"}, {"label": "絕緣檢查", "desc": "ACP-01絕緣檢查"}, {"label": "啟動時間繼電時間檢查及調整", "desc": "ACP-01啟動時間繼電時間檢查及調整"}, {"label": "設備外觀照", "desc": "ACP-01設備外觀照"}]}, "內湖機房": {"AHU-01": [{"label": "濾網及蒸發器檢查及清潔", "desc": "AHU-01濾網及蒸發器檢查及清潔"}, {"label": "壓縮機運轉電流值測試", "desc": "AHU-01壓縮機運轉電流值測試"}, {"label": "送風風扇軸承檢查", "desc": "AHU-01送風風扇軸承檢查"}, {"label": "蒸發器洩水管檢查及清潔", "desc": "AHU-01蒸發器洩水管檢查及清潔"}, {"label": "設備外觀照", "desc": "AHU-01設備外觀照"}], "AHU-02": [{"label": "濾網及蒸發器檢查及清潔", "desc": "AHU-02濾網及蒸發器檢查及清潔"}, {"label": "壓縮機運轉電流值測試", "desc": "AHU-02壓縮機運轉電流值測試"}, {"label": "送風風扇軸承檢查", "desc": "AHU-02送風風扇軸承檢查"}, {"label": "蒸發器洩水管檢查及清潔", "desc": "AHU-02蒸發器洩水管檢查及清潔"}, {"label": "設備外觀照", "desc": "AHU-02設備外觀照"}], "AHU-03": [{"label": "濾網及蒸發器檢查及清潔", "desc": "AHU-03濾網及蒸發器檢查及清潔"}, {"label": "壓縮機運轉電流值測試", "desc": "AHU-03壓縮機運轉電流值測試"}, {"label": "送風風扇軸承檢查", "desc": "AHU-03送風風扇軸承檢查"}, {"label": "蒸發器洩水管檢查及清潔", "desc": "AHU-03蒸發器洩水管檢查及清潔"}, {"label": "設備外觀照", "desc": "AHU-03設備外觀照"}], "AHU-04": [{"label": "濾網及蒸發器檢查及清潔", "desc": "AHU-04濾網及蒸發器檢查及清潔"}, {"label": "壓縮機運轉電流值測試", "desc": "AHU-04壓縮機運轉電流值測試"}, {"label": "送風風扇軸承檢查", "desc": "AHU-04送風風扇軸承檢查"}, {"label": "蒸發器洩水管檢查及清潔", "desc": "AHU-04蒸發器洩水管檢查及清潔"}, {"label": "設備外觀照", "desc": "AHU-04設備外觀照"}], "ACP-01": [{"label": "電壓檢查", "desc": "ACP-01電壓檢查"}, {"label": "各馬達電流檢查", "desc": "ACP-01各馬達電流檢查"}, {"label": "絕緣檢查", "desc": "ACP-01絕緣檢查"}, {"label": "啟動時間繼電時間檢查調整", "desc": "ACP-01啟動時間繼電時間檢查調整"}, {"label": "設備外觀照", "desc": "ACP-01設備外觀照"}], "ACP-02": [{"label": "電壓檢查", "desc": "ACP-02電壓檢查"}, {"label": "各馬達電流檢查", "desc": "ACP-02各馬達電流檢查"}, {"label": "絕緣檢查", "desc": "ACP-02絕緣檢查"}, {"label": "啟動時間繼電時間檢查及調整", "desc": "ACP-02啟動時間繼電時間檢查及調整"}, {"label": "設備外觀照", "desc": "ACP-02設備外觀照"}]}, "泰控機房": {"CHU-01": [{"label": "高低壓儀表檢查", "desc": "CHU-01高低壓儀表檢查"}, {"label": "壓縮機高低壓止閥、冷凝器之", "desc": "CHU-01壓縮機高低壓止閥、冷凝器之"}, {"label": "潤滑油位及油溫檢查", "desc": "CHU-01潤滑油位及油溫檢查"}, {"label": "馬達絕緣值檢查", "desc": "CHU-01馬達絕緣值檢查"}, {"label": "設備外觀照", "desc": "CHU-01設備外觀照"}], "其他": [{"label": "泰控-CHU-01液管視窗冷媒流量檢查", "desc": "泰控-CHU-01液管視窗冷媒流量檢查"}, {"label": "-CHU-03液管視窗冷媒流量檢查", "desc": "-CHU-03液管視窗冷媒流量檢查"}, {"label": ":CTR-02管路系統定期換水", "desc": ":CTR-02管路系統定期換水"}, {"label": "冰水空調送風機排水疏通之1", "desc": "冰水空調送風機排水疏通之1"}, {"label": "冰水空調送風機排水疏通之2", "desc": "冰水空調送風機排水疏通之2"}, {"label": "冰水空調送風機排水疏通之3", "desc": "冰水空調送風機排水疏通之3"}, {"label": "冰水空調送風機排水疏通之4", "desc": "冰水空調送風機排水疏通之4"}, {"label": "冰水空調送風機排水疏通之5", "desc": "冰水空調送風機排水疏通之5"}, {"label": "冰水空調送風機排水疏通之6", "desc": "冰水空調送風機排水疏通之6"}, {"label": "冰水空調送風機排水疏通之7", "desc": "冰水空調送風機排水疏通之7"}, {"label": "送風風扇軸承檢查", "desc": "送風風扇軸承檢查"}, {"label": "冷卻水加藥設備檢查", "desc": "冷卻水加藥設備檢查"}, {"label": "冷卻水加藥設備外觀照", "desc": "冷卻水加藥設備外觀照"}, {"label": "空氣濾清器電源檢測", "desc": "空氣濾清器電源檢測"}, {"label": "空氣濾清器功能檢測", "desc": "空氣濾清器功能檢測"}, {"label": "空氣濾清器外觀照", "desc": "空氣濾清器外觀照"}], "CHU-02": [{"label": "高低壓儀表檢查", "desc": "CHU-02高低壓儀表檢查"}, {"label": "油壓溫度等,開關檢查", "desc": "CHU-02油壓溫度等,開關檢查"}, {"label": "壓縮機高低壓止閥、冷凝器之", "desc": "CHU-02壓縮機高低壓止閥、冷凝器之"}, {"label": "潤滑油位及油溫檢查", "desc": "CHU-02潤滑油位及油溫檢查"}, {"label": "馬達絕緣值檢查", "desc": "CHU-02馬達絕緣值檢查"}, {"label": "設備外觀照", "desc": "CHU-02設備外觀照"}], "CHU-03": [{"label": "高低壓儀表檢查", "desc": "CHU-03高低壓儀表檢查"}, {"label": "壓縮機高低壓止閥、冷凝器之", "desc": "CHU-03壓縮機高低壓止閥、冷凝器之"}, {"label": "潤滑油位及油溫檢查", "desc": "CHU-03潤滑油位及油溫檢查"}, {"label": "馬達絕緣值檢查", "desc": "CHU-03馬達絕緣值檢查"}, {"label": "設備外觀照", "desc": "CHU-03設備外觀照"}], "CTR-01": [{"label": "水泵,水塔風扇電流值檢查", "desc": "CTR-01水泵,水塔風扇電流值檢查"}, {"label": "水泵,水封洩水及軸承異聲", "desc": "CTR-01水泵,水封洩水及軸承異聲"}, {"label": "水塔風扇軸承異聲檢查", "desc": "CTR-01水塔風扇軸承異聲檢查"}, {"label": "水塔盛水盤與外殼清洗", "desc": "CTR-01水塔盛水盤與外殼清洗"}, {"label": "噴水頭檢查", "desc": "CTR-01噴水頭檢查"}, {"label": "管路系統定期換水", "desc": "CTR-01管路系統定期換水"}, {"label": "設備外觀照-1", "desc": "CTR-01設備外觀照-1"}, {"label": "設備外觀照-2", "desc": "CTR-01設備外觀照-2"}], "CTR-02": [{"label": "水泵,水塔風扇電流值檢查", "desc": "CTR-02水泵,水塔風扇電流值檢查"}, {"label": "水泵,水封洩水及軸承異聲", "desc": "CTR-02水泵,水封洩水及軸承異聲"}, {"label": "水塔風扇軸承異聲檢查", "desc": "CTR-02水塔風扇軸承異聲檢查"}, {"label": "水塔盛水盤與外殼清洗", "desc": "CTR-02水塔盛水盤與外殼清洗"}, {"label": "噴水頭檢查", "desc": "CTR-02噴水頭檢查"}, {"label": "設備外觀照-1", "desc": "CTR-02設備外觀照-1"}, {"label": "設備外觀照-2", "desc": "CTR-02設備外觀照-2"}], "CTR-03": [{"label": "水泵,水塔風扇電流值檢查", "desc": "CTR-03水泵,水塔風扇電流值檢查"}, {"label": "水泵,水封洩水及軸承異聲", "desc": "CTR-03水泵,水封洩水及軸承異聲"}, {"label": "水塔風扇軸承異聲檢查", "desc": "CTR-03水塔風扇軸承異聲檢查"}, {"label": "水塔盛水盤與外殼清洗", "desc": "CTR-03水塔盛水盤與外殼清洗"}, {"label": "噴水頭檢查", "desc": "CTR-03噴水頭檢查"}, {"label": "管路系統定期換水", "desc": "CTR-03管路系統定期換水"}, {"label": "設備外觀照-1", "desc": "CTR-03設備外觀照-1"}, {"label": "設備外觀照-2", "desc": "CTR-03設備外觀照-2"}], "ACU-01": [{"label": "電源線,插頭及插座損壞或鬆動", "desc": "ACU-01電源線,插頭及插座損壞或鬆動"}, {"label": "壓縮機啟動,運轉電流值檢查", "desc": "ACU-01壓縮機啟動,運轉電流值檢查"}, {"label": "運轉風扇馬達軸承檢查", "desc": "ACU-01運轉風扇馬達軸承檢查"}, {"label": "排水管清潔", "desc": "ACU-01排水管清潔"}, {"label": "運轉時異聲及異常震動檢查", "desc": "ACU-01運轉時異聲及異常震動檢查"}, {"label": "冷氣出風口溫度量測", "desc": "ACU-01冷氣出風口溫度量測"}, {"label": "設備外觀照", "desc": "ACU-01設備外觀照"}], "ACU-02": [{"label": "電源線,插頭及插座損壞或鬆動", "desc": "ACU-02電源線,插頭及插座損壞或鬆動"}, {"label": "壓縮機啟動,運轉電流值檢查", "desc": "ACU-02壓縮機啟動,運轉電流值檢查"}, {"label": "運轉風扇馬達軸承檢查", "desc": "ACU-02運轉風扇馬達軸承檢查"}, {"label": "排水管清潔", "desc": "ACU-02排水管清潔"}, {"label": "運轉時異聲及異常震動檢查", "desc": "ACU-02運轉時異聲及異常震動檢查"}, {"label": "冷氣出風口溫度量測", "desc": "ACU-02冷氣出風口溫度量測"}, {"label": "設備外觀照", "desc": "ACU-02設備外觀照"}], "ACU-03": [{"label": "電源線,插頭及插座損壞或鬆動", "desc": "ACU-03電源線,插頭及插座損壞或鬆動"}, {"label": "壓縮機啟動,運轉電流值檢查", "desc": "ACU-03壓縮機啟動,運轉電流值檢查"}, {"label": "運轉風扇馬達軸承檢查", "desc": "ACU-03運轉風扇馬達軸承檢查"}, {"label": "排水管清潔", "desc": "ACU-03排水管清潔"}, {"label": "運轉時異聲及異常震動檢查", "desc": "ACU-03運轉時異聲及異常震動檢查"}, {"label": "冷氣出風口溫度量測", "desc": "ACU-03冷氣出風口溫度量測"}, {"label": "設備外觀照", "desc": "ACU-03設備外觀照"}], "ACU-04": [{"label": "電源線,插頭及插座損壞或鬆動", "desc": "ACU-04電源線,插頭及插座損壞或鬆動"}, {"label": "壓縮機啟動,運轉電流值檢查", "desc": "ACU-04壓縮機啟動,運轉電流值檢查"}, {"label": "運轉風扇馬達軸承檢查", "desc": "ACU-04運轉風扇馬達軸承檢查"}, {"label": "排水管清潔", "desc": "ACU-04排水管清潔"}, {"label": "運轉時異聲及異常震動檢查", "desc": "ACU-04運轉時異聲及異常震動檢查"}, {"label": "冷氣出風口溫度量測", "desc": "ACU-04冷氣出風口溫度量測"}, {"label": "設備外觀照", "desc": "ACU-04設備外觀照"}], "ACU-05": [{"label": "電源線,插頭及插座損壞或鬆動", "desc": "ACU-05電源線,插頭及插座損壞或鬆動"}, {"label": "空氣過濾器清潔", "desc": "ACU-05空氣過濾器清潔"}, {"label": "壓縮機啟動,運轉電流值檢查", "desc": "ACU-05壓縮機啟動,運轉電流值檢查"}, {"label": "運轉風扇馬達軸承檢查", "desc": "ACU-05運轉風扇馬達軸承檢查"}, {"label": "排水管清潔", "desc": "ACU-05排水管清潔"}, {"label": "運轉時異聲及異常震動檢查", "desc": "ACU-05運轉時異聲及異常震動檢查"}, {"label": "冷氣出風口溫度量測", "desc": "ACU-05冷氣出風口溫度量測"}, {"label": "設備外觀照", "desc": "ACU-05設備外觀照"}], "ACU-06": [{"label": "電源線,插頭及插座損壞或鬆動", "desc": "ACU-06電源線,插頭及插座損壞或鬆動"}, {"label": "空氣過濾器清潔", "desc": "ACU-06空氣過濾器清潔"}, {"label": "壓縮機啟動,運轉電流值檢查", "desc": "ACU-06壓縮機啟動,運轉電流值檢查"}, {"label": "運轉風扇馬達軸承檢查", "desc": "ACU-06運轉風扇馬達軸承檢查"}, {"label": "排水管清潔", "desc": "ACU-06排水管清潔"}, {"label": "運轉時異聲及異常震動檢查", "desc": "ACU-06運轉時異聲及異常震動檢查"}, {"label": "冷氣出風口溫度量測", "desc": "ACU-06冷氣出風口溫度量測"}, {"label": "設備外觀照", "desc": "ACU-06設備外觀照"}], "ACP-05": [{"label": "電壓檢查", "desc": "ACP-05電壓檢查"}, {"label": "各馬達電流檢查", "desc": "ACP-05各馬達電流檢查"}, {"label": "絕緣檢查", "desc": "ACP-05絕緣檢查"}, {"label": "啟動時間繼電時間檢查及調整", "desc": "ACP-05啟動時間繼電時間檢查及調整"}, {"label": "設備外觀照", "desc": "ACP-05設備外觀照"}], "ACP-08": [{"label": "電壓檢查", "desc": "ACP-08電壓檢查"}, {"label": "各馬達電流檢查", "desc": "ACP-08各馬達電流檢查"}, {"label": "絕緣檢查", "desc": "ACP-08絕緣檢查"}, {"label": "啟動時間繼電時間檢查及調整", "desc": "ACP-08啟動時間繼電時間檢查及調整"}, {"label": "設備外觀照", "desc": "ACP-08設備外觀照"}], "ACU-07": [{"label": "電源線,插頭及插座損壞或鬆動", "desc": "ACU-07電源線,插頭及插座損壞或鬆動"}, {"label": "空氣過濾器清潔", "desc": "ACU-07空氣過濾器清潔"}, {"label": "壓縮機啟動,運轉電流值檢查", "desc": "ACU-07壓縮機啟動,運轉電流值檢查"}, {"label": "運轉風扇馬達軸承檢查", "desc": "ACU-07運轉風扇馬達軸承檢查"}, {"label": "排水管清潔", "desc": "ACU-07排水管清潔"}, {"label": "運轉時異聲及異常震動檢查", "desc": "ACU-07運轉時異聲及異常震動檢查"}, {"label": "冷氣出風口溫度量測", "desc": "ACU-07冷氣出風口溫度量測"}, {"label": "設備外觀照", "desc": "ACU-07設備外觀照"}], "ACU-08": [{"label": "電源線,插頭及插座損壞或鬆動", "desc": "ACU-08電源線,插頭及插座損壞或鬆動"}, {"label": "空氣過濾器清潔", "desc": "ACU-08空氣過濾器清潔"}, {"label": "壓縮機啟動,運轉電流值檢查", "desc": "ACU-08壓縮機啟動,運轉電流值檢查"}, {"label": "運轉風扇馬達軸承檢查", "desc": "ACU-08運轉風扇馬達軸承檢查"}, {"label": "排水管清潔", "desc": "ACU-08排水管清潔"}, {"label": "運轉時異聲及異常震動檢查", "desc": "ACU-08運轉時異聲及異常震動檢查"}, {"label": "冷氣出風口溫度量測", "desc": "ACU-08冷氣出風口溫度量測"}, {"label": "設備外觀照", "desc": "ACU-08設備外觀照"}], "ACU-09": [{"label": "電源線,插頭及插座損壞或鬆動", "desc": "ACU-09電源線,插頭及插座損壞或鬆動"}, {"label": "空氣過濾器清潔", "desc": "ACU-09空氣過濾器清潔"}, {"label": "壓縮機啟動,運轉電流值檢查", "desc": "ACU-09壓縮機啟動,運轉電流值檢查"}, {"label": "運轉風扇馬達軸承檢查", "desc": "ACU-09運轉風扇馬達軸承檢查"}, {"label": "排水管清潔", "desc": "ACU-09排水管清潔"}, {"label": "運轉時異聲及異常震動檢查", "desc": "ACU-09運轉時異聲及異常震動檢查"}, {"label": "冷氣出風口溫度量測", "desc": "ACU-09冷氣出風口溫度量測"}, {"label": "設備外觀照", "desc": "ACU-09設備外觀照"}], "ACU-10": [{"label": "電源線,插頭及插座損壞或鬆動", "desc": "ACU-10電源線,插頭及插座損壞或鬆動"}, {"label": "空氣過濾器清潔", "desc": "ACU-10空氣過濾器清潔"}, {"label": "壓縮機啟動,運轉電流值檢查", "desc": "ACU-10壓縮機啟動,運轉電流值檢查"}, {"label": "運轉風扇馬達軸承檢查", "desc": "ACU-10運轉風扇馬達軸承檢查"}, {"label": "排水管清潔", "desc": "ACU-10排水管清潔"}, {"label": "運轉時異聲及異常震動檢查", "desc": "ACU-10運轉時異聲及異常震動檢查"}, {"label": "冷氣出風口溫度量測", "desc": "ACU-10冷氣出風口溫度量測"}, {"label": "設備外觀照", "desc": "ACU-10設備外觀照"}], "ACU-11": [{"label": "電源線,插頭及插座損壞或鬆動", "desc": "ACU-11電源線,插頭及插座損壞或鬆動"}, {"label": "空氣過濾器清潔", "desc": "ACU-11空氣過濾器清潔"}, {"label": "壓縮機啟動,運轉電流值檢查", "desc": "ACU-11壓縮機啟動,運轉電流值檢查"}, {"label": "運轉風扇馬達軸承檢查", "desc": "ACU-11運轉風扇馬達軸承檢查"}, {"label": "排水管清潔", "desc": "ACU-11排水管清潔"}, {"label": "運轉時異聲及異常震動檢查", "desc": "ACU-11運轉時異聲及異常震動檢查"}, {"label": "冷氣出風口溫度量測", "desc": "ACU-11冷氣出風口溫度量測"}, {"label": "設備外觀照", "desc": "ACU-11設備外觀照"}], "ACU-12": [{"label": "電源線,插頭及插座損壞或鬆動", "desc": "ACU-12電源線,插頭及插座損壞或鬆動"}, {"label": "壓縮機啟動,運轉電流值檢查", "desc": "ACU-12壓縮機啟動,運轉電流值檢查"}, {"label": "運轉風扇馬達軸承檢查", "desc": "ACU-12運轉風扇馬達軸承檢查"}, {"label": "排水管清潔", "desc": "ACU-12排水管清潔"}, {"label": "運轉時異聲及異常震動檢查", "desc": "ACU-12運轉時異聲及異常震動檢查"}, {"label": "冷氣出風口溫度量測", "desc": "ACU-12冷氣出風口溫度量測"}, {"label": "設備外觀照", "desc": "ACU-12設備外觀照"}], "ACU-13": [{"label": "電源線,插頭及插座損壞或鬆動", "desc": "ACU-13電源線,插頭及插座損壞或鬆動"}, {"label": "壓縮機啟動,運轉電流值檢查", "desc": "ACU-13壓縮機啟動,運轉電流值檢查"}, {"label": "運轉風扇馬達軸承檢查", "desc": "ACU-13運轉風扇馬達軸承檢查"}, {"label": "排水管清潔", "desc": "ACU-13排水管清潔"}, {"label": "運轉時異聲及異常震動檢查", "desc": "ACU-13運轉時異聲及異常震動檢查"}, {"label": "冷氣出風口溫度量測", "desc": "ACU-13冷氣出風口溫度量測"}, {"label": "設備外觀照", "desc": "ACU-13設備外觀照"}], "ACU-14": [{"label": "電源線,插頭及插座損壞或鬆動", "desc": "ACU-14電源線,插頭及插座損壞或鬆動"}, {"label": "壓縮機啟動,運轉電流值檢查", "desc": "ACU-14壓縮機啟動,運轉電流值檢查"}, {"label": "運轉風扇馬達軸承檢查", "desc": "ACU-14運轉風扇馬達軸承檢查"}, {"label": "排水管清潔", "desc": "ACU-14排水管清潔"}, {"label": "運轉時異聲及異常震動檢查", "desc": "ACU-14運轉時異聲及異常震動檢查"}, {"label": "冷氣出風口溫度量測", "desc": "ACU-14冷氣出風口溫度量測"}, {"label": "設備外觀照", "desc": "ACU-14設備外觀照"}], "ACU-15": [{"label": "電源線,插頭及插座損壞或鬆動", "desc": "ACU-15電源線,插頭及插座損壞或鬆動"}, {"label": "壓縮機啟動,運轉電流值檢查", "desc": "ACU-15壓縮機啟動,運轉電流值檢查"}, {"label": "運轉風扇馬達軸承檢查", "desc": "ACU-15運轉風扇馬達軸承檢查"}, {"label": "排水管清潔", "desc": "ACU-15排水管清潔"}, {"label": "運轉時異聲及異常震動檢查", "desc": "ACU-15運轉時異聲及異常震動檢查"}, {"label": "冷氣出風口溫度量測", "desc": "ACU-15冷氣出風口溫度量測"}, {"label": "設備外觀照", "desc": "ACU-15設備外觀照"}], "ACU-16": [{"label": "電源線,插頭及插座損壞或鬆動", "desc": "ACU-16電源線,插頭及插座損壞或鬆動"}, {"label": "空氣過濾器清潔", "desc": "ACU-16空氣過濾器清潔"}, {"label": "壓縮機啟動,運轉電流值檢查", "desc": "ACU-16壓縮機啟動,運轉電流值檢查"}, {"label": "運轉風扇馬達軸承檢查", "desc": "ACU-16運轉風扇馬達軸承檢查"}, {"label": "排水管清潔", "desc": "ACU-16排水管清潔"}, {"label": "運轉時異聲及異常震動檢查", "desc": "ACU-16運轉時異聲及異常震動檢查"}, {"label": "冷氣出風口溫度量測", "desc": "ACU-16冷氣出風口溫度量測"}, {"label": "設備外觀照", "desc": "ACU-16設備外觀照"}], "ACU-17": [{"label": "電源線,插頭及插座損壞或鬆動", "desc": "ACU-17電源線,插頭及插座損壞或鬆動"}, {"label": "空氣過濾器清潔", "desc": "ACU-17空氣過濾器清潔"}, {"label": "壓縮機啟動,運轉電流值檢查", "desc": "ACU-17壓縮機啟動,運轉電流值檢查"}, {"label": "運轉風扇馬達軸承檢查", "desc": "ACU-17運轉風扇馬達軸承檢查"}, {"label": "排水管清潔", "desc": "ACU-17排水管清潔"}, {"label": "運轉時異聲及異常震動檢查", "desc": "ACU-17運轉時異聲及異常震動檢查"}, {"label": "冷氣出風口溫度量測", "desc": "ACU-17冷氣出風口溫度量測"}, {"label": "設備外觀照", "desc": "ACU-17設備外觀照"}], "ACU-18": [{"label": "電源線,插頭及插座損壞或鬆動", "desc": "ACU-18電源線,插頭及插座損壞或鬆動"}, {"label": "空氣過濾器清潔", "desc": "ACU-18空氣過濾器清潔"}, {"label": "壓縮機啟動,運轉電流值檢查", "desc": "ACU-18壓縮機啟動,運轉電流值檢查"}, {"label": "運轉風扇馬達軸承檢查", "desc": "ACU-18運轉風扇馬達軸承檢查"}, {"label": "排水管清潔", "desc": "ACU-18排水管清潔"}, {"label": "運轉時異聲及異常震動檢查", "desc": "ACU-18運轉時異聲及異常震動檢查"}, {"label": "冷氣出風口溫度量測", "desc": "ACU-18冷氣出風口溫度量測"}, {"label": "設備外觀照", "desc": "ACU-18設備外觀照"}], "AHU-01": [{"label": "濾網及蒸發器檢查及清潔", "desc": "AHU-01濾網及蒸發器檢查及清潔"}, {"label": "送風風扇軸承檢查測試", "desc": "AHU-01送風風扇軸承檢查測試"}, {"label": "蒸發器洩水管檢查及清潔", "desc": "AHU-01蒸發器洩水管檢查及清潔"}, {"label": "抽風機皮帶輪及耦合器固定鎖", "desc": "AHU-01抽風機皮帶輪及耦合器固定鎖"}, {"label": "傳動及皮帶檢查", "desc": "AHU-01傳動及皮帶檢查"}, {"label": "設備外觀照", "desc": "AHU-01設備外觀照"}], "ACP-01": [{"label": "電壓檢查", "desc": "ACP-01電壓檢查"}, {"label": "各馬達電流檢查", "desc": "ACP-01各馬達電流檢查"}, {"label": "絕緣檢查", "desc": "ACP-01絕緣檢查"}, {"label": "啟動時間繼電時間檢查調整", "desc": "ACP-01啟動時間繼電時間檢查調整"}, {"label": "設備外觀照", "desc": "ACP-01設備外觀照"}], "AHU-02": [{"label": "濾網及蒸發器檢查及清潔", "desc": "AHU-02濾網及蒸發器檢查及清潔"}, {"label": "送風風扇軸承檢查測試", "desc": "AHU-02送風風扇軸承檢查測試"}, {"label": "蒸發器洩水管檢查及清潔", "desc": "AHU-02蒸發器洩水管檢查及清潔"}, {"label": "抽風機皮帶輪及耦合器固定鎖", "desc": "AHU-02抽風機皮帶輪及耦合器固定鎖"}, {"label": "傳動及皮帶檢查", "desc": "AHU-02傳動及皮帶檢查"}, {"label": "設備外觀照", "desc": "AHU-02設備外觀照"}], "ACP-02": [{"label": "電壓檢查", "desc": "ACP-02電壓檢查"}, {"label": "各馬達電流檢查", "desc": "ACP-02各馬達電流檢查"}, {"label": "絕緣檢查", "desc": "ACP-02絕緣檢查"}, {"label": "啟動時間繼電時間檢查調整", "desc": "ACP-02啟動時間繼電時間檢查調整"}, {"label": "設備外觀照", "desc": "ACP-02設備外觀照"}], "AHU-03": [{"label": "濾網及蒸發器檢查及清潔", "desc": "AHU-03濾網及蒸發器檢查及清潔"}, {"label": "送風風扇軸承檢查測試", "desc": "AHU-03送風風扇軸承檢查測試"}, {"label": "蒸發器洩水管檢查及清潔", "desc": "AHU-03蒸發器洩水管檢查及清潔"}, {"label": "抽風機皮帶輪及耦合器固定鎖", "desc": "AHU-03抽風機皮帶輪及耦合器固定鎖"}, {"label": "傳動及皮帶檢查", "desc": "AHU-03傳動及皮帶檢查"}, {"label": "設備外觀照", "desc": "AHU-03設備外觀照"}], "ACP-03": [{"label": "電壓檢查", "desc": "ACP-03電壓檢查"}, {"label": "各馬達電流檢查", "desc": "ACP-03各馬達電流檢查"}, {"label": "絕緣檢查", "desc": "ACP-03絕緣檢查"}, {"label": "啟動時間繼電時間檢查調整", "desc": "ACP-03啟動時間繼電時間檢查調整"}, {"label": "設備外觀照", "desc": "ACP-03設備外觀照"}], "AHU-04": [{"label": "濾網及蒸發器檢查及清潔", "desc": "AHU-04濾網及蒸發器檢查及清潔"}, {"label": "送風風扇軸承檢查測試", "desc": "AHU-04送風風扇軸承檢查測試"}, {"label": "蒸發器洩水管檢查及清潔", "desc": "AHU-04蒸發器洩水管檢查及清潔"}, {"label": "抽風機皮帶輪及耦合器固定鎖", "desc": "AHU-04抽風機皮帶輪及耦合器固定鎖"}, {"label": "傳動及皮帶檢查", "desc": "AHU-04傳動及皮帶檢查"}, {"label": "設備外觀照", "desc": "AHU-04設備外觀照"}], "ACP-04": [{"label": "電壓檢查", "desc": "ACP-04電壓檢查"}, {"label": "各馬達電流檢查", "desc": "ACP-04各馬達電流檢查"}, {"label": "絕緣檢查", "desc": "ACP-04絕緣檢查"}, {"label": "啟動時間繼電時間檢查調整", "desc": "ACP-04啟動時間繼電時間檢查調整"}, {"label": "設備外觀照", "desc": "ACP-04設備外觀照"}], "AHU-05": [{"label": "濾網及蒸發器檢查及清潔", "desc": "AHU-05濾網及蒸發器檢查及清潔"}, {"label": "壓縮機運轉電流值測試", "desc": "AHU-05壓縮機運轉電流值測試"}, {"label": "送風風扇軸承檢查測試", "desc": "AHU-05送風風扇軸承檢查測試"}, {"label": "蒸發器洩水管檢查及清潔", "desc": "AHU-05蒸發器洩水管檢查及清潔"}, {"label": "設備外觀照", "desc": "AHU-05設備外觀照"}], "AHU-06": [{"label": "濾網及蒸發器檢查及清潔", "desc": "AHU-06濾網及蒸發器檢查及清潔"}, {"label": "壓縮機運轉電流值測試", "desc": "AHU-06壓縮機運轉電流值測試"}, {"label": "送風風扇軸承檢查測試", "desc": "AHU-06送風風扇軸承檢查測試"}, {"label": "蒸發器洩水管檢查及清潔", "desc": "AHU-06蒸發器洩水管檢查及清潔"}, {"label": "設備外觀照", "desc": "AHU-06設備外觀照"}], "ACP-06": [{"label": "電壓檢查", "desc": "ACP-06電壓檢查"}, {"label": "各馬達電流檢查", "desc": "ACP-06各馬達電流檢查"}, {"label": "絕緣檢查", "desc": "ACP-06絕緣檢查"}, {"label": "啟動時間繼電時間檢查調整", "desc": "ACP-06啟動時間繼電時間檢查調整"}, {"label": "設備外觀照", "desc": "ACP-06設備外觀照"}], "AHU-07": [{"label": "濾網及蒸發器檢查及清潔", "desc": "AHU-07濾網及蒸發器檢查及清潔"}, {"label": "壓縮機運轉電流值測試", "desc": "AHU-07壓縮機運轉電流值測試"}, {"label": "送風風扇軸承檢查測試", "desc": "AHU-07送風風扇軸承檢查測試"}, {"label": "蒸發器洩水管檢查及清潔", "desc": "AHU-07蒸發器洩水管檢查及清潔"}, {"label": "設備外觀照", "desc": "AHU-07設備外觀照"}], "AHU-08": [{"label": "濾網及蒸發器檢查及清潔", "desc": "AHU-08濾網及蒸發器檢查及清潔"}, {"label": "壓縮機運轉電流值測試", "desc": "AHU-08壓縮機運轉電流值測試"}, {"label": "送風風扇軸承檢查測試", "desc": "AHU-08送風風扇軸承檢查測試"}, {"label": "蒸發器洩水管檢查及清潔", "desc": "AHU-08蒸發器洩水管檢查及清潔"}, {"label": "設備外觀照", "desc": "AHU-08設備外觀照"}], "ACP-07": [{"label": "電壓檢查", "desc": "ACP-07電壓檢查"}, {"label": "各馬達電流檢查", "desc": "ACP-07各馬達電流檢查"}, {"label": "絕緣檢查", "desc": "ACP-07絕緣檢查"}, {"label": "啟動時間繼電時間檢查調整", "desc": "ACP-07啟動時間繼電時間檢查調整"}, {"label": "設備外觀照", "desc": "ACP-07設備外觀照"}], "AHU-09": [{"label": "濾網及蒸發器檢查及清潔", "desc": "AHU-09濾網及蒸發器檢查及清潔"}, {"label": "壓縮機運轉電流值測試", "desc": "AHU-09壓縮機運轉電流值測試"}, {"label": "送風風扇軸承檢查測試", "desc": "AHU-09送風風扇軸承檢查測試"}, {"label": "蒸發器洩水管檢查及清潔", "desc": "AHU-09蒸發器洩水管檢查及清潔"}, {"label": "設備外觀照", "desc": "AHU-09設備外觀照"}], "AHU-10": [{"label": "濾網及蒸發器檢查及清潔", "desc": "AHU-10濾網及蒸發器檢查及清潔"}, {"label": "壓縮機運轉電流值測試", "desc": "AHU-10壓縮機運轉電流值測試"}, {"label": "送風風扇軸承檢查測試", "desc": "AHU-10送風風扇軸承檢查測試"}, {"label": "蒸發器洩水管檢查及清潔", "desc": "AHU-10蒸發器洩水管檢查及清潔"}, {"label": "設備外觀照", "desc": "AHU-10設備外觀照"}], "ACP-09": [{"label": "電壓檢查", "desc": "ACP-09電壓檢查"}, {"label": "各馬達電流檢查", "desc": "ACP-09各馬達電流檢查"}, {"label": "絕緣檢查", "desc": "ACP-09絕緣檢查"}, {"label": "啟動時間繼電時間檢查調整", "desc": "ACP-09啟動時間繼電時間檢查調整"}, {"label": "設備外觀照", "desc": "ACP-09設備外觀照"}], "AHU-11": [{"label": "濾網及蒸發器檢查及清潔", "desc": "AHU-11濾網及蒸發器檢查及清潔"}, {"label": "壓縮機運轉電流值測試", "desc": "AHU-11壓縮機運轉電流值測試"}, {"label": "送風風扇軸承檢查測試", "desc": "AHU-11送風風扇軸承檢查測試"}, {"label": "蒸發器洩水管檢查及清潔", "desc": "AHU-11蒸發器洩水管檢查及清潔"}, {"label": "設備外觀照", "desc": "AHU-11設備外觀照"}], "AHU-12": [{"label": "濾網及蒸發器檢查及清潔", "desc": "AHU-12濾網及蒸發器檢查及清潔"}, {"label": "壓縮機運轉電流值測試", "desc": "AHU-12壓縮機運轉電流值測試"}, {"label": "送風風扇軸承檢查測試", "desc": "AHU-12送風風扇軸承檢查測試"}, {"label": "蒸發器洩水管檢查及清潔", "desc": "AHU-12蒸發器洩水管檢查及清潔"}, {"label": "設備外觀照", "desc": "AHU-12設備外觀照"}], "ACP-10": [{"label": "電壓檢查", "desc": "ACP-10電壓檢查"}, {"label": "各馬達電流檢查", "desc": "ACP-10各馬達電流檢查"}, {"label": "絕緣檢查", "desc": "ACP-10絕緣檢查"}, {"label": "啟動時間繼電時間檢查調整", "desc": "ACP-10啟動時間繼電時間檢查調整"}, {"label": "設備外觀照", "desc": "ACP-10設備外觀照"}]}, "林口機房": {"AHU-01": [{"label": "濾網及蒸發器檢查及清潔", "desc": "AHU-01濾網及蒸發器檢查及清潔"}, {"label": "壓縮機運轉電流值測試", "desc": "AHU-01壓縮機運轉電流值測試"}, {"label": "送風風扇軸檢查", "desc": "AHU-01送風風扇軸檢查"}, {"label": "蒸發器洩水管檢查及清潔", "desc": "AHU-01蒸發器洩水管檢查及清潔"}, {"label": "設備外觀照", "desc": "AHU-01設備外觀照"}], "AHU-02": [{"label": "濾網及蒸發器檢查及清潔", "desc": "AHU-02濾網及蒸發器檢查及清潔"}, {"label": "壓縮機運轉電流值測試", "desc": "AHU-02壓縮機運轉電流值測試"}, {"label": "送風風扇軸檢查", "desc": "AHU-02送風風扇軸檢查"}, {"label": "蒸發器洩水管檢查及清潔", "desc": "AHU-02蒸發器洩水管檢查及清潔"}, {"label": "設備外觀照", "desc": "AHU-02設備外觀照"}], "ACP-01": [{"label": "電壓檢查", "desc": "ACP-01電壓檢查"}, {"label": "各馬達電流檢查", "desc": "ACP-01各馬達電流檢查"}, {"label": "絕緣檢查", "desc": "ACP-01絕緣檢查"}, {"label": "啟動時間繼電時間檢查及調整", "desc": "ACP-01啟動時間繼電時間檢查及調整"}, {"label": "設備外觀照", "desc": "ACP-01設備外觀照"}]}, "中壢機房": {"AHU-01": [{"label": "濾網及蒸發器檢查及清潔", "desc": "AHU-01濾網及蒸發器檢查及清潔"}, {"label": "壓縮機運轉電流值測試", "desc": "AHU-01壓縮機運轉電流值測試"}, {"label": "送風風扇軸檢查", "desc": "AHU-01送風風扇軸檢查"}, {"label": "蒸發器洩水管檢查及清潔", "desc": "AHU-01蒸發器洩水管檢查及清潔"}, {"label": "設備外觀照", "desc": "AHU-01設備外觀照"}], "AHU-02": [{"label": "濾網及蒸發器檢查及清潔", "desc": "AHU-02濾網及蒸發器檢查及清潔"}, {"label": "壓縮機運轉電流值測試", "desc": "AHU-02壓縮機運轉電流值測試"}, {"label": "送風風扇軸檢查", "desc": "AHU-02送風風扇軸檢查"}, {"label": "蒸發器洩水管檢查及清潔", "desc": "AHU-02蒸發器洩水管檢查及清潔"}, {"label": "設備外觀照", "desc": "AHU-02設備外觀照"}], "ACP-01": [{"label": "電壓檢查", "desc": "ACP-01電壓檢查"}, {"label": "各馬達電流檢查", "desc": "ACP-01各馬達電流檢查"}, {"label": "絕緣檢查", "desc": "ACP-01絕緣檢查"}, {"label": "啟動時間繼電時間檢查及調整", "desc": "ACP-01啟動時間繼電時間檢查及調整"}, {"label": "設備外觀照", "desc": "ACP-01設備外觀照"}], "AHU-03": [{"label": "濾網及蒸發器檢查及清潔", "desc": "AHU-03濾網及蒸發器檢查及清潔"}, {"label": "壓縮機運轉電流值測試", "desc": "AHU-03壓縮機運轉電流值測試"}, {"label": "送風風扇軸檢查", "desc": "AHU-03送風風扇軸檢查"}, {"label": "蒸發器洩水管檢查及清潔", "desc": "AHU-03蒸發器洩水管檢查及清潔"}, {"label": "設備外觀照", "desc": "AHU-03設備外觀照"}], "AHU-04": [{"label": "濾網及蒸發器檢查及清潔", "desc": "AHU-04濾網及蒸發器檢查及清潔"}, {"label": "壓縮機運轉電流值測試", "desc": "AHU-04壓縮機運轉電流值測試"}, {"label": "送風風扇軸檢查", "desc": "AHU-04送風風扇軸檢查"}, {"label": "蒸發器洩水管檢查及清潔", "desc": "AHU-04蒸發器洩水管檢查及清潔"}, {"label": "設備外觀照", "desc": "AHU-04設備外觀照"}], "ACP-02": [{"label": "電壓檢查", "desc": "ACP-02電壓檢查"}, {"label": "各馬達電流檢查", "desc": "ACP-02各馬達電流檢查"}, {"label": "絕緣檢查", "desc": "ACP-02絕緣檢查"}, {"label": "啟動時間繼電時間檢查及調整", "desc": "ACP-02啟動時間繼電時間檢查及調整"}, {"label": "設備外觀照", "desc": "ACP-02設備外觀照"}]}, "楊梅機房": {"AHU-01": [{"label": "濾網及蒸發器檢查及清潔", "desc": "AHU-01濾網及蒸發器檢查及清潔"}, {"label": "壓縮機運轉電流值測試", "desc": "AHU-01壓縮機運轉電流值測試"}, {"label": "送風風扇軸檢查", "desc": "AHU-01送風風扇軸檢查"}, {"label": "蒸發器洩水管檢查及清潔", "desc": "AHU-01蒸發器洩水管檢查及清潔"}, {"label": "設備外觀照", "desc": "AHU-01設備外觀照"}], "AHU-02": [{"label": "濾網及蒸發器檢查及清潔", "desc": "AHU-02濾網及蒸發器檢查及清潔"}, {"label": "壓縮機運轉電流值測試", "desc": "AHU-02壓縮機運轉電流值測試"}, {"label": "送風風扇軸檢查", "desc": "AHU-02送風風扇軸檢查"}, {"label": "蒸發器洩水管檢查及清潔", "desc": "AHU-02蒸發器洩水管檢查及清潔"}, {"label": "設備外觀照", "desc": "AHU-02設備外觀照"}], "ACP-01": [{"label": "電壓檢查", "desc": "ACP-01電壓檢查"}, {"label": "各馬達電流檢查", "desc": "ACP-01各馬達電流檢查"}, {"label": "絕緣檢查", "desc": "ACP-01絕緣檢查"}, {"label": "啟動時間繼電時間檢查及調整", "desc": "ACP-01啟動時間繼電時間檢查及調整"}, {"label": "設備外觀照", "desc": "ACP-01設備外觀照"}]}, "湖口機房": {"AHU-01": [{"label": "濾網及蒸發器檢查及清潔", "desc": "AHU-01濾網及蒸發器檢查及清潔"}, {"label": "壓縮機運轉電流值測試", "desc": "AHU-01壓縮機運轉電流值測試"}, {"label": "送風風扇軸檢查", "desc": "AHU-01送風風扇軸檢查"}, {"label": "蒸發器洩水管檢查及清潔", "desc": "AHU-01蒸發器洩水管檢查及清潔"}, {"label": "設備外觀照", "desc": "AHU-01設備外觀照"}], "AHU-02": [{"label": "濾網及蒸發器檢查及清潔", "desc": "AHU-02濾網及蒸發器檢查及清潔"}, {"label": "壓縮機運轉電流值測試", "desc": "AHU-02壓縮機運轉電流值測試"}, {"label": "送風風扇軸檢查", "desc": "AHU-02送風風扇軸檢查"}, {"label": "蒸發器洩水管檢查及清潔", "desc": "AHU-02蒸發器洩水管檢查及清潔"}, {"label": "設備外觀照", "desc": "AHU-02設備外觀照"}], "ACP-01": [{"label": "電壓檢查", "desc": "ACP-01電壓檢查"}, {"label": "各馬達電流檢查", "desc": "ACP-01各馬達電流檢查"}, {"label": "絕緣檢查", "desc": "ACP-01絕緣檢查"}, {"label": "啟動時間繼電時間檢查及調整", "desc": "ACP-01啟動時間繼電時間檢查及調整"}, {"label": "設備外觀照", "desc": "ACP-01設備外觀照"}]}, "新竹機房": {"AHU-01": [{"label": "濾網及蒸發器檢查及清潔", "desc": "AHU-01濾網及蒸發器檢查及清潔"}, {"label": "壓縮機運轉電流值測試", "desc": "AHU-01壓縮機運轉電流值測試"}, {"label": "送風風扇軸檢查", "desc": "AHU-01送風風扇軸檢查"}, {"label": "蒸發器洩水管檢查及清潔", "desc": "AHU-01蒸發器洩水管檢查及清潔"}, {"label": "設備外觀照", "desc": "AHU-01設備外觀照"}], "AHU-02": [{"label": "濾網及蒸發器檢查及清潔", "desc": "AHU-02濾網及蒸發器檢查及清潔"}, {"label": "壓縮機運轉電流值測試", "desc": "AHU-02壓縮機運轉電流值測試"}, {"label": "送風風扇軸檢查", "desc": "AHU-02送風風扇軸檢查"}, {"label": "蒸發器洩水管檢查及清潔", "desc": "AHU-02蒸發器洩水管檢查及清潔"}, {"label": "設備外觀照", "desc": "AHU-02設備外觀照"}], "ACP-01": [{"label": "電壓檢查", "desc": "ACP-01電壓檢查"}, {"label": "各馬達電流檢查", "desc": "ACP-01各馬達電流檢查"}, {"label": "絕緣檢查", "desc": "ACP-01絕緣檢查"}, {"label": "啟動時間繼電時間檢查及調整", "desc": "ACP-01啟動時間繼電時間檢查及調整"}, {"label": "設備外觀照", "desc": "ACP-01設備外觀照"}]}, "未分類": {"ACP-01": [{"label": "啟動時間繼電時間檢查及調整", "desc": "ACP-01啟動時間繼電時間檢查及調整"}]}, "木柵機房": {"AHU-01": [{"label": "濾網及蒸發器檢查及清潔", "desc": "AHU-01濾網及蒸發器檢查及清潔"}, {"label": "壓縮機運轉電流值測試", "desc": "AHU-01壓縮機運轉電流值測試"}, {"label": "送風風扇軸檢查", "desc": "AHU-01送風風扇軸檢查"}, {"label": "蒸發器洩水管檢查及清潔", "desc": "AHU-01蒸發器洩水管檢查及清潔"}, {"label": "設備外觀照", "desc": "AHU-01設備外觀照"}], "AHU-02": [{"label": "濾網及蒸發器檢查及清潔", "desc": "AHU-02濾網及蒸發器檢查及清潔"}, {"label": "壓縮機運轉電流值測試", "desc": "AHU-02壓縮機運轉電流值測試"}, {"label": "送風風扇軸檢查", "desc": "AHU-02送風風扇軸檢查"}, {"label": "蒸發器洩水管檢查及清潔", "desc": "AHU-02蒸發器洩水管檢查及清潔"}, {"label": "設備外觀照", "desc": "AHU-02設備外觀照"}], "ACP-01": [{"label": "電壓檢查", "desc": "ACP-01電壓檢查"}, {"label": "各馬達電流檢查", "desc": "ACP-01各馬達電流檢查"}, {"label": "絕緣檢查", "desc": "ACP-01絕緣檢查"}, {"label": "啟動時間繼電時間檢查及調整", "desc": "ACP-01啟動時間繼電時間檢查及調整"}, {"label": "設備外觀照", "desc": "ACP-01設備外觀照"}]}, "中和機房": {"AHU-01": [{"label": "濾網及蒸發器檢查及清潔", "desc": "AHU-01濾網及蒸發器檢查及清潔"}, {"label": "壓縮機運轉電流值測試", "desc": "AHU-01壓縮機運轉電流值測試"}, {"label": "送風風扇軸檢查", "desc": "AHU-01送風風扇軸檢查"}, {"label": "蒸發器洩水管檢查及清潔", "desc": "AHU-01蒸發器洩水管檢查及清潔"}, {"label": "設備外觀照", "desc": "AHU-01設備外觀照"}], "AHU-02": [{"label": "濾網及蒸發器檢查及清潔", "desc": "AHU-02濾網及蒸發器檢查及清潔"}, {"label": "壓縮機運轉電流值測試", "desc": "AHU-02壓縮機運轉電流值測試"}, {"label": "送風風扇軸檢查", "desc": "AHU-02送風風扇軸檢查"}, {"label": "蒸發器洩水管檢查及清潔", "desc": "AHU-02蒸發器洩水管檢查及清潔"}, {"label": "設備外觀照", "desc": "AHU-02設備外觀照"}], "ACP-01": [{"label": "電壓檢查", "desc": "ACP-01電壓檢查"}, {"label": "各馬達電流檢查", "desc": "ACP-01各馬達電流檢查"}, {"label": "絕緣檢查", "desc": "ACP-01絕緣檢查"}, {"label": "啟動時間繼電時間檢查及調整", "desc": "ACP-01啟動時間繼電時間檢查及調整"}, {"label": "設備外觀照", "desc": "ACP-01設備外觀照"}]}, "樹林機房": {"AHU-01": [{"label": "濾網及蒸發器檢查及清潔", "desc": "AHU-01濾網及蒸發器檢查及清潔"}, {"label": "壓縮機運轉電流值測試", "desc": "AHU-01壓縮機運轉電流值測試"}, {"label": "送風風扇軸檢查", "desc": "AHU-01送風風扇軸檢查"}, {"label": "蒸發器洩水管檢查及清潔", "desc": "AHU-01蒸發器洩水管檢查及清潔"}, {"label": "設備外觀照", "desc": "AHU-01設備外觀照"}], "AHU-02": [{"label": "濾網及蒸發器檢查及清潔", "desc": "AHU-02濾網及蒸發器檢查及清潔"}, {"label": "壓縮機運轉電流值測試", "desc": "AHU-02壓縮機運轉電流值測試"}, {"label": "送風風扇軸檢查", "desc": "AHU-02送風風扇軸檢查"}, {"label": "蒸發器洩水管檢查及清潔", "desc": "AHU-02蒸發器洩水管檢查及清潔"}, {"label": "設備外觀照", "desc": "AHU-02設備外觀照"}], "ACP-01": [{"label": "電壓檢查", "desc": "ACP-01電壓檢查"}, {"label": "各馬達電流檢查", "desc": "ACP-01各馬達電流檢查"}, {"label": "絕緣檢查", "desc": "ACP-01絕緣檢查"}, {"label": "啟動時間繼電時間檢查及調整", "desc": "ACP-01啟動時間繼電時間檢查及調整"}, {"label": "設備外觀照", "desc": "ACP-01設備外觀照"}]}, "大溪機房": {"AHU-01": [{"label": "濾網及蒸發器檢查及清潔", "desc": "AHU-01濾網及蒸發器檢查及清潔"}, {"label": "壓縮機運轉電流值測試", "desc": "AHU-01壓縮機運轉電流值測試"}, {"label": "送風風扇軸檢查", "desc": "AHU-01送風風扇軸檢查"}, {"label": "蒸發器洩水管檢查及清潔", "desc": "AHU-01蒸發器洩水管檢查及清潔"}, {"label": "設備外觀照", "desc": "AHU-01設備外觀照"}], "AHU-02": [{"label": "濾網及蒸發器檢查及清潔", "desc": "AHU-02濾網及蒸發器檢查及清潔"}, {"label": "壓縮機運轉電流值測試", "desc": "AHU-02壓縮機運轉電流值測試"}, {"label": "送風風扇軸檢查", "desc": "AHU-02送風風扇軸檢查"}, {"label": "蒸發器洩水管檢查及清潔", "desc": "AHU-02蒸發器洩水管檢查及清潔"}, {"label": "設備外觀照", "desc": "AHU-02設備外觀照"}], "ACP-01": [{"label": "電壓檢查", "desc": "ACP-01電壓檢查"}, {"label": "各馬達電流檢查", "desc": "ACP-01各馬達電流檢查"}, {"label": "絕緣檢查", "desc": "ACP-01絕緣檢查"}, {"label": "啟動時間繼電時間檢查及調整", "desc": "ACP-01啟動時間繼電時間檢查及調整"}, {"label": "設備外觀照", "desc": "ACP-01設備外觀照"}]}, "龍潭機房": {"AHU-01": [{"label": "濾網及蒸發器檢查及清潔", "desc": "AHU-01濾網及蒸發器檢查及清潔"}, {"label": "壓縮機運轉電流值測試", "desc": "AHU-01壓縮機運轉電流值測試"}, {"label": "送風風扇軸檢查", "desc": "AHU-01送風風扇軸檢查"}, {"label": "蒸發器洩水管檢查及清潔", "desc": "AHU-01蒸發器洩水管檢查及清潔"}, {"label": "設備外觀照", "desc": "AHU-01設備外觀照"}], "AHU-02": [{"label": "濾網及蒸發器檢查及清潔", "desc": "AHU-02濾網及蒸發器檢查及清潔"}, {"label": "壓縮機運轉電流值測試", "desc": "AHU-02壓縮機運轉電流值測試"}, {"label": "送風風扇軸檢查", "desc": "AHU-02送風風扇軸檢查"}, {"label": "蒸發器洩水管檢查及清潔", "desc": "AHU-02蒸發器洩水管檢查及清潔"}, {"label": "設備外觀照", "desc": "AHU-02設備外觀照"}], "ACP-01": [{"label": "電壓檢查", "desc": "ACP-01電壓檢查"}, {"label": "各馬達電流檢查", "desc": "ACP-01各馬達電流檢查"}, {"label": "絕緣檢查", "desc": "ACP-01絕緣檢查"}, {"label": "啟動時間繼電時間檢查及調整", "desc": "ACP-01啟動時間繼電時間檢查及調整"}, {"label": "設備外觀照", "desc": "ACP-01設備外觀照"}]}, "關西服務區": {"ACU-01": [{"label": "電源線,插頭及插座損壞或鬆動", "desc": "ACU-01電源線,插頭及插座損壞或鬆動"}, {"label": "空氣過濾器清潔", "desc": "ACU-01空氣過濾器清潔"}, {"label": "壓縮機啟動,運轉電流值檢查", "desc": "ACU-01壓縮機啟動,運轉電流值檢查"}, {"label": "運轉風扇馬達軸承檢查", "desc": "ACU-01運轉風扇馬達軸承檢查"}, {"label": "運轉時異聲及異常震動檢查", "desc": "ACU-01運轉時異聲及異常震動檢查"}, {"label": "排水管清潔", "desc": "ACU-01排水管清潔"}, {"label": "冷氣出風口溫度量測", "desc": "ACU-01冷氣出風口溫度量測"}, {"label": "設備外觀照", "desc": "ACU-01設備外觀照"}], "ACU-02": [{"label": "電源線,插頭及插座損壞或鬆動", "desc": "ACU-02電源線,插頭及插座損壞或鬆動"}, {"label": "空氣過濾器清潔", "desc": "ACU-02空氣過濾器清潔"}, {"label": "壓縮機啟動,運轉電流值檢查", "desc": "ACU-02壓縮機啟動,運轉電流值檢查"}, {"label": "運轉風扇馬達軸承檢查", "desc": "ACU-02運轉風扇馬達軸承檢查"}, {"label": "運轉時異聲及異常震動檢查", "desc": "ACU-02運轉時異聲及異常震動檢查"}, {"label": "排水管清潔", "desc": "ACU-02排水管清潔"}, {"label": "冷氣出風口溫度量測", "desc": "ACU-02冷氣出風口溫度量測"}, {"label": "設備外觀照", "desc": "ACU-02設備外觀照"}], "ACP-01": [{"label": "電壓檢查", "desc": "ACP-01電壓檢查"}, {"label": "各馬達電流檢查", "desc": "ACP-01各馬達電流檢查"}, {"label": "絕緣檢查檢查", "desc": "ACP-01絕緣檢查檢查"}, {"label": "啟動時間繼電時間檢查及調整", "desc": "ACP-01啟動時間繼電時間檢查及調整"}, {"label": "設備外觀照", "desc": "ACP-01設備外觀照"}]}, "關西機房": {"ACU-01": [{"label": "電源線,插頭及插座損壞或鬆動", "desc": "ACU-01電源線,插頭及插座損壞或鬆動"}, {"label": "壓縮機啟動,運轉電流值檢查", "desc": "ACU-01壓縮機啟動,運轉電流值檢查"}, {"label": "運轉風扇馬達軸承檢查", "desc": "ACU-01運轉風扇馬達軸承檢查"}, {"label": "排水管清潔", "desc": "ACU-01排水管清潔"}, {"label": "運轉時異聲及異常震動檢查", "desc": "ACU-01運轉時異聲及異常震動檢查"}, {"label": "冷氣出風口溫度量測", "desc": "ACU-01冷氣出風口溫度量測"}, {"label": "設備外觀照", "desc": "ACU-01設備外觀照"}], "ACU-02": [{"label": "電源線,插頭及插座損壞或鬆動", "desc": "ACU-02電源線,插頭及插座損壞或鬆動"}, {"label": "壓縮機啟動,運轉電流值檢查", "desc": "ACU-02壓縮機啟動,運轉電流值檢查"}, {"label": "運轉風扇馬達軸承檢查", "desc": "ACU-02運轉風扇馬達軸承檢查"}, {"label": "排水管清潔", "desc": ":ACU-02排水管清潔"}, {"label": "運轉時異聲及異常震動檢查", "desc": "ACU-02運轉時異聲及異常震動檢查"}, {"label": "冷氣出風口溫度量測", "desc": "ACU-02冷氣出風口溫度量測"}, {"label": "設備外觀照", "desc": "ACU-02設備外觀照"}], "ACP-01": [{"label": "電壓檢查", "desc": "ACP-01電壓檢查"}, {"label": "各馬達電流檢查", "desc": "ACP-01各馬達電流檢查"}, {"label": "絕緣檢查", "desc": "ACP-01絕緣檢查"}, {"label": "啟動時間繼電時間檢查及調整", "desc": "ACP-01啟動時間繼電時間檢查及調整"}, {"label": "設備外觀照", "desc": "ACP-01設備外觀照"}]}, "竹林機房": {"ACU-01": [{"label": "電源線,插頭及插座損壞或鬆動", "desc": "ACU-01電源線,插頭及插座損壞或鬆動"}, {"label": "空氣過濾器清潔", "desc": "ACU-01空氣過濾器清潔"}, {"label": "壓縮機啟動,運轉電流值檢查", "desc": "ACU-01壓縮機啟動,運轉電流值檢查"}, {"label": "運轉風扇馬達軸承檢查", "desc": "ACU-01運轉風扇馬達軸承檢查"}, {"label": "排水管清潔", "desc": "ACU-01排水管清潔"}, {"label": "運轉時異聲及異常震動檢查", "desc": "ACU-01運轉時異聲及異常震動檢查"}, {"label": "冷氣出風口溫度量測", "desc": "ACU-01冷氣出風口溫度量測"}, {"label": "設備外觀照", "desc": "ACU-01設備外觀照"}], "ACU-02": [{"label": "電源線,插頭及插座損壞或鬆動", "desc": "ACU-02電源線,插頭及插座損壞或鬆動"}, {"label": "空氣過濾器清潔", "desc": "ACU-02空氣過濾器清潔"}, {"label": "壓縮機啟動,運轉電流值檢查", "desc": "ACU-02壓縮機啟動,運轉電流值檢查"}, {"label": "運轉風扇馬達軸承檢查", "desc": "ACU-02運轉風扇馬達軸承檢查"}, {"label": "排水管清潔", "desc": ":ACU-02排水管清潔"}, {"label": "運轉時異聲及異常震動檢查", "desc": "ACU-02運轉時異聲及異常震動檢查"}, {"label": "冷氣出風口溫度量測", "desc": "ACU-02冷氣出風口溫度量測"}, {"label": "設備外觀照", "desc": "ACU-02設備外觀照"}], "ACP-01": [{"label": "電壓檢查", "desc": "ACP-01電壓檢查"}, {"label": "各馬達電流檢查", "desc": "ACP-01各馬達電流檢查"}, {"label": "絕緣檢查", "desc": "ACP-01絕緣檢查"}, {"label": "啟動時間繼電時間檢查及調整", "desc": "ACP-01啟動時間繼電時間檢查及調整"}, {"label": "設備外觀照", "desc": "ACP-01設備外觀照"}]}, "內湖北分局": {"ACU-01": [{"label": "電源線,插頭及插座損壞或鬆動", "desc": "ACU-01電源線,插頭及插座損壞或鬆動"}, {"label": "空氣過濾器清潔", "desc": "ACU-01空氣過濾器清潔"}, {"label": "壓縮機啟動,運轉電流值檢查", "desc": "ACU-01壓縮機啟動,運轉電流值檢查"}, {"label": "運轉風扇馬達軸承檢查", "desc": "ACU-01運轉風扇馬達軸承檢查"}, {"label": "排水管清潔", "desc": "ACU-01排水管清潔"}, {"label": "運轉時異聲及異常震動檢查", "desc": "ACU-01運轉時異聲及異常震動檢查"}, {"label": "冷氣出風口溫度量測", "desc": "ACU-01冷氣出風口溫度量測"}, {"label": "設備外觀照", "desc": "ACU-01設備外觀照"}], "ACU-02": [{"label": "電源線,插頭及插座損壞或鬆動", "desc": "ACU-02電源線,插頭及插座損壞或鬆動"}, {"label": "空氣過濾器清潔", "desc": "ACU-02空氣過濾器清潔"}, {"label": "壓縮機啟動,運轉電流值檢查", "desc": "ACU-02壓縮機啟動,運轉電流值檢查"}, {"label": "運轉風扇馬達軸承檢查", "desc": "ACU-02運轉風扇馬達軸承檢查"}, {"label": "排水管清潔", "desc": "ACU-02排水管清潔"}, {"label": "運轉時異聲及異常震動檢查", "desc": "ACU-02運轉時異聲及異常震動檢查"}, {"label": "冷氣出風口溫度量測", "desc": "ACU-02冷氣出風口溫度量測"}, {"label": "設備外觀照", "desc": "ACU-02設備外觀照"}], "ACU-03": [{"label": "電源線,插頭及插座損壞或鬆動", "desc": "ACU-03電源線,插頭及插座損壞或鬆動"}, {"label": "空氣過濾器清潔", "desc": "ACU-03空氣過濾器清潔"}, {"label": "壓縮機啟動,運轉電流值檢查", "desc": "ACU-03壓縮機啟動,運轉電流值檢查"}, {"label": "運轉風扇馬達軸承檢查", "desc": "ACU-03運轉風扇馬達軸承檢查"}, {"label": "排水管清潔", "desc": "ACU-03排水管清潔"}, {"label": "運轉時異聲及異常震動檢查", "desc": "ACU-03運轉時異聲及異常震動檢查"}, {"label": "冷氣出風口溫度量測", "desc": "ACU-03冷氣出風口溫度量測"}, {"label": "設備外觀照", "desc": "ACU-03設備外觀照"}], "ACU-04": [{"label": "電源線,插頭及插座損壞或鬆動", "desc": "ACU-04電源線,插頭及插座損壞或鬆動"}, {"label": "空氣過濾器清潔", "desc": "ACU-04空氣過濾器清潔"}, {"label": "壓縮機啟動,運轉電流值檢查", "desc": "ACU-04壓縮機啟動,運轉電流值檢查"}, {"label": "運轉風扇馬達軸承檢查", "desc": "ACU-04運轉風扇馬達軸承檢查"}, {"label": "排水管清潔", "desc": "ACU-04排水管清潔"}, {"label": "運轉時異聲及異常震動檢查", "desc": "ACU-04運轉時異聲及異常震動檢查"}, {"label": "冷氣出風口溫度量測", "desc": "ACU-04冷氣出風口溫度量測"}, {"label": "設備外觀照", "desc": "ACU-04設備外觀照"}], "ACU-05": [{"label": "電源線,插頭及插座損壞或鬆動", "desc": "ACU-05電源線,插頭及插座損壞或鬆動"}, {"label": "空氣過濾器清潔", "desc": "ACU-05空氣過濾器清潔"}, {"label": "壓縮機啟動,運轉電流值檢查", "desc": "ACU-05壓縮機啟動,運轉電流值檢查"}, {"label": "運轉風扇馬達軸承檢查", "desc": "ACU-05運轉風扇馬達軸承檢查"}, {"label": "排水管清潔", "desc": "ACU-05排水管清潔"}, {"label": "運轉時異聲及異常震動檢查", "desc": "ACU-05運轉時異聲及異常震動檢查"}, {"label": "冷氣出風口溫度量測", "desc": "ACU-05冷氣出風口溫度量測"}, {"label": "設備外觀照", "desc": "ACU-05設備外觀照"}], "ACU-06": [{"label": "電源線,插頭及插座損壞或鬆動", "desc": "ACU-06電源線,插頭及插座損壞或鬆動"}, {"label": "空氣過濾器清潔", "desc": "ACU-06空氣過濾器清潔"}, {"label": "壓縮機啟動,運轉電流值檢查", "desc": "ACU-06壓縮機啟動,運轉電流值檢查"}, {"label": "運轉風扇馬達軸承檢查", "desc": "ACU-06運轉風扇馬達軸承檢查"}, {"label": "排水管清潔", "desc": "ACU-06排水管清潔"}, {"label": "運轉時異聲及異常震動檢查", "desc": "ACU-06運轉時異聲及異常震動檢查"}, {"label": "冷氣出風口溫度量測", "desc": "ACU-06冷氣出風口溫度量測"}, {"label": "設備外觀照", "desc": "ACU-06設備外觀照"}], "ACP-01": [{"label": "電壓檢查", "desc": "ACP-01電壓檢查"}, {"label": "各馬達電流檢查", "desc": "ACP-01各馬達電流檢查"}, {"label": "絕緣檢查", "desc": "ACP-01絕緣檢查"}, {"label": "啟動時間繼電時間檢查及調整", "desc": "ACP-01啟動時間繼電時間檢查及調整"}, {"label": "設備外觀照", "desc": "ACP-01設備外觀照"}], "ACP-02": [{"label": "電壓檢查", "desc": "ACP-02電壓檢查"}, {"label": "各馬達電流檢查", "desc": "ACP-02各馬達電流檢查"}, {"label": "絶緣檢查", "desc": "ACP-02絶緣檢查"}, {"label": "啟動時間繼電時間檢查及調整", "desc": "ACP-02啟動時間繼電時間檢查及調整"}, {"label": "設備外觀照", "desc": "ACP-02設備外觀照"}]}};
