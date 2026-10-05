# CHANGELOG

> 依 CLAUDE.md【MUST LOG-1】：每次因回饋或需求調整而改動程式，記一筆日期、改了什麼、為什麼、影響範圍。

---

## 2026-09-15 — `uploadPhoto()` 加入階段計時（v4.7，純觀測）

**為什麼**
業主回報「上傳速度比較慢，約五秒」。原本 `uploadPhoto()` 只有 `t0` 總耗時（第 68 行），
看得到「5 秒」卻看不出「5 秒花在哪一段」，無法定位瓶頸。
在量到數字之前做任何優化都是猜測，可能改到不是瓶頸的地方。

**改了什麼**
`gas-maintenance-photo/程式碼.gs` — 25 行純新增，0 行刪除、0 行修改：

- 函式開頭新增 `tPrev` / `lap` / `mark()` 三個區域變數（名稱先 grep 確認無衝突）
- 在 10 個階段後各插入一行 `mark(...)`：
  解碼驗證、開啟試算表、取得月份分頁、Z欄定位、Z欄自我修復（僅觸發時）、
  槽位解析、Drive備份、欄寬列高、清舊浮動圖、照片入格、日期回寫
- 成功路徑結尾多印一行 `[uploadPhoto] 階段耗時(ms) ...`
- 記錄的是**每段 delta**（不是累計），最大的那個數字就是瓶頸

**影響範圍**
僅 `uploadPhoto()` 內部，且只有 `console.log`。
沒有改動任何判斷、順序、API 呼叫或回傳值；沒有動 `MENU_TREE`；沒有碰 Google Sheets 任何格式。
完全可逆（把新增的行刪掉即還原）。

**驗證輸出**

```
✅ node --check PASS
✅ MENU_TREE JSON.parse PASS ｜ 機房 17 ｜ 選單項目 666   （與改前基準一致）
✅ 17/17 機房 JSON 與 git show HEAD 完全一致，MENU_TREE 零波及
✅ git diff --stat：1 file changed, 25 insertions(+)      （無刪除、無修改）
```

**尚未驗證**
未端到端驗證 —— 需要業主把程式貼進 GAS 網頁 App 專案、部署新版本，
並實際用手機上傳 2～3 張，才會產生階段耗時數據。

**下一步**
依實測數字決定優化哪一段。四個待驗嫌疑犯（依可疑度）：
`sh.getImages()` 全表掃描 → `backupAndGetId_()` 的 6 次 Drive 往返 →
`Z:Z` TextFinder 整欄掃描 → `openById` 開 80 MB 試算表。

---

## 2026-09-15 — CLAUDE.md 補上「系統架構」一節

**為什麼**
業主希望不必每次都把整份 `程式碼.gs`（約 1000 行）貼給 AI。
原本的 `CLAUDE.md` 寫的是**工作守則與踩過的坑**（第 10 節，寫得很完整），
但**沒有一段在講「資料怎麼流」**——新 session 進來只知道「哪裡不能碰」，
不知道「這東西長什麼樣」，仍得從頭讀 1000 行才敢下手。

**改了什麼**
`CLAUDE.md` — 83 行純新增、0 行刪除。位置在「專案速覽」與「工作守則」之間，
新增一節「系統架構（高公局拍照歸檔）」，內含：

- 三個角色：手機前端 / GAS 後端 / 兩個資料落點
- 一張照片的完整路徑：前端 5 步 + 後端 `uploadPhoto()` 12 步逐段表
- 定位系統的推導鏈（F 欄文字 → Z 欄定位鍵 → 照片格），並點明這就是 P-4 的來由
- 三個資料落點各自壞掉會怎樣
- 19 個函式分成「自動跑的熱路徑」與「手動執行的工具」
- 效能特性：5 秒的成因、務實下限、四個待驗嫌疑犯、如何用階段計時量測

**影響範圍**
純文件。不動任何程式。既有 39 條 MUST/SHOULD 條文一條未改、一條未少。

**驗證輸出**

```
✅ git diff --numstat CLAUDE.md：新增 83 行，刪除 0 行
✅ 既有條文數 39 → 39，逐條比對無遺失
✅ node --check PASS（程式碼未再變動，重跑確認）
```

**尚未驗證**
這節的實用性要等下一個 session 實際使用才知道。若發現仍需補充，直接增修本節。

---

## 2026-10-05 — 前端改為「清單式批次上傳」＋ 後端新增唯讀 `getFilledSlots`（v12）

**為什麼**
現場流程是「選項目 → 拍 → 傳 → 站著等約 5 秒 → 再選下一項」，一個設備 5～8 項就要重複 5～8 輪。
改成選好設備後直接列出全部檢查項目、逐項拍照、最後按一次依序上傳。
**批次上傳不會減少網路傳輸量**；省下的是師傅站著等的時間（連拍完按一次就能走開）。

**改了什麼**
- `gas-maintenance-photo/index.html`（主要）
  - 拿掉「檢查項目」下拉與單張預覽卡；設備選好後依 `MENU_TREE` 展開該設備全部項目成清單，每項有「拍照/選圖」、縮圖、狀態圓點
  - 項目 >12 才分組並預設收合（只有泰控「其他」45 項：依「之N」「組號」分成 12 組）；其餘 105 個設備平鋪
  - 照片存 **Blob ＋ 160px 縮圖**，不再存整張 dataURL；壓縮改為一張接一張排隊（共用 `#cv`）
  - 「全部上傳」依序呼叫 `uploadPhoto`（不並發）；失敗的留在清單可重傳，全部成功才重置；連續 3 次連線失敗暫停
  - 選好設備時背景查 `getFilledSlots`，標示「雲端已有」；多槽位全滿才鎖按鈕，單槽位只提醒「再傳會覆蓋」
  - 底部黏底操作列顯示進度（`正在上傳 i/N…`＋進度條）；有未上傳照片或上傳中，機房/設備下拉鎖住；「清除全部」兩段式確認（HtmlService 沙箱不保證能用 `confirm()`）
  - 版本標籤 v11 → v12（讓現場看得出新版有沒有部署成功）
  - 圖示沿用既有 SVG 風格，未使用 emoji
- `gas-maintenance-photo/程式碼.gs`：**僅新增** `getFilledSlots`、`locateSlotRows_`（94 行，0 刪除）
- `CLAUDE.md`：架構節同步為清單式流程；新增 P-14～P-17

**刻意沒動**
- `uploadPhoto()` 與其餘 18 個既有後端函式：逐位元組相同（SHA256 `b2d1e8ae…` 與 HEAD 一致）
- `MENU_TREE`、Z 欄定位（P-6 `findAll`、P-7 合併錨點）、Google Sheets 任何格式／欄位／排版（P-1）
- Contain＋白底（P-12）、`MAX_SIDE`／`JPEG_Q`／`BOX_RATIO` 常數
- `parseExifDate`、`sortEquipKeys`、`nowYm`、`ymFromYmd`、`toast` 等 8 個前端函式：逐位元組相同
- payload 形狀與舊版單張上傳完全相同（`room/equip/item/ym/shotYmd/dataUrl`），`item` 一律是 `desc`

**影響範圍**
前端整個互動流程。後端寫入路徑零改動；新增的 `getFilledSlots` 只讀不寫、不加鎖。

**驗證輸出**
```
✅ node --check：程式碼.gs、index.html 內嵌 script 皆 PASS
✅ MENU_TREE：機房 17｜項目 666｜與 HEAD 逐機房比對差異 0 個
✅ uploadPhoto SHA256：HEAD b2d1e8ae4f8d／現在 b2d1e8ae4f8d（相同）
✅ 既有 19 個後端函式逐位元組相同；前端 8 個沿用函式逐位元組相同
✅ git diff（程式碼.gs）：+94／-0，單一 hunk 位於 getMenu 之後
✅ 前端離線實測（headless Chromium＋假的 google.script.run＋真的 MENU_TREE）：72/72 通過
   · 105/106 設備平鋪、泰控「其他」45 項分 12 組（7,7,5,2,3×8）
   · 依序上傳：最大並發 1、順序＝清單順序、payload 欄位同舊版、dataUrl 通過 validateJpegBytes_ 同等檢查
   · desc≠label 的三項（-CHU-03…、:CTR-02…、冷卻水加藥設備檢查檢查之2）送出的都是 desc
   · Contain：橫拍/直拍的紅(左上)、藍(右上)角落像素都在，上下或左右為純白
   · 部分失敗：失敗項保留、成功項不重送；斷網 3 次連續失敗後暫停，其餘維持 ready
   · 45 張處理後，強制 GC 後 JS heap 僅 +0.10 MB（Blob 合計 12.9 MB 不在 JS heap）
✅ 後端差異測試（模擬 GAS 服務，跑真正的原始碼）：18/18 通過
   · 11 種槽位情境下，前端依 getFilledSlots 的判斷（擋／覆蓋／填下一格）與 uploadPhoto 實際行為一致
   · API 呼叫次數與項目數無關：5 項與 45 項皆 getRange×3、getMergedRanges×1、TextFinder×0
```

**尚未端到端驗證**（需要業主實機）
- 沒有在真的 GAS／Google Sheets／手機上跑過；以上都是離線模擬。模擬的是 `uploadPhoto` 用到的 API 語意，不是真的 Sheets
- `getFilledSlots` 在 80 MB 母版上的實際耗時未知（預期 1～2 秒是估計值）；上線後看執行記錄的 `[getFilledSlots] … 耗時 N ms`
- 手機上「拍照/選圖」單一按鈕（無 `capture`）的實際選單長相，依機型而異
- `navigator.wakeLock` 在 HtmlService iframe 內可能被拒絕（已 try/catch，失敗只是少一層保護）
- 頁面被重新整理或手機回收頁面時，尚未上傳的照片會遺失（未做 IndexedDB 暫存）

**部署**
`index.html` 與 `程式碼.gs` **兩個都要**貼進「高公局空調報表自動化」獨立網頁 App 專案（有 `doGet` 的那個），再「部署新版本」。
只貼 `index.html` 也能用（查雲端狀態會顯示「查詢失敗」但不影響拍照上傳）。

**回滾／備份（貼 v12 之前的上線版）**
- 備份位置：commit `804ccfe1181f89b796b4e84124d66a109957ec97`（＝ `main` 目前的 `804ccfe`）內的 `gas-maintenance-photo/程式碼.gs`（v4.6 後端）與 `index.html`（v11 單張上傳）。
  瀏覽：`https://github.com/Fakepan/yofa-quote/tree/804ccfe1181f89b796b4e84124d66a109957ec97/gas-maintenance-photo`
  以 commit 編號固定、不另設標籤；只要不改寫歷史，這個編號永遠找得回來。即使日後 `main` 合併了 v12，舊版仍在歷史裡。
- 與業主 2026-10-05 貼出的「原本線上版」比對一致：程式碼.gs 21 行、index.html 11 行原文特徵全部命中，且不含任何 v12 內容。**這是特徵比對，不是逐位元組比對**；線上 GAS 專案本身讀不到（Drive 連接器沒有索引到該專案）。
- 最快的回滾：GAS「部署 → 管理部署作業 → ✎ → 版本」選回舊版本 → 部署（網址不變、不用貼程式碼；編輯器裡的程式碼仍是新版，只是對外服務的是舊版本）。**貼 v12 之前先記下現在的版本號**。
