# CLAUDE.md — yofa-quote

> 本檔為專案的長期記憶與工作守則。新的 session 進來請先讀本檔。

## 專案速覽

這個 repo 裝了兩個獨立系統，都是佑發空調（HVAC 工程）自用：

| 系統 | 位置 | 說明 |
|---|---|---|
| **高公局空調保養報表自動化**（主要） | `gas-maintenance-photo/` | 師傅在現場用手機拍照，照片自動歸檔進 Google Sheets 的保養檢查報表。取代原本「拍照→建檔→下載→分類→修圖→改格式」6 天的人工流程 |
| 報價／專案管理 | 根目錄 `index.html` | 451 行純前端單檔，Tailwind CDN + SheetJS，資料存 localStorage。無後端 |

- 目前階段：GAS 系統**已上線試運行**，師傅尚未正式全面啟用
- 相關文件：無獨立 PRD，需求散落在對話與本檔

**⚠️ 重要觀念：GitHub 不是程式跑的地方。** 這個 repo 只是備份與歷史。實際執行的是兩個 Google Apps Script 專案，改完必須**貼進 GAS 並「部署新版本」**才會生效。推 GitHub 對現場使用者沒有任何影響。

---

# 系統架構（高公局拍照歸檔）

> 這一節是**地圖**：講「資料怎麼流」。第 10 節講的是「哪裡不能碰」，兩者搭配看。
> 檔案本身在 `gas-maintenance-photo/`，需要細節直接讀檔，不要靠這節的摘要下判斷。

## 三個角色

| 角色 | 檔案 | 做什麼 |
|---|---|---|
| 手機前端 | `index.html`（約 900 行） | 師傅操作：三層選單 → 拍照 → 壓縮 → 送出 |
| GAS 後端 | `程式碼.gs`（約 1000 行，19 個函式） | `doGet` 送出網頁；`uploadPhoto` 處理每一張上傳 |
| 資料落點 | Google 雲端 | 母版試算表 + Drive 備份資料夾 |

## 一張照片的完整路徑

**前端**（`index.html`）
1. 三層連動選單：機房 → 設備 → 項目（資料來自後端 `getMenu()` 回傳的 `MENU_TREE`）
2. 拍照或從相簿選（時間相機 App 拍的要走「相簿」才保得住浮水印）
3. `parseExifDate()` 讀 EXIF tag `0x9003` 取拍攝日 → 決定月份分頁與施工/完工日期
4. `compress()` 用 Canvas **Contain 等比縮放 + 白色補邊**到 `MAX_SIDE=1024`、比例 404:330、`JPEG_Q=0.75`
5. `toDataURL` 轉 base64 → `google.script.run.uploadPhoto(payload)`

**後端**（`程式碼.gs` 的 `uploadPhoto()`，這是唯一的熱路徑）

| # | 做什麼 | 備註 |
|---|---|---|
| 0 | 參數防呆 → base64 解碼 → `validateJpegBytes_` 三重完整性檢查 | 擋「傳一半」的破圖 |
| 1 | 基準日推導：EXIF 拍攝日優先，讀不到退回上傳日 | 決定分頁與日期，跨年不撕裂 |
| 2 | `SpreadsheetApp.openById(MASTER_SHEET_ID)` | 母版約 80 MB（照片嵌在格子裡） |
| 3 | `getOrCreateMonthSheet_()` 取當月分頁，不存在就從模板複製 | 含 LockService 防併發 |
| 4 | `Z:Z` `createTextFinder().matchEntireCell(true).findAll()` 定位 | **整欄掃描** |
| 5 | 同鍵多筆 → 自我修復：重建該分頁 Z 欄再查一次（6 小時一次） | 由 CacheService 的 `zfix_` 旗標節流 |
| 6 | 逐個 match 算照片格起始列（`+1` 再用合併儲存格錨點校正），挑第一個空槽位 | 空槽位判定見 P-8 |
| 7 | `backupAndGetId_()` 存 Drive 並設公開可讀 | **約 6 次 Drive 往返** |
| 8 | `setColumnWidths` / `setRowHeights` 強制 404×330 | 批次版，各 1 次 |
| 9 | `sh.getImages()` 清舊浮動圖片 | **全表掃描**，僅為相容舊版 `insertImage` |
| 10 | `newCellImage().setSourceUrl(lh3 網址)` → `setValue`；失敗退回 `=IMAGE()` | 圖片成為儲存格「內容」 |
| 11 | 掃本槽位 10 列的 F 欄，回寫施工/完工日期（民國年） | 1 次 `getValues` + 1 次 `setValues` |

## 定位系統怎麼運作（為什麼 Z 欄是唯一真相）

```
母版 F 欄文字（照片內容說明:xxx）
   └─ rebuildZColumnFor_ 依規則推導出定位鍵「機房__設備+項目說明」
        └─ 寫進該分頁 Z 欄（第 26 欄），位置 = 照片格起始列
             └─ uploadPhoto 用 TextFinder 比對整格相等找到它
                  └─ 那一列（經合併儲存格校正）就是要塞照片的格子
```

**所以 `MENU_TREE` 的 `desc` 必須與母版 F 欄一字不差** —— 這就是 P-4 那條規則的來由，也是為什麼看起來像錯字的文字不能「順手修正」。

## 三個資料落點

| 落點 | 內容 | 壞掉會怎樣 |
|---|---|---|
| 母版試算表「模板」分頁 | 版面、F 欄說明文字、Z 欄定位鍵 | 新月份分頁都從這裡複製，改壞影響未來每一個月 |
| 月份分頁（`YYYYMM`，純 6 位數） | 該月實際報表與照片 | 是**複製當下的快照**，模板後來新增的項目救不回來（P-10） |
| Drive 備份資料夾 | 原圖，依 `年月/機房/` 分層 | 設為公開可讀，CellImage 才顯示得出來；權限改掉照片全變破圖 |

## 19 個函式分兩類

**自動跑的（師傅按上傳就會走到）**：`doGet`、`getMenu`、`uploadPhoto`、`getOrCreateMonthSheet_`、`backupAndGetId_`、`getOrCreate_`、`rebuildZColumnFor_`、`validateJpegBytes_`

**手動執行的工具（在 GAS 編輯器選函式跑，不需重新部署）**：
`buildZColumn`、`rebuildZColumnEverywhere`、`newMonthCopy`、`clearTemplatePhotos`、`normalizeTemplateLayout`(+Part1/Part2/`normalizeLayout_`)、`auditMenuVsMaster`、`auditAllMonthTabs`、`diagTest`

> 稽核用的三支：`auditMenuVsMaster`（選單 ↔ 模板）、`auditAllMonthTabs`（唯讀、逐月份分頁健檢）、`diagTest`（能不能開表）。

## 效能特性

- **一次上傳約 5 秒**（業主 2026-09 回報）。**這不是 bug，是 GAS + Sheets API + Drive API 一次往返的正常成本。**
- **下限**：這類流程壓到 2～3 秒是務實目標；要更快就得換架構（照片先落 Drive、歸檔改背景批次），代價是師傅拿不到「已歸檔」的即時確認。
- **已知成本較高的呼叫**（尚未實測，依可疑度排序）：
  1. 步驟 9 `sh.getImages()` —— 每次上傳都做全表掃描，但現行流程已不再產生浮動圖片，多半白掃
  2. 步驟 7 `backupAndGetId_()` —— 每張都重新解析年月/機房資料夾，連拍十張就解析十遍
  3. 步驟 4 `Z:Z` TextFinder —— 整欄掃描
  4. 步驟 2 `openById` —— 80 MB 母版
- **量測方式**：`uploadPhoto()` 已內建階段計時（v4.7），執行記錄會印
  `[uploadPhoto] 階段耗時(ms) 解碼驗證:xx｜開啟試算表:xx｜…`，**最大的那一項就是瓶頸**。
  優化前先看這行數字，不要憑猜測改。

---

# 工作守則

> 【MUST】違反即停下回報；【SHOULD】預設遵守，偏離須說明原因。

### 0. 事實來源

- 【MUST FACT-1】技術棧以實際 repo 為準。**本 repo 沒有 `package.json`**（`.gitignore` 主動忽略它）、沒有建置流程、沒有套件管理。不要引入任何框架或建置工具。
- 【MUST FACT-2】不臆造 API／函式／常數／欄位／檔案路徑。引用前先 grep 證實存在，無法證實就停下問。
- 技術棧：Google Apps Script（V8）＋ Google Sheets／Drive API；前端為單檔 HTML＋原生 JS（無框架、無 bundler）。

### 1. 程式碼原則

- 【SHOULD CODE-1】用能解決當前需求的最少程式碼，不為「將來可能用到」預先抽象。
- 【SHOULD CODE-2】函式只做一件事；超過約 50 行或承擔一個以上職責就拆。
- 【SHOULD CODE-3】只在「為什麼這樣做」不顯而易見時寫註解（解釋取捨、踩過的坑）。
- 【MUST CODE-4】外科手術式修改：只改非動不可的地方，不順手重構鄰近程式碼。看到無關死碼只回報、不擅自刪。

### 2. 何時必須停下來問

- 【MUST ASK-1】以下先停、列選項給業主裁決：需求有多種合理解讀／不可逆決策／刪除分頁或資料／預計動到超過 ~5 個檔／**要動 Google Sheets 的版面或結構**。
- 其餘日常操作自主執行。

### 3. 安全底線

- 【MUST SEC-1】Sheet ID、Folder ID 等常數集中在 `程式碼.gs` 檔頭，不散落各處。不提供「在 UI 輸入金鑰」的功能。
- 【MUST SEC-2】回給前端的錯誤只給可理解的訊息，不外洩堆疊或內部路徑。

### 4. 驗證

> 要的是**回報附證據**，不是多跑幾輪自我檢查。

- 【MUST VERIFY-1】回報「通過」時附實際輸出。沒跑過就別說跑過。
- 【MUST VERIFY-2】本專案**沒有 lint／build／test 指令**，禁止自編 npm 指令或謊稱通過。實際閘門是：
  - **語法**：`cp 程式碼.gs /tmp/x.js && node --check /tmp/x.js`（`.gs` 副檔名 node 不認，一定要先複製成 `.js`）
  - **MENU_TREE 可解析**：`JSON.parse` 取出 `const MENU_TREE = {...};` 那一行
  - **選單 ↔ 母版對帳**：下載模板 CSV，用與 `rebuildZColumnFor_` 相同規則解析 F 欄，雙向比對
  - **零波及**：與 `git show HEAD` 逐機房比對 JSON，證明沒動到不該動的
- 【MUST VERIFY-3】無法端到端驗證的部分（例如需要師傅實機上傳）明確標示「未端到端驗證」，不可用「應該可以」帶過。

### 5. 錯誤處理

- 【MUST ERR-1】`uploadPhoto` 一律回 `{ ok:false, error:'...' }`，不可拋出未捕捉例外讓前端看到空白。禁止空 catch 吞錯。
- 【MUST ERR-2】catch-all 的訊息不要寫具體原因——那時並不知道實際原因。

### 6. Git

- 【MUST GIT-1】未經明確指示不執行 commit / push / reset --hard / 強推。一個 commit 只含一件邏輯改動，訊息說明「為什麼」。
- 【MUST GIT-2】**做完就要推。** 本 session 曾遇到容器重建、工作目錄只剩根目錄 `index.html`，`gas-maintenance-photo/` 整個消失，全靠遠端分支救回。沒推上去的東西會真的不見。

### 7. 每次任務固定動作

- 【MUST FLOW-1 開場】(1) 讀本檔 (2) 一句話覆述任務目標 (3) 多步驟先列「步驟 → 驗證點」(4) 有歧義先問。
- 【MUST FLOW-2 收尾】回報三件事：改了哪些檔與為什麼、驗證輸出、**還有什麼沒做或未驗證**。

### 8. 溝通

- 【SHOULD COMM-1】回覆聚焦、簡短。第一句先回答「發生什麼事／找到什麼」。
- 【SHOULD COMM-2】使用者是現場工程師不是開發者。步驟要具體到「點哪裡、貼哪個檔、按哪個按鈕」，並明講**這次要貼進哪一個 GAS 專案**。
- 【SHOULD COMM-3】只在「錯誤會改變對方的決定」時才更正先前說法；更正就直接講然後繼續，不要道歉或反覆自責。

### 9. 範圍

- 【MUST SCOPE-1】交付對方要求的範圍。例行判斷自己做，只有不同解讀會導致明顯不同產出時才問。
- 【MUST SCOPE-2】認為需求有問題，用一兩句講出來，然後照原樣做完——不私自縮小或擴大範圍。
- 【MUST SCOPE-3】整個任務做完才回報完成。做不完的部分明講缺什麼、為什麼。
- 【SHOULD DELEGATE-1】不要為了驗證或複查開 subagent，那屬於主流程。

---

## 10. 專案特定規則 ★本檔核心★

### 最高優先

- 【MUST P-1】**絕對不可更改 Google Sheets 的任何格式、欄位順序、標題列或排版。** 業主明示的最高優先指令。任務僅限優化程式碼邏輯與效能，一律配合既有表格結構。

### 兩個 GAS 專案不可貼錯

| 專案 | 檔案 | 特徵 |
|---|---|---|
| **獨立網頁 App**「高公局空調報表自動化」 | `程式碼.gs`、`index.html` | 有 `doGet`；改完要**部署新版本** |
| **試算表綁定**（擴充功能 → Apps Script） | `選單按鈕.gs`（業主自有）、`報表工具.gs`、`健檢工具.gs`、`下載對話框.html` | 有 `onOpen`；只有綁定專案能觸發選單 |

- 【MUST P-2】曾發生把 PDF 程式碼貼進網頁 App 專案，覆蓋掉 `程式碼.gs` 與 `index.html`，整個網頁掛掉（`找不到指令碼函式：doGet`）。**每次交付都要明講貼進哪一個。**
- 【MUST P-3】**新增的 `.gs` 檔不要自帶 `onOpen()`** —— 綁定專案只能有一份，同名會互相覆蓋。改為請業主在既有 `onOpen()` 選單鏈加一行（`報表工具.gs`、`健檢工具.gs` 都是這個模式）。

### Z 欄定位系統 — 這是整個系統的唯一真相

Z 欄（第 26 欄）存放定位鍵 `機房__設備+項目說明`，`uploadPhoto` 用 `createTextFinder().matchEntireCell(true).findAll()` 找照片格。

- 【MUST P-4】**`MENU_TREE` 的 `desc` 就是定位鍵，必須與母版 F 欄文字一字不差。** 包含看起來像錯字的部分，全都要照抄：`-CHU-03液管視窗冷媒流量檢查`（開頭連字號）、`:CTR-02管路系統定期換水`（開頭冒號）、`冷卻水加藥設備檢查檢查之2`（重複「檢查」）。`label` 才是下拉顯示文字，可自由修飾。
- 【MUST P-5】母版 F 欄的說明文字**會折到下一列**（例如 `照片內容說明:空氣濾清器電源檢測` ＋ 下一列 `1-1`）。解析時要把兩列接起來，排除開頭是 `施工`／`完工`／`機房`／`保養`／`照片` 的列。這個規則在 `rebuildZColumnFor_`、`auditMenuVsMaster`、`auditAllMonthTabs`、`健檢工具.gs` 的 `chkSlots_` 四處**必須完全一致**，改一處要同步全部。
- 【MUST P-6】**用 `findAll()` 不可改回 `findNext()`** —— 泰控機房有同名多槽位，`findNext` 會靜默寫進錯的格子。
- 【MUST P-7】**合併儲存格錨點修正**（`targetRow = keyRow + 1` 再用 `getMergedRanges()[0].getRow()` 拉回）標記為勿移除，見 `程式碼.gs` 註解。
- 【MUST P-8】判斷格子是否已填**不能只看值是否為空**：`「照片」` 是佔位字要視為空格；已填要看 `typeof v === 'object'`（CellImage）或 `getFormula() !== ''`。

### 母版與月份分頁

- 【MUST P-9】**母版改了 → `MENU_TREE` 必須同步。** 改完在網頁版專案執行 `auditMenuVsMaster()`。不同步的症狀是「選項在手機上根本看不到、該格永遠空白」，**不是跳錯誤**，很容易漏掉。
- 【MUST P-10】**月份分頁是複製當下的模板快照。** `rebuildZColumnFor_` 只能從該分頁自己的 F 欄推導，**無法無中生有**——模板後來新增的項目，舊分頁永遠不會有。舊分頁救不回來，只能刪掉讓它重新複製。用綁定專案的「🔍 分頁健檢」（`checkMonthTabs`）檢查。
- 【MUST P-11】**備份分頁命名避開純 6 位數**（用 `202608備份` 而非 `202608`）。程式用 `/^\d{6}$/` 判斷月份分頁，純 6 位數會被寫入。

### 照片處理

- 【MUST P-12】**一律 Contain ＋ 白色留白，絕不可改成 Cover 裁切。** 時間相機浮水印在左上（GPS 地圖）與右上（時間戳），被裁掉會**送審退件**。實作見 `index.html` 的 `Math.min(outW/iw, outH/ih, 1)` ＋ `fillStyle` 白底（JPEG 無透明通道）。
- 照片格規格：A~D 共 4 欄 × 101px = 404px，跨 10 列 × 33px = 330px（比例約 1.224）。
- 施工／完工日期取自照片 EXIF 拍攝時間（tag `0x9003`），月份分頁與 Drive 資料夾也依此決定。

### Google Sheets 的坑

- 【MUST P-13】**Sheets 會把 `N-M` 自動轉成日期**（`3-2` → 3月2日）。空氣濾清器那 24 格用的就是這個格式。這類儲存格要**先設成「純文字」**或前置半形單引號 `'3-2`。中招的症狀：定位鍵變成 `空氣濾清器功能檢測Mon Mar 02 2026 16:00:00 GMT+0800`。
- 綁定專案的 `選單按鈕.gs` 用硬編正規式 `115/05/\d{2}` 批次換日期，所以 `clearTemplatePhotos()` 會把模板日期還原成 `115/05/07`，這是預期行為。

### 目前基準（2026-09 驗證）

```
母版照片格 666 個　選單項目 666 項
【A】選單有、母版沒格子：0 ✓
【B】母版有、選單漏掉：0 ✓
完全對應的機房：17 / 17
```

改動 `MENU_TREE` 或模板後，這四行都要重新驗證並更新。

### In Scope / Out of Scope

- **In Scope**：手機拍照上傳歸檔、月份分頁自動建立、PDF 匯出、選單／母版對帳與健檢
- **Out of Scope**：改成原生 App、引入前端框架或建置流程、動 Google Sheets 版面、報價 App 與 GAS 系統互通

---

## 11. 變更紀錄

- 【MUST LOG-1】每次因回饋或需求調整而改動程式，在 `CHANGELOG.md` 記一筆：日期、改了什麼、為什麼、影響範圍；對應一個 git commit。
- 【MUST LOG-2】動手前先確認語法閘門（`node --check`）與對帳為通過狀態當基準；改完再跑一次通過才算完成。
- 【MUST LOG-3】踩過的坑寫進本檔第 10 節，並註明「不可改回去」的理由。不寫下來，下一個人會順手改回去。

<tone_preference>
輸出保持精簡。
</tone_preference>
