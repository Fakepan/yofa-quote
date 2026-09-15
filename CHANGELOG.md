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
