# CoachFlow Prototype

## 專案說明
CoachFlow 是一套重訓課表與訓練紀錄系統原型，分成三個入口：

- `admin`：人員管理者控制台
- `coach`：教練工作台
- `student`：學生填寫入口

目前這份專案主要用來驗證：
- 多角色分流
- 教練建立課表
- 學生完成紀錄
- 多教練與學生指派
- 歷史資料匯入
- 學生個人歷史查詢與建議重量

---

## 入口頁

- [admin.html](C:\Users\User\Desktop\codex\admin.html)
- [coach.html](C:\Users\User\Desktop\codex\coach.html)
- [student.html](C:\Users\User\Desktop\codex\student.html)

---

## 角色分工

### 管理者 `admin`
- 管理教練
- 管理學生
- 指派學生的主要教練
- 啟用 / 停用 / 刪除教練與學生

### 教練 `coach`
- 教練登入
- 建立課表
- 管理自己學生
- 查看今日紀錄
- 查看歷史查詢
- 匯入歷史訓練紀錄

### 學生 `student`
- 學生登入
- 自動讀取自己教練的課表
- 填寫今日訓練紀錄
- 查看自己的歷史紀錄
- 下載本次訓練完成圖片

---

## 目前資料來源

目前前端仍以：

- `localStorage`

模擬正式資料流程，所以這是一份：

**可操作的穩定原型**

而不是最終正式上線版。

---

## Google Sheets / Apps Script 狀態

專案已重建一份對應目前架構的後端骨架：

- [google-apps-script/Code.gs](C:\Users\User\Desktop\codex\google-apps-script\Code.gs)
- [google-apps-script/appsscript.json](C:\Users\User\Desktop\codex\google-apps-script\appsscript.json)

對應的資料表為：

- `Coaches`
- `Students`
- `Programs`
- `ProgramItems`
- `WorkoutLogs`

### 重要說明
目前：
- **後端結構已完成**
- **前端尚未正式改成呼叫 Apps Script API**

也就是說：
- Google Sheets 這段可以先部署完成
- 下一階段再把 `admin / coach / student` 三個入口接上雲端資料

部署方式請看：
- [SETUP_GOOGLE_SHEETS.md](C:\Users\User\Desktop\codex\SETUP_GOOGLE_SHEETS.md)

---

## 主要檔案

- [index.html](C:\Users\User\Desktop\codex\index.html)
- [app.js](C:\Users\User\Desktop\codex\app.js)
- [styles.css](C:\Users\User\Desktop\codex\styles.css)
- [PROJECT_SPEC.md](C:\Users\User\Desktop\codex\PROJECT_SPEC.md)
- [PROTOTYPE_TO_PRODUCTION_PLAN.md](C:\Users\User\Desktop\codex\PROTOTYPE_TO_PRODUCTION_PLAN.md)
- [SETUP_GOOGLE_SHEETS.md](C:\Users\User\Desktop\codex\SETUP_GOOGLE_SHEETS.md)

---

## 穩定副本

目前已保留穩定版與快照：

- [stable](C:\Users\User\Desktop\codex\stable)
- [snapshots/2026-03-31-stable](C:\Users\User\Desktop\codex\snapshots\2026-03-31-stable)

說明檔：

- [stable/STABLE_NOTES.md](C:\Users\User\Desktop\codex\stable\STABLE_NOTES.md)
- [snapshots/2026-03-31-stable/SNAPSHOT_NOTES.md](C:\Users\User\Desktop\codex\snapshots\2026-03-31-stable\SNAPSHOT_NOTES.md)

---

## 建議下一步

接下來最推薦的順序是：

1. 建立 Google Sheets / Apps Script 後端
2. 建立前端 API 抽象層
3. 先把 `student` 接到雲端
4. 再接 `coach`
5. 最後接 `admin`

原因：
- 學生端資料流最單純
- 教練與管理者端有較多管理寫入行為
- 先打通學生端，能最快驗證雲端流程
