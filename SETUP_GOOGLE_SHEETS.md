# CoachFlow Google Sheets / Apps Script 部署說明

## 目前狀態
這份專案目前分成兩個部分：

1. 前端原型  
   - `admin.html`
   - `coach.html`
   - `student.html`
   - `index.html`
   - `app.js`
   - `styles.css`

2. Google Sheets / Apps Script 後端骨架  
   - `google-apps-script/Code.gs`
   - `google-apps-script/appsscript.json`

## 重要說明
目前 **Apps Script 後端資料結構已重建完成**，對應現在的：
- 管理者頁 `admin`
- 教練頁 `coach`
- 學生頁 `student`

但 **前端尚未正式改成直接讀寫 Apps Script API**。  
也就是說：

- 你現在可以先把 Google Sheets / Apps Script 建好
- 後端資料表會先就位
- 下一階段再把前端從 `localStorage` 改成正式呼叫雲端 API

這份文件的目的是先把 **Google Sheets 後端環境建立完成**。

---

## 一、建立 Google Sheet

1. 建立一份新的 Google Sheet
2. 建議命名為：`CoachFlow`
3. 打開這份 Google Sheet
4. 進入：`擴充功能 -> Apps Script`

---

## 二、貼上 Apps Script 檔案

### `Code.gs`
將下列檔案內容完整貼入 Apps Script：

- [google-apps-script/Code.gs](C:\Users\User\Desktop\codex\google-apps-script\Code.gs)

### `appsscript.json`
在 Apps Script 專案設定中開啟 `顯示 appsscript.json` 後，將下列檔案內容貼入：

- [google-apps-script/appsscript.json](C:\Users\User\Desktop\codex\google-apps-script\appsscript.json)

---

## 三、第一次執行與授權

貼上後，先在 Apps Script 編輯器中手動執行任一個會建立資料表的流程，例如：

- `doGet`

或先部署前直接執行：

- `ensureSheets_()`

系統會要求你授權 Google Sheets 存取權限。  
完成授權後，這份 Google Sheet 會自動建立下列工作表：

- `Coaches`
- `Students`
- `Programs`
- `ProgramItems`
- `WorkoutLogs`

---

## 四、資料表欄位

### `Coaches`
- `coach_id`
- `coach_name`
- `status`
- `role`
- `access_code`
- `token`
- `last_used_at`
- `created_at`
- `updated_at`

### `Students`
- `student_id`
- `student_name`
- `status`
- `access_code`
- `token`
- `primary_coach_id`
- `primary_coach_name`
- `last_used_at`
- `created_at`
- `updated_at`

### `Programs`
- `program_id`
- `program_code`
- `program_date`
- `title`
- `coach_id`
- `coach_name`
- `notes`
- `published`
- `created_at`
- `updated_at`

### `ProgramItems`
- `item_id`
- `program_id`
- `sort_order`
- `category`
- `exercise`
- `target_sets`
- `target_type`
- `target_value`
- `item_note`
- `updated_at`

### `WorkoutLogs`
- `log_id`
- `program_id`
- `program_code`
- `program_date`
- `coach_id`
- `coach_name`
- `student_id`
- `student_name`
- `category`
- `exercise`
- `target_sets`
- `target_type`
- `target_value`
- `actual_weight`
- `actual_sets`
- `actual_reps`
- `student_note`
- `submitted_at`
- `updated_at`

---

## 五、目前後端支援的 API 動作

### 讀取資料
- `bootstrap`
- `bootstrapAdmin`
- `bootstrapCoach`
- `bootstrapStudent`

### 教練管理
- `upsertCoach`
- `setCoachStatus`
- `deleteCoach`
- `touchCoach`

### 學生管理
- `upsertStudent`
- `assignStudentCoach`
- `setStudentStatus`
- `touchStudent`

### 課表管理
- `saveProgram`
- `publishProgram`

### 紀錄管理
- `submitWorkoutLogs`
- `importWorkoutLogs`

---

## 六、部署成 Web App

1. 在 Apps Script 編輯器右上角按 `部署`
2. 選擇 `新增部署作業`
3. 類型選 `網頁應用程式`
4. 設定：
   - 執行身分：你自己
   - 存取權限：依需求選擇（測試時可先用你自己可存取）
5. 按 `部署`
6. 複製 Web App URL

---

## 七、前端設定檔

專案目前有：

- [config.js](C:\Users\User\Desktop\codex\config.js)

目前內容如下：

```js
window.APP_CONFIG = {
  mode: "cloud",
  appsScriptUrl: "你的 Apps Script Web App 網址",
  requestTimeoutMs: 12000
};
```

## 重要提醒
雖然 `config.js` 已存在，而且可填入 Apps Script URL，  
但目前前端程式 **尚未完成正式雲端 API 串接**。

也就是說：
- `config.js` 目前屬於預先配置
- 真正讓前端改成呼叫 Apps Script，是下一階段工作

---

## 八、建議下一步

Google Sheets / Apps Script 建好後，下一步建議依序進行：

1. 前端 API 抽象層建立
2. `student` 先改成讀寫 Apps Script
3. `coach` 再接上雲端
4. `admin` 最後接上雲端

原因：
- 學生端資料流最單純
- 教練端與管理者端牽涉的寫入邏輯比較多
- 先把學生端打通，整體風險最低

---

## 九、目前可交付結果

如果你現在照這份文件完成部署，會得到：

- 一份已建立好結構的 Google Sheet
- 一個可回應 CoachFlow 資料表操作的 Apps Script Web App
- 一個可作為下一階段前端串接目標的正式後端

還不會立即得到：

- 已完全雲端化的 `admin / coach / student` 前端

那是下一階段要接的內容。
