# GitHub Pages 部署步驟

這份前端建議部署到 GitHub Pages。

## 1. 在 GitHub 建立新 repository

建議名稱：

- `coachflow`

建立時：

- 可以選 `Public`
- 不要勾 `Add a README`
- 不要勾 `.gitignore`

## 2. 把本地專案連到 GitHub

在 `C:\Users\User\Desktop\codex` 執行：

```powershell
git remote add origin https://github.com/<你的帳號>/coachflow.git
git add .
git commit -m "Initial CoachFlow frontend"
git push -u origin main
```

如果你的 GitHub 倉庫名稱不是 `coachflow`，把上面的網址改成你的實際倉庫網址。

## 3. 開啟 GitHub Pages

到 GitHub 倉庫頁面：

1. 點 `Settings`
2. 點左側 `Pages`
3. `Build and deployment`
4. `Source` 選 `Deploy from a branch`
5. `Branch` 選：
   - `main`
   - `/ (root)`
6. 按 `Save`

幾分鐘後會得到公開網址，通常像：

```text
https://<你的帳號>.github.io/coachflow/
```

## 4. 更新 publicBaseUrl

打開：

- [config.js](C:\Users\User\Desktop\codex\config.js)

把：

```js
publicBaseUrl: "",
```

改成：

```js
publicBaseUrl: "https://<你的帳號>.github.io/coachflow/",
```

注意：

- 結尾保留 `/`

## 5. 再次推送

```powershell
git add config.js
git commit -m "Set GitHub Pages public URL"
git push
```

## 6. 驗證

部署完成後，測這幾個網址：

- `https://<你的帳號>.github.io/coachflow/admin.html`
- `https://<你的帳號>.github.io/coachflow/coach.html`
- `https://<你的帳號>.github.io/coachflow/student.html`

再測：

- 學生專屬網址
- 教練專屬網址
- QR code 是否能正常掃描開啟

## 備註

- 前端放 GitHub Pages
- 後端仍使用 Google Sheets + Apps Script
- 如果之後換網址，記得同步更新 `publicBaseUrl`
