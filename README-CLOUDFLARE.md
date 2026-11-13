# 職健導航｜Cloudflare Pages 部署說明

這是一個 Vite + 原生 JavaScript 的純前端工具。健檢 Excel 只在使用者瀏覽器中讀取與處理，不使用後端 API、資料庫或伺服器儲存健檢資料。

## 方案一：連接 GitHub 自動部署

1. 將本專案完整資料夾推送到自己的 GitHub repository。
2. 開啟 Cloudflare Dashboard → **Workers & Pages** → **Create application** → **Pages** → **Connect to Git**。
3. 選取此 repository。
4. 建置設定如下：
   - Framework preset：`Vite`
   - Build command：`npm run build`
   - Build output directory：`dist`
   - Root directory：`/`
5. 儲存並部署。

## 方案二：直接上傳建置結果

若不使用 GitHub，可在專案根目錄執行：

```bash
npm install
npm run build
```

接著在 Cloudflare Pages 建立專案時，選擇 **Direct Upload**，上傳整個 `dist` 資料夾。


## 自動化測試

本專案使用 Node.js 內建測試執行器，測試檔位於 `tests/healthgrade-core.test.js`，並與 `src/healthgrade-core.js` 共用實際使用的核心邏輯。測試涵蓋數值解析、小數點一位四捨五入、區間與單邊條件、邊界重疊、欄名別名、序號去重，以及 Framingham 必要欄位與年齡範圍檢核。

```bash
npm install
npm run test
npm run build
```

## 本機測試

```bash
npm install
npm run dev
```

瀏覽器開啟 `http://localhost:3000`。

正式建置檢查：

```bash
npm run build
```

## 重要檔案

- `src/main.js`：分級、統計、Excel 匯入／匯出與介面邏輯
- `src/rules.json`：網站內建分級規則
- `src/styles.css`：介面樣式
- `public/下載參考.xlsx`：匯入範例檔案
- `public/logo.svg`：本地 Logo 與瀏覽器分頁圖示
- `public/manus-routes.json`：網站路由宣告
- `docs/`：規則與模型參考文件

## 部署注意事項

這是單頁靜態網站，不需要 Cloudflare Workers、D1、R2 或其他資料庫服務。請勿在未經個案同意的情況下加入第三方分析、錄影或會將健檢資料送出瀏覽器的程式碼。

若使用 Cloudflare Web Analytics 或廣告服務，請另行檢查其隱私政策與個資告知義務；目前本工具的健檢 Excel 處理流程不會上傳資料。

使用者自行新增、修改或刪除的分級規則不會自動保存到網站伺服器；若要保留，請在網站內下載 Excel 備份，下次再匯入。
