# 職健導航：AdSense 與 SEO 後續操作計畫

## 目前狀態

網站已經具備首頁單一廣告版位、隱私權與使用說明、`robots.txt`、`sitemap.xml`、canonical、Open Graph 基本標籤，並且 Cloudflare Pages 已連接 GitHub。首頁的健檢資料處理流程仍是瀏覽器本機處理。

目前 `index.html` 的 canonical、Open Graph 與 sitemap 使用 `https://healthgrade.pages.dev/`。如果之後改用自訂網域，必須同步修改 canonical、Open Graph URL、sitemap、robots.txt 的 Sitemap 行，以及 Search Console 資源；不能同時把兩個網址當成主要網址。

## 第一階段：完成 AdSense 審核

1. 打開 Google AdSense 的「網站」頁面，確認網站網址是 `healthgrade.pages.dev`，不是內部 Preview 網址。
2. 確認 AdSense 程式碼已發布在正式網站，且首頁可不登入、不輸入密碼就能讀取。Google 官方說明指出，網站審核通常需要幾天，部分情況可能需要 2–4 週；審核期間也要完成付款資料設定。[1]
3. 在 AdSense 的網站連接流程中確認程式碼、meta 標籤或 ads.txt 的驗證方式。Google 官方列出的連接方式包括 AdSense code snippet、ads.txt 與 meta tag；完成後要按 Verify，再按 Request review。[1]
4. 如果 AdSense 顯示需要 `ads.txt`，請從 AdSense 複製它提供的完整內容，再放到專案的 `public/ads.txt`，讓正式網址可以開啟 `https://healthgrade.pages.dev/ads.txt`。不要自行猜測或修改 Google 給的內容；修改後要重新推送 GitHub，再等待 Cloudflare 部署。
5. 暫時維持首頁只有一個廣告。工具內頁不載入 AdSense 程式，避免干擾 Excel 匯入、分析與結果檢視。Google SEO 指南也提醒不要讓廣告過度干擾內容或妨礙使用者使用頁面。[2]
6. AdSense 的隱私與同意設定要在「Privacy & messaging」檢查。若網站有來自 EEA、英國或瑞士的流量，Google 對個人化廣告要求使用 Google 認證、整合 IAB TCF 的 CMP；目前首頁的「同意並進入工具」是網站使用聲明，不應直接視為已完成 AdSense 的區域廣告同意要求。[3]

## 第二階段：Search Console 收錄

1. 在 Search Console 保留已驗證的 `https://healthgrade.pages.dev/` 資源。
2. 在「Sitemaps」提交：`sitemap.xml`。目前網站 sitemap 已放在根目錄，且使用絕對網址；Google 官方建議 sitemap 放在網站根目錄，並提醒提交 sitemap 只是爬取提示，不保證立即收錄。[4]
3. sitemap 成功後，不要每天重複提交。使用「網址審查」檢查首頁，確認 Google 是否能擷取、是否允許建立索引；遇到每日要求索引配額超限時，等待配額恢復即可。
4. 每月查看 Search Console 的「成效」、「網頁索引」、「手動處置」與「安全性問題」。Search Console 官方把這些列為網站管理者應優先查看的報告。[5]

## 第三階段：實際 SEO 優化

### 內容與標題

首頁標題目前已包含「職健導航」與「本機健檢分級工具」。後續可把首頁可見文字穩定集中在職護使用者會搜尋的詞，例如「職場健康檢查分級」、「健檢資料分級」、「職護個案管理」、「Excel 健檢分級」與「腦心血管風險試算」。不要為了塞關鍵字而重複堆疊；Google 建議以讀者可能使用的搜尋詞撰寫清楚、實用的內容。[2]

### 可被搜尋的公開說明頁

目前工具介面是單頁工具。下一步最值得做的是增加一個公開、可被 Google 讀取的「使用說明」區塊或頁面，說明工具用途、去識別化流程、分級規則、Excel 欄位要求、隱私限制與人工覆核。不要把真實健檢資料放入公開頁面。公開內容應有明確標題、段落與內部連結。

### 結構化資料

可以加入與頁面可見內容一致的 JSON-LD，例如 WebSite、Organization 或 WebApplication 的基本描述。Google 官方建議 JSON-LD 是較容易維護的結構化資料格式，但結構化資料不保證一定出現特殊搜尋結果，且不能標記頁面上沒有呈現的內容；完成後要用 Rich Results Test 驗證。[6]

### 圖片與連結

首頁醫護插圖已改為專案內資產，並有描述性的 alt text。Google 建議圖片放在相關文字附近，並使用能說明圖片內容的 alt text。[2] 對外分享時，優先分享首頁正式網址，不要分享內部 Preview 網址；也可在職護社群、職業安全衛生社群或相關文章中，以自然且有價值的方式介紹工具，不要大量重複貼連結。

## 建議操作順序

先確認正式網站可以開啟、醫護圖片正常顯示，再到 AdSense 完成 Verify 與 Request review。接著在 Search Console 確認 sitemap 狀態與首頁網址審查。AdSense 審核等待期間，持續使用 Search Console 觀察收錄，不要反覆要求索引；網站內容則以增加一個公開使用說明區塊與完善可讀文字為下一個 SEO 開發項目。

## 官方來源

[1]: https://support.google.com/adsense/answer/7584263?hl=en "Google AdSense Connect your site to AdSense"
[2]: https://developers.google.com/search/docs/fundamentals/seo-starter-guide "Google Search Central SEO Starter Guide"
[3]: https://support.google.com/adsense/answer/13554116?hl=en "Google AdSense consent management requirements"
[4]: https://developers.google.com/search/docs/crawling-indexing/sitemaps/build-sitemap "Google Search Central Build and submit a sitemap"
[5]: https://support.google.com/webmasters/answer/10267942?hl=en "Google Search Console getting started"
[6]: https://developers.google.com/search/docs/appearance/structured-data/intro-structured-data "Google Search Central structured data introduction"
