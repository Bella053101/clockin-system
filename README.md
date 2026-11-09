# 雲端上班打卡與薪資計算系統 (Google Sheet 後端 + HTML 前端)

一個兼具現代美觀介面與實用功能的雲端上班打卡系統：
- **前端技術**：HTML5 + Tailwind CSS + FontAwesome + JavaScript
- **後端技術**：Google Sheets + Google Apps Script (GAS)
- **核心特色**：
  - ⚡ **即時上班/下班打卡**：自動抓取當前時間並寫入資料庫
  - 📅 **日曆與自訂選時**：支援補登任何日期的上班、下班時間
  - 💰 **薪水智能計算**：設定時薪後，即時試算當日工時與累計應得薪資
  - 📊 **統計儀表板**：顯示總出勤次數、累積總工時、累計總薪資
  - ☁️ **Google 試算表串接**：支援將打卡資料無縫寫入 Google Sheet

---

## 快速設定 Google Sheet 後端 (3分鐘)

1. 開啟 [Google 試算表 (Google Sheets)](https://sheets.new)，建立一份新試算表（例如取名為：`員工打卡系統`）。
2. 在試算表上方選單點選 **擴充功能 (Extensions) ➔ Apps Script**。
3. 把本專案中的 `Code.gs` 內容全部複製並貼到 Apps Script 編輯器中，覆蓋原本內容。
4. 點選右上角 **部署 (Deploy) ➔ 新增部署 (New deployment)**：
   - 類型選取齒輪圖示 ➔ **網頁應用程式 (Web app)**
   - 說明輸入：`v1`
   - 誰可以存取 (Who has access)：選擇 **所有人 (Anyone)**
   - 點擊 **部署**
5. 複製生成的 **網頁應用程式網址 (Web App URL)**。
6. 打開本系統網頁，將網址貼入「Google Apps Script 串接網址」欄位並按下儲存即可！
