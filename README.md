#  塔羅占卜機器人 Tarot Card Fortune-Telling Robot

> 國立清華大學工業工程與工程管理學系「程式設計與應用」期末專題  
> 使用 C# Windows Forms 製作，結合 OpenAI GPT-4o 占卜解讀與會員互動功能之塔羅占卜應用程式！  
#### 🎵 每一個頁面都有對應的背景音樂，打造沉浸式占卜體驗。
---

## 📸 操作介面一覽

### 🔐 登入與註冊

![登入頁面](./screenshots/login.png)  
> 使用者可輸入帳號密碼登入，亦支援註冊與忘記密碼信箱驗證

---

### 🧙‍♀️ AI塔羅占卜

![AI塔羅主頁](./screenshots/ai_main.png)  
> 六種塔羅師風格可選，提供風格化 AI 占卜體驗（部分風格會員限定）

---

### ✍️ 問題輸入頁面

![問題輸入](./screenshots/question_input.png)  
> 使用者可輸入問題（限200字內），再進行抽牌與占卜

---

### 🎴 抽牌介面與動畫

![抽牌介面](./screenshots/card_draw.png)  
> 使用動畫與視覺特效呈現翻牌體驗，支援 1 張或 3 張牌抽取

---

### 📬 寄送占卜結果

![寄信功能](./screenshots/send_email.png)  
> 使用者可選擇是否透過 Email 寄送本次占卜解析與建議

---

### 📈 塔羅運勢 & 是/否占卜

![運勢塔羅](./screenshots/luck.png)  
> 選擇生活領域與塔羅風格後進行占卜（非會員無法使用完整功能）

---

### 📅 占卜預約系統

![預約介面](./screenshots/booking.png)  
> 使用者可挑選塔羅師與時段預約；管理者後台可匯出與分析預約資料

---

## ⚙️ 技術架構

- **開發語言**：C# (.NET 6)
- **GUI 框架**：Windows Forms
- **API 串接**：OpenAI GPT-4o (API Key 透過 `.env` 管理)
- **資料儲存**：CSV 檔案 (`users.csv`, `appointments.csv`, `schedules.csv`)
- **功能模組**：
  - 使用者登入、註冊、驗證、會員中心
  - AI塔羅占卜、是否占卜、運勢占卜
  - Email 寄信結果功能
  - 占卜預約 & 管理後台功能

---

## 📦 專案結構

📁 第一頁  
├─ bin/Debug  
│ └─ .env (已加入 .gitignore)  
├─ Form1.cs  
├─ TarotDraw.cs  
├─ User.cs  
├─ users.csv  
├─ appointments.csv  
└─ schedules.csv  

---

## 🔐 API 金鑰管理（.env）

請在專案根目錄下建立 `.env` 檔案，內容如下：

OPENAI_API_KEY=sk-xxxxxxx


程式碼需使用 `DotNetEnv` 套件，並於程式初始化時讀取：


using DotNetEnv;

Env.Load(); // 建議放在 Form1_Load  
string apiKey = Environment.GetEnvironmentVariable("OPENAI_API_KEY");  
如何執行?
安裝 .NET 4.8 SDK

git clone 本專案後，開啟 .sln 或 .csproj 檔

設定好 .env 後即可執行！

---

## 📝 License
僅限學術與教學用途，不得用於商業目的。
