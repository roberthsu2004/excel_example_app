# 通用型 Excel 佔位符單據產生器 (Excel Template Engine)

[![React](https://img.shields.io/badge/React-19.0-61DAFB?logo=react&logoColor=white)](https://react.dev/)
[![TypeScript](https://img.shields.io/badge/TypeScript-5.x-3178C6?logo=typescript&logoColor=white)](https://www.typescriptlang.org/)
[![Vite](https://img.shields.io/badge/Vite-6.x-646CFF?logo=vite&logoColor=white)](https://vitejs.dev/)
[![TailwindCSS](https://img.shields.io/badge/TailwindCSS-v4-38B2AC?logo=tailwind-css&logoColor=white)](https://tailwindcss.com/)
[![ExcelJS](https://img.shields.io/badge/ExcelJS-4.4-217346?logo=microsoft-excel&logoColor=white)](https://github.com/exceljs/exceljs)

純前端運作的 Excel 自動化單據與報表產生工具。只需在 Excel 檔案中標註 `{{變數名稱 | 預設值}}` 佔位符，系統即可自動解析、動態生成對應的輸入表單，並保留原始 Excel 的公式、圖片與儲存格樣式，一鍵匯出專業單據！

---

## ✨ 核心特色

- 🚀 **純前端極速運算**：免架設後端資料庫或伺服器，所有 Excel 解析與渲染皆在瀏覽器端使用 ExcelJS 完成，高隱私、零資安外洩風險。
- 🔍 **智慧佔位符解析**：支援雙括號語法 `{{變數名稱 | 預設值}}`，自動掃描所有工作表並產生對應的結構化表單。
- 🧮 **公式與樣式 100% 完整保留**：填入資料時精準替換指定文字，保留原有儲存格合併、字型、背景色、框線、圖片及計算公式（如小計、稅額、加總等）。
- 📝 **多種檢視與編輯模式**：
  - **表單模式 (Form View)**：依基本資料、雙方聯絡人、明細項目、備註條款分區呈現。
  - **品項明細管理 (Items View)**：動態增加/刪除明細列，快速填寫商品項目、數量與單價。
  - **佔位符清單 (Table View)**：一覽所有掃描到的儲存格座標（如 `B3`、`E8`）與當前設定值。
- 🖼️ **自訂商標 (Logo)**：支援上傳公司自訂 Logo 圖片，自動替換或嵌入至產生的單據中。
- 📂 **內建雙商務範本**：開箱即附「商務報價單」與「商業請款單」，支援一鍵快速載入與預填範例資料。
- 💾 **設定匯出與載入**：支援複製或匯出 JSON 資料設定，便於下次直接載入或串接批次流程。

---

## 📌 佔位符語法說明

在您的 Excel 範本 (`.xlsx`) 中，可以在任意儲存格輸入以下格式：

| 語法格式 | 說明 | 範例 |
| :--- | :--- | :--- |
| `{{變數名稱}}` | 基本佔位符，預設值為空 | `{{客戶公司名稱}}` |
| `{{變數名稱 \| 預設值}}` | 附帶預設值的佔位符 | `{{報價單號 \| QT-2025001}}` |
| 混合文字與佔位符 | 儲存格中可與固定文字混用 | `統一編號：{{客戶統編 \| 12345678}}` |

> 💡 **小撇步**：英文與中文變數名稱皆可自動識別（例如 `QUOTATION_NO` 或 `報價單號`），系統內建繁體中文自動對照與分類分組。

---

## 🛠️ 技術棧

- **前端框架**：React 19, TypeScript
- **建置工具**：Vite 6, Bun
- **樣式庫**：Tailwind CSS v4, Lucide React (圖示庫), Motion (動畫效果)
- **檔案處理**：ExcelJS (Excel 解析與建構), FileSaver.js (檔案下載)

---

## 🚀 快速開始

### 1. 環境需求
- [Node.js](https://nodejs.org/) (建議 v18 以上) 或 [Bun](https://bun.sh/)

### 2. 安裝相依套件

使用 Bun：
```bash
bun install
```

或使用 npm / pnpm / yarn：
```bash
npm install
```

### 3. 本地啟動開發伺服器

```bash
npm run dev
# 或
bun run dev
```
啟動後瀏覽器開啟 `http://localhost:3000` 即可預覽使用。

### 4. 建置生產版本 (Production Build)

```bash
npm run build
```

---

## 📁 專案結構

```text
excel_example_app/
├── public/                     # 靜態資源與預設範本
│   ├── 商務報價單範本.xlsx       # 內建精美報價單範本
│   ├── 商業請款單範本.xlsx       # 內建精美請款單範本
│   └── logo.png                # 範例商標
├── scripts/
│   └── generate_templates.cjs  # 透過 ExcelJS 程式化建立範本的腳本
├── src/
│   ├── types/
│   │   └── template.ts         # TypeScript 型別定義 (佔位符、單據項目、中繼資料)
│   ├── utils/
│   │   └── excelEngine.ts      # 核心 Excel 模板引擎 (解析、替換、公式維護、匯出)
│   ├── App.tsx                 # 主應用程式介面與互動邏輯
│   ├── main.tsx                # React 入口點
│   └── index.css               # Tailwind 樣式設定
├── package.json
└── README.md
```

---

## 💡 產生與客製 Excel 範本

專案內建了 `scripts/generate_templates.cjs` 腳本，可自動產生格式對齊、字型優雅、帶有計算公式的 Excel 範本：

```bash
node scripts/generate_templates.cjs
```
執行後會在 `public/` 目錄下重新建構標準範本檔案。

---

## 📄 授權條款

本專案採用 [MIT License](LICENSE) 授權釋出。
