# 司法裁判與法律爭點分析 Skill (Legal Issue Analysis) 使用與相依套件說明

> [!WARNING] **【免責聲明 / Disclaimer】**
> 1. **僅供學術研究與論文寫作參考**：本 Skill、自動化腳本及所產出之爭點分析報告、比對矩陣、司法三段論演繹與犯罪三階層心證整理，**僅供法學學術研究、專題探討與學位論文寫作參考，不構成任何形式之正式法律意見、法律諮詢、訴訟代理或裁判結果預測**。
> 2. **人工查核義務**：AI 智慧萃取與自動化腳本分析之內容可能因原始裁判文本品質或格式變更而有所誤差，**研究者於論文引用或實務論證前，應自行查閱司法院裁判書公開系統或金融消費評議中心之正式裁判公文書核對原貌**。
> 3. **無擔保聲明**：本套件依「現狀（AS-IS）」提供，開發者與維護者不對產出成果之完整性、精確性或特定目的之適用性提供任何明示或默示保證；使用者因信賴或引用本工具產出所致之任何直接、間接或衍生性損害，概不承擔任何法律責任。
> 4. **合規與合理使用**：使用自動下載工具存取公開司法資訊時，請遵守智慧財產權與目標主管機關網站之服務條款，秉持學術合理使用原則，避免頻繁請求對伺服器造成負載。

---

- **最後發佈日期 (Release Date)**：2026-11-22
- **最後修改日期 (Last Modified)**：2026-11-22

本說明文件與 `SKILL.md` 位於同一層目錄，提供本法學分析 Skill 之**所需安裝相依套件**、**環境檢測與自動安裝指令**、**自動化批次腳本使用指南**，以及**AI Agent 執行規範**。

---

## 目錄
1. [需安裝之環境與相依套件清單](#一需安裝之環境與相依套件清單)
2. [環境檢查與無人值守快速安裝](#二環境檢查與無人值守快速安裝)
3. [目錄架構規範](#三目錄架構規範)
4. [核心執行腳本與指令範例](#四核心執行腳本與指令範例)
5. [法學分析架構與輸出成果規範](#五法學分析架構與輸出成果規範)
6. [AI Agent 背景執行與對話原則](#六ai-agent-背景執行與對話原則)

---

## 一、需安裝之環境與相依套件清單

| 元件名稱 | 類型 | 建議版本 | 安裝指令 / 檢查方式 | 說明 |
| :--- | :--- | :--- | :--- | :--- |
| **Node.js** | 執行環境 (Runtime) | **v18.0.0 以上**<br>*(本機現有版本: v26.7.0)* | 終端機執行 `node -v`<br>若未安裝至 [nodejs.org](https://nodejs.org/) 下載 LTS 版 | 支援原生 `fetch`、`ES6+` 與大容量 PDF 文本異步解析 |
| **exceljs** | npm 套件 | **^4.4.0** | `npm --prefix .agents/skills/legal-issue-analysis/scripts install exceljs` | 產出符合學術論文與出版規格之專業 `.xlsx` 試算表（深藍 `#1F4E78` 表頭、首行凍結、斑馬紋、文字自動換行） |
| **pdf-parse** | npm 套件 | **^2.4.5** | `npm --prefix .agents/skills/legal-issue-analysis/scripts install pdf-parse` | 極速解析各級法院裁判書與金融消費評議決定書 PDF 內容 |
| **fs / path** | 核心內建模組 | 隨附於 Node.js | **無需額外安裝** | 負責目錄智慧偵測、檔案讀寫與備份防鎖定處理 |

---

## 二、環境檢查與無人值守快速安裝

### 1. 一鍵安裝相依套件指令（PowerShell / CMD）
在工作區根目錄下，執行以下指令即可於本 Skill 內建之 `scripts/` 安裝所有必要套件：

```powershell
npm --prefix .agents/skills/legal-issue-analysis/scripts install exceljs pdf-parse
```

### 2. AI Agent 無人值守環境檢測腳本（PowerShell）
AI Agent 或接手之 Subagent 可於背景執行以下指令，自動確認環境就緒狀態：

```powershell
# 檢測 Node.js
try {
    $nodeVer = node -v
    Write-Output "Node.js detected: $nodeVer"
} catch {
    Write-Output "Node.js not found. Installing Node.js LTS via winget..."
    winget install OpenJS.NodeJS.LTS --accept-package-agreements --accept-source-agreements --silent
    $env:Path = [System.Environment]::GetEnvironmentVariable("Path","Machine") + ";" + [System.Environment]::GetEnvironmentVariable("Path","User")
}

# 檢測並自動安裝 npm 套件
$scriptDir = ".agents\skills\legal-issue-analysis\scripts"
if (-not (Test-Path "$scriptDir\node_modules\exceljs") -or -not (Test-Path "$scriptDir\node_modules\pdf-parse")) {
    Write-Output "Installing required npm dependencies..."
    npm --prefix $scriptDir install exceljs pdf-parse
}
Write-Output "Environment check completed successfully."
```

---

## 三、目錄架構規範

本 Skill 採用**自包含獨立外掛包架構（Self-Contained Skill Package）**，腳本與工具均完整封裝於 Skill 資料夾內，工作區僅需保留必要之輸入與輸出工作目錄：

```text
[工作區根目錄]/
├── judgements_org/    # 判決書 裁判書 原始目錄（存放待分析之裁判書 PDF 或 TXT）
├── analysis_results/  # 分析成果目錄（存放產出之 Markdown 分析報告與 Excel 總表）
└── .agents/
    └── skills/
        └── legal-issue-analysis/       # 🌟 本 Skill 套件目錄（可直接拷貝至任何專案）
            ├── SKILL.md                # 法學分析核心規格規範
            ├── README.md               # 本說明文件
            └── scripts/                # 內建批次自動化工具鏈
                ├── batch_analyze_judgments.js  # 🚀 自動化批次裁判爭點分析引擎
                ├── download_foi_decisions.js   # 評議決定書自動下載工具 (ods.foi.org.tw)
                ├── export_matrix_to_excel.js   # Markdown 比對表轉 Excel 工具
                └── package.json                # 套件設定檔 (exceljs, pdf-parse)
```

---

## 四、核心執行腳本與指令範例

執行本 Skill 內建腳本之指令如下：

### 1. 自動化裁判批次分析腳本 (`batch_analyze_judgments.js`)
本腳本會自動讀取 `judgements_org/` 內之裁判書，進行**司法三段論演繹**、**犯罪三階層審查**及**完整量刑論述提取**，並產出雙軌雙格式成果至 `analysis_results/` 目錄。

- **全量批次執行**：
  ```bash
  node .agents/skills/legal-issue-analysis/scripts/batch_analyze_judgments.js
  ```
- **測試快速檢視（限制分析筆數，例如 5 筆）**：
  ```bash
  node .agents/skills/legal-issue-analysis/scripts/batch_analyze_judgments.js --limit 5
  ```
- **指定自訂來源、輸出目錄與論文主題**：
  ```bash
  node .agents/skills/legal-issue-analysis/scripts/batch_analyze_judgments.js --input "judgements_org" --output "analysis_results" --topic "加重詐欺與洗錢防制之量刑裁量研究"
  ```

### 2. 金融消費評議決定書批次下載腳本 (`download_foi_decisions.js`)
專門對接「金融消費評議中心查詢系統 (`https://ods.foi.org.tw/`)」。

- **互動式執行（推薦研究者手動操作）**：
  ```bash
  node .agents/skills/legal-issue-analysis/scripts/download_foi_decisions.js
  ```
- **帶參數執行（AI 經使用者確認條件後調用）**：
  ```bash
  node .agents/skills/legal-issue-analysis/scripts/download_foi_decisions.js --keyword "旅遊不便險" --start "1140101"
  ```

---

## 五、法學分析架構與輸出成果規範

### 1. 兩大核心法學論理體系
- **司法三段論法（Legal Syllogism）**：
  - **大前提（Major Premise）**：實體法法定構成要件、判例/大法庭裁定確立之法律原則。
  - **小前提（Minor Premise）**：法院採證審認之具體個案要件事實與行為情節。
  - **涵攝與結論（Subsumption & Conclusion）**：個案事實代入法規範之邏輯論斷與判決主文效果。
- **犯罪三階層理論（Three-Tier Crime Theory）與完整量刑論述**：
  - **(1) 構成要件該當性 (Tatbestand)**：客觀實行行為、因果歸責與主觀故意/未必故意/意圖。
  - **(2) 違法性審查 (Rechtswidrigkeit)**：阻卻違法事由（正當防衛、緊急避難）審查。
  - **(3) 有責性評價 (Schuld)**：責任能力（刑法第19條、精神/智能狀況、身心障礙訴訟照料）與違法性認識。
  - **(4) 罪數競合與完整量刑心證論述 (Sentencing & Discretion)**：
    - **全面放寬字數與行數限制**，拒絕以「量刑妥適」等一句套話帶過。
    - 完整收錄法官審酌**刑法第57條各款因子**（犯罪動機、目的、手段、分工地位、損害程度、犯後態度、和解賠償）、**刑法第59條酌量減輕**、**自白減刑要件**、**緩刑宣告**及**罪刑相當與比例原則**之完整說理。

### 2. 雙軌產出檔案規範
每次批次處理完成，系統將同步產出：
1. **Markdown 深度分析報告（`.md`）**：
   - 包含實務見解立場光譜、多篇裁判橫向對比大矩陣、實務見解演進與歧異深度剖析、論文寫作引用建議，以及每篇裁判之單篇獨立精要（FIRAC + 三段論 + 犯罪三階層 + 完整量刑論述）。
2. **Excel 各篇分析結果摘要總表（`.xlsx`）**：
   - 標準 12 大欄位：`序號`、`裁判日期`、`裁判/評議字號`、`案由/爭議類型`、`主文判決結果`、`核心法律爭點設問`、`核心涉訟法條與約款`、`法院/委員會立場`、`【司法三段論結構剖析】`、`【法官主要論點（含完整量刑心證）】`、`裁判核心論理原文金句 (Ratio Decidendi)`、`論文寫作引用建議`。
3. **Excel 裁判爭點橫向對比矩陣表（`.xlsx`）**：
   - 8 大精華比對欄位，首行凍結、深藍表頭、斑馬紋，適於論文橫向對照。

---

## 六、AI Agent 背景執行與對話原則

接手本專案的 AI Assistant / Subagent 必須嚴格恪守以下專案規範（依據 `AGENTS.md`）：

1. **安靜背景處理原則（Quiet Processing）**：
   - 處理判決分析或批次處理時，**嚴禁在對話框中傾倒冗長的中間過程、OCR 文本或除錯日誌**。
   - 分析報告與 Excel 總表直接寫入 `analysis_results/` 目錄。
   - 對話框僅需輸出簡明完成通知：
     1. 處理狀態與裁判總數。
     2. 產出檔案之可點擊連結（`file://` 格式）。
     3. 核心結論或爭點光譜重點摘要（3～5 點）。
2. **評議決定書下載前確認原則**：
   - 提出下載評議決定書需求時，**嚴禁私自預設關鍵字或日期區間直接執行下載**。
   - 必須先於對話框中向使用者確認：(1) 檢索關鍵字、(2) 決定日期區間、(3) 產業或爭議類型限制，經確認後方得調用下載腳本。
3. **檔案路徑超連結原則**：
   - 回報檔案時，一律使用標準 GitHub-flavored Markdown `file://` 超連結（例如：`[檔名](file:///C:/Users/...)`，路徑斜線一律使用正斜線 `/`）。
