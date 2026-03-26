# OSK — On-Screen Keyboard 螢幕小鍵盤 ⌨️

> 一款專為 Windows 打造的**羽量級、全功能、動態可視化**虛擬螢幕鍵盤，支援物理鍵盤同步、注音輸入法偵測、自訂版面與流體動畫特效。

<img width="1503" height="532" alt="image" src="https://github.com/user-attachments/assets/a82fd8f2-e670-4153-883e-ab1809dfe1a4" />
<img width="1509" height="539" alt="image" src="https://github.com/user-attachments/assets/3cde6b76-0c0b-4ea7-90d9-4ab3a1ae3500" />
<img width="1503" height="552" alt="image" src="https://github.com/user-attachments/assets/cff9fd65-ff6c-4f5a-a7a5-c8331eadbb09" />


## 🌟 核心特色

### 雙模式版面無縫切換

| 模式 | 說明 |
|------|------|
| **100% 全尺寸** | 含多媒體控制鍵（靜音、音量）、完整 Numpad、方向鍵 |
| **60% 精簡版** | 緊湊設計省螢幕空間，搭配 `Fn` 鍵實現 100% 功能 |

### 物理與虛擬雙向同步

點擊螢幕上的虛擬按鍵，或按下實體鍵盤，OSK 的按鍵都會即時亮起對應顏色，完整同步。

### 智慧「四角落」提示佈局

每顆按鍵同時顯示四種輸出對照（支援動態與靜態模式）：

- **左上角** — 預設英文字元
- **右上角** *(深藍色)* — 搭配 `Shift` 的符號 / 大寫
- **左下角** *(綠色)* — 搭配 `Fn` 的 F1–F12 與系統指令
- **右下角** *(橘色)* — 注音輸入法對應符號

> **動態模式**：按下 Shift 或 Fn 時，鍵盤字元即時切換，直接顯示即將輸出的結果。

### 拖移物理動畫特效（虛擬畫布架構）

- 鍵盤主體覆蓋整個虛擬螢幕（所有顯示器），按鍵可自由飛出邊框動畫，不受視窗邊界限制
- 拖移時按鍵呈流體彈跳跟隨效果，放開自動回彈至螢幕範圍內
- 可透過工具列按鈕即時開關特效

### 完美區分左右修飾鍵

支援左、右 `Shift` / `Ctrl` / `Alt` 獨立識別，對遊戲與快捷鍵軟體極度友善。

### 自動日夜主題適應
 
 跟隨 Windows 系統淺色 / 深色主題自動切換，也可手動覆寫。
 
### 🆘 SOS 緊急備援與 UIPI 突破 (新功能)

當遇到管理者權限 (Admin/UAC) 或 Splashtop 遠端桌面等嚴格限制模擬輸入的「死硬視窗」時：
- **一鍵轉換**：點擊工具列的 `🆘` 按鈕，本體會自動備份釘選狀態並隱藏。
- **系統協助**：透過 `cmd /c` 呼叫微軟內建 `osk.exe` 突破 UIPI 防線，協助您在關鍵時刻輸入密碼。
- **自動回歸**：一旦偵測到內建鍵盤關閉，本體會立刻在 1 秒內自動顯示並恢復原本的圖釘屬性。

---

## 🔘 工具列按鈕說明

| 按鈕 | 說明 |
|------|------|
| `⌨️ / 📱` | 切換 60% / 100% 版面 |
| `🗔 / 🗖` | 切換動態 / 靜態版面（Shift/Fn 時是否即時換字） |
| `📍 / 📌` | 固定顯示 / 自動隱藏（偵測輸入焦點） |
| `❄️ / 🌊` | 開關拖移物理特效（❄️ = 開啟，🌊 = 關閉） |
| `🌙 / ☀️` | 深色 / 淺色主題 |
| `透明度滑桿` | 調整視窗透明度，避免遮擋底層視窗 |
| `⟲` | 恢復預設尺寸與配置 |
| `⚙` | 進入版面編輯模式（可拖曳對調按鍵位置） |
| `🆘` | **SOS 緊急備援**：啟動微軟內建小鍵盤以突破防毒/管理者權限視窗攔截 |
| `🗕` | 最小化至系統托盤 |
| `✕` | 關閉鍵盤（自動執行秒退機制，不留殭屍執行緒） |

> 視窗左上角會顯示目前焦點的**程式名稱**與**輸入法狀態**（注音 / 英文）。

---

## 🚀 安裝與執行

1. 確認系統已安裝 [.NET 8.0 Desktop Runtime](https://dotnet.microsoft.com/download/dotnet/8.0)
2. Clone 或下載本專案：
   ```bash
   git clone https://github.com/<your-repo>/osk.git
   ```
3. 進入專案目錄並建置：
   ```bash
   cd osk/src
   dotnet build
   ```
4. 執行：
   ```bash
   dotnet run
   ```
   或直接雙擊 `bin/Debug/net8.0-windows/OSK.exe`

---

## 📁 專案結構

```
src/
├── MainWindow.xaml          # UI 定義（虛擬畫布架構）
├── MainWindow.xaml.cs       # 核心邏輯（物理、輸入法偵測、拖移）
├── InputDetector.cs         # 焦點與輸入法監測
├── Assets/                  # 圖示資源
└── osk.ini                  # 執行時期設定（自動產生）
```

---

## ⚙️ 組態設定（osk.ini）

程式關閉時自動儲存以下設定到 `osk.ini`：

| 鍵值 | 說明 |
|------|------|
| `Left` / `Top` | 鍵盤在虛擬畫布上的位置 |
| `Width` / `Height` | 鍵盤容器大小 |
| `IsFullLayout` | 是否使用 100% 版面 |
| `IsDynamicLayout` | 是否啟用動態版面 |
| `IsPinned` | 是否固定顯示 |
| `IsDarkMode` | 主題 |
| `Opacity` | 透明度 |
| `CompactRow0–4` / `FullRow0–5` | 自訂按鍵排列 |

---

## 🛠️ 技術架構重點

- **虛擬畫布 (Virtual Screen Canvas)**：`MainWindow` 鋪滿所有顯示器，鍵盤本體為其內部的 `KeyboardContainer`，`ClipToBounds="False"` 讓按鍵動畫可超出邊框。
- **VSync 對齊物理引擎**：物理計算掛載於 `CompositionTarget.Rendering`，與 WPF 渲染循環同步，消除頓挫感。
- **彈簧阻尼模型**：每顆按鍵擁有獨立的 `FollowSpeed`、`Friction`、`Mass` 參數，拖移時呈現流體跟隨效果。
- **零字串配置熱路徑**：狀態 Hash 使用位元整數，避免 30fps 物理迴圈每幀產生 GC 壓力。
- **輸入法偵測**：透過 IMM32 API (`ImmGetConversionStatus`) 即時偵測注音 / 英文模式。
- **雙重監控與 Latch 辨識**：監控 `osk.exe` 時結合「進程路徑比對」與「視窗類別 (`OSKMainClass`)」辨識，排除權限不足導致的偵測失效。
- **秒退優化 (Nuclear Shutdown)**：在 `OnClosed` 調用 `Environment.Exit(0)`，徹底終止 UIA 殘留執行緒，確保開發測試時可即關即開。
 
---
 
**© 2026 OSK Project - 打造最強大且具備 UIPI 突破能力的專業虛擬鍵盤**
