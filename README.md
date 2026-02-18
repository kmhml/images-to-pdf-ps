# Images to PDF (PowerShell)

フォルダ内の PNG / JPEG 画像をまとめて  
1つの PDF に変換する PowerShell スクリプトです。

フォルダを指定するだけで、その配下の画像を自動取得し、
日付形式（yyyyMMdd）のPDFを生成します。

---

## 🚀 特徴

- フォルダ選択ダイアログ付き
- PNG / JPG / JPEG に対応
- サブフォルダを含めて再帰取得
- 画像比率を維持して中央配置
- 1画像 = 1ページ
- 外部ライブラリ不要
- Windows標準機能のみで動作

---

## 📦 動作環境

- Windows 10 / 11
- PowerShell 5.1 以上
- Microsoft Print to PDF が有効

※ 追加インストール不要

---

## 🛠 使い方

### ① 実行ポリシー設定（初回のみ）

PowerShellを開き、以下を実行：

```powershell
Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy RemoteSigned
```

### ② PowerShell スクリプトを実行

```powershell
.\ImagesToPdf.ps1
```
