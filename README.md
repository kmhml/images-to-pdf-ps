# ImagesToPdf (PowerShell)

複数の PNG / JPEG 画像を 1 つの PDF にまとめる PowerShell スクリプトです。

フォルダを指定するだけで、その配下の画像を読み込み、
日付形式（yyyyMMdd）の PDF を自動生成します。

---

## 🔹 機能

- フォルダ選択ダイアログ表示
- 指定フォルダ配下の
  - `.png`
  - `.jpg`
  - `.jpeg`
  を再帰的に取得
- ファイル名順で並び替え
- 複数ページ PDF を生成
- 出力ファイル名：`yyyyMMdd.pdf`
- 画像比率を維持して中央配置

---

## 🔹 動作環境

- Windows 10 / 11
- PowerShell 5.1 以上
- Microsoft Print to PDF が有効であること

※ 外部ライブラリ不要

---

## 🔹 使い方

### 1. 実行ポリシーの設定（初回のみ）

```powershell
Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy RemoteSigned