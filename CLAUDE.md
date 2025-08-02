# CLAUDE.md

このファイルは、このリポジトリでコードを操作する際にClaude Code (claude.ai/code) にガイダンスを提供します。

## プロジェクト概要

このプロジェクトは、Visual StudioやSDKのインストール不要でC#のExcel操作アプリケーションを開発するためのフレームワークです。WindowsのcscコンパイラとDocumentFormat.OpenXmlライブラリを使用しています。

## 主要な開発コマンド

### ビルドコマンド
- **統合ビルドスクリプト**: `build.bat [debug|release]`
  - `build.bat` または `build.bat debug` - デバッグビルド（デフォルト）
  - `build.bat release` - リリースビルド（最適化）
  - `build.bat help` - ヘルプ表示
- **クリーン**: `clean.bat` - bin/ディレクトリのクリーンアップ

### セットアップコマンド
- **ライブラリセットアップ**: `setup-libraries.bat` または `setup-libraries.ps1` - NuGetパッケージの自動ダウンロードと配置

### 特殊な設定
- C# 7.0 (`/langversion:7`)を使用
- カスタムコンパイラパス: `packages\Microsoft.Net.Compilers\tools\csc.exe`
- WindowsBaseの参照パス: `WPF\WindowsBase.dll`

## アーキテクチャ

### プロジェクト構造
```
pure-csc-framework/
├── src/                    # ソースコード
│   ├── Program.cs         # メインエントリポイント - Excel操作のデモアプリ
│   └── ExcelHandler.cs    # Excel操作クラス - DocumentFormat.OpenXmlのラッパー
├── lib/                   # 実行時ライブラリ（DLL）
├── packages/              # NuGetパッケージ展開先
├── bin/                   # ビルド出力
│   ├── debug/            # デバッグビルド出力（App.exe + DLL）
│   └── release/          # リリースビルド出力（App.exe + DLL）
├── .vscode/              # VSCode設定（tasks.json, launch.json）
├── build.bat             # 統合ビルドスクリプト
└── setup-packages.ps1    # セットアップスクリプト
```

### 重要なクラス

#### ExcelHandler (`src/ExcelHandler.cs`)
DocumentFormat.OpenXmlを使用したExcel操作の中核クラス:
- `ReadExcel()` - Excelファイルの読み込み
- `WriteExcel()` - 単一シートへの書き込み
- `WriteMultipleSheets()` - 複数シートへの書き込み
- 列参照の自動計算、空セルの処理を含む

#### Program (`src/Program.cs`)
デモンストレーション用メインプログラム:
- サンプルデータ作成
- Excel読み書き操作
- データ統計処理（給与計算など）
- 複数シートサンプル生成

### 依存関係管理
- **DocumentFormat.OpenXml**: Excel操作のメインライブラリ
- **DocumentFormat.OpenXml.Framework**: フレームワーク拡張
- **System.IO.Packaging**: パッケージング操作
- PowerShellスクリプトによる自動NuGetパッケージ管理

### ビルドシステムの特徴
1. **統合ビルドスクリプト**: 1つのbuild.batでdebug/releaseを切り替え
2. **自動csc.exe検索**: システムパスと固定パスでコンパイラを探索
3. **カスタムコンパイラ**: Microsoft.Net.Compilersパッケージ内のcsc.exeを使用
4. **ポータブルPDB**: `/debug:portable`でクロスプラットフォーム対応のデバッグ情報
5. **DLL自動コピー**: ビルド後に必要なDLLを各出力ディレクトリにコピー
6. **フォルダ分離**: デバッグとリリースを`bin/debug/`と`bin/release/`に分離
7. **統一実行ファイル名**: 両方とも`App.exe`

### VSCode統合
- **tasks.json**: 統合ビルドタスクの定義（debug/release/clean/setup）
- **launch.json**: デバッグ設定（CLR）- フォルダ分離対応、実行時cwdを実行ファイルのディレクトリに設定
- IntelliSenseとデバッグ機能のサポート

## 開発時の注意点

- ビルド前に`setup-libraries.bat`の実行が必要
- 新しい統合ビルドシステムを使用（`build.bat debug` / `build.bat release`）
- WindowsBaseの参照パスが特殊（`WPF\WindowsBase.dll`）
- C# 7.0の機能を使用可能
- デバッグビルドではポータブルPDBが生成される
- lib/ディレクトリに必要なDLLが配置されていることを確認
- 出力ファイル: `bin/debug/App.exe` または `bin/release/App.exe`
- デバッグ実行時のカレントディレクトリ: 実行ファイルと同じディレクトリ（DLLが同居）

## トラブルシューティング

### よくある問題
1. **csc.exeが見つからない**: .NET Frameworkのインストール確認
2. **DocumentFormat.OpenXml.dllが見つからない**: `setup-libraries.bat`の実行
3. **WindowsBase.dllエラー**: パスの確認（WPF/WindowsBase.dll）

### デバッグ方法
- VSCodeでF5キーによるデバッグ実行
- ブレークポイント設置可能
- 変数監視とステップ実行をサポート