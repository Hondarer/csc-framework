# CLAUDE.md

このファイルは、このリポジトリでコードを操作する際に Claude Code (claude.ai/code) にガイダンスを提供します。

## プロジェクト概要

このプロジェクトは、Visual Studio や SDK のインストール不要で C# のアプリケーションを開発するためのフレームワークです。  
csc.exe は、Microsoft.Net.Compilers パッケージの csc.exe を使用するため、 C# 7.0 対応です。

## 主要な開発コマンド

### ビルドコマンド

- **統合ビルドスクリプト**: `build.bat [PROJECT_NAME] [debug|release]`
  - `build.bat {プロジェクト名} debug` - {プロジェクト名}.exeでデバッグビルド
  - `build.bat {プロジェクト名} release` - {プロジェクト名}.exeでリリースビルド
- **クリーン**: `clean.bat` - bin/ディレクトリのクリーンアップ

### セットアップコマンド

- **ライブラリセットアップ**: `setup-packages.bat` または `setup-packages.ps1` - NuGet パッケージの自動ダウンロードと配置
- **個別パッケージセットアップ**: `setup-package.ps1 -PackageName <パッケージ名> [-Version <バージョン>] [-TargetFramework <フレームワーク>]` - 単一 NuGet パッケージの手動ダウンロード

## アーキテクチャ

### プロジェクト構造

```text
csc-framework/
├── src/                  # サンプルソースコード
│   ├── Program.cs       # メインエントリポイント - Excel 操作のデモアプリ
│   └── ExcelHandler.cs  # Excel 操作クラス - DocumentFormat.OpenXml のラッパー
├── lib/                  # 実行時ライブラリ (DLL) 
├── packages/             # NuGet パッケージ展開先
├── bin/                  # ビルド出力
│   ├── debug/           # デバッグビルド出力 ({プロジェクト名}.exe + DLL)
│   └── release/         # リリースビルド出力 ({プロジェクト名}.exe + DLL)
├── .vscode/              # VSCode 設定 (tasks.json, launch.json)
├── build.bat             # 統合ビルドスクリプト
└── setup-packages.ps1    # セットアップスクリプト
```

#### パッケージセットアップスクリプトの機能

- **setup-packages.ps1**: `packages.config` からの一括パッケージセットアップ
- **setup-package.ps1**: 個別パッケージセットアップ (リアルタイム進捗表示機能)
  - WebClient によるダウンロード進捗のリアルタイム表示
  - MB 単位とパーセンテージでの進捗状況表示
  - 自動バージョン検出 (latest 指定時)
  - 複数ターゲットフレームワーク対応 (net48, net472, netstandard2.0 等)
  - 依存関係の自動検出と表示

### ビルドシステムの特徴

1. **統合ビルドスクリプト**: 1 つの build.bat で debug/release を切り替え
2. **動的プロジェクト名**: VSCode のワークスペース名から実行ファイル名を自動決定
3. **自動csc.exe検索**: システムパスと固定パスでコンパイラを探索
4. **カスタムコンパイラ**: Microsoft.Net.Compilers パッケージ内の csc.exe を使用
5. **ポータブルPDB**: `/debug:portable` でクロスプラットフォーム対応のデバッグ情報
6. **DLL自動コピー**: ビルド後に必要なDLLを各出力ディレクトリにコピー
7. **フォルダ分離**: デバッグとリリースを `bin/debug/` と `bin/release/` に分離

### VSCode統合

- **tasks.json**: 統合ビルドタスクの定義 (debug/release/clean/setup) - ワークスペース名を自動的に build.bat に渡す
- **launch.json**: デバッグ設定 (CLR) - ワークスペース名に基づく動的プログラムパス、実行時 cwd を実行ファイルのディレクトリに設定
- IntelliSense とデバッグ機能のサポート

## 開発時の注意点

- ビルド前に `setup-libraries.bat` の実行が必要
- C# 7.0の機能を使用可能
- デバッグビルドではポータブル PDB が生成される
- デバッグ実行時のカレントディレクトリ: 実行ファイルと同じディレクトリ (DLL が同居)

## トラブルシューティング

### よくある問題

1. **csc.exeが見つからない**: `setup-libraries.bat` の実行
2. **DocumentFormat.OpenXml.dllが見つからない**: `setup-libraries.bat` の実行

### デバッグ方法

- VSCode で F5 キーによるデバッグ実行
- ブレークポイント設置可能
- 変数監視とステップ実行をサポート
