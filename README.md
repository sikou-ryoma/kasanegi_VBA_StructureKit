# kasanegi - VBA Structure Kit

**kasanegi** は、Excelを基盤としたVBAプロジェクトを「重ねるように」構築・拡張・整理するための構造キットです。  
モジュールを積み重ねる設計思想をベースに、柔軟かつ拡張性の高いフレームワークとして機能します。

**kasanegi** is a modular VBA framework designed to layer, extend, and organize Excel-based solutions with clarity and structure.  
Inspired by the Japanese concept of “重ね着” (kasanegi – layering), this kit provides a flexible foundation for scalable, customizable VBA projects.


## 特長

- **設定の一元管理**: ロガー設定やアプリ設定をXMLファイルで柔軟に管理可能。コード変更なしで設定を調整できます。
- **モジュールの分離設計**: ConfigManager, LoggerManager, AppConfig などの専用モジュールで責任分担を明確化し、保守性を向上。
- **効率的なデバッグとログ管理**: レベル別ログ出力とファイル管理により、デバッグを支援。
- **フレームワークとしての拡張性**: VBA開発をフレームワーク化し、ブック・シート・プロセスなどのオブジェクトを一元管理。柔軟なカスタマイズと開発を可能に。
- **現場活用実績**: 実際の業務現場で活用されており、安定した運用が確認されています。

## インストール方法

- ~~releaseよりzipファイルをダウンロードして任意のフォルダに展開してください。~~
- ~~[最新リリースはこちら] ()~~

## 使い方

付属のドキュメントや各モジュールのサンプルコードやコメントを参照してください。

## フォルダ構成 (Release版)

```
ProjectTemplate/                    '--- プロジェクトルートフォルダ
├── archive/                        '--- 旧バージョン管理用フォルダ
├── build/                          '--- リリース版管理フォルダ
├── config/                         '--- 設定ファイル用フォルダ
│   └── log_config.ini              '--- ロガークラス用設定ファイル
├── log/                            '--- ログファイル用フォルダ
├── output/
│   ├── reports/                    '--- 出力ファイルの一時保存用フォルダ                    
│   └── temp/                       '--- 処理用テンプレートファイルの一時保存用フォルダ
├── src/
│   ├── Classes/                    '--- クラスモジュールのソースコードファイル用フォルダ
│   ├── Forms/                      '--- ユーザーフォームの各種ファイル用フォルダ
│   └── Modules/                    '--- 標準モジュールのソースコードファイル用フォルダ
├── test/
│   └── SampleData/                 '--- 開発に使用するテスト用ブック等の格納用フォルダ
├── workbook/
│   ├── format/                     '--- 処理に使用するフォーマットファイルを格納用フォルダ
│   └── MacroBookTemplate.xlsm      '--- プロジェクト用マクロ有効ブック
└── README.md
```

## ライセンス

MITライセンスのもとで公開しています。詳細は `LICENSE` ファイルをご確認ください。

## 貢献

バグ報告・機能追加の提案・プルリクエスト歓迎します。

---
