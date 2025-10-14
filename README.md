# kasanegi - VBA Structure Kit

**kasanegi** は、Excelを基盤としたVBAプロジェクトを「重ねるように」構築・拡張・整理するための構造キットです。  
モジュールを積み重ねる設計思想をベースに、柔軟かつ拡張性の高いフレームワークとして機能します。

**kasanegi** is a modular VBA framework designed to layer, extend, and organize Excel-based solutions with clarity and structure.  
Inspired by the Japanese concept of “重ね着” (kasanegi – layering), this kit provides a flexible foundation for scalable, customizable VBA projects.


## 特長

- マクロ全体のフロー管理及びブック、シートなどのオブジェクトも一元管理
- 効率的なデバッグを行うためのロガークラスを搭載
- VBA開発を行うためのフレームワークのようなモジュール群で柔軟なカスタマイズと開発を可能に

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
