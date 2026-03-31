Attribute VB_Name = "AppConfig"
'------------------------------------------------------------------------
' ---AppConfig---
' 全体の管理設定行います。
' パブリックレベルの定数、クラスの宣言など。
'------------------------------------------------------------------------
Option Explicit


' マクロ名、バージョン情報の定数宣言
'------------------------------------------------------------------------
Public Const VERSION As String = "v0.9.0"
Public Const MACRO_NAME As String = "Template_Macro"
Public Const KANRI_PASS As String = "9999"  '---IptPass用のパスワード


' パブリックレベルのクラス宣言
'------------------------------------------------------------------------
Public FO As FileOjt
Public bm As BookManager
Public sc As SheetCollection
Public du As DateUtility
Public Paths As PathConfig
Public ctx As ProcessContext


