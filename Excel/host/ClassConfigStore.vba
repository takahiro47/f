' 環境変数の管理を行うクラス
Option Explicit

' *** クラス変数 ***

Private objectCofigurations As Object ' 環境設定

' *** コンストラクタ ***

Private Sub Class_Initialize()

  ' 環境設定
  Set objectCofigurations = CreateObject("Scripting.Dictionary")
  With objectCofigurations
    ' *** ディレクトリ関係 ***

    ' 一時ファイル
    .Add "path_tmp", "C:\temp¥"

    ' メモ帳
    .Add "path_program_editor_note", "C:\Windows\System32\notepad.exe"

    ' IBM MVS Time Sharing Option(TSO)
    .Add "path_tso_exe",            ""
    .Add "path_tso_profile_dir",    "TSO\" ' 構成ファイル(TSO.exeに第1引数で渡す)
    .Add "path_tso_profile_regexp", "*.ws"

    ' IBM MVS Time Sharing Option(TSO) APIs
    .Add "path_tso_receiver_exe",   "C:\Pcswin\receive"
    .Add "path_tso_receiver_param", "JISCII CRLF"


    ' *** シートの設定 ***

    .Add "font_family", "ＭＳ ゴシック"

    .Add "font_size_title", 20
    .Add "font_size_description", 11
    .Add "font_size_log", 11
  End With


' *** 定数配列の代替 ***

' 環境設定
Friend Function getConfigurations() As Object
' use "Property Get" in Office 2013 or later
  Set getConfigurations = objectCofigurations
End Function
