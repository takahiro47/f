' メインクラス
Option Explicit

'*** Class Variables ***

Private logger As New ClassLogger
Private configStore As New ClassConfigStore

' ログ
Private cName As String

' 環境変数
Private conf As Object ' 環境設定


' *** Common Methods ***

' 半角英数字を取り出して返す
Private Function samplingHalfWidthAlphanumeric(ByVal target As String) As String
  Dim regExp
  Dim matchItems
  Dim strPattern As String
  Dim idx As Integer

  If target <> "" Then
    Set regExp = CreateObject("VBScript.RegExp")
    strPattern = "[0-9A-Z]"
    With regExp
      .Pattern = strPattern
      .IgnoreCase = True
      .Global = True
      Set matchItems = .Execute(StrConv(target, vbNarrow))
      If matchItems.Count > 0 Then
        For idx = 0 To matchItems.Count - 1
          samplingHalfWidthAlphanumeric = _
            samplingHalfWidthAlphanumeric + matchItems(idx).Value
        Next idx
      End If
    End With
    Set regExp = Nothing
  End If
End Function


' *** Methods ***

' コンストラクタ
Private Sub UserForm_Initialize()
  ' ログの初期化
  cName = "[MainApp] "
  Call logger.info(cName + "UserForm_Initialize()")

  ' 変数の初期化
  Call initVariables()
End Sub

' デストラクタ(クローズ時処理)
Sub closeUserInterface()
  Call logger(cName + "Window closed.")
  Unload Me
End Sub

Sub pageSource_initModuleList()
  Call logger.info(cName + "pageSource_initModuleList()")
  On Error GoTo ErrHandler

ErrHandler:
  Call logger.error(cName, Err, False)
End Sub

' UI設定状況と変数の初期化
Sub initVariables()
  Call logger.info(cName + "initVariables()")

  ' 環境変数
  Set conf = CreateObject("Scripting.Dictionary")
  Set conf = configStore.getConfigurations()

  ' 環境変数の取り出し
  filepath = conf.Item("path_baee_development") + conf.Item("path_ipo_modules_production") + "listD.txt"

End Sub
