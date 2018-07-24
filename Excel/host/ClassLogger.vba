' 実行ログの管理とファイル出力を行うクラス
Option Explicit

' *** クラス変数 ***

' 定数
Private Const LOG_FILE_PREFIX As String = "logs¥LOG_"
Private Const LOG_FILE_EXTENSION As String = ".log"
Private Const LOG_BUFFER_SIZE As Integer = 1024

' ファイルシステム
Public loggingDirectory As String
Private file As Object
Private Enum IOMode '
  forRead = 0 'Default
  forWrite = 1
  for_Appending = 8
End Enum
Private buffer As New Collection


' *** コンストラクタ ***

Private Sub Class_Initialize()
  Set file = CreateObject("Scripting.FileSystemObject")
  Call clearBuffer
  Call info("*** ----------------------------------------")
  Call info("[ClassLogger] initialized.")
  Call flush
End Sub

Private Sub Class_Terminate()
  Call info("[ClassLogger] terminated.")
  Call flush
  Set file = Nothing
End Sub


' *** アクセッサ ***

' Information
Friend Sub info(ByVal message As String, _
  Optional ByVal useBuffer As Boolean = False)

  Call log("INFO", message, useBuffer)
End Sub

' Warning
Friend Sub warn(ByVal message As String, _
  Optional ByVal useBuffer As Boolean = False)

  Call log("WARN", message, useBuffer)
End Sub

' Error
Friend Sub error(ByVal message As String, _
  Optional ByVal errorObject As ErrObject = Nothing, _
  Optional ByVal useBuffer As Boolean = False)

  ' Call log("DEBG", errorObject, False) 'ERROR?
  ' Call log("DEBG", TypeName(errorObject), False)
  ' Call log("DEBG", errorObject.Number, False)
  ' Call log("DEBG", TypeName(errorObject.Number), False)
  ' Call log("DEBG", errorObject.Description, False)
  ' Call log("DEBG", TypeName(errorObject.Description), False)

  Dim messageStr As String
  messageStr = message
  If Not (errorObject Is Nothing) Then
    ' messageStr = messageStr + "" + _
    '   "ErrNumber[" + errorObject.Number + "]" + _
    '   "ErrDescription[" + errorObject.Description + "]"
  End If

  Call log("EROR", messageStr, useBuffer)
End Sub


' *** 関数 ***

' ログをファイルに書き出す
Private Sub log(ByVal logType As String, _
  ByVal message As String, _
  ByVal useBuffer As Boolean)

  Dim messageStr As String
  messageStr = Format(Now, "yyyy-mm-dd hh:mm:ss") + " +900: [" + logType + "] " + message

  ' バッファにログを書き込む
  Debug.Print messageStr
  buffer.Add messageStr

  ' バッファが有効な場合は、バッファが指定サイズを超えた場合のみ書き出す
  If useBuffer = True Then
    If buffer.Count > LOG_BUFFER_SIZE Then
      Me.flush
    End If
  ' バッファが無効な場合は、逐次書き出す
  Else
    Me.flush
  End If
End Sub

' バッファをファイルに書き出す
Friend Sub flush()
  Dim textStream As Object ' TextStreamObject
  Dim i As Long

  On Error GoTo LABEL_ERROR:

  Set textStream = file.OpenTextFile( _
    fileName:=file.BuildPath(getLoggingDir(), LOG_FILE_PREFIX + Format(Now, "yyyymmdd") + LOG_FILE_EXTENSION), _
    IOMode:=IOMode.for_Appending, _
    Create:=True)
  For i = 1 To buffer.Count
    textStream.WriteLine buffer(i)
  Next i
  textStream.Close
  Set textStream = Nothing
  Call clearBuffer

  Exit Sub

' エラー処理:書き出しに失敗した場合
LABEL_ERROR:
  IF Not (textStream Is Nothing) Then
    textStream.Close
    Set textStream = Nothing
  End If
  Call clearBuffer

  Err.Raise _
    Number:=Err.Number, _
    Description:="ログの出力に失敗しました。" + Err.Description
End Sub

' ログ出力先のディレクトリを返す
' デフォルトはThisWorkBook(VBA実行ファイルのあるディレクトリ)
Private Function getLoggingDir() As String
  If Me.loggingDirectory = "" Then
    getLoggingDir = ThisWorkbook.Path
    Exit Function
  End If

  getLoggingDir = Me.loggingDirectory
End Function

' バッファーをクリア
Private Sub clearBuffer()
  Dim i As Long
  For i = buffer.Count To 1 Step -1
    buffer.Remove i
  Next i
End Sub
