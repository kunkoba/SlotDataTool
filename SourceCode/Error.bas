Attribute VB_Name = "Error"
Option Compare Database
Option Explicit


Public ErrObj As ErrorClass  'エラー保持変数
Public Const ERR_TMP As Long = 50000
Public Const ERR_BIZ As Long = 60000

' ==============================================================
' エラー情報を保持する構造体
' ==============================================================
Public Type ErrorClass

    Number As Long
    Description As String
    Source As String
    val1 As String
    val2 As String
    val3 As String
    
End Type


' ==============================================================
' エラーを保持
' ==============================================================
Public Sub ErrorClear()

    ErrObj.Number = 0
    ErrObj.Description = ""
    ErrObj.Source = ""
    ErrObj.val1 = ""
    ErrObj.val2 = ""
    ErrObj.val3 = ""
    
    Err.Clear
    
End Sub

' ==============================================================
' エラーを保持
' ==============================================================
Public Sub ErrorSave(methodName As String, Optional isMsg As Boolean = False, _
            Optional val1 As String = "", Optional val2 As String = "", Optional val3 As String = "")
            
    If Err.Number <> 0 And Err.Number <> ERR_TMP Then
        ErrObj.Number = Err.Number
        ErrObj.Description = Err.Description
        ErrObj.Source = methodName
        ErrObj.val1 = val1
        ErrObj.val2 = val2
        ErrObj.val3 = val3
        
        Call LogError(methodName)
        If val1 <> "" Then Call LogError(val1)
        If val2 <> "" Then Call LogError(val2)
        If val3 <> "" Then Call LogError(val3)
        
        Err.Clear
    End If
    
    If isMsg Then Call ErrorMessage
    
End Sub


' ==============================================================
' エラーメッセージ出力
' ==============================================================
Sub ErrorMessage()
    Dim msg As String
    
    If ErrObj.Number <> 0 Then
    
        msg = JoinText(msg, "以下のエラー情報を開発者へ連絡してください。", True)
        msg = JoinText(msg, "　・メソッド：" & ErrObj.Source, True)
        msg = JoinText(msg, "　・エラー番号：" & ErrObj.Number, True)
        msg = JoinText(msg, "　・エラー内容：" & ErrObj.Description, True)
        
        Call ErrorClear
        
'        MsgBox msg, vbCritical
        Call ShowInfomation("エラー情報", msg)
        
        
    End If
End Sub


