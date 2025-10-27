Attribute VB_Name = "Logger"
Option Compare Database
Option Explicit


Const delim As String = " >> "
Dim outputMsg As String

'********************************************************************************
' ログはあくまでエラー箇所を特定するためだけのもの
'********************************************************************************
Public Sub LogOpen(objName As String)
On Error Resume Next
    
    Call ErrorClear     ' 最初にクリア
    
    outputMsg = ""
    outputMsg = JoinText(outputMsg, "=======================================================================================", True)
    outputMsg = JoinText(outputMsg, Now & delim & "[OPEN ]" & delim & objName)
    Debug.Print outputMsg
    Call Output(outputMsg)
    
End Sub

'********************************************************************************
Public Sub LogStart(methodName As String)
On Error Resume Next
    
    Call ErrorClear     ' 最初にクリア
    
    outputMsg = ""
    outputMsg = JoinText(outputMsg, "-------------------------------------------------", True)
    outputMsg = JoinText(outputMsg, Now & delim & "[START]" & delim & methodName)
    Debug.Print outputMsg
    Call Output(outputMsg)
        
End Sub

'********************************************************************************
Public Sub LogEnd(methodName As String)
On Error Resume Next

    outputMsg = ""
    outputMsg = JoinText(outputMsg, Now & delim & "[END  ]" & delim & methodName)
    Debug.Print outputMsg
    Call Output(outputMsg)
    
    Call ErrorMessage   '出力してクリア
    
End Sub

'********************************************************************************
Public Sub LogClose(objName As String)
On Error Resume Next
    
    outputMsg = ""
    outputMsg = JoinText(outputMsg, Now & delim & "[CLOSE]" & delim & objName, True)
    outputMsg = JoinText(outputMsg, "=======================================================================================")
    Debug.Print outputMsg
    Call Output(outputMsg)
    
    Call ErrorMessage   '出力してクリア
    
End Sub

'********************************************************************************
Public Sub LogError(methodName As String)
On Error Resume Next

    outputMsg = ""
    outputMsg = JoinText(outputMsg, Now & delim & "[ ERR ]" & delim & methodName & " >>>> " & ErrObj.Number & vbTab & ErrObj.Description)
    Debug.Print outputMsg
    Call Output(outputMsg)
    
End Sub

'********************************************************************************
Public Sub LogDebug(methodName As String, ByVal log1 As String)
On Error Resume Next

    outputMsg = ""
    outputMsg = JoinText(outputMsg, Now & delim & "[DEBUG]" & delim & methodName & delim & log1)
    Debug.Print outputMsg
    Call Output(outputMsg)

End Sub


'********************************************************************************
' ログファイルに追記する（日付別ファイル）
'********************************************************************************
Private Sub Output(logText As String)
    Const methodName As String = "Output"
'    Const logFolder As String = "..\Log\"  ' App から見た相対パス
    Dim fso As Object
    Dim ts As Object
    Dim outPath As String
    Dim filaName As String
On Error Resume Next

    filaName = "log_" & Format(Date, "yyyymmdd") & ".log"
    outPath = CurrentProject.path & "\" & PATH_LOG

    Set fso = CreateObject("Scripting.FileSystemObject")
    Call Procフォルダが存在しない場合は作成(outPath)

    Set ts = fso.OpenTextFile(outPath & "\" & filaName, 8, True)  ' 8 = ForAppending, True = create if not exists
    ts.WriteLine logText

ErrHandler:
    Call ErrorSave(methodName)  '必ず先頭（発生したエラーを確実にキャッチするため）
    Call ProcNothing(ts)
    Call ProcNothing(fso)
    If ErrObj.Number <> 0 Then Err.Raise ERR_TMP, "", "疑似エラー"  ' 必ず最後(後続処理が動かないため）

End Sub



