Attribute VB_Name = "_AutoExe"
Option Compare Database
Option Explicit

'--- Access本体ウィンドウを隠す ---
Private Declare PtrSafe Function apiShowWindow Lib "user32" Alias "ShowWindow" ( _
    ByVal hwnd As LongPtr, ByVal nCmdShow As Long) As Long

Private Declare PtrSafe Function apiFindWindowA Lib "user32" Alias "FindWindowA" ( _
    ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr


' *********************************************************************
' AutoExe
' *********************************************************************
Public Function Procアプリ起動初期処理(Optional is非表示 As Boolean = True)
    Const methodName As String = "Procアプリ起動初期処理"
    ' 宣言部
On Error GoTo ErrHandler
    
    Call LogOpen(methodName)
    
    ' 二重起動チェック
    If Is二重起動チェック() Then
        MsgBox "すでに起動しています。二重起動はできません。", vbExclamation
        Application.Quit
        Exit Function
    End If
    
'    Application.Echo False

    If is非表示 Then
        ' Accessウィンドウを非表示に
        Dim hwnd As LongPtr
        hwnd = apiFindWindowA("OMain", vbNullString)
        If hwnd <> 0 Then
            Call apiShowWindow(hwnd, 0) ' SW_HIDE = 0
        End If
        
    End If

    flg最適化 = False   '最適化フラグ
    
    ' メニュー画面表示前処理
    If IsNull(App_アプリ有効期限) Then
        Call Setアプリ起動時設定更新
        Call MACアドレス更新
    End If
    
    ' メニュー画面表示
    DoCmd.OpenForm F11_MainMenu
    
    ' 接続先チェック
    If Not App_接続チェック Then
        ' リンク接続
        DoCmd.OpenForm F16_LinkManager
    End If

ErrHandler:
    Call ErrorSave(methodName, True) '必ず先頭（発生したエラーを確実にキャッチするため）
'    Application.Echo True
    Call LogClose(methodName)
        
End Function


' *********************************************************************
' 二重起動チェック（同一ファイル）
' *********************************************************************
Public Function Is二重起動チェック() As Boolean
    Const methodName As String = "Is二重起動チェック"
    ' 宣言部
    Dim objWMI As Object
    Dim colProcesses As Object
    Dim objProcess As Object
    Dim myPath As String
    Dim cmdLine As String
    Dim count As Integer
On Error GoTo ErrHandler
    
    ' メイン処理
    myPath = """" & CurrentDb.name & """" ' フルパスを囲んで比較（スペース対策）
    count = 0

    On Error Resume Next
        Set objWMI = GetObject("winmgmts:\\.\root\CIMV2")
        If objWMI Is Nothing Then
            Is二重起動チェック = False
            Exit Function
        End If
    
        Set colProcesses = objWMI.ExecQuery("SELECT * FROM Win32_Process WHERE Name = 'MSACCESS.EXE'")
    On Error GoTo ErrHandler

    For Each objProcess In colProcesses
        cmdLine = objProcess.CommandLine
        If InStr(cmdLine, myPath) > 0 Then
            count = count + 1
        End If
    Next

    ' 自分も含めて2つ以上あれば同一ファイルが多重起動中
    Is二重起動チェック = (count >= 2)

ErrHandler:
    Call ErrorSave(methodName)  '必ず先頭（発生したエラーを確実にキャッチするため）
    If ErrObj.Number <> 0 Then Err.Raise ERR_TMP, "", "疑似エラー"  ' 必ず最後(後続処理が動かないため）
    
End Function

' ======================================================
'■ 「S_アプリ設定」の指定項目を設定する
' ======================================================
Public Function Setアプリ起動時設定更新() As Boolean
    Const methodName As String = "Setアプリ起動時設定更新"
    ' 宣言部
    Dim strSQL As String
    Dim key As String
On Error GoTo ErrHandler

    Setアプリ起動時設定更新 = False
    
    ' メイン処理
    key = Lic秘密コード生成
    
    ' SQL作成
    strSQL = "UPDATE S_アプリ設定 " & _
             "SET " & _
             " アプリ起動日 = #" & Date & "#" & _
             ",アプリ有効期限 = #" & Date + 7 & "#" & _
             ",秘密コード = '" & key & "'" & _
             ",暗号キー = '" & Lic秘密コードから暗号キー(key) & "'"
    
    ' SQL実行
    CurrentDb.Execute strSQL, dbFailOnError
    
    Setアプリ起動時設定更新 = True

ErrHandler:
    Call ErrorSave(methodName)  '必ず先頭（発生したエラーを確実にキャッチするため）
'    Call ProcNothing(db)
    If ErrObj.Number <> 0 Then Err.Raise ERR_TMP, "", "疑似エラー"  ' 必ず最後(後続処理が動かないため）

End Function





