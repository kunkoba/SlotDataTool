Attribute VB_Name = "_Package"
Option Compare Database
Option Explicit




' ***********************************************
' パッケージ作成処理
' ***********************************************
Public Sub Pacパッケージ作成()
    Const methodName As String = "Pacパッケージ作成"
    Dim thisPath As String
    Dim tmpPath As String
On Error GoTo ErrHandler

    thisPath = CurrentProject.path & "\"
    
    ' zipフォルダ
    tmpPath = thisPath & PATH_BIN
    Call Procフォルダが存在しない場合は作成(tmpPath)
    
    ' accdeファイル作成
    tmpPath = tmpPath & App_バージョン
    Call Procフォルダが存在しない場合は作成(tmpPath)
    Call PacACCDEファイル作成(tmpPath)
    
    ' データファイル作成
    Call Procフォルダが存在しない場合は作成(tmpPath & PATH_DATA)
    Call Procファイルを指定フォルダへコピー(Procリンクテーブル接続先取得, tmpPath & PATH_DATA)
    
    'ログフォルダ作成
    Call Procフォルダが存在しない場合は作成(tmpPath & PATH_LOG)
    
    Call ShowToast("処理は正常に完了しました。", E_Color_Blue)

ErrHandler:
    Call ErrorSave(methodName)  '必ず先頭（発生したエラーを確実にキャッチするため）
'    If ErrObj.Number <> 0 Then Err.Raise ERR_TMP, "", "疑似エラー"  ' 必ず最後(後続処理が動かないため）

End Sub


' ***********************************************
' PacACCDEファイル作成（指定フォルダにコピー → コンパイル → コピー削除）
' ***********************************************
Public Sub PacACCDEファイル作成(destFolder As String)
    Const methodName As String = "PacACCDEファイル作成"
    Dim db As Object
    Dim parts As Variant
    Dim fileName As String
    Dim copiedFile As String
On Error GoTo ErrHandler

    ' 初期化
    Set db = CurrentDb
    parts = Getファイルパス分割配列(db.name)
    If IsNull(parts) Then GoTo ErrHandler

    ' ファイル名（拡張子付き）を構築
    fileName = parts(1) & parts(2)

    ' フォルダ末尾補完
    If Right(destFolder, 1) <> "\" Then destFolder = destFolder & "\"

    ' コピー先ファイルパス構築
    copiedFile = destFolder & fileName

    ' フォルダがなければ作成
    Call Procフォルダが存在しない場合は作成(destFolder)

    Call LogDebug(methodName, "コピー: " & copiedFile)

    ' ファイルをコピー（ファイル名変更なし）
    Call Procファイルを指定フォルダへコピー(db.name, destFolder)

    ' コンパイル（同フォルダに .accde を生成）
    Call PacACCDEコンパイル(copiedFile)

    ' コピーされた .accdb を削除
    Kill copiedFile
    Call LogDebug(methodName, "削除: " & copiedFile)

ErrHandler:
    Call ErrorSave(methodName)
    Call ProcNothing(db)
    If ErrObj.Number <> 0 Then Err.Raise ERR_TMP, "", "疑似エラー"
    
End Sub




' ***********************************************
' コンパイル
' ***********************************************
Private Sub PacACCDEコンパイル(srcPath As String)
    Const methodName As String = "PacACCDEコンパイル"
    ' 宣言部
    Dim fso As Object
    Dim Source As String
    Dim ts As Object
    Dim vbsPath As String
    Dim destPath As String
On Error GoTo ErrHandler
    
    ' コンパイル実行ファイル（VBSファイル）の生成
    vbsPath = Getファイルパス分割配列(srcPath)(0) & "build_accde.vbs"
    destPath = Replace(srcPath, ".accdb", ".accde")
    
    Source = _
        "Set acc = CreateObject(""Access.Application"")" & vbCrLf & _
        "acc.SysCmd 603, """ & srcPath & """, """ & destPath & """" & vbCrLf & _
        "acc.Quit" & vbCrLf & _
        "Set acc = Nothing"
    
    ' FileSystemObjectでファイル書き出し
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.CreateTextFile(vbsPath, True, False) ' 第3引数: Unicode=False（ANSIで保存）
    
    ts.Write Source
    ts.Close
    
    Call LogDebug(methodName, vbsPath)
    
    ' バッチ実行
    shell "wscript.exe """ & vbsPath & """", vbNormalFocus
    
    ' 元ファイル削除
    Sleep 3000  ' 3秒くらいあればコンパイル完了するでしょ
    Kill vbsPath
'    Kill srcPath

ErrHandler:
    Call ErrorSave(methodName)  '必ず先頭（発生したエラーを確実にキャッチするため）
    Call ProcNothing(ts)
    Call ProcNothing(fso)
    If ErrObj.Number <> 0 Then Err.Raise ERR_TMP, "", "疑似エラー"  ' 必ず最後(後続処理が動かないため）
        
End Sub




