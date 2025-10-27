Attribute VB_Name = "File"
Option Compare Database
Option Explicit


' ***********************************************
' フォルダをダイアログで指定する
' ***********************************************
Function Dialogフォルダ選択(Optional ByVal initialFolder As String = "") As String
    Const methodName As String = "Dialogフォルダ選択"
    ' 宣言部
    Dim fd As FileDialog
    Dim selectedPath As String
On Error GoTo ErrHandler
    
    ' メイン処理
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)

    With fd
        .title = "フォルダを選択してください"
        .AllowMultiSelect = False

        ' 初期フォルダが指定されていれば設定する
        If initialFolder <> "" Then
            .InitialFileName = initialFolder
        End If

        If .Show = -1 Then
            selectedPath = .SelectedItems(1) & "\"
        Else
            selectedPath = ""
        End If
    End With

    Set fd = Nothing

    Dialogフォルダ選択 = selectedPath
    
ErrHandler:
    Call ErrorSave(methodName)  '必ず先頭（発生したエラーを確実にキャッチするため）
    Call ProcNothing(fd)
    If ErrObj.Number <> 0 Then Err.Raise ERR_TMP, "", "疑似エラー"  ' 必ず最後(後続処理が動かないため）
        
End Function


' ***********************************************
' ファイルをダイアログで指定する
' ***********************************************
Function Dialogファイル選択(Optional ByVal initialFolder As String = "", _
            Optional filterName As String = "すべてのファイル", Optional filter As String = "*.*") As String
    Const methodName As String = "Dialogファイル選択"
    ' 宣言部
    Dim fd As FileDialog
    Dim selectedPath As String
    Dim desc As String
    Dim pattern As String
On Error GoTo ErrHandler
    Dialogファイル選択 = ""
    
    ' メイン処理
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    
    With fd
        .title = "ファイルを選択してください"
        .AllowMultiSelect = False
    
        ' 拡張子フィルターを設定（例：.slotファイル）
        .Filters.Clear
        
        If filter <> "" Then
            ' 例: "スロットデータ (*.slot)|*.slot"
            If InStr(filter, "|") > 0 Then
                desc = Split(filter, "|")(0)
                pattern = Split(filter, "|")(1)
            Else
                desc = "指定なし"
                pattern = filter
            End If
            .Filters.Add filterName, pattern
        Else
            ' 常に「すべてのファイル」も追加
            .Filters.Add "すべてのファイル", "*.*"
        End If
        
        ' 初期フォルダを設定
        If initialFolder <> "" Then
            .InitialFileName = initialFolder
        End If
    
        If .Show = -1 Then
            selectedPath = .SelectedItems(1)
        Else
            selectedPath = ""  ' キャンセル時は空文字
        End If
    End With

    Set fd = Nothing

    Dialogファイル選択 = selectedPath

ErrHandler:
    Call ErrorSave(methodName)  '必ず先頭（発生したエラーを確実にキャッチするため）
    If ErrObj.Number <> 0 Then Err.Raise ERR_TMP, "", "疑似エラー"  ' 必ず最後(後続処理が動かないため）
    
End Function


' ***********************************************
' フォルダ・ファイル名・拡張子を分割して配列で返す
' result(0): フォルダ
' result(1): ファイル名（拡張子なし）
' result(2): 拡張子
' ***********************************************
Function Getファイルパス分割配列(fullPath As String) As Variant
    Const methodName As String = "Getファイルパス分割配列"
    Dim pos As Long
    Dim dotPos As Long
    Dim folderPath As String
    Dim fileName As String
    Dim ext As String
    Dim result(2) As String
On Error GoTo ErrHandler
    Getファイルパス分割配列 = Null

    ' --- フォルダ部分とファイル部分を分ける ---
    pos = InStrRev(fullPath, "\")
    If pos > 0 Then
        folderPath = Left(fullPath, pos)   ' "\" を含む
        fileName = Mid(fullPath, pos + 1)
    Else
        folderPath = ""
        fileName = fullPath
    End If
    
    ' --- 拡張子を分ける ---
    dotPos = InStrRev(fileName, ".")
    If dotPos > 0 Then
        ext = Mid(fileName, dotPos)
        fileName = Left(fileName, dotPos - 1)
    Else
        ext = ""
    End If

    ' --- 配列に格納 ---
    result(0) = folderPath
    result(1) = fileName
    result(2) = ext
    
    Getファイルパス分割配列 = result

ErrHandler:
    Call ErrorSave(methodName)  '必ず先頭（発生したエラーを確実にキャッチするため）
    If ErrObj.Number <> 0 Then Err.Raise ERR_TMP, "", "疑似エラー"  ' 必ず最後(後続処理が動かないため）
    
End Function


' ***********************************************
' フォルダを指定してファイルを取得する（文字列配列）
' ***********************************************
Function Getファイル取得配列(folderPath As String, filterPattern As String) As String()
    Const methodName As String = "Getファイル取得配列"
    ' 宣言部
    Dim fileName As String
    Dim fileList() As String
    Dim count As Long
On Error GoTo ErrHandler
    
    ' フォルダ末尾の \ を補完
    If Right(folderPath, 1) <> "\" Then
        folderPath = folderPath & "\"
    End If

    ' 初期化
    count = 0
    fileName = Dir(folderPath & filterPattern)
    
    Do While fileName <> ""
        ReDim Preserve fileList(count)
        fileList(count) = fileName ' ← フルパスではなくファイル名のみ
        count = count + 1
        fileName = Dir()
    Loop
    
    ' ファイルが見つからなければ空配列を返す
    If count = 0 Then
        Getファイル取得配列 = Split("") ' 長さ0の配列を返す
    Else
        Getファイル取得配列 = fileList
    End If

ErrHandler:
    Call ErrorSave(methodName)  '必ず先頭（発生したエラーを確実にキャッチするため）
    If ErrObj.Number <> 0 Then Err.Raise ERR_TMP, "", "疑似エラー"  ' 必ず最後(後続処理が動かないため）
    
End Function


' ***********************************************
' ファイル名のみ取得する
' ***********************************************
Function GetFileName(fullPath As String) As String
    Const methodName As String = "GetFileName"
    Dim fso As Object
On Error GoTo ErrHandler

    GetFileName = ""
    
    ' メイン処理
    Set fso = CreateObject("Scripting.FileSystemObject")
    GetFileName = fso.GetFileName(fullPath)
    
ErrHandler:
    Call ErrorSave(methodName)  '必ず先頭（発生したエラーを確実にキャッチするため）
    Call ProcNothing(fso)
    If ErrObj.Number <> 0 Then Err.Raise ERR_TMP, "", "疑似エラー"  ' 必ず最後(後続処理が動かないため）
        
End Function


' **************************************************
' リンク先ファイルをコピー（オプションでファイル名を変更）
' **************************************************
Public Sub Cmdファイルコピー(srcPath As String, destFolder As String, Optional newFileName As String = "")
    Const methodName As String = "Cmdファイルコピー"
    On Error GoTo ErrHandler
    
    Dim destPath As String
    
    ' コピー先の完全パスを作成
    If Right(destFolder, 1) <> "\" Then
        destFolder = destFolder & "\"
    End If
    
    ' 新しいファイル名が指定されている場合
    If newFileName <> "" Then
        destPath = destFolder & newFileName
    Else
        destPath = destFolder & Dir(srcPath)
    End If
    
'    ' 既に同名ファイルが存在する場合は削除
'    If Dir(destPath) <> "" Then
'        Kill destPath
'    End If
    
    ' ファイルコピー
    FileCopy srcPath, destPath
    
    Exit Sub  ' 正常終了時はここで抜ける

ErrHandler:
    Call ErrorSave(methodName)  ' エラーを記録
    If Err.Number <> 0 Then Err.Raise vbObjectError + 1000, , "ファイルコピー中にエラーが発生しました"

End Sub





' ***********************************************
' バッチ実行関数（非同期）
' ***********************************************
Public Sub Runバッチプログラム実行(vbsPath As String)
    Const methodName As String = "Runバッチプログラム実行"
On Error Resume Next

    ' メイン処理
    shell "cscript //nologo """ & vbsPath & """", vbNormalFocus
    
ErrHandler:
    Call ErrorSave(methodName)  '必ず先頭（発生したエラーを確実にキャッチするため）
    If ErrObj.Number <> 0 Then Err.Raise ERR_TMP, "", "疑似エラー"  ' 必ず最後(後続処理が動かないため）
    
End Sub

' ***********************************************
' バッチ実行関数（同期）
' ***********************************************
Public Sub Runバッチプログラム実行AndWait(vbsPath As String)
    Const methodName As String = "Runバッチプログラム実行AndWait"
    Dim wsh As Object
On Error GoTo ErrHandler

    ' メイン処理
    Set wsh = CreateObject("WScript.Shell")
    
    wsh.Run "cscript //nologo """ & vbsPath & """", 1, True  ' True = 完了まで待機
    
ErrHandler:
    Call ErrorSave(methodName)  '必ず先頭（発生したエラーを確実にキャッチするため）
    Call ProcNothing(wsh)
    If ErrObj.Number <> 0 Then Err.Raise ERR_TMP, "", "疑似エラー"  ' 必ず最後(後続処理が動かないため）
    
End Sub



' **************************************************
' フォルダが存在しない場合は作成する
' **************************************************
Public Sub Procフォルダが存在しない場合は作成(folderPath As String)
    Const methodName As String = "Procフォルダが存在しない場合は作成"
    Dim fso As Object
On Error GoTo ErrHandler

    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(folderPath) Then
        fso.CreateFolder folderPath
    End If
    
ErrHandler:
    Call ErrorSave(methodName)  '必ず先頭（発生したエラーを確実にキャッチするため）
    Call ProcNothing(fso)
    If ErrObj.Number <> 0 Then Err.Raise ERR_TMP, "", "疑似エラー"  ' 必ず最後(後続処理が動かないため）

End Sub


' **************************************************
' ファイルが存在していれば True を返す
' **************************************************
Public Function Procファイル存在確認(filePath As String) As Boolean
    Dim fso As Object
On Error Resume Next
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Procファイル存在確認 = fso.FileExists(filePath)

    Call ProcNothing(fso)
    
End Function

' ***********************************************
' ファイルをコピーする（file → folder）
' ***********************************************
Public Sub Procファイルを指定フォルダへコピー(srcFile As String, destFolder As String)
    Const methodName As String = "Procファイルを指定フォルダへコピー"
    Dim fso As Object
    Dim fileName As String
    Dim destFile As String
On Error GoTo ErrHandler

    ' ファイル名抽出
    fileName = Mid(srcFile, InStrRev(srcFile, "\") + 1)

    ' フォルダ末尾補完
    If Right(destFolder, 1) <> "\" Then destFolder = destFolder & "\"

    ' コピー先ファイルパス構築
    destFile = destFolder & fileName

    ' コピー実行
    Set fso = CreateObject("Scripting.FileSystemObject")
    fso.CopyFile srcFile, destFile, True

ErrHandler:
    Call ErrorSave(methodName)
    Call ProcNothing(fso)
    If ErrObj.Number <> 0 Then Err.Raise ERR_TMP, "", "疑似エラー"
    
End Sub



