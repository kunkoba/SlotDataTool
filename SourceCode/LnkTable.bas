Attribute VB_Name = "LnkTable"
Option Compare Database
Option Explicit


' **************************************************
' リンクテーブルの接続先を返却（接続先が一つのみ）
' **************************************************
Public Function Procリンクテーブル接続先取得() As String
    Const methodName As String = "Procリンクテーブル接続先取得"
    ' 宣言部
    Dim db As DAO.Database
    Dim tdf As DAO.TableDef
    Dim conn As String
    Dim pos As Long
On Error GoTo ErrHandler
    
    Procリンクテーブル接続先取得 = ""
    
    ' メイン処理
    Set db = CurrentDb
    
    For Each tdf In db.TableDefs
        If Len(tdf.Connect) > 0 Then
            conn = tdf.Connect
            
            ' ";DATABASE=" を探してパス部分だけを抜き出す
            pos = InStr(conn, ";DATABASE=")
            If pos > 0 Then
                Procリンクテーブル接続先取得 = Mid(conn, pos + Len(";DATABASE="))
            Else
                Procリンクテーブル接続先取得 = conn
            End If
            
            Exit For
            
        End If
    Next

ErrHandler:
    Call ErrorSave(methodName)  '必ず先頭（発生したエラーを確実にキャッチするため）
    Call ProcNothing(tdf)
    Call ProcNothing(db)
    If ErrObj.Number <> 0 Then Err.Raise ERR_TMP, "", "疑似エラー"  ' 必ず最後(後続処理が動かないため）
        
End Function


' **************************************************
' リンクテーブルの接続先（DATABASE= の値）だけ置換する（ファイル名のみ）
' **************************************************
Public Sub Procリンクテーブル一括更新(newFileName As String)
    Const methodName As String = "Procリンクテーブル一括更新"
    ' 宣言部
    Dim db As DAO.Database
    Dim tdf As DAO.TableDef
    Dim conn As String
    Dim posDB As Long, startVal As Long, nextSemi As Long
    Dim oldDBValue As String, folderPath As String, tail As String
    Dim candidateNewDBValue As String, newConn As String, oldConn As String
    Dim lastErr As Long, lastDesc As String
    Dim isFullPath As Boolean
On Error GoTo ErrHandler
    
    
    ' メイン処理
    Set db = CurrentDb

    ' --- newFileName がフルパスかどうかを 1回だけ判定 ---
    isFullPath = (InStr(newFileName, "\") > 0)

    For Each tdf In db.TableDefs
        ' リンクテーブルのみ対象
        If Len(Trim$(tdf.Connect & "")) > 0 Then
            conn = tdf.Connect
            posDB = InStr(1, conn, "DATABASE=", vbTextCompare)
            If posDB > 0 Then
                startVal = posDB + Len("DATABASE=")
                nextSemi = InStr(startVal, conn, ";")
                If nextSemi > 0 Then
                    oldDBValue = Mid(conn, startVal, nextSemi - startVal)
                    tail = Mid(conn, nextSemi) ' ;以降
                Else
                    oldDBValue = Mid(conn, startVal)
                    tail = ""
                End If

                ' 元の接続文字列のフォルダ部分
                If InStrRev(oldDBValue, "\") > 0 Then
                    folderPath = Left(oldDBValue, InStrRev(oldDBValue, "\"))
                Else
                    folderPath = ""
                End If

                ' candidateNewDBValue を作成
                If isFullPath Then
                    candidateNewDBValue = newFileName
                ElseIf folderPath <> "" Then
                    candidateNewDBValue = folderPath & newFileName
                Else
                    candidateNewDBValue = newFileName
                End If

                ' ファイル存在チェック
                If Len(Dir(candidateNewDBValue)) = 0 Then
                    Call LogDebug(methodName, "リンク先ファイルが見つかりません。スキップ: " & candidateNewDBValue)
                    GoTo ContinueNext
                End If

                ' 新しい接続文字列を作る（DATABASE= の値だけ差し替え）
                newConn = Left(conn, startVal - 1) & candidateNewDBValue & tail
                oldConn = conn

                ' 実際に更新して再リンク。失敗したらロールバック
                On Error Resume Next
                tdf.Connect = newConn
                tdf.RefreshLink
                If Err.Number <> 0 Then
                    lastErr = Err.Number
                    lastDesc = Err.Description
                    Err.Clear
                    ' ロールバック
                    tdf.Connect = oldConn
                    On Error Resume Next
                    tdf.RefreshLink
                    On Error GoTo ErrHandler
                    Call LogDebug(methodName, "リンク更新失敗（ロールバック実施）: " & tdf.name & _
                                         " new=" & candidateNewDBValue & " err=" & lastErr & " / " & lastDesc)
                Else
                    ' 成功
                    On Error GoTo ErrHandler
'                    Call LogDebug(methodName, "リンク更新成功: " & tdf.name & " -> " & candidateNewDBValue)
                End If
            End If
        End If
ContinueNext:
    Next

ErrHandler:
    Call ErrorSave(methodName)  '必ず先頭（発生したエラーを確実にキャッチするため）
    Call ProcNothing(tdf)
    Call ProcNothing(db)
    If ErrObj.Number <> 0 Then Err.Raise ERR_TMP, "", "疑似エラー"  ' 必ず最後(後続処理が動かないため）
    
End Sub


' **************************************************
' 最適化（リンクファイル or 自ファイル）
' **************************************************
Public Sub Procファイル最適化(Optional srcPath As String = "")
    Const methodName As String = "Procファイル最適化"
    ' 宣言部
    Dim tmpPath As String
On Error GoTo ErrHandler
    
    
    ' パスが指定されていなければ自ファイルを対象
    If Len(srcPath) = 0 Then
        srcPath = Procリンクテーブル接続先取得
        If srcPath = "" Then GoTo ErrHandler
    End If
    
    ' 一時ファイル名
    tmpPath = Left(srcPath, InStrRev(srcPath, ".")) & "_最適化中.accdb"
    
    ' CompactDatabase 実行
    DBEngine.CompactDatabase srcPath, tmpPath
    
    ' 元ファイルを削除
    Kill srcPath
    ' 最適化済みのファイルを元の名前に戻す
    Name tmpPath As srcPath
    
    Call ShowToast("データファイルの最適化は完了しました。", E_Color_Blue)

ErrHandler:
    Call ErrorSave(methodName)  '必ず先頭（発生したエラーを確実にキャッチするため）
    If ErrObj.Number <> 0 Then Err.Raise ERR_TMP, "", "疑似エラー"  ' 必ず最後(後続処理が動かないため）
    
End Sub




