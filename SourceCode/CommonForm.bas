Attribute VB_Name = "CommonForm"
Option Compare Database
Option Explicit


' ======================================================
Public Sub Procクリップボードへ転送(ctrl As TextBox)
    Const methodName As String = "Procクリップボードへ転送"
On Error GoTo ErrHandler

    ' メイン処理
    With ctrl
        ' コピーしたいコントロール（例：テキスト0）にフォーカスを移動
        .SetFocus
        
        ' テキストボックス内の内容全体を選択
        .SelStart = 0
        .SelLength = Len(.text)
        
        DoCmd.RunCommand acCmdCopy
    End With
    
ErrHandler:
    Call ErrorSave(methodName)  '必ず先頭（発生したエラーを確実にキャッチするため）
    If ErrObj.Number <> 0 Then Err.Raise ERR_TMP, "", "疑似エラー"  ' 必ず最後(後続処理が動かないため）
    
End Sub



'' ******************************************************
'' コンボボックスの選択を移動させる（上下移動、見出し対応）
'' ******************************************************
'Public Sub changeCombbox(cmb As ComboBox, addValue As Long, Optional isHeader As Boolean = True)
'    Const methodName As String = "changeCombbox"
'    ' 宣言部
'    Dim idx As Long
'    Dim newIdx As Long
'    Dim offset As Long
'On Error GoTo ErrHandler
'
'    ' メイン処理
'    offset = IIf(isHeader, 1, 0)
'
'    ' 現在のインデックス
'    idx = cmb.ListIndex + offset
'    If idx < 0 Then Exit Sub  ' 未選択
'
'    newIdx = idx + addValue
'
'    ' 範囲チェック（ヘッダー考慮）
'    If newIdx < offset Then Exit Sub
'    If newIdx >= cmb.ListCount Then Exit Sub
'
'    ' 選択変更
'    cmb.value = cmb.ItemData(newIdx)
'
'ErrHandler:
'    ' エラー
'    If Err.Number <> 0 Then
'        Call LogError(methodName)
'        Err.Raise Err.Number, Err.source, Err.Description
'    End If
'ExitHandler:
'    ' 終了処理
''
'
'End Sub



' **************************************************
' サブフォーム更新（Filter / OrderBy を保持）
' subFormControl: サブフォームコントロール名（文字列）
' **************************************************
Public Sub Procフィルタ情報を保持したままサブフォーム更新(frm As Form, subFormControl As String)
    Const methodName As String = "Procフィルタ情報を保持したままサブフォーム更新"
    ' 宣言部
    Dim sf As subForm
    Dim ctrl As subForm
    Dim currentFilter As String
    Dim currentOrder As String
On Error GoTo ErrHandler
    
    ' サブフォームコントロールを取得
    Set ctrl = frm.Controls(subFormControl)
    
    ' 現在のフィルタ・オーダーを退避
    currentFilter = ctrl.Form.filter
    currentOrder = ctrl.Form.OrderBy
    
    ' SourceObject を再設定（更新）
    ctrl.SourceObject = ctrl.SourceObject
    
    ' フィルタ・オーダーを復元
    ctrl.Form.filter = currentFilter
    ctrl.Form.OrderBy = currentOrder
    
    ' フィルタ・オーダーの有効状態も復元
    If currentFilter <> "" Then ctrl.Form.FilterOn = True
    If currentOrder <> "" Then ctrl.Form.OrderByOn = True

ErrHandler:
    Call ErrorSave(methodName)  '必ず先頭（発生したエラーを確実にキャッチするため）
    If ErrObj.Number <> 0 Then Err.Raise ERR_TMP, "", "疑似エラー"  ' 必ず最後(後続処理が動かないため）
    
End Sub


' ******************************************************
' リストボックスの値ソースを一元配列に変換する
' ******************************************************
Function Procリストボックスの値取得配列(lst As ListBox, Optional delimiter As String = ";") As Variant
    Const methodName As String = "Procリストボックスの値取得配列"
    ' 宣言部
    Dim raw() As String
    Dim totalItems As Long
    Dim i As Long
    Dim arr1D() As String
On Error GoTo ErrHandler
    
    ' RowSource が空なら空配列を返す
    If lst.RowSource = "" Then
        Procリストボックスの値取得配列 = Array()
        Exit Function
    End If
    
    ' RowSource を delimiter で分割
    raw = Split(lst.RowSource, delimiter)
    totalItems = UBound(raw) - LBound(raw) + 1
    
    ' 1次元配列に詰める
    ReDim arr1D(0 To totalItems - 1)
    For i = 0 To totalItems - 1
        arr1D(i) = raw(i)
    Next i
    
    Procリストボックスの値取得配列 = arr1D

ErrHandler:
    Call ErrorSave(methodName)  '必ず先頭（発生したエラーを確実にキャッチするため）
    If ErrObj.Number <> 0 Then Err.Raise ERR_TMP, "", "疑似エラー"  ' 必ず最後(後続処理が動かないため）
    
End Function



'=============================================
' リストボックスの選択行の「連結列」の値を配列化
' lst : ListBox コントロール
'=============================================
Public Function Proc選択された要素を取得配列(lst As Access.ListBox, Optional is条件用 As Boolean = False) As Variant
    Const methodName As String = "Proc選択された要素を取得配列"
    ' 宣言部
    Dim i As Long
    Dim tmp As Collection
    Dim result() As Variant
On Error GoTo ErrHandler
    
    ' メイン処理
    Set tmp = New Collection
    
    ' 選択されている行の「Value」= 連結列の値を拾う
    For i = 0 To lst.ListCount - 1
        If lst.Selected(i) Then
            tmp.Add lst.ItemData(i)   ' ← これが連結列の値
        End If
    Next i
    
    ' 配列化
    If tmp.count > 0 Then
        ReDim result(0 To tmp.count - 1)
        For i = 1 To tmp.count
            If is条件用 Then
                result(i - 1) = Convert条件値文字列(tmp(i))
            Else
                result(i - 1) = tmp(i)
            End If
        Next i
        Proc選択された要素を取得配列 = result
    Else
        Proc選択された要素を取得配列 = Array()
    End If

ErrHandler:
    Call ErrorSave(methodName)  '必ず先頭（発生したエラーを確実にキャッチするため）
    If ErrObj.Number <> 0 Then Err.Raise ERR_TMP, "", "疑似エラー"  ' 必ず最後(後続処理が動かないため）
    
End Function



'=============================================
' リストボックスに項目を追加する関数
'=============================================
Public Sub Procリストボックス要素追加(ByVal newItem As String, lst As Access.ListBox)
    Const methodName As String = "Procリストボックス要素追加"
    ' 宣言部
On Error GoTo ErrHandler

    With lst
        If Len(.RowSource) > 0 Then
            .RowSource = .RowSource & ";" & newItem
        Else
            .RowSource = newItem
        End If
    End With
    
ErrHandler:
    Call ErrorSave(methodName)  '必ず先頭（発生したエラーを確実にキャッチするため）
    If ErrObj.Number <> 0 Then Err.Raise ERR_TMP, "", "疑似エラー"  ' 必ず最後(後続処理が動かないため）
    
End Sub


'=============================================
' リストボックス並べ替え
'=============================================
Public Sub Procリストボックス並べ替え(lst As ListBox)
    Const methodName As String = "Procリストボックス要素追加"
    ' 宣言部
    Dim arr() As String
    Dim i As Long
    Dim itemCount As Long
On Error GoTo ErrHandler
    
    ' リストボックスの行数取得
    itemCount = lst.ListCount
    If itemCount = 0 Then Exit Sub
    
    ' 配列に列1の値を格納
    ReDim arr(0 To itemCount - 1)
    For i = 0 To itemCount - 1
        arr(i) = lst.Column(0, i) ' 列1はColumn(0)
    Next i
    
    ' 配列を昇順ソート（バブルソート）
    Dim j As Long
    Dim temp As String
    For i = LBound(arr) To UBound(arr) - 1
        For j = i + 1 To UBound(arr)
            If arr(i) > arr(j) Then
                temp = arr(i)
                arr(i) = arr(j)
                arr(j) = temp
            End If
        Next j
    Next i
    
    ' リストボックスをクリアして再追加
    lst.RowSourceType = "Value List"
    lst.RowSource = ""
    For i = LBound(arr) To UBound(arr)
        If lst.RowSource = "" Then
            lst.RowSource = arr(i)
        Else
            lst.RowSource = lst.RowSource & ";" & arr(i)
        End If
    Next i
    
ErrHandler:
    Call ErrorSave(methodName)  '必ず先頭（発生したエラーを確実にキャッチするため）
    If ErrObj.Number <> 0 Then Err.Raise ERR_TMP, "", "疑似エラー"  ' 必ず最後(後続処理が動かないため）
    
End Sub


' **************************************************
' 自分以外のすべてのフォームを閉じる
' **************************************************
Public Sub Procすべてのフォームを閉じる(Optional keepFormName As String = "")
    Const methodName As String = "Procすべてのフォームを閉じる"
    ' 宣言部
    Dim frm As AccessObject
    Dim target As String
On Error GoTo ErrHandler
    
    ' デフォルト：呼び出し元のフォームを残す
    If keepFormName = "" Then
        If Screen.ActiveForm Is Nothing Then
            keepFormName = ""
        Else
            keepFormName = Screen.ActiveForm.name
        End If
    End If
    
    For Each frm In CurrentProject.AllForms
        target = frm.name
        If CurrentProject.AllForms(target).IsLoaded Then
            If target <> keepFormName Then
                DoCmd.Close acForm, target, acSaveNo
            End If
        End If
    Next

ErrHandler:
    DoCmd.Echo True
    Call ErrorSave(methodName, True) '必ず先頭（発生したエラーを確実にキャッチするため）
'    If ErrObj.Number <> 0 Then Err.Raise ERR_TMP, "", "疑似エラー"  ' 必ず最後(後続処理が動かないため）
    
End Sub


' **************************************************
' 自分以外のすべてのフォームを閉じる
' **************************************************
Public Sub Procリストボックス複数選択解除(ctrl As ListBox)
On Error Resume Next
    Dim i As Long
    With ctrl
        For i = 0 To .ListCount - 1
            .Selected(i) = False
        Next i
        If .ListCount > 0 Then
            .Selected(0) = True
            .Selected(0) = False
        End If
    End With
    
End Sub


