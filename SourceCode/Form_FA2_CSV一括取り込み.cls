VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_FA2_CSV一括取り込み"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Const HEADER As String = ""
Dim canselSwitch As Boolean


' ======================================================
Private Sub Form_Load()
On Error Resume Next
    Form.caption = App_アプリ名
    Me.KeyPreview = True
'    Call subタイトル.Form.Procタイトル更新("CSV一括取り込み", 0)

    Call LogOpen(Me.name)
    Call Procアプリ起動処理
    
End Sub
' ======================================================
Private Sub Form_Close()
On Error Resume Next

    Call ErrorSave(Me.name, True) '閉じる前にエラーリセット
    Call LogClose(Me.name)

End Sub
' ======================================================
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
    If KeyCode = vbKeyEscape Then
        Me.Undo
        DoCmd.Close acForm, Me.name
        DoCmd.Close acForm, Me.Parent.name
    End If
    
End Sub
' ======================================================
Private Sub Procアプリ起動処理()
    Const methodName As String = "Procアプリ起動処理"
On Error Resume Next

    ' メイン処理
    listファイル一覧.RowSource = HEADER
    
    If IsFormOpen(FA1_ImportData) Then
        list形式 = Forms(FA1_ImportData).list形式
        comb機種ID = Forms(FA1_ImportData).comb機種ID
    End If
    
    canselSwitch = False

ErrHandler:
    Call ErrorSave(methodName, True) '必ず先頭（メッセージ出力）
    
End Sub



' ======================================================
Private Sub btn中断_Click()
    canselSwitch = True
    
End Sub


' ======================================================
Private Sub comb機種ID_AfterUpdate()
On Error Resume Next
    If IsFormOpen(FA1_ImportData) Then
        Forms(FA1_ImportData).comb機種ID = comb機種ID
        Forms(FA1_ImportData).ProcComb機種ID_AfterUpdate
        Forms(FA1_ImportData).btn降順データ積み上げ日_Click
    End If

End Sub

' ======================================================
Private Sub list形式_AfterUpdate()
On Error Resume Next
    If IsFormOpen(FA1_ImportData) Then
        Forms(FA1_ImportData).list形式 = list形式
        Forms(FA1_ImportData).Proclist形式_AfterUpdate
    End If

End Sub


' ======================================================
Private Sub btnフォルダ選択_Click()
    Const methodName As String = "btnフォルダ選択_Click"
    ' 宣言部
    Dim textDir As String
On Error GoTo ErrHandler

    Call LogStart(methodName)

    ' メイン処理
    textDir = Dialogフォルダ選択()
    If textDir = "" Then Exit Sub
    
    listファイル一覧.RowSource = Join(Getファイル取得配列(textDir, "*.csv"), ";")
    textフォルダ名 = textDir

ErrHandler:
    Call ErrorSave(methodName, True) '必ず先頭（メッセージ出力）
    
End Sub



' ======================================================
Private Sub btn一括取り込み_Click()
    Const methodName As String = "btn一括取り込み_Click"
    ' 宣言部
    Const TBL_MINREPO As String = "Import_みんレポCSV"
    Dim arr As Variant
    Dim csvFolder As String
    Dim i As Long
    Dim filePath As String
    Dim status As String
    Dim msg As String
    Dim ret As Boolean
On Error GoTo ErrHandler

    Call LogStart(methodName)
        
    ' メイン処理
    canselSwitch = False
    list結果.RowSource = ""
    
    If ShowConfirm("一括取り込み前確認", "実行しますか？") <> vbYes Then Exit Sub
    
    btn中断.Visible = True
    btn中断.SetFocus
    btn一括取り込み.Visible = False
    
    ' フォルダパス
    csvFolder = Trim(Me.textフォルダ名.value)
    If Right(csvFolder, 1) <> "\" Then csvFolder = csvFolder & "\"
    
    ' 1. リストボックスを配列化（ここでは arr(i) = ファイル名）
    arr = Procリストボックスの値取得配列(listファイル一覧)
        
    ' 2. 配列ループ
    For i = LBound(arr) To UBound(arr)
        status = ""
        filePath = csvFolder & Trim(arr(i))
        
        If Dir(filePath) <> "" Then
            ' 取り込み＆積み上げ
            ret = Forms(FA1_ImportData).Proc取り込みCSV(filePath)
            
            If ret Then
                If Forms(FA1_ImportData).Proc積み上げ処理() Then
'                    status = "成功★: " & Trim(arr(i))
                Else
                    status = "失敗×: " & Trim(arr(i))
                    Call Procリストボックス要素追加(status, list結果)
                End If
            Else
                status = "失敗×: " & Trim(arr(i))
                Call Procリストボックス要素追加(status, list結果)
            End If
        Else
            status = "ファイルなし: " & Trim(arr(i))
            Call Procリストボックス要素追加(status, list結果)
        End If
            
        DoEvents
        Sleep 500
        
        ' 中断ボタン
        If canselSwitch Then
            If ShowConfirm("一括取り込み処理", "処理を中断しますか？", vbYesNo) = vbYes Then
                status = "処理を中断しました。"
            Else
                canselSwitch = False
            End If
        End If
        
        If canselSwitch Then GoTo ErrHandler
        
    Next i

'    Call ShowToast("一括取り込み処理は処理は完了しました。", E_Color_Blue)
    
    status = "処理は完了しました。"
    Call Procリストボックス要素追加(status, list結果)

ErrHandler:
    Call ErrorSave(methodName, True) '必ず先頭（メッセージ出力）
    btn一括取り込み.Visible = True
    btn一括取り込み.SetFocus
    btn中断.Visible = False
    
End Sub

'' ******************************************************
'Function Check積み上げ処理前() As String
'    Dim msg As String
'
'    'チェック
'    If Nz(list形式, "") = "" Or Nz(list形式, "") = "1" Then msg = "データ形式が選択されていません！"
''    If DCount("*", list形式.value) = 0 Then msg = "データ形式が選択されていません！"
'    If Nz(comb機種ID, "") = "" Or Nz(comb機種ID, "") = "0" Then msg = "機種が選択されていません！"
'
'    Check積み上げ処理前 = msg
'
'End Function


