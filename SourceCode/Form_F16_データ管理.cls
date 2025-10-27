VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_F16_データ管理"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


' ======================================================
Private Sub Form_Load()
On Error Resume Next
    Form.caption = App_アプリ名
    Me.KeyPreview = True
    Call subタイトル.Form.Procタイトル更新("データ管理", E_Color_Red)
    
    Call LogOpen(Me.name)
    Call Procアプリ起動処理
        
End Sub
' ======================================================
Private Sub Form_Close()
On Error Resume Next
    
    Call ErrorSave(Me.name, True) '閉じる前にエラーリセット
    Call LogClose(Me.name)
    
    If flg最適化 Then
        DoCmd.OpenForm F01_OptimizeForm, , , , , , F11_MainMenu
    End If
    
End Sub
' ======================================================
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
    If KeyCode = vbKeyEscape Then
'        Me.Undo
        DoCmd.Close acForm, Me.name
        DoCmd.Close acForm, Me.Parent.name
    End If
    
End Sub
' ======================================================
Private Sub Procアプリ起動処理()
    Const methodName As String = "Procアプリ起動処理"
    Dim parts As Variant
On Error Resume Next

    parts = Getファイルパス分割配列(Procリンクテーブル接続先取得)
    textフォルダ = parts(0)
'    textデータファイル新 = parts(0) & parts(1) & parts(2)
    textデータファイル = parts(1)
    textデータファイル新 = parts(1)

ErrHandler:
    Call ErrorSave(methodName, True) '必ず先頭（メッセージ出力）
    
End Sub

' ======================================================
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    ' データファイルがない場合は接続できない（強制終了しかない）
    If Not App_接続チェック Then
        Cancel = True
    End If
    
End Sub




' ======================================================
Private Sub btn戻す_Click()
On Error Resume Next
    Call Procアプリ起動処理

End Sub

' ======================================================
Private Sub btn付与店舗名_Click()
On Error Resume Next
    textデータファイル新 = textデータファイル新 & "_" & Get店舗名

End Sub

' ======================================================
Private Sub btn付与時刻_Click()
On Error Resume Next
    textデータファイル新 = textデータファイル新 & "_" & Format(Now, "_yymmdd_hhmm")
    
End Sub




' ======================================================
Private Sub btnファイル変更_Click()
    Const methodName As String = "btnファイル変更_Click"
    ' 宣言部
    Dim file As String
    Dim parts As Variant
On Error GoTo ErrHandler

    Call LogStart(methodName)
    
    ' メイン処理
    Call ShowConfirm("処理前確認", "データファイルの接続先を変更します。", vbYes)
    
'    file = Dialogファイル選択(textデータファイル新, "データファイル", "*.accdb")
    file = Dialogファイル選択(textフォルダ, "データファイル", "*" & SYS拡張子)
    If file = "" Then GoTo ErrHandler
    
    Call Procリンクテーブル一括更新(file)
    
    parts = Getファイルパス分割配列(file)
'    textデータファイル新 = parts(0) & parts(1) & parts(2)
    textフォルダ = parts(0)
    textデータファイル = parts(1)
    textデータファイル新 = parts(1)
    
'    Call ShowConfirm("システム", "正常に処理は完了しました。", vbYes)
'    Call ShowToast("正常に処理は完了しました。", E_Color_Blue)

    Me.Dirty = False
    
ErrHandler:
    Call ErrorSave(methodName, True) '必ず先頭（メッセージ出力）

End Sub

' ======================================================
Private Sub btnファイル複製_Click()
    Const methodName As String = "btnファイル複製_Click"
    ' 宣言部
    Dim parts As Variant
    Dim newFile As String
    Dim dataFile As String
On Error GoTo ErrHandler

    Call LogStart(methodName)
    
    ' メイン処理
    If ShowConfirm("処理前確認", "データファイルをコピーしてそちらに接続します。　よろしいですか？") <> vbYes Then GoTo ErrHandler
    
    dataFile = Procリンクテーブル接続先取得
    parts = Getファイルパス分割配列(dataFile)
    
    If parts(1) = textデータファイル新 Then
        Call ShowToast("複製するファイル名が変更前と同じです！", E_Color_Red)
        GoTo ErrHandler
    End If
    
    Call Cmdファイルコピー(dataFile, CStr(parts(0)), textデータファイル新 & parts(2))
    
    newFile = parts(0) & textデータファイル新 & parts(2)
    Call Procリンクテーブル一括更新(newFile)
    
    parts = Getファイルパス分割配列(Procリンクテーブル接続先取得)
    textフォルダ = parts(0)
    textデータファイル = parts(1)
    textデータファイル新 = parts(1)
    
    Me.Dirty = False    '保存
    Call ShowToast("正常に処理は完了しました。", E_Color_Blue)
    
ErrHandler:
    Call ErrorSave(methodName, True) '必ず先頭（メッセージ出力）

End Sub

' ======================================================
Private Sub btnバックアップ_Click()
    Const methodName As String = "btnバックアップ_Click"
    ' 宣言部
    Dim parts As Variant
    Dim newFile As String
    Dim dataFile As String
On Error GoTo ErrHandler

    Call LogStart(methodName)
    
    ' メイン処理
    If ShowConfirm("処理前確認", "データファイルをコピーしてそちらに接続します。　よろしいですか？") <> vbYes Then GoTo ErrHandler
    
    dataFile = Procリンクテーブル接続先取得
    parts = Getファイルパス分割配列(dataFile)
    
    If parts(1) = textデータファイル新 Then
        Call ShowToast("複製するファイル名が変更前と同じです！", E_Color_Red)
        GoTo ErrHandler
    End If
    
    Call Cmdファイルコピー(dataFile, CStr(parts(0)), textデータファイル新 & parts(2))
    
'    newFile = parts(0) & textデータファイル新 & parts(2)
'    Call Procリンクテーブル一括更新(newFile)
'
'    parts = Getファイルパス分割配列(Procリンクテーブル接続先取得)
'    textフォルダ = parts(0)
'    textデータファイル = parts(1)
'    textデータファイル新 = parts(1)
    
    Call ShowToast("正常に処理は完了しました。", E_Color_Blue)
    
ErrHandler:
    Call ErrorSave(methodName, True) '必ず先頭（メッセージ出力）

End Sub


' ======================================================
Private Sub btn初期化_Click()
    Const methodName As String = "btn初期化_Click"
    ' 宣言部
    Dim ret
On Error GoTo ErrHandler

    Call LogStart(methodName)
    
    ' メイン処理
    If ShowConfirm("初期化前確認", "データファイルのバックアップは大丈夫ですか？") <> vbYes Then GoTo ErrHandler
    If ShowConfirm("初期化前確認", "データを初期化します。　よろしいですか？") <> vbYes Then GoTo ErrHandler
    
    'テーブル削除
    Call Appデータ初期化

    '店舗マスタ（データ作成）
    Call Add店舗名

    Call ShowToast("正常に初期化は完了しました。", E_Color_Blue)
    
    flg最適化 = True

ErrHandler:
    Call ErrorSave(methodName, True) '必ず先頭（メッセージ出力）
    
End Sub

