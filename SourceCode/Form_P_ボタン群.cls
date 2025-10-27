VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_P_ボタン群"
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
Private Sub btn店舗マスタ_Click()
On Error Resume Next
    DoCmd.OpenForm FB0_MstStore
    
End Sub

' ======================================================
Private Sub btn日情報_Click()
On Error Resume Next
    DoCmd.OpenForm FC9_DataDateInfo

End Sub

' ======================================================
Private Sub btn日推移_Click()
On Error Resume Next
    DoCmd.OpenForm FD9_GraphCoinTrend

End Sub


