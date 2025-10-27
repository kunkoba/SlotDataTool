VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_F03_トースト"
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
    
End Sub


' ======================================================
Private Sub Form_Timer()
On Error Resume Next
    DoCmd.Close acForm, Me.name
    
End Sub


' ======================================================
Private Sub boxBackBoard_Click()
On Error Resume Next
    DoCmd.Close acForm, Me.name

End Sub

' ======================================================
Private Sub lblMessage_Click()
On Error Resume Next
    DoCmd.Close acForm, Me.name

End Sub
