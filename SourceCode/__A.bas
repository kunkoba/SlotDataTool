Attribute VB_Name = "__A"
Option Compare Database
Option Explicit


'' ======================================================
'Private Sub Form_Load()
'    Dim args() As String
'On Error Resume Next
'    Form.caption = App_�A�v����
'    Me.KeyPreview = True
''    Call sub�^�C�g��.Form.�^�C�g���ݒ�("�o�ʐ��ڃO���t", E_Color_DeepBlue)
'
'    Call LogOpen(Me.name)
'    Call Proc�A�v���N������
'
'End Sub
'' ======================================================
'Private Sub Form_Close()
'On Error Resume Next
'    Call LogClose(Me.name)
'    Me.Undo
'
'End Sub
'' ======================================================
'Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'On Error Resume Next
'    If KeyCode = vbKeyEscape Then
'        DoCmd.Close acForm, Me.name
'        DoCmd.Close acForm, Me.Parent.name
'    End If
'
'End Sub
'' ======================================================
'Private Sub Proc�A�v���N������()
'    Const methodName As String = "Proc�A�v���N������"
'    Dim args
'On Error GoTo ErrHandler
'
'    ' ���C������
'
'ErrHandler:
'    Call ErrorSave(methodName, True) '�K���擪�i���b�Z�[�W�o�́j
'
'End Sub




' ******************************************************
' �t�H�[���C�x���g�i���b�Z�[�W�o�͂���j
' ******************************************************
Private Sub XXXXXX1_Click()
    Const methodName As String = "XXXXXX1_Click"
On Error GoTo ErrHandler

    Call LogStart(methodName)
    
    ' ���C������
    
ErrHandler:
    Call ErrorSave(methodName, True) '�K���擪�i���b�Z�[�W�o�́j
    
End Sub


' ******************************************************
' ���ԏ����i���b�Z�[�W�o�͂Ȃ��j
' ******************************************************
Private Sub ZZZZZZ__3()
    Const methodName As String = "ZZZZZZ__3"
On Error GoTo ErrHandler

    ' ���C������

ErrHandler:
    Call ErrorSave(methodName)  '�K���擪�i���������G���[���m���ɃL���b�`���邽�߁j
'    Call ProcNothing(db)
    If ErrObj.Number <> 0 Then Err.Raise ERR_TMP, "", "�^���G���["  ' �K���Ō�(�㑱�����������Ȃ����߁j

End Sub





