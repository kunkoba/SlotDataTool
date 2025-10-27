Attribute VB_Name = "_AutoExe"
Option Compare Database
Option Explicit

'--- Access�{�̃E�B���h�E���B�� ---
Private Declare PtrSafe Function apiShowWindow Lib "user32" Alias "ShowWindow" ( _
    ByVal hwnd As LongPtr, ByVal nCmdShow As Long) As Long

Private Declare PtrSafe Function apiFindWindowA Lib "user32" Alias "FindWindowA" ( _
    ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr


' *********************************************************************
' AutoExe
' *********************************************************************
Public Function Proc�A�v���N����������(Optional is��\�� As Boolean = True)
    Const methodName As String = "Proc�A�v���N����������"
    ' �錾��
On Error GoTo ErrHandler
    
    Call LogOpen(methodName)
    
    ' ��d�N���`�F�b�N
    If Is��d�N���`�F�b�N() Then
        MsgBox "���łɋN�����Ă��܂��B��d�N���͂ł��܂���B", vbExclamation
        Application.Quit
        Exit Function
    End If
    
'    Application.Echo False

    If is��\�� Then
        ' Access�E�B���h�E���\����
        Dim hwnd As LongPtr
        hwnd = apiFindWindowA("OMain", vbNullString)
        If hwnd <> 0 Then
            Call apiShowWindow(hwnd, 0) ' SW_HIDE = 0
        End If
        
    End If

    flg�œK�� = False   '�œK���t���O
    
    ' ���j���[��ʕ\���O����
    If IsNull(App_�A�v���L������) Then
        Call Set�A�v���N�����ݒ�X�V
        Call MAC�A�h���X�X�V
    End If
    
    ' ���j���[��ʕ\��
    DoCmd.OpenForm F11_MainMenu
    
    ' �ڑ���`�F�b�N
    If Not App_�ڑ��`�F�b�N Then
        ' �����N�ڑ�
        DoCmd.OpenForm F16_LinkManager
    End If

ErrHandler:
    Call ErrorSave(methodName, True) '�K���擪�i���������G���[���m���ɃL���b�`���邽�߁j
'    Application.Echo True
    Call LogClose(methodName)
        
End Function


' *********************************************************************
' ��d�N���`�F�b�N�i����t�@�C���j
' *********************************************************************
Public Function Is��d�N���`�F�b�N() As Boolean
    Const methodName As String = "Is��d�N���`�F�b�N"
    ' �錾��
    Dim objWMI As Object
    Dim colProcesses As Object
    Dim objProcess As Object
    Dim myPath As String
    Dim cmdLine As String
    Dim count As Integer
On Error GoTo ErrHandler
    
    ' ���C������
    myPath = """" & CurrentDb.name & """" ' �t���p�X���͂�Ŕ�r�i�X�y�[�X�΍�j
    count = 0

    On Error Resume Next
        Set objWMI = GetObject("winmgmts:\\.\root\CIMV2")
        If objWMI Is Nothing Then
            Is��d�N���`�F�b�N = False
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

    ' �������܂߂�2�ȏ゠��Γ���t�@�C�������d�N����
    Is��d�N���`�F�b�N = (count >= 2)

ErrHandler:
    Call ErrorSave(methodName)  '�K���擪�i���������G���[���m���ɃL���b�`���邽�߁j
    If ErrObj.Number <> 0 Then Err.Raise ERR_TMP, "", "�^���G���["  ' �K���Ō�(�㑱�����������Ȃ����߁j
    
End Function

' ======================================================
'�� �uS_�A�v���ݒ�v�̎w�荀�ڂ�ݒ肷��
' ======================================================
Public Function Set�A�v���N�����ݒ�X�V() As Boolean
    Const methodName As String = "Set�A�v���N�����ݒ�X�V"
    ' �錾��
    Dim strSQL As String
    Dim key As String
On Error GoTo ErrHandler

    Set�A�v���N�����ݒ�X�V = False
    
    ' ���C������
    key = Lic�閧�R�[�h����
    
    ' SQL�쐬
    strSQL = "UPDATE S_�A�v���ݒ� " & _
             "SET " & _
             " �A�v���N���� = #" & Date & "#" & _
             ",�A�v���L������ = #" & Date + 7 & "#" & _
             ",�閧�R�[�h = '" & key & "'" & _
             ",�Í��L�[ = '" & Lic�閧�R�[�h����Í��L�[(key) & "'"
    
    ' SQL���s
    CurrentDb.Execute strSQL, dbFailOnError
    
    Set�A�v���N�����ݒ�X�V = True

ErrHandler:
    Call ErrorSave(methodName)  '�K���擪�i���������G���[���m���ɃL���b�`���邽�߁j
'    Call ProcNothing(db)
    If ErrObj.Number <> 0 Then Err.Raise ERR_TMP, "", "�^���G���["  ' �K���Ō�(�㑱�����������Ȃ����߁j

End Function





