Attribute VB_Name = "_Package"
Option Compare Database
Option Explicit




' ***********************************************
' �p�b�P�[�W�쐬����
' ***********************************************
Public Sub Pac�p�b�P�[�W�쐬()
    Const methodName As String = "Pac�p�b�P�[�W�쐬"
    Dim thisPath As String
    Dim tmpPath As String
On Error GoTo ErrHandler

    thisPath = CurrentProject.path & "\"
    
    ' zip�t�H���_
    tmpPath = thisPath & PATH_BIN
    Call Proc�t�H���_�����݂��Ȃ��ꍇ�͍쐬(tmpPath)
    
    ' accde�t�@�C���쐬
    tmpPath = tmpPath & App_�o�[�W����
    Call Proc�t�H���_�����݂��Ȃ��ꍇ�͍쐬(tmpPath)
    Call PacACCDE�t�@�C���쐬(tmpPath)
    
    ' �f�[�^�t�@�C���쐬
    Call Proc�t�H���_�����݂��Ȃ��ꍇ�͍쐬(tmpPath & PATH_DATA)
    Call Proc�t�@�C�����w��t�H���_�փR�s�[(Proc�����N�e�[�u���ڑ���擾, tmpPath & PATH_DATA)
    
    '���O�t�H���_�쐬
    Call Proc�t�H���_�����݂��Ȃ��ꍇ�͍쐬(tmpPath & PATH_LOG)
    
    Call ShowToast("�����͐���Ɋ������܂����B", E_Color_Blue)

ErrHandler:
    Call ErrorSave(methodName)  '�K���擪�i���������G���[���m���ɃL���b�`���邽�߁j
'    If ErrObj.Number <> 0 Then Err.Raise ERR_TMP, "", "�^���G���["  ' �K���Ō�(�㑱�����������Ȃ����߁j

End Sub


' ***********************************************
' PacACCDE�t�@�C���쐬�i�w��t�H���_�ɃR�s�[ �� �R���p�C�� �� �R�s�[�폜�j
' ***********************************************
Public Sub PacACCDE�t�@�C���쐬(destFolder As String)
    Const methodName As String = "PacACCDE�t�@�C���쐬"
    Dim db As Object
    Dim parts As Variant
    Dim fileName As String
    Dim copiedFile As String
On Error GoTo ErrHandler

    ' ������
    Set db = CurrentDb
    parts = Get�t�@�C���p�X�����z��(db.name)
    If IsNull(parts) Then GoTo ErrHandler

    ' �t�@�C�����i�g���q�t���j���\�z
    fileName = parts(1) & parts(2)

    ' �t�H���_�����⊮
    If Right(destFolder, 1) <> "\" Then destFolder = destFolder & "\"

    ' �R�s�[��t�@�C���p�X�\�z
    copiedFile = destFolder & fileName

    ' �t�H���_���Ȃ���΍쐬
    Call Proc�t�H���_�����݂��Ȃ��ꍇ�͍쐬(destFolder)

    Call LogDebug(methodName, "�R�s�[: " & copiedFile)

    ' �t�@�C�����R�s�[�i�t�@�C�����ύX�Ȃ��j
    Call Proc�t�@�C�����w��t�H���_�փR�s�[(db.name, destFolder)

    ' �R���p�C���i���t�H���_�� .accde �𐶐��j
    Call PacACCDE�R���p�C��(copiedFile)

    ' �R�s�[���ꂽ .accdb ���폜
    Kill copiedFile
    Call LogDebug(methodName, "�폜: " & copiedFile)

ErrHandler:
    Call ErrorSave(methodName)
    Call ProcNothing(db)
    If ErrObj.Number <> 0 Then Err.Raise ERR_TMP, "", "�^���G���["
    
End Sub




' ***********************************************
' �R���p�C��
' ***********************************************
Private Sub PacACCDE�R���p�C��(srcPath As String)
    Const methodName As String = "PacACCDE�R���p�C��"
    ' �錾��
    Dim fso As Object
    Dim Source As String
    Dim ts As Object
    Dim vbsPath As String
    Dim destPath As String
On Error GoTo ErrHandler
    
    ' �R���p�C�����s�t�@�C���iVBS�t�@�C���j�̐���
    vbsPath = Get�t�@�C���p�X�����z��(srcPath)(0) & "build_accde.vbs"
    destPath = Replace(srcPath, ".accdb", ".accde")
    
    Source = _
        "Set acc = CreateObject(""Access.Application"")" & vbCrLf & _
        "acc.SysCmd 603, """ & srcPath & """, """ & destPath & """" & vbCrLf & _
        "acc.Quit" & vbCrLf & _
        "Set acc = Nothing"
    
    ' FileSystemObject�Ńt�@�C�������o��
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.CreateTextFile(vbsPath, True, False) ' ��3����: Unicode=False�iANSI�ŕۑ��j
    
    ts.Write Source
    ts.Close
    
    Call LogDebug(methodName, vbsPath)
    
    ' �o�b�`���s
    shell "wscript.exe """ & vbsPath & """", vbNormalFocus
    
    ' ���t�@�C���폜
    Sleep 3000  ' 3�b���炢����΃R���p�C����������ł���
    Kill vbsPath
'    Kill srcPath

ErrHandler:
    Call ErrorSave(methodName)  '�K���擪�i���������G���[���m���ɃL���b�`���邽�߁j
    Call ProcNothing(ts)
    Call ProcNothing(fso)
    If ErrObj.Number <> 0 Then Err.Raise ERR_TMP, "", "�^���G���["  ' �K���Ō�(�㑱�����������Ȃ����߁j
        
End Sub




