Attribute VB_Name = "File"
Option Compare Database
Option Explicit


' ***********************************************
' �t�H���_���_�C�A���O�Ŏw�肷��
' ***********************************************
Function Dialog�t�H���_�I��(Optional ByVal initialFolder As String = "") As String
    Const methodName As String = "Dialog�t�H���_�I��"
    ' �錾��
    Dim fd As FileDialog
    Dim selectedPath As String
On Error GoTo ErrHandler
    
    ' ���C������
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)

    With fd
        .title = "�t�H���_��I�����Ă�������"
        .AllowMultiSelect = False

        ' �����t�H���_���w�肳��Ă���ΐݒ肷��
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

    Dialog�t�H���_�I�� = selectedPath
    
ErrHandler:
    Call ErrorSave(methodName)  '�K���擪�i���������G���[���m���ɃL���b�`���邽�߁j
    Call ProcNothing(fd)
    If ErrObj.Number <> 0 Then Err.Raise ERR_TMP, "", "�^���G���["  ' �K���Ō�(�㑱�����������Ȃ����߁j
        
End Function


' ***********************************************
' �t�@�C�����_�C�A���O�Ŏw�肷��
' ***********************************************
Function Dialog�t�@�C���I��(Optional ByVal initialFolder As String = "", _
            Optional filterName As String = "���ׂẴt�@�C��", Optional filter As String = "*.*") As String
    Const methodName As String = "Dialog�t�@�C���I��"
    ' �錾��
    Dim fd As FileDialog
    Dim selectedPath As String
    Dim desc As String
    Dim pattern As String
On Error GoTo ErrHandler
    Dialog�t�@�C���I�� = ""
    
    ' ���C������
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    
    With fd
        .title = "�t�@�C����I�����Ă�������"
        .AllowMultiSelect = False
    
        ' �g���q�t�B���^�[��ݒ�i��F.slot�t�@�C���j
        .Filters.Clear
        
        If filter <> "" Then
            ' ��: "�X���b�g�f�[�^ (*.slot)|*.slot"
            If InStr(filter, "|") > 0 Then
                desc = Split(filter, "|")(0)
                pattern = Split(filter, "|")(1)
            Else
                desc = "�w��Ȃ�"
                pattern = filter
            End If
            .Filters.Add filterName, pattern
        Else
            ' ��Ɂu���ׂẴt�@�C���v���ǉ�
            .Filters.Add "���ׂẴt�@�C��", "*.*"
        End If
        
        ' �����t�H���_��ݒ�
        If initialFolder <> "" Then
            .InitialFileName = initialFolder
        End If
    
        If .Show = -1 Then
            selectedPath = .SelectedItems(1)
        Else
            selectedPath = ""  ' �L�����Z�����͋󕶎�
        End If
    End With

    Set fd = Nothing

    Dialog�t�@�C���I�� = selectedPath

ErrHandler:
    Call ErrorSave(methodName)  '�K���擪�i���������G���[���m���ɃL���b�`���邽�߁j
    If ErrObj.Number <> 0 Then Err.Raise ERR_TMP, "", "�^���G���["  ' �K���Ō�(�㑱�����������Ȃ����߁j
    
End Function


' ***********************************************
' �t�H���_�E�t�@�C�����E�g���q�𕪊����Ĕz��ŕԂ�
' result(0): �t�H���_
' result(1): �t�@�C�����i�g���q�Ȃ��j
' result(2): �g���q
' ***********************************************
Function Get�t�@�C���p�X�����z��(fullPath As String) As Variant
    Const methodName As String = "Get�t�@�C���p�X�����z��"
    Dim pos As Long
    Dim dotPos As Long
    Dim folderPath As String
    Dim fileName As String
    Dim ext As String
    Dim result(2) As String
On Error GoTo ErrHandler
    Get�t�@�C���p�X�����z�� = Null

    ' --- �t�H���_�����ƃt�@�C�������𕪂��� ---
    pos = InStrRev(fullPath, "\")
    If pos > 0 Then
        folderPath = Left(fullPath, pos)   ' "\" ���܂�
        fileName = Mid(fullPath, pos + 1)
    Else
        folderPath = ""
        fileName = fullPath
    End If
    
    ' --- �g���q�𕪂��� ---
    dotPos = InStrRev(fileName, ".")
    If dotPos > 0 Then
        ext = Mid(fileName, dotPos)
        fileName = Left(fileName, dotPos - 1)
    Else
        ext = ""
    End If

    ' --- �z��Ɋi�[ ---
    result(0) = folderPath
    result(1) = fileName
    result(2) = ext
    
    Get�t�@�C���p�X�����z�� = result

ErrHandler:
    Call ErrorSave(methodName)  '�K���擪�i���������G���[���m���ɃL���b�`���邽�߁j
    If ErrObj.Number <> 0 Then Err.Raise ERR_TMP, "", "�^���G���["  ' �K���Ō�(�㑱�����������Ȃ����߁j
    
End Function


' ***********************************************
' �t�H���_���w�肵�ăt�@�C�����擾����i������z��j
' ***********************************************
Function Get�t�@�C���擾�z��(folderPath As String, filterPattern As String) As String()
    Const methodName As String = "Get�t�@�C���擾�z��"
    ' �錾��
    Dim fileName As String
    Dim fileList() As String
    Dim count As Long
On Error GoTo ErrHandler
    
    ' �t�H���_������ \ ��⊮
    If Right(folderPath, 1) <> "\" Then
        folderPath = folderPath & "\"
    End If

    ' ������
    count = 0
    fileName = Dir(folderPath & filterPattern)
    
    Do While fileName <> ""
        ReDim Preserve fileList(count)
        fileList(count) = fileName ' �� �t���p�X�ł͂Ȃ��t�@�C�����̂�
        count = count + 1
        fileName = Dir()
    Loop
    
    ' �t�@�C����������Ȃ���΋�z���Ԃ�
    If count = 0 Then
        Get�t�@�C���擾�z�� = Split("") ' ����0�̔z���Ԃ�
    Else
        Get�t�@�C���擾�z�� = fileList
    End If

ErrHandler:
    Call ErrorSave(methodName)  '�K���擪�i���������G���[���m���ɃL���b�`���邽�߁j
    If ErrObj.Number <> 0 Then Err.Raise ERR_TMP, "", "�^���G���["  ' �K���Ō�(�㑱�����������Ȃ����߁j
    
End Function


' ***********************************************
' �t�@�C�����̂ݎ擾����
' ***********************************************
Function GetFileName(fullPath As String) As String
    Const methodName As String = "GetFileName"
    Dim fso As Object
On Error GoTo ErrHandler

    GetFileName = ""
    
    ' ���C������
    Set fso = CreateObject("Scripting.FileSystemObject")
    GetFileName = fso.GetFileName(fullPath)
    
ErrHandler:
    Call ErrorSave(methodName)  '�K���擪�i���������G���[���m���ɃL���b�`���邽�߁j
    Call ProcNothing(fso)
    If ErrObj.Number <> 0 Then Err.Raise ERR_TMP, "", "�^���G���["  ' �K���Ō�(�㑱�����������Ȃ����߁j
        
End Function


' **************************************************
' �����N��t�@�C�����R�s�[�i�I�v�V�����Ńt�@�C������ύX�j
' **************************************************
Public Sub Cmd�t�@�C���R�s�[(srcPath As String, destFolder As String, Optional newFileName As String = "")
    Const methodName As String = "Cmd�t�@�C���R�s�["
    On Error GoTo ErrHandler
    
    Dim destPath As String
    
    ' �R�s�[��̊��S�p�X���쐬
    If Right(destFolder, 1) <> "\" Then
        destFolder = destFolder & "\"
    End If
    
    ' �V�����t�@�C�������w�肳��Ă���ꍇ
    If newFileName <> "" Then
        destPath = destFolder & newFileName
    Else
        destPath = destFolder & Dir(srcPath)
    End If
    
'    ' ���ɓ����t�@�C�������݂���ꍇ�͍폜
'    If Dir(destPath) <> "" Then
'        Kill destPath
'    End If
    
    ' �t�@�C���R�s�[
    FileCopy srcPath, destPath
    
    Exit Sub  ' ����I�����͂����Ŕ�����

ErrHandler:
    Call ErrorSave(methodName)  ' �G���[���L�^
    If Err.Number <> 0 Then Err.Raise vbObjectError + 1000, , "�t�@�C���R�s�[���ɃG���[���������܂���"

End Sub





' ***********************************************
' �o�b�`���s�֐��i�񓯊��j
' ***********************************************
Public Sub Run�o�b�`�v���O�������s(vbsPath As String)
    Const methodName As String = "Run�o�b�`�v���O�������s"
On Error Resume Next

    ' ���C������
    shell "cscript //nologo """ & vbsPath & """", vbNormalFocus
    
ErrHandler:
    Call ErrorSave(methodName)  '�K���擪�i���������G���[���m���ɃL���b�`���邽�߁j
    If ErrObj.Number <> 0 Then Err.Raise ERR_TMP, "", "�^���G���["  ' �K���Ō�(�㑱�����������Ȃ����߁j
    
End Sub

' ***********************************************
' �o�b�`���s�֐��i�����j
' ***********************************************
Public Sub Run�o�b�`�v���O�������sAndWait(vbsPath As String)
    Const methodName As String = "Run�o�b�`�v���O�������sAndWait"
    Dim wsh As Object
On Error GoTo ErrHandler

    ' ���C������
    Set wsh = CreateObject("WScript.Shell")
    
    wsh.Run "cscript //nologo """ & vbsPath & """", 1, True  ' True = �����܂őҋ@
    
ErrHandler:
    Call ErrorSave(methodName)  '�K���擪�i���������G���[���m���ɃL���b�`���邽�߁j
    Call ProcNothing(wsh)
    If ErrObj.Number <> 0 Then Err.Raise ERR_TMP, "", "�^���G���["  ' �K���Ō�(�㑱�����������Ȃ����߁j
    
End Sub



' **************************************************
' �t�H���_�����݂��Ȃ��ꍇ�͍쐬����
' **************************************************
Public Sub Proc�t�H���_�����݂��Ȃ��ꍇ�͍쐬(folderPath As String)
    Const methodName As String = "Proc�t�H���_�����݂��Ȃ��ꍇ�͍쐬"
    Dim fso As Object
On Error GoTo ErrHandler

    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(folderPath) Then
        fso.CreateFolder folderPath
    End If
    
ErrHandler:
    Call ErrorSave(methodName)  '�K���擪�i���������G���[���m���ɃL���b�`���邽�߁j
    Call ProcNothing(fso)
    If ErrObj.Number <> 0 Then Err.Raise ERR_TMP, "", "�^���G���["  ' �K���Ō�(�㑱�����������Ȃ����߁j

End Sub


' **************************************************
' �t�@�C�������݂��Ă���� True ��Ԃ�
' **************************************************
Public Function Proc�t�@�C�����݊m�F(filePath As String) As Boolean
    Dim fso As Object
On Error Resume Next
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Proc�t�@�C�����݊m�F = fso.FileExists(filePath)

    Call ProcNothing(fso)
    
End Function

' ***********************************************
' �t�@�C�����R�s�[����ifile �� folder�j
' ***********************************************
Public Sub Proc�t�@�C�����w��t�H���_�փR�s�[(srcFile As String, destFolder As String)
    Const methodName As String = "Proc�t�@�C�����w��t�H���_�փR�s�["
    Dim fso As Object
    Dim fileName As String
    Dim destFile As String
On Error GoTo ErrHandler

    ' �t�@�C�������o
    fileName = Mid(srcFile, InStrRev(srcFile, "\") + 1)

    ' �t�H���_�����⊮
    If Right(destFolder, 1) <> "\" Then destFolder = destFolder & "\"

    ' �R�s�[��t�@�C���p�X�\�z
    destFile = destFolder & fileName

    ' �R�s�[���s
    Set fso = CreateObject("Scripting.FileSystemObject")
    fso.CopyFile srcFile, destFile, True

ErrHandler:
    Call ErrorSave(methodName)
    Call ProcNothing(fso)
    If ErrObj.Number <> 0 Then Err.Raise ERR_TMP, "", "�^���G���["
    
End Sub



