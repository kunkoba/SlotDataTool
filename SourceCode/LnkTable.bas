Attribute VB_Name = "LnkTable"
Option Compare Database
Option Explicit


' **************************************************
' �����N�e�[�u���̐ڑ����ԋp�i�ڑ��悪��̂݁j
' **************************************************
Public Function Proc�����N�e�[�u���ڑ���擾() As String
    Const methodName As String = "Proc�����N�e�[�u���ڑ���擾"
    ' �錾��
    Dim db As DAO.Database
    Dim tdf As DAO.TableDef
    Dim conn As String
    Dim pos As Long
On Error GoTo ErrHandler
    
    Proc�����N�e�[�u���ڑ���擾 = ""
    
    ' ���C������
    Set db = CurrentDb
    
    For Each tdf In db.TableDefs
        If Len(tdf.Connect) > 0 Then
            conn = tdf.Connect
            
            ' ";DATABASE=" ��T���ăp�X���������𔲂��o��
            pos = InStr(conn, ";DATABASE=")
            If pos > 0 Then
                Proc�����N�e�[�u���ڑ���擾 = Mid(conn, pos + Len(";DATABASE="))
            Else
                Proc�����N�e�[�u���ڑ���擾 = conn
            End If
            
            Exit For
            
        End If
    Next

ErrHandler:
    Call ErrorSave(methodName)  '�K���擪�i���������G���[���m���ɃL���b�`���邽�߁j
    Call ProcNothing(tdf)
    Call ProcNothing(db)
    If ErrObj.Number <> 0 Then Err.Raise ERR_TMP, "", "�^���G���["  ' �K���Ō�(�㑱�����������Ȃ����߁j
        
End Function


' **************************************************
' �����N�e�[�u���̐ڑ���iDATABASE= �̒l�j�����u������i�t�@�C�����̂݁j
' **************************************************
Public Sub Proc�����N�e�[�u���ꊇ�X�V(newFileName As String)
    Const methodName As String = "Proc�����N�e�[�u���ꊇ�X�V"
    ' �錾��
    Dim db As DAO.Database
    Dim tdf As DAO.TableDef
    Dim conn As String
    Dim posDB As Long, startVal As Long, nextSemi As Long
    Dim oldDBValue As String, folderPath As String, tail As String
    Dim candidateNewDBValue As String, newConn As String, oldConn As String
    Dim lastErr As Long, lastDesc As String
    Dim isFullPath As Boolean
On Error GoTo ErrHandler
    
    
    ' ���C������
    Set db = CurrentDb

    ' --- newFileName ���t���p�X���ǂ����� 1�񂾂����� ---
    isFullPath = (InStr(newFileName, "\") > 0)

    For Each tdf In db.TableDefs
        ' �����N�e�[�u���̂ݑΏ�
        If Len(Trim$(tdf.Connect & "")) > 0 Then
            conn = tdf.Connect
            posDB = InStr(1, conn, "DATABASE=", vbTextCompare)
            If posDB > 0 Then
                startVal = posDB + Len("DATABASE=")
                nextSemi = InStr(startVal, conn, ";")
                If nextSemi > 0 Then
                    oldDBValue = Mid(conn, startVal, nextSemi - startVal)
                    tail = Mid(conn, nextSemi) ' ;�ȍ~
                Else
                    oldDBValue = Mid(conn, startVal)
                    tail = ""
                End If

                ' ���̐ڑ�������̃t�H���_����
                If InStrRev(oldDBValue, "\") > 0 Then
                    folderPath = Left(oldDBValue, InStrRev(oldDBValue, "\"))
                Else
                    folderPath = ""
                End If

                ' candidateNewDBValue ���쐬
                If isFullPath Then
                    candidateNewDBValue = newFileName
                ElseIf folderPath <> "" Then
                    candidateNewDBValue = folderPath & newFileName
                Else
                    candidateNewDBValue = newFileName
                End If

                ' �t�@�C�����݃`�F�b�N
                If Len(Dir(candidateNewDBValue)) = 0 Then
                    Call LogDebug(methodName, "�����N��t�@�C����������܂���B�X�L�b�v: " & candidateNewDBValue)
                    GoTo ContinueNext
                End If

                ' �V�����ڑ�����������iDATABASE= �̒l���������ւ��j
                newConn = Left(conn, startVal - 1) & candidateNewDBValue & tail
                oldConn = conn

                ' ���ۂɍX�V���čă����N�B���s�����烍�[���o�b�N
                On Error Resume Next
                tdf.Connect = newConn
                tdf.RefreshLink
                If Err.Number <> 0 Then
                    lastErr = Err.Number
                    lastDesc = Err.Description
                    Err.Clear
                    ' ���[���o�b�N
                    tdf.Connect = oldConn
                    On Error Resume Next
                    tdf.RefreshLink
                    On Error GoTo ErrHandler
                    Call LogDebug(methodName, "�����N�X�V���s�i���[���o�b�N���{�j: " & tdf.name & _
                                         " new=" & candidateNewDBValue & " err=" & lastErr & " / " & lastDesc)
                Else
                    ' ����
                    On Error GoTo ErrHandler
'                    Call LogDebug(methodName, "�����N�X�V����: " & tdf.name & " -> " & candidateNewDBValue)
                End If
            End If
        End If
ContinueNext:
    Next

ErrHandler:
    Call ErrorSave(methodName)  '�K���擪�i���������G���[���m���ɃL���b�`���邽�߁j
    Call ProcNothing(tdf)
    Call ProcNothing(db)
    If ErrObj.Number <> 0 Then Err.Raise ERR_TMP, "", "�^���G���["  ' �K���Ō�(�㑱�����������Ȃ����߁j
    
End Sub


' **************************************************
' �œK���i�����N�t�@�C�� or ���t�@�C���j
' **************************************************
Public Sub Proc�t�@�C���œK��(Optional srcPath As String = "")
    Const methodName As String = "Proc�t�@�C���œK��"
    ' �錾��
    Dim tmpPath As String
On Error GoTo ErrHandler
    
    
    ' �p�X���w�肳��Ă��Ȃ���Ύ��t�@�C����Ώ�
    If Len(srcPath) = 0 Then
        srcPath = Proc�����N�e�[�u���ڑ���擾
        If srcPath = "" Then GoTo ErrHandler
    End If
    
    ' �ꎞ�t�@�C����
    tmpPath = Left(srcPath, InStrRev(srcPath, ".")) & "_�œK����.accdb"
    
    ' CompactDatabase ���s
    DBEngine.CompactDatabase srcPath, tmpPath
    
    ' ���t�@�C�����폜
    Kill srcPath
    ' �œK���ς݂̃t�@�C�������̖��O�ɖ߂�
    Name tmpPath As srcPath
    
    Call ShowToast("�f�[�^�t�@�C���̍œK���͊������܂����B", E_Color_Blue)

ErrHandler:
    Call ErrorSave(methodName)  '�K���擪�i���������G���[���m���ɃL���b�`���邽�߁j
    If ErrObj.Number <> 0 Then Err.Raise ERR_TMP, "", "�^���G���["  ' �K���Ō�(�㑱�����������Ȃ����߁j
    
End Sub




