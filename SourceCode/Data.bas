Attribute VB_Name = "Data"
Option Compare Database
Option Explicit




' **************************************************
' �X�ܖ���1���擾�iDLookup�Łj
' **************************************************
Public Function Get�X�ܖ�() As String
    Const methodName As String = "Get�X�ܖ�"
On Error Resume Next
    Get�X�ܖ� = Nz(DLookup("�X�ܖ�", "M_�X�܃}�X�^"), "")

ErrHandler:
    Call ErrorSave(methodName)  '�K���擪�i���������G���[���m���ɃL���b�`���邽�߁j
    If ErrObj.Number <> 0 Then Err.Raise ERR_TMP, "", "�^���G���["  ' �K���Ō�(�㑱�����������Ȃ����߁j
        
End Function

' ***********************************************
' �X�ܖ���ݒ肷��iM_�X�܃}�X�^��1���O��j
' ***********************************************
Public Sub Set�X�ܖ�(newName As String)
    Const methodName As String = "Set�X�ܖ�"
On Error GoTo ErrHandler

    CurrentDb.Execute "UPDATE M_�X�܃}�X�^ SET �X�ܖ� = '" & Replace(newName, "'", "''") & "';", dbFailOnError

ErrHandler:
    Call ErrorSave(methodName)  '�K���擪�i���������G���[���m���ɃL���b�`���邽�߁j
    If ErrObj.Number <> 0 Then Err.Raise ERR_TMP, "", "�^���G���["  ' �K���Ō�(�㑱�����������Ȃ����߁j
        
End Sub

' ***********************************************
' �X�ܑ}���N�G�������s
' ***********************************************
Public Sub Add�X�ܖ�()
    Const methodName As String = "Add�X�ܖ�"
On Error GoTo ErrHandler

    CurrentDb.Execute "DELETE FROM M_�X�܃}�X�^;", dbFailOnError
    CurrentDb.Execute "INSERT INTO M_�X�܃}�X�^ (�X�ܖ�) VALUES ('�X�ܖ�����͂��Ă�������');", dbFailOnError

ErrHandler:
    Call ErrorSave(methodName)  '�K���擪�i���������G���[���m���ɃL���b�`���邽�߁j
    If ErrObj.Number <> 0 Then Err.Raise ERR_TMP, "", "�^���G���["  ' �K���Ō�(�㑱�����������Ȃ����߁j
        
End Sub



' **************************************************
' CSV�C���|�[�g
' **************************************************
Function data�C���|�[�gCSV(csvPath As String, targetTable As String) As Boolean
    Const methodName As String = "data�C���|�[�gCSV"
    ' �錾��
    Dim fso As Object, ts As Object
    Dim db As DAO.Database
    Dim fields As Collection
    Dim sql As String
    Dim colCount As Long
    Dim i As Long
On Error GoTo ErrHandler

    data�C���|�[�gCSV = 0

    ' ���C������
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.OpenTextFile(csvPath, 1)
    Set db = CurrentDb
    
    ' �񐔊m�F
    colCount = db.TableDefs(targetTable).fields.count
    
    Do While Not ts.AtEndOfStream
        Set fields = Parse�P�s�ǂݍ���(ts.ReadLine)
        
        If colCount <> fields.count Then
            ' ����G���[
            Err.Raise ERR_BIZ, , "��荞�݃t�@�C���̃f�[�^�`������v���Ă��܂���B"
        End If
        
        If fields.count = colCount Then
            sql = "INSERT INTO " & targetTable & " VALUES ('"
            For i = 1 To fields.count
                sql = sql & Replace(fields(i), "'", "''")
                If i < fields.count Then sql = sql & "','"
            Next
            sql = sql & "')"
            db.Execute sql
        End If
    Loop

    ' ���ʃZ�b�g
    data�C���|�[�gCSV = Not DCount("*", targetTable) = 0
        
ErrHandler:
    Call ErrorSave(methodName)  '�K���擪�i���������G���[���m���ɃL���b�`���邽�߁j
    Call ProcNothing(ts)
    Call ProcNothing(fso)
    Call ProcNothing(db)
    If ErrObj.Number <> 0 Then Err.Raise ERR_TMP, "", "�^���G���["  ' �K���Ō�(�㑱�����������Ȃ����߁j
        
End Function

' **************************************************
Function Parse�P�s�ǂݍ���(line As String) As Collection
    Const methodName As String = "Parse�P�s�ǂݍ���"
    ' �錾��
    Dim result As New Collection
    Dim inQuotes As Boolean
    Dim i As Long
    Dim ch As String
    Dim field As String
On Error GoTo ErrHandler
    
    ' ���C������
    inQuotes = False
    field = ""
    
    For i = 1 To Len(line)
        ch = Mid(line, i, 1)
        
        If ch = """" Then
            ' �_�u���N�H�[�g�̏ꍇ
            If inQuotes And i < Len(line) And Mid(line, i + 1, 1) = """" Then
                ' �A������_�u���N�H�[�g �� " ��1�ǉ�
                field = field & """"
                i = i + 1
            Else
                ' �N�H�[�g�̊J��
                inQuotes = Not inQuotes
            End If
        ElseIf ch = "," And Not inQuotes Then
            ' �J���}��؂�i�N�H�[�g�O�̂݁j
            result.Add field
            field = ""
        Else
            field = field & ch
        End If
    Next
    
    ' �Ō�̃t�B�[���h�ǉ�
    result.Add field
    
    Set Parse�P�s�ǂݍ��� = result

ErrHandler:
    Call ErrorSave(methodName)  '�K���擪�i���������G���[���m���ɃL���b�`���邽�߁j
    If ErrObj.Number <> 0 Then Err.Raise ERR_TMP, "", "�^���G���["  ' �K���Ō�(�㑱�����������Ȃ����߁j
    
End Function


' ***********************************************
' �e�[�u���̃��R�[�h�������t���ō폜
' ***********************************************
Public Function data�e�[�u���N���A(tableName As String, Optional whereCondition As String = "") As Boolean
    Const methodName As String = "data�e�[�u���N���A"
    ' �錾��
    Dim sql As String
On Error GoTo ErrHandler

    data�e�[�u���N���A = False
    
    ' ��{��DELETE��
    sql = "DELETE FROM [" & tableName & "]"

    ' �������n���ꂽ��WHERE���t����
    If Trim(whereCondition) <> "" Then
        sql = sql & " WHERE " & whereCondition
    End If

    ' ���s
    CurrentDb.Execute sql, dbFailOnError
    data�e�[�u���N���A = True

ErrHandler:
    Call ErrorSave(methodName)  '�K���擪�i���������G���[���m���ɃL���b�`���邽�߁j
    If ErrObj.Number <> 0 Then Err.Raise ERR_TMP, "", "�^���G���["  ' �K���Ō�(�㑱�����������Ȃ����߁j
    
End Function


'' ***********************************************
'' �P��N�G�������s����
'' ***********************************************
'Function �N�G�����s_�P��(queryName As String) As Boolean
'    Const methodName As String = "�N�G�����s_�P��"
'On Error Resume Next
'    �N�G�����s_�P�� = False
'
'    ' ���C������
'    CurrentDb.QueryDefs(queryName).Execute dbFailOnError
'
'    �N�G�����s_�P�� = True
'
'ErrHandler:
'    Call ErrorSave(methodName)  '�K���擪�i���������G���[���m���ɃL���b�`���邽�߁j
'    If ErrObj.Number <> 0 Then Err.Raise ERR_TMP, "", "�^���G���["  ' �K���Ō�(�㑱�����������Ȃ����߁j
'
'End Function


' ***********************************************
' �z����̃N�G�������ׂĎ��s����
' ***********************************************
Function Proc�N�G�����X�g�ꊇ���s(queryNames As Variant) As Boolean
    Const methodName As String = "Proc�N�G�����X�g�ꊇ���s"
    ' �錾��
    Dim i As Long
    Dim db As DAO.Database
    Dim queryName As String
On Error GoTo ErrHandler
    
    ' ���C������
    Proc�N�G�����X�g�ꊇ���s = False
    
    ' ���C������
    Set db = CurrentDb
    
    For i = LBound(queryNames) To UBound(queryNames)
        queryName = queryNames(i)
        db.QueryDefs(queryName).Execute dbFailOnError
    Next i
    
    Proc�N�G�����X�g�ꊇ���s = True

ErrHandler:
    Call ErrorSave(methodName, False, queryName) '�K���擪�i���������G���[���m���ɃL���b�`���邽�߁j
    Call ProcNothing(db)
    If ErrObj.Number <> 0 Then Err.Raise ERR_TMP, "", "�^���G���["  ' �K���Ō�(�㑱�����������Ȃ����߁j
    
End Function

