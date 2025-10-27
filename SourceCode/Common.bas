Attribute VB_Name = "Common"
Option Compare Database
Option Explicit



' ***********************************************
' �Ƃɂ����N���A
' ***********************************************
Public Sub ProcNothing(ByRef obj)
On Error Resume Next
    obj.Close
    Set obj = Nothing

'    Debug.Print Now, "ProcNothing", Err.Number, Err.Description
    Err.Clear
End Sub

'' ***********************************************
'' ���ʊJ�n����
'' ***********************************************
'Public Sub ProcInitSetting()
'On Error Resume Next
'    DoCmd.SetWarnings False
'
'End Sub

'' ***********************************************
'' ���ʏI������
'' ***********************************************
'Public Sub ProcFinally()
'On Error Resume Next
'    DoCmd.SetWarnings True
'    DoCmd.Echo True
'    Screen.MousePointer = 0    ' 0 = �ʏ�̖��
'
'End Sub


' ******************************************************
' �t�H�[�����J���Ă��邩�ǂ���
' ******************************************************
Function IsFormOpen(formName As String) As Boolean
On Error Resume Next
    If (SysCmd(acSysCmdGetObjectState, acForm, formName) And acObjStateOpen) <> 0 Then
        If Forms(formName).CurrentView = 1 Then
            IsFormOpen = True
        End If
    End If
   
End Function

' ***********************************************
' �����񌋍��֐�
' ***********************************************
Function JoinText(word1 As String, word2 As String, Optional addNewLine As Boolean = False) As String
On Error Resume Next

    If addNewLine Then
        JoinText = word1 & word2 & vbCrLf
    Else
        JoinText = word1 & word2
    End If
    
End Function



' ***********************************************
' �Z���`�ϊ�
' ***********************************************
Public Function CmToTwips(cm As Double) As Long
On Error Resume Next
    CmToTwips = cm * 567
    
End Function



' ***********************************************
' �I�u�W�F�N�g���擾����i�N�G���j
' ***********************************************
Public Function Get�N�G���ꗗ(Optional filter As String = "") As Variant
    Const methodName As String = "Get�N�G���ꗗ"
    ' �錾��
    Dim db As DAO.Database
    Dim qdf As DAO.QueryDef
    Dim names() As String
    Dim count As Long
On Error GoTo ErrHandler
    
    ' ���C������
    Set db = CurrentDb
    count = 0

    For Each qdf In db.QueryDefs
        If Left(qdf.name, 1) <> "~" Then ' �e���|�����N�G�����O
            If filter = "" Or qdf.name Like filter Then
                ReDim Preserve names(count)
                names(count) = qdf.name
                count = count + 1
            End If
        End If
    Next

    If count = 0 Then
        Get�N�G���ꗗ = Array() ' ��z��
    Else
        Get�N�G���ꗗ = names
    End If

ErrHandler:
    Call ErrorSave(methodName)  '�K���擪�i���������G���[���m���ɃL���b�`���邽�߁j
    Call ProcNothing(qdf)
    Call ProcNothing(db)
    If ErrObj.Number <> 0 Then Err.Raise ERR_TMP, "", "�^���G���["  ' �K���Ō�(�㑱�����������Ȃ����߁j
        
End Function






' ***********************************************
' �F�R�[�h�ϊ�
' ***********************************************
Public Function ConvertHexColorToRGB(hexColor As String) As Long
    Dim r As Integer
    Dim g As Integer
    Dim b As Integer
On Error Resume Next
    ' ��: hexColor = "#CDDCAF"

    ' "#" ����菜��
    If Left(hexColor, 1) = "#" Then
        hexColor = Mid(hexColor, 2)
    End If

    ' R, G, B ��16�i��10�i�ɕϊ�
    r = CInt("&H" & Mid(hexColor, 1, 2))
    g = CInt("&H" & Mid(hexColor, 3, 2))
    b = CInt("&H" & Mid(hexColor, 5, 2))

    ' RGB�֐��ŐF��Ԃ��iLong�^�j
    ConvertHexColorToRGB = RGB(r, g, b)

End Function


' ***********************************************
' �N�G����CSV�ŏo�͂���֐�
' ����:
'   queryName  �c �N�G�����i��: "Q_�o�̓f�[�^"�j
'   folderPath �c �o�̓t�H���_�i������ \ �͕s�v�ł��j
'   fileName   �c �o�̓t�@�C�����i.csv �g���q�͎����t�^�j
' ***********************************************
Public Sub �N�G���o��ToCSV(queryName As String, folderPath As String, fileName As String)
    Const methodName As String = "�N�G���o��ToCSV"
    ' �錾��
    Dim fullPath As String
On Error GoTo ErrHandler
    
    ' �t�H���_�����Ƀo�b�N�X���b�V����t����
    If Right(folderPath, 1) <> "\" Then
        folderPath = folderPath & "\"
    End If

    ' .csv �g���q���Ȃ���Βǉ�
    If LCase(Right(fileName, 4)) <> ".csv" Then
        fileName = fileName & ".csv"
    End If

    ' �t���p�X�쐬
    fullPath = folderPath & fileName

    ' �N�G����CSV�ŃG�N�X�|�[�g
    DoCmd.TransferText _
        TransferType:=acExportDelim, _
        tableName:=queryName, _
        fileName:=fullPath, _
        HasFieldNames:=True

ErrHandler:
    Call ErrorSave(methodName)  '�K���擪�i���������G���[���m���ɃL���b�`���邽�߁j
    If ErrObj.Number <> 0 Then Err.Raise ERR_TMP, "", "�^���G���["  ' �K���Ō�(�㑱�����������Ȃ����߁j
    
End Sub





' ******************************************************
' ���t�����^����t�֕ϊ��iyyyymmdd / yymmdd ���Ή��j
' yymmdd �̏ꍇ�͐擪�� 20 ��t���ĕϊ�
' ******************************************************
Function Convert�����񂩂���t(strDate As Variant) As Variant
    Const methodName As String = "Convert�����񂩂���t"
    Dim s As String
    Dim Y As Long, m As Long, d As Long
    Dim yyyy As Long
On Error GoTo ErrHandler
    
    ' Null�`�F�b�N
    If IsNull(strDate) Then
        Convert�����񂩂���t = Null
        Exit Function
    End If

    s = Trim(CStr(strDate))
    
    ' yyyymmdd
    If Len(s) = 8 And s Like "########" Then
        Y = CLng(Left(s, 4))
        m = CLng(Mid(s, 5, 2))
        d = CLng(Right(s, 2))
        Convert�����񂩂���t = DateSerial(Y, m, d)
        Exit Function
    End If
    
    ' yymmdd �� 20yy mm dd
    If Len(s) = 6 And s Like "######" Then
        yyyy = 2000 + CLng(Left(s, 2))
        m = CLng(Mid(s, 3, 2))
        d = CLng(Right(s, 2))
        
        Convert�����񂩂���t = DateSerial(yyyy, m, d)
        Exit Function
    End If
    
    ' ����ȊO�� Null
    Convert�����񂩂���t = Null
    
ErrHandler:
    Call ErrorSave(methodName)  '�K���擪�i���������G���[���m���ɃL���b�`���邽�߁j
    If ErrObj.Number <> 0 Then Err.Raise ERR_TMP, "", "�^���G���["  ' �K���Ō�(�㑱�����������Ȃ����߁j
    
End Function




' ***********************************************
' ���l�Ȃ炻�̂܂܁A������Ȃ�V���O���N�H�[�e�[�V�����ň͂�
' ���t�^�Ȃ� # �ň͂�
' ***********************************************
Public Function Convert�����l������(value As Variant) As String
    Const methodName As String = "Convert�����l������"
On Error GoTo ErrHandler

    If IsNull(value) Then
        Convert�����l������ = "Null"
    ElseIf IsDate(value) Then
        ' ���t�� # �ň͂��SQL�p�ɐ��`
        Convert�����l������ = "#" & Format(value, "yyyy/mm/dd") & "#"
    ElseIf IsNumeric(value) Then
        Convert�����l������ = CStr(value)
    Else
        ' ������ �� �V���O���N�H�[�g�ň͂݁A�V���O���N�H�[�g���G�X�P�[�v
        Convert�����l������ = "'" & Replace(value, "'", "''") & "'"
    End If
    
ErrHandler:
    Call ErrorSave(methodName)  '�K���擪�i���������G���[���m���ɃL���b�`���邽�߁j
    If ErrObj.Number <> 0 Then Err.Raise ERR_TMP, "", "�^���G���["  ' �K���Ō�(�㑱�����������Ȃ����߁j
    
End Function



' ======================================================
' PC���擾�֐��i�z��ŕԂ��j
' ======================================================
Public Function Get�p�\�R�����() As Variant
    Const methodName As String = "Get�p�\�R�����"
    Dim infoArr(2) As String
    Dim pcName As String
    Dim userName As String
    Dim ipAddr As String
    Dim wmi As Object, colItems As Object, objItem As Object
On Error GoTo ErrHandler
    
    ' �p�\�R����
    pcName = Environ("COMPUTERNAME")
    
    ' ���[�U�[��
    userName = Environ("USERNAME")
    
    ' IP�A�h���X�擾�i�A�N�e�B�u�Ȑڑ�1���j
    Set wmi = GetObject("winmgmts:\\.\root\cimv2")
    Set colItems = wmi.ExecQuery("SELECT * FROM Win32_NetworkAdapterConfiguration WHERE IPEnabled = True")
    
    ipAddr = ""
    For Each objItem In colItems
        If Not IsNull(objItem.IPAddress) Then
            ipAddr = objItem.IPAddress(0)
            Exit For
        End If
    Next
    
    ' �z��Ɋi�[
    infoArr(0) = pcName
    infoArr(1) = userName
    infoArr(2) = ipAddr
    
    Get�p�\�R����� = infoArr
    
ErrHandler:
    Call ErrorSave(methodName)  '�K���擪�i���������G���[���m���ɃL���b�`���邽�߁j
    If ErrObj.Number <> 0 Then Err.Raise ERR_TMP, "", "�^���G���["  ' �K���Ō�(�㑱�����������Ȃ����߁j
    
End Function


' ============================================================
' �֐��� : �z��v�f���݃`�F�b�N
' �@�@�\ : �w�肵�������񂪔z��̂����ꂩ�̗v�f�Ɋ܂܂�Ă��邩�𔻒肷��
' ���@�� : targetArray - �����Ώۂ̔z��iVariant�^�����j
'�@�@�@ : searchValue  - �T������������
' �߂�l : Boolean�i�܂܂�Ă���� True�j
' ============================================================
Public Function �z��v�f���݃`�F�b�N(targetArray As Variant, searchValue As String) As Boolean
    Const methodName As String = "�z��v�f���݃`�F�b�N"
    Dim v As Variant
On Error GoTo ErrHandler
    �z��v�f���݃`�F�b�N = False  ' �����l
    
    ' ��z��`�F�b�N
    If IsEmpty(targetArray) Then Exit Function
    If IsNull(searchValue) Or searchValue = "" Then Exit Function
    
    ' �z��̊e�v�f�����[�v���Ĉ�v���m�F
    For Each v In targetArray
        If v = searchValue Then
            �z��v�f���݃`�F�b�N = True
            Exit Function
        End If
    Next v
    
ErrHandler:
    Call ErrorSave(methodName)  '�K���擪�i���������G���[���m���ɃL���b�`���邽�߁j
    If ErrObj.Number <> 0 Then Err.Raise ERR_TMP, "", "�^���G���["  ' �K���Ō�(�㑱�����������Ȃ����߁j
    
End Function



