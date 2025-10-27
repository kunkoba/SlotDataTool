Attribute VB_Name = "_License"
Option Compare Database
Option Explicit



Const KEY_LEN As Integer = 15  ' �������镶����
Const CODE_LEN_1 As Integer = 5 ' �����̒���
Const CODE_LEN_2 As Integer = 10 ' �����̒���
    
    
' **************************************************************
' �A�v����������
' **************************************************************
Public Sub TEST�A�v����������()
    Dim key1 As String
'    Dim key2 As String
'    Dim key_tmp As String
    Dim code1 As String
    Dim code2 As String

    Debug.Print Now
    
    Debug.Print "----�@�閧�R�[�h�����@----"
    key1 = Lic�閧�R�[�h����

    Debug.Print "----�@�Í��L�[�����@----"
    code1 = Lic�閧�R�[�h����Í��L�[(key1)

    Debug.Print "----�@�����R�[�h�����@----"
    code2 = Lic�Í��L�[��������R�[�h(code1)
    code2 = str�w�蕶�������ƂɎw�蕶��������(code2, "-", 5)
    

    Debug.Print "----�@���ʁ@----"
    Debug.Print Lic�����R�[�h�`�F�b�N(code1, code2)
    
    Debug.Print

End Sub



' **************************************************************
' �����R�[�h�`�F�b�N�i�Í��L�[�Ɣ�r�j
' **************************************************************
Public Function Lic�����R�[�h�`�F�b�N(�Í��L�[ As String, �����R�[�h As String) As String
    Const methodName As String = "Lic�����R�[�h�`�F�b�N"
    Dim key1 As String, key2 As String
On Error GoTo ErrHandler

    Lic�����R�[�h�`�F�b�N = False

    Debug.Print "Lic�����R�[�h�`�F�b�N1 >> ", �Í��L�[, �����R�[�h
    
    �����R�[�h = Replace(�����R�[�h, "-", "")   '�n�C�t�����O
    
    If Len(�����R�[�h) <> KEY_LEN + CODE_LEN_2 Then GoTo ErrHandler
    
    key1 = Lic�w�蕶���񂩂�閧�R�[�h(�Í��L�[)
    key2 = Lic�w�蕶���񂩂�閧�R�[�h(�����R�[�h)
    
    Debug.Print "Lic�����R�[�h�`�F�b�N2-> ", key1, key2
    Lic�����R�[�h�`�F�b�N = (str������������\�[�g(key1) = str������������\�[�g(key2))
    
ErrHandler:
    Call ErrorSave(methodName)  '�K���擪�i���������G���[���m���ɃL���b�`���邽�߁j
    If ErrObj.Number <> 0 Then Err.Raise ERR_TMP, "", "�^���G���["  ' �K���Ō�(�㑱�����������Ȃ����߁j

End Function



' **************************************************************
' **************************************************************
' **************************************************************
' **************************************************************
' **************************************************************
' �閧�R�[�h(10)�@���@�Í��L�[(20)�@���@�閧�R�[�h(10)�@���@�����R�[�h(25)
' **************************************************************
' �����_���p������
' **************************************************************
Public Function Lic�閧�R�[�h����() As String
    Const methodName As String = "Lic�閧�R�[�h����"
    Dim code1 As String, code2 As String
    Dim i As Integer
    Dim result As String
    Dim rndChar As Integer
On Error GoTo ErrHandler

    Randomize ' ����������
    
    result = ""
    For i = 1 To KEY_LEN
        ' A�`Z (ASCII�R�[�h 65�`90)
        rndChar = Int((26 * Rnd) + 65)
        result = result & Chr(rndChar)
    Next i
    
    Lic�閧�R�[�h���� = result
    Debug.Print "Lic�閧�R�[�h���� >> " & result
    
ErrHandler:
    Call ErrorSave(methodName)  '�K���擪�i���������G���[���m���ɃL���b�`���邽�߁j
    If ErrObj.Number <> 0 Then Err.Raise ERR_TMP, "", "�^���G���["  ' �K���Ō�(�㑱�����������Ȃ����߁j

End Function


' **************************************************************
Public Function Lic�閧�R�[�h����Í��L�[(ByVal key As String) As String
On Error Resume Next
    Lic�閧�R�[�h����Í��L�[ = Lic�閧�R�[�h����w�蕶��������(key, CODE_LEN_1)
    
End Function


' **************************************************************
Public Function Lic�Í��L�[��������R�[�h(ByVal code As String) As String
    Dim temp As String
On Error Resume Next
    temp = Lic�w�蕶���񂩂�閧�R�[�h(code)
    temp = Lic�閧�R�[�h����w�蕶��������(temp, CODE_LEN_2)
    Lic�Í��L�[��������R�[�h = str�w�蕶�������ƂɎw�蕶��������(temp, "-", 5)
    
End Function



' **************************************************************
' �p�����J���t���[�W������
' **************************************************************
Private Function Lic�閧�R�[�h����w�蕶��������(ByVal src As String, codeNum As Integer) As String
    Const methodName As String = "Lic�閧�R�[�h����w�蕶��������"
    Dim i As Integer
    Dim ch As String
    Dim result As String
    Dim rndNum As String
    Dim mixStr As String
    Dim shuffled As String
On Error GoTo ErrHandler

    Randomize
    
    '--- �p���������_���ŏ������� ---
    result = ""
    For i = 1 To Len(src)
        ch = Mid(src, i, 1)
        If Rnd < 0.5 Then
            result = result & LCase(ch)
        Else
            result = result & UCase(ch)
        End If
    Next i
    
    '--- �����_�����l�𐶐� ---
    rndNum = ""
    For i = 1 To codeNum
        rndNum = rndNum & CStr(Int((10 * Rnd)))
    Next i
    
    '--- �p���{���������� ---
    mixStr = result & rndNum
    
    '--- �����_�����ёւ� ---
    shuffled = str��������o���o���\�[�g(mixStr)
    
    Lic�閧�R�[�h����w�蕶�������� = shuffled
    
    Debug.Print "Lic�閧�R�[�h����w�蕶�������� >> " & shuffled
    
ErrHandler:
    Call ErrorSave(methodName)  '�K���擪�i���������G���[���m���ɃL���b�`���邽�߁j
    If ErrObj.Number <> 0 Then Err.Raise ERR_TMP, "", "�^���G���["  ' �K���Ō�(�㑱�����������Ȃ����߁j

End Function

' **************************************************************
' �����_���������ꂽ���������Ƃɖ߂�
' **************************************************************
Private Function Lic�w�蕶���񂩂�閧�R�[�h(ByVal camo As String) As String
    Const methodName As String = "Lic�w�蕶���񂩂�閧�R�[�h"
    Dim i As Integer
    Dim ch As String
    Dim result As String
On Error GoTo ErrHandler
    
    Lic�w�蕶���񂩂�閧�R�[�h = "X"
    
    ' 1�������`�F�b�N
    For i = 1 To Len(camo)
        ch = Mid(camo, i, 1)
        
        ' ���������O���ĉp�����������
        If ch Like "[A-Za-z]" Then
            result = result & UCase(ch)
        End If
    Next i
    
    Lic�w�蕶���񂩂�閧�R�[�h = result
    Debug.Print "Lic�w�蕶���񂩂�閧�R�[�h >> " & result
    
ErrHandler:
    Call ErrorSave(methodName)  '�K���擪�i���������G���[���m���ɃL���b�`���邽�߁j
    If ErrObj.Number <> 0 Then Err.Raise ERR_TMP, "", "�^���G���["  ' �K���Ō�(�㑱�����������Ȃ����߁j

End Function




' **************************************************************
' ������������_���ɕ��ёւ���w���p�[�֐�
' **************************************************************
Private Function str��������o���o���\�[�g(ByVal src As String) As String
    Const methodName As String = "str��������o���o���\�[�g"
    Dim i As Integer, j As Integer
    Dim arr() As String
    Dim temp As String
    Dim result As String
On Error GoTo ErrHandler
    
    ReDim arr(1 To Len(src))
    
    ' ������z��Ɋi�[
    For i = 1 To Len(src)
        arr(i) = Mid(src, i, 1)
    Next i
    
    ' Fisher?Yates shuffle
    For i = Len(src) To 2 Step -1
        j = Int(i * Rnd) + 1
        temp = arr(i)
        arr(i) = arr(j)
        arr(j) = temp
    Next i
    
    ' �z��𕶎���ɖ߂�
    result = ""
    For i = 1 To Len(src)
        result = result & arr(i)
    Next i
    
    str��������o���o���\�[�g = result
    
ErrHandler:
    Call ErrorSave(methodName)  '�K���擪�i���������G���[���m���ɃL���b�`���邽�߁j
    If ErrObj.Number <> 0 Then Err.Raise ERR_TMP, "", "�^���G���["  ' �K���Ō�(�㑱�����������Ȃ����߁j

End Function

' **************************************************************
' ������������ŕ��ёւ���֐�
' **************************************************************
Private Function str������������\�[�g(ByVal src As String) As String
    Const methodName As String = "str������������\�[�g"
    Dim arr() As String
    Dim i As Long, j As Long
    Dim temp As String
    Dim result As String
On Error GoTo ErrHandler
    
    ' �������z��ɕ���
    ReDim arr(1 To Len(src))
    For i = 1 To Len(src)
        arr(i) = Mid(src, i, 1)
    Next i
    
    ' �o�u���\�[�g�i�P���\�[�g�j
    For i = 1 To UBound(arr) - 1
        For j = i + 1 To UBound(arr)
            If arr(i) > arr(j) Then
                temp = arr(i)
                arr(i) = arr(j)
                arr(j) = temp
            End If
        Next j
    Next i
    
    ' �z�������
    result = ""
    For i = 1 To UBound(arr)
        result = result & arr(i)
    Next i
    
    str������������\�[�g = result
    
ErrHandler:
    Call ErrorSave(methodName)  '�K���擪�i���������G���[���m���ɃL���b�`���邽�߁j
    If ErrObj.Number <> 0 Then Err.Raise ERR_TMP, "", "�^���G���["  ' �K���Ō�(�㑱�����������Ȃ����߁j

End Function


' **************************************************
Private Function str�w�蕶�������ƂɎw�蕶��������(src As String, strChar As String, num As Integer) As String
    Const methodName As String = "str�w�蕶�������ƂɎw�蕶��������"
    Dim i As Long
    Dim result As String
On Error GoTo ErrHandler
    
    For i = 1 To Len(src)
        result = result & Mid(src, i, 1)
        ' �w�萔���ƂɃn�C�t����ǉ��i�����ȊO�j
        If i Mod num = 0 And i <> Len(src) Then
            result = result & strChar
        End If
    Next i
    
    str�w�蕶�������ƂɎw�蕶�������� = result
    
ErrHandler:
    Call ErrorSave(methodName)  '�K���擪�i���������G���[���m���ɃL���b�`���邽�߁j
    If ErrObj.Number <> 0 Then Err.Raise ERR_TMP, "", "�^���G���["  ' �K���Ō�(�㑱�����������Ȃ����߁j
    
End Function



'' ==============================================================
'' ������S�̂��V�t�g
'' ==============================================================
'Private Function str�V�t�g������(ByVal s As String, ByVal Shift As Integer) As String
'    Dim i As Long
'    Dim result As String
'    result = ""
'
'    For i = 1 To Len(s)
'        result = result & str�V�t�g����(Mid$(s, i, 1), Shift)
'    Next i
'
'    str�V�t�g������ = result
'
'End Function
'' ==============================================================
'' �⏕�F�������V�t�g
'' ==============================================================
'Private Function str�V�t�g����(ByVal ch As String, ByVal Shift As Integer) As String
'    Dim code As Integer
'    code = Asc(ch)
'
'    ' 0-9 -> 48-57
'    If code >= 48 And code <= 57 Then
'        str�V�t�g���� = Chr(((code - 48 + Shift) Mod 10) + 48)
'        Exit Function
'    End If
'
'    ' A-Z -> 65-90
'    If code >= 65 And code <= 90 Then
'        str�V�t�g���� = Chr(((code - 65 + Shift) Mod 26) + 65)
'        Exit Function
'    End If
'
'    ' a-z -> 97-122
'    If code >= 97 And code <= 122 Then
'        str�V�t�g���� = Chr(((code - 97 + Shift) Mod 26) + 97)
'        Exit Function
'    End If
'
'    ' ����ȊO�͂��̂܂ܕԂ�
'    str�V�t�g���� = ch
'
'End Function






