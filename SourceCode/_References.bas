Attribute VB_Name = "_References"
Option Compare Database
Option Explicit

'Const ProcName As String = "LibDocmd"


'********************************************************************************
' 参照設定
'********************************************************************************

'------------------------------------------------------------------
'VBA           4.2           {000204EF-0000-0000-C000-000000000046}
'Access        9.0           {4AFFC9A0-5F99-101B-AF4E-00AA003F0F07}
'stdole        2.0           {00020430-0000-0000-C000-000000000046}
'DAO           12.0          {4AC9E1DA-5BAD-4AC7-86E3-24F4CDCECA28}
'Office        2.8           {2DF8D04C-5BFA-101B-BDE5-00AA0044DE52}
'ADODB         6.1           {B691E011-1797-432E-907A-4D8C69339129}
'ADOX          6.0           {00000600-0000-0010-8000-00AA006D2EA4}
'VBIDE         5.3           {0002E157-0000-0000-C000-000000000046}
'Scripting     1.0           {420B2830-E718-11CF-893D-00A0C9054228}


'-------------------------------------------------------------
'■ 参照設定のguidを取得する
'-------------------------------------------------------------
Private Sub Pr_Check_RefGuid()
On Error Resume Next
    Dim ref As Reference
    Debug.Print "------------------------------------------------------------------"
    For Each ref In References
        Debug.Print ref.name, ref.Major & "." & ref.Minor, ref.Guid
    Next
    Set ref = Nothing
End Sub


'-------------------------------------------------------------
'■ 文字配列を指定文字で挟み込む
'-------------------------------------------------------------
Public Function P_AddRef() As Boolean
On Error Resume Next
    Application.VBE.ActiveVBProject.References.AddFromGuid "{2DF8D04C-5BFA-101B-BDE5-00AA0044DE52}", 2, 8
    Application.VBE.ActiveVBProject.References.AddFromGuid "{B691E011-1797-432E-907A-4D8C69339129}", 6, 1
    Application.VBE.ActiveVBProject.References.AddFromGuid "{00000600-0000-0010-8000-00AA006D2EA4}", 6, 0
    Application.VBE.ActiveVBProject.References.AddFromGuid "{0002E157-0000-0000-C000-000000000046}", 5, 3
    Application.VBE.ActiveVBProject.References.AddFromGuid "{420B2830-E718-11CF-893D-00A0C9054228}", 1, 0
    Application.VBE.ActiveVBProject.References.AddFromGuid "{73E709EA-7E85-11D1-ABD7-00C04FD97575}", 14, 0
    Application.VBE.ActiveVBProject.References.AddFromGuid "{2DF8D04C-5BFA-101B-BDE5-00AA0044DE52}", 2, 0
    
End Function


