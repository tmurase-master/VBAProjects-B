Attribute VB_Name = "Get_Param"
Public Function getString(cell As String) As String
    '#####################
    '### �w�肳�ꂽ�Z����ǂݎ��
    '###  �����@�@�ǂݎ�茳�̃Z���@ex)"A1"
    '###  �߂�l�@�w��Z���̒l
    '#####################
    
    Dim tmp As String
    tmp = ThisWorkbook.Sheets(1).Range(cell)
    
    getString = tmp

End Function

Public Function getLong(cell As String) As Long
    
    Dim tmp As Long
    tmp = ThisWorkbook.Sheets(1).Range(cell)
    
    getLong = tmp

End Function
