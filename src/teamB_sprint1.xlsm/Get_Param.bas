Attribute VB_Name = "Get_Param"
Public Function getString(cell As String) As String
    '#####################
    '### 指定されたセルを読み取る
    '###  引数　　読み取り元のセル　ex)"A1"
    '###  戻り値　指定セルの値
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
