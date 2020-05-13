Attribute VB_Name = "Main"
Public Function copyCell()
    '  MsgBox "Hello"
  
    ' 設定するパラメータ
    ' pathCell    集約対象ファイルのディレクトリが記載されたセル
    ' sheetCell   集約対象ファイルのシート名が記載されたセル
    ' targetCell  集約対象ファイルのコピー対象のセルが記載されたセル
    ' startRowCell 集約後のシートを書き始める位置（行）が記載されたセル
    ' startColCell 集約後のシートを書き始める位置（列）が記載されたセル
    Dim pathCell As String, path As String
    Dim sheetCell As String, sheet As String
    Dim targetCell As String, target As String
    Dim startRowCell As String, startRow As Long
    Dim startColCell As String, startCol As Long
  
    pathCell = "C6"
    sheetCell = "C7"
    targetCell = "C8"
    startRowCell = "C9"
    startColCell = "C10"
  
    path = Get_Param.getString(pathCell)
    sheet = Get_Param.getString(sheetCell)
    target = Get_Param.getString(targetCell)
    startRow = Get_Param.getLong(startRowCell)
    startCol = Get_Param.getLong(startColCell)
    
    Debug.Print ("path: " & path)
    Debug.Print ("sheet: " & sheet)
    Debug.Print ("target: " & target)
    Debug.Print ("start row: " & startRow)
    Debug.Print ("start col: " & startCol)
    'Debug.Print (path & " " & sheet & " " & target & " " & startRow & " " & startCol)
    
    Call Open_File_and_Copy.openFileAndCopy(path, sheet, target, startRow, startCol)
  
    Debug.Print ("End")
End Function
