Attribute VB_Name = "Open_File_and_Copy"
Public Sub openFileAndCopy(path As String, sheetName As String, scp As String, startRow As Long, startCol As Long)
    '#####################
    '### 変数で指定されたディレクトリ内のファイルを１ファイルずつ開く
    '###  引数　ディレクト 名 ,  シート名, 集約範囲, 集約後のスタートセル(行・列)
    '#####################
    
    Debug.Print ("called open file function")
    
    '  対象はxlsx
    'mac だとDir関数が正しく動作しない可能性あり
    Dim fileNames As String
    fileNames = dir(path & "*.xlsx")
    Debug.Print (fileNames)
    
    If fileNames = "" Then
        MsgBox "Excelファイルがありません。"
        Exit Sub
    End If
    
    Dim openFileName As String
    rowNum = startRow
    '  指定ディレクトリ内のファイルを順次開く
    Do While fileNames <> ""
        ' 集約対象のファイル名
        openFileName = path & fileNames
        Debug.Print (openFileName)
        ' 念のためにreadonlyで開く
        Workbooks.Open fileName:=openFileName, ReadOnly:=True
        
        ' 指定セルの内容をコピー
        Debug.Print (" row num: " & rowNum)
        Debug.Print (" content: " & Worksheets(sheetName).Range(scp))
        ThisWorkbook.Sheets(2).Cells(rowNum, startCol).Value = Worksheets(sheetName).Range(scp)
        
        ' 集約対象ファイルのクローズ
        Workbooks(fileNames).Close
        
        '次のファイル名を指定
        fileNames = dir()
        '次の行へ移動
        rowNum = rowNum + 1
    Loop
  
End Sub
