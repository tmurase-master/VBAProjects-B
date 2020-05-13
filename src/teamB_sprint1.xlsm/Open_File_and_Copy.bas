Attribute VB_Name = "Open_File_and_Copy"
Public Sub openFileAndCopy(path As String, sheetName As String, scp As String, startRow As Long, startCol As Long)
    '#####################
    '### �ϐ��Ŏw�肳�ꂽ�f�B���N�g�����̃t�@�C�����P�t�@�C�����J��
    '###  �����@�f�B���N�g �� ,  �V�[�g��, �W��͈�, �W���̃X�^�[�g�Z��(�s�E��)
    '#####################
    
    Debug.Print ("called open file function")
    
    '  �Ώۂ�xlsx
    'mac ����Dir�֐������������삵�Ȃ��\������
    Dim fileNames As String
    fileNames = dir(path & "*.xlsx")
    Debug.Print (fileNames)
    
    If fileNames = "" Then
        MsgBox "Excel�t�@�C��������܂���B"
        Exit Sub
    End If
    
    Dim openFileName As String
    rowNum = startRow
    '  �w��f�B���N�g�����̃t�@�C���������J��
    Do While fileNames <> ""
        ' �W��Ώۂ̃t�@�C����
        openFileName = path & fileNames
        Debug.Print (openFileName)
        ' �O�̂��߂�readonly�ŊJ��
        Workbooks.Open fileName:=openFileName, ReadOnly:=True
        
        ' �w��Z���̓��e���R�s�[
        Debug.Print (" row num: " & rowNum)
        Debug.Print (" content: " & Worksheets(sheetName).Range(scp))
        ThisWorkbook.Sheets(2).Cells(rowNum, startCol).Value = Worksheets(sheetName).Range(scp)
        
        ' �W��Ώۃt�@�C���̃N���[�Y
        Workbooks(fileNames).Close
        
        '���̃t�@�C�������w��
        fileNames = dir()
        '���̍s�ֈړ�
        rowNum = rowNum + 1
    Loop
  
End Sub
