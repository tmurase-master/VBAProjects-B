Attribute VB_Name = "Main"
Public Function copyCell()
    '  MsgBox "Hello"
  
    ' �ݒ肷��p�����[�^
    ' pathCell    �W��Ώۃt�@�C���̃f�B���N�g�����L�ڂ��ꂽ�Z��
    ' sheetCell   �W��Ώۃt�@�C���̃V�[�g�����L�ڂ��ꂽ�Z��
    ' targetCell  �W��Ώۃt�@�C���̃R�s�[�Ώۂ̃Z�����L�ڂ��ꂽ�Z��
    ' startRowCell �W���̃V�[�g�������n�߂�ʒu�i�s�j���L�ڂ��ꂽ�Z��
    ' startColCell �W���̃V�[�g�������n�߂�ʒu�i��j���L�ڂ��ꂽ�Z��
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
