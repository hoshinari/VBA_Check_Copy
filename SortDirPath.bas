Attribute VB_Name = "Module2"


Sub SortDirPath()
Attribute SortDirPath.VB_ProcData.VB_Invoke_Func = " \n14"
'
' SortDirPath Macro
'

'
  Dim col As Integer
  Dim row As Integer
  
  Dim count As Integer
  
  
  Dim start_col As Integer
  Dim start_row As Integer
  
  Dim env_group_line_num As Integer
  Dim server_line_num As Integer
  Dim env_user_line_num As Integer
  Dim dir_num As Integer
  
  env_group_line_num = 5
  server_line_num = 6
  env_user_line_num = 7
  
  start_col = 8
  start_row = 9
  
  col = start_col
  row = start_row
  
  
  
  
  
  While Cells(row, 6).Value <> ""
    If Cells(row, 1).Value = "�Z" Then
      dir_num = server_line_num
    Else
      dir_num = env_user_line_num
    End If
  
    While Cells(dir_num, col).Value <> ""
      If Cells(row, col).Value = "�Z" Then
        Cells(row, col).Select
        Selection.Copy
        Sheets("�o��").Select

        
        count = count + 1
      Else
          Debug.Print col; row
      End If
      col = col + 1
    Wend
    col = start_col
    row = row + 1
  Wend
  
  Debug.Print count
  
'        Cells(env_user_line_num).PasteSpecial Paste:=xlPasteValues
'    Range("H5:I5").Select
'    Selection.Copy
'    Sheets("�o��").Select
'    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
'        :=False, Transpose:=False
'    Sheets("�ꗗ").Select
'    Range("H6").Select
'    Application.CutCopyMode = False
'    Selection.Copy
'    Sheets("�o��").Select
'    Range("D8").Select
'    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
'        :=False, Transpose:=False
'    Sheets("�ꗗ").Select
'    Range("B9:F9").Select
'    Application.CutCopyMode = False
'    Selection.Copy
'    Sheets("�o��").Select
'    Range("E8").Select
'    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
'        :=False, Transpose:=False
End Sub
