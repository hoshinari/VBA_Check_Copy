Attribute VB_Name = "Module2"


Sub SortDirPath()
Attribute SortDirPath.VB_ProcData.VB_Invoke_Func = " \n14"
'
' SortDirPath Macro
' test

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
  
  Dim output_row As Integer
  
  env_group_line_num = 5
  server_line_num = 6
  env_user_line_num = 7
  
  start_col = 8
  start_row = 9
  
  col = start_col
  row = start_row
  
  output_row = 10
  
  Workbooks.Open "C:\Users\hoshi\Desktop\macro2.xlsm"
  Worksheets("List").Select
  While Cells(row, 6).Value <> ""
    If Cells(row, 1).Value = "ÅZ" Then
      dir_num = server_line_num
    Else
      dir_num = env_user_line_num
    End If

    While Cells(dir_num, col).Value <> ""
      ' to right edge
      If Cells(row, col).Value = "ÅZ" Then
        Cells(row, col).Select
        Range(Cells(row, 2), Cells(row, 6)).Select
        Selection.Copy
        Worksheets("Output2").Select
        Range(Cells(output_row, 4), Cells(output_row, 8)).PasteSpecial xlPasteValues
        
        Worksheets("List").Select
        Cells(dir_num, col).Select
        Selection.Copy
        Worksheets("Output2").Select
        Cells(output_row, 3).PasteSpecial xlPasteValues
        
        Worksheets("List").Select
        Cells(5, col).Select
        Selection.Copy
        Worksheets("Output2").Select
        Cells(output_row, 2).PasteSpecial xlPasteValues
        Worksheets("List").Select
        
        count = count + 1
        output_row = output_row + 1

      Else
          Debug.Print col; row
      End If
      col = col + 1
    Wend
    col = start_col
    row = row + 1
  Wend
  
  Debug.Print count
  
End Sub
