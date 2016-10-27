# Test

Sub LoopThroughDirectory()
Dim MyFile As String
Dim erow As Integer
Dim Filepath As String
Dim shName As String
Dim rCell As Range
Dim i&

Filepath = "M:\CTOps\CSD\CSD-T&P\Process\MIS\Individual Countsheets\2016\October\"
MyFile = Dir(Filepath)
Do While Len(MyFile) > 0
    If MyFile = "zmaster.xlsm" Then
    Exit Sub
    End If
    
    
    Workbooks.Open Filepath & MyFile, ReadOnly:=True
    Application.DisplayAlerts = False
    With Application
   .ScreenUpdating = False
   .EnableEvents = False
    End With
    shName = InputBox("Enter sheet name")
    Set rCell = Worksheets(shName).Range("B2")
    
   If Len(rCell.Formula) = 0 Then
     
     MsgBox "This Sheet is empty! Please check this file later"
        Else
        Worksheets(shName).Range("B2:N5000").Copy
       End If

    Worksheets(shName).Range("B2:O5000").Copy
    
    ActiveWorkbook.Close
    
    erow = Sheet2.Cells(Rows.Count, 1).End(xlUp).Offset(1, 0).Row
    
    ActiveSheet.Paste Destination:=Worksheets("Sheet2").Range(Cells(erow, 1), Cells(erow, 15))
    Application.DisplayAlerts = False
    Application.DisplayAlerts = True
    
    MyFile = Dir
    
Loop

Set Rng = Worksheets("Sheet2").Range("A2:R2000")
  
  'Apply new borders
    With Application
    Rng.BorderAround xlContinuous
    Rng.Borders(xlInsideHorizontal).LineStyle = xlContinuous
    Rng.Borders(xlInsideVertical).LineStyle = xlContinuous
    Rng.HorizontalAlignment = xlCenter
    End With
    
  'Delete rows if cells from D to I are empty
  For i = 1000 To 1 Step -1
    If WorksheetFunction.CountA(Range("D" & i, "I" & i)) = 0 Then
        Rows(i).EntireRow.Delete
    End If
Next i

With Application
.ScreenUpdating = True
.EnableEvents = True
End With

End Sub
