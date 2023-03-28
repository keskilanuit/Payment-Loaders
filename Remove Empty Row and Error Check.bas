Sub RemovedEmptyRowAndErrorCheck()
    
    
    On Error Resume Next
    
    Worksheets("Master EFT").Range("Q4:Q" & Worksheets("Master EFT").UsedRange.Rows.Count).SpecialCells(xlCellTypeBlanks).EntireRow.Delete
    Worksheets("Master EFT").Activate
    Worksheets("Master EFT").Range("T2:Z2").Select
    Selection.Copy
    Worksheets("Master EFT").Range("T4:Z4").PasteSpecial
    Worksheets("Master EFT").Range("T4:Z4").Copy
    Worksheets("Master EFT").Range("Q4").Select
    Selection.End(xlDown).Select
    Selection.Offset(0, 3).Select
    ActiveSheet.Paste
    Worksheets("Master EFT").Range("T4:Z4").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.PasteSpecial
    
    
    Worksheets("Master EFT").Activate
    Worksheets("Master EFT").Range("A4:N2000").Select
    With Selection.Font
        .Name = "Arial"
        .Size = 9
    End With
    Worksheets("Master EFT").Range("A4:S2000").Font.Bold = False
    
    Worksheets("Master EFT").Range("I4:I2000").Font.Bold = True
    Worksheets("Master EFT").Range("O4:O2000").Select
    With Selection.Font
        .Name = "Arial"
        .Size = 9
    End With
    
    Worksheets("Master EFT").Range("P4:S2000").Select
    With Selection.Font
        .Name = "Arial"
        .Size = 9
    End With
    
  
    Worksheets("Master EFT").Range("A4:C2000").Select
    With Selection.Font
        .Color = RGB(0, 0, 255)
      End With
      Worksheets("Master EFT").Range("D4:H2000").Select
    With Selection.Font
        .Color = RGB(0, 0, 0)
      End With
    Worksheets("Master EFT").Range("I4:I2000").Select
    With Selection.Font
        .Color = RGB(0, 0, 255)
      End With
     Worksheets("Master EFT").Range("K4:M2000").Select
    With Selection.Font
        .Color = RGB(0, 0, 0)
    End With
    Worksheets("Master EFT").Range("N4:N2000").Select
    With Selection.Font
        .Color = RGB(0, 0, 255)
    End With
    Worksheets("Master EFT").Range("O4:O2000").Select
    With Selection.Font
        .Color = RGB(0, 0, 0)
    End With
    Worksheets("Master EFT").Range("R4:S2000").Select
    With Selection.Font
        .Color = RGB(0, 0, 255)
    End With
    
   
    Worksheets("Master EFT").Range("J4").Interior.ColorIndex = 48
    Worksheets("Master EFT").Range("Q4").Select
    Selection.End(xlDown).Select
    Selection.Offset(0, -7).Select
    Range(Selection, Range("J4")).Interior.ColorIndex = 48
    Worksheets("Master EFT").Range("K4:S2000").Interior.ColorIndex = 0
    Worksheets("Master EFT").Range("K4:S2000").Borders.LineStyle = xlNone
    Worksheets("Master EFT").Range("F4:F2000").Interior.ColorIndex = 0
    
    Worksheets("Master EFT").Range("A4:A2000").HorizontalAlignment = xlCenter
    Worksheets("Master EFT").Range("I4:I2000").HorizontalAlignment = xlCenter
    Worksheets("Master EFT").Range("K4:K2000").HorizontalAlignment = xlRight
    Worksheets("Master EFT").Range("I4:I2000").HorizontalAlignment = xlCenter
    Worksheets("Master EFT").Range("M4:N2000").HorizontalAlignment = xlCenter
    Worksheets("Master EFT").Range("P4:Q2000").HorizontalAlignment = xlRight
    Worksheets("Master EFT").Range("R4:R2000").HorizontalAlignment = xlCenter
    
       
      Count = Worksheets("Master EFT").Cells(Rows.Count, "P").End(xlUp).Row
        i = 4
    Do While i <= Count
    If Cells(i, 16) = "0" Then
    Rows(i).EntireRow.Delete
    i = i - 1
    End If
    i = i + 1
    Count = Worksheets("Master EFT").Cells(Rows.Count, "P").End(xlUp).Row
    
    Loop

   
    Worksheets("Master EFT").Range("B2").Copy
    Worksheets("Master EFT").Range("B2").PasteSpecial Paste:=xlPasteValues
    Worksheets("Master EFT").Range("H2").Copy
    Worksheets("Master EFT").Range("H2").PasteSpecial Paste:=xlPasteValues
    
   
    Worksheets("Tool").Activate
    Worksheets("Tool").Cells(1, 1).Select
    
    MsgBox "Master EFT File has been optimized, blank rows been deleted, Error checking done!"
    
End Sub

