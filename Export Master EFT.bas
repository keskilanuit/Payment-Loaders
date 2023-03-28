Sub Export_Master_payment_sheet()

Dim ws As Worksheet
Dim RangeName As Name


On Error Resume Next
For Each RangeName In Names
    ActiveWorkbook.Names(RangeName.Name).Delete
    Next
On Error GoTo 0

For Each ws In ThisWorkbook.Worksheets
Application.DisplayAlerts = False

If ws.Name <> "Master EFT" And ws.Name <> "Tool" Then
ws.Delete
End If

Next ws
Application.DisplayAlerts = True


Worksheets("Master EFT").Activate
Worksheets("Master EFT").Cells(1, 1).Select


Dim varResult As Variant
Dim dirPath, fileName As String
 
 
dirPath = Application.ActiveWorkbook.Path
fileName = "_Master EFT Loader " & Format(Now, "MM.DD.YY")
 
 
Application.ScreenUpdating = False
Application.DisplayAlerts = False
 
With ActiveSheet
    ActiveSheet.Copy
    Application.ActiveWorkbook.SaveAs fileName:=dirPath & "\" & fileName & ".xlsx"
    Application.ActiveWorkbook.Close False

End With
 
 
 
Application.DisplayAlerts = True
Application.ScreenUpdating = True
 
Range("A4:Z5000").Delete

    
    
    Worksheets("Master EFT").Range("B2").Formula = "=TODAY()-30"
    Worksheets("Master EFT").Range("H2").Formula = "=TODAY()"

Worksheets("Tool").Activate
Worksheets("Tool").Cells(1, 1).Select


MsgBox "Master EFT File has been exported, all sub-EFT worksheets are been deleted from this Workbook!"

End Sub

