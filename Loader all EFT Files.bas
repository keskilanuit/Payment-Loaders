Sub Load_All_Subpayment_Files()


Dim Path As String
Dim fileName As String
Dim Sheet As Worksheet

    Path = "your path\"
    'load all .xlsx files as long as it has a name.'
fileName = Dir(Path & "*.xlsx")
Do While fileName <> ""
    
Workbooks.Open fileName:=Path & fileName, ReadOnly:=True

    For Each Sheet In ActiveWorkbook.Sheets
        Sheet.Copy After:=ThisWorkbook.Sheets(1)
         Application.DisplayAlerts = False
    Next Sheet

Workbooks(fileName).Close

fileName = Dir()
Loop

Worksheets("Master EFT").Activate
Worksheets("Master EFT").Cells(1, 1).Select
Worksheets("Tool").Activate
Worksheets("Tool").Cells(1, 1).Select
        'indicates how many files been loaded, -2 to get accurate number (since template contains two ws)'
        MsgBox "Consolidate Payment sheet Tool is done loading, total of " & Application.Sheets.Count - 2 & " files has been loaded."

End Sub

