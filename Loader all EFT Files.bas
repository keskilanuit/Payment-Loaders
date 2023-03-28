Attribute VB_Name = "Module1"
Sub Load_All_EFT_Files()


Dim Path As String
Dim fileName As String
Dim Sheet As Worksheet

Path = "K:\Dept\Finance\Tax\SALT\S & U\Return Processing\_Return Processing post implementation\_Master Loader File\_2023 EFT LOADER\EFT Files\"
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

MsgBox "Consolidate EFT Loader Tool is done loading, total of " & Application.Sheets.Count - 2 & " files has been loaded."

End Sub

