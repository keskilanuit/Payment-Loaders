Attribute VB_Name = "Module2"
Option Explicit
Public Sub CombineDataFromAllSheets()

    Dim wksSrc As Worksheet, wksDst As Worksheet
    Dim rngSrc As Range, rngDst As Range
    Dim lngLastCol As Long, lngSrcLastRow As Long, lngDstLastRow As Long
    
   
    Set wksDst = ThisWorkbook.Worksheets("Master EFT")
    lngDstLastRow = LastOccupiedRowNum(wksDst)
    lngLastCol = LastOccupiedColNum(wksDst)
    
 
    Set rngDst = wksDst.Cells(lngDstLastRow + 1, 1)
    
  
    For Each wksSrc In ThisWorkbook.Worksheets
    
       
        If wksSrc.Name <> "Master EFT" Then
            
           
            lngSrcLastRow = LastOccupiedRowNum(wksSrc)
            
          
            With wksSrc
                Set rngSrc = .Range(.Cells(4, 1), .Cells(lngSrcLastRow, lngLastCol))
                rngSrc.Copy Destination:=rngDst
            End With
            
           
            lngDstLastRow = LastOccupiedRowNum(wksDst)
            Set rngDst = wksDst.Cells(lngDstLastRow + 1, 1)
            
        End If
    
    Next wksSrc

MsgBox ("All Sub EFT tabs has been combined into the Master EFT Sheet!")

End Sub


Public Function LastOccupiedRowNum(Sheet As Worksheet) As Long
    Dim lng As Long
    If Application.WorksheetFunction.CountA(Sheet.Cells) <> 0 Then
        With Sheet
            lng = .Cells.Find(What:="*", _
                              After:=.Range("A1"), _
                              Lookat:=xlPart, _
                              LookIn:=xlFormulas, _
                              SearchOrder:=xlByRows, _
                              SearchDirection:=xlPrevious, _
                              MatchCase:=False).Row
        End With
    Else
        lng = 1
    End If
    LastOccupiedRowNum = lng
End Function


Public Function LastOccupiedColNum(Sheet As Worksheet) As Long
    Dim lng As Long
    If Application.WorksheetFunction.CountA(Sheet.Cells) <> 0 Then
        With Sheet
            lng = .Cells.Find(What:="*", _
                              After:=.Range("A1"), _
                              Lookat:=xlPart, _
                              LookIn:=xlFormulas, _
                              SearchOrder:=xlByColumns, _
                              SearchDirection:=xlPrevious, _
                              MatchCase:=False).Column
        End With
    Else
        lng = 1
    End If
    LastOccupiedColNum = lng

    
End Function

