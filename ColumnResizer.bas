Attribute VB_Name = "ColumnResizer"
Option Explicit

Sub ResizeColumnsInSelection()
    
    ResizeColumnsInRange Selection
    
End Sub

Sub ResizeColumnsInActiveWorksheet()
    
    ResizeColumnsInWorksheet ActiveWorksheet
    
End Sub

Sub ResizeColumnsInActiveWorkbook()
    
    ResizeColumnsInWorkbook ActiveWorkbook
    
End Sub

Sub ResizeColumnsInWorkbook(wb As Workbook)
    
    Dim ws As Worksheet
    
    For Each ws In wb.Sheets
        
        ResizeColumnsInWorksheet ws
    
    Next ws
    
End Sub

Sub ResizeColumnsInWorksheet(ws As Worksheet)
    
    ResizeColumnsInRange ws.UsedRange
    
End Sub

Sub ResizeColumnsInRange(tr As Range)
    
    Dim col As Range
    
    Set tr = Intersect(tr, tr.Parent.UsedRange)
    
    For Each col In tr.Columns
        
        ResizeColumn col
        
    Next col
    
End Sub


Sub ResizeColumn(col As Range)

    Dim c As Range
    Dim f As Integer
    Dim l As Integer
    Dim ratio As Double
    Dim newSize As Integer
    Dim filterButtonSize As Integer
    Dim baseSize As Integer
    
    Set c = col.Cells(1)
    
    l = Len(c.Value2)
    f = c.Font.Size
    filterButtonSize = 3
    newSize = l + filterButtonSize
    
    col.ColumnWidth = newSize
    
End Sub
