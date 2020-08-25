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
    Dim cellValueLength As Integer
    Dim fontSize As Integer
    Dim newSize As Integer
    Dim filterButtonSize As Integer
    Dim fontSizeToLengthRatio As Double
    
    Set c = col.Cells(1)
    
    cellValueLength = Len(c.Value2)
    fontSize = c.Font.Size
    filterButtonSize = 5
    fontSizeToLengthRatio = 0.95
    
    newSize = fontSizeToLengthRatio * cellValueLength + filterButtonSize
    
    If newSize > 50 Then newSize = 50
    If newSize < 8 Then newSize = 8
    col.ColumnWidth = newSize
    
End Sub
