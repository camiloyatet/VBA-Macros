Sub ResizeCharts() 
'cambiar pra que sea dinamico el libro de entrada, y las dimensiones del grafico.
    Workbooks("Workbook.xlsx").activate
    Dim ws As Worksheet
    
    For Each ws In ActiveWorkbook.Worksheets
    ws.activate
        Dim chart As ChartObject
        For Each chart In ActiveSheet.ChartObjects
            With chart.Parent
                chart.Height = Application.InchesToPoints(3.9)
                chart.Width = Application.InchesToPoints(5.9)
            End With
        Next
    Next
End Sub
