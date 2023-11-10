Sub HighlightTableCells()
    Dim tbl As Table
    Dim cell As cell
    Dim value As Double
    Dim isValueValid As Boolean
    
    validValues = Array(3500, 750, 2350, 900, 2600, 1800, 2100)
             
                
    Dim x As Long, y As Variant
    
    With ActiveDocument.Tables(1).Range
      For x = 1 To .Cells.Count
        With .Cells(x)
          y = Split(.Range.text, vbCr)(0)
          If IsNumeric(y) Then
            isValueValid = False
            Select Case y
            Case 3500, 750, 2350, 900, 2600, 1800, 2100, 0
                isValueValid = True
            End Select
            If Not isValueValid Then
                .Shading.BackgroundPatternColor = wdColorBlue
            End If
          End If
        End With
      Next
    End With

End Sub
