Sub CalculPERT()
    Dim ws As Worksheet
    Set ws = Worksheets("PERT")
    
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    Dim i As Long
    For i = 2 To lastRow
        Dim O As Double, M As Double, P As Double
        O = ws.Cells(i, 2).Value
        M = ws.Cells(i, 3).Value
        P = ws.Cells(i, 4).Value
        
        If IsNumeric(O) And IsNumeric(M) And IsNumeric(P) Then
            ws.Cells(i, 5).Value = Round((O + 4 * M + P) / 6, 2)
            ws.Cells(i, 6).Value = Round(((P - O) / 6) ^ 2, 2)
        End If
    Next i
    
    MsgBox "✅ Calculs terminés avec succès !", vbInformation
End Sub
