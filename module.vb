Sub insertpagebreaks()
'updateby Extendoffice 20151228
    Dim I As Long, J As Long
    J = ActiveSheet.Cells(Rows.Count, "D").End(xlUp).Row
    For I = J To 2 Step -1
        If Range("D" & I).Value <> Range("D" & I - 1).Value Then
            ActiveSheet.HPageBreaks.Add Before:=Range("D" & I)
        End If
    Next I
End Sub
