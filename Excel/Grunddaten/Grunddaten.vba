Sub subHallo()
    
    Dim IntI As Integer
    Dim IntWrite As Integer
    Dim intMinJahr As Integer
    Dim intMaxJahr As Integer
    Dim intMainJahr As Integer
    Dim RowIndex As Integer
    Dim intAktMonat As Integer
    
    ''Anfang des Zaehlers
    IntI = 2
    IntWrite = 0
    intAktMonat = Month(Date) - 1

    
    ''Jahres ermitlung Min und Max Jahr
    Do
        intMainJahr = tabGrunddaten.Cells(IntI, "A")
        If intMainJahr >= intMaxJahr Then
            intMaxJahr = tabGrunddaten.Cells(IntI, "A")
        End If
        IntI = IntI + 1
    Loop While "" <> tabGrunddaten.Cells(IntI, "A")
    
    For IntJ = 2 To IntI
        If tabGrunddaten.Cells(IntJ, "A") = intMaxJahr Then
            If intMainJahr = intMaxJahr Then
                tabGrunddaten.Cells(IntI + IntWrite, "A") = tabGrunddaten.Cells(IntJ, "A") + 1
                tabGrunddaten.Cells(IntI + IntWrite, "B") = tabGrunddaten.Cells(IntJ, "B")
                tabGrunddaten.Cells(IntI + IntWrite, "C") = tabGrunddaten.Cells(IntJ, "C")
                tabGrunddaten.Cells(IntI + IntWrite, "D") = tabGrunddaten.Cells(IntJ, "D")
                tabGrunddaten.Cells(IntI + IntWrite, "E") = tabGrunddaten.Cells(IntJ, "E")
                tabGrunddaten.Cells(IntI + IntWrite, "F") = tabGrunddaten.Cells(IntJ, "F")
                tabGrunddaten.Cells(IntI + IntWrite, "G") = tabGrunddaten.Cells(IntJ, "G") * 1.05
                IntWrite = IntWrite + 1
            End If
        End If
    Next IntJ
    
End Sub

