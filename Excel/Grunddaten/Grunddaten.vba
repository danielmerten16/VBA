Sub subHallo()
    
    Dim intI As Integer
    Dim IntWrite As Integer
    Dim intMinJahr As Integer
    Dim intMaxJahr As Integer
    Dim intMainJahr As Integer
    Dim RowIndex As Integer
    Dim intAktMonat As Integer
    
    ''Anfang des Zaehlers
    intI = 2
    IntWrite = 0
    intAktMonat = Month(Date) - 1

    
    ''Jahres ermitlung Min und Max Jahr
    Do
        intMainJahr = tabGrunddaten.Cells(intI, "A")
        If intMainJahr >= intMaxJahr Then
            intMaxJahr = tabGrunddaten.Cells(intI, "A")
        End If
        intI = intI + 1
    Loop While "" <> tabGrunddaten.Cells(intI, "A")
    
    For intJ = 2 To intI
        If tabGrunddaten.Cells(intJ, "A") = intMaxJahr Then
            tabGrunddaten.Cells(intI + IntWrite, "A") = tabGrunddaten.Cells(intJ, "A") + 1
            tabGrunddaten.Cells(intI + IntWrite, "B") = tabGrunddaten.Cells(intJ, "B")
            tabGrunddaten.Cells(intI + IntWrite, "C") = tabGrunddaten.Cells(intJ, "C")
            tabGrunddaten.Cells(intI + IntWrite, "D") = tabGrunddaten.Cells(intJ, "D")
            tabGrunddaten.Cells(intI + IntWrite, "E") = tabGrunddaten.Cells(intJ, "E")
            tabGrunddaten.Cells(intI + IntWrite, "F") = tabGrunddaten.Cells(intJ, "F")
            tabGrunddaten.Cells(intI + IntWrite, "G") = tabGrunddaten.Cells(intJ, "G") * 1.05
            IntWrite = IntWrite + 1
        End If
    Next intJ
    
End Sub


