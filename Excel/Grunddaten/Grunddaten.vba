Sub subHallo()
    
    Dim intI As Integer
    Dim intWrite As Integer
    Dim intMinJahr As Integer
    Dim intMaxJahr As Integer
    Dim intMainJahr As Integer
    Dim RowIndex As Integer
    Dim intAktMonat As Integer
    
    ''Anfang des Zaehlers
    intI = 2
    intWrite = 0
    intAktMonat = Month(Date) - 1

    
    ''Jahres ermitlung Min und Max Jahr
    Do
        intMainJahr = tabGrunddaten.Cells(intI, "A")
        If intMainJahr <= intMinJahr Or intMinJahr = 0 Then
            intMinJahr = tabGrunddaten.Cells(intI, "A")
        End If
        If intMainJahr >= intMaxJahr Then
            intMaxJahr = tabGrunddaten.Cells(intI, "A")
        End If
        intI = intI + 1
    Loop While "" <> tabGrunddaten.Cells(intI, "A")
    
    For intJ = 2 To intI
        For intMainJahr = intMinJahr To intMaxJahr
            If tabGrunddaten.Cells(intJ, "A") = intMainJahr Then
                If intMainJahr = intMaxJahr Then
                    tabGrunddaten.Cells(intI + intWrite, "A") = tabGrunddaten.Cells(intJ, "A") + 1
                    tabGrunddaten.Cells(intI + intWrite, "B") = tabGrunddaten.Cells(intJ, "B")
                    tabGrunddaten.Cells(intI + intWrite, "C") = tabGrunddaten.Cells(intJ, "C")
                    tabGrunddaten.Cells(intI + intWrite, "D") = tabGrunddaten.Cells(intJ, "D")
                    tabGrunddaten.Cells(intI + intWrite, "E") = tabGrunddaten.Cells(intJ, "E")
                    tabGrunddaten.Cells(intI + intWrite, "F") = tabGrunddaten.Cells(intJ, "F")
                    tabGrunddaten.Cells(intI + intWrite, "G") = tabGrunddaten.Cells(intJ, "G") * 1.05
                    intWrite = intWrite + 1
                End If
            End If
        Next intMainJahr
    Next intJ
    
End Sub
