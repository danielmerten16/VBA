Sub subHallo()
    
    Dim intI As Integer
    Dim intMinJahr As Integer
    Dim intMaxJahr As Integer
    Dim intMainJahr As Integer
    Dim RowIndex As Integer
    Dim intAktMonat As Integer
    
    ''Anfang des Zaehlers
    intI = 2
    
    intAktMonat = Month(Date) - 1
    
    ''Titel Übersicht
    tabLösung.Cells(1, "A") = "Jahre"
    tabLösung.Cells(1, "B") = "Lösung"
    tabLösung.Cells(1, "C") = "Prozentual"
    tabLösung.Cells(1, "D") = "Diferenz"
    tabLösung.Cells(1, "E") = "Monats Durchschnit"
    
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
    
    ''intMaxJahr = 2015
    
    ''Das Lehre Blat
    For intMainJahr = intMinJahr To intMaxJahr
        RowIndex = ((2 + (intMaxJahr - intMinJahr)) - (intMaxJahr - intMainJahr))
        tabLösung.Cells(RowIndex, "A") = ""
        tabLösung.Cells(RowIndex, "B") = ""
        tabLösung.Cells(RowIndex, "C") = ""
        tabLösung.Cells(RowIndex, "D") = ""
        tabLösung.Cells(RowIndex, "E") = ""
    Next intMainJahr
    

    For intJ = 2 To intI
        For intMainJahr = intMinJahr To intMaxJahr
            RowIndex = ((2 + (intMaxJahr - intMinJahr)) - (intMaxJahr - intMainJahr))
            If tabGrunddaten.Cells(intJ, "A") = intMainJahr Then
                If intMainJahr = intMaxJahr Then
                    If intAktMonat >= Month(DateValue("1 " & tabGrunddaten.Cells(intJ, "B") & " 2000")) And intAktMonat >= 1 Then
                        tabLösung.Cells(RowIndex, "B") = tabLösung.Cells(RowIndex, "B") + tabGrunddaten.Cells(intJ, "G")
                    End If
                Else
                    tabLösung.Cells(RowIndex, "B") = tabLösung.Cells(RowIndex, "B") + tabGrunddaten.Cells(intJ, "G")
                End If
            End If
        Next intMainJahr
    Next intJ
    
    If intAktMonat = 0 Then
        intMaxJahr = intMaxJahr - 1
    End If
    
    For intMainJahr = intMinJahr To intMaxJahr
        RowIndex = ((2 + (intMaxJahr - intMinJahr)) - (intMaxJahr - intMainJahr))
        tabLösung.Cells(RowIndex, "A") = intMainJahr
        If intMainJahr = intMaxJahr And intAktMonat >= 1 Then
            tabLösung.Cells(RowIndex, "E") = tabLösung.Cells(RowIndex, "B") / intAktMonat
        Else
            tabLösung.Cells(RowIndex, "E") = tabLösung.Cells(RowIndex, "B") / 12
        End If
    Next intMainJahr
    
    For intMainJahr = intMinJahr To intMaxJahr
        RowIndex = ((2 + (intMaxJahr - intMinJahr)) - (intMaxJahr - intMainJahr))
        If intMainJahr <> intMaxJahr Then
            tabLösung.Cells(RowIndex + 1, "D") = Round((((tabLösung.Cells(RowIndex + 1, "E") / tabLösung.Cells(RowIndex, "E")) - 1) * 100), 2)
        End If
        tabLösung.Cells(RowIndex, "C") = 100 + tabLösung.Cells(RowIndex, "D")
    Next intMainJahr

End Sub

