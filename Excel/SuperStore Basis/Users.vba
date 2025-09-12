Sub Users()
    Dim intI As Integer
    Dim intJ As Integer
    Dim strZwischenSpeicher As String
    
    ''Anfang des Zaehlers
    intI = 1
    
    
    
    Do Until Sheet3.Cells(intI, "A") = ""
        For intRem = 4 To 5
            Sheet3.Cells(intI, intRem) = ""
        Next intRem
        intI = intI + 1
    Loop
    
    ''Titel Übersicht
    Sheet3.Cells(1, "D") = "Sales volume"
    Sheet3.Cells(1, "E") = "Packages"
    
    intI = 2
    Do Until Sheet3.Cells(intI, "A") = ""
        intJ = 2
        Do Until Tabelle2.Cells(intJ, "A") = ""
            If Sheet3.Cells(intI, "B") = Tabelle2.Cells(intJ, "Z") Then
                Sheet3.Cells(intI, "D") = Sheet3.Cells(intI, "D") + Tabelle2.Cells(intJ, "X")
            End If
            If Sheet3.Cells(intI, "B") = Tabelle2.Cells(intJ, "Z") Then
                Sheet3.Cells(intI, "E") = Sheet3.Cells(intI, "E") + Tabelle2.Cells(intJ, "W")
            End If
            intJ = intJ + 1
        Loop
        intI = intI + 1
    Loop
    
End Sub

