Sub Orders()
    Dim intI As Integer
    Dim intA As Integer
    
    ''Anfang des Zaehlers
    intI = 2
    intA = 2
    
    ''Titel Übersicht
    Tabelle2.Cells(1, "Z") = "Manager"
    Tabelle2.Cells(1, "AA") = "Status"
    
    ''Jahres ermitlung Min und Max Jahr
    Do Until "" = Tabelle2.Cells(intI, "A")
        Tabelle2.Cells(intI, "Z") = ""
        Tabelle2.Cells(intI, "AA") = ""
        Do Until "" = Sheet3.Cells(intA, "A")
            If Tabelle2.Cells(intI, "P") = Sheet3.Cells(intA, "A") Then
                Tabelle2.Cells(intI, "Z") = Sheet3.Cells(intA, "B")
            End If
            intA = intA + 1
        Loop
        intA = 2
        Do Until "" = Tabelle3.Cells(intA, "A")
            If Tabelle2.Cells(intI, "Y") = Tabelle3.Cells(intA, "A") Then
                Tabelle2.Cells(intI, "AA") = Tabelle3.Cells(intA, "B")
            Else
                Tabelle2.Cells(intI, "AA") = "Not Returned"
            End If
            intA = intA + 1
        Loop
        intI = intI + 1
        intA = 2
    Loop
    
End Sub