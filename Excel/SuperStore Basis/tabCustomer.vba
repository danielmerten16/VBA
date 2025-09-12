Sub Customer()
    Dim intI As Integer
    Dim intJ As Integer
    Dim intZ As Integer
    Dim strZwischenSpeicher As String
    
    ''Anfang des Zaehlers
    intI = 1
    intJ = 2
    intZ = 0
    intLets = intJ
    
    Do Until tabCustomer.Cells(intI, "A") = "" And tabCustomer.Cells(intI, "G") = ""
        For intRem = 1 To 20
            tabCustomer.Cells(intI, intRem) = ""
        Next intRem
        intI = intI + 1
    Loop
    
    ''Titel Übersicht
    tabCustomer.Cells(1, "A") = "Customer ID"
    tabCustomer.Cells(1, "B") = "Customer Name"
    
    
    
    
    
    intI = 2
    
    Do Until Tabelle2.Cells(intI, "A") = ""
        If Tabelle2.Cells(intI, "G") <> tabCustomer.Cells(intJ - 1, "B") Then
            tabCustomer.Cells(intJ, "A") = Tabelle2.Cells(intI, "F")
            tabCustomer.Cells(intJ, "B") = Tabelle2.Cells(intI, "G")
            intJ = intJ + 1
        End If
        'If Tabelle2.Cells(intI, "G") = tabCustomer.Cells(intJ - 1, "B") Then
        '    If intJ <> 2 Then
        '        tabCustomer.Cells(intJ - 1, "C") = tabCustomer.Cells(intJ - 1, "C") + 1
        '        If Tabelle2.Cells(intI, "AA") <> "" Then
        '            tabCustomer.Cells(intJ - 1, "D") = tabCustomer.Cells(intJ - 1, "D") + 1
        '        Else
        '            tabCustomer.Cells(intJ - 1, "E") = tabCustomer.Cells(intJ - 1, "E") + Tabelle2.Cells(intI, "X")
        '        End If
        '        tabCustomer.Cells(intJ - 1, "F") = tabCustomer.Cells(intJ - 1, "F") + Tabelle2.Cells(intI, "E")
        '    End If
        'End If
        'If intJ > 4 Then
        '    If tabCustomer.Cells(intJ - 2, "A") < tabCustomer.Cells(intJ - 3, "A") And intZ = 0 Then
                'MsgBox intJ
                'intZ = 0
        '        For intMov = 1 To 6
        '             tabCustomer.Cells(intJ, intMov) = tabCustomer.Cells(intJ - 3, intMov)
        '             tabCustomer.Cells(intJ - 3, intMov) = tabCustomer.Cells(intJ - 2, intMov)
        '             tabCustomer.Cells(intJ - 2, intMov) = tabCustomer.Cells(intJ, intMov)
        '             tabCustomer.Cells(intJ, intMov) = ""
        '        Next intMov
        '    End If
        'End If
        intI = intI + 1
    Loop
    
    intI = 2
    intJ = 2
    
    tabCustomer.Cells(1, "C") = "Orders"
    Do Until Tabelle2.Cells(intI, "A") = ""
        If Tabelle2.Cells(intI, "G") <> tabCustomer.Cells(intJ - 1, "B") Then
            intJ = intJ + 1
        End If
        If Tabelle2.Cells(intI, "G") = tabCustomer.Cells(intJ - 1, "B") Then
            If intJ <> 2 Then
                tabCustomer.Cells(intJ - 1, "C") = tabCustomer.Cells(intJ - 1, "C") + 1
            End If
        End If
        intI = intI + 1
    Loop
    
    intI = 2
    intJ = 2

    tabCustomer.Cells(1, "D") = "Canceled"
    Do Until Tabelle2.Cells(intI, "A") = ""
        If Tabelle2.Cells(intI, "G") <> tabCustomer.Cells(intJ - 1, "B") Then
            intJ = intJ + 1
        End If
        If Tabelle2.Cells(intI, "G") = tabCustomer.Cells(intJ - 1, "B") Then
            If intJ <> 2 Then
                If Tabelle2.Cells(intI, "AA") <> "" Then
                    tabCustomer.Cells(intJ - 1, "D") = tabCustomer.Cells(intJ - 1, "D") + 1
                End If
            End If
        End If
        intI = intI + 1
    Loop
    
    intI = 2
    intJ = 2

    tabCustomer.Cells(1, "E") = "Sales volume"
    Do Until Tabelle2.Cells(intI, "A") = ""
        If Tabelle2.Cells(intI, "G") <> tabCustomer.Cells(intJ - 1, "B") Then
            intJ = intJ + 1
        End If
        If Tabelle2.Cells(intI, "G") = tabCustomer.Cells(intJ - 1, "B") Then
            If intJ <> 2 Then
                tabCustomer.Cells(intJ - 1, "C") = tabCustomer.Cells(intJ - 1, "C") + 1
                If Tabelle2.Cells(intI, "AA") = "" Then
                    tabCustomer.Cells(intJ - 1, "E") = tabCustomer.Cells(intJ - 1, "E") + Tabelle2.Cells(intI, "X")
                End If
            End If
        End If
        intI = intI + 1
    Loop
    
    intI = 2
    intJ = 2

    tabCustomer.Cells(1, "F") = "Postage"
    Do Until Tabelle2.Cells(intI, "A") = ""
        If Tabelle2.Cells(intI, "G") <> tabCustomer.Cells(intJ - 1, "B") Then
            intJ = intJ + 1
        End If
        If Tabelle2.Cells(intI, "G") = tabCustomer.Cells(intJ - 1, "B") Then
            If intJ <> 2 Then
                tabCustomer.Cells(intJ - 1, "F") = tabCustomer.Cells(intJ - 1, "F") + Tabelle2.Cells(intI, "E")
            End If
        End If
        intI = intI + 1
    Loop
    
    intI = 2
    intJ = 2

    Do Until Tabelle2.Cells(intI, "A") = ""
        If Tabelle2.Cells(intI, "G") <> tabCustomer.Cells(intJ - 1, "B") Then
            intJ = intJ + 1
        End If
        If intJ > 4 Then
            If tabCustomer.Cells(intJ - 2, "A") < tabCustomer.Cells(intJ - 3, "A") And intZ = 0 Then
                'MsgBox intJ
                'intZ = 0
                For intMov = 1 To 6
                     tabCustomer.Cells(intJ, intMov) = tabCustomer.Cells(intJ - 3, intMov)
                     tabCustomer.Cells(intJ - 3, intMov) = tabCustomer.Cells(intJ - 2, intMov)
                     tabCustomer.Cells(intJ - 2, intMov) = tabCustomer.Cells(intJ, intMov)
                     tabCustomer.Cells(intJ, intMov) = ""
                Next intMov
            End If
        End If
        intI = intI + 1
    Loop
    
    intI = 2
    intJ = 2
    
    tabCustomer.Cells(1, "G") = "Not uset Customer ID"
    intZ = 0
    Do Until tabCustomer.Cells(intJ, "A") = ""
        For intnotuse = intZ To tabCustomer.Cells(intJ, "A") - 1
                tabCustomer.Cells(intnotuse + 2, "G") = intnotuse
                intZ = intZ + 1
        Next intnotuse
        intJ = intJ + 1
    Loop
    
    intI = 2
    intJ = 2
    
    tabCustomer.Cells(1, "I") = "Product Category"
    intZ = 0
    Do Until Tabelle2.Cells(intI, "A") = ""
        For intA = 2 To intJ
            If tabCustomer.Cells(intA, "I") = Tabelle2.Cells(intI, "J") Then
                intZ = 1
                tabCustomer.Cells(intA, "J") = tabCustomer.Cells(intA, "J") + 1
            End If
        Next intA
        If intZ = 0 Then
            tabCustomer.Cells(intA - 1, "I") = Tabelle2.Cells(intI, "J")
            tabCustomer.Cells(intA - 1, "J") = tabCustomer.Cells(intA, "J") + 1
            intJ = intJ + 1
        End If
        intZ = 0
        intI = intI + 1
    Loop
    
    intI = 2
    intJ = 2
    intZ = 0
    
    tabCustomer.Cells(1, "L") = "Product Sub-Category"
    Do Until Tabelle2.Cells(intI, "A") = ""
        For intA = 2 To intJ
            If tabCustomer.Cells(intA, "L") = Tabelle2.Cells(intI, "K") Then
                tabCustomer.Cells(intA, "M") = tabCustomer.Cells(intA, "M") + 1
                intZ = 1
            End If
        Next intA
        If intZ = 0 Then
            tabCustomer.Cells(intA - 1, "L") = Tabelle2.Cells(intI, "K")
            tabCustomer.Cells(intA - 1, "M") = tabCustomer.Cells(intA, "M") + 1
            intJ = intJ + 1
        End If
        intZ = 0
        intI = intI + 1
    Loop
    
    intI = 2
    intJ = 2
    intZ = 0
    
    tabCustomer.Cells(1, "O") = "Product Container"
    Do Until Tabelle2.Cells(intI, "A") = ""
        For intA = 2 To intJ
            If tabCustomer.Cells(intA, "O") = Tabelle2.Cells(intI, "L") Then
                tabCustomer.Cells(intA, "P") = tabCustomer.Cells(intA, "P") + 1
                intZ = 1
            End If
        Next intA
        If intZ = 0 Then
            tabCustomer.Cells(intA - 1, "O") = Tabelle2.Cells(intI, "L")
            tabCustomer.Cells(intA - 1, "P") = tabCustomer.Cells(intA, "P") + 1
            intJ = intJ + 1
        End If
        intZ = 0
        intI = intI + 1
    Loop
    
End Sub

