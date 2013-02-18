Sub Macro1()
    Range("C3,E3,I3,K3,B5,D5,F5,H5,J5,L5,I16:I40,B41,D42,H42,I45:I48,B49,D50,H50,J53:J64,B65,D66,H66,I69:I74,B78,D79,H79,J82:J89,B92,D93,H93,J96:J100,B101:B102,A105,A107,C111,B114,D113,F117:F122,D123,B124,F127:F132,F134,I134,F137:F142,F144,I144,F147:F152").Select
    With Selection.Font
        .Name = "標楷體"
        .Size = 12
        .ColorIndex = xlAutomatic
    End With
    Selection.Font.Bold = False
    With Selection
        .HorizontalAlignment = xlLeft
    End With
End Sub
