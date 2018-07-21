Sub EXISTING_CONTRACTS_VOLUME()

Application.ScreenUpdating = False

[A:A, B:B, D:D, E:E, G:G, H:H, I:I, N:N, R:R, T:T, W:W, X:X, Z:Z, AC:AC, AD:AD].Delete

    [A:O].AutoFilter Field:=5, Criteria1:="*existing*"
    [A:O].AutoFilter Field:=10, Criteria1:=Array("NEW", "PAID"), Operator:=xlFilterValues
    [A:O].AutoFilter Field:=13, Criteria1:="SUCCESS"
    [A:O].AutoFilter Field:=15, Criteria1:=Array("DELIVERED", "NEW"), Operator:=xlFilterValues

Range("E2:E" & Cells(Rows.Count, "B").End(xlUp).Row).Select
Selection.Copy
Sheets.Add
ActiveSheet.Name = "Top_Offers"
ActiveSheet.Paste
Application.CutCopyMode = False

CurrentRowA = 1
LastRowA = Range("A50000").End(xlUp).Row
Dim r As Range
While CurrentRowA <= LastRowA
    CurrentRowB = 1
    LastRowB = Range("B50000").End(xlUp).Row
    Do While CurrentRowB <= LastRowB
        If Cells(CurrentRowA, "A").Value = Cells(CurrentRowB, "B").Value Then
            Exit Do
        Else
        CurrentRowB = CurrentRowB + 1
        End If
    Loop
    If CurrentRowB > LastRowB Then
        Cells(CurrentRowB, "B").Value = Cells(CurrentRowA, "A").Value
        Set r = Range("A1", "A" & LastRowA)
        Cells(CurrentRowB, "C").Value = Application.CountIf(r, Cells(CurrentRowA, "A").Value)
    End If
    CurrentRowA = CurrentRowA + 1
Wend
LastRowB = Range("B50000").End(xlUp).Row
Range("B2", "C" & LastRowB).Cut
Range("B1").Select
ActiveSheet.Paste
Application.CutCopyMode = False

Range("B:C").Sort Key1:=Range("C1"), Header:=xlNo, Order1:=xlDescending
[A:A].Delete
Range("A1").Select

Rows("1:1").Select
Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove

Range("A1") = "OFFER_NAME"
[A1].Font.Bold = True

Range("B1") = "Volume"
[B1].Font.Bold = True

Sheets("Consumption_Report").Select

Range("A2:A" & Cells(Rows.Count, "B").End(xlUp).Row).Select
Selection.Copy
Sheets.Add
ActiveSheet.Name = "Top_Companies"
ActiveSheet.Paste
Application.CutCopyMode = False

CurrentRowA = 1
LastRowA = Range("A50000").End(xlUp).Row
'Dim r As Range
While CurrentRowA <= LastRowA
    CurrentRowB = 1
    LastRowB = Range("B50000").End(xlUp).Row
    Do While CurrentRowB <= LastRowB
        If Cells(CurrentRowA, "A").Value = Cells(CurrentRowB, "B").Value Then
            Exit Do
        Else
        CurrentRowB = CurrentRowB + 1
        End If
    Loop
    If CurrentRowB > LastRowB Then
        Cells(CurrentRowB, "B").Value = Cells(CurrentRowA, "A").Value
        Set r = Range("A1", "A" & LastRowA)
        Cells(CurrentRowB, "C").Value = Application.CountIf(r, Cells(CurrentRowA, "A").Value)
    End If
    CurrentRowA = CurrentRowA + 1
Wend
LastRowB = Range("B50000").End(xlUp).Row
Range("B2", "C" & LastRowB).Cut
Range("B1").Select
ActiveSheet.Paste

Range("B:C").Sort Key1:=Range("C1"), Header:=xlNo, Order1:=xlDescending
[A:A].Delete
Range("A1").Select

Rows("1:1").Select
Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove

Range("A1") = "COMPANY_NAME"
[A1].Font.Bold = True

Range("B1") = "Volume"
[B1].Font.Bold = True

Sheets("Consumption_Report").Select

Range("D2:D" & Cells(Rows.Count, "B").End(xlUp).Row).Select
Selection.Copy
Sheets.Add
ActiveSheet.Name = "Top_Boards"
ActiveSheet.Paste
Application.CutCopyMode = False

CurrentRowA = 1
LastRowA = Range("A50000").End(xlUp).Row
'Dim r As Range
While CurrentRowA <= LastRowA
    CurrentRowB = 1
    LastRowB = Range("B50000").End(xlUp).Row
    Do While CurrentRowB <= LastRowB
        If Cells(CurrentRowA, "A").Value = Cells(CurrentRowB, "B").Value Then
            Exit Do
        Else
        CurrentRowB = CurrentRowB + 1
        End If
    Loop
    If CurrentRowB > LastRowB Then
        Cells(CurrentRowB, "B").Value = Cells(CurrentRowA, "A").Value
        Set r = Range("A1", "A" & LastRowA)
        Cells(CurrentRowB, "C").Value = Application.CountIf(r, Cells(CurrentRowA, "A").Value)
    End If
    CurrentRowA = CurrentRowA + 1
Wend
LastRowB = Range("B50000").End(xlUp).Row
Range("B2", "C" & LastRowB).Cut
Range("B1").Select
ActiveSheet.Paste
Application.CutCopyMode = False

Range("B:C").Sort Key1:=Range("C1"), Header:=xlNo, Order1:=xlDescending
[A:A].Delete
Range("A1").Select

Rows("1:1").Select
Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove

Range("A1") = "PARTNER_NAME"
[A1].Font.Bold = True

Range("B1") = "Volume"
[B1].Font.Bold = True

Sheets("Consumption_Report").Select

Range("I2:I" & Cells(Rows.Count, "B").End(xlUp).Row).Select
Selection.Copy
Sheets.Add
ActiveSheet.Name = "Top_Existing_Payment"
ActiveSheet.Paste
Application.CutCopyMode = False

CurrentRowA = 1
LastRowA = Range("A50000").End(xlUp).Row
'Dim r As Range
While CurrentRowA <= LastRowA
    CurrentRowB = 1
    LastRowB = Range("B50000").End(xlUp).Row
    Do While CurrentRowB <= LastRowB
        If Cells(CurrentRowA, "A").Value = Cells(CurrentRowB, "B").Value Then
            Exit Do
        Else
        CurrentRowB = CurrentRowB + 1
        End If
    Loop
    If CurrentRowB > LastRowB Then
        Cells(CurrentRowB, "B").Value = Cells(CurrentRowA, "A").Value
        Set r = Range("A1", "A" & LastRowA)
        Cells(CurrentRowB, "C").Value = Application.CountIf(r, Cells(CurrentRowA, "A").Value)
    End If
    CurrentRowA = CurrentRowA + 1
Wend
LastRowB = Range("B50000").End(xlUp).Row
Range("B2", "C" & LastRowB).Cut
Range("B1").Select
ActiveSheet.Paste
Application.CutCopyMode = False

Range("B:C").Sort Key1:=Range("C1"), Header:=xlNo, Order1:=xlDescending
[A:A].Delete
Range("A1").Select

Rows("1:1").Select
Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove

Range("A1") = "PAYMENT_METHOD"
[A1].Font.Bold = True

Range("B1") = "Volume"
[B1].Font.Bold = True

Sheets("Consumption_Report").Select
Range("A1").Select
Selection.AutoFilter

Sheets.Add
ActiveSheet.Name = "Existing_Contracts_Volume"

Sheets("Top_Offers").Select
Range("A1:B" & Cells(Rows.Count, "B").End(xlUp).Row).Select
Selection.Copy
Sheets("Existing_Contracts_Volume").Select
Range("A1").Select
ActiveSheet.Paste
Application.CutCopyMode = False

Sheets("Top_Companies").Select
Range("A1:B" & Cells(Rows.Count, "B").End(xlUp).Row).Select
Selection.Copy
Sheets("Existing_Contracts_Volume").Select
Range("D1").Select
ActiveSheet.Paste
Application.CutCopyMode = False

Sheets("Top_Boards").Select
Range("A1:B" & Cells(Rows.Count, "B").End(xlUp).Row).Select
Selection.Copy
Sheets("Existing_Contracts_Volume").Select
Range("G1").Select
ActiveSheet.Paste
Application.CutCopyMode = False

Sheets("Top_Existing_Payment").Select
Range("A1:B" & Cells(Rows.Count, "B").End(xlUp).Row).Select
Selection.Copy
Sheets("Existing_Contracts_Volume").Select
Range("J1").Select
ActiveSheet.Paste
Application.CutCopyMode = False

Application.DisplayAlerts = False
Sheets("Top_Offers").Delete
Sheets("Top_Companies").Delete
Sheets("Top_Boards").Delete
Sheets("Top_Existing_Payment").Delete
Application.DisplayAlerts = True

[C:C, F:F, I:I, L:L, M:M, N:N, O:O, P:P, Q:Q, R:R, S:S, T:T, U:U, V:V, W:W, X:X, Y:Y, Z:Z, AA:AA, AB:AB, AC:AC, AD:AD, AE:AE, AF:AF, AG:AG, AH:AH, AI:AI, AJ:AJ, AK:AK, AL:AL, AM:AM, AN:AN, AO:AO].Interior.Color = RGB(232, 232, 232)

Range("A1:B1").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous: .Weight = xlMedium: .ColorIndex = xlAutomatic
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous: .Weight = xlMedium: .ColorIndex = xlAutomatic
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous: .Weight = xlMedium: .ColorIndex = xlAutomatic
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous: .Weight = xlMedium: .ColorIndex = xlAutomatic
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous: .Weight = xlThin: .ColorIndex = 1
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous: .Weight = xlThin: .ColorIndex = 1
    End With

Range("D1:E1").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous: .Weight = xlMedium: .ColorIndex = xlAutomatic
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous: .Weight = xlMedium: .ColorIndex = xlAutomatic
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous: .Weight = xlMedium: .ColorIndex = xlAutomatic
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous: .Weight = xlMedium: .ColorIndex = xlAutomatic
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous: .Weight = xlThin: .ColorIndex = 1
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous: .Weight = xlThin: .ColorIndex = 1
    End With

Range("G1:H1").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous: .Weight = xlMedium: .ColorIndex = xlAutomatic
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous: .Weight = xlMedium: .ColorIndex = xlAutomatic
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous: .Weight = xlMedium: .ColorIndex = xlAutomatic
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous: .Weight = xlMedium: .ColorIndex = xlAutomatic
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous: .Weight = xlThin: .ColorIndex = 1
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous: .Weight = xlThin: .ColorIndex = 1
    End With

Range("J1:K1").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous: .Weight = xlMedium: .ColorIndex = xlAutomatic
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous: .Weight = xlMedium: .ColorIndex = xlAutomatic
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous: .Weight = xlMedium: .ColorIndex = xlAutomatic
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous: .Weight = xlMedium: .ColorIndex = xlAutomatic
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous: .Weight = xlThin: .ColorIndex = 1
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous: .Weight = xlThin: .ColorIndex = 1
    End With

Columns("A:K").EntireColumn.AutoFit
Columns("C:C").ColumnWidth = 3
Columns("F:F").ColumnWidth = 3
Columns("I:I").ColumnWidth = 3

[1:1].Font.Bold = True

Range("A1").Select

Application.ScreenUpdating = True

MsgBox "YOLO!"

End Sub