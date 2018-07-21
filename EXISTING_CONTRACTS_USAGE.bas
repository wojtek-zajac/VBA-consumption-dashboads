Sub EXISTING_CONTRACTS_USAGE()

Application.ScreenUpdating = False

Range("A:AD").Sort Key1:=Range("C1"), Header:=xlYes

[A:A, B:B, D:D, E:E, F:F, G:G, H:H, I:I, J:J, N:N, Q:Q, R:R, S:S, T:T, W:W, X:X, Z:Z, AC:AC, AD:AD].Delete

[A1:O1].Font.Bold = True

    [A:K].AutoFilter Field:=2, Criteria1:="<>*Linkedin*"
    [A:K].AutoFilter Field:=3, Criteria1:="*existing*"
    [A:K].AutoFilter Field:=9, Criteria1:="SUCCESS"
    [A:K].AutoFilter Field:=11, Criteria1:=Array("DELIVERED", "NEW"), Operator:=xlFilterValues

Range("A1:K" & Cells(Rows.Count, "A").End(xlUp).Row).Select
Selection.Copy
Sheets.Add
ActiveSheet.Name = "Existing_Contracts"
ActiveSheet.Paste
Application.CutCopyMode = False

Application.DisplayAlerts = False
Sheets("Consumption_Report").Delete
Application.DisplayAlerts = True

Application.ScreenUpdating = True

'_________________________
    Dim My_Range As Range
    Dim FieldNum As Long
    Dim CalcMode As Long
    Dim ViewMode As Long
    Dim ws2 As Worksheet
    Dim Lrow As Long
    Dim cell As Range
    Dim CCount As Long
    Dim WSNew As Worksheet
    Dim ErrNum As Long
    Dim DestRange As Range
    Dim Lr As Long

    Set My_Range = Range("A1:K" & LastRow(ActiveSheet))
    My_Range.Parent.Select

        If ActiveWorkbook.ProtectStructure = True Or _
           My_Range.Parent.ProtectContents = True Then
           MsgBox "Sorry, not working when the workbook or worksheet is protected", vbOKOnly, "Copy to new worksheet"
           Exit Sub
       End If

    FieldNum = 1

    My_Range.Parent.AutoFilterMode = False

    With Application
        CalcMode = .Calculation
        .Calculation = xlCalculationManual
        .ScreenUpdating = False
        .EnableEvents = False
    End With
    
    ViewMode = ActiveWindow.View
    ActiveWindow.View = xlNormalView
    ActiveSheet.DisplayPageBreaks = False

    Set ws2 = Worksheets.Add

    With ws2

    My_Range.Columns(FieldNum).AdvancedFilter Action:=xlFilterCopy, CopyToRange:=.Range("A1"), UNIQUE:=True

    Lrow = .Cells(Rows.Count, "A").End(xlUp).Row
    For Each cell In .Range("A2:A" & Lrow)

    My_Range.Parent.Select

    My_Range.AutoFilter Field:=FieldNum, Criteria1:="=" & Replace(Replace(Replace(cell.Value, "~", "~~"), "*", "~*"), "?", "~?")

            CCount = 0
            On Error Resume Next
            CCount = My_Range.Columns(1).SpecialCells(xlCellTypeVisible).Areas(1).Cells.Count
            On Error GoTo 0
            If CCount = 0 Then
                MsgBox "There are more than 8192 areas for the value: " & cell.Value & vbNewLine & "It is not possible to copy the visible data." & vbNewLine & "Tip: Sort your data before you use this macro.", vbOKOnly, "Split in worksheets"
            
       Else
            If SheetExists(cell.Text) = False Then
                    Set WSNew = Worksheets.Add(After:=Sheets(Sheets.Count))
                    On Error Resume Next
                    WSNew.Name = cell.Value
                    
            If Err.Number > 0 Then
                        ErrNum = ErrNum + 1
                        WSNew.Name = "Error_" & Format(ErrNum, "0000")
                        Err.Clear
            End If
            
                    On Error GoTo 0
                    Set DestRange = WSNew.Range("A1")
             
        Else
                    Set WSNew = Sheets(cell.Text)
                    Lr = LastRow(WSNew)
                    Set DestRange = WSNew.Range("A" & Lr + 1)
            End If

                My_Range.SpecialCells(xlCellTypeVisible).Copy
                With DestRange
                    .Parent.Select
                    .PasteSpecial Paste:=8
                    .PasteSpecial xlPasteValues
                    .PasteSpecial xlPasteFormats
                    Application.CutCopyMode = False
                    .Select
                End With
            End If

'____________________________________________
For Each Worksheet In ThisWorkbook.Worksheets
     
Range("A:K").Sort Key1:=Range("B1"), Header:=xlYes

ActiveSheet.Range("B:B").AdvancedFilter Action:=xlFilterCopy, CopyToRange:=ActiveSheet.Range("N1"), UNIQUE:=True

Range("O1") = "Volume"
[O1].Font.Bold = True

       Range("O2").Select
   ActiveCell.Formula = "=COUNTIF(B:B,N2)"
   If Selection = 0 Then Selection.ClearContents

        Range("O3").Select
    ActiveCell.Formula = "=COUNTIF(B:B,N3)"
    If Selection = 0 Then Selection.ClearContents
    
        Range("O4").Select
    ActiveCell.Formula = "=COUNTIF(B:B,N4)"
    If Selection = 0 Then Selection.ClearContents

       Range("O5").Select
    ActiveCell.Formula = "=COUNTIF(B:B,N5)"
    If Selection = 0 Then Selection.ClearContents

       Range("O6").Select
    ActiveCell.Formula = "=COUNTIF(B:B,N6)"
    If Selection = 0 Then Selection.ClearContents

       Range("O7").Select
    ActiveCell.Formula = "=COUNTIF(B:B,N7)"
    If Selection = 0 Then Selection.ClearContents

       Range("O8").Select
    ActiveCell.Formula = "=COUNTIF(B:B,N8)"
    If Selection = 0 Then Selection.ClearContents
     
       Range("O9").Select
    ActiveCell.Formula = "=COUNTIF(B:B,N9)"
    If Selection = 0 Then Selection.ClearContents
     
       Range("O10").Select
    ActiveCell.Formula = "=COUNTIF(B:B,N10)"
    If Selection = 0 Then Selection.ClearContents
    
       Range("O11").Select
    ActiveCell.Formula = "=COUNTIF(B:B,N11)"
    If Selection = 0 Then Selection.ClearContents
     
       Range("O12").Select
    ActiveCell.Formula = "=COUNTIF(B:B,N12)"
    If Selection = 0 Then Selection.ClearContents
     
       Range("O13").Select
    ActiveCell.Formula = "=COUNTIF(B:B,N13)"
    If Selection = 0 Then Selection.ClearContents
     
      Range("O14").Select
    ActiveCell.Formula = "=COUNTIF(B:B,N14)"
    If Selection = 0 Then Selection.ClearContents
     
        Range("O15").Select
    ActiveCell.Formula = "=COUNTIF(B:B,N15)"
    If Selection = 0 Then Selection.ClearContents
     
        Range("O16").Select
    ActiveCell.Formula = "=COUNTIF(B:B,N16)"
    If Selection = 0 Then Selection.ClearContents
     
        Range("O17").Select
    ActiveCell.Formula = "=COUNTIF(B:B,N17)"
    If Selection = 0 Then Selection.ClearContents
     
        Range("O18").Select
    ActiveCell.Formula = "=COUNTIF(B:B,N18)"
    If Selection = 0 Then Selection.ClearContents
    
       Range("O19").Select
    ActiveCell.Formula = "=COUNTIF(B:B,N19)"
    If Selection = 0 Then Selection.ClearContents
     
       Range("O20").Select
    ActiveCell.Formula = "=COUNTIF(B:B,N20)"
    If Selection = 0 Then Selection.ClearContents
     
        Range("O21").Select
    ActiveCell.Formula = "=COUNTIF(B:B,N21)"
    If Selection = 0 Then Selection.ClearContents
     
        Range("O22").Select
    ActiveCell.Formula = "=COUNTIF(B:B,N22)"
    If Selection = 0 Then Selection.ClearContents
     
        Range("O23").Select
    ActiveCell.Formula = "=COUNTIF(B:B,N23)"
    If Selection = 0 Then Selection.ClearContents
     
        Range("O24").Select
    ActiveCell.Formula = "=COUNTIF(B:B,N24)"
    If Selection = 0 Then Selection.ClearContents
     
        Range("O25").Select
    ActiveCell.Formula = "=COUNTIF(B:B,N25)"
    If Selection = 0 Then Selection.ClearContents
     
        Range("O26").Select
    ActiveCell.Formula = "=COUNTIF(B:B,N26)"
    If Selection = 0 Then Selection.ClearContents
     
        Range("O27").Select
    ActiveCell.Formula = "=COUNTIF(B:B,N27)"
    If Selection = 0 Then Selection.ClearContents
    
       Range("O28").Select
    ActiveCell.Formula = "=COUNTIF(B:B,N28)"
    If Selection = 0 Then Selection.ClearContents
     
        Range("O29").Select
    ActiveCell.Formula = "=COUNTIF(B:B,N29)"
    If Selection = 0 Then Selection.ClearContents
    
       Range("O30").Select
    ActiveCell.Formula = "=COUNTIF(B:B,N30)"
    If Selection = 0 Then Selection.ClearContents
     
        Range("O31").Select
    ActiveCell.Formula = "=COUNTIF(B:B,N31)"
    If Selection = 0 Then Selection.ClearContents
     
        Range("O32").Select
    ActiveCell.Formula = "=COUNTIF(B:B,N32)"
    If Selection = 0 Then Selection.ClearContents
     
        Range("O33").Select
    ActiveCell.Formula = "=COUNTIF(B:B,N33)"
    If Selection = 0 Then Selection.ClearContents
     
        Range("O34").Select
    ActiveCell.Formula = "=COUNTIF(B:B,N34)"
    If Selection = 0 Then Selection.ClearContents
     
        Range("O35").Select
    ActiveCell.Formula = "=COUNTIF(B:B,N35)"
    If Selection = 0 Then Selection.ClearContents
     
       Range("O36").Select
    ActiveCell.Formula = "=COUNTIF(B:B,N36)"
    If Selection = 0 Then Selection.ClearContents
     
        Range("O37").Select
    ActiveCell.Formula = "=COUNTIF(B:B,N37)"
    If Selection = 0 Then Selection.ClearContents
     
       Range("O38").Select
    ActiveCell.Formula = "=COUNTIF(B:B,N38)"
    If Selection = 0 Then Selection.ClearContents
     
        Range("O39").Select
    ActiveCell.Formula = "=COUNTIF(B:B,N39)"
    If Selection = 0 Then Selection.ClearContents
     
        Range("O40").Select
    ActiveCell.Formula = "=COUNTIF(B:B,N40)"
    If Selection = 0 Then Selection.ClearContents
     
        Range("O41").Select
    ActiveCell.Formula = "=COUNTIF(B:B,N41)"
    If Selection = 0 Then Selection.ClearContents
     
       Range("O42").Select
    ActiveCell.Formula = "=COUNTIF(B:B,N42)"
    If Selection = 0 Then Selection.ClearContents
     
        Range("O43").Select
    ActiveCell.Formula = "=COUNTIF(B:B,N43)"
    If Selection = 0 Then Selection.ClearContents
     
        Range("O44").Select
    ActiveCell.Formula = "=COUNTIF(B:B,N44)"
    If Selection = 0 Then Selection.ClearContents
     
        Range("O45").Select
    ActiveCell.Formula = "=COUNTIF(B:B,N45)"
    If Selection = 0 Then Selection.ClearContents
     
        Range("O46").Select
    ActiveCell.Formula = "=COUNTIF(B:B,N46)"
    If Selection = 0 Then Selection.ClearContents
     
        Range("O47").Select
    ActiveCell.Formula = "=COUNTIF(B:B,N47)"
    If Selection = 0 Then Selection.ClearContents
     
        Range("O48").Select
    ActiveCell.Formula = "=COUNTIF(B:B,N48)"
    If Selection = 0 Then Selection.ClearContents
     
        Range("O49").Select
    ActiveCell.Formula = "=COUNTIF(B:B,N49)"
    If Selection = 0 Then Selection.ClearContents
     
        Range("O50").Select
    ActiveCell.Formula = "=COUNTIF(B:B,N50)"
    If Selection = 0 Then Selection.ClearContents
     
        Range("O51").Select
    ActiveCell.Formula = "=COUNTIF(B:B,N51)"
    If Selection = 0 Then Selection.ClearContents
     
        Range("O52").Select
    ActiveCell.Formula = "=COUNTIF(B:B,N52)"
    If Selection = 0 Then Selection.ClearContents
     
        Range("O53").Select
    ActiveCell.Formula = "=COUNTIF(B:B,N53)"
    If Selection = 0 Then Selection.ClearContents
     
        Range("O54").Select
    ActiveCell.Formula = "=COUNTIF(B:B,N54)"
    If Selection = 0 Then Selection.ClearContents
     
        Range("O55").Select
    ActiveCell.Formula = "=COUNTIF(B:B,N55)"
    If Selection = 0 Then Selection.ClearContents
     
        Range("O56").Select
    ActiveCell.Formula = "=COUNTIF(B:B,N56)"
    If Selection = 0 Then Selection.ClearContents
     
        Range("O57").Select
    ActiveCell.Formula = "=COUNTIF(B:B,N57)"
    If Selection = 0 Then Selection.ClearContents
     
        Range("O58").Select
    ActiveCell.Formula = "=COUNTIF(B:B,N58)"
    If Selection = 0 Then Selection.ClearContents
     
        Range("O59").Select
    ActiveCell.Formula = "=COUNTIF(B:B,N59)"
    If Selection = 0 Then Selection.ClearContents
     
        Range("O60").Select
    ActiveCell.Formula = "=COUNTIF(B:B,N60)"
    If Selection = 0 Then Selection.ClearContents
     
        Range("O61").Select
    ActiveCell.Formula = "=COUNTIF(B:B,N61)"
    If Selection = 0 Then Selection.ClearContents

Range("N:O").Sort Key1:=Range("O1"), Header:=xlYes, Order1:=xlDescending
        
Columns("A:B").EntireColumn.AutoFit
Columns("K:K").EntireColumn.AutoFit
Columns("N:O").EntireColumn.AutoFit

'[L:L, M:M, P:P, Q:Q, R:R, S:S, T:T, U:U, V:V, W:W, X:X, Y:Y, Z:Z, AA:AA, AB:AB, AC:AC, AD:AD, AE:AE, AF:AF, AG:AG, AH:AH, AI:AI, AJ:AJ, AK:AK, AL:AL, AM:AM, AN:AN, AO:AO].Interior.Color = RGB(232, 232, 232)

Range("A1:K1").Select
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

Range("N1:O1").Select
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

Range("A1").Select
     
    'Dim myRange As Range
    Set myRange = Range("L1:BB600")
    For Each myCell In myRange
        If myCell.Text = "" Then
            myCell.Interior.Color = RGB(232, 232, 232)
        End If
    Next
    
Next

'____________________________________________________________
    If Lr > 1 Then WSNew.Range("A" & Lr + 1).EntireRow.Delete

        My_Range.AutoFilter Field:=FieldNum

        Next cell

        On Error Resume Next
        Application.DisplayAlerts = False
        .Delete
        Application.DisplayAlerts = True
        On Error GoTo 0

    End With

    My_Range.Parent.AutoFilterMode = False

    If ErrNum > 0 Then
        MsgBox "Zmien recznie nazwy zkladek zaczynajacych sie od ""Error_""." & vbNewLine & "Wystepuja niedozwolone dla nazwy zakladki znaki" & vbNewLine & "lub nazwa jest za dluga."
    End If

    My_Range.Parent.Select
    ActiveWindow.View = ViewMode
    
    With Application
        .ScreenUpdating = True
        .EnableEvents = True
        .Calculation = CalcMode
    End With

End Sub

Function LastRow(sh As Worksheet)

    On Error Resume Next
    LastRow = sh.Cells.Find(What:="*", _
                            After:=sh.Range("A1"), _
                            Lookat:=xlPart, _
                            LookIn:=xlValues, _
                            SearchOrder:=xlByRows, _
                            SearchDirection:=xlPrevious, _
                            MatchCase:=False).Row
    On Error GoTo 0
    
End Function

Function SheetExists(SName As String, Optional ByVal WB As Workbook) As Boolean

    On Error Resume Next
    If WB Is Nothing Then Set WB = ThisWorkbook
    SheetExists = CBool(Len(WB.Sheets(SName).Name))
    
End Function