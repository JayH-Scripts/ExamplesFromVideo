Attribute VB_Name = "ModBinning"
Option Explicit
Public Const Row1 As Long = 4
Public Const ColumnsPer As Long = 1  '2, when RM, % pairs
Public Const BinCount As Long = 6   '6 on Maguire
Public Const ColCount As Long = 6

Sub RUNrandomizeAndTEST()
    randomizeBM4BR9
        Application.Wait (Now + TimeValue("0:00:04"))
    testReorder
End Sub

Sub testReorder()
    Sheet2.Range("BM4").Activate
    reorder_items_max_left_to_right_repeats ActiveSheet, _
        ActiveCell.Column, 6, 6, 1
    Sheet2.Range("BM4").Activate
End Sub
Private Sub testReorderBottomUp()
'_bottomUp
    reorder_items_max_left_to_right_repeats ActiveSheet, _
        ActiveCell.Column, 6, 6, -1
End Sub

Private Sub test_left_to_right()
'Does not help
reorder_items_max_left_to_right_repeats_Left_to_Right ActiveSheet, _
        ActiveCell.Column, 6, 6, 1
End Sub

Private Sub reorder_items_max_left_to_right_repeats(wksht As Worksheet, _
    col1 As Long, maxBins As Long, maxRecipes As Long, ByVal direction As Integer)

    Dim here As Range
    Set here = wksht.Cells(Row1, col1)
        here.Activate
        
    Dim cond
    For cond = 1 To maxRecipes - 1
        Do While WithinTheBox(here, col1, direction)
            If Not Adjacent(here, ColumnsPer).Value = here.Value Then
                   Dim there As Range
                   Set there = Matching_R_ange(here, direction)
                If Not there Is Nothing Then swapThem Adjacent(here, ColumnsPer), there
            End If
NextItemDown:
            Set here = here.Offset(direction, 0)
                here.Activate
                'Debug.Assert here.Address <> "$AZ$6"
          DoEvents
        Loop
NextCond:
        Select Case direction
            Case 1
                Set here = Cells(Row1, here.Column + ColumnsPer)
            Case -1
                Set here = Cells(Row1 + maxBins - 1, here.Column + ColumnsPer)
        End Select
        here.Activate
    Next cond
End Sub

Function Adjacent(fromHereOnLeft As Range, colsRight As Long) As Range
    Set Adjacent = fromHereOnLeft.Offset(0, colsRight)
End Function

Function Matching_R_ange(fromHereOnLeft As Range, _
                         ByVal direction As Integer) As Range
    
    Dim rowStart As Long
        rowStart = Row1
        
    Dim colLook As Long
        colLook = fromHereOnLeft.Offset(0, ColumnsPer).Column
        
    Dim c As Range
    Set c = Cells(rowStart, colLook)
    
    Dim col1 As Long
    col1 = c.Column
    
    Do While WithinTheBox(c, col1, direction)
        Debug.Print "C " & c.Address
    
        If c.Value = fromHereOnLeft.Value _
        And c.Row <> fromHereOnLeft.Row Then
            Set Matching_R_ange = c
            Exit Function
        Else
                Set c = c.Offset(1 * direction, 0)
        End If
      DoEvents
    Loop
    'returning NOTHING is expected, often
End Function

Function WithinTheBox(ByVal c As Range, ByVal col1 As Long, ByVal direction As Integer)
    Select Case direction
        Case 1
            WithinTheBox = c.Row <= Row1 + BinCount - 1 And c.Row >= Row1
        Case -1
            WithinTheBox = c.Row <= Row1 + BinCount - 1 And c.Row > Row1
    End Select
    WithinTheBox = WithinTheBox And _
               c.Column >= col1 And c.Column < col1 + ColCount - 1
End Function

Private Sub swapThem(range10 As Range, range20 As Range)
    'Unlike with SUB 'Matching_R_ange', we have to swap the %s as well as the items
    'So set temporary range vars to hold %s, to avoid confusion due to referencing items/r_anges
    If ColumnsPer = 2 Then
        Dim range11 As Range
        Set range11 = range10.Offset(0, 1)
        
        Dim range21 As Range
        Set range21 = range20.Offset(0, 1)
        'sit on them for now
    End If
    
    Dim Stak As Object
    Set Stak = CreateObject("System.Collections.Stack")
        Stak.push (range10.Value)           'A
        Stak.push (range20.Value)           'BA
                   range10.Value = Stak.pop 'A
                   range20.Value = Stak.pop '_  Stak is empty now, can re-use
                   
    If ColumnsPer = 2 Then
        Stak.push (range11.Value)
        Stak.push (range21.Value)
                   range11.Value = Stak.pop
                   range21.Value = Stak.pop
    End If
End Sub


'Not used. But it did help before we made the right tweaks.
Private Sub reorder_items_max_left_to_right_repeats_bottomUp(wksht As Worksheet, _
    col1 As Long, maxBins As Long, maxRecipes As Long, ByVal direction As Integer)
    Dim here As Range
    Set here = wksht.Cells(ActiveCell.Row, col1)
        here.Activate
    Dim cond
    For cond = 1 To maxRecipes - 1
        Do While WithinTheBox(here, direction)
            If Not Adjacent(here, ColumnsPer).Value = here.Value Then
                   Dim there As Range
                   Set there = Matching_R_ange(here, -1)
                If Not there Is Nothing Then swapThem Adjacent(here, ColumnsPer), there
            End If
NextItemUp:
            Set here = here.Offset(direction, 0)
                here.Activate
          DoEvents
        Loop
NextCond:
        Select Case direction
            Case 1
                Set here = Cells(Row1, here.Column + ColumnsPer)
            Case -1
                Set here = Cells(Row1 + maxBins - 1, here.Column + ColumnsPer)
        End Select
        here.Activate
    Next cond
End Sub

'Stop reading. Recorded macro for testing. No use outside of a very particular context.
Private Sub randomizeBM4BR9()
Application.ScreenUpdating = False

ActiveWorkbook.Worksheets("Sheet1").Activate
    ActiveWorkbook.Worksheets("Sheet1").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Sheet1").Sort.SortFields.Add2 Key:=Range("A2:A9") _
        , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortTextAsNumbers
    With ActiveWorkbook.Worksheets("Sheet1").Sort
        .SetRange Range("A2:B9")
        .Header = xlGuess
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
ActiveWorkbook.Worksheets("Sheet1").Activate
    Range("B2:B9").Select
    Selection.Copy
    Range("B11").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("A2:B9").Select
    Application.CutCopyMode = False
    ActiveWorkbook.Worksheets("Sheet1").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Sheet1").Sort.SortFields.Add2 Key:=Range("A2:A9") _
        , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortTextAsNumbers
    With ActiveWorkbook.Worksheets("Sheet1").Sort
        .SetRange Range("A2:B9")
        .Header = xlGuess
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    ActiveWorkbook.Worksheets("Sheet1").Activate
    Range("B2:B9").Select
    Selection.Copy
    Range("C11").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("A2:B9").Select
    Application.CutCopyMode = False
    ActiveWorkbook.Worksheets("Sheet1").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Sheet1").Sort.SortFields.Add2 Key:=Range("A2:A9") _
        , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortTextAsNumbers
    With ActiveWorkbook.Worksheets("Sheet1").Sort
        .SetRange Range("A2:B9")
        .Header = xlGuess
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    ActiveWorkbook.Worksheets("Sheet1").Activate
    Range("B2:B9").Select
    Selection.Copy
    Range("D11").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("A2:B9").Select
    Application.CutCopyMode = False
    ActiveWorkbook.Worksheets("Sheet1").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Sheet1").Sort.SortFields.Add2 Key:=Range("A2:A9") _
        , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortTextAsNumbers
    With ActiveWorkbook.Worksheets("Sheet1").Sort
        .SetRange Range("A2:B9")
        .Header = xlGuess
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    ActiveWorkbook.Worksheets("Sheet1").Activate
    Range("B2:B9").Select
    Selection.Copy
    Range("E11").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("A2:B9").Select
    Application.CutCopyMode = False
    ActiveWorkbook.Worksheets("Sheet1").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Sheet1").Sort.SortFields.Add2 Key:=Range("A2:A9") _
        , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortTextAsNumbers
    With ActiveWorkbook.Worksheets("Sheet1").Sort
        .SetRange Range("A2:B9")
        .Header = xlGuess
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    ActiveWorkbook.Worksheets("Sheet1").Activate
    Range("B2:B9").Select
    Selection.Copy
    Range("F11").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("A2:B9").Select
    Application.CutCopyMode = False
    ActiveWorkbook.Worksheets("Sheet1").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Sheet1").Sort.SortFields.Add2 Key:=Range("A2:A9") _
        , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortTextAsNumbers
    With ActiveWorkbook.Worksheets("Sheet1").Sort
        .SetRange Range("A2:B9")
        .Header = xlGuess
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    ActiveWorkbook.Worksheets("Sheet1").Activate
    Range("B2:B9").Select
    Selection.Copy
    Range("G11").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("B11:G16").Select
    Application.CutCopyMode = False
    Selection.Copy

Application.ScreenUpdating = True

    Sheets("A (2)").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
End Sub


