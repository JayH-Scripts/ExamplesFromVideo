Attribute VB_Name = "ParallelizeSerialSeries"
Option Explicit

Public Const Up1Row As Long = 1

Public Const ColDATE As Long = 1
Public Const ColData As Long = 2

Public Const Col1p    As Long = 5
Public Const Col2p    As Long = 6
Public Const Col1C    As Long = 7
Public Const Col2C    As Long = 8

Public Const Min14Sec59 As Double = 1.04050925874617E-02   '14 minutes:59 seconds
Public Const Min4Sec59 As Double = 3.46064814948477E-03    ' 4 minutes:59 seconds


Sub Parallelize4TimeSeries()
'works.
'   Sorted by Date and Time,
Dim s:  s = ""
    s = s & " If you did not sort by date/time yet, CTRL-BRK or CTRL-PAUSE NOW, and do it!"

    s = s & " After it runs, drag all the"
    s = s & " numbers in EFGH to BCDE, clobbering whats there."
    MsgBox s
        'A                 B     C     D  (sort by A, C, D)
    '3/1/2021  10:00      0.1    1     p
    '3/1/2021  10:00        2    2     C  <--START IT HERE ON ROW '2'
    '3/1/2021  10:00     0.05    2     p
    '3/1/2021  10:30       28    1     C
    '3/1/2021  10:30      7.2    1     p
    '3/1/2021  10:30        2    2     C
    '3/1/2021  10:30     0.05    2     p
    '3/1/2021  11:00       28    1     C
    '3/1/2021  11:00     7.25    1     p
    '3/1/2021  11:00        2    2     C
    '3/1/2021  11:00     0.05    2     p
    '3/1/2021  11:30       28    1     C
    '3/1/2021  11:30      7.1    1     p
    '3/1/2021  11:30        2    2     C
    '3/1/2021  11:30     0.05    2     p
    '3/1/2021  12:00       28    1     C
    '3/1/2021  12:00     7.05    1     p
    '3/1/2021  12:00        2    2     C
    '3/1/2021  12:00     0.05    2     p
    '3/1/2021  12:30       28    1     C
    '3/1/2021  12:30      7.3    1     p
    '3/1/2021  12:30        2    2     C
    '3/1/2021  12:30     0.05    2     p
    '3/1/2021  13:00       28    1     C
    
    
    'Get this as result (after manual dragging step mentioned above):
    '   A            B       C       D       E
    '3/1/2021 13:00  7.2    7.35    28      16
    '3/1/2021 13:30  7.1    7.35    28      16
    '3/1/2021 14:00  7      7.35    26      14
    '3/1/2021 14:30  7.25   7.15    26      16
    '3/1/2021 15:00  7.2    7.1     28      20
    '3/1/2021 15:30  7.1    7.05    30      20
    '3/1/2021 16:00  7.05   7.05    30      20
    '3/1/2021 16:30  7.45   7.1     30      16

    On Error Resume Next
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    If Not ActiveSheet.Name = "Merge" Then Exit Sub

    Do Until ActiveCell.Value = ""

        If DateTimeWithinXofAbove(Min14Sec59) Then
            copyDataToColumnForDevice Up1Row
            ActiveCell.EntireRow.Delete
        Else
            copyDataToColumnForDevice
            go_down
        End If
        
    Loop

    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
End Sub




Sub copyDataToColumnForDevice(Optional ByVal rowsAbove As Integer = 0)
      Cells(ActiveCell.Row - rowsAbove, myDevicecolumn(concatCandD)).Value2 _
    = Cells(ActiveCell.Row, ColData).Value2
End Sub







Function DateTimeWithinXofAbove(Xmin As Double) As Boolean
    DateTimeWithinXofAbove = False
    If cellDiffersFromCellAboveByLessThan(ColDATE, Xmin) Then
        DateTimeWithinXofAbove = True
    End If
End Function








Function cellDiffersFromCellAboveByLessThan(iColumn As Long, threshld As Double) As Boolean
    Dim activRow As Long
    activRow = ActiveCell.Row
    cellDiffersFromCellAboveByLessThan = False
    '14 minutes:59 seconds
    If Cells(activRow, iColumn).Value2 - Cells(activRow - 1, iColumn).Value2 <= threshld Then
        cellDiffersFromCellAboveByLessThan = True
    End If
End Function







Function isSomeCellSameAsAbove(iColumn As Long) As Boolean
    Dim activRow As Long
    activRow = ActiveCell.Row
    isSomeCellSameAsAbove = False
    If Cells(activRow, iColumn).Value2 = Cells(activRow - 1, iColumn).Value2 Then
        isSomeCellSameAsAbove = True
    End If
End Function







Function concatCandD()
    'C or p, 1 or 2:  --> 1C, 1p, 2C, 2p for func myDevicecolumn
    Dim activRow As Long
    activRow = ActiveCell.Row
    concatCandD = Range("C" & activRow).Value & Range("D" & activRow).Value
End Function






Function myDevicecolumn(NumberLetter As String) As Long
    Select Case NumberLetter
        Case "1p": myDevicecolumn = Col1p: Exit Function   '"E"
        Case "2p": myDevicecolumn = Col2p: Exit Function   '"F"
        Case "1C": myDevicecolumn = Col1C: Exit Function   '"G"
        Case "2C": myDevicecolumn = Col2C: Exit Function   '"H"
    End Select
End Function






Private Sub go_down()
    ActiveCell.Offset(1, 0).Select
End Sub





Function isSomeRangeSameAsAbove(iCols As Integer, Optional ByVal skipCol As Integer = 0) As Boolean
    On Error GoTo errHandler
    Dim i As Integer, activRow
    isSomeRangeSameAsAbove = True
    activRow = ActiveCell.Row
    For i = 1 To iCols
        If i = skipCol Then GoTo skippedThatCol
        If Not (Cells(activRow, i).Value2 = Cells(activRow - 1, i).Value2) Then
            isSomeRangeSameAsAbove = False
            Exit Function
        End If

skippedThatCol:
    Next i
    Exit Function

errHandler:
    If Not Err.Number = 0 Then isSomeRangeSameAsAbove = False
    Exit Function
End Function

