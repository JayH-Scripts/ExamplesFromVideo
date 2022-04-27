Attribute VB_Name = "mod4Devices"
Option Explicit
Option Private Module


Public Sub CycleChartSeriesColors()
    
    Dim seriesCount As Integer, i As Integer, smod As Integer
    Const WT = 1.5
    Static s    '"Remember" how many times we clicked "Cycle Focus" button since
                ' the file has been open.
    
    ActiveSheet.ChartObjects(ActiveSheet.ChartObjects.Count).Activate
    
    'This is 4 in this example, always. But not set to constant 4, because
    '   we might add or remove chart series (more or fewer devices.) dynamic
    seriesCount = ActiveChart.SeriesCollection.Count
    
    'We didn't need to intitialize "s" to 0. VBA does that anyway.
    '   Every time we click "Cycle Focus" button, s gets incremented
    '   Doesn't matter how big it grows, because MODULO operation will just
    '   return a cycle of  1 ,  2 ,  3 ,  4 ,  0 ,  1 ,  2 ,  3 ,  4 ,  0 , ...
    
    s = s + 1   'we're just incrementing a count of button-clicks
                'but now we divide it by now many line-graphs we have
    smod = s Mod (seriesCount + 1)  ':which is s Mod 5 (4 + 1), but again...dynamic!

    If Not smod = 0 Then    's/5 has a remainder that isn't 0. 1, 2, 3, or 4
                            'in YouTube example. And 1,2,3, & 4 are our series!
                                        '1,2,3,4
        With ActiveChart.SeriesCollection(smod).Format.Line
            .Visible = msoTrue
            .ForeColor.RGB = RGB(192, 0, 0) 'so one of our series will be red
            .Transparency = 0       '       /
            .Weight = WT            '      /
        End With                    '     /
        For i = 1 To seriesCount    '    /
            If Not i = smod Then   'the series,  #i is to backgrounded (gray)
                With ActiveChart.SeriesCollection(i).Format.Line
                    .Visible = msoTrue
                    .ForeColor.RGB = RGB(240, 240, 240)
                    .Transparency = 0.85
                    .Weight = WT
                End With
            End If
        Next i
        
    Else    'else smod IS = 0, so we've cycled all our series red/gray
            'but now we to show all series, wildly colored.
       For i = 1 To seriesCount
            
            With ActiveChart.SeriesCollection(i).Format.Line
                .Visible = msoTrue
                Select Case i
                    Case 1: .ForeColor.RGB = RGB(200, 0, 0)
                    Case 2: .ForeColor.RGB = RGB(0, 200, 0)
                    Case 3: .ForeColor.RGB = RGB(0, 0, 200)
                    Case Else: .ForeColor.RGB = RGB(200, 66, 192)
                End Select
                .Transparency = 0
                .Weight = WT
            End With
            
        Next i
    End If
End Sub

