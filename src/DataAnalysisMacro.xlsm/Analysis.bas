Attribute VB_Name = "Analysis"
Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal ms As LongPtr)

'
'各向きの時間用構造体
'
Type directionTime
    up As Integer
    rightUp As Integer
    right As Integer
    rightDown As Integer
    down As Integer
    leftDown As Integer
    left As Integer
    leftUp As Integer
End Type

'
'データ解析
'
Sub dataAnalysis()
    Dim snoreState As Integer           'いびき判定状態'
    Dim apneaState As Integer           '無呼吸判定状態'
    Dim beforeSnoreState As Integer     '１つ前のいびき判定状態'
    Dim beforeApneaState As Integer     '１つ前の無呼吸判定状態'
    Dim time As Integer                 '経過時間(秒)'
    Dim snoreCnt As Integer             'いびき回数'
    Dim apneaCnt As Integer             '無呼吸回数'
    Dim startTime As Date               '開始時刻'
    Dim dataLine As Long                '現在解析中のデータの行'
    Dim retLine As Long                 '現在結果入力中の行'
    Dim no As Long                      'ナンバー'
    Dim lastNo As Long                  '最終判定ナンバー'
    Dim remark As Long                  '備考用'
    Dim breath As directionTime         '通常呼吸の向き構造体'
    Dim snore As directionTime          'いびきの向き構造体'
    Dim apnea As directionTime          '無呼吸の向き構造体'
    
    ''''''初期化''''''
    snoreState = 0
    apneaState = 0
    time = 0
    snoreCnt = 0
    apneaCnt = 0
    no = 0
    dataLine = constInitDataLine
    retLine = constInitRetLine
    
    ''''''初期値設定''''''
    '回数0初期化'
    Sheets(constRetSheetName).Range("B24:H24").Value = 0
    Sheets(constRetSheetName).Range("B28:H28").Value = 0
    
    ''''''開始時刻設定''''''
    startTime = Sheets(constRetSheetName).Range("B3").Value
    
    ''''''解析''''''
    Application.ScreenUpdating = False
    Application.Calculation = xlManual
    
    While IsEmpty(Sheets(constDataSheetName).Cells(dataLine, constRawRow)) = False
        DoEvents
        Sheets(constDataSheetName).Cells(dataLine, constNoRow).Value = no 'ナンバー挿入'

        '呼吸の移動平均'
        Call movAverage(dataLine, no)

        If IsEmpty(Sheets(constDataSheetName).Cells(dataLine, constSnoreStateRow)) = False Then
            'いびき判定結果が入力されている'
            beforeSnoreState = snoreState                       '１つ前のいびき判定状態を保存'
            beforeApneaState = apneaState                       '１つ前の無呼吸判定状態を保存'
            snoreState = Sheets(constDataSheetName).Cells(dataLine, constSnoreStateRow).Value   'いびき状態取得'
            apneaState = Sheets(constDataSheetName).Cells(dataLine, constApneaStateRow).Value   '無呼吸状態取得'

            If snoreState = 1 Then
            'いびき判定あり'
                If beforeApneaState = 1 Or beforeApneaState = 2 Then
                '１つ前で無呼吸判定ありだった'
                    Call setRemarks(retLine, startTime, time, remark, no)
                    retLine = retLine + 1     '結果入力を次の行へ'
                End If

                If beforeSnoreState = 0 Then
                '１つ前はいびき判定なし'
                    Call setStart(retLine, startTime, time, constSnore)
                    snoreCnt = snoreCnt + 1
                    remark = no
                End If

                'いびきのトータル時間'
                Call calculationDirectionTime(no, snore)
            ElseIf apneaState = 1 Or apneaState = 2 Then
            '無呼吸判定あり'
                If beforeSnoreState = 1 Then
                '１つ前でいびき判定ありだった'
                    Call setRemarks(retLine, startTime, time, remark, no)
                    retLine = retLine + 1     '結果入力を次の行へ'
                End If

                If beforeApneaState = 0 Then
                '１つ前は無呼吸判定なし'
                    Call setStart(retLine, startTime, time, constApnea)
                    apneaCnt = apneaCnt + 1
                    remark = no
                End If

                '無呼吸のトータル時間'
                Call calculationDirectionTime(no, apnea)
            Else
                If beforeApneaState = 1 Or beforeApneaState = 2 Or beforeSnoreState = 1 Then
                '１つ前で無呼吸判定あり、もしくはいびき判定ありだった'
                    Call setRemarks(retLine, startTime, time, remark, no)
                    retLine = retLine + 1     '結果入力を次の行へ'
                End If

                '通常呼吸のトータル時間'
                Call calculationDirectionTime(no, breath)
            End If
            time = time + 10    '時間を10秒増やす'
            lastNo = no + 1
        End If

        no = no + 1
        dataLine = dataLine + 1     '次の行のデータ02へ'
    Wend

    If IsEmpty(Sheets(constRetSheetName).Cells(retLine, constRetTypeRow).Value) = False Then
        '最後の判定の停止時刻など'
        Call setRemarks(retLine, startTime, time, remark, lastNo)
    End If

    ''''''各種データ記入''''''
    Call setData(time, startTime, snoreCnt, apneaCnt)

    '各向きごとの通常呼吸の時間'
    Call setDirectionTime(breath, 9, 2)

    '各向きごとのいびきの時間'
    Call setDirectionTime(snore, 14, 2)

    '各向きごとの無呼吸の時間'
    Call setDirectionTime(apnea, 19, 2)

    '睡眠時間の割合'
    Call sleepTimeRatio

    'いびき抑制の割合'
    Call perOfSuppression(24, 36, 2, Sheets(constRetSheetName).Range("E3").Value)

    '無呼吸抑制の割合'
    Call perOfSuppression(28, 40, 2, Sheets(constRetSheetName).Range("F3").Value)

    ''''''加速度センサー''''''
    Dim endLine As Long
    Dim i As Long
    i = 1

    '向き判定の最終行'
    endLine = Sheets(constDataSheetName).Cells(rows.Count, constRetAcceStartRow).End(xlUp).row

    '最終の向きの行数検索'
    While i <= 7
        If endLine <= Sheets(constDataSheetName).Cells(rows.Count, constRetAcceStartRow + i).End(xlUp).row Then
            endLine = Sheets(constDataSheetName).Cells(rows.Count, constRetAcceStartRow + i).End(xlUp).row
        End If
        i = i + 1
    Wend
    
    ''''''グラフ作成''''''
    '既にグラフがあれば一旦削除'
    If Sheets(constRetSheetName).ChartObjects.Count > 0 Then
        Sheets(constRetSheetName).ChartObjects.Delete
    End If
    
    'グラフ作成'
    Call createGraph(endLine)
    
    'データを1行にコピー'
    Call copyData
    
    Application.Calculation = xlAutomatic
    Application.ScreenUpdating = True
    
    MsgBox "完了しました。"
End Sub

'
'いびき・無呼吸の開始時刻セット
'
Sub setStart(ByVal retLine As Long, ByVal startTime As Date, ByVal time As Long, ByVal kind As String)
    Sheets(constRetSheetName).Range(Cells(retLine, constRetStartTimeRow), Cells(retLine, constRetRemarkRow)).Font.Size = 10
    Sheets(constRetSheetName).Cells(retLine, constRetStartTimeRow).Value = DateAdd("s", time, startTime)   '開始時刻セット'
    Sheets(constRetSheetName).Cells(retLine, constRetStartTimeRow).NumberFormatLocal = "hh:mm:ss"         '時刻書式設定'
    Sheets(constRetSheetName).Cells(retLine, constRetTypeRow).Value = kind                                '種別セット'
    Sheets(constRetSheetName).Cells(retLine, constRetTypeRow).HorizontalAlignment = xlCenter
End Sub

'
'いびき・無呼吸の終了時刻セット
'
Sub setEnd(ByVal retLine As Long, ByVal startTime As Date, ByVal time As Long)
    Dim kind As String
    Dim duration As Date
    
    '停止時刻セット'
    Sheets(constRetSheetName).Cells(retLine, constRetStopTimeRow).Value = DateAdd("s", time, startTime)
    Sheets(constRetSheetName).Cells(retLine, constRetStopTimeRow).NumberFormatLocal = "hh:mm:ss"         '時刻書式設定'
    
    '継続時間セット'
    duration = Sheets(constRetSheetName).Cells(retLine, constRetStopTimeRow).Value - Sheets(constRetSheetName).Cells(retLine, constRetStartTimeRow).Value
    Sheets(constRetSheetName).Cells(retLine, constRetContinuTimeRow).Value = duration
    Sheets(constRetSheetName).Cells(retLine, constRetContinuTimeRow).NumberFormatLocal = "hh:mm:ss"      '継続時間書式設定'
    
    '再発経過時間セット'
    If retLine = constInitRetLine Then
        Sheets(constRetSheetName).Cells(retLine, constRetStartFromStopTimeRow).Value = "-"
    Else
        Sheets(constRetSheetName).Cells(retLine, constRetStartFromStopTimeRow).Value = Sheets(constRetSheetName).Cells(retLine, constRetStartTimeRow).Value - Sheets(constRetSheetName).Cells(retLine - 1, constRetStopTimeRow).Value
        Sheets(constRetSheetName).Cells(retLine, constRetStartFromStopTimeRow).NumberFormatLocal = "hh:mm:ss" '前回停止から今回発生までの時間書式設定'
    End If
    Sheets(constRetSheetName).Cells(retLine, constRetStartFromStopTimeRow).HorizontalAlignment = xlRight

    
    '継続時間ごとに回数をセット'
    '種別'
    kind = Sheets(constRetSheetName).Cells(retLine, constRetTypeRow).Value
    If kind = "いびき" Then
        Call setNumPerDuration(duration, 24)
    Else
        Call setNumPerDuration(duration, 28)
    End If
End Sub

'
'継続時間ごとに回数をセット
'
Sub setNumPerDuration(ByVal duration As Date, ByVal line As Integer)
    Dim durationInt As Integer
    
    'DateをIntegerに変換'
    durationInt = duration * 86400

    If durationInt = 10 Then
        '10秒'
        Sheets(constRetSheetName).Cells(line, 2).Value = Sheets(constRetSheetName).Cells(line, 2).Value + 1
    ElseIf durationInt = 20 Then
        '20秒'
        Sheets(constRetSheetName).Cells(line, 3).Value = Sheets(constRetSheetName).Cells(line, 3).Value + 1
    ElseIf durationInt >= 30 And durationInt < 60 Then
        '30秒以上1分未満'
        Sheets(constRetSheetName).Cells(line, 4).Value = Sheets(constRetSheetName).Cells(line, 4).Value + 1
    ElseIf durationInt >= 60 And durationInt < 120 Then
        '1分以上2分未満'
        Sheets(constRetSheetName).Cells(line, 5).Value = Sheets(constRetSheetName).Cells(line, 5).Value + 1
    ElseIf durationInt >= 120 And durationInt < 300 Then
        '2分以上5分未満'
        Sheets(constRetSheetName).Cells(line, 6).Value = Sheets(constRetSheetName).Cells(line, 6).Value + 1
    ElseIf durationInt >= 300 And durationInt < 600 Then
        '5分以上10分未満'
        Sheets(constRetSheetName).Cells(line, 7).Value = Sheets(constRetSheetName).Cells(line, 7).Value + 1
    Else
        '10分以上'
        Sheets(constRetSheetName).Cells(line, 8).Value = Sheets(constRetSheetName).Cells(line, 8).Value + 1
    End If
End Sub


'
'移動平均を求める
'
Sub movAverage(ByVal dataLine As Long, ByVal no As Long)
    If no >= 4 Then
        Sheets(constDataSheetName).Cells(dataLine, constRawMovAvrRow).Value = WorksheetFunction.Sum(Range(Sheets(constDataSheetName).Cells(dataLine - 4, constRawRow), Sheets(constDataSheetName).Cells(dataLine, constRawRow))) / 5
    Else
        Sheets(constDataSheetName).Cells(dataLine, constRawMovAvrRow).Value = "-"
    End If
End Sub

'
'備考欄に記入
'
Sub setRemarks(ByVal retLine As Long, ByVal startTime As Date, ByVal time As Long, ByVal remark As Long, ByVal lastNo As Long)
    Call setEnd(retLine, startTime, time)
    Sheets(constRetSheetName).Cells(retLine, constRetRemarkRow).Value = remark & "から" & lastNo
    Sheets(constRetSheetName).Cells(retLine, constRetRemarkRow).HorizontalAlignment = xlRight
End Sub

'
'各種データをセットする
'
Sub setData(ByVal time As Long, ByVal startTime As Date, ByVal snoreCnt As Integer, ByVal apneaCnt As Integer)
    '終了時刻'
    Sheets(constRetSheetName).Range("C3").Value = DateAdd("s", time, startTime)
    
    'データ取得時間'
    Sheets(constRetSheetName).Range("D3").Value = CStr(CDate(DateDiff("s", startTime, Sheets(constRetSheetName).Range("C3").Value) / 86400#))
    
    'いびき回数'
    Sheets(constRetSheetName).Range("E3").Value = snoreCnt
    
    '無呼吸回数'
    Sheets(constRetSheetName).Range("F3").Value = apneaCnt
End Sub

'
'各状態ごとの各向きの時間をセットする
'
Sub setDirectionTime(directTime As directionTime, ByVal line As Integer, ByVal row As Integer)
    Dim time As Date
    Dim totalTime As Integer
    
    '上'
    time = TimeSerial(0, 0, directTime.up)
    Sheets(constRetSheetName).Cells(line, row).NumberFormatLocal = "hh:mm:ss"
    Sheets(constRetSheetName).Cells(line, row).Value = time
    row = row + 1
    
    '右上'
    time = TimeSerial(0, 0, directTime.rightUp)
    Sheets(constRetSheetName).Cells(line, row).NumberFormatLocal = "hh:mm:ss"
    Sheets(constRetSheetName).Cells(line, row).Value = time
    row = row + 1
    
    '右'
    time = TimeSerial(0, 0, directTime.right)
    Sheets(constRetSheetName).Cells(line, row).NumberFormatLocal = "hh:mm:ss"
    Sheets(constRetSheetName).Cells(line, row).Value = time
    row = row + 1
    
    '右下'
    time = TimeSerial(0, 0, directTime.rightDown)
    Sheets(constRetSheetName).Cells(line, row).NumberFormatLocal = "hh:mm:ss"
    Sheets(constRetSheetName).Cells(line, row).Value = time
    row = row + 1
    
    '下'
    time = TimeSerial(0, 0, directTime.down)
    Sheets(constRetSheetName).Cells(line, row).NumberFormatLocal = "hh:mm:ss"
    Sheets(constRetSheetName).Cells(line, row).Value = time
    row = row + 1
    
    '左下'
    time = TimeSerial(0, 0, directTime.leftDown)
    Sheets(constRetSheetName).Cells(line, row).NumberFormatLocal = "hh:mm:ss"
    Sheets(constRetSheetName).Cells(line, row).Value = time
    row = row + 1
    
    '左'
    time = TimeSerial(0, 0, directTime.left)
    Sheets(constRetSheetName).Cells(line, row).NumberFormatLocal = "hh:mm:ss"
    Sheets(constRetSheetName).Cells(line, row).Value = time
    row = row + 1
    
    '左上'
    time = TimeSerial(0, 0, directTime.leftUp)
    Sheets(constRetSheetName).Cells(line, row).NumberFormatLocal = "hh:mm:ss"
    Sheets(constRetSheetName).Cells(line, row).Value = time
    row = row + 1
    
    '合計'
    totalTime = directTime.up + directTime.rightUp + directTime.right + directTime.rightDown + directTime.down + directTime.leftDown + directTime.left + directTime.leftUp
    time = TimeSerial(0, 0, totalTime)
    Sheets(constRetSheetName).Cells(line, row).NumberFormatLocal = "hh:mm:ss"
    Sheets(constRetSheetName).Cells(line, row).Value = time
End Sub

'
'グラフ作成
'
Sub createGraph(ByVal endLine As Long)
    Dim i As Long
'いびき/呼吸の大きさ'
    If IsEmpty(Sheets(constDataSheetName).Cells(constInitDataLine, constRawRow)) = False And IsEmpty(Sheets(constDataSheetName).Cells(constInitDataLine, constRawSnoreRow)) = False Then
        With Sheets(constRetSheetName).ChartObjects.Add(30, 50, 300, 200).Chart
            .ChartType = xlLine
            .SetSourceData Source:=Sheets(constDataSheetName).Range(Sheets(constDataSheetName).Cells(constInitDataLine, constRawRow), Sheets(constDataSheetName).Cells(rows.Count, constRawSnoreRow).End(xlUp))
            .ChartArea.Top = Sheets(constRetSheetName).Range("L7").Top
            .ChartArea.left = Sheets(constRetSheetName).Range("L7").left
            .SeriesCollection(1).Name = "=""呼吸音"""
            .SeriesCollection(2).Name = "=""いびき"""
            .Legend.Position = xlLegendPositionLeft
            .Axes(xlValue).MinimumScale = 0
            .Axes(xlValue).MaximumScale = 1024
            .Axes(xlValue).MajorUnit = 256
            .Axes(xlCategory).HasMajorGridlines = False
            .Axes(xlCategory).TickLabels.NumberFormatLocal = "G/標準"
            .Axes(xlCategory).MajorTickMark = xlNone
            .ChartArea.Width = 36000
            .ChartArea.Height = 150
            With .PlotArea
                Application.ScreenUpdating = False
               .Select
                With Selection
                    .left = 69
                    .Height = 140
                    .Width = 35940
                End With
                Application.ScreenUpdating = True
            End With
        End With
    End If

    'いびき/呼吸の判定'
    If IsEmpty(Sheets(constDataSheetName).Cells(constInitDataLine, constSnoreStateRow)) = False And IsEmpty(Sheets(constDataSheetName).Cells(constInitDataLine, constApneaStateRow)) = False Then
        With Sheets(constRetSheetName).ChartObjects.Add(30, 50, 300, 200).Chart
            .ChartType = xlLine
            .SetSourceData Source:=Sheets(constDataSheetName).Range(Sheets(constDataSheetName).Cells(constInitDataLine, constSnoreStateRow), Sheets(constDataSheetName).Cells(rows.Count, constApneaStateRow).End(xlUp))
            .ChartArea.Width = 36000
            .ChartArea.Height = 150
            .ChartArea.Top = Sheets(constRetSheetName).Range("L19").Top
            .ChartArea.left = Sheets(constRetSheetName).Range("L19").left
            .SeriesCollection(1).Name = "=""無呼吸"""
            .SeriesCollection(2).Name = "=""いびき"""
            .Axes(xlValue).MinimumScale = 0
            .Axes(xlValue).MaximumScale = 2
            .Axes(xlValue).MajorUnit = 1
            .Axes(xlCategory).HasMajorGridlines = False
            .Axes(xlCategory).TickLabels.NumberFormatLocal = "G/標準"
            .Legend.Position = xlLegendPositionLeft
            With .PlotArea
                Application.ScreenUpdating = False
               .Select
                With Selection
                    .Width = 50
                    .left = 96
                    .Height = 140
                    .Width = 35940
                End With
                Application.ScreenUpdating = True
            End With
        End With
    End If

    '体の向き'
    If endLine > 1 Then
        With Sheets(constRetSheetName).ChartObjects.Add(30, 50, 300, 200).Chart
            .ChartType = xlLine
            .SetSourceData Source:=Sheets(constDataSheetName).Range(Sheets(constDataSheetName).Cells(constInitDataLine - 1, constRetAcceStartRow), Sheets(constDataSheetName).Cells(endLine, constRetAcceEndRow))
            .ChartArea.Width = 36000
            .ChartArea.Height = 150
            .ChartArea.Top = Sheets(constRetSheetName).Range("L30").Top
            .ChartArea.left = Sheets(constRetSheetName).Range("L30").left
            .SeriesCollection(1).Name = "=""上"""
            .SeriesCollection(2).Name = "=""右上"""
            .SeriesCollection(3).Name = "=""右"""
            .SeriesCollection(4).Name = "=""右下"""
            .SeriesCollection(5).Name = "=""下"""
            .SeriesCollection(6).Name = "=""左下"""
            .SeriesCollection(7).Name = "=""左"""
            .SeriesCollection(8).Name = "=""左上"""
            .Axes(xlValue).MinimumScale = 0
            .Axes(xlValue).MaximumScale = 7
            .Axes(xlValue).MajorUnit = 1
            .Axes(xlCategory).HasMajorGridlines = False
            .Axes(xlCategory).TickLabels.NumberFormatLocal = "G/標準"
            .Axes(xlCategory).MajorTickMark = xlNone
            .Legend.Position = xlLegendPositionLeft
            With .PlotArea
                Application.ScreenUpdating = False
               .Select
                With Selection
                    .Width = 50
                    .left = 84
                    .Height = 140
                    .Width = 35940
                End With
                Application.ScreenUpdating = True
            End With
            With .SeriesCollection
                For i = 1 To .Count
                    .Item(i).Format.line.Weight = 3
                Next i
            End With
        End With
    End If

    '加速度センサー値'
    If IsEmpty(Sheets(constDataSheetName).Cells(constInitDataLine, constAcceXRow)) = False And IsEmpty(Sheets(constDataSheetName).Cells(constInitDataLine, constAcceYRow)) = False And IsEmpty(Sheets(constDataSheetName).Cells(constInitDataLine, constAcceZRow)) = False Then
        With Sheets(constRetSheetName).ChartObjects.Add(30, 50, 300, 200).Chart
            .ChartType = xlLine
            .SetSourceData Source:=Sheets(constDataSheetName).Range(Sheets(constDataSheetName).Cells(constInitDataLine, constAcceXRow), Sheets(constDataSheetName).Cells(rows.Count, constAcceZRow).End(xlUp))
            .ChartArea.Width = 36000
            .ChartArea.Height = 150
            .ChartArea.Top = Sheets(constRetSheetName).Range("L41").Top
            .ChartArea.left = Sheets(constRetSheetName).Range("L41").left
            .SeriesCollection(1).Name = "=""Ｘ軸"""
            .SeriesCollection(2).Name = "=""Ｙ軸"""
            .SeriesCollection(3).Name = "=""Ｚ軸"""
            .Axes(xlValue).MinimumScale = -100
            .Axes(xlValue).MaximumScale = 100
            .Axes(xlValue).MajorUnit = 50
            .Axes(xlCategory).HasMajorGridlines = False
            .Axes(xlCategory).TickLabels.NumberFormatLocal = "G/標準"
            .Axes(xlCategory).MajorTickMark = xlNone
            .Legend.Position = xlLegendPositionLeft
            With .PlotArea
                Application.ScreenUpdating = False
               .Select
                With Selection
                    .left = 100
                    .Height = 140
                    .Width = 35900
                End With
                Application.ScreenUpdating = True
            End With
            With .SeriesCollection
                For i = 1 To .Count
                    .Item(i).Format.line.Weight = 3
                Next i
            End With
        End With
    End If
    
    'フォトセンサー値'
    If IsEmpty(Sheets(constDataSheetName).Cells(constInitDataLine, constPhotorefRow)) = False Then
        With Sheets(constRetSheetName).ChartObjects.Add(30, 50, 300, 200).Chart
            .ChartType = xlLine
            .SetSourceData Source:=Sheets(constDataSheetName).Range(Sheets(constDataSheetName).Cells(constInitDataLine, constPhotorefRow), Sheets(constDataSheetName).Cells(rows.Count, constPhotorefRow).End(xlUp))
            .ChartArea.Width = 36000
            .ChartArea.Height = 150
            .ChartArea.Top = Sheets(constRetSheetName).Range("L53").Top
            .ChartArea.left = Sheets(constRetSheetName).Range("L53").left
            .SeriesCollection(1).Name = "=""ﾌｫﾄｾﾝｻｰ"""
            .Axes(xlValue).MinimumScale = 0
            .Axes(xlValue).MaximumScale = 1000
            .Axes(xlValue).MajorUnit = 200
            .Axes(xlCategory).HasMajorGridlines = False
            .Axes(xlCategory).TickLabels.NumberFormatLocal = "G/標準"
            .Axes(xlCategory).MajorTickMark = xlNone
            .Legend.Position = xlLegendPositionLeft
            With .PlotArea
                Application.ScreenUpdating = False
               .Select
                With Selection
                    .Top = 5
                    .left = 68
                    .Height = 140
                    .Width = 35900
                End With
                Application.ScreenUpdating = True
            End With
            With .SeriesCollection
                For i = 1 To .Count
                    .Item(i).Format.line.Weight = 3
                Next i
            End With
        End With
    End If
End Sub

'
'各状態ごとの各向きの時間を求める
'
Sub calculationDirectionTime(ByVal no As Long, directTime As directionTime)
    Dim line As Long
    Dim rows As Integer
    
    rows = 10
    
    '該当の向きの行'
    line = (no * 20) + 20
    
    '加速度センサーの値がカラなら上の行の値があるところまで遡る'
    While WorksheetFunction.CountA(Sheets(constDataSheetName).Cells(line, constAcceXRow)) = 0
        line = line - 1
    Wend
    
    '向きを検索'
    While WorksheetFunction.CountA(Sheets(constDataSheetName).Cells(line, rows)) = 0
        '空白'
        rows = rows + 1
    Wend

    Select Case rows
        Case 10 'J列(上)'
            directTime.up = directTime.up + 10
        Case 11 'K列(右上)'
            directTime.rightUp = directTime.rightUp + 10
        Case 12 'L列(右)'
            directTime.right = directTime.right + 10
        Case 13 'M列(右下)'
            directTime.rightDown = directTime.rightDown + 10
        Case 14 'N列(下)'
            directTime.down = directTime.down + 10
        Case 15 'O列(左下)'
            directTime.leftDown = directTime.leftDown + 10
        Case 16 'P列(左)'
            directTime.left = directTime.left + 10
        Case 17 'Q列(左上)'
            directTime.leftUp = directTime.leftUp + 10
    End Select
End Sub

'
'睡眠時間の割合
'
Sub sleepTimeRatio()
    Dim dataAcqTime As Variant 'データ取得時間'
    
    dataAcqTime = Sheets(constRetSheetName).Range("D3").Value
    
    '通常呼吸'
    Sheets(constRetSheetName).Range("B32").Value = Sheets(constRetSheetName).Range("J9").Value / dataAcqTime
    Sheets(constRetSheetName).Range("B32").NumberFormatLocal = "0.0%"
    
    'いびき'
    Sheets(constRetSheetName).Range("C32").Value = Sheets(constRetSheetName).Range("J14").Value / dataAcqTime
    Sheets(constRetSheetName).Range("C32").NumberFormatLocal = "0.0%"
    
    '無呼吸'
    Sheets(constRetSheetName).Range("D32").Value = Sheets(constRetSheetName).Range("J19").Value / dataAcqTime
    Sheets(constRetSheetName).Range("D32").NumberFormatLocal = "0.0%"
End Sub

'
'いびき・無呼吸抑制の割合
'
Sub perOfSuppression(ByVal line As Integer, ByVal retLine As Integer, ByVal row As Integer, ByVal totalCnt As Integer)
    Dim i As Integer
    '10秒　～　10分以上まで7項目分'
    If totalCnt = 0 Then
        'totalCntが0'
        Sheets(constRetSheetName).Range(Cells(retLine, row), Cells(retLine, row + 6)).Value = 0
        Sheets(constRetSheetName).Range(Cells(retLine, row), Cells(retLine, row + 6)).NumberFormatLocal = "0.0%"
    Else
        'totalCntが0以外'
        For i = 1 To 7
            Sheets(constRetSheetName).Cells(retLine, row).Value = Sheets(constRetSheetName).Cells(line, row).Value / totalCnt
            Sheets(constRetSheetName).Cells(retLine, row).NumberFormatLocal = "0.0%"
            row = row + 1
        Next i
    End If
End Sub

'
'解析結果コピー
'
Sub copyData()
    Dim line As Integer
    Dim row As Integer
    
    line = 1
    row = 1
    
    Sheets(constRetSheetName).Range("B3:F3").Copy Sheets(constCopySheetName).Cells(line, row)   '開始時刻, 終了時刻, データ取得時間, いびき回数, 無呼吸回数 + 空列
    row = row + 6
    
    Sheets(constRetSheetName).Range("J9").Copy Sheets(constCopySheetName).Cells(line, row)      '通常呼吸時間
    row = row + 1
    
    Sheets(constRetSheetName).Range("J14").Copy Sheets(constCopySheetName).Cells(line, row)     'いびき時間
    row = row + 1
    
    Sheets(constRetSheetName).Range("J19").Copy Sheets(constCopySheetName).Cells(line, row)     '無呼吸時間 + 空列
    row = row + 2
    
    Sheets(constRetSheetName).Range("B24:H24").Copy Sheets(constCopySheetName).Cells(line, row) 'いびき時間（回数）- 10秒, 20秒, 30秒以上1分未満, 1分以上2分未満, 2分以上5分未満, 5分以上10分未満, 10分以上 + 空列
    row = row + 8
    
    Sheets(constRetSheetName).Range("B28:H28").Copy Sheets(constCopySheetName).Cells(line, row) '無呼吸時間（回数）- 10秒, 20秒, 30秒以上1分未満, 1分以上2分未満, 2分以上5分未満, 5分以上10分未満, 10分以上 + 空列
    row = row + 8
    
    Sheets(constRetSheetName).Range("B32:D32").Copy Sheets(constCopySheetName).Cells(line, row) '割合 - 通常呼吸, いびき, 無呼吸 + 空列
    row = row + 4
    
    Sheets(constRetSheetName).Range("B36:H36").Copy Sheets(constCopySheetName).Cells(line, row) 'いびき時間（割合）- 10秒, 20秒, 30秒以上1分未満, 1分以上2分未満, 2分以上5分未満, 5分以上10分未満, 10分以上 + 空列
    row = row + 8
    
    Sheets(constRetSheetName).Range("B40:H40").Copy Sheets(constCopySheetName).Cells(line, row) '無呼吸時間（割合）- 10秒, 20秒, 30秒以上1分未満, 1分以上2分未満, 2分以上5分未満, 5分以上10分未満, 10分以上 + 空列
    row = row + 8
    
    Sheets(constRetSheetName).Range("B9:I9").Copy Sheets(constCopySheetName).Cells(line, row)   '通常呼吸時間 - 上, 右上, 右, 右下, 下, 左下, 左, 左上 + 空列
    row = row + 9
    
    Sheets(constRetSheetName).Range("B14:I14").Copy Sheets(constCopySheetName).Cells(line, row)   'いびき時間 - 上, 右上, 右, 右下, 下, 左下, 左, 左上 + 空列
    row = row + 9
    
    Sheets(constRetSheetName).Range("B19:I19").Copy Sheets(constCopySheetName).Cells(line, row)   '無呼吸時間 - 上, 右上, 右, 右下, 下, 左下, 左, 左上 + 空列
    row = row + 9
End Sub











