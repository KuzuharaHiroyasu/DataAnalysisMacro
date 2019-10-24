Attribute VB_Name = "Read"
'
'データ読み込み
'
Sub readData()
    Dim ret As Boolean
    Dim msg As String
    Dim startTime As String

    Dim fileName As String
    Dim sheetNameCSV As String
    Dim Path As String
    
    Application.Calculation = xlManual
    
    'パス取得'
    Path = ThisWorkbook.Path + "\"
    
    'csvファイル検索しファイル名取得'
    fileName = Dir(Path & "*.csv")
    
    Do While Len(fileName) > 0
        'シート名取得'
        sheetNameCSV = left(fileName, 14)
        
        'コピー元のシートがあるファイルを開く
        Workbooks.Open (Path + fileName)
         
        'シートをコピー(新しいファイルに作成)
        Workbooks(fileName).Worksheets(sheetNameCSV).Copy After:=ThisWorkbook.Sheets(1)
         
        'コピー元ファイルを閉じる
        Workbooks(fileName).Close
        
        'csvからデータシートにデータをセット'
        dataSet (sheetNameCSV)
        
        '開始時間記入'
        startTime = setStartTime(sheetNameCSV)
        ThisWorkbook.Sheets(constRetSheetName).Range("B3").Value = startTime
        
        Application.DisplayAlerts = False ' メッセージを非表示
    
        'コピーしたcsvファイルのシート削除'
        ThisWorkbook.Sheets(sheetNameCSV).Delete
        
        'データ解析'
        Analysis.dataAnalysis
        
        'セットしたデータクリア'
        Clear.dataClear
        Clear.retClear
        
        '次のcsvファイル検索しファイル名取得'
        fileName = Dir()
    Loop
       
        Application.Calculation = xlAutomatic

'    If Not msg = "" Then
'        msg = buf + "を読み込めませんでした。"
'    Else
'        msg = buf
'    End If
'
'    MsgBox msg
    MsgBox "完了しました。"
    
    Worksheets(constCopySheetName).Activate ' 「Sheet1」のシートをアクティブ
End Sub

'
'データテキスト読み込み
'
Public Function readText(ByVal fileName As String, ByVal inputRow As Long) As Boolean
    Dim a
    Dim inputLine As Long
    
    inputLine = 2
    
    a = Dir(fileName)
    If (a <> "") Then
        Open fileName For Input As #1
        Do Until EOF(1)
            Line Input #1, buf
            Sheets(constDataSheetName).Cells(inputLine, inputRow) = buf
            'DoEvents
            inputLine = inputLine + 1
        Loop
        Close #1
        readText = True
    Else
        readText = False
    End If
End Function

'
'加速度センサー値の絶対値
'
Sub absoluteValue(ByVal x As Integer, ByVal y As Integer, ByVal z As Integer)
    x = Abs(x)
    y = Abs(y)
    z = Abs(z)
End Sub

'
'体の向きを決める
'
Sub acceAnalysis()
    Dim x As Long                    '加速度センサー_X軸'
    Dim y As Long                    '加速度センサー_Y軸'
    Dim z As Long                    '加速度センサー_Z軸'
    Dim line As Long
    Dim x_abs As Integer
    Dim z_abs As Integer
    
    line = constInitDataLine
    
    While IsEmpty(Sheets(constDataSheetName).Cells(line, constAcceXRow)) = False
        DoEvents
        x = Sheets(constDataSheetName).Cells(line, constAcceXRow).Value
        y = Sheets(constDataSheetName).Cells(line, constAcceYRow).Value
        z = Sheets(constDataSheetName).Cells(line, constAcceZRow).Value
        
        If x > 200 Or y > 200 Or z > 200 Then
            'エラー回避(イレギュラーな値)'
            Sheets(constDataSheetName).Range(Cells(line, constAcceXRow), Cells(line, constAcceZRow)).Delete
            GoTo Continue
        End If
        
        x_abs = Abs(x)
        z_abs = Abs(z)
        
        'ヘッド部の場所（体の向きではない）'
        If 0 <= x Then
            '上or左'
            If 0 <= z Then
                '左'
                Sheets(constDataSheetName).Cells(line, constRetAcceStartRow + 6).Value = 1
            Else
                '上'
                Sheets(constDataSheetName).Cells(line, constRetAcceStartRow).Value = 7
            End If
        Else
            '下or右'
            If 0 <= z Then
                '下'
                Sheets(constDataSheetName).Cells(line, constRetAcceStartRow + 4).Value = 3
            Else
                '右'
                Sheets(constDataSheetName).Cells(line, constRetAcceStartRow + 2).Value = 5
            End If
        End If
        line = line + 1
Continue:
    Wend
End Sub

'
'データセット
'
Public Function dataSet(ByVal sheetNameCSV As String) As Boolean
    Dim cnt_csv_line As Long
    Dim cnt_csv_row_kokyu As Long
    Dim cnt_csv_row_acce As Long
    Dim cnt_dst_line As Long
    
    Set sh_dst = Sheets("データ")
    
    'csvファイルのデータの開始位置'
    cnt_csv_line = 4
    cnt_csv_row_kokyu = 4
    
    'データをセットする開始行'
    cnt_dst_line = 2
    
    While IsEmpty(Worksheets(sheetNameCSV).Cells(cnt_csv_line, cnt_csv_row_kokyu)) = False
        'データセット'
        If cnt_csv_row_kokyu <= 6 Then
            'いびき、無呼吸判定セット'
            If Worksheets(sheetNameCSV).Cells(cnt_csv_line, cnt_csv_row_kokyu) = 0 Then
                sh_dst.Range(sh_dst.Cells(cnt_dst_line, "E"), sh_dst.Cells(cnt_dst_line, "F")).Value = 0
            ElseIf Worksheets(sheetNameCSV).Cells(cnt_csv_line, cnt_csv_row_kokyu) = 1 Then
                sh_dst.Cells(cnt_dst_line, "E").Value = 0
                sh_dst.Cells(cnt_dst_line, "F").Value = 1
            ElseIf Worksheets(sheetNameCSV).Cells(cnt_csv_line, cnt_csv_row_kokyu) = 2 Then
                sh_dst.Cells(cnt_dst_line, "E").Value = 2
                sh_dst.Cells(cnt_dst_line, "F").Value = 0
            End If
            
            cnt_csv_row_acce = cnt_csv_row_kokyu + 7
            '首の向きセット'
            If Worksheets(sheetNameCSV).Cells(cnt_csv_line, cnt_csv_row_acce) = 0 Then
                '左'
                sh_dst.Cells(cnt_dst_line, constRetAcceStartRow + 6).Value = 1
            ElseIf Worksheets(sheetNameCSV).Cells(cnt_csv_line, cnt_csv_row_acce) = 1 Then
                '上'
                sh_dst.Cells(cnt_dst_line, constRetAcceStartRow).Value = 7
            ElseIf Worksheets(sheetNameCSV).Cells(cnt_csv_line, cnt_csv_row_acce) = 2 Then
                '右'
                sh_dst.Cells(cnt_dst_line, constRetAcceStartRow + 2).Value = 5
            Else
                '下'
                sh_dst.Cells(cnt_dst_line, constRetAcceStartRow + 4).Value = 3
            End If
        End If
        
        If cnt_csv_row_kokyu < 6 Then
            '呼吸状態の次の値へ'
            cnt_csv_row_kokyu = cnt_csv_row_kokyu + 1
        Else
            '次の行の呼吸状態の最初の値へ'
            cnt_csv_row_kokyu = 4
            cnt_csv_line = cnt_csv_line + 1
        End If
        
        'データをセットする行を次へ'
        cnt_dst_line = cnt_dst_line + 1
    Wend
End Function

'
'開始時間セット
'
Public Function setStartTime(ByVal sheetNameCSV As String) As String
    Dim year As String
    Dim time As Date
    
    
    year = Worksheets(sheetNameCSV).Range("A3").Value
    time = Worksheets(sheetNameCSV).Range("C3").Value
    
    setStartTime = year + " " + time
    
End Function


