Attribute VB_Name = "Clear"
Sub retClear()
    Dim cnt As Long
    
    Application.ScreenUpdating = False
    Application.Calculation = xlManual
    
    'グラフ削除'
    If ActiveSheet.ChartObjects.Count > 0 Then
        ActiveSheet.ChartObjects.Delete
    End If
    
    '結果シート'
    rows(3).Clear
    rows(9).Clear
    rows(14).Clear
    rows(19).Clear
    rows(24).Clear
    rows(28).Clear
    rows(32).Clear
    rows(36).Clear
    rows(40).Clear
    
    '睡眠時間の表の枠線を復活させる'
    Range("B9:J20").BorderAround LineStyle:=xlContinuous
 
    cnt = 44    '44行目から'
    While IsEmpty(Cells(cnt, constRetStartTimeRow)) = False
        rows(cnt).Clear
        cnt = cnt + 1
    Wend
    
    'copyシートの出力結果削除'
    Sheets(constCopySheetName).rows(1).Clear
    
    Application.Calculation = xlAutomatic
    Application.ScreenUpdating = True
    
    MsgBox "削除完了しました。"
End Sub

Sub dataClear()
    Dim endLine As Long
    
    Application.ScreenUpdating = False
    Application.Calculation = xlManual
    
    'データシート'
    endLine = Cells(rows.Count, 2).End(xlUp).row
    Range(Cells(2, 1), Cells(endLine, 17)).Clear
    
    Application.Calculation = xlAutomatic
    Application.ScreenUpdating = True
    
    MsgBox "削除完了しました。"
End Sub
