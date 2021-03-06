Attribute VB_Name = "Header"
'定数'
'データシート'
Public Const constDataSheetName = "データ"  'シート名'
'行'
Public Const constInitDataLine = 2          'データの最初の行'
'列'
Public Const constNoRow = 1                 'Noの列(A列)'
Public Const constRawRow = 2                '呼吸音の列(B列)'
Public Const constRawHBRemovRow = 3         '心拍除去後の呼吸音の列(C列)'
Public Const constRawSnoreRow = 4           'いびき音の列(D列)'
'Public Const constRawMovAvrRow = 4          '呼吸音の移動平均の列(D列)''
Public Const constApneaStateRow = 5         '無呼吸判定結果の入った列(E列)'
Public Const constSnoreStateRow = 6         'いびき判定結果の入った列(F列)'
Public Const constPhotorefRow = 7           'フォトセンサー値の入った列(G列)'
Public Const constAcceXRow = 8              '加速度(X)の入った列(H列)'
Public Const constAcceYRow = 9              '加速度(Y)の入った列(I列)'
Public Const constAcceZRow = 10             '加速度(Z)の入った列(J列)'
Public Const constRetAcceStartRow = 11      '向き(K列)'
Public Const constRetAcceEndRow = 18        '向き(R列)'


'結果シート'
Public Const constRetSheetName = "結果"     'シート名'
'行'
Public Const constInitRetLine = 44           '結果が入力される最初の行'
'列'
Public Const constRetStartTimeRow = 2       '判定時刻の列(B列)'
Public Const constRetStopTimeRow = 3        '停止時刻の列(C列)'
Public Const constRetContinuTimeRow = 4     '継続時間の列(D列)'
Public Const constRetTypeRow = 5            '種別の列(E列)'
Public Const constRetStartFromStopTimeRow = 6   '前回停止から今回発生までの時間の列(F列)'
Public Const constRetRemarkRow = 7          '備考の列(G列)'
'入力文字'
Public Const constSnore = "いびき"
Public Const constApnea = "無呼吸"

'コピーシート'
Public Const constCopySheetName = "copy"     'シート名'
