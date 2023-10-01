Sub うるう年判定_2()
    '応用編_1
    '割り切れるだの割り切れないだのあるけど、そもそも「2月が29日あればうるう年じゃね：という方法
    'だって定義は示したけど、それを使えとは言っていないからね（あれ？）
    'DateSerial関数を使用して、2月末日を取得し、29日だったらうるう年と判定する

    Worksheets("うるう年").Activate
    
    Dim r As Long '列ループカウンタ
    Dim determine_year As Long '判定対象年
    Dim last_date As Variant 'yyyy/m/d形式の前月末日取得して格納する変数
    Dim last_day As Variant '前月末日から日付だけを取得して格納する変数
    
    For r = 2 To 25
        determine_year = Cells(r, 1).Value
        'DateSerial関数を使用して、2月の末日を取得する
        last_day = DateSerial(determine_year, 3, 0)
        last_date = Day(last_day) 'Day関数で日付だけ取得する
        
        If last_date = 29 Then 'もし2月の末日が29日だったら
            Cells(r, 2).Value = "うるう年"
        End If
    Next r
    

End Sub
