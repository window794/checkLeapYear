Sub うるう年判定_3()
    '応用編_2
    'うるう年は366日/年であることを利用して、対象年が366日かどうかを、DatePart関数を使って調べる
    'Ref: https://vba-create.jp/vba-tips-del-leap-year/
    
    Worksheets("うるう年").Activate
    
    Dim r As Long '列ループカウンタ
    Dim determine_year As Long '判定対象年
    Dim last_date As Variant '判定年の12/31を変数に格納
    
    For r = 2 To 25
        determine_year = Cells(r, 1).Value
        last_date = determine_year & "/12/31"
        If DatePart("y", last_date) = 366 Then '1/1から12/31までの日数が366日だったらうるう年
            Cells(r, 2).Value = "うるう年"
        End If
    Next r
    
End Sub
