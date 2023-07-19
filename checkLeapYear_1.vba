Sub うるう年判定_1()
    'うるう年定義
    '1. 西暦年号が4で割り切れる年
    '2. 上の例外として、100で割り切れて400で割り切れない年は平年

    Worksheets("うるう年").Activate
    
    Dim r As Long '列ループカウンタ
    Dim determine_year As Long '判定対象年
    
    For r = 2 To 25
        determine_year = Cells(r, 1).Value
        
        '定義を以下のように読み替える
        '4で割り切れる　かつ　100で割り切れない　または　400で割り切れる
        If (determine_year Mod 4 = 0) And ((determine_year Mod 100 <> 0) Or (determine_year Mod 400) = 0) Then
            Cells(r, 2).Value = "うるう年"
        End If
    Next r
    
    MsgBox "Has Completed!"

End Sub
