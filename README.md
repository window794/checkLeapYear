# このリポジトリについて
このリポジトリはVBA講座代6回で出た応用問題「うるう年判定」の解答例をご紹介しているものです。<br>
回答は必ずしもこれらだけとは限りません。何か良い回答があればTeamsなりでご共有いただけると嬉しいです。

# [checkLeapYear_1](https://github.com/window794/checkLeapYear/blob/main/checkLeapYear_1.vba)
国立天文台の定義するうるう年の定義から、うるう年かどうかを判定するプログラムです。<br>
[国立天文台](https://www.nao.ac.jp/faq/a0306.html)曰く
```
グレゴリオ暦法では、うるう年を次のように決めています。

（1）西暦年号が4で割り切れる年をうるう年とする。

（2）（1）の例外として、西暦年号が100で割り切れて400で割り切れない年は平年とする。
```
とのこと。これを実直にIf文で書いてやってもいいのですが……読み替えると `うるう年は：4で割り切れる　かつ　100で割り切れない　または　400で割り切れる年`になるんじゃないか？と思いました。なので条件式は` (determine_year Mod 4 = 0) And ((determine_year Mod 100 <> 0) Or (determine_year Mod 400) = 0) `としています。<br>
## 補足
If文の条件式はカッコで括って書くこともできます。カッコを省略して書くことももちろんできます。ただ、可読性（読み取れる度合い）の観点からいえば、カッコで括ったほうがいいのかな？と思いはじめてきました。
### If文の条件式をカッコで括った例
```vb
If (file_year > this_year) Then
    MsgBox "ファイル内に記載のある年は未来です。"
End If
```

# [checkLeaypYear_2](https://github.com/window794/checkLeapYear/blob/main/checkLeapYear_2.vba)
うるう年判定の考え方は何も国立天文台の定義のみではありません。あくまで定義は定義であって、使えとは誰も言っていないから……。<br>
うるう年といえば、2月は29日まである！そこを活用して、その年がうるう年かどうかを判定するプログラムです。<br>
[DateSerial関数](https://excel-ubara.com/excelvba8/EXCELVBA843.html)を活用して、2月の末日を取得します。そこから[Day関数](https://excel-ubara.com/excelvba8/EXCELVBA847.html)で日付のみを取得します。流れは以下のとおりです。<br>
1. DateSerial関数で、Dayに0を指定すると前月の末日が返ってくる性質を活用し、`DateSerial(determine_year, 3, 0)`とする。<br>
前月の末日が返ってくるので、Monthに指定するのは末日を出したい月の翌月の数字。またDateSerial関数で返ってくるのは`yyyy/mm/dd`の日付
2. Day関数を使って、DateSerial関数で帰ってきたyyyy/mm/ddの末日から、日付のみを抽出する
3. もし日付が29だったら、その年はうるう年ということになる

# [checkLeapYear_3](https://github.com/window794/checkLeapYear/blob/main/checkLeapYear_3.vba)
うるう年では2月は29日まである！ということは一年は366日である！というのを活用して、その年がうるう年かどうかを判定するプログラムです。<br>
[DatePart関数](https://excel-ubara.com/excelvba8/EXCELVBA842.html)を活用し、年始めから年末まで何日あるかを出します。これが366であればうるう年であると判定します。流れは以下のとおりです。<br>
1. 判定年の12/31を格納する変数`last_date`を宣言
2. `last_date`に、当年の末日（`determine_year & 12/31`）を格納
3. DatePart関数で年始めから年末までの日数を計算し、366日であればうるう年であるということになる
