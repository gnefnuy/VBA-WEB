' ExcelHome推荐方法一
Function Py$(ByVal rng$)
    Dim i%, pyArr, str$, ch$
    pyArr = [{"吖","A";"八","B";"攃","C";"咑","D";"妸","E";"发","F";"旮","G";"哈","H";"丌","J";"咔","K";"垃","L";"妈","M";"乸","N";"噢","O";"帊","P";"七","Q";"冄","R";"仨","S";"他","T";"屲","W";"夕","X";"丫","Y";"帀","Z"}]
    str = Replace(Replace(rng, " ", ""), "　", "")          '去空格和Tab
    For i = 1 To Len(str)
        ch = Mid(str, i, 1)
        If ch Like "[一-龥]" Then   '如果是汉字，进行转换
            Py = Py & WorksheetFunction.Lookup(Mid(str, i, 1), pyArr)
        Else
            'Py = Py & UCase(ch)     '如果不是汉字，直接输出
        End If
    Next
End Function
' ExcelHome推荐方法二
'注意：本函数须配合声明中的“Option Compare Text”使用
Function Pyy$(ByVal rng$)
    Dim i%, k%, str$, ch$
    str = Replace(Replace(rng, " ", ""), "　", "")          '去空格和Tab
    For i = 1 To Len(str)
        k = 1
        ch = Mid(str, i, 1)
        If ch Like "[一-龥]" Then       '如果是汉字，进行转换
            Do Until Mid("八攃咑妸发旮哈丌丌咔垃妈乸噢帊七冄仨他屲屲屲夕丫帀咗", k, 1) > ch
                k = k + 1
            Loop
            Pyy = Pyy & Chr(64 + k)
        Else
            'Pyy = Pyy & UCase(ch)       '如果不是汉字，直接输出
        End If
    Next
End Function
