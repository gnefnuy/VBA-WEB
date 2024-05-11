Function isVIN(VIN As String) As Boolean
    ' 检查车架号VIN是否符合标准
    ' 参数：
    '   VIN：需要检查的车架号字符串
    ' 返回值:
    ' Boolean: 正确返回True，错误返回False
    If TypeName(VIN) <> "String" Then ' 如果不是文本，退出检查
        isVIN = False
        Exit Function
    End If
    
    If Len(Trim(VIN)) <> 17 Then ' 如果没有17位，退出检查
        isVIN = False
        Exit Function
    End If

    VIN = UCase(VIN)
    Dim RE As Object
    Set RE = CreateObject("VBScript.RegExp")
    RE.Pattern = "^[A-HJ-NPR-Z\d]{8}[X\d][A-HJ-NPR-Z\d]{3}\d{5}$"

    If Not RE.test(VIN) Then ' 如果不符合正则要求，退出检查
        isVIN = False
        Exit Function
    End If

    Dim cOT As Object
    Set cOT = CreateObject("Scripting.Dictionary")
    cOT.Add "0", 0
    cOT.Add "1", 1
    cOT.Add "2", 2
    cOT.Add "3", 3
    cOT.Add "4", 4
    cOT.Add "5", 5
    cOT.Add "6", 6
    cOT.Add "7", 7
    cOT.Add "8", 8
    cOT.Add "9", 9
    cOT.Add "A", 1
    cOT.Add "B", 2
    cOT.Add "C", 3
    cOT.Add "D", 4
    cOT.Add "E", 5
    cOT.Add "F", 6
    cOT.Add "G", 7
    cOT.Add "H", 8
    cOT.Add "J", 1
    cOT.Add "K", 2
    cOT.Add "L", 3
    cOT.Add "M", 4
    cOT.Add "N", 5
    cOT.Add "P", 7
    cOT.Add "R", 9
    cOT.Add "S", 2
    cOT.Add "T", 3
    cOT.Add "U", 4
    cOT.Add "V", 5
    cOT.Add "W", 6
    cOT.Add "X", 7
    cOT.Add "Y", 8
    cOT.Add "Z", 9

    Dim xWT As Variant
    xWT = Array(8, 7, 6, 5, 4, 3, 2, 10, 0, 9, 8, 7, 6, 5, 4, 3, 2)

    Dim sum As Long
    Dim i As Integer
    For i = 1 To 17
        sum = sum + cOT(Mid(VIN, i, 1)) * xWT(i - 1)
    Next i

    Dim cT As Variant
    cT = Array("0", "1", "2", "3", "4", "5", "6", "7", "8", "9", "X")

    isVIN = (cT(sum Mod 11) = Mid(VIN, 9, 1))

End Function

Function isPhoneNumber(PhoneNumber As String) As Boolean
    '检查手机号码是否符合标准
    Dim regIdCard As Object
    Set regIdCard = CreateObject("VBScript.RegExp")
    regIdCard.Pattern = "^1(3[0-9]|4[01456879]|5[0-35-9]|6[2567]|7[0-8]|8[0-9]|9[0-35-9])\d{8}$"
    isPhoneNumber = regIdCard.test(PhoneNumber)
End Function

Function isIdCard(idCard As String) As Boolean
    ' 检查身份证号码是否符合标准
    ' 参数：
    '   idCard：需要检查的身份证号码字符串
    ' 返回值:
    ' Boolean: 正确返回True，错误返回False
    
    ' 15位和18位身份证号码的正则表达式
    Dim regIdCard As Object
    Set regIdCard = CreateObject("VBScript.RegExp")
    regIdCard.Pattern = "^(^[1-9]\d{7}((0\d)|(1[0-2]))(([0|1|2]\d)|3[0-1])\d{3}$)|(^[1-9]\d{5}[1-9]\d{3}((0\d)|(1[0-2]))(([0|1|2]\d)|3[0-1])((\d{4})|\d{3}[Xx])$)$"

    ' 如果通过该验证，说明身份证格式正确，但准确性还需计算
    If regIdCard.test(idCard) Then

        If Len(idCard) = 18 Then

            Dim idCardWi As Variant
            idCardWi = Array(7, 9, 10, 5, 8, 4, 2, 1, 6, 3, 7, 9, 10, 5, 8, 4, 2)

            Dim idCardY As Variant
            idCardY = Array(1, 0, 10, 9, 8, 7, 6, 5, 4, 3, 2)

            Dim idCardWiSum As Long
            idCardWiSum = 0

            Dim i As Long
            For i = 0 To 16
                idCardWiSum = idCardWiSum + CLng(Mid(idCard, i + 1, 1)) * idCardWi(i)
            Next i

            Dim idCardMod As Long
            idCardMod = idCardWiSum Mod 11

            Dim idCardLast As String
            idCardLast = Right(idCard, 1)

            ' 如果等于2，则说明校验码是10，身份证号码最后一位应该是X
            If idCardMod = 2 Then

                If idCardLast = "X" Or idCardLast = "x" Then
                    isIdCard = True
                Else
                    isIdCard = False
                End If

            Else

                ' 用计算出的验证码与最后一位身份证号码匹配，如果一致，说明通过，否则是无效的身份证号码
                If idCardLast = idCardY(idCardMod) Then
                    isIdCard = True
                Else
                    isIdCard = False
                End If

            End If

        Else
            isIdCard = True
        End If

    Else
        isIdCard = False
    End If
End Function

Function isSocialCreditCode(code As String) As Boolean
    ' 检查统一社会信用代码是否符合标准
    ' 参数：
    '   code：需要检查的统一社会信用代码字符串
    ' 返回值:
    ' Boolean: 正确返回True，错误返回False
   
    ' 空值直接返回false
    If code = "" Then
        isSocialCreditCode = False
        Exit Function
    End If
    code = UCase(code)
    
    '18位及正则校验
    Dim reg As Object: Set reg = CreateObject("VBScript.RegExp")
    reg.Pattern = "^[0-9A-HJ-NPQRTUWXY]{2}\d{6}[0-9A-HJ-NPQRTUWXY]{10}$"
    If (Len(code) <> 18) Or Not reg.test(code) Then
        isSocialCreditCode = False
        Exit Function
    End If
    
    Dim codeDict  As Object
    Set codeDict = CreateObject("Scripting.Dictionary")
    codeDict.Add "0", 0
    codeDict.Add "1", 1
    codeDict.Add "2", 2
    codeDict.Add "3", 3
    codeDict.Add "4", 4
    codeDict.Add "5", 5
    codeDict.Add "6", 6
    codeDict.Add "7", 7
    codeDict.Add "8", 8
    codeDict.Add "9", 9
    codeDict.Add "A", 10
    codeDict.Add "B", 11
    codeDict.Add "C", 12
    codeDict.Add "D", 13
    codeDict.Add "E", 14
    codeDict.Add "F", 15
    codeDict.Add "G", 16
    codeDict.Add "H", 17
    codeDict.Add "J", 18
    codeDict.Add "K", 19
    codeDict.Add "L", 20
    codeDict.Add "M", 21
    codeDict.Add "N", 22
    codeDict.Add "P", 23
    codeDict.Add "Q", 24
    codeDict.Add "R", 25
    codeDict.Add "T", 26
    codeDict.Add "U", 27
    codeDict.Add "W", 28
    codeDict.Add "X", 29
    codeDict.Add "Y", 30
    
    Dim xWT As Variant
    xWT = Array(1, 3, 9, 27, 19, 26, 16, 17, 20, 29, 25, 13, 8, 24, 10, 30, 28)
    
    Dim sum As Long
    Dim i As Integer
    Dim modResult As Variant
    For i = 1 To 17
        sum = sum + codeDict(Mid(code, i, 1)) * xWT(i - 1)
    Next i
    modResult = 31 - sum Mod 31
    
    Dim cT As Variant
    cT = Array("0", "1", "2", "3", "4", "5", "6", "7", "8", "9", "A", "B", "C", "D", "E", "F", "G", "H", "J", "K", "L", "M", "N", "P", "Q", "R", "T", "U", "W", "X", "Y", "0")
    
    If cT(modResult) = Right(code, 1) Or (modResult = 31 And Right(code, 1) = "0") Then
        isSocialCreditCode = True
    Else
        isSocialCreditCode = False
    End If
End Function