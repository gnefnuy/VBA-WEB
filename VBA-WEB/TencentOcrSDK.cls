' 实现腾讯云
' tencentOcrSdk.cls
'' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
Option Explicit
' --------------------------------------------- '
' 属性
' --------------------------------------------- '
Public secretId As String
Public secretKey As String

' ============================================= '
' 公共方法
' ============================================= '

''
' Setup
'
' @param {String} SecretId 用于标识 API 调用者身份，可以简单类比为用户名。
' @param {String} SecretKey 用于验证 API 调用者的身份，可以简单类比为密码。
''
Public Sub Setup(secretId As String, secretKey As String)
    Me.secretId = secretId
    Me.secretKey = secretKey
End Sub
'签名方法1.0
Public Function GetAuth(host As String, body As Dictionary) As String
    Dim segs As New Dictionary
    Dim key As Variant
    Dim Value As Variant
    Dim formattedParams() As String
    
    Set segs = body
    ' 对参数名称进行排序
    Set segs = SortDictionaryKeys(segs)
    
    ' 初始化格式化后的参数数组
    ReDim formattedParams(segs.Count - 1)
    
    ' 遍历Dictionary中的每个键值对
    Dim Index As Integer
    For Each key In segs.Keys
        Value = segs(key)
        
        ' 将键和值拼接成"参数名称=参数值"的形式，并将结果添加到数组中
        formattedParams(Index) = key & "=" & Value
        Index = Index + 1
    Next key
    
    ' 使用Join函数将数组中的元素连接成一个字符串，并用"&"分隔
    Dim result As String
    result = Join(formattedParams, "&")
    
    result = "POST" & host & "/?" & result '拼接签名原文字符串
    Debug.Print result
    Dim signStr As String
    signStr = WebHelpers.HmacSha1(result, Me.secretKey, "Base64")
    Debug.Print signStr
    'signStr = WebHelpers.UrlEncode(signStr)
    'Debug.Print signStr
    GetAuth = signStr
End Function

' 签名
Public Function sign(key, msg) As String
    sign = WebHelpers.HmacSha256(msg, key, "Base64")
End Function
' 腾讯签名方法3.0
Public Function GetAuth3(secretId As String, secretKey As String, host As String, contentType As String, timestamp As String, tc_action As String, body As Dictionary) As String
    Dim canonicalURI As String
    Dim canonicalHeaders As String
    Dim signedHeaders As String
    Dim hashedRequestPayload As String
    Dim canonicalRequest As String
    Dim algorithm As String
    Dim date_s As String
    Dim service As String
    Dim credentialScope As String
    Dim hashedCanonicalRequest As String
    Dim stringToSign As String
    Dim tc3SecretKey() As Byte
    Dim secretDate() As Byte
    Dim secretService() As Byte
    Dim secretSigning() As Byte
    Dim signatureBytes() As Byte
    Dim Signature As String
    
    canonicalURI = "/"
    canonicalHeaders = "content-type:" + contentType + vbLf + "host:" + host + vbLf + "x-tc-action:" + LCase(tc_action) + vbLf ' 此处有坑，必须小写
    signedHeaders = "content-type;host;x-tc-action"
    hashedRequestPayload = WebHelpers.Sha256(WebHelpers.ConvertToJson(body)) 'Sha256Hex(body)
    canonicalRequest = "POST" + vbLf + canonicalURI + vbLf + vbLf + canonicalHeaders + vbLf + signedHeaders + vbLf + hashedRequestPayload
    'Debug.Print canonicalRequest ' 拼接规范请求串, 正确
    
    algorithm = "TC3-HMAC-SHA256"
    date_s = Format(WebHelpers.ConvertToUtc(Now()), "yyyy-MM-dd") 'Format$(DateAdd("s", CLng(timestamp), #1/1/1970#), "yyyy-MM-dd")
    service = Split(host, ".")(0)
    credentialScope = date_s + "/" + service + "/" + "tc3_request"
    hashedCanonicalRequest = WebHelpers.Sha256(canonicalRequest) 'Sha256Hex(canonicalRequest)
    stringToSign = algorithm + vbLf + timestamp + vbLf + credentialScope + vbLf + hashedCanonicalRequest
    'Debug.Print stringToSign ' 拼接待签名字符串,正确
    
    tc3SecretKey = VBA.StrConv("TC3" + secretKey, vbFromUnicode)
    secretDate = WebHelpers.HmacSha256_Byte(VBA.StrConv(date_s, vbFromUnicode), tc3SecretKey)
    secretService = WebHelpers.HmacSha256_Byte(VBA.StrConv(service, vbFromUnicode), secretDate)
    secretSigning = WebHelpers.HmacSha256_Byte(VBA.StrConv("tc3_request", vbFromUnicode), secretService)
    signatureBytes = WebHelpers.HmacSha256_Byte(VBA.StrConv(stringToSign, vbFromUnicode), secretSigning)
    Signature = WebHelpers.web_AnsiBytesToHex(signatureBytes)
    'Debug.Print Signature ' 计算签名, 正确
    
    '拼接 Authorization
    GetAuth3 = algorithm + " " + _
              "Credential=" + secretId + "/" + credentialScope + ", " + _
              "SignedHeaders=" + signedHeaders + ", " + _
              "Signature=" + Signature
End Function

' 通过本地行驶证图片文件识别
Public Function VehicleLicenseFile(Imgpath As String) As Dictionary
    Set VehicleLicenseFile = OCR("VehicleLicenseOCR", Imgpath)
End Function


' 通过本地车牌图片文件识别
Public Function LicensePlateFile(Imgpath As String) As Dictionary
    Set LicensePlateFile = OCR("LicensePlateOCR", Imgpath)
End Function

' 识别器
' OcrType: 识别类别，按照供应商提供的数据来
' UriType: 0->本地文件路径（默认），2->网络地址，3->BASE64
' body: 对应的相应参数，默认为空
Private Function OCR(action As String, Uri As String, Optional UriType As Integer = 0, Optional body As Dictionary = Nothing) As Dictionary
    Dim client As New WebClient
    Dim request As New WebRequest
    Dim response As WebResponse
    Dim body_ As New Dictionary
    Dim result As New Dictionary
    Dim host As String
    Dim url As String
    Dim contentType As String
    Dim Authorization As String 'v3版本
    Dim Signature As String ' v1版本
    Dim timestamp As String
    
    host = "ocr.tencentcloudapi.com"
    'host = "httpbin.org"
    url = "https://" & host & "/"
    contentType = "application/json; charset=utf-8"
    timestamp = WebHelpers.GetTimestamp
    'timestamp = "1715143697"
    
    Set body_ = body
    '-------v1--------
    'body_.Add "Action", action
    'body_.Add "Language", "zh-CN"
    'body_.Add "Nonce", Int(Rnd * 10000) ' 随机正整数
    'body_.Add "Region", "ap-shanghai"
    'body_.Add "SecretId", Me.secretId '
    'body_.Add "Timestamp", timestamp
    'body_.Add "Version", "2018-11-19"
    '-------------
    Select Case UriType
        Case 0
            body_.Add "ImageBase64", WebHelpers.Base64EncodeFile(Uri)
            'body_.Add "ImageBase64", "abab"
        Case 1
            body_.Add "ImageUrl", Uri
        Case 2
            body_.Add "ImageBase64", Uri
    End Select
    '生成Auth v3
    Authorization = Me.GetAuth3(Me.secretId, Me.secretKey, host, contentType, timestamp, action, body_)
    'Debug.Print Authorization
    '---生成签名 v1
    'Signature = Me.GetAuth(host, body_)
    'Debug.Print Signature
    'body_.Add "Signature", Signature
    ' 设置请求参数
    With request
        .Resource = url
        .Method = WebMethod.HttpPost
        '--------v1-----------
        '.RequestFormat = FormUrlEncoded
        '--------v3-----------
        .AddHeader "Host", host
        .AddHeader "X-TC-Timestamp", timestamp
        .AddHeader "X-TC-Version", "2018-11-19"
        .AddHeader "X-TC-Action", action
        .AddHeader "X-TC-Region", "ap-shanghai"
        .AddHeader "X-TC-Language", "zh-CN"
        '.AddHeader "X-TC-Token", ""
        '.AddHeader "X-TC-RequestClient", ""
        .AddHeader "Content-Type", contentType
        .AddHeader "Authorization", Authorization '
        '.RequestFormat = Json
        '----------------------
    End With
    ' 将请求正文添加到请求中
    Set request.body = body_
    
    ' 发送请求并获取响应
    Set response = client.Execute(request)
    
   Dim aStr As String
    aStr = DecodeToBytes(response.body) '将返回的UTF8字符转为VB可显示的ANSI字符
    Set result = WebHelpers.ParseJson(aStr)
     ' 输出响应内容
   Set OCR = result
End Function