' 实现百度云OCR的支持
'' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
Option Explicit

' --------------------------------------------- '
' 常量和私有变量
' --------------------------------------------- '

Private Const BASE_URL As String = "https://aip.baidubce.com/rest/2.0"
Private Const OCR_URL As String = BASE_URL & "/ocr/v1/"
Private Const IMAGE_CLASSIFY_URL As String = BASE_URL & "/image-classify/v1/"

' --------------------------------------------- '
' 属性
' --------------------------------------------- '

Public apiKey As String
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
Public Sub Setup(apiKey As String, secretKey As String)
    Me.apiKey = apiKey
    Me.secretKey = secretKey
End Sub

' 使用 AK，SK 生成鉴权签名（Access Token）,有效期30天，需加入避免重复获取
Public Function GetAccessToken() As String
    Dim token As String
    ' TODO:先获取本地数据，是否有Token，如果没有就去平台获取
    token = "******30天有效*****"
    GetAccessToken = token: Exit Function
    Dim client As New WebClient
    Dim request As New WebRequest
    Dim response As WebResponse
    
    With request
        .Resource = "https://aip.baidubce.com/oauth/2.0/token"
        .Method = WebMethod.HttpPost
        .AddQuerystringParam "grant_type", "client_credentials"
        .AddQuerystringParam "client_id", Me.apiKey
        .AddQuerystringParam "client_secret", Me.secretKey
    End With
    Set response = client.Execute(request)
    GetAccessToken = response.data("access_token")
End Function
' 通过本地行驶证图片文件识别
Public Function VehicleLicenseFile(Imgpath As String) As Dictionary
    Set VehicleLicenseFile = OCR("vehicle_license", Imgpath)
End Function
' 通过行驶证图片网络地址识别
Public Function VehicleLicenseUrl(ImgUrl As String) As Dictionary
    Set VehicleLicenseUrl = OCR("vehicle_license", ImgUrl, 2)
End Function
' 通过本地驾驶证图片文件识别
Public Function DrivingLicenseFile(Imgpath As String) As Dictionary
    Set DrivingLicenseFile = OCR("driving_license", Imgpath)
End Function
' 通过本地图片识别车牌
Public Function LicensePlateFile(Imgpath As String) As Dictionary
    Set LicensePlateFile = OCR("license_plate", Imgpath)
End Function
' 通过本地图片识别VIN
Public Function VinCodeFile(Imgpath As String) As Dictionary
    Set VinCodeFile = OCR("vin_code", Imgpath)
End Function
' 通过本地图片识别机动车销售发票
Public Function VehicleInvoiceFile(Imgpath As String) As Dictionary
    Set VehicleInvoiceFile = OCR("vehicle_invoice", Imgpath)
End Function
' 通过本地图片识别身份证
Public Function IdcardFile(Imgpath As String) As Dictionary
    Set IdcardFile = OCR("idcard", Imgpath)
End Function

' 识别器
' OcrType: 识别类别，按照供应商提供的数据来
' UriType: 0->本地文件路径（默认），2->网络地址，3->BASE64
' body: 对应的相应参数，默认为空
Private Function OCR(OcrType As String, Uri As String, Optional UriType As Integer = 0, Optional body As Dictionary = Nothing) As Dictionary
    Dim client As New WebClient
    Dim request As New WebRequest
    Dim response As WebResponse
    Dim body_ As New Dictionary
    Dim result As New Dictionary
    
    Set body_ = body
    ' 设置请求参数
    With request
        .Resource = OCR_URL & OcrType & "?access_token=" & GetAccessToken()
        .Method = WebMethod.HttpPost
        .AddHeader "Content-Type", "application/x-www-form-urlencoded"
        .AddHeader "Accept", "application/json"
        .RequestFormat = FormUrlEncoded
    End With
    
    Select Case UriType
        Case 0
            body_.Add "image", WebHelpers.Base64EncodeFile(Uri)
        Case 1
            body_.Add "url", Uri
        Case 2
            body_.Add "image", Uri
    End Select
    ' 将请求正文添加到请求中
    Set request.body = body_
    ' 发送请求并获取响应
    Set response = client.Execute(request)
   Dim aStr As String
    aStr = DecodeToBytes(response.body) '将返回的UTF8字符转为VB可显示的ANSI字符
    'aStr = WebHelpers.StringToAnsiBytes(response.body)
    Set result = WebHelpers.ParseJson(aStr)
     ' 输出响应内容
   Set OCR = result
End Function

' 行驶证识别
Private Function VehicleLicense(Uri As String, isFile As Boolean, Optional body As Dictionary = Nothing) As Dictionary
    'body.Add "detect_direction", "false" ' - false：默认值不进行图像方向自动矫正 - true: 开启图像方向自动矫正功能，可对旋转 90/180/270 度的图片进行自动矫正并识别
    'body.Add "vehicle_license_side","front" ' - front：默认值，识别行驶证主页 - back：识别行驶证副页
    'body.Add "unified", "false" ' - false：默认值，不进行归一化处理 - true：对输出字段进行归一化处理，将新/老版行驶证的“注册登记日期/注册日期”统一为”注册日期“进行输出
    'body.Add "quality_warn", "false" ' 是否开启质量检测功能，仅在行驶证正页识别时生效， - false：默认值，不输出质量告警信息 - true： 输出行驶证遮挡、不完整质量告警信息
    'body.Add "risk_warn", "false" ' 是否开启风险检测功能， - false：默认值，不输出风险告警信息 - true：开启，输出行驶证复印、翻拍、PS等告警信息
End Function
' 驾驶证识别
Private Function DrivingLicense(Uri As String, isFile As Boolean, Optional body As Dictionary = Nothing) As Dictionary
    'body.Add "detect_direction", "false" ' - false：默认值，不检测朝向，朝向是指输入图像是正常方向、逆时针旋转90/180/270度 - true：检测朝向
    'body.Add "driving_license_side", "front" ' - front：默认值，识别驾驶证正页 - back：识别驾驶证副页
    'body.Add "unified_valid_period", "false" '- false: 默认值，不进行归一化处理 - true: 归一化格式输出驾驶证的「有效起始日期」+「有效期限」及「有效期限」+「至」两种输出格式归一化为「有效起始日期」+「失效日期」
    'body.Add "quality_warn", "false" ' 是否开启质量检测功能，仅在驾驶证正页识别时生效， - false：默认值，不输出质量告警信息 - true： 输出驾驶证遮挡、不完整质量告警信息
    'body.Add "risk_warn", "false" ' 是否开启风险检测功能， - false：默认值，不输出风险告警信息 - true：开启，输出驾驶证复印、翻拍、PS等告警信息
End Function