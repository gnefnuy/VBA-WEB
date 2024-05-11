' API 声明
Private Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type

Private Type GdiplusStartupInput
    GdiplusVersion As Long
    DebugEventCallback As LongPtr
    SuppressBackgroundThread As Long
    SuppressExternalCodecs As Long
End Type

Private Type EncoderParameter
    GUID As GUID
    NumberOfValues As Long
    type As Long
    Value As LongPtr
End Type

Private Type EncoderParameters
    Count As Long
    Parameter As EncoderParameter
End Type

Private Declare PtrSafe Function GdiplusStartup Lib "GDIPlus" ( _
    token As LongPtr, _
    inputbuf As GdiplusStartupInput, _
    Optional ByVal outputbuf As LongPtr = 0) As Long

Private Declare PtrSafe Function GdiplusShutdown Lib "GDIPlus" ( _
    ByVal token As LongPtr) As Long

Private Declare PtrSafe Function GdipCreateBitmapFromHBITMAP Lib "GDIPlus" ( _
    ByVal hbm As LongPtr, _
    ByVal hpal As LongPtr, _
    Bitmap As LongPtr) As Long

Private Declare PtrSafe Function GdipDisposeImage Lib "GDIPlus" ( _
    ByVal Image As LongPtr) As Long

Private Declare PtrSafe Function GdipSaveImageToFile Lib "GDIPlus" ( _
    ByVal Image As LongPtr, _
    ByVal filename As LongPtr, _
    clsidEncoder As GUID, _
    encoderParams As Any) As Long

Private Declare PtrSafe Function CLSIDFromString Lib "ole32" ( _
    ByVal str As LongPtr, _
    id As GUID) As Long

' SaveJPG
Public Sub SaveJPG( _
    ByVal pict As StdPicture, _
    ByVal filename As String, _
    Optional ByVal quality As Byte = 80)

    Dim tSI As GdiplusStartupInput
    Dim lRes As Long
    Dim lGDIP As LongPtr
    Dim lBitmap As LongPtr

    ' 初始化 GDI+
    tSI.GdiplusVersion = 1
    lRes = GdiplusStartup(lGDIP, tSI)

    If lRes = 0 Then
        ' 从图片句柄创建 GDI+ 位图
        lRes = GdipCreateBitmapFromHBITMAP(pict.handle, 0, lBitmap)

        If lRes = 0 Then
            Dim tJpgEncoder As GUID
            Dim tParams As EncoderParameters

            ' 初始化编码器 GUID
            CLSIDFromString StrPtr("{557CF401-1A04-11D3-9A73-0000F81EF32E}"), _
                              tJpgEncoder

            ' 初始化编码器参数
            tParams.Count = 1
            With tParams.Parameter     ' 质量
                ' 设置质量 GUID
                CLSIDFromString StrPtr("{1D5BE4B5-FA4A-452D-9CDD-5DB3505E7EB}"), .GUID
                .NumberOfValues = 1
                .type = 1
                .Value = VarPtr(quality)
            End With

            ' 保存图像
            lRes = GdipSaveImageToFile( _
                          lBitmap, _
                          StrPtr(filename), _
                          tJpgEncoder, _
                          tParams)

            ' 销毁位图
            GdipDisposeImage lBitmap

        End If

        ' 关闭 GDI+
        GdiplusShutdown lGDIP

    End If

    If lRes Then
        Err.Raise 5, , "无法保存图像。GDI+ 错误:" & lRes
    End If

End Sub

