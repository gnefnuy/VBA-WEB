Option Explicit
'***********************************************************************************************
'   * 请保留任何商标或版权信息。
'   *
'   * 致谢贡献者：
'   *       STEPHEN BULLEN, 1998年11月15日 - 原始PastPicture代码
'   *       G HUDSON, 2010年4月5日 - 暂停功能
'   *       LUTZ GENTKOW, 2011年7月23日 - Alt + PrtScrn
'   *       PAUL FRANCIS, 2013年4月11日 - 将所有部分整合在一起，桥接32位和64位版本。
'   *       CHRIS O, 2013年4月12日 - 代码建议适用于早期版本的Access。
'   *
'   * 描述：从剪贴板上创建一个标准的图片对象。
'   *              然后将此对象保存到磁盘上的某个位置。请注意，这也可以分配给（例如）用户表单上的图像控件。
'   *
'   * 该代码需要引用“OLE自动化”类型库。
'   *
'   * 本模块中的代码源自MSDN、Access世界论坛、VBForums上发现的多个来源。
'   *
'   * 要使用它，只需将此模块复制到您的项目中，然后您可以使用：
'   * SaveClip2Bit("C:\Pics\Sample.bmp") 或 SaveClip2Bit("D:\TEST.JPG")
'   * 将其保存到磁盘上的某个位置。
'   * （或者）
'   * Set ImageControl.Image = PastePicture
'   * 将剪贴板上的任何内容的图片粘贴到标准图像控件中。
'   *
'   * 过程：
'   *   PastePicture  :   设置图像的入口点
'   *   CreatePicture :   将位图或元文件句柄转换为OLE引用的私有函数
'   *   fnOLEError    :   获取OLE错误代码的错误文本
'   *   SaveClip2Bit  :   保存图像的入口点，调用PastePicture
'   *   AltPrintScreen:   执行Alt + PrtScrn的自动化，用于获取活动窗口。
'   *   Pause         :   使程序等待，以确保进行适当的屏幕捕获。
'**************************************************************************************************
 
'声明一个UDT来存储IPicture OLE接口的GUID
Private Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type
 
'声明一个UDT来存储位图信息
Private Type uPicDesc
    Size As Long
    type As Long
    hPic As Long
    hpal As Long
End Type
 
'Windows API Function Declarations
'Does the clipboard contain a bitmap/metafile?剪贴板中是否包含位图/元文件？
Private Declare PtrSafe Function IsClipboardFormatAvailable Lib "user32" (ByVal wFormat As Integer) As Long
'Open the clipboard to read打开剪贴板以进行读取
Private Declare PtrSafe Function OpenClipboard Lib "user32" (ByVal hwnd As Long) As Long
'Get a pointer to the bitmap/metafile获取指向位图/元文件的指针
Private Declare PtrSafe Function GetClipboardData Lib "user32" (ByVal wFormat As Integer) As Long
'Close the clipboard关闭剪贴板
Private Declare PtrSafe Function CloseClipboard Lib "user32" () As Long
'Convert the handle into an OLE IPicture interface.将句柄转换为OLE IPicture接口。
Private Declare PtrSafe Function OleCreatePictureIndirect Lib "oleaut32.dll" (PicDesc As uPicDesc, RefIID As GUID, ByVal fPictureOwnsHandle As Long, IPic As IPicture) As Long
 'Create our own copy of the metafile, so it doesn't get wiped out by subsequent clipboard updates.创建我们自己的元文件副本，以便它不会因后续剪贴板更新而被清除。
Declare PtrSafe Function CopyEnhMetaFile Lib "gdi32" Alias "CopyEnhMetaFileA" (ByVal hemfSrc As Long, ByVal lpszFile As String) As Long
'Create our own copy of the bitmap, so it doesn't get wiped out by subsequent clipboard updates.创建我们自己的位图副本，以便它不会因后续剪贴板更新而被清除。
Declare PtrSafe Function CopyImage Lib "user32" (ByVal handle As Long, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long
'Uses the Keyboard simulation使用键盘模拟
Private Declare PtrSafe Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
 
   
'The API format types we're interested in
Const CF_BITMAP = 2
Const CF_PALETTE = 9
Const CF_ENHMETAFILE = 14
Const IMAGE_BITMAP = 0
Const LR_COPYRETURNORG = &H4
 
Private Const KEYEVENTF_KEYUP = &H2
Private Const VK_SNAPSHOT = &H2C
Private Const VK_MENU = &H12
 
' 子程序    : AltPrintScreen
' 目的       : 捕获活动窗口，并将其放置在剪贴板上。
 
Sub AltPrintScreen()
    keybd_event VK_MENU, 0, 0, 0
    keybd_event VK_SNAPSHOT, 0, 0, 0
    keybd_event VK_SNAPSHOT, 0, KEYEVENTF_KEYUP, 0
    keybd_event VK_MENU, 0, KEYEVENTF_KEYUP, 0
End Sub
 
' 子程序 : PastePicture
' 目的 : 获取一个显示剪贴板上内容的图片对象。
Function PastePicture() As IPicture
    '一些指针
    Dim h As Long, hPtr As Long, hpal As Long, lPicType As Long, hCopy As Long
 
    '检查剪贴板是否包含所需的格式
    If IsClipboardFormatAvailable(CF_BITMAP) Then
        '获取对剪贴板的访问权限
        h = OpenClipboard(0&)
        If h > 0 Then
            '获取图像数据的句柄
            hPtr = GetClipboardData(CF_BITMAP)
 
            hCopy = CopyImage(hPtr, IMAGE_BITMAP, 0, 0, LR_COPYRETURNORG)
 
            '释放剪贴板以供其他程序使用
            h = CloseClipboard
            '如果我们获得了图像的句柄，将其转换为图片对象并返回
            If hPtr <> 0 Then Set PastePicture = CreatePicture(hCopy, 0, CF_BITMAP)
        End If
    End If
End Function
 
 
' 子程序 : CreatePicture
' 目的 : 将图像（和调色板）句柄转换为图片对象。
' 注意 : 需要引用“OLE自动化”类型库

Private Function CreatePicture(ByVal hPic As Long, ByVal hpal As Long, ByVal lPicType) As IPicture
    ' IPicture 需要引用 "OLE Automation"
    Dim r As Long, uPicInfo As uPicDesc, IID_IDispatch As GUID, IPic As IPicture
    'OLE 图片类型
    Const PICTYPE_BITMAP = 1
    Const PICTYPE_ENHMETAFILE = 4
    '  创建接口 GUID（用于 IPicture 接口）
    With IID_IDispatch
        .Data1 = &H7BF80980
        .Data2 = &HBF32
        .Data3 = &H101A
        .Data4(0) = &H8B
        .Data4(1) = &HBB
        .Data4(2) = &H0
        .Data4(3) = &HAA
        .Data4(4) = &H0
        .Data4(5) = &H30
        .Data4(6) = &HC
        .Data4(7) = &HAB
    End With
    '  用必要的部分填充 uPicInfo。
 
    With uPicInfo
        .Size = Len(uPicInfo) ' 结构的长度。
        .type = PICTYPE_BITMAP ' 图片类型
        .hPic = hPic ' 图像的句柄。
        .hpal = hpal ' 调色板的句柄（如果是位图）。
    End With
 
    ' 创建图片对象。
    r = OleCreatePictureIndirect(uPicInfo, IID_IDispatch, True, IPic)
 
    ' 如果发生错误，显示描述
    If r <> 0 Then Debug.Print "Create Picture: " & fnOLEError(r)
 
    ' 返回新的图片对象
    Set CreatePicture = IPic
End Function
 
 
' Subroutine    : fnOLEError
' Purpose       : Gets the message text for standard OLE errors
 
Private Function fnOLEError(lErrNum As Long) As String
    'OLECreatePictureIndirect return values
    Const E_ABORT = &H80004004
    Const E_ACCESSDENIED = &H80070005
    Const E_FAIL = &H80004005
    Const E_HANDLE = &H80070006
    Const E_INVALIDARG = &H80070057
    Const E_NOINTERFACE = &H80004002
    Const E_NOTIMPL = &H80004001
    Const E_OUTOFMEMORY = &H8007000E
    Const E_POINTER = &H80004003
    Const E_UNEXPECTED = &H8000FFFF
    Const S_OK = &H0
 
    Select Case lErrNum
        Case E_ABORT
            fnOLEError = " Aborted"
        Case E_ACCESSDENIED
            fnOLEError = " Access Denied"
        Case E_FAIL
            fnOLEError = " General Failure"
        Case E_HANDLE
            fnOLEError = " Bad/Missing Handle"
        Case E_INVALIDARG
            fnOLEError = " Invalid Argument"
        Case E_NOINTERFACE
            fnOLEError = " No Interface"
        Case E_NOTIMPL
            fnOLEError = " Not Implemented"
        Case E_OUTOFMEMORY
            fnOLEError = " Out of Memory"
        Case E_POINTER
            fnOLEError = " Invalid Pointer"
        Case E_UNEXPECTED
            fnOLEError = " Unknown Error"
        Case S_OK
            fnOLEError = " Success!"
    End Select
End Function
 
' Routine   : SaveClip2Bit
' Purpose   : Saves Picture object to desired location.
' Arguments : Path to save the file
 
Public Sub SaveClip2Bit(savePath As String)
On Error GoTo errHandler:
    AltPrintScreen
    Pause (3)
    SavePicture PastePicture, savePath
errExit:
        Exit Sub
errHandler:
    Debug.Print "Save Picture: (" & Err.Number & ") - " & Err.description
    Resume errExit
End Sub
 
' Routine   : Pause
' Purpose   : Gives a short interval for proper image capture.
' Arguments : Seconds to wait.
 
Public Function Pause(NumberOfSeconds As Variant)
On Error GoTo Err_Pause
    Dim PauseTime As Variant, start As Variant
    PauseTime = NumberOfSeconds
    start = Timer
    Do While Timer < start + PauseTime
        DoEvents
    Loop
Exit_Pause:
    Exit Function
Err_Pause:
    MsgBox Err.Number & " - " & Err.description, vbCritical, "Pause()"
    Resume Exit_Pause
End Function