VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsYPNCommonVB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'---------------------------------------------------------------------------------------
' Module    : ClsCommon
' Author    : YPN
' Date      : 2017-06-28 12:06
' Purpose   : 公共类
'---------------------------------------------------------------------------------------
Option Explicit


Public Sub Help(Optional ByVal i_Procedure As String, Optional ByVal i_IsMsg As Boolean = False)
    
    Dim v_Help As String
    
    v_Help = "请传入过程名，显示详细描述！若i_IsMsg为True，则弹出详细描述，否则会在立即窗口打印详细描述。"
    
    Select Case i_Procedure
        
    Case "ApendText"
        v_Help = ""
        v_Help = v_Help & vbCrLf & "过程：ApendText      追加内容到指定文件"
        v_Help = v_Help & vbCrLf & "参数：i_TextFile     指定文件"
        v_Help = v_Help & vbCrLf & "参数：i_ApendContent 追加内容"
        
    Case "GetFileNameInPath"
        v_Help = ""
        v_Help = v_Help & vbCrLf & "过程：GetFileNameInPath 从指定全路径中获取文件名"
        v_Help = v_Help & vbCrLf & "参数：i_Path            指定全路径"
        v_Help = v_Help & vbCrLf & "参数：i_NeedSuffix      文件名是否需要后缀名，默认不需要"
        
    Case "GetGUID"
        v_Help = ""
        v_Help = v_Help & vbCrLf & "过程：GetGUID         生成一个GUID"
        v_Help = v_Help & vbCrLf & "参数：i_Format (可选) 格式：'B'、'D'"
        
    Case "GetSuffixInFileName"
        v_Help = ""
        v_Help = v_Help & vbCrLf & "过程：GetSuffixInFileName 从文件名中获取后缀名"
        v_Help = v_Help & vbCrLf & "参数：i_FileName          文件名"
        
    Case "GetHardDriveSerialNumber"
        v_Help = ""
        v_Help = v_Help & vbCrLf & "过程：GetHardDriveSerialNumber 取得逻辑盘序列号（非唯一）"
        v_Help = v_Help & vbCrLf & "参数：i_DiskName               目标逻辑盘名称"
        
    Case "GetHardDiskSerialNumber"
        v_Help = ""
        v_Help = v_Help & vbCrLf & "过程：GetHardDiskSerialNumber 获取物理盘的系列号SN（唯一）"
        
    Case "GetHardDiskModel"
        v_Help = ""
        v_Help = v_Help & vbCrLf & "过程：GetHardDiskModel 获取物理盘的型号"
        
    Case "GetInitialFirst"
        v_Help = ""
        v_Help = v_Help & vbCrLf & "过程：GetInitialFirst 获取第一个汉字的首字母"
        v_Help = v_Help & vbCrLf & "参数：i_Str           汉字字符串"
        
    Case "GetInitialAll"
        v_Help = ""
        v_Help = v_Help & vbCrLf & "过程：GetInitialAll 获取所有汉字的首字母"
        v_Help = v_Help & vbCrLf & "参数：i_Str         汉字字符串"
        
    Case "GetYear2"
        v_Help = ""
        v_Help = v_Help & vbCrLf & "过程：GetYear2 获取日期中的年份后2位"
        v_Help = v_Help & vbCrLf & "参数：i_Date   日期"
        
    Case "HexToText"
        v_Help = ""
        v_Help = v_Help & vbCrLf & "过程：HexToText 将16进制编码串转换为文本。没有写异常处理，但只要是用 TextToHex() 转换的结果就没问题"
        v_Help = v_Help & vbCrLf & "参数：i_Code    要转换的16进制编码"
        
    Case "IsNull"
        v_Help = ""
        v_Help = v_Help & vbCrLf & "过程：IsNull 判断变量是否为空"
        v_Help = v_Help & vbCrLf & "参数：i_Var  变量"
    
    Case "JSONAnalyze"
        v_Help = ""
        v_Help = v_Help & vbCrLf & "过程：JSONAnalyze  JSON解析"
        v_Help = v_Help & vbCrLf & "参数：i_JSONString 待解析的JSON字符串"
        v_Help = v_Help & vbCrLf & "参数：i_JSONKey    解析的关键字 例如：body.data.username"
        
    Case "MD5"
        v_Help = ""
        v_Help = v_Help & vbCrLf & "过程：MD5   MD5加密"
        v_Help = v_Help & vbCrLf & "参数：i_Str 加密内容"
        v_Help = v_Help & vbCrLf & "参数：i_Bit 加密位数：16位或32位。默认16位"
        
    Case "QRCode"
        v_Help = ""
        v_Help = v_Help & vbCrLf & "过程：QRCode      生成QR Code码制的二维码"
        v_Help = v_Help & vbCrLf & "参数：i_QRText    二维码内容"
        v_Help = v_Help & vbCrLf & "参数：i_Version  （可选）生成版本，支持40种，从1到40，默认自动，即0"
        v_Help = v_Help & vbCrLf & "参数：i_ECLevel  （可选）容错等级，支持4种：L-7%、M-15%、Q-25%、H-30%，默认M（传入首字母L、M、Q、H即可）"
        v_Help = v_Help & vbCrLf & "参数：i_MaskType （可选）模糊类型，支持8种，从0到7，默认自动，即-1"
        v_Help = v_Help & vbCrLf & "参数：i_Encoding （可选）字符编码，支持2种：UTF-8 和 ANSI，默认UTF-8"
    
    Case "RequestREST"
        v_Help = ""
        v_Help = v_Help & vbCrLf & "过程：RequestREST        请求/调用REST接口"
        v_Help = v_Help & vbCrLf & "参数：i_RequstURL        请求地址 例如：http://218.21.3.20:5076/brp/services/avplan/aAUserData/queryAAUserData"
        v_Help = v_Help & vbCrLf & "参数：i_RequestParameter 请求参数 例如：{""sysid"": ""SYS_LogWeb"",""sidv"": ""1.0"",""body"": {""innerid"": ""568dd7cc1ba68779fe295fb9ebe3288c"",""userid"": ""czd"",""userno"": ""50271""}}"
        
    Case "SetStyleToolBar"
        v_Help = ""
        v_Help = v_Help & vbCrLf & "过程：SetStyleToolBar       设置控件ToolBar的样式"
        v_Help = v_Help & vbCrLf & "参数：i_Color       （可选）指定背景颜色，默认RGB(188, 210, 233)"
        v_Help = v_Help & vbCrLf & "参数：i_PicturePath （可选）指定背景图片"
        
    Case "TextToHex"
        v_Help = ""
        v_Help = v_Help & vbCrLf & "过程：TextToHex 将文本转换为16进制编码串"
        v_Help = v_Help & vbCrLf & "参数：i_Text    要转换的文本"
    End Select
    
    If i_IsMsg Then
        MsgBox v_Help, vbInformation, TS
    Else
        Debug.Print v_Help
    End If
    
End Sub

'---------------------------------------------------------------------------------------
' Procedure : ApendText
' Author    : YPN
' Date      : 2017-12-12 16:20
' Purpose   : 追加内容到指定文件
' Param     : i_TextFile     指定文件
'             i_ApendContent 追加内容
' Return    :
' Remark    :
'---------------------------------------------------------------------------------------
'
Public Sub ApendText(ByVal i_TextFile As String, ByVal i_ApendContent As String)

    Call MApendText(i_TextFile, i_ApendContent)
    
End Sub
'---------------------------------------------------------------------------------------
' Procedure : BarCode128
' Author    : YPN
' Date      : 2017-07-18 16:11
' Purpose   : 生成Code128标准的条形码
' Param     :
' Return    :
' Remark    : 调用实例：Set Picture1 = BarCode128(Text1, 6, True)
'---------------------------------------------------------------------------------------
'
Public Function BarCode128(i_BarText As String, i_BarHeight As Integer, Optional ByVal i_HasCaption As Boolean = False) As StdPicture

    Set BarCode128 = MBarCode128(i_BarText, i_BarHeight, i_HasCaption)
    
End Function

'---------------------------------------------------------------------------------------
' Procedure : BarCode39
' Author    : YPN
' Date      : 2017-07-18 16:12
' Purpose   : 生成Code39标准的条形码
' Param     :
' Return    :
' Remark    :
'---------------------------------------------------------------------------------------
'
Public Function BarCode39(i_BarText As String, i_BarHeight As Integer, Optional i_WithCheckSum As Boolean = False, Optional ByVal i_HasCaption As Boolean = False) As StdPicture
    
    Set BarCode39 = MBarCode39(i_BarText, i_BarHeight, i_WithCheckSum, i_HasCaption)
    
End Function

'---------------------------------------------------------------------------------------
' Procedure : BarEAN13
' Author    : YPN
' Date      : 2017-07-18 16:12
' Purpose   : 生成EAN13标准的条形码
' Param     :
' Return    :
' Remark    :
'---------------------------------------------------------------------------------------
'
Public Function BarEAN13(i_BarText As String, i_BarHeight As Integer, Optional ByVal i_HasCaption As Boolean = False) As StdPicture
    
    Set BarEAN13 = MBarEAN13(i_BarText, i_BarHeight, i_HasCaption)
    
End Function

'---------------------------------------------------------------------------------------
' Procedure : BarEAN8
' Author    : YPN
' Date      : 2017-07-18 16:12
' Purpose   : 生成EAN8标准的条形码
' Param     :
' Return    :
' Remark    :
'---------------------------------------------------------------------------------------
'
Public Function BarEAN8(i_BarText As String, i_BarHeight As Integer, Optional ByVal i_HasCaption As Boolean = False) As StdPicture
    
    Set BarEAN8 = MBarEAN8(i_BarText, i_BarHeight, i_HasCaption)
    
End Function

'---------------------------------------------------------------------------------------
' Procedure : Base64ToText
' Author    : YPN
' Date      : 2017-12-13 11:29
' Purpose   : 将Base64位编码字符串转换为文本
' Param     : i_Text 要转换的文本
' Return    :
' Remark    : Base64 解码
'---------------------------------------------------------------------------------------
'
Public Function Base64ToText(ByVal i_Base64 As String) As String

    Base64ToText = MBase64ToText(i_Base64)
    
End Function

'---------------------------------------------------------------------------------------
' Procedure : GetFileNameInPath
' Author    : YPN
' Date      : 2017-06-28 17:45
' Purpose   : 从指定全路径中获取文件名
' Param     : i_Path       指定全路径
'             i_NeedSuffix 文件名是否需要后缀名，默认不需要
' Return    :
' Remark    :
'---------------------------------------------------------------------------------------
'
Public Function GetFileNameInPath(ByVal i_Path As String, Optional ByVal i_NeedSuffix As Boolean = False) As String
    
    GetFileNameInPath = MGetFileNameInPath(i_Path, i_NeedSuffix)
    
End Function

'---------------------------------------------------------------------------------------
' Procedure : GetGUID
' Author    : YPN
' Date      : 2017-07-12 17:24
' Purpose   : 生成一个GUID
' Param     : i_Format(可选项) 格式："B"、"D"
' Return    : 默认为GetGUID(): bf8b9c642ea5426f82499bb60681671b
'             GetGUID("B"): {903c1236-fe24-43c2-b9b5-bec35d9a43a8}
'             GetGUID("D"): 17e316f4-3f5b-46a0-ad68-58abb816a969
' Remark    :
'---------------------------------------------------------------------------------------
'
Public Function GetGUID(Optional ByVal i_Format As String) As String

    If UCase$(Trim(i_Format)) = "B" Then
        GetGUID = MGetGUID2("B")
    ElseIf UCase$(Trim(i_Format)) = "D" Then
        GetGUID = MGetGUID2("D")
    Else
        GetGUID = MGetGUID
    End If
    
End Function

'---------------------------------------------------------------------------------------
' Procedure : GetSuffixInFileName
' Author    : YPN
' Date      : 2017-06-28 17:50
' Purpose   : 从文件名中获取后缀名
' Param     : i_FileName 文件名
' Return    :
' Remark    :
'---------------------------------------------------------------------------------------
'
Public Function GetSuffixInFileName(ByVal i_FileName As String) As String

    GetSuffixInFileName = MGetSuffixInFileName(i_FileName)
    
End Function

'---------------------------------------------------------------------------------------
' Procedure : GetTextMaxLine
' Author    : YPN
' Date      : 2017-12-12 17:28
' Purpose   : 获取text文件最大行数（二进制方式）
' Param     : i_TextFile     指定文件
' Return    :
' Remark    :
'---------------------------------------------------------------------------------------
'
Public Function GetTextMaxLine(ByVal i_TextFile As String) As Long

    GetTextMaxLine = MGetTextMaxLine(i_TextFile)
    
End Function

'---------------------------------------------------------------------------------------
' Procedure : MGetHardDriveSerialNumber
' Author    : YPN
' Date      : 2017-06-28 12:29
' Purpose   : 取得逻辑盘序列号（非唯一）
' Param     : i_DiskName 目标逻辑盘名称
' Return    :
' Remark    : 默认C盘
'---------------------------------------------------------------------------------------
'
Public Function GetHardDriveSerialNumber(Optional ByVal i_DiskName = "C") As String

    GetHardDriveSerialNumber = MGetHardDriveSerialNumber(i_DiskName)
    
End Function

'---------------------------------------------------------------------------------------
' Procedure : MGetHardDiskSerialNumber
' Author    : YPN
' Date      : 2017-06-28 12:30
' Purpose   : 获取物理盘的系列号SN（唯一）
' Param     :
' Return    :
' Remark    :
'---------------------------------------------------------------------------------------
'
Public Function GetHardDiskSerialNumber() As String

    GetHardDiskSerialNumber = MGetHardDiskSerialNumber
    
End Function

'---------------------------------------------------------------------------------------
' Procedure : MGetHardDiskModel
' Author    : YPN
' Date      : 2017-06-28 15:48
' Purpose   : 获取物理盘的型号
' Param     :
' Return    :
' Remark    :
'---------------------------------------------------------------------------------------
'
Public Function GetHardDiskModel() As String

    GetHardDiskModel = MGetHardDiskModel
    
End Function

'---------------------------------------------------------------------------------------
' Procedure : GetInitialAll
' Author    : YPN
' Date      : 2017-06-28 17:04
' Purpose   : 获取所有汉字的首字母
' Param     : i_Str 汉字字符串
' Return    :
' Remark    :
'---------------------------------------------------------------------------------------
'
Public Function GetInitialAll(ByVal i_Str As String) As String

    GetInitialAll = MGetInitialAll(i_Str)
    
End Function

'---------------------------------------------------------------------------------------
' Procedure : GetInitialFirst
' Author    : YPN
' Date      : 2017-06-28 17:07
' Purpose   : 获取第一个汉字的首字母
' Param     : i_Str 汉字字符串
' Return    :
' Remark    :
'---------------------------------------------------------------------------------------
'
Public Function GetInitialFirst(ByVal i_Str As String) As String

    GetInitialFirst = MGetInitialFirst(i_Str)
    
End Function

'---------------------------------------------------------------------------------------
' Procedure : GetIniValue
' Author    : YPN
' Date      : 2017-12-15 11:23
' Purpose   : 获取初始化文件（.ini）指定键（Key）的值（Value）
' Param     : i_Section    节
'             i_Key        键
'             i_FileName   完整的INI文件名
' Return    :
' Remark    :
'---------------------------------------------------------------------------------------
'
Public Function GetIniValue(ByVal i_Section As String, ByVal i_Key As String, ByVal i_FileName As String) As String

    GetIniValue = MGetIniValue(i_Section, i_Key, i_FileName)
    
End Function

'---------------------------------------------------------------------------------------
' Procedure : GetLocalComputerName
' Author    : YPN
' Date      : 2018-04-08 16:00
' Purpose   : 获取本机电脑名称
' Param     :
' Return    : String
' Remark    :
'---------------------------------------------------------------------------------------
'
Public Function GetLocalComputerName() As String

    GetLocalComputerName = ModStringUtils.MGetLocalComputerName

End Function

'---------------------------------------------------------------------------------------
' Procedure : GetLocalIP
' Author    : YPN
' Date      : 2018-04-08 16:00
' Purpose   : 获取本机IP地址
' Param     :
' Return    : String
' Remark    :
'---------------------------------------------------------------------------------------
'
Public Function GetLocalIP() As String

    GetLocalIP = ModStringUtils.MGetLocalIP

End Function

'---------------------------------------------------------------------------------------
' Procedure : GetLocalUserName
' Author    : YPN
' Date      : 2018-04-08 16:00
' Purpose   : 获取本机用户名称
' Param     :
' Return    : String
' Remark    :
'---------------------------------------------------------------------------------------
'
Public Function GetLocalUserName() As String

    GetLocalUserName = ModStringUtils.MGetLocalUserName

End Function

'---------------------------------------------------------------------------------------
' Procedure : GetYear2
' Author    : YPN
' Date      : 2017-07-10 17:06
' Purpose   : 获取日期中的年份后2位
' Param     : i_Date 日期
' Return    :
' Remark    :
'---------------------------------------------------------------------------------------
'
Public Function GetYear2(ByVal i_Date As String) As Integer

    GetYear2 = MGetYear2(i_Date)
    
End Function

'---------------------------------------------------------------------------------------
' Procedure : GetMonthBegin
' Author    : YPN
' Date      : 2018/03/08 16:51
' Purpose   : 获取指定日期的月初日期
' Param     :
' Return    :
' Remark    :
'---------------------------------------------------------------------------------------
'
Public Function GetMonthBegin(ByVal i_Date As String) As String

    GetMonthBegin = MGetMonthBegin(i_Date)
    
End Function

'---------------------------------------------------------------------------------------
' Procedure : GetMonthEnd
' Author    : YPN
' Date      : 2018/03/08 16:52
' Purpose   : 获取指定日期的月末日期
' Param     :
' Return    :
' Remark    :
'---------------------------------------------------------------------------------------
'
Public Function GetMonthEnd(ByVal i_Date As String) As String

    GetMonthEnd = MGetMonthEnd(i_Date)
    
End Function

'---------------------------------------------------------------------------------------
' Procedure : HexToText
' Author    : YPN
' Date      : 2017-07-05 15:55
' Purpose   : 将16进制编码串转换为文本。没有写异常处理，但只要是用 TextToHex() 转换的结果就没问题
' Param     : i_Code 要转换的16进制编码
' Return    :
' Remark    :
'---------------------------------------------------------------------------------------
'
Public Function HexToText(i_Code As String) As String

    HexToText = MHexToText(i_Code)
    
End Function

'---------------------------------------------------------------------------------------
' Procedure : IsNull
' Author    : YPN
' Date      : 2017-06-29 14:51
' Purpose   : 判断变量是否为空
' Param     : i_Var 变量
' Return    :
' Remark    :
'---------------------------------------------------------------------------------------
'
Public Function IsNull(ByVal i_Var As Variant) As Boolean

    IsNull = MIsNull(i_Var)
    
End Function

'---------------------------------------------------------------------------------------
' Procedure : JSONAnalyze
' Author    : YPN
' Date      : 2018-4-5
' Purpose   : JSON解析
' Param     : i_JSONString 待解析的JSON字符串
'           : i_JSONKey    解析的关键字 例如：body.data.username
' Return    : String
' Remark    :
'---------------------------------------------------------------------------------------
'
Public Function JSONAnalyze(ByVal i_JSONString As String, ByVal i_JSONKey As String) As String

    JSONAnalyze = ModJSON.MJSONAnalyze(i_JSONString, i_JSONKey)
    
End Function

'---------------------------------------------------------------------------------------
' Procedure : MD5
' Author    : YPN
' Date      : 2017-06-28 17:28
' Purpose   : MD5加密
' Param     : i_Str 加密内容
'             i_Bit 加密位数：16位或32位。默认16位
' Return    :
' Remark    :
'---------------------------------------------------------------------------------------
'
Public Function MD5(ByVal i_Str As String, Optional ByVal i_Bit As Long = 16) As String

    MD5 = MMD5(i_Str, i_Bit)
    
End Function

'---------------------------------------------------------------------------------------
' Procedure : QRCode
' Author    : YPN
' Date      : 2017-09-25 14:33
' Purpose   : 生成QR Code码制的二维码
' Param     : i_QRText    二维码内容
'             i_Version  （可选参数）生成版本，支持40种，从1到40，默认自动，即0
'             i_ECLevel  （可选参数）容错等级，支持4种：L-7%、M-15%、Q-25%、H-30%，默认M（传入首字母L、M、Q、H即可）
'             i_MaskType （可选参数）模糊类型，支持8种，从0到7，默认自动，即-1
'             i_Encoding （可选参数）字符编码，支持2种：UTF-8 和 ANSI，默认UTF-8
' Return    :
' Remark    :
'---------------------------------------------------------------------------------------
'
Public Function QRCode(ByVal i_QRText As String, Optional ByVal i_Version As Long = 0, Optional ByVal i_ECLevel As String = "M", Optional ByVal i_MaskType As Long = -1, Optional ByVal i_Encoding As String = "UTF-8") As StdPicture
    
    Set QRCode = MQRCode(i_QRText, i_Version, i_ECLevel, i_MaskType, i_Encoding)
    
End Function

'---------------------------------------------------------------------------------------
' Procedure : RequestREST
' Author    : YPN
' Date      : 2018-4-5
' Purpose   : 请求/调用REST接口
' Param     : i_RequstURL        请求地址 例如：http://218.21.3.20:5076/brp/services/avplan/aAUserData/queryAAUserData
'           : i_RequestParameter 请求参数 例如：{"sysid": "SYS_LogWeb","sidv": "1.0","body": {"innerid": "568dd7cc1ba68779fe295fb9ebe3288c","userid": "czd","userno": "50271"}}
' Return    : String
' Remark    :
'---------------------------------------------------------------------------------------
'
Public Function RequestREST(ByVal i_RequstURL As String, ByVal i_RequestParameter As String) As String
    
    RequestREST = ModNetUtils.MRequestREST(i_RequstURL, i_RequestParameter)
    
End Function

'---------------------------------------------------------------------------------------
' Procedure : RndInt
' Author    : YPN
' Date      : 2018/03/05 16:55
' Purpose   : 生成指定范围内的随机整数
' Param     : i_Min 最小数
'             i_Max 最大数
' Return    : 随机整数
' Remark    :
'---------------------------------------------------------------------------------------
'
Public Function RndInt(ByVal i_Min As Long, ByVal i_Max As Long) As Integer
    
    RndInt = MRndInt(i_Min, i_Max)
    
End Function

'---------------------------------------------------------------------------------------
' Procedure : SetIniValue
' Author    : YPN
' Date      : 2017-12-15 12:19
' Purpose   : 写入初始化文件（.ini）指定键（Key）和值（Value）
' Param     : i_Section    节
'             i_Key        键
'             i_Value      值
'             i_FileName   完整的INI文件名
' Return    :
' Remark    :
'---------------------------------------------------------------------------------------
'
Public Function SetIniValue(ByVal i_Section As String, ByVal i_Key As String, ByVal i_Value As String, ByVal i_FileName As String) As Boolean
    
    SetIniValue = MSetIniValue(i_Section, i_Key, i_Value, i_FileName)
    
End Function

'---------------------------------------------------------------------------------------
' Procedure : SetStyleToolBar
' Author    : YPN
' Date      : 2017-06-30 14:34
' Purpose   : 设置控件ToolBar的样式
' Param     : i_Color       （可选参数）指定背景颜色，默认RGB(188, 210, 233)
'             i_PicturePath （可选参数）指定背景图片
' Return    :
' Remark    :
'---------------------------------------------------------------------------------------
'
Public Sub SetStyleToolBar(Optional ByVal i_Color As Long = 15323836, Optional ByVal i_PicturePath As String)

    FrmStyleToolBar.p_Color = i_Color
    FrmStyleToolBar.p_PicturePath = i_PicturePath
    FrmStyleToolBar.Hide
    
End Sub

'---------------------------------------------------------------------------------------
' Procedure : ShowMessage
' Author    : YPN
' Date      : 2018/01/23 14:31
' Purpose   : 显示右下角的弹出窗口消息
' Param     : i_MsgSource  消息来源
'             i_FrmIcon    窗口图标
'             i_FrmCaption 窗口标题
'             i_MsgTitle   消息标题
'             i_MsgContent 消息内容
'             i_WiteTime   窗口关闭等等时间
' Return    :
' Remark    :
'---------------------------------------------------------------------------------------
'
Public Sub ShowMessage(ByVal i_MsgSource As String, ByVal i_FrmIcon As StdPicture, ByVal i_FrmCaption As String, ByVal i_MsgTitle As String, ByVal i_MsgContent As String, Optional ByVal i_WiteTime As Integer)
    
    On Error GoTo ShowMessage_Error
    
    FrmMessage.Icon = i_FrmIcon
    FrmMessage.Caption = i_FrmCaption
    FrmMessage.lbl_MsgTitle = i_MsgTitle
    FrmMessage.lbl_MsgContent = i_MsgContent
    FrmMessage.lbl_MsgSource = i_MsgSource
    FrmMessage.f_WaitTime = i_WiteTime
    FrmMessage.Show
    
    On Error GoTo 0
    Exit Sub
    
ShowMessage_Error:

    ' MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure ShowMessage of Class Module ClsYPNCommonVB"
    
End Sub

'---------------------------------------------------------------------------------------
' Procedure : ShowProgress
' Author    : 杨东
' Date      : 2014/01/08
' Purpose   : 缓冲显示进度窗口
' Parameter : i_Form 要显示进度条的窗体，传入参数则显示进度窗口，否则不显示
'             i_Tip  进度条提示内容
'---------------------------------------------------------------------------------------

Public Sub ShowProgress(Optional ByVal i_Form As Object, Optional ByVal i_Tip As String)

    On Error GoTo ShowProgress_Error
    
    If Not i_Form Is Nothing Then
        
        If Not (TypeOf i_Form Is Form) Then Err.Raise 1, "ypn.common.by", "该类型不是Form类型"
        
        If i_Tip <> "" Then FrmProgress.LblContent.Caption = i_Tip
        FrmProgress.Show vbModeless, i_Form
        DoEvents
    Else
        Unload FrmProgress
    End If
    
    On Error GoTo 0
    Exit Sub
    
ShowProgress_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure ShowProgress of Class Module ClsYPNCommonVB"
    
End Sub

'---------------------------------------------------------------------------------------
' Procedure : SSTabInit
' Author    : YPN
' Date      : 2018-03-24 22:59
' Purpose   : 初始化SSTab（重绘SSTab）
' Param     : i_SSTab             SSTab类型
'             i_Style （可选参数）样式类型：0 XP样式；1 QQ样式；2 Office样式
' Return    :
' Remark    :
'---------------------------------------------------------------------------------------
'
Public Sub SSTabInit(ByVal i_SSTab As Object, Optional ByVal i_Style As Integer = 0)

    Call FrmSSTab.FSSTabInit(i_SSTab, i_Style)
    
End Sub

'---------------------------------------------------------------------------------------
' Procedure : TextToBase64
' Author    : YPN
' Date      : 2017-12-13 11:19
' Purpose   : 将文本转换为Base64位编码字符串
' Param     : i_Text 要转换的文本
' Return    :
' Remark    :
'---------------------------------------------------------------------------------------
'
Public Function TextToBase64(ByVal i_Text As String) As String

    TextToBase64 = MTextToBase64(i_Text)
    
End Function

'---------------------------------------------------------------------------------------
' Procedure : TextToHex
' Author    : YPN
' Date      : 2017-07-05 15:54
' Purpose   : 将文本转换为16进制编码串
' Param     : i_Text 要转换的文本
' Return    :
' Remark    :
'---------------------------------------------------------------------------------------
'
Public Function TextToHex(i_Text As String) As String

    TextToHex = MTextToHex(i_Text)
    
End Function

'---------------------------------------------------------------------------------------
' Procedure : TrimText
' Author    : YPN
' Date      : 2018/01/29 11:50
' Purpose   : 去除字符串两边空格，支持去除两边空的换行符
' Param     :
' Return    :
' Remark    :
'---------------------------------------------------------------------------------------
'
Public Function TrimText(i_Text As String) As String

    TrimText = MTrimText(i_Text)
    
End Function
