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
' Purpose   : YPN公共类
'---------------------------------------------------------------------------------------
Option Explicit


Public Function Help(Optional ByVal i_SubName As String, Optional ByVal i_IsMsg As Boolean = False) As String
    Dim v_Help As String
    Dim ApendText, GetFileNameInPath, GetGUID, GetSuffixInFileName, GetHardDriveSerialNumber, GetHardDiskSerialNumber, GetHardDiskModel
    Dim GetInitialFirst, GetInitialAll, GetYear2, HexToText, IsIncludeChinese, IsNull, JSONAnalyze, JSONParse, MD5, QRCode, RequestREST
    Dim RndInt, SendEmail, SendEmail2, SendHTMLEmail, SendHTMLEmail2, SetStyleToolBar, TextToHex
    
    v_Help = "传入过程名，显示详细描述！若i_IsMsg为True，则弹出详细描述，否则返回所有过程的详细描述。"
    
    ApendText = vbCrLf & "----------------------------- " _
    & vbCrLf & "过程：ApendText      追加内容到指定文件" _
    & vbCrLf & "参数：i_TextFile     指定文件" _
    & vbCrLf & "参数：i_ApendContent 追加内容"
    GetFileNameInPath = vbCrLf & "-----------------------------" _
    & vbCrLf & "过程：GetFileNameInPath 从指定全路径中获取文件名" _
    & vbCrLf & "参数：i_Path            指定全路径" _
    & vbCrLf & "参数：i_NeedSuffix      文件名是否需要后缀名，默认不需要"
    GetGUID = vbCrLf & "-----------------------------" _
    & vbCrLf & "过程：GetGUID         生成一个GUID" _
    & vbCrLf & "参数：i_Format (可选) 格式：'B'、'D'" _
    & vbCrLf & "GetGUID()默认: bf8b9c642ea5426f82499bb60681671b" _
    & vbCrLf & "GetGUID('B'): {903c1236-fe24-43c2-b9b5-bec35d9a43a8}" _
    & vbCrLf & "GetGUID('D'): 17e316f4-3f5b-46a0-ad68-58abb816a969"
    GetSuffixInFileName = vbCrLf & "-----------------------------" _
    & vbCrLf & "过程：GetSuffixInFileName 从文件名中获取后缀名" _
    & vbCrLf & "参数：i_FileName          文件名"
    GetHardDriveSerialNumber = vbCrLf & "-----------------------------" _
    & vbCrLf & "过程：GetHardDriveSerialNumber 取得逻辑盘序列号（非唯一）" _
    & vbCrLf & "参数：i_DiskName               目标逻辑盘名称"
    GetHardDiskSerialNumber = vbCrLf & "-----------------------------" _
    & vbCrLf & "过程：GetHardDiskSerialNumber 获取物理盘的系列号SN（唯一）"
    GetHardDiskModel = vbCrLf & "-----------------------------" _
    & vbCrLf & "过程：GetHardDiskModel 获取物理盘的型号"
    GetInitialFirst = vbCrLf & "-----------------------------" _
    & vbCrLf & "过程：GetInitialFirst 获取第一个汉字的首字母" _
    & vbCrLf & "参数：i_Str           汉字字符串"
    GetInitialAll = vbCrLf & "-----------------------------" _
    & vbCrLf & "过程：GetInitialAll 获取所有汉字的首字母" _
    & vbCrLf & "参数：i_Str         汉字字符串"
    GetYear2 = vbCrLf & "-----------------------------" _
    & vbCrLf & "过程：GetYear2 获取日期中的年份后2位" _
    & vbCrLf & "参数：i_Date   日期"
    HexToText = vbCrLf & "-----------------------------" _
    & vbCrLf & "过程：HexToText 将16进制编码串转换为文本。没有写异常处理，但只要是用 TextToHex() 转换的结果就没问题" _
    & vbCrLf & "参数：i_Code    要转换的16进制编码"
    IsIncludeChinese = vbCrLf & "-----------------------------" _
    & vbCrLf & "过程：IsIncludeChinese 判断是否包含中文" _
    & vbCrLf & "参数：i_Var            目标变量"
    IsNull = vbCrLf & "-----------------------------" _
    & vbCrLf & "过程：IsNull 判断变量是否为空" _
    & vbCrLf & "参数：i_Var  变量"
    JSONAnalyze = vbCrLf & "-----------------------------" _
    & vbCrLf & "过程：JSONAnalyze  JSON解析" _
    & vbCrLf & "参数：i_JSONString 待解析的JSON字符串" _
    & vbCrLf & "参数：i_JSONKey    解析的关键字 例如：body.data.username"
    JSONParse = vbCrLf & "-----------------------------" _
    & vbCrLf & "过程：JSONParse    JSON解析" _
    & vbCrLf & "参数：i_JSONString JSON格式源数据" _
    & vbCrLf & "参数：i_JSONPath   数据访问路径 例如：body.data.username。获取datas的长度，使用：body.datas.length"
    MD5 = vbCrLf & "-----------------------------" _
    & vbCrLf & "过程：MD5   MD5加密" _
    & vbCrLf & "参数：i_Str 加密内容" _
    & vbCrLf & "参数：i_Bit 加密位数：16位或32位。默认16位"
    QRCode = vbCrLf & "-----------------------------" _
    & vbCrLf & "过程：QRCode      生成QR Code码制的二维码" _
    & vbCrLf & "参数：i_QRText    二维码内容" _
    & vbCrLf & "参数：i_Version  （可选）生成版本，支持40种，从1到40，默认自动，即0" _
    & vbCrLf & "参数：i_ECLevel  （可选）容错等级，支持4种：L-7%、M-15%、Q-25%、H-30%，默认M（传入首字母L、M、Q、H即可）" _
    & vbCrLf & "参数：i_MaskType （可选）模糊类型，支持8种，从0到7，默认自动，即-1" _
    & vbCrLf & "参数：i_Encoding （可选）字符编码，支持2种：UTF-8 和 ANSI，默认UTF-8"
    RequestREST = vbCrLf & "-----------------------------" _
    & vbCrLf & "过程：RequestREST        请求/调用REST接口" _
    & vbCrLf & "参数：i_RequstURL        请求地址 例如：http://218.21.3.20:5076/brp/services/avplan/aAUserData/queryAAUserData" _
    & vbCrLf & "参数：i_RequestParameter 请求参数 例如：{""sysid"": ""SYS_LogWeb"",""sidv"": ""1.0"",""body"": {""innerid"": ""568dd7cc1ba68779fe295fb9ebe3288c"",""userid"": ""czd"",""userno"": ""50271""}}"
    RndInt = vbCrLf & "-----------------------------" _
    & vbCrLf & "过程：RndInt 生成指定范围内的随机整数" _
    & vbCrLf & "参数：i_Min  最小数" _
    & vbCrLf & "参数：i_Max  最大数"
    SendEmail = vbCrLf & "-----------------------------" _
    & vbCrLf & "过程：SendEmail     发送电子邮件，按文本格式发送" _
    & vbCrLf & "参数：i_smtpServer  SMTP服务器地址" _
    & vbCrLf & "参数：i_from        发信人邮箱地址" _
    & vbCrLf & "参数：i_password    发信人邮箱密码" _
    & vbCrLf & "参数：i_to          收信人邮箱地址，多个地址间用英文分号;隔开" _
    & vbCrLf & "参数：i_subject     邮件主题" _
    & vbCrLf & "参数：i_body        邮件正文，文本格式" _
    & vbCrLf & "参数：i_attachment （可选）附件地址，例如：D:\1.txt"
    SendEmail2 = vbCrLf & "-----------------------------" _
    & vbCrLf & "过程：SendEmail2    发送电子邮件，按文本格式发送" _
    & vbCrLf & "参数：i_smtpServer  SMTP服务器地址" _
    & vbCrLf & "参数：i_from        发信人邮箱地址" _
    & vbCrLf & "参数：i_password    发信人邮箱密码" _
    & vbCrLf & "参数：i_to          收信人邮箱地址，多个地址间用英文分号;隔开" _
    & vbCrLf & "参数：i_subject     邮件主题" _
    & vbCrLf & "参数：i_body        邮件正文，文本格式" _
    & vbCrLf & "参数：i_attachment （可选）附件地址，例如：D:\1.txt" _
    & vbCrLf & "参数：i_cc         （可选）抄送人邮箱地址" _
    & vbCrLf & "参数：i_bcc        （可选）加密抄送人邮箱地址"
    SendHTMLEmail = vbCrLf & "-----------------------------" _
    & vbCrLf & "过程：SendHTMLEmail 发送电子邮件，按HTML格式发送" _
    & vbCrLf & "参数：i_smtpServer  SMTP服务器地址" _
    & vbCrLf & "参数：i_from        发信人邮箱地址" _
    & vbCrLf & "参数：i_password    发信人邮箱密码" _
    & vbCrLf & "参数：i_to          收信人邮箱地址，多个地址间用英文分号;隔开" _
    & vbCrLf & "参数：i_subject     邮件主题" _
    & vbCrLf & "参数：i_body        邮件正文，文本格式" _
    & vbCrLf & "参数：i_attachment （可选）附件地址，例如：D:\1.txt"
    SendHTMLEmail2 = vbCrLf & "-----------------------------" _
    & vbCrLf & "过程：SendHTMLEmail2 发送电子邮件，按HTML格式发送，支持抄送和加密抄送" _
    & vbCrLf & "参数：i_smtpServer   SMTP服务器地址" _
    & vbCrLf & "参数：i_from         发信人邮箱地址" _
    & vbCrLf & "参数：i_password     发信人邮箱密码" _
    & vbCrLf & "参数：i_to           收信人邮箱地址，多个地址间用英文分号;隔开" _
    & vbCrLf & "参数：i_subject      邮件主题" _
    & vbCrLf & "参数：i_body         邮件正文，文本格式" _
    & vbCrLf & "参数：i_attachment  （可选）附件地址，例如：D:\1.txt" _
    & vbCrLf & "参数：i_cc          （可选）抄送人邮箱地址" _
    & vbCrLf & "参数：i_bcc         （可选）加密抄送人邮箱地址"
    SetStyleToolBar = vbCrLf & "-----------------------------" _
    & vbCrLf & "过程：SetStyleToolBar 设置控件ToolBar的样式" _
    & vbCrLf & "参数：i_Color        （可选）指定背景颜色，默认RGB(188, 210, 233)" _
    & vbCrLf & "参数：i_PicturePath  （可选）指定背景图片"
    TextToHex = vbCrLf & "-----------------------------" _
    & vbCrLf & "过程：TextToHex 将文本转换为16进制编码串" _
    & vbCrLf & "参数：i_Text    要转换的文本"
    
    Select Case i_SubName
    Case "ApendText"
        v_Help = v_Help & ApendText
    Case "GetFileNameInPath"
        v_Help = v_Help & GetFileNameInPath
    Case "GetGUID"
        v_Help = v_Help & GetGUID
    Case "GetSuffixInFileName"
        v_Help = v_Help & GetSuffixInFileName
    Case "GetHardDriveSerialNumber"
        v_Help = v_Help & GetHardDriveSerialNumber
    Case "GetHardDiskSerialNumber"
        v_Help = v_Help & GetHardDiskSerialNumber
    Case "GetHardDiskModel"
        v_Help = v_Help & GetHardDiskModel
    Case "GetInitialFirst"
        v_Help = v_Help & GetInitialFirst
    Case "GetInitialAll"
        v_Help = v_Help & GetInitialAll
    Case "GetYear2"
        v_Help = v_Help & GetYear2
    Case "HexToText"
        v_Help = v_Help & HexToText
    Case "IsIncludeChinese"
        v_Help = v_Help & IsIncludeChinese
    Case "IsNull"
        v_Help = v_Help & IsNull
    Case "JSONAnalyze"
        v_Help = v_Help & JSONAnalyze
    Case "JSONParse"
        v_Help = v_Help & JSONParse
    Case "MD5"
        v_Help = v_Help & MD5
    Case "QRCode"
        v_Help = v_Help & QRCode
    Case "RequestREST"
        v_Help = v_Help & RequestREST
    Case "RndInt"
        v_Help = v_Help & RndInt
    Case "SendEmail"
        v_Help = v_Help & SendEmail
    Case "SendEmail2"
        v_Help = v_Help & SendEmail2
    Case "SendHTMLEmail"
        v_Help = v_Help & SendHTMLEmail
    Case "SendHTMLEmail2"
        v_Help = v_Help & SendHTMLEmail2
    Case "SetStyleToolBar"
        v_Help = v_Help & SetStyleToolBar
    Case "TextToHex"
        v_Help = v_Help & TextToHex
    Case Else
        v_Help = "以下为YPNCommonVB的所有过程："
        v_Help = v_Help & ApendText & GetFileNameInPath & GetGUID & GetSuffixInFileName & GetHardDriveSerialNumber & GetHardDiskSerialNumber _
        & GetHardDiskModel & GetInitialFirst & GetInitialAll & GetYear2 & HexToText & IsIncludeChinese & IsNull & JSONAnalyze & JSONParse _
        & MD5 & QRCode & RequestREST & RndInt & SendEmail & SendEmail2 & SendHTMLEmail & SendHTMLEmail2 & SetStyleToolBar & TextToHex
        v_Help = v_Help & vbCrLf & "-----------------------------" & "《结束》"
    End Select
    Help = v_Help
    If i_IsMsg Then MsgBox v_Help, vbInformation, TS
End Function

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
' Procedure : FormatDate
' Author    : YPN
' Date      : 2018-04-26
' Purpose   : 按指定分隔符格式化日期
' Param     : i_Date      目标日期
'             i_Delimiter 指定分隔符
' Return    : String
' Remark    :
'---------------------------------------------------------------------------------------
'
Public Function FormatDate(ByVal i_Date As String, ByVal i_Delimiter As String) As String
    FormatDate = ModStringUtils.MFormatDate(i_Date, i_Delimiter)
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
Public Function GetInitialAll(ByVal i_str As String) As String
    GetInitialAll = MGetInitialAll(i_str)
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
Public Function GetInitialFirst(ByVal i_str As String) As String
    GetInitialFirst = MGetInitialFirst(i_str)
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
' Procedure : GetVersion
' Author    : YPN
' Date      : 2018/08/21 17:14
' Purpose   : 获取App的版本号
' Param     : i_App 指定App
' Return    : String
' Remark    :
'---------------------------------------------------------------------------------------
'
Public Function GetVersion(Optional ByVal i_App As Object) As String
    GetVersion = ModApp.MGetVersion(i_App)
End Function

'---------------------------------------------------------------------------------------
' Procedure : GetVersionFile
' Author    : YPN
' Date      : 2019-01-08 17:25
' Purpose   : 获取文件的版本号
' Param     : i_Path 文件的全路径，包括文件名
' Return    : String
' Remark    :
'---------------------------------------------------------------------------------------
'
Public Function GetVersionFile(ByVal i_Path As String) As String
    GetVersionFile = ModApp.MGetVersionFile(i_Path)
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
' Procedure : IsIncludeChinese
' Author    : YPN
' Date      : 2019/03/28 16:15
' Purpose   : 判断是否包含中文
' Param     : i_Var 目标变量
' Return    : Boolean
' Remark    :
'---------------------------------------------------------------------------------------
'
Public Function IsIncludeChinese(ByVal i_Var As Variant) As Boolean
    IsIncludeChinese = ModStringUtils.MIsIncludeChinese(i_Var)
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
' Procedure : JSONParse
' Author    : YPN
' Date      : 2019/03/08 11:39
' Purpose   : JSON解析
' Param     : i_JSONString JSON格式源数据
'             i_JSONPath   数据访问路径
' Return    : Variant
' Remark    :
'---------------------------------------------------------------------------------------
'
Public Function JSONParse(ByVal i_JSONString As String, ByVal i_JSONPath As String) As Variant
    JSONParse = ModJSON.MJSONParse(i_JSONString, i_JSONPath)
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
Public Function MD5(ByVal i_str As String, Optional ByVal i_Bit As Long = 16) As String
    MD5 = MMD5(i_str, i_Bit)
End Function

'---------------------------------------------------------------------------------------
' Procedure : MsgBoxTimeout
' Author    : YPN
' Date      : 2018/08/24 12:18
' Purpose   : 弹出指定时间后消失的消息框
' Param     : i_Form     要弹出消息框的窗体
'             i_Msg      消息框内容
'             i_Type     消息框类型
'             i_Tip      消息框标题
'             i_Timeout  消息框显示时间
' Return    :
' Remark    :
'---------------------------------------------------------------------------------------
'
Public Sub MsgBoxTimeout(ByVal i_Form As Object, ByVal i_Msg As String, ByVal i_Type As Long, ByVal i_Tip As String, ByVal i_Timeout As Long)
    ModControl.MMsgBoxTimeout i_Form, i_Msg, i_Type, i_Tip, i_Timeout
End Sub

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
Public Function QRCode(ByVal i_QRText As String, Optional ByVal i_Version As Long = 0, Optional ByVal i_ECLevel As String = "M", Optional ByVal i_MaskType As Long = -1, Optional ByVal i_encoding As String = "UTF-8") As StdPicture
    Set QRCode = MQRCode(i_QRText, i_Version, i_ECLevel, i_MaskType, i_encoding)
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
' Procedure : SendEmail
' Author    : YPN
' Date      : 2018-04-25 17:29
' Purpose   : 发送电子邮件，按文本格式发送
' Param     : i_smtpServer SMTP服务器地址，例如：smtp.qiye.163.com
'             i_from       发信人邮箱地址
'             i_password   发信人邮箱密码
'             i_to         收信人邮箱地址，多个地址间用英文分号;隔开
'             i_subject    邮件主题
'             i_body       邮件正文，文本格式
'             i_attachment （可选）附件地址，例如："D:\1.txt"
' Return    : String       发送成功则返回success，失败则返回failure加失败原因
' Remark    :
'---------------------------------------------------------------------------------------
'
Public Function SendEmail(ByVal i_smtpServer As String, ByVal i_from As String, ByVal i_password As String, ByVal i_to As String, _
    ByVal i_subject As String, ByVal i_body As String, Optional ByVal i_attachment As String) As String
    SendEmail = ModNetUtils.MSendEmail(i_smtpServer, i_from, i_password, i_to, i_subject, i_body, i_attachment)
End Function

'---------------------------------------------------------------------------------------
' Procedure : SendEmail2
' Author    : YPN
' Date      : 2019/04/10 14:35
' Purpose   : 发送电子邮件，按文本格式发送，支持抄送和加密抄送
' Param     : i_smtpServer SMTP服务器地址
'             i_from       发信人邮箱地址
'             i_password   发信人邮箱密码
'             i_to         收信人邮箱地址，多个地址间用英文分号;隔开
'             i_subject    邮件主题
'             i_body       邮件正文，文本格式
'             i_attachment （可选）附件地址，例如："D:\1.txt"
'             i_cc         （可选）抄送人邮箱地址
'             i_bcc        （可选）加密抄送人邮箱地址
' Return    : String       发送成功则返回success，失败则返回failure加失败原因
' Remark    :
'---------------------------------------------------------------------------------------
'
Public Function SendEmail2(ByVal i_smtpServer As String, ByVal i_from As String, ByVal i_password As String, ByVal i_to As String, _
    ByVal i_subject As String, ByVal i_body As String, Optional ByVal i_attachment As String, Optional ByVal i_cc As String, Optional ByVal i_bcc As String) As String
    SendEmail2 = ModNetUtils.MSendEmail2(i_smtpServer, i_from, i_password, i_to, i_subject, i_body, i_attachment, i_cc, i_bcc)
End Function

'---------------------------------------------------------------------------------------
' Procedure : SendHTMLEmail
' Author    : YPN
' Date      : 2018-04-25 17:29
' Purpose   : 发送电子邮件，按HTML格式发送
' Param     : i_smtpServer SMTP服务器地址，例如：smtp.qiye.163.com
'             i_from       发信人邮箱地址
'             i_password   发信人邮箱密码
'             i_to         收信人邮箱地址，多个地址间用英文分号;隔开
'             i_subject    邮件主题
'             i_body       邮件正文，HTML格式
'             i_attachment （可选）附件地址，例如："D:\1.txt"
' Return    : String       发送成功则返回success，失败则返回failure加失败原因
' Remark    :
'---------------------------------------------------------------------------------------
'
Public Function SendHTMLEmail(ByVal i_smtpServer As String, ByVal i_from As String, ByVal i_password As String, ByVal i_to As String, _
    ByVal i_subject As String, ByVal i_body As String, Optional ByVal i_attachment As String) As String
    SendHTMLEmail = ModNetUtils.MSendHTMLEmail(i_smtpServer, i_from, i_password, i_to, i_subject, i_body, i_attachment)
End Function

'---------------------------------------------------------------------------------------
' Procedure : SendHTMLEmail2
' Author    : YPN
' Date      : 2019/04/10 14:29
' Purpose   : 发送电子邮件，按HTML格式发送，支持抄送和加密抄送
' Param     : i_smtpServer SMTP服务器地址
'             i_from       发信人邮箱地址
'             i_password   发信人邮箱密码
'             i_to         收信人邮箱地址，多个地址间用英文分号;隔开
'             i_subject    邮件主题
'             i_body       邮件正文，HTML格式
'             i_attachment （可选）附件地址，例如："D:\1.txt"
'             i_cc         （可选）抄送人邮箱地址
'             i_bcc        （可选）加密抄送人邮箱地址
' Return    : String       发送成功则返回success，失败则返回failure加失败原因
' Remark    :
'---------------------------------------------------------------------------------------
'
Public Function SendHTMLEmail2(ByVal i_smtpServer As String, ByVal i_from As String, ByVal i_password As String, ByVal i_to As String, _
    ByVal i_subject As String, ByVal i_body As String, Optional ByVal i_attachment As String, Optional ByVal i_cc As String, Optional ByVal i_bcc As String) As String
    SendHTMLEmail2 = ModNetUtils.MSendHTMLEmail2(i_smtpServer, i_from, i_password, i_to, i_subject, i_body, i_attachment, i_cc, i_bcc)
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
    FrmStyleToolBar.F_Color = i_Color
    FrmStyleToolBar.F_PicturePath = i_PicturePath
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
    FrmMessage.F_WaitTime = i_WiteTime
    FrmMessage.Show
    
    On Error GoTo 0
    Exit Sub
    
ShowMessage_Error:
    
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure ShowMessage of Class Module ClsYPNCommonVB"
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
    On Error Resume Next
    
    If Not i_Form Is Nothing Then
        
        If Not (TypeOf i_Form Is Form) Then Err.Raise 1, "ypn.common.vb", "该类型不是Form类型"
        
        If i_Tip <> "" Then FrmProgress.LblContent.Caption = i_Tip
        FrmProgress.Show vbModeless, i_Form
        DoEvents
    Else
        Unload FrmProgress
    End If
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
Public Function TextToBase64(ByVal i_text As String) As String
    TextToBase64 = MTextToBase64(i_text)
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
Public Function TextToHex(i_text As String) As String
    TextToHex = MTextToHex(i_text)
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
Public Function TrimText(i_text As String) As String
    TrimText = MTrimText(i_text)
End Function
