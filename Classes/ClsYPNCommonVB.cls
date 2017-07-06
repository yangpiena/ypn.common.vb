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



Public Sub Help(Optional ByVal i_IsMsg As Boolean = False)
    
    Dim v_Help As String
    
    v_Help = "以下为YPNCommonVB的所有过程：" & vbCrLf
    
    v_Help = v_Help & vbCrLf & "过程：GetFileNameInPath         从指定全路径中获取文件名"
    v_Help = v_Help & vbCrLf & "参数：i_Path       指定全路径"
    v_Help = v_Help & vbCrLf & "参数：i_NeedSuffix 文件名是否需要后缀名，默认不需要"
    v_Help = v_Help & vbCrLf & "--------------------------------"
    v_Help = v_Help & vbCrLf & "过程：GetSuffixInFileName       从文件名中获取后缀名"
    v_Help = v_Help & vbCrLf & "参数：i_FileName 文件名"
    v_Help = v_Help & vbCrLf & "--------------------------------"
    v_Help = v_Help & vbCrLf & "过程：GetHardDriveSerialNumber  取得逻辑盘序列号（非唯一）"
    v_Help = v_Help & vbCrLf & "参数：i_DiskName 目标逻辑盘名称"
    v_Help = v_Help & vbCrLf & "--------------------------------"
    v_Help = v_Help & vbCrLf & "过程：GetHardDiskSerialNumber   获取物理盘的系列号SN（唯一）"
    v_Help = v_Help & vbCrLf & "--------------------------------"
    v_Help = v_Help & vbCrLf & "过程：GetHardDiskModel          获取物理盘的型号"
    v_Help = v_Help & vbCrLf & "--------------------------------"
    v_Help = v_Help & vbCrLf & "过程：GetInitialFirst           获取第一个汉字的首字母"
    v_Help = v_Help & vbCrLf & "参数：i_Str 汉字字符串"
    v_Help = v_Help & vbCrLf & "--------------------------------"
    v_Help = v_Help & vbCrLf & "过程：GetInitialAll             获取所有汉字的首字母"
    v_Help = v_Help & vbCrLf & "参数：i_Str 汉字字符串"
    v_Help = v_Help & vbCrLf & "--------------------------------"
    v_Help = v_Help & vbCrLf & "过程：HexToText                 将16进制编码串转换为文本。没有写异常处理，但只要是用 TextToHex() 转换的结果就没问题"
    v_Help = v_Help & vbCrLf & "参数：i_Code 要转换的16进制编码"
    v_Help = v_Help & vbCrLf & "--------------------------------"
    v_Help = v_Help & vbCrLf & "过程：TextToHex                 将文本转换为16进制编码串"
    v_Help = v_Help & vbCrLf & "参数：i_Text 要转换的文本"
    v_Help = v_Help & vbCrLf & "--------------------------------"
    v_Help = v_Help & vbCrLf & "过程：isNull                    判断变量是否为空"
    v_Help = v_Help & vbCrLf & "参数：i_Var 变量"
    v_Help = v_Help & vbCrLf & "--------------------------------"
    v_Help = v_Help & vbCrLf & "过程：SetStyleToolBar           设置控件ToolBar的样式"
    v_Help = v_Help & vbCrLf & "参数：i_Color       （可选参数）指定背景颜色，默认RGB(188, 210, 233)"
    v_Help = v_Help & vbCrLf & "参数：i_PicturePath （可选参数）指定背景图片"
    v_Help = v_Help & vbCrLf & "--------------------------------"
    v_Help = v_Help & vbCrLf & "过程：MD5                       MD5加密"
    v_Help = v_Help & vbCrLf & "参数：i_Str 加密内容"
    v_Help = v_Help & vbCrLf & "参数：i_Bit 加密位数：16位或32位。默认16位"
    
    Debug.Print v_Help
    
    If i_IsMsg Then MsgBox v_Help, vbInformation, "提示"
    
End Sub



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
' Procedure : isNull
' Author    : YPN
' Date      : 2017-06-29 14:51
' Purpose   : 判断变量是否为空
' Param     : i_Var 变量
' Return    :
' Remark    :
'---------------------------------------------------------------------------------------
'
Public Function isNull(ByVal i_Var As Variant) As Boolean
    
    isNull = MIsNull(i_Var)
    
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
' Procedure : ShowProgress
' Author    : 杨东
' Date      : 2014/01/08
' Purpose   : 缓冲显示进度窗口
' Parameter : i_Form 要显示进度条的窗体，传入参数则显示进度窗口，否则不显示
'             i_Tip  进度条提示内容
'---------------------------------------------------------------------------------------

Public Sub ShowProgress(Optional ByVal i_Form As Object, Optional ByVal i_Tip As String)
    
    If Not i_Form Is Nothing Then
        
        If Not (TypeOf i_Form Is Form) Then Err.Raise 5
        
        If i_Tip <> "" Then FrmProgress.LblContent.Caption = i_Tip
        FrmProgress.Show vbModeless, i_Form
    Else
        Unload FrmProgress
    End If
    
End Sub

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
