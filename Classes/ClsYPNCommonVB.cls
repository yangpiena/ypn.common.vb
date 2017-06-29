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



Public Sub Help()
    Debug.Print "过程：GetHardDriveSerialNumber  取得逻辑盘序列号（非唯一）"
    Debug.Print "参数：i_DiskName  目标逻辑盘名称"
    Debug.Print "------"
    Debug.Print "过程：GetHardDiskSerialNumber  获取物理盘的系列号SN（唯一）"
    Debug.Print "------"
    Debug.Print "过程：GetHardDiskModel  获取物理盘的型号"
    Debug.Print "------"
    Debug.Print "过程：GetInitialFirst  获取第一个汉字的首字母"
    Debug.Print "参数：i_Str 汉字字符串"
    Debug.Print "------"
    Debug.Print "过程：GetInitialAll  获取所有汉字的首字母"
    Debug.Print "参数：i_Str 汉字字符串"
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
' Procedure : isNull
' Author    : YPN
' Date      : 2017-06-29 14:51
' Purpose   : 判断字符串是否为空
' Param     : i_Str 字符串
' Return    :
' Remark    :
'---------------------------------------------------------------------------------------
'
Public Function isNull(ByVal i_Str As Variant) As Boolean

    isNull = MIsNull(i_Str)
    
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
