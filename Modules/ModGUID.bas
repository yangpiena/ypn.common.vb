Attribute VB_Name = "ModGUID"
'---------------------------------------------------------------------------------------
' Module    : ModGUID
' Author    : YPN
' Date      : 2017-07-12 17:21
' Purpose   : 生成一个GUID
'
' UUID和GUID
'
'1. UUID: (Universally Unique Identifier) 通用唯一标识符,是一个标识符标准用于软件架构,由开放软件基金会(OSF)作为分布式计算环境(DCE)的一部分而制作的标准。
'
'         UUID的目的是让分布式系统中的所有元素都能有唯一的辨识资讯，不需要透过中央控制端来做辨认资讯的制定。如此一来每个人都建立一个与其他人不同的标
'
'         识符，这样在存储到数据库中时，就不用担心名称相同的事情(功能类似数据库中的主键,但是数据库的主键只是在一张表中有效).
'
'         这个标准现在被广泛应用在微软的全球唯一标识上面(GUID)。
'
'2. GUID:(Globally Unique Identifier) 全球唯一标识符,是一个假随机数用于软件中。
'
'   GUID的特点:
'
'    (1). 全球唯一性：世界上两台计算机生成的GUID都不相同,GUID主要用于拥有多个节点、多台计算机组成的计算机网络和系统中，分配具有唯一性的标志符。
'
'         在时间和空间上都能保证唯一性 , 保证在同一时间不同的地点生成的GUID值不同?
'
'    (2). 组成结构：通过特定算法生成的一个二进制长度为为128的字符串,在用GUID时是由算法自动生成,不需要任何机构来帮助。
'
'         GUID 的格式为“xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx”，其中每个 x 是 0-9 或 a-f 范围内的一个十六进制的数字。
'
'         例如：6F9619FF-8B86-D011-B42D-00C04FC964FF 即为有效的 GUID 值。------>一个16进制是4个二进制，所以共32位。
'
'    (3). 应用：世界上所有用户的每一个Office文档计算机都会自动生成一个GUID值，并作为这个Office的唯一标识符;而且这个GUID值与计算机的网卡是相关的，
'
'         但是这个GUID值对作者是不可见的?作者的信息可以通过GUID的值找到?
'---------------------------------------------------------------------------------------

Option Explicit

Private Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(7) As Byte
End Type
Private Declare Function CoCreateGuid Lib "ole32.dll" (pGuid As GUID) As Long
Private Declare Function StringFromGUID2 Lib "ole32.dll" (rguid As Any, ByVal lpstrClsId As Long, ByVal cbMax As Long) As Long


'---------------------------------------------------------------------------------------
' Procedure : MGetGUID
' Author    : YPN
' Date      : 2017-07-12 17:24
' Purpose   : 生成一个GUID
' Param     :
' Return    : GetGUID(): bf8b9c642ea5426f82499bb60681671b
' Remark    : GUID:(Globally Unique Identifier) 全球唯一标识符,是一个假随机数用于软件中。
'---------------------------------------------------------------------------------------
'
Public Function MGetGUID() As String
    
    Dim v_GUID As GUID
    
    If (CoCreateGuid(v_GUID) = 0) Then
        MGetGUID = _
        String(8 - Len(Hex$(v_GUID.Data1)), "0") & Hex$(v_GUID.Data1) & _
        String(4 - Len(Hex$(v_GUID.Data2)), "0") & Hex$(v_GUID.Data2) & _
        String(4 - Len(Hex$(v_GUID.Data3)), "0") & Hex$(v_GUID.Data3) & _
        IIf((v_GUID.Data4(0) < &H10), "0", "") & Hex$(v_GUID.Data4(0)) & _
        IIf((v_GUID.Data4(1) < &H10), "0", "") & Hex$(v_GUID.Data4(1)) & _
        IIf((v_GUID.Data4(2) < &H10), "0", "") & Hex$(v_GUID.Data4(2)) & _
        IIf((v_GUID.Data4(3) < &H10), "0", "") & Hex$(v_GUID.Data4(3)) & _
        IIf((v_GUID.Data4(4) < &H10), "0", "") & Hex$(v_GUID.Data4(4)) & _
        IIf((v_GUID.Data4(5) < &H10), "0", "") & Hex$(v_GUID.Data4(5)) & _
        IIf((v_GUID.Data4(6) < &H10), "0", "") & Hex$(v_GUID.Data4(6)) & _
        IIf((v_GUID.Data4(7) < &H10), "0", "") & Hex$(v_GUID.Data4(7))
    End If
    
End Function


'---------------------------------------------------------------------------------------
' Procedure : MGetGUID2
' Author    : YPN
' Date      : 2017-07-13 15:31
' Purpose   :
' Param     : i_Format 格式："B"、"D"
' Return    : GetGUID("B"): {903c1236-fe24-43c2-b9b5-bec35d9a43a8}
'             GetGUID("D"): 17e316f4-3f5b-46a0-ad68-58abb816a969
' Remark    :
'---------------------------------------------------------------------------------------
'
Public Function MGetGUID2(ByVal i_Format As String) As String
    
    Dim v_GUID    As GUID
    Dim v_Str     As String
    Dim v_Byte()  As Byte
    Dim v_Len     As Long
    Dim v_RealLen As Long
    
    v_Len = 40
    v_Byte = String(v_Len, 0)
    CoCreateGuid v_GUID                                                         '把结构转换为一个可显示的字符串
    v_RealLen = StringFromGUID2(v_GUID, VarPtr(v_Byte(0)), v_Len)
    v_Str = v_Byte
    If (Asc(Mid$(v_Str, v_RealLen, 1)) = 0) Then v_RealLen = v_RealLen - 1
        
    If UCase$(Trim(i_Format)) = "B" Then
        MGetGUID2 = Left$(v_Str, v_RealLen)
    ElseIf UCase$(Trim(i_Format)) = "D" Then
        MGetGUID2 = Mid$(v_Str, 2, 36)
    End If
    
End Function
