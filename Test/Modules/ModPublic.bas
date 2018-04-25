Attribute VB_Name = "ModPublic"
Option Explicit

Public YPN As New ClsYPNCommonVB

Private Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(8) As Byte
End Type
Private Declare Function CoCreateGuid Lib "ole32.dll" (pguid As GUID) As Long
Private Declare Function StringFromGUID2 Lib "ole32.dll" (rguid As Any, ByVal lpstrClsId As Long, ByVal cbMax As Long) As Long


Public Function GUIDGen() As String '生成GUID
    
    Dim uGUID As GUID
    Dim sGUID As String
    Dim bGUID() As Byte
    Dim lLen As Long
    Dim RetVal As Long
    
    lLen = 40
    bGUID = String(lLen, 0)
    CoCreateGuid uGUID '把结构转换为一个可显示的字符串
    RetVal = StringFromGUID2(uGUID, VarPtr(bGUID(0)), lLen)
    sGUID = bGUID
    If (Asc(Mid$(sGUID, RetVal, 1)) = 0) Then RetVal = RetVal - 1
    GUIDGen = Left$(sGUID, RetVal)
    GUIDGen = Mid$(sGUID, 2, 36)
    
End Function

Public Function CreateGUID() As String
    
    Dim udtGUID         As GUID
    Dim sGUID         As String
    Dim lResult         As Long
    
    lResult = CoCreateGuid(udtGUID)
    If lResult Then
        sGUID = " "
    Else
        sGUID = String$(38, 0)
        StringFromGUID2 udtGUID, StrPtr(sGUID), 39
    End If
    CreateGUID = sGUID
    
End Function

