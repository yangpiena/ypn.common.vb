Attribute VB_Name = "ModApp"
'---------------------------------------------------------------------------------------
' Module    : ModApp
' Author    : YPN
' Date      : 2018/08/21 12:11
' Purpose   : 主应用程序方法类
'---------------------------------------------------------------------------------------
Option Explicit
Private Declare Function GetFileVersionInfo Lib "Version.dll" Alias "GetFileVersionInfoA" (ByVal lptstrFilename As String, ByVal dwhandle As Long, ByVal dwlen As Long, lpData As Any) As Long
Private Declare Function GetFileVersionInfoSize Lib "Version.dll" Alias "GetFileVersionInfoSizeA" (ByVal lptstrFilename As String, lpdwHandle As Long) As Long
Private Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, ByVal Source As Long, ByVal length As Long)
Private Declare Function VerQueryValue Lib "Version.dll" Alias "VerQueryValueA" (pBlock As Any, ByVal lpSubBlock As String, lplpBuffer As Any, puLen As Long) As Long
Private Type VS_FIXEDFILEINFO
    dwSignature As Long
    dwStrucVersionl As Integer     ' e.g. = &h0000 = 0
    dwStrucVersionh As Integer     ' e.g. = &h0042 = .42
    dwFileVersionMSl As Integer    ' e.g. = &h0003 = 3
    dwFileVersionMSh As Integer    ' e.g. = &h0075 = .75
    dwFileVersionLSl As Integer    ' e.g. = &h0000 = 0
    dwFileVersionLSh As Integer    ' e.g. = &h0031 = .31
End Type


'---------------------------------------------------------------------------------------
' Procedure : MGetVersion
' Author    : YPN
' Date      : 2018/08/21 17:13
' Purpose   : 获取App的版本号
' Param     :
' Return    : String
' Remark    : YPN Edit 2019-01-16 原方法仅能获取3位版本号，现统一为4为版本号
'---------------------------------------------------------------------------------------
'
Public Function MGetVersion(Optional ByVal i_App As Object) As String
    
    On Error GoTo MGetVersion_Error
    
    If i_App Is Nothing Then
        ' MGetVersion = App.Major & "." & App.Minor & "." & App.Revision
        MGetVersion = MGetVersionFile(App.Path & "\" & App.EXEName & ".dll")
    Else
        If Not (TypeOf i_App Is App) Then Err.Raise 5
        ' MGetVersion = i_App.Major & "." & i_App.Minor & "." & i_App.Revision
        MGetVersion = MGetVersionFile(i_App.Path & "\" & i_App.EXEName & ".dll")
    End If
    
    On Error GoTo 0
    Exit Function
    
MGetVersion_Error:
    
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure MGetVersion of Module ModApp"
    MGetVersion = ""
End Function

'---------------------------------------------------------------------------------------
' Procedure : MGetVersionFile
' Author    : YPN
' Date      : 2019-01-08 17:20
' Purpose   : 获取文件的版本号
' Param     : i_Path 文件的全路径，包括文件名
' Return    : String
' Remark    :
'---------------------------------------------------------------------------------------
'
Public Function MGetVersionFile(ByVal i_Path As String) As String
    
    Dim v_retCode      As Long, v_buffer()   As Byte
    Dim v_bufferLen    As Long, v_verPointer As Long, v_verBuffer As VS_FIXEDFILEINFO
    Dim v_verBufferLen As Long
    
    On Error GoTo MGetVersionFile_Error
    
    ' 检查文件需要多大的缓冲区
    v_bufferLen = GetFileVersionInfoSize(i_Path, 0&)
    If v_bufferLen < 1 Then
        MGetVersionFile = ""
        Exit Function
    End If
    
    ReDim v_buffer(v_bufferLen)
    ' 读取文件版本信息
    v_retCode = GetFileVersionInfo(i_Path, 0&, v_bufferLen, v_buffer(0))
    v_retCode = VerQueryValue(v_buffer(0), "\", v_verPointer, v_verBufferLen)
    MoveMemory v_verBuffer, v_verPointer, Len(v_verBuffer)
    
    MGetVersionFile = Format$(v_verBuffer.dwFileVersionMSh) & "." & Format$(v_verBuffer.dwFileVersionMSl) & "." & Format$(v_verBuffer.dwFileVersionLSh) & "." & Format$(v_verBuffer.dwFileVersionLSl)
    
    On Error GoTo 0
    Exit Function
    
MGetVersionFile_Error:
    
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure MGetVersionFile of Module ModApp"
End Function
