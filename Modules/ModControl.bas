Attribute VB_Name = "ModControl"
'---------------------------------------------------------------------------------------
' Module    : ModControl
' Author    : YPN
' Date      : 2018/08/24 12:12
' Purpose   : 控件类
'---------------------------------------------------------------------------------------

Option Explicit
'一定时间后自动关闭MsgBox
Private Declare Function MessageBoxTimeout Lib "user32" Alias "MessageBoxTimeoutA" (ByVal hWnd As Long, ByVal lpText As String, ByVal lpCaption As String, ByVal wType As Long, ByVal wlange As Long, ByVal dwTimeout As Long) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Const SW_SHOWNORMAL = 1


'---------------------------------------------------------------------------------------
' Procedure : MMsgBoxTimeout
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
Public Sub MMsgBoxTimeout(ByVal i_Form As Object, ByVal i_Msg As String, ByVal i_Type As Long, ByVal i_Tip As String, ByVal i_Timeout As Long)
    
    On Error GoTo MMsgBoxTimeout_Error
    
    If Not i_Form Is Nothing Then
        If Not (TypeOf i_Form Is Form) Then Err.Raise 1, "ypn.common.vb", "该类型不是Form类型"
        
        MessageBoxTimeout i_Form.hWnd, i_Msg, i_Tip, i_Type, 0, i_Timeout
    End If
    
    On Error GoTo 0
    Exit Sub
    
MMsgBoxTimeout_Error:
    
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure MMsgBoxTimeout of Module ModControl"
    
End Sub

'---------------------------------------------------------------------------------------
' Procedure : MOpenBrowser
' Author    : YPN
' Date      : 2019/11/27
' Purpose   : 调用默认浏览器，显示URL
'---------------------------------------------------------------------------------------
'
Public Sub MOpenBrowser(i_URL As String)
    On Error GoTo MOpenBrowser_Error
    
    Dim Result
    Result = ShellExecute(0, vbNullString, i_URL, vbNullString, vbNullString, SW_SHOWNORMAL)
    If Result <= 32 Then
        MsgBox "调用默认浏览器错误！请先设置您操作系统的默认浏览器。", vbOKOnly + vbCritical, "错误：", 0
    End If
    
    On Error GoTo 0
    Exit Sub
    
MOpenBrowser_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure MOpenBrowser of Module ModControl"
End Sub
