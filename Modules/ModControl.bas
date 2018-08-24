Attribute VB_Name = "ModControl"
'---------------------------------------------------------------------------------------
' Module    : ModControl
' Author    : YPN
' Date      : 2018/08/24 12:12
' Purpose   : �ؼ���
'---------------------------------------------------------------------------------------

Option Explicit
'һ��ʱ����Զ��ر�MsgBox
Private Declare Function MessageBoxTimeout Lib "user32" Alias "MessageBoxTimeoutA" (ByVal hWnd As Long, ByVal lpText As String, ByVal lpCaption As String, ByVal wType As Long, ByVal wlange As Long, ByVal dwTimeout As Long) As Long


'---------------------------------------------------------------------------------------
' Procedure : MMsgBoxTimeout
' Author    : YPN
' Date      : 2018/08/24 12:18
' Purpose   : ����ָ��ʱ�����ʧ����Ϣ��
' Param     : i_Form     Ҫ������Ϣ��Ĵ���
'             i_Msg      ��Ϣ������
'             i_Type     ��Ϣ������
'             i_Tip      ��Ϣ�����
'             i_Timeout  ��Ϣ����ʾʱ��
' Return    :
' Remark    :
'---------------------------------------------------------------------------------------
'
Public Sub MMsgBoxTimeout(ByVal i_Form As Object, ByVal i_Msg As String, ByVal i_Type As Long, ByVal i_Tip As String, ByVal i_Timeout As Long)
    
    On Error GoTo MMsgBoxTimeout_Error
    
    If Not i_Form Is Nothing Then
        If Not (TypeOf i_Form Is Form) Then Err.Raise 1, "ypn.common.vb", "�����Ͳ���Form����"
        
        MessageBoxTimeout i_Form.hWnd, i_Msg, i_Tip, i_Type, 0, i_Timeout
    End If
    
    On Error GoTo 0
    Exit Sub
    
MMsgBoxTimeout_Error:
    
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure MMsgBoxTimeout of Module ModControl"
    
End Sub
