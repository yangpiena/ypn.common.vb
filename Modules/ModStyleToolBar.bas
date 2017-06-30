Attribute VB_Name = "ModStyleToolBar"
'---------------------------------------------------------------------------------------
' Module    : ModStyleToolBar
' Author    : YPN
' Date      : 2017-06-30 14:32
' Purpose   : ToolBarƤ��
'---------------------------------------------------------------------------------------

Option Explicit
Public Declare Function InvalidateRect Lib "user32" (ByVal hwnd As Long, lpRect As Long, ByVal bErase As Long) As Long      ' �ú�����ָ���Ĵ�������������һ�����Σ�Ȼ�󴰿ڿͻ��������һ���ֽ������»��ơ�
Public Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long                                        ' �ú�������һ������ָ����ɫ���߼�ˢ��
Public Declare Function CreatePatternBrush Lib "gdi32" (ByVal hBitmap As Long) As Long
Public Enum m_EnuTBStyleType    ' ö��Toolbar�ķ��
    m_EnuTB_FLAT = 1       ' ��ƽ���
    m_EnuTB_STANDARD = 2   ' ��׼���
End Enum

Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long ' �ú���ɾ��һ���߼��ʡ����ʡ����塢λͼ��������ߵ�ɫ�壬�ͷ�������ö����йص�ϵͳ��Դ���ڶ���ɾ��֮��ָ���ľ��Ҳ��ʧЧ�ˡ�
Private Declare Function SetClassLong Lib "user32" Alias "SetClassLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Const GCL_HBRBACKGROUND = (-10)



'---------------------------------------------------------------------------------------
' Procedure : MChangeTBBack
' Author    : YPN
' Date      : 2017-06-30 12:15
' Purpose   : �ı�Toolbar�ı���
' Param     :
' Return    :
' Remark    :
'---------------------------------------------------------------------------------------
'
Public Sub MChangeTBBack(i_ToolBar As Object, i_BackColor As Long, i_StyleType As m_EnuTBStyleType)
    
    Dim lTBWnd As Long
    
    Select Case i_StyleType
    Case m_EnuTB_FLAT                                                               ' FLAT Button Style Toolbar
        DeleteObject SetClassLong(i_ToolBar.hwnd, GCL_HBRBACKGROUND, i_BackColor)   ' Apply directly to i_ToolBar Hwnd
        
    Case m_EnuTB_STANDARD                                                           ' STANDARD Button Style Toolbar
        lTBWnd = FindWindowEx(i_ToolBar.hwnd, 0, "msvb_lib_toolbar", vbNullString)  ' Find Hwnd first
        DeleteObject SetClassLong(lTBWnd, GCL_HBRBACKGROUND, i_BackColor)           ' Set new Back
    End Select
    
End Sub
