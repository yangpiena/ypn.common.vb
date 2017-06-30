Attribute VB_Name = "ModStyleToolBar"
'---------------------------------------------------------------------------------------
' Module    : ModStyleToolBar
' Author    : YPN
' Date      : 2017-06-30 14:32
' Purpose   : ToolBar皮肤
'---------------------------------------------------------------------------------------

Option Explicit
Public Declare Function InvalidateRect Lib "user32" (ByVal hwnd As Long, lpRect As Long, ByVal bErase As Long) As Long      ' 该函数向指定的窗体更新区域添加一个矩形，然后窗口客户区域的这一部分将被重新绘制。
Public Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long                                        ' 该函数创建一个具有指定颜色的逻辑刷子
Public Declare Function CreatePatternBrush Lib "gdi32" (ByVal hBitmap As Long) As Long
Public Enum m_EnuTBStyleType    ' 枚举Toolbar的风格
    m_EnuTB_FLAT = 1       ' 扁平风格
    m_EnuTB_STANDARD = 2   ' 标准风格
End Enum

Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long ' 该函数删除一个逻辑笔、画笔、字体、位图、区域或者调色板，释放所有与该对象有关的系统资源，在对象被删除之后，指定的句柄也就失效了。
Private Declare Function SetClassLong Lib "user32" Alias "SetClassLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Const GCL_HBRBACKGROUND = (-10)



'---------------------------------------------------------------------------------------
' Procedure : MChangeTBBack
' Author    : YPN
' Date      : 2017-06-30 12:15
' Purpose   : 改变Toolbar的背景
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
