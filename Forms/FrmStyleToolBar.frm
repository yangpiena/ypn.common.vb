VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmStyleToolBar 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  '����ȱʡ
   Begin VB.PictureBox Picture1 
      Height          =   495
      Left            =   1560
      ScaleHeight     =   435
      ScaleWidth      =   1635
      TabIndex        =   1
      Top             =   1200
      Width           =   1695
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   615
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4680
      _ExtentX        =   8255
      _ExtentY        =   1085
      ButtonWidth     =   609
      ButtonHeight    =   926
      Appearance      =   1
      _Version        =   393216
   End
End
Attribute VB_Name = "FrmStyleToolBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : FrmStyleToolBar
' Author    : YPN
' Date      : 2017-06-30 14:36
' Purpose   : ToolBar��ʽ
'---------------------------------------------------------------------------------------

Option Explicit
Public p_Color As Long            ' ��ɫֵ
Public p_PicturePath As String    ' ͼƬ·��������ͼƬ���ƣ�



Private Sub Form_Load()
    
    Call applyStyle
    
End Sub

'---------------------------------------------------------------------------------------
' Procedure : applyStyle
' Author    : YPN
' Date      : 2017-06-30 12:24
' Purpose   : Ӧ��Toolbar����ʽ
' Param     :
' Return    :
' Remark    :
'---------------------------------------------------------------------------------------
'
Private Sub applyStyle()
    
    Dim v_BG As Long
    
    If Trim(p_PicturePath) <> "" Then
        ' ʹ��ͼƬ��ΪToolBar�ı���
        Picture1.Picture = LoadPicture(p_PicturePath)
        v_BG = CreatePatternBrush(Picture1.Picture.Handle)     ' Creates the background from a Picture Handle
        Call MChangeTBBack(Me.Toolbar1, v_BG, m_EnuTB_FLAT)      ' ������ʽ��m_EnuTB_FLAT �� m_EnuTB_STANDARD
    Else
        v_BG = CreateSolidBrush(p_Color)                       ' ����ָ����ɫ����һ������ (Long)
        Call MChangeTBBack(Me.Toolbar1, v_BG, m_EnuTB_FLAT)
    End If
    
    ' ˢ����Ļ�Կ�����ʽ
    InvalidateRect 0&, 0&, False
    
End Sub
