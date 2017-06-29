VERSION 5.00
Begin VB.Form FrmTest 
   Caption         =   "≤‚ ‘"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  '¥∞ø⁄»± °
   Begin VB.TextBox Text2 
      Height          =   345
      Left            =   1485
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   555
      Width           =   2940
   End
   Begin VB.TextBox Text1 
      Height          =   345
      Left            =   1500
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   173
      Width           =   2940
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "ŒÔ¿Ì≈Ã–Ú¡–∫≈£∫"
      Height          =   180
      Left            =   180
      TabIndex        =   3
      Top             =   630
      Width           =   1260
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "¬ﬂº≠≈Ã–Ú¡–∫≈£∫"
      Height          =   180
      Left            =   195
      TabIndex        =   1
      Top             =   255
      Width           =   1260
   End
End
Attribute VB_Name = "FrmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()

    Me.Text1.Text = MGetHardDriveSerialNumber
    Me.Text2.Text = MGetHardDiskInfo
    
End Sub

