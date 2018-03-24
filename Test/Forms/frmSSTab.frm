VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmSSTab 
   Caption         =   "SSTab"
   ClientHeight    =   5820
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13200
   LinkTopic       =   "Form1"
   ScaleHeight     =   5820
   ScaleWidth      =   13200
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin TabDlg.SSTab SSTab2 
      Height          =   1935
      Left            =   240
      TabIndex        =   6
      Top             =   3600
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   3413
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "Tab 0"
      TabPicture(0)   =   "frmSSTab.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).ControlCount=   0
      TabCaption(1)   =   "Tab 1"
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
      TabCaption(2)   =   "Tab 2"
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3015
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   5318
      _Version        =   393216
      Style           =   1
      Tabs            =   5
      TabsPerRow      =   5
      TabHeight       =   520
      TabCaption(0)   =   "Tab 0"
      TabPicture(0)   =   "frmSSTab.frx":001C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Tab 1"
      TabPicture(1)   =   "frmSSTab.frx":0038
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Command1"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Tab 2"
      TabPicture(2)   =   "frmSSTab.frx":0054
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Check1"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Tab 3"
      TabPicture(3)   =   "frmSSTab.frx":0070
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Option1"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "Tab 4"
      TabPicture(4)   =   "frmSSTab.frx":008C
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Combo1"
      Tab(4).ControlCount=   1
      Begin VB.ComboBox Combo1 
         Height          =   300
         Left            =   -69000
         TabIndex        =   5
         Text            =   "Combo1"
         Top             =   480
         Width           =   2055
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Option1"
         Height          =   1215
         Left            =   -73680
         TabIndex        =   4
         Top             =   1320
         Width           =   3855
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Check1"
         Height          =   1095
         Left            =   -67200
         TabIndex        =   3
         Top             =   1320
         Width           =   2655
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Command1"
         Height          =   1575
         Left            =   -71760
         TabIndex        =   2
         Top             =   1200
         Width           =   5175
      End
      Begin VB.Frame Frame1 
         Caption         =   "Frame1"
         Height          =   1095
         Left            =   840
         TabIndex        =   1
         Top             =   1080
         Width           =   3015
      End
   End
End
Attribute VB_Name = "frmSSTab"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub Form_Load()
    Dim YPN As ClsYPNCommonVB
    Dim YPN2 As ClsYPNCommonVB
    
    Set YPN = New ClsYPNCommonVB
    Call YPN.SSTabInit(Me.SSTab1, 0)
    
    Set YPN2 = New ClsYPNCommonVB
    Call YPN2.SSTabInit(Me.SSTab2, 1)
End Sub

