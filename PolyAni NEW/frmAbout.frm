VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " About"
   ClientHeight    =   2505
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4860
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2505
   ScaleWidth      =   4860
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   375
      Left            =   3855
      TabIndex        =   4
      Top             =   2025
      Width           =   870
   End
   Begin VB.Line Line1 
      X1              =   398
      X2              =   4463
      Y1              =   1095
      Y2              =   1095
   End
   Begin VB.Label lblCright 
      Alignment       =   2  'Center
      Caption         =   $"frmAbout.frx":0000
      Height          =   780
      Left            =   630
      TabIndex        =   3
      Top             =   1260
      Width           =   3615
   End
   Begin VB.Label lblCright1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright (c) 2002 - yar interactive software"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   870
      TabIndex        =   2
      Top             =   735
      Width           =   3585
   End
   Begin VB.Label lblVersion 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Version 0.96 Beta"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   870
      TabIndex        =   1
      Top             =   510
      Width           =   1425
   End
   Begin VB.Label lblCompany 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "yar-X Polygon Movie Editor"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   870
      TabIndex        =   0
      Top             =   270
      Width           =   2190
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   210
      Picture         =   "frmAbout.frx":00B2
      Top             =   285
      Width           =   480
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdOk_Click()
Me.Hide
Unload Me
End Sub

Private Sub Form_load()
Me.Show
Me.Refresh
Load frmMain
Load frmAddPoint
Me.cmdOk.Enabled = True
frmMain.Show
Me.ZOrder 0
End Sub
