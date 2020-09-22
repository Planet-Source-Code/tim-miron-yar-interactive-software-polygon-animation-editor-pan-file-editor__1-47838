VERSION 5.00
Begin VB.Form frmAddPoint 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Add Point"
   ClientHeight    =   480
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3030
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   480
   ScaleWidth      =   3030
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdAddPoint 
      Caption         =   "Add"
      Default         =   -1  'True
      Height          =   360
      Left            =   2145
      TabIndex        =   4
      Top             =   60
      Width           =   840
   End
   Begin VB.TextBox txtY 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1395
      TabIndex        =   1
      Text            =   "0"
      Top             =   98
      Width           =   675
   End
   Begin VB.TextBox txtX 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   300
      TabIndex        =   0
      Text            =   "0"
      Top             =   98
      Width           =   675
   End
   Begin VB.Label lblY 
      AutoSize        =   -1  'True
      Caption         =   "Y:"
      Height          =   195
      Left            =   1200
      TabIndex        =   3
      Top             =   143
      Width           =   150
   End
   Begin VB.Label lblX 
      AutoSize        =   -1  'True
      Caption         =   "X:"
      Height          =   195
      Left            =   75
      TabIndex        =   2
      Top             =   143
      Width           =   150
   End
End
Attribute VB_Name = "frmAddPoint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CmdAddPoint_Click()
CurShape.PntCount = CurShape.PntCount + 1
frmMain.lblStat3.Caption = "Point Count: " & CurShape.PntCount
frmMain.cmdTB(11).Enabled = True

CurShape.PolyPnt(CurShape.PntCount).x = txtX.Text
CurShape.PolyPnt(CurShape.PntCount).y = txtY.Text

frmMain.FirstPoint = True
frmMain.linTemp.Visible = False
frmMain.linTemp.X1 = txtX.Text
frmMain.linTemp.Y1 = txtY.Text

frmMain.picMain.Refresh
frmMain.lblStat.Caption = "POLYGON: Point Count - " & CurShape.PntCount
End Sub
