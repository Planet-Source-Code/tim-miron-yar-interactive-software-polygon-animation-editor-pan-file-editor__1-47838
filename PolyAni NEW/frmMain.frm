VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H8000000C&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "yar-X  -  Polygon Movie Editor  [untitled.pan]"
   ClientHeight    =   5205
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   8580
   HasDC           =   0   'False
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   347
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   572
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picRTB 
      Align           =   4  'Align Right
      BorderStyle     =   0  'None
      HasDC           =   0   'False
      Height          =   4425
      Left            =   6420
      ScaleHeight     =   295
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   144
      TabIndex        =   8
      Top             =   435
      Width           =   2160
      Begin VB.PictureBox cmdTB 
         BorderStyle     =   0  'None
         HasDC           =   0   'False
         Height          =   180
         Index           =   18
         Left            =   1785
         Picture         =   "frmMain.frx":0CCA
         ScaleHeight     =   12
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   23
         TabIndex        =   33
         ToolTipText     =   "Move Up"
         Top             =   4050
         Width           =   345
         Begin VB.Line lin_Br 
            BorderColor     =   &H80000010&
            Index           =   18
            X1              =   22
            X2              =   22
            Y1              =   0
            Y2              =   11
         End
         Begin VB.Line lin_Bb 
            BorderColor     =   &H80000010&
            Index           =   18
            X1              =   0
            X2              =   23
            Y1              =   11
            Y2              =   11
         End
         Begin VB.Line lin_Bt 
            BorderColor     =   &H80000014&
            Index           =   18
            X1              =   0
            X2              =   22
            Y1              =   0
            Y2              =   0
         End
         Begin VB.Line lin_bL 
            BorderColor     =   &H80000014&
            Index           =   18
            X1              =   0
            X2              =   0
            Y1              =   0
            Y2              =   11
         End
      End
      Begin VB.PictureBox cmdTB 
         BorderStyle     =   0  'None
         HasDC           =   0   'False
         Height          =   180
         Index           =   17
         Left            =   1785
         Picture         =   "frmMain.frx":101F
         ScaleHeight     =   12
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   23
         TabIndex        =   32
         ToolTipText     =   "Move Down"
         Top             =   4245
         Width           =   345
         Begin VB.Line lin_bL 
            BorderColor     =   &H80000014&
            Index           =   17
            X1              =   0
            X2              =   0
            Y1              =   0
            Y2              =   11
         End
         Begin VB.Line lin_Bt 
            BorderColor     =   &H80000014&
            Index           =   17
            X1              =   0
            X2              =   22
            Y1              =   0
            Y2              =   0
         End
         Begin VB.Line lin_Bb 
            BorderColor     =   &H80000010&
            Index           =   17
            X1              =   0
            X2              =   23
            Y1              =   11
            Y2              =   11
         End
         Begin VB.Line lin_Br 
            BorderColor     =   &H80000010&
            Index           =   17
            X1              =   22
            X2              =   22
            Y1              =   0
            Y2              =   11
         End
      End
      Begin VB.PictureBox cmdTB 
         BorderStyle     =   0  'None
         HasDC           =   0   'False
         Height          =   180
         Index           =   16
         Left            =   1785
         Picture         =   "frmMain.frx":1376
         ScaleHeight     =   12
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   23
         TabIndex        =   31
         ToolTipText     =   "Move Down"
         Top             =   3120
         Width           =   345
         Begin VB.Line lin_Br 
            BorderColor     =   &H80000010&
            Index           =   16
            X1              =   22
            X2              =   22
            Y1              =   0
            Y2              =   11
         End
         Begin VB.Line lin_Bb 
            BorderColor     =   &H80000010&
            Index           =   16
            X1              =   0
            X2              =   23
            Y1              =   11
            Y2              =   11
         End
         Begin VB.Line lin_Bt 
            BorderColor     =   &H80000014&
            Index           =   16
            X1              =   0
            X2              =   22
            Y1              =   0
            Y2              =   0
         End
         Begin VB.Line lin_bL 
            BorderColor     =   &H80000014&
            Index           =   16
            X1              =   0
            X2              =   0
            Y1              =   0
            Y2              =   11
         End
      End
      Begin VB.PictureBox cmdTB 
         BorderStyle     =   0  'None
         HasDC           =   0   'False
         Height          =   180
         Index           =   15
         Left            =   1785
         Picture         =   "frmMain.frx":16CD
         ScaleHeight     =   12
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   23
         TabIndex        =   30
         ToolTipText     =   "Move Up"
         Top             =   2925
         Width           =   345
         Begin VB.Line lin_bL 
            BorderColor     =   &H80000014&
            Index           =   15
            X1              =   0
            X2              =   0
            Y1              =   0
            Y2              =   11
         End
         Begin VB.Line lin_Bt 
            BorderColor     =   &H80000014&
            Index           =   15
            X1              =   0
            X2              =   22
            Y1              =   0
            Y2              =   0
         End
         Begin VB.Line lin_Bb 
            BorderColor     =   &H80000010&
            Index           =   15
            X1              =   0
            X2              =   23
            Y1              =   11
            Y2              =   11
         End
         Begin VB.Line lin_Br 
            BorderColor     =   &H80000010&
            Index           =   15
            X1              =   22
            X2              =   22
            Y1              =   0
            Y2              =   11
         End
      End
      Begin VB.PictureBox cmdTB 
         BorderStyle     =   0  'None
         HasDC           =   0   'False
         Height          =   330
         Index           =   14
         Left            =   1785
         Picture         =   "frmMain.frx":1A22
         ScaleHeight     =   22
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   23
         TabIndex        =   29
         ToolTipText     =   "Delete Selected Object"
         Top             =   3630
         Width           =   345
         Begin VB.Line lin_Br 
            BorderColor     =   &H80000010&
            Index           =   14
            X1              =   22
            X2              =   22
            Y1              =   0
            Y2              =   22
         End
         Begin VB.Line lin_Bb 
            BorderColor     =   &H80000010&
            Index           =   14
            X1              =   0
            X2              =   23
            Y1              =   21
            Y2              =   21
         End
         Begin VB.Line lin_Bt 
            BorderColor     =   &H80000014&
            Index           =   14
            X1              =   0
            X2              =   22
            Y1              =   0
            Y2              =   0
         End
         Begin VB.Line lin_bL 
            BorderColor     =   &H80000014&
            Index           =   14
            X1              =   0
            X2              =   0
            Y1              =   0
            Y2              =   21
         End
      End
      Begin VB.PictureBox cmdTB 
         BorderStyle     =   0  'None
         HasDC           =   0   'False
         Height          =   330
         Index           =   13
         Left            =   1785
         Picture         =   "frmMain.frx":1D92
         ScaleHeight     =   22
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   23
         TabIndex        =   28
         ToolTipText     =   "Delete Last Frame"
         Top             =   2505
         Width           =   345
         Begin VB.Line lin_Br 
            BorderColor     =   &H80000010&
            Index           =   13
            X1              =   22
            X2              =   22
            Y1              =   0
            Y2              =   22
         End
         Begin VB.Line lin_Bb 
            BorderColor     =   &H80000010&
            Index           =   13
            X1              =   0
            X2              =   23
            Y1              =   21
            Y2              =   21
         End
         Begin VB.Line lin_Bt 
            BorderColor     =   &H80000014&
            Index           =   13
            X1              =   0
            X2              =   22
            Y1              =   0
            Y2              =   0
         End
         Begin VB.Line lin_bL 
            BorderColor     =   &H80000014&
            Index           =   13
            X1              =   0
            X2              =   0
            Y1              =   0
            Y2              =   21
         End
      End
      Begin VB.ListBox lstObjects 
         Appearance      =   0  'Flat
         Height          =   810
         Left            =   75
         TabIndex        =   26
         Top             =   3615
         Width           =   1680
      End
      Begin VB.ListBox lstFrames 
         Appearance      =   0  'Flat
         Height          =   810
         Left            =   75
         TabIndex        =   24
         Top             =   2490
         Width           =   1680
      End
      Begin VB.PictureBox picZoomer 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   1920
         Left            =   135
         ScaleHeight     =   1920
         ScaleWidth      =   1920
         TabIndex        =   22
         Top             =   255
         Width           =   1920
      End
      Begin VB.Label lblTB2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Object List"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000009&
         Height          =   210
         Index           =   2
         Left            =   75
         TabIndex        =   27
         Top             =   3375
         Width           =   885
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H80000002&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   225
         Index           =   2
         Left            =   45
         Top             =   3375
         Width           =   2100
      End
      Begin VB.Label lblTB2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Frame List"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000009&
         Height          =   210
         Index           =   1
         Left            =   75
         TabIndex        =   25
         Top             =   2250
         Width           =   885
      End
      Begin VB.Label lblTB2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Zoom Box"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000009&
         Height          =   210
         Index           =   0
         Left            =   90
         TabIndex        =   23
         Top             =   0
         Width           =   825
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H80000002&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   225
         Index           =   0
         Left            =   45
         Top             =   0
         Width           =   2100
      End
      Begin VB.Line li5 
         BorderColor     =   &H80000014&
         Index           =   5
         X1              =   0
         X2              =   0
         Y1              =   0
         Y2              =   324
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H80000002&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   225
         Index           =   1
         Left            =   45
         Top             =   2250
         Width           =   2100
      End
   End
   Begin VB.PictureBox picDraw 
      Align           =   3  'Align Left
      BorderStyle     =   0  'None
      HasDC           =   0   'False
      Height          =   4425
      Left            =   0
      ScaleHeight     =   295
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   31
      TabIndex        =   7
      Top             =   435
      Width           =   465
      Begin VB.PictureBox picCurPX 
         Appearance      =   0  'Flat
         BackColor       =   &H00202020&
         ForeColor       =   &H80000008&
         Height          =   360
         Left            =   30
         ScaleHeight     =   330
         ScaleWidth      =   345
         TabIndex        =   40
         ToolTipText     =   "Color of Pixel..."
         Top             =   3180
         Width           =   375
      End
      Begin VB.PictureBox cmdTB 
         BorderStyle     =   0  'None
         Height          =   330
         Index           =   7
         Left            =   45
         MouseIcon       =   "frmMain.frx":2102
         Picture         =   "frmMain.frx":2254
         ScaleHeight     =   22
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   23
         TabIndex        =   16
         ToolTipText     =   "Color Sample - Push and Drag mouse..."
         Top             =   1890
         Width           =   345
         Begin VB.Line lin_Br 
            BorderColor     =   &H80000010&
            Index           =   7
            X1              =   22
            X2              =   22
            Y1              =   0
            Y2              =   22
         End
         Begin VB.Line lin_Bb 
            BorderColor     =   &H80000010&
            Index           =   7
            X1              =   0
            X2              =   23
            Y1              =   21
            Y2              =   21
         End
         Begin VB.Line lin_Bt 
            BorderColor     =   &H80000014&
            Index           =   7
            X1              =   0
            X2              =   22
            Y1              =   0
            Y2              =   0
         End
         Begin VB.Line lin_bL 
            BorderColor     =   &H80000014&
            Index           =   7
            X1              =   0
            X2              =   0
            Y1              =   0
            Y2              =   21
         End
      End
      Begin VB.PictureBox cmdTB 
         BorderStyle     =   0  'None
         Height          =   330
         Index           =   6
         Left            =   45
         Picture         =   "frmMain.frx":25D5
         ScaleHeight     =   22
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   23
         TabIndex        =   15
         ToolTipText     =   "Ellipse Tool"
         Top             =   1320
         Width           =   345
         Begin VB.Line lin_bL 
            BorderColor     =   &H80000014&
            Index           =   6
            X1              =   0
            X2              =   0
            Y1              =   0
            Y2              =   21
         End
         Begin VB.Line lin_Bt 
            BorderColor     =   &H80000014&
            Index           =   6
            X1              =   0
            X2              =   22
            Y1              =   0
            Y2              =   0
         End
         Begin VB.Line lin_Bb 
            BorderColor     =   &H80000010&
            Index           =   6
            X1              =   0
            X2              =   23
            Y1              =   21
            Y2              =   21
         End
         Begin VB.Line lin_Br 
            BorderColor     =   &H80000010&
            Index           =   6
            X1              =   22
            X2              =   22
            Y1              =   0
            Y2              =   22
         End
      End
      Begin VB.PictureBox cmdTB 
         BorderStyle     =   0  'None
         Height          =   330
         Index           =   5
         Left            =   45
         Picture         =   "frmMain.frx":293C
         ScaleHeight     =   22
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   23
         TabIndex        =   14
         ToolTipText     =   "Line Tool"
         Top             =   960
         Width           =   345
         Begin VB.Line lin_bL 
            BorderColor     =   &H80000014&
            Index           =   5
            X1              =   0
            X2              =   0
            Y1              =   0
            Y2              =   21
         End
         Begin VB.Line lin_Bt 
            BorderColor     =   &H80000014&
            Index           =   5
            X1              =   0
            X2              =   22
            Y1              =   0
            Y2              =   0
         End
         Begin VB.Line lin_Bb 
            BorderColor     =   &H80000010&
            Index           =   5
            X1              =   0
            X2              =   23
            Y1              =   21
            Y2              =   21
         End
         Begin VB.Line lin_Br 
            BorderColor     =   &H80000010&
            Index           =   5
            X1              =   22
            X2              =   22
            Y1              =   0
            Y2              =   22
         End
      End
      Begin VB.PictureBox cmdTB 
         BorderStyle     =   0  'None
         Height          =   330
         Index           =   4
         Left            =   45
         Picture         =   "frmMain.frx":2C96
         ScaleHeight     =   22
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   23
         TabIndex        =   13
         ToolTipText     =   "Box Tool"
         Top             =   600
         Width           =   345
         Begin VB.Line lin_bL 
            BorderColor     =   &H80000014&
            Index           =   4
            X1              =   0
            X2              =   0
            Y1              =   0
            Y2              =   21
         End
         Begin VB.Line lin_Bt 
            BorderColor     =   &H80000014&
            Index           =   4
            X1              =   0
            X2              =   22
            Y1              =   0
            Y2              =   0
         End
         Begin VB.Line lin_Bb 
            BorderColor     =   &H80000010&
            Index           =   4
            X1              =   0
            X2              =   23
            Y1              =   21
            Y2              =   21
         End
         Begin VB.Line lin_Br 
            BorderColor     =   &H80000010&
            Index           =   4
            X1              =   22
            X2              =   22
            Y1              =   0
            Y2              =   22
         End
      End
      Begin VB.PictureBox cmdTB 
         BackColor       =   &H00FFC080&
         BorderStyle     =   0  'None
         Height          =   330
         Index           =   3
         Left            =   45
         Picture         =   "frmMain.frx":2FFC
         ScaleHeight     =   22
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   23
         TabIndex        =   11
         ToolTipText     =   "Polygon Tool"
         Top             =   240
         Width           =   345
         Begin VB.Line lin_Br 
            BorderColor     =   &H80000010&
            Index           =   3
            X1              =   22
            X2              =   22
            Y1              =   0
            Y2              =   22
         End
         Begin VB.Line lin_Bb 
            BorderColor     =   &H80000010&
            Index           =   3
            X1              =   0
            X2              =   23
            Y1              =   21
            Y2              =   21
         End
         Begin VB.Line lin_Bt 
            BorderColor     =   &H80000014&
            Index           =   3
            X1              =   0
            X2              =   22
            Y1              =   0
            Y2              =   0
         End
         Begin VB.Line lin_bL 
            BorderColor     =   &H80000014&
            Index           =   3
            X1              =   0
            X2              =   0
            Y1              =   0
            Y2              =   21
         End
      End
      Begin VB.PictureBox cmdOutLine 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         ForeColor       =   &H80000008&
         HasDC           =   0   'False
         Height          =   720
         Left            =   30
         ScaleHeight     =   690
         ScaleWidth      =   345
         TabIndex        =   9
         ToolTipText     =   "Outline Color"
         Top             =   2340
         Width           =   375
         Begin VB.PictureBox cmdBG 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H80000008&
            Height          =   570
            Left            =   60
            ScaleHeight     =   540
            ScaleWidth      =   195
            TabIndex        =   10
            ToolTipText     =   "Fill Color"
            Top             =   60
            Width           =   225
         End
      End
      Begin VB.Line linSep 
         BorderColor     =   &H80000014&
         Index           =   1
         X1              =   3
         X2              =   25
         Y1              =   118
         Y2              =   118
      End
      Begin VB.Line linSep 
         BorderColor     =   &H80000010&
         Index           =   0
         X1              =   3
         X2              =   25
         Y1              =   117
         Y2              =   117
      End
      Begin VB.Label lblTB 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Draw"
         ForeColor       =   &H80000009&
         Height          =   195
         Left            =   15
         TabIndex        =   12
         Top             =   -15
         Width           =   375
      End
      Begin VB.Shape shpTB 
         BackColor       =   &H80000002&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   210
         Left            =   -30
         Top             =   -15
         Width           =   480
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000015&
         X1              =   30
         X2              =   30
         Y1              =   295
         Y2              =   -1
      End
   End
   Begin VB.PictureBox picMain 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00202020&
      BorderStyle     =   0  'None
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   4320
      Left            =   525
      MousePointer    =   2  'Cross
      ScaleHeight     =   288
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   388
      TabIndex        =   6
      Top             =   480
      Width           =   5820
      Begin VB.Line linTemp 
         BorderColor     =   &H0000FF00&
         DrawMode        =   6  'Mask Pen Not
         Visible         =   0   'False
         X1              =   166
         X2              =   277
         Y1              =   124
         Y2              =   200
      End
      Begin VB.Shape shpTemp 
         BorderColor     =   &H0000FF00&
         BorderStyle     =   6  'Inside Solid
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   495
         Left            =   3900
         Top             =   3225
         Visible         =   0   'False
         Width           =   1035
      End
      Begin VB.Shape shpMark 
         BorderColor     =   &H000000FF&
         Height          =   15
         Left            =   2100
         Top             =   1650
         Width           =   15
      End
      Begin VB.Line linBCH_v 
         DrawMode        =   6  'Mask Pen Not
         X1              =   -1
         X2              =   -1
         Y1              =   0
         Y2              =   288
      End
      Begin VB.Line linBCH_h 
         DrawMode        =   6  'Mask Pen Not
         X1              =   0
         X2              =   388
         Y1              =   -1
         Y2              =   -1
      End
   End
   Begin VB.PictureBox picSBar 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      HasDC           =   0   'False
      Height          =   345
      Left            =   0
      ScaleHeight     =   23
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   572
      TabIndex        =   4
      Top             =   4860
      Width           =   8580
      Begin VB.PictureBox picProg 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   3960
         ScaleHeight     =   255
         ScaleWidth      =   2400
         TabIndex        =   36
         Top             =   60
         Visible         =   0   'False
         Width           =   2400
      End
      Begin VB.Label lblTool 
         AutoSize        =   -1  'True
         Caption         =   "Polygon Tool"
         Height          =   195
         Left            =   4020
         TabIndex        =   35
         Top             =   90
         Width           =   1575
      End
      Begin VB.Label lblCoords 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "0, 0"
         Height          =   195
         Left            =   8205
         TabIndex        =   34
         Top             =   90
         Width           =   270
      End
      Begin VB.Line li5 
         BorderColor     =   &H80000014&
         Index           =   4
         X1              =   30
         X2              =   429
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Line linSB 
         BorderColor     =   &H80000014&
         Index           =   7
         X1              =   424
         X2              =   424
         Y1              =   3
         Y2              =   21
      End
      Begin VB.Line linSB 
         BorderColor     =   &H80000010&
         Index           =   6
         X1              =   263
         X2              =   263
         Y1              =   3
         Y2              =   21
      End
      Begin VB.Line linSB 
         BorderColor     =   &H80000014&
         Index           =   5
         X1              =   263
         X2              =   425
         Y1              =   21
         Y2              =   21
      End
      Begin VB.Line linSB 
         BorderColor     =   &H80000010&
         Index           =   4
         X1              =   264
         X2              =   425
         Y1              =   3
         Y2              =   3
      End
      Begin VB.Line linSB 
         BorderColor     =   &H80000014&
         Index           =   3
         X1              =   569
         X2              =   569
         Y1              =   3
         Y2              =   21
      End
      Begin VB.Line linSB 
         BorderColor     =   &H80000010&
         Index           =   2
         X1              =   433
         X2              =   433
         Y1              =   3
         Y2              =   21
      End
      Begin VB.Line linSB 
         BorderColor     =   &H80000014&
         Index           =   1
         X1              =   433
         X2              =   570
         Y1              =   21
         Y2              =   21
      End
      Begin VB.Line linSB 
         BorderColor     =   &H80000010&
         Index           =   0
         X1              =   433
         X2              =   570
         Y1              =   3
         Y2              =   3
      End
      Begin VB.Label lblStat 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "        "
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
         Left            =   75
         TabIndex        =   5
         Top             =   75
         Width           =   360
      End
   End
   Begin VB.PictureBox picTBholder 
      Align           =   1  'Align Top
      BorderStyle     =   0  'None
      HasDC           =   0   'False
      Height          =   435
      Left            =   0
      ScaleHeight     =   29
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   572
      TabIndex        =   0
      Top             =   0
      Width           =   8580
      Begin VB.PictureBox cmdTB 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         HasDC           =   0   'False
         Height          =   330
         Index           =   12
         Left            =   2850
         Picture         =   "frmMain.frx":3365
         ScaleHeight     =   22
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   23
         TabIndex        =   21
         ToolTipText     =   "Done Frame..."
         Top             =   45
         Width           =   345
         Begin VB.Line lin_bL 
            BorderColor     =   &H80000014&
            Index           =   12
            X1              =   0
            X2              =   0
            Y1              =   0
            Y2              =   21
         End
         Begin VB.Line lin_Bt 
            BorderColor     =   &H80000014&
            Index           =   12
            X1              =   0
            X2              =   22
            Y1              =   0
            Y2              =   0
         End
         Begin VB.Line lin_Bb 
            BorderColor     =   &H80000010&
            Index           =   12
            X1              =   0
            X2              =   23
            Y1              =   21
            Y2              =   21
         End
         Begin VB.Line lin_Br 
            BorderColor     =   &H80000010&
            Index           =   12
            X1              =   22
            X2              =   22
            Y1              =   0
            Y2              =   22
         End
      End
      Begin VB.PictureBox cmdTB 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         HasDC           =   0   'False
         Height          =   330
         Index           =   11
         Left            =   2490
         Picture         =   "frmMain.frx":3724
         ScaleHeight     =   22
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   23
         TabIndex        =   20
         ToolTipText     =   "Done Polygon..."
         Top             =   45
         Width           =   345
         Begin VB.Line lin_bL 
            BorderColor     =   &H80000014&
            Index           =   11
            X1              =   0
            X2              =   0
            Y1              =   0
            Y2              =   21
         End
         Begin VB.Line lin_Bt 
            BorderColor     =   &H80000014&
            Index           =   11
            X1              =   0
            X2              =   22
            Y1              =   0
            Y2              =   0
         End
         Begin VB.Line lin_Bb 
            BorderColor     =   &H80000010&
            Index           =   11
            X1              =   0
            X2              =   23
            Y1              =   21
            Y2              =   21
         End
         Begin VB.Line lin_Br 
            BorderColor     =   &H80000010&
            Index           =   11
            X1              =   22
            X2              =   22
            Y1              =   0
            Y2              =   22
         End
      End
      Begin VB.PictureBox cmdTB 
         BorderStyle     =   0  'None
         HasDC           =   0   'False
         Height          =   330
         Index           =   10
         Left            =   2130
         Picture         =   "frmMain.frx":3ABB
         ScaleHeight     =   22
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   23
         TabIndex        =   19
         ToolTipText     =   "Coordinate List Input..."
         Top             =   45
         Width           =   345
         Begin VB.Line lin_bL 
            BorderColor     =   &H80000014&
            Index           =   10
            X1              =   0
            X2              =   0
            Y1              =   0
            Y2              =   21
         End
         Begin VB.Line lin_Bt 
            BorderColor     =   &H80000014&
            Index           =   10
            X1              =   0
            X2              =   22
            Y1              =   0
            Y2              =   0
         End
         Begin VB.Line lin_Bb 
            BorderColor     =   &H80000010&
            Index           =   10
            X1              =   0
            X2              =   23
            Y1              =   21
            Y2              =   21
         End
         Begin VB.Line lin_Br 
            BorderColor     =   &H80000010&
            Index           =   10
            X1              =   22
            X2              =   22
            Y1              =   0
            Y2              =   22
         End
      End
      Begin VB.PictureBox cmdTB 
         BorderStyle     =   0  'None
         HasDC           =   0   'False
         Height          =   330
         Index           =   8
         Left            =   1275
         Picture         =   "frmMain.frx":3E5B
         ScaleHeight     =   22
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   23
         TabIndex        =   18
         ToolTipText     =   "Load picture into background for tracing..."
         Top             =   45
         Width           =   345
         Begin VB.Line lin_Br 
            BorderColor     =   &H80000010&
            Index           =   8
            X1              =   22
            X2              =   22
            Y1              =   0
            Y2              =   22
         End
         Begin VB.Line lin_Bb 
            BorderColor     =   &H80000010&
            Index           =   8
            X1              =   0
            X2              =   23
            Y1              =   21
            Y2              =   21
         End
         Begin VB.Line lin_Bt 
            BorderColor     =   &H80000014&
            Index           =   8
            X1              =   0
            X2              =   22
            Y1              =   0
            Y2              =   0
         End
         Begin VB.Line lin_bL 
            BorderColor     =   &H80000014&
            Index           =   8
            X1              =   0
            X2              =   0
            Y1              =   0
            Y2              =   21
         End
      End
      Begin VB.PictureBox cmdTB 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         HasDC           =   0   'False
         Height          =   330
         Index           =   9
         Left            =   1635
         Picture         =   "frmMain.frx":41E7
         ScaleHeight     =   22
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   23
         TabIndex        =   17
         ToolTipText     =   "Play / Preview Animation"
         Top             =   45
         Width           =   345
         Begin VB.Line lin_Br 
            BorderColor     =   &H80000010&
            Index           =   9
            X1              =   22
            X2              =   22
            Y1              =   0
            Y2              =   22
         End
         Begin VB.Line lin_Bb 
            BorderColor     =   &H80000010&
            Index           =   9
            X1              =   0
            X2              =   23
            Y1              =   21
            Y2              =   21
         End
         Begin VB.Line lin_Bt 
            BorderColor     =   &H80000014&
            Index           =   9
            X1              =   0
            X2              =   22
            Y1              =   0
            Y2              =   0
         End
         Begin VB.Line lin_bL 
            BorderColor     =   &H80000014&
            Index           =   9
            X1              =   0
            X2              =   0
            Y1              =   0
            Y2              =   21
         End
      End
      Begin VB.PictureBox cmdTB 
         BorderStyle     =   0  'None
         HasDC           =   0   'False
         Height          =   330
         Index           =   0
         Left            =   60
         Picture         =   "frmMain.frx":458C
         ScaleHeight     =   22
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   23
         TabIndex        =   3
         ToolTipText     =   "New Polygon Movie..."
         Top             =   45
         Width           =   345
         Begin VB.Line lin_bL 
            BorderColor     =   &H80000014&
            Index           =   0
            X1              =   0
            X2              =   0
            Y1              =   0
            Y2              =   21
         End
         Begin VB.Line lin_Bt 
            BorderColor     =   &H80000014&
            Index           =   0
            X1              =   0
            X2              =   22
            Y1              =   0
            Y2              =   0
         End
         Begin VB.Line lin_Bb 
            BorderColor     =   &H80000010&
            Index           =   0
            X1              =   0
            X2              =   23
            Y1              =   21
            Y2              =   21
         End
         Begin VB.Line lin_Br 
            BorderColor     =   &H80000010&
            Index           =   0
            X1              =   22
            X2              =   22
            Y1              =   0
            Y2              =   22
         End
      End
      Begin VB.PictureBox cmdTB 
         BorderStyle     =   0  'None
         HasDC           =   0   'False
         Height          =   330
         Index           =   1
         Left            =   420
         Picture         =   "frmMain.frx":4912
         ScaleHeight     =   22
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   23
         TabIndex        =   2
         ToolTipText     =   "Open a Polygon Movie file..."
         Top             =   45
         Width           =   345
         Begin VB.Line lin_Br 
            BorderColor     =   &H80000010&
            Index           =   1
            X1              =   22
            X2              =   22
            Y1              =   0
            Y2              =   22
         End
         Begin VB.Line lin_Bb 
            BorderColor     =   &H80000010&
            Index           =   1
            X1              =   0
            X2              =   23
            Y1              =   21
            Y2              =   21
         End
         Begin VB.Line lin_Bt 
            BorderColor     =   &H80000014&
            Index           =   1
            X1              =   0
            X2              =   22
            Y1              =   0
            Y2              =   0
         End
         Begin VB.Line lin_bL 
            BorderColor     =   &H80000014&
            Index           =   1
            X1              =   0
            X2              =   0
            Y1              =   0
            Y2              =   21
         End
      End
      Begin VB.PictureBox cmdTB 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         HasDC           =   0   'False
         Height          =   330
         Index           =   2
         Left            =   780
         Picture         =   "frmMain.frx":4CA7
         ScaleHeight     =   22
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   23
         TabIndex        =   1
         ToolTipText     =   "Save current Polygon Movie..."
         Top             =   45
         Width           =   345
         Begin VB.Line lin_Br 
            BorderColor     =   &H80000010&
            Index           =   2
            X1              =   22
            X2              =   22
            Y1              =   0
            Y2              =   22
         End
         Begin VB.Line lin_Bb 
            BorderColor     =   &H80000010&
            Index           =   2
            X1              =   0
            X2              =   23
            Y1              =   21
            Y2              =   21
         End
         Begin VB.Line lin_Bt 
            BorderColor     =   &H80000014&
            Index           =   2
            X1              =   0
            X2              =   22
            Y1              =   0
            Y2              =   0
         End
         Begin VB.Line lin_bL 
            BorderColor     =   &H80000014&
            Index           =   2
            X1              =   0
            X2              =   0
            Y1              =   0
            Y2              =   21
         End
      End
      Begin VB.Timer tmrHov 
         Interval        =   500
         Left            =   -360
         Top             =   285
      End
      Begin VB.Label lblStat3 
         AutoSize        =   -1  'True
         Caption         =   "Point Count: 0"
         Height          =   195
         Left            =   6975
         TabIndex        =   39
         Top             =   105
         Width           =   1005
      End
      Begin VB.Label lblStat2 
         AutoSize        =   -1  'True
         Caption         =   "Object Count: 0"
         Height          =   195
         Left            =   5265
         TabIndex        =   38
         Top             =   105
         Width           =   1110
      End
      Begin VB.Label lblStat1 
         AutoSize        =   -1  'True
         Caption         =   "Frame Count: 0"
         Height          =   195
         Left            =   3555
         TabIndex        =   37
         Top             =   105
         Width           =   1080
      End
      Begin VB.Line linSB 
         BorderColor     =   &H80000010&
         Index           =   19
         X1              =   235
         X2              =   235
         Y1              =   5
         Y2              =   23
      End
      Begin VB.Line linSB 
         BorderColor     =   &H80000014&
         Index           =   18
         X1              =   235
         X2              =   341
         Y1              =   23
         Y2              =   23
      End
      Begin VB.Line linSB 
         BorderColor     =   &H80000010&
         Index           =   17
         X1              =   235
         X2              =   341
         Y1              =   5
         Y2              =   5
      End
      Begin VB.Line linSB 
         BorderColor     =   &H80000014&
         Index           =   16
         X1              =   340
         X2              =   340
         Y1              =   5
         Y2              =   23
      End
      Begin VB.Line linSB 
         BorderColor     =   &H80000014&
         Index           =   12
         X1              =   454
         X2              =   454
         Y1              =   5
         Y2              =   23
      End
      Begin VB.Line linSB 
         BorderColor     =   &H80000010&
         Index           =   15
         X1              =   349
         X2              =   455
         Y1              =   5
         Y2              =   5
      End
      Begin VB.Line linSB 
         BorderColor     =   &H80000014&
         Index           =   14
         X1              =   349
         X2              =   455
         Y1              =   23
         Y2              =   23
      End
      Begin VB.Line linSB 
         BorderColor     =   &H80000010&
         Index           =   13
         X1              =   349
         X2              =   349
         Y1              =   5
         Y2              =   23
      End
      Begin VB.Line linSB 
         BorderColor     =   &H80000010&
         Index           =   11
         X1              =   463
         X2              =   569
         Y1              =   5
         Y2              =   5
      End
      Begin VB.Line linSB 
         BorderColor     =   &H80000014&
         Index           =   10
         X1              =   463
         X2              =   569
         Y1              =   23
         Y2              =   23
      End
      Begin VB.Line linSB 
         BorderColor     =   &H80000010&
         Index           =   9
         X1              =   463
         X2              =   463
         Y1              =   5
         Y2              =   23
      End
      Begin VB.Line linSB 
         BorderColor     =   &H80000014&
         Index           =   8
         X1              =   568
         X2              =   568
         Y1              =   5
         Y2              =   23
      End
      Begin VB.Line li6 
         BorderColor     =   &H80000010&
         Index           =   1
         X1              =   0
         X2              =   572
         Y1              =   27
         Y2              =   27
      End
      Begin VB.Line li4 
         BorderColor     =   &H80000014&
         X1              =   0
         X2              =   572
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Line li5 
         BorderColor     =   &H80000010&
         Index           =   0
         X1              =   79
         X2              =   79
         Y1              =   3
         Y2              =   25
      End
      Begin VB.Line li5 
         BorderColor     =   &H80000014&
         Index           =   1
         X1              =   80
         X2              =   80
         Y1              =   3
         Y2              =   25
      End
      Begin VB.Line li6 
         BorderColor     =   &H80000015&
         Index           =   0
         X1              =   30
         X2              =   428
         Y1              =   28
         Y2              =   28
      End
      Begin VB.Line li5 
         BorderColor     =   &H80000014&
         Index           =   2
         X1              =   137
         X2              =   137
         Y1              =   3
         Y2              =   25
      End
      Begin VB.Line li5 
         BorderColor     =   &H80000010&
         Index           =   3
         X1              =   136
         X2              =   136
         Y1              =   3
         Y2              =   25
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000014&
         X1              =   428
         X2              =   428
         Y1              =   29
         Y2              =   27
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuNew 
         Caption         =   "&New"
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "&Open..."
      End
      Begin VB.Menu mnuSave 
         Caption         =   "&Save"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuSaveAs 
         Caption         =   "Save &As..."
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuDash1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLoadBGPicture 
         Caption         =   "Load Picture into Background for Tracing..."
      End
      Begin VB.Menu mnuExportBMP 
         Caption         =   "Export &BMP"
      End
      Begin VB.Menu mnuDash4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit..."
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuFinishPoly 
         Caption         =   "Finish / Close Polygon"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuDoneFrame 
         Caption         =   "Done Current Frame..."
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuDash2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFillColor 
         Caption         =   "Fill Color..."
      End
      Begin VB.Menu mnuOutlineColor 
         Caption         =   "Outline Color... [Applies to entire animation]"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuPreview 
         Caption         =   "Preview Polygon Movie"
      End
      Begin VB.Menu mnuAnimationProps 
         Caption         =   "Animation Properties"
      End
      Begin VB.Menu mnuLargeCH 
         Caption         =   "Show Large CrossHairs"
         Checked         =   -1  'True
         Shortcut        =   ^W
      End
   End
   Begin VB.Menu mnuToolS 
      Caption         =   "&Tools"
      Begin VB.Menu mnuToolSel 
         Caption         =   "Polygon Tool"
         Checked         =   -1  'True
         Index           =   0
      End
      Begin VB.Menu mnuToolSel 
         Caption         =   "Square Tool"
         Index           =   1
      End
      Begin VB.Menu mnuToolSel 
         Caption         =   "Line Tool"
         Index           =   2
      End
      Begin VB.Menu mnuToolSel 
         Caption         =   "Ellipse Tool"
         Index           =   3
      End
      Begin VB.Menu mnuDash 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCoordList 
         Caption         =   "Coordinate List Input..."
      End
      Begin VB.Menu mnuAddPointManually 
         Caption         =   "Manually Add Point..."
      End
      Begin VB.Menu mnuPointReduce 
         Caption         =   "Point Reduction..."
      End
      Begin VB.Menu mnuDash3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAutoApply 
         Caption         =   "Auto-Apply Shapes"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetTickCount Lib "kernel32" () As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long

Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long


Private FD As New cFileDialog 'save/open dialog

Private CurPFrame As Long
Public FirstPoint As Boolean
Private tmrPrev As Boolean
Private CurTool As Byte
Private LargeCH As Boolean
Private PHC As POINTAPI
Private SaveFileName As String

Private Sub cmdBG_DblClick()
On Error Resume Next
 cmdBG.BackColor = ShowColor(Me)
 picMain.FillColor = cmdBG.BackColor
    CurShape.PolyColor = cmdBG.BackColor
    shpTemp.FillColor = cmdBG.BackColor
End Sub

Private Sub cmdOutLine_DblClick()
On Error Resume Next
cmdOutLine.BackColor = ShowColor(Me)
    picMain.ForeColor = cmdOutLine.BackColor
    TempPANI.OutLineColor = cmdOutLine.BackColor
    shpTemp.BorderColor = cmdOutLine.BackColor
End Sub

Private Sub cmdTB_Click(Index As Integer)
Select Case Index
   Case 0 'New File button...
    If MsgBox("Are you sure you want to create a new file?" & vbNewLine & "WARNING: Any unsaved work on this animation will be lost!", vbYesNo + vbExclamation, "Create New File?") = vbYes Then _
       NewFile
   Case 1 'Open File button
    OpenFile
   Case 2 'save file, save file as...
    If Len(SaveFileName) = 0 Then
       SaveFile
       Else
       SavePan SaveFileName
    End If
   Case 3 To 6 'SWITCH TOOL TYPE...
      mnuToolSel(CurTool).Checked = False
      cmdTB(CurTool + 3).BackColor = &H8000000F
      cmdTB(Index).BackColor = &HFFC080
      CurTool = (Index - 3)
      mnuToolSel(CurTool).Checked = True
      CurShape.PolyType = CurTool
      FirstPoint = False
      cmdTB(11).Enabled = False
      cmdTB(12).Enabled = False
      lblTool.Caption = cmdTB(Index).ToolTipText
    Case 8
     LoadBGpic
    Case 9 'preview / play animation
    
         If tmrPrev = False Then
            tmrPrev = True
                Call PrevTimer
            Else
                 tmrPrev = False
         End If
         
    Case 11 'finish shape
        Call DoneObject
    Case 12 'finish frame
        Call DoneFrame
    Case 13 'delete last frame
        If lstFrames.ListCount > 0 Then
           lstFrames.RemoveItem (lstFrames.ListCount - 1)
           TempPANI.FrameCount = TempPANI.FrameCount - 1
    
    If TempPANI.FrameCount > 0 Then
       ReDim Preserve TempPANI.Polys(1 To TempPANI.FrameCount)
    Else
       Erase TempPANI.Polys
    End If
    
        End If
   End Select
End Sub

Private Sub cmdTB_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
 If lin_bL(Index).BorderColor = &H80000014 Then
    'left and top button border
    lin_bL(Index).BorderColor = &H80000010
    lin_Bt(Index).BorderColor = &H80000010
       
       'bottom and right button border
       lin_Br(Index).BorderColor = &H80000014
       lin_Bb(Index).BorderColor = &H80000014
 End If
 If Index = 7 Then cmdTB(7).MousePointer = 99
End Sub

Private Sub cmdTB_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
'sample color from screen...
Dim P As POINTAPI

If x >= 0 And x <= cmdTB(Index).ScaleWidth And _
   y >= 0 And y <= cmdTB(Index).ScaleHeight Then
 'make visible
 If lin_bL(Index).Visible = False Then
    Call Htb_CLEAR
    lin_bL(Index).Visible = True
    lin_Br(Index).Visible = True
    lin_Bt(Index).Visible = True
    lin_Bb(Index).Visible = True
 End If
   Else
   'make invisible
 If lin_bL(Index).Visible = True Then
    Call Htb_CLEAR
    lin_bL(Index).Visible = False
    lin_Br(Index).Visible = False
    lin_Bt(Index).Visible = False
    lin_Bb(Index).Visible = False
 End If
End If
If lblStat.Caption <> cmdTB(Index).ToolTipText Then _
      lblStat.Caption = cmdTB(Index).ToolTipText


 If Index = 7 Then 'sample color when dragging...
  If Button = 0 Then Exit Sub
   GetCursorPos P
    Select Case Button
        Case Is = 1
            cmdBG.BackColor = GetPixel(GetDC(0), P.x, P.y)
            picMain.FillColor = cmdBG.BackColor
            CurShape.PolyColor = cmdBG.BackColor
            shpTemp.FillColor = cmdBG.BackColor
        Case Is = 2
            cmdOutLine.BackColor = GetPixel(GetDC(0), P.x, P.y)
            linTemp.BorderColor = cmdOutLine.BackColor
            shpTemp.BorderColor = cmdOutLine.BackColor
     End Select
          
 End If
End Sub

Private Sub cmdTB_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
 If lin_bL(Index).BorderColor = &H80000010 Then
    'left and top button border
    lin_bL(Index).BorderColor = &H80000014
    lin_Bt(Index).BorderColor = &H80000014
       
       'bottom and right button border
       lin_Br(Index).BorderColor = &H80000010
       lin_Bb(Index).BorderColor = &H80000010
 End If
 If Index = 7 Then cmdTB(7).MousePointer = 0
End Sub

Private Sub Form_load()
Call Htb_CLEAR
Call PrepColorDlg 'prepare the color dialog
   'make default fill color of cur shape struct white
   CurShape.PolyColor = vbWhite
   ReDim CurShape.PolyPnt(1 To 4096)
   LargeCH = True
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Call Htb_CLEAR
'clear button when mouse moves over the
'form
End Sub

Private Sub Form_Terminate()
tmrPrev = False
End Sub
Private Sub Form_Unload(Cancel As Integer)
If MsgBox("Are you sure you want to exit?", vbYesNo + vbQuestion) = vbNo Then
  Cancel = 1
  Else

Set FD = Nothing
Unload frmAbout
Unload frmAddPoint
Unload frmPointReduce
End If
End Sub

Private Sub lstFrames_Click()
Dim I As Long, T As Long
picMain.Cls
DoEvents
DrawFrame (lstFrames.ListIndex + 1), TempPANI, picMain.hdc

lblStat1.Caption = "Frame Count: " & TempPANI.FrameCount
lblStat2.Caption = "Object Count: " & TempPANI.Polys(lstFrames.ListIndex + 1).PolyCount

For I = 1 To TempPANI.Polys(lstFrames.ListIndex + 1).PolyCount
 T = T + TempPANI.Polys(lstFrames.ListIndex + 1).PolyShp(I).PntCount
Next
lblStat3.Caption = "Point Count: " & T
picMain.Refresh
End Sub

Private Sub lstFrames_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Call Htb_CLEAR
End Sub

Private Sub lstObjects_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Call Htb_CLEAR
End Sub


Private Sub mnuAbout_Click()
frmAbout.Show 1
End Sub

Private Sub mnuAddPointManually_Click()
frmAddPoint.Show
End Sub

Private Sub mnuAnimationProps_Click()
On Error GoTo ErrOut:
Dim FS As Long 'filesize
Dim PC As Long 'point count
Dim OC As Long 'object count
Dim FN As String 'filename
Dim LT As Single 'load time...
Dim T1 As Long, T2 As Long 'timer..
'counters
Dim I As Long
Dim Z As Long
Dim InfoString As String
lblStat.Caption = "Evaluating File..."

SavePan (App.Path & "\tmp~1")

FS = FileLen(App.Path & "\tmp~1")
FN = SaveFileName
If FN = "" Then FN = "untitled"

For I = 1 To TempPANI.FrameCount
  OC = OC + TempPANI.Polys(I).PolyCount
  For Z = 1 To TempPANI.Polys(1).PolyCount
      PC = PC + TempPANI.Polys(I).PolyShp(Z).PntCount
  Next
Next

T1 = GetTickCount
LoadPan (App.Path & "\tmp~1")
T2 = GetTickCount

LT = Round((T2 - T1) / 1000, 3)

InfoString = "File Information" & vbNewLine & vbNewLine & _
"Filename:          " & FN & vbNewLine & _
"Filesize:          " & Round((FS / 1024), 3) & " Kb [" & FS & " bytes]" & vbNewLine & _
vbNewLine & _
"Total Frames:      " & TempPANI.FrameCount & vbNewLine & _
"Total Polygons:    " & OC & vbNewLine & _
"Total Points:      " & PC & vbNewLine & vbNewLine & _
"File Load Time:    " & LT & " Seconds."

'delete temp file...
Kill App.Path & "\tmp~1"

'display information
MsgBox InfoString, vbOKOnly, "File Information"

ErrOut:
End Sub

Private Sub mnuAutoApply_Click()
 If mnuAutoApply.Checked = True Then
   mnuAutoApply.Checked = False
  Else
   mnuAutoApply.Checked = True
 End If
End Sub


Private Sub mnuExit_Click()
If MsgBox("Are you sure you want to exit? Any unsaved changes made to the current file will be lost.", vbYesNo + vbExclamation, "Exit Polygon Movie Editor?") = vbYes Then
   Unload frmMain
   End
End If
End Sub

Private Sub mnuExportBMP_Click()
Call ExportBMP
End Sub

Private Sub mnuFillColor_Click()
Call cmdBG_DblClick
End Sub

Private Sub mnuLargeCH_Click()
If mnuLargeCH.Checked = False Then
   mnuLargeCH.Checked = True
        linBCH_h.Visible = True
        linBCH_v.Visible = True
        LargeCH = True
            Else
    mnuLargeCH.Checked = False
      linBCH_h.Visible = False
      linBCH_v.Visible = False
      LargeCH = False
End If
End Sub

Private Sub mnuLoadBGPicture_Click()
LoadBGpic
End Sub

Private Sub mnuNew_Click()
If MsgBox("Are you sure you want to create a new file?" & vbNewLine & "WARNING: Any unsaved work on this animation will be lost!", vbYesNo + vbExclamation, "Create New File?") = vbYes Then _
   NewFile
End Sub

Private Sub mnuOpen_Click()
    OpenFile
End Sub

Private Sub mnuOutlineColor_Click()
cmdOutLine_DblClick
End Sub

Private Sub mnuPointReduce_Click()
frmPointReduce.Show 1
End Sub

Private Sub mnuSave_Click()
SavePan SaveFileName
End Sub

Private Sub mnuSaveAs_Click()
SaveFile
End Sub

Private Sub mnuToolSel_Click(Index As Integer)
    mnuToolSel(CurTool).Checked = False
    cmdTB(CurTool + 3).BackColor = &H8000000F
    mnuToolSel(Index).Checked = True
    CurTool = Index
    cmdTB(Index + 3).BackColor = &HFFC080
    lblTool.Caption = cmdTB(Index + 3).ToolTipText
    CurShape.PolyType = CurTool
    FirstPoint = False
    cmdTB(11).Enabled = False
    cmdTB(12).Enabled = False
End Sub

Private Sub picDraw_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Call Htb_CLEAR
End Sub

Private Sub picRTB_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Call Htb_CLEAR
End Sub


Private Sub picSBar_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Call Htb_CLEAR
End Sub

Private Sub picTBholder_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Call Htb_CLEAR 'clear hover buttons when
'its over the toolbar, and not over any
'of the buttons
End Sub

Private Sub tmrHov_Timer()
'clear hover buttons if mouse leaves form
GetCursorPos PHC
If PHC.x < (Me.Left / Screen.TwipsPerPixelX) Or _
   PHC.x > (Me.Left / Screen.TwipsPerPixelX) + (Me.Width / Screen.TwipsPerPixelX) Or _
   PHC.y < (Me.Top / Screen.TwipsPerPixelY) + 38 Or _
   PHC.y > (Me.Top / Screen.TwipsPerPixelY) + (Me.Height / Screen.TwipsPerPixelY) _
   Then _
       Call Htb_CLEAR
End Sub

Public Sub Htb_CLEAR()
'clear hover buttons
Dim I As Long
 For I = cmdTB.LBound To cmdTB.UBound
  If lin_bL(I).Visible = True Then
    lin_bL(I).Visible = False
    lin_Br(I).Visible = False
    lin_Bt(I).Visible = False
    lin_Bb(I).Visible = False
  End If
 Next
 
 'clear status bar
 If Len(lblStat.Caption) > 0 Then _
        lblStat.Caption = vbNullString
End Sub

'===================================
'=====================================
'=======================================
'
'THIS IS WHERE CODE FROM AN OLDER VERSION
'OF THIS PROGRAM STARTS. I RE-WROTE THE ENTIRE
'EDITOR TO MAKE IT MORE FUNCTIONAL AND USER-FRIENDLY
'
'=======================================
'=====================================
'===================================

Private Sub picMain_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim P As POINTAPI
Select Case CurTool
    Case Is = 0  '==============
                 '   Polygon
                 '==============
        If FirstPoint = False Then
            'set up our line object
        With linTemp
            .Visible = True
            .X1 = x
            .Y1 = y
            .X2 = x
            .Y2 = y
        End With
            'loose the shape object we dont need it
            'cuz we're drawing polygons baby! yea!
        shpTemp.Visible = False
        
            MoveToEx picMain.hdc, x, y, P
            
            'add point to polygon here...
              CurShape.PntCount = CurShape.PntCount + 1
              cmdTB(11).Enabled = True
              CurShape.PolyPnt(CurShape.PntCount).x = x
              CurShape.PolyPnt(CurShape.PntCount).y = y
        Else
             'we aint adding a point here
             'cuz this could be the final point
             'and it will be made when they finish the
             'polygon
             
             'so make the line connect from
             'the origin point, to the
             'new point...
            With linTemp
                .Visible = True
                .X2 = x
                .Y2 = y
            End With
        End If
    Case Is = 1 '=============
                '  rectangle
                '=============
    
    'we only put the mouse down once when
    'drawing rectangles so we dont need to worry
    'about the FirstPoint flag...
    
        'we dont need the line no more
            linTemp.Visible = False
        
            'make the shape object a rect and make it visible
            With shpTemp
                .Shape = 0 'make rectangle
                .Left = x
                .Top = y
                .Width = 0
                .Height = 0
                .Visible = True
            End With
            
            'add data that we can to the temp shape struct
            
            'we only have 2 points in a rect
            'representing the opposite corners
            CurShape.PntCount = 2
            'add top-left corner x and y coords
            CurShape.PolyPnt(1).x = x
            CurShape.PolyPnt(1).y = y
            
    Case Is = 2 '=============
                '    line
                '=============
            
            'dont need the shape object
            shpTemp.Visible = False
            
            'set up the line object for us...
            With linTemp
               .X1 = x
               .X2 = x
               .Y1 = y
               .Y2 = y
               .Visible = True
            End With
            
            'add data that we can to the temp shape struct
            
            'we only have 2 points in a LINE
            'representing the 2 points at either end...
            CurShape.PntCount = 2
            'add top-left corner x and y coords
            CurShape.PolyPnt(1).x = x
            CurShape.PolyPnt(1).y = y
    
    Case Is = 3 '=============
                '   ELLIPS
                '=============
                
                'dont need line object
                linTemp.Visible = False
                
                With shpTemp
                 .Left = x
                 .Top = y
                 .Width = 0
                 .Height = 0
                 .Shape = 2 'oval
                 .Visible = True
                End With
            
             'we only have 2 points in an ellips
             'representing the 2 corners of its
             'rectangular frame...
            CurShape.PntCount = 2
            
            'add top-left corner x and y coords
            CurShape.PolyPnt(1).x = x
            CurShape.PolyPnt(1).y = y
End Select
End Sub

Private Sub picMain_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim CP As POINTAPI
On Error Resume Next
If Button = 1 Then
Select Case CurTool
        Case 0, 2
            linTemp.X2 = x
            linTemp.Y2 = y
                lblCoords.Caption = x & ", " & y
                txtpX.Text = x
                txtpY.Text = y
        Case 1, 3
            With shpTemp
             .Width = (x - .Left)
             .Height = (y - .Top)
             lblCoords.Caption = .Left & ", " & .Top & " | W=" & .Width & ", H=" & .Height
            End With
End Select
    Else
     lblCoords.Caption = x & ", " & y
End If
  
  Call Htb_CLEAR
  
  If LargeCH = True Then
     With linBCH_h
      .Y1 = y
      .Y2 = y
     End With
       With linBCH_v
        .X1 = x
        .X2 = x
       End With
  End If
  
  With shpMark
    .Left = x
    .Top = y
  End With
  
DoEvents
 GetCursorPos CP
   StretchBlt picZoomer.hdc, 0, 0, 128, 128, GetDC(0), (CP.x - 16), (CP.y - 16), 32, 32, vbSrcCopy
   picCurPX.BackColor = GetPixel(picMain.hdc, x, y)
End Sub

Private Sub picMain_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim P As POINTAPI
Dim Sx As Long
Dim Sy As Long

Select Case CurTool
        Case Is = 0 '=============
                    '   POLYGON
                    '=============

'check if this is the first point
If x = CurShape.PolyPnt(1).x And y = CurShape.PolyPnt(1).y Then
   If MsgBox("This point is close to the starting point, do you wish to close the polygon now?", vbYesNo + vbQuestion, "Close Polygon...") = vbYes Then
      Call DoneObject
      cmdTB(12).Enabled = True
      cmdTB(11).Enabled = False
      Exit Sub
   End If
End If

LineTo picMain.hdc, x, y

'increase the point count for our current polygon
'in our current frame...
'.PolyShp(CurFrame.PolyCount).PntCount = _
        CurFrame.PolyShp(CurFrame.PolyCount).PntCount + 1
CurShape.PntCount = CurShape.PntCount + 1
lblStat3.Caption = "Point Count: " & CurShape.PntCount
cmdTB(11).Enabled = True

CurShape.PolyPnt(CurShape.PntCount).x = x
CurShape.PolyPnt(CurShape.PntCount).y = y

FirstPoint = True
linTemp.Visible = False
linTemp.X1 = x
linTemp.Y1 = y
'MoveToEx picMain.hdc, X, Y, P
picMain.Refresh
lblStat.Caption = "POLYGON: Point Count - " & CurShape.PntCount

        Case Is = 1 '=============
                    '  Rectangle
                    '=============
        
        'look at the canvas first and verify that you wanna
        'add this shape to the frame...
        If MsgBox("Do you wish to add this shape to the frame?", vbYesNo, "Add Shape...") = vbNo Then
           shpTemp.Visible = False
           'reset flag..
           FirstPoint = False
           CurShape.PntCount = 0
           Exit Sub
        End If
           
        'ok they said yes... draw the shape..
        CurShape.PolyPnt(2).x = x
        CurShape.PolyPnt(2).y = y
            
            'draw onto canvas
            Rectangle picMain.hdc, _
                            CurShape.PolyPnt(1).x, _
                            CurShape.PolyPnt(1).y, _
                            CurShape.PolyPnt(2).x, _
                            CurShape.PolyPnt(2).y
        
        shpTemp.Visible = False 'hide shape object we
                                'dont need it anymore
        
        'now add the shape to our frame data
        
            Call CurShpTOCurFram
            
            'clear cur shape..
            CurShape.PntCount = 0
            cmdTB(11).Enabled = False
            cmdTB(12).Enabled = True
            
            'status bar text
            lblStat.Caption = "FRAME: " & CurFrame.PolyCount & " Objects"
        
        Case Is = 2 '==========
                    '   line
                    '==========
        'look at the canvas first and verify that you wanna
        'add this shape to the frame...
        If MsgBox("Do you wish to add this shape to the frame?", vbYesNo, "Add Shape...") = vbNo Then
           'reset flag..
           FirstPoint = False
           CurShape.PntCount = 0
           
           linTemp.Visible = False
           Exit Sub
        End If
        
        'ok they said yes... draw the shape..
        CurShape.PolyPnt(2).x = x
        CurShape.PolyPnt(2).y = y
        
        
        'now add the shape to our frame data
        Call CurShpTOCurFram
        'clear cur shape..
                lblStat.Caption = "POLYGON: Point Count - " & CurShape.PntCount
                
        CurShape.PntCount = 0
        cmdTB(11).Enabled = False
        cmdTB(12).Enabled = True
                         
        'draw onto canvas
                MoveToEx picMain.hdc, CurShape.PolyPnt(1).x, _
                         CurShape.PolyPnt(1).y, _
                         P
                         
                LineTo picMain.hdc, CurShape.PolyPnt(2).x, _
                         CurShape.PolyPnt(2).y

                
        'hide line
        linTemp.Visible = False 'hide shape object we
                                'dont need it anymore
        
        'show canvas now..
        picMain.Refresh
               
        Case Is = 3 '==========
                    '   Ellipse
                    '==========
        'look at the canvas first and verify that you wanna
        'add this shape to the frame...
        If MsgBox("Do you wish to add this shape to the frame?", vbYesNo, "Add Shape...") = vbNo Then
           'reset flag..
           FirstPoint = False
           CurShape.PntCount = 0
           
           shpTemp.Visible = False
           Exit Sub
        End If
        
        'ok they said yes... draw the shape..
        CurShape.PolyPnt(2).x = x
        CurShape.PolyPnt(2).y = y
        
        'draw onto canvas
        
               Ellipse picMain.hdc, _
                         CurShape.PolyPnt(1).x, _
                         CurShape.PolyPnt(1).y, _
                         CurShape.PolyPnt(2).x, _
                         CurShape.PolyPnt(2).y
                
        'hide shape
        shpTemp.Visible = False 'hide shape object we
                                'dont need it anymore
             
        'now add the shape to our frame data
        Call CurShpTOCurFram
        
        'clear cur shape..
        CurShape.PntCount = 0
        cmdTB(11).Enabled = False
        cmdTB(12).Enabled = True
        'show canvas now..
        picMain.Refresh
        
        'status bar text
End Select
lblStat2.Caption = "Object Count: " & CurFrame.PolyCount
End Sub

Private Sub PrevTimer()
'preview animation...
Dim T1 As Long
Dim PBuf As Long, Pbmp As Long

PBuf = CreateCompatibleDC(Me.hdc)
Pbmp = CreateCompatibleBitmap(Me.hdc, 388, 288)
SelectObject PBuf, Pbmp

tmrHov.Enabled = False
picMain.AutoRedraw = False
CurPFrame = 0
T1 = GetTickCount
Do
 DoEvents
 If (GetTickCount - T1) >= 50 Then
        CurPFrame = CurPFrame + 1
         
         BitBlt PBuf, 0, 0, 388, 288, 0, 0, 0, vbBlackness
         
            DrawFrame CurPFrame, TempPANI, PBuf
            
            BitBlt picMain.hdc, 0, 0, 388, 288, PBuf, 0, 0, vbSrcCopy
            
            If CurPFrame = (TempPANI.FrameCount) Then CurPFrame = 0
            T1 = GetTickCount
End If
Loop Until tmrPrev = False
tmrHov.Enabled = True
picMain.AutoRedraw = True
picMain.Cls
picMain.Refresh
'buffer cleanup...
DeleteObject Pbmp
DeleteDC PBuf
End Sub

Public Sub CurShpTOCurFram()
Dim I As Long
'increase the number of polygons in this frame

           CurFrame.PolyCount = CurFrame.PolyCount + 1
           
           ReDim Preserve CurFrame.PolyShp(1 To CurFrame.PolyCount)
        
            'copy data from temp shape structure to
            'the temp frame structure...
            CurFrame.PolyShp(CurFrame.PolyCount).PntCount _
                        = CurShape.PntCount
            CurFrame.PolyShp(CurFrame.PolyCount).PolyColor _
                            = CurShape.PolyColor
      'POE! \/
         ReDim Preserve CurFrame.PolyShp(CurFrame.PolyCount).PolyPnt(1 To CurShape.PntCount)
            For I = 1 To CurShape.PntCount
            CurFrame.PolyShp(CurFrame.PolyCount).PolyPnt(I).x _
                            = CurShape.PolyPnt(I).x
            CurFrame.PolyShp(CurFrame.PolyCount).PolyPnt(I).y _
                            = CurShape.PolyPnt(I).y
            Next
            
            CurFrame.PolyShp(CurFrame.PolyCount).PolyType _
                            = CurShape.PolyType
    'status text
lblStat.Caption = "Frame: " & CurFrame.PolyCount & " objects"
End Sub

Public Sub DoneObject()
Dim I As Long
Dim P As POINTAPI
Dim x As Long, y As Long

'copy data from our current shape structure to
'our tempt frame structure...
Call CurShpTOCurFram

'draw to screen
Polygon picMain.hdc, CurShape.PolyPnt(1), CurShape.PntCount

'clear current shape
      CurShape.PntCount = 0
      cmdTB(12).Enabled = False
      'loose line object
      linTemp.Visible = False
      
        'reset flag
        FirstPoint = False

'refresh screen...
cmdTB(12).Enabled = True
cmdTB(11).Enabled = False
picMain.Refresh

'status bar text
lblStat2.Caption = "Object Count: " & CurFrame.PolyCount
lblStat3.Caption = "Point Count: 0"
End Sub

Private Sub NewFile()
'clear for new canvas... new file...
CurFrame.PolyCount = 0
CurShape.PntCount = 0

SaveFileName = ""
mnuSave.Enabled = False

cmdTB(11).Enabled = False
cmdTB(12).Enabled = False

picMain.Cls
picMain.Refresh
lstFrames.Clear
lstObjects.Clear
TempPANI.FrameCount = 0
cmdTB(9).Enabled = False
cmdTB(2).Enabled = False
mnuSaveAs.Enabled = False
FirstPoint = False
End Sub

Private Sub OpenFile()
Dim A As String
    FD.hwnd = Me.hwnd
    FD.DefaultExt = "pan"
    FD.Filter = "PolyAnimation (*.PAN) | *.pan|All Files (*.*) | *.*"
    FD.ShowOpen
    A = FD.Filename
    LoadPan A
    If TempPANI.FrameCount > 0 Then
       SaveFileName = A
       mnuSave.Enabled = True
       cmdTB(2).Enabled = True
       mnuSaveAs.Enabled = True
    End If
End Sub

Private Sub SaveFile()
Dim A As String
    FD.hwnd = Me.hwnd
    FD.DefaultExt = "pan"
    FD.Filter = "PolyAnimation (*.PAN) | *.pan|All Files (*.*) | *.*"
    FD.ShowSave
    A = FD.Filename
     If Len(A) > 0 Then
       SaveFileName = A
       mnuSave.Enabled = True
     End If
    SavePan A
End Sub

Private Sub ExportBMP()
Dim A As String
    FD.hwnd = Me.hwnd
    FD.DefaultExt = "bmp"
    FD.Filter = "Windows Bitmap (*.bmp) | *.bmp|All Files (*.*) | *.*"
    FD.ShowSave
    A = FD.Filename
     If Len(A) > 0 Then
     lblStat.Caption = "Saving Bitmap..."
     lblStat.Refresh
     SavePicture picMain.Image, A
     lblStat.Caption = ""
     End If
End Sub

Private Sub LoadBGpic()
On Error GoTo ErrOut:
Dim A As String
FD.Filter = "All Files (*.*) | *.*"
FD.ShowOpen
A = FD.Filename
If Len(A) > 0 Then
   picMain.Picture = LoadPicture(A)
   picMain.Cls
   picMain.Refresh
   FirstPoint = False
End If

Exit Sub
ErrOut:
 MsgBox Err.Description, vbExclamation
End Sub

Private Sub DoneFrame()
Dim I As Long
Dim C As Long
Dim B As Long
If MsgBox("Are you sure you want to add this frame?", vbYesNo + vbQuestion, "Add Frame?") = vbNo Then Exit Sub
'increase temp structs framecount
TempPANI.FrameCount = TempPANI.FrameCount + 1
cmdTB(9).Enabled = True
cmdTB(2).Enabled = False
'hold enough frames...
ReDim Preserve TempPANI.Polys(1 To TempPANI.FrameCount)

    'number of polygons/shapes/lines
    TempPANI.Polys(TempPANI.FrameCount).PolyCount = CurFrame.PolyCount
        'for each polygon/shape/line
        
        'hold enough shapes in each of these frames
        'that we're going through
    ReDim Preserve TempPANI.Polys(TempPANI.FrameCount).PolyShp(1 To CurFrame.PolyCount)
        For C = 1 To CurFrame.PolyCount
            
            'match point count
            TempPANI.Polys(TempPANI.FrameCount).PolyShp(C).PntCount = _
                    CurFrame.PolyShp(C).PntCount
            
            'match color
            TempPANI.Polys(TempPANI.FrameCount).PolyShp(C).PolyColor = _
                    CurFrame.PolyShp(C).PolyColor
                    
            'match shape type
            TempPANI.Polys(TempPANI.FrameCount).PolyShp(C).PolyType = _
                    CurFrame.PolyShp(C).PolyType
            
                    'match each point
        ReDim Preserve TempPANI.Polys(TempPANI.FrameCount).PolyShp(C).PolyPnt(1 To CurFrame.PolyShp(C).PntCount)
           For B = 1 To CurFrame.PolyShp(C).PntCount
           'hold enough points in this shape...
                        
                TempPANI.Polys(TempPANI.FrameCount).PolyShp(C).PolyPnt(B).x = _
                    CurFrame.PolyShp(C).PolyPnt(B).x
                    
                TempPANI.Polys(TempPANI.FrameCount).PolyShp(C).PolyPnt(B).y = _
                    CurFrame.PolyShp(C).PolyPnt(B).y
           Next
        Next


lstFrames.AddItem "Frame " & TempPANI.FrameCount & ": " & TempPANI.Polys(TempPANI.FrameCount).PolyCount & " Objects"
'clear data
For I = 1 To CurFrame.PolyCount
CurFrame.PolyShp(I).PntCount = 0
Next
CurFrame.PolyCount = 0
'clear the canvas
Set picMain.Picture = Nothing
picMain.Cls
picMain.Refresh

FirstPoint = False
CurShape.PntCount = 0
cmdTB(11).Enabled = False
cmdTB(12).Enabled = False
cmdTB(2).Enabled = True 'enable save button...
mnuSaveAs.Enabled = True 'enable save-as menu item

'update status text displays...
lblStat2.Caption = "Object Count: 0"
lblStat3.Caption = "Point Count: 0"
lblStat1.Caption = "Frame Count: " & TempPANI.FrameCount
End Sub
