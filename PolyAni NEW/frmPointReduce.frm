VERSION 5.00
Begin VB.Form frmPointReduce 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Point Reduction Mini Wizard..."
   ClientHeight    =   2730
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4470
   Icon            =   "frmPointReduce.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2730
   ScaleWidth      =   4470
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraStep 
      Height          =   1890
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   90
      Width           =   4230
      Begin VB.Label Label2 
         Caption         =   $"frmPointReduce.frx":0CCA
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   255
         TabIndex        =   4
         Top             =   555
         Width           =   3780
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "What is this?"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   195
         TabIndex        =   3
         Top             =   225
         Width           =   1215
      End
   End
   Begin VB.Frame fraStep 
      Height          =   1785
      Index           =   2
      Left            =   120
      TabIndex        =   2
      Top             =   90
      Width           =   4230
   End
   Begin VB.Frame fraStep 
      Height          =   1785
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   90
      Width           =   4230
   End
End
Attribute VB_Name = "frmPointReduce"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
