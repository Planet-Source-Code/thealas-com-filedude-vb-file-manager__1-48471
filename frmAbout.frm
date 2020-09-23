VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About File Dude"
   ClientHeight    =   3480
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4155
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   232
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   277
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Caption         =   "What ever..."
      Default         =   -1  'True
      Height          =   465
      Left            =   1755
      TabIndex        =   5
      Top             =   2880
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "VOTE FOR ME !!!    :-)"
      Height          =   195
      Left            =   1485
      TabIndex        =   6
      Top             =   2070
      Width           =   2580
   End
   Begin VB.Label lblSite 
      Caption         =   "www.hallsoft.tk"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Index           =   4
      Left            =   1485
      MouseIcon       =   "frmAbout.frx":000C
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Top             =   1665
      Width           =   1335
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00C00000&
      FillColor       =   &H00C00000&
      FillStyle       =   7  'Diagonal Cross
      Height          =   2265
      Left            =   135
      Top             =   1080
      Width           =   735
   End
   Begin VB.Label lblAb 
      AutoSize        =   -1  'True
      Caption         =   "Sala Bojan, alas@eunet.yu"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   240
      Index           =   3
      Left            =   1485
      TabIndex        =   3
      Top             =   1260
      Width           =   2430
   End
   Begin VB.Label lblAb 
      AutoSize        =   -1  'True
      Caption         =   "Author:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   240
      Index           =   2
      Left            =   1485
      TabIndex        =   2
      Top             =   900
      Width           =   615
   End
   Begin VB.Label lblAb 
      AutoSize        =   -1  'True
      Caption         =   "Copyright (C) Hallsoft"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   1
      Left            =   1485
      TabIndex        =   1
      Top             =   450
      Width           =   1875
   End
   Begin VB.Label lblAb 
      AutoSize        =   -1  'True
      Caption         =   "File Dude 1.0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   0
      Left            =   1485
      TabIndex        =   0
      Top             =   135
      Width           =   1170
   End
   Begin VB.Image Image1 
      Height          =   900
      Left            =   90
      Picture         =   "frmAbout.frx":015E
      Top             =   90
      Width           =   825
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub lblSite_Click(Index As Integer)
    frmMain.OpenFile "www.hallsoft.tk", "Open"
End Sub
