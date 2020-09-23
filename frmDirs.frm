VERSION 5.00
Begin VB.Form frmDirs 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Select Directory"
   ClientHeight    =   4065
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4455
   Icon            =   "frmDirs.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   271
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   297
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   420
      Left            =   3150
      TabIndex        =   4
      Top             =   3600
      Width           =   1230
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Ok"
      Default         =   -1  'True
      Height          =   420
      Left            =   1800
      TabIndex        =   3
      Top             =   3600
      Width           =   1230
   End
   Begin VB.TextBox txtDir 
      Height          =   330
      Left            =   0
      TabIndex        =   2
      Top             =   3195
      Width           =   4425
   End
   Begin VB.DirListBox tvDir 
      Height          =   2790
      Left            =   0
      TabIndex        =   1
      Top             =   360
      Width           =   4425
   End
   Begin VB.DriveListBox drvDrive 
      Height          =   315
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4425
   End
End
Attribute VB_Name = "frmDirs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Path As String

Private Sub Command1_Click()
    Path = tvDir.Path
    Unload Me
End Sub

Private Sub Command2_Click()
    Path = ""
    Unload Me
End Sub

Private Sub drvDrive_Change()
    tvDir.Path = drvDrive.Drive
End Sub

