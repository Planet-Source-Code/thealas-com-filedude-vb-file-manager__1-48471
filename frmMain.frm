VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   Caption         =   "FileDude"
   ClientHeight    =   5835
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   5580
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5835
   ScaleWidth      =   5580
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picSplitter 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      FillColor       =   &H00808080&
      Height          =   4800
      Left            =   5400
      ScaleHeight     =   2090.126
      ScaleMode       =   0  'User
      ScaleWidth      =   780
      TabIndex        =   6
      Top             =   705
      Visible         =   0   'False
      Width           =   72
   End
   Begin VB.DriveListBox drvDrive 
      Height          =   315
      Left            =   0
      TabIndex        =   11
      Top             =   720
      Width           =   1995
   End
   Begin VB.PictureBox picIconFrame 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   600
      Left            =   3915
      ScaleHeight     =   40
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   34
      TabIndex        =   8
      Top             =   2160
      Visible         =   0   'False
      Width           =   510
      Begin VB.PictureBox picIcon 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   0
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   9
         Top             =   0
         Width           =   240
      End
   End
   Begin VB.DirListBox tvTreeView 
      Height          =   4365
      Left            =   0
      TabIndex        =   7
      Top             =   1080
      Width           =   1995
   End
   Begin MSComctlLib.ImageList imgListViewIcons 
      Left            =   3015
      Top             =   1350
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":27A2
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picTitles 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   0
      ScaleHeight     =   300
      ScaleWidth      =   5580
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   420
      Width           =   5580
      Begin MSComctlLib.ProgressBar pbFiles 
         Height          =   150
         Left            =   2700
         TabIndex        =   10
         Top             =   90
         Width           =   2490
         _ExtentX        =   4392
         _ExtentY        =   265
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
      End
      Begin VB.Label lblTitle 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "   Files:"
         Height          =   270
         Index           =   1
         Left            =   2078
         TabIndex        =   4
         Tag             =   " ListView:"
         Top             =   12
         Width           =   3216
      End
      Begin VB.Label lblTitle 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "   Directories:"
         Height          =   270
         Index           =   0
         Left            =   0
         TabIndex        =   3
         Tag             =   " TreeView:"
         Top             =   12
         Width           =   2016
      End
   End
   Begin MSComctlLib.StatusBar sbStatusBar 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   0
      Top             =   5565
      Width           =   5580
      _ExtentX        =   9843
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Text            =   "Status"
            TextSave        =   "Status"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            AutoSize        =   2
            TextSave        =   "7/24/2003"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            AutoSize        =   2
            TextSave        =   "12:39 h"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imlToolbarIcons 
      Left            =   2385
      Top             =   1350
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   13
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":28FE
            Key             =   "Forward"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2A10
            Key             =   "Back"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2B22
            Key             =   "Up One Level"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2C34
            Key             =   "Cut"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2D46
            Key             =   "Copy"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2E58
            Key             =   "Paste"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2F6A
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":307C
            Key             =   "Properties"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":318E
            Key             =   "View Details"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":32A0
            Key             =   "View Large Icons"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":33B2
            Key             =   "View List"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":34C4
            Key             =   "View Small Icons"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":35D6
            Key             =   "Help"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbToolBar 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   5580
      _ExtentX        =   9843
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "imlToolbarIcons"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   15
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Back"
            Object.ToolTipText     =   "Back"
            ImageKey        =   "Back"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Up One Level"
            Object.ToolTipText     =   "Up One Level"
            ImageKey        =   "Up One Level"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Cut"
            Object.ToolTipText     =   "Cut"
            ImageKey        =   "Cut"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Copy"
            Object.ToolTipText     =   "Copy"
            ImageKey        =   "Copy"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Paste"
            Object.ToolTipText     =   "Paste"
            ImageKey        =   "Paste"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Delete"
            Object.ToolTipText     =   "Delete"
            ImageKey        =   "Delete"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Properties"
            Object.ToolTipText     =   "Properties"
            ImageKey        =   "Properties"
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "View Details"
            Object.ToolTipText     =   "View Details"
            ImageKey        =   "View Details"
            Style           =   2
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "View Large Icons"
            Object.ToolTipText     =   "View Large Icons"
            ImageKey        =   "View Large Icons"
            Style           =   2
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "View List"
            Object.ToolTipText     =   "View List"
            ImageKey        =   "View List"
            Style           =   2
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "View Small Icons"
            Object.ToolTipText     =   "View Small Icons"
            ImageKey        =   "View Small Icons"
            Style           =   2
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Help"
            Object.ToolTipText     =   "Help"
            ImageKey        =   "Help"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvListView 
      Height          =   4800
      Left            =   2070
      TabIndex        =   5
      Top             =   705
      Width           =   3210
      _ExtentX        =   5662
      _ExtentY        =   8467
      SortKey         =   1
      View            =   3
      Arrange         =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      SmallIcons      =   "imgListViewIcons"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "File Name"
         Object.Width           =   3881
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Type"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Size (KB)"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Modified"
         Object.Width           =   7056
      EndProperty
   End
   Begin VB.Image imgSplitter 
      Height          =   4785
      Left            =   3015
      MousePointer    =   9  'Size W E
      Top             =   720
      Width           =   150
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Open..."
      End
      Begin VB.Menu mnuFileNew 
         Caption         =   "&New Directory"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuFileBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileBack 
         Caption         =   "&Back"
      End
      Begin VB.Menu mnuFileUpOneLevel 
         Caption         =   "&Up one level"
      End
      Begin VB.Menu mnuFileBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileSendTo 
         Caption         =   "Mo&ve"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuFileCopy 
         Caption         =   "&Copy"
      End
      Begin VB.Menu mnuFileDelete 
         Caption         =   "&Delete"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu mnuFileRename 
         Caption         =   "Rena&me"
      End
      Begin VB.Menu mnuFileProperties 
         Caption         =   "Propert&ies"
         Shortcut        =   ^{F4}
      End
      Begin VB.Menu mnuFileBar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileClose 
         Caption         =   "&Close"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuEditUndo 
         Caption         =   "&Undo"
      End
      Begin VB.Menu mnuEditBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditCut 
         Caption         =   "Cu&t"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuEditCopy 
         Caption         =   "&Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuEditPaste 
         Caption         =   "&Paste"
         Shortcut        =   ^V
      End
      Begin VB.Menu mnuEditBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditSelectAll 
         Caption         =   "Select &All"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuEditInvertSelection 
         Caption         =   "&Invert Selection"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuViewToolbar 
         Caption         =   "&Toolbar"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuViewStatusBar 
         Caption         =   "Status &Bar"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuViewBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuListViewMode 
         Caption         =   "Lar&ge Icons"
         Enabled         =   0   'False
         Index           =   0
      End
      Begin VB.Menu mnuListViewMode 
         Caption         =   "S&mall Icons"
         Index           =   1
      End
      Begin VB.Menu mnuListViewMode 
         Caption         =   "&List"
         Index           =   2
      End
      Begin VB.Menu mnuListViewMode 
         Caption         =   "&Details"
         Index           =   3
      End
      Begin VB.Menu mnuViewBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewArrangeIcons 
         Caption         =   "Arrange &Icons"
      End
      Begin VB.Menu mnuViewBar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewRefresh 
         Caption         =   "&Refresh"
      End
      Begin VB.Menu mnuViewOptions 
         Caption         =   "&Options..."
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpContents 
         Caption         =   "&Contents"
      End
      Begin VB.Menu mnuHelpBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About "
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mbMoving As Boolean

Const NAME_COLUMN = 0
Const TYPE_COLUMN = 1
Const SIZE_COLUMN = 2
Const DATE_COLUMN = 3
Const sglSplitLimit = 500

Const LISTVIEW_MODE0 = "View Large Icons"
Const LISTVIEW_MODE1 = "View Small Icons"
Const LISTVIEW_MODE2 = "View List"
Const LISTVIEW_MODE3 = "View Details"

Private imlCashe As Long
Private tvMemo(1 To 999) As String
Private tvMemoLevel As Long
Private tvMemoB As Boolean

Public Sub SendToRecycleBin(FileName As String)
On Error GoTo e
    'Same as rest, but different op
     Dim FileOperation As SHFILEOPSTRUCT
     Dim lReturn As Long
     Dim sTempFilename As String * 100
     Dim sSendMeToTheBin As String
     sSendMeToTheBin = FileName
     With FileOperation
        .wFunc = FO_DELETE ' = 3
        .pFrom = sSendMeToTheBin
        .fFlags = FOF_SILENT Or FOF_ALLOWUNDO 'U can undo it from explorer !  :-)
     End With
     lReturn = SHFileOperation(FileOperation)
Exit Sub
e:
    Err_Raise Err.Number, Err.Description, "frm_main", "sendtorecyclebin": Resume Next
End Sub

Private Sub drvDrive_Change()
    tvTreeView.Path = drvDrive.Drive
End Sub

Private Sub Form_Load()
On Error GoTo e
    Me.Left = GetSetting(App.Title, "Settings", "MainLeft", 1000)
    Me.Top = GetSetting(App.Title, "Settings", "MainTop", 1000)
    Me.Width = GetSetting(App.Title, "Settings", "MainWidth", 6500)
    Me.Height = GetSetting(App.Title, "Settings", "MainHeight", 6500)
    tvTreeView.Path = "C:\" 'drvDrive.Drive
Exit Sub
e:
    Err_Raise Err.Number, Err.Description, "frm_main", "formload": Resume Next '>:-( never add error handling to form_load !!!!
End Sub


Private Sub Form_Paint()
On Error GoTo e
    lvListView.View = Val(GetSetting(App.Title, "Settings", "ViewMode", "0"))
    Select Case lvListView.View
        Case lvwIcon
            tbToolBar.Buttons(LISTVIEW_MODE0).Value = tbrPressed
        Case lvwSmallIcon
            tbToolBar.Buttons(LISTVIEW_MODE1).Value = tbrPressed
        Case lvwList
            tbToolBar.Buttons(LISTVIEW_MODE2).Value = tbrPressed
        Case lvwReport
            tbToolBar.Buttons(LISTVIEW_MODE3).Value = tbrPressed
    End Select
Exit Sub
e:
    Err_Raise Err.Number, Err.Description, "frm_main", "formpaint": Resume Next
End Sub


Private Sub Form_Unload(Cancel As Integer)
On Error GoTo e
    Dim I As Integer
    'close all sub forms
    For I = Forms.Count - 1 To 1 Step -1
        Unload Forms(I)
    Next
    If Not Me.WindowState = vbMinimized Then
        If Not Me.WindowState = vbMaximized Then
            SaveSetting App.Title, "Settings", "MainLeft", Me.Left
            SaveSetting App.Title, "Settings", "MainTop", Me.Top
            SaveSetting App.Title, "Settings", "MainWidth", Me.Width
            SaveSetting App.Title, "Settings", "MainHeight", Me.Height
        End If
    End If
    SaveSetting App.Title, "Settings", "ViewMode", lvListView.View
Exit Sub
e:
    Err_Raise Err.Number, Err.Description, "frm_main", "formunload": Resume Next
End Sub



Private Sub Form_Resize()
    On Error Resume Next
    If Me.Width < 3000 Then Me.Width = 3000
    SizeControls imgSplitter.Left
End Sub


Private Sub imgSplitter_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    With imgSplitter
        picSplitter.Move .Left, .Top, .Width \ 2, .Height - 20
    End With
    picSplitter.Visible = True
    mbMoving = True
End Sub


Private Sub imgSplitter_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim sglPos As Single
    If mbMoving Then
        sglPos = x + imgSplitter.Left
        If sglPos < sglSplitLimit Then
            picSplitter.Left = sglSplitLimit
        ElseIf sglPos > Me.Width - sglSplitLimit Then
            picSplitter.Left = Me.Width - sglSplitLimit
        Else
            picSplitter.Left = sglPos
        End If
    End If
End Sub


Private Sub imgSplitter_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    SizeControls picSplitter.Left
    picSplitter.Visible = False
    mbMoving = False
End Sub





Sub SizeControls(x As Single)
    On Error Resume Next
    'set the width
    If x < 1500 Then x = 1500
    If x > (Me.Width - 1500) Then x = Me.Width - 1500
    tvTreeView.Width = x
    drvDrive.Width = x
    imgSplitter.Left = x
    lvListView.Left = x + 40
    lvListView.Width = Me.Width - (tvTreeView.Width + 140)
    lblTitle(0).Width = tvTreeView.Width
    lblTitle(1).Left = lvListView.Left + 20
    lblTitle(1).Width = lvListView.Width - 40
    pbFiles.Left = lblTitle(1).Left + 50 * 15
    pbFiles.Width = lblTitle(1).Width - 60 * 15
    'set the top
    If tbToolBar.Visible Then
        drvDrive.Top = tbToolBar.Height + picTitles.Height
        tvTreeView.Top = tbToolBar.Height + drvDrive.Height * 2
    Else
        drvDrive.Top = picTitles.Height
        tvTreeView.Top = drvDrive.Height * 2
    End If
    lvListView.Top = drvDrive.Top
    'set the height
    If sbStatusBar.Visible Then
        tvTreeView.Height = Me.ScaleHeight - (picTitles.Top + picTitles.Height + sbStatusBar.Height) - drvDrive.Height
    Else
        tvTreeView.Height = Me.ScaleHeight - (picTitles.Top + picTitles.Height) - drvDrive.Height
    End If
    lvListView.Height = tvTreeView.Height + drvDrive.Height
    imgSplitter.Top = tvTreeView.Top
    imgSplitter.Height = tvTreeView.Height
End Sub


Private Sub lblTitle_Click(Index As Integer)

End Sub

Private Sub lvListView_AfterLabelEdit(Cancel As Integer, NewString As String)
On Error GoTo e
    'If its Drive:\ then dont use "\"
    If Not Right(tvTreeView.Path, 1) = "\" Then
        Name tvTreeView.Path & "\" & lvListView.SelectedItem.Text As tvTreeView.Path & "\" & NewString
    Else
        Name tvTreeView.Path & lvListView.SelectedItem.Text As tvTreeView.Path & NewString
    End If
Exit Sub
e:
    Err_Raise Err.Number, Err.Description, "frm_main", "lwlistviewafterlabeledit": Resume Next
End Sub

Private Sub lvListView_BeforeLabelEdit(Cancel As Integer)
    If lvListView.SelectedItem.Text = "<...>" Then Cancel = 1
End Sub

Private Sub lvListView_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    lvListView.Sorted = True
    lvListView.SortKey = ColumnHeader.Index - 1
End Sub


Private Sub lvListView_DblClick()
On Error Resume Next
    Dim fPath$
    With lvListView
        fPath = tvGetPath
        If .SelectedItem.SubItems(1) = "<DIR>" Then
            If .SelectedItem.Text = "<...>" Then
                tvTreeView.Path = tvTreeView.Path & "\.."
            Else
                tvTreeView.Path = fPath & "\" & .SelectedItem.Text
            End If
        Else
            OpenFile fPath & "\" & lvListView.SelectedItem.Text, "Open" 'There is "Edit" also !
        End If
    End With
End Sub

Private Sub lvListView_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then PopupMenu mnuFile
End Sub

Private Sub mnuFileBack_Click()
    If tvMemoLevel > 1 Then
        tvMemoLevel = tvMemoLevel - 1
        tvTreeView.Path = tvMemo(tvMemoLevel)
    End If
End Sub

Private Sub mnuFileCopy_Click()
On Error GoTo e
    'copies selected files or directories
    Dim fPath$, I&, fCount&
    fPath = ShowDirs(tvGetPath)
    If Right(fPath, 1) = "\" Then
        fPath = Left(fPath, Len(fPath) - 1)
    End If
    If Not fPath = "" Then
        With lvListView
            For I = 1 To .ListItems.Count
                If .ListItems(I).Selected = True Then
                    fCount = fCount + 1
                End If
            Next I
            pbFiles.Max = fCount + 1
            sbStatusBar.Panels(1).Text = "Copying files..."
            For I = 1 To .ListItems.Count
                If .ListItems(I).Selected = True Then
                    If InStr(1, .ListItems(I).SubItems(1), "<DIR>") > 0 Then
                        MkDir fPath & "\" & .ListItems(I).Text
                        XCopy tvGetPath & "\" & .ListItems(I).Text, fPath & "\" & .ListItems(I).Text, True, "*"
                    Else
                        FileCopy tvGetPath & "\" & .ListItems(I).Text, fPath & "\" & .ListItems(I).Text
                    End If
                    pbFiles.Value = pbFiles.Value + 1
                End If
            Next I
        End With
    End If
    pbFiles.Value = 0
Exit Sub
e:
    Err_Raise Err.Number, Err.Description, "frm_main", "mnufilecopy": Resume Next
End Sub

Private Sub mnuFileUpOneLevel_Click()
On Error Resume Next
    tvTreeView.Path = tvTreeView.Path & "\.."
End Sub

Private Sub mnuViewArrangeIcons_Click()
    lvListView.Arrange = lvwAutoLeft
End Sub

Private Sub tbToolBar_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "Back"
            mnuFileBack_Click
        Case "Up One Level"
            mnuFileUpOneLevel_Click
        Case "Cut"
            mnuEditCut_Click
        Case "Copy"
            mnuEditCopy_Click
        Case "Paste"
            mnuEditPaste_Click
        Case "Delete"
            mnuFileDelete_Click
        Case "Properties"
            mnuFileProperties_Click
        Case "View Details"
            lvListView.View = lvwReport
        Case "View Large Icons"
            lvListView.View = lvwIcon
        Case "View List"
            lvListView.View = lvwList
        Case "View Small Icons"
            lvListView.View = lvwSmallIcon
        Case "Help"
            mnuHelpContents_Click
    End Select
End Sub

Private Sub mnuHelpAbout_Click()
    frmAbout.Show vbModal, Me
End Sub


Private Sub mnuHelpContents_Click()
    Dim nRet As Integer
    'if there is no helpfile for this project display a message to the user
    'you can set the HelpFile for your application in the
    'Project Properties dialog
    If Len(App.HelpFile) = 0 Then
        MsgBox "Seems there is no help, see the comment at the top of this line", vbInformation, Me.Caption
    Else
        On Error Resume Next
        nRet = OSWinHelp(Me.hWnd, App.HelpFile, 3, 0)
        If Err Then
            MsgBox Err.Description
        End If
    End If
End Sub



Private Sub mnuViewRefresh_Click()
    Call tvTreeView_Change
End Sub



Private Sub mnuViewStatusBar_Click()
    mnuViewStatusBar.Checked = Not mnuViewStatusBar.Checked
    sbStatusBar.Visible = mnuViewStatusBar.Checked
    SizeControls imgSplitter.Left
End Sub

Private Sub mnuViewToolbar_Click()
    mnuViewToolbar.Checked = Not mnuViewToolbar.Checked
    tbToolBar.Visible = mnuViewToolbar.Checked
    SizeControls imgSplitter.Left
End Sub

Private Sub mnuEditInvertSelection_Click()
    Dim I&
    With lvListView
        For I = 1 To .ListItems.Count
            If .ListItems(I).Selected = True Then
                .ListItems(I).Selected = False
            Else
                .ListItems(I).Selected = True
            End If
        Next I
    End With
End Sub

Private Sub mnuEditSelectAll_Click()
    Dim I&
    With lvListView
        For I = 1 To .ListItems.Count
            .ListItems(I).Selected = True
        Next I
    End With
End Sub



Private Sub mnuEditPaste_Click()
    MsgBox "Well, you've found me... if u know how to get and set files to clipboard (like explorer) tell me !!"
End Sub

Private Sub mnuEditCopy_Click()
    MsgBox "Well, you've found me... if u know how to get and set files to clipboard (like explorer) tell me !!"
End Sub

Private Sub mnuEditCut_Click()
    MsgBox "Well, you've found me... if u know how to get and set files to clipboard (like explorer) tell me !!"
End Sub

Private Sub mnuEditUndo_Click()
    MsgBox "Not yet done, just undo it from explorer (if you've done something SO bad :-)"
End Sub

Private Sub mnuFileClose_Click()
    'unload the form
    Unload Me
End Sub

Private Sub mnuFileProperties_Click()
On Error GoTo e
    'Here, you can see how to retrieve all the important properties
    'for single file. You only need tvCAttr function to copy, just for making Yes/No from
    'numbers, no apis at all, just code down there. It is all in VBA.
    Dim fInfo$
    With lvListView.SelectedItem
        fInfo = "File Name:" & vbTab & .Text & vbCrLf
        fInfo = fInfo & "File Size:" & vbTab & vbTab & FileLen(tvGetPath & "\" & .Text) / 1024 & " Kb" & vbCrLf
        fInfo = fInfo & "File Date:" & vbTab & vbTab & FileDateTime(tvGetPath & "\" & .Text) & vbCrLf
        fInfo = fInfo & "Is Hidden:" & vbTab & tvCAttr((GetAttr(tvGetPath & "\" & .Text) And vbHidden)) & vbCrLf
        fInfo = fInfo & "Is System:" & vbTab & tvCAttr((GetAttr(tvGetPath & "\" & .Text) And vbSystem)) & vbCrLf
        fInfo = fInfo & "Is ReadOnly:" & vbTab & tvCAttr((GetAttr(tvGetPath & "\" & .Text) And vbReadOnly)) & vbCrLf
        fInfo = fInfo & "Is Archive:" & vbTab & vbTab & tvCAttr((GetAttr(tvGetPath & "\" & .Text) And vbArchive)) & vbCrLf
    End With
    MsgBox fInfo, vbInformation, "File Properties"
Exit Sub
e:
    Err_Raise Err.Number, Err.Description, "frm_main", "mnufileproperties": Resume Next
End Sub

Private Sub mnuFileRename_Click()
    lvListView.StartLabelEdit
End Sub

Private Sub mnuFileDelete_Click()
On Error GoTo e
    'deletes selected files or directories
    Dim I&, Msg, fPath$, Stat&, Max&
    With lvListView
        .Enabled = False
        For I = 1 To .ListItems.Count
            If .ListItems(I).Selected = True Then
                Max = Max + 1
            End If
        Next I
        pbFiles.Max = Max
clean:
        sbStatusBar.Panels(1).Text = "Sending files to recycle bin, may take a while :-( "
        For I = 1 To .ListItems.Count
            If .ListItems(I).Selected = True Then
                fPath = tvGetPath
                Call SendToRecycleBin(fPath & "\" & .ListItems(I).Text)
                .ListItems.Remove I
                pbFiles.Value = pbFiles.Value + 1
                GoTo clean
            End If
        Next I
        .Enabled = True
    End With
    pbFiles.Value = 0
Exit Sub
e:
    Err_Raise Err.Number, Err.Description, "frm_main", "mnufiledelete": Resume Next
End Sub

Private Sub mnuFileNew_Click()
On Error GoTo e
    Dim fDir$, I&
    fDir = InputBox("Enter directory name: ", "New directory")
    If Not fDir = "" Then
        MkDir tvGetPath & "\" & fDir
    End If
    tvTreeView.Path = tvGetPath & "\" & fDir
    tvTreeView.Path = tvTreeView.Path & "\.."
    For I = 1 To lvListView.ListItems.Count
    Next I
Exit Sub
e:
    Err_Raise Err.Number, Err.Description, "frm_main", "mnufilenew": Resume Next
End Sub

Private Sub mnuFileSendTo_Click()
On Error GoTo e
    'Moves selected files or directories
    Dim fPath$, I&, fCount&
    fPath = ShowDirs(tvGetPath)
    If Right(fPath, 1) = "\" Then
        fPath = Left(fPath, Len(fPath) - 1)
    End If
    With lvListView
        sbStatusBar.Panels(1).Text = "Moving files, please wait..."
        If Not fPath = "" Then
            For I = 1 To .ListItems.Count
                If .ListItems(I).Selected = True Then
                    fCount = fCount + 1
                End If
            Next I
            pbFiles.Max = fCount + 1
            For I = 1 To .ListItems.Count
                If .ListItems(I).Selected = True Then
                    XMove tvGetPath & "\" & .ListItems(I).Text, fPath & "\"
                    pbFiles.Value = pbFiles.Value + 1
                End If
            Next I
        End If
clean:
        For I = 1 To .ListItems.Count
            If .ListItems(I).Selected = True Then
                .ListItems.Remove I
                GoTo clean
            End If
        Next I
    End With

    pbFiles.Value = 0

Exit Sub
e:
    Err_Raise Err.Number, Err.Description, "frm_main", "mnufilesendto": Resume Next
End Sub


Private Sub mnuFileOpen_Click()
    Call lvListView_DblClick
End Sub


Public Function tvIsDir(FileName As String) As Boolean
On Error GoTo e
    'If you need, I've made this for easier understanding, it tells you is it dir or not
    Dim fData As WIN32_FIND_DATA
    Dim fHwnd As Long
    Dim fReg As Long
    
    fHwnd = FindFirstFile(FileName, fData)
    tvIsDir = ((fData.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) = FILE_ATTRIBUTE_DIRECTORY)
    fReg = FindClose(fHwnd)
Exit Function
e:
    Err_Raise Err.Number, Err.Description, "frm_main", "tvsdir": Resume Next
End Function

Private Sub tvTreeView_Change()
On Error GoTo e
    Dim fDir$, Itm As ListItem, fPath$, I&, fCount&
    With lvListView
        .ListItems.Clear
        fPath = tvGetPath
        sbStatusBar.Panels(1).Text = "Adding folders..."
        If Not Len(fPath) = 2 Then
            Set Itm = .ListItems.Add(, , "<...>", , 1)
            Itm.SubItems(1) = "<DIR>"
        End If
        For I = tvTreeView.ListIndex + 1 To tvTreeView.ListCount - 1
            Set Itm = .ListItems.Add(, , tvStripPath(tvTreeView.List(I)), , tvAddIconToIML(tvTreeView.List(I), "<DIR>", imgListViewIcons))
            Itm.SubItems(1) = "<DIR>"
        Next I
        
        fCount = 0
        fDir = Dir(fPath & "\", vbHidden)
        While fDir <> ""
            fCount = fCount + 1
            fDir = Dir
        Wend
        
        sbStatusBar.Panels(2).Text = "Files: " & fCount & " Folders: " & tvTreeView.ListCount & "   "
        
        pbFiles.Max = fCount + 1
        pbFiles.Value = 0
        fDir = Dir(fPath & "\", vbHidden)
        sbStatusBar.Panels(1).Text = "Adding files..."
        While fDir <> ""
            Set Itm = .ListItems.Add(, , fDir, , tvAddIconToIML(fPath & "\" & fDir, UCase(Right(fDir, 3)), imgListViewIcons))
            pbFiles.Value = pbFiles.Value + 1
            Itm.SubItems(1) = UCase(Right(fDir, 3))
            Itm.SubItems(2) = CLng(FileLen(fPath & "\" & fDir) / 1024)
            Itm.SubItems(3) = FileDateTime(fPath & "\" & fDir)
            fDir = Dir
        Wend
        pbFiles.Value = 0
        tvMemoLevel = tvMemoLevel + 1
        tvMemo(tvMemoLevel) = tvTreeView.Path
        sbStatusBar.Panels(1).Text = "Finished loading files."
    End With
Exit Sub
e:
    Err_Raise Err.Number, Err.Description, "frm_main", "tvtreeviewvhange": Resume Next
End Sub
Private Function tvStripPath(t$) As String
    Dim x%, ct%
    tvStripPath = t$
    x% = InStr(t$, "\")
    Do While x%
        ct% = x%
        x% = InStr(ct% + 1, t$, "\")
    Loop
    If ct% > 0 Then tvStripPath = Mid$(t$, ct% + 1)
End Function


Public Sub tvExtractIcon(FileName As String, PictureBox As PictureBox)
On Error GoTo e
    Dim Icon As Long
    Icon = SHGetFileInfo(FileName, 0&, IFileInfo, Len(IFileInfo), IFlags Or SHGFI_SMALLICON)
    If Icon <> 0 Then
      With PictureBox
        .Picture = LoadPicture("")
        Icon = ImageList_Draw(Icon, IFileInfo.iIcon, PictureBox.hDC, 0, 0, ILD_TRANSPARENT)
        .Refresh
      End With
    End If
Exit Sub
e:
    Err_Raise Err.Number, Err.Description, "frm_main", "tvextracticon": Resume Next
End Sub

Public Function tvAddIconToIML(FileName As String, FType As String, IML As ImageList) As Long
On Error GoTo e
    Dim I&
    If IsNumeric(FType) Then FType = "XXX"
    If FType = "EXE" Or FType = "ICO" Then
        Call tvExtractIcon(FileName, picIcon)
        tvAddIconToIML = IML.ListImages.Add(, , picIcon.Image).Index
    Else
        For I = 1 To IML.ListImages.Count
            If IML.ListImages(I).Key = FType Then
                tvAddIconToIML = I
                Exit Function
            End If
        Next I
        Call tvExtractIcon(FileName, picIcon)
        tvAddIconToIML = IML.ListImages.Add(, FType, picIcon.Image).Index
    End If
Exit Function
e:
    Err_Raise Err.Number, Err.Description, "frm_main", "tvaddicontoiml": Resume Next
End Function

Public Function OpenFile(FileName As String, Action As String) As Long
On Error GoTo e
    Dim Scr_hDC As Long
    Scr_hDC = GetDesktopWindow()
    OpenFile = ShellExecute(Scr_hDC, Action, FileName, "", Left(FileName, 3), 1)
Exit Function
e:
    Err_Raise Err.Number, Err.Description, "frm_main", "openfile": Resume Next
End Function

Public Function tvGetPath() As String
    If Right(tvTreeView.Path, 1) = "\" Then
        tvGetPath = Left(tvTreeView.Path, Len(tvTreeView.Path) - 1)
    Else
        tvGetPath = tvTreeView.Path
    End If
End Function

Public Function ShowDirs(Default As String) As String
    frmDirs.Show vbModal, Me
    frmDirs.tvDir.Path = Default
    ShowDirs = frmDirs.Path
End Function
Function XCopy(srcPath As String, dstPath As String, IncludeSubDirs As Integer, FilePat As String) As Integer
    'This is classic XCopy code, it can be done in few lines of code, but
    'this is faster, manual version :-) .
    
    Dim DirOK As Integer, I As Integer
    Dim DirReturn As String
    ReDim D(100) As String
    Dim dCount As Integer
    Dim CurrFile$
    Dim CurrDir$
    Dim dstPathBackup As String
    Dim f%

    On Error GoTo DirErr
    CurrDir$ = CurDir$
   
    If Right$(srcPath, 1) <> "\" Then srcPath = srcPath & "\"
    srcPath = UCase$(srcPath)
    
    If Right$(dstPath, 1) <> "\" Then dstPath = dstPath & "\"
    dstPath = UCase$(dstPath)
    dstPathBackup = dstPath
   
    'Dirs
    DirReturn = Dir(srcPath & "*.*", FILE_ATTRIBUTE_DIRECTORY)
    Do While DirReturn <> ""
        If DirReturn <> "." And DirReturn <> ".." Then
           'Well, here you see there is better way for attributes, VBA instead API
           If (GetAttr(srcPath & DirReturn) And FILE_ATTRIBUTE_DIRECTORY) = FILE_ATTRIBUTE_DIRECTORY Then
              dCount = dCount + 1
              D(dCount) = srcPath & DirReturn
           End If
        End If
        DirReturn = Dir
    Loop
   
    'Files
    DirReturn = Dir(srcPath & FilePat, 0)
    Do While DirReturn <> ""
        If Not ((GetAttr(srcPath & DirReturn) And FILE_ATTRIBUTE_DIRECTORY) = FILE_ATTRIBUTE_DIRECTORY) Then
            On Error Resume Next
            f% = FreeFile
            Open dstPath & DirReturn For Input As #f%
            Close #f%
            'Here is fast error handling, using numbers instead constants, and %,$,& to avoid Dim
            If Err = 0 Then
                f% = MsgBox("The file " & dstPath & DirReturn & " already exists. Do you wish to overwrite it?", 3 + 32 + 256)
                If f% = 6 Then FileCopy srcPath & DirReturn, dstPath & DirReturn
                If f% = 2 Then Exit Function
            Else
                FileCopy srcPath & DirReturn, dstPath & DirReturn
            End If
      End If
      DirReturn = Dir
   Loop
   
    'Now the hard part :-P
    For I = 1 To dCount
        If IncludeSubDirs Then
             On Error GoTo PathErr
             dstPath = dstPath & Right$(D(I), Len(D(I)) - Len(srcPath))
             ChDir dstPath
             On Error GoTo DirErr
        Else
            XCopy = True
            GoTo ExitFunc
        End If
        DirOK = XCopy(D(I), dstPath, IncludeSubDirs, FilePat)
        dstPath = dstPathBackup
    Next
    XCopy = True
    
ExitFunc:
   ChDir CurrDir$
   Exit Function
DirErr:
   XCopy = False
   Resume ExitFunc
PathErr:
   If Err = 75 Or Err = 76 Then
      MkDir dstPath
      Resume Next
   End If
   GoTo DirErr
End Function




Public Function XMove(srcPath As String, dstPath As String) As Long
On Error GoTo e
    'Now this is windows-default move, same as xcopy, but
    'MUCH easier version, use this in you apps, but there
    'are other operations, like delete or copy, AND you can
    'see that amazing animation box !! ;-) (not here...)
    Dim FileOperation As SHFILEOPSTRUCT   'Wanted operation
    
    srcPath = srcPath & Chr$(0) & Chr$(0)
    With FileOperation
       .wFunc = 1 'Move
       .pFrom = srcPath
       .pTo = dstPath
       .fFlags = FOF_SILENT
    End With
    XMove = SHFileOperation(FileOperation)
Exit Function
e:
    Err_Raise Err.Number, Err.Description, "frm_main", "xmove": Resume Next
End Function

Public Function tvCAttr(Value As Long) As String
    If Value > 0 Then tvCAttr = "YES" Else tvCAttr = "NOPE"
End Function
