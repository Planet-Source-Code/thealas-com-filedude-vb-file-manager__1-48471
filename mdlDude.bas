Attribute VB_Name = "mdlDude"
Option Explicit

Public Declare Function OSWinHelp% Lib "user32" Alias "WinHelpA" (ByVal hWnd&, ByVal HelpFile$, ByVal wCommand%, dwData As Any)
Public Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Public Declare Function GetFileType Lib "kernel32" (ByVal hFile As Long) As Long
Public Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
Public Declare Function SHGetFileInfo Lib "shell32.dll" Alias "SHGetFileInfoA" (ByVal pszPath As String, ByVal dwFileAttributes As Long, psfi As typSHFILEINFO, ByVal cbSizeFileInfo As Long, ByVal uFlags As Long) As Long
Public Declare Function ImageList_Draw Lib "comctl32.dll" (ByVal himl&, ByVal I&, ByVal hDCDest&, ByVal x&, ByVal y&, ByVal fLags&) As Long
Public Declare Function SHFileOperation Lib "shell32.dll" Alias "SHFileOperationA" (lpFileOp As SHFILEOPSTRUCT) As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare Function GetDesktopWindow Lib "user32" () As Long

'FO's are not used much sence I know numbers, but I've placed them here
'so you can use them.
Public Const FO_DELETE = &H3
Public Const FO_MOVE As Long = &H1
Public Const FO_COPY As Long = &H2
Public Const FO_RENAME As Long = &H4
Public Const FOF_ALLOWUNDO = &H40
Public Const FOF_SIMPLEPROGRESS = &H100
Public Const FOF_SILENT As Long = &H4
Public Const FILE_ATTRIBUTE_DIRECTORY = &H10
Public Const FILE_ATTRIBUTE_HIDDEN = &H2
Public Const FILE_ATTRIBUTE_NORMAL = &H80
Public Const FILE_ATTRIBUTE_READONLY = &H1
Public Const FILE_ATTRIBUTE_SYSTEM = &H4
Public Const SHGFI_DISPLAYNAME = &H200
Public Const SHGFI_EXETYPE = &H2000
Public Const SHGFI_SYSICONINDEX = &H4000
Public Const SHGFI_SHELLICONSIZE = &H4
Public Const SHGFI_TYPENAME = &H400
Public Const SHGFI_LARGEICON = &H0
Public Const SHGFI_SMALLICON = &H1
Public Const ILD_TRANSPARENT = &H1
Public Const IFlags = SHGFI_TYPENAME Or SHGFI_SHELLICONSIZE Or SHGFI_SYSICONINDEX Or SHGFI_DISPLAYNAME Or SHGFI_EXETYPE 'Too stuffs, just put it in decs

Public IFileInfo As typSHFILEINFO

Public Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type
Public Type WIN32_FIND_DATA
    dwFileAttributes As Long
    ftCreationTime As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime As FILETIME
    nFileSizeHigh As Long
    nFileSizeLow As Long
    dwReserved0 As Long
    dwReserved1 As Long
    cFileName As String * 260
    cAlternate As String * 14
End Type
Public Type typSHFILEINFO
    hIcon As Long
    iIcon As Long
    dwAttributes As Long
    szDisplayName As String * 260
    szTypeName As String * 80
End Type
Public Type SHFILEOPSTRUCT
     hWnd As Long
     wFunc As Long
     pFrom As String
     pTo As String
     fFlags As Integer
     fAnyOperationsAborted As Boolean
     hNameMappings As Long
     lpszProgressTitle As String
End Type

Public Sub Log_Append(Text As String)
On Error GoTo e 'Hehe...
    Open App.Path & "\ERRORS.LOG" For Append As #2
        Print #2, "DATE: " & Date & ",TIME: " & Time & " - " & Text
    Close #2
e:
    Exit Sub
End Sub

Public Sub Log_Clear()
    Open App.Path & "\ERRORS.LOG" For Output As #3
        Print #3, ""
    Close #3
End Sub

Public Sub Err_Raise(Err_Num As String, Err_Description As String, Err_Module As String, Err_Function As String)
    Log_Append "ERROR " & Err_Num & " - " & Err_Module & "\" & Err_Function & " ::: " & Err_Description
End Sub

