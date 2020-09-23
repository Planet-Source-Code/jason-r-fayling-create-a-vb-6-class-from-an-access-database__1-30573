Attribute VB_Name = "modShellFileDisk"

' Module      : modShellFileDisk
' Description : Routines for working with the Windows 95/NT 4.0 shell
' Source      : Total VB SourceBook 6
' Update      : Code Service Pack 3
'
Private Const CSIDL_DESKTOP = &H0
Private Const CSIDL_PROGRAMS = &H2
Private Const CSIDL_CONTROLS = &H3
Private Const CSIDL_PRINTERS = &H4
Private Const CSIDL_PERSONAL = &H5
Private Const CSIDL_FAVORITES = &H6
Private Const CSIDL_STARTUP = &H7
Private Const CSIDL_RECENT = &H8
Private Const CSIDL_SENDTO = &H9
Private Const CSIDL_BITBUCKET = &HA
Private Const CSIDL_STARTMENU = &HB
Private Const CSIDL_DESKTOPDIRECTORY = &H10
Private Const CSIDL_DRIVES = &H11
Private Const CSIDL_NETWORK = &H12
Private Const CSIDL_NETHOOD = &H13
Private Const CSIDL_FONTS = &H14
Private Const CSIDL_TEMPLATES = &H15
Private Const CSIDL_COMMON_STARTMENU = &H16
Private Const CSIDL_COMMON_PROGRAMS = &H17
Private Const CSIDL_COMMON_STARTUP = &H18
Private Const CSIDL_COMMON_DESKTOPDIRECTORY = &H19
Private Const CSIDL_APPDATA = &H1A
Private Const CSIDL_PRINTHOOD = &H1B

Private Const BIF_RETURNONLYFSDIRS = &H1
Private Const BIF_NEWDIALOGSTYLE = &H40
Private Const BIF_EDITBOX = &H10
Private Const BIF_VALIDATE = &H20



Private Type BROWSEINFO
  hOwner As Long
  pidlRoot As Long
  pszDisplayName As String
  lpszTitle As String
  ulFlags As Long
  lpfn As Long
  lParam As Long
  lImage As Long
End Type

Private Type SHITEMID
  cb As Long
  abID As Byte
End Type

Private Type ITEMIDLIST
  mkid As SHITEMID
End Type

Private Type SHFILEOPSTRUCT
  hwnd As Long
  wFunc As Long
  pFrom As String
  pTo As String
  fFlags As Integer
  fAnyOperationsAborted As Long
  hNameMappings As Long
  lpszProgressTitle As String
End Type

Private Declare Sub CopyMemory _
  Lib "kernel32" _
  Alias "RtlMoveMemory" _
  (lpDest As Any, _
    lpSource As Any, _
    ByVal nCount As Long)

Private Declare Function SHGetPathFromIDList _
  Lib "shell32.dll" _
  Alias "SHGetPathFromIDListA" _
  (ByVal pidl As Long, _
    ByVal pszPath As String) _
  As Long

Private Declare Function SHGetSpecialFolderLocation _
  Lib "shell32.dll" _
  (ByVal hwndOwner As Long, _
    ByVal nFolder As Long, _
    pidl As ITEMIDLIST) _
  As Long

Private Declare Function SHFileOperation _
  Lib "shell32.dll" _
  Alias "SHFileOperationA" _
  (lpFileOp As Any) _
  As Long

Private Declare Function SHBrowseForFolder _
  Lib "shell32.dll" _
  Alias "SHBrowseForFolderA" _
  (lpBrowseInfo As BROWSEINFO) _
  As Long

Private Declare Sub SHAddToRecentDocs _
  Lib "shell32.dll" _
  (ByVal uFlags As Long, _
    ByVal pszPath As String)

Private Declare Function SHFormatDrive _
  Lib "shell32" _
  (ByVal hwnd As Long, _
    ByVal Drive As Long, _
    ByVal fmtID As Long, _
    ByVal options As Long) _
  As Long
    
Private Declare Function GetDriveType _
  Lib "kernel32" _
  Alias "GetDriveTypeA" _
  (ByVal nDrive As String) _
  As Long

Private Declare Sub CoTaskMemFree _
  Lib "ole32.dll" _
  (ByVal pv As Long)

Private Const FO_MOVE = &H1
Private Const FO_COPY = &H2
Private Const FO_DELETE = &H3
Private Const FO_RENAME = &H4

Private Const FOF_MULTIDESTFILES = &H1
Private Const FOF_CONFIRMMOUSE = &H2
Private Const FOF_SILENT = &H4
Private Const FOF_RENAMEONCOLLISION = &H8
Private Const FOF_NOCONFIRMATION = &H10
Private Const FOF_WANTMAPPINGHANDLE = &H20

Private Const FOF_ALLOWUNDO = &H40
Private Const FOF_FILESONLY = &H80
Private Const FOF_SIMPLEPROGRESS = &H100
Private Const FOF_NOCONFIRMMKDIR = &H200
Private Const FOF_NOERRORUI = &H400
Private Const SHARD_PATH = &H2&

' GetDriveType return values
Private Const DRIVE_NO_ROOT_DIR = 1
Private Const DRIVE_REMOVABLE = 2
Private Const DRIVE_FIXED = 3
Private Const DRIVE_REMOTE = 4
Private Const DRIVE_CDROM = 5
Private Const DRIVE_RAMDISK = 6

Private Const SHFMT_OPT_FULL = &H1
Private Const SHFMT_OPT_SYSONLY = &H2

Public Sub AddToRecentDocs(strFileName As String)
  ' Comments  : Adds a file to the 'Documents' submenu on the
  '             Windows Start menu
  ' Parameters: strFileName - full path to the document. The file must
  '             have a registered extension
  ' Returns   : Nothing
  ' Source    : Total VB SourceBook 6
  '
  On Error GoTo PROC_ERR
  
  SHAddToRecentDocs SHARD_PATH, strFileName

PROC_EXIT:
  Exit Sub

PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
    "AddToRecentDocs"
  Resume PROC_EXIT
  
End Sub

Public Function BrowseForFolder( _
  lnghWnd As Long, _
  strMessage As String, _
  Optional strDefault As String) _
  As String
  ' Comments  : Prompts the user for the location of an
  '             existing directory
  ' Parameters: lnghWnd - handle to window to serve as
  '             the parent for the dialog. Use a form's
  '             hWnd property for example
  '             strMessage - prompt message to display on the
  '             dialog
  '             strDefault - value to return if the user
  '             hits 'cancel' to close the dialog
  ' Returns   : The path the user selected
  ' Source    : Total VB SourceBook 6
  '
  Dim biFolder As BROWSEINFO
  Dim idlList As ITEMIDLIST
  Dim lngIDLptr As Long
  Dim lngResult As Long
  Dim strPath As String

  On Error GoTo PROC_ERR

  ' get the location of the user's desktop
  SHGetSpecialFolderLocation lnghWnd, CSIDL_DESKTOP, idlList

  ' set BROWSEINFO options
  With biFolder
    .hOwner = lnghWnd
    .pidlRoot = idlList.mkid.cb
    .lpszTitle = strMessage
    .ulFlags = BIF_RETURNONLYFSDIRS Or BIF_NEWDIALOGSTYLE
    

  End With
  
  ' Show the browse for folder dialog
  lngIDLptr = SHBrowseForFolder(biFolder)
  
  ' Get the Path indicated in the id list
  strPath = Space$(260)
  lngResult = SHGetPathFromIDList( _
    ByVal lngIDLptr, ByVal strPath)
  
  If lngResult <> 0 Then
    strPath = Left$(strPath, InStr(1, strPath, vbNullChar) - 1)
  Else
    ' user hit 'cancel', use default
    strPath = strDefault
  End If
  
  BrowseForFolder = strPath

PROC_EXIT:
  Exit Function

PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
    "BrowseForFolder"
  Resume PROC_EXIT

End Function

Public Sub ClearRecentDocs()
  ' Comments  : Clears the list of recently-opened documents
  '             from the Windows Start menu
  ' Parameters: None
  ' Returns   : Nothing
  ' Source    : Total VB SourceBook 6
  '
  On Error GoTo PROC_ERR

  SHAddToRecentDocs SHARD_PATH, vbNullString

PROC_EXIT:
  Exit Sub

PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
    "ClearRecentDocs"
  Resume PROC_EXIT

End Sub

Public Function GetShellAppdataLocation( _
  lnghWnd As Long) _
  As String
  ' Comments  : Returns the path of the user's
  '             "Appdata" folder
  ' Parameters: lnghWnd - handle to window to serve as
  '             the parent for the dialog. Use a form's
  '             hWnd property for example
  ' Returns   : Path of the user's Appdata folder
  ' Source    : Total VB SourceBook 6
  '
  Dim lngResult As Long
  Dim strPath As String
  Dim idlist As ITEMIDLIST
  
  On Error GoTo PROC_ERR
    
  ' populate an ITEMIDLIST struct with the specified folder information
  lngResult = SHGetSpecialFolderLocation( _
    lnghWnd, CSIDL_APPDATA, idlist)
    
  If lngResult = 0 Then
    
    ' if the information is present, get the path information
    strPath = Space$(260)
    lngResult = SHGetPathFromIDList( _
        ByVal idlist.mkid.cb, _
        ByVal strPath)
      
    ' free memory allocated by shell
    CoTaskMemFree idlist.mkid.cb
    
    'if a path was found, trim off trailing null char
    strPath = Left$(strPath, InStr(1, strPath, vbNullChar) - 1)
    GetShellAppdataLocation = strPath
    
  End If
  
PROC_EXIT:
  Exit Function

PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
    "GetShellAppdataLocation"
  Resume PROC_EXIT
  
End Function

Public Function GetShellCommonDesktopDirectoryLocation( _
  lnghWnd As Long) _
  As String
  ' Comments  : Returns the path of the user's
  '             "CommonDesktopDirectory" folder
  ' Parameters: lnghWnd - handle to window to serve as
  '             the parent for the dialog. Use a form's
  '             hWnd property for example
  ' Returns   : Path of the user's CommonDesktopDirectory folder
  ' Source    : Total VB SourceBook 6
  '
  Dim lngResult As Long
  Dim strPath As String
  Dim idlist As ITEMIDLIST
  
  On Error GoTo PROC_ERR
    
  ' populate an ITEMIDLIST struct with the specified folder information
  lngResult = SHGetSpecialFolderLocation( _
    lnghWnd, CSIDL_COMMON_DESKTOPDIRECTORY, idlist)
    
  If lngResult = 0 Then
    
    ' if the information is present, get the path information
    strPath = Space$(260)
    lngResult = SHGetPathFromIDList( _
        ByVal idlist.mkid.cb, _
        ByVal strPath)
      
    ' free memory allocated by shell
    CoTaskMemFree idlist.mkid.cb
    
    ' free memory allocated by shell
    CoTaskMemFree idlist.mkid.cb
    
    'if a path was found, trim off trailing null char
    strPath = Left$(strPath, InStr(1, strPath, vbNullChar) - 1)
    GetShellCommonDesktopDirectoryLocation = strPath
    
  End If
  
PROC_EXIT:
  Exit Function

PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
    "GetShellCommonDesktopDirectoryLocation"
  Resume PROC_EXIT
  
End Function

Public Function GetShellCommonProgramsLocation( _
  lnghWnd As Long) _
  As String
  ' Comments  : Returns the path of the user's
  '             "CommonPrograms" folder
  ' Parameters: lnghWnd - handle to window to serve as
  '             the parent for the dialog. Use a form's
  '             hWnd property for example
  ' Returns   : Path of the user's CommonPrograms folder
  ' Source    : Total VB SourceBook 6
  '
  Dim lngResult As Long
  Dim strPath As String
  Dim idlist As ITEMIDLIST
  
  On Error GoTo PROC_ERR
    
  ' populate an ITEMIDLIST struct with the specified folder information
  lngResult = SHGetSpecialFolderLocation( _
    lnghWnd, CSIDL_COMMON_PROGRAMS, idlist)
    
  If lngResult = 0 Then
    
    ' if the information is present, get the path information
    strPath = Space$(260)
    lngResult = SHGetPathFromIDList( _
        ByVal idlist.mkid.cb, _
        ByVal strPath)
      
    ' free memory allocated by shell
    CoTaskMemFree idlist.mkid.cb
    
    'if a path was found, trim off trailing null char
    strPath = Left$(strPath, InStr(1, strPath, vbNullChar) - 1)
    GetShellCommonProgramsLocation = strPath
    
  End If
  
PROC_EXIT:
  Exit Function

PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
    "GetShellCommonProgramsLocation"
  Resume PROC_EXIT
  
End Function

Public Function GetShellCommonStartMenuLocation( _
  lnghWnd As Long) _
  As String
  ' Comments  : Returns the path of the user's
  '             "CommonStartMenu" folder
  ' Parameters: lnghWnd - handle to window to serve as
  '             the parent for the dialog. Use a form's
  '             hWnd property for example
  ' Returns   : Path of the user's CommonStartMenu folder
  ' Source    : Total VB SourceBook 6
  '
  Dim lngResult As Long
  Dim strPath As String
  Dim idlist As ITEMIDLIST
  
  On Error GoTo PROC_ERR
    
  ' populate an ITEMIDLIST struct with the specified folder information
  lngResult = SHGetSpecialFolderLocation( _
    lnghWnd, CSIDL_COMMON_STARTMENU, idlist)
    
  If lngResult = 0 Then
    
    ' if the information is present, get the path information
    strPath = Space$(260)
    lngResult = SHGetPathFromIDList( _
        ByVal idlist.mkid.cb, _
        ByVal strPath)
      
    ' free memory allocated by shell
    CoTaskMemFree idlist.mkid.cb
    
    'if a path was found, trim off trailing null char
    strPath = Left$(strPath, InStr(1, strPath, vbNullChar) - 1)
    GetShellCommonStartMenuLocation = strPath
    
  End If
  
PROC_EXIT:
  Exit Function

PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
    "GetShellCommonStartMenuLocation"
  Resume PROC_EXIT
  
End Function

Public Function GetShellCommonStartupLocation( _
  lnghWnd As Long) _
  As String
  ' Comments  : Returns the path of the user's
  '             "CommonStartup" folder
  ' Parameters: lnghWnd - handle to window to serve as
  '             the parent for the dialog. Use a form's
  '             hWnd property for example
  ' Returns   : Path of the user's CommonStartup folder
  ' Source    : Total VB SourceBook 6
  '
  Dim lngResult As Long
  Dim strPath As String
  Dim idlist As ITEMIDLIST
  
  On Error GoTo PROC_ERR
    
  ' populate an ITEMIDLIST struct with the specified folder information
  lngResult = SHGetSpecialFolderLocation( _
    lnghWnd, CSIDL_COMMON_STARTUP, idlist)
    
  If lngResult = 0 Then
    
    ' if the information is present, get the path information
    strPath = Space$(260)
    lngResult = SHGetPathFromIDList( _
        ByVal idlist.mkid.cb, _
        ByVal strPath)
      
    ' free memory allocated by shell
    CoTaskMemFree idlist.mkid.cb
    
    'if a path was found, trim off trailing null char
    strPath = Left$(strPath, InStr(1, strPath, vbNullChar) - 1)
    GetShellCommonStartupLocation = strPath
    
  End If
  
PROC_EXIT:
  Exit Function

PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
    "GetShellCommonStartupLocation"
  Resume PROC_EXIT
  
End Function

Public Function GetShellDesktopDirectoryLocation( _
  lnghWnd As Long) _
  As String
  ' Comments  : Returns the path of the user's
  '             "DesktopDirectory" folder
  ' Parameters: lnghWnd - handle to window to serve as
  '             the parent for the dialog. Use a form's
  '             hWnd property for example
  ' Returns   : Path of the user's DesktopDirectory folder
  ' Source    : Total VB SourceBook 6
  '
  Dim lngResult As Long
  Dim strPath As String
  Dim idlist As ITEMIDLIST
  
  On Error GoTo PROC_ERR
    
  ' populate an ITEMIDLIST struct with the specified folder information
  lngResult = SHGetSpecialFolderLocation( _
    lnghWnd, CSIDL_DESKTOPDIRECTORY, idlist)
    
  If lngResult = 0 Then
    
    ' if the information is present, get the path information
    strPath = Space$(260)
    lngResult = SHGetPathFromIDList( _
        ByVal idlist.mkid.cb, _
        ByVal strPath)
      
    ' free memory allocated by shell
    CoTaskMemFree idlist.mkid.cb
    
    'if a path was found, trim off trailing null char
    strPath = Left$(strPath, InStr(1, strPath, vbNullChar) - 1)
    GetShellDesktopDirectoryLocation = strPath
    
  End If
  
PROC_EXIT:
  Exit Function

PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
    "GetShellDesktopDirectoryLocation"
  Resume PROC_EXIT
  
End Function

Public Function GetShellDesktopLocation( _
  lnghWnd As Long) _
  As String
  ' Comments  : Returns the path of the user's
  '             "Desktop" folder
  ' Parameters: lnghWnd - handle to window to serve as
  '             the parent for the dialog. Use a form's
  '             hWnd property for example
  ' Returns   : Path of the user's Desktop folder
  ' Source    : Total VB SourceBook 6
  '
  Dim lngResult As Long
  Dim strPath As String
  Dim idlist As ITEMIDLIST
  
  On Error GoTo PROC_ERR
    
  ' populate an ITEMIDLIST struct with the specified folder information
  lngResult = SHGetSpecialFolderLocation( _
    lnghWnd, CSIDL_DESKTOP, idlist)
    
  If lngResult = 0 Then
    
    ' if the information is present, get the path information
    strPath = Space$(260)
    lngResult = SHGetPathFromIDList( _
        ByVal idlist.mkid.cb, _
        ByVal strPath)
      
    ' free memory allocated by shell
    CoTaskMemFree idlist.mkid.cb
    
    'if a path was found, trim off trailing null char
    strPath = Left$(strPath, InStr(1, strPath, vbNullChar) - 1)
    GetShellDesktopLocation = strPath
    
  End If
  
PROC_EXIT:
  Exit Function

PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
    "GetShellDesktopLocation"
  Resume PROC_EXIT
  
End Function

Public Function GetShellFavoritesLocation( _
  lnghWnd As Long) _
  As String
  ' Comments  : Returns the path of the user's
  '             "Favorites" folder
  ' Parameters: lnghWnd - handle to window to serve as
  '             the parent for the dialog. Use a form's
  '             hWnd property for example
  ' Returns   : Path of the user's Favorites folder
  ' Source    : Total VB SourceBook 6
  '
  Dim lngResult As Long
  Dim strPath As String
  Dim idlist As ITEMIDLIST
  
  On Error GoTo PROC_ERR
    
  ' populate an ITEMIDLIST struct with the specified folder information
  lngResult = SHGetSpecialFolderLocation( _
    lnghWnd, CSIDL_FAVORITES, idlist)
    
  If lngResult = 0 Then
    
    ' if the information is present, get the path information
    strPath = Space$(260)
    lngResult = SHGetPathFromIDList( _
        ByVal idlist.mkid.cb, _
        ByVal strPath)
        
    ' free memory allocated by shell
    CoTaskMemFree idlist.mkid.cb
    
    'if a path was found, trim off trailing null char
    strPath = Left$(strPath, InStr(1, strPath, vbNullChar) - 1)
    GetShellFavoritesLocation = strPath
    
  End If
  
PROC_EXIT:
  Exit Function

PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
    "GetShellFavoritesLocation"
  Resume PROC_EXIT
  
End Function

Public Function GetShellFontsLocation( _
  lnghWnd As Long) _
  As String
  ' Comments  : Returns the path of the user's
  '             "Fonts" folder
  ' Parameters: lnghWnd - handle to window to serve as
  '             the parent for the dialog. Use a form's
  '             hWnd property for example
  ' Returns   : Path of the user's Fonts folder
  ' Source    : Total VB SourceBook 6
  '
  Dim lngResult As Long
  Dim strPath As String
  Dim idlist As ITEMIDLIST
  
  On Error GoTo PROC_ERR
    
  ' populate an ITEMIDLIST struct with the specified folder information
  lngResult = SHGetSpecialFolderLocation( _
    lnghWnd, CSIDL_FONTS, idlist)
    
  If lngResult = 0 Then
    
    ' if the information is present, get the path information
    strPath = Space$(260)
    lngResult = SHGetPathFromIDList( _
        ByVal idlist.mkid.cb, _
        ByVal strPath)
      
    ' free memory allocated by shell
    CoTaskMemFree idlist.mkid.cb
    
    'if a path was found, trim off trailing null char
    strPath = Left$(strPath, InStr(1, strPath, vbNullChar) - 1)
    GetShellFontsLocation = strPath
    
  End If
  
PROC_EXIT:
  Exit Function

PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
    "GetShellFontsLocation"
  Resume PROC_EXIT
  
End Function

Public Function GetShellPersonalLocation( _
  lnghWnd As Long) _
  As String
  ' Comments  : Returns the path of the user's
  '             "Personal" folder
  ' Parameters: lnghWnd - handle to window to serve as
  '             the parent for the dialog. Use a form's
  '             hWnd property for example
  ' Returns   : Path of the user's Personal folder
  ' Source    : Total VB SourceBook 6
  '
  Dim lngResult As Long
  Dim strPath As String
  Dim idlist As ITEMIDLIST
  
  On Error GoTo PROC_ERR
    
  ' populate an ITEMIDLIST struct with the specified folder information
  lngResult = SHGetSpecialFolderLocation( _
    lnghWnd, CSIDL_PERSONAL, idlist)
    
  If lngResult = 0 Then
    
    ' if the information is present, get the path information
    strPath = Space$(260)
    lngResult = SHGetPathFromIDList( _
        ByVal idlist.mkid.cb, _
        ByVal strPath)
      
    ' free memory allocated by shell
    CoTaskMemFree idlist.mkid.cb
    
    'if a path was found, trim off trailing null char
    strPath = Left$(strPath, InStr(1, strPath, vbNullChar) - 1)
    GetShellPersonalLocation = strPath
    
  End If
  
PROC_EXIT:
  Exit Function

PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
    "GetShellPersonalLocation"
  Resume PROC_EXIT
  
End Function

Public Function GetShellProgramsLocation( _
  lnghWnd As Long) _
  As String
  ' Comments  : Returns the path of the user's
  '             "Programs" folder
  ' Parameters: lnghWnd - handle to window to serve as
  '             the parent for the dialog. Use a form's
  '             hWnd property for example
  ' Returns   : Path of the user's Programs folder
  ' Source    : Total VB SourceBook 6
  '
  Dim lngResult As Long
  Dim strPath As String
  Dim idlist As ITEMIDLIST
  
  On Error GoTo PROC_ERR
    
  ' populate an ITEMIDLIST struct with the specified folder information
  lngResult = SHGetSpecialFolderLocation( _
    lnghWnd, CSIDL_PROGRAMS, idlist)
    
  If lngResult = 0 Then
    
    ' if the information is present, get the path information
    strPath = Space$(260)
    lngResult = SHGetPathFromIDList( _
        ByVal idlist.mkid.cb, _
        ByVal strPath)
      
    ' free memory allocated by shell
    CoTaskMemFree idlist.mkid.cb
    
    'if a path was found, trim off trailing null char
    strPath = Left$(strPath, InStr(1, strPath, vbNullChar) - 1)
    GetShellProgramsLocation = strPath
    
  End If
  
PROC_EXIT:
  Exit Function

PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
    "GetShellProgramsLocation"
  Resume PROC_EXIT
  
End Function

Public Function GetShellRecentLocation( _
  lnghWnd As Long) _
  As String
  ' Comments  : Returns the path of the user's
  '             "Recent" folder
  ' Parameters: lnghWnd - handle to window to serve as
  '             the parent for the dialog. Use a form's
  '             hWnd property for example
  ' Returns   : Path of the user's Recent folder
  ' Source    : Total VB SourceBook 6
  '
  Dim lngResult As Long
  Dim strPath As String
  Dim idlist As ITEMIDLIST
  
  On Error GoTo PROC_ERR
    
  ' populate an ITEMIDLIST struct with the specified folder information
  lngResult = SHGetSpecialFolderLocation( _
    lnghWnd, CSIDL_RECENT, idlist)
    
  If lngResult = 0 Then
    
    ' if the information is present, get the path information
    strPath = Space$(260)
    lngResult = SHGetPathFromIDList( _
        ByVal idlist.mkid.cb, _
        ByVal strPath)
      
    ' free memory allocated by shell
    CoTaskMemFree idlist.mkid.cb
        
    'if a path was found, trim off trailing null char
    strPath = Left$(strPath, InStr(1, strPath, vbNullChar) - 1)
    GetShellRecentLocation = strPath
    
  End If
  
PROC_EXIT:
  Exit Function

PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
    "GetShellRecentLocation"
  Resume PROC_EXIT
  
End Function

Public Function GetShellSendToLocation( _
  lnghWnd As Long) _
  As String
  ' Comments  : Returns the path of the user's
  '             "SendTo" folder
  ' Parameters: lnghWnd - handle to window to serve as
  '             the parent for the dialog. Use a form's
  '             hWnd property for example
  ' Returns   : Path of the user's SendTo folder
  ' Source    : Total VB SourceBook 6
  '
  Dim lngResult As Long
  Dim strPath As String
  Dim idlist As ITEMIDLIST
  
  On Error GoTo PROC_ERR
    
  ' populate an ITEMIDLIST struct with the specified folder information
  lngResult = SHGetSpecialFolderLocation( _
    lnghWnd, CSIDL_SENDTO, idlist)
    
  If lngResult = 0 Then
    
    ' if the information is present, get the path information
    strPath = Space$(260)
    lngResult = SHGetPathFromIDList( _
        ByVal idlist.mkid.cb, _
        ByVal strPath)
      
    ' free memory allocated by shell
    CoTaskMemFree idlist.mkid.cb
    
    'if a path was found, trim off trailing null char
    strPath = Left$(strPath, InStr(1, strPath, vbNullChar) - 1)
    GetShellSendToLocation = strPath
    
  End If
  
PROC_EXIT:
  Exit Function

PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
    "GetShellSendToLocation"
  Resume PROC_EXIT
  
End Function

Public Function GetShellStartMenuLocation( _
  lnghWnd As Long) _
  As String
  ' Comments  : Returns the path of the user's
  '             "StartMenu" folder
  ' Parameters: lnghWnd - handle to window to serve as
  '             the parent for the dialog. Use a form's
  '             hWnd property for example
  ' Returns   : Path of the user's StartMenu folder
  ' Source    : Total VB SourceBook 6
  '
  Dim lngResult As Long
  Dim strPath As String
  Dim idlist As ITEMIDLIST
  
  On Error GoTo PROC_ERR
    
  ' populate an ITEMIDLIST struct with the specified folder information
  lngResult = SHGetSpecialFolderLocation( _
    lnghWnd, CSIDL_STARTMENU, idlist)
    
  If lngResult = 0 Then
    
    ' if the information is present, get the path information
    strPath = Space$(260)
    lngResult = SHGetPathFromIDList( _
        ByVal idlist.mkid.cb, _
        ByVal strPath)
      
    ' free memory allocated by shell
    CoTaskMemFree idlist.mkid.cb
    
    'if a path was found, trim off trailing null char
    strPath = Left$(strPath, InStr(1, strPath, vbNullChar) - 1)
    GetShellStartMenuLocation = strPath
    
  End If
  
PROC_EXIT:
  Exit Function

PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
    "GetShellStartMenuLocation"
  Resume PROC_EXIT
  
End Function

Public Function GetShellStartupLocation( _
  lnghWnd As Long) _
  As String
  ' Comments  : Returns the path of the user's
  '             "Startup" folder
  ' Parameters: lnghWnd - handle to window to serve as
  '             the parent for the dialog. Use a form's
  '             hWnd property for example
  ' Returns   : Path of the user's Startup folder
  ' Source    : Total VB SourceBook 6
  '
  Dim lngResult As Long
  Dim strPath As String
  Dim idlist As ITEMIDLIST
  
  On Error GoTo PROC_ERR
    
  ' populate an ITEMIDLIST struct with the specified folder information
  lngResult = SHGetSpecialFolderLocation( _
    lnghWnd, CSIDL_STARTUP, idlist)
    
  If lngResult = 0 Then
    
    ' if the information is present, get the path information
    strPath = Space$(260)
    lngResult = SHGetPathFromIDList( _
        ByVal idlist.mkid.cb, _
        ByVal strPath)
      
    ' free memory allocated by shell
    CoTaskMemFree idlist.mkid.cb
          
    'if a path was found, trim off trailing null char
    strPath = Left$(strPath, InStr(1, strPath, vbNullChar) - 1)
    GetShellStartupLocation = strPath
    
  End If
  
PROC_EXIT:
  Exit Function

PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
    "GetShellStartupLocation"
  Resume PROC_EXIT
  
End Function

Public Function GetShellTemplatesLocation( _
  lnghWnd As Long) _
  As String
  ' Comments  : Returns the path of the user's
  '             "Templates" folder
  ' Parameters: lnghWnd - handle to window to serve as
  '             the parent for the dialog. Use a form's
  '             hWnd property for example
  ' Returns   : Path of the user's Templates folder
  ' Source    : Total VB SourceBook 6
  '
  Dim lngResult As Long
  Dim strPath As String
  Dim idlist As ITEMIDLIST
  
  On Error GoTo PROC_ERR
    
  ' populate an ITEMIDLIST struct with the specified folder information
  lngResult = SHGetSpecialFolderLocation( _
    lnghWnd, CSIDL_TEMPLATES, idlist)
    
  If lngResult = 0 Then
    
    ' if the information is present, get the path information
    strPath = Space$(260)
    lngResult = SHGetPathFromIDList( _
        ByVal idlist.mkid.cb, _
        ByVal strPath)
      
    ' free memory allocated by shell
    CoTaskMemFree idlist.mkid.cb
    
    'if a path was found, trim off trailing null char
    strPath = Left$(strPath, InStr(1, strPath, vbNullChar) - 1)
    GetShellTemplatesLocation = strPath
    
  End If
  
PROC_EXIT:
  Exit Function

PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
    "GetShellTemplatesLocation"
  Resume PROC_EXIT
  
End Function

Public Sub ShellCopyFile( _
  lnghWnd As Long, _
  ByVal strSource As String, _
  ByVal strDestination As String, _
  Optional ByVal fSilent As Boolean = False, _
  Optional strTitle As String = "")
  ' Comments  : Copies a file or files to a single destination
  ' Parameters: lnghWnd - handle to window to serve as
  '             the parent for the dialog. Use a form's
  '             hWnd property for example
  '             strSource - file spec for files to copy
  '             strDestination - destination file name or directory
  '             fSilent - if true, no warnings are displayed
  '             strTitle - title of the progress dialog
  ' Returns   : Nothing
  ' Source    : Total VB SourceBook 6
  ' Update    : Code Service Pack 3
  '
  Dim foCopy As SHFILEOPSTRUCT
  Dim lngFlags As Long
  Dim lngResult As Long
  Dim lngStructLen As Long
  Dim abytBuf() As Byte
    
  On Error GoTo PROC_ERR
  
  ' check to be sure file exists
  If Dir$(strSource) <> "" Then
    
    ' set flags for no prompting
    If fSilent Then
      lngFlags = FOF_NOCONFIRMMKDIR Or FOF_NOCONFIRMATION Or FOF_SILENT
    End If
    
    lngStructLen = LenB(foCopy)
    ReDim abytBuf(1 To lngStructLen)
  
    ' set shell file operations settings
    With foCopy
      .hwnd = lnghWnd
      .pFrom = strSource & vbNullChar & vbNullChar
      .pTo = strDestination & vbNullChar & vbNullChar
      .fFlags = lngFlags
      .lpszProgressTitle = strTitle & vbNullChar & vbNullChar
      .wFunc = FO_COPY
      
      If strTitle <> "" Then
        .fFlags = .fFlags Or FOF_SIMPLEPROGRESS
        
        ' Adjust alignment by copying to byte array
        Call CopyMemory(abytBuf(1), foCopy, lngStructLen)
        Call CopyMemory(abytBuf(19), abytBuf(21), 12)
      
        lngResult = SHFileOperation(abytBuf(1))
      Else
        lngResult = SHFileOperation(foCopy)
      End If
    
    End With
    
  End If

PROC_EXIT:
  Exit Sub

PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
    "ShellCopyFile"
  Resume PROC_EXIT

End Sub

Public Function ShellFormatDisk( _
  ByVal lnghWnd As Long, _
  ByVal strDriveLetter As String, _
  Optional fRemoveableOnly As Boolean = True, _
  Optional fQuickFormat As Boolean = True, _
  Optional fSysOnly As Boolean = False) _
  As Boolean
  ' Comments  : Shows the Windows shell format disk dialog
  ' Parameters: lnghWnd - handle to window to be parent of the dialog
  '             strDriveLetter - drive letter of the drive to format
  '             fRemoveableOnly - specifies whether the function should
  '             be limited to removeable media such as floppy drives
  '             fQuickFormat - specifies whether the default for the
  '             format operation should be to do a Quick Format
  '             fSysOnly - specifies whether the default should be to
  '             format the disk with system files only
  ' Returns   : True if the format was executed, otherwise false
  ' Source    : Total VB SourceBook 6
  '
  Dim lngDriveNumber As Long
  Dim lngDriveType As Long
  Dim lngResult As Long
  Dim lngFlags As Long
  Dim fReturn As Boolean
  Dim fContinue As Boolean
  
  On Error GoTo PROC_ERR
  
  strDriveLetter = UCase(strDriveLetter)
  
  ' turn the drive letter into a drive number
  lngDriveNumber = (Asc(strDriveLetter) - 65)
    
  ' find out what type of drive it is
  lngDriveType = GetDriveType(strDriveLetter & ":\")
    
  ' determine whether to continue with the format
  If lngDriveType = DRIVE_REMOVABLE Then
    fContinue = True
  Else
    If fRemoveableOnly = False Then
      fContinue = True
    End If
  End If
  
  If fContinue Then
    ' set the flags to control the default options for the format dialog
    If fQuickFormat = False Then
      lngFlags = lngFlags Or SHFMT_OPT_FULL
    End If
    If fSysOnly = True Then
      lngFlags = lngFlags Or SHFMT_OPT_SYSONLY
    End If
    
    ' show the format dialog
    lngResult = SHFormatDrive(lnghWnd, lngDriveNumber, 0&, lngFlags)
    If lngResult Then
      fReturn = True
    End If
  Else
    fReturn = False
  End If

PROC_EXIT:
  Exit Function

PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
    "ShellFormatDisk"
  Resume PROC_EXIT

End Function

Public Sub ShellRecycleFile( _
  lnghWnd As Long, _
  ByVal strFileSpec As String, _
  Optional fUndoable As Boolean = True, _
  Optional strTitle As String = "")
  ' Comments  : Sends the specified file or files
  '             to the Windows 95/NT recycle bin
  ' Parameters: lnghWnd - handle to window to serve as the parent for the
  '             dialog. Use a form's hWnd property for example
  '             strFileSpec - full path to the file(s) todelete. May include
  '             wildcard characters
  '             fUndoable - If true, the files are permanently deleted
  '             strTitle - title of the progress dialog
  ' Returns   : Nothing
  ' Source    : Total VB SourceBook 6
  ' Update    : Code Service Pack 3
  '
  Dim foDelete As SHFILEOPSTRUCT
  Dim lngResult As Long
  Dim lngFlags As Long
  Dim lngStructLen As Long
  Dim abytBuf() As Byte
      
  On Error GoTo PROC_ERR
      
  ' skip empty file specs
  If Not strFileSpec = vbNullString Then
  
    lngStructLen = LenB(foDelete)
    ReDim abytBuf(1 To lngStructLen)
  
    ' set optional flag to permanently delete the files
    If fUndoable = True Then
      lngFlags = FOF_ALLOWUNDO
    End If
    
    With foDelete
      .hwnd = lnghWnd
      .wFunc = FO_DELETE
      .pFrom = strFileSpec & vbNullChar & vbNullChar
      .fFlags = lngFlags
      .lpszProgressTitle = strTitle & vbNullChar & vbNullChar
    
      If strTitle <> "" Then
        .fFlags = .fFlags Or FOF_SIMPLEPROGRESS
        
        ' Adjust alignment by copying to byte array
        Call CopyMemory(abytBuf(1), foDelete, lngStructLen)
        Call CopyMemory(abytBuf(19), abytBuf(21), 12)
      
        lngResult = SHFileOperation(abytBuf(1))
      Else
        lngResult = SHFileOperation(foDelete)
      End If

    End With

  End If
  
PROC_EXIT:
  Exit Sub

PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
    "ShellRecycleFile"
  Resume PROC_EXIT
  
End Sub

