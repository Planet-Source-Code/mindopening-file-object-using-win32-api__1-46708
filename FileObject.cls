VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FileObject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'Event Declarations
Public Event Progress(Percent As Integer)

'Win32 API Declarations
Private Declare Function GetFileTime Lib "kernel32" (ByVal hFile As Long, lpCreationTime As FILETIME, lpLastAccessTime As FILETIME, lpLastWriteTime As FILETIME) As Long
Private Declare Function FileTimeToLocalFileTime Lib "kernel32" (lpFileTime As FILETIME, lpLocalFileTime As FILETIME) As Long
Private Declare Function GetFileAttributes Lib "kernel32" Alias "GetFileAttributesA" (ByVal lpFileSpec As String) As Long
Private Declare Function SetFileAttributes Lib "kernel32" Alias "SetFileAttributesA" (ByVal lpFileSpec As String, ByVal dwFileAttributes As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function FlushFileBuffers Lib "kernel32" (ByVal hFile As Long) As Long
Private Declare Function SHFileOperation Lib "shell32.dll" Alias "SHFileOperationA" (lpFileOp As SHFILEOPSTRUCT) As Long
Private Declare Function SHGetFileInfo Lib "shell32.dll" Alias "SHGetFileInfoA" (ByVal pszPath As String, ByVal dwFileAttributes As Long, psfi As SHFILEINFO, ByVal cbFileInfo As Long, ByVal uFlags As Long) As Long
Private Declare Function GetExpandedName Lib "lz32.dll" Alias "GetExpandedNameA" (ByVal lpszSource As String, ByVal lpszBuffer As String) As Long
Private Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long
Private Declare Function FileTimeToSystemTime Lib "kernel32" (lpFileTime As FILETIME, lpSystemTime As SYSTEMTIME) As Long
Private Declare Function OpenFile Lib "kernel32" (ByVal lpFileName As String, lpReOpenBuff As OFSTRUCT, ByVal wStyle As Long) As Long
Private Declare Function GetFileInformationByHandle Lib "kernel32" (ByVal hFile As Long, lpFileInformation As BY_HANDLE_FILE_INFORMATION) As Long
Private Declare Function CharLower Lib "user32" Alias "CharLowerA" (ByVal lpsz As String) As Long
Private Declare Function SystemTimeToFileTime Lib "kernel32.dll" (lpSystemTime As SYSTEMTIME, lpFileTime As FILETIME) As Long
Private Declare Function SetFileTime Lib "kernel32.dll" (ByVal hFile As Long, lpCreationTime As FILETIME, lpLastAccessTime As FILETIME, lpLastWriteTime As FILETIME) As Long
Private Declare Function SetFilePointer Lib "kernel32.dll" (ByVal iFileHandler As Long, ByVal lDistanceToMove As Long, lpDistanceToMoveHigh As Long, ByVal dwMoveMethod As Long) As Long
Private Declare Function WriteFile Lib "kernel32.dll" (ByVal iFileHandler As Long, lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, ByVal lpOverlapped As Long) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, ByVal lpOverlapped As Any) As Long
Private Declare Function GetFileSize Lib "kernel32" (ByVal hFile As Long, lpFileSizeHigh As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)

Private Const OFS_MAXPATHNAME = 128
Private Const GENERIC_READ = &H80000000
'File Destroy API Declarations
'CreateFile Flags
Private Const GENERIC_WRITE = &H40000000
Private Const TRUNCATE_EXISTING = 5
Private Const FILE_SHARE_WRITE = &H2
'Bypass any disk cache and instead read and write directly to the file.
Private Const FILE_FLAG_WRITE_THROUGH = &H80000000
'Do not use any buffers or caches
Private Const FILE_FLAG_NO_BUFFERING = &H20000000
'File Attributes
Private Const FILE_ATTRIBUTE_READONLY = &H1
Private Const FILE_ATTRIBUTE_HIDDEN = &H2
Private Const FILE_ATTRIBUTE_SYSTEM = &H4
Private Const FILE_ATTRIBUTE_ARCHIVE = &H20
Private Const FILE_ATTRIBUTE_NORMAL = &H80
Private Const FILE_ATTRIBUTE_TEMPORARY = &H100
'File Functions
Private Const FO_DELETE = &H3
Private Const FO_COPY = &H2
Private Const FO_MOVE = &H1
Private Const FO_RENAME = &H4
Private Const FOF_ALLOWUNDO = &H40
Private Const FOF_RENAMEONCOLLISION = &H8
Private Const FOF_NOCONFIRMATION = &H10       ' Don't prompt the user.
Private Const FOF_NOERRORUI = &H400
Private Const FOF_FILESONLY = &H80
Private Const FILE_SHARE_READ = &H1
Private Const OPEN_EXISTING = 3


Private Const CREATE_ALWAYS = 2
Private Const CREATE_NEW = 1
Private Const OPEN_ALWAYS = 4



Private Const FILE_FLAG_DELETE_ON_CLOSE = &H4000000

Private Const FILE_FLAG_OVERLAPPED = &H40000000
Private Const FILE_FLAG_POSIX_SEMANTICS = &H1000000
Private Const FILE_FLAG_RANDOM_ACCESS = &H10000000
Private Const FILE_FLAG_SEQUENTIAL_SCAN = &H8000000

Private Const FILE_BEGIN = 0
Private Const INVALID_HANDLE_VALUE = -1
'File Wiping
Private Const CWR_BUFFER As Long = 1048576


'File Execute
'Constant Values
Private Const SW_HIDE = 0
Private Const SW_SHOWNORMAL = 1
Private Const SW_SHOWMINIMIZED = 2
Private Const SW_SHOWMAXIMIZED = 3
Private Const SW_SHOWNOACTIVE = 4
Private Const SW_SHOW = 5
Private Const SW_MINIMIZE = 6
Private Const SW_SHOWMINNOACTIVE = 7
Private Const SW_SHOWNA = 8
Private Const SW_RESTORE = 9
Private Const SW_SHOWDEFAULT = 10

'Return Values
Private Const SE_ERR_FNF = 2&
Private Const SE_ERR_PNF = 3&
Private Const SE_ERR_ACCESSDENIED = 5&
Private Const SE_ERR_OOM = 8&
Private Const SE_ERR_DLLNOTFOUND = 32&
Private Const SE_ERR_SHARE = 26&
Private Const SE_ERR_ASSOCINCOMPLETE = 27&
Private Const SE_ERR_DDETIMEOUT = 28&
Private Const SE_ERR_DDEFAIL = 29&
Private Const SE_ERR_DDEBUSY = 30&
Private Const SE_ERR_NOASSOC = 31&
Private Const ERROR_BAD_FORMAT = 11&

Private Const SHGFI_ATTRIBUTES = &H800
Private Const SHGFI_DISPLAYNAME = &H200
Private Const SHGFI_EXETYPE = &H2000
Private Const SHGFI_ICON = &H100
Private Const SHGFI_ICONLOCATION = &H1000
Private Const SHGFI_LARGEICON = &H0
Private Const SHGFI_LINKOVERLAY = &H8000
Private Const SHGFI_OPENICON = &H2
Private Const SHGFI_PIDL = &H8
Private Const SHGFI_SELECTED = &H10000
Private Const SHGFI_SHELLICONSIZE = &H4
Private Const SHGFI_SMALLICON = &H1
Private Const SHGFI_SYSICONINDEX = &H4000
Private Const SHGFI_TYPENAME = &H400
Private Const SHGFI_USEFILEATTRIBUTES = &H10

'Constansts for Alder CRC 32 Method
Private Const BASE = 65521  ' largest prime smaller than 65536
Private Const NMAX = 5552   ' Largest n such that 255n(n+1)/2 + (n+1)(BASE-1) <= 2^32-1
Private Const CHUNK_SIZE = 2048

Private Enum CRCAlgorithms
    enmCRC32 = 0
    enmCRCAlder = 1
End Enum

Public Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

Private Type SHFILEINFO
    hIcon As Long
    iIcon As Long
    dwAttributes As Long
    szDisplayName As String * 260
    szTypeName As String * 80
End Type

Public Type OFSTRUCT
    cBytes As Byte
    fFixedDisk As Byte
    nErrCode As Integer
    Reserved1 As Integer
    Reserved2 As Integer
    szPathName(OFS_MAXPATHNAME) As Byte
End Type

Public Type BY_HANDLE_FILE_INFORMATION
    dwFileAttributes As Long
    ftCreationTime As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime As FILETIME
    dwVolumeSerialNumber As Long
    nFileSizeHigh As Long
    nFileSizeLow As Long
    nNumberOfLinks As Long
    nFileIndexHigh As Long
    nFileIndexLow As Long
End Type

Public Type SYSTEMTIME
    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Integer
End Type
Public Enum AccessRights
    FileRead = &H80000000                   ' Read Access
    FileWrite = &H40000000                  ' Write Access
    FileReadWrite = &HC0000000              ' Read and Write Access
    FileExe = &H20000000                ' Execute Access
    FileAllAccess = &H10000000              ' All Access Rights
End Enum

' =============================================================================
' Share Mode
' =============================================================================
Public Enum ShareMode
    FileShareNone = &H0                     ' No Sharing
    FileShareRead = &H1                     ' Share Read of File
    FileShareWrite = &H2                    ' Share Write of File
    FileShareReadWrite = &H3                ' Share Read/Write of File
End Enum

' =============================================================================
' Creation Distribution
' =============================================================================
Public Enum CreationDistribution
    FileCreateNew = &H1                     ' Create New (No Overwriting)
    FileCreateAlways = &H2                  ' Create New (Overwrite Existing)
    FileOpenExisting = &H3                  ' Open Existing
    FileOpenAlways = &H4                    ' Open Existing (Create New if not found)
    FileTruncateExisting = &H5              ' Append Existing
End Enum

' =============================================================================
' File Attributes
' =============================================================================
Public Enum FlagsAndAttributes
    FileAttributeReadOnly = &H1             ' Read-Only
    FileAttributeHidden = &H2               ' Hidden
    FileAttributeSystem = &H4               ' System
    FileAttributeDirectory = &H10           ' Directory
    FileAttributeArchive = &H20             ' Archive
    FileAttributeNormal = &H80              ' Normal
    FileAttributeTemporary = &H100          ' Temporary
    FileAttributeCompressed = &H800         ' Compressed
    FileAttributeAll = &H9B7                ' All Attributes
    FlagDeleteOnClose = &H4000000
    FlagNoBuffering = &H20000000
    FlagOverlapped = &H20000000
    FlagPosixSemantics = &H1000000
    FlagRandomAccess = &H10000000
    FlagSequentialScan = &H8000000
    FlagWriteThrough = &H80000000
End Enum


'File Operation API declarations
Private Type SHFILEOPSTRUCT
    hWnd As Long
    wFunc As Long
    pFrom As String
    pTo As String
    fFlags As Integer
    fAnyOperationsAborted As Long
    hNameMappings As Long
    lpszProgressTitle As Long
End Type


Enum DeleteType
    vKill = 0
    vRecycle = 1
    vDestroy = 2
    vWipe = 3
End Enum

Private mvarsFileName As String
Private mvarComputerName As String
Private mvarDrive As String
Private mvarHidden As Boolean
Private mvarReadOnly As Boolean
Private mCRCMethod As CRCAlgorithms
Private mvarFileSize As Long
Private mvarFileHandle As Long
Private mvarOpened As Boolean


'Variables for CRC
Private lngArrayCRCByte(0 To 255) As Long
Public Property Get IsOpened() As Boolean
    IsOpened = mvarOpened
End Property
Public Property Get FileHandle() As Long
    FileHandle = mvarFileHandle
End Property
Public Property Get FileSize() As Long
    FileSize = mvarFileSize
End Property

Public Property Get CRC32() As Long
    If LenB(mvarsFileName) > 0 Then
        If Exists Then
            CRC32 = CRCFile(enmCRC32)
        End If
    End If
End Property

Public Property Get CRCAlder32() As Long
    If LenB(mvarsFileName) > 0 Then
        If Exists Then
            CRCAlder32 = CRCFile(enmCRCAlder)
        End If
    End If
End Property
Public Property Get Hidden() As Boolean
    If LenB(mvarsFileName) > 0 Then
        If Exists Then
            Select Case GetAttr(mvarsFileName)
                Case 1
                    mvarReadOnly = True
                Case 2
                    mvarHidden = True
                Case 3
                    mvarHidden = True
                    mvarReadOnly = True
            End Select
        End If
    End If
                
End Property
Public Property Let Hidden(vData As Boolean)
    mvarHidden = vData
    SetFileAtt
End Property
Public Property Let ReadOnly(vData As Boolean)
    mvarReadOnly = vData
    SetFileAtt
End Property
Public Property Get FileType() As String
    Dim lngResult As Long
    Dim FileInfo As SHFILEINFO
    
    lngResult = SHGetFileInfo(mvarsFileName, &O0, FileInfo, Len(FileInfo), SHGFI_USEFILEATTRIBUTES Or SHGFI_TYPENAME)
    
    If lngResult > 0 Then
        FileType = Left(FileInfo.szTypeName, InStr(FileInfo.szTypeName, vbNullChar) - 1)
    Else
    
    End If
End Property
Public Property Get ReadOnly() As Boolean
    If LenB(mvarsFileName) > 0 Then
        If Exists Then
            Select Case GetAttr(mvarsFileName)
                Case 1
                    mvarReadOnly = True
                Case 2
                    mvarHidden = True
                Case 3
                    mvarHidden = True
                    mvarReadOnly = True
            End Select
        End If
    End If
End Property
Public Property Get DateModified() As Date
    DateModified = GetTimeModified(mvarsFileName)
End Property
Public Property Let DateModified(vData As Date)
    Dim cctime As SYSTEMTIME
    Dim ctime As FILETIME
    Dim atime As FILETIME
    Dim mTime As FILETIME
    Dim sTempTime As String
    Dim retVal As Long
    
    If LenB(mvarsFileName) > 0 Then
        If Exists Then
            cctime.wDay = Day(vData)
            cctime.wMonth = Month(vData)
            cctime.wYear = Year(vData)
            If SystemTimeToFileTime(cctime, mTime) = 0 Then
                If CreateFile(mvarsFileName, GENERIC_READ Or GENERIC_WRITE, FILE_SHARE_READ, ByVal CLng(0), OPEN_EXISTING, FILE_ATTRIBUTE_ARCHIVE, 0) = -1 Then
                    MsgBox "Can not change file date"
                Else
                    sTempTime = DateAccessed
                    cctime.wDay = CInt(Format(sTempTime, "dd"))
                    cctime.wMonth = CInt(Format(sTempTime, "mm"))
                    cctime.wYear = CInt(Format(sTempTime, "yyyy"))
                    SystemTimeToFileTime cctime, atime
                    sTempTime = DateCreated
                    cctime.wDay = CInt(Format(sTempTime, "dd"))
                    cctime.wMonth = CInt(Format(sTempTime, "mm"))
                    cctime.wYear = CInt(Format(sTempTime, "yyyy"))
                    SystemTimeToFileTime cctime, ctime
                    
                    retVal = SetFileTime(mvarsFileName, ctime, atime, mTime)
                    retVal = CloseHandle(mvarsFileName)
                End If
            End If
        End If
    End If
    
End Property
Public Property Get DateAccessed() As Date
    DateAccessed = GetTimeAccessed(mvarsFileName)
End Property
Public Property Let DateAccessed(vData As Date)
    Dim cctime As SYSTEMTIME
    Dim ctime As FILETIME
    Dim atime As FILETIME
    Dim mTime As FILETIME
    Dim sTempTime As String
    Dim retVal As Long
    
    If LenB(mvarsFileName) > 0 Then
        If Exists Then
            cctime.wDay = CInt(Format(vData, "dd"))
            cctime.wMonth = CInt(Format(vData, "mm"))
            cctime.wYear = CInt(Format(vData, "yyyy"))
            If SystemTimeToFileTime(cctime, atime) = 0 Then
                If CreateFile(mvarsFileName, GENERIC_READ Or GENERIC_WRITE, FILE_SHARE_READ, ByVal CLng(0), OPEN_EXISTING, FILE_ATTRIBUTE_ARCHIVE, 0) = -1 Then
                    MsgBox "Can not change file date"
                Else
                    sTempTime = DateModified
                    cctime.wDay = CInt(Format(sTempTime, "dd"))
                    cctime.wMonth = CInt(Format(sTempTime, "mm"))
                    cctime.wYear = CInt(Format(sTempTime, "yyyy"))
                    SystemTimeToFileTime cctime, mTime
                    sTempTime = DateCreated
                    cctime.wDay = CInt(Format(sTempTime, "dd"))
                    cctime.wMonth = CInt(Format(sTempTime, "mm"))
                    cctime.wYear = CInt(Format(sTempTime, "yyyy"))
                    SystemTimeToFileTime cctime, ctime
                    retVal = SetFileTime(mvarsFileName, ctime, atime, mTime)
                    retVal = CloseHandle(mvarsFileName)
                End If
            End If
        End If
    End If
    
End Property
Public Property Get DateCreated() As Date
    DateCreated = GetTimeCreated(mvarsFileName)
End Property
Public Property Let DateCreated(vData As Date)
    Dim cctime As SYSTEMTIME
    Dim ctime As FILETIME
    Dim atime As FILETIME
    Dim mTime As FILETIME
    Dim sTempTime As String
    Dim retVal As Long
    
    If LenB(mvarsFileName) > 0 Then
        If Exists Then
            cctime.wDay = CInt(Format(vData, "dd"))
            cctime.wMonth = CInt(Format(vData, "mm"))
            cctime.wYear = CInt(Format(vData, "yyyy"))
            If SystemTimeToFileTime(cctime, ctime) = 0 Then
                If CreateFile(mvarsFileName, GENERIC_READ Or GENERIC_WRITE, FILE_SHARE_READ, ByVal CLng(0), OPEN_EXISTING, FILE_ATTRIBUTE_ARCHIVE, 0) = -1 Then
                    MsgBox "Can not change file date"
                Else
                    sTempTime = DateModified
                    cctime.wDay = CInt(Format(sTempTime, "dd"))
                    cctime.wMonth = CInt(Format(sTempTime, "mm"))
                    cctime.wYear = CInt(Format(sTempTime, "yyyy"))
                    SystemTimeToFileTime cctime, atime
                    sTempTime = DateAccessed
                    cctime.wDay = CInt(Format(sTempTime, "dd"))
                    cctime.wMonth = CInt(Format(sTempTime, "mm"))
                    cctime.wYear = CInt(Format(sTempTime, "yyyy"))
                    SystemTimeToFileTime cctime, atime
                    retVal = SetFileTime(mvarsFileName, ctime, atime, mTime)
                    retVal = CloseHandle(mvarsFileName)
                End If
            End If
        End If
    End If
End Property
Public Property Get Drive() As String
    Drive = GetDrive(mvarsFileName)
End Property
Public Property Get ComputerName() As String
    ComputerName = GetComputeName(mvarsFileName)
End Property
Public Property Get FixedDisk() As Boolean
    FixedDisk = fileOpenStructure(mvarsFileName).fFixedDisk
End Property
Public Property Get Title() As String
    Title = ExtractFileName(mvarsFileName)
End Property
Public Property Get LastAccessed() As Date
    LastAccessed = timeFileToDate(FileInformation(mvarsFileName).ftLastAccessTime)
End Property
Public Property Get LongName() As String
    LongName = GetLongName(mvarsFileName)
End Property
Public Property Get ShortName() As String
    ShortName = GetShortName(mvarsFileName)
End Property
Public Property Let FileName(ByVal vData As String)
    mvarsFileName = vData
End Property
Public Property Get FileName() As String
    FileName = mvarsFileName
End Property
Public Property Get Exists() As Boolean
    Exists = IsFileThere(mvarsFileName)
End Property
Public Property Get Extension() As String
    Extension = GetExt(mvarsFileName)
End Property
Public Property Get ParentDirectory() As String
    Dim lngPosition As Long
    Dim strTemp As String
    strTemp = Directory
    strTemp = StrReverse(strTemp)
    lngPosition = InStr(2, strTemp, "\")
    If lngPosition > 0 Then
        strTemp = Left$(strTemp, lngPosition)
        ParentDirectory = StrReverse(strTemp)
    Else
        ParentDirectory = ""
    End If
End Property
Public Property Get Directory() As String
    Directory = GetDirectory(mvarsFileName)
End Property
Private Function GetDirectory(ByVal strPath As String) As String
    Dim lngPosition As Long
    
    strPath = StrReverse(strPath)
    lngPosition = InStr(1, strPath, "\")
    If lngPosition > 0 Then
        strPath = Right$(strPath, Len(strPath) - lngPosition + 1)
        GetDirectory = StrReverse(strPath)
    Else
        GetDirectory = ""
    End If
End Function
Private Function SetFileAtt()
    Dim intFileState As Integer
    If LenB(mvarsFileName) > 0 Then
        If Exists Then
            If Not mvarHidden And Not mvarReadOnly Then
                intFileState = 0
            ElseIf Not mvarHidden And mvarReadOnly Then
                intFileState = 1
            ElseIf mvarHidden And Not mvarReadOnly Then
                intFileState = 2
            ElseIf mvarHidden And mvarReadOnly Then
                intFileState = 3
            End If
            SetAttr mvarsFileName, intFileState
        End If
    End If
End Function
Public Function FileDelete(typ As DeleteType, Optional sFileName As String = "", Optional blnConfirm As Boolean, Optional intPasses As Integer) As Boolean
    Dim SHop As SHFILEOPSTRUCT
    Dim lngResult As Long
        
    On Error GoTo LocalError
        
    If LenB(sFileName) = 0 Then sFileName = mvarsFileName
        
    With SHop
        .fFlags = FOF_NOCONFIRMATION
        Select Case typ
            Case vRecycle
                .fFlags = .fFlags + FOF_ALLOWUNDO
            Case vDestroy
                If Not (DestroyFile(sFileName)) Then
                End If
            Case vWipe
                SecureWipeFile sFileName, intPasses
                FileDelete = True
                Exit Function
        End Select
    
        .wFunc = FO_DELETE
        .pFrom = sFileName
        
    End With
    
    lngResult = SHFileOperation(SHop)
    If lngResult = 0 Then FileDelete = True
    Exit Function
LocalError:
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
End Function

Public Function FileCopy(strToFileName As String, Optional sFileName As String = "") As Boolean
    Dim SHop As SHFILEOPSTRUCT
    Dim lngResult As Long
    
    If LenB(sFileName) = 0 Then sFileName = mvarsFileName
    
    With SHop
        .wFunc = FO_COPY
        .pTo = strToFileName
        .pFrom = sFileName
        .fFlags = FOF_NOCONFIRMATION
    End With
    
    lngResult = SHFileOperation(SHop)
    If lngResult = 0 Then FileCopy = True
End Function

Public Function FileMove(strDestination As String, Optional sFileName As String) As Boolean
    Dim SHop As SHFILEOPSTRUCT
    Dim lngResult As Long
    
    If LenB(sFileName) = 0 Then sFileName = mvarsFileName
    
    With SHop
        .wFunc = FO_MOVE
        .pFrom = sFileName
        .pTo = strDestination
        .fFlags = FOF_ALLOWUNDO
    End With
    
    lngResult = SHFileOperation(SHop)
    If lngResult = 0 Then FileMove = True
End Function
Public Function FileExecute(Optional strFileName As String = "", Optional strParameters As String = "", Optional RunFrom As String = "C:\") As Long
    Dim r As Long
    Dim hWnd As Long
    
    If LenB(strFileName) > 0 Then mvarsFileName = strFileName
    If LenB(mvarsFileName) > 0 Then
        hWnd = GetDesktopWindow()
        r = ShellExecute(hWnd, "open", mvarsFileName, strParameters, RunFrom, SW_SHOWNORMAL)
        If r <= 32 Then
            Err.Raise r, "FileObject:FileExecute", WhatError(r)
        End If
        FileExecute = r
    End If
    
End Function
Public Sub FilePropertiesDialog(Optional strFileName As String)
    Dim hWnd As Long
    hWnd = GetDesktopWindow()
    ShellExecute hWnd, "properties", strFileName, vbNullString, vbNullString, SW_SHOWNORMAL
End Sub


Private Function WhatError(lng As Long) As String
    Select Case lng
        Case SE_ERR_FNF
            WhatError = "File not found"
        Case SE_ERR_PNF
            WhatError = "Path not found"
        Case SE_ERR_ACCESSDENIED
            WhatError = "Access denied"
        Case SE_ERR_OOM
            WhatError = "Out of memory"
        Case SE_ERR_DLLNOTFOUND
            WhatError = "DLL not found"
        Case SE_ERR_SHARE
            WhatError = "A sharing violation occurred"
        Case SE_ERR_ASSOCINCOMPLETE
            WhatError = "Incomplete or invalid file association"
        Case SE_ERR_DDETIMEOUT
            WhatError = "DDE Time out"
        Case SE_ERR_DDEFAIL
            WhatError = "DDE transaction failed"
        Case SE_ERR_DDEBUSY
            WhatError = "DDE busy"
        Case SE_ERR_NOASSOC
            WhatError = "No association for file extension"
        Case ERROR_BAD_FORMAT
            WhatError = "Invalid EXE file or error in EXE image"
        Case Else
            WhatError = "Unknown error"
    End Select
End Function
Public Sub FileExplorDir(Optional strFileName As String)
    Dim hWnd As Long
    hWnd = GetDesktopWindow()
    ShellExecute hWnd, "explore", strFileName, vbNullString, vbNullString, SW_SHOWNORMAL
End Sub
Public Function FileRename(strDestination As String, Optional sFileName As String) As Boolean
    Dim SHop As SHFILEOPSTRUCT
    Dim lngResult As Long
    
    If LenB(sFileName) = 0 Then sFileName = mvarsFileName
    
    With SHop
        .wFunc = FO_RENAME
        .pFrom = sFileName
        .pTo = strDestination
        .fFlags = FOF_ALLOWUNDO
    End With
    
    lngResult = SHFileOperation(SHop)
    If lngResult = 0 Then FileRename = True
End Function
Private Function GetShortName(ByVal sFileName As String) As String
                    
    Dim sBuffer As String
    sBuffer = Space(1024)
    GetShortPathName sFileName, sBuffer, Len(sBuffer)
    GetShortName = NullTrim(sBuffer)
    
End Function
Private Function GetLongName(ByVal sFileName As String) As String
                    
    Dim sBuffer As String
    sBuffer = Space(1024)
    GetExpandedName sFileName, sBuffer
    GetLongName = NullTrim(sBuffer)
    
End Function
Private Function NullTrim(ByVal sString As String) As String
                
    Dim iNullLoc As Integer
    iNullLoc = InStr(sString, Chr(0))
    
    If iNullLoc > 0 Then
        NullTrim = Left(sString, iNullLoc - 1)
    Else
        NullTrim = sString
    End If
    
End Function

Public Function timeFileToDate(ft As FILETIME) As Date
                    
    Dim tSysTime As SYSTEMTIME
    
    FileTimeToSystemTime ft, tSysTime
    timeFileToDate = timeSysToDate(tSysTime)
                    
End Function
Private Function FileInformation(ByVal sFileName As String) As BY_HANDLE_FILE_INFORMATION
                    
    Dim tInfo As BY_HANDLE_FILE_INFORMATION
    Dim tOF As OFSTRUCT
    Dim lHandle As Long
    
    lHandle = OpenFile(sFileName, tOF, 0)
    
    If lHandle > 0 Then
        GetFileInformationByHandle lHandle, tInfo
    End If
    
    FileInformation = tInfo
    
    CloseHandle lHandle
    
End Function
Private Function timeSysToDate(st As SYSTEMTIME) As Date
    
    timeSysToDate = CDate(Format(st.wMonth, "00") & "/" & Format(st.wDay, "00") & "/" & Format(st.wYear, "0000") & " " & Format(st.wHour, "00") & ":" & Format(st.wMinute, "00") & ":" & Format(st.wSecond, "00"))
    
End Function
Private Function fileOpenStructure(ByVal sFileName As String) As OFSTRUCT
    
    Dim tOF As OFSTRUCT
    Dim lHandle As Long
    
    lHandle = OpenFile(sFileName, tOF, 0)
    CloseHandle lHandle
    fileOpenStructure = tOF
    
End Function

Private Function ExtractFileName(ByVal strPath As String) As String
    Dim lngPosition As Long
   
    ' StrReverse is only working in VB6
    If LenB(strPath) > 0 Then
        strPath = StrReverse(strPath)
        lngPosition = InStr(strPath, "\")
        If lngPosition > 0 Then
            strPath = Left(strPath, lngPosition - 1)
        End If
        strPath = StrReverse(strPath)
        ExtractFileName = strPath
    Else
        ExtractFileName = ""
    End If
    
End Function
Private Function IsFileThere(ByRef inFileSpec As String) As Boolean
   Dim i     As Variant
   Dim mFile As String
   
   On Error Resume Next
   
   mFile = GetShortName(inFileSpec)
   
   i = FreeFile
   Open inFileSpec For Input As i
      If Err Then
         IsFileThere = False
      Else
      Close i
      IsFileThere = True
   End If
End Function
Private Function DestroyFile(strFile As String) As Boolean

    Dim fhnd As Long
    Dim fSize As Long
     
    fhnd = CreateFile(strFile, GENERIC_WRITE, FILE_SHARE_WRITE, 0, TRUNCATE_EXISTING, FILE_FLAG_NO_BUFFERING Or FILE_FLAG_WRITE_THROUGH, 0)
   
    If fhnd = -1 Then
        Call CloseHandle(fhnd)
        DestroyFile = False
        Exit Function
    Else
        Call FlushFileBuffers(fhnd) 'Flush any buffers the system used for the file
        Call CloseHandle(fhnd)
        fSize = FileLen(strFile)

            If fSize <> 0 Then
                DestroyFile = False
            Else
                DestroyFile = True
            End If
    End If
 
End Function
Private Sub SecureWipeFile(sPath As String, iPasses As Integer)
    Dim iCurrentPass As Integer
    Dim iFileHandler As Integer
    Dim iNumberOfPasses As Integer
    Dim iProgress1 As Integer
    Dim iProgress2 As Integer
    Dim iProgress3 As Integer
    Dim iProgress4 As Integer
    Dim iProgress5 As Integer
    Dim iProgress6 As Integer
    Dim lBytesWritten As Long
    Dim lChunk As Long
    Dim dRemainingSize As Double
    Dim lFileSize As Long
    Dim sBitPattern1 As String
    Dim sBitPattern2 As String
    Dim sBitPattern3 As String
    Dim sTemporaryBitPattern As String
    On Error Resume Next
    sBitPattern1 = String$(CWR_BUFFER, Chr$(0))
    sBitPattern2 = String$(CWR_BUFFER, Chr$(255))
    lFileSize = FileLen(sPath)
    iNumberOfPasses = Int((lFileSize) / CWR_BUFFER)
    dRemainingSize = lFileSize - (iNumberOfPasses * CWR_BUFFER)
    iFileHandler = CreateFile(sPath, GENERIC_WRITE, 0, 0, OPEN_EXISTING, FILE_ATTRIBUTE_ARCHIVE Or FILE_FLAG_SEQUENTIAL_SCAN Or FILE_FLAG_DELETE_ON_CLOSE, 0)


    If iFileHandler = -1 Then
        MsgBox "File does Not exist, or is already open!", vbCritical, "Error!"
    End If
    
    iProgress1 = 0
    iProgress2 = 0
    iProgress3 = 0


    For iCurrentPass = 1 To iPasses
        DoEvents
        iProgress4 = 0
        iProgress5 = 0
        iProgress6 = 0


        If lFileSize > CWR_BUFFER Then


            For lChunk = 1 To iNumberOfPasses
                DoEvents
                WriteFile iFileHandler, ByVal sBitPattern1, CWR_BUFFER, lBytesWritten, 0
                iProgress4 = iProgress4 + 1
                iProgress1 = iProgress3 + iProgress4
                RaiseEvent Progress(Int(iProgress1 * 100 / (iNumberOfPasses * 3 * iPasses)))
                DoEvents
            Next lChunk
        End If
        sBitPattern3 = String$(dRemainingSize, Chr$(0))
        WriteFile iFileHandler, ByVal sBitPattern3, Len(sBitPattern3), lBytesWritten, 0
        sBitPattern3 = Empty
        SetFilePointer iFileHandler, 0, 0, 0
        Randomize
        dRemainingSize = lFileSize


        If lFileSize > CWR_BUFFER Then
            sBitPattern3 = String$(CWR_BUFFER, Chr$(Int(255 * Rnd) + 1))
            sTemporaryBitPattern = Mid(sBitPattern3, 1, 1)
            For lChunk = 1 To iNumberOfPasses
                DoEvents
                WriteFile iFileHandler, ByVal sBitPattern3, CWR_BUFFER, lBytesWritten, 0
                iProgress5 = iProgress5 + 1
                iProgress2 = iProgress1 + iProgress5
                RaiseEvent Progress(Int(iProgress2 * 100 / (iNumberOfPasses * 3 * iPasses)))
                DoEvents
            Next lChunk
            sBitPattern3 = Empty
            
        End If


        If sTemporaryBitPattern = "" Then
            sTemporaryBitPattern = String$(1, Chr$(Int(255 * Rnd) + 1))
        End If
        sBitPattern3 = String$(dRemainingSize, sTemporaryBitPattern)
        sTemporaryBitPattern = Empty
        WriteFile iFileHandler, ByVal sBitPattern3, Len(sBitPattern3), lBytesWritten, 0
        sBitPattern3 = Empty
        SetFilePointer iFileHandler, 0, 0, 0
        dRemainingSize = lFileSize


        If lFileSize > CWR_BUFFER Then


            For lChunk = 1 To iNumberOfPasses
                DoEvents
                WriteFile iFileHandler, ByVal sBitPattern2, CWR_BUFFER, lBytesWritten, 0
                iProgress6 = iProgress6 + 1
                iProgress3 = iProgress2 + iProgress6
                RaiseEvent Progress(Int(iProgress3 * 100 / (iNumberOfPasses * 3 * iPasses)))
                DoEvents
            Next lChunk
        End If
        sBitPattern3 = String$(dRemainingSize, Chr$(255))
        WriteFile iFileHandler, ByVal sBitPattern3, Len(sBitPattern3), lBytesWritten, 0
        sBitPattern3 = Empty
        SetFilePointer iFileHandler, 0, 0, 0
        If lFileSize <= CWR_BUFFER Then RaiseEvent Progress(Int(iCurrentPass * 100 / iPasses))
        DoEvents
    Next iCurrentPass
    CloseHandle iFileHandler
    RaiseEvent Progress(100)
End Sub

Private Function GetExt(ByVal Name As String) As String
   Dim j As Integer
   
   On Error GoTo LocalError
      
   j = InStrRev(Name, ".")
   
   If j > 0 And j < (Len(Name)) Then
      GetExt = Mid$(Name, j + 1)
      CharLower GetExt
   End If

    Exit Function
LocalError:
     
End Function
Private Function GetTimeCreated(strFile As String) As Date
    Dim FTimeCreated As FILETIME
    Dim FTimeModified As FILETIME
    Dim FTimeAccessed As FILETIME
    Dim FLocalTime As FILETIME
    Dim lngReturn As Long
    Dim lngFHandle As Long
    Dim OpenBuffer As OFSTRUCT
    Dim TSystem As SYSTEMTIME
    
    'Open the file up, and get the handle
    lngFHandle = OpenFile(strFile, OpenBuffer, GENERIC_READ)
    'Use the handle to get the created, open
    '     ed and modified times
    lngReturn = GetFileTime(lngFHandle, FTimeCreated, FTimeAccessed, FTimeModified)
    'Close the file
    CloseHandle lngFHandle
    
    Call FileTimeToLocalFileTime(FTimeCreated, FLocalTime)
    
    'Convert the local file time to system time
    Call FileTimeToSystemTime(FLocalTime, TSystem)
    
    GetTimeCreated = TSystem.wMonth & "/" & TSystem.wDay & "/" & TSystem.wYear & " " & TSystem.wHour & ":" & TSystem.wMinute & ":" & TSystem.wSecond
End Function
Private Function GetTimeModified(strFile As String) As Date
    Dim FTimeCreated As FILETIME
    Dim FTimeModified As FILETIME
    Dim FTimeAccessed As FILETIME
    Dim FLocalTime As FILETIME
    Dim lngReturn As Long
    Dim lngFHandle As Long
    Dim OpenBuffer As OFSTRUCT
    Dim TSystem As SYSTEMTIME
    
    'Open the file up, and get the handle
    lngFHandle = OpenFile(strFile, OpenBuffer, GENERIC_READ)
    
    'Use the handle to get the created, opened and modified times
    lngReturn = GetFileTime(lngFHandle, FTimeCreated, FTimeAccessed, FTimeModified)
    
    'Close the file
    CloseHandle lngFHandle
    
    Call FileTimeToLocalFileTime(FTimeModified, FLocalTime)
    
    'Convert the local file time to system time
    Call FileTimeToSystemTime(FLocalTime, TSystem)
    
    GetTimeModified = TSystem.wMonth & "/" & TSystem.wDay & "/" & TSystem.wYear & " " & TSystem.wHour & ":" & TSystem.wMinute & ":" & TSystem.wSecond
End Function
Private Function GetTimeAccessed(strFile As String) As Date
    Dim FTimeCreated As FILETIME
    Dim FTimeModified As FILETIME
    Dim FTimeAccessed As FILETIME
    Dim FLocalTime As FILETIME
    Dim lngReturn As Long
    Dim lngFHandle As Long
    Dim OpenBuffer As OFSTRUCT
    Dim TSystem As SYSTEMTIME
    
    'Open the file up, and get the handle
    lngFHandle = OpenFile(strFile, OpenBuffer, GENERIC_READ)
    'Use the handle to get the created, open
    '     ed and modified times
    lngReturn = GetFileTime(lngFHandle, FTimeCreated, FTimeAccessed, FTimeModified)
    'Close the file
    CloseHandle lngFHandle
    
    Call FileTimeToLocalFileTime(FTimeAccessed, FLocalTime)
    
    'Convert the local file time to system time
    Call FileTimeToSystemTime(FLocalTime, TSystem)
    
    GetTimeAccessed = TSystem.wMonth & "/" & TSystem.wDay & "/" & TSystem.wYear & " " & TSystem.wHour & ":" & TSystem.wMinute & ":" & TSystem.wSecond
End Function
'Functions for Alder CRC 32

Private Function CRCFile(Optional Method As CRCAlgorithms = enmCRC32) As Long
    Dim bArrayFile() As Byte
    Dim lngCRC32 As Long

    Dim lngChunkSize As Long
    Dim lngSize As Long

    lngSize = FileLen(mvarsFileName)
    lngChunkSize = CHUNK_SIZE
    
    
    
    If lngSize <> 0 Then

        ' Read byte array from file
        Open mvarsFileName For Binary Access Read As #1

        Do While Seek(1) < lngSize

            If (lngSize - Seek(1)) > lngChunkSize Then
                ' Process data in chunks. Chunky!
                Do While Seek(1) < (lngSize - lngChunkSize)
                    ReDim bArrayFile(lngChunkSize - 1)
                    Get #1, , bArrayFile()
                    If Method = enmCRC32 Then
                        lngCRC32 = CRCFletcher32(lngCRC32, bArrayFile, UBound(bArrayFile))
                    Else
                        lngCRC32 = CRCAdler(lngCRC32, bArrayFile, UBound(bArrayFile))
                    End If
                    
                Loop
            Else
                ' Blast it at them
                ReDim bArrayFile(lngSize - Seek(1))
                Get #1, , bArrayFile()
                
                If Method = enmCRC32 Then
                    lngCRC32 = CRCFletcher32(lngCRC32, bArrayFile, UBound(bArrayFile))
                Else
                    lngCRC32 = CRCAdler(lngCRC32, bArrayFile, UBound(bArrayFile))
                End If
            End If

        Loop

        Close #1

        ' Everyone expects to view checksums in Hex strings.  Add buffer zeros if
        ' needed by smaller values.
        CRCFile = lngCRC32
    Else
        ' File of zero bytes has a CRC of 0
        CRCFile = lngCRC32
    End If
End Function
'METHODS FOR ALDER
Private Function CRCAdler(ByVal lngAdler32 As Long, ByRef bArrayIn() As Byte, _
    ByVal dblLength As Double) As Long
    
    Dim intPos As Integer
    Dim lngPosInArray As Long
    Dim lngLengthRemaining As Long
    Dim dblLow As Double
    Dim dblHigh As Double
    
    If lngAdler32 <> 0 Then
        dblLow = lngAdler32 And 65535
        dblHigh = RShiftNoRound(lngAdler32, 16) And 65535
    End If
    
    ' Array could be empty
    If UBound(bArrayIn) < LBound(bArrayIn) Then
        CRCAdler = 1
    Else
        ' Get the initial length of the data
        lngLengthRemaining = dblLength
        
        Do While (lngLengthRemaining > 0)
            ' Are we at the end of the file?
            If lngLengthRemaining < NMAX Then
                ' Process the remaining data
                intPos = lngLengthRemaining
                lngLengthRemaining = 0
            Else
                ' No, process a chunk of data the size of NMAX
                intPos = NMAX
                ' We start counting at zero so add one to the NMAX
                lngLengthRemaining = lngLengthRemaining - (NMAX + 1)
            End If
            
            Do
                dblLow = dblLow + bArrayIn(lngPosInArray)
                dblHigh = dblHigh + dblLow
                
                lngPosInArray = lngPosInArray + 1
                intPos = intPos - 1
            Loop While intPos >= 0
            
            ' Use the overflow safe modulus function instead of the
            ' operator.
            dblLow = Modulus(dblLow, BASE)
            dblHigh = Modulus(dblHigh, BASE)
            
        Loop
        
        ' Done, return the result (rightmost 4 bytes).
        CRCAdler = LShift4Byte(dblHigh, 16) Or dblLow
    End If
End Function

' ----------------------------
' Support Routines
' ----------------------------

' We can not use the Mod operator because floating point number could
' be greater than the maximum value of a Long (2,147,483,647).  The
' Mod operator in VB converts numerator and denominator to Longs just to
' mess with your head. ;)
Private Function Modulus(ByVal dblValue As Double, ByVal dblModValue As Double) As Double
    Modulus = dblValue - (dblModValue * Fix(dblValue / dblModValue))
End Function

' Thank you Redeye.  It seems VB does not automaticly trunicate leftmost bits when
' trying to type cast a Double into a Long.  Large numbers give overflow. Not good.
Private Function LShift4Byte(ByVal pnValue As Double, ByVal pnShift As Integer) As Long
    ' Mask For fixing msb
    Dim lngMask As Long
    
    ' Determine whether the sign bit should be set in the bit-shifted result
    If pnValue And (2 ^ (31 - pnShift)) Then
        lngMask = &H80000000
    End If
    
    ' Unset leftmost bits that will be discarded and do the bit shift.
    LShift4Byte = ((pnValue And ((2 ^ (31 - pnShift)) - 1)) * (2 ^ pnShift)) Or lngMask
End Function

Private Function RShiftNoRound(ByVal pnValue As Double, ByVal pnShift As Integer) As Double
    ' Equivilant to C's Bitwise >> operator
    RShiftNoRound = Int(pnValue / (2 ^ pnShift))
End Function
Public Function FileOpen(Optional FileName As String, Optional ofAccessRights As AccessRights = FileRead, _
                        Optional ofShareMode As ShareMode = FileShareRead, Optional ofCreationDisposition As CreationDistribution = FileOpenExisting, _
                        Optional ofFlagsAndAttributes As FlagsAndAttributes = FileAttributeNormal) As Long
                        
    If LenB(FileName) > 0 Then mvarsFileName = FileName
    If LenB(mvarsFileName) > 0 Then
        mvarFileHandle = CreateFile(mvarsFileName, ofAccessRights, ofShareMode, ByVal 0&, ofCreationDisposition, ofFlagsAndAttributes, 0)
        If mvarFileHandle = INVALID_HANDLE_VALUE Then
            FileOpen = 0
        Else
            FileOpen = mvarFileHandle
            mvarFileSize = GetFileSize(mvarFileHandle, 0)
            If mvarFileSize > 0 Then mvarOpened = True
        End If
    Else
    End If
End Function
Public Sub FileClose(Optional FileHandle As Long = 0)
    Dim lngToClose As Long
    
    If FileHandle > 0 Then
        lngToClose = FileHandle
    Else
        lngToClose = mvarFileHandle
    End If
    
    If lngToClose > 0 Then
        FlushFileBuffers lngToClose
        CloseHandle lngToClose
        If lngToClose = mvarFileHandle Then
            mvarOpened = False
            mvarsFileName = vbNullChar
        End If
    End If
End Sub
Public Function SetFilePosition(Optional FileHandle As Long, Optional StartPos As Long) As Boolean
    If FileHandle > 0 Then mvarFileHandle = FileHandle
    If mvarFileHandle > 0 Then
        If (SetFilePointer(mvarFileHandle, StartPos, ByVal 0&, FILE_BEGIN) = &HFFFFFFFF) Then         'exit sub if the pointer did not set correctly
            CloseHandle mvarFileHandle
            mvarFileHandle = 0
            mvarsFileName = vbNullChar
            SetFilePosition = False
        Else
            SetFilePosition = True
        End If
    End If
End Function

Public Function GetDataFromFile(Optional FileName As String, Optional StartPos As Long = 0, Optional ByRef Data As Variant) As Long
    
    Dim nSize As Long, lBytesRead As Long
    
    
    If LenB(FileName) > 0 Then mvarsFileName = FileName
    If LenB(mvarsFileName) > 0 Then
        If Not IsEmpty(Data) Then
            If FileOpen(mvarsFileName, GENERIC_READ, FILE_SHARE_READ Or FILE_SHARE_WRITE, , OPEN_EXISTING) > 0 Then
                If SetFilePosition(, StartPos) Then        'exit sub if the pointer did not set correctly
                    GetDataFromFile = GetData(, Data)
                    mvarsFileName = ""
                    FileClose
                End If
            End If
        End If
    End If
End Function
Public Function GetData(Optional FileHandle As Long, Optional ByRef Data As Variant) As Long
    Dim nSize As Long, lBytesRead As Long
    If FileHandle > 0 Then mvarFileHandle = FileHandle
    If mvarFileHandle > 0 Then
        If Not IsEmpty(Data) Then
            Select Case VarType(Data)
                Case vbArray + vbByte
                    Dim bBytes() As Byte
                    
                    If ReadFile(mvarFileHandle, Data, UBound(Data), lBytesRead, ByVal 0&) > 0 Then
                        Data = bBytes
                        GetData = lBytesRead
                    End If
                    
                    
                                            
                Case vbBoolean
                    Dim blnData As Boolean
                    
                    If ReadFile(mvarFileHandle, blnData, LenB(blnData), lBytesRead, ByVal 0&) > 0 Then
                        Data = blnData
                        GetData = lBytesRead
                    End If
                    
                Case vbByte
                    Dim bytData As Byte
                    If ReadFile(mvarFileHandle, bytData, 1, lBytesRead, ByVal 0&) > 0 Then
                        Data = bytData
                        GetData = lBytesRead
                    End If
        
                Case vbCurrency
                    Dim curData As Currency
                    
                    If ReadFile(mvarFileHandle, curData, LenB(curData), lBytesRead, ByVal 0&) > 0 Then
                        Data = curData
                        GetData = lBytesRead
                    End If
                    
                Case vbDate
                    Dim datData As Date
                    
                    If ReadFile(mvarFileHandle, datData, LenB(datData), lBytesRead, ByVal 0&) > 0 Then
                        Data = datData
                        GetData = lBytesRead
                    End If
                            
                Case vbDouble
                    Dim dblData As Double
                    
                    If ReadFile(mvarFileHandle, dblData, LenB(dblData), lBytesRead, ByVal 0&) > 0 Then
                        Data = dblData
                        GetData = lBytesRead
                    End If
                    
                Case vbInteger
                    Dim intData As Integer
                    If ReadFile(mvarFileHandle, intData, LenB(intData), lBytesRead, ByVal 0&) > 0 Then   'Read from the file
                        Data = intData
                        GetData = lBytesRead
                    End If
                                   
                Case vbLong
                    Dim lngData As Long
                    If ReadFile(mvarFileHandle, lngData, 4, lBytesRead, ByVal 0&) > 0 Then   'Read from the file
                        Data = lngData
                        GetData = lBytesRead
                    End If
                            
                Case vbSingle
                    Dim sngData As Single
                    
                    If ReadFile(mvarFileHandle, sngData, LenB(sngData), lBytesRead, ByVal 0&) > 0 Then    'Read from the file
                        Data = sngData
                        GetData = lBytesRead
                    End If
                    
                Case vbString
                    Dim strData As String
                    If ReadFile(mvarFileHandle, strData, Len(strData), lBytesRead, ByVal 0&) > 0 Then 'Read from the file
                        Data = strData
                        GetData = lBytesRead
                    End If
                            
                Case Else
                    '
                    'Unknown data type
                    '
            End Select
        End If
    Else
    
    End If
End Function

Public Function PutDataInFile(Optional FileName As String, Optional ByVal StartPos As Long, Optional ByRef Data As Variant) As Long

    Dim lngFileLen As Long
    Dim lngTemp As Long
    Dim TempArray As Variant
        
    If LenB(FileName) > 0 Then mvarsFileName = FileName
    If LenB(mvarsFileName) > 0 Then
    
        If Len(Data) > 0 Then
            'get the size of the file (if it exists)
            lngFileLen = 0
            
            If (Len(Dir(FileName))) Then
                lngFileLen = FileLen(FileName)
            End If
            
            If StartPos = 0 Then StartPos = lngFileLen
            
            If FileOpen(FileName, GENERIC_WRITE, FileShareRead, IIf(Len(Dir(FileName)), OPEN_EXISTING, CREATE_NEW), FILE_ATTRIBUTE_NORMAL) > 0 Then
                If SetFilePosition(, StartPos) Then
                    PutDataInFile = PutData(, Data)
                End If
            End If
            
            FileClose
            
        End If
    End If
            
End Function
Public Function PutData(Optional FileHandle As Long, Optional ByRef Data As Variant) As Long
    Dim lngFileLen As Long
    Dim lngTemp As Long
    Dim TempArray As Variant
    
    If FileHandle > 0 Then mvarFileHandle = FileHandle
    If mvarFileHandle > 0 Then
        If Len(Data) > 0 Then
                        
            
            
            Select Case VarType(Data)
                Case vbArray + vbByte
                    
                    TempArray() = Data
                    If (WriteFile(mvarFileHandle, TempArray, UBound(TempArray) + 1, lngTemp, ByVal 0&) <> 0) Then        'the data was written correctly
                        PutData = lngTemp
                    End If
                    
                Case vbBoolean
                    Dim blnData As Boolean
                    blnData = CBool(Data)
                    
                    If (WriteFile(mvarFileHandle, blnData, LenB(blnData), lngTemp, ByVal 0&) <> 0) Then        'the data was written correctly
                        PutData = lngTemp
                    End If
                    
                Case vbByte
                    Dim bytData As Byte
                    bytData = CByte(Data)
                    
                    If (WriteFile(mvarFileHandle, bytData, 1, lngTemp, ByVal 0&) <> 0) Then        'the data was written correctly
                        PutData = lngTemp
                    End If
                            
                Case vbCurrency
                    Dim curData As Currency
                    curData = CCur(Data)
                    
                    If (WriteFile(mvarFileHandle, curData, LenB(curData), lngTemp, ByVal 0&) <> 0) Then        'the data was written correctly
                        PutData = lngTemp
                    End If
                    
                Case vbDate
                    Dim datData As Date
                    datData = CDate(Data)
                    
                    If (WriteFile(mvarFileHandle, datData, LenB(datData), lngTemp, ByVal 0&) <> 0) Then        'the data was written correctly
                        PutData = lngTemp
                    End If
                    
                Case vbDouble
                    Dim dblData As Double
                    dblData = CDbl(Data)
                    
                    If (WriteFile(mvarFileHandle, dblData, LenB(dblData), lngTemp, ByVal 0&) <> 0) Then            'the data was written correctly
                        PutData = lngTemp
                    End If
                Case vbInteger
                    Dim intData As Integer
                    intData = CInt(Data)
                    
                    If (WriteFile(mvarFileHandle, intData, LenB(intData), lngTemp, ByVal 0&) <> 0) Then        'the data was written correctly
                        PutData = lngTemp
                    End If
    
                Case vbLong
                    Dim lngData As Long
                    lngData = CLng(Data)
                    If (WriteFile(mvarFileHandle, lngData, LenB(lngData), lngTemp, ByVal 0&) <> 0) Then        'the data was written correctly
                        PutData = lngTemp
                    End If
                            
                Case vbSingle
                    Dim sngData As Single
                    sngData = CSng(Data)
                    
                    If (WriteFile(mvarFileHandle, sngData, LenB(sngData), lngTemp, ByVal 0&) <> 0) Then        'the data was written correctly
                        PutData = lngTemp
                    End If
    
                Case vbString
                    Dim strData As String
                    strData = CStr(Data)
                    If (WriteFile(mvarFileHandle, strData, LenB(strData), lngTemp, ByVal 0&) <> 0) Then        'the data was written correctly
                        PutData = lngTemp
                    End If
                    
                Case Else
                    '
                    'Unknown data type
                    '
            End Select
            
        End If
    End If
End Function

'METHODS FOR FLETCHER

' ----------------------------
' Public Methods
' ----------------------------

Private Function CRCFletcher32(ByVal lngCRC32 As Long, ByRef bArrayIn() As Byte, ByVal dblLength As Double) As Long
    Dim dblCurPos As Double
    Dim lngTemp As Long
    
    ' Make sure the table is ready, just in case.
    'If Not bolTableReady Then bolTableReady = BuildTable
    
    ' Array could be empty
    If UBound(bArrayIn) < LBound(bArrayIn) Then
        ' Files with no data have a CRC of 0
        CRCFletcher32 = 0
    Else
        lngTemp = lngCRC32 Xor &HFFFFFFFF
        
        For dblCurPos = 0 To dblLength
            lngTemp = (((lngTemp And &HFFFFFF00) \ &H100) And &HFFFFFF) Xor (lngArrayCRCByte((lngTemp And &HFF) Xor bArrayIn(dblCurPos)))
        Next dblCurPos
        
        ' Return the value
        CRCFletcher32 = lngTemp Xor &HFFFFFFFF
    End If
End Function

' ----------------------------
' Support Routines
' ----------------------------

' Create a lookup table of the CRC's of all possible eight
' bit values so we can do a byte-wise 32-bit CRC calculation.
' This will speed things up because we work at the byte
' level, else, we have to go bit by bit. Not good.
Private Function BuildTable() As Boolean
    Dim intBytePos As Integer
    Dim bBitPos As Byte
    
    ' Definition of polynomial for this type of CRC
    Const lngLimit = &HEDB88320
    
    Dim lngCRC As Long
    
    For intBytePos = 0 To 255 Step 1
        lngCRC = intBytePos
        
        For bBitPos = 0 To 7 Step 1
            If lngCRC And 1 Then
              lngCRC = (((lngCRC And &HFFFFFFFE) \ 2) And &H7FFFFFFF) Xor lngLimit
            Else
              lngCRC = ((lngCRC And &HFFFFFFFE) \ 2) And &H7FFFFFFF
            End If
        Next bBitPos
        
        ' Add the checksum of the current byte to the lookup table
        lngArrayCRCByte(intBytePos) = lngCRC
    Next intBytePos
    
    ' Table is ready.  Error handling routine would return false.
    BuildTable = True
End Function


Private Sub Class_Initialize()
    BuildTable
End Sub
