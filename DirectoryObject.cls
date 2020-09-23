VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DirectoryObject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private Type SHFILEOPSTRUCT
   hWnd        As Long
   wFunc       As Long
   pFrom       As String
   pTo         As String
   fFlags      As Integer
   fAborted    As Boolean
   hNameMaps   As Long
   sProgress   As String
 End Type
  


Private Const FO_MOVE As Long = &H1
Private Const FO_COPY As Long = &H2
Private Const FO_DELETE As Long = &H3
Private Const FO_RENAME As Long = &H4

Private Const FOF_SILENT As Long = &H4
Private Const FOF_RENAMEONCOLLISION As Long = &H8
Private Const FOF_NOCONFIRMATION As Long = &H10
Private Const FOF_SIMPLEPROGRESS As Long = &H100
Private Const FOF_ALLOWUNDO As Long = &H40

'local variable(s) to hold property value(s)
Private mvarExists As Boolean
Private mvarPath As String
Private mvarDrive As String
Private mvarComputerName As String

Private Declare Function CreateDirectory Lib "kernel32.dll" Alias "CreateDirectoryA" (ByVal lpPathName As String, lpSecurityAttributes As SECURITY_ATTRIBUTES) As Long
Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function RemoveDirectory Lib "kernel32.dll" Alias "RemoveDirectoryA" (ByVal lpPathName As String) As Long
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
Private Declare Function SHFileOperation Lib "shell32.dll" Alias "SHFileOperationA" (lpFileOp As SHFILEOPSTRUCT) As Long
Private Const MaxLFNPath = 260
Public Type SECURITY_ATTRIBUTES
  nLength As Long
  lpSecurityDescriptor As Long
  bInheritHandle As Boolean
End Type
Private Type WIN32_FIND_DATA
    dwFileAttributes As Long
    ftCreationTime As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime As FILETIME
    nFileSizeHigh As Long
    nFileSizeLow As Long
    dwReserved0 As Long
    dwReserved1 As Long
    cFileName As String * MaxLFNPath
    cShortFileName As String * 14
End Type

Public Property Get ComputerName() As String
    ComputerName = GetComputeName(mvarPath)
End Property
Public Property Get Drive() As String
    Drive = GetDrive(mvarPath)
End Property
Public Property Let Path(ByVal vData As String)
    mvarPath = vData
End Property
Public Property Get Path() As String
    Path = mvarPath
End Property
Public Property Get Exists() As Boolean
    Exists = CHECKFORDIRECTORY(mvarPath)
End Property
Public Function Create(Optional strDirectoryPath As String) As Boolean
    Dim SECATR As SECURITY_ATTRIBUTES
    
    If LenB(strDirectoryPath) = 0 Then strDirectoryPath = mvarPath
    If LenB(strDirectoryPath) = 0 Then Err.Raise 100, "DirectoryObject", "No path set"
    If CreateDirectory(strDirectoryPath, SECATR) = 0 Then Create = True
    
End Function
Public Function Delete(Optional strDirectoryPath As String) As Boolean
    If LenB(strDirectoryPath) = 0 Then strDirectoryPath = mvarPath
    
    If LenB(strDirectoryPath) = 0 Then Err.Raise 100, "DirectoryObject", "No path set"
    If RemoveDirectory(strDirectoryPath) = 0 Then Delete = True
End Function
Public Function Copy(DestinationPath As String, Optional strDirPath As String, _
                     Optional blnConfirm As Boolean = False, Optional blnSilent As Boolean = True, _
                     Optional blnRenameOnCollision As Boolean, Optional blnSimpleProgress As Boolean = False) As Long
                     
    Dim SHFileOp As SHFILEOPSTRUCT
    Dim Flags As Long
    
    If LenB(strDirPath) > 0 Then mvarPath = strDirPath
    If LenB(mvarPath) = 0 Then
        If blnConfirm Then MsgBox "Invalid Path Name"
        Exit Function
    End If
    
    'these can be multiple
    If blnSilent Then Flags = Flags Or FOF_SILENT
    If blnSimpleProgress Then Flags = Flags Or FOF_SIMPLEPROGRESS
    If Not blnConfirm Then Flags = Flags Or FOF_NOCONFIRMATION
    If blnRenameOnCollision Then Flags = Flags Or FOF_RENAMEONCOLLISION
    
    'terminate the folder string with a pair of nulls
    mvarPath = mvarPath & Chr$(0) & Chr$(0)
      
    'set up the options
    With SHFileOp
        .wFunc = FO_COPY
        .pFrom = mvarPath
        .pTo = DestinationPath
        .fFlags = Flags
    End With
    
    Copy = SHFileOperation(SHFileOp)
End Function
Public Function Move(DestinationPath As String, Optional strDirPath As String, _
                     Optional blnConfirm As Boolean = False, Optional blnSilent As Boolean = True, _
                     Optional blnRenameOnCollision As Boolean, Optional blnSimpleProgress As Boolean = False) As Long
    
    Dim SHFileOp As SHFILEOPSTRUCT
    Dim Flags As Long
    
    If LenB(strDirPath) > 0 Then mvarPath = strDirPath
    If LenB(mvarPath) = 0 Then
        If blnConfirm Then MsgBox "Invalid Path Name"
        Exit Function
    Else
        If LenB(DestinationPath) = 0 Then
            If blnConfirm Then MsgBox "Invalid Destination Path"
            Exit Function
        End If
    End If
    
    'these can be multiple
    If blnSilent Then Flags = Flags Or FOF_SILENT
    If blnSimpleProgress Then Flags = Flags Or FOF_SIMPLEPROGRESS
    If Not blnConfirm Then Flags = Flags Or FOF_NOCONFIRMATION
    If blnRenameOnCollision Then Flags = Flags Or FOF_RENAMEONCOLLISION
    
    'terminate the folder string with a pair of nulls
    mvarPath = mvarPath & Chr$(0) & Chr$(0)
      
    'set up the options
    With SHFileOp
        .wFunc = FO_MOVE
        .pFrom = mvarPath
        .pTo = DestinationPath
        .fFlags = Flags
    End With
    
    Move = SHFileOperation(SHFileOp)
End Function

Private Function CHECKFORDIRECTORY(ByVal DIRECTORYNAME As String) As Boolean
   
    Dim lpFindFileData As WIN32_FIND_DATA
    Dim hFindFirst As Long
    hFindFirst = FindFirstFile(DIRECTORYNAME, lpFindFileData)


    If hFindFirst > 0 Then
        FindClose hFindFirst
        CHECKFORDIRECTORY = True
    Else
        CHECKFORDIRECTORY = False
    End If
End Function

