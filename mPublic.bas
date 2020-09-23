Attribute VB_Name = "mPublic"
Option Explicit

Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long

Public Function GetDrive(strPath As String) As String
    On Error Resume Next
    Dim strWrkPath As String
    'Set working path variable to avoid send
    '     ing back a changed strPath
    strWrkPath = strPath


    If (InStr(strWrkPath, ":") > 0) Then
        'Drive is there in simplest form
        GetDrive = Left$(strWrkPath, InStr(strWrkPath, ":"))
    Else
        'Check for double back slash usually ind
        '     icating UNC path
        If Mid(strWrkPath, 1, 2) = "\\" Then
            'Remove double backslash
            GetDrive = "OTH_COMPUTER"
        Else
            'Drive\Server not there
            GetDrive = "NO_DRIVE"
        End If
    End If
End Function
Public Function GetComputeName(strPath As String) As String
    On Error Resume Next
    Dim strWrkPath As String
        
    'Set working path variable to avoid send
    '     ing back a changed strPath
    strWrkPath = strPath


    If (InStr(strWrkPath, ":") > 0) Then
        
        GetComputeName = GetCompName
    Else
        'Check for double back slash usually ind
        '     icating UNC path
        If Mid(strWrkPath, 1, 2) = "\\" Then
            'Remove double backslash
            strWrkPath = Mid(strWrkPath, 3, Len(strWrkPath))
            'Set drive to UNC server with \\
            strWrkPath = Left$(strWrkPath, (InStr(strWrkPath, "\") - 1))
            GetComputeName = strWrkPath
        Else
            'Drive\Server not there
            GetComputeName = "NO_DRIVE"
        End If
    End If
End Function
Private Function GetCompName() As String
    Dim buf As String
    Dim str As String
    
    buf = Space$(255)
    GetComputerName buf, Len(buf)
    buf = Trim$(buf)
    GetCompName = Left$(buf, Len(buf) - 1)
    
End Function
