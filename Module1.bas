Attribute VB_Name = "Module1"
' Some common API's
Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Const HWND_TOPMOST = -1
Public Const SWP_SHOWWINDOW = &H40
Public Const SWP_NOSIZE = &H1
Public Const SWP_NOMOVE = &H2

Public masmdir As String

Public Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal lBuffer As String) As Long

Private Const GCT_INVALID = &H0
Private Const GCT_LFNCHAR = &H1
Private Const GCT_SEPARATOR = &H8
Private Const GCT_SHORTCHAR = &H2
Private Const GCT_WILD = &H4
Public Declare Function PathFileExists Lib "shlwapi.dll" Alias "PathFileExistsA" (ByVal pszPath As String) As Long



Public Function GetShortPath(LongPath As String) As String
    'E-Mail: KPDTeam@Allapi.net
    Dim Buffer As String, ret As Long
    'create a buffer
    Buffer = Space(255)
        
        'KPD-Team 1999
    'URL: http://www.allapi.net/
    'E-Mail: KPDTeam@Allapi.net
    Dim lngRes As Long, strPath As String
    'Create a buffer
    strPath = String$(165, 0)
    'retrieve the short pathname
    lngRes = GetShortPathName(LongPath, strPath, Buffer)
    'remove all unnecessary chr$(0)'s
    
    GetShortPath = Left$(strPath, lngRes)

End Function
