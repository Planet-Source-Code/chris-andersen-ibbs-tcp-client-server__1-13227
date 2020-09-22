Attribute VB_Name = "Declares"
'---------------------------------------------------
'System tray Icon control API
Declare Function Shell_NotifyIcon Lib "shell32.dll" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long

Public Const NIM_ADD = &H0
Public Const NIM_DELETE = &H2
Public Const NIM_MODIFY = &H1

Public Const NIF_ICON = &H2
Public Const NIF_MESSAGE = &H1
Public Const NIF_TIP = &H4

Public Const WM_LBUTTONDBLCLK = &H203
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202
Public Const WM_MOUSEMOVE = &H200
Public Const WM_RBUTTONDBLCLK = &H206
Public Const WM_RBUTTONDOWN = &H204
Public Const WM_RBUTTONUP = &H205
Public Const MAX_PATH = 260

Type NOTIFYICONDATA
        cbSize As Long
        hwnd As Long
        uID As Long
        uFlags As Long
        uCallbackMessage As Long
        hIcon As Long
        szTip As String * 64
End Type

Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

Type WIN32_FIND_DATA
    dwFileAttributes As Long
    ftCreationTime As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime As FILETIME
    nFileSizeHigh As Long
    nFileSizeLow As Long
    dwReserved0 As Long
    dwReserved1 As Long
    cFileName As String * MAX_PATH
    cAlternate As String * 14
End Type

Public Type TFindInfo
    arrPos() As Long
    lngCount As Long
End Type

Public IconData As NOTIFYICONDATA

Public Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" ( _
    ByVal lpFileName As String, _
    lpFindFileData As WIN32_FIND_DATA _
    ) As Long

Public Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" ( _
    ByVal hFindFile As Long, _
    lpFindFileData As WIN32_FIND_DATA _
    ) As Long

'Public Declarations for iBBS Server

Public blnLoginGood As Boolean
Public colUsers As New Collection
Public User As New UserProperty
Public strServerName As String
Public strAdminName As String
Public strMaxClients As String
Public strDays As String
Public blnServerStartState As Boolean
Public blnTStamp As Boolean

'ini file reading and writing API functions
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias _
"GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal _
lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As _
String, ByVal nSize As Long, ByVal lpFileName As String) As Long

Public Declare Function WritePrivateProfileString Lib "kernel32" Alias _
"WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal _
lpKeyName As String, ByVal lpString As String, ByVal lpFileName As String) As _
Long

'Pause API function
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)





