Attribute VB_Name = "Module1"
'A few project wide variables

Public strChannels As String
Public strChannel As String
Public fFile As Long
Public strFileLen As Long
Public strHandle As String
Public strChatHandle As String
Public strChatFormState As String
Public IMForm(100) As New frmIM
Public IMNumber As Long
Public strStatus As String
Public strHost As String
Public strUser As String
Public strPass As String
Public lngMBGroup As Long
Public lngMessage As Long
Public strMBGroup As String

'Pause API function
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

