Attribute VB_Name = "projRTagModule"
Option Explicit
Option Base 1

'API declarations
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Public Const WM_CLOSE = &H10

Public g_TimeSendCMS As Integer       'time to send alarm to CMS
Public g_TimeSendCOMM As Integer      'time to send tag to Communicator
Public g_CMSIP As String              'CMSIPAddr 10.15.15.43
Public g_COMMIPAddr As String         'COMMIPAddr 10.15.15.63
Public g_COMM_OPCLocalPortNo As Long
Public g_OPC_COMMLocalPortNo As Long

Public g_Debug As Boolean             'true to show more information

'CMS Server Alarm type
Type CMSMsg
  SiteIni As String * 3               'site initial, AAA
  SubSys As String * 3                'site subsystem, AAA
  AlmStatus As String * 1             'alarm status, U up, D down, ? unknown
  NetworkPath As String               'network path
  ServerName As String                'Server name
  TagName As String                   'tagname
End Type

'OPC tag define
Public Const MaxNoOfTag = 100         'max no of tag to support
Public g_TagToRead(MaxNoOfTag) As CMSMsg
Public g_NoOfTag                      'No of tag read from Excel file

' The following objects represent various OPC objects.
' The first is crucial to the process, the last two are
' mere conveniences.
Public Const MaxOPCServer = 10                      'max number of OPC to support
Public gNoOPCServer As Integer                      'no of OPC Server read from EXCEL File
Public gMyServer(MaxOPCServer) As New OPCServer     'OPC Server object, used to connect
Public gMygroup(MaxOPCServer) As OPCGroup           'OPC Group we will create
Public gMyItem(MaxOPCServer) As OPCItem             'OPC Item we will create
Public gServerName(MaxOPCServer, 2) As String       '1 server name, 2 connect or not

'main form
Public fMainForm As frmMain

Public Sub ReadFromFile(ByVal fP As Form)
'read .ini file
Dim TextInfo$   'holds text from INI file
Dim res         'holds results
Dim i, j As Integer
Dim iniFName$
Dim s As String
  
  'ini file name
  iniFName$ = App.Path & "\projRTag.ini"

  'change cursor
  Screen.MousePointer = vbHourglass
  
  DoEvents
    
  'read CMS IP Address
  TextInfo$ = Space(80)
  res = GetPrivateProfileString("CMSIP", "CMSIPAddr", "", TextInfo$, 100, iniFName)
  g_CMSIP = Trim(Left$(TextInfo$, res))
  fP.DispMsg "Read initial file: CMS IP Address is " & g_CMSIP
  
  'read OPC local port no, default 55557
  TextInfo$ = Space(80)
  res = GetPrivateProfileString("OPCIP", "OPC_COMMLocalPort", "", TextInfo$, 100, iniFName)
  g_OPC_COMMLocalPortNo = Val(Trim(Left$(TextInfo$, res)))
  fP.DispMsg "Read initial file: local port number is " & Trim(Str(g_OPC_COMMLocalPortNo))
  
  
  'read COMM IP Address
  TextInfo$ = Space(80)
  res = GetPrivateProfileString("COMMIP", "COMMIPAddr", "", TextInfo$, 100, iniFName)
  g_COMMIPAddr = Trim(Left$(TextInfo$, res))
  fP.DispMsg "Read initial file: Communicator IP Address is " & g_COMMIPAddr
  
  'read COMM Local Port
  TextInfo$ = Space(80)
  res = GetPrivateProfileString("COMMIP", "COMM_OPCLocalPort", "", TextInfo$, 100, iniFName)
  g_COMM_OPCLocalPortNo = Val(Left$(TextInfo$, res))
  fP.DispMsg "Read initial file: Communicator local port number is " & Trim(Str(g_COMM_OPCLocalPortNo))
  
  'read OPC Local Port
  TextInfo$ = Space(80)
  res = GetPrivateProfileString("OPCIP", "OPC_COMMLocalPort", "", TextInfo$, 100, iniFName)
  g_OPC_COMMLocalPortNo = Val(Left$(TextInfo$, res))
  fP.DispMsg "Read initial file: Communicator Remote port number is " & Trim(Str(g_OPC_COMMLocalPortNo))
  
  'read time to send alarm to Communicator
  TextInfo$ = Space(80)
  res = GetPrivateProfileString("TIMESEND", "OPC_COMMTime", "", TextInfo$, 100, iniFName)
  g_TimeSendCOMM = Val(Left$(TextInfo$, res))
  fP.DispMsg "Read initial file: time to send tags to Communicator is " & Trim(Str(g_TimeSendCOMM)) & "sec"
  
  
  'change cursor
  Screen.MousePointer = vbDefault
  
End Sub


Function InitOPConnection(ByRef connectErr As String, ByVal mySNo As Integer, _
  ByVal ServerName As String) As Boolean
''mySNo  which server
''connectErr  Error message return
''
Dim result As Boolean
    
  On Error GoTo InitError
  
  connectErr = "nil"
  result = False
    
  '' Connect to the OPC Server
  Dim svrlist
  svrlist = gMyServer(mySNo).GetOPCServers
  
  ' check if serverName in Registered Server name
  Dim findServer As Boolean
  Dim i As Integer
  findServer = False
  For i = LBound(svrlist) To UBound(svrlist)
    If svrlist(i) = ServerName Then
      findServer = True
      Exit For
    End If
  Next
    
  'if server found, connect to the server
  If findServer Then
    gMyServer(mySNo).Connect ServerName
    
    '' Check for success: if the server isn't running, abort:
    If gMyServer(mySNo).ServerState <> OPCRunning Then
      connectErr = "Unable to connect to server"
      Exit Function
    End If
    
    '' Add a new group: if it fails, abort:
    'Set myGroup = myServer.OPCGroups.Add("MyNewGroup")
    Set gMygroup(mySNo) = gMyServer(mySNo).OPCGroups.Add
    If TypeName(gMygroup(mySNo)) = TypeName(Nothing) Then
      connectErr = "Unable to create OPC group"
      Exit Function
    End If
  Else
    'server not found
    connectErr = "Server not in registered OPC Server list"
    Exit Function
  End If
  
  result = True
  InitOPConnection = result
  Exit Function
  
InitError:
  connectErr = Err.Description
  InitOPConnection = False
  
End Function

Function UnInitOPCConnection(ByVal mySNo As Integer, ByRef myGroup As OPCGroup) As Boolean
''
'' Remove all groups and close connection
''

  ' Remove all OPC groups and disconnect from the server
  gMyServer(mySNo).OPCGroups.RemoveAll
  gMyServer(mySNo).Disconnect
  Set gMyServer(mySNo) = Nothing
  Set gMygroup(mySNo) = Nothing
  
  UnInitOPCConnection = True
  
End Function


Function GetOpcV(ByVal gp As Integer, ByVal TagName As String, _
  ByRef ValueOut As Variant) As Boolean
'
'   Read one OPC value from a specified OPC tag.
'   The value is returned in ValueOut variable
'
Dim ItemID(1 To 10) As String       'Strings containing connection string
Dim ClntHdl(1 To 10) As Long        'User-defined handles for items
Dim SvrHdl() As Long                'OUTPUT: handles defined in item creation
Dim Errors() As Long                'OUTPUT: any error codes generated
Dim ReqDataTypes As Variant         'OUTPUT: requested data types from server
Dim AccessPath(1 To 10) As String   'Access paths of OPC items
Dim Values() As Variant             'OUTPUT: Target for OPC Values

  On Error GoTo FAILURE
  
  ItemID(1) = TagName
  ClntHdl(1) = 1
  AccessPath(1) = ""
  
  gMygroup(gp).OPCItems.AddItems 1, ItemID(), ClntHdl(), _
      SvrHdl(), Errors(), ReqDataTypes, AccessPath()
    
  gMygroup(gp).SyncRead OPCCache, 1, SvrHdl(), Values(), Errors()
    
  gMygroup(gp).OPCItems.Remove 1, SvrHdl(), Errors()
  
  ValueOut = Values(1)
  GetOpcV = True
  
  Exit Function
  
FAILURE:
  GetOpcV = False

End Function


Sub Main()
 
  'change working directory to the directory wher the application was executed.
  ChDrive App.Path
  ChDir App.Path
  
  'load frmMain
  Set fMainForm = New frmMain
  Load fMainForm
  fMainForm.Show
  
End Sub
