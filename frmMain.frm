VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   Caption         =   "Read Tag - OPC Client"
   ClientHeight    =   8130
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8880
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8130
   ScaleWidth      =   8880
   WindowState     =   1  'Minimized
   Begin VB.PictureBox picButtong 
      Align           =   2  'Align Bottom
      BackColor       =   &H00FF8080&
      Height          =   735
      Left            =   0
      ScaleHeight     =   675
      ScaleWidth      =   8820
      TabIndex        =   25
      Top             =   7020
      Width           =   8880
      Begin VB.CommandButton cmdDebug 
         Caption         =   "Debug On"
         Enabled         =   0   'False
         Height          =   400
         Left            =   3360
         TabIndex        =   30
         Top             =   120
         Width           =   1095
      End
      Begin VB.CommandButton cmdAbout 
         Caption         =   "About"
         Enabled         =   0   'False
         Height          =   400
         Left            =   7800
         TabIndex        =   29
         Top             =   120
         Width           =   855
      End
      Begin VB.CommandButton cmdExit 
         Caption         =   "Exit"
         Enabled         =   0   'False
         Height          =   400
         Left            =   6600
         TabIndex        =   28
         Top             =   120
         Width           =   975
      End
      Begin VB.CommandButton cmdSim 
         Caption         =   "Sim Read"
         Enabled         =   0   'False
         Height          =   400
         Left            =   2040
         TabIndex        =   27
         Top             =   120
         Width           =   1095
      End
      Begin VB.CommandButton cmdOpenSocket 
         Caption         =   "Open Socket"
         Enabled         =   0   'False
         Height          =   400
         Left            =   240
         TabIndex        =   26
         Top             =   120
         Width           =   1215
      End
   End
   Begin MSComctlLib.StatusBar sbStatusBar 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   24
      Top             =   7755
      Width           =   8880
      _ExtentX        =   15663
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   0
            Enabled         =   0   'False
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   6271
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   6271
         EndProperty
      EndProperty
   End
   Begin VB.Timer tmrRetry 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   4440
      Top             =   6480
   End
   Begin VB.Timer tmrChHeartBeat 
      Interval        =   60000
      Left            =   6240
      Top             =   6480
   End
   Begin VB.Timer tmrSimRead 
      Enabled         =   0   'False
      Left            =   6840
      Top             =   6480
   End
   Begin VB.Timer tmrReadTag 
      Enabled         =   0   'False
      Left            =   6600
      Top             =   6480
   End
   Begin VB.Timer tmrSckState 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   4800
      Top             =   6480
   End
   Begin MSWinsockLib.Winsock sckCOMM 
      Left            =   6960
      Top             =   6480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
   End
   Begin VB.Timer tmrInit 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   5400
      Top             =   6480
   End
   Begin VB.Frame fraTagStatus 
      Caption         =   "Tags Status"
      Height          =   1215
      Left            =   120
      TabIndex        =   14
      Top             =   2040
      Width           =   8655
      Begin VB.Label lblNoServer 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5880
         TabIndex        =   23
         Top             =   360
         Width           =   2415
      End
      Begin VB.Label lblTagStatus 
         Alignment       =   1  'Right Justify
         Caption         =   "OPC Server Connected:"
         Height          =   375
         Index           =   3
         Left            =   3480
         TabIndex        =   22
         Top             =   360
         Width           =   2295
      End
      Begin VB.Label lblSimMode 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5880
         TabIndex        =   20
         Top             =   720
         Width           =   2415
      End
      Begin VB.Label lblTagStatus 
         Alignment       =   1  'Right Justify
         Caption         =   "Sim Mode:"
         Height          =   375
         Index           =   2
         Left            =   4560
         TabIndex        =   19
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label lblTimeToUpdate 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1560
         TabIndex        =   18
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label lblNoOfTag 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1560
         TabIndex        =   17
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label lblTagStatus 
         Alignment       =   1  'Right Justify
         Caption         =   "Time to Update:"
         Height          =   375
         Index           =   1
         Left            =   240
         TabIndex        =   16
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label lblTagStatus 
         Alignment       =   1  'Right Justify
         Caption         =   "No of Tags:"
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   15
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Frame fraMsg 
      Caption         =   "Message"
      Height          =   3375
      Left            =   120
      TabIndex        =   11
      Top             =   3360
      Width           =   8655
      Begin VB.Timer tmrDelay 
         Enabled         =   0   'False
         Interval        =   50
         Left            =   2640
         Top             =   2760
      End
      Begin VB.ListBox lstMsg 
         Height          =   2790
         Left            =   120
         TabIndex        =   21
         Top             =   360
         Width           =   8415
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Winsock Information"
      Height          =   1815
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8655
      Begin VB.Label lblLocalPort 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1560
         TabIndex        =   13
         Top             =   1080
         Width           =   1815
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Local Port:"
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "Local Name:"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label lblHostName 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1560
         TabIndex        =   9
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Caption         =   "Local IP:"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label lblLocalIP 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1560
         TabIndex        =   7
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         Caption         =   "State:"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label lblSenderState 
         Caption         =   "sckClosed"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1560
         TabIndex        =   5
         Top             =   1440
         Width           =   1815
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         Caption         =   "Communicator IP:"
         Height          =   255
         Left            =   4200
         TabIndex        =   4
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label lblRemoteIP 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5880
         TabIndex        =   3
         Top             =   720
         Width           =   2415
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         Caption         =   "Communicator  Port:"
         Height          =   255
         Left            =   4200
         TabIndex        =   2
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Label lblRemotePort 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5880
         TabIndex        =   1
         Top             =   1080
         Width           =   2415
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Base 1

Private exclApp As Excel.Application
Private exclBook As Excel.Workbook
Private exclSheet As Excel.Worksheet
Private exclOpen As Boolean

Private simMode As Boolean
Private HeartBeatRx As Boolean

Private delayCn As Integer


Private Function SendTagToCOMM() As Boolean
'send data to COMM
Dim s$
Dim i As Integer

  If Me.sckCOMM.State <> sckConnected Then
    SendTagToCOMM = False
    Exit Function
  End If
  
  s$ = "^STATUS " & Format(Now, "YYYYMMDD hhmmss")
  For i = 1 To g_NoOfTag
    With g_TagToRead(i)
      s$ = s$ + .SiteIni & .SubSys & .AlmStatus & " "
    End With
  Next
  
  s$ = s$ & "^"
  Me.sckCOMM.SendData s$ & vbCrLf
  DispMsg ("Packet send: " & s$)
  
End Function

Private Function WhichServer(ByVal sName As String, ByRef sNo As Integer) As Boolean
'match the server number and return true if found
'Public gServerName(MaxOPCServer, 2) As String       '1 server name, 2 connect or not
Dim i As Integer

  For i = 1 To MaxOPCServer
    If gServerName(i, 1) = sName Then
      If gServerName(i, 2) = 1 Then
        sNo = i
        WhichServer = True
        Exit For
      Else
        sNo = 0
        WhichServer = False
      End If
    End If
  Next
  
  If i = MaxOPCServer Then
  'not found
    sNo = 0
    WhichServer = False
  End If
    
End Function

Private Function SimReadTag() As Boolean
'simulator read tag from RTU
Dim OpcTagName As String
Dim OPCNetworkName As String
Dim OpcServerName As String
Dim OpcServerNo As Integer
Dim OpcValue As Variant     'hold return value
Dim i As Integer


  On Error GoTo SimReadTag_Error
  
  For i = 1 To g_NoOfTag
    OPCNetworkName = g_TagToRead(i).NetworkPath
    OpcServerName = g_TagToRead(i).ServerName
    OpcTagName = g_TagToRead(i).TagName
    
    'generate current value
    OpcValue = Int((2 * Rnd) + 1)
    Select Case OpcValue
    Case 1
      g_TagToRead(i).AlmStatus = "U"
    Case 2
      g_TagToRead(i).AlmStatus = "D"
    Case Else
      g_TagToRead(i).AlmStatus = "?"
    End Select
    'DispMsg ("* " & OpcTagName & " = " & g_TagToRead(i).AlmStatus)
  Next
    
  SimReadTag = True
  Exit Function
  
SimReadTag_Error:
  SimReadTag = False

End Function
Private Function ReadTag() As Boolean
'read the tag from RTU
Dim OpcTagName As String
Dim OPCNetworkName As String
Dim OpcServerName As String
Dim OpcServerNo As Integer
Dim OpcValue As Variant     'hold return value
Dim i As Integer


  On Error GoTo ReadTag_Error
  
  DispMsg ("Start read tags!")
    
    
  OPCNetworkName = g_TagToRead(1).NetworkPath
  OpcServerName = g_TagToRead(1).ServerName
  OpcTagName = g_TagToRead(1).TagName
  If WhichServer(OpcServerName, OpcServerNo) Then
  End If

  For i = 1 To g_NoOfTag
    delayCn = 1
    Me.tmrDelay.Enabled = True
    Do While delayCn = 1
      DoEvents
    Loop
    
    OpcTagName = g_TagToRead(i).TagName

    ' Read current value
    'If WhichServer(OpcServerName, OpcServerNo) Then
      If GetOpcV(OpcServerNo, OpcTagName, OpcValue) Then
        If g_Debug Then
          DispMsg (OpcTagName & " = " & OpcValue)
        End If
        'update tag
        Select Case OpcValue
        Case True
          g_TagToRead(i).AlmStatus = "D"
        Case False
          g_TagToRead(i).AlmStatus = "U"
        Case 1
          g_TagToRead(i).AlmStatus = "D"
        Case 0
          g_TagToRead(i).AlmStatus = "U"
        Case Else
          g_TagToRead(i).AlmStatus = "?"
        End Select
        If g_Debug Then
          DispMsg (OpcTagName & " = " & g_TagToRead(i).AlmStatus)
        End If
      Else
        DispMsg (OpcTagName & "Failed to read OPC value from ")
        g_TagToRead(i).AlmStatus = "?"
      End If
    'Else
    '  DispMsg ("Server name not found or not connected!!")
    '  g_TagToRead(i).AlmStatus = "?"
    'End If
  Next
    
  DispMsg ("Read tags ended!")
    
  ReadTag = True
  Exit Function
  
ReadTag_Error:
  ReadTag = False
  
End Function
Private Sub Open_ExcelFile(ByVal na$)

  Set exclApp = New Excel.Application
  Set exclBook = exclApp.Workbooks.Open(na$)
  Set exclSheet = exclBook.ActiveSheet
  DispMsg ("Opened Excel file " & na$)
  exclOpen = True
End Sub

Private Sub Close_ExcelFile()
  
  exclBook.Close (False)
 
  Set exclSheet = Nothing
  Set exclBook = Nothing
  Set exclSheet = Nothing

  
End Sub

Private Sub Read_ExcelFile_Tag()
'read the excel file and determine tags to be read
'g_NoOfTag read from file
'g_TagToRead(i,1) = network path
'g_TagToRead(i,2) = server name
'g_TagToRead(i,3) = tag name
'
'gNoOPCServer
Dim i As Integer
Dim tServer As String

  gNoOPCServer = 0
  tServer = ""
  
  If exclOpen Then
    g_NoOfTag = 0
    For i = 1 To MaxNoOfTag
    
      With g_TagToRead(i)
        'get network path
        .NetworkPath = exclSheet.Cells(i, 1)
        If Len(Trim(.NetworkPath)) = 0 Then Exit For
      
        'get Server path
        .ServerName = exclSheet.Cells(i, 2)
        If Len(Trim(.ServerName)) = 0 Then Exit For
        
        'get no of OPC Server to connect
        If tServer <> .ServerName Then
          tServer = .ServerName
          gNoOPCServer = gNoOPCServer + 1
          If gNoOPCServer = MaxOPCServer Then
            DispMsg ("Max no of OPC Server reached!!")
            Exit For
          End If
        End If
        
        'get tag name
        .TagName = exclSheet.Cells(i, 3)
        If Len(Trim(.TagName)) = 0 Then Exit For
        
        'get SiteIni
        .SiteIni = exclSheet.Cells(i, 4)
        If Len(Trim(.SiteIni)) = 0 Then Exit For
        
        'get Subsys
        .SubSys = exclSheet.Cells(i, 5)
        If Len(Trim(.SubSys)) = 0 Then Exit For
        
        DispMsg ("Tags " & Str(i) & ": \\" & .NetworkPath & "\" & _
          .ServerName & "\" & .TagName & " " & .SiteIni & " " & .SubSys)
      End With
    Next
    
    If i = MaxNoOfTag Then DispMsg ("Max no of Tag reached!!")
    g_NoOfTag = i - 1
    DispMsg ("No of tag read: " & Str(g_NoOfTag))
    DispMsg ("No of OPC Server read: " & Str(gNoOPCServer))
  End If
End Sub


Public Sub DispMsg(ByVal s$)
'display message using listbox
'security level, g_Security 1: show all, 2: not show
Dim a$
Dim i As Integer
Dim m$

  'clear listbox if item too many to prevent overflow
  If Me.lstMsg.ListCount > 2048 Then
    Me.lstMsg.Clear
  End If
  
  a$ = Format(Now, "YYYYMMDD hh:mm:ss")
  
  i = Len(s$)
  Do Until i < 66
    m$ = Mid$(s$, 1, 66)
    s$ = Mid$(s$, 67)
    Call Me.lstMsg.AddItem(a$ & " -- " & m$)
    i = Len(s$)
  Loop
  If i > 0 Then
    Call Me.lstMsg.AddItem(a$ & " -- " & s$)
  End If
  
  
  
  'Call Me.lstMsg.AddItem(a$ & " -- " & s$)
  
  If Me.lstMsg.ListCount > 4096 Then Me.lstMsg.Clear
  
  'select last index
  If Me.lstMsg.ListIndex < 0 Then
    Me.lstMsg.AddItem ("")
    Me.lstMsg.ListIndex = 0
  Else
    Me.lstMsg.ListIndex = Me.lstMsg.ListCount - 1
  End If

End Sub

Private Sub cmdAbout_Click()
  
  frmAbout.Show vbModal, Me
End Sub

Private Sub cmdDebug_Click()

  If Me.cmdDebug.Caption = "Debug On" Then
    Me.cmdDebug.Caption = "Debug Off"
    g_Debug = True
  Else
    Me.cmdDebug.Caption = "Debug On"
    g_Debug = False
  End If
End Sub

Private Sub cmdExit_Click()
  
  Unload Me

End Sub


Private Sub cmdOpenSocket_Click()
' If the socket state is closed, we need to bind to a local
' port and also to the remote host's IP address and port
Dim Msg   ' Declare variable.
   
  
  Select Case Me.cmdOpenSocket.Caption
  Case "Close Socket"
    'confirm to Close socket
    Msg = "Do you really want to close the UDP Socket?"
    If MsgBox(Msg, vbQuestion + vbYesNo, Me.Caption) = vbNo Then Exit Sub
    
    Me.sckCOMM.Close
    DispMsg ("Close Socket button clicked")
    Me.cmdOpenSocket.Caption = "Open Socket"
    Me.tmrRetry.Enabled = False
    
    If simMode Then
      Me.tmrSimRead.Enabled = False
    Else
      Me.tmrReadTag.Enabled = False
    End If
  Case "Open Socket"
    Me.sckCOMM.Connect
    DispMsg ("Open Socket button clicked")
    Me.cmdOpenSocket.Caption = "Close Socket"
    Me.tmrRetry.Enabled = True
    
    If simMode Then
      Me.tmrSimRead.Enabled = True
    Else
      Me.tmrReadTag.Enabled = True
    End If
  End Select
  
End Sub


Private Sub cmdSim_Click()
  
  If Me.cmdSim.Caption = "Sim Read" Then
    DispMsg ("Enable Tag read simulation button clicked")
    If Me.sckCOMM.State = sckConnected Then
    'start simulation
      Me.tmrReadTag.Enabled = False
      Me.cmdSim.Caption = "Disable Sim"
      Me.tmrSimRead.Enabled = True
      simMode = True
      Me.lblSimMode.Caption = simMode
      Me.lblSimMode.ForeColor = vbRed
    Else
      DispMsg ("Simulation failed, scoket closed")
      simMode = False
    End If
    
  Else
    DispMsg ("Disable Tag read simulation button clicked")
    If Me.sckCOMM.State = sckConnected Then
      Me.cmdSim.Caption = "Sim Read"
      Me.tmrReadTag.Enabled = True
      Me.tmrSimRead.Enabled = False
      simMode = False
      Me.lblSimMode.Caption = simMode
      Me.lblSimMode.ForeColor = vbNormal
    Else
      Me.cmdSim.Caption = "Sim Read"
      Me.tmrReadTag.Enabled = False
      Me.tmrSimRead.Enabled = False
      simMode = False
      Me.lblSimMode.Caption = simMode
      Me.lblSimMode.ForeColor = vbNormal
    End If
  End If
  
  
End Sub

Private Sub Form_Initialize()

  'read the ini file
  ReadFromFile Me
  
  g_Debug = False
  
  Me.lblHostName.Caption = Me.sckCOMM.LocalHostName
  Me.lblLocalIP.Caption = Me.sckCOMM.LocalIP
  'Me.lblRemoteIP.Caption = g_COMMIPAddr
  'Me.lblRemotePort.Caption = g_COMM_OPCLocalPortNo
  'Me.lblLocalPort.Caption = g_OPC_COMMLocalPortNo

  'start init timer
  Me.tmrInit.Enabled = True
  Me.MousePointer = vbHourglass
  
End Sub


Private Sub Form_Load()
Dim i As Integer

  'init CMS signal
  For i = 1 To MaxNoOfTag
    With g_TagToRead(i)
      .NetworkPath = ""
      .ServerName = ""
      .TagName = ""
      .SiteIni = "XXX"
      .SubSys = "XXX"
      .AlmStatus = "?"
    End With
  Next
  
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim Msg   ' Declare variable.
   
  ' Set the message text.
  Msg = "Do you really want to exit the application?"
  
  ' If user clicks the No button, stop QueryUnload.
  If MsgBox(Msg, vbQuestion + vbYesNo, Me.Caption) = vbNo Then Cancel = True

End Sub

Private Sub Form_Resize()
  
  
  On Error GoTo FormSizeError
  If Me.Width < 300 Then GoTo FormSizeError
  
  If Me.WindowState <> vbMinimized Then
    Me.Height = 8535
    Me.fraMsg.Width = Me.ScaleWidth
    Me.lstMsg.Width = Me.fraMsg.Width - 400
    Me.cmdExit.Left = Me.picButtong.Width - Me.cmdExit.Width - Me.cmdExit.Width - 200
    
    Me.cmdAbout.Left = Me.cmdExit.Left + Me.cmdExit.Width + 200
    
  End If
  
  Exit Sub
  
FormSizeError:
  If Me.WindowState <> vbMinimized Then
    If Me.WindowState <> vbMaximized Then
      Me.Width = 9000
      Me.Height = 8535
      Me.cmdExit.Left = 6660
    End If
  End If
  
End Sub


Private Sub Form_Unload(Cancel As Integer)
Dim i As Integer

  Me.MousePointer = vbHourglass
  
  Me.cmdExit.Enabled = False
  Me.cmdOpenSocket.Enabled = False
  Me.cmdSim.Enabled = False
  Me.cmdAbout.Enabled = False
  Me.cmdDebug.Enabled = False
  
  Me.tmrReadTag.Enabled = False
  Me.tmrSimRead.Enabled = False
  
  'init OPC Server connection
  For i = 1 To gNoOPCServer
    If gServerName(i, 2) = "1" Then
      Call UnInitOPCConnection(i, gMygroup(1))
      DispMsg ("Disconnected OPC Server " & gServerName(i, 1))
    End If
  Next
  
End Sub

Private Sub lstMsg_Click()
  
  Me.lstMsg.ToolTipText = Me.lstMsg.List(Me.lstMsg.ListIndex)

End Sub

Private Sub sckCOMM_DataArrival(ByVal bytesTotal As Long)
Dim data As String
    
  On Error Resume Next
  
  ' Allocate a string of sufficient size and get the data,
  ' then add it to the list box.
  data = String(bytesTotal + 2, Chr$(0))
  sckCOMM.GetData data, , bytesTotal
  
  'Me.DispMsg "Data from COMM: " & data
  
  'update the label if heart beat received
  HeartBeatRx = True
  
End Sub

Private Sub tmrChHeartBeat_Timer()
Static i

  If i < 2 Then
    i = i + 1
    Exit Sub
  End If
  
  If HeartBeatRx Then
    HeartBeatRx = False
    Exit Sub
  End If
  
  Me.DispMsg "** No heartbeat received from COMM within 3 min **"
  
End Sub



Private Sub tmrDelay_Timer()

  Me.tmrDelay.Enabled = False
  delayCn = 0

End Sub

Private Sub tmrInit_Timer()
'''
'program init after form loaded 1.5sec
'''
Dim i, j As Integer
Dim sTempOPC As String
Dim sErr As String
Dim sServer(100) As String

  
  'read Excel file for Tag to read
  Open_ExcelFile (App.Path & "\TagRead.xls")
  Read_ExcelFile_Tag
  Me.lblNoOfTag.Caption = g_NoOfTag
  Close_ExcelFile
  
  
  'find OPC Server Name
  j = 1
  gServerName(j, 1) = ""
  For i = 1 To g_NoOfTag
    If j = 1 Then
      gServerName(j, 1) = g_TagToRead(i).ServerName
      gServerName(j, 2) = "0"
      j = j + 1
    Else
      If gServerName(j - 1, 1) <> g_TagToRead(i).ServerName Then
        gServerName(j, 1) = g_TagToRead(i).ServerName
        gServerName(j, 2) = "0"
        j = j + 1
      End If
    End If
  Next
  
  'show no of OPC Server
  Me.lblNoServer.Caption = gNoOPCServer

  'init OPC Server connection
  For i = 1 To gNoOPCServer
    'init new opc server
    If InitOPConnection(sErr, i, gServerName(i, 1)) Then
      gServerName(i, 2) = "1"
      DispMsg ("OPC Server connection: " & gServerName(i, 1) & " connected")
    Else
      DispMsg ("OPC Server connection: " & gServerName(i, 1) & " failed to connect, error msg: " & sErr)
      gServerName(i, 2) = "0"
    End If
  Next
  
  'setup timer to read tag from RTU
  Me.tmrReadTag.Interval = g_TimeSendCOMM * 1000
  Me.lblTimeToUpdate.Caption = Me.tmrReadTag.Interval / 1000 & " sec"
  Me.tmrReadTag.Enabled = True
  
  'setup timer to read tag in simulation mode
  Me.tmrSimRead.Interval = Me.tmrReadTag.Interval
  Me.tmrSimRead.Enabled = False
  simMode = False
  Me.lblSimMode = simMode
  
  'TCP setup
  Me.sckCOMM.Protocol = sckTCPProtocol
  Me.sckCOMM.RemoteHost = g_COMMIPAddr
  Me.sckCOMM.RemotePort = g_COMM_OPCLocalPortNo
  'Me.sckCOMM.LocalPort = g_OPC_COMMLocalPortNo
  
  'command button setup
  Me.cmdOpenSocket.Enabled = True
  Me.cmdOpenSocket.Caption = "Open Socket"
  Me.cmdAbout.Enabled = True
  Me.cmdDebug.Enabled = True
  Me.cmdExit.Enabled = True
  Me.cmdSim.Enabled = True
  
  Me.cmdOpenSocket.ToolTipText = "Open or Close UDP Scoket"
  Me.cmdAbout.ToolTipText = "About this application"
  Me.cmdDebug.ToolTipText = "Program debugging"
  Me.cmdSim.ToolTipText = "Simulate to read from RTU"
  Me.cmdExit.ToolTipText = "Exit this application"
  
  'timer for update scoket status
  Me.tmrSckState.Interval = 1000
  Me.tmrSckState.Enabled = True
  
  Me.tmrInit.Enabled = False
  
  Call cmdOpenSocket_Click
  
  Me.sbStatusBar.Panels.Item(1).Text = "Ready"
  Me.MousePointer = vbNormal
  
End Sub

Private Sub tmrReadTag_Timer()

  If Me.sckCOMM.State = sckConnected Then
    'read tag from RTU
    ReadTag
    
    'send tag to Communicator
    SendTagToCOMM
    
  Else
    DispMsg ("Connection broken, please check!!")
  End If
  
End Sub



Private Sub tmrRetry_Timer()
'retry to connect
  
  If Me.sckCOMM.State <> sckConnected Then
    Me.sckCOMM.Close
    Me.sckCOMM.Connect
    DispMsg ("Retry to connect")
  Else
    Me.tmrRetry.Enabled = False
  End If
  
End Sub

Private Sub tmrSckState_Timer()
' When the timer goes off, update the socket status labels
'
Static lastStatus As Integer

  'time information update
  Me.sbStatusBar.Panels.Item(3).Text = Format(Now, "YYYY/MM/DD HH:MM:SS")

  If lastStatus = Me.sckCOMM.State Then Exit Sub
  
  
  lastStatus = Me.sckCOMM.State
  Select Case Me.sckCOMM.State
  Case 0
    Me.lblSenderState.ForeColor = vbRed
    lblSenderState.Caption = "sckClosed"
    Me.lblRemotePort.Caption = "N/A"
    Me.lblLocalPort.Caption = "N/A"
    Me.lblRemoteIP.Caption = "N/A"
  Case 1
    Me.lblSenderState.ForeColor = vbNormal
    lblSenderState.Caption = "sckOpen"
  Case 2
    Me.lblSenderState.ForeColor = vbRed
    lblSenderState.Caption = "sckListening"
  Case 3
    Me.lblSenderState.ForeColor = vbRed
    lblSenderState.Caption = "sckConnectionPending"
  Case 4
    Me.lblSenderState.ForeColor = vbRed
    lblSenderState.Caption = "sckResolvingHost"
  Case 5
    Me.lblSenderState.ForeColor = vbRed
    lblSenderState.Caption = "sckHostResolved"
  Case 6
    Me.lblSenderState.ForeColor = vbRed
    lblSenderState.Caption = "sckConnecting"
  Case 7
    Me.lblSenderState.ForeColor = vbNormal
    lblSenderState.Caption = "sckConnected"
    Me.lblRemotePort.Caption = Me.sckCOMM.RemotePort
    Me.lblLocalPort.Caption = Me.sckCOMM.LocalPort
    Me.lblRemoteIP.Caption = Me.sckCOMM.RemoteHostIP
  Case 8
    Me.lblSenderState.ForeColor = vbRed
    lblSenderState.Caption = "sckClosing"
    If Me.cmdOpenSocket.Caption = "Close Socket" Then
      Me.tmrRetry.Enabled = True
    End If
  Case 9
    Me.lblSenderState.ForeColor = vbRed
    lblSenderState.Caption = "sckError"
    If Me.cmdOpenSocket.Caption = "Close Socket" Then
      Me.tmrRetry.Enabled = True
    End If
    Me.lblRemotePort.Caption = "N/A"
    Me.lblLocalPort.Caption = "N/A"
    Me.lblRemoteIP.Caption = "N/A"
  Case Else
    Me.lblSenderState.ForeColor = vbRed
    lblSenderState.Caption = "sckUnknown"
  End Select

  
End Sub

Private Sub tmrSimRead_Timer()
  
  'simulation read tag
  SimReadTag
  
  'Send tag to Communicator
  SendTagToCOMM
  
End Sub
