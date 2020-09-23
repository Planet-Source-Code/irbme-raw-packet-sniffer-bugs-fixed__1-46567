VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Packet Sniffer"
   ClientHeight    =   8370
   ClientLeft      =   150
   ClientTop       =   540
   ClientWidth     =   9645
   Icon            =   "frmTest.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8370
   ScaleWidth      =   9645
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CD 
      Left            =   105
      Top             =   8400
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox picICMP 
      Height          =   6315
      Left            =   105
      ScaleHeight     =   6255
      ScaleWidth      =   9195
      TabIndex        =   19
      Top             =   1785
      Width           =   9255
      Begin VB.CommandButton cmdBrowseICMP 
         Caption         =   "..."
         Height          =   285
         Left            =   6720
         TabIndex        =   25
         Top             =   5880
         Width           =   330
      End
      Begin VB.CheckBox chkICMPLogging 
         Caption         =   "Enable Logging"
         Height          =   225
         Left            =   7245
         TabIndex        =   24
         Top             =   5880
         Width           =   1485
      End
      Begin VB.TextBox txtICMPLog 
         Height          =   285
         Left            =   105
         TabIndex        =   23
         Top             =   5880
         Width           =   6630
      End
      Begin MSComctlLib.TreeView tvICMP 
         Height          =   5685
         Left            =   105
         TabIndex        =   22
         Top             =   105
         Width           =   9045
         _ExtentX        =   15954
         _ExtentY        =   10028
         _Version        =   393217
         Indentation     =   529
         LabelEdit       =   1
         Style           =   7
         ImageList       =   "ImgLst"
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSComctlLib.ImageList ImgLst 
         Left            =   5355
         Top             =   1995
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   3
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmTest.frx":0442
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmTest.frx":075C
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmTest.frx":0A76
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin VB.PictureBox picUDP 
      Height          =   6315
      Left            =   105
      ScaleHeight     =   6255
      ScaleWidth      =   9195
      TabIndex        =   18
      Top             =   1785
      Width           =   9255
      Begin VB.CommandButton cmdBrowseUDP 
         Caption         =   "..."
         Height          =   285
         Left            =   6720
         TabIndex        =   28
         Top             =   5880
         Width           =   330
      End
      Begin VB.CheckBox chkUDPLogging 
         Caption         =   "Enable Logging"
         Height          =   225
         Left            =   7245
         TabIndex        =   27
         Top             =   5880
         Width           =   1485
      End
      Begin VB.TextBox txtUDPLog 
         Height          =   285
         Left            =   105
         TabIndex        =   26
         Top             =   5880
         Width           =   6630
      End
      Begin MSComctlLib.TreeView tvUDP 
         Height          =   5685
         Left            =   105
         TabIndex        =   21
         Top             =   105
         Width           =   9045
         _ExtentX        =   15954
         _ExtentY        =   10028
         _Version        =   393217
         Indentation     =   529
         LabelEdit       =   1
         Style           =   7
         ImageList       =   "ImgLst"
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.PictureBox picTCP 
      Height          =   6315
      Left            =   105
      ScaleHeight     =   6255
      ScaleWidth      =   9195
      TabIndex        =   17
      Top             =   1785
      Width           =   9255
      Begin VB.CommandButton cmdBrowseTCP 
         Caption         =   "..."
         Height          =   285
         Left            =   6720
         TabIndex        =   31
         Top             =   5880
         Width           =   330
      End
      Begin VB.CheckBox chkTCPLogging 
         Caption         =   "Enable Logging"
         Height          =   225
         Left            =   7245
         TabIndex        =   30
         Top             =   5880
         Width           =   1485
      End
      Begin VB.TextBox txtTCPLog 
         Height          =   285
         Left            =   105
         TabIndex        =   29
         Top             =   5880
         Width           =   6630
      End
      Begin MSComctlLib.TreeView tvTCP 
         Height          =   5685
         Left            =   105
         TabIndex        =   20
         Top             =   105
         Width           =   9045
         _ExtentX        =   15954
         _ExtentY        =   10028
         _Version        =   393217
         Indentation     =   353
         LabelEdit       =   1
         Style           =   7
         ImageList       =   "ImgLst"
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.Frame fStats 
      Caption         =   "Statistics"
      Height          =   1275
      Left            =   3150
      TabIndex        =   6
      Top             =   105
      Width           =   6420
      Begin VB.Label lblBytesRecieved 
         Caption         =   "0"
         Height          =   225
         Left            =   4830
         TabIndex        =   16
         Top             =   630
         Width           =   1485
      End
      Begin VB.Label Label5 
         Caption         =   "Bytes Recieved (Excluding Headers):"
         Height          =   225
         Left            =   2100
         TabIndex        =   15
         Top             =   630
         Width           =   2640
      End
      Begin VB.Label lblBytesRecievedPackets 
         Caption         =   "0"
         Height          =   225
         Left            =   4830
         TabIndex        =   14
         Top             =   315
         Width           =   1485
      End
      Begin VB.Label Label4 
         Caption         =   "Bytes Recieved (Entire Packet):"
         Height          =   225
         Left            =   2100
         TabIndex        =   13
         Top             =   315
         Width           =   2640
      End
      Begin VB.Label lblICMPPackets 
         Caption         =   "0"
         Height          =   225
         Left            =   1365
         TabIndex        =   12
         Top             =   945
         Width           =   750
      End
      Begin VB.Label lblTCPPackets 
         Caption         =   "0"
         Height          =   225
         Left            =   1365
         TabIndex        =   11
         Top             =   315
         Width           =   750
      End
      Begin VB.Label lblUDPPackets 
         Caption         =   "0"
         Height          =   225
         Left            =   1365
         TabIndex        =   10
         Top             =   630
         Width           =   750
      End
      Begin VB.Label Label3 
         Caption         =   "ICMP Packets:"
         Height          =   225
         Left            =   105
         TabIndex        =   9
         Top             =   945
         Width           =   1380
      End
      Begin VB.Label Label2 
         Caption         =   "TCP Packets:"
         Height          =   225
         Left            =   105
         TabIndex        =   8
         Top             =   315
         Width           =   1380
      End
      Begin VB.Label Label1 
         Caption         =   "UDP Packets:"
         Height          =   225
         Left            =   105
         TabIndex        =   7
         Top             =   630
         Width           =   1380
      End
   End
   Begin MSComctlLib.TabStrip Tabs 
      Height          =   6840
      Left            =   0
      TabIndex        =   5
      Top             =   1365
      Width           =   9465
      _ExtentX        =   16695
      _ExtentY        =   12065
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   3
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "TCP"
            Key             =   "TCP"
            Object.Tag             =   "TCP"
            Object.ToolTipText     =   "Transmission Control Protocol"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "UDP"
            Key             =   "UDP"
            Object.Tag             =   "UDP"
            Object.ToolTipText     =   "user Datagram Protocol"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "ICMP"
            Key             =   "ICMP"
            Object.Tag             =   "ICMP"
            Object.ToolTipText     =   "Internet Control Message Protocol"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Frame fInterface 
      Caption         =   "Network Interfaces"
      Height          =   1275
      Left            =   105
      TabIndex        =   0
      Top             =   105
      Width           =   2850
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   855
         Left            =   105
         ScaleHeight     =   855
         ScaleWidth      =   2640
         TabIndex        =   1
         Top             =   315
         Width           =   2640
         Begin VB.ComboBox cmbInterface 
            Height          =   315
            Left            =   0
            TabIndex        =   4
            Text            =   "Interface List"
            Top             =   0
            Width           =   2640
         End
         Begin VB.CommandButton cmdStart 
            Caption         =   "Start Logging"
            Height          =   435
            Left            =   0
            TabIndex        =   3
            Top             =   420
            Width           =   1275
         End
         Begin VB.CommandButton cmdStop 
            Caption         =   "Stop Logging"
            Enabled         =   0   'False
            Height          =   435
            Left            =   1365
            TabIndex        =   2
            Top             =   420
            Width           =   1275
         End
      End
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "Main"
      Visible         =   0   'False
      Begin VB.Menu mnuRemNode 
         Caption         =   "Remove Node"
      End
      Begin VB.Menu mnuExpandNode 
         Caption         =   "Expand Node"
      End
      Begin VB.Menu mnuCollapseNode 
         Caption         =   "Collapse Node"
      End
      Begin VB.Menu mnuViewData 
         Caption         =   "View Data"
      End
   End
   Begin VB.Menu mnuPopup2 
      Caption         =   "Main"
      Visible         =   0   'False
      Begin VB.Menu mnuExpandNode2 
         Caption         =   "Expand Node"
      End
      Begin VB.Menu mnuCollapseNode2 
         Caption         =   "Collapse Node"
      End
   End
   Begin VB.Menu mnuPopup3 
      Caption         =   "Main"
      Visible         =   0   'False
      Begin VB.Menu mnuRemoveNode3 
         Caption         =   "Remove Node"
      End
      Begin VB.Menu mnuExpandNode3 
         Caption         =   "Expand Node"
      End
      Begin VB.Menu mnuCollapseNode3 
         Caption         =   "Collapse Node"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Main protocol builder interface
Private ProtocolBuilder         As clsProtocolInterface

'The Drivers which plug into the protocol interface in order to capture packets
Private WithEvents TCPDriver    As clsTCPProtocol
Attribute TCPDriver.VB_VarHelpID = -1
Private WithEvents UDPDriver    As clsUDPProtocol
Attribute UDPDriver.VB_VarHelpID = -1
Private WithEvents ICMPDriver   As clsICMPProtocol
Attribute ICMPDriver.VB_VarHelpID = -1

'The complete number of bytes sent including those that make up the headers
Private BytesRecievedPackets    As Long

'The number of bytes of data sent (i.e. exlcuding the packet headers)
Private BytesRecieved           As Long

'The number of packets recieved for each protocol
Private TCPPackets              As Long
Private UDPPackets              As Long
Private ICMPPackets             As Long

'The file numbers of the log files for each protocol
Private TCPLog                  As Integer
Private UDPLog                  As Integer
Private ICMPLog                 As Integer


Private Sub chkICMPLogging_Click()
    
    If chkICMPLogging.Value = vbChecked Then
        ICMPLog = FreeFile

        Open txtICMPLog.Text For Append As #ICMPLog
        Print #ICMPLog, "Best Viewed in font 'Courier New'"
        Print #ICMPLog, "========== Logging Enabled " & Now & " =========="
    Else
        Print #ICMPLog, "========== Logging Disabled " & Now & " ========="
        Close #ICMPLog
        ICMPLog = 0
    End If
    
End Sub


Private Sub chkTCPLogging_Click()

    If chkTCPLogging.Value = vbChecked Then
        TCPLog = FreeFile
        
        Open txtTCPLog.Text For Append As #TCPLog
        Print #TCPLog, "Best Viewed in font 'Courier New'"
        Print #TCPLog, "========== Logging Enabled " & Now & " =========="
    Else
        Print #TCPLog, "========== Logging Disabled " & Now & " ========="
        Close #TCPLog
        TCPLog = 0
    End If
    
End Sub


Private Sub chkUDPLogging_Click()

    If chkUDPLogging.Value = vbChecked Then
        UDPLog = FreeFile
        
        Open txtUDPLog.Text For Append As #UDPLog
        Print #UDPLog, "Best Viewed in font 'Courier New'"
        Print #UDPLog, "========== Logging Enabled " & Now & " =========="
    Else
        Print #UDPLog, "========== Logging Disabled " & Now & " ========="
        Close #UDPLog
        UDPLog = 0
    End If
    
End Sub


Private Sub cmdBrowseTCP_Click()
  On Error GoTo Hell
    
    CD.CancelError = True
    CD.Filter = "Log Files|*.txt"
    CD.DialogTitle = "Open Log File"
    CD.ShowOpen
    txtTCPLog.Text = CD.FileName
    
Hell:
End Sub


Private Sub cmdBrowseUDP_Click()
  On Error GoTo Hell
    
    CD.CancelError = True
    CD.Filter = "Log Files|*.txt"
    CD.DialogTitle = "Open Log File"
    CD.ShowOpen
    txtUDPLog.Text = CD.FileName
    
Hell:
End Sub


Private Sub cmdBrowseICMP_Click()
  On Error GoTo Hell
    
    CD.CancelError = True
    CD.Filter = "Log Files|*.txt"
    CD.DialogTitle = "Open Log File"
    CD.ShowOpen
    txtICMPLog.Text = CD.FileName
    
Hell:
End Sub



Private Sub cmdStart_Click()
    If ProtocolBuilder.CreateRawSocket(Left$(cmbInterface.Text, InStr(1, cmbInterface, " ")), 7000, Me.hWnd) <> 0 Then
        cmdStart.Enabled = Not cmdStart.Enabled
        cmdStop.Enabled = Not cmdStop.Enabled
    End If
End Sub


Private Sub cmdStop_Click()
    cmdStart.Enabled = Not cmdStart.Enabled
    cmdStop.Enabled = Not cmdStop.Enabled
    ProtocolBuilder.CloseRawSocket
End Sub


Private Sub Form_Load()

  Dim str() As String, i As Integer

    Set ProtocolBuilder = New clsProtocolInterface
    Set TCPDriver = New clsTCPProtocol
    Set UDPDriver = New clsUDPProtocol
    Set ICMPDriver = New clsICMPProtocol

    ProtocolBuilder.AddinProtocol TCPDriver, "TCP", IPPROTO_TCP
    ProtocolBuilder.AddinProtocol UDPDriver, "UDP", IPPROTO_UDP
    ProtocolBuilder.AddinProtocol ICMPDriver, "ICMP", IPPROTO_ICMP

    str = Split(EnumNetworkInterfaces(), ";")
        
    For i = 0 To UBound(str)
        If str(i) <> "127.0.0.1" Then
            cmbInterface.AddItem str(i) & " [" & GetHostNameByAddr(inet_addr(str(i))) & "]"
        End If
    Next
    
    cmbInterface.Text = cmbInterface.List(0)

    picTCP.Visible = True
    picUDP.Visible = False
    picICMP.Visible = False
    
    txtTCPLog.Text = App.Path & "\TCPLog.txt"
    txtUDPLog.Text = App.Path & "\UDPLog.txt"
    txtICMPLog.Text = App.Path & "\ICMPLog.txt"

End Sub


Private Sub Form_Unload(Cancel As Integer)
    ProtocolBuilder.CloseRawSocket
    Set ProtocolBuilder = Nothing
    
    Set ICMPDriver = Nothing
    Set UDPDriver = Nothing
    Set TCPDriver = Nothing
End Sub


Private Sub mnuCollapseNode_Click()
    If picTCP.Visible Then tvTCP.SelectedItem.Expanded = False
    If picUDP.Visible Then tvUDP.SelectedItem.Expanded = False
End Sub

Private Sub mnuCollapseNode2_Click()
    If picTCP.Visible Then tvTCP.SelectedItem.Expanded = False
    If picUDP.Visible Then tvUDP.SelectedItem.Expanded = False
    If picICMP.Visible Then tvICMP.SelectedItem.Expanded = False
End Sub

Private Sub mnuCollapseNode3_Click()
    If picICMP.Visible Then tvICMP.SelectedItem.Expanded = False
End Sub

Private Sub mnuExpandNode_Click()
    If picTCP.Visible Then tvTCP.SelectedItem.Expanded = True
    If picUDP.Visible Then tvUDP.SelectedItem.Expanded = True
End Sub

Private Sub mnuExpandNode2_Click()
    If picTCP.Visible Then tvTCP.SelectedItem.Expanded = True
    If picUDP.Visible Then tvUDP.SelectedItem.Expanded = True
    If picICMP.Visible Then tvICMP.SelectedItem.Expanded = True
End Sub

Private Sub mnuExpandNode3_Click()
    If picICMP.Visible Then tvICMP.SelectedItem.Expanded = True
End Sub

Private Sub mnuRemNode_Click()
    If picTCP.Visible Then tvTCP.Nodes.Remove tvTCP.SelectedItem.Index
    If picUDP.Visible Then tvUDP.Nodes.Remove tvUDP.SelectedItem.Index
End Sub

Private Sub mnuRemoveNode3_Click()
    If picICMP.Visible Then tvICMP.Nodes.Remove tvICMP.SelectedItem.Index
End Sub


Private Sub mnuViewData_Click()

    If picTCP.Visible Then
        If tvTCP.SelectedItem.Tag <> vbNullString Then
            frmData.txtData.Text = tvTCP.SelectedItem.Tag
            frmData.Show
        End If
    End If
            
    If picUDP.Visible Then
        If tvUDP.SelectedItem.Tag <> vbNullString Then
            frmData.txtData.Text = tvUDP.SelectedItem.Tag
            frmData.Show
        End If
    End If
    
End Sub


Private Sub Tabs_Click()

    picTCP.Visible = False
    picUDP.Visible = False
    picICMP.Visible = False

    Select Case Tabs.SelectedItem.key
        Case "TCP"
            picTCP.Visible = True
        Case "UDP"
            picUDP.Visible = True
        Case "ICMP"
            picICMP.Visible = True
    End Select
    
End Sub


Private Sub tvTCP_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 And Not tvTCP.SelectedItem Is Nothing Then
        If tvTCP.SelectedItem.Parent Is Nothing Then
            Me.PopupMenu mnuPopup, , x + picTCP.Left + tvTCP.Left, y + picTCP.Top + tvTCP.Top
        ElseIf Not tvTCP.SelectedItem.Child Is Nothing Then
            Me.PopupMenu mnuPopup2, , x + picTCP.Left + tvTCP.Left, y + picTCP.Top + tvTCP.Top
        End If
    End If
End Sub


Private Sub tvUDP_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 And Not tvUDP.SelectedItem Is Nothing Then
        If tvUDP.SelectedItem.Parent Is Nothing Then
            Me.PopupMenu mnuPopup, , x + picUDP.Left + tvUDP.Left, y + picUDP.Top + tvUDP.Top
        ElseIf Not tvUDP.SelectedItem.Child Is Nothing Then
            Me.PopupMenu mnuPopup2, , x + picUDP.Left + tvUDP.Left, y + picUDP.Top + tvUDP.Top
        End If
    End If
End Sub


Private Sub tvICMP_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 And Not tvICMP.SelectedItem Is Nothing Then
        If tvICMP.SelectedItem.Parent Is Nothing Then
            Me.PopupMenu mnuPopup3, , x + picICMP.Left + tvICMP.Left, y + picICMP.Top + tvICMP.Top
        ElseIf Not tvICMP.SelectedItem.Child Is Nothing Then
            Me.PopupMenu mnuPopup2, , x + picICMP.Left + tvICMP.Left, y + picICMP.Top + tvICMP.Top
        End If
    End If
End Sub



Private Sub TCPDriver_RecievedPacket(IPHeader As clsIPHeader, TCPProtocol As clsTCPProtocol, Data As String)

  Dim Parent    As Node
  Dim IPH       As Node
  Dim TCPH      As Node
  Dim Flags     As Node
  
  Dim strHeader As String
  Dim strData   As String
    
    strHeader = IPHeader.SourceIP & " -> " & IPHeader.DestIP
    strData = Space(40 - Len(strHeader)) & IIf(Len(Data) <= 35, Data, Left$(Data, 35) & "...")
    
    Set Parent = tvTCP.Nodes.Add(, , , strHeader & strData, 1)
    Set IPH = tvTCP.Nodes.Add(Parent, tvwChild, , "IP Header", 3)
    Set TCPH = tvTCP.Nodes.Add(Parent, tvwChild, , "TCP Header", 3)
    
    Parent.Tag = Data
    
    If TCPLog <> 0 Then
        Print #TCPLog, "New TCP Packet [" & LenB(Data) & "]"
        Print #TCPLog, "IP Header:"
    End If
    
    With IPHeader
        tvTCP.Nodes.Add IPH, tvwChild, , "Source IP  . . . . . . . . " & .SourceIP, 2
        tvTCP.Nodes.Add IPH, tvwChild, , "Dest IP  . . . . . . . . . " & .DestIP, 2

        tvTCP.Nodes.Add IPH, tvwChild, , "Time To Live (TTL) . . . . " & .TimeToLive, 2
        tvTCP.Nodes.Add IPH, tvwChild, , "IP Version . . . . . . . . IPv" & .Version, 2
        tvTCP.Nodes.Add IPH, tvwChild, , "ID . . . . . . . . . . . . " & .ID, 2
        
        tvTCP.Nodes.Add IPH, tvwChild, , "Checksum . . . . . . . . . " & .Checksum, 2
        
        If TCPLog <> 0 Then
            Print #TCPLog, "          Source IP  . . . . . . . . " & .SourceIP
            Print #TCPLog, "          Dest IP  . . . . . . . . . " & .DestIP

            Print #TCPLog, "          Time To Live (TTL) . . . . " & .TimeToLive
            Print #TCPLog, "          IP Version . . . . . . . . IPv" & .Version
            Print #TCPLog, "          ID . . . . . . . . . . . . " & .ID
        
            Print #TCPLog, "          Checksum . . . . . . . . . " & .Checksum
        End If
    End With

    With TCPProtocol
        tvTCP.Nodes.Add TCPH, tvwChild, , "Source Port . . . . . . . " & GetPortName(.SourcePort), 2
        tvTCP.Nodes.Add TCPH, tvwChild, , "Dest Port . . . . . . . . " & GetPortName(.DestPort), 2
        
        tvTCP.Nodes.Add TCPH, tvwChild, , "Acknowledgement Number  . " & .AckNumber, 2
        tvTCP.Nodes.Add TCPH, tvwChild, , "Sequence Number . . . . . " & .SequenceNumber, 2
        
        tvTCP.Nodes.Add TCPH, tvwChild, , "Urgent Pointer  . . . . . " & .UrgentPointer, 2
        
        Set Flags = tvTCP.Nodes.Add(TCPH, tvwChild, , "Flags", 2)
        
        tvTCP.Nodes.Add TCPH, tvwChild, , "Windows . . . . . . . . . " & .Windows, 2
        tvTCP.Nodes.Add TCPH, tvwChild, , "Checksum  . . . . . . . . " & .Checksum, 2
        
        If .IsFlagSet(TCPF_ACK) Then tvTCP.Nodes.Add Flags, tvwChild, , "ACK"
        If .IsFlagSet(TCPF_FIN) Then tvTCP.Nodes.Add Flags, tvwChild, , "FIN"
        If .IsFlagSet(TCPF_PSH) Then tvTCP.Nodes.Add Flags, tvwChild, , "PSH"
        If .IsFlagSet(TCPF_RST) Then tvTCP.Nodes.Add Flags, tvwChild, , "RST"
        If .IsFlagSet(TCPF_SYN) Then tvTCP.Nodes.Add Flags, tvwChild, , "SYN"
        If .IsFlagSet(TCPF_URG) Then tvTCP.Nodes.Add Flags, tvwChild, , "URG"
    
        If TCPLog <> 0 Then
            Print #TCPLog, "TCP Header:"
        
            Print #TCPLog, "          Source Port . . . . . . . " & GetPortName(.SourcePort)
            Print #TCPLog, "          Dest Port . . . . . . . . " & GetPortName(.DestPort)
            
            Print #TCPLog, "          Acknowledgement Number  . " & .AckNumber
            Print #TCPLog, "          Sequence Number . . . . . " & .SequenceNumber
            
            Print #TCPLog, "          Urgent Pointer  . . . . . " & .UrgentPointer
            
            Print #TCPLog, "          Flags"
            
            If .IsFlagSet(TCPF_ACK) Then Print #TCPLog, "               ACK"
            If .IsFlagSet(TCPF_FIN) Then Print #TCPLog, "               FIN"
            If .IsFlagSet(TCPF_PSH) Then Print #TCPLog, "               PSH"
            If .IsFlagSet(TCPF_RST) Then Print #TCPLog, "               RST"
            If .IsFlagSet(TCPF_SYN) Then Print #TCPLog, "               SYN"
            If .IsFlagSet(TCPF_URG) Then Print #TCPLog, "               URG"
            
            Print #TCPLog, "          Windows . . . . . . . . . " & .Windows
            Print #TCPLog, "          Checksum  . . . . . . . . " & .Checksum
        
            Print #TCPLog, "Data:"
            Print #TCPLog, Data
            Print #TCPLog, vbCrLf
        End If
        
    End With
    
    
    TCPPackets = TCPPackets + 1
    BytesRecieved = BytesRecieved + LenB(Data)
    BytesRecievedPackets = BytesRecievedPackets + LenB(Data) + 40
    
    lblTCPPackets.Caption = TCPPackets
    lblBytesRecieved.Caption = BytesRecieved
    lblBytesRecievedPackets.Caption = BytesRecievedPackets

End Sub


Private Sub UDPDriver_RecievedPacket(IPHeader As clsIPHeader, UDPProtocol As clsUDPProtocol, Data As String)

  Dim Parent    As Node
  Dim IPH       As Node
  Dim UDPH      As Node

  Dim strHeader As String
  Dim strData   As String
    
    strHeader = IPHeader.SourceIP & " -> " & IPHeader.DestIP
    strData = Space(40 - Len(strHeader)) & IIf(Len(Data) <= 35, Data, Left$(Data, 35) & "...")

    Set Parent = tvUDP.Nodes.Add(, , , strHeader & strData, 1)
    Set IPH = tvUDP.Nodes.Add(Parent, tvwChild, , "IP Header", 3)
    Set UDPH = tvUDP.Nodes.Add(Parent, tvwChild, , "UDP Header", 3)
    
    Parent.Tag = Data
    
    If UDPLog <> 0 Then
        Print #UDPLog, "New UDP Packet [" & LenB(Data) & "]"
        Print #UDPLog, "IP Header:"
    End If
    
    With IPHeader
        tvUDP.Nodes.Add IPH, tvwChild, , "Source IP  . . . . . . . . " & .SourceIP, 2
        tvUDP.Nodes.Add IPH, tvwChild, , "Dest IP  . . . . . . . . . " & .DestIP, 2

        tvUDP.Nodes.Add IPH, tvwChild, , "Time To Live (TTL) . . . . " & .TimeToLive, 2
        tvUDP.Nodes.Add IPH, tvwChild, , "IP Version . . . . . . . . IPv" & .Version, 2
        tvUDP.Nodes.Add IPH, tvwChild, , "ID . . . . . . . . . . . . " & .ID, 2
        
        tvUDP.Nodes.Add IPH, tvwChild, , "Checksum . . . . . . . . . " & .Checksum, 2
        
        If UDPLog <> 0 Then
            Print #UDPLog, "          Source IP  . . . . . . . . " & .SourceIP
            Print #UDPLog, "          Dest IP  . . . . . . . . . " & .DestIP

            Print #UDPLog, "          Time To Live (TTL) . . . . " & .TimeToLive
            Print #UDPLog, "          IP Version . . . . . . . . IPv" & .Version
            Print #UDPLog, "          ID . . . . . . . . . . . . " & .ID
        
            Print #UDPLog, "          Checksum . . . . . . . . . " & .Checksum
        End If
        
    End With

    With UDPProtocol
        tvUDP.Nodes.Add UDPH, tvwChild, , "Source Port . . . . . . . " & GetPortName(.SourcePort), 2
        tvUDP.Nodes.Add UDPH, tvwChild, , "Dest Port . . . . . . . . " & GetPortName(.DestPort), 2
        
        tvUDP.Nodes.Add UDPH, tvwChild, , "Checksum  . . . . . . . . " & .Checksum, 2
        
        If UDPLog <> 0 Then
            Print #UDPLog, "UDP Header:"
            Print #UDPLog, "          Source Port . . . . . . . " & GetPortName(.SourcePort)
            Print #UDPLog, "          Dest Port . . . . . . . . " & GetPortName(.DestPort)
            Print #UDPLog, "          Checksum  . . . . . . . . " & .Checksum
            
            Print #UDPLog, "Data:"
            Print #UDPLog, Data
            Print #UDPLog, vbCrLf
        End If
    End With
    
    
    UDPPackets = UDPPackets + 1
    BytesRecieved = BytesRecieved + LenB(Data)
    BytesRecievedPackets = BytesRecievedPackets + LenB(Data) + 28
    
    lblUDPPackets.Caption = UDPPackets
    lblBytesRecieved.Caption = BytesRecieved
    lblBytesRecievedPackets.Caption = BytesRecievedPackets

End Sub



Private Sub ICMPDriver_RecievedPacket(IPHeader As clsIPHeader, ICMPProtocol As clsICMPProtocol)

  Dim Parent    As Node
  Dim IPH       As Node
  Dim ICMPH    As Node

    Set Parent = tvICMP.Nodes.Add(, , , IPHeader.SourceIP & " -> " & IPHeader.DestIP, 1)
    Set IPH = tvICMP.Nodes.Add(Parent, tvwChild, , "IP Header", 3)
    Set ICMPH = tvICMP.Nodes.Add(Parent, tvwChild, , "ICMP Header", 3)

    If ICMPLog <> 0 Then
        Print #ICMPLog, "New ICMP Packet"
        Print #ICMPLog, "IP Header:"
    End If

    With IPHeader
        tvICMP.Nodes.Add IPH, tvwChild, , "Source IP  . . . . . . . . " & .SourceIP, 2
        tvICMP.Nodes.Add IPH, tvwChild, , "Dest IP  . . . . . . . . . " & .DestIP, 2

        tvICMP.Nodes.Add IPH, tvwChild, , "Time To Live (TTL) . . . . " & .TimeToLive, 2
        tvICMP.Nodes.Add IPH, tvwChild, , "IP Version . . . . . . . . IPv" & .Version, 2
        tvICMP.Nodes.Add IPH, tvwChild, , "ID . . . . . . . . . . . . " & .ID, 2
        
        tvICMP.Nodes.Add IPH, tvwChild, , "Checksum . . . . . . . . . " & .Checksum, 2
        
        If ICMPLog <> 0 Then
            Print #ICMPLog, "          Source IP  . . . . . . . . " & .SourceIP
            Print #ICMPLog, "          Dest IP  . . . . . . . . . " & .DestIP

            Print #ICMPLog, "          Time To Live (TTL) . . . . " & .TimeToLive
            Print #ICMPLog, "          IP Version . . . . . . . . IPv" & .Version
            Print #ICMPLog, "          ID . . . . . . . . . . . . " & .ID
        
            Print #ICMPLog, "          Checksum . . . . . . . . . " & .Checksum
        End If
        
    End With

    With ICMPProtocol
        tvICMP.Nodes.Add ICMPH, tvwChild, , "Type  . . . . . . . . . . " & .GetICMPTypeStr, 2
        tvICMP.Nodes.Add ICMPH, tvwChild, , "Code  . . . . . . . . . . " & .GetICMPCodeStr, 2
        tvICMP.Nodes.Add ICMPH, tvwChild, , "ID  . . . . . . . . . . . " & .ID, 2
        tvICMP.Nodes.Add ICMPH, tvwChild, , "Checksum  . . . . . . . . " & .Checksum, 2
        
        If ICMPLog <> 0 Then
            Print #ICMPLog, "ICMP Header:"
            Print #ICMPLog, "          Type  . . . . . . . . . . " & .GetICMPTypeStr
            Print #ICMPLog, "          Code  . . . . . . . . . . " & .GetICMPCodeStr
            Print #ICMPLog, "          ID  . . . . . . . . . . . " & .ID
            Print #ICMPLog, "          Checksum  . . . . . . . . " & .Checksum
            Print #ICMPLog, vbCrLf
        End If
        
    End With
    
    ICMPPackets = ICMPPackets + 1
    BytesRecievedPackets = BytesRecievedPackets + 28
    
    lblICMPPackets.Caption = ICMPPackets
    lblBytesRecieved.Caption = BytesRecieved
    lblBytesRecievedPackets.Caption = BytesRecievedPackets
    
End Sub
