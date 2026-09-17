VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmSplash 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DC-2020 port selector"
   ClientHeight    =   5895
   ClientLeft      =   255
   ClientTop       =   1695
   ClientWidth     =   6705
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   5895
   ScaleWidth      =   6705
   Begin VB.Frame Frame1 
      Height          =   5820
      Left            =   90
      TabIndex        =   0
      Top             =   0
      Width           =   6540
      Begin VB.CommandButton CmdDEMO 
         BackColor       =   &H00C0FFFF&
         Caption         =   "DEMO mode"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   180
         Style           =   1  'Graphical
         TabIndex        =   29
         ToolTipText     =   "No serial connection to controller"
         Top             =   4365
         Width           =   1365
      End
      Begin VB.CommandButton cmdSIM_DLL 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Simulation DLL mode"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   180
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "No serial connection to controller"
         Top             =   4860
         Width           =   1365
      End
      Begin VB.CommandButton cmdConnect 
         BackColor       =   &H00C0FFC0&
         Caption         =   "                           Connect to Controller"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   1755
         Picture         =   "frmSplash.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   28
         ToolTipText     =   "Attempt to connect to controller using Port# and Baud Rate"
         Top             =   4365
         Width           =   4650
      End
      Begin VB.OptionButton OptConnection 
         BackColor       =   &H00FFC0FF&
         Caption         =   "USB->COM"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   0
         Left            =   3375
         TabIndex        =   10
         Top             =   2475
         Width           =   1500
      End
      Begin VB.OptionButton OptConnection 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Telnet"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   1
         Left            =   4995
         TabIndex        =   9
         Top             =   2475
         Value           =   -1  'True
         Width           =   1410
      End
      Begin VB.CommandButton cmdSaveSetup 
         BackColor       =   &H0080FF80&
         Caption         =   "Save Port Setup"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   1755
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Save Port type"
         Top             =   3510
         Width           =   1365
      End
      Begin VB.CommandButton cmdLoadPort 
         BackColor       =   &H00FFFFC0&
         Caption         =   " Load Port Setup"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   180
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Load saved Port type"
         Top             =   3510
         Width           =   1365
      End
      Begin MSComDlg.CommonDialog dlgCommonDialog 
         Left            =   2835
         Top             =   1260
         _ExtentX        =   688
         _ExtentY        =   688
         _Version        =   393216
      End
      Begin MSWinsockLib.Winsock Winsock 
         Left            =   2790
         Top             =   2700
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin MSCommLib.MSComm SerialComm 
         Left            =   2790
         Top             =   1980
         _ExtentX        =   794
         _ExtentY        =   794
         _Version        =   393216
         DTREnable       =   0   'False
         InBufferSize    =   32000
         RThreshold      =   1
         RTSEnable       =   -1  'True
         BaudRate        =   115200
      End
      Begin VB.Frame FrameTelnet 
         Caption         =   "Telnet"
         Height          =   1410
         Left            =   3330
         TabIndex        =   22
         Top             =   2835
         Width           =   3075
         Begin VB.TextBox txtHost 
            Height          =   285
            Left            =   1260
            TabIndex        =   24
            Text            =   "169.254.20.20"
            Top             =   235
            Width           =   1395
         End
         Begin VB.ComboBox ComboIPlist 
            Height          =   315
            ItemData        =   "frmSplash.frx":D9C0
            Left            =   1260
            List            =   "frmSplash.frx":D9D6
            Style           =   2  'Dropdown List
            TabIndex        =   25
            Top             =   225
            Width           =   1660
         End
         Begin VB.TextBox txtPort 
            Height          =   285
            Left            =   1260
            TabIndex        =   23
            Text            =   "2020"
            Top             =   630
            Width           =   705
         End
         Begin VB.Label Label1 
            Caption         =   "IP Address:"
            Height          =   210
            Index           =   8
            Left            =   270
            TabIndex        =   8
            Top             =   270
            Width           =   945
         End
         Begin VB.Label Label1 
            Caption         =   "Port:"
            Height          =   255
            Index           =   9
            Left            =   720
            TabIndex        =   27
            Top             =   675
            Width           =   435
         End
         Begin VB.Label lbl_IP_tip 
            Alignment       =   2  'Center
            Caption         =   "DC-2020 default static IP address"
            Height          =   420
            Left            =   45
            TabIndex        =   26
            Top             =   945
            Width           =   2940
         End
      End
      Begin VB.Frame FrameCOM 
         Caption         =   "COM port"
         Height          =   1410
         Left            =   3330
         TabIndex        =   11
         Top             =   2835
         Visible         =   0   'False
         Width           =   3075
         Begin VB.ComboBox s_Baud 
            Height          =   315
            ItemData        =   "frmSplash.frx":DA34
            Left            =   1575
            List            =   "frmSplash.frx":DA4A
            TabIndex        =   16
            Text            =   "115200"
            ToolTipText     =   "default=115200"
            Top             =   720
            Width           =   975
         End
         Begin VB.ComboBox s_port 
            Height          =   315
            ItemData        =   "frmSplash.frx":DA79
            Left            =   1575
            List            =   "frmSplash.frx":DA89
            TabIndex        =   15
            Text            =   "COM4"
            ToolTipText     =   "valid range COM1-COM16"
            Top             =   270
            Width           =   972
         End
         Begin VB.ComboBox s_Length 
            Height          =   315
            ItemData        =   "frmSplash.frx":DAA9
            Left            =   1575
            List            =   "frmSplash.frx":DAB9
            TabIndex        =   14
            Text            =   "8"
            ToolTipText     =   "default=8"
            Top             =   1575
            Width           =   975
         End
         Begin VB.ComboBox s_Parity 
            Height          =   315
            ItemData        =   "frmSplash.frx":DAC9
            Left            =   1575
            List            =   "frmSplash.frx":DAD9
            TabIndex        =   13
            Text            =   "None"
            ToolTipText     =   "default 'None'"
            Top             =   1980
            Width           =   975
         End
         Begin VB.ComboBox s_stop 
            Height          =   315
            ItemData        =   "frmSplash.frx":DAF4
            Left            =   1575
            List            =   "frmSplash.frx":DB01
            TabIndex        =   12
            Text            =   "1"
            ToolTipText     =   "default=1"
            Top             =   2385
            Width           =   975
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Stop Bit(s)"
            Height          =   255
            Index           =   7
            Left            =   540
            TabIndex        =   21
            Top             =   2430
            Width           =   975
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Parity"
            Height          =   255
            Index           =   6
            Left            =   495
            TabIndex        =   20
            Top             =   2025
            Width           =   975
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Word Length"
            Height          =   255
            Index           =   5
            Left            =   495
            TabIndex        =   19
            Top             =   1620
            Width           =   975
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Select Baud"
            Height          =   255
            Index           =   4
            Left            =   360
            TabIndex        =   18
            Top             =   765
            Width           =   975
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Select Port"
            Height          =   255
            Index           =   1
            Left            =   360
            TabIndex        =   17
            Top             =   270
            Width           =   975
         End
      End
      Begin VB.Image Image2 
         Height          =   2025
         Left            =   3690
         Picture         =   "frmSplash.frx":DB10
         Top             =   405
         Width           =   2715
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "V1.5.6 @ 01-05-2026"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   10
         Left            =   525
         TabIndex        =   7
         Top             =   495
         Width           =   1875
      End
      Begin VB.Label lblVersion 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "PN 95-38961-02 rev.G"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   240
         TabIndex        =   2
         ToolTipText     =   "TMC part number"
         Top             =   180
         Width           =   2445
      End
      Begin VB.Label lblWarning 
         Caption         =   "Choose Communication port and hit 'Connect' or hit 'DEMO'"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   240
         Left            =   315
         TabIndex        =   4
         Top             =   5490
         Width           =   6105
      End
      Begin VB.Image imgLogo 
         Height          =   570
         Index           =   1
         Left            =   1215
         Picture         =   "frmSplash.frx":1FA32
         Stretch         =   -1  'True
         Top             =   630
         Width           =   600
      End
      Begin VB.Image imgLogo 
         Height          =   2520
         Index           =   0
         Left            =   180
         Picture         =   "frmSplash.frx":1FE74
         Stretch         =   -1  'True
         Top             =   765
         Width           =   2580
      End
      Begin VB.Label lblInternalVersion 
         Alignment       =   1  'Right Justify
         Caption         =   "internal version"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3780
         TabIndex        =   1
         Top             =   135
         Width           =   2600
      End
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'On Error Resume Next ' cannot be here
Const str_Select As String = "Select"
Const str_Ignore As String = "Ignore"

Public just_show As Boolean

Private Sub cmdSaveSetup_Click()
    Call SavePortSetup
End Sub

Private Sub cmdLoadPort_Click()
    Call ReadPortSetup
    Call update_ports_from_file
End Sub

Sub show_IP_tip()
    If (InStr(txtHost.Text, "192.168.") <> 0) Then
        lbl_IP_tip = "Public Wireless IP address usually starts from 192.168.1."
    ElseIf (InStr(txtHost.Text, "192.168.168.") <> 0) Then
        lbl_IP_tip = "Custom IP address starts from 192.168.168."
    ElseIf (InStr(txtHost.Text, "169.254.20.20") <> 0) Then
        lbl_IP_tip = "DC-2020 default static IP address"
    ElseIf (InStr(txtHost.Text, "169.254.") <> 0) Then
        lbl_IP_tip = "Auto generated IP address usually starts from 169.254."
    Else
        lbl_IP_tip = "IP address loaded from TMC.port file or manually entered"
    End If
End Sub

Private Sub ComboIPlist_Change()
    txtHost.Text = ComboIPlist.List(ComboIPlist.ListIndex)
    Call show_IP_tip
End Sub

Private Sub ComboIPlist_Click()
    txtHost.Text = ComboIPlist.List(ComboIPlist.ListIndex)
    Call show_IP_tip
End Sub

Sub search_COM_ports()
    On Error Resume Next
    Dim p As Long
    s_port.Clear
    DLL_is_available = Is_DLL_available 'call DLL function
    If (DLL_is_available = True) Then '
        CmdDEMO.Height = 375
        Demo_Simul_string = "Simulation DLL" ' this will go to cwCOMdisconnect.Caption
        cmdSIM_DLL.Visible = True
    Else 'DLL not found
        Demo_Simul_string = "DEMO mode"
        cmdSIM_DLL.Visible = False
        CmdDEMO.Height = 690
    End If
'    s_port.AddItem (Demo_Simul_string)

    For p = 1 To 16
        frmSplash.SerialComm.CommPort = p
        frmSplash.SerialComm.PortOpen = True
        If frmSplash.SerialComm.PortOpen = True Then
            frmSplash.SerialComm.PortOpen = False
            s_port.AddItem ("COM" & CStr(p))
        End If
    Next p
    If (s_port.ListCount = 1) Then ' IK20221110 added if single port, show it immediately
        s_port.ListIndex = 0
        cmdConnect.Enabled = True
    End If
End Sub

Private Sub CmdDEMO_Click()
    ConnectionType = ConnDEMO
    strPortStatus = ""
    strPortSetting = ""
    strPortNum = " DEMO mode"
    Demo_Simul_string = "DEMO mode"
    Port_chosen = True
    Analyzer.Show
'    Analyzer.LblCONNparams = " DEMO mode" & strPortStatus
End Sub

Private Sub cmdSIM_DLL_Click()
    If (DLL_is_available = True) Then
        strPortNum = Demo_Simul_string
        ConnectionType = ConnSimDLL
        strPortSetting = DLL_Name & ".dll"
        strPortStatus = " open"
        Call Init_DLL
'    LogMessage = Analyzer.in_time & " " & Analyzer.in_date
'    AddLogMessage (LogMessage)
        Port_chosen = True
        Analyzer.Show
        ' if form already loaded, it does not call Form_Load again, update connection status
        Analyzer.cwCOMdisconnect.Caption = strPortNum
        Analyzer.cwCOMdisconnect.Value = True
        Analyzer.ImageConnected.Visible = True
        Analyzer.ImageDisconnected.Visible = False
        frmTerminal.TimerDLL.Enabled = True
        DoEvents
    End If
End Sub

Private Sub Form_Activate()
    Call search_COM_ports
End Sub
Function IsComPortPresent(Com_N As Long) As Boolean
    Dim p As Long
    Dim Lst As Long
    'search if port is already discovered and is in the list
    For p = 0 To s_port.ListCount
        Lst = Val(Mid(s_port.List(p), 4))
        If (Com_N = Lst) Then
            s_port.Text = s_port.List(p)
'                s_port.TopIndex = p
            IsComPortPresent = True
            Exit For
        Else
            IsComPortPresent = False
        End If
    Next p
    If IsComPortPresent = True Then
        s_port.BackColor = &H80000005 'white
        cmdConnect.Enabled = True
    Else
        s_port.BackColor = &HC0C0FF  'orange, not available
        cmdConnect.Enabled = False
    End If
End Function

Sub update_ports_from_file()
    Dim p As Long
    Dim N As Long
    Dim timeStamp As Single
    N = InStr(strPortNum, ".")
    If N <> 0 Then 'ethernet address
        OptConnection(1).Value = True
        ConnectionType = ConnTelnet
        txtHost.Text = strPortNum
        txtPort.Text = strPortSetting
        cmdConnect.Enabled = True
    End If

    N = InStr(UCase(strPortNum), "COM")

    If N <> 0 Then   'COM port
        OptConnection(0).Value = True
        ConnectionType = ConnCOMport
        timeStamp = Timer
        Do ' small delay to let control re-paint and show COM#
            s_port.Text = strPortNum
            DoEvents
        Loop Until (Timer - timeStamp > 0.1)

        N = Val(Mid(strPortNum, N + 3)) 'port number As Long
        'search if port is already discovered and is in the list
        If (IsComPortPresent(N) = True) Then
            cmdConnect.Enabled = True
        Else
            cmdConnect.Enabled = False
        End If
    End If
End Sub

Private Sub Form_Load()
On Error Resume Next
    Dim OS_64 As Boolean
    Application_Dir = App.Path
    lblInternalVersion.Caption = "Build ver." & App.Major & "." & App.Minor & "." & App.Revision
    strBaud = "115200"
    strBitnum = "8"
    strParity = "n"
    strBitstop = "1"

    OS_64 = IsDirectory("C:\Program Files (x86)")
    Windows_is_64bit = OS_64
'    If OS_64 Then
'        MsgBox ("Win64")
'    Else
'        MsgBox ("Win32")
'    End If
    If IsIDE Then
        ChDrive App.Path
        ChDir App.Path
    End If



    frmSplash.Show

    KeyPreview = True 'form will check keyboard first
'    Analyzer.Hide
    Call search_COM_ports
'    Call Get_port_settings
    strPortNum = ""     'either IP address or "COMx"
    strPortSetting = "" 'either "2020" for telnet or "115200,8,n,1"
    If FileExists(".\TMC.port") Then
        DataFileName = Application_Dir & "\TMC.port"
        Call ReadPortSetup(True) 'in will update strPortNum and strPortSetting
    Else
        Call ReadPortSetup 'in will update strPortNum and strPortSetting
    End If
    Call update_ports_from_file

    If txtHost.Text <> "" Then strTelnet_IP_ADR = txtHost.Text 'default
    If txtPort.Text <> "" Then strTelnet_PORT = txtPort.Text 'default
'    lblDescription.Caption = Prog_description
'    lblVersion.Caption = Prog_version
'    lblProductName.Caption = App.Title
    Port_chosen = False
'    just_show = False

'    Do
'        DoEvents
'    Loop Until Port_chosen = True

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
        Analyzer.Visible = True
        Me.Hide
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
        'someone hit the "x" to shut down, form unloads if "Cancel" is not set True
  Dim i As Long
  If (Port_chosen = False) Then
    i = MsgBox("Close Test program - are you sure?", vbYesNo + vbExclamation, "TMC Analyzer CLOSING")
    If i = vbNo Then
'        If ((Instrument_chosen = True) Or (just_show = True)) Then frmSplash.Visible = False
        Cancel = True
        Exit Sub 'it exits right here if "Cancel" is not set True
    End If
    If (frmSplash.SerialComm.PortOpen = True) Then
        frmSplash.SerialComm.PortOpen = False
    End If
'    LogMessage = Format(Time, "HH:MM:SS") & " " & Format(Date, "MM/DD/YYYY") & ", program closed"
'    AddLogMessage (LogMessage)

    Close
    For i = Forms.Count - 1 To 0 Step -1
'            Forms(i).Show
'            MsgBox ("form #" & i & ", " & Forms(i).Caption)
        Unload Forms(i)
    Next i
    End
  Else
'    If just_show = True Then 'final closing from Analyzer
'        LogMessage = Format(Time, "HH:MM:SS") & " " & Format(Date, "MM/DD/YYYY") & ", program closed"
'        AddLogMessage (LogMessage)
'        Unload Me
'    Else
        Me.Hide
        Cancel = True 'form unloads if "Cancel" is not set True
        Analyzer.Show
'    End If
  End If
End Sub

'Private Sub Form_Paint()
'    If just_show = True Then lblWarning.Caption = "Hit any key or click to close..."
'End Sub


Private Sub cmdConnect_Click()
    On Error Resume Next
    If OptConnection(0).Value = True Then 'COM port
        If frmSplash.SerialComm.PortOpen = True Then frmSplash.SerialComm.PortOpen = False

        strPortNum = s_port.Text
        strBaud = s_Baud.Text
        strParity = s_Parity.List(s_Parity.TopIndex)
        strBitnum = s_Length.List(s_Length.TopIndex)
        strBitstop = s_stop.List(s_stop.TopIndex)
        strPortSetting = strBaud + "," + strParity + "," + strBitnum + "," + strBitstop
        'Cls
        'Print "  strPortSetting = com "; Mid$(strPortNum, 4, 1); ": "; strPortSetting
        If frmSplash.SerialComm.PortOpen = False Then
            frmSplash.SerialComm.CommPort = Val(Mid(strPortNum, 4))
        End If
        frmSplash.SerialComm.Settings = strPortSetting
        frmTerminal.Update_Form_Caption 'causes load form and com port buffer clean
        If frmSplash.SerialComm.PortOpen = False Then
            frmSplash.SerialComm.PortOpen = True
            If frmSplash.SerialComm.PortOpen = False Then 'port cannot be open
                s_port.BackColor = &HC0C0FF 'orange
                strPortStatus = " not avail"
                Analyzer.Show 'show main form
                Call portproblems
                If (ConnectionType = ConnDEMO) Then '-!- 'demo mode
                    strPortStatus = " DEMO"
'updated once in frmSplash.search_COM_ports Analyzer.ChkCOMdisconnect.Caption = Demo_Simul_string
'                    LogMessage = Analyzer.in_time & " " & Analyzer.in_date & ", DEMO mode"
'                    AddLogMessage (LogMessage)
                    Port_chosen = True
                Else 'try another port
                    Exit Sub
                End If
            Else ' port opens successfully
'                frmSplash.SerialComm.PortOpen = True
                s_port.BackColor = &H80000005 'white
                Port = frmSplash.SerialComm.CommPort
                strPortStatus = " open"
                ConnectionType = ConnCOMport
                If Init_sys_data = False Then
                    Init_sys_data = True
                    Analyzer.UpdateConnectControls (True)
                    Analyzer.cwCOMdisconnect.Caption = "COM" & frmSplash.SerialComm.CommPort & " Connected"
                    Init_sys_data = False
                Else
                    Analyzer.UpdateConnectControls (True)
                End If
                frmTerminal.Update_Form_Caption
                Me.Hide
                Analyzer.Show 'show main form
            End If
'        Analyzer.txtFileName.Text = strPortNum & " open as " & strPortSetting
        End If
    Else 'telnet
        If frmSplash.Winsock.State <> sckClosed Then
            frmSplash.Winsock.Close
        End If
        frmSplash.Winsock.RemoteHost = txtHost.Text
        frmSplash.Winsock.RemotePort = txtPort.Text
        frmSplash.Winsock.Connect
        If frmSplash.Winsock.State <> sckConnecting And frmSplash.Winsock.State <> sckConnected And frmSplash.Winsock.State <> sckConnectionPending Then
            MsgBox ("Telnet connection error: " & (frmSplash.Winsock.State))
        Else
            strPortStatus = " open"
            If Init_sys_data = False Then
                Init_sys_data = True
                Analyzer.UpdateConnectControls (True)
                Init_sys_data = False
                ConnectionType = False
            Else
                Analyzer.UpdateConnectControls (True)
            End If
        End If
        frmTerminal.Update_Form_Caption 'causes load form and com port buffer clean
        Me.Hide
    End If
    Port_chosen = True
    AnalyzerVisible = True
    frmTerminal.Update_Form_Caption
    Me.Hide
    Analyzer.Show
End Sub

Public Sub Get_port_settings()
On Error Resume Next
    Dim first_comma_number As Long
    Dim i As Long

    strPortNum = CStr(frmSplash.SerialComm.CommPort)
    strPortSetting = frmSplash.SerialComm.Settings
    first_comma_number = InStr(strPortSetting, ",")
    strBaud = Left(strPortSetting, first_comma_number - 1)
    strParity = Mid(strPortSetting, first_comma_number + 1, 1)
    strBitnum = Mid(strPortSetting, first_comma_number + 3, 1)
    strBitstop = RTrim(Mid(strPortSetting, first_comma_number + 5))

    For i = 0 To s_Baud.ListCount
        If s_Baud.List(i) = strBaud Then
           s_Baud.TopIndex = i
           s_Baud.Text = s_Baud.List(i)
           Exit For
        End If
    Next i
    For i = 0 To s_port.ListCount
        If Mid(s_port.List(i), 4) = strPortNum Then
           s_port.TopIndex = i
           s_port.Text = s_port.List(i)
           Exit For
        End If
    Next i
    For i = 0 To s_Length.ListCount
        If s_Length.List(i) = strBitnum Then
           s_Length.TopIndex = i
           s_Length.Text = s_Length.List(i)
           Exit For
        End If
    Next i
    For i = 0 To s_Parity.ListCount
        If Left(s_Parity.List(i), 1) = UCase(strParity) Then
           s_Parity.TopIndex = i
           s_Parity.Text = s_Parity.List(i)
           Exit For
        End If
    Next i
    For i = 0 To s_stop.ListCount
        If Val(s_stop.List(i)) = Val(strBitstop) Then
           s_stop.TopIndex = i
           s_stop.Text = strBitstop
           Exit For
        End If
    Next i
End Sub


Private Sub OptConnection_Click(Index As Integer)
 If Index = 0 Then 'COM port
    FrameTelnet.Visible = False
    FrameCOM.Visible = True
    Call search_COM_ports
    ConnectionType = ConnCOMport
    'cmdConnect.Enabled = False
 Else 'Telnet
    FrameCOM.Visible = False
    FrameTelnet.Visible = True
    ConnectionType = False
    strPortNum = txtHost.Text
    strTelnet_IP_ADR = txtHost.Text
    cmdConnect.Enabled = True
 End If
End Sub

Private Sub cmdCancel_Click()
'If frmTerminal.MSComm1.PortOpen = False Then
'    frmTerminal.MSComm1.PortOpen = True
'        frmTerminal.lblPortOpen.Caption = "open"
'End If
    Hide
End Sub

Function port_problems() As Long
    Dim Title As String, Msg As String, Style As Integer, answer As Integer
'    doing_command = False ' release com port for next request
    If (ConnectionType = ConnCOMport) Then
        Msg = "PC port: " & strPortNum & " is not available, probably used by other application." & vbCrLf & "Retry again / try other port? 'Cancel' will close 'Port Setup'"
        Style = vbRetryCancel + vbDefaultButton1 + vbExclamation ' Define buttons.
        Title = " COM port is not available"   'Define title.
    Else
        Msg = "Ethernet connection IP address: " & strTelnet_IP_ADR & ", port" & strTelnet_PORT & " is not available" & vbCrLf & "Retry again / try other port? 'Cancel' will close 'Port Setup'"
        Style = vbRetryCancel + vbDefaultButton1 + vbExclamation ' Define buttons.
        Title = " Telnet Ethernet connection is not available"   'Define title.
    End If
    answer = MsgBox(Msg, Style, Title)
    port_problems = answer 'return value

'    If answer = vbRetry Then   ' User chose Retry.
'        Exit Function ' Perform some action.
'    Else   ' User chose No.
'        ConnectionType = ConnDEMO  ' port is not available
'        Call UpdateConnectControls(False)
'        lblHelp.Caption = "DEMO MODE"
'    End If
End Function

Private Sub txtHost_Change()
    strTelnet_IP_ADR = txtHost.Text
    cmdConnect.Enabled = True
    Call show_IP_tip
End Sub

Private Sub txtHost_LostFocus()
    strTelnet_IP_ADR = txtHost.Text
    cmdConnect.Enabled = True
    Call show_IP_tip
End Sub

Sub ComPortChange()
    Dim tstr As String
    Dim tport As Long
    tstr = UCase(s_port.Text)
    If (Len(tstr) >= 4) Then tstr = Mid(tstr, 4)
    If IsNumeric(tstr) Then
        tport = Val(tstr)
    End If
    Call IsComPortPresent(tport)
End Sub
Private Sub s_port_GotFocus()
    s_port.BackColor = &H80000005 'white
End Sub
Private Sub s_port_LostFocus()
    Call ComPortChange
End Sub
Private Sub s_port_Click()
    Call ComPortChange
End Sub
Private Sub s_port_Change()
'    Call ComPortChange
End Sub

'Private Sub open_port_if_closed()
'On Error Resume Next
'  If frmSplash.SerialComm.PortOpen = False Then
'    'MsgBox ("Port COM " & Port & ": closed - opening...")
'    frmSplash.SerialComm.CommPort = Port
'    frmSplash.SerialComm.PortOpen = True
'    If frmSplash.SerialComm.PortOpen = False Then ' port cannot be open
'        strPortStatus = " not avail"
''        FrmDebug.Add_Debug_Message_to_File ("caller: Open_port_if_closed(), port COM" & Port & ": is not available")
''        Call Analyzer.portproblems
''        frmTerminal.show
''        Load frmsplash
''        frmsplash.Show
'    Else
'        strPortStatus = " open"
'    End If
'  End If
''  Call frmTerminal.Update_Form_Caption
'End Sub

Private Sub SerialComm_OnComm()
    Dim errCOM As Long
    Select Case frmSplash.SerialComm.CommEvent
        'Events
        Case comEvCD    'change in Carrier Detect line
        Case comEvCTS    ' Change in the CTS line.
        Case comEvDSR   'change in DSR line (Ready-Not ready)
        '    If frmSplash.SerialComm.DSRHolding Then
        '        str_Help = "Port Not Ready"
        '    Else
        '        str_Help = "Port Ready"
        '    End If
        Case comEvReceive    ' Received RThreshold # of chars.
          If frmTerminal.Visible = True Then 'causes frmTerminal.Form_load event
            'Print sym
            errCOM = frmTerminal.Receive(frmSplash.SerialComm.Input)
          Else ' form Terminal is not visible, simply update buffer
            Call Get_Serial_Input
          End If ' if form Terminal is visible

        Case comEvRing   ' Change in the Ring Indicator.

        Case comEvSend   ' There are SThreshold number of
                      ' characters in the transmit buffer.
    'errors
        Case comEventBreak   'A Break was received
            str_Help = "BREAK"
        Case comEventCDTO    'Carrier Detect Time Out
            str_Help = "Carrier Detect Time Out"
        Case comEventDSRTO   'Data Send Ready Time Out
            str_Help = "Data Send Ready Time Out"
        'Case Frame
        '    str_Help = "Error Frame"
        Case comEventOverrun
            str_Help = "Overrun, Data LOST"
        Case comEventRxOver
            str_Help = "Receive Queue Overloaded"
        Case comEventRxParity
            str_Help = "Error received Parity"
        Case comEventTxFull
            str_Help = "Transmit Queue Full"
    End Select
    'Call frmTerminal.Update_Diag
End Sub

Private Sub Winsock_DataArrival(ByVal bytesTotal As Long)
    'On Error Resume Next 'or the first time data arrives state might still be "connecting" and winsock gives error 40006
    Dim strData As String
'    Dim strState As String
'    strState = "state=" & Winsock.State
'    If Winsock.State = sckConnecting Then
'        strPortStatus = " connecting..."
'        Exit Sub
'    End If
    If Winsock.State = sckConnected Then
'        strPortStatus = " open"
        Winsock.GetData strData, vbString
        If frmTerminal.Visible = True Then
            frmTerminal.Receive (strData)
        Else ' form Terminal is not visible, simply update buffer
'            TelnetBuf = strData
'            Call Get_Serial_Input
            Buffer = Buffer & strData
        End If ' if form Terminal is visible
    End If
End Sub


