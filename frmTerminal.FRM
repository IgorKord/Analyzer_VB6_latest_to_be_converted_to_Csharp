VERSION 5.00
Begin VB.Form frmTerminal 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   Caption         =   "Terminal"
   ClientHeight    =   6420
   ClientLeft      =   1515
   ClientTop       =   2490
   ClientWidth     =   10275
   BeginProperty Font 
      Name            =   "Courier New"
      Size            =   11.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   FontTransparent =   0   'False
   ForeColor       =   &H00000000&
   Icon            =   "frmTerminal.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6420
   ScaleWidth      =   10275
   Begin VB.Timer TimerDLL 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   765
      Top             =   270
   End
   Begin VB.Timer Timer1 
      Interval        =   250
      Left            =   90
      Top             =   270
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      NegotiatePosition=   2  'Middle
      Begin VB.Menu mnuLoadSetup 
         Caption         =   "&Load setup (Ctrl+O)"
      End
      Begin VB.Menu mnuSaveSetup 
         Caption         =   "&Save setup (Ctrl+S)"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&Hide"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu copClip 
         Caption         =   "&Copy to Clipboard (Ctrl+C)"
      End
      Begin VB.Menu ClearScreen 
         Caption         =   "Clear &Screen (Ctrl+D)"
      End
      Begin VB.Menu screen_to_history 
         Caption         =   "Copy to &History (Ctrl+A)"
      End
      Begin VB.Menu Debug 
         Caption         =   "&Debug"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu setPort 
      Caption         =   "COM &Port"
   End
   Begin VB.Menu connDial 
      Caption         =   "&Connect"
   End
   Begin VB.Menu connDisconnect 
      Caption         =   "&DISconnect"
   End
   Begin VB.Menu mnuCallMenu 
      Caption         =   "Call &Menu"
   End
   Begin VB.Menu mnuQuitMenu 
      Caption         =   "&Quit Menu"
   End
   Begin VB.Menu mnuHistory 
      Caption         =   "History&Log (Ctrl+L)"
   End
   Begin VB.Menu termHelp 
      Caption         =   "&Help"
   End
End
Attribute VB_Name = "frmTerminal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'On Error Resume Next ' cannot be here
Option Base 0 'first element in array has index 0
' Values from the VB object library.
Const vbKeyF1 As Integer = &H70
Const vbKeyF2 As Integer = &H71
Const vbKeyF3 As Integer = &H72
Const vbKeyF4 As Integer = &H73
Const vbKeyHome As Integer = &H24 '    HOME key.
Const vbKeyLeft As Integer = &H25 '    LEFT ARROW key.
Const vbKeyUp As Integer = &H26 '   UP ARROW key.
Const vbKeyRight As Integer = &H27 '    RIGHT ARROW key.
Const vbKeyDown As Integer = &H28 '    DOWN ARROW key.
Const vbKeySelect As Integer = &H29 '   SELECT key.
Const vbCFScreenFonts As Integer = &H1 ' show only Screen Font dialog

Const Rect_height As Single = 0.95


Dim Err As Long

Dim Font_Size As Long ' font size
Dim border As Long
Dim header As Long
Dim x_max As Long        'screen size
Dim y_max As Long
Dim x_gap As Long     'gap from left & botton
Dim y_gap As Long
Dim x_save As Long
Dim y_save As Long
Dim sym_width As Long    'sizes of screen symbol position
Dim sym_height As Long
Dim orgH As Long
Dim orgW As Long
Dim strtst As String
' public for Debug access
Public line_X_index As Long       ' Cursor_X current position
Public line_Y_index As Long       ' Cursor_Y current position
Public prev_y As Long             ' Cursor_Y previous position - to detect line change and print whole previous line
Public char_at_mouse_click As String    ' instead of many times print partial building char-by-char line

Dim m_x As Long    'Mouse_X current position
Dim m_y As Long    'Mouse_Y current position
Public px As Single    ' Cursor_X PREVIOUS position (used to erase cursor)
Public py As Single    ' Cursor_Y PREVIOUS position (used to erase cursor)
Dim font_change As Boolean
Dim strSpaces As String ' * 80
Dim form_resizing As Boolean
Dim ExitForm_inProgress As Boolean
Dim ScreenLine(30) As String ' array of screen lines

' raising events:
' need to create a class (? save form as class?)
' then create an instance of the class as form, declase 'WithEvents :
' Public WithEvents Instance As Class1
'Private Sub Command1_Click()
'    Instance.Text = Text1.Text
'End Sub
'
'Private Sub Form_Load()
'    Set Instance = New Class1
'End Sub
'
'Private Sub Instance_OnChange(ByVal Text As String)
'  Call MsgBox("Changed to: " + Text)
'End Sub
 
'Private FText As String
'Public Event OnCharArrival(ByVal Text As String)
'
'Private Sub CharArrival(ByVal Text As String)
'  RaiseEvent OnCharArrival(Text)
'End Sub
'
'Public Property Get Text() As String
'  Text = FText
'End Property
'
'Public Property Let Text(ByVal Value As String)
'  FText = Value
'  Call CharArrival(Value)
'End Property


'Public Sub Send_GUI_command(cmd As String) 'sets doing_cmd=True and sends cmd to port
'    If (ConnectionType = ConnCOMport) Then 'COM port
'        frmSplash.SerialComm.Output = cmd 'send command
'    Else 'telnet
'        Winsock.SendData (cmd)
'    End If
'End Sub

Sub TerminalSend(str_msg As String)
    On Error Resume Next
    Call open_port_if_closed
'    If frmSplash.SerialComm.PortOpen = False Then
'        If (ConnectionType <> ConnDEMO) Then MsgBox ("Port COM" & frmSplash.SerialComm.CommPort & ": cannot be open, switching to 'DEMO' mode")
'            ConnectionType = ConnDEMO
'    End If
'    If (ConnectionType <> ConnDEMO) Then
    Call Send_GUI_command(str_msg)
'-!-    Else
'-!-        Exit Sub
'-!-    End If
End Sub

'Sub Update_Diag()
'    DebugForm.Update_data
'End Sub

Public Sub Copy_screen_to_history()
    Dim N As Long
    Dim tmp_str As String
    For N = 1 To 25
        tmp_str = RTrim(ScreenLine(N))
        FrmHistory.AddToHistory (tmp_str)
    Next N
End Sub

Public Sub Copy_screen_to_clipboard()
    Dim N As Long
    Dim tmp_str As String
    For N = 1 To 25
        tmp_str = tmp_str & RTrim(ScreenLine(N)) & vbCrLf
    Next N
    Clipboard.Clear
    Clipboard.SetText tmp_str
End Sub

Public Sub Copy_clipboard_to_screen()
    Dim N As Long
    Dim tmp_str As String
    Dim OneChar As String
    tmp_str = Clipboard.GetText
    For N = 1 To Len(tmp_str)
        If Len(ScreenLine(line_Y_index)) > line_X_index Then
            OneChar = Mid(tmp_str, N, 1)
'            Mid(ScreenLine(line_Y_index), line_X_index, 1) = OneChar ' THIS WILL CAUSE double string "STATSTAT" change line in array
            TerminalSend (OneChar) 'send char-by char clipboard string, it will echo and print on screen
        End If

'        CurrentX = 1
'        CurrentY = line_Y_index
'
'        If line_X_index < x_max Then
'            line_X_index = line_X_index + 1
'        Else
'            Print_line (ScreenLine(line_Y_index)) ' ONE CHAR in line has been updated, -NEED THIS TO SHOW ENTERING NEW VALUE IN THE MENU
'            If line_Y_index < y_max Then
'                line_Y_index = line_Y_index + 1
'                line_X_index = 1
'            Else '************  here should be scroll
'                line_X_index = 1: line_Y_index = 1
'            End If
'        End If 'If line_X_index < x_max
    Next N
'    Print_line (ScreenLine(line_Y_index)) ' ONE CHAR in line has been updated, -NEED THIS TO SHOW ENTERING NEW VALUE IN THE MENU
End Sub

Sub Update_Form_Caption()
  Dim tmp_stat As Boolean
  tmp_stat = Init_sys_data
    Init_sys_data = True
    If (ConnectionType = ConnSimDLL) Then 'DLL is found, but we need it only if serial port not open
        frmTerminal.Caption = "Terminal <==> Simulation DLL"
    End If
    If (ConnectionType = ConnDEMO) Then
        frmTerminal.Caption = "Terminal in DEMO mode "
    End If
    ' real serial ports
    If (ConnectionType = ConnCOMport) Then 'COM port
        If strPortNum = "" Then
            strPortNum = "COM1"
        End If
        Port = Mid(strPortNum, 4)
        If frmSplash.SerialComm.PortOpen = True Then
            frmTerminal.Caption = "Terminal <==> " & strPortNum & ":  " & strPortStatus & "; at " & strPortSetting
            If AnalyzerVisible = True Then
                Analyzer.UpdateConnectControls (True)
                Analyzer.cwCOMdisconnect.Caption = strPortNum
            End If
            Analyzer.ImageDisconnected.Visible = False
            Analyzer.ImageConnected.Visible = True
            Analyzer.LblCONNparams.Caption = strPortNum & strPortStatus
        Else
            frmTerminal.Caption = "Terminal <-X-> " & strPortNum & ":  " & strPortStatus
            If AnalyzerVisible = True Then
                Analyzer.UpdateConnectControls (False)
                Analyzer.cwCOMdisconnect.Caption = strPortNum
            End If
            Analyzer.ImageDisconnected.Visible = True
            Analyzer.ImageConnected.Visible = False
        End If
        Analyzer.LblCONNparams.Caption = strPortNum & strPortStatus
        'Call Update_Diag
    End If 'COM port
    If (ConnectionType = ConnTelnet) Then 'telnet
    '    If Len(strPortNum) < 3 Then
    '        strPortNum = strTelnet_IP_ADR
    '    End If
        If frmSplash.Winsock.State = sckConnecting Then
            frmTerminal.Caption = "Terminal <-?-> IP: " & strTelnet_IP_ADR & ",  port: " & strTelnet_PORT & "  connecting..."
            If AnalyzerVisible = True Then
                Analyzer.UpdateConnectControls (True)
                Analyzer.cwCOMdisconnect.Caption = strTelnet_IP_ADR 'strPortNum
            End If
        Else
            If frmSplash.Winsock.State = sckConnected Then
                frmTerminal.Caption = "Terminal <==> IP: " & strTelnet_IP_ADR & ",  port: " & strTelnet_PORT & "  " & strPortStatus
                If AnalyzerVisible = True Then
                    Analyzer.UpdateConnectControls (True)
                    Analyzer.cwCOMdisconnect.Caption = strTelnet_IP_ADR 'strPortNum
                End If
            Else
                strPortStatus = "closed"
                frmTerminal.Caption = "Terminal <-X-> " & strTelnet_IP_ADR & ",  port: " & strTelnet_PORT & "  " & strPortStatus
                If AnalyzerVisible = True Then
                    Analyzer.UpdateConnectControls (False)
                    Analyzer.cwCOMdisconnect.Caption = strTelnet_IP_ADR 'strPortNum
                End If
            End If
        End If
    End If

  Init_sys_data = tmp_stat
End Sub

Private Sub Form_Load() 'initialization port
    On Error Resume Next
    Dim commapos As Long
    Dim t_str As String
    '****************set screen ***************
'    Me.Show 'bad: it calls non-initialuzed vars
'    Me.MaxButton = False ' does not work
    ExitForm_inProgress = False
    Application_Dir = App.Path
    KeyPreview = True 'form will check keyboard first
    Font_Size = 10   'starting font size
    strtst = "."
    sym_height = TextHeight(strtst)
    sym_width = TextWidth(strtst)
    x_max = 80 ' standard screen 80 * 25
    y_max = 25
       border = (Width - ScaleWidth) / 2           'calculate size of border (=60)
       header = (Height - ScaleHeight) - 2 * border    ' and header (=450)
       orgW = (x_max) * sym_width + 2 * border         'this automatically change ScaleWidth
       orgH = (y_max + y_gap) * sym_height + header + 2 * border ' + h_label 'this automatically change ScaleHeight
       Height = orgH 'store starting Height & Width
       Width = orgW
       FontName = "Courier New"  ' Change to proportional font.
       Font.Bold = True
       font_change = False
    x_gap = border / 2 'gap between modem controls
    y_gap = 0 'in symbols gap on botton for modem controls
    'DrawWidth = 2
    '    Line (1, 1)-(x_max, y_max), , B 'draw terminal frame
    'DebugForm.lblMousePos.Height = TextHeight(sym) * 0.8  'adjast size
    'DebugForm.lblMousePos.Width = TextWidth(sym) * 0.9    'of mouse rectangle

    'strSpaces = ""
    'For line_X_index = 1 To x_max
    '   strSpaces = strSpaces & sym
    'Next line_X_index

    strSpaces = ""
    'str = "1234567890"
    For line_X_index = 1 To 8
    ' str = str & "123456789."
       strSpaces = strSpaces & "          "
    Next line_X_index
    set_screen
    Clr_screen
    m_x = 1 ' mouse click positions
    m_y = 1
    If FormTerminalVisible = True Then Me.Show
'***************** set port **********
'    DataFileName = Application_Dir & "\TMC.port"
'    If FileExists(DataFileName) Then
'        If (ReadPortSetup(True)) <> True Then 'error reading file
'            'set port manually
'            If frmSplash.OptConnection(0).Value = True Then
'                ConnectionType = ConnCOMport
'                Port = 1
'                frmSplash.SerialComm.CommPort = Port
'                frmSplash.SerialComm.Settings = "115200,n,8,1"
'            Else 'telnet
'                ConnectionType = False
'                strTelnet_IP_ADR = frmSplash.txtHost
'                strTelnet_PORT = frmSplash.txtPort
'            End If
'        Else
'       'if port saved into file does not exist, ReadPortSetup already opens new port via dialog and updated CommPort and Settings
'            Port = frmSplash.SerialComm.CommPort
'        End If
        If (ConnectionType = ConnCOMport) Then
            If frmSplash.SerialComm.PortOpen = True Then frmSplash.SerialComm.PortOpen = False
            frmSplash.SerialComm.InputLen = 0    ' set to read entire buffer
            frmSplash.SerialComm.RThreshold = 1  ' act on each single receved character

'next_port:
            frmSplash.SerialComm.CommPort = Port
            If frmSplash.SerialComm.PortOpen = False Then frmSplash.SerialComm.PortOpen = True
            strPortStatus = " open"
            strPortNum = "COM" & LTrim(Str(Port))
            strPortSetting = frmSplash.SerialComm.Settings

            commapos = InStr(strPortSetting, ",")
            strBaud = Left(strPortSetting, commapos - 1) '"115200"
            t_str = Mid(strPortSetting, commapos + 1)
            strParity = Left(t_str, 1)   '"N"
            strBitnum = Mid(t_str, 3, 1) '"8"
            strBitstop = Mid(t_str, 5)   ' "1"

            frmSplash.SerialComm.Output = "" '//"  " & Asc(8) 'backspace '//"vers" & vbCr 'show version
        Else 'telnet
            If frmSplash.Winsock.State = sckConnecting Then strPortStatus = " connecting"
            If frmSplash.Winsock.State = sckConnected Then strPortStatus = " connected"
            If frmSplash.Winsock.State = sckClosed Then strPortStatus = " closed"
        End If
        Call Update_Form_Caption
        Exit Sub

port_error:
'    Me.Show
'    FormTerminalVisible = True
'    Load frmsplash
'    MsgBox ("Port COM" & Port & " is not available")
'    frmsplash.Show 1
End Sub 'form_load

Private Sub Form_Unload(Cancel As Integer)
    Dim tmp_stat As Boolean
    tmp_stat = Init_sys_data
    Init_sys_data = False
    Cancel = -1
'    FormTerminalVisible = False
'    If AnalyzerVisible = True Then
     Call Analyzer.ToggleTerminalControls 'changes FormTerminalVisible = not FormTerminalVisible
'    End If
    Me.Hide
    Init_sys_data = tmp_stat
End Sub

Sub Delay_sec(delay_time As Single)
    Dim CmdStartTime As Single
    Dim WaitTime As Single
    CmdStartTime = Timer
    Do
        DoEvents
        WaitTime = Timer - CmdStartTime
    Loop Until (WaitTime > delay_time)
End Sub

Private Sub mnuCallMenu_Click()
'    TerminalSend (vbCr) 'send dummy command
'    Analyzer.chkPOSmonitoring.Value = Unchecked ' disable continuous monitoring WINDOW if call menu
    If IsFormLoaded("frmMonitoring") Then
        frmMonitoring.Visible = False
        If (frmMonitoring.Visible = True) Then Call Analyzer.TogglePOSmonitoringControls
        Analyzer.cwBut_RT_Monitoring.Value = Unchecked ' disable continuous Real Time monitoring (upper part) if call menu
    End If
    Delay_sec (0.1)
    TerminalSend ("menu" & vbCr) 'send "menu" command
    Menu_Active = True
End Sub

Private Sub mnuHistory_Click()
    FrmHistory.Show
End Sub

Sub QuitMenu()
'-!- need delays between esc
    GUI_doing_command = False 'enable "stop" command or it will wait and give error
'send and wait between chars leave advanced menu, if it is there
    TerminalSend (Chr(27))
    Delay_sec (0.15)
    TerminalSend (Chr(27))
    Delay_sec (0.15)
    TerminalSend (Chr(27))
    Delay_sec (0.15)
    TerminalSend (Chr(27))
    Delay_sec (0.15)
    TerminalSend ("q")  'send "q" command
    Delay_sec (0.05)
    TerminalSend (vbCr)  'send "\r" command - get >~Unrecognized command; it sets Menu_Active = False and LastStringSent = "" after each vbCr
    Delay_sec (0.05)
'    LastStringSent = ""
'    line_X_index = 1
'    line_Y_index = line_Y_index + 1
'    Menu_Active = False
End Sub

Private Sub mnuQuitMenu_Click()
    Call QuitMenu
End Sub

Private Sub connDisconnect_Click()
    If (ConnectionType = ConnCOMport) Then
        If frmSplash.SerialComm.PortOpen <> False Then frmSplash.SerialComm.PortOpen = False
    Else 'telnet
        If frmSplash.Winsock.State = sckConnecting Or frmSplash.Winsock.State = sckConnected Then frmSplash.Winsock.Close
    End If
    strPortStatus = " closed"
    Call Update_Form_Caption
End Sub

Sub Port_Setup_Dialog()
    Load frmSplash
    If (frmSplash.Visible = False) Then frmSplash.Show vbModal
    Call Update_Form_Caption
End Sub

Private Sub setPort_Click()
    Call Port_Setup_Dialog
End Sub

Private Sub ClearScreen_Click()
    Call Clr_screen
End Sub

Private Sub Screen_to_History_Click()
    Call Copy_screen_to_history
End Sub

Private Sub connDial_Click()
    Call open_port_if_closed
'    strPortStatus = " open"
    Call Update_Form_Caption
'   Port_Setup_Dialog
End Sub

Private Sub copClip_Click()
    Call Copy_screen_to_clipboard
End Sub

'Private Sub Debug_Click()
'    DebugForm.Show
'End Sub

Function ExitForm() As Boolean
'    Dim Title As String, Msg As String, Style As Integer, answer As Integer
'    If frmSplash.SerialComm.PortOpen = True Then
'        Msg = "You are currently connected" & vbCrLf & "Are you sure you want to Exit?"
'        Style = vbYesNo + vbDefaultButton1 + vbExclamation ' Define buttons.
'        Title = "Confirm Exit"   '
'        answer = MsgBox(Msg, Style, Title)
'        If answer = vbNo Then   ' User chose No
'            ExitForm = False
'            Exit Function ' stop exit
'        Else   ' User chose Yes.
'            ExitForm = True
'            End
'        End If
'    Else
'        ExitForm = True
'        End
'    End If
    ExitForm = False
End Function

Private Sub mnuExit_Click()
    Me.Hide
End Sub

Private Sub mnuLoadSetup_Click()
    Call ReadPortSetup
End Sub

Private Sub mnuSaveSetup_Click()
    Call SavePortSetup
End Sub

'Private Sub setFont_Click()
'    Call setScreenFont
'End Sub

Private Sub termHelp_Click()
    Call Term_Help
End Sub

'Sub setScreenFont()
'    On Error Resume Next
'    dlgCommonDialog.CancelError = True
'    dlgCommonDialog.DialogTitle = "Screen Font"
'    dlgCommonDialog.Flags = vbCFScreenFonts
'    FontName = "Courier New"  ' Change to new font.
'    dlgCommonDialog.FontSize = size ' pass existing value of screen font size
'    dlgCommonDialog.FontBold = True ' pass existing value of screen font size
'    dlgCommonDialog.ShowFont  ' display Open common dialog box.
'    size = dlgCommonDialog.FontSize
'    Font.Bold = dlgCommonDialog.FontBold
'    set_screen
'End Sub


Sub set_screen() ' sets height & width and scale in symbols
    Font.Size = Font_Size
    Height = orgH 'restore twips sizes
    Width = orgW
    str_Help = "previous original Scale H,W =" & ScaleHeight & ", " & ScaleWidth
    ScaleMode = 1 ' restore scale in twips
    strtst = "."
    sym_height = TextHeight(strtst)
    sym_width = TextWidth(strtst)
    orgW = (x_max) * sym_width + 2 * border         'this automatically change ScaleWidth
    orgH = (y_max + y_gap) * sym_height + header + 2 * border ' + h_label 'this automatically change ScaleHeight
    Height = orgH 'save sizes for next call
    Width = orgW
    ScaleHeight = y_max + y_gap       ' Set height units.
    ScaleWidth = x_max '        ' Set width units.
    ScaleTop = 1                ' Set scale for Top of grid.
    ScaleLeft = 1               ' Set scale for left of grid.
'TEST:
'Cls
'Print "form width X height="; Width; " X "; Height
'Print " font size="; Font.size
'Print "  sym_height="; sym_height
'Print "   sym_width="; sym_width
'Print str

End Sub
             '**** regen **** sub, redraw screen
Sub Regen_screen() ' writes x_max spaces in each line
        Cls
        For line_Y_index = 1 To y_max + y_gap
         CurrentX = 1: CurrentY = line_Y_index: Print_line (ScreenLine(line_Y_index))
        Next line_Y_index '
    line_Y_index = py: line_X_index = px
End Sub

Private Sub Form_Resize()
    If Me.WindowState = 2 Then Me.WindowState = 0 'Prevent go maximized 1 - Minimized or 2 - Maximized.
    If Me.WindowState <> 0 Then Exit Sub '1 - Minimized or 2 - Maximized.
'*******************this code is not stable - nested recall from set_screen***********
'    If form_resizing = True Then Exit Sub
    form_resizing = True
    If orgH = 0 Then orgH = 1
    If orgW = 0 Then orgW = 1
    If Font_Size = 0 Then Font_Size = 10
    If (Height / orgH) > (Font_Size + 1) / Font_Size Or (Width / orgW) > (Font_Size + 1) / Font_Size Then
          If Font_Size < 15 Then Font_Size = Font_Size + 1
            font_change = True
    Else    ' not sagnificant change-restore previous size
        If (Width / orgW) <> 1 Then font_change = True
    End If

    If (Height / orgH) < (Font_Size - 1) / Font_Size Or (Width / orgW) < (Font_Size - 1) / Font_Size Then
        If Font_Size > 5 Then Font_Size = Font_Size - 1
            font_change = True
    Else    ' not sagnificant change-restore previous size
        If (Height / orgH) <> 1 Then font_change = True
    End If

    If font_change = False Then
        ScaleHeight = Height - 2 * border - header ' + h_label 'restore twips scale sizes
        ScaleWidth = Width - 2 * border
        'Call Update_Diag
        ScaleHeight = y_max + y_gap     ' Set height units in symbols.
        ScaleWidth = x_max              ' Set width units in symbols.
        ScaleTop = 1              ' Set scale for Top of grid.
        ScaleLeft = 1             ' Set scale for left of grid.
        'DrawWidth = 2
        '    Line (1, y_max + 1.4)-(x_max + 1, y_max + 1.5), , B 'draw terminal frame
        'DrawWidth = 1
    Else ' here form needs to change font - font_change = True
        Font.Size = Font_Size
        font_change = False
        ScaleMode = 1 ' restore scale in twips
        str_Help = "previous original Scale H,W =" & ScaleHeight & ", " & ScaleWidth
        strtst = "."
        sym_height = TextHeight(strtst)
        sym_width = TextWidth(strtst)
        orgW = (x_max) * sym_width + 2 * border         'this automatically change ScaleWidth
        orgH = (y_max + y_gap) * sym_height + header + 2 * border ' + h_label 'this automatically change ScaleHeight
        If frmTerminal.WindowState = vbNormal Then
            Height = orgH    '!! recursive self-recall 2 times
            Width = orgW     '   save sizes for next call
            ScaleHeight = y_max + y_gap ' Set height units.
            ScaleWidth = x_max          ' Set width units.
            ScaleTop = 1                ' Set scale for Top of grid.
            ScaleLeft = 1               ' Set scale for left of grid.
        End If
    End If
    Regen_screen
    form_resizing = False
'*******************************************************
End Sub 'set_screen


'Sub show_pos()
' DebugForm.lblCursorX.Caption = "CX=" & line_X_index
' DebugForm.lblCursorY.Caption = "CY=" & line_Y_index
' DebugForm.lblMouseX.Caption = "m_x=" & m_x
' DebugForm.lblMouseY.Caption = "m_y=" & m_y
'End Sub

Sub Cursor_home()
    line_X_index = 1: line_Y_index = 1
End Sub
Sub Cursor_left()
    line_X_index = line_X_index - 1
       If line_X_index = 0 Then 'if cursor was on first column
         If line_Y_index > 1 Then '- it goes up and to the end of line
             line_Y_index = line_Y_index - 1
             line_X_index = x_max
         Else
             line_X_index = 1: line_Y_index = 1
         End If
       Else   ' cursor was NOT on first column
       End If
    If LastKeyPressed <> 8 Then 'if it was not "back space"
        TerminalSend (Chr(27) & "[D")
    Else
        LastKeyPressed = 0
    End If
End Sub

Sub Cursor_up()
    If line_Y_index > 1 Then line_Y_index = line_Y_index - 1
    TerminalSend (Chr(27) & "[A")
End Sub

Sub Cursor_right()
    line_X_index = line_X_index + 1
       If line_X_index > x_max Then 'if cursor was on last column
         If line_Y_index < y_max Then '- it goes down and to the beginning of line
             line_Y_index = line_Y_index + 1
             line_X_index = 1
         Else
             line_X_index = 1: line_Y_index = 1
         End If
       Else   ' cursor was NOT on first column
       End If
    TerminalSend (Chr(27) & "[C")
End Sub
Sub Cursor_down()
    If line_Y_index < y_max Then line_Y_index = line_Y_index + 1
    TerminalSend (Chr(27) & "[B")
End Sub

Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    LastKeyPressed = KeyCode
    Select Case KeyCode
        Case vbKeyF1: Call Term_Help
'        Case vbKeyF2: MsgBox "F2 could copy text."
'        Case vbKeyF3: MsgBox "F3 could paste text."
'        Case vbKeyF4: MsgBox "F4 could format text."
        Case vbKeyHome:  Call Cursor_home
        Case vbKeyLeft:  Call Cursor_left
        Case vbKeyUp:    Call Cursor_up
        Case vbKeyRight: Call Cursor_right
        Case vbKeyDown:  Call Cursor_down
        Case 27: ' esc
        Case 8:
            Call TerminalSend(Chr(8))  'back space
    End Select
    str_Help = "cx,cy =" & line_X_index & "," & line_Y_index & "char=" & Chr(KeyCode)
    'Call Update_Diag
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
   Dim char As String
    char = Chr(KeyAscii)
'    sym = char ' for testing purpuses this sub is changing sym
    str_Help = "KeyAscii, Char = " & KeyAscii & ", " & char
    'Call Update_Diag
    'KeyAscii = Asc(UCase(char))
    Select Case KeyAscii
        Case 13: 'Enter 'entering new variable value in the menu
'        KeyAscii = InStr(1, ScreenLine(line_Y_index), ". ", 0) - 1 'binary search & compare
'        If KeyAscii > 0 Then
'           TerminalSend(Mid(ScreenLine(line_Y_index), KeyAscii, 1))
'              lblHelp.Caption = "symbol sent: " & Mid(ScreenLine(line_Y_index), KeyAscii, 1)
'        End If
        Case 3:  'CTRL-C
            Call Copy_screen_to_clipboard
            Exit Sub
        Case 22: 'CTRL-V
            Call Copy_clipboard_to_screen
            Exit Sub
        Case 12: 'CTRL-L
            FrmHistory.Show
            Exit Sub
        Case 19: 'CTRL-S
            Call SavePortSetup
            Exit Sub
        Case 15: 'CTRL-O
            Call ReadPortSetup
            Exit Sub
        Case 1: 'CTRL-A
            Call Copy_screen_to_history
            Exit Sub
        Case 4: 'CTRL-D
            Call Clr_screen
            Exit Sub
    End Select

    TerminalSend (char) 'KeyAscii
'        errCOM = TransmitCommChar(idComDev, KeyAscii)
'    If errCOM < 1 Then
'        MsgBox "error sending keyboard symbol"
'        lblHelp.Caption = "err sending code =" & errCOM
'    End If
End Sub

Sub Term_Help()
    Dim Title As String, Msg As String, Style As Integer, answer As Integer
    Title = "TMC Terminal for Analyzer"
      Msg = "Emulates ANSI terminal, Connects to COM and Telnet ports" & vbCrLf & _
            "Keeps Log of communication history: open CTRL-L, add whole screen CTRL-A" & vbCrLf & _
            "Accepts keyboard input, CTRL-C, CTRL-V" & vbCrLf & _
            "Accepts mouse clicks if menu is called" & vbCrLf & _
            "Left Button chooses menu line, Right Button works as ESC" & vbCrLf & _
            "Window is resizable and automatically adjusts font size"
    Style = vbOKOnly + vbInformation ' Define buttons.
    answer = MsgBox(Msg, Style, Title)
End Sub


Sub portproblems(Optional prog_called As String = "")
    Dim Title As String, Msg As String, Style As Integer, answer As Integer
'    GUI_doing_command = False ' release com port for next request
    If (ConnectionType = ConnCOMport) Then 'COM port
        Msg = "PC port: COM" & frmSplash.SerialComm.CommPort & " is not available, or used by other application." & vbCrLf & "Retry again / try other port? 'Cancel' will close dialog"
    End If
    If (ConnectionType = ConnTelnet) Then
        Msg = "Ethernet connection IP:" & strTelnet_IP_ADR & " port " & strTelnet_PORT & " is not available" & vbCrLf & "Retry again / try other port? 'Cancel' will close dialog"
    End If
    If (ConnectionType = ConnSimDLL) Then
        Msg = "Simulation DLL" & DLL_Name & ".dll or " & DLL_Name & ".lib not found, or wrong version: " & DLL_version & vbCrLf & "Retry again ? 'Cancel' will close dialog"
    End If
    Style = vbRetryCancel + vbDefaultButton1 + vbExclamation ' Define buttons.
    If prog_called <> "" Then
        Title = prog_called & " COM port is not available"   ' Define title.
    Else 'no argument string
        Title = " COM port is not available"   'Function_requested_TX &  Define title.
    End If
    answer = MsgBox(Msg, Style, Title)
    If answer = vbRetry Then   ' User chose Retry.
        Call Port_Setup_Dialog
        Exit Sub   ' Perform some action.
    Else   ' User chose No.
        ConnectionType = ConnDEMO ' port is not available
'        Call UpdateConnectControls(False)
        str_Help = "DEMO MODE"
    End If
    'Call Update_Diag

End Sub

'Public Sub open_port_if_closed()
'  Call open_port_if_closed
'  Call Update_Form_Caption
'End Sub

Sub Print_line(scr_line As String)
' no point to print long buffer, it will be owerwritten.
' print takes time to re=paint form, only to be replaced few moments later if buffer is too long
'    If BufferLEN < 100 Then
        Print scr_line
'        DoEvents ' enable events breaks normal menu update
'    End If
End Sub


' *****  Receive one by one, set & check commode
Function Receive(ByRef sInBuffer As String) As Long
    On Error Resume Next
    Dim i As Long
    Dim coord As Long
    Dim y_pos As Long
    Static first_arg As Byte
    Static second_arg As Byte
    Static byteOneTwo As Byte
    Static strCom As String
    Static strRxLine As String
    Static commode As Byte
    Dim sym As String ' one char
    Dim DebugStr As String 'to see Buffer as local variable in C debugger
    If ConnectionType = ConnDEMO Then Exit Function 'demo mode

'''>>>>>>>>>>>>>>>>>>>>>>>> vvvvv THIS CAUSED TO LOOSE Buffer and UNSTABLE WORK WHEN DOING LTF or Real Time
'        Buffer = sInBuffer 'read whole com buffer for test
'''>>>>>>>>>>>>>>>>>>>>>>>> ^^^^^ THIS CAUSED TO LOOSE Buffer and UNSTABLE WORK WHEN DOING LTF or Real Time
    BufferLEN = Len(sInBuffer)    ' how many symbols received
    If BufferLEN = 0 Then Exit Function
    For i = 1 To BufferLEN
        sym = Mid(sInBuffer, i, 1)
        Buffer = Buffer & sym
        DebugStr = Buffer ' For test in Visual C debugger
'        DoEvents 'if enabled, with each incomming char it causes SerialComm_OnComm event which calls this Receive.
'        If these events are too fast, and not enough time to process and exit this sub, stack is overflown
'        Buffer = Buffer & sInBuffer 'read whole com buffer, it overfloats if not clean periodically
        strRxLine = strRxLine & sym
        Select Case Asc(sym)
         ' check incoming chars: is it ESC?
            Case 27
                commode = 1 'if ESC, SET COMMAND MODE: read next byte(s) and interprete it(them) as command
                strCom = sym 'begin building command string
                GoTo nextsym
           ' if symbol is not ESC, check, is it a printable symbol?
            Case 13  ' CODE of Enter
                line_X_index = 1
                GoTo nextsym
            Case 10 'Lf
scroll_scr:
                If Menu_Active = False Then
                    If Right(strRxLine, 1) = vbCr Then strRxLine = Left(strRxLine, Len(strRxLine) - 1)
                    If Right(strRxLine, 1) = vbLf Then strRxLine = Left(strRxLine, Len(strRxLine) - 1)
                    If Right(strRxLine, 1) = vbCr Then strRxLine = Left(strRxLine, Len(strRxLine) - 1)
                    FrmHistory.AddToHistory (strRxLine)
                End If
                strRxLine = ""
                line_Y_index = line_Y_index + 1:
                If line_Y_index < y_max + 1 Then ' ***** CR + Line_forward *****
                Else ' ************ >>>>>>>>>>> scroll screen
                    For y_pos = 1 To y_max - 1
                       ScreenLine(y_pos) = ScreenLine(y_pos + 1)
                       CurrentY = y_pos
                       CurrentX = 1
                       Print ScreenLine(y_pos);
                    Next y_pos
                    ScreenLine(y_max) = strSpaces
                    CurrentY = y_max
                    CurrentX = 1
                    line_Y_index = y_max
                    Print_line (ScreenLine(y_max))
                End If
                GoTo show_cursor
            Case 9 ' Tab: move cursor ->, to the next position muptiple of 8
                line_X_index = (Int((CurrentX - 1) / 8) + 1) * 8
                If line_X_index > 80 Then
                    line_X_index = line_X_index - 80
                    GoTo scroll_scr
                End If
                For coord = CurrentX To line_X_index
                    Mid(ScreenLine(line_Y_index), coord, 1) = " "
                Next coord
                CurrentX = line_X_index
                GoTo nextsym
            Case 8 ' Back Space: move cursor <-, erase position, where cursor
                LastKeyPressed = 8
                Cursor_left
                Mid(ScreenLine(line_Y_index), line_X_index, 1) = " " ' blank previous position
                CurrentX = 1: CurrentY = line_Y_index: Print_line (ScreenLine(line_Y_index))
                GoTo nextsym
            Case 7 ' Beep:
                Beep
                GoTo nextsym
            Case Else
        End Select

        If commode = 0 Then
            GoTo put_sym ' NOT command, print symbol on screen and exit
        '**********************************
        Else    ' if command mode - START WORKING with command sequence
            strCom = strCom & sym ' ADD new symbol to command string
            commode = commode + 1 ' increase byte counter
            If Mid(strCom, 2, 1) <> "[" Then
                commode = 0
                GoTo put_sym        ' NOT command, print symbol on screen and exit
            Else                    ' we've captured beginning of command sequence and next byte must be part of it
                Menu_Active = True  ' next time we call function, we would receive next byte
            End If

            If commode = 3 Then '3-rd received byte in command mode = sym
              Select Case sym
                Case "K"  '*clear to end of line* sub, == ESC[K
ClearEnd:
                    ScreenLine(line_Y_index) = Left(ScreenLine(line_Y_index), line_X_index) + Left(strSpaces, x_max - line_X_index)
                    GoTo Print_Row ' end command mode

                Case "s"  '*saves cursor position for recall later*,  == ESC[s
Save_pos:
                    x_save = line_X_index
                    y_save = line_Y_index
                    GoTo show_cursor ' end command mode

                Case "u"  '*Return to saved cursor position*,         == ESC[u
Restore_pos:
                    line_X_index = x_save
                    line_Y_index = y_save
                    GoTo show_cursor ' end command mode
              End Select
              GoTo nextsym
            End If ' End decoding 3-rd byte
            If commode < 4 Then GoTo nextsym
        End If 'If commode = 0
          'If commode < 6 Then ' we've received at least 4-th and 5-th bytes
        Select Case sym ' check current position - if it is count or letter (NOT count)-exit
          Case "0", "1", "2", "3", "4", "5", "6", "7", "8", "9" ' - wait for command - exit
            GoTo nextsym '---->>>>>>>  ------->>>>>  ------>>>>>
          Case Else  ' if sym is NOT count(NOT ARGUMENT), it must be COMMAND
            If second_arg = 0 Then
                first_arg = Val(Mid$(strCom, 3, commode - 3)) '-3? convert string from
                byteOneTwo = commode                          '3-rd to symbol before last to integer
            Else ' conversion for second argument
                second_arg = Val(Mid$(strCom, byteOneTwo + 1, commode - byteOneTwo - 1))
            End If
        End Select

        Select Case sym
            Case "A"           '*go_up # LINES*        == ESC[#A
Cursor_up:
                For coord = 1 To first_arg
                    Cursor_up
                Next coord
            Case "B"           '*go_down # LINES*      == ESC[#B
Cursor_down:
                For coord = 1 To first_arg
                  Cursor_down
                Next coord
            Case "C"           '*go_left #symbols*     == ESC[#C
Cursor_right:
                For coord = 1 To first_arg
                  Cursor_right
                Next coord
            Case "D"           '*go_right # SYMBOLS*   == ESC[#D
Cursor_left:
                For coord = 1 To first_arg
                  Cursor_left
                Next coord
            Case "J"           '*clear screen & cursor home*   == ESC[2J
                Clr_screen

            Case ";"  '*Reports current cursor line & column* ,   == ESC[#;#R
                      '*put cursor in X,Y screen position ==ESC[#;#H or ESC[#;#f
                      ' expect second argument

                second_arg = 1 ' second argument will follow
                GoTo nextsym '------>>>> ------->>>>> ------->>>>>
             ' cases "H","f" & "R" - TWO ARGUMENTS
            Case "H", "f"       '==ESC[#;#H or ESC[#;#f
set_scr_pos:
                line_X_index = second_arg
                line_Y_index = first_arg
            Case "R"            '== ESC[#,#R
Report_pos:
                GoTo show_cursor
        End Select
          'End If ' end if commode < 6 then - decoding first argument

        GoTo show_cursor
        'End If
        ' if yes - put char on the screen  position line_X_index,yx, move cursor -->
put_sym:
          'Print Left(ScreenLine(line_Y_index), 20), "0_sym="; sym
        If Len(ScreenLine(line_Y_index)) > line_X_index Then
                Mid(ScreenLine(line_Y_index), line_X_index, 1) = sym ' change line in array
        End If
Print_Row:
        CurrentX = 1
        CurrentY = line_Y_index
        Print_line (ScreenLine(line_Y_index)) ' ONE CHAR in line has been updated, -NEED THIS TO SHOW ENTERING NEW VALUE IN THE MENU
'        If prev_y <> line_Y_index Then            ' IT DOES NOT PRINT  while entering new value, because line_Y_index does not change
'            CurrentX = 1
'            CurrentY = prev_y
'            Print ScreenLine(prev_y)    ' "prev_y" line has been updated, and ANOTHER line is building now - print whole "prev_y" line
'            prev_y = line_Y_index
'        End If

        If line_X_index < x_max Then
            line_X_index = line_X_index + 1
        Else
            If line_Y_index < y_max Then
                line_Y_index = line_Y_index + 1
                line_X_index = 1
            Else '************  here should be scroll
                line_X_index = 1: line_Y_index = 1
            End If
        End If 'If line_X_index < x_max
show_cursor:
        If char_at_mouse_click <> " " Then '-!-
            Mid(ScreenLine(m_y), m_x, 1) = char_at_mouse_click 'restore original symbol on screen - it might be overwritten
            CurrentX = 1
            CurrentY = m_y
            Print_line (ScreenLine(m_y)) ' ONE CHAR in line has been updated, -NEED THIS TO kill chars sometimes appearing in MENU after mouse click
        End If
        commode = 0: strCom = "": byteOneTwo = 0
        first_arg = 0: second_arg = 0
        If line_Y_index > y_max Then line_Y_index = y_max - 1
        If line_Y_index < 1 Then line_Y_index = 1
'            MsgBox "wrong Y position=" & Str(line_Y_index)
'            GoTo nextsym
'        End If
        If line_X_index > x_max Then line_X_index = x_max - 1
        If line_X_index < 1 Then line_X_index = 1
'            MsgBox "wrong X position=" & Str(line_X_index)
'            GoTo nextsym
'        End If
        CurrentX = line_X_index: CurrentY = line_Y_index

        'Timer1_Timer ' make flashing rectangle

        'GoTo nextsym

nextsym:
        str_Help = "cx,cy =" & line_X_index & "," & line_Y_index
        'Call Update_Diag
'        DoEvents ' enable events breaks normal menu update ' allow time to paint
    Next i
End Function
             '**** clean **** sub, cursor home == ESC[2J
Sub Clr_screen(): ' writes x_max spaces in each line
        CurrentX = 1
        For line_Y_index = 1 To y_max + 5
         ScreenLine(line_Y_index) = strSpaces ' assign 80 spaces
'         ScreenLine(line_Y_index) = "!" & strSpaces ' assign 80 spaces
         CurrentX = 1: CurrentY = line_Y_index: Print_line (ScreenLine(line_Y_index))
        Next line_Y_index '
    line_Y_index = 1: line_X_index = 1
    py = 1: px = 1
End Sub

'Private Sub form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Static BeginWord As Long
'Static EndWord As Long
'  line_X_index = Int(X): line_Y_index = Int(Y)
'DebugForm.lblMousePos.Left = line_X_index + 0.05 ' ajast center of mouse cursor
'If Mid(ScreenLine(line_Y_index), line_X_index) <> " " Then
'    EndWord = InStr(line_X_index, ScreenLine(line_Y_index), " ", 0) 'binary search & compare
'        For BeginWord = line_X_index To 1 Step -1
'            If Mid(ScreenLine(line_Y_index), BeginWord, 1) = " " Then GoTo beg_w
'        Next BeginWord
'beg_w:
'BeginWord = BeginWord + 1
'End If

'DebugForm.lblMousePos.Top = line_Y_index + 0.1
'DebugForm.lblMousePos.Left = line_X_index + 0.1 'BeginWord
''If EndWord < BeginWord Then EndWord = BeginWord + 1
''DebugForm.lblMousePos.Width = EndWord - BeginWord
'DebugForm.lblMousePos.Caption = Mid(ScreenLine(line_Y_index), line_X_index, 1) ' BeginWord, (EndWord - BeginWord))
'    DebugForm.lblMouseX.Caption = "begin" & BeginWord
'    DebugForm.lblMouseY.Caption = "end" & EndWord
'
'End Sub

Private Sub form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Static xcommand As Long
    Static dot_pos As Long
    Static tempcolor As Long
    Dim menu_char(4) As String ' one char to check if it is menu choice item: 0...9, A...F (may be up to K)
    Dim menu_item(4) As Boolean ' if true, the char before ". " is menu choice, i.e: "0. " or "D. " NOTE: "d. " IS NOT Menu choice (small letter)

    If Button = 2 Then
        GoTo send_esc        ' RIGHT button is ESC
    End If
    str_MouseX = "X=" & X
    str_MouseY = "Y=" & Y

'    CurrentX = Int(px)     ' erase previous symbol
'    CurrentY = Int(py)

    m_x = Int(X)     'convert mouse click coordinates into rectangular cursor coordinates
    m_y = Int(Y)
    If (m_x > 80) Then m_x = 80 ' sometimes it goes beyound array?
    If (m_y > 29) Then m_y = 29 ' sometimes it goes beyound array?
'    line_X_index = m_x ' save coordinates and char on screen - it might be updated in menu
'    line_Y_index = m_y
    char_at_mouse_click = Mid(ScreenLine(m_y), m_x, 1)
'test
    str_Help = "m_x,m_y =" & m_x & "," & m_y
    'Call Update_Diag
'test

    If m_y > 0 And m_y <= y_max Then  'mouse within working area
      If InStr(1, ScreenLine(m_y), "<ESC>", 0) Then ' mouse click sends ESC
send_esc:
        TerminalSend (Chr(27))
        str_Help = "symbol sent: ESC"
        px = m_x
        py = m_y ' store current x,y
        'Call Update_Diag
        Exit Sub
    End If

    xcommand = InStr(1, ScreenLine(m_y), ". ", 0) - 1 'binary search & compare of menu token ". "
    End If

    Line (px, py)-(px + 1, py + Rect_height), BackColor, B   ' erase previous cursor mark

    'Call Update_Diag

    px = m_x
    py = m_y ' store current x,y

'/////////////////////////// MENU DISCOVERY //////////////////////////////////////////////////////

    If xcommand > 0 Then ' MENU DISCOVERED, search the second, third, etc instance of ". "
      menu_char(0) = (Mid(ScreenLine(m_y), xcommand, 1))
      menu_item(0) = IsMenuChoice(menu_char(0)) 'menu choice
      dot_pos = InStr(xcommand + 2, ScreenLine(m_y), ". ", 0) - 1 'binary search & compare
      If dot_pos > 0 Then  ' there is second instance of ". ", but it can be inside text, like: "  1. FLOAT vs. DOCKpos, rel.Micr  3200"
        If dot_pos > m_x Then GoTo menu_item_found ' if click is to the left of the menu, when there is choice and subchoice in TWO COLUMNS, and user clicks on FIRST subchoice
        menu_char(1) = (Mid(ScreenLine(m_y), dot_pos, 1))
        menu_item(1) = IsMenuChoice(menu_char(1)) 'menu choice or NOT menu choice, ". " inside text, search next
        xcommand = dot_pos ', search next instance
        dot_pos = InStr(xcommand + 2, ScreenLine(m_y), ". ", 0) - 1 'binary search & compare
        If dot_pos > 0 Then  ' there is third instance of ". "
            If dot_pos > m_x Then GoTo menu_item_found ' if click is to the left of the menu, when there is choice and subchoice in TWO COLUMNS, and user clicks on FIRST subchoice
            menu_char(2) = (Mid(ScreenLine(m_y), dot_pos, 1))
            menu_item(2) = IsMenuChoice(menu_char(2)) 'menu choice
            xcommand = dot_pos
            dot_pos = InStr(xcommand + 2, ScreenLine(m_y), ". ", 0) - 1 'binary search & compare
            If dot_pos > 0 Then  ' there is forth instance of ". "
              If dot_pos > m_x Then GoTo menu_item_found ' if click is to the left of the menu, when there is choice and subchoice in TWO COLUMNS, and user clicks on FIRST subchoice
              menu_char(3) = (Mid(ScreenLine(m_y), dot_pos, 1))
              menu_item(3) = IsMenuChoice(menu_char(3)) 'menu choice
            End If
        End If
      End If
menu_item_found:
        For dot_pos = 3 To 0 Step -1
          If menu_item(dot_pos) = True Then
            TerminalSend (menu_char(dot_pos))
            str_Help = "symbol sent: " & menu_char(dot_pos)
            Exit For    'sens only the right-most char which is a menu choice, and exit
          End If
        Next dot_pos
      Exit Sub ' do not update place where to print if it is menu
    End If
'///END OF MENU DISCOVERY /////////////////////////////////////////////////////////////////////////

    Line (m_x, m_y)-(m_x + 1, m_y + Rect_height), , B            ' draw new cursor mark
    CurrentX = m_x   'update position where next print will happen
    CurrentY = m_y
    line_X_index = m_x ' save coordinates and char on screen - it might be updated in menu
    line_Y_index = m_y
End Sub

Function IsMenuChoice(str_in As String) As Boolean
    Dim menu_char As String
    menu_char = (str_in)
    If IsNumeric(menu_char) Or menu_char = "U" Or (Asc(menu_char) >= Asc("A") And Asc(menu_char) <= Asc("N")) Then 'add hex
        IsMenuChoice = True
    Else
        IsMenuChoice = False
    End If
End Function

Private Sub Timer1_Timer()
  Static flash As Byte
  Line (px, py)-(px + 1, py + Rect_height), BackColor, B ' erase previous cursor mark
  Line (line_X_index, line_Y_index)-(line_X_index + 1, line_Y_index + Rect_height), , B 'draw new cursor mark
  px = line_X_index: py = line_Y_index ' store current x,y
' Following lines cause Terminal form to be "modal" - it prevent opening other dialogs and want to be on top of other forms
'  If FormTerminalVisible = True Then
'    Me.Show vbModeless 'this does not help, still modal
'  Else
'    Me.Hide
'  End If
  If flash = 0 Then
    If (ConnectionType = ConnTelnet) Then 'telnet, update form caption
      If frmSplash.Winsock.State = sckConnecting Then
          strPortStatus = " connecting..."
      End If
      If frmSplash.Winsock.State = sckConnected Then
          strPortStatus = " open"
      End If
      Call Update_Form_Caption
    End If
    'Line (line_X_index, line_Y_index)-(line_X_index + TextWidth(sym), line_Y_index + TextHeight(sym)), BackColor, B ' erase previous cursor mark
    Line (line_X_index, line_Y_index)-(line_X_index + 1, line_Y_index + Rect_height), BackColor, B ' erase previous cursor mark
    flash = 1
  Else
    'Line (line_X_index, line_Y_index)-(line_X_index + TextWidth(sym), line_Y_index + TextHeight(sym)), FillColor, B 'draw new cursor mark
    Line (line_X_index, line_Y_index)-(line_X_index + 1, line_Y_index + Rect_height), FillColor, B 'draw new cursor mark
    flash = 0
  End If
    If Len(Buffer) > 32000 Then
      Buffer = Mid(Buffer, 32000) 'cut the beginning of buffer if it gets too long
    End If

End Sub

Private Sub TimerDLL_Timer() ' check if global flag has been changed and if it is - call function
' if DLL has separate thread, this thread comes here and calls non static member
' and gives: Unhandled exception at 0x72949197 in TMC Analyzer.exe: 0xC0000005: Access violation reading location 0x00000018.
' it actually needs Invoke method
' instead Invoke this timer checks if global flag is set
    Dim Current_buf As String
    If DLL_is_Initialized = False Then Exit Sub
    If frmTerminal.Visible = True Then 'Print symbol
'        If DLL_CharArrived = True Then
'            frmTerminal.Receive (CharFromDLL)
'            DLL_CharArrived = False ' signal to DLL interface that char has been taken
'        End If
        Current_buf = Read_DLL_whole_buffer
        DoEvents
        frmTerminal.Receive (Current_buf)
    End If
End Sub
