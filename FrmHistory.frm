VERSION 5.00
Begin VB.Form FrmHistory 
   Caption         =   "History"
   ClientHeight    =   5070
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8715
   Icon            =   "FrmHistory.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5070
   ScaleWidth      =   8715
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtResponse 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4380
      Left            =   90
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   3
      Top             =   585
      Width           =   8565
   End
   Begin VB.CheckBox Chk_AutoClear 
      BackColor       =   &H0080FFFF&
      Caption         =   "Auto clear each 2000 lines"
      Height          =   285
      Left            =   6345
      TabIndex        =   2
      ToolTipText     =   "If checked, list will be Auto cleared when 1000 lines are recorded and started over to prevent slow-down"
      Top             =   180
      Value           =   1  'Checked
      Width           =   2265
   End
   Begin VB.CommandButton cmdClearHistory 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Clear History"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   4905
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "clears history list"
      Top             =   135
      Width           =   1380
   End
   Begin VB.CheckBox ChkLocalEcho 
      BackColor       =   &H0080FFFF&
      Caption         =   "Include sent commands"
      Height          =   285
      Left            =   135
      TabIndex        =   0
      ToolTipText     =   "If checked, the sending chars will be added to the list with ""-->"" preffix"
      Top             =   135
      Value           =   1  'Checked
      Width           =   2040
   End
   Begin VB.Label CharN 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "00000"
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   4095
      TabIndex        =   7
      Top             =   180
      Width           =   510
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Chars:"
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   3600
      TabIndex        =   6
      Top             =   180
      Width           =   510
   End
   Begin VB.Label LineN 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "0000"
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   2790
      TabIndex        =   5
      Top             =   180
      Width           =   510
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Lanes:"
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   2295
      TabIndex        =   4
      Top             =   180
      Width           =   510
   End
End
Attribute VB_Name = "FrmHistory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Xsize As Long
Dim Ysize As Long
Dim LineCount As Long
Dim AutoClearBit As Long

Public Sub ClearHistory()
    txtResponse.Text = ""
    LineCount = 0
    LineN.Caption = LineCount
    CharN.Caption = Len(txtResponse)
End Sub

Private Sub Chk_AutoClear_Click()
    AutoClearBit = Chk_AutoClear.Value
End Sub

Private Sub cmdClearHistory_Click()
    Call ClearHistory
End Sub

Private Sub Form_Load()
    Chk_AutoClear.Value = AutoClearBit
End Sub

'Private Sub Form_Paint()
'    Chk_AutoClear.Value = AutoClearBit
'End Sub

Private Sub Form_Resize()
    Xsize = Me.Width
    Ysize = Me.Height
    
    If (Xsize > 300) Then txtResponse.Width = Xsize - 300
    If (Ysize > 1200) Then txtResponse.Height = Ysize - 1200
    If (Xsize > Chk_AutoClear.Width + 300) Then Chk_AutoClear.Left = (Xsize - Chk_AutoClear.Width - 300)
    cmdClearHistory.Left = (Chk_AutoClear.Left - cmdClearHistory.Width - 100)
End Sub

Public Sub AddToHistory(msg_str As String)
    txtResponse.Text = txtResponse.Text & msg_str & vbCrLf
    If ((Chk_AutoClear.Value = True) And (LineCount > 2000)) Then
        ClearHistory
        LineCount = -1
    End If
    LineCount = LineCount + 1
End Sub


Private Sub txtResponse_DblClick()
    LineN.Caption = LineCount
    CharN.Caption = Len(txtResponse)
End Sub
