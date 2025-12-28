VERSION 5.00
Begin VB.Form frmPrint 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000009&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   11130
   ClientLeft      =   120
   ClientTop       =   360
   ClientWidth     =   14910
   Icon            =   "frmPrint.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2719.942
   ScaleMode       =   0  'User
   ScaleWidth      =   15000
   Begin VB.TextBox txtFW_Version 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0FFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   5580
      TabIndex        =   22
      Text            =   "?FW version"
      ToolTipText     =   "Firmware Version"
      Top             =   720
      Width           =   5430
   End
   Begin VB.Frame frameAxis 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Transfer Function"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6810
      Left            =   1665
      TabIndex        =   2
      Top             =   2610
      Width           =   2460
      Begin VB.Frame Frame1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Axis"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1500
         Left            =   90
         TabIndex        =   10
         ToolTipText     =   "ElectroDamp LTF #: 1==X,Y,Z; 2==tX,tY,tZ Vel (LM drive) via filters; 3==tX,tY,Z Pos (Valves drive); 4==X,Y,tZ Pos (LM drive)"
         Top             =   3150
         Width           =   2265
         Begin VB.OptionButton optAxis 
            BackColor       =   &H00FFFFC0&
            Caption         =   "  Z"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   2
            Left            =   135
            TabIndex        =   16
            Top             =   1080
            Width           =   810
         End
         Begin VB.OptionButton optAxis 
            BackColor       =   &H00FFC0FF&
            Caption         =   "  Y"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   1
            Left            =   135
            TabIndex        =   15
            Top             =   675
            Width           =   810
         End
         Begin VB.OptionButton optAxis 
            BackColor       =   &H00C0FFC0&
            Caption         =   "  X"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   0
            Left            =   135
            TabIndex        =   14
            Top             =   270
            Value           =   -1  'True
            Width           =   810
         End
         Begin VB.OptionButton optAxis 
            BackColor       =   &H00C0FFC0&
            Caption         =   "theta X"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   3
            Left            =   1125
            TabIndex        =   13
            Top             =   270
            Width           =   975
         End
         Begin VB.OptionButton optAxis 
            BackColor       =   &H00FFC0FF&
            Caption         =   "theta Y"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   4
            Left            =   1125
            TabIndex        =   12
            Top             =   675
            Width           =   975
         End
         Begin VB.OptionButton optAxis 
            BackColor       =   &H00FFFFC0&
            Caption         =   "theta Z"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   5
            Left            =   1125
            TabIndex        =   11
            Top             =   1080
            Width           =   975
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Driving Output"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1140
         Left            =   90
         TabIndex        =   7
         Top             =   1620
         Width           =   2265
         Begin VB.OptionButton optExcitOutput 
            BackColor       =   &H00FFFFC0&
            Caption         =   "Pheumatic Valves"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   1
            Left            =   135
            TabIndex        =   9
            ToolTipText     =   "Excitation signal is applied to valves which control pressure in isolators"
            Top             =   720
            Width           =   1980
         End
         Begin VB.OptionButton optExcitOutput 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Linear Motors"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   0
            Left            =   135
            TabIndex        =   8
            ToolTipText     =   "Excitation signal is applied to linear motors"
            Top             =   315
            Value           =   -1  'True
            Width           =   1965
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Test Type"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1545
         Left            =   90
         TabIndex        =   3
         Top             =   5040
         Width           =   2265
         Begin VB.OptionButton optTestInput 
            BackColor       =   &H0080C0FF&
            Caption         =   "Open Loop Velocity"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   0
            Left            =   90
            TabIndex        =   6
            ToolTipText     =   "Open Loop Velocity Transfer Function when signal goes via filters but loop is open"
            Top             =   1125
            Value           =   -1  'True
            Width           =   2055
         End
         Begin VB.OptionButton optTestInput 
            BackColor       =   &H00FFFFC0&
            Caption         =   "Prox (Position Plant)"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   2
            Left            =   90
            TabIndex        =   5
            ToolTipText     =   "Position Axis Input (combined proximity sensors signals) when driving motors"
            Top             =   315
            Width           =   2055
         End
         Begin VB.OptionButton optTestInput 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Geophone (Vel Plant)"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   1
            Left            =   90
            TabIndex        =   4
            ToolTipText     =   "Raw Signal on input of axis without filters"
            Top             =   720
            Width           =   2055
         End
      End
      Begin VB.Label lblTime 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Time"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1170
         TabIndex        =   21
         Top             =   810
         Width           =   1185
      End
      Begin VB.Label lblDate 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Date"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   945
         TabIndex        =   20
         Top             =   405
         Width           =   1410
      End
      Begin VB.Label lblAxisOrt 
         BackColor       =   &H80000004&
         Caption         =   "?"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   540
         TabIndex        =   19
         Top             =   765
         Width           =   420
      End
      Begin VB.Label Lable6 
         Alignment       =   2  'Center
         BackColor       =   &H80000004&
         Caption         =   "AXIS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   135
         TabIndex        =   18
         Top             =   405
         Width           =   825
      End
      Begin VB.Label lblAxisN 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000004&
         Caption         =   "?"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   135
         TabIndex        =   17
         Top             =   765
         Width           =   420
      End
   End
   Begin VB.PictureBox picGain 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   4910
      Left            =   4455
      ScaleHeight     =   122.952
      ScaleMode       =   0  'User
      ScaleWidth      =   135.296
      TabIndex        =   1
      Top             =   1080
      Width           =   9960
   End
   Begin VB.PictureBox picPhase 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   4910
      Left            =   4455
      ScaleHeight     =   4875
      ScaleWidth      =   9930
      TabIndex        =   0
      Top             =   6165
      Width           =   9960
   End
   Begin VB.Image Image1 
      Height          =   1605
      Left            =   1620
      Picture         =   "frmPrint.frx":640A
      Top             =   675
      Width           =   2715
   End
   Begin VB.Label Lbl_BAUD 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "115200,n,8,1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   13050
      TabIndex        =   26
      Top             =   720
      Width           =   1365
   End
   Begin VB.Label LblCOMport 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "COM1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   12150
      TabIndex        =   25
      Top             =   720
      Width           =   735
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Port"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   286
      Index           =   2
      Left            =   11610
      TabIndex        =   24
      Top             =   720
      Width           =   570
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00E0FFFF&
      Caption         =   "FW version"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   6
      Left            =   4455
      TabIndex        =   23
      Top             =   720
      Width           =   1155
   End
End
Attribute VB_Name = "frmPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
