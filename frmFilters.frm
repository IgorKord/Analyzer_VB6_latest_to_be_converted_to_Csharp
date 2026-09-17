VERSION 5.00
Object = "{D940E4E4-6079-11CE-88CB-0020AF6845F6}#1.6#0"; "cwui.ocx"
Begin VB.Form frmFilters
   Caption         =   "Filters Tuning"
   ClientHeight    =   5865
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13665
   Icon            =   "frmFilters.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5865
   ScaleWidth      =   13665
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame FramePneumatic
      Caption         =   "Air Isolators"
      Height          =   1995
      Left            =   11835
      TabIndex        =   100
      Top             =   1035
      Width           =   1815
      Begin VB.CheckBox CheckPneumFilters
         BackColor       =   &H00FFFFC0&
         Caption         =   "Pneumatic filters"
         BeginProperty Font
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   90
         TabIndex        =   102
         ToolTipText     =   "Calculates Transfer Function for pneumatic axis, where filters 3 & 4 are ADDED to result"
         Top             =   225
         Value           =   1  'Checked
         Width           =   1590
      End
      Begin VB.CheckBox ChkAuxDamping
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF80&
         Caption         =   "Aux Damping"
         BeginProperty Font
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   90
         Style           =   1  'Graphical
         TabIndex        =   101
         ToolTipText     =   "Enables/Disables Auxilary Damping using valves - adding fast reaction peak filters #3,4 at 15-25 Hz "
         Top             =   1485
         Width           =   1635
      End
      Begin CWUIControlsLib.CWNumEdit cwAuxDampGain
         Height          =   360
         Left            =   585
         TabIndex        =   103
         Tag             =   "gain0C"
         ToolTipText     =   "Auxiliary Damping Gain of the axis"
         Top             =   1035
         Width           =   1095
         _Version        =   524288
         _ExtentX        =   1940
         _ExtentY        =   635
         _StockProps     =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}
            Name            =   "Arial"
            Size            =   12
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Reset_0         =   0   'False
         CompatibleVers_0=   524288
         NumEdit_0       =   1
         ClassName_1     =   "CCWNumEdit"
         opts_1          =   458814
         C[0]_1          =   16777152
         C[1]_1          =   0
         BorderStyle_1   =   1
         TextAlignment_1 =   2
         format_1        =   2
         ClassName_2     =   "CCWFormat"
         Format_2        =   ".#0"
         scale_1         =   3
         ClassName_3     =   "CCWScale"
         opts_3          =   65536
         dMax_3          =   10
         discInterval_3  =   1
         ValueVarType_1  =   5
         Value_Val_1     =   0.11
         IncValueVarType_1=   5
         IncValue_Val_1  =   0.1
         AccelIncVarType_1=   5
         AccelInc_Val_1  =   0.5
         RangeMinVarType_1=   5
         RangeMaxVarType_1=   5
         RangeMax_Val_1  =   50
         ButtonStyle_1   =   0
         Bindings_1      =   4
         ClassName_4     =   "CCWBindingHolderArray"
         Editor_4        =   5
         ClassName_5     =   "CCWBindingHolderArrayEditor"
         Owner_5         =   1
      End
      Begin VB.Label LblDampGain
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Caption         =   "Damping Gain"
         BeginProperty Font
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   90
         TabIndex        =   104
         Top             =   630
         Width           =   1590
      End
   End
   Begin VB.CommandButton cmdClearFilterHistory
      BackColor       =   &H008080FF&
      Caption         =   "Clear Filters History"
      BeginProperty Font
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   12105
      Style           =   1  'Graphical
      TabIndex        =   92
      ToolTipText     =   "!! WHEN CLICKED IT WILL BE A JOLT TO THE SYSTEM !!"
      Top             =   180
      Width           =   1410
   End
   Begin VB.Frame FrameFiltCoefficients
      BackColor       =   &H00E0E0E0&
      Caption         =   "Filter Coefficients"
      BeginProperty Font
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1545
      Left            =   45
      TabIndex        =   67
      Top             =   4230
      Width           =   13605
      Begin VB.CheckBox ChkEnablePRecisionLogging
         BackColor       =   &H0000C000&
         Caption         =   "Precision Log"
         BeginProperty Font
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   11970
         TabIndex        =   110
         ToolTipText     =   "If enabled writes double-single precision difference in each operation when caclulating filter factors"
         Top             =   315
         Width           =   1410
      End
      Begin VB.CheckBox ChkUseDoubleCalcWhenFalse
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         Caption         =   "32-bit float"
         BeginProperty Font
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   12060
         Style           =   1  'Graphical
         TabIndex        =   108
         ToolTipText     =   "Filter coefficients calculation precision: 32-bit float used in DC-2020, 64-bit double can be used in this simulation"
         Top             =   990
         Value           =   1  'Checked
         Width           =   1140
      End
      Begin VB.CommandButton CmdCalcMagPhaseNOW
         BackColor       =   &H00C0FFC0&
         Caption         =   "Calculate Mag-Phase"
         BeginProperty Font
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   10485
         Style           =   1  'Graphical
         TabIndex        =   69
         ToolTipText     =   "Calculates Magnitude and Phase of a filter at selected frequency based on filter coefficients"
         Top             =   225
         Width           =   1185
      End
      Begin VB.TextBox TxtFreq
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
         BeginProperty Font
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   9540
         TabIndex        =   68
         Text            =   "1.0"
         ToolTipText     =   "Enter frequency to calculate Magnitude-phase of the filter"
         Top             =   450
         Width           =   780
      End
      Begin VB.Label Label1
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "PC precision"
         BeginProperty Font
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   690
         Index           =   5
         Left            =   11970
         TabIndex        =   109
         Top             =   720
         Width           =   1320
      End
      Begin VB.Label Label1
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFE0&
         Caption         =   "Difference PPM"
         BeginProperty Font
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   250
         Index           =   3
         Left            =   90
         TabIndex        =   99
         ToolTipText     =   "Parts Per Million; =100000*(PC_coeff - DC2020_Coeff)/PC_Coeff"
         Top             =   850
         Width           =   1365
      End
      Begin VB.Label Lbl_a1_diff
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFE0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "a1"
         BeginProperty Font
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   1530
         TabIndex        =   98
         Top             =   820
         Width           =   1500
      End
      Begin VB.Label Lbl_a2_diff
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFE0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "a2"
         BeginProperty Font
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   3105
         TabIndex        =   97
         Top             =   820
         Width           =   1500
      End
      Begin VB.Label Lbl_b0_diff
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFE0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "b0"
         BeginProperty Font
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   4680
         TabIndex        =   96
         Top             =   820
         Width           =   1500
      End
      Begin VB.Label Lbl_b1_diff
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFE0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "b1"
         BeginProperty Font
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   6255
         TabIndex        =   95
         Top             =   820
         Width           =   1500
      End
      Begin VB.Label Lbl_b2_diff
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFE0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "b2"
         BeginProperty Font
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   7830
         TabIndex        =   94
         Top             =   820
         Width           =   1500
      End
      Begin VB.Label LblPhase
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "phase"
         BeginProperty Font
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   10665
         TabIndex        =   78
         ToolTipText     =   "Phase (degrees) of the chosen filter"
         Top             =   1140
         Width           =   1005
      End
      Begin VB.Label LblMag
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "mag"
         BeginProperty Font
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   9495
         TabIndex        =   77
         ToolTipText     =   "Magnitude (dB) of the chosen filter"
         Top             =   1140
         Width           =   1005
      End
      Begin VB.Label Label1
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         Caption         =   "Freq"
         BeginProperty Font
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   19
         Left            =   9540
         TabIndex        =   91
         ToolTipText     =   "Choose axis OUTPUT for excitation"
         Top             =   225
         Width           =   780
      End
      Begin VB.Label Label1
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Caption         =   "Phase"
         BeginProperty Font
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   18
         Left            =   10665
         TabIndex        =   90
         ToolTipText     =   "Phase (degrees) of the chosen filter"
         Top             =   890
         Width           =   1005
      End
      Begin VB.Label Label1
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Caption         =   "Mag"
         BeginProperty Font
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   17
         Left            =   9495
         TabIndex        =   89
         ToolTipText     =   "Magnitude (dB) of the chosen filter"
         Top             =   890
         Width           =   1005
      End
      Begin VB.Label Label1
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "b2"
         BeginProperty Font
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   16
         Left            =   8235
         TabIndex        =   88
         Top             =   225
         Width           =   690
      End
      Begin VB.Label Lbl_b2
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "b2"
         BeginProperty Font
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   7830
         TabIndex        =   87
         Top             =   495
         Width           =   1500
      End
      Begin VB.Label Label1
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "b1"
         BeginProperty Font
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   14
         Left            =   6660
         TabIndex        =   86
         Top             =   225
         Width           =   690
      End
      Begin VB.Label Lbl_b1
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "b1"
         BeginProperty Font
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   6255
         TabIndex        =   85
         Top             =   495
         Width           =   1500
      End
      Begin VB.Label Label1
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "b0"
         BeginProperty Font
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   10
         Left            =   5085
         TabIndex        =   84
         Top             =   225
         Width           =   690
      End
      Begin VB.Label Lbl_b0
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "b0"
         BeginProperty Font
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   4680
         TabIndex        =   83
         Top             =   495
         Width           =   1500
      End
      Begin VB.Label Label1
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "a2"
         BeginProperty Font
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   9
         Left            =   3510
         TabIndex        =   82
         Top             =   225
         Width           =   690
      End
      Begin VB.Label Lbl_a2
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "a2"
         BeginProperty Font
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   3105
         TabIndex        =   81
         Top             =   495
         Width           =   1500
      End
      Begin VB.Label Label1
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "a1"
         BeginProperty Font
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   8
         Left            =   1935
         TabIndex        =   80
         Top             =   225
         Width           =   690
      End
      Begin VB.Label Lbl_a1
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "a1"
         BeginProperty Font
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   1530
         TabIndex        =   79
         Top             =   495
         Width           =   1500
      End
      Begin VB.Label Label1
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "From Controller"
         BeginProperty Font
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   520
         Index           =   1
         Left            =   90
         TabIndex        =   76
         Top             =   270
         Width           =   1365
      End
      Begin VB.Label Label1
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "PC calculated"
         BeginProperty Font
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   2
         Left            =   90
         TabIndex        =   75
         Top             =   1160
         Width           =   1365
      End
      Begin VB.Label Lbl_PC_a1
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "a1"
         BeginProperty Font
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   1530
         TabIndex        =   74
         ToolTipText     =   "PC-calculated value"
         Top             =   1160
         Width           =   1500
      End
      Begin VB.Label Lbl_PC_a2
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "a2"
         BeginProperty Font
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   3105
         TabIndex        =   73
         ToolTipText     =   "PC-calculated value"
         Top             =   1160
         Width           =   1500
      End
      Begin VB.Label Lbl_PC_b0
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "b0"
         BeginProperty Font
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   4680
         TabIndex        =   72
         ToolTipText     =   "PC-calculated value"
         Top             =   1160
         Width           =   1500
      End
      Begin VB.Label Lbl_PC_b1
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "b1"
         BeginProperty Font
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   6255
         TabIndex        =   71
         ToolTipText     =   "PC-calculated value"
         Top             =   1160
         Width           =   1500
      End
      Begin VB.Label Lbl_PC_b2
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "b2"
         BeginProperty Font
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   7830
         TabIndex        =   70
         ToolTipText     =   "PC-calculated value"
         Top             =   1160
         Width           =   1500
      End
   End
   Begin VB.Frame FrameTuning
      BackColor       =   &H00E0E0E0&
      BeginProperty Font
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4230
      Left            =   45
      TabIndex        =   0
      ToolTipText     =   "After changing axis, filters are NOT updated and must be red if needed. This allows to copy filter table to the OTHER axis"
      Top             =   0
      Width           =   11756
      Begin VB.CheckBox ChkRealTimeUpdateTF
         BackColor       =   &H00FFFF80&
         Caption         =   "Predict TF instant update"
         BeginProperty Font
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2745
         TabIndex        =   66
         ToolTipText     =   "Re-calculates in PC (not in controller)  'a' and 'b' coefficients to show changing Transfer Function "
         Top             =   720
         Width           =   2310
      End
      Begin VB.OptionButton Opt_ALLfilters_or_ChangedOnly
         BackColor       =   &H00C0C0FF&
         Caption         =   "Changed"
         BeginProperty Font
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   7470
         TabIndex        =   64
         Top             =   720
         Value           =   -1  'True
         Width           =   1005
      End
      Begin VB.OptionButton Opt_ALLfilters_or_ChangedOnly
         BackColor       =   &H00C0FFFF&
         Caption         =   "ALL"
         BeginProperty Font
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   6750
         TabIndex        =   63
         Top             =   720
         Width           =   690
      End
      Begin VB.CommandButton CmdFilterRefresh
         BackColor       =   &H00FF80FF&
         Caption         =   "Read Axis Filters"
         BeginProperty Font
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3195
         Style           =   1  'Graphical
         TabIndex        =   62
         ToolTipText     =   "Read filters of axis. NOTE: this table will be overwritten, SAVE if needed"
         Top             =   270
         Width           =   1860
      End
      Begin VB.CommandButton cmdSaveParamsIntoFLASH
         BackColor       =   &H00FFFF80&
         Caption         =   "Save Params into FLASH"
         BeginProperty Font
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   0
         Left            =   10215
         Style           =   1  'Graphical
         TabIndex        =   61
         ToolTipText     =   "Saves controller parameters into internal non-volatile memory"
         Top             =   270
         Width           =   1410
      End
      Begin VB.CommandButton cmdLoadDefaults
         BackColor       =   &H0080C0FF&
         Caption         =   "Load Default Parameters"
         BeginProperty Font
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   0
         Left            =   8640
         Style           =   1  'Graphical
         TabIndex        =   60
         ToolTipText     =   "Temporary sets controller parameters to factory default in RAM. Parameters in FLASH are not affected until ""save"" command"
         Top             =   270
         Width           =   1410
      End
      Begin VB.CommandButton CmdUpdateFilter
         BackColor       =   &H00C0FFC0&
         Caption         =   "Update Filters"
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
         Left            =   6750
         Style           =   1  'Graphical
         TabIndex        =   59
         ToolTipText     =   "Update Controller filters with this table. DO NOT FORGET to save into FLASH!"
         Top             =   270
         Width           =   1725
      End
      Begin VB.CheckBox ChkShowOneFilterTF
         BackColor       =   &H0000C000&
         Caption         =   "Show Filter TF"
         BeginProperty Font
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5220
         TabIndex        =   58
         ToolTipText     =   "Shows/hides Transfer Function of active filter"
         Top             =   315
         Value           =   1  'Checked
         Width           =   1410
      End
      Begin VB.CheckBox ChkShowAxisTF
         BackColor       =   &H000080FF&
         Caption         =   "Show Axis TF"
         BeginProperty Font
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5220
         TabIndex        =   57
         ToolTipText     =   "Shows/hides Transfer Function of whole axis"
         Top             =   720
         Value           =   1  'Checked
         Width           =   1410
      End
      Begin VB.ComboBox ComboFilterAxis
         BeginProperty Font
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "frmFilters.frx":08CA
         Left            =   1575
         List            =   "frmFilters.frx":0912
         Style           =   2  'Dropdown List
         TabIndex        =   1
         ToolTipText     =   "Choose Filter Axis to view and edit"
         Top             =   270
         Width           =   1470
      End
      Begin VB.Frame Frame5
         BackColor       =   &H00C0FFFF&
         Height          =   2985
         Left            =   135
         TabIndex        =   2
         Top             =   1125
         Width           =   11490
         Begin VB.CommandButton CmdCopyFilters
            BackColor       =   &H008080FF&
            Caption         =   "Set all 4 axes"
            BeginProperty Font
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   1935
            Style           =   1  'Graphical
            TabIndex        =   93
            ToolTipText     =   "Update similar axes with this table (i.e. if Z2 is chosen, Z1, Z2, Z3, Z4 will be updated) "
            Top             =   90
            Visible         =   0   'False
            Width           =   1545
         End
         Begin VB.Frame FrameFilter
            BackColor       =   &H00C0FFC0&
            BorderStyle     =   0  'None
            Height          =   420
            Left            =   585
            TabIndex        =   3
            Top             =   450
            Width           =   10845
            Begin VB.ComboBox ComboFilterTYPE
               BeginProperty Font
                  Name            =   "Tahoma"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               ItemData        =   "frmFilters.frx":09E7
               Left            =   45
               List            =   "frmFilters.frx":0A2C
               Style           =   2  'Dropdown List
               TabIndex        =   4
               ToolTipText     =   "Choose filter type"
               Top             =   30
               Width           =   3090
            End
            Begin CWUIControlsLib.CWNumEdit cwNumFilterParam
               Height          =   360
               Index           =   1
               Left            =   3330
               TabIndex        =   5
               ToolTipText     =   "Corner Frequency 1"
               Top             =   30
               Width           =   1400
               _Version        =   524288
               _ExtentX        =   2469
               _ExtentY        =   635
               _StockProps     =   4
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}
                  Name            =   "Arial"
                  Size            =   11.25
                  Charset         =   204
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Reset_0         =   0   'False
               CompatibleVers_0=   524288
               NumEdit_0       =   1
               ClassName_1     =   "CCWNumEdit"
               opts_1          =   458814
               C[0]_1          =   12632319
               C[1]_1          =   0
               BorderStyle_1   =   1
               TextAlignment_1 =   1
               format_1        =   2
               ClassName_2     =   "CCWFormat"
               Format_2        =   ".##0"
               scale_1         =   3
               ClassName_3     =   "CCWScale"
               opts_3          =   65536
               dMax_3          =   10
               discInterval_3  =   1
               ValueVarType_1  =   5
               Value_Val_1     =   0.001
               IncValueVarType_1=   5
               IncValue_Val_1  =   0.1
               AccelIncVarType_1=   5
               AccelInc_Val_1  =   1
               RangeMinVarType_1=   5
               RangeMin_Val_1  =   0.001
               RangeMaxVarType_1=   5
               RangeMax_Val_1  =   5000
               ButtonStyle_1   =   0
               Bindings_1      =   4
               ClassName_4     =   "CCWBindingHolderArray"
               Editor_4        =   5
               ClassName_5     =   "CCWBindingHolderArrayEditor"
               Owner_5         =   1
            End
            Begin CWUIControlsLib.CWNumEdit cwNumFilterParam
               Height          =   360
               Index           =   3
               Left            =   4950
               TabIndex        =   6
               ToolTipText     =   "Q factor at Frequency 1"
               Top             =   30
               Width           =   1275
               _Version        =   524288
               _ExtentX        =   2249
               _ExtentY        =   635
               _StockProps     =   4
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}
                  Name            =   "Arial"
                  Size            =   11.25
                  Charset         =   204
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Reset_0         =   0   'False
               CompatibleVers_0=   524288
               NumEdit_0       =   1
               ClassName_1     =   "CCWNumEdit"
               opts_1          =   458814
               C[0]_1          =   12632319
               C[1]_1          =   0
               BorderStyle_1   =   1
               TextAlignment_1 =   1
               format_1        =   2
               ClassName_2     =   "CCWFormat"
               Format_2        =   ".#0"
               scale_1         =   3
               ClassName_3     =   "CCWScale"
               opts_3          =   65536
               dMax_3          =   10
               discInterval_3  =   1
               ValueVarType_1  =   5
               Value_Val_1     =   1.111
               IncValueVarType_1=   5
               IncValue_Val_1  =   0.1
               AccelIncVarType_1=   5
               AccelInc_Val_1  =   1
               RangeMinVarType_1=   5
               RangeMin_Val_1  =   0.001
               RangeMaxVarType_1=   5
               RangeMax_Val_1  =   5000
               ButtonStyle_1   =   0
               Bindings_1      =   4
               ClassName_4     =   "CCWBindingHolderArray"
               Editor_4        =   5
               ClassName_5     =   "CCWBindingHolderArrayEditor"
               Owner_5         =   1
            End
            Begin CWUIControlsLib.CWNumEdit cwNumFilterParam
               Height          =   360
               Index           =   2
               Left            =   6435
               TabIndex        =   7
               ToolTipText     =   "Corner Frequency 2"
               Top             =   30
               Width           =   1395
               _Version        =   524288
               _ExtentX        =   2469
               _ExtentY        =   635
               _StockProps     =   4
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}
                  Name            =   "Arial"
                  Size            =   11.25
                  Charset         =   204
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Reset_0         =   0   'False
               CompatibleVers_0=   524288
               NumEdit_0       =   1
               ClassName_1     =   "CCWNumEdit"
               opts_1          =   458814
               C[0]_1          =   12632319
               C[1]_1          =   0
               BorderStyle_1   =   1
               TextAlignment_1 =   1
               format_1        =   2
               ClassName_2     =   "CCWFormat"
               Format_2        =   ".##0"
               scale_1         =   3
               ClassName_3     =   "CCWScale"
               opts_3          =   65536
               dMax_3          =   10
               discInterval_3  =   1
               ValueVarType_1  =   5
               Value_Val_1     =   1.111
               IncValueVarType_1=   5
               IncValue_Val_1  =   0.1
               AccelIncVarType_1=   5
               AccelInc_Val_1  =   1
               RangeMinVarType_1=   5
               RangeMin_Val_1  =   0.001
               RangeMaxVarType_1=   5
               RangeMax_Val_1  =   5000
               ButtonStyle_1   =   0
               Bindings_1      =   4
               ClassName_4     =   "CCWBindingHolderArray"
               Editor_4        =   5
               ClassName_5     =   "CCWBindingHolderArrayEditor"
               Owner_5         =   1
            End
            Begin CWUIControlsLib.CWNumEdit cwNumFilterParam
               Height          =   360
               Index           =   4
               Left            =   8055
               TabIndex        =   8
               ToolTipText     =   "Q factor at Frequency 2"
               Top             =   30
               Width           =   1275
               _Version        =   524288
               _ExtentX        =   2249
               _ExtentY        =   635
               _StockProps     =   4
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}
                  Name            =   "Arial"
                  Size            =   11.25
                  Charset         =   204
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Reset_0         =   0   'False
               CompatibleVers_0=   524288
               NumEdit_0       =   1
               ClassName_1     =   "CCWNumEdit"
               opts_1          =   458814
               C[0]_1          =   12632319
               C[1]_1          =   0
               BorderStyle_1   =   1
               TextAlignment_1 =   1
               format_1        =   2
               ClassName_2     =   "CCWFormat"
               Format_2        =   ".#0"
               scale_1         =   3
               ClassName_3     =   "CCWScale"
               opts_3          =   65536
               dMax_3          =   10
               discInterval_3  =   1
               ValueVarType_1  =   5
               Value_Val_1     =   1.11
               IncValueVarType_1=   5
               IncValue_Val_1  =   0.1
               AccelIncVarType_1=   5
               AccelInc_Val_1  =   1
               RangeMinVarType_1=   5
               RangeMin_Val_1  =   0.001
               RangeMaxVarType_1=   5
               RangeMax_Val_1  =   5000
               ButtonStyle_1   =   0
               Bindings_1      =   4
               ClassName_4     =   "CCWBindingHolderArray"
               Editor_4        =   5
               ClassName_5     =   "CCWBindingHolderArrayEditor"
               Owner_5         =   1
            End
            Begin CWUIControlsLib.CWNumEdit cwNumFilterParam
               Height          =   360
               Index           =   0
               Left            =   9540
               TabIndex        =   9
               ToolTipText     =   "Gain of the Filter"
               Top             =   30
               Width           =   1275
               _Version        =   524288
               _ExtentX        =   2249
               _ExtentY        =   635
               _StockProps     =   4
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}
                  Name            =   "Arial"
                  Size            =   11.25
                  Charset         =   204
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Reset_0         =   0   'False
               CompatibleVers_0=   524288
               NumEdit_0       =   1
               ClassName_1     =   "CCWNumEdit"
               opts_1          =   458814
               C[0]_1          =   12632319
               C[1]_1          =   0
               BorderStyle_1   =   1
               TextAlignment_1 =   1
               format_1        =   2
               ClassName_2     =   "CCWFormat"
               Format_2        =   ".#0"
               scale_1         =   3
               ClassName_3     =   "CCWScale"
               opts_3          =   65536
               dMax_3          =   10
               discInterval_3  =   1
               ValueVarType_1  =   5
               Value_Val_1     =   -1.11
               IncValueVarType_1=   5
               IncValue_Val_1  =   0.1
               AccelIncVarType_1=   5
               AccelInc_Val_1  =   1
               RangeMinVarType_1=   5
               RangeMin_Val_1  =   -100
               RangeMaxVarType_1=   5
               RangeMax_Val_1  =   100
               ButtonStyle_1   =   0
               Bindings_1      =   4
               ClassName_4     =   "CCWBindingHolderArray"
               Editor_4        =   5
               ClassName_5     =   "CCWBindingHolderArrayEditor"
               Owner_5         =   1
            End
         End
         Begin VB.Label LblFilter
            Appearance      =   0  'Flat
            BackColor       =   &H008080FF&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "  1"
            BeginProperty Font
               Name            =   "Tahoma"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   420
            Index           =   0
            Left            =   0
            TabIndex        =   15
            ToolTipText     =   "Feed Back Filter"
            Top             =   450
            Width           =   11445
         End
         Begin VB.Label LblFilter
            Appearance      =   0  'Flat
            BackColor       =   &H008080FF&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "  2"
            BeginProperty Font
               Name            =   "Tahoma"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   420
            Index           =   1
            Left            =   0
            TabIndex        =   14
            ToolTipText     =   "Feed Back Filter"
            Top             =   855
            Width           =   11445
         End
         Begin VB.Label LblFilter
            Appearance      =   0  'Flat
            BackColor       =   &H008080FF&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "  3"
            BeginProperty Font
               Name            =   "Tahoma"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   420
            Index           =   2
            Left            =   0
            TabIndex        =   13
            ToolTipText     =   "Feed Back Filter"
            Top             =   1260
            Width           =   11445
         End
         Begin VB.Label LblFilter
            Appearance      =   0  'Flat
            BackColor       =   &H008080FF&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "  4"
            BeginProperty Font
               Name            =   "Tahoma"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   420
            Index           =   3
            Left            =   0
            TabIndex        =   12
            ToolTipText     =   "Feed Back Filter"
            Top             =   1665
            Width           =   11445
         End
         Begin VB.Label LblFilter
            Appearance      =   0  'Flat
            BackColor       =   &H008080FF&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "  5"
            BeginProperty Font
               Name            =   "Tahoma"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   420
            Index           =   4
            Left            =   0
            TabIndex        =   11
            ToolTipText     =   "Feed Back Filter"
            Top             =   2070
            Width           =   11445
         End
         Begin VB.Label LblFilter
            Appearance      =   0  'Flat
            BackColor       =   &H0080C0FF&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "  6"
            BeginProperty Font
               Name            =   "Tahoma"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   420
            Index           =   5
            Left            =   0
            TabIndex        =   10
            ToolTipText     =   "This is input filter for stage FF adaptive algorithm (input error)"
            Top             =   2475
            Width           =   11445
         End
         Begin VB.Label Lbl_FilterGain
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Gain"
            BeginProperty Font
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   330
            Index           =   0
            Left            =   10395
            TabIndex        =   46
            ToolTipText     =   "Usually =1, for PID it is Proportional Gain"
            Top             =   495
            Width           =   1005
         End
         Begin VB.Label Lbl_FilterQ2
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Q_2"
            BeginProperty Font
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   330
            Index           =   0
            Left            =   8910
            TabIndex        =   47
            ToolTipText     =   "For 2nd order filters controls peak at Frequency 2"
            Top             =   495
            Width           =   1005
         End
         Begin VB.Label Lbl_FilterQ1
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Q_1"
            BeginProperty Font
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   330
            Index           =   0
            Left            =   5790
            TabIndex        =   48
            Top             =   495
            Width           =   1005
         End
         Begin VB.Label Lbl_FilterFreq2
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Freq_2"
            BeginProperty Font
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   330
            Index           =   0
            Left            =   7245
            TabIndex        =   49
            ToolTipText     =   "For PID it is Differential Gain"
            Top             =   495
            Width           =   1155
         End
         Begin VB.Label Lbl_FilterFreq1
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Freq_1"
            BeginProperty Font
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   330
            Index           =   0
            Left            =   4140
            TabIndex        =   50
            ToolTipText     =   "For PID it is Integral Gain"
            Top             =   495
            Width           =   1150
         End
         Begin VB.Label Lbl_FilterType
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Filter Type"
            BeginProperty Font
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   330
            Index           =   0
            Left            =   630
            TabIndex        =   51
            Top             =   495
            Width           =   2850
         End
         Begin VB.Label Lbl_FilterGain
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Gain"
            BeginProperty Font
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   330
            Index           =   5
            Left            =   10395
            TabIndex        =   16
            ToolTipText     =   "Usually =1, for PID it is Proportional Gain"
            Top             =   2520
            Width           =   1005
         End
         Begin VB.Label Lbl_FilterQ2
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Q_2"
            BeginProperty Font
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   330
            Index           =   5
            Left            =   8910
            TabIndex        =   17
            ToolTipText     =   "For 2nd order filters controls peak at Frequency 2"
            Top             =   2520
            Width           =   1005
         End
         Begin VB.Label Lbl_FilterQ1
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Q_1"
            BeginProperty Font
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   330
            Index           =   5
            Left            =   5790
            TabIndex        =   18
            Top             =   2520
            Width           =   1005
         End
         Begin VB.Label Lbl_FilterFreq2
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Freq_2"
            BeginProperty Font
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   330
            Index           =   5
            Left            =   7245
            TabIndex        =   19
            ToolTipText     =   "For PID it is Differential Gain"
            Top             =   2520
            Width           =   1155
         End
         Begin VB.Label Lbl_FilterGain
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Gain"
            BeginProperty Font
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   330
            Index           =   4
            Left            =   10395
            TabIndex        =   22
            ToolTipText     =   "Usually =1, for PID it is Proportional Gain"
            Top             =   2115
            Width           =   1005
         End
         Begin VB.Label Lbl_FilterQ2
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Q_2"
            BeginProperty Font
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   330
            Index           =   4
            Left            =   8910
            TabIndex        =   23
            ToolTipText     =   "For 2nd order filters controls peak at Frequency 2"
            Top             =   2115
            Width           =   1005
         End
         Begin VB.Label Lbl_FilterQ1
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Q_1"
            BeginProperty Font
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   330
            Index           =   4
            Left            =   5790
            TabIndex        =   24
            Top             =   2115
            Width           =   1005
         End
         Begin VB.Label Lbl_FilterFreq2
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Freq_2"
            BeginProperty Font
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   330
            Index           =   4
            Left            =   7245
            TabIndex        =   25
            ToolTipText     =   "For PID it is Differential Gain"
            Top             =   2115
            Width           =   1155
         End
         Begin VB.Label Lbl_FilterFreq1
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Freq_1"
            BeginProperty Font
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   330
            Index           =   4
            Left            =   4140
            TabIndex        =   26
            ToolTipText     =   "For PID it is Integral Gain"
            Top             =   2115
            Width           =   1150
         End
         Begin VB.Label Lbl_FilterType
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Filter Type"
            BeginProperty Font
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   330
            Index           =   4
            Left            =   630
            TabIndex        =   27
            Top             =   2115
            Width           =   2850
         End
         Begin VB.Label Lbl_FilterGain
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Gain"
            BeginProperty Font
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   330
            Index           =   3
            Left            =   10395
            TabIndex        =   28
            ToolTipText     =   "Usually =1, for PID it is Proportional Gain"
            Top             =   1710
            Width           =   1005
         End
         Begin VB.Label Lbl_FilterQ2
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Q_2"
            BeginProperty Font
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   330
            Index           =   3
            Left            =   8910
            TabIndex        =   29
            ToolTipText     =   "For 2nd order filters controls peak at Frequency 2"
            Top             =   1710
            Width           =   1005
         End
         Begin VB.Label Lbl_FilterQ1
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Q_1"
            BeginProperty Font
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   330
            Index           =   3
            Left            =   5790
            TabIndex        =   30
            Top             =   1710
            Width           =   1005
         End
         Begin VB.Label Lbl_FilterFreq2
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Freq_2"
            BeginProperty Font
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   330
            Index           =   3
            Left            =   7245
            TabIndex        =   31
            ToolTipText     =   "For PID it is Differential Gain"
            Top             =   1710
            Width           =   1155
         End
         Begin VB.Label Lbl_FilterFreq1
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Freq_1"
            BeginProperty Font
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   330
            Index           =   3
            Left            =   4140
            TabIndex        =   32
            ToolTipText     =   "For PID it is Integral Gain"
            Top             =   1710
            Width           =   1150
         End
         Begin VB.Label Lbl_FilterType
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Filter Type"
            BeginProperty Font
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   330
            Index           =   3
            Left            =   630
            TabIndex        =   33
            Top             =   1710
            Width           =   2850
         End
         Begin VB.Label Lbl_FilterGain
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Gain"
            BeginProperty Font
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   330
            Index           =   2
            Left            =   10395
            TabIndex        =   34
            ToolTipText     =   "Usually =1, for PID it is Proportional Gain"
            Top             =   1305
            Width           =   1005
         End
         Begin VB.Label Lbl_FilterQ2
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Q_2"
            BeginProperty Font
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   330
            Index           =   2
            Left            =   8910
            TabIndex        =   35
            ToolTipText     =   "For 2nd order filters controls peak at Frequency 2"
            Top             =   1305
            Width           =   1005
         End
         Begin VB.Label Lbl_FilterQ1
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Q_1"
            BeginProperty Font
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   330
            Index           =   2
            Left            =   5790
            TabIndex        =   36
            Top             =   1305
            Width           =   1005
         End
         Begin VB.Label Lbl_FilterFreq2
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Freq_2"
            BeginProperty Font
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   330
            Index           =   2
            Left            =   7245
            TabIndex        =   37
            ToolTipText     =   "For PID it is Differential Gain"
            Top             =   1305
            Width           =   1155
         End
         Begin VB.Label Lbl_FilterFreq1
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Freq_1"
            BeginProperty Font
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   330
            Index           =   2
            Left            =   4140
            TabIndex        =   38
            ToolTipText     =   "For PID it is Integral Gain"
            Top             =   1305
            Width           =   1150
         End
         Begin VB.Label Lbl_FilterType
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Filter Type"
            BeginProperty Font
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   330
            Index           =   2
            Left            =   630
            TabIndex        =   39
            Top             =   1305
            Width           =   2850
         End
         Begin VB.Label Lbl_FilterGain
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Gain"
            BeginProperty Font
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   330
            Index           =   1
            Left            =   10395
            TabIndex        =   40
            ToolTipText     =   "Usually =1, for PID it is Proportional Gain"
            Top             =   900
            Width           =   1005
         End
         Begin VB.Label Lbl_FilterQ2
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Q_2"
            BeginProperty Font
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   330
            Index           =   1
            Left            =   8910
            TabIndex        =   41
            ToolTipText     =   "For 2nd order filters controls peak at Frequency 2"
            Top             =   900
            Width           =   1005
         End
         Begin VB.Label Lbl_FilterQ1
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Q_1"
            BeginProperty Font
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   330
            Index           =   1
            Left            =   5790
            TabIndex        =   42
            Top             =   900
            Width           =   1005
         End
         Begin VB.Label Lbl_FilterFreq2
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Freq_2"
            BeginProperty Font
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   330
            Index           =   1
            Left            =   7245
            TabIndex        =   43
            ToolTipText     =   "For PID it is Differential Gain"
            Top             =   900
            Width           =   1155
         End
         Begin VB.Label Lbl_FilterType
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Filter Type"
            BeginProperty Font
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   330
            Index           =   1
            Left            =   630
            TabIndex        =   45
            Top             =   900
            Width           =   2850
         End
         Begin VB.Label Lbl_FiltType
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "F#    Filter Type"
            BeginProperty Font
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   90
            TabIndex        =   52
            Top             =   135
            Width           =   2850
         End
         Begin VB.Label Lbl_FiltGain
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Filter Gain"
            BeginProperty Font
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   1
            Left            =   10260
            TabIndex        =   53
            Top             =   135
            Width           =   1185
         End
         Begin VB.Label Lbl_F2_Q2
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Frequency 2    Q-factor 2"
            BeginProperty Font
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   7110
            TabIndex        =   54
            Top             =   135
            Width           =   2895
         End
         Begin VB.Label Lbl_F1_Q1
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Frequency 1    Q-factor 1"
            BeginProperty Font
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   4050
            TabIndex        =   55
            Top             =   135
            Width           =   3210
         End
         Begin VB.Label Lbl_FilterFreq1
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Freq_1"
            BeginProperty Font
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   330
            Index           =   5
            Left            =   4140
            TabIndex        =   20
            ToolTipText     =   "For PID it is Integral Gain"
            Top             =   2520
            Width           =   1150
         End
         Begin VB.Label Lbl_FilterType
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Filter Type"
            BeginProperty Font
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   330
            Index           =   5
            Left            =   630
            TabIndex        =   21
            Top             =   2520
            Width           =   2850
         End
         Begin VB.Label Lbl_FilterFreq1
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Freq_1"
            BeginProperty Font
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   330
            Index           =   1
            Left            =   4140
            TabIndex        =   44
            ToolTipText     =   "For PID it is Integral Gain"
            Top             =   900
            Width           =   1150
         End
      End
      Begin CWUIControlsLib.CWNumEdit cwNumAxisGain
         Height          =   360
         Left            =   1575
         TabIndex        =   65
         Tag             =   "gain00"
         ToolTipText     =   "Gain of the axis"
         Top             =   675
         Width           =   1050
         _Version        =   524288
         _ExtentX        =   1852
         _ExtentY        =   635
         _StockProps     =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}
            Name            =   "Arial"
            Size            =   10.5
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Reset_0         =   0   'False
         CompatibleVers_0=   524288
         NumEdit_0       =   1
         ClassName_1     =   "CCWNumEdit"
         opts_1          =   458814
         C[0]_1          =   12632319
         C[1]_1          =   0
         AccelTime_1     =   1
         BorderStyle_1   =   1
         ButtonPosition_1=   1
         TextAlignment_1 =   1
         format_1        =   2
         ClassName_2     =   "CCWFormat"
         Format_2        =   ".#0"
         scale_1         =   3
         ClassName_3     =   "CCWScale"
         opts_3          =   65536
         dMax_3          =   10
         discInterval_3  =   1
         ValueVarType_1  =   5
         Value_Val_1     =   -99
         IncValueVarType_1=   5
         IncValue_Val_1  =   0.1
         AccelIncVarType_1=   5
         AccelInc_Val_1  =   1
         RangeMinVarType_1=   5
         RangeMin_Val_1  =   -500
         RangeMaxVarType_1=   5
         RangeMax_Val_1  =   500
         ButtonStyle_1   =   0
         Bindings_1      =   4
         ClassName_4     =   "CCWBindingHolderArray"
         Editor_4        =   5
         ClassName_5     =   "CCWBindingHolderArrayEditor"
         Owner_5         =   1
      End
      Begin VB.Label Label1
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Gain"
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
         Height          =   330
         Index           =   0
         Left            =   990
         TabIndex        =   107
         ToolTipText     =   "Choose axis gain"
         Top             =   675
         Width           =   600
      End
      Begin VB.Label Label1
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "AXIS"
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
         Height          =   330
         Index           =   15
         Left            =   990
         TabIndex        =   56
         ToolTipText     =   "Choose axis to see/edit its filters"
         Top             =   270
         Width           =   600
      End
      Begin VB.Image ImagePictogram
         Height          =   735
         Index           =   13
         Left            =   135
         Picture         =   "frmFilters.frx":0BAA
         Stretch         =   -1  'True
         Top             =   270
         Width           =   735
      End
   End
   Begin CWUIControlsLib.CWNumEdit cwNumExperimentalFFscale
      Height          =   360
      Left            =   12060
      TabIndex        =   105
      Tag             =   "gain00"
      ToolTipText     =   "Experimentally found addtional scale for the FloorFF axis, default = 2.0"
      Top             =   3690
      Width           =   1365
      _Version        =   524288
      _ExtentX        =   2408
      _ExtentY        =   635
      _StockProps     =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}
         Name            =   "Arial"
         Size            =   10.5
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Reset_0         =   0   'False
      CompatibleVers_0=   524288
      NumEdit_0       =   1
      ClassName_1     =   "CCWNumEdit"
      opts_1          =   458814
      C[0]_1          =   16777215
      C[1]_1          =   0
      AccelTime_1     =   1
      BorderStyle_1   =   1
      ButtonPosition_1=   1
      TextAlignment_1 =   1
      format_1        =   2
      ClassName_2     =   "CCWFormat"
      Format_2        =   ".#0"
      scale_1         =   3
      ClassName_3     =   "CCWScale"
      opts_3          =   65536
      dMax_3          =   10
      discInterval_3  =   1
      ValueVarType_1  =   5
      Value_Val_1     =   2
      IncValueVarType_1=   5
      IncValue_Val_1  =   0.1
      AccelIncVarType_1=   5
      AccelInc_Val_1  =   1
      RangeMinVarType_1=   5
      RangeMin_Val_1  =   0.1
      RangeMaxVarType_1=   5
      RangeMax_Val_1  =   10
      ButtonStyle_1   =   0
      Bindings_1      =   4
      ClassName_4     =   "CCWBindingHolderArray"
      Editor_4        =   5
      ClassName_5     =   "CCWBindingHolderArrayEditor"
      Owner_5         =   1
   End
   Begin VB.Label Label1
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Experimental FF scale"
      BeginProperty Font
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   555
      Index           =   4
      Left            =   12105
      TabIndex        =   106
      ToolTipText     =   "Experimentally found addtional scale for the Floor FF axis"
      Top             =   3150
      Width           =   1275
   End
End
Attribute VB_Name = "frmFilters"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim tst As Long
    Dim retval As String
    Dim FilterPrevPhase As Single ' the previous phase of the filter TF at previous frequency, to detect jump in phase

    Const FILTER_PARAMS_TOTAL   As Long = 11 'filt type, 5 params, 5 coefficients
'#define PID_P_gain FilterGain
'#define PID_I_gain freq_1
'#define PID_D_gain freq_2
    Const MAX_FILTERS_IN_AXIS   As Long = 6
    '7 filters (0 to 6) in Axis chain, 10 parameters (0 to 5)-user, 6=a1,7=a2,8=b0,9=b1,0xA=b2
    Const FB_FILTERS_IN_AXIS As Long = MAX_FILTERS_IN_AXIS - 1 '0 to 5 for ELDAMP, the 5th filter in ElectroDamp is for stage FF adaptive algorithm
    Const FF_ADAPTIVE_FILT_NUM  As Long = MAX_FILTERS_IN_AXIS - 1 '#5
    Const MAX_ED_W_STAGE_FF_AXES_W_FILTERS As Long = 6 + 6 + 4 ' 6 position, 6 velocity, 4 feed forward
    Const MAX_ED_W_FLOOR_FF_AXES_W_FILTERS As Long = 6 + 6 + 3 ' 6 position, 6 velocity, 3 feed forward

    Dim OriginalFilterParamArray(MAX_FILTERS_IN_AXIS - 1, FILTER_PARAMS_TOTAL - 1) As Single ' (0-5, 0-9) read from controller filters
    Dim CHANGED_FilterParamArray(MAX_FILTERS_IN_AXIS - 1, FILTER_PARAMS_TOTAL - 1) As Single 'editing copy of filters
    Dim Filter_CHANGED(MAX_FILTERS_IN_AXIS - 1, USER_PARAM_NUMBER) As Boolean  'check change of 5 params AND FILTER TYPE
    Dim FrameFilterTop As Integer 'graphical Y coordinate
    Dim Filters_were_read As Boolean
    Dim PreviousEditedFilterNumber As Integer
    Const FiltCoeffFormat As String = "0.0#######"
    Const FiltParFormat As String = "###0.000#"
    Dim CopiedFilterCoefficients(5) As Single
    Dim FilterNumberToUpdate As Long  'filter number in the axis to update format and min-max values


Sub SetDefaultFiltParams() ' or it will assume all zeros and Filter_FT will be -100 dB and Axis_TF will be -500 dB
    Dim f_num
    Dim f_par As Long
    For f_num = 0 To MAX_FILTERS_IN_AXIS - 1
        OriginalFilterParamArray(f_num, 0) = 0 'filter type = NO_FILTER
        OriginalFilterParamArray(f_num, 1) = 1  'gain
        OriginalFilterParamArray(f_num, 2) = 1  'F1
        OriginalFilterParamArray(f_num, 3) = 1  'F2
        OriginalFilterParamArray(f_num, 4) = 0  'Q1
        OriginalFilterParamArray(f_num, 5) = 0  'Q2
        OriginalFilterParamArray(f_num, 6) = 0  'a1
        OriginalFilterParamArray(f_num, 7) = 0  'a2
        OriginalFilterParamArray(f_num, 8) = 1  'b0
        OriginalFilterParamArray(f_num, 9) = 0  'b1
        OriginalFilterParamArray(f_num, 10) = 0  'b2

        For f_par = 0 To FILTER_PARAMS_TOTAL - 1
            CHANGED_FilterParamArray(f_num, f_par) = OriginalFilterParamArray(f_num, f_par)
        Next f_par
    Next f_num
End Sub

Private Sub CheckPneumFilters_Click()
    If CheckPneumFilters.Value = Checked Then
        PneumFilterChain = True
        cwAuxDampGain.Visible = True
        LblDampGain.Visible = True
        ChkAuxDamping.Visible = True
    Else
        PneumFilterChain = False
        cwAuxDampGain.Visible = False
        LblDampGain.Visible = False
        ChkAuxDamping.Visible = False
    End If
'    Call CalculateAxis_TF
'    Call Analyzer.Calc_prediction
End Sub

Private Sub ChkAuxDamping_Click()
    Call Analyzer.Calc_prediction
End Sub

Private Sub ChkEnablePRecisionLogging_Click()
On Error Resume Next
    If ChkEnablePRecisionLogging = Checked Then
        EnablePrecisionLog = True
        PrecisionLogFile = App.Path & "\PrecisionLog.csv"
        ' Optional: write CSV header
        Dim f As Integer
        f = FreeFile
        Open PrecisionLogFile For Output As #f
        Print #f, "Operation,Input1,Input2,DoubleResult,SingleResult,Error PPM"
        Close #f
    Else
        EnablePrecisionLog = False
    End If
End Sub

Private Sub ChkRealTimeUpdateTF_Click()
    'Call CalculateAxis_TF
    Call Analyzer.Calc_prediction
End Sub

Private Sub ChkUseDoubleCalcWhenFalse_Click() ' IK20260820
    If (ChkUseDoubleCalcWhenFalse.Value = Unchecked) Then
        ChkUseDoubleCalcWhenFalse.BackColor = &HC0FFC0
        ChkUseDoubleCalcWhenFalse.Caption = "64-bit double"
        UseSingleCalculations = False
    Else 'checked
        ChkUseDoubleCalcWhenFalse.BackColor = &HC0E0FF
        ChkUseDoubleCalcWhenFalse.Caption = "32-bit float"
        UseSingleCalculations = True
    End If
    Call CalculateAxis_TF
    Call Analyzer.Calc_prediction
End Sub

Private Sub cwAuxDampGain_ValueChanged(Value As Variant, PreviousValue As Variant, ByVal OutOfRange As Boolean)
    If PneumFilterChain = True Then
        Call Analyzer.Calc_prediction
    End If
End Sub

Private Sub cwNumExperimentalFFscale_ValueChanged(Value As Variant, PreviousValue As Variant, ByVal OutOfRange As Boolean)
    Call CalculateAxis_TF
    Call Analyzer.Calc_prediction
End Sub

Private Sub Form_Load()
    If (AirIsoPresent = True) Then
        FramePneumatic.Enabled = True
        CheckPneumFilters.Value = Checked
    Else
        FramePneumatic.Enabled = False
        CheckPneumFilters.Value = Unchecked
    End If
    If ComboFilterAxis.ListCount >= 22 Then
        If (HasFloorFF = True) Then
            ComboFilterAxis.List(12) = "ZfloorFF"
            ComboFilterAxis.List(13) = "XfloorFF"
            ComboFilterAxis.List(14) = "YfloorFF"
            ComboFilterAxis.List(15) = "UnusedFF"
            ComboFilterAxis.List(16) = "UnusedV1"
            ComboFilterAxis.List(17) = "UnusedV2"
            ComboFilterAxis.List(18) = "UnusedV3"
            ComboFilterAxis.List(19) = "UnusedV4"
            ComboFilterAxis.List(20) = "UnusedBal"
            ComboFilterAxis.List(21) = "Aux Filters"
        Else
            ComboFilterAxis.List(12) = "FF X acc"
            ComboFilterAxis.List(13) = "FF Y acc"
            ComboFilterAxis.List(14) = "FF X pos"
            ComboFilterAxis.List(15) = "FF Y pos"
            ComboFilterAxis.List(16) = "pressure V1"
            ComboFilterAxis.List(17) = "pressure V2"
            ComboFilterAxis.List(18) = "pressure V3"
            ComboFilterAxis.List(19) = "pressure V4"
            ComboFilterAxis.List(20) = "pressure Bal"
            ComboFilterAxis.List(21) = "Aux Filters"
        End If
    Else
    End If
    Call SetDefaultFiltParams
    ComboFilterAxis.Text = ComboFilterAxis.List(ComboFilterAxis.TopIndex) '"Z1"
    ComboFilterTYPE.Text = ComboFilterTYPE.List(ComboFilterTYPE.TopIndex) '"NO_FILTER"
    FrameFilterTop = FrameFilter.Top ' coordiate of the top, because it will move up-down
    Call fill_Freq_array
    If FormFiltersVisible = True Then Me.Show
End Sub


Private Sub cmdLoadDefaults_Click(Index As Integer)
    cmd_sent = False
    Call Analyzer.GetSend("dflt", True)      'default cmd
End Sub

Private Sub cwNumAxisGain_ValueChanged(Value As Variant, PreviousValue As Variant, ByVal OutOfRange As Boolean)
    On Error Resume Next
    Dim gain_str As String
    Dim AxisIndex As Long
    AxisIndex = ComboFilterAxis.ListIndex
    If (AxisIndex) > MAX_ED_W_STAGE_FF_AXES_W_FILTERS - 1 Then 'error, there are only 21 axis gains for Pos and Vel, gain#$, #=0,1,2,3, $=0-A
        cwNumAxisGain.BackColorText = &HC0C0FF
        cwNumAxisGain.Enabled = False
        Exit Sub
    Else
        cwNumAxisGain.BackColorText = vbWhite
        cwNumAxisGain.Enabled = True
    End If
    If (AxisIndex < 12) Then
        gain_str = "gain0" & ExtendedHEX(AxisIndex) & "=" & cwNumAxisGain.Value
    Else
        If (AxisIndex >= 12 And AxisIndex <= 15) Then
            If (HasFloorFF = True) Then ' Floor FF: floorZ=0; FloorX=1; FloorY=2; UnUsed=3; Linear Motor adaptive rate=4
                gain_str = "g_fff" & ExtendedHEX(AxisIndex - 12) & "wn=" & cwNumAxisGain.Value
                response = Analyzer.GetSend(gain_str, True)
            Else 'IK20251208 the is no single gain for StageFF, gains are in the matrix.
                ' Stage FF: Xacc=0; Yacc=1; Xpos=2; Ypos=3; Linear Motor adaptive rate=4
                ' Parameter "$": feed forward output channel: 0=MassCenterPos; 1=tXpos; 2=tYpos; 3=Xmot, 4=Ymot, 5=tZmot, 6=tXmot, 7=tYmot
                ' to show Stage FF TF, just use gain = 1.0
                ' create FAKE response with gain=1
                gain_str = "g_ff0" & ExtendedHEX(AxisIndex - 12) & "wn=" & cwNumAxisGain.Value
                response = "g_ff" + CStr(AxisIndex) + "0=1.00"
            End If

        End If
    End If
    'response = Analyzer.GetSend(gain_str, True)
    If (ChkRealTimeUpdateTF.Value = Checked) Then
        Call CalculateAxis_TF
        Call Analyzer.Calc_prediction
    End If
End Sub

Private Sub ChkShowAxisTF_Click()
    If ChkShowAxisTF.Value = Checked Then
        Call CalculateAxis_TF
        Analyzer.CWGraph.Plots(Axis_TF_Num).Visible = True
        Analyzer.CWGraphPhase.Plots(Axis_TF_Num).Visible = True
    Else
        Analyzer.CWGraph.Plots(Axis_TF_Num).Visible = False
        Analyzer.CWGraphPhase.Plots(Axis_TF_Num).Visible = False
    End If
End Sub

Private Sub ChkShowOneFilterTF_Click()
    If ChkShowOneFilterTF.Value = Checked Then
        Analyzer.CWGraph.Plots(OneFilterTF_Num).Visible = True
        Analyzer.CWGraphPhase.Plots(OneFilterTF_Num).Visible = True
    Else
        Analyzer.CWGraph.Plots(OneFilterTF_Num).Visible = False
        Analyzer.CWGraphPhase.Plots(OneFilterTF_Num).Visible = False
    End If
End Sub

Sub fill_Freq_array()
    Dim i As Long
    Dim delta_freq As Double
    Dim freq_values_per_decade As Single
    Freq_points(0) = 0.1
    freq_values_per_decade = TEST_LTF_ARRAY_LENGTH / 4
    delta_freq = 10 ^ (1 / freq_values_per_decade)
    For i = 1 To TEST_LTF_ARRAY_LENGTH
       Freq_points(i) = Freq_points(i - 1) * delta_freq
    Next i
End Sub

'Function calculate_test_transfer_function(frequency As Single) As Double
'    Dim displacement_x As Double
'    Dim displacement_u As Double
'    Dim resonance As Single
'    Dim damping As Single
'    resonance = 2.5
'    damping = 0.25
'    displacement_x = 1 + (2 * damping * frequency / resonance) ^ 2
'       displacement_u = Abs((1 - (frequency ^ 2 / resonance ^ 2) ^ 2 + (2 * damping * frequency / resonance) ^ 2))
'       calculate_test_transfer_function = 5 * Log((displacement_x / displacement_u) ^ 0.5)
'
'End Function
'Function calculate_test_phase(frequency As Single) As Double
'    calculate_test_phase = (-180 + 2 * frequency)
'    calculate_test_phase = ((calculate_test_phase + 180) Mod 360) - 180
'End Function

Sub CalculateFilterTF(Filt_num As Long)
    Dim i As Long
    Dim Mag_Phase As tFreqHz_MagTimes_PhaseDeg ''temp structure
    Dim ThisFilterCoeff As float_iir
    Dim freq_sum_for_test As Single
    Dim prev_freq As Single
    Dim t_float As Double
    Call set_filt_params_from_user_input(Filt_num)
    ThisFilterCoeff.ftyp = CHANGED_FilterParamArray(Filt_num, 0) 'IK 20210425 fixed, need this or ftyp is always 0

    ThisFilterCoeff.b2 = CHANGED_FilterParamArray(Filt_num, 6)
    ThisFilterCoeff.b1 = CHANGED_FilterParamArray(Filt_num, 7)
    ThisFilterCoeff.b0 = CHANGED_FilterParamArray(Filt_num, 8)
    ThisFilterCoeff.a2 = CHANGED_FilterParamArray(Filt_num, 9)
    ThisFilterCoeff.a1 = CHANGED_FilterParamArray(Filt_num, 10)

'check if frequency array is filled, not zero
    freq_sum_for_test = 0
    For i = 0 To TEST_LTF_ARRAY_LENGTH  '0 to 200
        If Freq_points(i) > 0 Then 'prepare test array for axis OLTF calculation
            prev_freq = Freq_points(i)         'prepare test array for axis OLTF calculation
            Freq_points(i) = prev_freq
        Else 'Freq_points are empty, maybe Freq_Reference is filled?
            If Freq_Reference(i) > 0 Then
                prev_freq = Freq_Reference(i) ' update array
                Freq_points(i) = prev_freq
            Else
'                Freq_points(i) = prev_freq ' fill up to the end with the last valid frequency
            End If
        End If
        freq_sum_for_test = freq_sum_for_test + prev_freq
    Next i
    i = i - 1
    If freq_sum_for_test = 0 Then ' array is not filled
        Call fill_Freq_array
    End If

 ' calculate filter TF
    For i = TEST_LTF_ARRAY_LENGTH - 1 To 0 Step -1 'IK 20210425 fixed, starts from lower frequency, 0.1 Hz; was starting from index 0 which is high freq
       Mag_Phase = Mag_Ph_vs_Freq(ThisFilterCoeff, CDbl(Freq_points(i)), LOOP_FREQUENCY)
       OneFilter_Mag(i) = 20 * Log(Mag_Phase.mag) / Log(10)
       t_float = Mag_Phase.Phase ' returns between -90 and +90
       If i = TEST_LTF_ARRAY_LENGTH - 1 Then
           FilterPrevPhase = t_float 'IK20210614 was if i=0 - when changed search order, forgot to set initial prev_phase. That caused some phase plots to jump 180 deg.
       End If
'       If ThisFilterCoeff.ftyp = Position Then '7 = LPF 2nd order, need to reverse phase because it calculates phase 180 at low freq IK 20210425 fixed LP2 2nd order ThisFilterCoeff.ftyp = LPF_2ndOr_wPk Or
'          t_float = t_float - 180
'       End If
'        If (FilterNumberInChain = Filt_num) Then
'            t_phase(i) = t_float
'            If (Freq_points(i) = 8) Then
'                t1 = t_float ' test trap for break point
'            End If
'            If Freq_points(i) = 9 Then
'                t2 = t_float ' test trap for break point
'            End If
'            If Freq_points(i) = 10 Then
'                t3 = t_float ' test trap for break point
'                ' test trap for break point
'            End If
'        End If
       OneFilter_Phase(i) = Phase_Filter_Expand_plus_minus_180(t_float) ' returns between -180 and +180
       'OneFilter_Phase(i) = (t_float) ' returns between -180 and +180
    Next i
    If ((FilterNumberInChain = Filt_num) And (ChkShowOneFilterTF.Value = Checked)) Then
        Call Analyzer.CWGraph.Plots(OneFilterTF_Num).PlotXvsY(Freq_points, OneFilter_Mag)
        Call Analyzer.CWGraphPhase.Plots(OneFilterTF_Num).PlotXvsY(Freq_points, OneFilter_Phase)
    End If
End Sub

Function Phase_Filter_Expand_plus_minus_180(phase_in As Double)
    Dim phase_diff As Single
    ' cleaning phase, which might jump because ATAN returns between -90 and +90
    ' example: prev phase is still +87, but this phase has jumped to -88
    ' example: prev phase is still -87, but this phase has jumped to +88
    ' check starts from lower frequency, 0.1 Hz

    If (FilterPrevPhase >= 180) Then ' > 179.9999999
        FilterPrevPhase = FilterPrevPhase - 180 '
    Else
        If (FilterPrevPhase <= -180) Then '<= -179.9999999
            FilterPrevPhase = FilterPrevPhase + 180 '
        End If
    End If

    phase_diff = phase_in - FilterPrevPhase
    'ex: phase_diff=-136.0: prevPhase=+46.5,phase_in=-89.5
    'ex: phase_diff=+175.0: prevPhase = -87,phase_in = +88
    If (phase_diff >= 179.9999999) Then
        phase_in = phase_in - 180 '
    Else
        If (phase_diff >= 90) Then 'ex: phase_diff=+175.0: prevPhase = -87,phase_in = +88
            phase_in = phase_in - 180 'becomes +88-180=83
        Else
           If (phase_diff <= -179.9999999) Then
                phase_in = phase_in + 180 '
            Else
                If (phase_diff < -90) Then 'ex: phase_diff=-136.0: prevPhase=+46.5,phase_in=-89.5
                    phase_in = phase_in + 180 'becomes -89.5+180=+90.5
                Else
                    phase_in = phase_in
                End If
            End If
        End If
    End If
    FilterPrevPhase = phase_in
    Phase_Filter_Expand_plus_minus_180 = phase_in
End Function

Function Phase_Limit_TF_plus_minus_180(phase_in As Single)
    Dim phase_diff As Single
    ' cleaning phase, which might jump because sum of all filters can go beyond 360
    ' example: prev phase is still +174, but this phase has jumped to -188
    If (phase_in > 360) Then
        phase_in = phase_in - 360 '
    Else
        If (phase_in > 180) Then
            phase_in = phase_in - 360
        Else
           If (phase_in < -360) Then
                phase_in = phase_in + 360 '
            Else
                If (phase_in < -180) Then
                    phase_in = phase_in + 360
                Else
                    phase_in = phase_in
                End If
            End If
        End If
    End If
    'FilterPrevPhase = phase_in
    Phase_Limit_TF_plus_minus_180 = phase_in
End Function

Sub set_filt_params_from_user_input(Filt_Number As Long) 'FilterParams As float_iir)
    Dim FilterParams As float_iir
    Dim param As Long
    Dim testVar As Single
    FilterParams.ftyp = CHANGED_FilterParamArray(Filt_Number, 0) 'ComboFilterTYPE.ListIndex
    FilterParams.par(FilterGain) = CHANGED_FilterParamArray(Filt_Number, 1)
    FilterParams.par(freq_1) = CHANGED_FilterParamArray(Filt_Number, 2)
    FilterParams.par(freq_2) = CHANGED_FilterParamArray(Filt_Number, 3)
    FilterParams.par(Q_f_1) = CHANGED_FilterParamArray(Filt_Number, 4)
    FilterParams.par(Q_f_2) = CHANGED_FilterParamArray(Filt_Number, 5)
    For param = 0 To USER_PARAM_NUMBER - 1
        If FilterParams.par(param) <> 0 Then testVar = testVar + 1
    Next param
    If (testVar = 0) Then
        Exit Sub 'all params are zero, not initiated yet. it will cause 'dived by zero' exception
    End If
    If ChkRealTimeUpdateTF.Value = Checked Then
        Call set_float_iir(FilterParams)  'this recalculates a, b coefficients
        ' update coefficients so we can re-cal axis TF and re-plot
        CHANGED_FilterParamArray(Filt_Number, 6) = FilterParams.b2
        CHANGED_FilterParamArray(Filt_Number, 7) = FilterParams.b1
        CHANGED_FilterParamArray(Filt_Number, 8) = FilterParams.b0
        CHANGED_FilterParamArray(Filt_Number, 9) = FilterParams.a2
        CHANGED_FilterParamArray(Filt_Number, 10) = FilterParams.a1
    Else
    End If
    ' update screen
    If Filt_Number = FilterNumberInChain Then
        Lbl_PC_b2 = Format(CHANGED_FilterParamArray(Filt_Number, 6), FiltCoeffFormat) '"#0.0#######")  'FilterParams.b2
        Lbl_PC_b1 = Format(CHANGED_FilterParamArray(Filt_Number, 7), FiltCoeffFormat)  'FilterParams.b1
        Lbl_PC_b0 = Format(CHANGED_FilterParamArray(Filt_Number, 8), FiltCoeffFormat)  'FilterParams.b0
        Lbl_PC_a2 = Format(CHANGED_FilterParamArray(Filt_Number, 9), FiltCoeffFormat)  'FilterParams.a2
        Lbl_PC_a1 = Format(CHANGED_FilterParamArray(Filt_Number, 10), FiltCoeffFormat) 'FilterParams.a1
        Call ShowDifference(Filt_Number)
    End If
End Sub

Function Calc_Diff(Filt_Number As Long, coeff_num As Long) As String
    Dim diff As Double
    Dim PC_coeff As Double
    Dim Contr_coeff As Double
    PC_coeff = CHANGED_FilterParamArray(Filt_Number, coeff_num)
    Contr_coeff = OriginalFilterParamArray(Filt_Number, coeff_num)
    diff = PC_coeff - Contr_coeff
    If (PC_coeff <> 0) Then
        Calc_Diff = Format((1000000 * diff / PC_coeff), FiltCoeffFormat)
    Else
        Calc_Diff = "Coeff=0"
    End If
End Function

Sub ShowDifference(Filt_Number As Long)
   Lbl_b2_diff.Caption = Calc_Diff(Filt_Number, 6)
   Lbl_b1_diff.Caption = Calc_Diff(Filt_Number, 7)
   Lbl_b0_diff.Caption = Calc_Diff(Filt_Number, 8)
   Lbl_a2_diff.Caption = Calc_Diff(Filt_Number, 9)
   Lbl_a1_diff.Caption = Calc_Diff(Filt_Number, 10)
End Sub

Public Function PredictionCorrectionFloorFF(mag As Double)
    Dim retval As Double
    retval = mag
    If ((ComboFilterAxis.ListIndex >= 12) And (ComboFilterAxis.ListIndex <= 15) And (HasFloorFF = True)) Then
        retval = mag + 20 * Log(cwNumExperimentalFFscale.Value) / Log(10)  'add as dB experimental index for feed forward
    End If
    PredictionCorrectionFloorFF = retval
End Function

Public Sub CalculateAxis_TF(Optional reverse_Ph As Boolean = False)
    Dim i As Long
    Dim Filt_num As Long
    Dim Mag_Phase As tFreqHz_MagTimes_PhaseDeg ''temp structure
    Dim Axis_mag_phase As tFreqHz_MagTimes_PhaseDeg
    Dim ThisFilterCoeff As float_iir
    Dim TmpFiltPlot_Mag(TEST_LTF_ARRAY_LENGTH) As Single
    Dim TmpFiltPlot_Phase(TEST_LTF_ARRAY_LENGTH) As Single
    Dim NegGainCorrection As Integer
    Dim phase_reversal As Integer
    If reverse_Ph = True Then
        phase_reversal = 180
    Else
        phase_reversal = 0
    End If

'initial conditions
    For i = 0 To TEST_LTF_ARRAY_LENGTH ' 0 to 200
        TmpFiltPlot_Phase(i) = 0
        If cwNumAxisGain.Value <> -99 Then 'means it was updated
            Dim tmp_gain As Single
            If (cwNumAxisGain.Value < 0) Then
                tmp_gain = -cwNumAxisGain.Value
                NegGainCorrection = -180
            ElseIf (cwNumAxisGain.Value = 0) Then
                tmp_gain = 0.01 ' -80 db, but positive for log()
                cwNumAxisGain.Value = tmp_gain ' -80 db, but positive for log()
            Else
                tmp_gain = cwNumAxisGain.Value
                NegGainCorrection = 0
            End If
            TmpFiltPlot_Mag(i) = 20 * Log(tmp_gain) / Log(10) 'convert Axis gain to dB
        Else
            TmpFiltPlot_Mag(i) = 0 ' gain not updated: assume gain of 1 = 0 dB
        End If
    Next i

  If (PneumFilterChain <> True) Then
    Dim AxisEnabledFiltersNumber As Integer
    If (ComboFilterAxis.ListIndex >= 0 And ComboFilterAxis.ListIndex <= 5) Then 'Xpos, Ypos, tZpos, Zpos, tXpos, tYpos
        AxisEnabledFiltersNumber = 4
    ElseIf (ComboFilterAxis.ListIndex >= 6 And ComboFilterAxis.ListIndex <= 11) Then 'Xvel, Yvel, tZvel, Zvel, tXvel, tYvel
        AxisEnabledFiltersNumber = FB_FILTERS_IN_AXIS
    ElseIf (ComboFilterAxis.ListIndex >= 12 And ComboFilterAxis.ListIndex <= 15) Then 'FF, FF, FF, FF,
        If (HasFloorFF = True) Then
            AxisEnabledFiltersNumber = MAX_FILTERS_IN_AXIS - 1 ' ElDamp Floor FF has 6 filters in chain
        Else
            AxisEnabledFiltersNumber = 2 ' ElDamp Stage FF has complicated filters; 2 filters in chain go to the motors
            '/************************************************************************
            '*************************************************************************
            '.                               /--> motion detect: (abs(FFi1.Fout) > threshold)?
            '.                              /        LPF 1st 1 Hz
            '.                             /-------> FFp1=filt[3] -----> * pneum FF gains ---> pneumatic axes --> valves
            '.input-->opt DC block        /--> LPF>50Hz noise kill (for motors)
            '.ACC_FF_in -> FFi1=filt[0] -/--> FFi2=filt[1] ----> FF output to motors----+-> * FF gains for Vel axes --> Motors
            '.     \----------------------> FFjk=filt[2] ----> jerk added to FF out----/
            '.      \input----------------> HPF~10Hz detect accel fronts=jerk:
            '.       \
            '.        \input--> FFg1=filt[4]-----------> FFg2=filt[5] --> Velocity emulation from Acc to compare with geo signals during adaptive gain adjustment
            '.                  LPF_2ndOr_wPk,0.5,1.0 -> LEAD_2ndOr_wPk,5.5,2.0,1.0,0.5,5.0
            '************************************************************************/
        End If
    End If
' calculation "horizontally" NOT changing Filt_num untill all freq are calculated
    For Filt_num = 0 To AxisEnabledFiltersNumber '0 to 4 for velocity, the 5th filter in ElectroDamp is for stage FF adaptive algorithm
        If ChkRealTimeUpdateTF.Value = Checked Then
            Call set_filt_params_from_user_input(Filt_num)
        End If

        Call CalculateFilterTF(Filt_num) ' it fills       OneFilter_Mag(200),   OneFilter_Phase(200)
        For i = 0 To TEST_LTF_ARRAY_LENGTH ' 0 to 200
'            If i = 126 Then ' for break point
'             str_Help = Freq_points(i)
'             str_MouseX = TMpFiltPlot_Phase(i)
'            End If
'            If i = 127 Then ' for break point
'             str_MouseY = TMpFiltPlot_Phase(i)
'            End If
            TmpFiltPlot_Mag(i) = TmpFiltPlot_Mag(i) + OneFilter_Mag(i) ' dB added
            TmpFiltPlot_Phase(i) = TmpFiltPlot_Phase(i) + OneFilter_Phase(i) ' phase added
        Next i
    Next Filt_num

    For i = 0 To TEST_LTF_ARRAY_LENGTH ' 0 to 200
        Axis_Mag(i) = TmpFiltPlot_Mag(i)
        Axis_Phase(i) = Phase_Limit_TF_plus_minus_180(TmpFiltPlot_Phase(i) - phase_reversal + NegGainCorrection)
    Next i

    i = i - 1
    Call Analyzer.CWGraph.Plots(Axis_TF_Num).PlotXvsY(Freq_points, Axis_Mag)
    Call Analyzer.CWGraphPhase.Plots(Axis_TF_Num).PlotXvsY(Freq_points, Axis_Phase)

  Else ' PneumFilterChain == True
' IK20251112 Pneumatics has 2 signal paths: each has 2 fiters:
' regular path: [0] ==> LPF and [1] ==> PID
' electronic damping path: combination of [2]==> LPF + [3]==>HPF usually at the same frequency around 15-25 Hz, forming peak filter.
' It serves as "derivative" filter but without higher frequency gain increase which brings noise.
' user can choose to add or not this extra derivative component using "Electronic" or "Axilary" damping feature.
'//////////////////////////////////
' calculation "vertically"  changing Filt_num for each frequency untill all freq are calculated
''initial conditions
    Dim filt_mag_times(MAX_FILTERS_IN_AXIS) As Double
    Dim filt_mag_dB(MAX_FILTERS_IN_AXIS) As Double
    Dim filt_phase(MAX_FILTERS_IN_AXIS) As Double
    Dim AuxDamp_mag_times(2) As Double
    Dim AuxDamp_mag_dB(2) As Double
    Dim AuxDamp_phase(2) As Double

    Axis_mag_phase.mag = 1
    Axis_mag_phase.Phase = 0

    For i = 0 To TEST_LTF_ARRAY_LENGTH ' 0 to 200
''initial conditions
        If cwNumAxisGain.Value <> -99 Then 'means it was updated
            Axis_mag_phase.mag = cwNumAxisGain.Value
        Else
            Axis_mag_phase.mag = 1
        End If
        Axis_mag_phase.Phase = 0

        If cwAuxDampGain.Value > 0 Then 'means it was updated
            AuxDamp_mag_times(0) = cwAuxDampGain.Value
        Else
            AuxDamp_mag_times(0) = 1
        End If
        AuxDamp_phase(0) = 0

        For Filt_num = 0 To FB_FILTERS_IN_AXIS 'IK20251112 0 to 5, but FW has 4 filters for SoftDock
            If ChkRealTimeUpdateTF.Value = Checked Then
                Call set_filt_params_from_user_input(Filt_num)
            End If
            ThisFilterCoeff.b2 = CHANGED_FilterParamArray(Filt_num, 6)
            ThisFilterCoeff.b1 = CHANGED_FilterParamArray(Filt_num, 7)
            ThisFilterCoeff.b0 = CHANGED_FilterParamArray(Filt_num, 8)
            ThisFilterCoeff.a2 = CHANGED_FilterParamArray(Filt_num, 9)
            ThisFilterCoeff.a1 = CHANGED_FilterParamArray(Filt_num, 10)

            Mag_Phase = Mag_Ph_vs_Freq(ThisFilterCoeff, CDbl(Freq_points(i)), LOOP_FREQUENCY)
            filt_mag_times(Filt_num) = Mag_Phase.mag
            filt_phase(Filt_num) = Mag_Phase.Phase
            filt_mag_dB(Filt_num) = 20 * Log(filt_mag_times(Filt_num)) / Log(10)
        Next Filt_num

        Axis_mag_phase.mag = Axis_mag_phase.mag * filt_mag_times(0)
        Axis_mag_phase.mag = Axis_mag_phase.mag * filt_mag_times(1)

        If (ChkAuxDamping.Value = Checked) Then
            AuxDamp_mag_times(0) = AuxDamp_mag_times(0) * filt_mag_times(2)
            AuxDamp_mag_times(1) = AuxDamp_mag_times(0) * filt_mag_times(3)
            AuxDamp_mag_dB(0) = 20 * Log(AuxDamp_mag_times(0)) / Log(10)
            AuxDamp_mag_dB(1) = 20 * Log(AuxDamp_mag_times(1)) / Log(10)
            AuxDamp_phase(0) = AuxDamp_phase(0) * filt_phase(2)
            AuxDamp_phase(1) = AuxDamp_phase(0) + filt_phase(3)

            Axis_mag_phase.mag = Axis_mag_phase.mag + AuxDamp_mag_times(1)
            'Axis_mag_phase.Phase = Axis_mag_phase.Phase + AuxDamp_phase(1) '??? add with coefficients reverse proportional to mag?
            Axis_mag_phase.Phase = Axis_mag_phase.Phase + Atn(AuxDamp_mag_times(1) / Axis_mag_phase.mag) * 180 / PI ' returns between -90 and +90
        End If

        '////////////////////////////// final
        Axis_Mag(i) = 20 * Log(Axis_mag_phase.mag) / Log(10)
        Axis_Phase(i) = Axis_mag_phase.Phase
    Next i
    i = i - 1
    Call Analyzer.CWGraph.Plots(Axis_TF_Num).PlotXvsY(Freq_points, Axis_Mag)
    Call Analyzer.CWGraphPhase.Plots(Axis_TF_Num).PlotXvsY(Freq_points, Axis_Phase)
  End If
End Sub


'============================================================================================================================
'''''''''''''''''''''''''''''''''''''' FILTER PARAMETERS ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'============================================================================================================================
Sub update_damp_controls()
    If (AirIsoPresent = True) Then
        If (CheckPneumFilters.Value = Checked) Then ' read aux damping gain
            If (ComboFilterAxis.ListIndex >= 3 And ComboFilterAxis.ListIndex <= 5) Then 'Zpos, tX pos, tY pos
                cwAuxDampGain.Enabled = True
                PneumFilterChain = True
                ChkAuxDamping.Enabled = True
            Else
                PneumFilterChain = False
                cwAuxDampGain.Enabled = False
                ChkAuxDamping.Enabled = False
            End If
        End If
    Else
        PneumFilterChain = False
        cwAuxDampGain.Enabled = False
        ChkAuxDamping.Enabled = False
    End If
End Sub

Function GetAxisGain(AxisIndex As Long) As Double
    Dim gain_str As String
    Dim sep_pos As Long
    If (AxisIndex) > MAX_ED_W_STAGE_FF_AXES_W_FILTERS - 1 Then
        cwNumAxisGain.BackColorText = &HC0C0FF
        cwNumAxisGain.Enabled = False
    Else
        cwNumAxisGain.BackColorText = vbWhite
        cwNumAxisGain.Enabled = True
        Call update_damp_controls ' pneumatic auxilary damping gains CONTROLS states, not the gains themselves
        If (CheckPneumFilters = Checked) Then ' read aux damping gain
            If (AxisIndex >= 3 And AxisIndex <= 5) Then 'Zpos, tX pos, tY pos
                ' reading gains with indexes [C]=Zdamping, [D]=tXdamping, [E]=tYdamping
                gain_str = "gain0" & ExtendedHEX(AxisIndex + 9) 'gain0C
                response = Analyzer.GetSend(gain_str, True)
                sep_pos = InStr(response, "=")
                If (sep_pos <> 0) Then
                    retval = Mid(response, sep_pos + 1)
                    sep_pos = InStr(retval, "//") ''if verbose response
                    If (sep_pos <> 0) Then
                        retval = Left(retval, sep_pos - 1)
                    End If
                    If IsNumber(retval) Then
                        cwAuxDampGain.Value = Val(retval)
                        cwAuxDampGain.BackColorText = vbWhite
                    Else
                        cwAuxDampGain.BackColorText = &HC0C0FF
                    End If
                End If
                ' get Auxilary Pneumatic damping loop status
                response = Analyzer.GetSend("loop_FBD", True) 'loop_FBD
                sep_pos = InStr(response, ">")
                If (sep_pos <> 0) Then
                    If (LCase(Mid(response, sep_pos + 1, 4)) = "acti") Then 'acti
                       ChkAuxDamping.Value = Checked
                    End If
                    If (LCase(Mid(response, sep_pos + 1, 4)) = "pass") Then 'pass
                       ChkAuxDamping.Value = Unchecked
                    End If
                End If
            Else ' not damping gains
'                PneumFilterChain = False
'                cwAuxDampGain.Enabled = False
'                LblDampGain.Visible = False
'                ChkAuxDamping.Enabled = False
            End If
        Else 'CheckPneumFilters = Unchecked
            PneumFilterChain = False
        End If
    End If
    If (AxisIndex) > 21 Then 'error, there are only 22 filter chains,
        Call Refresh_FilterParams ' but, refresh  filter chain
        Exit Function
    End If
    ' read FB or FF gains
    Init_sys_data = True
    'AxisIndex = AxisIndex ' depends on selection of axis from dropdown
    If (AxisIndex < 12) Then
        gain_str = "gain0" + ExtendedHEX(AxisIndex)
        response = Analyzer.GetSend(gain_str, True)
    Else
        If ((AxisIndex >= 12) And (AxisIndex <= 15)) Then ' four feedforward gains
            If (HasFloorFF = True) Then ' Floor FF: floorZ=0; FloorX=1; FloorY=2; UnUsed=3; Linear Motor adaptive rate=4
                gain_str = "g_fff" + CStr(AxisIndex - 12)
                response = Analyzer.GetSend(gain_str, True)
            Else 'IK20251208 the is no single gain for StageFF, gains are in the matrix.
                ' Stage FF: Xacc=0; Yacc=1; Xpos=2; Ypos=3; Linear Motor adaptive rate=4
                ' Parameter "$": feed forward output channel: 0=MassCenterPos; 1=tXpos; 2=tYpos; 3=Xmot, 4=Ymot, 5=tZmot, 6=tXmot, 7=tYmot
                ' to show Stage FF TF, just use gain = 1.0
                ' create FAKE response with gain=1
                response = "g_ff" + CStr(AxisIndex) + "0=1.00"
            End If
        Else 'other axes:
        'pressure V1
        'pressure V2
        'pressure V3
        'pressure V4
        'pressure Bal
        'Aux Filters
        ' create FAKE response with gain=1
        response = "g_ff" + CStr(AxisIndex) + "0=1.00"
        End If
    End If
    sep_pos = InStr(response, "=")
    If (sep_pos <> 0) Then
        retval = Mid(response, sep_pos + 1)
        sep_pos = InStr(retval, "//") ''if verbose response
        If (sep_pos <> 0) Then
            retval = Left(retval, sep_pos - 1)
        End If
        If IsNumber(retval) Then
            cwNumAxisGain.Value = Val(retval)
            cwNumAxisGain.BackColorText = vbWhite
            GetAxisGain = Val(retval)
        Else
            cwNumAxisGain.BackColorText = &HC0C0FF
            GetAxisGain = -9.99 ' wrong value
        End If
    End If
    Init_sys_data = False
End Function

'event handlers
Private Sub CmdFilterRefresh_Click()
    Dim gain_str As String
    Dim retval As String
    Dim sep_pos As Long
    Dim AxisIndex As Long
    Dim AxisGain As Double
    ' design-time set axes in ComboFilterAxis.List, and its index, IK 20251208:
    'X pos   = 0
    'Y pos   = 1
    'tZ pos  = 2
    'Z pos   = 3
    'tX pos  = 4
    'tY pos  = 5
    'X vel   = 6
    'Y vel   = 7
    'tZ vel  = 8
    'Z vel   = 9
    'tX vel  = 10
    'tY vel  = 11
    'FF X acc     = 12 <<< NAME changed to "ZfloorFF" during Form_load() if use floor feed forward
    'FF Y acc     = 13 <<< NAME changed to "XfloorFF" during Form_load() if use floor feed forward
    'FF X pos     = 14 <<< NAME changed to "YfloorFF" during Form_load() if use floor feed forward
    'FF Y pos     = 15 <<< NAME changed to "UnusedFF" during Form_load() if use floor feed forward
    'pressure V1  = 16 <<< NAME changed to "UnusedV1" during Form_load() if use FFF. //-!- IK20251208 SHOULD BE "if no air isolators"
    'pressure V2  = 17 <<< NAME changed to "UnusedV1" during Form_load() if use FFF. //-!- IK20251208 SHOULD BE "if no air isolators"
    'pressure V3  = 18 <<< NAME changed to "UnusedV2" during Form_load() if use FFF. //-!- IK20251208 SHOULD BE "if no air isolators"
    'pressure V4  = 19 <<< NAME changed to "UnusedV3" during Form_load() if use FFF. //-!- IK20251208 SHOULD BE "if no air isolators"
    'pressure Bal = 20 <<< NAME changed to "UnusedBal" during Form_load() if use FFF. //-!- IK20251208 SHOULD BE "if no air isolators"
    'Aux Filters  = 21

    ' Read axis gain:

    'gain$#, $=0,1,2,3, "0=Axes, 1=Test_gains, 2=Test_phases, 3=Pressure, 4=PneumGainBoost"
    ' if $=0:
    ' #=0-F:  XpLM,  YpLM,  tZpLM, Zpos,  tXpos, tYpos, Xvel,  Yvel,  tZv ,  Zvel,  tXv,   tYv,   Zp_aux, tXp_aux, tYp_aux, diag balance
    ' first 3 gains for soft dock: pos to LinMotors, 3 pneumatic gains, 6 velocity gains, 3 auxilary pneum damping, last - pneumatic diagonal balancing
    ' cputs("Xp=0,Yp=1,tZp=2,Zp=3,tXp=4,tYp=5,Xv=6,Yv=7,tZv=8,Zv=9,tXv=A,tYv=B,Zd=C,tXd=D,tYd=E, bd=F\r\n");

    ' IK20251208 up to SysData.gain[11] old way, with FB axes
    ' /***************************************************************************************************
    ' for Stage or Floor FF, the syntax changes from"gain$#=XXX" to "g_FF#$@&=GGG.GGG" Set/get Feed Forward gain and attributes.
    'Parameter "#": feed forward input channel and adaptive rate: Xacc=0; Yacc=1; Xpos=2; Ypos=3; Linear Motor adaptive rate=4, Valve adaptive rate=5,
    '            AccErrorSLEW=6,PosErrorSLEW=7,MotionTreshold=8,MotionDetectDelay=9
    'Parameter "$": feed forward output channel: 0=MassCenterPos; 1=tXpos; 2=tYpos; 3=Xmot, 4=Ymot, 5=tZmot, 6=tXmot, 7=tYmot
    'Parameter "@": feed forward gain attribute: either "W" (Working gain) or "P" (Passive, not working)
    'Parameter "&": feed forward adaptive attribute: either "A" (Adaptive gain) or "N" (Not Adaptive)
    'Value "ggg.ggg"= can be in regular or scientific notation "+-m.mmE+-pp"; range: -100.0 ...+100
    'Get gain "g_FF$#?": returns gain and attributes: ?"g_FF@#WA=ggg.gg" or "g_FF@#WN=ggg.gg" or "g_FF@#PN=ggg.gg" or "g_FF@#PA=ggg.gg"
    'g_FF40xx= Linear Motor Adaptive Rate        SysData.StageFF_LinMotorLearningRate,
    'g_FF50xx= Valve Adaptive Rate               SysData.StageFF_ValveLearningRate,

    '/***************************************************************************************************
    '    "g_FF#$@&=GGG.GGG" Set / get Feed Forward gain and attributes.
    '    Parameter "#" for Electrodamp just 'F': Floor feed forward
    '    Parameter "$" : feed forward output channel : Z_fff = 0; X_fff = 1; Y_fff = 2;
    '    Parameter "@": feed forward gain attribute : either "W" (Working gain) or "P" (Passive, not working)
    '    Parameter "&" : feed forward adaptive attribute : either "A" (Adaptive gain) or "N" (Not Adaptive)
    '    Value "ggg.ggg" = can be in regular or scientific notation "+/-m.mmE+/-pp"; valid range : +/-100.0
    '
    '    Get gain "g_FFF#?" : returns gain and attributes : ? "g_FFF#WA=ggg.gg" or "g_FFF#WN=ggg.gg" or "g_FFF#PN=ggg.gg" or "g_FFF#PA=ggg.gg"
    '    g_FF40xx = Linear Motor Adaptive Rate       SysData.FloorFF_LearningRate,
    '
    '    g_fff0pn = 0      // FloorFF Z axis gain
    '    g_fff1pn = 0      // FloorFF X axis gain
    '    g_fff2pn = 0      // FloorFF Y axis gain
    '    g_ff40wa = 0      // Floor FF Adaptive rate
    '*********************************************************************************************************/

    'command g_fff#wn=+1.800, # =0 to 2 FFF
    AxisGain = GetAxisGain(ComboFilterAxis.ListIndex)
'    If (ComboFilterAxis.ListIndex) > MAX_ED_W_STAGE_FF_AXES_W_FILTERS - 1 Then
'        cwNumAxisGain.BackColorText = &HC0C0FF
'        cwNumAxisGain.Enabled = False
'    Else
'        cwNumAxisGain.BackColorText = vbWhite
'        cwNumAxisGain.Enabled = True
'        Call update_damp_controls ' pneumatic auxilary damping gains CONTROLS states, not the gains themselves
'        If (CheckPneumFilters = Checked) Then ' read aux damping gain
'            If (ComboFilterAxis.ListIndex >= 3 And ComboFilterAxis.ListIndex <= 5) Then 'Zpos, tX pos, tY pos
'                ' reading gains with indexes [C]=Zdamping, [D]=tXdamping, [E]=tYdamping
'                gain_str = "gain0" & ExtendedHEX(ComboFilterAxis.ListIndex + 9) 'gain0C
'                response = Analyzer.GetSend(gain_str, True)
'                sep_pos = InStr(response, "=")
'                If (sep_pos <> 0) Then
'                    retval = Mid(response, sep_pos + 1)
'                    sep_pos = InStr(retval, "//") ''if verbose response
'                    If (sep_pos <> 0) Then
'                        retval = Left(retval, sep_pos - 1)
'                    End If
'                    If IsNumber(retval) Then
'                        cwAuxDampGain.Value = Val(retval)
'                        cwAuxDampGain.BackColorText = vbWhite
'                    Else
'                        cwAuxDampGain.BackColorText = &HC0C0FF
'                    End If
'                End If
'                ' get Auxilary Pneumatic damping loop status
'                response = Analyzer.GetSend("loop_FBD", True) 'loop_FBD
'                sep_pos = InStr(response, ">")
'                If (sep_pos <> 0) Then
'                    If (LCase(Mid(response, sep_pos + 1, 4)) = "acti") Then 'acti
'                       ChkAuxDamping.Value = Checked
'                    End If
'                    If (LCase(Mid(response, sep_pos + 1, 4)) = "pass") Then 'pass
'                       ChkAuxDamping.Value = Unchecked
'                    End If
'                End If
'            Else ' not damping gains
''                PneumFilterChain = False
''                cwAuxDampGain.Enabled = False
''                LblDampGain.Visible = False
''                ChkAuxDamping.Enabled = False
'            End If
'        Else 'CheckPneumFilters = Unchecked
'            PneumFilterChain = False
'        End If
'    End If
'    If (ComboFilterAxis.ListIndex) > 21 Then 'error, there are only 22 filter chains,
'        Call Refresh_FilterParams ' but, refresh  filter chain
'        Exit Sub
'    End If
'    ' read FB gains
'    Init_sys_data = True
'    AxisIndex = ComboFilterAxis.ListIndex ' depends on selection of axis from dropdown
'    If (AxisIndex < 12) Then
'        gain_str = "gain0" + ExtendedHEX(AxisIndex)
'    Else
'        If ((AxisIndex >= 12) And (AxisIndex <= 15)) Then ' four feedforward gains
'            If (HasFloorFF = True) Then ' Floor FF: floorZ=0; FloorX=1; FloorY=2; UnUsed=3; Linear Motor adaptive rate=4
'                gain_str = "g_fff" + CStr(AxisIndex - 12)
'                response = Analyzer.GetSend(gain_str, True)
'            Else 'IK20251208 the is no single gain for StageFF, gains are in the matrix.
'                ' Stage FF: Xacc=0; Yacc=1; Xpos=2; Ypos=3; Linear Motor adaptive rate=4
'                ' Parameter "$": feed forward output channel: 0=MassCenterPos; 1=tXpos; 2=tYpos; 3=Xmot, 4=Ymot, 5=tZmot, 6=tXmot, 7=tYmot
'                ' to show Stage FF TF, just use gain = 1.0
'                ' create FAKE response with gain=1
'                response = "g_ff" + CStr(AxisIndex) + "0=1.00"
'            End If
'        Else 'other axes:
'        'pressure V1
'        'pressure V2
'        'pressure V3
'        'pressure V4
'        'pressure Bal
'        'Aux Filters
'        ' create FAKE response with gain=1
'        response = "g_ff" + CStr(AxisIndex) + "0=1.00"
'        End If
'    End If
'    sep_pos = InStr(response, "=")
'    If (sep_pos <> 0) Then
'        retval = Mid(response, sep_pos + 1)
'        sep_pos = InStr(retval, "//") ''if verbose response
'        If (sep_pos <> 0) Then
'            retval = Left(retval, sep_pos - 1)
'        End If
'        If IsNumber(retval) Then
'            cwNumAxisGain.Value = Val(retval)
'            cwNumAxisGain.BackColorText = vbWhite
'        Else
'            cwNumAxisGain.BackColorText = &HC0C0FF
'        End If
'    End If
'    Init_sys_data = False
    Call Refresh_FilterParams
End Sub


Private Sub cmdSaveParamsIntoFLASH_Click(Index As Integer)
    cmd_sent = False
    Call Analyzer.GetSend("save", True)      'save cmd
End Sub

Private Sub ComboFilterTYPE_Click()
    If Init_sys_data = True Then Exit Sub
    FilterTypeNumber = ComboFilterTYPE.ListIndex ' between 0 and 20
    CHANGED_FilterParamArray(FilterNumberInChain, 0) = FilterTypeNumber ' FilterTYPE is float in VB, but integer param

    Call ChangeFilterFontsAndMins(FilterNumberInChain, FilterTypeNumber)

    If (CHANGED_FilterParamArray(FilterNumberInChain, 0) <> OriginalFilterParamArray(FilterNumberInChain, 0)) Then
        ComboFilterTYPE.BackColor = &HFFFF80 ' cyan indicates change
        Lbl_FilterType(FilterNumberInChain).BackColor = &HC0C0FF
        Lbl_FilterType(FilterNumberInChain).Caption = ComboFilterTYPE.List(ComboFilterTYPE.ListIndex)
'        Filter_CHANGED(FilterNumberInChain, 0) = True
    Else
        ComboFilterTYPE.BackColor = &HFFFFFF ' white indicates NO change
        Lbl_FilterType(FilterNumberInChain).BackColor = vbWhite
'        Filter_CHANGED(FilterNumberInChain, 0) = False
    End If
    Call CheckIfFilterChanged(FilterNumberInChain)
    If ChkRealTimeUpdateTF.Value = Checked Then
        Call Analyzer.Calc_prediction
    End If
End Sub

Sub ChangeFilterFontsAndMins(FilterNumberInAxis As Long, FilterTypeNum As Long) ' IK20220224 Fuc!$#~% Putin invaded Ukraine
    Dim DirectCoeffFontSize, RegularFontSize As Single
    Dim i As Long
    DirectCoeffFontSize = 9.25
    RegularFontSize = 11.25
    If (FilterTypeNum = DIRECT_COEFF) Then ' smaller font to fit
        For i = 0 To 4 ' smaller font to fit
            cwNumFilterParam(i).Font.Size = DirectCoeffFontSize
            cwNumFilterParam(i).FormatString = FiltCoeffFormat '".######0" ' change numeric format to show 6 dec points for float coefficient
            cwNumFilterParam(i).Minimum = -5000 ' IK20220319 changed from -10 to filter coeff change range min to -5000 because a & b coeff can be in thousands
            cwNumFilterParam(i).IncDecValue = 0.001
        Next i
        Lbl_FilterFreq1(FilterNumberInAxis).Font.Size = DirectCoeffFontSize
        Lbl_FilterFreq2(FilterNumberInAxis).Font.Size = DirectCoeffFontSize
        Lbl_FilterQ1(FilterNumberInAxis).Font.Size = DirectCoeffFontSize
        Lbl_FilterQ2(FilterNumberInAxis).Font.Size = DirectCoeffFontSize
        Lbl_FilterGain(FilterNumberInAxis).Font.Size = DirectCoeffFontSize
         ' IK20260310 filter coefficients names
        cwNumFilterParam(1).ToolTipText = "b1 coeff" ' F1
        cwNumFilterParam(2).ToolTipText = "b2 coeff" ' F2
        cwNumFilterParam(3).ToolTipText = "a1 coeff" ' Q1
        cwNumFilterParam(4).ToolTipText = "a2 coeff" ' Q2
        cwNumFilterParam(0).ToolTipText = "b0 coeff" ' Gain
    Else
        For i = 0 To 4
            FilterNumberToUpdate = i ' pass extra parameter to cwNumFilterParam(i)_Changed event to avoid miss-formating direct coeff when neg values substituted with 0.001
            cwNumFilterParam(i).Font.Size = RegularFontSize
            cwNumFilterParam(i).FormatString = ".##0" ' change numeric format to show 3 dec points for filter param
            cwNumFilterParam(i).IncDecValue = 0.1       ' IK20260310
            If i > 0 Then cwNumFilterParam(i).Minimum = 0.001 ' filter param change range to +0.001,+5000 except gain -100...+100
        Next i
        Lbl_FilterFreq1(FilterNumberInAxis).Font.Size = RegularFontSize
        Lbl_FilterFreq2(FilterNumberInAxis).Font.Size = RegularFontSize
        Lbl_FilterQ1(FilterNumberInAxis).Font.Size = RegularFontSize
        Lbl_FilterQ2(FilterNumberInAxis).Font.Size = RegularFontSize
        Lbl_FilterGain(FilterNumberInAxis).Font.Size = RegularFontSize
         ' IK20260310 restore std filter parameter names
        If (FilterTypeNum = PID) Then
            cwNumFilterParam(1).ToolTipText = "Integral gain"
            cwNumFilterParam(3).ToolTipText = "Leaky - HPF frequency: larger = more leaky"
            cwNumFilterParam(2).ToolTipText = "Differential gain"
            cwNumFilterParam(4).ToolTipText = "Not used in PID"
            cwNumFilterParam(0).ToolTipText = "Proportional gain"
        Else
            cwNumFilterParam(1).ToolTipText = "Frequency 1"
            cwNumFilterParam(3).ToolTipText = "Q-factor 1"
            cwNumFilterParam(2).ToolTipText = "Frequency 2"
            cwNumFilterParam(4).ToolTipText = "Q-factor 2"
            cwNumFilterParam(0).ToolTipText = "Filter Gain"
        End If
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
'    FormTerminalVisible = False
'    If AnalyzerVisible = True Then
     Call Analyzer.ToggleFiltersControls 'changes FormFiltersVisible = not FormFiltersVisible
'    End If
End Sub

Private Sub LblFilter_Click(Index As Integer)
    If Init_sys_data = True Then Exit Sub
    FilterNumberInChain = Index
    FrameFilter.Top = FrameFilterTop + FilterNumberInChain * (LblFilter(1).Top - LblFilter(0).Top) '450
    FrameFilter.Visible = True
    Lbl_FilterType(PreviousEditedFilterNumber).Caption = ComboFilterTYPE.List(ComboFilterTYPE.ListIndex)
    Copy_Params_for_Edit (Index)
    CalculateFilterTF (FilterNumberInChain)
    PreviousEditedFilterNumber = Index
End Sub

Private Sub ComboFilterAxis_Click()
    If Init_sys_data = True Then Exit Sub   ' fires first
    FilterAxis = ComboFilterAxis.ListIndex
    Call update_damp_controls
End Sub

Private Sub cwNumFilterParam_ValueChanged(Index As Integer, Value As Variant, PreviousValue As Variant, ByVal OutOfRange As Boolean)
    If Init_sys_data = True Then Exit Sub
    CHANGED_FilterParamArray(FilterNumberInChain, Index + 1) = (Value) 'Index+1 because we skip FilterTYPE
    If (CHANGED_FilterParamArray(FilterNumberInChain, Index + 1) <> OriginalFilterParamArray(FilterNumberInChain, Index + 1)) Then
        cwNumFilterParam(Index).BackColorText = &HFFFF80 ' cyan indicates change
    Else
        cwNumFilterParam(Index).BackColorText = &HFFFFFF ' white indicates NO change
    End If
    If (FilterTypeNumber = DIRECT_COEFF) Then ' IK20220224 Fuc!$#~% Putin invaded Ukraine
        If Index = FilterGain Then Lbl_FilterFreq1(FilterNumberInChain).Caption = Format(Value, FiltCoeffFormat) '"0.#######0")
        If Index = freq_1 Then Lbl_FilterFreq1(FilterNumberInChain).Caption = Format(Value, FiltCoeffFormat)
        If Index = freq_2 Then
            If Value > 0 And Value < 0.002 Then
                Call MsgBox("response = 0", vbOKOnly, "too small response")
            End If
            Lbl_FilterFreq2(FilterNumberInChain).Caption = Format(Value, FiltCoeffFormat)
        End If
        If Index = Q_f_1 Then Lbl_FilterQ1(FilterNumberInChain).Caption = Format(Value, FiltCoeffFormat)
        If Index = Q_f_2 Then Lbl_FilterQ2(FilterNumberInChain).Caption = Format(Value, FiltCoeffFormat)
    Else ' normal precision
        If Index = FilterGain Then Lbl_FilterGain(FilterNumberInChain).Caption = Format(Value, FiltParFormat) '"###0.000"    'index=0
        If Index = freq_1 Then
            If Value <= 0 Then Value = 0.001 ' if zero, calculation might give "division by zero" error
            Lbl_FilterFreq1(FilterNumberInChain).Caption = Format(Value, FiltParFormat)
        End If
        If Index = freq_2 Then
            If Value <= 0 Then Value = 0.001 ' if zero, calculation might give "division by zero" error
            Lbl_FilterFreq2(FilterNumberInChain).Caption = Format(Value, FiltParFormat)
        End If
        If Index = Q_f_1 Then
            If Value <= 0 Then Value = 0.001 ' if zero, calculation might give "division by zero" error
            Lbl_FilterQ1(FilterNumberInChain).Caption = Format(Value, FiltParFormat)
        End If
        If Index = Q_f_2 Then
            If Value <= 0 Then Value = 0.001 ' if zero, calculation might give "division by zero" error
            Lbl_FilterQ2(FilterNumberInChain).Caption = Format(Value, FiltParFormat)
        End If
    End If
    CheckIfFilterChanged (FilterNumberInChain)
    If ChkRealTimeUpdateTF.Value = Checked Then
        Call Analyzer.Calc_prediction
    End If
End Sub

Private Sub CmdUpdateFilter_Click()
    Dim Fnum As Long
    Dim Fpar As Long
    Dim filter_param_str As String
    Dim response As String
    For Fnum = 0 To MAX_FILTERS_IN_AXIS - 1
        For Fpar = 0 To USER_PARAM_NUMBER 'check change of 5 params AND FILTER TYPE
          If (Opt_ALLfilters_or_ChangedOnly(1).Value = True) Then ' update only changed filters
            If (Filter_CHANGED(Fnum, Fpar) = True) Then 'flag for individual param
                filter_param_str = "fpar" & ExtendedHEX(FilterAxis) & CStr(Fnum) & CStr(Fpar)
                ' IK 20220130 the Str(CHANGED_FilterParamArray(Fnum, Fpar)) produces string with leading space, " 4", and firmware gives error because "fpar000= 4" is not recognized prior 20220130
                ' IK20250828 fixed the bug: value was rounding to integer: if value was 0.876, "fpar902 = 1.0" ' just appending numeric to string (with '&') converts it correctly
                filter_param_str = filter_param_str & "=" & Format(CHANGED_FilterParamArray(Fnum, Fpar), "0.00#######") ' Format produces "4"
                response = Analyzer.GetSend(filter_param_str, True)
            End If
          Else 'update ALL filters
            filter_param_str = "fpar" & ExtendedHEX(FilterAxis) & CStr(Fnum) & CStr(Fpar) ' coordinates of the filter
                filter_param_str = filter_param_str & "=" & CHANGED_FilterParamArray(Fnum, Fpar) ' just appending numeric to string (with '&') converts it correctly
            response = Analyzer.GetSend(filter_param_str, True)
          End If
        Next Fpar
        Call Fill_in_filter_params(Fnum) 'read back changed filter
        Call CheckIfFilterChanged(Fnum)  'update label colors
    Next Fnum
    'Call CalculateFilterTF(Fnum)   'Fnum there is out of range update plot of changed filter
End Sub

Private Sub CmdCopyFilters_Click()
    Dim Iso_axis As Long
    Dim Fnum As Long
    Dim Fpar As Long
    Dim filter_param_str As String
    Dim response As String
    Dim axis_number As Long

    Iso_axis = Int(FilterAxis / 4) 'IK 20210416 added int(7/4) it gives 1; without Int: 7/4 = 2!
    'axis type: 0 == Z axis; 1== X axis; 2== Y axis
'   for STACIS, it updates all SIMILAR filters, like Z1, Z2, Z3, Z4 at once.
'   for ELDAMP, it is too dangeropus.
'    For axis_number = 0 To 3
'        For Fnum = 0 To MAX_FILTERS_IN_AXIS - 1
'            For Fpar = 0 To USER_PARAM_NUMBER 'check change of 5 params AND FILTER TYPE
'                filter_param_str = "fpar" & ExtendedHEX(Iso_axis * 4 + axis_number) & CStr(Fnum) & CStr(Fpar)
'                filter_param_str = filter_param_str & "=" & Str(CHANGED_FilterParamArray(Fnum, Fpar))
'                response = Analyzer.GetSend(filter_param_str, True)
'            Next Fpar
'        Next Fnum
'    Next axis_number
'    For Fnum = 0 To MAX_FILTERS_IN_AXIS - 1
'        Call Fill_in_filter_params(Fnum) 'read back changed filter
'        Call CheckIfFilterChanged(Fnum)  'update label colors
'    Next Fnum
End Sub

'functions
''''''''''changes color of control and set/clear Boolean Filter_CHANGED(filt_NUM,, FilterParamNum)
Sub CheckIfFilterChanged(Filt_num As Long)
    Dim NN As Long
    Dim FilterChanged As Boolean
    FilterChanged = False
    For NN = 0 To USER_PARAM_NUMBER - 1 'check params change
        Filter_CHANGED(Filt_num, NN) = False 'preset not changed - if discovered it was - it will be overwritten
    Next NN
    For NN = 0 To USER_PARAM_NUMBER 'check 5 params change AND FILTER TYPE
        If CHANGED_FilterParamArray(Filt_num, NN) <> OriginalFilterParamArray(Filt_num, NN) Then
           Filter_CHANGED(Filt_num, NN) = True 'flag for individual param
           FilterChanged = True 'flag for whole filter line
           If NN - 1 = FilterGain Then Lbl_FilterGain(Filt_num).BackColor = &HC0C0FF ' -1 because we skip FilterTYPE with index 0
           If NN - 1 = freq_1 Then Lbl_FilterFreq1(Filt_num).BackColor = &HC0C0FF
           If NN - 1 = freq_2 Then Lbl_FilterFreq2(Filt_num).BackColor = &HC0C0FF
           If NN - 1 = Q_f_1 Then Lbl_FilterQ1(Filt_num).BackColor = &HC0C0FF
           If NN - 1 = Q_f_2 Then Lbl_FilterQ2(Filt_num).BackColor = &HC0C0FF
        End If
    Next NN
    If (FilterChanged = True) Then
        LblFilter(Filt_num).ForeColor = &H2020FF
        LblFilter(Filt_num).Caption = "*" & CStr(Filt_num + 1) & "*"
    Else 'not changed, restore white color
        LblFilter(Filt_num).Caption = "  " & CStr(Filt_num + 1) & " "
        LblFilter(Filt_num).ForeColor = &H80000008     'black
        Lbl_FilterGain(Filt_num).BackColor = &H80000005  'white
        Lbl_FilterFreq1(Filt_num).BackColor = &H80000005 'white
        Lbl_FilterFreq2(Filt_num).BackColor = &H80000005 'white
        Lbl_FilterQ1(Filt_num).BackColor = &H80000005    'white
        Lbl_FilterQ2(Filt_num).BackColor = &H80000005    'white
    End If
    CalculateFilterTF (Filt_num)
    If ChkShowAxisTF.Value = Checked Then
        Call CalculateAxis_TF
        Analyzer.CWGraph.Plots(Axis_TF_Num).Visible = True
        Analyzer.CWGraphPhase.Plots(Axis_TF_Num).Visible = True
    Else
        Analyzer.CWGraph.Plots(Axis_TF_Num).Visible = False
        Analyzer.CWGraphPhase.Plots(Axis_TF_Num).Visible = False
    End If
End Sub

Sub Fill_in_filter_params(filter_num As Long)
'reads filters from controller
'Dim OriginalFilterParamArray( 7,6) As Single '5 parameters, 7 filters in chain
'Dim Changed_FilterParamArray( 7,6) As Single '5 parameters, 7 filters in chain
    Dim i As Long
    Dim sep_pos As Long
    Dim filt_axis_code As String
    Dim filter_param_str As String
    Dim paramValue As Single
    Init_sys_data = True
    If (ConnectionType = ConnDEMO) Then Exit Sub

    filt_axis_code = ExtendedHEX(FilterAxis)

       filter_param_str = Analyzer.GetSend("fpar" & filt_axis_code & CStr(filter_num) & "0", True)
       sep_pos = InStr(filter_param_str, "=")
       If (sep_pos <> 0) Then
           retval = Mid(filter_param_str, sep_pos + 1)
           sep_pos = InStr(retval, "//") ''if verbose response
           If (sep_pos <> 0) Then
               retval = Left(retval, sep_pos - 1)
           End If
           If IsNumber(retval) Then
               paramValue = Val(retval)
               OriginalFilterParamArray(filter_num, 0) = paramValue 'filter type
               CHANGED_FilterParamArray(filter_num, 0) = paramValue 'filter type
               Filter_CHANGED(filter_num, 0) = False
           End If
       End If
'''''''''''''' end filter type
  ''fill parameters
    For i = 1 To FILTER_PARAMS_TOTAL - 1 '0 til 10
        filter_param_str = CStr(i)  '#= filter parameter number [0,1,2,3,4,5], where 0== long ftyp; the rest-> float par[5];
        If i = 10 Then filter_param_str = "A"
        filter_param_str = CStr(filter_num) & filter_param_str
        filter_param_str = filt_axis_code & filter_param_str
        filter_param_str = "fpar" & filter_param_str

        filter_param_str = Analyzer.GetSend(filter_param_str, True)
        sep_pos = InStr(filter_param_str, "=")
        If (sep_pos <> 0) Then
            retval = Mid(filter_param_str, sep_pos + 1)
            sep_pos = InStr(retval, "//") ''if verbose response
            If (sep_pos <> 0) Then
                retval = Left(retval, sep_pos - 1)
            End If
            If IsNumber(retval) Then
                paramValue = Val(retval)
                OriginalFilterParamArray(filter_num, i) = paramValue 'filter parameter as float[0 to 4], coefficient [5 to 9]
                CHANGED_FilterParamArray(filter_num, i) = paramValue 'make a copy for editing
                If (i < USER_PARAM_NUMBER) Then Filter_CHANGED(filter_num, i) = False
            End If
        End If
    Next i
    Init_sys_data = False
End Sub

'Sub SetNumericIncDec(nEdit As CWNumEdit) ' VB tells object required
Sub SetNumericIncDec(idx As Integer)
    Dim paramValue As Double
    Dim nEdit As CWNumEdit
    Dim FilterTypeNum As Integer
    FilterTypeNum = ComboFilterTYPE.ListIndex
    If (FilterTypeNum = DIRECT_COEFF) Then Exit Sub
    Set nEdit = cwNumFilterParam(idx)
    paramValue = nEdit.Value
    If (paramValue < 0.1) Then
        nEdit.IncDecValue = 0.001
    ElseIf (paramValue < 1) Then
        nEdit.IncDecValue = 0.01
    ElseIf (paramValue < 10) Then
        nEdit.IncDecValue = 0.1
    ElseIf (paramValue < 100) Then
        nEdit.IncDecValue = 1#
    Else
        nEdit.IncDecValue = 10#
    End If
End Sub

Sub Copy_Params_for_Edit(filter_num As Long) ' from original array to the moving "edit frame"
    Dim i As Long
    Dim sep_pos As Long
    Dim filter_param_str As String
    Dim paramValue As Single
    Init_sys_data = True

    FilterNumberInChain = filter_num 'update global variable
  ''fill filter type
    For i = 0 To ComboFilterTYPE.ListCount
      If Lbl_FilterType(filter_num).Caption = ComboFilterTYPE.List(i) Then
          ComboFilterTYPE.ListIndex = i
          Lbl_FilterType(filter_num).BackColor = vbWhite
          GoTo name_found
      End If
    Next i
    'string not found
    Lbl_FilterType(filter_num).BackColor = &HC0C0FF 'orange
name_found:
   '     ComboFilterTYPE.List (filter_num) ' = Lbl_FilterType(filter_num).Caption
  'Copy parameters from labels to numeric edits
    Call ChangeFilterFontsAndMins(filter_num, i)

    If IsNumber(Lbl_FilterGain(filter_num).Caption) Then ' <<<<<<<<<<<<<<<<< Filter GAIN
        paramValue = CSng(Lbl_FilterGain(filter_num).Caption) 'IK20220227 Label can have both comma and dot because of user input, need CSng
        If (ComboFilterTYPE.ListIndex <> DIRECT_COEFF) Then cwNumFilterParam(0).Minimum = 0.01 ' do not allow negative parameters, IK 20210417 set to -40 dB
        cwNumFilterParam(0).BackColorText = &HFFFFFF 'white
        cwNumFilterParam(0).Value = paramValue
        SetNumericIncDec (0)
    Else
        cwNumFilterParam(0).BackColorText = &HC0C0FF 'pink, cannot be converted
    End If

    If IsNumber(Lbl_FilterFreq1(filter_num).Caption) Then ' <<<<<<<<<<<<<<<<< F1 frequency
        paramValue = CSng(Lbl_FilterFreq1(filter_num).Caption) 'IK20220227 Label can have both comma and dot because of user input, need CSng
        If (ComboFilterTYPE.ListIndex <> DIRECT_COEFF) Then cwNumFilterParam(1).Minimum = 0.001 ' do not allow zero or negative parameters, IK 20210417 set to 0.001
        cwNumFilterParam(1).BackColorText = &HFFFFFF
        cwNumFilterParam(1).Value = paramValue
        SetNumericIncDec (1)
    Else
        cwNumFilterParam(1).BackColorText = &HC0C0FF
    End If

    If IsNumber(Lbl_FilterFreq2(filter_num).Caption) Then ' <<<<<<<<<<<<<<<<< F2 frequency
        paramValue = CSng(Lbl_FilterFreq2(filter_num).Caption) 'IK20220227 Label can have both comma and dot because of user input, need CSng
        If (ComboFilterTYPE.ListIndex <> DIRECT_COEFF) Then cwNumFilterParam(2).Minimum = 0.001 ' do not allow zero or negative parameters, IK 20210417 set to 0.001
        cwNumFilterParam(2).BackColorText = &HFFFFFF
        cwNumFilterParam(2).Value = paramValue
        SetNumericIncDec (2)
    Else
        cwNumFilterParam(2).BackColorText = &HC0C0FF
    End If

    If IsNumber(Lbl_FilterQ1(filter_num).Caption) Then ' <<<<<<<<<<<<<<<<< Q1
        paramValue = CSng(Lbl_FilterQ1(filter_num).Caption) 'IK20220227 Label can have both comma and dot because of user input, need CSng
        If InStr(Lbl_FilterType(filter_num).Caption, "PID") Then
            cwNumFilterParam(3).Minimum = 0  ' do not allow negative parameter Q1 for PID (leaky term), allow zero
            cwNumFilterParam(3).FormatString = ".#00"
        Else
            If (ComboFilterTYPE.ListIndex <> DIRECT_COEFF) Then
                cwNumFilterParam(3).Minimum = 0.001 ' do not allow zero or negative parameters, IK 20210417 set to 0.001
                cwNumFilterParam(3).FormatString = ".#00"
            End If
        End If
        cwNumFilterParam(3).BackColorText = &HFFFFFF
        cwNumFilterParam(3).Value = paramValue
        SetNumericIncDec (3)
    Else
        cwNumFilterParam(3).BackColorText = &HC0C0FF
    End If
    If IsNumber(Lbl_FilterQ2(filter_num).Caption) Then ' <<<<<<<<<<<<<<<<< Q2
        paramValue = CSng(Lbl_FilterQ2(filter_num).Caption) 'IK20220227 Label can have both comma and dot because of user input, need CSng
        If (ComboFilterTYPE.ListIndex <> DIRECT_COEFF) Then cwNumFilterParam(4).Minimum = 0.001 ' do not allow zero or negative parameters, IK 20210417 set to 0.001
        cwNumFilterParam(4).BackColorText = &HFFFFFF
        cwNumFilterParam(4).Value = paramValue
        SetNumericIncDec (4)
    Else
        cwNumFilterParam(4).BackColorText = &HC0C0FF
    End If

    Lbl_b2 = Format(OriginalFilterParamArray(filter_num, 6), FiltCoeffFormat)
    Lbl_b1 = Format(OriginalFilterParamArray(filter_num, 7), FiltCoeffFormat)
    Lbl_b0 = Format(OriginalFilterParamArray(filter_num, 8), FiltCoeffFormat)
    Lbl_a2 = Format(OriginalFilterParamArray(filter_num, 9), FiltCoeffFormat)
    Lbl_a1 = Format(OriginalFilterParamArray(filter_num, 10), FiltCoeffFormat)
    Init_sys_data = False
    Call Analyzer.Calc_prediction
    Call ShowDifference(filter_num)
End Sub

Sub Refresh_FilterParams()
'//define indexes for float_iir.par[index]
'#define FilterGain   0
'#define freq_1 1
'#define freq_2 2
'#define Q_f_1  3
'#define Q_f_2  4
'#define PID_P_gain FilterGain
'#define PID_I_gain freq_1
'#define PID_D_gain freq_2
'fpar&$#? returns "fpar&$#=GGG.gg" scientific format displayed; fpar&$#=ggg.gg sets filter parameter
'& = axis [0-F]== first index of SysData.filters[6][max_filt_in_channel]; SysData.PROXfilt[6][]; SysData.FF_Filt[4][],
'$= filter number in the axis [0,1,2,3,4,5] = = second index in SysData.filters[NUM_FILT_CHANs][max_filt_in_channel];
'
'#= filter parameter number [0,1,2,3,4,5] == long ftyp; float par[5];
'#= read-omly coefficient, read [6]=b2; [7]=b1; [8]=b0; [9]=a2; [A]=a1
    Dim filter_num As Long
    Dim f As Long
    Dim format_str As String

    FrameFilter.Visible = False ' hide editing panel to see filter behind

    retval = Analyzer.GetSend("echo>enab", True) 'disable "echo>verb" to speed up

    For filter_num = 0 To MAX_FILTERS_IN_AXIS - 1 '0 to 5, 5 params AND FILTER_TYPE
        Call Fill_in_filter_params(filter_num) ' read one filter
        If (OriginalFilterParamArray(filter_num, 0) = DIRECT_COEFF) Then
            format_str = FiltCoeffFormat '"#0.######0"
            For f = 0 To 5 ' test
                If (OriginalFilterParamArray(filter_num, f) = 0) Then
                    Call MsgBox("response = 0", vbOKOnly, "too small response")
                End If
            Next f
        Else
            format_str = FiltParFormat '"###0.000"
        End If
        Call ChangeFilterFontsAndMins(filter_num, CInt(OriginalFilterParamArray(filter_num, 0)))

        Lbl_FilterType(filter_num).Caption = ComboFilterTYPE.List(OriginalFilterParamArray(filter_num, 0))
        Lbl_FilterGain(filter_num).Caption = Format(OriginalFilterParamArray(filter_num, 1), format_str)    'index=0
        Lbl_FilterFreq1(filter_num).Caption = Format(OriginalFilterParamArray(filter_num, 2), format_str) 'index=1
        Lbl_FilterFreq2(filter_num).Caption = Format(OriginalFilterParamArray(filter_num, 3), format_str) 'index=2
        Lbl_FilterQ1(filter_num).Caption = Format(OriginalFilterParamArray(filter_num, 4), format_str)    'index=3
        Lbl_FilterQ2(filter_num).Caption = Format(OriginalFilterParamArray(filter_num, 5), format_str)    'index=4
        Call CheckIfFilterChanged(filter_num)
    Next filter_num
    Call Copy_Params_for_Edit(FilterNumberInChain)
End Sub

Private Sub CmdCalcMagPhaseNOW_Click()
    Dim TF As tFreqHz_MagTimes_PhaseDeg
    Dim FiltCoef As float_iir
    Dim dB As Double
    Dim frq As Double
    Dim pos As Long
    Dim frq_str As String
    If IsNumber(Lbl_a1) Then FiltCoef.a1 = CDbl(Lbl_a1)
    If IsNumber(Lbl_a2) Then FiltCoef.a2 = CDbl(Lbl_a2)
    If IsNumber(Lbl_b0) Then FiltCoef.b0 = CDbl(Lbl_b0)
    If IsNumber(Lbl_b1) Then FiltCoef.b1 = CDbl(Lbl_b1)
    If IsNumber(Lbl_b2) Then FiltCoef.b2 = CDbl(Lbl_b2)
    frq_str = TxtFreq.Text
    If IsNumber(frq_str) Then
        pos = InStr(frq_str, ",")
        If pos <> 0 Then Mid(frq_str, pos, 1) = "."
        frq = Val(frq_str)
    Else
        Exit Sub
    End If
    TF = Mag_Ph_vs_Freq(FiltCoef, frq, LOOP_FREQUENCY)
    dB = 20 * Log(TF.mag) / Log(10)
    LblMag.Caption = Format(dB, "###0.00")
    LblPhase.Caption = Format(TF.Phase, "###0.00")
End Sub

Private Sub cmdClearFilterHistory_Click()
    cmd_sent = False
    Call Analyzer.GetSend("fpar>histclr", True)      'clear filter history cmd
End Sub

'====================================  E N D   O F  =========================================================================
'''''''''''''''''''''''''''''''''''''' FILTER PARAMETERS ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''



