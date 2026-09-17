VERSION 5.00
Object = "{D940E4E4-6079-11CE-88CB-0020AF6845F6}#1.6#0"; "cwui.ocx"
Begin VB.Form frmDrivingOutput 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Choose Driving Output"
   ClientHeight    =   6270
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5010
   Icon            =   "frmDrivingOutput.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6270
   ScaleWidth      =   5010
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton optExcitOutput 
      BackColor       =   &H00C0E0FF&
      Caption         =   "DAC  OUTPUT #"
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
      Index           =   32
      Left            =   180
      TabIndex        =   41
      Tag             =   "BNC0"
      ToolTipText     =   "Excitation signal is applied ONLY to BNC Diag.0, to be used as shaker signal, par=BNC0"
      Top             =   5850
      Width           =   1635
   End
   Begin VB.OptionButton optExcitOutput 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Z axis Vel In"
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
      Index           =   31
      Left            =   180
      TabIndex        =   37
      Tag             =   "GA_Z"
      ToolTipText     =   "Excitation signal is applied to a velocity axis INPUT (after geophones values are combined), par=GA_Z"
      Top             =   3330
      Width           =   1450
   End
   Begin VB.OptionButton optExcitOutput 
      BackColor       =   &H00C0FFC0&
      Caption         =   "tY axis Vel In"
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
      Index           =   30
      Left            =   3330
      TabIndex        =   36
      Tag             =   "GATY"
      ToolTipText     =   "Excitation signal is applied to a velocity axis INPUT (after geophones values are combined), par=GATY"
      Top             =   3330
      Width           =   1450
   End
   Begin VB.OptionButton optExcitOutput 
      BackColor       =   &H00C0FFC0&
      Caption         =   "tX axis Vel In"
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
      Index           =   29
      Left            =   1755
      TabIndex        =   35
      Tag             =   "GATX"
      ToolTipText     =   "Excitation signal is applied to a velocity axis INPUT (after geophones values are combined), par=GATX"
      Top             =   3330
      Width           =   1450
   End
   Begin VB.OptionButton optExcitOutput 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Z axis Pos In"
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
      Index           =   28
      Left            =   180
      TabIndex        =   34
      Tag             =   "PA_Z"
      ToolTipText     =   "Excitation signal is applied to a position axis INPUT (after proximity values are combined), par=PA_Z"
      Top             =   2880
      Value           =   -1  'True
      Width           =   1450
   End
   Begin VB.OptionButton optExcitOutput 
      BackColor       =   &H00FFFFC0&
      Caption         =   "tY axis Pos In"
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
      Index           =   27
      Left            =   3330
      TabIndex        =   33
      Tag             =   "PATY"
      ToolTipText     =   "Excitation signal is applied to a position axis INPUT (after proximity values are combined), par=PATY"
      Top             =   2880
      Width           =   1450
   End
   Begin VB.OptionButton optExcitOutput 
      BackColor       =   &H00FFFFC0&
      Caption         =   "tX axis Pos In"
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
      Index           =   26
      Left            =   1755
      TabIndex        =   32
      Tag             =   "PATX"
      ToolTipText     =   "Excitation signal is applied to a position axis INPUT (after proximity values are combined), par=PATX"
      Top             =   2880
      Width           =   1450
   End
   Begin VB.OptionButton optExcitOutput 
      BackColor       =   &H00C0C0C0&
      Caption         =   "BNC Diag.0"
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
      Index           =   25
      Left            =   3465
      TabIndex        =   31
      Tag             =   "BNC0"
      ToolTipText     =   "Excitation signal is applied ONLY to BNC Diag.0, to be used as shaker signal, par=BNC0"
      Top             =   5850
      Visible         =   0   'False
      Width           =   1320
   End
   Begin VB.CheckBox chkZmotors 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Vertical Motors Present"
      Height          =   285
      Left            =   45
      TabIndex        =   30
      ToolTipText     =   "Enable selection of vertical motors"
      Top             =   495
      Width           =   2370
   End
   Begin VB.OptionButton optSystemType 
      BackColor       =   &H000080FF&
      Caption         =   "Non standard Horiz Motors"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   2
      Left            =   2565
      TabIndex        =   29
      ToolTipText     =   "Isolators have non stardard motor layout"
      Top             =   495
      Width           =   2370
   End
   Begin VB.OptionButton optSystemType 
      BackColor       =   &H00FFC0FF&
      Caption         =   "Y1-X2-Y3-X4 Horiz Motors"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   1
      Left            =   2565
      TabIndex        =   28
      ToolTipText     =   "Isolator #1 has motor Y, Iso#2 has motor X, Iso #3 has motor Y, Iso #4 has motor X"
      Top             =   135
      Width           =   2370
   End
   Begin VB.OptionButton optSystemType 
      BackColor       =   &H00FFC0C0&
      Caption         =   "X1-Y2-X3-Y4 Horiz Motors"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   0
      Left            =   45
      TabIndex        =   27
      ToolTipText     =   "Isolator #1 has motor X, Iso#2 has motor Y, Iso #3 has motor X, Iso #4 has motor Y"
      Top             =   135
      Width           =   2370
   End
   Begin VB.OptionButton optExcitOutput 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Z axis mot"
      Enabled         =   0   'False
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
      Index           =   22
      Left            =   3330
      TabIndex        =   24
      Tag             =   "MA_Z"
      ToolTipText     =   "Excitation signal is applied to axis, each linear motor gets its part of the signal, par=MA_Z"
      Top             =   2070
      Width           =   1450
   End
   Begin VB.OptionButton optExcitOutput 
      BackColor       =   &H00C0FFC0&
      Caption         =   "tY axis mot"
      Enabled         =   0   'False
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
      Index           =   24
      Left            =   1755
      TabIndex        =   23
      Tag             =   "MATY"
      ToolTipText     =   "Excitation signal is applied to axis, each linear motor gets its part of the signal, par=MATY"
      Top             =   2070
      Width           =   1450
   End
   Begin VB.OptionButton optExcitOutput 
      BackColor       =   &H00C0FFC0&
      Caption         =   "tX axis mot"
      Enabled         =   0   'False
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
      Index           =   23
      Left            =   180
      TabIndex        =   22
      Tag             =   "MATX"
      ToolTipText     =   "Excitation signal is applied to axis, each linear motor gets its part of the signal, par=MATX"
      Top             =   2070
      Width           =   1450
   End
   Begin VB.OptionButton optExcitOutput 
      BackColor       =   &H00C0FFFF&
      Caption         =   "tZ axis mot"
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
      Index           =   21
      Left            =   3330
      TabIndex        =   21
      Tag             =   "MATZ"
      ToolTipText     =   "Excitation signal is applied to axis, each linear motor gets its part of the signal, par=MATZ"
      Top             =   1665
      Width           =   1450
   End
   Begin VB.OptionButton optExcitOutput 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Y axis mot"
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
      Index           =   20
      Left            =   1755
      TabIndex        =   20
      Tag             =   "MA_Y"
      ToolTipText     =   "Excitation signal is applied to axis, each linear motor gets its part of the signal, par=MA_Y"
      Top             =   1665
      Width           =   1450
   End
   Begin VB.OptionButton optExcitOutput 
      BackColor       =   &H00C0FFFF&
      Caption         =   "X axis mot"
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
      Index           =   19
      Left            =   180
      TabIndex        =   19
      Tag             =   "MA_X"
      ToolTipText     =   "Excitation signal is applied to axis, each linear motor gets its part of the signal, par=MA_X"
      Top             =   1665
      Width           =   1450
   End
   Begin VB.OptionButton optExcitOutput 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Z Pneumatic"
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
      Index           =   16
      Left            =   3330
      TabIndex        =   18
      Tag             =   "VA_Z"
      ToolTipText     =   "Excitation signal is applied to axis, each valve gets its part of the signal, par=VA_Z"
      Top             =   1260
      Width           =   1450
   End
   Begin VB.OptionButton optExcitOutput 
      BackColor       =   &H00FFFFC0&
      Caption         =   "tY Pneumatic"
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
      Index           =   17
      Left            =   1755
      TabIndex        =   17
      Tag             =   "VATY"
      ToolTipText     =   "Excitation signal is applied to axis, each valve gets its part of the signal, par=VATY"
      Top             =   1260
      Width           =   1450
   End
   Begin VB.OptionButton optExcitOutput 
      BackColor       =   &H00FFFFC0&
      Caption         =   "tX Pneumatic"
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
      Index           =   18
      Left            =   180
      TabIndex        =   16
      Tag             =   "VATX"
      ToolTipText     =   "Excitation signal is applied to axis, each valve gets its part of the signal, par=VATX"
      Top             =   1260
      Width           =   1450
   End
   Begin VB.OptionButton optExcitOutput 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Motor Z4"
      Enabled         =   0   'False
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
      Index           =   15
      Left            =   3690
      TabIndex        =   15
      Tag             =   "M_Z4"
      ToolTipText     =   "Excitation signal is applied to ONE linear motor, par=M_Z4"
      Top             =   5355
      Width           =   1070
   End
   Begin VB.OptionButton optExcitOutput 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Motor Y4"
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
      Index           =   14
      Left            =   3690
      TabIndex        =   14
      Tag             =   "M_Y4"
      ToolTipText     =   "Excitation signal is applied to ONE linear motor, par=M_Y4"
      Top             =   4950
      Width           =   1070
   End
   Begin VB.OptionButton optExcitOutput 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Motor X4"
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
      Index           =   13
      Left            =   3690
      TabIndex        =   13
      Tag             =   "M_X4"
      ToolTipText     =   "Excitation signal is applied to ONE linear motor, par=M_X4"
      Top             =   4545
      Width           =   1070
   End
   Begin VB.OptionButton optExcitOutput 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Motor Z3"
      Enabled         =   0   'False
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
      Index           =   12
      Left            =   2520
      TabIndex        =   12
      Tag             =   "M_Z3"
      ToolTipText     =   "Excitation signal is applied to ONE linear motor, par=M_Z3"
      Top             =   5355
      Width           =   1070
   End
   Begin VB.OptionButton optExcitOutput 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Motor Y3"
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
      Index           =   11
      Left            =   2520
      TabIndex        =   11
      Tag             =   "M_Y3"
      ToolTipText     =   "Excitation signal is applied to ONE linear motor, par=M_Y3"
      Top             =   4950
      Width           =   1070
   End
   Begin VB.OptionButton optExcitOutput 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Motor X3"
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
      Index           =   10
      Left            =   2520
      TabIndex        =   10
      Tag             =   "M_X3"
      ToolTipText     =   "Excitation signal is applied to ONE linear motor, par=M_X3"
      Top             =   4545
      Width           =   1070
   End
   Begin VB.OptionButton optExcitOutput 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Motor Z2"
      Enabled         =   0   'False
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
      Index           =   9
      Left            =   1350
      TabIndex        =   9
      Tag             =   "M_Z2"
      ToolTipText     =   "Excitation signal is applied to ONE linear motor, par=M_Z2"
      Top             =   5355
      Width           =   1070
   End
   Begin VB.OptionButton optExcitOutput 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Motor Y2"
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
      Index           =   8
      Left            =   1350
      TabIndex        =   8
      Tag             =   "M_Y2"
      ToolTipText     =   "Excitation signal is applied to ONE linear motor, par=M_Y2"
      Top             =   4950
      Width           =   1070
   End
   Begin VB.OptionButton optExcitOutput 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Motor X2"
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
      Index           =   7
      Left            =   1350
      TabIndex        =   7
      Tag             =   "M_X2"
      ToolTipText     =   "Excitation signal is applied to ONE linear motor, par=M_X2"
      Top             =   4545
      Width           =   1070
   End
   Begin VB.OptionButton optExcitOutput 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Motor Z1"
      Enabled         =   0   'False
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
      Index           =   6
      Left            =   180
      TabIndex        =   6
      Tag             =   "M_Z1"
      ToolTipText     =   "Excitation signal is applied to ONE linear motor, par=M_Z1"
      Top             =   5355
      Width           =   1070
   End
   Begin VB.OptionButton optExcitOutput 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Motor Y1"
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
      Index           =   5
      Left            =   180
      TabIndex        =   5
      Tag             =   "M_Y1"
      ToolTipText     =   "Excitation signal is applied to ONE linear motor, par=M_Y1"
      Top             =   4950
      Width           =   1070
   End
   Begin VB.OptionButton optExcitOutput 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Valve 4"
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
      Index           =   3
      Left            =   3690
      TabIndex        =   4
      Tag             =   "V_Z4"
      ToolTipText     =   "Excitation signal is applied to a single valve, par=V_Z4"
      Top             =   4140
      Width           =   1070
   End
   Begin VB.OptionButton optExcitOutput 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Valve 3"
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
      Index           =   2
      Left            =   2520
      TabIndex        =   3
      Tag             =   "V_Z3"
      ToolTipText     =   "Excitation signal is applied to a single valve, par=V_Z3"
      Top             =   4140
      Width           =   1070
   End
   Begin VB.OptionButton optExcitOutput 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Motor X1"
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
      Index           =   4
      Left            =   180
      TabIndex        =   2
      Tag             =   "M_X1"
      ToolTipText     =   "Excitation signal is applied to ONE linear motor, par=M_X1"
      Top             =   4545
      Width           =   1070
   End
   Begin VB.OptionButton optExcitOutput 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Valve 2"
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
      Left            =   1350
      TabIndex        =   1
      Tag             =   "V_Z2"
      ToolTipText     =   "Excitation signal is applied to a single valve, par=V_Z2"
      Top             =   4140
      Width           =   1070
   End
   Begin VB.OptionButton optExcitOutput 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Valve 1"
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
      Left            =   180
      TabIndex        =   0
      Tag             =   "V_Z1"
      ToolTipText     =   "Excitation signal is applied to a single valve, par=V_Z1"
      Top             =   4140
      Width           =   1070
   End
   Begin CWUIControlsLib.CWNumEdit cwNumDACchannel 
      Height          =   300
      Left            =   1845
      TabIndex        =   39
      Tag             =   "DA"
      ToolTipText     =   "Output DAC channel number, 0 to 23. NOTE: DAC23 ->BNC0; DAC22->BNC1"
      Top             =   5850
      Width           =   810
      _Version        =   524288
      _ExtentX        =   1429
      _ExtentY        =   529
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
      C[0]_1          =   12640511
      C[1]_1          =   0
      BorderStyle_1   =   1
      TextAlignment_1 =   2
      format_1        =   2
      ClassName_2     =   "CCWFormat"
      Format_2        =   "."
      scale_1         =   3
      ClassName_3     =   "CCWScale"
      opts_3          =   65536
      dMax_3          =   10
      discInterval_3  =   1
      ValueVarType_1  =   5
      Value_Val_1     =   20
      IncValueVarType_1=   5
      IncValue_Val_1  =   1
      AccelIncVarType_1=   5
      AccelInc_Val_1  =   0.5
      RangeMinVarType_1=   5
      RangeMaxVarType_1=   5
      RangeMax_Val_1  =   23
      BackColor_1     =   16777215
      ButtonStyle_1   =   0
      Bindings_1      =   4
      ClassName_4     =   "CCWBindingHolderArray"
      Editor_4        =   5
      ClassName_5     =   "CCWBindingHolderArrayEditor"
      Owner_5         =   1
   End
   Begin VB.Label lblAXES 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   420
      Index           =   1
      Left            =   45
      TabIndex        =   40
      ToolTipText     =   "Double-click updates input and closes window"
      Top             =   5805
      Width           =   4875
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Caption         =   "D R  I  V  I  N  G      A  X  I  S      I N P U T"
      ForeColor       =   &H80000008&
      Height          =   1140
      Left            =   45
      TabIndex        =   38
      ToolTipText     =   "Double-click updates output and closes window"
      Top             =   2610
      Width           =   4875
   End
   Begin VB.Label lblACTUATORS 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Caption         =   "D R  I  V  I  N  G      A  C  T  U  A  T  O  R"
      ForeColor       =   &H80000008&
      Height          =   1905
      Left            =   45
      TabIndex        =   26
      ToolTipText     =   "Double-click updates output and closes window"
      Top             =   3870
      Width           =   4875
   End
   Begin VB.Label lblAXES 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Caption         =   "D R  I  V  I  N  G      A  X  I  S      O U T P U T"
      ForeColor       =   &H80000008&
      Height          =   1545
      Index           =   0
      Left            =   45
      TabIndex        =   25
      ToolTipText     =   "Double-click updates output and closes window"
      Top             =   945
      Width           =   4875
   End
End
Attribute VB_Name = "frmDrivingOutput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cwNumDACchannel_DblClick()
    Call Leaving_Form
End Sub

Private Sub cwNumDACchannel_ValueChanged(Value As Variant, PreviousValue As Variant, ByVal OutOfRange As Boolean)
    optExcitOutput(32).Value = True
    Me.Caption = "Choose Driving Output. Now = DAC " & Format(Value, "00")
End Sub

Private Sub Form_Activate()
    If (LTFadvanced = True) Then
        Me.optExcitOutput(25).Visible = True 'Excitation signal is applied ONLY to BNC Diag.0, to be used as shaker signal, par=BNC0
    Else ' regular TF
        Me.optExcitOutput(25).Visible = False ' Hide BNC0 option as input for driving
    End If
    Call Form_click
'    If (OutputChosen = 32) Then ' radio button "DAC channel #" is selected
'        Me.Caption = "Choose Driving Output. Now = DAC " & cwNumDACchannel.Value
'        LTFtestOutputName = cwNumDACchannel.Tag & cwNumDACchannel.Value
'    Else
'        Me.Caption = "Choose Driving Output. Now = " & optExcitOutput(OutputChosen).Caption
'        LTFtestOutputName = optExcitOutput(OutputChosen).Tag
'    End If
End Sub

Private Sub Form_Load()
    Call Form_click
'    If (OutputChosen = 32) Then ' radio button "DAC channel #" is selected
'        Me.Caption = "Choose Driving Output. Now = DAC " & Format(cwNumDACchannel.Value, "00")
'        LTFtestOutputName = cwNumDACchannel.Tag & cwNumDACchannel.Value
'    Else
'        Me.Caption = "Choose Driving Output. Now = " & optExcitOutput(OutputChosen).Caption
'        LTFtestOutputName = optExcitOutput(OutputChosen).Tag
'    End If
End Sub

Private Sub Form_click()
    If (OutputChosen = 32) Then ' radio button "DAC channel #" is selected
        Me.Caption = "Choose Driving Output. Now = DAC " & Format(cwNumDACchannel.Value, "00")
        LTFtestOutputName = cwNumDACchannel.Tag & Format(cwNumDACchannel.Value, "00")
    Else
        Me.Caption = "Choose Driving Output. Now = " & optExcitOutput(OutputChosen).Caption
        LTFtestOutputName = optExcitOutput(OutputChosen).Tag
    End If
End Sub

Private Sub Form_LostFocus()
    If (OutputChosen = 32) Then ' radio button "DAC channel #" is selected
        Me.Caption = "Choose Driving Output. Now = DAC " & cwNumDACchannel.Value
        LTFtestOutputName = cwNumDACchannel.Tag & Format(cwNumDACchannel.Value, "00")
        Analyzer.cmdChooseExcitation.Caption = "Choose Output" & vbCrLf & "Now = DAC " & Format(cwNumDACchannel.Value, "00")
    Else
        Me.Caption = "Choose Driving Output. Now = " & optExcitOutput(OutputChosen).Caption
        LTFtestOutputName = optExcitOutput(OutputChosen).Tag
        Analyzer.cmdChooseExcitation.Caption = "Choose Output" & vbCrLf & "Now = " & optExcitOutput(OutputChosen).Caption
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call Leaving_Form
End Sub

Private Sub optExcitOutput_Click(Index As Integer)
    OutputChosen = Index
    Call Form_click
'    If (OutputChosen = 32) Then ' radio button "DAC channel #" is selected
'        Me.Caption = "Choose Driving Output. Now = DAC " & cwNumDACchannel.Value
'        LTFtestOutputName = cwNumDACchannel.Tag & cwNumDACchannel.Value
'    Else
'        Me.Caption = "Choose Driving Output. Now = " & optExcitOutput(OutputChosen).Caption
'        LTFtestOutputName = optExcitOutput(OutputChosen).Tag
'    End If
End Sub

Private Sub optExcitOutput_DblClick(Index As Integer)
    OutputChosen = Index
    Call Leaving_Form
End Sub

Sub Leaving_Form()
    Call Form_LostFocus
'    If (OutputChosen = 32) Then ' radio button "DAC channel #" is selected
'        Me.Caption = "Choose Driving Output. Now = DAC " & cwNumDACchannel.Value
'        LTFtestOutputName = cwNumDACchannel.Tag & cwNumDACchannel.Value
'        Analyzer.cmdChooseExcitation.Caption = "Choose Output" & vbCrLf & "Now = DAC " & cwNumDACchannel.Value
'    Else
'        Me.Caption = "Choose Driving Output. Now = " & optExcitOutput(OutputChosen).Caption
'        LTFtestOutputName = optExcitOutput(OutputChosen).Tag
'        Analyzer.cmdChooseExcitation.Caption = "Choose Output" & vbCrLf & "Now = " & optExcitOutput(OutputChosen).Caption
'    End If
    Me.Hide
    Call Analyzer.Check_Selected_Actuator_Sensor ' this will load proper limit file OR if selection is wrong, message and hide limit plots
End Sub

Private Sub optSystemType_Click(Index As Integer)
    If Index = 0 Then 'X1-Y2-X3-Y4
        optExcitOutput(4).Enabled = True 'X1
        optExcitOutput(5).Enabled = False 'Y1
        optExcitOutput(7).Enabled = False 'X2
        optExcitOutput(8).Enabled = True 'Y2
        optExcitOutput(10).Enabled = True 'X3
        optExcitOutput(11).Enabled = False 'Y3
        optExcitOutput(13).Enabled = False 'X4
        optExcitOutput(14).Enabled = True 'Y4
    End If
    If Index = 1 Then 'Y1-X2-Y3-X4
        optExcitOutput(4).Enabled = False 'X1
        optExcitOutput(5).Enabled = True 'Y1
        optExcitOutput(7).Enabled = True 'X2
        optExcitOutput(8).Enabled = False 'Y2
        optExcitOutput(10).Enabled = False 'X3
        optExcitOutput(11).Enabled = True 'Y3
        optExcitOutput(13).Enabled = True 'X4
        optExcitOutput(14).Enabled = False 'Y4
    End If
    If Index = 2 Then 'Non standard
        optExcitOutput(4).Enabled = True 'X1
        optExcitOutput(5).Enabled = True 'Y1
        optExcitOutput(7).Enabled = True 'X2
        optExcitOutput(8).Enabled = True 'Y2
        optExcitOutput(10).Enabled = True 'X3
        optExcitOutput(11).Enabled = True 'Y3
        optExcitOutput(13).Enabled = True 'X4
        optExcitOutput(14).Enabled = True 'Y4
    End If
    If chkZmotors.Value = Checked Then 'Vertical Motors Present
        optExcitOutput(6).Enabled = True 'Z1
        optExcitOutput(9).Enabled = True 'Z2
        optExcitOutput(12).Enabled = True 'Z3
        optExcitOutput(15).Enabled = True 'Z4
        optExcitOutput(22).Enabled = True 'Z2
        optExcitOutput(23).Enabled = True 'Z3
        optExcitOutput(24).Enabled = True 'Z4
    Else  'NO Vertical Motors
        optExcitOutput(6).Enabled = False 'Z1
        optExcitOutput(9).Enabled = False 'Z2
        optExcitOutput(12).Enabled = False 'Z3
        optExcitOutput(15).Enabled = False 'Z4
        optExcitOutput(22).Enabled = False 'Z2
        optExcitOutput(23).Enabled = False 'Z3
        optExcitOutput(24).Enabled = False 'Z4
    End If
End Sub

Private Sub chkZmotors_Click()
    If chkZmotors.Value = Checked Then 'Vertical Motors Present
        optExcitOutput(6).Enabled = True 'Z1
        optExcitOutput(9).Enabled = True 'Z2
        optExcitOutput(12).Enabled = True 'Z3
        optExcitOutput(15).Enabled = True 'Z4
        optExcitOutput(22).Enabled = True 'Z2
        optExcitOutput(23).Enabled = True 'Z3
        optExcitOutput(24).Enabled = True 'Z4
    Else  'NO Vertical Motors
        optExcitOutput(6).Enabled = False 'Z1
        optExcitOutput(9).Enabled = False 'Z2
        optExcitOutput(12).Enabled = False 'Z3
        optExcitOutput(15).Enabled = False 'Z4
        optExcitOutput(22).Enabled = False 'Z2
        optExcitOutput(23).Enabled = False 'Z3
        optExcitOutput(24).Enabled = False 'Z4
    End If
End Sub

