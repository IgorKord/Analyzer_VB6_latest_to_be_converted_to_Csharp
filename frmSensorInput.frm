VERSION 5.00
Object = "{D940E4E4-6079-11CE-88CB-0020AF6845F6}#1.6#0"; "cwui.ocx"
Begin VB.Form frmSensorInput 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Choose Test Input"
   ClientHeight    =   8250
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5010
   Icon            =   "frmSensorInput.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8250
   ScaleWidth      =   5010
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton optSensorInput 
      BackColor       =   &H00FFC0C0&
      Caption         =   "A  D  C   I N P U T #"
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
      Index           =   31
      Left            =   135
      TabIndex        =   53
      Tag             =   "GA_Z"
      ToolTipText     =   "Input signal from a velocity axis (combined geophones value), par=GA_Z"
      Top             =   7840
      Width           =   2310
   End
   Begin VB.OptionButton optSensorInput 
      BackColor       =   &H00FFFFC0&
      Caption         =   "FloorFF_Z"
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
      TabIndex        =   51
      Tag             =   "FFFZ"
      ToolTipText     =   "Floor FF Z Signal, par=FFFZ"
      Top             =   6885
      Width           =   1450
   End
   Begin VB.OptionButton optSensorInput 
      BackColor       =   &H00FFFFC0&
      Caption         =   "FloorFF_Y"
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
      TabIndex        =   49
      Tag             =   "FFFY"
      ToolTipText     =   "Floor FF Z Signal, par=FFFY"
      Top             =   6885
      Width           =   1450
   End
   Begin VB.OptionButton optSensorInput 
      BackColor       =   &H00FFFFC0&
      Caption         =   "FloorFF_X"
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
      TabIndex        =   48
      Tag             =   "FFFX"
      ToolTipText     =   "Floor FF Z Signal, par=FFFX"
      Top             =   6885
      Width           =   1450
   End
   Begin VB.OptionButton optSensorInput 
      BackColor       =   &H00C0FFC0&
      Caption         =   "StFF Ypos"
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
      Left            =   3690
      TabIndex        =   47
      Tag             =   "SFYP"
      ToolTipText     =   "Stage FF Y position signal, par=SFYP"
      Top             =   7290
      Width           =   1105
   End
   Begin VB.OptionButton optSensorInput 
      BackColor       =   &H00C0FFC0&
      Caption         =   "StFF Xacc"
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
      Left            =   135
      TabIndex        =   46
      Tag             =   "SFXA"
      ToolTipText     =   "Stage FF X acceleration signal, par=SFXA"
      Top             =   7290
      Width           =   1080
   End
   Begin VB.OptionButton optSensorInput 
      BackColor       =   &H00C0FFC0&
      Caption         =   "StFF Yacc"
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
      Left            =   1305
      TabIndex        =   45
      Tag             =   "SFYA"
      ToolTipText     =   "Stage FF Y acceleration signal, par=SFYA"
      Top             =   7290
      Width           =   1080
   End
   Begin VB.OptionButton optSensorInput 
      BackColor       =   &H00C0FFC0&
      Caption         =   "StFF Xpos"
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
      Left            =   2475
      TabIndex        =   44
      Tag             =   "SFXP"
      ToolTipText     =   "Stage FF Y position signal, par=SFXP"
      Top             =   7290
      Width           =   1125
   End
   Begin VB.OptionButton optSensorInput 
      BackColor       =   &H00FF80FF&
      Caption         =   "Excitation"
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
      TabIndex        =   43
      Tag             =   "EXCI"
      ToolTipText     =   "Use Excitation as TF parameter is common for P_TF and OLTF. Choose some other input for advanced TF with third parameter, par=EXCI"
      Top             =   6480
      Width           =   1455
   End
   Begin VB.CheckBox chkHprox 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Horizontal Position Present"
      Height          =   285
      Left            =   2565
      TabIndex        =   42
      ToolTipText     =   "Check if system has Horizontal position sensors (proximity)"
      Top             =   90
      Value           =   1  'Checked
      Width           =   2325
   End
   Begin VB.CheckBox chkHgeophones 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Horizontal Velocity Present"
      Height          =   285
      Left            =   2565
      TabIndex        =   41
      ToolTipText     =   "Check if system has horizontal velocity sensors (geophones)"
      Top             =   495
      Width           =   2325
   End
   Begin VB.CheckBox chkZgeophones 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Vertical Velocity Present"
      Height          =   285
      Left            =   2565
      TabIndex        =   40
      ToolTipText     =   "Check if system has vertical velocity sensors (geophones)"
      Top             =   900
      Width           =   2325
   End
   Begin VB.OptionButton optSystemType 
      BackColor       =   &H000080FF&
      Caption         =   "Non standard Layout"
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
      Left            =   90
      TabIndex        =   39
      ToolTipText     =   "Isolators have non stardard sensors layout"
      Top             =   900
      Value           =   -1  'True
      Width           =   2370
   End
   Begin VB.OptionButton optSystemType 
      BackColor       =   &H00FFC0FF&
      Caption         =   "Y1-X2-Y3 Horiz Sensors"
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
      Left            =   90
      TabIndex        =   38
      ToolTipText     =   "Isolator #1 has sensor Y, Iso#2 has sensor X, Iso #3 has sensor Y"
      Top             =   495
      Width           =   2370
   End
   Begin VB.OptionButton optSystemType 
      BackColor       =   &H00FFC0C0&
      Caption         =   "X1-Y2-X3 Horiz Sensors"
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
      Left            =   90
      TabIndex        =   37
      ToolTipText     =   "Isolator #1 has sensor X, Iso#2 has sensor Y, Iso #3 has sensor X"
      Top             =   90
      Width           =   2370
   End
   Begin VB.OptionButton optSensorInput 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Z axis Vel"
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
      Index           =   41
      Left            =   3330
      TabIndex        =   36
      Tag             =   "GA_Z"
      ToolTipText     =   "Input signal from a velocity axis (combined geophones value), par=GA_Z"
      Top             =   2790
      Width           =   1450
   End
   Begin VB.OptionButton optSensorInput 
      BackColor       =   &H00C0FFC0&
      Caption         =   "tY axis Vel"
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
      Index           =   42
      Left            =   1755
      TabIndex        =   35
      Tag             =   "GATY"
      ToolTipText     =   "Input signal from a velocity axis (combined geophones value), par=GATY"
      Top             =   2790
      Width           =   1450
   End
   Begin VB.OptionButton optSensorInput 
      BackColor       =   &H00C0FFC0&
      Caption         =   "tX axis Vel"
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
      Index           =   43
      Left            =   180
      TabIndex        =   34
      Tag             =   "GATX"
      ToolTipText     =   "Input signal from a velocity axis (combined geophones value), par=GATX"
      Top             =   2790
      Width           =   1450
   End
   Begin VB.OptionButton optSensorInput 
      BackColor       =   &H00C0FFFF&
      Caption         =   "tZ axis Vel"
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
      Index           =   40
      Left            =   3330
      TabIndex        =   33
      Tag             =   "GATZ"
      ToolTipText     =   "Input signal from a velocity axis (combined geophones value), par=GATZ"
      Top             =   2385
      Width           =   1450
   End
   Begin VB.OptionButton optSensorInput 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Y axis Vel"
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
      Index           =   39
      Left            =   1755
      TabIndex        =   32
      Tag             =   "GA_Y"
      ToolTipText     =   "Input signal from a velocity axis (combined geophones value), par=GA_Y"
      Top             =   2385
      Width           =   1450
   End
   Begin VB.OptionButton optSensorInput 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Z axis Pos"
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
      Index           =   35
      Left            =   3330
      TabIndex        =   31
      Tag             =   "PA_Z"
      ToolTipText     =   "Input signal from a position axis (combined proximity value), par=PA_Z"
      Top             =   1980
      Width           =   1450
   End
   Begin VB.OptionButton optSensorInput 
      BackColor       =   &H00FFFFC0&
      Caption         =   "tY axis Pos"
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
      Index           =   36
      Left            =   1755
      TabIndex        =   30
      Tag             =   "PATY"
      ToolTipText     =   "Input signal from a position axis (combined proximity value), par=PATY"
      Top             =   1980
      Width           =   1450
   End
   Begin VB.OptionButton optSensorInput 
      BackColor       =   &H00FFFFC0&
      Caption         =   "tX axis Pos"
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
      Index           =   37
      Left            =   180
      TabIndex        =   29
      Tag             =   "PATX"
      ToolTipText     =   "Input signal from a position axis (combined proximity value), par=PATX"
      Top             =   1980
      Width           =   1450
   End
   Begin VB.OptionButton optSensorInput 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Inp Pressure"
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
      Left            =   180
      TabIndex        =   22
      Tag             =   "P_I0"
      ToolTipText     =   "Input signal from a pressure sensor, par=P_I0"
      Top             =   6075
      Width           =   1450
   End
   Begin VB.OptionButton optSensorInput 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Pressure 4"
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
      Left            =   1755
      TabIndex        =   21
      Tag             =   "P_I4"
      ToolTipText     =   "Input signal from a pressure sensor, par=P_I4"
      Top             =   6480
      Width           =   1450
   End
   Begin VB.OptionButton optSensorInput 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Pressure 3"
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
      Left            =   3330
      TabIndex        =   20
      Tag             =   "P_I3"
      ToolTipText     =   "Input signal from a pressure sensor, par=P_I3"
      Top             =   6480
      Width           =   1450
   End
   Begin VB.OptionButton optSensorInput 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Pressure 2"
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
      Left            =   3330
      TabIndex        =   19
      Tag             =   "P_I2"
      ToolTipText     =   "Input signal from a pressure sensor, par=P_I2"
      Top             =   6075
      Width           =   1450
   End
   Begin VB.OptionButton optSensorInput 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Pressure 1"
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
      Left            =   1755
      TabIndex        =   18
      Tag             =   "P_I1"
      ToolTipText     =   "Input signal from a pressure sensor, par=P_I1"
      Top             =   6075
      Width           =   1450
   End
   Begin VB.OptionButton optSensorInput 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Velocity Z3"
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
      Index           =   17
      Left            =   3330
      TabIndex        =   17
      Tag             =   "G_Z3"
      ToolTipText     =   "Input signal from a velocity sensor (geophone), par=G_Z3"
      Top             =   5670
      Width           =   1450
   End
   Begin VB.OptionButton optSensorInput 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Velocity Y3"
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
      TabIndex        =   16
      Tag             =   "G_Y3"
      ToolTipText     =   "Input signal from a velocity sensor (geophone), par=G_Y3"
      Top             =   5265
      Width           =   1450
   End
   Begin VB.OptionButton optSensorInput 
      BackColor       =   &H00C0FFFF&
      Caption         =   "X axis Vel"
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
      Index           =   38
      Left            =   180
      TabIndex        =   27
      Tag             =   "GA_X"
      ToolTipText     =   "Input signal from a velocity axis (combined geophones value), par=GA_X"
      Top             =   2385
      Width           =   1450
   End
   Begin VB.OptionButton optSensorInput 
      BackColor       =   &H00C0E0FF&
      Caption         =   "tZ axis Pos"
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
      Index           =   34
      Left            =   3330
      TabIndex        =   24
      Tag             =   "PATZ"
      ToolTipText     =   "Input signal from a position axis (combined proximity value), par=PATZ"
      Top             =   1575
      Width           =   1450
   End
   Begin VB.OptionButton optSensorInput 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Y axis Pos"
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
      Index           =   33
      Left            =   1755
      TabIndex        =   23
      Tag             =   "PA_Y"
      ToolTipText     =   "Input signal from a position axis (combined proximity value), par=PA_Y"
      Top             =   1575
      Width           =   1450
   End
   Begin VB.OptionButton optSensorInput 
      BackColor       =   &H00C0E0FF&
      Caption         =   "X axis Pos "
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
      TabIndex        =   28
      Tag             =   "PA_X"
      ToolTipText     =   "Input signal from a position axis (combined proximity value), par=PA_X"
      Top             =   1575
      Width           =   1450
   End
   Begin VB.OptionButton optSensorInput 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Velocity X3"
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
      Left            =   3330
      TabIndex        =   15
      Tag             =   "G_X3"
      ToolTipText     =   "Input signal from a velocity sensor (geophone), par=G_X3"
      Top             =   4860
      Width           =   1450
   End
   Begin VB.OptionButton optSensorInput 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Velocity Z2"
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
      Index           =   14
      Left            =   1755
      TabIndex        =   14
      Tag             =   "G_Z2"
      ToolTipText     =   "Input signal from a velocity sensor (geophone), par=G_Z2"
      Top             =   5670
      Width           =   1450
   End
   Begin VB.OptionButton optSensorInput 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Velocity Y2"
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
      Left            =   1755
      TabIndex        =   13
      Tag             =   "G_Y2"
      ToolTipText     =   "Input signal from a velocity sensor (geophone), par=G_Y2"
      Top             =   5265
      Width           =   1450
   End
   Begin VB.OptionButton optSensorInput 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Velocity X2"
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
      Left            =   1755
      TabIndex        =   12
      Tag             =   "G_X2"
      ToolTipText     =   "Input signal from a velocity sensor (geophone), par=G_X2"
      Top             =   4860
      Width           =   1450
   End
   Begin VB.OptionButton optSensorInput 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Velocity Z1"
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
      Index           =   11
      Left            =   180
      TabIndex        =   11
      Tag             =   "G_Z1"
      ToolTipText     =   "Input signal from a velocity sensor (geophone), par=G_Z1"
      Top             =   5670
      Width           =   1450
   End
   Begin VB.OptionButton optSensorInput 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Velocity Y1"
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
      Left            =   180
      TabIndex        =   10
      Tag             =   "G_Y1"
      ToolTipText     =   "Input signal from a velocity sensor (geophone), par=G_Y1"
      Top             =   5265
      Width           =   1450
   End
   Begin VB.OptionButton optSensorInput 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Velocity X1"
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
      Left            =   180
      TabIndex        =   9
      Tag             =   "G_X1"
      ToolTipText     =   "Input signal from a velocity sensor (geophone), par=G_X1"
      Top             =   4860
      Width           =   1450
   End
   Begin VB.OptionButton optSensorInput 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Position Z3"
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
      Left            =   3330
      TabIndex        =   8
      Tag             =   "PRZ3"
      ToolTipText     =   "Input signal from a position (proximity) sensor, par=PRZ3"
      Top             =   4410
      Width           =   1450
   End
   Begin VB.OptionButton optSensorInput 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Position Y3"
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
      Left            =   3330
      TabIndex        =   7
      Tag             =   "PRY3"
      ToolTipText     =   "Input signal from a position (proximity) sensor, par=PRY3"
      Top             =   4005
      Width           =   1450
   End
   Begin VB.OptionButton optSensorInput 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Position X3"
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
      Left            =   3330
      TabIndex        =   6
      Tag             =   "PRX3"
      ToolTipText     =   "Input signal from a position (proximity) sensor, par=PRX3"
      Top             =   3600
      Width           =   1450
   End
   Begin VB.OptionButton optSensorInput 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Position Z2"
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
      Left            =   1755
      TabIndex        =   5
      Tag             =   "PRZ2"
      ToolTipText     =   "Input signal from a position (proximity) sensor, par=PRZ1"
      Top             =   4410
      Width           =   1450
   End
   Begin VB.OptionButton optSensorInput 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Position X2"
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
      Left            =   1755
      TabIndex        =   4
      Tag             =   "PRX2"
      ToolTipText     =   "Input signal from a position (proximity) sensor, par=PRX2"
      Top             =   3600
      Width           =   1450
   End
   Begin VB.OptionButton optSensorInput 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Position Z1"
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
      Left            =   180
      TabIndex        =   3
      Tag             =   "PRZ1"
      ToolTipText     =   "Input signal from a position (proximity) sensor, par=PRZ1"
      Top             =   4410
      Width           =   1450
   End
   Begin VB.OptionButton optSensorInput 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Position Y2"
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
      Left            =   1755
      TabIndex        =   2
      Tag             =   "PRY2"
      ToolTipText     =   "Input signal from a position (proximity) sensor, par=PRY2"
      Top             =   4005
      Width           =   1450
   End
   Begin VB.OptionButton optSensorInput 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Position Y1"
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
      Left            =   180
      TabIndex        =   1
      Tag             =   "PRY1"
      ToolTipText     =   "Input signal from a position (proximity) sensor, par=PRY1"
      Top             =   4005
      Width           =   1450
   End
   Begin VB.OptionButton optSensorInput 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Position X1"
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
      Tag             =   "PRX1"
      ToolTipText     =   "Input signal from a position (proximity) sensor, par=PRX1"
      Top             =   3600
      Width           =   1450
   End
   Begin CWUIControlsLib.CWNumEdit cwNumADCchannel 
      Height          =   300
      Left            =   2385
      TabIndex        =   52
      Tag             =   "AD"
      ToolTipText     =   "Input ADC channel number, 0 to 31"
      Top             =   7830
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
      C[0]_1          =   16761024
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
      Value_Val_1     =   31
      IncValueVarType_1=   5
      IncValue_Val_1  =   1
      AccelIncVarType_1=   5
      AccelInc_Val_1  =   0.5
      RangeMinVarType_1=   5
      RangeMaxVarType_1=   5
      RangeMax_Val_1  =   31
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
         Size            =   12
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
      TabIndex        =   50
      ToolTipText     =   "Double-click updates input and closes window"
      Top             =   7785
      Width           =   4875
   End
   Begin VB.Label lblSensors 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Caption         =   "I  N  D  I  V  I  D  U  A  L        S E  N  S  O  R  S"
      ForeColor       =   &H80000008&
      Height          =   4380
      Left            =   45
      TabIndex        =   26
      ToolTipText     =   "Double-click updates input and closes window"
      Top             =   3330
      Width           =   4875
   End
   Begin VB.Label lblAXES 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Caption         =   "S E  N  S  I  N  G      A  X  E  S     I N P U T"
      ForeColor       =   &H80000008&
      Height          =   1905
      Index           =   0
      Left            =   45
      TabIndex        =   25
      ToolTipText     =   "Double-click updates input and closes window"
      Top             =   1305
      Width           =   4875
   End
End
Attribute VB_Name = "frmSensorInput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ADCchannelSelection As Boolean

Private Sub cwNumADCchannel_DblClick()
    Call Leaving_Form
End Sub

Private Sub cwNumADCchannel_ValueChanged(Value As Variant, PreviousValue As Variant, ByVal OutOfRange As Boolean)
    optSensorInput(31).Value = True ' Radio Button "ADC INPUT #" has index 31
    InputChosen = 31 ' index of Radio Button "ADC INPUT #"
    If (SecondInputSelection = False) Then
        Me.Caption = "Choose 1st Test Input. Now = ADC " & Value
    Else ' second input
        Me.Caption = "Choose 2nd Test Input. Now = ADC " & Value
    End If
End Sub

Private Sub Form_Activate()
    Call Form_click
End Sub

Private Sub Form_Load()
    Call Form_click
'    If (SecondInputSelection = False) Then
''        optSensorInput(23).Visible = False ' Radio button "Excitation"
'        If (InputChosen = 31) Then ' index of Radio Button "ADC INPUT #"
'            Me.Caption = "Choose 1st Test Input. Now = ADC" & cwNumADCchannel.Value
'        Else
'            Me.Caption = "Choose 1st Test Input. Now = " & optSensorInput(InputChosen).Caption
'        End If
'    Else ' second input
''        optSensorInput(23).Visible = True ' Radio button "Excitation"
'        If (InputChosen = 31) Then ' index of Radio Button "ADC INPUT #"
'            Me.Caption = "Choose 2nd Test Input. Now = ADC" & cwNumADCchannel.Value
'        Else
'            Me.Caption = "Choose 2nd Test Input. Now = " & optSensorInput(Input2ndChosen).Caption
'        End If
'    End If
End Sub

Private Sub Form_click()
    If (SecondInputSelection = False) Then
'        optSensorInput(23).Visible = False ' Radio button "Excitation"
        If (InputChosen = 31) Then ' index of Radio Button "ADC INPUT #"
            LTFtestInputName = cwNumADCchannel.Tag & Format(cwNumADCchannel.Value, "00")
            Me.Caption = "Choose 1st Test Input. Now = ADC" & Format(cwNumADCchannel.Value, "00")
            Analyzer.CmdChooseInputSignal.Caption = "Input = ADC" & Format(cwNumADCchannel.Value, "00")
        Else
            LTFtestInputName = optSensorInput(InputChosen).Tag
            Me.Caption = "Choose 1st Test Input. Now = " & optSensorInput(InputChosen).Caption
            Analyzer.CmdChooseInputSignal.Caption = "Input = " & optSensorInput(InputChosen).Caption
        End If
    Else ' second input
'        optSensorInput(23).Visible = True ' Radio button "Excitation"
        If (Input2ndChosen = 31) Then ' index of Radio Button "ADC INPUT #"
            LTFtestInputName = cwNumADCchannel.Tag & Format(cwNumADCchannel.Value, "00")
            Me.Caption = "Choose 2nd Test Input. Now = ADC" & Format(cwNumADCchannel.Value, "00")
            Analyzer.CmdChoose2ndInputSignal.Caption = "Inp2 = ADC" & Format(cwNumADCchannel.Value, "00")
        Else
            LTFtestInputName = optSensorInput(InputChosen).Tag
            Me.Caption = "Choose 2nd Test Input. Now = " & optSensorInput(Input2ndChosen).Caption
            Analyzer.CmdChoose2ndInputSignal.Caption = "Inp2 = " & optSensorInput(Input2ndChosen).Caption
        End If
    End If
End Sub

Private Sub Form_LostFocus()
    Call Form_click
    'Call UndateInputs
End Sub

Private Sub Form_Paint() 'calls on each control update, many times
'    Call UndateInputs,
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call Leaving_Form
End Sub

Private Sub optSensorInput_Click(Index As Integer)
    If (Index = 31) Then ' ADC channel selection
        ADCchannelSelection = True
    Else
        ADCchannelSelection = False
    End If
    If (SecondInputSelection <> True) Then
        InputChosen = Index
    Else ' second input
        Input2ndChosen = Index
    End If
    Call Form_click
End Sub

Private Sub optSensorInput_DblClick(Index As Integer)
    If (SecondInputSelection <> True) Then
        InputChosen = Index
    Else ' second input
        Input2ndChosen = Index
    End If
    Call Leaving_Form
End Sub

Sub UndateInputs()
    Dim ch_name As String
    If (ADCchannelSelection = True) Then
        ch_name = "ADC" + cwNumADCchannel.Value
        If (SecondInputSelection <> True) Then
            LTFtestInputName = "AD" & cwNumADCchannel.Value
            Me.Caption = "Choose 1st Test Input. Now = " & ch_name
            Analyzer.CmdChooseInputSignal.Caption = "Input = " & ch_name
        Else ' second input
            LTFtest2ndInputName = ch_name
            Me.Caption = "Choose 2nd Test Input. Now = " & ch_name
            Analyzer.CmdChoose2ndInputSignal.Caption = "Inp2 = " & ch_name
        End If
    Else
        ch_name = optSensorInput(Input2ndChosen).Caption

        If (SecondInputSelection <> True) Then
            LTFtestInputName = optSensorInput(InputChosen).Tag
            Me.Caption = "Choose 1st Test Input. Now = " & optSensorInput(InputChosen).Caption
            If (InputChosen >= 32) Then DiagnosticAxis = InputChosen - 32 ' 32 is the index of "X axis Pos" choice
            Analyzer.CmdChooseInputSignal.Caption = "Input = " & optSensorInput(InputChosen).Caption
        Else ' second input
            LTFtest2ndInputName = optSensorInput(Input2ndChosen).Tag
            Me.Caption = "Choose 2nd Test Input. Now = " & optSensorInput(Input2ndChosen).Caption
            Analyzer.CmdChoose2ndInputSignal.Caption = "Inp2 = " & ch_name
        End If
    End If
End Sub

Sub Leaving_Form()
    Call UndateInputs
    Call Form_click
    Me.Hide
    Call Analyzer.Check_Selected_Actuator_Sensor ' this will load proper limit file OR if selection is wrong, message and hide limit plots
End Sub

Private Sub optSystemType_Click(Index As Integer)
    If Index = 0 Then 'X1-Y2-X3
      If chkHprox.Value = Checked Then 'Horizontal Proximity Present
        optSensorInput(0).Enabled = True   'X1 prox
        optSensorInput(1).Enabled = False  'Y1 prox
        optSensorInput(3).Enabled = False  'X2 prox
        optSensorInput(4).Enabled = True   'Y2 prox
        optSensorInput(6).Enabled = True   'X3 prox
        optSensorInput(7).Enabled = False  'Y3 prox
      Else ' No horiz prox
        optSensorInput(0).Enabled = False   'X1 prox
        optSensorInput(1).Enabled = False   'Y1 prox
        optSensorInput(3).Enabled = False   'X2 prox
        optSensorInput(4).Enabled = False   'Y2 prox
        optSensorInput(6).Enabled = False   'X3 prox
        optSensorInput(7).Enabled = False   'Y3 prox
      End If
      If chkHgeophones.Value = Checked Then 'Horizontal Geophones Present
        optSensorInput(9).Enabled = True    'X1 geo
        optSensorInput(10).Enabled = False  'Y1 geo
        optSensorInput(12).Enabled = False  'X2 geo
        optSensorInput(13).Enabled = True   'Y2 geo
        optSensorInput(15).Enabled = True   'X3 geo
        optSensorInput(16).Enabled = False  'Y3 geo
      Else ' no horiz geo
        optSensorInput(9).Enabled = False   'X1 geo
        optSensorInput(10).Enabled = False  'Y1 geo
        optSensorInput(12).Enabled = False  'X2 geo
        optSensorInput(13).Enabled = False  'Y2 geo
        optSensorInput(15).Enabled = False  'X3 geo
        optSensorInput(16).Enabled = False  'Y3 geo
      End If
   End If

   If Index = 1 Then 'Y1-X2-Y3
      If chkHprox.Value = Checked Then 'Horizontal Proximity Present
        optSensorInput(0).Enabled = False   'X1 prox
        optSensorInput(1).Enabled = True    'Y1 prox
        optSensorInput(3).Enabled = True    'X2 prox
        optSensorInput(4).Enabled = False   'Y2 prox
        optSensorInput(6).Enabled = False   'X3 prox
        optSensorInput(7).Enabled = True    'Y3 prox
      Else
        optSensorInput(0).Enabled = False   'X1 prox
        optSensorInput(1).Enabled = False   'Y1 prox
        optSensorInput(3).Enabled = False   'X2 prox
        optSensorInput(4).Enabled = False   'Y2 prox
        optSensorInput(6).Enabled = False   'X3 prox
        optSensorInput(7).Enabled = False   'Y3 prox
      End If
      If chkHgeophones.Value = Checked Then 'Horizontal Geophones Present
        optSensorInput(9).Enabled = False   'X1 geo
        optSensorInput(10).Enabled = True   'Y1 geo
        optSensorInput(12).Enabled = True   'X2 geo
        optSensorInput(13).Enabled = False  'Y2 geo
        optSensorInput(15).Enabled = False  'X3 geo
        optSensorInput(16).Enabled = True   'Y3 geo
      Else ' no horiz geo
        optSensorInput(9).Enabled = False   'X1 geo
        optSensorInput(10).Enabled = False  'Y1 geo
        optSensorInput(12).Enabled = False  'X2 geo
        optSensorInput(13).Enabled = False  'Y2 geo
        optSensorInput(15).Enabled = False  'X3 geo
        optSensorInput(16).Enabled = False  'Y3 geo
      End If
   End If

   If Index = 2 Then 'Non standard
      If chkHprox.Value = Checked Then 'Horizontal Proximity Present
        optSensorInput(0).Enabled = True  'X1 prox
        optSensorInput(1).Enabled = True  'Y1 prox
        optSensorInput(3).Enabled = True  'X2 prox
        optSensorInput(4).Enabled = True  'Y2 prox
        optSensorInput(6).Enabled = True  'X3 prox
        optSensorInput(7).Enabled = True  'Y3 prox
      Else ' no horiz present
        optSensorInput(0).Enabled = False 'X1 prox
        optSensorInput(1).Enabled = False 'Y1 prox
        optSensorInput(3).Enabled = False 'X2 prox
        optSensorInput(4).Enabled = False 'Y2 prox
        optSensorInput(6).Enabled = False 'X3 prox
        optSensorInput(7).Enabled = False 'Y3 prox
      End If
      If chkHgeophones.Value = Checked Then 'Horizontal Geophones Present
        optSensorInput(9).Enabled = True  'X1 geo
        optSensorInput(10).Enabled = True 'Y1 geo
        optSensorInput(12).Enabled = True 'X2 geo
        optSensorInput(13).Enabled = True 'Y2 geo
        optSensorInput(15).Enabled = True 'X3 geo
        optSensorInput(16).Enabled = True 'Y3 geo
      Else
        optSensorInput(9).Enabled = False  'X1 geo
        optSensorInput(10).Enabled = False 'Y1 geo
        optSensorInput(12).Enabled = False 'X2 geo
        optSensorInput(13).Enabled = False 'Y2 geo
        optSensorInput(15).Enabled = False 'X3 geo
        optSensorInput(16).Enabled = False 'Y3 geo
      End If
    End If

    If chkZgeophones.Value = Checked Then 'Vertical Geophones Present
        optSensorInput(11).Enabled = True 'Z1
        optSensorInput(14).Enabled = True 'Z2
        optSensorInput(17).Enabled = True 'Z3
        optSensorInput(41).Enabled = True 'velocity tX
        optSensorInput(42).Enabled = True 'velocity tY
        optSensorInput(43).Enabled = True 'velocity Z
    Else  'NO Vertical Geophones
        optSensorInput(11).Enabled = False 'Z1
        optSensorInput(14).Enabled = False 'Z2
        optSensorInput(17).Enabled = False 'Z3
        optSensorInput(41).Enabled = False 'velocity tX
        optSensorInput(42).Enabled = False 'velocity tY
        optSensorInput(43).Enabled = False 'velocity Z
    End If
End Sub

Private Sub chkZgeophones_Click()
    If chkZgeophones.Value = Checked Then 'Vertical Geophones Present
        optSensorInput(11).Enabled = True 'Z1
        optSensorInput(14).Enabled = True 'Z2
        optSensorInput(17).Enabled = True 'Z3
        optSensorInput(41).Enabled = True 'velocity tX
        optSensorInput(42).Enabled = True 'velocity tY
        optSensorInput(43).Enabled = True 'velocity Z
    Else  'NO Vertical Geophones
        optSensorInput(11).Enabled = False 'Z1
        optSensorInput(14).Enabled = False 'Z2
        optSensorInput(17).Enabled = False 'Z3
        optSensorInput(41).Enabled = False 'velocity tX
        optSensorInput(42).Enabled = False 'velocity tY
        optSensorInput(43).Enabled = False 'velocity Z
    End If
End Sub

Private Sub chkHgeophones_Click()
    If chkHgeophones.Value = Checked Then 'Horizontal Geophones Present
        optSensorInput(9).Enabled = True  'X1
        optSensorInput(10).Enabled = True 'Y1
        optSensorInput(12).Enabled = True 'X2
        optSensorInput(13).Enabled = True 'Y2
        optSensorInput(15).Enabled = True 'X3
        optSensorInput(16).Enabled = True 'Y3
        optSensorInput(38).Enabled = True 'velocity X
        optSensorInput(39).Enabled = True 'velocity Y
        optSensorInput(40).Enabled = True 'velocity tZ
    Else  'NO horix geophones
        optSensorInput(9).Enabled = False  'X1
        optSensorInput(10).Enabled = False 'Y1
        optSensorInput(12).Enabled = False 'X2
        optSensorInput(13).Enabled = False 'Y2
        optSensorInput(15).Enabled = False 'X3
        optSensorInput(16).Enabled = False 'Y3
        optSensorInput(38).Enabled = False 'velocity X
        optSensorInput(39).Enabled = False 'velocity Y
        optSensorInput(40).Enabled = False 'velocity tZ
    End If
End Sub
Private Sub chkHprox_Click()
    If chkHprox.Value = Checked Then 'Horizontal Proximity Present
        optSensorInput(0).Enabled = True  'X1
        optSensorInput(1).Enabled = True  'Y1
        optSensorInput(3).Enabled = True  'X2
        optSensorInput(4).Enabled = True  'Y2
        optSensorInput(6).Enabled = True  'X3
        optSensorInput(7).Enabled = True  'Y3
        optSensorInput(32).Enabled = True 'position X
        optSensorInput(33).Enabled = True 'position Y
        optSensorInput(34).Enabled = True 'position tZ
    Else  'NO horix geophones
        optSensorInput(0).Enabled = False  'X1
        optSensorInput(1).Enabled = False  'Y1
        optSensorInput(3).Enabled = False  'X2
        optSensorInput(4).Enabled = False  'Y2
        optSensorInput(6).Enabled = False  'X3
        optSensorInput(7).Enabled = False  'Y3
        optSensorInput(32).Enabled = False 'position X
        optSensorInput(33).Enabled = False 'position Y
        optSensorInput(34).Enabled = False 'position tZ
    End If
End Sub

