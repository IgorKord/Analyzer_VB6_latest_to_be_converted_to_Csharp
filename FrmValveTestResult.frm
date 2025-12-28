VERSION 5.00
Object = "{D940E4E4-6079-11CE-88CB-0020AF6845F6}#1.6#0"; "cwui.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form FrmValveTestResult 
   Caption         =   "Valve performance test result"
   ClientHeight    =   5835
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10275
   Icon            =   "FrmValveTestResult.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5835
   ScaleWidth      =   10275
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame 
      Caption         =   "C R I T E R I A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2040
      Left            =   90
      TabIndex        =   55
      Top             =   3780
      Width           =   9015
      Begin VB.CommandButton cmdSaveCriteria 
         BackColor       =   &H00FFC0FF&
         Caption         =   "Save"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   90
         Style           =   1  'Graphical
         TabIndex        =   80
         ToolTipText     =   "Saves test creteria and 'vtst' command parameters"
         Top             =   1080
         Width           =   690
      End
      Begin VB.CommandButton cmdReadCriteria 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Read"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   90
         Style           =   1  'Graphical
         TabIndex        =   79
         ToolTipText     =   "Reads saved test creteria and re-evaluates test"
         Top             =   360
         Width           =   690
      End
      Begin CWUIControlsLib.CWNumEdit cwDeflRateMAX 
         Height          =   300
         Left            =   2565
         TabIndex        =   56
         ToolTipText     =   "Maximum accepted deflation rate; default=1.3 psi/sec from 2500 lbs Iso. If more - there is leak in exhaust"
         Top             =   675
         Width           =   975
         _Version        =   196609
         _ExtentX        =   1720
         _ExtentY        =   529
         _StockProps     =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Reset_0         =   0   'False
         CompatibleVers_0=   196609
         NumEdit_0       =   1
         ClassName_1     =   "CCWNumEdit"
         opts_1          =   458814
         C[0]_1          =   12632319
         C[1]_1          =   0
         ButtonPosition_1=   1
         TextAlignment_1 =   1
         Appearance_1    =   0
         format_1        =   2
         ClassName_2     =   "CCWFormat"
         Format_2        =   ".00"
         scale_1         =   3
         ClassName_3     =   "CCWScale"
         opts_3          =   65536
         dMax_3          =   10
         discInterval_3  =   1
         ValueVarType_1  =   5
         Value_Val_1     =   -1.3
         IncValueVarType_1=   5
         IncValue_Val_1  =   0.02
         AccelIncVarType_1=   5
         AccelInc_Val_1  =   0.2
         RangeMinVarType_1=   5
         RangeMin_Val_1  =   -5
         RangeMaxVarType_1=   5
         Bindings_1      =   4
         ClassName_4     =   "CCWBindingHolderArray"
         Editor_4        =   5
         ClassName_5     =   "CCWBindingHolderArrayEditor"
         Owner_5         =   1
      End
      Begin CWUIControlsLib.CWNumEdit cwDeflRateMIN 
         Height          =   300
         Left            =   2565
         TabIndex        =   57
         ToolTipText     =   "Minimum accepted deflation rate; default=0.8 psi/sec from 2500 lbs Iso. If less - there is restriction in exhaust"
         Top             =   1260
         Width           =   975
         _Version        =   196609
         _ExtentX        =   1720
         _ExtentY        =   529
         _StockProps     =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Reset_0         =   0   'False
         CompatibleVers_0=   196609
         NumEdit_0       =   1
         ClassName_1     =   "CCWNumEdit"
         opts_1          =   458814
         C[0]_1          =   16777152
         C[1]_1          =   0
         ButtonPosition_1=   1
         TextAlignment_1 =   1
         Appearance_1    =   0
         format_1        =   2
         ClassName_2     =   "CCWFormat"
         Format_2        =   ".00"
         scale_1         =   3
         ClassName_3     =   "CCWScale"
         opts_3          =   65536
         dMax_3          =   10
         discInterval_3  =   1
         ValueVarType_1  =   5
         Value_Val_1     =   -0.8
         IncValueVarType_1=   5
         IncValue_Val_1  =   0.02
         AccelIncVarType_1=   5
         AccelInc_Val_1  =   0.2
         RangeMinVarType_1=   5
         RangeMin_Val_1  =   -5
         RangeMaxVarType_1=   5
         Bindings_1      =   4
         ClassName_4     =   "CCWBindingHolderArray"
         Editor_4        =   5
         ClassName_5     =   "CCWBindingHolderArrayEditor"
         Owner_5         =   1
      End
      Begin CWUIControlsLib.CWNumEdit cwPSIvalveOffMAX 
         Height          =   300
         Left            =   3825
         TabIndex        =   58
         ToolTipText     =   "Maximum accepted Iso pressure when valves are OFF; default=2.5 psi. If more - there is leak from input to Iso in the valve"
         Top             =   675
         Width           =   1035
         _Version        =   196609
         _ExtentX        =   1817
         _ExtentY        =   529
         _StockProps     =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Reset_0         =   0   'False
         CompatibleVers_0=   196609
         NumEdit_0       =   1
         ClassName_1     =   "CCWNumEdit"
         opts_1          =   458814
         C[0]_1          =   12632319
         C[1]_1          =   0
         ButtonPosition_1=   1
         TextAlignment_1 =   1
         Appearance_1    =   0
         format_1        =   2
         ClassName_2     =   "CCWFormat"
         Format_2        =   ".0"
         scale_1         =   3
         ClassName_3     =   "CCWScale"
         opts_3          =   65536
         dMax_3          =   10
         discInterval_3  =   1
         ValueVarType_1  =   5
         Value_Val_1     =   2.5
         IncValueVarType_1=   5
         IncValue_Val_1  =   0.1
         AccelIncVarType_1=   5
         AccelInc_Val_1  =   0.3
         RangeMinVarType_1=   5
         RangeMaxVarType_1=   5
         RangeMax_Val_1  =   5
         Bindings_1      =   4
         ClassName_4     =   "CCWBindingHolderArray"
         Editor_4        =   5
         ClassName_5     =   "CCWBindingHolderArrayEditor"
         Owner_5         =   1
      End
      Begin CWUIControlsLib.CWNumEdit cwValveOpeningMINmA 
         Height          =   300
         Left            =   5175
         TabIndex        =   59
         ToolTipText     =   "Minimum accepted valve's opening current; default=90 mA. If less - either input pressure is too low or defective valve"
         Top             =   1260
         Width           =   1035
         _Version        =   196609
         _ExtentX        =   1817
         _ExtentY        =   529
         _StockProps     =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Reset_0         =   0   'False
         CompatibleVers_0=   196609
         NumEdit_0       =   1
         ClassName_1     =   "CCWNumEdit"
         opts_1          =   458814
         C[0]_1          =   16777152
         C[1]_1          =   0
         ButtonPosition_1=   1
         TextAlignment_1 =   1
         Appearance_1    =   0
         format_1        =   2
         ClassName_2     =   "CCWFormat"
         Format_2        =   ".0"
         scale_1         =   3
         ClassName_3     =   "CCWScale"
         opts_3          =   65536
         dMax_3          =   10
         discInterval_3  =   1
         ValueVarType_1  =   5
         Value_Val_1     =   90
         IncValueVarType_1=   5
         IncValue_Val_1  =   1
         AccelIncVarType_1=   5
         AccelInc_Val_1  =   5
         RangeMinVarType_1=   5
         RangeMin_Val_1  =   65
         RangeMaxVarType_1=   5
         RangeMax_Val_1  =   95
         Bindings_1      =   4
         ClassName_4     =   "CCWBindingHolderArray"
         Editor_4        =   5
         ClassName_5     =   "CCWBindingHolderArrayEditor"
         Owner_5         =   1
      End
      Begin CWUIControlsLib.CWNumEdit cwValveOpeningMAXmA 
         Height          =   300
         Left            =   5175
         TabIndex        =   60
         ToolTipText     =   "Maximum accepted valve's opening current; default=110 mA. If more- either input pressure is too high or defective valve"
         Top             =   675
         Width           =   1035
         _Version        =   196609
         _ExtentX        =   1817
         _ExtentY        =   529
         _StockProps     =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Reset_0         =   0   'False
         CompatibleVers_0=   196609
         NumEdit_0       =   1
         ClassName_1     =   "CCWNumEdit"
         opts_1          =   458814
         C[0]_1          =   12632319
         C[1]_1          =   0
         ButtonPosition_1=   1
         TextAlignment_1 =   1
         Appearance_1    =   0
         format_1        =   2
         ClassName_2     =   "CCWFormat"
         Format_2        =   ".0"
         scale_1         =   3
         ClassName_3     =   "CCWScale"
         opts_3          =   65536
         dMax_3          =   10
         discInterval_3  =   1
         ValueVarType_1  =   5
         Value_Val_1     =   110
         IncValueVarType_1=   5
         IncValue_Val_1  =   1
         AccelIncVarType_1=   5
         AccelInc_Val_1  =   5
         RangeMinVarType_1=   5
         RangeMin_Val_1  =   75
         RangeMaxVarType_1=   5
         RangeMax_Val_1  =   115
         Bindings_1      =   4
         ClassName_4     =   "CCWBindingHolderArray"
         Editor_4        =   5
         ClassName_5     =   "CCWBindingHolderArrayEditor"
         Owner_5         =   1
      End
      Begin CWUIControlsLib.CWNumEdit cwInflationRateMAX 
         Height          =   300
         Left            =   6570
         TabIndex        =   61
         ToolTipText     =   $"FrmValveTestResult.frx":030A
         Top             =   675
         Width           =   855
         _Version        =   196609
         _ExtentX        =   1508
         _ExtentY        =   529
         _StockProps     =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Reset_0         =   0   'False
         CompatibleVers_0=   196609
         NumEdit_0       =   1
         ClassName_1     =   "CCWNumEdit"
         opts_1          =   458814
         C[0]_1          =   12632319
         C[1]_1          =   0
         ButtonPosition_1=   1
         TextAlignment_1 =   1
         Appearance_1    =   0
         format_1        =   2
         ClassName_2     =   "CCWFormat"
         Format_2        =   ".0"
         scale_1         =   3
         ClassName_3     =   "CCWScale"
         opts_3          =   65536
         dMax_3          =   10
         discInterval_3  =   1
         ValueVarType_1  =   5
         Value_Val_1     =   5
         IncValueVarType_1=   5
         IncValue_Val_1  =   0.1
         AccelIncVarType_1=   5
         AccelInc_Val_1  =   0.3
         RangeMinVarType_1=   5
         RangeMaxVarType_1=   5
         RangeMax_Val_1  =   10
         Bindings_1      =   4
         ClassName_4     =   "CCWBindingHolderArray"
         Editor_4        =   5
         ClassName_5     =   "CCWBindingHolderArrayEditor"
         Owner_5         =   1
      End
      Begin CWUIControlsLib.CWNumEdit cwInflationRateMIN 
         Height          =   300
         Left            =   6570
         TabIndex        =   62
         ToolTipText     =   $"FrmValveTestResult.frx":03C7
         Top             =   1260
         Width           =   855
         _Version        =   196609
         _ExtentX        =   1508
         _ExtentY        =   529
         _StockProps     =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Reset_0         =   0   'False
         CompatibleVers_0=   196609
         NumEdit_0       =   1
         ClassName_1     =   "CCWNumEdit"
         opts_1          =   458814
         C[0]_1          =   16777152
         C[1]_1          =   0
         ButtonPosition_1=   1
         TextAlignment_1 =   1
         Appearance_1    =   0
         format_1        =   2
         ClassName_2     =   "CCWFormat"
         Format_2        =   ".0"
         scale_1         =   3
         ClassName_3     =   "CCWScale"
         opts_3          =   65536
         dMax_3          =   10
         discInterval_3  =   1
         ValueVarType_1  =   5
         Value_Val_1     =   3
         IncValueVarType_1=   5
         IncValue_Val_1  =   0.1
         AccelIncVarType_1=   5
         AccelInc_Val_1  =   0.3
         RangeMinVarType_1=   5
         RangeMaxVarType_1=   5
         RangeMax_Val_1  =   10
         Bindings_1      =   4
         ClassName_4     =   "CCWBindingHolderArray"
         Editor_4        =   5
         ClassName_5     =   "CCWBindingHolderArrayEditor"
         Owner_5         =   1
      End
      Begin CWUIControlsLib.CWNumEdit cwPSIvalveFullyOpenMIN 
         Height          =   300
         Left            =   7785
         TabIndex        =   63
         ToolTipText     =   $"FrmValveTestResult.frx":0494
         Top             =   1260
         Width           =   1035
         _Version        =   196609
         _ExtentX        =   1817
         _ExtentY        =   529
         _StockProps     =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Reset_0         =   0   'False
         CompatibleVers_0=   196609
         NumEdit_0       =   1
         ClassName_1     =   "CCWNumEdit"
         opts_1          =   458814
         C[0]_1          =   16777152
         C[1]_1          =   0
         ButtonPosition_1=   1
         TextAlignment_1 =   1
         Appearance_1    =   0
         format_1        =   2
         ClassName_2     =   "CCWFormat"
         Format_2        =   ".0"
         scale_1         =   3
         ClassName_3     =   "CCWScale"
         opts_3          =   65536
         dMax_3          =   10
         discInterval_3  =   1
         ValueVarType_1  =   5
         Value_Val_1     =   -6
         IncValueVarType_1=   5
         IncValue_Val_1  =   1
         AccelIncVarType_1=   5
         AccelInc_Val_1  =   5
         RangeMinVarType_1=   5
         RangeMin_Val_1  =   -50
         RangeMaxVarType_1=   5
         RangeMax_Val_1  =   -5
         Bindings_1      =   4
         ClassName_4     =   "CCWBindingHolderArray"
         Editor_4        =   5
         ClassName_5     =   "CCWBindingHolderArrayEditor"
         Owner_5         =   1
      End
      Begin CWUIControlsLib.CWNumEdit cwPSIvalveFullyOpenMAX 
         Height          =   300
         Left            =   7785
         TabIndex        =   64
         ToolTipText     =   $"FrmValveTestResult.frx":055C
         Top             =   675
         Width           =   1035
         _Version        =   196609
         _ExtentX        =   1817
         _ExtentY        =   529
         _StockProps     =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Reset_0         =   0   'False
         CompatibleVers_0=   196609
         NumEdit_0       =   1
         ClassName_1     =   "CCWNumEdit"
         opts_1          =   458814
         C[0]_1          =   12632319
         C[1]_1          =   0
         ButtonPosition_1=   1
         TextAlignment_1 =   1
         Appearance_1    =   0
         format_1        =   2
         ClassName_2     =   "CCWFormat"
         Format_2        =   ".0"
         scale_1         =   3
         ClassName_3     =   "CCWScale"
         opts_3          =   65536
         dMax_3          =   10
         discInterval_3  =   1
         ValueVarType_1  =   5
         Value_Val_1     =   -20
         IncValueVarType_1=   5
         IncValue_Val_1  =   1
         AccelIncVarType_1=   5
         AccelInc_Val_1  =   5
         RangeMinVarType_1=   5
         RangeMin_Val_1  =   -50
         RangeMaxVarType_1=   5
         RangeMax_Val_1  =   -10
         Bindings_1      =   4
         ClassName_4     =   "CCWBindingHolderArray"
         Editor_4        =   5
         ClassName_5     =   "CCWBindingHolderArrayEditor"
         Owner_5         =   1
      End
      Begin CWUIControlsLib.CWNumEdit cwInpPsiMAX 
         Height          =   300
         Left            =   1485
         TabIndex        =   76
         ToolTipText     =   "Input pressure @ valves shut - should be standard test pressure 90+-5psi."
         Top             =   675
         Width           =   795
         _Version        =   196609
         _ExtentX        =   1402
         _ExtentY        =   529
         _StockProps     =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Reset_0         =   0   'False
         CompatibleVers_0=   196609
         NumEdit_0       =   1
         ClassName_1     =   "CCWNumEdit"
         opts_1          =   458814
         C[0]_1          =   12632319
         C[1]_1          =   0
         ButtonPosition_1=   1
         TextAlignment_1 =   2
         Appearance_1    =   0
         format_1        =   2
         ClassName_2     =   "CCWFormat"
         Format_2        =   "0."
         scale_1         =   3
         ClassName_3     =   "CCWScale"
         opts_3          =   65536
         dMax_3          =   10
         discInterval_3  =   1
         ValueVarType_1  =   5
         Value_Val_1     =   95
         IncValueVarType_1=   5
         IncValue_Val_1  =   1
         AccelIncVarType_1=   5
         AccelInc_Val_1  =   5
         RangeMinVarType_1=   5
         RangeMin_Val_1  =   20
         RangeMaxVarType_1=   5
         RangeMax_Val_1  =   111
         Bindings_1      =   4
         ClassName_4     =   "CCWBindingHolderArray"
         Editor_4        =   5
         ClassName_5     =   "CCWBindingHolderArrayEditor"
         Owner_5         =   1
      End
      Begin CWUIControlsLib.CWNumEdit cwInpPsiMIN 
         Height          =   300
         Left            =   1485
         TabIndex        =   77
         ToolTipText     =   "Input pressure @ valves shut - should be standard test pressure 90+-5psi."
         Top             =   1260
         Width           =   795
         _Version        =   196609
         _ExtentX        =   1402
         _ExtentY        =   529
         _StockProps     =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Reset_0         =   0   'False
         CompatibleVers_0=   196609
         NumEdit_0       =   1
         ClassName_1     =   "CCWNumEdit"
         opts_1          =   458814
         C[0]_1          =   16777152
         C[1]_1          =   0
         ButtonPosition_1=   1
         TextAlignment_1 =   2
         Appearance_1    =   0
         format_1        =   2
         ClassName_2     =   "CCWFormat"
         Format_2        =   "."
         scale_1         =   3
         ClassName_3     =   "CCWScale"
         opts_3          =   65536
         dMax_3          =   10
         discInterval_3  =   1
         ValueVarType_1  =   5
         Value_Val_1     =   85
         IncValueVarType_1=   5
         IncValue_Val_1  =   1
         AccelIncVarType_1=   5
         AccelInc_Val_1  =   5
         RangeMinVarType_1=   5
         RangeMin_Val_1  =   20
         RangeMaxVarType_1=   5
         RangeMax_Val_1  =   110
         Bindings_1      =   4
         ClassName_4     =   "CCWBindingHolderArray"
         Editor_4        =   5
         ClassName_5     =   "CCWBindingHolderArrayEditor"
         Owner_5         =   1
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
         Caption         =   "Good between"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   270
         Index           =   30
         Left            =   3735
         TabIndex        =   81
         Top             =   990
         Width           =   2715
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Input psi"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   270
         Index           =   29
         Left            =   1440
         TabIndex        =   78
         Top             =   405
         Width           =   870
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
         Caption         =   "Good between"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   270
         Index           =   28
         Left            =   6345
         TabIndex        =   75
         Top             =   990
         Width           =   2490
      End
      Begin VB.Label lblMsg 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "message"
         Height          =   285
         Left            =   135
         TabIndex        =   74
         Top             =   1665
         Width           =   8700
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Caption         =   "min:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   300
         Index           =   21
         Left            =   945
         TabIndex        =   73
         Top             =   1260
         Width           =   555
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
         Caption         =   "max:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   300
         Index           =   12
         Left            =   945
         TabIndex        =   72
         Top             =   675
         Width           =   555
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
         Caption         =   "Good between"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   270
         Index           =   13
         Left            =   945
         TabIndex        =   71
         Top             =   990
         Width           =   2760
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "no min"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   300
         Index           =   14
         Left            =   3825
         TabIndex        =   70
         Top             =   1260
         Width           =   1050
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "psi drop vs. input"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   495
         Index           =   15
         Left            =   7830
         TabIndex        =   69
         Top             =   180
         Width           =   870
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Inflation psi / mA"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   495
         Index           =   17
         Left            =   6570
         TabIndex        =   68
         Top             =   180
         Width           =   870
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Deflation psi / sec"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   495
         Index           =   16
         Left            =   2610
         TabIndex        =   67
         Top             =   180
         Width           =   870
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Valve OFF psi"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   495
         Index           =   18
         Left            =   3825
         TabIndex        =   66
         Top             =   180
         Width           =   870
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Opening mA"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   495
         Index           =   19
         Left            =   5220
         TabIndex        =   65
         Top             =   180
         Width           =   870
      End
   End
   Begin VB.CommandButton CmdRecalculate 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Re - Evaluate"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   4365
      Style           =   1  'Graphical
      TabIndex        =   35
      ToolTipText     =   "Change test Min - Max criteria and re evaluate"
      Top             =   180
      Width           =   1680
   End
   Begin MSComDlg.CommonDialog dlgCommonDialog 
      Left            =   135
      Top             =   945
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin CWUIControlsLib.CWNumEdit CwStandardSlope 
      Height          =   300
      Left            =   9225
      TabIndex        =   94
      ToolTipText     =   "Desiered valve slope, PSI/mA"
      Top             =   855
      Width           =   990
      _Version        =   196609
      _ExtentX        =   1746
      _ExtentY        =   529
      _StockProps     =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Reset_0         =   0   'False
      CompatibleVers_0=   196609
      NumEdit_0       =   1
      ClassName_1     =   "CCWNumEdit"
      opts_1          =   458814
      C[0]_1          =   16777215
      C[1]_1          =   0
      ButtonPosition_1=   1
      TextAlignment_1 =   1
      Appearance_1    =   0
      format_1        =   2
      ClassName_2     =   "CCWFormat"
      Format_2        =   ".00"
      scale_1         =   3
      ClassName_3     =   "CCWScale"
      opts_3          =   65536
      dMax_3          =   10
      discInterval_3  =   1
      ValueVarType_1  =   5
      Value_Val_1     =   4
      IncValueVarType_1=   5
      IncValue_Val_1  =   0.02
      AccelIncVarType_1=   5
      AccelInc_Val_1  =   0.3
      RangeMinVarType_1=   5
      RangeMin_Val_1  =   1
      RangeMaxVarType_1=   5
      RangeMax_Val_1  =   10
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
      BackStyle       =   0  'Transparent
      Caption         =   "psi/mA"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   36
      Left            =   9225
      TabIndex        =   96
      Top             =   540
      Width           =   870
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Standard Inflation"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   480
      Index           =   34
      Left            =   9180
      TabIndex        =   95
      Top             =   90
      Width           =   1005
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Rate Correction"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   540
      Index           =   33
      Left            =   9180
      TabIndex        =   93
      Top             =   1305
      Width           =   1050
   End
   Begin VB.Label lbl_SlopeCorrection 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "88.8"
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
      Height          =   285
      Index           =   0
      Left            =   9225
      TabIndex        =   92
      ToolTipText     =   $"FrmValveTestResult.frx":0621
      Top             =   1890
      Width           =   750
   End
   Begin VB.Label lbl_SlopeCorrection 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "88.8"
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
      Height          =   285
      Index           =   1
      Left            =   9225
      TabIndex        =   91
      ToolTipText     =   $"FrmValveTestResult.frx":06C0
      Top             =   2385
      Width           =   750
   End
   Begin VB.Label lbl_SlopeCorrection 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "88.8"
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
      Height          =   285
      Index           =   2
      Left            =   9225
      TabIndex        =   90
      ToolTipText     =   $"FrmValveTestResult.frx":075F
      Top             =   2880
      Width           =   750
   End
   Begin VB.Label lbl_SlopeCorrection 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "88.8"
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
      Height          =   285
      Index           =   3
      Left            =   9225
      TabIndex        =   89
      ToolTipText     =   $"FrmValveTestResult.frx":07FE
      Top             =   3375
      Width           =   750
   End
   Begin VB.Label Lbl_V4sn 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "44444"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   300
      Left            =   1170
      TabIndex        =   88
      Top             =   3375
      Width           =   735
   End
   Begin VB.Label Lbl_V3sn 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "3333"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   300
      Left            =   1170
      TabIndex        =   87
      Top             =   2880
      Width           =   735
   End
   Begin VB.Label Lbl_V2sn 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "V2 SN"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   300
      Left            =   1170
      TabIndex        =   86
      Top             =   2385
      Width           =   735
   End
   Begin VB.Label Lbl_V1sn 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "V1 SN"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   300
      Left            =   1170
      TabIndex        =   85
      Top             =   1890
      Width           =   735
   End
   Begin VB.Label lbl_MaxInpPSI_ValvesClosed 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "95"
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
      Height          =   285
      Left            =   1935
      TabIndex        =   84
      ToolTipText     =   "Input pressure @ valves shut - should be standard test pressure +-5psi."
      Top             =   180
      Width           =   600
   End
   Begin VB.Label lbl_MinInpPSI_ValvesClosed 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "75"
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
      Height          =   285
      Left            =   1170
      TabIndex        =   83
      ToolTipText     =   "Input pressure @ valves shut - should be standard test pressure +-5psi."
      Top             =   180
      Width           =   600
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0080FF80&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "--"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   31
      Left            =   1755
      TabIndex        =   82
      Top             =   180
      Width           =   195
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0080FF80&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "--"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   27
      Left            =   8280
      TabIndex        =   54
      Top             =   180
      Width           =   195
   End
   Begin VB.Label lbl_MinInpPSI_at_FullyOpen 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "63.88"
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
      Height          =   285
      Left            =   7695
      TabIndex        =   53
      ToolTipText     =   $"FrmValveTestResult.frx":089D
      Top             =   180
      Width           =   600
   End
   Begin VB.Label lbl_MaxInpPSI_at_FullyOpen 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "67.77"
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
      Height          =   285
      Left            =   8460
      TabIndex        =   52
      ToolTipText     =   $"FrmValveTestResult.frx":094A
      Top             =   180
      Width           =   600
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0080FF80&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "--"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   26
      Left            =   6840
      TabIndex        =   51
      Top             =   1395
      Width           =   195
   End
   Begin VB.Label lbl_minInflSlope 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "3.0"
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
      Height          =   285
      Left            =   6390
      TabIndex        =   50
      ToolTipText     =   $"FrmValveTestResult.frx":09FB
      Top             =   1395
      Width           =   465
   End
   Begin VB.Label lbl_maxInflSlope 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "5.0"
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
      Height          =   285
      Left            =   7020
      TabIndex        =   49
      ToolTipText     =   $"FrmValveTestResult.frx":0AAC
      Top             =   1395
      Width           =   465
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0080FF80&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "--"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   25
      Left            =   5400
      TabIndex        =   48
      Top             =   1395
      Width           =   195
   End
   Begin VB.Label lbl_minOpening_mA 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "90.0"
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
      Height          =   285
      Left            =   4815
      TabIndex        =   47
      ToolTipText     =   $"FrmValveTestResult.frx":0B5D
      Top             =   1395
      Width           =   600
   End
   Begin VB.Label lbl_maxOpening_mA 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "110.0"
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
      Height          =   285
      Left            =   5580
      TabIndex        =   46
      ToolTipText     =   $"FrmValveTestResult.frx":0C0E
      Top             =   1395
      Width           =   600
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0080FF80&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "--"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   24
      Left            =   3960
      TabIndex        =   45
      Top             =   1395
      Width           =   195
   End
   Begin VB.Label lbl_minPSI_off 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.0"
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
      Height          =   285
      Left            =   3510
      TabIndex        =   44
      ToolTipText     =   $"FrmValveTestResult.frx":0CBF
      Top             =   1395
      Width           =   465
   End
   Begin VB.Label lbl_maxPSI_off 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "2.5"
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
      Height          =   285
      Left            =   4140
      TabIndex        =   43
      ToolTipText     =   $"FrmValveTestResult.frx":0D70
      Top             =   1395
      Width           =   465
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0080FF80&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "--"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   23
      Left            =   2565
      TabIndex        =   42
      Top             =   1395
      Width           =   195
   End
   Begin VB.Label lbl_minDeflSlope 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "-0.8"
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
      Height          =   285
      Left            =   2025
      TabIndex        =   41
      ToolTipText     =   $"FrmValveTestResult.frx":0E21
      Top             =   1395
      Width           =   555
   End
   Begin VB.Label lbl_maxDeflSlope 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "-1.3"
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
      Height          =   285
      Left            =   2745
      TabIndex        =   40
      ToolTipText     =   $"FrmValveTestResult.frx":0ED2
      Top             =   1395
      Width           =   555
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "RANGES"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   270
      Index           =   22
      Left            =   1125
      TabIndex        =   39
      Top             =   1395
      Width           =   870
   End
   Begin VB.Label lbl_maxPSI_at_FULLY_OPEN 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "-20.0"
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
      Height          =   285
      Left            =   8460
      TabIndex        =   37
      ToolTipText     =   $"FrmValveTestResult.frx":0F83
      Top             =   1395
      Width           =   600
   End
   Begin VB.Label lbl_minPSI_at_FULLY_OPEN 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "-6.0"
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
      Height          =   285
      Left            =   7695
      TabIndex        =   36
      ToolTipText     =   $"FrmValveTestResult.frx":1034
      Top             =   1395
      Width           =   600
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0080FF80&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "--"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   20
      Left            =   8280
      TabIndex        =   38
      Top             =   1395
      Width           =   195
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Test Delay, s"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Index           =   11
      Left            =   2970
      TabIndex        =   34
      Top             =   45
      Width           =   960
   End
   Begin VB.Label LblDelay 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "2.22"
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
      Height          =   285
      Left            =   3015
      TabIndex        =   33
      ToolTipText     =   $"FrmValveTestResult.frx":10E5
      Top             =   540
      Width           =   870
   End
   Begin VB.Label lbl_Opening_Current 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "88.8"
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
      Height          =   285
      Index           =   3
      Left            =   4815
      TabIndex        =   32
      ToolTipText     =   "Valve opening current - measured when pressure reaches 5 PSI. It can vary due to valve assembly"
      Top             =   3375
      Width           =   1365
   End
   Begin VB.Label lbl_PSI_at_FULLY_OPEN 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "88.8"
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
      Height          =   285
      Index           =   3
      Left            =   7695
      TabIndex        =   31
      ToolTipText     =   $"FrmValveTestResult.frx":117C
      Top             =   3375
      Width           =   1365
   End
   Begin VB.Label lbl_InflationSlope_PSI_per_mA 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "88.8"
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
      Height          =   285
      Index           =   3
      Left            =   6390
      TabIndex        =   30
      ToolTipText     =   $"FrmValveTestResult.frx":122D
      Top             =   3375
      Width           =   1095
   End
   Begin VB.Label lbl_PSI_deflated 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "88.8"
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
      Height          =   285
      Index           =   3
      Left            =   3510
      TabIndex        =   29
      ToolTipText     =   $"FrmValveTestResult.frx":12CA
      Top             =   3375
      Width           =   1095
   End
   Begin VB.Label lbl_DeflationSlope_PSI_per_sec 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "88.8"
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
      Height          =   285
      Index           =   3
      Left            =   2025
      TabIndex        =   28
      ToolTipText     =   "Deflation rate, PSI / sec -  characterizes exhaust rate. Measured when valve is shut and pressure drops from 40 to 20 psi"
      Top             =   3375
      Width           =   1275
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Valve 4 SN"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   270
      Index           =   10
      Left            =   90
      TabIndex        =   27
      Top             =   3375
      Width           =   1215
   End
   Begin VB.Label lbl_Opening_Current 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "88.8"
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
      Height          =   285
      Index           =   2
      Left            =   4815
      TabIndex        =   26
      ToolTipText     =   "Valve opening current - measured when pressure reaches 5 PSI. It can vary due to valve assembly"
      Top             =   2880
      Width           =   1365
   End
   Begin VB.Label lbl_PSI_at_FULLY_OPEN 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "88.8"
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
      Height          =   285
      Index           =   2
      Left            =   7695
      TabIndex        =   25
      ToolTipText     =   $"FrmValveTestResult.frx":1381
      Top             =   2880
      Width           =   1365
   End
   Begin VB.Label lbl_InflationSlope_PSI_per_mA 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "88.8"
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
      Height          =   285
      Index           =   2
      Left            =   6390
      TabIndex        =   24
      ToolTipText     =   $"FrmValveTestResult.frx":1432
      Top             =   2880
      Width           =   1095
   End
   Begin VB.Label lbl_PSI_deflated 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "88.8"
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
      Height          =   285
      Index           =   2
      Left            =   3510
      TabIndex        =   23
      ToolTipText     =   $"FrmValveTestResult.frx":14CF
      Top             =   2880
      Width           =   1095
   End
   Begin VB.Label lbl_DeflationSlope_PSI_per_sec 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "88.8"
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
      Height          =   285
      Index           =   2
      Left            =   2025
      TabIndex        =   22
      ToolTipText     =   "Deflation rate, PSI / sec -  characterizes exhaust rate. Measured when valve is shut and pressure drops from 40 to 20 psi"
      Top             =   2880
      Width           =   1275
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Valve 3 SN"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   270
      Index           =   9
      Left            =   90
      TabIndex        =   21
      Top             =   2880
      Width           =   1215
   End
   Begin VB.Label lbl_Opening_Current 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "88.8"
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
      Height          =   285
      Index           =   1
      Left            =   4815
      TabIndex        =   20
      ToolTipText     =   "Valve opening current - measured when pressure reaches 5 PSI. It can vary due to valve assembly"
      Top             =   2385
      Width           =   1365
   End
   Begin VB.Label lbl_PSI_at_FULLY_OPEN 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "88.8"
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
      Height          =   285
      Index           =   1
      Left            =   7695
      TabIndex        =   19
      ToolTipText     =   $"FrmValveTestResult.frx":1586
      Top             =   2385
      Width           =   1365
   End
   Begin VB.Label lbl_InflationSlope_PSI_per_mA 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "88.8"
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
      Height          =   285
      Index           =   1
      Left            =   6390
      TabIndex        =   18
      ToolTipText     =   $"FrmValveTestResult.frx":1637
      Top             =   2385
      Width           =   1095
   End
   Begin VB.Label lbl_PSI_deflated 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "88.8"
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
      Height          =   285
      Index           =   1
      Left            =   3510
      TabIndex        =   17
      ToolTipText     =   $"FrmValveTestResult.frx":16D4
      Top             =   2385
      Width           =   1095
   End
   Begin VB.Label lbl_DeflationSlope_PSI_per_sec 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "88.8"
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
      Height          =   285
      Index           =   1
      Left            =   2025
      TabIndex        =   16
      ToolTipText     =   "Deflation rate, PSI / sec -  characterizes exhaust rate. Measured when valve is shut and pressure drops from 40 to 20 psi"
      Top             =   2385
      Width           =   1275
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Valve 2 SN"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   270
      Index           =   8
      Left            =   90
      TabIndex        =   15
      Top             =   2385
      Width           =   1215
   End
   Begin VB.Label lbl_Opening_Current 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "88.8"
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
      Height          =   285
      Index           =   0
      Left            =   4815
      TabIndex        =   14
      ToolTipText     =   "Valve opening current - measured when pressure reaches 5 PSI. It can vary due to valve assembly"
      Top             =   1890
      Width           =   1365
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Opening current mA"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   540
      Index           =   7
      Left            =   4815
      TabIndex        =   13
      ToolTipText     =   "Valve opening current - measured when pressure reaches 5 PSI. It can vary due to valve assembly"
      Top             =   855
      Width           =   1230
   End
   Begin VB.Label lbl_PSI_at_FULLY_OPEN 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "88.8"
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
      Height          =   285
      Index           =   0
      Left            =   7695
      TabIndex        =   12
      ToolTipText     =   $"FrmValveTestResult.frx":178B
      Top             =   1890
      Width           =   1365
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "PSI @ valve FULLY OPEN"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   540
      Index           =   6
      Left            =   7695
      TabIndex        =   11
      ToolTipText     =   $"FrmValveTestResult.frx":183C
      Top             =   900
      Width           =   1365
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Inflation rate PSI / mA"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Index           =   5
      Left            =   6300
      TabIndex        =   10
      ToolTipText     =   $"FrmValveTestResult.frx":18ED
      Top             =   900
      Width           =   1230
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "PSI when valve OFF"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   540
      Index           =   4
      Left            =   3510
      TabIndex        =   9
      ToolTipText     =   $"FrmValveTestResult.frx":198A
      Top             =   855
      Width           =   1050
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Deflation rate PSI / sec"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   540
      Index           =   3
      Left            =   1980
      TabIndex        =   8
      ToolTipText     =   "Deflation rate, PSI / sec -  characterizes exhaust rate. Measured when valve is shut and pressure drops from 40 to 20 psi"
      Top             =   855
      Width           =   1320
   End
   Begin VB.Label lbl_InflationSlope_PSI_per_mA 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "88.8"
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
      Height          =   285
      Index           =   0
      Left            =   6390
      TabIndex        =   7
      ToolTipText     =   $"FrmValveTestResult.frx":1A41
      Top             =   1890
      Width           =   1095
   End
   Begin VB.Label lbl_PSI_deflated 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "88.8"
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
      Height          =   285
      Index           =   0
      Left            =   3510
      TabIndex        =   6
      ToolTipText     =   $"FrmValveTestResult.frx":1ADE
      Top             =   1890
      Width           =   1095
   End
   Begin VB.Label lbl_DeflationSlope_PSI_per_sec 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "noTimeInfo"
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
      Height          =   285
      Index           =   0
      Left            =   2025
      TabIndex        =   5
      ToolTipText     =   "Deflation rate, PSI / sec -  characterizes exhaust rate. Measured when valve is shut and pressure drops from 40 to 20 psi"
      Top             =   1890
      Width           =   1275
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Valve 1 SN"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   270
      Index           =   2
      Left            =   90
      TabIndex        =   4
      Top             =   1890
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Input PSI valves OFF"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Index           =   0
      Left            =   -45
      TabIndex        =   3
      ToolTipText     =   "Input pressure @ valves shut - should be standard test pressure 90 PSI"
      Top             =   180
      Width           =   1185
   End
   Begin VB.Label lbl_Inp_PSI_valves_FULL_ON 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "88.8"
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
      Height          =   285
      Left            =   7695
      TabIndex        =   2
      ToolTipText     =   $"FrmValveTestResult.frx":1B95
      Top             =   540
      Width           =   1365
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Input PSI valves ON"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   585
      Index           =   1
      Left            =   6615
      TabIndex        =   1
      ToolTipText     =   $"FrmValveTestResult.frx":1C2C
      Top             =   180
      Width           =   1050
   End
   Begin VB.Label lbl_Inp_PSI_valves_OFF 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "-NA-"
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
      Height          =   285
      Left            =   1170
      TabIndex        =   0
      ToolTipText     =   "Input pressure @ valves shut - should be standard test pressure 90 PSI"
      Top             =   540
      Width           =   1365
   End
End
Attribute VB_Name = "FrmValveTestResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const NAstring As String = "-NA-"
Dim criteria_loaded As Boolean
Dim CriteriaMAXinputPSI As Single
Dim CriteriaMINinputPSI As Single
Dim CriteriaMAXdeflationSLOPE As Single
Dim CriteriaMINdeflationSLOPE As Single
Dim CriteriaMAXremainingPressure As Single
Dim CriteriaMAXopeningCurrent As Single
Dim CriteriaMINopeningCurrent As Single
Dim CriteriaMAXinflationSLOPE As Single
Dim CriteriaMINinflationSLOPE As Single
Dim CriteriaMAXvalveOPENpsi As Single
Dim CriteriaMINvalveOPENpsi As Single

Const Criteria_file As Integer = 4
Dim ParamLINEScount As Long




Private Sub CwStandardSlope_ValueChanged(Value As Variant, PreviousValue As Variant, ByVal OutOfRange As Boolean)
    frmMonitoring.Calc_SlopeCorrection
End Sub

Sub Form_Load()
    CriteriaMAXinputPSI = cwInpPsiMAX.Value '95
    CriteriaMINinputPSI = cwInpPsiMIN.Value '85
    CriteriaMAXdeflationSLOPE = cwDeflRateMAX.Value
    CriteriaMINdeflationSLOPE = cwDeflRateMIN.Value
    CriteriaMAXremainingPressure = cwPSIvalveOffMAX.Value
    CriteriaMAXopeningCurrent = cwValveOpeningMAXmA.Value
    CriteriaMINopeningCurrent = cwValveOpeningMINmA.Value
    CriteriaMAXinflationSLOPE = cwInflationRateMAX.Value
    CriteriaMINinflationSLOPE = cwInflationRateMIN.Value
    CriteriaMAXvalveOPENpsi = cwPSIvalveFullyOpenMAX.Value
    CriteriaMINvalveOPENpsi = cwPSIvalveFullyOpenMIN.Value

    Call setDefaults
    criteria_loaded = False
End Sub

Sub setDefaults()
    Dim indx As Long
    lbl_Inp_PSI_valves_OFF.Caption = NAstring
    lbl_Inp_PSI_valves_FULL_ON.Caption = NAstring
    LblDelay.Caption = NAstring
    For indx = 0 To 3
        lbl_DeflationSlope_PSI_per_sec(indx).Caption = NAstring
        lbl_PSI_deflated(indx).Caption = NAstring
        lbl_Opening_Current(indx).Caption = NAstring
        lbl_InflationSlope_PSI_per_mA(indx).Caption = NAstring
        lbl_PSI_at_FULLY_OPEN(indx).Caption = NAstring
    Next indx
End Sub

'Private Sub CmdDefault_Click()
'    Dim captn As String
'    cwDeflRateMAX.Value = CriteriaMAXdeflationSLOPE
'    cwDeflRateMIN.Value = CriteriaMINdeflationSLOPE
'    cwPSIvalveOffMAX.Value = CriteriaMAXremainingPressure
'    cwValveOpeningMAXmA.Value = CriteriaMAXopeningCurrent
'    cwValveOpeningMINmA.Value = CriteriaMINopeningCurrent
'    cwInflationRateMAX.Value = CriteriaMAXinflationSLOPE
'    cwInflationRateMIN.Value = CriteriaMINinflationSLOPE
'    cwPSIvalveFullyOpenMAX.Value = CriteriaMAXvalveOPENpsi
'    cwPSIvalveFullyOpenMIN.Value = CriteriaMINvalveOPENpsi
'
''    captn = LblDelay.Caption
''    If IsNumber(captn) Then
''        If Val(LblDelay.Caption) <= 2 Then ' at 89 psi in
''            cwInflationRateMIN.Value = 2
''            cwInflationRateMAX.Value = 3.6
''' at 80 psi in, 2 sec delay rate is 2.35 to 2.6
''' at 88 psi in, 2 sec delay rate is 2.75 to 3.0
''' at 90 psi in, 2 sec delay rate is 3.0
''        Else '>2, 3 sec
''' at 80 psi in, 3 sec delay rate is 3.07 - 3.45
''' at 88 psi in, 3 sec delay rate is 3.32 - 3.66
''' at 90 psi in, 3 sec delay rate is 3.65 - 3.91
''            cwInflationRateMIN.Value = 2.5
''            cwInflationRateMAX.Value = 4.5
''        End If
''    Else ' not a number
''            cwInflationRateMIN.Value = 2
''            cwInflationRateMAX.Value = 5
''    End If
''    criteria_loaded = True ' allow user to change them
'End Sub

Private Sub CmdRecalculate_Click()
    Call frmMonitoring.Analyze_Valve_plot
End Sub

Private Sub cwInpPsiMAX_ValueChanged(Value As Variant, PreviousValue As Variant, ByVal OutOfRange As Boolean)
    lbl_MaxInpPSI_ValvesClosed.Caption = cwInpPsiMAX.Text
    CriteriaMAXinputPSI = cwInpPsiMAX.Value
End Sub

Private Sub cwInpPsiMIN_ValueChanged(Value As Variant, PreviousValue As Variant, ByVal OutOfRange As Boolean)
    lbl_MinInpPSI_ValvesClosed.Caption = cwInpPsiMIN.Text
    CriteriaMINinputPSI = cwInpPsiMIN.Value
End Sub

Private Sub cwDeflRateMAX_ValueChanged(Value As Variant, PreviousValue As Variant, ByVal OutOfRange As Boolean)
    lbl_maxDeflSlope.Caption = cwDeflRateMAX.Text
    CriteriaMAXdeflationSLOPE = cwDeflRateMAX.Value
End Sub

Private Sub cwDeflRateMIN_ValueChanged(Value As Variant, PreviousValue As Variant, ByVal OutOfRange As Boolean)
    lbl_minDeflSlope.Caption = cwDeflRateMIN.Text
    CriteriaMINdeflationSLOPE = cwDeflRateMIN.Value
End Sub

Private Sub cwPSIvalveOffMAX_ValueChanged(Value As Variant, PreviousValue As Variant, ByVal OutOfRange As Boolean)
    lbl_maxPSI_off.Caption = cwPSIvalveOffMAX.Text
    CriteriaMAXremainingPressure = cwPSIvalveOffMAX.Value
End Sub

Private Sub cwValveOpeningMAXmA_ValueChanged(Value As Variant, PreviousValue As Variant, ByVal OutOfRange As Boolean)
    lbl_maxOpening_mA.Caption = cwValveOpeningMAXmA.Text
    CriteriaMAXopeningCurrent = cwValveOpeningMAXmA.Value
End Sub

Private Sub cwValveOpeningMINmA_ValueChanged(Value As Variant, PreviousValue As Variant, ByVal OutOfRange As Boolean)
    lbl_minOpening_mA.Caption = cwValveOpeningMINmA.Text
    CriteriaMINopeningCurrent = cwValveOpeningMINmA.Value
End Sub

Private Sub cwInflationRateMAX_ValueChanged(Value As Variant, PreviousValue As Variant, ByVal OutOfRange As Boolean)
    lbl_maxInflSlope.Caption = cwInflationRateMAX.Text
    CriteriaMAXinflationSLOPE = cwInflationRateMAX.Value
End Sub

Private Sub cwInflationRateMIN_ValueChanged(Value As Variant, PreviousValue As Variant, ByVal OutOfRange As Boolean)
    lbl_minInflSlope.Caption = cwInflationRateMIN.Text
    CriteriaMINinflationSLOPE = cwInflationRateMIN.Value
End Sub

Private Sub cwPSIvalveFullyOpenMAX_ValueChanged(Value As Variant, PreviousValue As Variant, ByVal OutOfRange As Boolean)
    Dim new_val As Single
    new_val = Val(lbl_Inp_PSI_valves_FULL_ON.Caption) + cwPSIvalveFullyOpenMAX.Value
    lbl_maxPSI_at_FULLY_OPEN.Caption = new_val
    'CriteriaMAXvalveOPENpsi = new_val
End Sub

Private Sub cwPSIvalveFullyOpenMIN_ValueChanged(Value As Variant, PreviousValue As Variant, ByVal OutOfRange As Boolean)
    Dim new_val As Single
    new_val = Val(lbl_Inp_PSI_valves_FULL_ON.Caption) + cwPSIvalveFullyOpenMIN.Value
    lbl_minPSI_at_FULLY_OPEN.Caption = new_val
    'CriteriaMINvalveOPENpsi = new_val
End Sub

Private Sub lbl_Inp_PSI_valves_OFF_Change()
    Dim captn As String
    Dim inp_pressure As Single
    captn = lbl_Inp_PSI_valves_OFF.Caption
    If captn = NAstring Then
        lbl_Inp_PSI_valves_OFF.BackColor = &HFFFFC0 ' blue
    Else
        If IsNumber(captn) Then
           inp_pressure = Val(captn)
            lbl_MaxInpPSI_ValvesClosed.Caption = cwInpPsiMAX.Text
            lbl_MinInpPSI_ValvesClosed.Caption = cwInpPsiMIN.Text
                If inp_pressure > cwInpPsiMIN.Value And inp_pressure < cwInpPsiMAX.Value Then
                    lbl_Inp_PSI_valves_OFF.BackColor = &H80FF80 'green
                Else
                    lbl_Inp_PSI_valves_OFF.BackColor = &HC0FFFF 'yellow &HFFFFC0 'blue
                End If
        End If
    End If
End Sub

Private Sub lbl_DeflationSlope_PSI_per_sec_Change(Index As Integer)
    Dim captn As String
    Dim slope As Single
    captn = lbl_DeflationSlope_PSI_per_sec(Index).Caption
    If captn = "noTimeInfo" Or captn = NAstring Then
        lbl_DeflationSlope_PSI_per_sec(Index).BackColor = &HFFFFC0 ' blue
    Else
        If IsNumber(captn) Then
            slope = Val(captn) ' negative
            lbl_minDeflSlope.Caption = cwDeflRateMIN.Text
            lbl_maxDeflSlope.Caption = cwDeflRateMAX.Text
            If slope < cwDeflRateMIN.Value And slope > cwDeflRateMAX.Value Then
                lbl_DeflationSlope_PSI_per_sec(Index).BackColor = &H80FF80 'green
            Else
                lbl_DeflationSlope_PSI_per_sec(Index).BackColor = &HC0FFFF 'yellow &HFFFFC0 'blue
            End If
        End If
    End If
End Sub

Private Sub lbl_PSI_deflated_Change(Index As Integer)
    Dim captn As String
    Dim min_pressure As Single
    captn = lbl_PSI_deflated(Index).Caption
    If captn = NAstring Then
        lbl_PSI_deflated(Index).BackColor = &HFFFFC0 ' blue
    Else
        If IsNumber(captn) Then
            min_pressure = Val(captn)
            lbl_minPSI_off.Caption = "0.0"
            lbl_maxPSI_off.Caption = cwPSIvalveOffMAX.Text
                If min_pressure < cwPSIvalveOffMAX.Value Then
                    lbl_PSI_deflated(Index).BackColor = &H80FF80 'green
                Else
                    lbl_PSI_deflated(Index).BackColor = &HC0FFFF 'yellow &HFFFFC0 'blue
                End If
        End If
    End If
End Sub

Private Sub lbl_Opening_Current_Change(Index As Integer)
    Dim captn As String
    Dim opening_mA As Single
    captn = lbl_Opening_Current(Index).Caption
    If captn = NAstring Then
        lbl_Opening_Current(Index).BackColor = &HFFFFC0 ' blue
    Else
        If IsNumber(captn) Then
            opening_mA = Val(captn)
            lbl_minOpening_mA = cwValveOpeningMINmA.Text
            lbl_maxOpening_mA = cwValveOpeningMAXmA.Text
            If opening_mA > cwValveOpeningMINmA.Value And opening_mA < cwValveOpeningMAXmA.Value Then
                lbl_Opening_Current(Index).BackColor = &H80FF80 ' green
            Else
                lbl_Opening_Current(Index).BackColor = &HC0FFFF ' yellow
            End If
        End If
    End If
End Sub

Private Sub lbl_InflationSlope_PSI_per_mA_Change(Index As Integer)
    Dim captn As String
    Dim slope As Single
    captn = lbl_InflationSlope_PSI_per_mA(Index).Caption
    If captn = NAstring Then
        lbl_InflationSlope_PSI_per_mA(Index).BackColor = &HFFFFC0 ' blue
    Else
        If IsNumber(captn) Then
            slope = Val(captn)
            lbl_minInflSlope.Caption = cwInflationRateMIN.Text
            lbl_maxInflSlope.Caption = cwInflationRateMAX.Text
            If slope > cwInflationRateMIN.Value And slope < cwInflationRateMAX.Value Then
                lbl_InflationSlope_PSI_per_mA(Index).BackColor = &H80FF80 'green
            Else
                lbl_InflationSlope_PSI_per_mA(Index).BackColor = &HC0FFFF 'yellow &HFFFFC0 'blue
            End If
        End If
    End If
End Sub

Private Sub lbl_PSI_at_FULLY_OPEN_Change(Index As Integer)
    Dim captn As String
    Dim ValveFullyOpenPSI As Single
    Dim calculated_InPSI As Single
    Dim min_psi As Single
    Dim max_psi As Single
    captn = lbl_PSI_at_FULLY_OPEN(Index).Caption
    If captn = NAstring Or lbl_Inp_PSI_valves_FULL_ON.Caption = NAstring Then
        lbl_PSI_at_FULLY_OPEN(Index).BackColor = &HFFFFC0 ' blue
    Else
        If IsNumber(captn) And IsNumber(lbl_Inp_PSI_valves_FULL_ON.Caption) Then
            ValveFullyOpenPSI = Val(captn)
            calculated_InPSI = Val(lbl_Inp_PSI_valves_FULL_ON.Caption)
            min_psi = calculated_InPSI + cwPSIvalveFullyOpenMAX.Value
            max_psi = calculated_InPSI + cwPSIvalveFullyOpenMIN.Value
            lbl_minPSI_at_FULLY_OPEN.Caption = Format(min_psi, "0.00")
            lbl_maxPSI_at_FULLY_OPEN.Caption = Format(max_psi, "0.00")

            If ValveFullyOpenPSI > min_psi And ValveFullyOpenPSI < max_psi Then
                lbl_PSI_at_FULLY_OPEN(Index).BackColor = &H80FF80 ' green
            Else
                lbl_PSI_at_FULLY_OPEN(Index).BackColor = &HC0FFFF ' yellow
            End If
        End If
    End If
End Sub

Private Sub lbl_Inp_PSI_valves_FULL_ON_Change()
    Dim captn As String
    Dim pressure As Single
    Dim min_pressure As Single
    Dim in_pressure As Single
    captn = lbl_Inp_PSI_valves_FULL_ON.Caption
    If captn = NAstring Then
        lbl_Inp_PSI_valves_FULL_ON.BackColor = &HFFFFC0  ' blue
    Else
        If IsNumber(captn) And IsNumber(lbl_Inp_PSI_valves_OFF.Caption) Then
            pressure = Val(captn)
            ' update criteria based on input pressure
            'at 88 psi,opening current is up 108 to 113 mA
            'at 85 psi,opening current is 108 mA
            'at 80 psi,opening current is 105 mA
            'at 78 psi,opening current is 103 mA
            ' 105 + (pressure - 80)*1.5
'            cwValveOpeningMAXmA.Value = 105 + (pressure - 80) * 1.5 'PSI to mA cheating conversion
'            cwValveOpeningMINmA.Value = cwValveOpeningMAXmA.Value - 20
'            If (pressure >= 85) Then
'                cwValveOpeningMAXmA.Value = 115
'                cwValveOpeningMINmA.Value = 90
'            Else
'                If ((pressure > 75) And (pressure < 85)) Then
'                    cwValveOpeningMAXmA.Value = 110
'                    cwValveOpeningMINmA.Value = 85
'                End If
'            End If
            in_pressure = Val(lbl_Inp_PSI_valves_OFF.Caption)
            min_pressure = in_pressure - 10
            lbl_MinInpPSI_at_FullyOpen.Caption = Format(min_pressure, "0.00")
            lbl_MaxInpPSI_at_FullyOpen.Caption = Format(in_pressure, "0.00")
            If pressure > min_pressure Then ' allowed drop 10 psi
                lbl_Inp_PSI_valves_FULL_ON.BackColor = &H80FF80 'green
            Else
                lbl_Inp_PSI_valves_FULL_ON.BackColor = &HC0FFFF 'yellow &HFFFFC0 'blue
            End If
        End If
    End If
End Sub


Private Sub cmdReadCriteria_Click()
   Call ReadCriteriaFile
End Sub

Sub ReadCriteriaFile()
    On Error Resume Next
    Dim cont_field_monitoring As Boolean
    Dim DefaultFilename As String
    Dim DataFileName As String
    Dim FileName As String
    Dim i As Long
    Dim temp_str As String
    Dim val_str As String
    Dim sep_pos As Long
    Dim line_read As Long
    Dim line_No_str As String
    Dim line_error_number As Long
    Dim Msg As String, Style As Long, Title As String, answer As Long
    Dim search_str As String

   DefaultFilename = "DC-2020 valve test criteria " & Date$ & " " & Left$(Time$, 2) & "-" & Mid$(Time$, 4, 2) & ".criteria"
    dlgCommonDialog.CancelError = True
    dlgCommonDialog.DialogTitle = "Choose name for DC-2020 parameters to load"
    ' dlgCommonDialog.FileName = DefaultFilename
    dlgCommonDialog.InitDir = "c:\test"
    dlgCommonDialog.DefaultExt = "criteria"
    dlgCommonDialog.Filter = "Data File (*.criteria)|*.criteria|All files (*.*)|*.*"
    dlgCommonDialog.Action = 1
    'dlgCommonDialog.ShowSave
    If Err = 32755 Then 'Cancel is selected
        Exit Sub
    Else
        lblMsg.Caption = "Reading Saved Criteria file '" & dlgCommonDialog.FileName & " ....."
    End If
    'SaveFileName becomes "C:\test\my_test.txt"

    DataFileName = dlgCommonDialog.FileName
    'SaveFileName becomes "C:\test\my_test"
    i = InStrRev(DataFileName, "\")
    FileName = Right(DataFileName, Len(DataFileName) - i) 'i.e. "my_test"

    i = InStrRev(dlgCommonDialog.FileName, "\")
    dlgCommonDialog.InitDir = Left(dlgCommonDialog.FileName, i - 1)
    DataFileName = dlgCommonDialog.FileName
    'now lets open the file and read the data

    line_read = 0
    Open DataFileName For Input As Criteria_file 'Values

    Line Input #Criteria_file, temp_str 'Print #Criteria_file, "File name, "; FileName
    ' decode line
    sep_pos = InStr(temp_str, ",")
    If InStr(LCase(temp_str), "file name") = 0 Then
        GoTo error_format
    Else
        search_str = Mid(temp_str, sep_pos + 1)
    End If
' next line is "DC-2020 valve test criteria"
    Line Input #Criteria_file, temp_str '    Print #Criteria_file, "DC-2020 parameters"
    ' decode line
    sep_pos = InStr(temp_str, "2020")
    If sep_pos = 0 Then
        GoTo error_format
    End If
    Line Input #Criteria_file, temp_str '    Print #Criteria_file, "Date, "; Date$
    Line Input #Criteria_file, temp_str '    Print #Criteria_file, "Time, "; Time$


' load parameters
    Msg = ""

'//// File name, DC-2020 valve test criteria 09-19-2017 00-19.criteria
'//// DC-2020 valve test criteria
'//// Date, 09-19-2017
'//// Time, 00:57:14,
'////'vtst' command parameters
'vtst_start_curr 70  // mA
'vtst_increment  0.5 // mA
'vtst_end_curr   140 // mA
'vtst_step_sec   3   // sec
'vtst_hold_curr  60  // mA
'slope_high_curr 30  // mA slope measurement upper threshold current
'slope_low_curr  15  // mA slope measurement lower threshold current
'//// MAX & MIN input pressure PSI affect 'opening Current','remaining pressure','inflation SLOPE','valve OPEN PSI'
'max_input_psi    95     // PSI      MAX input pressure
'min_input_psi    85     // PSI      MIN input pressure
'max_defl_slope  -1.3    // PSI/sec  MAX deflation SLOPE
'min_defl_slope  -0.8    // PSI/sec  MIN deflation SLOPE
'max_closed_psi   2.5    // PSI      Valve closed MAX remaining pressure
'max_opening_ma   110    // mA       MAX opening current
'min_opening_ma   90     // mA       MIN opening current
'max_infl_slope   5  // PSI/mA   MAX inflation SLOPE
'min_infl_slope   3  // PSI/mA   MIN inflation SLOPE
'max_open_psi    -20     // PSI      Valve fully OPEN MAX pressure
'max_open_psi    -6  // PSI      Valve fully OPEN MIN pressure

read_again:
    Do While (Not EOF(Criteria_file))  ' Loop until end of file.

        Line Input #Criteria_file, temp_str '   read settings line
        line_read = line_read + 1

        sep_pos = InStr(temp_str, "////") 'find comment
        ' it can be comment WITHOUT command, like //// version
        If sep_pos <> 0 Then 'comment
            GoTo read_again
        End If
        sep_pos = InStr(temp_str, "//") 'find comment
        ' or comment AFTER command
        If sep_pos <> 0 Then 'comment
            temp_str = Left(temp_str, sep_pos - 1) 'cut comment
        End If
        If (Len(temp_str) > 0) Then 'it was criteria with comment, parse
            search_str = "vtst_start_curr" '====
            sep_pos = InStr(temp_str, search_str)
            If (sep_pos <> 0) Then
               val_str = Mid(temp_str, Len(search_str) + sep_pos)
               If (IsNumber(val_str)) Then frmMonitoring.cwValveStartCurrent.Value = CDbl(val_str)
            End If

            search_str = "vtst_increment" '====
            sep_pos = InStr(temp_str, search_str)
            If (sep_pos <> 0) Then
               val_str = Mid(temp_str, Len(search_str) + sep_pos)
               If (IsNumber(val_str)) Then frmMonitoring.cwValveIncrementCurrent.Value = CDbl(val_str)
            End If

            search_str = "vtst_end_curr" '===
            sep_pos = InStr(temp_str, search_str)
            If (sep_pos <> 0) Then
               val_str = Mid(temp_str, Len(search_str) + sep_pos)
               If (IsNumber(val_str)) Then frmMonitoring.cwValveEndCurrent.Value = CDbl(val_str)
            End If

            search_str = "vtst_step_sec" '===
            sep_pos = InStr(temp_str, search_str)
            If (sep_pos <> 0) Then
               val_str = Mid(temp_str, Len(search_str) + sep_pos)
               If (IsNumber(val_str)) Then frmMonitoring.cwValveMeasTime.Value = CDbl(val_str)
            End If

            search_str = "vtst_hold_curr" '====
            sep_pos = InStr(temp_str, search_str)
            If (sep_pos <> 0) Then
               val_str = Mid(temp_str, Len(search_str) + sep_pos)
               If (IsNumber(val_str)) Then frmMonitoring.cwValveHoldCurrent.Value = CDbl(val_str)
            End If

            search_str = "max_input_psi" '-----
            sep_pos = InStr(temp_str, search_str)
            If (sep_pos <> 0) Then
               val_str = Mid(temp_str, Len(search_str) + sep_pos)
               If (IsNumber(val_str)) Then cwInpPsiMAX.Value = CDbl(val_str) 'cwInpPsiMAX_ValueChanged updates : CriteriaMAXinputPSI = CDbl(val_str)
            End If

            search_str = "min_input_psi" '-----
            sep_pos = InStr(temp_str, search_str)
            If (sep_pos <> 0) Then
               val_str = Mid(temp_str, Len(search_str) + sep_pos)
               If (IsNumber(val_str)) Then cwInpPsiMIN.Value = CDbl(val_str) ': CriteriaMINinputPSI = CDbl(val_str)
            End If

            search_str = "max_defl_slope" '===
            sep_pos = InStr(temp_str, search_str)
            If (sep_pos <> 0) Then
               val_str = Mid(temp_str, Len(search_str) + sep_pos)
               If (IsNumber(val_str)) Then cwDeflRateMAX.Value = CDbl(val_str) ': CriteriaMAXdeflationSLOPE = CDbl(val_str):
            End If

            search_str = "min_defl_slope" '===
            sep_pos = InStr(temp_str, search_str)
            If (sep_pos <> 0) Then
               val_str = Mid(temp_str, Len(search_str) + sep_pos)
               If (IsNumber(val_str)) Then cwDeflRateMIN.Value = CDbl(val_str) ': CriteriaMINdeflationSLOPE = CDbl(val_str):
            End If

            search_str = "max_closed_psi" '===
            sep_pos = InStr(temp_str, search_str)
            If (sep_pos <> 0) Then
               val_str = Mid(temp_str, Len(search_str) + sep_pos)
               If (IsNumber(val_str)) Then cwPSIvalveOffMAX.Value = CDbl(val_str) ': CriteriaMAXremainingPressure = CDbl(val_str):
            End If

            search_str = "max_opening_ma" '===
            sep_pos = InStr(temp_str, search_str)
            If (sep_pos <> 0) Then
               val_str = Mid(temp_str, Len(search_str) + sep_pos)
               If (IsNumber(val_str)) Then cwValveOpeningMAXmA.Value = CDbl(val_str) ': CriteriaMAXopeningCurrent = CDbl(val_str):
            End If

            search_str = "min_opening_ma" '===
            sep_pos = InStr(temp_str, search_str)
            If (sep_pos <> 0) Then
               val_str = Mid(temp_str, Len(search_str) + sep_pos)
               If (IsNumber(val_str)) Then cwValveOpeningMINmA.Value = CDbl(val_str) ': CriteriaMINopeningCurrent = CDbl(val_str):
            End If

            search_str = "max_infl_slope" '===
            sep_pos = InStr(temp_str, search_str)
            If (sep_pos <> 0) Then
               val_str = Mid(temp_str, Len(search_str) + sep_pos)
               If (IsNumber(val_str)) Then cwInflationRateMAX.Value = CDbl(val_str) ': CriteriaMAXinflationSLOPE = CDbl(val_str)
            End If

            search_str = "min_infl_slope" '===
            sep_pos = InStr(temp_str, search_str)
            If (sep_pos <> 0) Then
               val_str = Mid(temp_str, Len(search_str) + sep_pos)
               If (IsNumber(val_str)) Then cwInflationRateMIN.Value = CDbl(val_str) ': CriteriaMINinflationSLOPE = CDbl(val_str)
            End If

            search_str = "max_open_psi" '===
            sep_pos = InStr(temp_str, search_str)
            If (sep_pos <> 0) Then
               val_str = Mid(temp_str, Len(search_str) + sep_pos)
               If (IsNumber(val_str)) Then cwPSIvalveFullyOpenMAX.Value = CDbl(val_str) ': CriteriaMAXvalveOPENpsi = CDbl(val_str)
            End If

            search_str = "min_open_psi" '===
            sep_pos = InStr(temp_str, search_str)
            If (sep_pos <> 0) Then
               val_str = Mid(temp_str, Len(search_str) + sep_pos)
               If (IsNumber(val_str)) Then cwPSIvalveFullyOpenMIN.Value = CDbl(val_str) ': CriteriaMINvalveOPENpsi = CDbl(val_str)
            End If

'    Print #Criteria_file, "slope_high_curr"; frmMonitoring.CwUpperPSIthreshold.Value; "// mA slope measurement upper threshold current"
'    Print #Criteria_file, "slope_low_curr"; frmMonitoring.CwLowerPSIthreshold.Value; "// mA slope measurement lower threshold current"
            search_str = "slope_high_curr" '
            sep_pos = InStr(temp_str, search_str)
            If (sep_pos <> 0) Then
               val_str = Mid(temp_str, Len(search_str) + sep_pos)
               If (IsNumber(val_str)) Then frmMonitoring.CwUpperPSIthreshold.Value = CDbl(val_str) ': CriteriaMINvalveOPENpsi = CDbl(val_str)
            End If

            search_str = "slope_low_curr" '
            sep_pos = InStr(temp_str, search_str)
            If (sep_pos <> 0) Then
               val_str = Mid(temp_str, Len(search_str) + sep_pos)
               If (IsNumber(val_str)) Then frmMonitoring.CwLowerPSIthreshold.Value = CDbl(val_str) ': CriteriaMINvalveOPENpsi = CDbl(val_str)
            End If
        End If
    Loop
    Close #Criteria_file
    criteria_loaded = True
    ' re-evaluate test
    Call frmMonitoring.Analyze_Valve_plot

    Exit Sub
error_format:
    Close #Criteria_file
    MsgBox ("Unrecognizable file format")
End Sub

Private Sub cmdSaveCriteria_Click()
    Call SaveCriteriaFile
End Sub
Sub SaveCriteriaFile()
    On Error Resume Next
    'user wants to save a test file
    'First get the filename
    Dim DefaultFilename As String
    Dim SaveFileName As String
    Dim FileName As String
    Dim i As Long
    Dim colon_pos As Long
    Dim param_line As String


   DefaultFilename = "DC-2020 valve test criteria " & Date$ & " " & Left$(Time$, 2) & "-" & Mid$(Time$, 4, 2) & ".criteria"
    dlgCommonDialog.CancelError = True
    dlgCommonDialog.DialogTitle = "Choose name for DC-2020 valve test criteria"
    ' dlgCommonDialog.FileName = DefaultFilename
    dlgCommonDialog.InitDir = "c:\test"
    dlgCommonDialog.DefaultExt = "criteria"
    dlgCommonDialog.Filter = "Data File (*.criteria)|*.criteria|All files (*.*)|*.*"
    dlgCommonDialog.FileName = DefaultFilename 'if enabled,  FILE NAME SHOULD NOT HAVE ':'
    dlgCommonDialog.ShowSave
    If Err = 32755 Then 'Cancel is selected
        Exit Sub
    Else
        lblMsg.Caption = "File=" & dlgCommonDialog.FileName
    End If
    'SaveFileName becomes "C:\test\my_test.txt"

'    SaveFileName = Left(dlgCommonDialog.FileName, Len(dlgCommonDialog.FileName) - 4) ' remove ".txt" from full file path
    SaveFileName = dlgCommonDialog.FileName
    'SaveFileName becomes "C:\test\my_test"
    i = InStrRev(SaveFileName, "\")
    FileName = Right(SaveFileName, Len(SaveFileName) - i) 'i.e. "my_test"

    SaveFileName = dlgCommonDialog.FileName
    i = InStrRev(dlgCommonDialog.FileName, "\")
    dlgCommonDialog.InitDir = Left(dlgCommonDialog.FileName, i - 1)

    'now lets open the file and write the data
    Open SaveFileName For Output As Criteria_file 'parameters
    Print #Criteria_file, "//// File name, "; FileName
    Print #Criteria_file, "//// DC-2020 valve test criteria"
    Print #Criteria_file, "//// Date, "; Date$
    Print #Criteria_file, "//// Time, "; Time$; ", "

'-    Print #Criteria_file, textParams.Text
    Print #Criteria_file, "////'vtst' command parameters"
    Print #Criteria_file, "vtst_start_curr"; vbTab; CStr(frmMonitoring.cwValveStartCurrent.Value); vbTab; "// mA"
    Print #Criteria_file, "vtst_increment"; vbTab; CStr(frmMonitoring.cwValveIncrementCurrent.Value); vbTab; "// mA"
    Print #Criteria_file, "vtst_end_curr"; vbTab; CStr(frmMonitoring.cwValveEndCurrent.Value); vbTab; "// mA"
    Print #Criteria_file, "vtst_step_sec"; vbTab; CStr(frmMonitoring.cwValveMeasTime.Value); vbTab; "// sec"
    Print #Criteria_file, "vtst_hold_curr"; vbTab; CStr(frmMonitoring.cwValveHoldCurrent.Value); vbTab; "// mA"
    Print #Criteria_file, "slope_high_curr"; vbTab; frmMonitoring.CwUpperPSIthreshold.Value; "// mA slope measurement upper threshold current"
    Print #Criteria_file, "slope_low_curr"; vbTab; frmMonitoring.CwLowerPSIthreshold.Value; "// mA slope measurement lower threshold current"

    Print #Criteria_file, "//// MAX & MIN input pressure PSI affect 'opening Current','remaining pressure','inflation SLOPE','valve OPEN PSI'"
    Print #Criteria_file, "max_input_psi"; vbTab; CriteriaMAXinputPSI; vbTab; "// PSI   "; vbTab; "MAX input pressure"
    Print #Criteria_file, "min_input_psi"; vbTab; CriteriaMINinputPSI; vbTab; "// PSI   "; vbTab; "MIN input pressure"
    Print #Criteria_file, "max_defl_slope"; vbTab; CriteriaMAXdeflationSLOPE; vbTab; "// PSI/sec"; vbTab; "MAX deflation SLOPE"
    Print #Criteria_file, "min_defl_slope"; vbTab; CriteriaMINdeflationSLOPE; vbTab; "// PSI/sec"; vbTab; "MIN deflation SLOPE"
    Print #Criteria_file, "max_closed_psi"; vbTab; CriteriaMAXremainingPressure; vbTab; "// PSI   "; vbTab; "Valve closed MAX remaining pressure"
    Print #Criteria_file, "max_opening_ma"; vbTab; CriteriaMAXopeningCurrent; vbTab; "// mA    "; vbTab; "MAX opening current"
    Print #Criteria_file, "min_opening_ma"; vbTab; CriteriaMINopeningCurrent; vbTab; "// mA    "; vbTab; "MIN opening current"
    Print #Criteria_file, "max_infl_slope"; vbTab; CriteriaMAXinflationSLOPE; vbTab; "// PSI/mA"; vbTab; "MAX inflation SLOPE"
    Print #Criteria_file, "min_infl_slope"; vbTab; CriteriaMINinflationSLOPE; vbTab; "// PSI/mA"; vbTab; "MIN inflation SLOPE"
    Print #Criteria_file, "max_open_psi"; vbTab; CriteriaMAXvalveOPENpsi; vbTab; "// PSI   "; vbTab; "Valve fully OPEN MAX pressure"
    Print #Criteria_file, "min_open_psi"; vbTab; CriteriaMINvalveOPENpsi; vbTab; "// PSI   "; vbTab; "Valve fully OPEN MIN pressure"

    Close #Criteria_file
End Sub

