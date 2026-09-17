VERSION 5.00
Object = "{D940E4E4-6079-11CE-88CB-0020AF6845F6}#1.6#0"; "cwui.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Analyzer
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E0E0E0&
   Caption         =   "TMC Analyzer"
   ClientHeight    =   8805
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   14610
   Icon            =   "Analyzer.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   12720.88
   ScaleMode       =   0  'User
   ScaleWidth      =   15518.83
   Begin VB.Frame FramePerformance
      BackColor       =   &H00FFFFC0&
      Caption         =   "Performance"
      BeginProperty Font
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8355
      Left            =   2745
      TabIndex        =   152
      Top             =   0
      Width           =   11760
      Begin VB.CheckBox ChkUseAsStartTFfreq
         Caption         =   "TF start"
         Height          =   285
         Left            =   9600
         TabIndex        =   314
         ToolTipText     =   "When checked, this value is used ias start frequency for Transfer Function command"
         Top             =   4200
         Width           =   915
      End
      Begin CWUIControlsLib.CWNumEdit cwNumMaxFreq
         Height          =   360
         Left            =   10530
         TabIndex        =   196
         ToolTipText     =   "Max Frequency, 100...5000 (default = 1000) Hz"
         Top             =   4140
         Width           =   1110
         _Version        =   524288
         _ExtentX        =   1958
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
         C[0]_1          =   16777215
         C[1]_1          =   0
         BorderStyle_1   =   1
         TextAlignment_1 =   2
         format_1        =   2
         ClassName_2     =   "CCWFormat"
         Format_2        =   ".###"
         scale_1         =   3
         ClassName_3     =   "CCWScale"
         opts_3          =   65536
         dMax_3          =   10
         discInterval_3  =   1
         ValueVarType_1  =   5
         Value_Val_1     =   1000
         IncValueVarType_1=   5
         IncValue_Val_1  =   500
         AccelIncVarType_1=   5
         AccelInc_Val_1  =   10
         RangeMinVarType_1=   5
         RangeMin_Val_1  =   100
         RangeMaxVarType_1=   5
         RangeMax_Val_1  =   1000
         ButtonStyle_1   =   0
         Bindings_1      =   4
         ClassName_4     =   "CCWBindingHolderArray"
         Editor_4        =   5
         ClassName_5     =   "CCWBindingHolderArrayEditor"
         Owner_5         =   1
      End
      Begin CWUIControlsLib.CWNumEdit cwMin_dB
         Height          =   300
         Left            =   2610
         TabIndex        =   187
         ToolTipText     =   "Plot Min scale"
         Top             =   3870
         Width           =   675
         _Version        =   524288
         _ExtentX        =   1199
         _ExtentY        =   529
         _StockProps     =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
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
         BorderStyle_1   =   1
         TextAlignment_1 =   1
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
         Value_Val_1     =   -40
         IncValueVarType_1=   5
         IncValue_Val_1  =   5
         AccelIncVarType_1=   5
         AccelInc_Val_1  =   20
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
      Begin CWUIControlsLib.CWNumEdit cwMax_dB
         Height          =   300
         Left            =   2610
         TabIndex        =   188
         ToolTipText     =   "Plot Max scale"
         Top             =   135
         Width           =   705
         _Version        =   524288
         _ExtentX        =   1235
         _ExtentY        =   529
         _StockProps     =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
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
         BorderStyle_1   =   1
         TextAlignment_1 =   1
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
         Value_Val_1     =   20
         IncValueVarType_1=   5
         IncValue_Val_1  =   5
         AccelIncVarType_1=   5
         AccelInc_Val_1  =   20
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
      Begin VB.Frame frameAxis
         BackColor       =   &H00E0E0E0&
         Caption         =   "Transfer Function"
         BeginProperty Font
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   7360
         Left            =   135
         TabIndex        =   159
         Top             =   930
         Width           =   2460
         Begin VB.CommandButton cmdChooseExcitation
            BackColor       =   &H00C0C0FF&
            Caption         =   "Choose Output"
            Height          =   465
            Left            =   90
            Style           =   1  'Graphical
            TabIndex        =   166
            ToolTipText     =   "Choose where to apply Excitation for the test"
            Top             =   1845
            Width           =   2265
         End
         Begin VB.Frame FrameLTF
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   0  'None
            Height          =   4450
            Left            =   45
            TabIndex        =   168
            Top             =   2835
            Width           =   2355
            Begin VB.CommandButton CmdClearPlots
               BackColor       =   &H0000FFFF&
               Caption         =   "Clear"
               Height          =   300
               Left            =   45
               MaskColor       =   &H80000004&
               Style           =   1  'Graphical
               TabIndex        =   220
               ToolTipText     =   "Clear all plots"
               Top             =   3420
               UseMaskColor    =   -1  'True
               Width           =   645
            End
            Begin VB.CommandButton CmdReadScan
               BackColor       =   &H00C0C000&
               Caption         =   "Read Test"
               Height          =   310
               Left            =   1260
               MaskColor       =   &H80000004&
               Style           =   1  'Graphical
               TabIndex        =   217
               ToolTipText     =   "Reads saved data and shows plots"
               Top             =   2670
               UseMaskColor    =   -1  'True
               Width           =   1050
            End
            Begin VB.CommandButton CmdPlotReference
               BackColor       =   &H0060B0FF&
               Caption         =   "Get Ref Plot"
               Height          =   310
               Left            =   1260
               MaskColor       =   &H80000004&
               Style           =   1  'Graphical
               TabIndex        =   216
               ToolTipText     =   "Reads saved data and shows persistent Reference plot"
               Top             =   1910
               UseMaskColor    =   -1  'True
               Width           =   1050
            End
            Begin VB.CommandButton CmdReadLimits
               BackColor       =   &H00C0C0C0&
               Caption         =   "Get Limits"
               Height          =   310
               Left            =   1260
               MaskColor       =   &H80000004&
               Style           =   1  'Graphical
               TabIndex        =   215
               ToolTipText     =   $"Analyzer.frx":030A
               Top             =   2290
               UseMaskColor    =   -1  'True
               Width           =   1050
            End
            Begin VB.CommandButton cmdPrint
               BackColor       =   &H0000E000&
               Caption         =   "Print Test"
               Height          =   310
               Left            =   1260
               MaskColor       =   &H80000004&
               Style           =   1  'Graphical
               TabIndex        =   210
               ToolTipText     =   "Prints plots to any system printer"
               Top             =   3060
               UseMaskColor    =   -1  'True
               Width           =   1050
            End
            Begin VB.CheckBox ChkShowLimits
               BackColor       =   &H00808080&
               Caption         =   "Limits"
               ForeColor       =   &H00FFFFFF&
               Height          =   270
               Left            =   30
               TabIndex        =   209
               ToolTipText     =   "Shows/hides desired test limits if they are loaded"
               Top             =   2320
               Width           =   1120
            End
            Begin VB.CheckBox ChkShowRefPlot
               BackColor       =   &H0060B0FF&
               Caption         =   "Reference"
               Height          =   270
               Left            =   30
               TabIndex        =   208
               ToolTipText     =   "Shows/hides Reference plot"
               Top             =   1940
               Value           =   1  'Checked
               Width           =   1120
            End
            Begin VB.CheckBox ChkShowPrevPlot
               BackColor       =   &H0000C000&
               Caption         =   "Previous"
               Height          =   270
               Left            =   30
               TabIndex        =   207
               ToolTipText     =   "Shows/hides PREVIOUS plot"
               Top             =   1180
               Value           =   1  'Checked
               Width           =   1120
            End
            Begin VB.CheckBox ChkShowPredictionPlot
               BackColor       =   &H00FFA0A0&
               Caption         =   "Prediction"
               Height          =   270
               Left            =   30
               TabIndex        =   206
               ToolTipText     =   "Shows/hides Open Loop TF Prediction plot based on Reference plant and filters"
               Top             =   2700
               Width           =   1120
            End
            Begin VB.CheckBox ChkShowCurrentPlot
               BackColor       =   &H00FF0000&
               Caption         =   "Current"
               ForeColor       =   &H00FFFFFF&
               Height          =   270
               Left            =   30
               TabIndex        =   205
               ToolTipText     =   "Shows/hides most recent plot"
               Top             =   800
               Value           =   1  'Checked
               Width           =   1120
            End
            Begin VB.CheckBox ChkShowCoherencePlot
               BackColor       =   &H008080FF&
               Caption         =   "Coherence"
               BeginProperty Font
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   270
               Left            =   30
               TabIndex        =   204
               ToolTipText     =   "Shows/hides Coherence of the measurement on the Phase plot, in percents"
               Top             =   1560
               Value           =   1  'Checked
               Width           =   1120
            End
            Begin VB.CommandButton Cmd_SavePrediction
               BackColor       =   &H00FFA0A0&
               Caption         =   "Save Predict"
               Height          =   310
               Left            =   30
               MaskColor       =   &H80000004&
               Style           =   1  'Graphical
               TabIndex        =   202
               ToolTipText     =   "Saves prediction into data file"
               Top             =   3060
               UseMaskColor    =   -1  'True
               Width           =   1140
            End
            Begin VB.Frame Frame5
               Caption         =   "Cursor  Pan    Zoom   Scale "
               Height          =   730
               Left            =   0
               TabIndex        =   197
               Top             =   0
               Width           =   2350
               Begin VB.OptionButton OptGraphMode
                  Height          =   450
                  Index           =   0
                  Left            =   60
                  Picture         =   "Analyzer.frx":0391
                  Style           =   1  'Graphical
                  TabIndex        =   200
                  ToolTipText     =   "Snap cursor to chosen plot"
                  Top             =   210
                  Width           =   450
               End
               Begin VB.OptionButton OptGraphMode
                  Height          =   450
                  Index           =   1
                  Left            =   585
                  Picture         =   "Analyzer.frx":0BF3
                  Style           =   1  'Graphical
                  TabIndex        =   199
                  ToolTipText     =   "Pan window"
                  Top             =   210
                  Width           =   450
               End
               Begin VB.OptionButton OptGraphMode
                  Height          =   450
                  Index           =   2
                  Left            =   1100
                  Picture         =   "Analyzer.frx":13A1
                  Style           =   1  'Graphical
                  TabIndex        =   198
                  ToolTipText     =   "Zoom window"
                  Top             =   210
                  Width           =   450
               End
               Begin CWUIControlsLib.CWButton cwAutoscale
                  Height          =   450
                  Left            =   1610
                  TabIndex        =   201
                  ToolTipText     =   "Scale of Plots Toggle"
                  Top             =   210
                  Width           =   645
                  _Version        =   524288
                  _ExtentX        =   1138
                  _ExtentY        =   794
                  _StockProps     =   68
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}
                     Name            =   "Arial"
                     Size            =   9.01
                     Charset         =   204
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Reset_0         =   0   'False
                  CompatibleVers_0=   524288
                  Boolean_0       =   1
                  ClassName_1     =   "CCWBoolean"
                  opts_1          =   2614
                  SclRef_1.l      =   2
                  SclRef_1.t      =   2
                  SclRef_1.r      =   45
                  SclRef_1.b      =   45
                  Scl_1.l         =   2
                  Scl_1.t         =   2
                  Scl_1.r         =   45
                  Scl_1.b         =   46
                  C[0]_1          =   13821670
                  Enum_1          =   2
                  ClassName_2     =   "CCWEnum"
                  Array_2         =   2
                  Editor_2        =   0
                  Array[0]_2      =   3
                  ClassName_3     =   "CCWEnumElt"
                  opts_3          =   1
                  Name_3          =   "Off"
                  frame_3         =   286
                  DrawList_3      =   4
                  ClassName_4     =   "CDrawList"
                  count_4         =   4
                  list[4]_4       =   5
                  ClassName_5     =   "CCWDrawObj"
                  opts_5          =   62
                  C[0]_5          =   -2147483641
                  C[1]_5          =   -2147483641
                  Image_5         =   6
                  ClassName_6     =   "CCWTextImage"
                  font_6          =   0
                  Animator_5      =   0
                  Blinker_5       =   0
                  list[3]_4       =   7
                  ClassName_7     =   "CCWDrawObj"
                  opts_7          =   52
                  SclRef_7.l      =   2
                  SclRef_7.t      =   2
                  SclRef_7.r      =   45
                  SclRef_7.b      =   33
                  Scl_7.l         =   11
                  Scl_7.t         =   11
                  Scl_7.r         =   35
                  Scl_7.b         =   26
                  C[0]_7          =   0
                  C[1]_7          =   0
                  Image_7         =   8
                  ClassName_8     =   "CCWTextImage"
                  szText_8        =   "Auto"
                  font_8          =   0
                  Animator_7      =   0
                  Blinker_7       =   0
                  list[2]_4       =   9
                  ClassName_9     =   "CCWDrawObj"
                  opts_9          =   54
                  SclRef_9.l      =   2
                  SclRef_9.t      =   2
                  SclRef_9.r      =   45
                  SclRef_9.b      =   33
                  Scl_9.l         =   9
                  Scl_9.t         =   14
                  Scl_9.r         =   36
                  Scl_9.b         =   29
                  C[0]_9          =   16777215
                  C[1]_9          =   16777215
                  Image_9         =   10
                  ClassName_10    =   "CCWTextImage"
                  szText_10       =   "User"
                  font_10         =   0
                  Animator_9      =   0
                  Blinker_9       =   0
                  list[1]_4       =   11
                  ClassName_11    =   "CCWDrawObj"
                  opts_11         =   54
                  SclRef_11.l     =   2
                  SclRef_11.t     =   18
                  SclRef_11.r     =   66
                  SclRef_11.b     =   48
                  Scl_11.l        =   2
                  Scl_11.t        =   16
                  Scl_11.r        =   66
                  Scl_11.b        =   48
                  C[0]_11         =   16711680
                  C[1]_11         =   16711680
                  Image_11        =   12
                  ClassName_12    =   "CCWPictImage"
                  opts_12         =   1280
                  Rows_12         =   1
                  Cols_12         =   1
                  Pict_12         =   44
                  F_12            =   16711680
                  B_12            =   16711680
                  ColorReplaceWith_12=   8421504
                  ColorReplace_12 =   8421504
                  Tolerance_12    =   2
                  Animator_11     =   0
                  Blinker_11      =   0
                  varVarType_3    =   5
                  Array[1]_2      =   13
                  ClassName_13    =   "CCWEnumElt"
                  opts_13         =   1
                  Name_13         =   "On"
                  frame_13        =   286
                  DrawList_13     =   14
                  ClassName_14    =   "CDrawList"
                  count_14        =   4
                  list[4]_14      =   15
                  ClassName_15    =   "CCWDrawObj"
                  opts_15         =   62
                  C[0]_15         =   -2147483641
                  C[1]_15         =   -2147483641
                  Image_15        =   6
                  Animator_15     =   0
                  Blinker_15      =   0
                  list[3]_14      =   16
                  ClassName_16    =   "CCWDrawObj"
                  opts_16         =   54
                  SclRef_16.l     =   2
                  SclRef_16.t     =   2
                  SclRef_16.r     =   45
                  SclRef_16.b     =   33
                  Scl_16.l        =   11
                  Scl_16.t        =   11
                  Scl_16.r        =   35
                  Scl_16.b        =   26
                  C[0]_16         =   0
                  C[1]_16         =   0
                  Image_16        =   8
                  Animator_16     =   0
                  Blinker_16      =   0
                  list[2]_14      =   17
                  ClassName_17    =   "CCWDrawObj"
                  opts_17         =   52
                  SclRef_17.l     =   2
                  SclRef_17.t     =   2
                  SclRef_17.r     =   45
                  SclRef_17.b     =   33
                  Scl_17.l        =   9
                  Scl_17.t        =   14
                  Scl_17.r        =   36
                  Scl_17.b        =   29
                  C[0]_17         =   16777215
                  C[1]_17         =   16777215
                  Image_17        =   10
                  Animator_17     =   0
                  Blinker_17      =   0
                  list[1]_14      =   18
                  ClassName_18    =   "CCWDrawObj"
                  opts_18         =   54
                  SclRef_18.l     =   2
                  SclRef_18.t     =   2
                  SclRef_18.r     =   45
                  SclRef_18.b     =   36
                  Scl_18.l        =   2
                  Scl_18.t        =   2
                  Scl_18.r        =   45
                  Scl_18.b        =   42
                  C[0]_18         =   16744703
                  C[1]_18         =   16744703
                  Image_18        =   19
                  ClassName_19    =   "CCWPictImage"
                  opts_19         =   1280
                  Rows_19         =   1
                  Cols_19         =   1
                  Pict_19         =   42
                  F_19            =   16744703
                  B_19            =   16744703
                  ColorReplaceWith_19=   8421504
                  ColorReplace_19 =   8421504
                  Tolerance_19    =   2
                  Animator_18     =   0
                  Blinker_18      =   0
                  varVarType_13   =   5
                  Bindings_1      =   20
                  ClassName_20    =   "CCWBindingHolderArray"
                  Editor_20       =   21
                  ClassName_21    =   "CCWBindingHolderArrayEditor"
                  Owner_21        =   1
                  Style_1         =   7
                  drawState_1     =   1
                  mechAction_1    =   1
                  BGImg_1         =   22
                  ClassName_22    =   "CCWDrawObj"
                  opts_22         =   62
                  Image_22        =   23
                  ClassName_23    =   "CCWPictImage"
                  opts_23         =   1280
                  Rows_23         =   1
                  Cols_23         =   1
                  Pict_23         =   286
                  F_23            =   -2147483633
                  B_23            =   -2147483633
                  ColorReplaceWith_23=   8421504
                  ColorReplace_23 =   8421504
                  Tolerance_23    =   2
                  Animator_22     =   0
                  Blinker_22      =   0
                  Array_1         =   6
                  Editor_1        =   0
                  Array[0]_1      =   11
                  Array[1]_1      =   18
                  Array[2]_1      =   0
                  Array[3]_1      =   0
                  Array[4]_1      =   24
                  ClassName_24    =   "CCWDrawObj"
                  opts_24         =   62
                  Image_24        =   8
                  Animator_24     =   0
                  Blinker_24      =   0
                  Array[5]_1      =   25
                  ClassName_25    =   "CCWDrawObj"
                  opts_25         =   62
                  Image_25        =   10
                  Animator_25     =   0
                  Blinker_25      =   0
                  Label_1         =   26
                  ClassName_26    =   "CCWDrawObj"
                  opts_26         =   62
                  C[0]_26         =   -2147483640
                  Image_26        =   6
                  Animator_26     =   0
                  Blinker_26      =   0
               End
            End
            Begin CWUIControlsLib.CWSlide cwSldProgress
               Height          =   510
               Left            =   0
               TabIndex        =   169
               ToolTipText     =   "Approximate percentage of test completion"
               Top             =   3735
               Width           =   2355
               _Version        =   524288
               _ExtentX        =   4154
               _ExtentY        =   900
               _StockProps     =   68
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   204
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Reset_0         =   0   'False
               CompatibleVers_0=   524288
               Slider_0        =   1
               ClassName_1     =   "CCWSlider"
               opts_1          =   2102
               SclRef_1.l      =   2
               SclRef_1.t      =   2
               SclRef_1.r      =   126
               SclRef_1.b      =   33
               Scl_1.l         =   2
               Scl_1.t         =   2
               Scl_1.r         =   120
               Scl_1.b         =   23
               C[0]_1          =   14737632
               BGImg_1         =   2
               ClassName_2     =   "CCWDrawObj"
               opts_2          =   54
               SclRef_2.l      =   2
               SclRef_2.t      =   2
               SclRef_2.r      =   159
               SclRef_2.b      =   36
               Scl_2.l         =   2
               Scl_2.t         =   2
               Scl_2.r         =   159
               Scl_2.b         =   32
               Image_2         =   3
               ClassName_3     =   "CCWPictImage"
               opts_3          =   1280
               Rows_3          =   1
               Cols_3          =   1
               Pict_3          =   286
               F_3             =   -2147483633
               B_3             =   -2147483633
               ColorReplaceWith_3=   8421504
               ColorReplace_3  =   8421504
               Tolerance_3     =   2
               Animator_2      =   0
               Blinker_2       =   0
               BFImg_1         =   4
               ClassName_4     =   "CCWDrawObj"
               opts_4          =   62
               Image_4         =   5
               ClassName_5     =   "CCWPictImage"
               opts_5          =   1280
               Rows_5          =   1
               Cols_5          =   1
               Pict_5          =   286
               F_5             =   -2147483633
               B_5             =   -2147483633
               ColorReplaceWith_5=   8421504
               ColorReplace_5  =   8421504
               Tolerance_5     =   2
               Animator_4      =   0
               Blinker_4       =   0
               style_1         =   1
               Label_1         =   6
               ClassName_6     =   "CCWDrawObj"
               opts_6          =   54
               SclRef_6.l      =   2
               SclRef_6.t      =   2
               SclRef_6.r      =   162
               SclRef_6.b      =   36
               Scl_6.l         =   51
               Scl_6.t         =   17
               Scl_6.r         =   111
               Scl_6.b         =   31
               C[0]_6          =   -2147483640
               Image_6         =   7
               ClassName_7     =   "CCWTextImage"
               szText_7        =   "Progress, %"
               font_7          =   0
               Animator_6      =   0
               Blinker_6       =   0
               Border_1        =   8
               ClassName_8     =   "CCWDrawObj"
               opts_8          =   52
               SclRef_8.l      =   2
               SclRef_8.t      =   2
               SclRef_8.r      =   159
               SclRef_8.b      =   33
               Scl_8.l         =   2
               Scl_8.t         =   2
               Scl_8.r         =   159
               Scl_8.b         =   32
               Image_8         =   9
               ClassName_9     =   "CCWPictImage"
               opts_9          =   1280
               Rows_9          =   1
               Cols_9          =   1
               Pict_9          =   25
               F_9             =   -2147483633
               B_9             =   -2147483633
               ColorReplaceWith_9=   8421504
               ColorReplace_9  =   8421504
               Tolerance_9     =   2
               Animator_8      =   0
               Blinker_8       =   0
               FillBound_1     =   10
               ClassName_10    =   "CCWGuiObject"
               opts_10         =   60
               FillTok_1       =   11
               ClassName_11    =   "CCWGuiObject"
               opts_11         =   62
               Axis_1          =   12
               ClassName_12    =   "CCWAxis"
               opts_12         =   567
               Name_12         =   "Axis"
               SclRef_12.l     =   2
               SclRef_12.t     =   2
               SclRef_12.r     =   154
               SclRef_12.b     =   25
               Scl_12.l        =   9
               Scl_12.t        =   1
               Scl_12.r        =   149
               Scl_12.b        =   18
               Orientation_12  =   548
               format_12       =   13
               ClassName_13    =   "CCWFormat"
               Scale_12        =   14
               ClassName_14    =   "CCWScale"
               opts_14         =   90112
               rMin_14         =   7
               rMax_14         =   143
               dMax_14         =   100
               discInterval_14 =   1
               Radial_12       =   0
               Enum_12         =   15
               ClassName_15    =   "CCWEnum"
               Editor_15       =   16
               ClassName_16    =   "CCWEnumArrayEditor"
               Owner_16        =   12
               Font_12         =   0
               tickopts_12     =   2690
               major_12        =   20
               minor_12        =   5
               Caption_12      =   17
               ClassName_17    =   "CCWDrawObj"
               opts_17         =   62
               C[0]_17         =   -2147483640
               Image_17        =   18
               ClassName_18    =   "CCWTextImage"
               font_18         =   0
               Animator_17     =   0
               Blinker_17      =   0
               DrawLst_1       =   19
               ClassName_19    =   "CDrawList"
               count_19        =   10
               list[10]_19     =   8
               list[9]_19      =   20
               ClassName_20    =   "CCWThumb"
               opts_20         =   1048631
               Name_20         =   "Pointer-1"
               SclRef_20.l     =   2
               SclRef_20.t     =   2
               SclRef_20.r     =   151
               SclRef_20.b     =   25
               Scl_20.l        =   6
               Scl_20.t        =   14
               Scl_20.r        =   14
               Scl_20.b        =   32
               C[0]_20         =   16711680
               C[1]_20         =   16711680
               C[2]_20         =   16776960
               Image_20        =   21
               ClassName_21    =   "CCWPictImage"
               opts_21         =   1280
               Rows_21         =   1
               Cols_21         =   1
               Pict_21         =   286
               F_21            =   16711680
               B_21            =   16711680
               ColorReplaceWith_21=   8421504
               ColorReplace_21 =   8421504
               Tolerance_21    =   2
               Animator_20     =   0
               Blinker_20      =   0
               mode_20         =   1
               Value_20        =   1
               FillStyle_20    =   1
               Fill_20         =   22
               ClassName_22    =   "CCWDrawObj"
               opts_22         =   62
               Image_22        =   23
               ClassName_23    =   "CCWPictImage"
               opts_23         =   1280
               Rows_23         =   1
               Cols_23         =   1
               Pict_23         =   286
               F_23            =   -2147483633
               B_23            =   -2147483633
               ColorReplaceWith_23=   8421504
               ColorReplace_23 =   8421504
               Tolerance_23    =   2
               Animator_22     =   0
               Blinker_22      =   0
               list[8]_19      =   12
               list[7]_19      =   6
               list[6]_19      =   11
               list[5]_19      =   4
               list[4]_19      =   24
               ClassName_24    =   "CCWDrawObj"
               opts_24         =   54
               SclRef_24.l     =   2
               SclRef_24.t     =   2
               SclRef_24.r     =   151
               SclRef_24.b     =   25
               Scl_24.l        =   4
               Scl_24.t        =   14
               Scl_24.r        =   154
               Scl_24.b        =   32
               C[0]_24         =   16777215
               C[1]_24         =   16777215
               Image_24        =   25
               ClassName_25    =   "CCWPictImage"
               opts_25         =   1280
               Rows_25         =   1
               Cols_25         =   1
               Pict_25         =   7
               F_25            =   16777215
               B_25            =   16777215
               ColorReplaceWith_25=   8421504
               ColorReplace_25 =   8421504
               Tolerance_25    =   2
               Animator_24     =   0
               Blinker_24      =   0
               list[3]_19      =   26
               ClassName_26    =   "CCWDrawObj"
               opts_26         =   60
               C[0]_26         =   16777215
               C[1]_26         =   16777215
               Image_26        =   27
               ClassName_27    =   "CCWPictImage"
               opts_27         =   1280
               Rows_27         =   1
               Cols_27         =   1
               Pict_27         =   93
               F_27            =   16777215
               B_27            =   16777215
               ColorReplaceWith_27=   8421504
               ColorReplace_27 =   8421504
               Tolerance_27    =   2
               Animator_26     =   0
               Blinker_26      =   0
               list[2]_19      =   28
               ClassName_28    =   "CCWDrawObj"
               opts_28         =   60
               C[0]_28         =   16777215
               C[1]_28         =   16777215
               Image_28        =   29
               ClassName_29    =   "CCWPictImage"
               opts_29         =   1280
               Rows_29         =   1
               Cols_29         =   1
               Pict_29         =   94
               F_29            =   16777215
               B_29            =   16777215
               ColorReplaceWith_29=   8421504
               ColorReplace_29 =   8421504
               Tolerance_29    =   2
               Animator_28     =   0
               Blinker_28      =   0
               list[1]_19      =   2
               IncDec_1        =   0.05
               Ptrs_1          =   30
               ClassName_30    =   "CCWPointerArray"
               Array_30        =   1
               Editor_30       =   31
               ClassName_31    =   "CCWPointerArrayEditor"
               Owner_31        =   1
               Array[0]_30     =   20
               Bindings_1      =   32
               ClassName_32    =   "CCWBindingHolderArray"
               Editor_32       =   33
               ClassName_33    =   "CCWBindingHolderArrayEditor"
               Owner_33        =   1
               Stats_1         =   34
               ClassName_34    =   "CCWStats"
               doInc_1         =   28
               doDec_1         =   26
               doFrame_1       =   24
            End
            Begin CWUIControlsLib.CWButton cwButStartStopLTF
               Height          =   690
               Left            =   1245
               TabIndex        =   203
               ToolTipText     =   "Starts / stops Transfer Function test"
               Top             =   765
               Width           =   1050
               _Version        =   524288
               _ExtentX        =   1852
               _ExtentY        =   1217
               _StockProps     =   68
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
               Boolean_0       =   1
               ClassName_1     =   "CCWBoolean"
               opts_1          =   2614
               SclRef_1.l      =   2
               SclRef_1.t      =   2
               SclRef_1.r      =   72
               SclRef_1.b      =   51
               Scl_1.l         =   3
               Scl_1.t         =   18
               Scl_1.r         =   71
               Scl_1.b         =   52
               C[0]_1          =   16777152
               Enum_1          =   2
               ClassName_2     =   "CCWEnum"
               Array_2         =   2
               Editor_2        =   0
               Array[0]_2      =   3
               ClassName_3     =   "CCWEnumElt"
               opts_3          =   1
               Name_3          =   "Off"
               frame_3         =   286
               DrawList_3      =   4
               ClassName_4     =   "CDrawList"
               count_4         =   4
               list[4]_4       =   5
               ClassName_5     =   "CCWDrawObj"
               opts_5          =   62
               C[0]_5          =   0
               C[1]_5          =   0
               Image_5         =   6
               ClassName_6     =   "CCWTextImage"
               szText_6        =   "TF Test"
               font_6          =   7
               ClassName_7     =   "CCWFont"
               bFont_7         =   -1  'True
               BeginProperty Font_7 {0BE35203-8F91-11CE-9DE3-00AA004BB851}
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   204
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Animator_5      =   0
               Blinker_5       =   0
               list[3]_4       =   8
               ClassName_8     =   "CCWDrawObj"
               opts_8          =   60
               C[0]_8          =   0
               C[1]_8          =   0
               Image_8         =   9
               ClassName_9     =   "CCWTextImage"
               szText_9        =   "STOP"
               font_9          =   10
               ClassName_10    =   "CCWFont"
               bFont_10        =   -1  'True
               BeginProperty Font_10 {0BE35203-8F91-11CE-9DE3-00AA004BB851}
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   204
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Animator_8      =   0
               Blinker_8       =   0
               list[2]_4       =   11
               ClassName_11    =   "CCWDrawObj"
               opts_11         =   54
               SclRef_11.l     =   3
               SclRef_11.t     =   18
               SclRef_11.r     =   71
               SclRef_11.b     =   52
               Scl_11.l        =   14
               Scl_11.t        =   31
               Scl_11.r        =   57
               Scl_11.b        =   47
               C[0]_11         =   0
               C[1]_11         =   0
               Image_11        =   12
               ClassName_12    =   "CCWTextImage"
               szText_12       =   "START"
               font_12         =   13
               ClassName_13    =   "CCWFont"
               bFont_13        =   -1  'True
               BeginProperty Font_13 {0BE35203-8F91-11CE-9DE3-00AA004BB851}
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   204
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Animator_11     =   0
               Blinker_11      =   0
               list[1]_4       =   14
               ClassName_14    =   "CCWDrawObj"
               opts_14         =   54
               SclRef_14.l     =   3
               SclRef_14.t     =   20
               SclRef_14.r     =   73
               SclRef_14.b     =   59
               Scl_14.l        =   7
               Scl_14.t        =   20
               Scl_14.r        =   68
               Scl_14.b        =   55
               C[0]_14         =   65280
               C[1]_14         =   65280
               Image_14        =   15
               ClassName_15    =   "CCWPictImage"
               opts_15         =   1280
               Rows_15         =   1
               Cols_15         =   1
               Pict_15         =   44
               F_15            =   65280
               B_15            =   65280
               ColorReplaceWith_15=   8421504
               ColorReplace_15 =   8421504
               Tolerance_15    =   2
               Animator_14     =   0
               Blinker_14      =   0
               varVarType_3    =   5
               Array[1]_2      =   16
               ClassName_16    =   "CCWEnumElt"
               opts_16         =   1
               Name_16         =   "On"
               frame_16        =   286
               DrawList_16     =   17
               ClassName_17    =   "CDrawList"
               count_17        =   4
               list[4]_17      =   18
               ClassName_18    =   "CCWDrawObj"
               opts_18         =   62
               C[0]_18         =   0
               C[1]_18         =   0
               Image_18        =   6
               Animator_18     =   0
               Blinker_18      =   0
               list[3]_17      =   19
               ClassName_19    =   "CCWDrawObj"
               opts_19         =   62
               C[0]_19         =   0
               C[1]_19         =   0
               Image_19        =   9
               Animator_19     =   0
               Blinker_19      =   0
               list[2]_17      =   20
               ClassName_20    =   "CCWDrawObj"
               opts_20         =   52
               SclRef_20.l     =   3
               SclRef_20.t     =   18
               SclRef_20.r     =   71
               SclRef_20.b     =   52
               Scl_20.l        =   14
               Scl_20.t        =   31
               Scl_20.r        =   57
               Scl_20.b        =   47
               C[0]_20         =   0
               C[1]_20         =   0
               Image_20        =   12
               Animator_20     =   0
               Blinker_20      =   0
               list[1]_17      =   21
               ClassName_21    =   "CCWDrawObj"
               opts_21         =   62
               C[0]_21         =   8421631
               C[1]_21         =   8421631
               Image_21        =   22
               ClassName_22    =   "CCWPictImage"
               opts_22         =   1280
               Rows_22         =   1
               Cols_22         =   1
               Pict_22         =   42
               F_22            =   8421631
               B_22            =   8421631
               ColorReplaceWith_22=   8421504
               ColorReplace_22 =   8421504
               Tolerance_22    =   2
               Animator_21     =   0
               Blinker_21      =   0
               varVarType_16   =   5
               Bindings_1      =   23
               ClassName_23    =   "CCWBindingHolderArray"
               Editor_23       =   24
               ClassName_24    =   "CCWBindingHolderArrayEditor"
               Owner_24        =   1
               Style_1         =   7
               mechAction_1    =   1
               BGImg_1         =   25
               ClassName_25    =   "CCWDrawObj"
               opts_25         =   62
               Image_25        =   26
               ClassName_26    =   "CCWPictImage"
               opts_26         =   1280
               Rows_26         =   1
               Cols_26         =   1
               Pict_26         =   286
               F_26            =   -2147483633
               B_26            =   -2147483633
               ColorReplaceWith_26=   8421504
               ColorReplace_26 =   8421504
               Tolerance_26    =   2
               Animator_25     =   0
               Blinker_25      =   0
               Array_1         =   6
               Editor_1        =   0
               Array[0]_1      =   14
               Array[1]_1      =   21
               Array[2]_1      =   0
               Array[3]_1      =   0
               Array[4]_1      =   27
               ClassName_27    =   "CCWDrawObj"
               opts_27         =   62
               Image_27        =   9
               Animator_27     =   0
               Blinker_27      =   0
               Array[5]_1      =   28
               ClassName_28    =   "CCWDrawObj"
               opts_28         =   62
               Image_28        =   12
               Animator_28     =   0
               Blinker_28      =   0
               Label_1         =   29
               ClassName_29    =   "CCWDrawObj"
               opts_29         =   62
               C[0]_29         =   -2147483640
               Image_29        =   6
               Animator_29     =   0
               Blinker_29      =   0
            End
            Begin CWUIControlsLib.CWButton cwSaveScan
               Height          =   330
               Left            =   1260
               TabIndex        =   211
               ToolTipText     =   "Saves most recent plots from screen into data file"
               Top             =   1500
               Width           =   1050
               _Version        =   524288
               _ExtentX        =   1852
               _ExtentY        =   582
               _StockProps     =   68
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   204
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Reset_0         =   0   'False
               CompatibleVers_0=   524288
               Boolean_0       =   1
               ClassName_1     =   "CCWBoolean"
               opts_1          =   2622
               C[0]_1          =   14737632
               Enum_1          =   2
               ClassName_2     =   "CCWEnum"
               Array_2         =   2
               Editor_2        =   0
               Array[0]_2      =   3
               ClassName_3     =   "CCWEnumElt"
               opts_3          =   1
               Name_3          =   "Off"
               frame_3         =   288
               DrawList_3      =   4
               ClassName_4     =   "CDrawList"
               count_4         =   4
               list[4]_4       =   5
               ClassName_5     =   "CCWDrawObj"
               opts_5          =   62
               C[0]_5          =   0
               C[1]_5          =   0
               Image_5         =   6
               ClassName_6     =   "CCWTextImage"
               font_6          =   0
               Animator_5      =   0
               Blinker_5       =   0
               list[3]_4       =   7
               ClassName_7     =   "CCWDrawObj"
               opts_7          =   60
               C[0]_7          =   -2147483640
               C[1]_7          =   -2147483640
               Image_7         =   8
               ClassName_8     =   "CCWTextImage"
               szText_8        =   "CANCEL"
               style_8         =   48598384
               font_8          =   0
               Animator_7      =   0
               Blinker_7       =   0
               list[2]_4       =   9
               ClassName_9     =   "CCWDrawObj"
               opts_9          =   62
               C[0]_9          =   -2147483640
               C[1]_9          =   -2147483640
               Image_9         =   10
               ClassName_10    =   "CCWTextImage"
               szText_10       =   "Save Test"
               font_10         =   0
               Animator_9      =   0
               Blinker_9       =   0
               list[1]_4       =   11
               ClassName_11    =   "CCWDrawObj"
               opts_11         =   62
               C[0]_11         =   49152
               C[1]_11         =   49152
               Image_11        =   12
               ClassName_12    =   "CCWPictImage"
               opts_12         =   1280
               Rows_12         =   1
               Cols_12         =   1
               Pict_12         =   2
               F_12            =   49152
               B_12            =   49152
               ColorReplaceWith_12=   8421504
               ColorReplace_12 =   8421504
               Tolerance_12    =   2
               Animator_11     =   0
               Blinker_11      =   0
               varVarType_3    =   5
               Array[1]_2      =   13
               ClassName_13    =   "CCWEnumElt"
               opts_13         =   1
               Name_13         =   "On"
               frame_13        =   287
               DrawList_13     =   14
               ClassName_14    =   "CDrawList"
               count_14        =   4
               list[4]_14      =   15
               ClassName_15    =   "CCWDrawObj"
               opts_15         =   62
               C[0]_15         =   0
               C[1]_15         =   0
               Image_15        =   6
               Animator_15     =   0
               Blinker_15      =   0
               list[3]_14      =   16
               ClassName_16    =   "CCWDrawObj"
               opts_16         =   62
               C[0]_16         =   -2147483640
               C[1]_16         =   -2147483640
               Image_16        =   8
               Animator_16     =   0
               Blinker_16      =   0
               list[2]_14      =   17
               ClassName_17    =   "CCWDrawObj"
               opts_17         =   60
               C[0]_17         =   -2147483640
               C[1]_17         =   -2147483640
               Image_17        =   10
               Animator_17     =   0
               Blinker_17      =   0
               list[1]_14      =   18
               ClassName_18    =   "CCWDrawObj"
               opts_18         =   62
               C[0]_18         =   65535
               C[1]_18         =   65535
               Image_18        =   19
               ClassName_19    =   "CCWPictImage"
               opts_19         =   1280
               Rows_19         =   1
               Cols_19         =   1
               Pict_19         =   2
               F_19            =   65535
               B_19            =   65535
               ColorReplaceWith_19=   8421504
               ColorReplace_19 =   8421504
               Tolerance_19    =   2
               Animator_18     =   0
               Blinker_18      =   0
               varVarType_13   =   5
               Bindings_1      =   20
               ClassName_20    =   "CCWBindingHolderArray"
               Editor_20       =   21
               ClassName_21    =   "CCWBindingHolderArrayEditor"
               Owner_21        =   1
               Style_1         =   12
               frameStyle_1    =   1
               mechAction_1    =   1
               BGImg_1         =   22
               ClassName_22    =   "CCWDrawObj"
               opts_22         =   62
               Image_22        =   23
               ClassName_23    =   "CCWPictImage"
               opts_23         =   1280
               Rows_23         =   1
               Cols_23         =   1
               Pict_23         =   286
               F_23            =   -2147483633
               B_23            =   -2147483633
               ColorReplaceWith_23=   8421504
               ColorReplace_23 =   8421504
               Tolerance_23    =   2
               Animator_22     =   0
               Blinker_22      =   0
               Array_1         =   6
               Editor_1        =   0
               Array[0]_1      =   11
               Array[1]_1      =   18
               Array[2]_1      =   0
               Array[3]_1      =   0
               Array[4]_1      =   24
               ClassName_24    =   "CCWDrawObj"
               opts_24         =   62
               Image_24        =   8
               Animator_24     =   0
               Blinker_24      =   0
               Array[5]_1      =   25
               ClassName_25    =   "CCWDrawObj"
               opts_25         =   62
               Image_25        =   10
               Animator_25     =   0
               Blinker_25      =   0
               Label_1         =   26
               ClassName_26    =   "CCWDrawObj"
               opts_26         =   62
               C[0]_26         =   -2147483640
               Image_26        =   6
               Animator_26     =   0
               Blinker_26      =   0
            End
            Begin CWUIControlsLib.CWButton CWButton1
               Height          =   375
               Left            =   0
               TabIndex        =   214
               ToolTipText     =   "Reads saved data and shows plots"
               Top             =   0
               Width           =   1050
               _Version        =   524288
               _ExtentX        =   1852
               _ExtentY        =   661
               _StockProps     =   68
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}
                  Name            =   "Arial"
                  Size            =   8.24
                  Charset         =   204
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Reset_0         =   0   'False
               CompatibleVers_0=   524288
               Boolean_0       =   1
               ClassName_1     =   "CCWBoolean"
               opts_1          =   2622
               C[0]_1          =   14737632
               Enum_1          =   2
               ClassName_2     =   "CCWEnum"
               Array_2         =   2
               Editor_2        =   0
               Array[0]_2      =   3
               ClassName_3     =   "CCWEnumElt"
               opts_3          =   1
               Name_3          =   "Off"
               frame_3         =   288
               DrawList_3      =   4
               ClassName_4     =   "CDrawList"
               count_4         =   4
               list[4]_4       =   5
               ClassName_5     =   "CCWDrawObj"
               opts_5          =   62
               C[0]_5          =   0
               C[1]_5          =   0
               Image_5         =   6
               ClassName_6     =   "CCWTextImage"
               font_6          =   0
               Animator_5      =   0
               Blinker_5       =   0
               list[3]_4       =   7
               ClassName_7     =   "CCWDrawObj"
               opts_7          =   60
               C[0]_7          =   -2147483640
               C[1]_7          =   -2147483640
               Image_7         =   8
               ClassName_8     =   "CCWTextImage"
               szText_8        =   "CANCEL"
               style_8         =   48598384
               font_8          =   0
               Animator_7      =   0
               Blinker_7       =   0
               list[2]_4       =   9
               ClassName_9     =   "CCWDrawObj"
               opts_9          =   62
               C[0]_9          =   -2147483640
               C[1]_9          =   -2147483640
               Image_9         =   10
               ClassName_10    =   "CCWTextImage"
               szText_10       =   "Read Test"
               font_10         =   0
               Animator_9      =   0
               Blinker_9       =   0
               list[1]_4       =   11
               ClassName_11    =   "CCWDrawObj"
               opts_11         =   62
               C[0]_11         =   12632064
               C[1]_11         =   12632064
               Image_11        =   12
               ClassName_12    =   "CCWPictImage"
               opts_12         =   1280
               Rows_12         =   1
               Cols_12         =   1
               Pict_12         =   2
               F_12            =   12632064
               B_12            =   12632064
               ColorReplaceWith_12=   8421504
               ColorReplace_12 =   8421504
               Tolerance_12    =   2
               Animator_11     =   0
               Blinker_11      =   0
               varVarType_3    =   5
               Array[1]_2      =   13
               ClassName_13    =   "CCWEnumElt"
               opts_13         =   1
               Name_13         =   "On"
               frame_13        =   287
               DrawList_13     =   14
               ClassName_14    =   "CDrawList"
               count_14        =   4
               list[4]_14      =   15
               ClassName_15    =   "CCWDrawObj"
               opts_15         =   62
               C[0]_15         =   0
               C[1]_15         =   0
               Image_15        =   6
               Animator_15     =   0
               Blinker_15      =   0
               list[3]_14      =   16
               ClassName_16    =   "CCWDrawObj"
               opts_16         =   62
               C[0]_16         =   -2147483640
               C[1]_16         =   -2147483640
               Image_16        =   8
               Animator_16     =   0
               Blinker_16      =   0
               list[2]_14      =   17
               ClassName_17    =   "CCWDrawObj"
               opts_17         =   60
               C[0]_17         =   -2147483640
               C[1]_17         =   -2147483640
               Image_17        =   10
               Animator_17     =   0
               Blinker_17      =   0
               list[1]_14      =   18
               ClassName_18    =   "CCWDrawObj"
               opts_18         =   62
               C[0]_18         =   8438015
               C[1]_18         =   8438015
               Image_18        =   19
               ClassName_19    =   "CCWPictImage"
               opts_19         =   1280
               Rows_19         =   1
               Cols_19         =   1
               Pict_19         =   2
               F_19            =   8438015
               B_19            =   8438015
               ColorReplaceWith_19=   8421504
               ColorReplace_19 =   8421504
               Tolerance_19    =   2
               Animator_18     =   0
               Blinker_18      =   0
               varVarType_13   =   5
               Bindings_1      =   20
               ClassName_20    =   "CCWBindingHolderArray"
               Editor_20       =   21
               ClassName_21    =   "CCWBindingHolderArrayEditor"
               Owner_21        =   1
               Style_1         =   12
               frameStyle_1    =   1
               mechAction_1    =   1
               BGImg_1         =   22
               ClassName_22    =   "CCWDrawObj"
               opts_22         =   62
               Image_22        =   23
               ClassName_23    =   "CCWPictImage"
               opts_23         =   1280
               Rows_23         =   1
               Cols_23         =   1
               Pict_23         =   286
               F_23            =   -2147483633
               B_23            =   -2147483633
               ColorReplaceWith_23=   8421504
               ColorReplace_23 =   8421504
               Tolerance_23    =   2
               Animator_22     =   0
               Blinker_22      =   0
               Array_1         =   6
               Editor_1        =   0
               Array[0]_1      =   11
               Array[1]_1      =   18
               Array[2]_1      =   0
               Array[3]_1      =   0
               Array[4]_1      =   24
               ClassName_24    =   "CCWDrawObj"
               opts_24         =   62
               Image_24        =   8
               Animator_24     =   0
               Blinker_24      =   0
               Array[5]_1      =   25
               ClassName_25    =   "CCWDrawObj"
               opts_25         =   62
               Image_25        =   10
               Animator_25     =   0
               Blinker_25      =   0
               Label_1         =   26
               ClassName_26    =   "CCWDrawObj"
               opts_26         =   62
               C[0]_26         =   -2147483640
               Image_26        =   6
               Animator_26     =   0
               Blinker_26      =   0
            End
            Begin CWUIControlsLib.CWButton cwReadScan
               Height          =   330
               Left            =   1260
               TabIndex        =   218
               ToolTipText     =   "Reads saved data and shows plots"
               Top             =   2655
               Width           =   1050
               _Version        =   524288
               _ExtentX        =   1852
               _ExtentY        =   582
               _StockProps     =   68
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   204
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Reset_0         =   0   'False
               CompatibleVers_0=   524288
               Boolean_0       =   1
               ClassName_1     =   "CCWBoolean"
               opts_1          =   2622
               C[0]_1          =   14737632
               Enum_1          =   2
               ClassName_2     =   "CCWEnum"
               Array_2         =   2
               Editor_2        =   0
               Array[0]_2      =   3
               ClassName_3     =   "CCWEnumElt"
               opts_3          =   1
               Name_3          =   "Off"
               frame_3         =   288
               DrawList_3      =   4
               ClassName_4     =   "CDrawList"
               count_4         =   4
               list[4]_4       =   5
               ClassName_5     =   "CCWDrawObj"
               opts_5          =   62
               C[0]_5          =   0
               C[1]_5          =   0
               Image_5         =   6
               ClassName_6     =   "CCWTextImage"
               font_6          =   0
               Animator_5      =   0
               Blinker_5       =   0
               list[3]_4       =   7
               ClassName_7     =   "CCWDrawObj"
               opts_7          =   60
               C[0]_7          =   -2147483640
               C[1]_7          =   -2147483640
               Image_7         =   8
               ClassName_8     =   "CCWTextImage"
               szText_8        =   "CANCEL"
               style_8         =   48598384
               font_8          =   0
               Animator_7      =   0
               Blinker_7       =   0
               list[2]_4       =   9
               ClassName_9     =   "CCWDrawObj"
               opts_9          =   62
               C[0]_9          =   -2147483640
               C[1]_9          =   -2147483640
               Image_9         =   10
               ClassName_10    =   "CCWTextImage"
               szText_10       =   "Read Test"
               font_10         =   0
               Animator_9      =   0
               Blinker_9       =   0
               list[1]_4       =   11
               ClassName_11    =   "CCWDrawObj"
               opts_11         =   62
               C[0]_11         =   12632064
               C[1]_11         =   12632064
               Image_11        =   12
               ClassName_12    =   "CCWPictImage"
               opts_12         =   1280
               Rows_12         =   1
               Cols_12         =   1
               Pict_12         =   2
               F_12            =   12632064
               B_12            =   12632064
               ColorReplaceWith_12=   8421504
               ColorReplace_12 =   8421504
               Tolerance_12    =   2
               Animator_11     =   0
               Blinker_11      =   0
               varVarType_3    =   5
               Array[1]_2      =   13
               ClassName_13    =   "CCWEnumElt"
               opts_13         =   1
               Name_13         =   "On"
               frame_13        =   287
               DrawList_13     =   14
               ClassName_14    =   "CDrawList"
               count_14        =   4
               list[4]_14      =   15
               ClassName_15    =   "CCWDrawObj"
               opts_15         =   62
               C[0]_15         =   0
               C[1]_15         =   0
               Image_15        =   6
               Animator_15     =   0
               Blinker_15      =   0
               list[3]_14      =   16
               ClassName_16    =   "CCWDrawObj"
               opts_16         =   62
               C[0]_16         =   -2147483640
               C[1]_16         =   -2147483640
               Image_16        =   8
               Animator_16     =   0
               Blinker_16      =   0
               list[2]_14      =   17
               ClassName_17    =   "CCWDrawObj"
               opts_17         =   60
               C[0]_17         =   -2147483640
               C[1]_17         =   -2147483640
               Image_17        =   10
               Animator_17     =   0
               Blinker_17      =   0
               list[1]_14      =   18
               ClassName_18    =   "CCWDrawObj"
               opts_18         =   62
               C[0]_18         =   8438015
               C[1]_18         =   8438015
               Image_18        =   19
               ClassName_19    =   "CCWPictImage"
               opts_19         =   1280
               Rows_19         =   1
               Cols_19         =   1
               Pict_19         =   2
               F_19            =   8438015
               B_19            =   8438015
               ColorReplaceWith_19=   8421504
               ColorReplace_19 =   8421504
               Tolerance_19    =   2
               Animator_18     =   0
               Blinker_18      =   0
               varVarType_13   =   5
               Bindings_1      =   20
               ClassName_20    =   "CCWBindingHolderArray"
               Editor_20       =   21
               ClassName_21    =   "CCWBindingHolderArrayEditor"
               Owner_21        =   1
               Style_1         =   12
               frameStyle_1    =   1
               mechAction_1    =   1
               BGImg_1         =   22
               ClassName_22    =   "CCWDrawObj"
               opts_22         =   62
               Image_22        =   23
               ClassName_23    =   "CCWPictImage"
               opts_23         =   1280
               Rows_23         =   1
               Cols_23         =   1
               Pict_23         =   286
               F_23            =   -2147483633
               B_23            =   -2147483633
               ColorReplaceWith_23=   8421504
               ColorReplace_23 =   8421504
               Tolerance_23    =   2
               Animator_22     =   0
               Blinker_22      =   0
               Array_1         =   6
               Editor_1        =   0
               Array[0]_1      =   11
               Array[1]_1      =   18
               Array[2]_1      =   0
               Array[3]_1      =   0
               Array[4]_1      =   24
               ClassName_24    =   "CCWDrawObj"
               opts_24         =   62
               Image_24        =   8
               Animator_24     =   0
               Blinker_24      =   0
               Array[5]_1      =   25
               ClassName_25    =   "CCWDrawObj"
               opts_25         =   62
               Image_25        =   10
               Animator_25     =   0
               Blinker_25      =   0
               Label_1         =   26
               ClassName_26    =   "CCWDrawObj"
               opts_26         =   62
               C[0]_26         =   -2147483640
               Image_26        =   6
               Animator_26     =   0
               Blinker_26      =   0
            End
            Begin VB.Label LabelCurrentLTFline
               BackColor       =   &H00FFFFFF&
               Caption         =   "975,-37.6, +166#"
               Height          =   240
               Left            =   30
               TabIndex        =   186
               Top             =   4230
               Width           =   2250
            End
            Begin VB.Label lblTestTime
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "226"
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
               Height          =   300
               Left            =   1710
               TabIndex        =   170
               ToolTipText     =   "Elapsed test time"
               Top             =   3420
               Width           =   600
            End
            Begin VB.Label Label1
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E0E0E0&
               Caption         =   "Elapsed s"
               BeginProperty Font
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Index           =   7
               Left            =   765
               TabIndex        =   171
               Top             =   3420
               Width           =   915
            End
         End
         Begin VB.Frame FrameChooseTestType
            BackColor       =   &H00C0C0C0&
            Caption         =   "Choose Test Type"
            BeginProperty Font
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1560
            Left            =   90
            TabIndex        =   161
            Top             =   240
            Width           =   2265
            Begin VB.OptionButton optTestInput
               BackColor       =   &H00C0FFFF&
               Caption         =   "Freq Sweep OLTF"
               BeginProperty Font
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   204
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   1
               Left            =   90
               TabIndex        =   165
               Tag             =   "oltf "
               ToolTipText     =   $"Analyzer.frx":1D2B
               Top             =   585
               Value           =   -1  'True
               Width           =   2055
            End
            Begin VB.OptionButton optTestInput
               BackColor       =   &H00FFFFC0&
               Caption         =   "Freq Sweep Plant"
               BeginProperty Font
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   204
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   270
               Index           =   2
               Left            =   90
               TabIndex        =   164
               Tag             =   "p_tf "
               ToolTipText     =   "For each frequency in sweep: Input signal when driving Output, without filters in control loop"
               Top             =   270
               Width           =   2055
            End
            Begin VB.OptionButton optTestInput
               BackColor       =   &H0080C0FF&
               Caption         =   "Measure at one freq"
               BeginProperty Font
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   204
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   270
               Index           =   0
               Left            =   90
               TabIndex        =   163
               Tag             =   "mtst "
               ToolTipText     =   "Open Loop Velocity Transfer Function when signal goes via filters but loop is open"
               Top             =   900
               Width           =   2055
            End
            Begin VB.CheckBox ChkAdvTF
               BackColor       =   &H00FFC0FF&
               Caption         =   "Advanced TF sweep"
               Height          =   270
               Left            =   90
               TabIndex        =   162
               ToolTipText     =   "Advanced TF llows to compare two inputs to each other and select averaging; regular TF compares input to driving output"
               Top             =   1215
               Width           =   2040
            End
         End
         Begin VB.CommandButton CmdChoose2ndInputSignal
            BackColor       =   &H00FFFFC0&
            Caption         =   "Choose Inp2"
            Height          =   465
            Left            =   1260
            Style           =   1  'Graphical
            TabIndex        =   160
            ToolTipText     =   "Choose where to take second input signal for advanced TF"
            Top             =   2385
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.CommandButton CmdChooseInputSignal
            BackColor       =   &H00C0FFC0&
            Caption         =   "Choose Input"
            Height          =   465
            Left            =   90
            Style           =   1  'Graphical
            TabIndex        =   167
            ToolTipText     =   "Choose where to take input signal"
            Top             =   2385
            Width           =   2265
         End
         Begin VB.Frame FrameMeasureAtFreq
            BackColor       =   &H0080C0FF&
            BorderStyle     =   0  'None
            Height          =   4300
            Left            =   135
            TabIndex        =   172
            Top             =   2880
            Width           =   2175
            Begin VB.TextBox txtMag
               Alignment       =   2  'Center
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
               Left            =   405
               TabIndex        =   174
               Text            =   "Magnitude"
               ToolTipText     =   "Magnitude ratio in dB"
               Top             =   2970
               Width           =   1365
            End
            Begin VB.TextBox txtPhase
               Alignment       =   2  'Center
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
               Left            =   405
               TabIndex        =   173
               Text            =   "Phase"
               ToolTipText     =   "Phase ratio in degrees"
               Top             =   3780
               Width           =   1365
            End
            Begin CWUIControlsLib.CWButton cwButStartStopMotSenTest
               Height          =   915
               Left            =   405
               TabIndex        =   175
               ToolTipText     =   "Starts / stops measurement at one frequency"
               Top             =   1620
               Width           =   1410
               _Version        =   524288
               _ExtentX        =   2487
               _ExtentY        =   1614
               _StockProps     =   68
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
               Boolean_0       =   1
               ClassName_1     =   "CCWBoolean"
               opts_1          =   2614
               SclRef_1.l      =   2
               SclRef_1.t      =   2
               SclRef_1.r      =   96
               SclRef_1.b      =   63
               Scl_1.l         =   9
               Scl_1.t         =   19
               Scl_1.r         =   90
               Scl_1.b         =   58
               C[0]_1          =   16777152
               Enum_1          =   2
               ClassName_2     =   "CCWEnum"
               Array_2         =   2
               Editor_2        =   0
               Array[0]_2      =   3
               ClassName_3     =   "CCWEnumElt"
               opts_3          =   1
               Name_3          =   "Off"
               frame_3         =   288
               DrawList_3      =   4
               ClassName_4     =   "CDrawList"
               count_4         =   4
               list[4]_4       =   5
               ClassName_5     =   "CCWDrawObj"
               opts_5          =   62
               C[0]_5          =   -2147483640
               C[1]_5          =   -2147483640
               Image_5         =   6
               ClassName_6     =   "CCWTextImage"
               szText_6        =   "Sine Test"
               font_6          =   7
               ClassName_7     =   "CCWFont"
               bFont_7         =   -1  'True
               BeginProperty Font_7 {0BE35203-8F91-11CE-9DE3-00AA004BB851}
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   204
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Animator_5      =   0
               Blinker_5       =   0
               list[3]_4       =   8
               ClassName_8     =   "CCWDrawObj"
               opts_8          =   60
               C[0]_8          =   0
               C[1]_8          =   0
               Image_8         =   9
               ClassName_9     =   "CCWTextImage"
               szText_9        =   "DOING"
               font_9          =   0
               Animator_8      =   0
               Blinker_8       =   0
               list[2]_4       =   10
               ClassName_10    =   "CCWDrawObj"
               opts_10         =   62
               C[0]_10         =   -2147483640
               C[1]_10         =   -2147483640
               Image_10        =   11
               ClassName_11    =   "CCWTextImage"
               szText_11       =   "START"
               font_11         =   0
               Animator_10     =   0
               Blinker_10      =   0
               list[1]_4       =   12
               ClassName_12    =   "CCWDrawObj"
               opts_12         =   62
               C[0]_12         =   65280
               C[1]_12         =   65280
               Image_12        =   13
               ClassName_13    =   "CCWPictImage"
               opts_13         =   1280
               Rows_13         =   1
               Cols_13         =   1
               Pict_13         =   2
               F_13            =   65280
               B_13            =   65280
               ColorReplaceWith_13=   8421504
               ColorReplace_13 =   8421504
               Tolerance_13    =   2
               Animator_12     =   0
               Blinker_12      =   0
               varVarType_3    =   5
               Array[1]_2      =   14
               ClassName_14    =   "CCWEnumElt"
               opts_14         =   1
               Name_14         =   "On"
               frame_14        =   287
               DrawList_14     =   15
               ClassName_15    =   "CDrawList"
               count_15        =   4
               list[4]_15      =   16
               ClassName_16    =   "CCWDrawObj"
               opts_16         =   62
               C[0]_16         =   -2147483640
               C[1]_16         =   -2147483640
               Image_16        =   6
               Animator_16     =   0
               Blinker_16      =   0
               list[3]_15      =   17
               ClassName_17    =   "CCWDrawObj"
               opts_17         =   62
               C[0]_17         =   0
               C[1]_17         =   0
               Image_17        =   9
               Animator_17     =   0
               Blinker_17      =   0
               list[2]_15      =   18
               ClassName_18    =   "CCWDrawObj"
               opts_18         =   60
               C[0]_18         =   -2147483640
               C[1]_18         =   -2147483640
               Image_18        =   11
               Animator_18     =   0
               Blinker_18      =   0
               list[1]_15      =   19
               ClassName_19    =   "CCWDrawObj"
               opts_19         =   62
               C[0]_19         =   65535
               C[1]_19         =   65535
               Image_19        =   20
               ClassName_20    =   "CCWPictImage"
               opts_20         =   1280
               Rows_20         =   1
               Cols_20         =   1
               Pict_20         =   2
               F_20            =   65535
               B_20            =   65535
               ColorReplaceWith_20=   8421504
               ColorReplace_20 =   8421504
               Tolerance_20    =   2
               Animator_19     =   0
               Blinker_19      =   0
               varVarType_14   =   5
               Bindings_1      =   21
               ClassName_21    =   "CCWBindingHolderArray"
               Editor_21       =   22
               ClassName_22    =   "CCWBindingHolderArrayEditor"
               Owner_22        =   1
               Style_1         =   12
               frameStyle_1    =   1
               mechAction_1    =   1
               BGImg_1         =   23
               ClassName_23    =   "CCWDrawObj"
               opts_23         =   62
               Image_23        =   24
               ClassName_24    =   "CCWPictImage"
               opts_24         =   1280
               Rows_24         =   1
               Cols_24         =   1
               Pict_24         =   286
               F_24            =   -2147483633
               B_24            =   -2147483633
               ColorReplaceWith_24=   8421504
               ColorReplace_24 =   8421504
               Tolerance_24    =   2
               Animator_23     =   0
               Blinker_23      =   0
               Array_1         =   6
               Editor_1        =   0
               Array[0]_1      =   12
               Array[1]_1      =   19
               Array[2]_1      =   0
               Array[3]_1      =   0
               Array[4]_1      =   25
               ClassName_25    =   "CCWDrawObj"
               opts_25         =   62
               Image_25        =   9
               Animator_25     =   0
               Blinker_25      =   0
               Array[5]_1      =   26
               ClassName_26    =   "CCWDrawObj"
               opts_26         =   62
               Image_26        =   11
               Animator_26     =   0
               Blinker_26      =   0
               Label_1         =   27
               ClassName_27    =   "CCWDrawObj"
               opts_27         =   62
               C[0]_27         =   -2147483640
               Image_27        =   6
               Animator_27     =   0
               Blinker_27      =   0
            End
            Begin CWUIControlsLib.CWNumEdit cwNumExcitFreq
               Height          =   360
               Index           =   0
               Left            =   405
               TabIndex        =   176
               ToolTipText     =   "Driving frequency, Hz, 0.01...1000 Hz (default = 5 Hz)"
               Top             =   405
               Width           =   1380
               _Version        =   524288
               _ExtentX        =   2434
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
               C[0]_1          =   12648384
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
               Value_Val_1     =   5
               IncValueVarType_1=   5
               IncValue_Val_1  =   1
               AccelIncVarType_1=   5
               AccelInc_Val_1  =   10
               RangeMinVarType_1=   5
               RangeMaxVarType_1=   5
               RangeMax_Val_1  =   200
               ButtonStyle_1   =   0
               Bindings_1      =   4
               ClassName_4     =   "CCWBindingHolderArray"
               Editor_4        =   5
               ClassName_5     =   "CCWBindingHolderArrayEditor"
               Owner_5         =   1
            End
            Begin CWUIControlsLib.CWNumEdit cwNumExcitAmpl
               Height          =   360
               Index           =   0
               Left            =   405
               TabIndex        =   177
               ToolTipText     =   "Driving Amplitude, 0...30000 (default = 300)"
               Top             =   1125
               Width           =   1380
               _Version        =   524288
               _ExtentX        =   2434
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
               C[0]_1          =   12648447
               C[1]_1          =   0
               BorderStyle_1   =   1
               TextAlignment_1 =   2
               format_1        =   2
               ClassName_2     =   "CCWFormat"
               Format_2        =   ".###"
               scale_1         =   3
               ClassName_3     =   "CCWScale"
               opts_3          =   65536
               dMax_3          =   10
               discInterval_3  =   1
               ValueVarType_1  =   5
               Value_Val_1     =   500
               IncValueVarType_1=   5
               IncValue_Val_1  =   10
               AccelIncVarType_1=   5
               AccelInc_Val_1  =   10
               RangeMinVarType_1=   5
               RangeMaxVarType_1=   5
               RangeMax_Val_1  =   30000
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
               BackColor       =   &H00C0FFC0&
               Caption         =   " Sine Freq, Hz"
               BeginProperty Font
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   204
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   285
               Index           =   1
               Left            =   405
               TabIndex        =   181
               ToolTipText     =   "Driving frequency, Hz, 0.01...1000 Hz (default = 5 Hz)"
               Top             =   135
               Width           =   1380
            End
            Begin VB.Label Label1
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0FFFF&
               Caption         =   "Amplitude"
               BeginProperty Font
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   204
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   285
               Index           =   5
               Left            =   405
               TabIndex        =   180
               ToolTipText     =   "Driving Amplitude, 0...30000 (default = 300)"
               Top             =   855
               Width           =   1380
            End
            Begin VB.Label Label1
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Caption         =   "Phase deg"
               BeginProperty Font
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   204
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   285
               Index           =   2
               Left            =   405
               TabIndex        =   179
               ToolTipText     =   "Phase ratio in degrees"
               Top             =   3510
               Width           =   1380
            End
            Begin VB.Label Label1
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0FFC0&
               Caption         =   "Magnitude DB"
               BeginProperty Font
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   204
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   285
               Index           =   0
               Left            =   405
               TabIndex        =   178
               ToolTipText     =   "Magnitude ratio in dB"
               Top             =   2700
               Width           =   1365
            End
         End
         Begin VB.Frame FrameLTFaveraging
            BorderStyle     =   0  'None
            Height          =   600
            Left            =   1170
            TabIndex        =   182
            Top             =   1755
            Width           =   1230
            Begin CWUIControlsLib.CWNumEdit cwLTF_time_scale
               Height          =   285
               Left            =   90
               TabIndex        =   183
               ToolTipText     =   "Averaging factor vs standard time (factor=1.0). Range 0.3 to 10, Test time reduced approximately 3 times or increased 10 times "
               Top             =   135
               Width           =   375
               _Version        =   524288
               _ExtentX        =   661
               _ExtentY        =   503
               _StockProps     =   4
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}
                  Name            =   "MS Sans Serif"
                  Size            =   9.76
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Reset_0         =   0   'False
               CompatibleVers_0=   524288
               NumEdit_0       =   1
               ClassName_1     =   "CCWNumEdit"
               opts_1          =   196670
               BorderStyle_1   =   1
               format_1        =   2
               ClassName_2     =   "CCWFormat"
               Format_2        =   ".0"
               scale_1         =   3
               ClassName_3     =   "CCWScale"
               dMax_3          =   10
               discInterval_3  =   0.1
               ValueVarType_1  =   5
               Value_Val_1     =   1
               IncValueVarType_1=   5
               IncValue_Val_1  =   0.1
               AccelIncVarType_1=   5
               AccelInc_Val_1  =   5
               RangeMinVarType_1=   5
               RangeMin_Val_1  =   0.3
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
               Caption         =   "TF averaging"
               BeginProperty Font
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   420
               Index           =   39
               Left            =   495
               TabIndex        =   184
               Top             =   90
               Width           =   735
            End
         End
      End
      Begin VB.TextBox TextXunits
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00D2E6E6&
         BorderStyle     =   0  'None
         BeginProperty Font
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   5850
         TabIndex        =   158
         Text            =   "Frequency"
         Top             =   3870
         Width           =   960
      End
      Begin VB.TextBox TextXunits
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00D2E6E6&
         BorderStyle     =   0  'None
         BeginProperty Font
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   5895
         TabIndex        =   157
         Text            =   "Frequency"
         Top             =   7965
         Width           =   960
      End
      Begin VB.CheckBox ChkReversePhase
         BackColor       =   &H008080FF&
         Caption         =   "Reverse Phase"
         Height          =   285
         Left            =   2700
         TabIndex        =   156
         ToolTipText     =   "When checked, 180 degrees is added to Phase plot"
         Top             =   4275
         Width           =   1545
      End
      Begin VB.Frame FrameCallTuning
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   0  'None
         Height          =   600
         Left            =   1080
         TabIndex        =   153
         Top             =   315
         Width           =   1590
         Begin VB.CommandButton CmdFilters
            BackColor       =   &H00C0C0C0&
            Caption         =   "Tuning"
            BeginProperty Font
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   465
            Index           =   0
            Left            =   45
            Style           =   1  'Graphical
            TabIndex        =   154
            ToolTipText     =   "Opens axis filters tuning window"
            Top             =   45
            Width           =   915
         End
         Begin VB.Image ImageTuning
            Height          =   465
            Left            =   990
            Picture         =   "Analyzer.frx":1DB2
            Stretch         =   -1  'True
            ToolTipText     =   "Opens axis filters tuning window"
            Top             =   45
            Width           =   465
         End
         Begin VB.Label LblTuning
            BackColor       =   &H00C0C0C0&
            Height          =   555
            Left            =   0
            TabIndex        =   155
            Top             =   0
            Width           =   1500
         End
      End
      Begin CWUIControlsLib.CWGraph CWGraph
         Height          =   4000
         Left            =   2700
         TabIndex        =   185
         ToolTipText     =   "Click - show cursor, dblClick - hide cursor"
         Top             =   170
         Width           =   8940
         _Version        =   524288
         _ExtentX        =   15769
         _ExtentY        =   7056
         _StockProps     =   71
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}
            Name            =   "Arial"
            Size            =   8.24
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Reset_0         =   0   'False
         CompatibleVers_0=   524288
         Graph_0         =   1
         ClassName_1     =   "CCWGraphFrame"
         opts_1          =   46
         C[0]_1          =   16777215
         C[1]_1          =   13821670
         C[2]_1          =   16777215
         Event_1         =   2
         ClassName_2     =   "CCWGFPlotEvent"
         Owner_2         =   1
         Plots_1         =   3
         ClassName_3     =   "CCWDataPlots"
         Array_3         =   8
         Editor_3        =   4
         ClassName_4     =   "CCWGFPlotArrayEditor"
         Owner_4         =   1
         Array[0]_3      =   5
         ClassName_5     =   "CCWDataPlot"
         opts_5          =   4194367
         Name_5          =   "Plot-Current"
         C[0]_5          =   16711680
         C[1]_5          =   255
         C[2]_5          =   16711680
         C[3]_5          =   16776960
         Event_5         =   2
         X_5             =   6
         ClassName_6     =   "CCWAxis"
         opts_6          =   575
         Name_6          =   "XAxis"
         C[3]_6          =   0
         C[4]_6          =   0
         Orientation_6   =   2944
         LabelWidth_6    =   8
         format_6        =   7
         ClassName_7     =   "CCWFormat"
         Format_7        =   ".#"
         Scale_6         =   8
         ClassName_8     =   "CCWScale"
         opts_8          =   94208
         rMin_8          =   45
         rMax_8          =   567
         dMin_8          =   0.1
         dMax_8          =   1000
         discInterval_8  =   1
         discBase_8      =   100
         Radial_6        =   0
         Enum_6          =   9
         ClassName_9     =   "CCWEnum"
         Editor_9        =   10
         ClassName_10    =   "CCWEnumArrayEditor"
         Owner_10        =   6
         Font_6          =   0
         tickopts_6      =   1759
         major_6         =   4
         minor_6         =   9
         Caption_6       =   11
         ClassName_11    =   "CCWDrawObj"
         opts_11         =   62
         C[0]_11         =   0
         Image_11        =   12
         ClassName_12    =   "CCWTextImage"
         style_12        =   1076101120
         font_12         =   0
         Animator_11     =   0
         Blinker_11      =   0
         Y_5             =   13
         ClassName_13    =   "CCWAxis"
         opts_13         =   575
         Name_13         =   "Y"
         C[2]_13         =   16711680
         C[3]_13         =   0
         C[4]_13         =   0
         Orientation_13  =   2067
         LabelWidth_13   =   3
         format_13       =   14
         ClassName_14    =   "CCWFormat"
         Scale_13        =   15
         ClassName_15    =   "CCWScale"
         opts_15         =   122880
         rMin_15         =   25
         rMax_15         =   239
         dMin_15         =   -40
         dMax_15         =   20
         discInterval_15 =   1
         Radial_13       =   0
         Enum_13         =   16
         ClassName_16    =   "CCWEnum"
         Editor_16       =   17
         ClassName_17    =   "CCWEnumArrayEditor"
         Owner_17        =   13
         Font_13         =   0
         tickopts_13     =   2751
         major_13        =   10
         minor_13        =   5
         Caption_13      =   18
         ClassName_18    =   "CCWDrawObj"
         opts_18         =   62
         C[0]_18         =   16711680
         Image_18        =   19
         ClassName_19    =   "CCWTextImage"
         szText_19       =   "Gain"
         style_19        =   1076101120
         font_19         =   0
         Animator_18     =   0
         Blinker_18      =   0
         LineStyle_5     =   1
         LineWidth_5     =   3
         BaseVal_5       =   -3.40282346638529E+38
         BasePlot_5      =   0
         DefaultXInc_5   =   1
         DefaultPlotPerRow_5=   -1  'True
         Array[1]_3      =   20
         ClassName_20    =   "CCWDataPlot"
         opts_20         =   4194367
         Name_20         =   "Plot-Previous"
         C[0]_20         =   32768
         C[1]_20         =   255
         C[2]_20         =   16711680
         C[3]_20         =   16776960
         Event_20        =   2
         X_20            =   6
         Y_20            =   13
         LineStyle_20    =   1
         LineWidth_20    =   2
         BaseVal_20      =   -3.40282346638529E+38
         BasePlot_20     =   0
         DefaultXInc_20  =   1
         DefaultPlotPerRow_20=   -1  'True
         Array[2]_3      =   21
         ClassName_21    =   "CCWDataPlot"
         opts_21         =   4194367
         Name_21         =   "Plot-Reference"
         C[0]_21         =   4227327
         C[1]_21         =   16711680
         C[2]_21         =   16711680
         C[3]_21         =   16776960
         Event_21        =   2
         X_21            =   6
         Y_21            =   13
         LineStyle_21    =   1
         LineWidth_21    =   2
         BaseVal_21      =   -3.40282346638529E+38
         BasePlot_21     =   0
         DefaultXInc_21  =   1
         DefaultPlotPerRow_21=   -1  'True
         Array[3]_3      =   22
         ClassName_22    =   "CCWDataPlot"
         opts_22         =   4194367
         Name_22         =   "Plot-OneFilterTF"
         C[0]_22         =   16711935
         C[1]_22         =   255
         C[2]_22         =   16711680
         C[3]_22         =   16776960
         Event_22        =   2
         X_22            =   6
         Y_22            =   13
         LineStyle_22    =   1
         LineWidth_22    =   2
         BasePlot_22     =   0
         DefaultXInc_22  =   1
         DefaultPlotPerRow_22=   -1  'True
         Array[4]_3      =   23
         ClassName_23    =   "CCWDataPlot"
         opts_23         =   4194367
         Name_23         =   "Plot-AxisTF"
         C[0]_23         =   8388736
         C[1]_23         =   255
         C[2]_23         =   16711680
         C[3]_23         =   16776960
         Event_23        =   2
         X_23            =   6
         Y_23            =   13
         LineStyle_23    =   1
         LineWidth_23    =   2
         BasePlot_23     =   0
         DefaultXInc_23  =   1
         DefaultPlotPerRow_23=   -1  'True
         Array[5]_3      =   24
         ClassName_24    =   "CCWDataPlot"
         opts_24         =   4194367
         Name_24         =   "Plot-Prediction"
         C[0]_24         =   12877824
         C[1]_24         =   8421376
         C[2]_24         =   16711680
         C[3]_24         =   16776960
         Event_24        =   2
         X_24            =   6
         Y_24            =   13
         LineStyle_24    =   1
         LineWidth_24    =   2
         BasePlot_24     =   0
         DefaultXInc_24  =   1
         DefaultPlotPerRow_24=   -1  'True
         Array[6]_3      =   25
         ClassName_25    =   "CCWDataPlot"
         opts_25         =   4194367
         Name_25         =   "Plot-lower_lim"
         C[0]_25         =   0
         C[1]_25         =   255
         C[2]_25         =   16711680
         C[3]_25         =   16776960
         Event_25        =   2
         X_25            =   6
         Y_25            =   13
         LineStyle_25    =   1
         LineWidth_25    =   2
         BasePlot_25     =   0
         DefaultXInc_25  =   1
         DefaultPlotPerRow_25=   -1  'True
         Array[7]_3      =   26
         ClassName_26    =   "CCWDataPlot"
         opts_26         =   4194367
         Name_26         =   "Plot-upper_lim"
         C[0]_26         =   0
         C[1]_26         =   255
         C[2]_26         =   16711680
         C[3]_26         =   16776960
         Event_26        =   2
         X_26            =   6
         Y_26            =   13
         LineStyle_26    =   1
         LineWidth_26    =   2
         BasePlot_26     =   0
         DefaultXInc_26  =   1
         DefaultPlotPerRow_26=   -1  'True
         Axes_1          =   27
         ClassName_27    =   "CCWAxes"
         Array_27        =   2
         Editor_27       =   28
         ClassName_28    =   "CCWGFAxisArrayEditor"
         Owner_28        =   1
         Array[0]_27     =   6
         Array[1]_27     =   13
         DefaultPlot_1   =   29
         ClassName_29    =   "CCWDataPlot"
         opts_29         =   4194367
         Name_29         =   "[Template]"
         C[0]_29         =   65280
         C[1]_29         =   255
         C[2]_29         =   16711680
         C[3]_29         =   16776960
         Event_29        =   2
         X_29            =   6
         Y_29            =   13
         LineStyle_29    =   1
         LineWidth_29    =   1
         BasePlot_29     =   0
         DefaultXInc_29  =   1
         DefaultPlotPerRow_29=   -1  'True
         Cursors_1       =   30
         ClassName_30    =   "CCWCursors"
         Array_30        =   1
         Editor_30       =   31
         ClassName_31    =   "CCWGFCursorArrayEditor"
         Owner_31        =   1
         Array[0]_30     =   32
         ClassName_32    =   "CCWCursor"
         opts_32         =   61
         Name_32         =   "Cursor-1"
         C[0]_32         =   255
         Event_32        =   2
         X_32            =   6
         Y_32            =   13
         XPos_32         =   97.3880672847497
         YPos_32         =   34.5685085075867
         SnapMode_32     =   4
         PointIndex_32   =   -1
         ChrosshairStyle_32=   8
         PointStyle_32   =   9
         LockPlot_32     =   5
         TrackMode_1     =   2
         GraphBackground_1=   0
         GraphFrame_1    =   33
         ClassName_33    =   "CCWDrawObj"
         opts_33         =   62
         C[0]_33         =   13821670
         C[1]_33         =   13821670
         Image_33        =   34
         ClassName_34    =   "CCWPictImage"
         opts_34         =   1280
         Rows_34         =   1
         Cols_34         =   1
         F_34            =   13821670
         B_34            =   13821670
         ColorReplaceWith_34=   8421504
         ColorReplace_34 =   8421504
         Tolerance_34    =   2
         Animator_33     =   0
         Blinker_33      =   0
         PlotFrame_1     =   35
         ClassName_35    =   "CCWDrawObj"
         opts_35         =   62
         C[0]_35         =   13821670
         C[1]_35         =   16777215
         Image_35        =   36
         ClassName_36    =   "CCWPictImage"
         opts_36         =   1280
         Rows_36         =   1
         Cols_36         =   1
         Pict_36         =   1
         F_36            =   13821670
         B_36            =   16777215
         ColorReplaceWith_36=   8421504
         ColorReplace_36 =   8421504
         Tolerance_36    =   2
         Animator_35     =   0
         Blinker_35      =   0
         Caption_1       =   37
         ClassName_37    =   "CCWDrawObj"
         opts_37         =   62
         C[0]_37         =   16711680
         Image_37        =   38
         ClassName_38    =   "CCWTextImage"
         szText_38       =   "Freq=1.00, Gain=0.00"
         style_38        =   1076101120
         font_38         =   0
         Animator_37     =   0
         Blinker_37      =   0
         DefaultXInc_1   =   1
         DefaultPlotPerRow_1=   -1  'True
         Bindings_1      =   39
         ClassName_39    =   "CCWBindingHolderArray"
         Editor_39       =   40
         ClassName_40    =   "CCWBindingHolderArrayEditor"
         Owner_40        =   1
         Annotations_1   =   41
         ClassName_41    =   "CCWAnnotations"
         Editor_41       =   42
         ClassName_42    =   "CCWAnnotationArrayEditor"
         Owner_42        =   1
         AnnotationTemplate_1=   43
         ClassName_43    =   "CCWAnnotation"
         opts_43         =   63
         Name_43         =   "[Template]"
         Plot_43         =   29
         Text_43         =   "[Template]"
         TextXPoint_43   =   670.033
         TextYPoint_43   =   670.033
         TextColor_43    =   16777215
         TextFont_43     =   44
         ClassName_44    =   "CCWFont"
         bFont_44        =   -1  'True
         BeginProperty Font_44 {0BE35203-8F91-11CE-9DE3-00AA004BB851}
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ShapeXPoints_43 =   45
         ClassName_45    =   "CDataBuffer"
         Type_45         =   5
         m_cDims;_45     =   1
         m_cElts_45      =   1
         Element[0]_45   =   330.067
         ShapeYPoints_43 =   46
         ClassName_46    =   "CDataBuffer"
         Type_46         =   5
         m_cDims;_46     =   1
         m_cElts_46      =   1
         Element[0]_46   =   330.067
         ShapeFillColor_43=   16777215
         ShapeLineColor_43=   16777215
         ShapeLineWidth_43=   1
         ShapeLineStyle_43=   1
         ShapePointStyle_43=   10
         ShapeImage_43   =   47
         ClassName_47    =   "CCWDrawObj"
         opts_47         =   62
         Image_47        =   48
         ClassName_48    =   "CCWPictImage"
         opts_48         =   1280
         Rows_48         =   1
         Cols_48         =   1
         Pict_48         =   7
         F_48            =   -2147483633
         B_48            =   -2147483633
         ColorReplaceWith_48=   8421504
         ColorReplace_48 =   8421504
         Tolerance_48    =   2
         Animator_47     =   0
         Blinker_47      =   0
         ArrowVisible_43 =   -1  'True
         ArrowColor_43   =   16777215
         ArrowWidth_43   =   1
         ArrowLineStyle_43=   1
         ArrowHeadStyle_43=   1
      End
      Begin CWUIControlsLib.CWGraph CWGraphPhase
         Height          =   4000
         Left            =   2700
         TabIndex        =   219
         ToolTipText     =   "Click - show cursor, dblClick - hide cursor"
         Top             =   4275
         Width           =   8940
         _Version        =   524288
         _ExtentX        =   15769
         _ExtentY        =   7056
         _StockProps     =   71
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}
            Name            =   "Arial"
            Size            =   8.24
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Reset_0         =   0   'False
         CompatibleVers_0=   524288
         Graph_0         =   1
         ClassName_1     =   "CCWGraphFrame"
         opts_1          =   62
         C[0]_1          =   16777215
         C[1]_1          =   13821670
         C[2]_1          =   16777215
         Event_1         =   2
         ClassName_2     =   "CCWGFPlotEvent"
         Owner_2         =   1
         Plots_1         =   3
         ClassName_3     =   "CCWDataPlots"
         Array_3         =   10
         Editor_3        =   4
         ClassName_4     =   "CCWGFPlotArrayEditor"
         Owner_4         =   1
         Array[0]_3      =   5
         ClassName_5     =   "CCWDataPlot"
         opts_5          =   4194367
         Name_5          =   "Plot-Current"
         C[0]_5          =   16711680
         C[1]_5          =   255
         C[2]_5          =   16711680
         C[3]_5          =   16776960
         Event_5         =   2
         X_5             =   6
         ClassName_6     =   "CCWAxis"
         opts_6          =   575
         Name_6          =   "XAxis"
         C[3]_6          =   0
         C[4]_6          =   0
         Orientation_6   =   2944
         LabelWidth_6    =   8
         format_6        =   7
         ClassName_7     =   "CCWFormat"
         Format_7        =   ".#"
         Scale_6         =   8
         ClassName_8     =   "CCWScale"
         opts_8          =   94208
         rMin_8          =   45
         rMax_8          =   567
         dMin_8          =   0.1
         dMax_8          =   1000
         discInterval_8  =   1
         discBase_8      =   100
         Radial_6        =   0
         Enum_6          =   9
         ClassName_9     =   "CCWEnum"
         Editor_9        =   10
         ClassName_10    =   "CCWEnumArrayEditor"
         Owner_10        =   6
         Font_6          =   0
         tickopts_6      =   1759
         major_6         =   4
         minor_6         =   9
         Caption_6       =   11
         ClassName_11    =   "CCWDrawObj"
         opts_11         =   62
         C[0]_11         =   -2147483640
         Image_11        =   12
         ClassName_12    =   "CCWTextImage"
         style_12        =   1076101120
         font_12         =   0
         Animator_11     =   0
         Blinker_11      =   0
         Y_5             =   13
         ClassName_13    =   "CCWAxis"
         opts_13         =   575
         Name_13         =   "Y"
         C[2]_13         =   255
         C[3]_13         =   0
         C[4]_13         =   0
         Orientation_13  =   2067
         LabelWidth_13   =   3
         format_13       =   14
         ClassName_14    =   "CCWFormat"
         Scale_13        =   15
         ClassName_15    =   "CCWScale"
         opts_15         =   122880
         rMin_15         =   25
         rMax_15         =   239
         dMin_15         =   -180
         dMax_15         =   180
         discInterval_15 =   1
         Radial_13       =   0
         Enum_13         =   16
         ClassName_16    =   "CCWEnum"
         Editor_16       =   17
         ClassName_17    =   "CCWEnumArrayEditor"
         Owner_17        =   13
         Font_13         =   0
         tickopts_13     =   2759
         major_13        =   60
         minor_13        =   30
         Caption_13      =   18
         ClassName_18    =   "CCWDrawObj"
         opts_18         =   62
         C[0]_18         =   255
         Image_18        =   19
         ClassName_19    =   "CCWTextImage"
         szText_19       =   "Phase"
         style_19        =   1076101120
         font_19         =   0
         Animator_18     =   0
         Blinker_18      =   0
         LineStyle_5     =   1
         LineWidth_5     =   3
         BasePlot_5      =   0
         DefaultXInc_5   =   1
         DefaultPlotPerRow_5=   -1  'True
         Array[1]_3      =   20
         ClassName_20    =   "CCWDataPlot"
         opts_20         =   4194367
         Name_20         =   "Plot-Previous"
         C[0]_20         =   32768
         C[1]_20         =   255
         C[2]_20         =   16711680
         C[3]_20         =   16776960
         Event_20        =   2
         X_20            =   6
         Y_20            =   13
         LineStyle_20    =   1
         LineWidth_20    =   2
         BasePlot_20     =   0
         DefaultXInc_20  =   1
         DefaultPlotPerRow_20=   -1  'True
         Array[2]_3      =   21
         ClassName_21    =   "CCWDataPlot"
         opts_21         =   4194367
         Name_21         =   "Plot-Ref"
         C[0]_21         =   4227327
         C[1]_21         =   16711680
         C[2]_21         =   16711680
         C[3]_21         =   16776960
         Event_21        =   2
         X_21            =   6
         Y_21            =   13
         LineStyle_21    =   1
         LineWidth_21    =   2
         BasePlot_21     =   0
         DefaultXInc_21  =   1
         DefaultPlotPerRow_21=   -1  'True
         Array[3]_3      =   22
         ClassName_22    =   "CCWDataPlot"
         opts_22         =   4194367
         Name_22         =   "Plot-OneFilterTF"
         C[0]_22         =   16711935
         C[1]_22         =   255
         C[2]_22         =   16711680
         C[3]_22         =   16776960
         Event_22        =   2
         X_22            =   6
         Y_22            =   13
         LineStyle_22    =   1
         LineWidth_22    =   2
         BasePlot_22     =   0
         DefaultXInc_22  =   1
         DefaultPlotPerRow_22=   -1  'True
         Array[4]_3      =   23
         ClassName_23    =   "CCWDataPlot"
         opts_23         =   4194367
         Name_23         =   "Plot-Axis_TF"
         C[0]_23         =   8388736
         C[1]_23         =   255
         C[2]_23         =   16711680
         C[3]_23         =   16776960
         Event_23        =   2
         X_23            =   6
         Y_23            =   13
         LineStyle_23    =   1
         LineWidth_23    =   2
         BasePlot_23     =   0
         DefaultXInc_23  =   1
         DefaultPlotPerRow_23=   -1  'True
         Array[5]_3      =   24
         ClassName_24    =   "CCWDataPlot"
         opts_24         =   4194367
         Name_24         =   "Plot-Prediction"
         C[0]_24         =   12877824
         C[1]_24         =   16711680
         C[2]_24         =   16711680
         C[3]_24         =   16776960
         Event_24        =   2
         X_24            =   6
         Y_24            =   13
         LineStyle_24    =   1
         LineWidth_24    =   2
         BasePlot_24     =   0
         DefaultXInc_24  =   1
         DefaultPlotPerRow_24=   -1  'True
         Array[6]_3      =   25
         ClassName_25    =   "CCWDataPlot"
         opts_25         =   4194367
         Name_25         =   "Plot-lower_lim"
         C[0]_25         =   0
         C[1]_25         =   255
         C[2]_25         =   16711680
         C[3]_25         =   16776960
         Event_25        =   2
         X_25            =   6
         Y_25            =   13
         LineStyle_25    =   1
         LineWidth_25    =   2
         BasePlot_25     =   0
         DefaultXInc_25  =   1
         DefaultPlotPerRow_25=   -1  'True
         Array[7]_3      =   26
         ClassName_26    =   "CCWDataPlot"
         opts_26         =   4194367
         Name_26         =   "Plot-upper_lim"
         C[0]_26         =   0
         C[1]_26         =   255
         C[2]_26         =   16711680
         C[3]_26         =   16776960
         Event_26        =   2
         X_26            =   6
         Y_26            =   13
         LineStyle_26    =   1
         LineWidth_26    =   2
         BasePlot_26     =   0
         DefaultXInc_26  =   1
         DefaultPlotPerRow_26=   -1  'True
         Array[8]_3      =   27
         ClassName_27    =   "CCWDataPlot"
         opts_27         =   4194367
         Name_27         =   "Plot-Coherence"
         C[0]_27         =   255
         C[1]_27         =   255
         C[2]_27         =   16711680
         C[3]_27         =   16776960
         Event_27        =   2
         X_27            =   6
         Y_27            =   13
         LineStyle_27    =   1
         LineWidth_27    =   2
         BasePlot_27     =   0
         DefaultXInc_27  =   1
         DefaultPlotPerRow_27=   -1  'True
         Array[9]_3      =   28
         ClassName_28    =   "CCWDataPlot"
         opts_28         =   4194367
         Name_28         =   "Plot-PrevConerence"
         C[0]_28         =   128
         C[1]_28         =   255
         C[2]_28         =   16711680
         C[3]_28         =   16776960
         Event_28        =   2
         X_28            =   6
         Y_28            =   13
         LineStyle_28    =   1
         LineWidth_28    =   2
         BasePlot_28     =   0
         DefaultXInc_28  =   1
         DefaultPlotPerRow_28=   -1  'True
         Axes_1          =   29
         ClassName_29    =   "CCWAxes"
         Array_29        =   2
         Editor_29       =   30
         ClassName_30    =   "CCWGFAxisArrayEditor"
         Owner_30        =   1
         Array[0]_29     =   6
         Array[1]_29     =   13
         DefaultPlot_1   =   31
         ClassName_31    =   "CCWDataPlot"
         opts_31         =   4194367
         Name_31         =   "[Template]"
         C[0]_31         =   65280
         C[1]_31         =   255
         C[2]_31         =   16711680
         C[3]_31         =   16776960
         Event_31        =   2
         X_31            =   6
         Y_31            =   13
         LineStyle_31    =   1
         LineWidth_31    =   1
         BasePlot_31     =   0
         DefaultXInc_31  =   1
         DefaultPlotPerRow_31=   -1  'True
         Cursors_1       =   32
         ClassName_32    =   "CCWCursors"
         Array_32        =   1
         Editor_32       =   33
         ClassName_33    =   "CCWGFCursorArrayEditor"
         Owner_33        =   1
         Array[0]_32     =   34
         ClassName_34    =   "CCWCursor"
         opts_34         =   61
         Name_34         =   "Cursor-Phase"
         C[0]_34         =   255
         Event_34        =   2
         X_34            =   6
         Y_34            =   13
         SnapMode_34     =   2
         ChrosshairStyle_34=   8
         PointStyle_34   =   9
         LockPlot_34     =   5
         TrackMode_1     =   2
         GraphBackground_1=   0
         GraphFrame_1    =   35
         ClassName_35    =   "CCWDrawObj"
         opts_35         =   62
         C[0]_35         =   13821670
         C[1]_35         =   13821670
         Image_35        =   36
         ClassName_36    =   "CCWPictImage"
         opts_36         =   1280
         Rows_36         =   1
         Cols_36         =   1
         F_36            =   13821670
         B_36            =   13821670
         ColorReplaceWith_36=   8421504
         ColorReplace_36 =   8421504
         Tolerance_36    =   2
         Animator_35     =   0
         Blinker_35      =   0
         PlotFrame_1     =   37
         ClassName_37    =   "CCWDrawObj"
         opts_37         =   62
         C[0]_37         =   13821670
         C[1]_37         =   16777215
         Image_37        =   38
         ClassName_38    =   "CCWPictImage"
         opts_38         =   1280
         Rows_38         =   1
         Cols_38         =   1
         Pict_38         =   1
         F_38            =   13821670
         B_38            =   16777215
         ColorReplaceWith_38=   8421504
         ColorReplace_38 =   8421504
         Tolerance_38    =   2
         Animator_37     =   0
         Blinker_37      =   0
         Caption_1       =   39
         ClassName_39    =   "CCWDrawObj"
         opts_39         =   62
         C[0]_39         =   255
         Image_39        =   40
         ClassName_40    =   "CCWTextImage"
         szText_40       =   "Freq=1.00, Phase=0.00"
         style_40        =   1076101120
         font_40         =   0
         Animator_39     =   0
         Blinker_39      =   0
         DefaultXInc_1   =   1
         DefaultPlotPerRow_1=   -1  'True
         Bindings_1      =   41
         ClassName_41    =   "CCWBindingHolderArray"
         Editor_41       =   42
         ClassName_42    =   "CCWBindingHolderArrayEditor"
         Owner_42        =   1
         Annotations_1   =   43
         ClassName_43    =   "CCWAnnotations"
         Editor_43       =   44
         ClassName_44    =   "CCWAnnotationArrayEditor"
         Owner_44        =   1
         AnnotationTemplate_1=   45
         ClassName_45    =   "CCWAnnotation"
         opts_45         =   63
         Name_45         =   "[Template]"
         Plot_45         =   31
         Text_45         =   "[Template]"
         TextXPoint_45   =   670.033
         TextYPoint_45   =   670.033
         TextColor_45    =   16777215
         TextFont_45     =   46
         ClassName_46    =   "CCWFont"
         bFont_46        =   -1  'True
         BeginProperty Font_46 {0BE35203-8F91-11CE-9DE3-00AA004BB851}
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ShapeXPoints_45 =   47
         ClassName_47    =   "CDataBuffer"
         Type_47         =   5
         m_cDims;_47     =   1
         m_cElts_47      =   1
         Element[0]_47   =   330.067
         ShapeYPoints_45 =   48
         ClassName_48    =   "CDataBuffer"
         Type_48         =   5
         m_cDims;_48     =   1
         m_cElts_48      =   1
         Element[0]_48   =   330.067
         ShapeFillColor_45=   16777215
         ShapeLineColor_45=   16777215
         ShapeLineWidth_45=   1
         ShapeLineStyle_45=   1
         ShapePointStyle_45=   10
         ShapeImage_45   =   49
         ClassName_49    =   "CCWDrawObj"
         opts_49         =   62
         Image_49        =   50
         ClassName_50    =   "CCWPictImage"
         opts_50         =   1280
         Rows_50         =   1
         Cols_50         =   1
         Pict_50         =   7
         F_50            =   -2147483633
         B_50            =   -2147483633
         ColorReplaceWith_50=   8421504
         ColorReplace_50 =   8421504
         Tolerance_50    =   2
         Animator_49     =   0
         Blinker_49      =   0
         ArrowVisible_45 =   -1  'True
         ArrowColor_45   =   16777215
         ArrowWidth_45   =   1
         ArrowLineStyle_45=   1
         ArrowHeadStyle_45=   1
      End
      Begin VB.Image ImagePerformance
         Height          =   600
         Left            =   135
         Picture         =   "Analyzer.frx":165F4
         Stretch         =   -1  'True
         Top             =   260
         Width           =   600
      End
   End
   Begin VB.CommandButton CmdScope
      BackColor       =   &H00E0FFE0&
      Caption         =   "Scope"
      BeginProperty Font
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   485
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   151
      Top             =   5625
      Width           =   1600
   End
   Begin VB.Frame frameAccessWindow
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      BeginProperty Font
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3030
      Left            =   -45
      TabIndex        =   38
      Top             =   2475
      Width           =   2785
      Begin VB.OptionButton optDialogWindow
         BackColor       =   &H00FFFFC0&
         Caption         =   "Performance"
         BeginProperty Font
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Index           =   2
         Left            =   660
         Style           =   1  'Graphical
         TabIndex        =   44
         ToolTipText     =   "Access to Loop Transfer function Test"
         Top             =   1305
         Width           =   2000
      End
      Begin VB.OptionButton optDialogWindow
         BackColor       =   &H008080FF&
         Caption         =   "Alarms"
         BeginProperty Font
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Index           =   5
         Left            =   660
         MaskColor       =   &H8000000F&
         Style           =   1  'Graphical
         TabIndex        =   43
         ToolTipText     =   "Shows Alarm page"
         Top             =   3045
         Visible         =   0   'False
         Width           =   2000
      End
      Begin VB.OptionButton optDialogWindow
         BackColor       =   &H0080FFFF&
         Caption         =   "System Control"
         BeginProperty Font
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Index           =   1
         Left            =   660
         Style           =   1  'Graphical
         TabIndex        =   42
         ToolTipText     =   "Access to System Control / Config "
         Top             =   720
         Value           =   -1  'True
         Width           =   2000
      End
      Begin VB.OptionButton optDialogWindow
         BackColor       =   &H00FFB0B0&
         Caption         =   "Diagnostic-tuning"
         BeginProperty Font
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Index           =   3
         Left            =   660
         Style           =   1  'Graphical
         TabIndex        =   41
         ToolTipText     =   "Access to Diagnostic & Tuning Features"
         Top             =   1875
         Width           =   2000
      End
      Begin VB.OptionButton optDialogWindow
         BackColor       =   &H0070D0FF&
         Caption         =   "Firmware Params"
         BeginProperty Font
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Index           =   4
         Left            =   660
         MaskColor       =   &H8000000F&
         Style           =   1  'Graphical
         TabIndex        =   40
         ToolTipText     =   "Access to Upgrading Firmware"
         Top             =   2460
         Width           =   2000
      End
      Begin VB.OptionButton optDialogWindow
         BackColor       =   &H0080FF80&
         Caption         =   "Communication"
         BeginProperty Font
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Index           =   0
         Left            =   660
         MaskColor       =   &H8000000F&
         Style           =   1  'Graphical
         TabIndex        =   39
         ToolTipText     =   "Shows Alarm Log"
         Top             =   135
         Width           =   2000
      End
      Begin VB.Image ImagePictogram
         Height          =   495
         Index           =   0
         Left            =   60
         Picture         =   "Analyzer.frx":1B432
         Stretch         =   -1  'True
         Top             =   60
         Width           =   555
      End
      Begin VB.Image ImagePictogram
         Height          =   495
         Index           =   1
         Left            =   60
         Picture         =   "Analyzer.frx":2F4B4
         Stretch         =   -1  'True
         Top             =   660
         Width           =   495
      End
      Begin VB.Image ImagePictogram
         Height          =   495
         Index           =   4
         Left            =   60
         Picture         =   "Analyzer.frx":51936
         Stretch         =   -1  'True
         Top             =   2415
         Width           =   495
      End
      Begin VB.Image ImagePictogram
         Height          =   495
         Index           =   5
         Left            =   60
         Picture         =   "Analyzer.frx":719B8
         Stretch         =   -1  'True
         Top             =   3000
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Image ImagePictogram
         Height          =   495
         Index           =   2
         Left            =   60
         Picture         =   "Analyzer.frx":79A3A
         Stretch         =   -1  'True
         Top             =   1245
         Width           =   495
      End
      Begin VB.Image ImagePictogram
         Height          =   495
         Index           =   3
         Left            =   60
         Picture         =   "Analyzer.frx":7E878
         Stretch         =   -1  'True
         Top             =   1830
         Width           =   495
      End
      Begin VB.Label LabelTabSelect
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   555
         Index           =   0
         Left            =   45
         TabIndex        =   45
         Top             =   45
         Width           =   2805
      End
      Begin VB.Label LabelTabSelect
         Appearance      =   0  'Flat
         BackColor       =   &H0080E0FF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   555
         Index           =   4
         Left            =   45
         TabIndex        =   50
         Top             =   2385
         Visible         =   0   'False
         Width           =   2805
      End
      Begin VB.Label LabelTabSelect
         Appearance      =   0  'Flat
         BackColor       =   &H00FFE0E0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   555
         Index           =   3
         Left            =   45
         TabIndex        =   49
         Top             =   1800
         Visible         =   0   'False
         Width           =   2805
      End
      Begin VB.Label LabelTabSelect
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   555
         Index           =   2
         Left            =   45
         TabIndex        =   48
         Top             =   1215
         Visible         =   0   'False
         Width           =   2805
      End
      Begin VB.Label LabelTabSelect
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   555
         Index           =   5
         Left            =   45
         TabIndex        =   47
         Top             =   2970
         Visible         =   0   'False
         Width           =   2805
      End
      Begin VB.Label LabelTabSelect
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   555
         Index           =   1
         Left            =   45
         TabIndex        =   46
         Top             =   630
         Visible         =   0   'False
         Width           =   2805
      End
   End
   Begin VB.Timer timeTimer
      Interval        =   100
      Left            =   630
      Top             =   450
   End
   Begin VB.Timer TimerTestProgram
      Interval        =   100
      Left            =   1890
      Top             =   450
   End
   Begin MSComDlg.CommonDialog dlgCommonDialog
      Left            =   45
      Top             =   405
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.Frame FrameDiagnosticExcitation
      BackColor       =   &H00FFFFFF&
      BeginProperty Font
         Name            =   "Arial"
         Size            =   13.5
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1635
      Index           =   0
      Left            =   14715
      TabIndex        =   83
      Top             =   6795
      Visible         =   0   'False
      Width           =   2580
      Begin VB.CommandButton cmdSetControllerTime
         BackColor       =   &H00E0E0E0&
         Caption         =   "Set Controller Time from PC time"
         BeginProperty Font
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   550
         Left            =   135
         Style           =   1  'Graphical
         TabIndex        =   84
         ToolTipText     =   "Copies PC time to controller time"
         Top             =   945
         Width           =   2310
      End
      Begin VB.Label LblTemperature
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "+28"
         BeginProperty Font
            Name            =   "Arial"
            Size            =   10.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   1845
         TabIndex        =   89
         ToolTipText     =   "How long it takes to response on the last command"
         Top             =   585
         Width           =   600
      End
      Begin VB.Label Label1
         BackColor       =   &H00F0F0F0&
         Caption         =   "T_controller, degC"
         BeginProperty Font
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   27
         Left            =   135
         TabIndex        =   88
         Top             =   585
         Width           =   1680
      End
      Begin VB.Label Label3
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "?0.025"
         BeginProperty Font
            Name            =   "Arial"
            Size            =   10.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   1575
         TabIndex        =   87
         ToolTipText     =   "How long it takes to response on the last command"
         Top             =   270
         Width           =   870
      End
      Begin VB.Label Label1
         BackColor       =   &H00F0F0F0&
         Caption         =   "cmd exec. time"
         BeginProperty Font
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   28
         Left            =   135
         TabIndex        =   86
         Top             =   270
         Width           =   1365
      End
      Begin VB.Label Label1
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "miscellaneous"
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
         Height          =   345
         Index           =   20
         Left            =   405
         TabIndex        =   85
         ToolTipText     =   "3-D vector,   projection on Y axis"
         Top             =   0
         Width           =   1785
      End
   End
   Begin VB.Frame FrameButtons
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   2745
      TabIndex        =   118
      Top             =   7614
      Width           =   11805
      Begin VB.CommandButton CmdRestorefromFlash
         BackColor       =   &H00C0FFC0&
         Caption         =   "Restore Params from FLASH"
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
         Left            =   6120
         Style           =   1  'Graphical
         TabIndex        =   141
         ToolTipText     =   "Sends command to controller to restore parameters from FLASH (parameters must be already saved)"
         Top             =   90
         Width           =   1665
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
         Height          =   600
         Index           =   0
         Left            =   8100
         Style           =   1  'Graphical
         TabIndex        =   142
         ToolTipText     =   "Temporary sets controller parameters to factory default in RAM. Parameters in FLASH are not affected until ""save"" command"
         Top             =   90
         Width           =   1665
      End
      Begin VB.CommandButton cmdSaveParamsIntoFLASH
         BackColor       =   &H00FFFF80&
         Caption         =   "Save Parameters into FLASH"
         BeginProperty Font
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Index           =   0
         Left            =   10080
         Style           =   1  'Graphical
         TabIndex        =   143
         ToolTipText     =   "Saves controller parameters into internal non-volatile memory"
         Top             =   90
         Width           =   1665
      End
      Begin CWUIControlsLib.CWButton cwBut_FF_ALL
         Height          =   600
         Left            =   4005
         TabIndex        =   144
         Tag             =   "loop_ffa"
         ToolTipText     =   "Controls All Feed Forward loops"
         Top             =   90
         Width           =   1995
         _Version        =   524288
         _ExtentX        =   3519
         _ExtentY        =   1058
         _StockProps     =   68
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Reset_0         =   0   'False
         CompatibleVers_0=   524288
         Boolean_0       =   1
         ClassName_1     =   "CCWBoolean"
         opts_1          =   2614
         SclRef_1.l      =   2
         SclRef_1.t      =   2
         SclRef_1.r      =   135
         SclRef_1.b      =   42
         Scl_1.l         =   3
         Scl_1.t         =   5
         Scl_1.r         =   79
         Scl_1.b         =   39
         C[0]_1          =   -2147483643
         Enum_1          =   2
         ClassName_2     =   "CCWEnum"
         Array_2         =   2
         Editor_2        =   0
         Array[0]_2      =   3
         ClassName_3     =   "CCWEnumElt"
         opts_3          =   1
         Name_3          =   "Off"
         frame_3         =   288
         DrawList_3      =   4
         ClassName_4     =   "CDrawList"
         count_4         =   4
         list[4]_4       =   5
         ClassName_5     =   "CCWDrawObj"
         opts_5          =   54
         SclRef_5.l      =   5
         SclRef_5.t      =   9
         SclRef_5.r      =   74
         SclRef_5.b      =   36
         Scl_5.l         =   79
         Scl_5.t         =   15
         Scl_5.r         =   119
         Scl_5.b         =   31
         C[0]_5          =   -2147483640
         C[1]_5          =   -2147483640
         Image_5         =   6
         ClassName_6     =   "CCWTextImage"
         szText_6        =   "ALL FF"
         font_6          =   0
         Animator_5      =   0
         Blinker_5       =   0
         list[3]_4       =   7
         ClassName_7     =   "CCWDrawObj"
         opts_7          =   60
         C[0]_7          =   -2147483640
         C[1]_7          =   -2147483640
         Image_7         =   8
         ClassName_8     =   "CCWTextImage"
         font_8          =   0
         Animator_7      =   0
         Blinker_7       =   0
         list[2]_4       =   9
         ClassName_9     =   "CCWDrawObj"
         opts_9          =   62
         C[0]_9          =   -2147483640
         C[1]_9          =   -2147483640
         Image_9         =   10
         ClassName_10    =   "CCWTextImage"
         font_10         =   0
         Animator_9      =   0
         Blinker_9       =   0
         list[1]_4       =   11
         ClassName_11    =   "CCWDrawObj"
         opts_11         =   54
         SclRef_11.l     =   6
         SclRef_11.t     =   6
         SclRef_11.r     =   102
         SclRef_11.b     =   51
         Scl_11.l        =   4
         Scl_11.t        =   6
         Scl_11.r        =   102
         Scl_11.b        =   51
         Image_11        =   12
         ClassName_12    =   "CCWPictImage"
         Rows_12         =   1
         Cols_12         =   1
         URL_12          =   "\\tsclient\C\In_Out\Software\DC2020_WEB_Analyzer\OFF_gray_slider.bmp"
         Pict_12         =   "Analyzer.frx":9E8FA
         ColorReplaceWith_12=   8421504
         ColorReplace_12 =   8421504
         Tolerance_12    =   2
         Animator_11     =   0
         Blinker_11      =   0
         varVarType_3    =   5
         Array[1]_2      =   13
         ClassName_13    =   "CCWEnumElt"
         opts_13         =   1
         Name_13         =   "On"
         frame_13        =   287
         DrawList_13     =   14
         ClassName_14    =   "CDrawList"
         count_14        =   4
         list[4]_14      =   15
         ClassName_15    =   "CCWDrawObj"
         opts_15         =   54
         SclRef_15.l     =   5
         SclRef_15.t     =   9
         SclRef_15.r     =   74
         SclRef_15.b     =   36
         Scl_15.l        =   79
         Scl_15.t        =   15
         Scl_15.r        =   119
         Scl_15.b        =   31
         C[0]_15         =   -2147483640
         C[1]_15         =   -2147483640
         Image_15        =   6
         Animator_15     =   0
         Blinker_15      =   0
         list[3]_14      =   16
         ClassName_16    =   "CCWDrawObj"
         opts_16         =   62
         C[0]_16         =   -2147483640
         C[1]_16         =   -2147483640
         Image_16        =   8
         Animator_16     =   0
         Blinker_16      =   0
         list[2]_14      =   17
         ClassName_17    =   "CCWDrawObj"
         opts_17         =   60
         C[0]_17         =   -2147483640
         C[1]_17         =   -2147483640
         Image_17        =   10
         Animator_17     =   0
         Blinker_17      =   0
         list[1]_14      =   18
         ClassName_18    =   "CCWDrawObj"
         opts_18         =   54
         SclRef_18.l     =   6
         SclRef_18.t     =   6
         SclRef_18.r     =   102
         SclRef_18.b     =   51
         Scl_18.l        =   4
         Scl_18.t        =   6
         Scl_18.r        =   102
         Scl_18.b        =   51
         Image_18        =   19
         ClassName_19    =   "CCWPictImage"
         Rows_19         =   1
         Cols_19         =   1
         URL_19          =   "\\tsclient\C\In_Out\Software\DC2020_WEB_Analyzer\ON_blue_slider.bmp"
         Pict_19         =   "Analyzer.frx":B2644
         ColorReplaceWith_19=   8421504
         ColorReplace_19 =   8421504
         Tolerance_19    =   2
         Animator_18     =   0
         Blinker_18      =   0
         varVarType_13   =   5
         Bindings_1      =   20
         ClassName_20    =   "CCWBindingHolderArray"
         Editor_20       =   21
         ClassName_21    =   "CCWBindingHolderArrayEditor"
         Owner_21        =   1
         Style_1         =   14
         drawState_1     =   1
         frameStyle_1    =   1
         mechAction_1    =   1
         BGImg_1         =   22
         ClassName_22    =   "CCWDrawObj"
         opts_22         =   62
         Image_22        =   23
         ClassName_23    =   "CCWPictImage"
         opts_23         =   1280
         Rows_23         =   1
         Cols_23         =   1
         Pict_23         =   286
         F_23            =   -2147483633
         B_23            =   -2147483633
         ColorReplaceWith_23=   8421504
         ColorReplace_23 =   8421504
         Tolerance_23    =   2
         Animator_22     =   0
         Blinker_22      =   0
         Array_1         =   6
         Editor_1        =   0
         Array[0]_1      =   11
         Array[1]_1      =   18
         Array[2]_1      =   0
         Array[3]_1      =   0
         Array[4]_1      =   24
         ClassName_24    =   "CCWDrawObj"
         opts_24         =   62
         Image_24        =   8
         Animator_24     =   0
         Blinker_24      =   0
         Array[5]_1      =   25
         ClassName_25    =   "CCWDrawObj"
         opts_25         =   62
         Image_25        =   10
         Animator_25     =   0
         Blinker_25      =   0
         Label_1         =   26
         ClassName_26    =   "CCWDrawObj"
         opts_26         =   62
         C[0]_26         =   -2147483640
         Image_26        =   6
         Animator_26     =   0
         Blinker_26      =   0
      End
      Begin CWUIControlsLib.CWButton cw_Status_Bit_Indicator
         Height          =   645
         Index           =   1
         Left            =   0
         TabIndex        =   145
         ToolTipText     =   "Indicator shows payload is within or outside ""Go-NoGO"" position window"
         Top             =   45
         Width           =   975
         _Version        =   524288
         _ExtentX        =   1720
         _ExtentY        =   1138
         _StockProps     =   68
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Reset_0         =   0   'False
         CompatibleVers_0=   524288
         Boolean_0       =   1
         ClassName_1     =   "CCWBoolean"
         opts_1          =   2622
         C[0]_1          =   16761024
         Enum_1          =   2
         ClassName_2     =   "CCWEnum"
         Array_2         =   2
         Editor_2        =   0
         Array[0]_2      =   3
         ClassName_3     =   "CCWEnumElt"
         opts_3          =   1
         Name_3          =   "Off"
         frame_3         =   286
         DrawList_3      =   4
         ClassName_4     =   "CDrawList"
         count_4         =   4
         list[4]_4       =   5
         ClassName_5     =   "CCWDrawObj"
         opts_5          =   62
         C[0]_5          =   -2147483640
         C[1]_5          =   -2147483640
         Image_5         =   6
         ClassName_6     =   "CCWTextImage"
         szText_6        =   "Status"
         font_6          =   0
         Animator_5      =   0
         Blinker_5       =   0
         list[3]_4       =   7
         ClassName_7     =   "CCWDrawObj"
         opts_7          =   62
         C[0]_7          =   -2147483640
         C[1]_7          =   -2147483640
         Image_7         =   8
         ClassName_8     =   "CCWTextImage"
         szText_8        =   "GO"
         font_8          =   0
         Animator_7      =   0
         Blinker_7       =   0
         list[2]_4       =   9
         ClassName_9     =   "CCWDrawObj"
         opts_9          =   60
         C[0]_9          =   -2147483640
         C[1]_9          =   -2147483640
         Image_9         =   10
         ClassName_10    =   "CCWTextImage"
         szText_10       =   "NoGO"
         font_10         =   0
         Animator_9      =   0
         Blinker_9       =   0
         list[1]_4       =   11
         ClassName_11    =   "CCWDrawObj"
         opts_11         =   62
         C[0]_11         =   65280
         C[1]_11         =   65280
         Image_11        =   12
         ClassName_12    =   "CCWPictImage"
         opts_12         =   1280
         Rows_12         =   1
         Cols_12         =   1
         Pict_12         =   75
         F_12            =   65280
         B_12            =   65280
         ColorReplaceWith_12=   8421504
         ColorReplace_12 =   8421504
         Tolerance_12    =   2
         Animator_11     =   0
         Blinker_11      =   0
         varVarType_3    =   5
         Array[1]_2      =   13
         ClassName_13    =   "CCWEnumElt"
         opts_13         =   1
         Name_13         =   "On"
         frame_13        =   286
         DrawList_13     =   14
         ClassName_14    =   "CDrawList"
         count_14        =   4
         list[4]_14      =   15
         ClassName_15    =   "CCWDrawObj"
         opts_15         =   62
         C[0]_15         =   -2147483640
         C[1]_15         =   -2147483640
         Image_15        =   6
         Animator_15     =   0
         Blinker_15      =   0
         list[3]_14      =   16
         ClassName_16    =   "CCWDrawObj"
         opts_16         =   60
         C[0]_16         =   -2147483640
         C[1]_16         =   -2147483640
         Image_16        =   8
         Animator_16     =   0
         Blinker_16      =   0
         list[2]_14      =   17
         ClassName_17    =   "CCWDrawObj"
         opts_17         =   62
         C[0]_17         =   -2147483640
         C[1]_17         =   -2147483640
         Image_17        =   10
         Animator_17     =   0
         Blinker_17      =   0
         list[1]_14      =   18
         ClassName_18    =   "CCWDrawObj"
         opts_18         =   62
         C[0]_18         =   32767
         C[1]_18         =   32767
         Image_18        =   19
         ClassName_19    =   "CCWPictImage"
         opts_19         =   1280
         Rows_19         =   1
         Cols_19         =   1
         Pict_19         =   76
         F_19            =   32767
         B_19            =   32767
         ColorReplaceWith_19=   8421504
         ColorReplace_19 =   8421504
         Tolerance_19    =   2
         Animator_18     =   0
         Blinker_18      =   20
         ClassName_20    =   "CCWBlinker"
         Interval_20     =   600
         EventID_20      =   7
         Obj_20          =   18
         varVarType_13   =   5
         Bindings_1      =   21
         ClassName_21    =   "CCWBindingHolderArray"
         Editor_21       =   22
         ClassName_22    =   "CCWBindingHolderArrayEditor"
         Owner_22        =   1
         Style_1         =   5
         mechAction_1    =   3
         BGImg_1         =   23
         ClassName_23    =   "CCWDrawObj"
         opts_23         =   62
         Image_23        =   24
         ClassName_24    =   "CCWPictImage"
         opts_24         =   1280
         Rows_24         =   1
         Cols_24         =   1
         Pict_24         =   286
         F_24            =   -2147483633
         B_24            =   -2147483633
         ColorReplaceWith_24=   8421504
         ColorReplace_24 =   8421504
         Tolerance_24    =   2
         Animator_23     =   0
         Blinker_23      =   0
         Array_1         =   6
         Editor_1        =   0
         Array[0]_1      =   11
         Array[1]_1      =   18
         Array[2]_1      =   0
         Array[3]_1      =   0
         Array[4]_1      =   25
         ClassName_25    =   "CCWDrawObj"
         opts_25         =   62
         Image_25        =   8
         Animator_25     =   0
         Blinker_25      =   0
         Array[5]_1      =   26
         ClassName_26    =   "CCWDrawObj"
         opts_26         =   62
         Image_26        =   10
         Animator_26     =   0
         Blinker_26      =   0
         Label_1         =   27
         ClassName_27    =   "CCWDrawObj"
         opts_27         =   62
         C[0]_27         =   -2147483640
         Image_27        =   6
         Animator_27     =   0
         Blinker_27      =   0
      End
      Begin CWUIControlsLib.CWButton cw_Status_Bit_Indicator
         Height          =   645
         Index           =   2
         Left            =   1350
         TabIndex        =   146
         ToolTipText     =   "Indicator shows whether motor power was reduced due to too much motion and violating of safety limits"
         Top             =   45
         Width           =   975
         _Version        =   524288
         _ExtentX        =   1720
         _ExtentY        =   1138
         _StockProps     =   68
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Reset_0         =   0   'False
         CompatibleVers_0=   524288
         Boolean_0       =   1
         ClassName_1     =   "CCWBoolean"
         opts_1          =   2622
         C[0]_1          =   16761024
         Enum_1          =   2
         ClassName_2     =   "CCWEnum"
         Array_2         =   2
         Editor_2        =   0
         Array[0]_2      =   3
         ClassName_3     =   "CCWEnumElt"
         opts_3          =   1
         Name_3          =   "Off"
         frame_3         =   286
         DrawList_3      =   4
         ClassName_4     =   "CDrawList"
         count_4         =   4
         list[4]_4       =   5
         ClassName_5     =   "CCWDrawObj"
         opts_5          =   62
         C[0]_5          =   -2147483640
         C[1]_5          =   -2147483640
         Image_5         =   6
         ClassName_6     =   "CCWTextImage"
         szText_6        =   "LM safety"
         font_6          =   0
         Animator_5      =   0
         Blinker_5       =   0
         list[3]_4       =   7
         ClassName_7     =   "CCWDrawObj"
         opts_7          =   60
         C[0]_7          =   -2147483640
         C[1]_7          =   -2147483640
         Image_7         =   8
         ClassName_8     =   "CCWTextImage"
         szText_8        =   "Recover"
         font_8          =   0
         Animator_7      =   0
         Blinker_7       =   0
         list[2]_4       =   9
         ClassName_9     =   "CCWDrawObj"
         opts_9          =   62
         C[0]_9          =   -2147483640
         C[1]_9          =   -2147483640
         Image_9         =   10
         ClassName_10    =   "CCWTextImage"
         szText_10       =   "NORM"
         font_10         =   0
         Animator_9      =   0
         Blinker_9       =   0
         list[1]_4       =   11
         ClassName_11    =   "CCWDrawObj"
         opts_11         =   62
         C[0]_11         =   65280
         C[1]_11         =   65280
         Image_11        =   12
         ClassName_12    =   "CCWPictImage"
         opts_12         =   1280
         Rows_12         =   1
         Cols_12         =   1
         Pict_12         =   73
         F_12            =   65280
         B_12            =   65280
         ColorReplaceWith_12=   8421504
         ColorReplace_12 =   8421504
         Tolerance_12    =   2
         Animator_11     =   0
         Blinker_11      =   0
         varVarType_3    =   5
         Array[1]_2      =   13
         ClassName_13    =   "CCWEnumElt"
         opts_13         =   1
         Name_13         =   "On"
         frame_13        =   286
         DrawList_13     =   14
         ClassName_14    =   "CDrawList"
         count_14        =   4
         list[4]_14      =   15
         ClassName_15    =   "CCWDrawObj"
         opts_15         =   62
         C[0]_15         =   -2147483640
         C[1]_15         =   -2147483640
         Image_15        =   6
         Animator_15     =   0
         Blinker_15      =   0
         list[3]_14      =   16
         ClassName_16    =   "CCWDrawObj"
         opts_16         =   62
         C[0]_16         =   -2147483640
         C[1]_16         =   -2147483640
         Image_16        =   8
         Animator_16     =   0
         Blinker_16      =   0
         list[2]_14      =   17
         ClassName_17    =   "CCWDrawObj"
         opts_17         =   60
         C[0]_17         =   -2147483640
         C[1]_17         =   -2147483640
         Image_17        =   10
         Animator_17     =   0
         Blinker_17      =   0
         list[1]_14      =   18
         ClassName_18    =   "CCWDrawObj"
         opts_18         =   62
         C[0]_18         =   255
         C[1]_18         =   255
         Image_18        =   19
         ClassName_19    =   "CCWPictImage"
         opts_19         =   1280
         Rows_19         =   1
         Cols_19         =   1
         Pict_19         =   74
         F_19            =   255
         B_19            =   255
         ColorReplaceWith_19=   8421504
         ColorReplace_19 =   8421504
         Tolerance_19    =   2
         Animator_18     =   0
         Blinker_18      =   0
         varVarType_13   =   5
         Bindings_1      =   20
         ClassName_20    =   "CCWBindingHolderArray"
         Editor_20       =   21
         ClassName_21    =   "CCWBindingHolderArrayEditor"
         Owner_21        =   1
         Style_1         =   4
         mechAction_1    =   3
         BGImg_1         =   22
         ClassName_22    =   "CCWDrawObj"
         opts_22         =   62
         Image_22        =   23
         ClassName_23    =   "CCWPictImage"
         opts_23         =   1280
         Rows_23         =   1
         Cols_23         =   1
         Pict_23         =   286
         F_23            =   -2147483633
         B_23            =   -2147483633
         ColorReplaceWith_23=   8421504
         ColorReplace_23 =   8421504
         Tolerance_23    =   2
         Animator_22     =   0
         Blinker_22      =   0
         Array_1         =   6
         Editor_1        =   0
         Array[0]_1      =   11
         Array[1]_1      =   18
         Array[2]_1      =   0
         Array[3]_1      =   0
         Array[4]_1      =   24
         ClassName_24    =   "CCWDrawObj"
         opts_24         =   62
         Image_24        =   8
         Animator_24     =   0
         Blinker_24      =   0
         Array[5]_1      =   25
         ClassName_25    =   "CCWDrawObj"
         opts_25         =   62
         Image_25        =   10
         Animator_25     =   0
         Blinker_25      =   0
         Label_1         =   26
         ClassName_26    =   "CCWDrawObj"
         opts_26         =   62
         C[0]_26         =   -2147483640
         Image_26        =   6
         Animator_26     =   0
         Blinker_26      =   0
      End
      Begin CWUIControlsLib.CWButton cw_Status_Bit_Indicator
         Height          =   645
         Index           =   6
         Left            =   2700
         TabIndex        =   147
         ToolTipText     =   "Shows status of ALL Velocity Feed Back loops"
         Top             =   45
         Width           =   975
         _Version        =   524288
         _ExtentX        =   1720
         _ExtentY        =   1138
         _StockProps     =   68
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Reset_0         =   0   'False
         CompatibleVers_0=   524288
         Boolean_0       =   1
         ClassName_1     =   "CCWBoolean"
         opts_1          =   2622
         C[0]_1          =   16761024
         Enum_1          =   2
         ClassName_2     =   "CCWEnum"
         Array_2         =   2
         Editor_2        =   0
         Array[0]_2      =   3
         ClassName_3     =   "CCWEnumElt"
         opts_3          =   1
         Name_3          =   "Off"
         frame_3         =   286
         DrawList_3      =   4
         ClassName_4     =   "CDrawList"
         count_4         =   4
         list[4]_4       =   5
         ClassName_5     =   "CCWDrawObj"
         opts_5          =   62
         C[0]_5          =   -2147483640
         C[1]_5          =   -2147483640
         Image_5         =   6
         ClassName_6     =   "CCWTextImage"
         szText_6        =   "Doing Test"
         font_6          =   0
         Animator_5      =   0
         Blinker_5       =   0
         list[3]_4       =   7
         ClassName_7     =   "CCWDrawObj"
         opts_7          =   60
         C[0]_7          =   -2147483640
         C[1]_7          =   -2147483640
         Image_7         =   8
         ClassName_8     =   "CCWTextImage"
         szText_8        =   "YES"
         font_8          =   0
         Animator_7      =   0
         Blinker_7       =   0
         list[2]_4       =   9
         ClassName_9     =   "CCWDrawObj"
         opts_9          =   62
         C[0]_9          =   -2147483640
         C[1]_9          =   -2147483640
         Image_9         =   10
         ClassName_10    =   "CCWTextImage"
         szText_10       =   "No Test"
         font_10         =   0
         Animator_9      =   0
         Blinker_9       =   0
         list[1]_4       =   11
         ClassName_11    =   "CCWDrawObj"
         opts_11         =   62
         C[0]_11         =   65280
         C[1]_11         =   65280
         Image_11        =   12
         ClassName_12    =   "CCWPictImage"
         opts_12         =   1280
         Rows_12         =   1
         Cols_12         =   1
         Pict_12         =   75
         F_12            =   65280
         B_12            =   65280
         ColorReplaceWith_12=   8421504
         ColorReplace_12 =   8421504
         Tolerance_12    =   2
         Animator_11     =   0
         Blinker_11      =   0
         varVarType_3    =   5
         Array[1]_2      =   13
         ClassName_13    =   "CCWEnumElt"
         opts_13         =   1
         Name_13         =   "On"
         frame_13        =   286
         DrawList_13     =   14
         ClassName_14    =   "CDrawList"
         count_14        =   4
         list[4]_14      =   15
         ClassName_15    =   "CCWDrawObj"
         opts_15         =   62
         C[0]_15         =   -2147483640
         C[1]_15         =   -2147483640
         Image_15        =   6
         Animator_15     =   0
         Blinker_15      =   0
         list[3]_14      =   16
         ClassName_16    =   "CCWDrawObj"
         opts_16         =   62
         C[0]_16         =   -2147483640
         C[1]_16         =   -2147483640
         Image_16        =   8
         Animator_16     =   0
         Blinker_16      =   0
         list[2]_14      =   17
         ClassName_17    =   "CCWDrawObj"
         opts_17         =   60
         C[0]_17         =   -2147483640
         C[1]_17         =   -2147483640
         Image_17        =   10
         Animator_17     =   0
         Blinker_17      =   0
         list[1]_14      =   18
         ClassName_18    =   "CCWDrawObj"
         opts_18         =   62
         C[0]_18         =   65535
         C[1]_18         =   65535
         Image_18        =   19
         ClassName_19    =   "CCWPictImage"
         opts_19         =   1280
         Rows_19         =   1
         Cols_19         =   1
         Pict_19         =   76
         F_19            =   65535
         B_19            =   65535
         ColorReplaceWith_19=   8421504
         ColorReplace_19 =   8421504
         Tolerance_19    =   2
         Animator_18     =   0
         Blinker_18      =   20
         ClassName_20    =   "CCWBlinker"
         Interval_20     =   600
         EventID_20      =   7
         Obj_20          =   18
         varVarType_13   =   5
         Bindings_1      =   21
         ClassName_21    =   "CCWBindingHolderArray"
         Editor_21       =   22
         ClassName_22    =   "CCWBindingHolderArrayEditor"
         Owner_22        =   1
         Style_1         =   5
         mechAction_1    =   3
         BGImg_1         =   23
         ClassName_23    =   "CCWDrawObj"
         opts_23         =   62
         Image_23        =   24
         ClassName_24    =   "CCWPictImage"
         opts_24         =   1280
         Rows_24         =   1
         Cols_24         =   1
         Pict_24         =   286
         F_24            =   -2147483633
         B_24            =   -2147483633
         ColorReplaceWith_24=   8421504
         ColorReplace_24 =   8421504
         Tolerance_24    =   2
         Animator_23     =   0
         Blinker_23      =   0
         Array_1         =   6
         Editor_1        =   0
         Array[0]_1      =   11
         Array[1]_1      =   18
         Array[2]_1      =   0
         Array[3]_1      =   0
         Array[4]_1      =   25
         ClassName_25    =   "CCWDrawObj"
         opts_25         =   62
         Image_25        =   8
         Animator_25     =   0
         Blinker_25      =   0
         Array[5]_1      =   26
         ClassName_26    =   "CCWDrawObj"
         opts_26         =   62
         Image_26        =   10
         Animator_26     =   0
         Blinker_26      =   0
         Label_1         =   27
         ClassName_27    =   "CCWDrawObj"
         opts_27         =   62
         C[0]_27         =   -2147483640
         Image_27        =   6
         Animator_27     =   0
         Blinker_27      =   0
      End
   End
   Begin VB.Frame FrameRTmon
      BackColor       =   &H00404040&
      Height          =   1725
      Left            =   2745
      TabIndex        =   70
      Top             =   -62
      Width           =   11760
      Begin VB.Frame Frame_RT_Mag_Ph
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         Height          =   1095
         Left            =   7335
         TabIndex        =   132
         Top             =   540
         Visible         =   0   'False
         Width           =   4380
         Begin CWUIControlsLib.CWButton cw_Excitation_is_On
            Height          =   375
            Left            =   45
            TabIndex        =   133
            ToolTipText     =   "Excitation is on and might interfere with normal work"
            Top             =   360
            Width           =   2040
            _Version        =   524288
            _ExtentX        =   3598
            _ExtentY        =   661
            _StockProps     =   68
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}
               Name            =   "Arial"
               Size            =   9.74
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Reset_0         =   0   'False
            CompatibleVers_0=   524288
            Boolean_0       =   1
            ClassName_1     =   "CCWBoolean"
            opts_1          =   2622
            C[0]_1          =   4210752
            Enum_1          =   2
            ClassName_2     =   "CCWEnum"
            Array_2         =   2
            Editor_2        =   0
            Array[0]_2      =   3
            ClassName_3     =   "CCWEnumElt"
            opts_3          =   1
            Name_3          =   "Off"
            frame_3         =   286
            DrawList_3      =   4
            ClassName_4     =   "CDrawList"
            count_4         =   4
            list[4]_4       =   5
            ClassName_5     =   "CCWDrawObj"
            opts_5          =   62
            C[0]_5          =   -2147483640
            C[1]_5          =   -2147483640
            Image_5         =   6
            ClassName_6     =   "CCWTextImage"
            font_6          =   0
            Animator_5      =   0
            Blinker_5       =   0
            list[3]_4       =   7
            ClassName_7     =   "CCWDrawObj"
            opts_7          =   60
            C[0]_7          =   6579455
            C[1]_7          =   6579455
            Image_7         =   8
            ClassName_8     =   "CCWTextImage"
            szText_8        =   "! EXCITATION IS ON !"
            font_8          =   0
            Animator_7      =   0
            Blinker_7       =   0
            list[2]_4       =   9
            ClassName_9     =   "CCWDrawObj"
            opts_9          =   62
            C[0]_9          =   -2147483640
            C[1]_9          =   -2147483640
            Image_9         =   10
            ClassName_10    =   "CCWTextImage"
            font_10         =   0
            Animator_9      =   0
            Blinker_9       =   0
            list[1]_4       =   11
            ClassName_11    =   "CCWDrawObj"
            opts_11         =   62
            Image_11        =   12
            ClassName_12    =   "CCWPictImage"
            opts_12         =   1280
            Rows_12         =   1
            Cols_12         =   1
            Pict_12         =   51
            F_12            =   -2147483633
            B_12            =   -2147483633
            ColorReplaceWith_12=   8421504
            ColorReplace_12 =   8421504
            Tolerance_12    =   2
            Animator_11     =   0
            Blinker_11      =   0
            varVarType_3    =   5
            Array[1]_2      =   13
            ClassName_13    =   "CCWEnumElt"
            opts_13         =   1
            Name_13         =   "On"
            frame_13        =   286
            DrawList_13     =   14
            ClassName_14    =   "CDrawList"
            count_14        =   4
            list[4]_14      =   15
            ClassName_15    =   "CCWDrawObj"
            opts_15         =   60
            C[0]_15         =   -2147483640
            C[1]_15         =   -2147483640
            Image_15        =   6
            Animator_15     =   0
            Blinker_15      =   0
            list[3]_14      =   16
            ClassName_16    =   "CCWDrawObj"
            opts_16         =   62
            C[0]_16         =   6579455
            C[1]_16         =   6579455
            Image_16        =   8
            Animator_16     =   0
            Blinker_16      =   17
            ClassName_17    =   "CCWBlinker"
            Interval_17     =   600
            Enable_17       =   -1  'True
            EventID_17      =   7
            Obj_17          =   16
            list[2]_14      =   18
            ClassName_18    =   "CCWDrawObj"
            opts_18         =   60
            C[0]_18         =   -2147483640
            C[1]_18         =   -2147483640
            Image_18        =   10
            Animator_18     =   0
            Blinker_18      =   0
            list[1]_14      =   19
            ClassName_19    =   "CCWDrawObj"
            opts_19         =   60
            Image_19        =   20
            ClassName_20    =   "CCWPictImage"
            opts_20         =   1280
            Rows_20         =   1
            Cols_20         =   1
            Pict_20         =   52
            F_20            =   -2147483633
            B_20            =   -2147483633
            ColorReplaceWith_20=   8421504
            ColorReplace_20 =   8421504
            Tolerance_20    =   2
            Animator_19     =   0
            Blinker_19      =   0
            varVarType_13   =   5
            Bindings_1      =   21
            ClassName_21    =   "CCWBindingHolderArray"
            Editor_21       =   22
            ClassName_22    =   "CCWBindingHolderArrayEditor"
            Owner_22        =   1
            Style_1         =   11
            drawState_1     =   1
            mechAction_1    =   1
            BGImg_1         =   23
            ClassName_23    =   "CCWDrawObj"
            opts_23         =   62
            Image_23        =   24
            ClassName_24    =   "CCWPictImage"
            opts_24         =   1280
            Rows_24         =   1
            Cols_24         =   1
            Pict_24         =   286
            F_24            =   -2147483633
            B_24            =   -2147483633
            ColorReplaceWith_24=   8421504
            ColorReplace_24 =   8421504
            Tolerance_24    =   2
            Animator_23     =   0
            Blinker_23      =   0
            Array_1         =   6
            Editor_1        =   0
            Array[0]_1      =   11
            Array[1]_1      =   19
            Array[2]_1      =   0
            Array[3]_1      =   0
            Array[4]_1      =   25
            ClassName_25    =   "CCWDrawObj"
            opts_25         =   62
            Image_25        =   8
            Animator_25     =   0
            Blinker_25      =   0
            Array[5]_1      =   26
            ClassName_26    =   "CCWDrawObj"
            opts_26         =   62
            Image_26        =   10
            Animator_26     =   0
            Blinker_26      =   0
            Label_1         =   27
            ClassName_27    =   "CCWDrawObj"
            opts_27         =   62
            C[0]_27         =   -2147483640
            Image_27        =   6
            Animator_27     =   0
            Blinker_27      =   0
         End
         Begin VB.Label lbl_BNC_0
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00E0FFFF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Zpos_mm"
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
            Left            =   90
            TabIndex        =   140
            ToolTipText     =   "Valve #1 current, mA"
            Top             =   45
            Width           =   1860
         End
         Begin VB.Label lbl_BNC_1
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00E0FFFF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Excitation"
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
            Left            =   2340
            TabIndex        =   139
            ToolTipText     =   "Valve #1 current, mA"
            Top             =   45
            Width           =   1860
         End
         Begin VB.Label lbl_Mag_dB
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
            Left            =   3330
            TabIndex        =   138
            ToolTipText     =   "Valve #1 current, mA"
            Top             =   405
            Width           =   870
         End
         Begin VB.Label lbl_Phase_diff
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
            Left            =   3330
            TabIndex        =   137
            ToolTipText     =   "Valve #1 current, mA"
            Top             =   765
            Width           =   870
         End
         Begin VB.Label lbValveNum
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "vs."
            BeginProperty Font
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   270
            Index           =   5
            Left            =   1845
            TabIndex        =   136
            Top             =   90
            Width           =   645
         End
         Begin VB.Label lbValveNum
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Relative dB"
            BeginProperty Font
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   270
            Index           =   6
            Left            =   1710
            TabIndex        =   135
            Top             =   405
            Width           =   1590
         End
         Begin VB.Label lbValveNum
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Relative degrees"
            BeginProperty Font
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   270
            Index           =   7
            Left            =   1260
            TabIndex        =   134
            Top             =   765
            Width           =   2040
         End
      End
      Begin VB.TextBox txtManualCMD
         BeginProperty Font
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   7425
         TabIndex        =   131
         Text            =   "command"
         Top             =   180
         Width           =   1860
      End
      Begin VB.Frame FrameBNCs
         BackColor       =   &H00808080&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   480
         Left            =   45
         TabIndex        =   125
         Top             =   1180
         Width           =   7225
         Begin VB.ComboBox CmbBNC0
            BeginProperty Font
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            ItemData        =   "Analyzer.frx":C669E
            Left            =   45
            List            =   "Analyzer.frx":C66A0
            Style           =   2  'Dropdown List
            TabIndex        =   127
            Tag             =   "bncd0="
            ToolTipText     =   "Choose signal on BNC Diag.0"
            Top             =   45
            Width           =   1890
         End
         Begin VB.ComboBox CmbBNC1
            BeginProperty Font
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            ItemData        =   "Analyzer.frx":C66A2
            Left            =   5310
            List            =   "Analyzer.frx":C66A4
            Style           =   2  'Dropdown List
            TabIndex        =   126
            Tag             =   "bncd1="
            ToolTipText     =   "Choose signal on BNC Diag.1"
            Top             =   45
            Width           =   1890
         End
         Begin CWUIControlsLib.CWButton cwBNCsliders
            Height          =   450
            Left            =   2610
            TabIndex        =   128
            ToolTipText     =   "Show / hide BNC sliders instead of drop-dowm lists"
            Top             =   20
            Width           =   1995
            _Version        =   524288
            _ExtentX        =   3519
            _ExtentY        =   794
            _StockProps     =   68
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Reset_0         =   0   'False
            CompatibleVers_0=   524288
            Boolean_0       =   1
            ClassName_1     =   "CCWBoolean"
            opts_1          =   2614
            SclRef_1.l      =   2
            SclRef_1.t      =   2
            SclRef_1.r      =   135
            SclRef_1.b      =   32
            Scl_1.l         =   41
            Scl_1.t         =   7
            Scl_1.r         =   98
            Scl_1.b         =   28
            C[0]_1          =   16777215
            Enum_1          =   2
            ClassName_2     =   "CCWEnum"
            Array_2         =   2
            Editor_2        =   0
            Array[0]_2      =   3
            ClassName_3     =   "CCWEnumElt"
            opts_3          =   1
            Name_3          =   "Off"
            frame_3         =   288
            DrawList_3      =   4
            ClassName_4     =   "CDrawList"
            count_4         =   4
            list[4]_4       =   5
            ClassName_5     =   "CCWDrawObj"
            opts_5          =   60
            C[0]_5          =   0
            C[1]_5          =   0
            Image_5         =   6
            ClassName_6     =   "CCWTextImage"
            font_6          =   0
            Animator_5      =   0
            Blinker_5       =   0
            list[3]_4       =   7
            ClassName_7     =   "CCWDrawObj"
            opts_7          =   52
            SclRef_7.l      =   45
            SclRef_7.t      =   11
            SclRef_7.r      =   95
            SclRef_7.b      =   25
            Scl_7.l         =   44
            Scl_7.t         =   11
            Scl_7.r         =   94
            Scl_7.b         =   24
            C[0]_7          =   -2147483640
            C[1]_7          =   -2147483640
            Image_7         =   8
            ClassName_8     =   "CCWTextImage"
            szText_8        =   "drop down"
            style_8         =   48598384
            font_8          =   9
            ClassName_9     =   "CCWFont"
            bFont_9         =   -1  'True
            BeginProperty Font_9 {0BE35203-8F91-11CE-9DE3-00AA004BB851}
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Animator_7      =   0
            Blinker_7       =   0
            list[2]_4       =   10
            ClassName_10    =   "CCWDrawObj"
            opts_10         =   62
            C[0]_10         =   -2147483640
            C[1]_10         =   -2147483640
            Image_10        =   11
            ClassName_11    =   "CCWTextImage"
            szText_11       =   "SLIDERS"
            font_11         =   0
            Animator_10     =   0
            Blinker_10      =   0
            list[1]_4       =   12
            ClassName_12    =   "CCWDrawObj"
            opts_12         =   62
            C[0]_12         =   12648384
            C[1]_12         =   12648384
            Image_12        =   13
            ClassName_13    =   "CCWPictImage"
            opts_13         =   1280
            Rows_13         =   1
            Cols_13         =   1
            Pict_13         =   2
            F_13            =   12648384
            B_13            =   12648384
            ColorReplaceWith_13=   8421504
            ColorReplace_13 =   8421504
            Tolerance_13    =   2
            Animator_12     =   0
            Blinker_12      =   0
            varVarType_3    =   5
            Array[1]_2      =   14
            ClassName_14    =   "CCWEnumElt"
            opts_14         =   1
            Name_14         =   "On"
            frame_14        =   287
            DrawList_14     =   15
            ClassName_15    =   "CDrawList"
            count_15        =   4
            list[4]_15      =   16
            ClassName_16    =   "CCWDrawObj"
            opts_16         =   62
            C[0]_16         =   0
            C[1]_16         =   0
            Image_16        =   6
            Animator_16     =   0
            Blinker_16      =   0
            list[3]_15      =   17
            ClassName_17    =   "CCWDrawObj"
            opts_17         =   54
            SclRef_17.l     =   45
            SclRef_17.t     =   11
            SclRef_17.r     =   95
            SclRef_17.b     =   25
            Scl_17.l        =   44
            Scl_17.t        =   11
            Scl_17.r        =   94
            Scl_17.b        =   24
            C[0]_17         =   -2147483640
            C[1]_17         =   -2147483640
            Image_17        =   8
            Animator_17     =   0
            Blinker_17      =   0
            list[2]_15      =   18
            ClassName_18    =   "CCWDrawObj"
            opts_18         =   60
            C[0]_18         =   -2147483640
            C[1]_18         =   -2147483640
            Image_18        =   11
            Animator_18     =   0
            Blinker_18      =   0
            list[1]_15      =   19
            ClassName_19    =   "CCWDrawObj"
            opts_19         =   62
            C[0]_19         =   16777152
            C[1]_19         =   16777152
            Image_19        =   20
            ClassName_20    =   "CCWPictImage"
            opts_20         =   1280
            Rows_20         =   1
            Cols_20         =   1
            Pict_20         =   2
            F_20            =   16777152
            B_20            =   16777152
            ColorReplaceWith_20=   8421504
            ColorReplace_20 =   8421504
            Tolerance_20    =   2
            Animator_19     =   0
            Blinker_19      =   0
            varVarType_14   =   5
            Bindings_1      =   21
            ClassName_21    =   "CCWBindingHolderArray"
            Editor_21       =   22
            ClassName_22    =   "CCWBindingHolderArrayEditor"
            Owner_22        =   1
            Style_1         =   12
            frameStyle_1    =   1
            mechAction_1    =   1
            BGImg_1         =   23
            ClassName_23    =   "CCWDrawObj"
            opts_23         =   62
            Image_23        =   24
            ClassName_24    =   "CCWPictImage"
            opts_24         =   256
            Rows_24         =   1
            Cols_24         =   1
            URL_24          =   "C:\in_out\Software\DC-2020_ED_Analyzer_VB6\Resources\BNC0_BNC1_400x90.bmp"
            Pict_24         =   "Analyzer.frx":C66A6
            ColorReplaceWith_24=   8421504
            ColorReplace_24 =   8421504
            Tolerance_24    =   2
            Animator_23     =   0
            Blinker_23      =   0
            Array_1         =   6
            Editor_1        =   0
            Array[0]_1      =   12
            Array[1]_1      =   19
            Array[2]_1      =   0
            Array[3]_1      =   0
            Array[4]_1      =   25
            ClassName_25    =   "CCWDrawObj"
            opts_25         =   62
            Image_25        =   8
            Animator_25     =   0
            Blinker_25      =   0
            Array[5]_1      =   26
            ClassName_26    =   "CCWDrawObj"
            opts_26         =   62
            Image_26        =   11
            Animator_26     =   0
            Blinker_26      =   0
            Label_1         =   27
            ClassName_27    =   "CCWDrawObj"
            opts_27         =   62
            C[0]_27         =   -2147483640
            Image_27        =   6
            Animator_27     =   0
            Blinker_27      =   0
         End
         Begin VB.Label lbl
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Diag.0"
            BeginProperty Font
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   0
            Left            =   1965
            TabIndex        =   130
            Top             =   90
            Width           =   600
         End
         Begin VB.Label lbl
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Diag.1"
            BeginProperty Font
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   1
            Left            =   4635
            TabIndex        =   129
            Top             =   90
            Width           =   645
         End
      End
      Begin VB.Timer aTimerMonitoring
         Interval        =   1000
         Left            =   3150
         Top             =   90
      End
      Begin CWUIControlsLib.CWButton cwBut_RT_Monitoring
         Height          =   855
         Left            =   90
         TabIndex        =   71
         ToolTipText     =   "Enables checking position and status of the payload once a second "
         Top             =   300
         Width           =   855
         _Version        =   524288
         _ExtentX        =   1508
         _ExtentY        =   1508
         _StockProps     =   68
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}
            Name            =   "Arial"
            Size            =   9
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Reset_0         =   0   'False
         CompatibleVers_0=   524288
         Boolean_0       =   1
         ClassName_1     =   "CCWBoolean"
         opts_1          =   2622
         C[0]_1          =   8421504
         Enum_1          =   2
         ClassName_2     =   "CCWEnum"
         Array_2         =   2
         Editor_2        =   0
         Array[0]_2      =   3
         ClassName_3     =   "CCWEnumElt"
         opts_3          =   1
         Name_3          =   "Off"
         frame_3         =   286
         DrawList_3      =   4
         ClassName_4     =   "CDrawList"
         count_4         =   4
         list[4]_4       =   5
         ClassName_5     =   "CCWDrawObj"
         opts_5          =   62
         C[0]_5          =   16777215
         C[1]_5          =   16777215
         Image_5         =   6
         ClassName_6     =   "CCWTextImage"
         font_6          =   0
         Animator_5      =   0
         Blinker_5       =   0
         list[3]_4       =   7
         ClassName_7     =   "CCWDrawObj"
         opts_7          =   52
         SclRef_7.l      =   2
         SclRef_7.t      =   13
         SclRef_7.r      =   56
         SclRef_7.b      =   70
         Scl_7.l         =   6
         Scl_7.t         =   1
         Scl_7.r         =   51
         Scl_7.b         =   16
         C[0]_7          =   65535
         C[1]_7          =   -2147483640
         Image_7         =   8
         ClassName_8     =   "CCWTextImage"
         font_8          =   0
         Animator_7      =   0
         Blinker_7       =   0
         list[2]_4       =   9
         ClassName_9     =   "CCWDrawObj"
         opts_9          =   54
         SclRef_9.l      =   2
         SclRef_9.t      =   13
         SclRef_9.r      =   56
         SclRef_9.b      =   70
         Scl_9.l         =   4
         Scl_9.r         =   53
         Scl_9.b         =   15
         C[0]_9          =   65535
         C[1]_9          =   -2147483640
         Image_9         =   10
         ClassName_10    =   "CCWTextImage"
         style_10        =   17301745
         font_10         =   0
         Animator_9      =   0
         Blinker_9       =   0
         list[1]_4       =   11
         ClassName_11    =   "CCWDrawObj"
         opts_11         =   62
         C[0]_11         =   65535
         C[1]_11         =   65535
         Image_11        =   12
         ClassName_12    =   "CCWPictImage"
         opts_12         =   256
         Rows_12         =   1
         Cols_12         =   1
         URL_12          =   "C:\in_out\Software\DC-2020_ED_Analyzer_VB6\Resources\SwitchOffGreen.bmp"
         Pict_12         =   "Analyzer.frx":CF798
         ColorReplaceWith_12=   8421504
         ColorReplace_12 =   8421504
         Tolerance_12    =   2
         Animator_11     =   0
         Blinker_11      =   0
         varVarType_3    =   5
         Array[1]_2      =   13
         ClassName_13    =   "CCWEnumElt"
         opts_13         =   1
         Name_13         =   "On"
         frame_13        =   286
         DrawList_13     =   14
         ClassName_14    =   "CDrawList"
         count_14        =   4
         list[4]_14      =   15
         ClassName_15    =   "CCWDrawObj"
         opts_15         =   62
         C[0]_15         =   16777215
         C[1]_15         =   16777215
         Image_15        =   6
         Animator_15     =   0
         Blinker_15      =   0
         list[3]_14      =   16
         ClassName_16    =   "CCWDrawObj"
         opts_16         =   54
         SclRef_16.l     =   2
         SclRef_16.t     =   13
         SclRef_16.r     =   56
         SclRef_16.b     =   70
         Scl_16.l        =   6
         Scl_16.t        =   1
         Scl_16.r        =   51
         Scl_16.b        =   16
         C[0]_16         =   65280
         C[1]_16         =   -2147483640
         Image_16        =   8
         Animator_16     =   0
         Blinker_16      =   0
         list[2]_14      =   17
         ClassName_17    =   "CCWDrawObj"
         opts_17         =   52
         SclRef_17.l     =   2
         SclRef_17.t     =   13
         SclRef_17.r     =   56
         SclRef_17.b     =   70
         Scl_17.l        =   4
         Scl_17.r        =   53
         Scl_17.b        =   15
         C[0]_17         =   65280
         C[1]_17         =   -2147483640
         Image_17        =   10
         Animator_17     =   0
         Blinker_17      =   0
         list[1]_14      =   18
         ClassName_18    =   "CCWDrawObj"
         opts_18         =   62
         C[0]_18         =   65280
         C[1]_18         =   65280
         Image_18        =   19
         ClassName_19    =   "CCWPictImage"
         opts_19         =   256
         Rows_19         =   1
         Cols_19         =   1
         URL_19          =   "C:\in_out\Software\DC-2020_ED_Analyzer_VB6\Resources\SwitchOnYellow.bmp"
         Pict_19         =   "Analyzer.frx":E8ECA
         ColorReplaceWith_19=   8421504
         ColorReplace_19 =   8421504
         Tolerance_19    =   2
         Animator_18     =   0
         Blinker_18      =   0
         varVarType_13   =   5
         Bindings_1      =   20
         ClassName_20    =   "CCWBindingHolderArray"
         Editor_20       =   21
         ClassName_21    =   "CCWBindingHolderArrayEditor"
         Owner_21        =   1
         mechAction_1    =   1
         BGImg_1         =   22
         ClassName_22    =   "CCWDrawObj"
         opts_22         =   62
         Image_22        =   23
         ClassName_23    =   "CCWPictImage"
         opts_23         =   1280
         Rows_23         =   1
         Cols_23         =   1
         Pict_23         =   286
         F_23            =   -2147483633
         B_23            =   -2147483633
         ColorReplaceWith_23=   8421504
         ColorReplace_23 =   8421504
         Tolerance_23    =   2
         Animator_22     =   0
         Blinker_22      =   0
         Array_1         =   6
         Editor_1        =   0
         Array[0]_1      =   11
         Array[1]_1      =   18
         Array[2]_1      =   0
         Array[3]_1      =   0
         Array[4]_1      =   24
         ClassName_24    =   "CCWDrawObj"
         opts_24         =   62
         Image_24        =   8
         Animator_24     =   0
         Blinker_24      =   0
         Array[5]_1      =   25
         ClassName_25    =   "CCWDrawObj"
         opts_25         =   62
         Image_25        =   10
         Animator_25     =   0
         Blinker_25      =   0
         Label_1         =   26
         ClassName_26    =   "CCWDrawObj"
         opts_26         =   62
         C[0]_26         =   -2147483640
         Image_26        =   6
         Animator_26     =   0
         Blinker_26      =   0
      End
      Begin CWUIControlsLib.CWButton cw_Status_Bit_Indicator
         Height          =   285
         Index           =   3
         Left            =   3870
         TabIndex        =   106
         ToolTipText     =   "Shows status of ALL Velocity Feed Back loops"
         Top             =   165
         Width           =   1680
         _Version        =   524288
         _ExtentX        =   2963
         _ExtentY        =   503
         _StockProps     =   68
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Reset_0         =   0   'False
         CompatibleVers_0=   524288
         Boolean_0       =   1
         ClassName_1     =   "CCWBoolean"
         opts_1          =   2622
         C[0]_1          =   -2147483640
         Enum_1          =   2
         ClassName_2     =   "CCWEnum"
         Array_2         =   2
         Editor_2        =   0
         Array[0]_2      =   3
         ClassName_3     =   "CCWEnumElt"
         opts_3          =   1
         Name_3          =   "Off"
         frame_3         =   288
         DrawList_3      =   4
         ClassName_4     =   "CDrawList"
         count_4         =   4
         list[4]_4       =   5
         ClassName_5     =   "CCWDrawObj"
         opts_5          =   62
         C[0]_5          =   -2147483640
         C[1]_5          =   -2147483640
         Image_5         =   6
         ClassName_6     =   "CCWTextImage"
         font_6          =   0
         Animator_5      =   0
         Blinker_5       =   0
         list[3]_4       =   7
         ClassName_7     =   "CCWDrawObj"
         opts_7          =   60
         C[0]_7          =   -2147483640
         C[1]_7          =   -2147483640
         Image_7         =   8
         ClassName_8     =   "CCWTextImage"
         szText_8        =   "Velocity FB off"
         font_8          =   0
         Animator_7      =   0
         Blinker_7       =   0
         list[2]_4       =   9
         ClassName_9     =   "CCWDrawObj"
         opts_9          =   62
         C[0]_9          =   -2147483640
         C[1]_9          =   -2147483640
         Image_9         =   10
         ClassName_10    =   "CCWTextImage"
         szText_10       =   "Velocity FB on"
         font_10         =   0
         Animator_9      =   0
         Blinker_9       =   0
         list[1]_4       =   11
         ClassName_11    =   "CCWDrawObj"
         opts_11         =   62
         C[0]_11         =   65280
         C[1]_11         =   65280
         Image_11        =   12
         ClassName_12    =   "CCWPictImage"
         opts_12         =   1280
         Rows_12         =   1
         Cols_12         =   1
         Pict_12         =   2
         F_12            =   65280
         B_12            =   65280
         ColorReplaceWith_12=   8421504
         ColorReplace_12 =   8421504
         Tolerance_12    =   2
         Animator_11     =   0
         Blinker_11      =   0
         varVarType_3    =   5
         Array[1]_2      =   13
         ClassName_13    =   "CCWEnumElt"
         opts_13         =   1
         Name_13         =   "On"
         frame_13        =   287
         DrawList_13     =   14
         ClassName_14    =   "CDrawList"
         count_14        =   4
         list[4]_14      =   15
         ClassName_15    =   "CCWDrawObj"
         opts_15         =   62
         C[0]_15         =   -2147483640
         C[1]_15         =   -2147483640
         Image_15        =   6
         Animator_15     =   0
         Blinker_15      =   0
         list[3]_14      =   16
         ClassName_16    =   "CCWDrawObj"
         opts_16         =   62
         C[0]_16         =   -2147483640
         C[1]_16         =   -2147483640
         Image_16        =   8
         Animator_16     =   0
         Blinker_16      =   0
         list[2]_14      =   17
         ClassName_17    =   "CCWDrawObj"
         opts_17         =   60
         C[0]_17         =   -2147483640
         C[1]_17         =   -2147483640
         Image_17        =   10
         Animator_17     =   0
         Blinker_17      =   0
         list[1]_14      =   18
         ClassName_18    =   "CCWDrawObj"
         opts_18         =   62
         C[0]_18         =   8438015
         C[1]_18         =   8438015
         Image_18        =   19
         ClassName_19    =   "CCWPictImage"
         opts_19         =   1280
         Rows_19         =   1
         Cols_19         =   1
         Pict_19         =   2
         F_19            =   8438015
         B_19            =   8438015
         ColorReplaceWith_19=   8421504
         ColorReplace_19 =   8421504
         Tolerance_19    =   2
         Animator_18     =   0
         Blinker_18      =   0
         varVarType_13   =   5
         Bindings_1      =   20
         ClassName_20    =   "CCWBindingHolderArray"
         Editor_20       =   21
         ClassName_21    =   "CCWBindingHolderArrayEditor"
         Owner_21        =   1
         Style_1         =   12
         drawState_1     =   1
         frameStyle_1    =   1
         mechAction_1    =   3
         BGImg_1         =   22
         ClassName_22    =   "CCWDrawObj"
         opts_22         =   62
         Image_22        =   23
         ClassName_23    =   "CCWPictImage"
         opts_23         =   1280
         Rows_23         =   1
         Cols_23         =   1
         Pict_23         =   286
         F_23            =   -2147483633
         B_23            =   -2147483633
         ColorReplaceWith_23=   8421504
         ColorReplace_23 =   8421504
         Tolerance_23    =   2
         Animator_22     =   0
         Blinker_22      =   0
         Array_1         =   6
         Editor_1        =   0
         Array[0]_1      =   11
         Array[1]_1      =   18
         Array[2]_1      =   0
         Array[3]_1      =   0
         Array[4]_1      =   24
         ClassName_24    =   "CCWDrawObj"
         opts_24         =   62
         Image_24        =   8
         Animator_24     =   0
         Blinker_24      =   0
         Array[5]_1      =   25
         ClassName_25    =   "CCWDrawObj"
         opts_25         =   62
         Image_25        =   10
         Animator_25     =   0
         Blinker_25      =   0
         Label_1         =   26
         ClassName_26    =   "CCWDrawObj"
         opts_26         =   62
         C[0]_26         =   -2147483640
         Image_26        =   6
         Animator_26     =   0
         Blinker_26      =   0
      End
      Begin CWUIControlsLib.CWButton cw_Status_Bit_Indicator
         Height          =   285
         Index           =   7
         Left            =   3870
         TabIndex        =   107
         ToolTipText     =   "Shows Status of All Feed Forward loops"
         Top             =   855
         Width           =   1100
         _Version        =   524288
         _ExtentX        =   1940
         _ExtentY        =   503
         _StockProps     =   68
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Reset_0         =   0   'False
         CompatibleVers_0=   524288
         Boolean_0       =   1
         ClassName_1     =   "CCWBoolean"
         opts_1          =   2622
         C[0]_1          =   -2147483640
         Enum_1          =   2
         ClassName_2     =   "CCWEnum"
         Array_2         =   2
         Editor_2        =   0
         Array[0]_2      =   3
         ClassName_3     =   "CCWEnumElt"
         opts_3          =   1
         Name_3          =   "Off"
         frame_3         =   288
         DrawList_3      =   4
         ClassName_4     =   "CDrawList"
         count_4         =   4
         list[4]_4       =   5
         ClassName_5     =   "CCWDrawObj"
         opts_5          =   62
         C[0]_5          =   -2147483640
         C[1]_5          =   -2147483640
         Image_5         =   6
         ClassName_6     =   "CCWTextImage"
         font_6          =   0
         Animator_5      =   0
         Blinker_5       =   0
         list[3]_4       =   7
         ClassName_7     =   "CCWDrawObj"
         opts_7          =   60
         C[0]_7          =   -2147483640
         C[1]_7          =   -2147483640
         Image_7         =   8
         ClassName_8     =   "CCWTextImage"
         szText_8        =   "FF off"
         font_8          =   0
         Animator_7      =   0
         Blinker_7       =   0
         list[2]_4       =   9
         ClassName_9     =   "CCWDrawObj"
         opts_9          =   62
         C[0]_9          =   -2147483640
         C[1]_9          =   -2147483640
         Image_9         =   10
         ClassName_10    =   "CCWTextImage"
         szText_10       =   "FF on"
         font_10         =   0
         Animator_9      =   0
         Blinker_9       =   0
         list[1]_4       =   11
         ClassName_11    =   "CCWDrawObj"
         opts_11         =   62
         C[0]_11         =   16777088
         C[1]_11         =   16777088
         Image_11        =   12
         ClassName_12    =   "CCWPictImage"
         opts_12         =   1280
         Rows_12         =   1
         Cols_12         =   1
         Pict_12         =   2
         F_12            =   16777088
         B_12            =   16777088
         ColorReplaceWith_12=   8421504
         ColorReplace_12 =   8421504
         Tolerance_12    =   2
         Animator_11     =   0
         Blinker_11      =   0
         varVarType_3    =   5
         Array[1]_2      =   13
         ClassName_13    =   "CCWEnumElt"
         opts_13         =   1
         Name_13         =   "On"
         frame_13        =   287
         DrawList_13     =   14
         ClassName_14    =   "CDrawList"
         count_14        =   4
         list[4]_14      =   15
         ClassName_15    =   "CCWDrawObj"
         opts_15         =   62
         C[0]_15         =   -2147483640
         C[1]_15         =   -2147483640
         Image_15        =   6
         Animator_15     =   0
         Blinker_15      =   0
         list[3]_14      =   16
         ClassName_16    =   "CCWDrawObj"
         opts_16         =   62
         C[0]_16         =   -2147483640
         C[1]_16         =   -2147483640
         Image_16        =   8
         Animator_16     =   0
         Blinker_16      =   0
         list[2]_14      =   17
         ClassName_17    =   "CCWDrawObj"
         opts_17         =   60
         C[0]_17         =   -2147483640
         C[1]_17         =   -2147483640
         Image_17        =   10
         Animator_17     =   0
         Blinker_17      =   0
         list[1]_14      =   18
         ClassName_18    =   "CCWDrawObj"
         opts_18         =   62
         C[0]_18         =   12632256
         C[1]_18         =   12632256
         Image_18        =   19
         ClassName_19    =   "CCWPictImage"
         opts_19         =   1280
         Rows_19         =   1
         Cols_19         =   1
         Pict_19         =   2
         F_19            =   12632256
         B_19            =   12632256
         ColorReplaceWith_19=   8421504
         ColorReplace_19 =   8421504
         Tolerance_19    =   2
         Animator_18     =   0
         Blinker_18      =   0
         varVarType_13   =   5
         Bindings_1      =   20
         ClassName_20    =   "CCWBindingHolderArray"
         Editor_20       =   21
         ClassName_21    =   "CCWBindingHolderArrayEditor"
         Owner_21        =   1
         Style_1         =   12
         frameStyle_1    =   1
         mechAction_1    =   3
         BGImg_1         =   22
         ClassName_22    =   "CCWDrawObj"
         opts_22         =   62
         Image_22        =   23
         ClassName_23    =   "CCWPictImage"
         opts_23         =   1280
         Rows_23         =   1
         Cols_23         =   1
         Pict_23         =   286
         F_23            =   -2147483633
         B_23            =   -2147483633
         ColorReplaceWith_23=   8421504
         ColorReplace_23 =   8421504
         Tolerance_23    =   2
         Animator_22     =   0
         Blinker_22      =   0
         Array_1         =   6
         Editor_1        =   0
         Array[0]_1      =   11
         Array[1]_1      =   18
         Array[2]_1      =   0
         Array[3]_1      =   0
         Array[4]_1      =   24
         ClassName_24    =   "CCWDrawObj"
         opts_24         =   62
         Image_24        =   8
         Animator_24     =   0
         Blinker_24      =   0
         Array[5]_1      =   25
         ClassName_25    =   "CCWDrawObj"
         opts_25         =   62
         Image_25        =   10
         Animator_25     =   0
         Blinker_25      =   0
         Label_1         =   26
         ClassName_26    =   "CCWDrawObj"
         opts_26         =   62
         C[0]_26         =   -2147483640
         Image_26        =   6
         Animator_26     =   0
         Blinker_26      =   0
      End
      Begin CWUIControlsLib.CWButton cwAxisFBstatus
         Height          =   285
         Index           =   0
         Left            =   3870
         TabIndex        =   110
         ToolTipText     =   "Shows status of a Velocity Feed Back loop"
         Top             =   495
         Width           =   510
         _Version        =   524288
         _ExtentX        =   891
         _ExtentY        =   503
         _StockProps     =   68
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Reset_0         =   0   'False
         CompatibleVers_0=   524288
         Boolean_0       =   1
         ClassName_1     =   "CCWBoolean"
         opts_1          =   2622
         C[0]_1          =   -2147483640
         Enum_1          =   2
         ClassName_2     =   "CCWEnum"
         Array_2         =   2
         Editor_2        =   0
         Array[0]_2      =   3
         ClassName_3     =   "CCWEnumElt"
         opts_3          =   1
         Name_3          =   "Off"
         frame_3         =   288
         DrawList_3      =   4
         ClassName_4     =   "CDrawList"
         count_4         =   4
         list[4]_4       =   5
         ClassName_5     =   "CCWDrawObj"
         opts_5          =   62
         C[0]_5          =   -2147483640
         C[1]_5          =   -2147483640
         Image_5         =   6
         ClassName_6     =   "CCWTextImage"
         font_6          =   0
         Animator_5      =   0
         Blinker_5       =   0
         list[3]_4       =   7
         ClassName_7     =   "CCWDrawObj"
         opts_7          =   60
         C[0]_7          =   -2147483640
         C[1]_7          =   -2147483640
         Image_7         =   8
         ClassName_8     =   "CCWTextImage"
         szText_8        =   "X on"
         font_8          =   0
         Animator_7      =   0
         Blinker_7       =   0
         list[2]_4       =   9
         ClassName_9     =   "CCWDrawObj"
         opts_9          =   54
         SclRef_9.l      =   5
         SclRef_9.t      =   5
         SclRef_9.r      =   31
         SclRef_9.b      =   17
         Scl_9.l         =   7
         Scl_9.t         =   5
         Scl_9.r         =   29
         Scl_9.b         =   19
         C[0]_9          =   -2147483640
         C[1]_9          =   -2147483640
         Image_9         =   10
         ClassName_10    =   "CCWTextImage"
         szText_10       =   "Xoff"
         font_10         =   0
         Animator_9      =   0
         Blinker_9       =   0
         list[1]_4       =   11
         ClassName_11    =   "CCWDrawObj"
         opts_11         =   62
         C[0]_11         =   8438015
         C[1]_11         =   8438015
         Image_11        =   12
         ClassName_12    =   "CCWPictImage"
         opts_12         =   1280
         Rows_12         =   1
         Cols_12         =   1
         Pict_12         =   2
         F_12            =   8438015
         B_12            =   8438015
         ColorReplaceWith_12=   8421504
         ColorReplace_12 =   8421504
         Tolerance_12    =   2
         Animator_11     =   0
         Blinker_11      =   0
         varVarType_3    =   5
         Array[1]_2      =   13
         ClassName_13    =   "CCWEnumElt"
         opts_13         =   1
         Name_13         =   "On"
         frame_13        =   287
         DrawList_13     =   14
         ClassName_14    =   "CDrawList"
         count_14        =   4
         list[4]_14      =   15
         ClassName_15    =   "CCWDrawObj"
         opts_15         =   62
         C[0]_15         =   -2147483640
         C[1]_15         =   -2147483640
         Image_15        =   6
         Animator_15     =   0
         Blinker_15      =   0
         list[3]_14      =   16
         ClassName_16    =   "CCWDrawObj"
         opts_16         =   62
         C[0]_16         =   -2147483640
         C[1]_16         =   -2147483640
         Image_16        =   8
         Animator_16     =   0
         Blinker_16      =   0
         list[2]_14      =   17
         ClassName_17    =   "CCWDrawObj"
         opts_17         =   52
         SclRef_17.l     =   5
         SclRef_17.t     =   5
         SclRef_17.r     =   31
         SclRef_17.b     =   17
         Scl_17.l        =   7
         Scl_17.t        =   5
         Scl_17.r        =   29
         Scl_17.b        =   19
         C[0]_17         =   -2147483640
         C[1]_17         =   -2147483640
         Image_17        =   10
         Animator_17     =   0
         Blinker_17      =   0
         list[1]_14      =   18
         ClassName_18    =   "CCWDrawObj"
         opts_18         =   62
         C[0]_18         =   65280
         C[1]_18         =   65280
         Image_18        =   19
         ClassName_19    =   "CCWPictImage"
         opts_19         =   1280
         Rows_19         =   1
         Cols_19         =   1
         Pict_19         =   2
         F_19            =   65280
         B_19            =   65280
         ColorReplaceWith_19=   8421504
         ColorReplace_19 =   8421504
         Tolerance_19    =   2
         Animator_18     =   0
         Blinker_18      =   0
         varVarType_13   =   5
         Bindings_1      =   20
         ClassName_20    =   "CCWBindingHolderArray"
         Editor_20       =   21
         ClassName_21    =   "CCWBindingHolderArrayEditor"
         Owner_21        =   1
         Style_1         =   12
         frameStyle_1    =   1
         mechAction_1    =   3
         BGImg_1         =   22
         ClassName_22    =   "CCWDrawObj"
         opts_22         =   62
         Image_22        =   23
         ClassName_23    =   "CCWPictImage"
         opts_23         =   1280
         Rows_23         =   1
         Cols_23         =   1
         Pict_23         =   286
         F_23            =   -2147483633
         B_23            =   -2147483633
         ColorReplaceWith_23=   8421504
         ColorReplace_23 =   8421504
         Tolerance_23    =   2
         Animator_22     =   0
         Blinker_22      =   0
         Array_1         =   6
         Editor_1        =   0
         Array[0]_1      =   11
         Array[1]_1      =   18
         Array[2]_1      =   0
         Array[3]_1      =   0
         Array[4]_1      =   24
         ClassName_24    =   "CCWDrawObj"
         opts_24         =   62
         Image_24        =   8
         Animator_24     =   0
         Blinker_24      =   0
         Array[5]_1      =   25
         ClassName_25    =   "CCWDrawObj"
         opts_25         =   62
         Image_25        =   10
         Animator_25     =   0
         Blinker_25      =   0
         Label_1         =   26
         ClassName_26    =   "CCWDrawObj"
         opts_26         =   62
         C[0]_26         =   -2147483640
         Image_26        =   6
         Animator_26     =   0
         Blinker_26      =   0
      End
      Begin CWUIControlsLib.CWButton cwAxisFBstatus
         Height          =   285
         Index           =   1
         Left            =   4455
         TabIndex        =   111
         ToolTipText     =   "Shows status of a Velocity Feed Back loop"
         Top             =   495
         Width           =   510
         _Version        =   524288
         _ExtentX        =   891
         _ExtentY        =   503
         _StockProps     =   68
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Reset_0         =   0   'False
         CompatibleVers_0=   524288
         Boolean_0       =   1
         ClassName_1     =   "CCWBoolean"
         opts_1          =   2622
         C[0]_1          =   -2147483640
         Enum_1          =   2
         ClassName_2     =   "CCWEnum"
         Array_2         =   2
         Editor_2        =   0
         Array[0]_2      =   3
         ClassName_3     =   "CCWEnumElt"
         opts_3          =   1
         Name_3          =   "Off"
         frame_3         =   288
         DrawList_3      =   4
         ClassName_4     =   "CDrawList"
         count_4         =   4
         list[4]_4       =   5
         ClassName_5     =   "CCWDrawObj"
         opts_5          =   62
         C[0]_5          =   -2147483640
         C[1]_5          =   -2147483640
         Image_5         =   6
         ClassName_6     =   "CCWTextImage"
         font_6          =   0
         Animator_5      =   0
         Blinker_5       =   0
         list[3]_4       =   7
         ClassName_7     =   "CCWDrawObj"
         opts_7          =   60
         C[0]_7          =   -2147483640
         C[1]_7          =   -2147483640
         Image_7         =   8
         ClassName_8     =   "CCWTextImage"
         szText_8        =   "Y on"
         font_8          =   0
         Animator_7      =   0
         Blinker_7       =   0
         list[2]_4       =   9
         ClassName_9     =   "CCWDrawObj"
         opts_9          =   54
         SclRef_9.l      =   5
         SclRef_9.t      =   5
         SclRef_9.r      =   31
         SclRef_9.b      =   17
         Scl_9.l         =   7
         Scl_9.t         =   5
         Scl_9.r         =   29
         Scl_9.b         =   19
         C[0]_9          =   -2147483640
         C[1]_9          =   -2147483640
         Image_9         =   10
         ClassName_10    =   "CCWTextImage"
         szText_10       =   "Yoff"
         font_10         =   0
         Animator_9      =   0
         Blinker_9       =   0
         list[1]_4       =   11
         ClassName_11    =   "CCWDrawObj"
         opts_11         =   62
         C[0]_11         =   8438015
         C[1]_11         =   8438015
         Image_11        =   12
         ClassName_12    =   "CCWPictImage"
         opts_12         =   1280
         Rows_12         =   1
         Cols_12         =   1
         Pict_12         =   2
         F_12            =   8438015
         B_12            =   8438015
         ColorReplaceWith_12=   8421504
         ColorReplace_12 =   8421504
         Tolerance_12    =   2
         Animator_11     =   0
         Blinker_11      =   0
         varVarType_3    =   5
         Array[1]_2      =   13
         ClassName_13    =   "CCWEnumElt"
         opts_13         =   1
         Name_13         =   "On"
         frame_13        =   287
         DrawList_13     =   14
         ClassName_14    =   "CDrawList"
         count_14        =   4
         list[4]_14      =   15
         ClassName_15    =   "CCWDrawObj"
         opts_15         =   62
         C[0]_15         =   -2147483640
         C[1]_15         =   -2147483640
         Image_15        =   6
         Animator_15     =   0
         Blinker_15      =   0
         list[3]_14      =   16
         ClassName_16    =   "CCWDrawObj"
         opts_16         =   62
         C[0]_16         =   -2147483640
         C[1]_16         =   -2147483640
         Image_16        =   8
         Animator_16     =   0
         Blinker_16      =   0
         list[2]_14      =   17
         ClassName_17    =   "CCWDrawObj"
         opts_17         =   52
         SclRef_17.l     =   5
         SclRef_17.t     =   5
         SclRef_17.r     =   31
         SclRef_17.b     =   17
         Scl_17.l        =   7
         Scl_17.t        =   5
         Scl_17.r        =   29
         Scl_17.b        =   19
         C[0]_17         =   -2147483640
         C[1]_17         =   -2147483640
         Image_17        =   10
         Animator_17     =   0
         Blinker_17      =   0
         list[1]_14      =   18
         ClassName_18    =   "CCWDrawObj"
         opts_18         =   62
         C[0]_18         =   65280
         C[1]_18         =   65280
         Image_18        =   19
         ClassName_19    =   "CCWPictImage"
         opts_19         =   1280
         Rows_19         =   1
         Cols_19         =   1
         Pict_19         =   2
         F_19            =   65280
         B_19            =   65280
         ColorReplaceWith_19=   8421504
         ColorReplace_19 =   8421504
         Tolerance_19    =   2
         Animator_18     =   0
         Blinker_18      =   0
         varVarType_13   =   5
         Bindings_1      =   20
         ClassName_20    =   "CCWBindingHolderArray"
         Editor_20       =   21
         ClassName_21    =   "CCWBindingHolderArrayEditor"
         Owner_21        =   1
         Style_1         =   12
         frameStyle_1    =   1
         mechAction_1    =   3
         BGImg_1         =   22
         ClassName_22    =   "CCWDrawObj"
         opts_22         =   62
         Image_22        =   23
         ClassName_23    =   "CCWPictImage"
         opts_23         =   1280
         Rows_23         =   1
         Cols_23         =   1
         Pict_23         =   286
         F_23            =   -2147483633
         B_23            =   -2147483633
         ColorReplaceWith_23=   8421504
         ColorReplace_23 =   8421504
         Tolerance_23    =   2
         Animator_22     =   0
         Blinker_22      =   0
         Array_1         =   6
         Editor_1        =   0
         Array[0]_1      =   11
         Array[1]_1      =   18
         Array[2]_1      =   0
         Array[3]_1      =   0
         Array[4]_1      =   24
         ClassName_24    =   "CCWDrawObj"
         opts_24         =   62
         Image_24        =   8
         Animator_24     =   0
         Blinker_24      =   0
         Array[5]_1      =   25
         ClassName_25    =   "CCWDrawObj"
         opts_25         =   62
         Image_25        =   10
         Animator_25     =   0
         Blinker_25      =   0
         Label_1         =   26
         ClassName_26    =   "CCWDrawObj"
         opts_26         =   62
         C[0]_26         =   -2147483640
         Image_26        =   6
         Animator_26     =   0
         Blinker_26      =   0
      End
      Begin CWUIControlsLib.CWButton cwAxisFBstatus
         Height          =   285
         Index           =   2
         Left            =   5040
         TabIndex        =   112
         ToolTipText     =   "Shows status of a Velocity Feed Back loop"
         Top             =   495
         Width           =   510
         _Version        =   524288
         _ExtentX        =   891
         _ExtentY        =   503
         _StockProps     =   68
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Reset_0         =   0   'False
         CompatibleVers_0=   524288
         Boolean_0       =   1
         ClassName_1     =   "CCWBoolean"
         opts_1          =   2622
         C[0]_1          =   -2147483640
         Enum_1          =   2
         ClassName_2     =   "CCWEnum"
         Array_2         =   2
         Editor_2        =   0
         Array[0]_2      =   3
         ClassName_3     =   "CCWEnumElt"
         opts_3          =   1
         Name_3          =   "Off"
         frame_3         =   288
         DrawList_3      =   4
         ClassName_4     =   "CDrawList"
         count_4         =   4
         list[4]_4       =   5
         ClassName_5     =   "CCWDrawObj"
         opts_5          =   62
         C[0]_5          =   -2147483640
         C[1]_5          =   -2147483640
         Image_5         =   6
         ClassName_6     =   "CCWTextImage"
         font_6          =   0
         Animator_5      =   0
         Blinker_5       =   0
         list[3]_4       =   7
         ClassName_7     =   "CCWDrawObj"
         opts_7          =   60
         C[0]_7          =   -2147483640
         C[1]_7          =   -2147483640
         Image_7         =   8
         ClassName_8     =   "CCWTextImage"
         szText_8        =   "Z on"
         font_8          =   0
         Animator_7      =   0
         Blinker_7       =   0
         list[2]_4       =   9
         ClassName_9     =   "CCWDrawObj"
         opts_9          =   54
         SclRef_9.l      =   5
         SclRef_9.t      =   5
         SclRef_9.r      =   31
         SclRef_9.b      =   17
         Scl_9.l         =   7
         Scl_9.t         =   5
         Scl_9.r         =   29
         Scl_9.b         =   19
         C[0]_9          =   -2147483640
         C[1]_9          =   -2147483640
         Image_9         =   10
         ClassName_10    =   "CCWTextImage"
         szText_10       =   "Zoff"
         font_10         =   0
         Animator_9      =   0
         Blinker_9       =   0
         list[1]_4       =   11
         ClassName_11    =   "CCWDrawObj"
         opts_11         =   62
         C[0]_11         =   8438015
         C[1]_11         =   8438015
         Image_11        =   12
         ClassName_12    =   "CCWPictImage"
         opts_12         =   1280
         Rows_12         =   1
         Cols_12         =   1
         Pict_12         =   2
         F_12            =   8438015
         B_12            =   8438015
         ColorReplaceWith_12=   8421504
         ColorReplace_12 =   8421504
         Tolerance_12    =   2
         Animator_11     =   0
         Blinker_11      =   0
         varVarType_3    =   5
         Array[1]_2      =   13
         ClassName_13    =   "CCWEnumElt"
         opts_13         =   1
         Name_13         =   "On"
         frame_13        =   287
         DrawList_13     =   14
         ClassName_14    =   "CDrawList"
         count_14        =   4
         list[4]_14      =   15
         ClassName_15    =   "CCWDrawObj"
         opts_15         =   62
         C[0]_15         =   -2147483640
         C[1]_15         =   -2147483640
         Image_15        =   6
         Animator_15     =   0
         Blinker_15      =   0
         list[3]_14      =   16
         ClassName_16    =   "CCWDrawObj"
         opts_16         =   62
         C[0]_16         =   -2147483640
         C[1]_16         =   -2147483640
         Image_16        =   8
         Animator_16     =   0
         Blinker_16      =   0
         list[2]_14      =   17
         ClassName_17    =   "CCWDrawObj"
         opts_17         =   52
         SclRef_17.l     =   5
         SclRef_17.t     =   5
         SclRef_17.r     =   31
         SclRef_17.b     =   17
         Scl_17.l        =   7
         Scl_17.t        =   5
         Scl_17.r        =   29
         Scl_17.b        =   19
         C[0]_17         =   -2147483640
         C[1]_17         =   -2147483640
         Image_17        =   10
         Animator_17     =   0
         Blinker_17      =   0
         list[1]_14      =   18
         ClassName_18    =   "CCWDrawObj"
         opts_18         =   62
         C[0]_18         =   65280
         C[1]_18         =   65280
         Image_18        =   19
         ClassName_19    =   "CCWPictImage"
         opts_19         =   1280
         Rows_19         =   1
         Cols_19         =   1
         Pict_19         =   2
         F_19            =   65280
         B_19            =   65280
         ColorReplaceWith_19=   8421504
         ColorReplace_19 =   8421504
         Tolerance_19    =   2
         Animator_18     =   0
         Blinker_18      =   0
         varVarType_13   =   5
         Bindings_1      =   20
         ClassName_20    =   "CCWBindingHolderArray"
         Editor_20       =   21
         ClassName_21    =   "CCWBindingHolderArrayEditor"
         Owner_21        =   1
         Style_1         =   12
         frameStyle_1    =   1
         mechAction_1    =   3
         BGImg_1         =   22
         ClassName_22    =   "CCWDrawObj"
         opts_22         =   62
         Image_22        =   23
         ClassName_23    =   "CCWPictImage"
         opts_23         =   1280
         Rows_23         =   1
         Cols_23         =   1
         Pict_23         =   286
         F_23            =   -2147483633
         B_23            =   -2147483633
         ColorReplaceWith_23=   8421504
         ColorReplace_23 =   8421504
         Tolerance_23    =   2
         Animator_22     =   0
         Blinker_22      =   0
         Array_1         =   6
         Editor_1        =   0
         Array[0]_1      =   11
         Array[1]_1      =   18
         Array[2]_1      =   0
         Array[3]_1      =   0
         Array[4]_1      =   24
         ClassName_24    =   "CCWDrawObj"
         opts_24         =   62
         Image_24        =   8
         Animator_24     =   0
         Blinker_24      =   0
         Array[5]_1      =   25
         ClassName_25    =   "CCWDrawObj"
         opts_25         =   62
         Image_25        =   10
         Animator_25     =   0
         Blinker_25      =   0
         Label_1         =   26
         ClassName_26    =   "CCWDrawObj"
         opts_26         =   62
         C[0]_26         =   -2147483640
         Image_26        =   6
         Animator_26     =   0
         Blinker_26      =   0
      End
      Begin CWUIControlsLib.CWButton cwAxisFBstatus
         Height          =   285
         Index           =   3
         Left            =   5622
         TabIndex        =   113
         ToolTipText     =   "Shows status of a Velocity Feed Back loop"
         Top             =   495
         Width           =   510
         _Version        =   524288
         _ExtentX        =   891
         _ExtentY        =   503
         _StockProps     =   68
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Reset_0         =   0   'False
         CompatibleVers_0=   524288
         Boolean_0       =   1
         ClassName_1     =   "CCWBoolean"
         opts_1          =   2622
         C[0]_1          =   -2147483640
         Enum_1          =   2
         ClassName_2     =   "CCWEnum"
         Array_2         =   2
         Editor_2        =   0
         Array[0]_2      =   3
         ClassName_3     =   "CCWEnumElt"
         opts_3          =   1
         Name_3          =   "Off"
         frame_3         =   288
         DrawList_3      =   4
         ClassName_4     =   "CDrawList"
         count_4         =   4
         list[4]_4       =   5
         ClassName_5     =   "CCWDrawObj"
         opts_5          =   62
         C[0]_5          =   -2147483640
         C[1]_5          =   -2147483640
         Image_5         =   6
         ClassName_6     =   "CCWTextImage"
         font_6          =   0
         Animator_5      =   0
         Blinker_5       =   0
         list[3]_4       =   7
         ClassName_7     =   "CCWDrawObj"
         opts_7          =   60
         C[0]_7          =   -2147483640
         C[1]_7          =   -2147483640
         Image_7         =   8
         ClassName_8     =   "CCWTextImage"
         szText_8        =   "tX on"
         font_8          =   0
         Animator_7      =   0
         Blinker_7       =   0
         list[2]_4       =   9
         ClassName_9     =   "CCWDrawObj"
         opts_9          =   54
         SclRef_9.l      =   5
         SclRef_9.t      =   5
         SclRef_9.r      =   31
         SclRef_9.b      =   17
         Scl_9.l         =   5
         Scl_9.t         =   5
         Scl_9.r         =   31
         Scl_9.b         =   19
         C[0]_9          =   -2147483640
         C[1]_9          =   -2147483640
         Image_9         =   10
         ClassName_10    =   "CCWTextImage"
         szText_10       =   "tXoff"
         font_10         =   0
         Animator_9      =   0
         Blinker_9       =   0
         list[1]_4       =   11
         ClassName_11    =   "CCWDrawObj"
         opts_11         =   62
         C[0]_11         =   8438015
         C[1]_11         =   8438015
         Image_11        =   12
         ClassName_12    =   "CCWPictImage"
         opts_12         =   1280
         Rows_12         =   1
         Cols_12         =   1
         Pict_12         =   2
         F_12            =   8438015
         B_12            =   8438015
         ColorReplaceWith_12=   8421504
         ColorReplace_12 =   8421504
         Tolerance_12    =   2
         Animator_11     =   0
         Blinker_11      =   0
         varVarType_3    =   5
         Array[1]_2      =   13
         ClassName_13    =   "CCWEnumElt"
         opts_13         =   1
         Name_13         =   "On"
         frame_13        =   287
         DrawList_13     =   14
         ClassName_14    =   "CDrawList"
         count_14        =   4
         list[4]_14      =   15
         ClassName_15    =   "CCWDrawObj"
         opts_15         =   62
         C[0]_15         =   -2147483640
         C[1]_15         =   -2147483640
         Image_15        =   6
         Animator_15     =   0
         Blinker_15      =   0
         list[3]_14      =   16
         ClassName_16    =   "CCWDrawObj"
         opts_16         =   62
         C[0]_16         =   -2147483640
         C[1]_16         =   -2147483640
         Image_16        =   8
         Animator_16     =   0
         Blinker_16      =   0
         list[2]_14      =   17
         ClassName_17    =   "CCWDrawObj"
         opts_17         =   52
         SclRef_17.l     =   5
         SclRef_17.t     =   5
         SclRef_17.r     =   31
         SclRef_17.b     =   17
         Scl_17.l        =   5
         Scl_17.t        =   5
         Scl_17.r        =   31
         Scl_17.b        =   19
         C[0]_17         =   -2147483640
         C[1]_17         =   -2147483640
         Image_17        =   10
         Animator_17     =   0
         Blinker_17      =   0
         list[1]_14      =   18
         ClassName_18    =   "CCWDrawObj"
         opts_18         =   62
         C[0]_18         =   65280
         C[1]_18         =   65280
         Image_18        =   19
         ClassName_19    =   "CCWPictImage"
         opts_19         =   1280
         Rows_19         =   1
         Cols_19         =   1
         Pict_19         =   2
         F_19            =   65280
         B_19            =   65280
         ColorReplaceWith_19=   8421504
         ColorReplace_19 =   8421504
         Tolerance_19    =   2
         Animator_18     =   0
         Blinker_18      =   0
         varVarType_13   =   5
         Bindings_1      =   20
         ClassName_20    =   "CCWBindingHolderArray"
         Editor_20       =   21
         ClassName_21    =   "CCWBindingHolderArrayEditor"
         Owner_21        =   1
         Style_1         =   12
         frameStyle_1    =   1
         mechAction_1    =   3
         BGImg_1         =   22
         ClassName_22    =   "CCWDrawObj"
         opts_22         =   62
         Image_22        =   23
         ClassName_23    =   "CCWPictImage"
         opts_23         =   1280
         Rows_23         =   1
         Cols_23         =   1
         Pict_23         =   286
         F_23            =   -2147483633
         B_23            =   -2147483633
         ColorReplaceWith_23=   8421504
         ColorReplace_23 =   8421504
         Tolerance_23    =   2
         Animator_22     =   0
         Blinker_22      =   0
         Array_1         =   6
         Editor_1        =   0
         Array[0]_1      =   11
         Array[1]_1      =   18
         Array[2]_1      =   0
         Array[3]_1      =   0
         Array[4]_1      =   24
         ClassName_24    =   "CCWDrawObj"
         opts_24         =   62
         Image_24        =   8
         Animator_24     =   0
         Blinker_24      =   0
         Array[5]_1      =   25
         ClassName_25    =   "CCWDrawObj"
         opts_25         =   62
         Image_25        =   10
         Animator_25     =   0
         Blinker_25      =   0
         Label_1         =   26
         ClassName_26    =   "CCWDrawObj"
         opts_26         =   62
         C[0]_26         =   -2147483640
         Image_26        =   6
         Animator_26     =   0
         Blinker_26      =   0
      End
      Begin CWUIControlsLib.CWButton cwAxisFBstatus
         Height          =   285
         Index           =   4
         Left            =   6208
         TabIndex        =   114
         ToolTipText     =   "Shows status of a Velocity Feed Back loop"
         Top             =   495
         Width           =   510
         _Version        =   524288
         _ExtentX        =   891
         _ExtentY        =   503
         _StockProps     =   68
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Reset_0         =   0   'False
         CompatibleVers_0=   524288
         Boolean_0       =   1
         ClassName_1     =   "CCWBoolean"
         opts_1          =   2622
         C[0]_1          =   -2147483640
         Enum_1          =   2
         ClassName_2     =   "CCWEnum"
         Array_2         =   2
         Editor_2        =   0
         Array[0]_2      =   3
         ClassName_3     =   "CCWEnumElt"
         opts_3          =   1
         Name_3          =   "Off"
         frame_3         =   288
         DrawList_3      =   4
         ClassName_4     =   "CDrawList"
         count_4         =   4
         list[4]_4       =   5
         ClassName_5     =   "CCWDrawObj"
         opts_5          =   62
         C[0]_5          =   -2147483640
         C[1]_5          =   -2147483640
         Image_5         =   6
         ClassName_6     =   "CCWTextImage"
         font_6          =   0
         Animator_5      =   0
         Blinker_5       =   0
         list[3]_4       =   7
         ClassName_7     =   "CCWDrawObj"
         opts_7          =   60
         C[0]_7          =   -2147483640
         C[1]_7          =   -2147483640
         Image_7         =   8
         ClassName_8     =   "CCWTextImage"
         szText_8        =   "tY on"
         font_8          =   0
         Animator_7      =   0
         Blinker_7       =   0
         list[2]_4       =   9
         ClassName_9     =   "CCWDrawObj"
         opts_9          =   54
         SclRef_9.l      =   5
         SclRef_9.t      =   5
         SclRef_9.r      =   31
         SclRef_9.b      =   17
         Scl_9.l         =   6
         Scl_9.t         =   5
         Scl_9.r         =   32
         Scl_9.b         =   19
         C[0]_9          =   -2147483640
         C[1]_9          =   -2147483640
         Image_9         =   10
         ClassName_10    =   "CCWTextImage"
         szText_10       =   "tYoff"
         font_10         =   0
         Animator_9      =   0
         Blinker_9       =   0
         list[1]_4       =   11
         ClassName_11    =   "CCWDrawObj"
         opts_11         =   62
         C[0]_11         =   8438015
         C[1]_11         =   8438015
         Image_11        =   12
         ClassName_12    =   "CCWPictImage"
         opts_12         =   1280
         Rows_12         =   1
         Cols_12         =   1
         Pict_12         =   2
         F_12            =   8438015
         B_12            =   8438015
         ColorReplaceWith_12=   8421504
         ColorReplace_12 =   8421504
         Tolerance_12    =   2
         Animator_11     =   0
         Blinker_11      =   0
         varVarType_3    =   5
         Array[1]_2      =   13
         ClassName_13    =   "CCWEnumElt"
         opts_13         =   1
         Name_13         =   "On"
         frame_13        =   287
         DrawList_13     =   14
         ClassName_14    =   "CDrawList"
         count_14        =   4
         list[4]_14      =   15
         ClassName_15    =   "CCWDrawObj"
         opts_15         =   62
         C[0]_15         =   -2147483640
         C[1]_15         =   -2147483640
         Image_15        =   6
         Animator_15     =   0
         Blinker_15      =   0
         list[3]_14      =   16
         ClassName_16    =   "CCWDrawObj"
         opts_16         =   62
         C[0]_16         =   -2147483640
         C[1]_16         =   -2147483640
         Image_16        =   8
         Animator_16     =   0
         Blinker_16      =   0
         list[2]_14      =   17
         ClassName_17    =   "CCWDrawObj"
         opts_17         =   52
         SclRef_17.l     =   5
         SclRef_17.t     =   5
         SclRef_17.r     =   31
         SclRef_17.b     =   17
         Scl_17.l        =   6
         Scl_17.t        =   5
         Scl_17.r        =   32
         Scl_17.b        =   19
         C[0]_17         =   -2147483640
         C[1]_17         =   -2147483640
         Image_17        =   10
         Animator_17     =   0
         Blinker_17      =   0
         list[1]_14      =   18
         ClassName_18    =   "CCWDrawObj"
         opts_18         =   62
         C[0]_18         =   65280
         C[1]_18         =   65280
         Image_18        =   19
         ClassName_19    =   "CCWPictImage"
         opts_19         =   1280
         Rows_19         =   1
         Cols_19         =   1
         Pict_19         =   2
         F_19            =   65280
         B_19            =   65280
         ColorReplaceWith_19=   8421504
         ColorReplace_19 =   8421504
         Tolerance_19    =   2
         Animator_18     =   0
         Blinker_18      =   0
         varVarType_13   =   5
         Bindings_1      =   20
         ClassName_20    =   "CCWBindingHolderArray"
         Editor_20       =   21
         ClassName_21    =   "CCWBindingHolderArrayEditor"
         Owner_21        =   1
         Style_1         =   12
         frameStyle_1    =   1
         mechAction_1    =   3
         BGImg_1         =   22
         ClassName_22    =   "CCWDrawObj"
         opts_22         =   62
         Image_22        =   23
         ClassName_23    =   "CCWPictImage"
         opts_23         =   1280
         Rows_23         =   1
         Cols_23         =   1
         Pict_23         =   286
         F_23            =   -2147483633
         B_23            =   -2147483633
         ColorReplaceWith_23=   8421504
         ColorReplace_23 =   8421504
         Tolerance_23    =   2
         Animator_22     =   0
         Blinker_22      =   0
         Array_1         =   6
         Editor_1        =   0
         Array[0]_1      =   11
         Array[1]_1      =   18
         Array[2]_1      =   0
         Array[3]_1      =   0
         Array[4]_1      =   24
         ClassName_24    =   "CCWDrawObj"
         opts_24         =   62
         Image_24        =   8
         Animator_24     =   0
         Blinker_24      =   0
         Array[5]_1      =   25
         ClassName_25    =   "CCWDrawObj"
         opts_25         =   62
         Image_25        =   10
         Animator_25     =   0
         Blinker_25      =   0
         Label_1         =   26
         ClassName_26    =   "CCWDrawObj"
         opts_26         =   62
         C[0]_26         =   -2147483640
         Image_26        =   6
         Animator_26     =   0
         Blinker_26      =   0
      End
      Begin CWUIControlsLib.CWButton cwAxisFBstatus
         Height          =   285
         Index           =   5
         Left            =   6790
         TabIndex        =   115
         ToolTipText     =   "Shows status of a Velocity Feed Back loop"
         Top             =   495
         Width           =   510
         _Version        =   524288
         _ExtentX        =   891
         _ExtentY        =   503
         _StockProps     =   68
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Reset_0         =   0   'False
         CompatibleVers_0=   524288
         Boolean_0       =   1
         ClassName_1     =   "CCWBoolean"
         opts_1          =   2622
         C[0]_1          =   -2147483640
         Enum_1          =   2
         ClassName_2     =   "CCWEnum"
         Array_2         =   2
         Editor_2        =   0
         Array[0]_2      =   3
         ClassName_3     =   "CCWEnumElt"
         opts_3          =   1
         Name_3          =   "Off"
         frame_3         =   288
         DrawList_3      =   4
         ClassName_4     =   "CDrawList"
         count_4         =   4
         list[4]_4       =   5
         ClassName_5     =   "CCWDrawObj"
         opts_5          =   62
         C[0]_5          =   -2147483640
         C[1]_5          =   -2147483640
         Image_5         =   6
         ClassName_6     =   "CCWTextImage"
         font_6          =   0
         Animator_5      =   0
         Blinker_5       =   0
         list[3]_4       =   7
         ClassName_7     =   "CCWDrawObj"
         opts_7          =   60
         C[0]_7          =   -2147483640
         C[1]_7          =   -2147483640
         Image_7         =   8
         ClassName_8     =   "CCWTextImage"
         szText_8        =   "tZ on"
         font_8          =   0
         Animator_7      =   0
         Blinker_7       =   0
         list[2]_4       =   9
         ClassName_9     =   "CCWDrawObj"
         opts_9          =   54
         SclRef_9.l      =   5
         SclRef_9.t      =   5
         SclRef_9.r      =   31
         SclRef_9.b      =   17
         Scl_9.l         =   5
         Scl_9.t         =   5
         Scl_9.r         =   31
         Scl_9.b         =   19
         C[0]_9          =   -2147483640
         C[1]_9          =   -2147483640
         Image_9         =   10
         ClassName_10    =   "CCWTextImage"
         szText_10       =   "tZoff"
         font_10         =   0
         Animator_9      =   0
         Blinker_9       =   0
         list[1]_4       =   11
         ClassName_11    =   "CCWDrawObj"
         opts_11         =   62
         C[0]_11         =   8438015
         C[1]_11         =   8438015
         Image_11        =   12
         ClassName_12    =   "CCWPictImage"
         opts_12         =   1280
         Rows_12         =   1
         Cols_12         =   1
         Pict_12         =   2
         F_12            =   8438015
         B_12            =   8438015
         ColorReplaceWith_12=   8421504
         ColorReplace_12 =   8421504
         Tolerance_12    =   2
         Animator_11     =   0
         Blinker_11      =   0
         varVarType_3    =   5
         Array[1]_2      =   13
         ClassName_13    =   "CCWEnumElt"
         opts_13         =   1
         Name_13         =   "On"
         frame_13        =   287
         DrawList_13     =   14
         ClassName_14    =   "CDrawList"
         count_14        =   4
         list[4]_14      =   15
         ClassName_15    =   "CCWDrawObj"
         opts_15         =   62
         C[0]_15         =   -2147483640
         C[1]_15         =   -2147483640
         Image_15        =   6
         Animator_15     =   0
         Blinker_15      =   0
         list[3]_14      =   16
         ClassName_16    =   "CCWDrawObj"
         opts_16         =   62
         C[0]_16         =   -2147483640
         C[1]_16         =   -2147483640
         Image_16        =   8
         Animator_16     =   0
         Blinker_16      =   0
         list[2]_14      =   17
         ClassName_17    =   "CCWDrawObj"
         opts_17         =   52
         SclRef_17.l     =   5
         SclRef_17.t     =   5
         SclRef_17.r     =   31
         SclRef_17.b     =   17
         Scl_17.l        =   5
         Scl_17.t        =   5
         Scl_17.r        =   31
         Scl_17.b        =   19
         C[0]_17         =   -2147483640
         C[1]_17         =   -2147483640
         Image_17        =   10
         Animator_17     =   0
         Blinker_17      =   0
         list[1]_14      =   18
         ClassName_18    =   "CCWDrawObj"
         opts_18         =   62
         C[0]_18         =   65280
         C[1]_18         =   65280
         Image_18        =   19
         ClassName_19    =   "CCWPictImage"
         opts_19         =   1280
         Rows_19         =   1
         Cols_19         =   1
         Pict_19         =   2
         F_19            =   65280
         B_19            =   65280
         ColorReplaceWith_19=   8421504
         ColorReplace_19 =   8421504
         Tolerance_19    =   2
         Animator_18     =   0
         Blinker_18      =   0
         varVarType_13   =   5
         Bindings_1      =   20
         ClassName_20    =   "CCWBindingHolderArray"
         Editor_20       =   21
         ClassName_21    =   "CCWBindingHolderArrayEditor"
         Owner_21        =   1
         Style_1         =   12
         frameStyle_1    =   1
         mechAction_1    =   3
         BGImg_1         =   22
         ClassName_22    =   "CCWDrawObj"
         opts_22         =   62
         Image_22        =   23
         ClassName_23    =   "CCWPictImage"
         opts_23         =   1280
         Rows_23         =   1
         Cols_23         =   1
         Pict_23         =   286
         F_23            =   -2147483633
         B_23            =   -2147483633
         ColorReplaceWith_23=   8421504
         ColorReplace_23 =   8421504
         Tolerance_23    =   2
         Animator_22     =   0
         Blinker_22      =   0
         Array_1         =   6
         Editor_1        =   0
         Array[0]_1      =   11
         Array[1]_1      =   18
         Array[2]_1      =   0
         Array[3]_1      =   0
         Array[4]_1      =   24
         ClassName_24    =   "CCWDrawObj"
         opts_24         =   62
         Image_24        =   8
         Animator_24     =   0
         Blinker_24      =   0
         Array[5]_1      =   25
         ClassName_25    =   "CCWDrawObj"
         opts_25         =   62
         Image_25        =   10
         Animator_25     =   0
         Blinker_25      =   0
         Label_1         =   26
         ClassName_26    =   "CCWDrawObj"
         opts_26         =   62
         C[0]_26         =   -2147483640
         Image_26        =   6
         Animator_26     =   0
         Blinker_26      =   0
      End
      Begin VB.Label lbl_RTstatus
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "TooLowInpPSI"
         BeginProperty Font
            Name            =   "Arial"
            Size            =   9
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   5625
         TabIndex        =   193
         Top             =   155
         Width           =   1680
      End
      Begin VB.Label Label1
         BackColor       =   &H00C0FFFF&
         Caption         =   "cmd      exec.time, s  "
         Height          =   245
         Index           =   3
         Left            =   9360
         TabIndex        =   192
         Top             =   225
         Width           =   1455
      End
      Begin VB.Label lblCMDtime
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "?11.33"
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
         Height          =   275
         Left            =   10845
         TabIndex        =   191
         ToolTipText     =   "Execution time of the last command sent to controller"
         Top             =   210
         Width           =   780
      End
      Begin VB.Label lblAxis
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "RT Monitoring"
         BeginProperty Font
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   3
         Left            =   0
         TabIndex        =   116
         ToolTipText     =   "Checking position and status of the system once a second"
         Top             =   100
         Width           =   1560
      End
      Begin VB.Label lbl
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Motor Power"
         BeginProperty Font
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   2
         Left            =   5220
         TabIndex        =   109
         Top             =   855
         Width           =   1365
      End
      Begin VB.Label lbl_MotPower
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "88%"
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
         Left            =   6600
         TabIndex        =   108
         ToolTipText     =   "Horizontal #1 prox sensor signal in microns"
         Top             =   855
         Width           =   700
      End
      Begin VB.Label lbl_ValveCurrent
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "122.2"
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
         Left            =   10125
         TabIndex        =   105
         ToolTipText     =   "Valve #4 current, mA"
         Top             =   1260
         Width           =   645
      End
      Begin VB.Label lbl_ValveCurrent
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "111.0"
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
         Left            =   9405
         TabIndex        =   104
         ToolTipText     =   "Valve #3 current, mA"
         Top             =   1260
         Width           =   645
      End
      Begin VB.Label lbl_ValveCurrent
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "99.9"
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
         Left            =   8685
         TabIndex        =   103
         ToolTipText     =   "Valve #2 current, mA"
         Top             =   1260
         Width           =   645
      End
      Begin VB.Label lbl_ValveCurrent
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0FFFF&
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
         Left            =   7965
         TabIndex        =   102
         ToolTipText     =   "Valve #1 current, mA"
         Top             =   1260
         Width           =   645
      End
      Begin VB.Label lbl_IsoPressure
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "77.77"
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
         Index           =   4
         Left            =   10845
         TabIndex        =   101
         ToolTipText     =   "Input pressure, PSI"
         Top             =   855
         Width           =   690
      End
      Begin VB.Label lbl_IsoPressure
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "44.44"
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
         Left            =   10125
         TabIndex        =   100
         ToolTipText     =   "Isolator #4 pressure, PSI"
         Top             =   855
         Width           =   645
      End
      Begin VB.Label lbl_IsoPressure
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "33.33"
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
         Left            =   9405
         TabIndex        =   99
         ToolTipText     =   "Isolator #3 pressure, PSI"
         Top             =   855
         Width           =   645
      End
      Begin VB.Label lbl_IsoPressure
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "22.22"
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
         Left            =   8685
         TabIndex        =   98
         ToolTipText     =   "Isolator #2 pressure, PSI"
         Top             =   855
         Width           =   645
      End
      Begin VB.Label lbl_IsoPressure
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "12.34"
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
         Left            =   7965
         TabIndex        =   97
         ToolTipText     =   "Isolator #1 pressure, PSI"
         Top             =   855
         Width           =   645
      End
      Begin VB.Label lbl_mA
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "mA"
         BeginProperty Font
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   7425
         TabIndex        =   96
         Top             =   1260
         Width           =   465
      End
      Begin VB.Label lbl_PSI
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "PSI"
         BeginProperty Font
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   7425
         TabIndex        =   95
         Top             =   855
         Width           =   465
      End
      Begin VB.Label lbValveNum
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Input"
         BeginProperty Font
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   270
         Index           =   4
         Left            =   10845
         TabIndex        =   94
         Top             =   540
         Width           =   645
      End
      Begin VB.Label lbValveNum
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Iso 4"
         BeginProperty Font
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   270
         Index           =   3
         Left            =   10125
         TabIndex        =   93
         Top             =   540
         Width           =   645
      End
      Begin VB.Label lbValveNum
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Iso 3"
         BeginProperty Font
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   270
         Index           =   2
         Left            =   9405
         TabIndex        =   92
         Top             =   540
         Width           =   645
      End
      Begin VB.Label lbValveNum
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Iso 2"
         BeginProperty Font
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   270
         Index           =   1
         Left            =   8685
         TabIndex        =   91
         Top             =   540
         Width           =   645
      End
      Begin VB.Label lbValveNum
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Iso 1"
         BeginProperty Font
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   270
         Index           =   0
         Left            =   7965
         TabIndex        =   90
         Top             =   540
         Width           =   645
      End
      Begin VB.Label lbl_ProxSensor
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0080FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "-1200"
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
         Left            =   1575
         TabIndex        =   82
         ToolTipText     =   "Vertical #1 prox sensor signal in microns"
         Top             =   495
         Width           =   645
      End
      Begin VB.Label lbIsoNum
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Prox 1"
         BeginProperty Font
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   0
         Left            =   1575
         TabIndex        =   81
         Top             =   225
         Width           =   690
      End
      Begin VB.Label lbl_ProxSensor
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0080FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "-6789"
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
         Index           =   4
         Left            =   2295
         TabIndex        =   80
         ToolTipText     =   "Vertical #2 prox sensor signal in microns"
         Top             =   495
         Width           =   645
      End
      Begin VB.Label lbIsoNum
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Prox 2"
         BeginProperty Font
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   1
         Left            =   2295
         TabIndex        =   79
         Top             =   225
         Width           =   690
      End
      Begin VB.Label lbl_ProxSensor
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0080FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "-6789"
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
         Index           =   5
         Left            =   3015
         TabIndex        =   78
         ToolTipText     =   "Vertical #3 prox sensor signal in microns"
         Top             =   495
         Width           =   645
      End
      Begin VB.Label lbIsoNum
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Prox 3"
         BeginProperty Font
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   2
         Left            =   3015
         TabIndex        =   77
         Top             =   225
         Width           =   690
      End
      Begin VB.Label lbl_ProxSensor
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0080FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "1234"
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
         Left            =   1575
         TabIndex        =   76
         ToolTipText     =   "Horizontal #1 prox sensor signal in microns"
         Top             =   855
         Width           =   645
      End
      Begin VB.Label lbl_RTM
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Vert"
         BeginProperty Font
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   0
         Left            =   945
         TabIndex        =   75
         Top             =   540
         Width           =   600
      End
      Begin VB.Label lbl_RTM
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Horiz"
         BeginProperty Font
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   1
         Left            =   945
         TabIndex        =   74
         Top             =   900
         Width           =   600
      End
      Begin VB.Label lbl_ProxSensor
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0080FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "1234"
         BeginProperty Font
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   1
         Left            =   2295
         TabIndex        =   73
         ToolTipText     =   "Horizontal #2 prox sensor signal in microns"
         Top             =   855
         Width           =   645
      End
      Begin VB.Label lbl_ProxSensor
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0080FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "1234"
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
         Left            =   3015
         TabIndex        =   72
         ToolTipText     =   "Horizontal #3 prox sensor signal in microns"
         Top             =   855
         Width           =   645
      End
   End
   Begin VB.TextBox txtStatus
      BackColor       =   &H00FFFFC0&
      BeginProperty Font
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   284
      Left            =   2745
      TabIndex        =   24
      Text            =   "  Status"
      Top             =   8415
      Width           =   9375
   End
   Begin VB.Frame FrameBNC0
      BorderStyle     =   0  'None
      Height          =   7575
      Left            =   2835
      TabIndex        =   119
      Top             =   1170
      Visible         =   0   'False
      Width           =   1930
      Begin CWUIControlsLib.CWSlide cwBNC0select
         Height          =   7530
         Left            =   10
         TabIndex        =   121
         ToolTipText     =   "Changes BNC Diag.0 signal"
         Top             =   10
         Width           =   1875
         _Version        =   524288
         _ExtentX        =   3307
         _ExtentY        =   13282
         _StockProps     =   68
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Reset_0         =   0   'False
         CompatibleVers_0=   524288
         Slider_0        =   1
         ClassName_1     =   "CCWSlider"
         opts_1          =   2102
         SclRef_1.l      =   2
         SclRef_1.t      =   2
         SclRef_1.r      =   132
         SclRef_1.b      =   484
         Scl_1.l         =   2
         Scl_1.t         =   2
         Scl_1.r         =   132
         Scl_1.b         =   478
         C[0]_1          =   16777215
         BGImg_1         =   2
         ClassName_2     =   "CCWDrawObj"
         opts_2          =   54
         SclRef_2.l      =   2
         SclRef_2.t      =   2
         SclRef_2.r      =   132
         SclRef_2.b      =   502
         Scl_2.l         =   2
         Scl_2.t         =   2
         Scl_2.r         =   133
         Scl_2.b         =   502
         Image_2         =   3
         ClassName_3     =   "CCWPictImage"
         opts_3          =   1280
         Rows_3          =   1
         Cols_3          =   1
         Pict_3          =   286
         F_3             =   -2147483633
         B_3             =   -2147483633
         ColorReplaceWith_3=   8421504
         ColorReplace_3  =   8421504
         Tolerance_3     =   2
         Animator_2      =   0
         Blinker_2       =   0
         BFImg_1         =   4
         ClassName_4     =   "CCWDrawObj"
         opts_4          =   62
         Image_4         =   5
         ClassName_5     =   "CCWPictImage"
         opts_5          =   1280
         Rows_5          =   1
         Cols_5          =   1
         Pict_5          =   286
         F_5             =   -2147483633
         B_5             =   -2147483633
         ColorReplaceWith_5=   8421504
         ColorReplace_5  =   8421504
         Tolerance_5     =   2
         Animator_4      =   0
         Blinker_4       =   0
         style_1         =   4
         Label_1         =   6
         ClassName_6     =   "CCWDrawObj"
         opts_6          =   54
         SclRef_6.l      =   2
         SclRef_6.t      =   2
         SclRef_6.r      =   132
         SclRef_6.b      =   504
         Scl_6.l         =   58
         Scl_6.t         =   2
         Scl_6.r         =   99
         Scl_6.b         =   18
         C[0]_6          =   -2147483640
         Image_6         =   7
         ClassName_7     =   "CCWTextImage"
         szText_7        =   "BNC_0"
         font_7          =   0
         Animator_6      =   0
         Blinker_6       =   0
         Border_1        =   8
         ClassName_8     =   "CCWDrawObj"
         opts_8          =   60
         Image_8         =   9
         ClassName_9     =   "CCWPictImage"
         opts_9          =   1280
         Rows_9          =   1
         Cols_9          =   1
         Pict_9          =   25
         F_9             =   -2147483633
         B_9             =   -2147483633
         ColorReplaceWith_9=   8421504
         ColorReplace_9  =   8421504
         Tolerance_9     =   2
         Animator_8      =   0
         Blinker_8       =   0
         FillBound_1     =   10
         ClassName_10    =   "CCWGuiObject"
         opts_10         =   60
         FillTok_1       =   11
         ClassName_11    =   "CCWGuiObject"
         opts_11         =   62
         Axis_1          =   12
         ClassName_12    =   "CCWAxis"
         opts_12         =   1079
         Name_12         =   "Axis"
         SclRef_12.l     =   2
         SclRef_12.t     =   2
         SclRef_12.r     =   132
         SclRef_12.b     =   498
         Scl_12.l        =   2
         Scl_12.t        =   27
         Scl_12.r        =   123
         Scl_12.b        =   481
         Orientation_12  =   133172
         format_12       =   13
         ClassName_13    =   "CCWFormat"
         Scale_12        =   14
         ClassName_14    =   "CCWScale"
         opts_14         =   57344
         rMin_14         =   25
         rMax_14         =   478
         dMin_14         =   1
         dMax_14         =   38
         discInterval_14 =   1
         Radial_12       =   0
         Enum_12         =   15
         ClassName_15    =   "CCWEnum"
         Array_15        =   38
         Editor_15       =   16
         ClassName_16    =   "CCWEnumArrayEditor"
         Owner_16        =   12
         Array[0]_15     =   17
         ClassName_17    =   "CCWEnumElt"
         opts_17         =   1
         Name_17         =   "Excitation"
         DrawList_17     =   0
         varVarType_17   =   2
         var_Val_17      =   120
         Array[1]_15     =   18
         ClassName_18    =   "CCWEnumElt"
         opts_18         =   1
         Name_18         =   "Pos_tY"
         DrawList_18     =   0
         varVarType_18   =   2
         var_Val_18      =   59
         Array[2]_15     =   19
         ClassName_19    =   "CCWEnumElt"
         opts_19         =   1
         Name_19         =   "Pos_tX"
         DrawList_19     =   0
         varVarType_19   =   2
         var_Val_19      =   58
         Array[3]_15     =   20
         ClassName_20    =   "CCWEnumElt"
         opts_20         =   1
         Name_20         =   "Pos_Z"
         DrawList_20     =   0
         varVarType_20   =   2
         var_Val_20      =   57
         Array[4]_15     =   21
         ClassName_21    =   "CCWEnumElt"
         opts_21         =   1
         Name_21         =   "Pos_tZ"
         DrawList_21     =   0
         varVarType_21   =   2
         var_Val_21      =   56
         Array[5]_15     =   22
         ClassName_22    =   "CCWEnumElt"
         opts_22         =   1
         Name_22         =   "Pos_Y"
         DrawList_22     =   0
         varVarType_22   =   2
         var_Val_22      =   55
         Array[6]_15     =   23
         ClassName_23    =   "CCWEnumElt"
         opts_23         =   1
         Name_23         =   "Pos_X"
         DrawList_23     =   0
         varVarType_23   =   2
         var_Val_23      =   54
         Array[7]_15     =   24
         ClassName_24    =   "CCWEnumElt"
         opts_24         =   1
         Name_24         =   "Vel_tY"
         DrawList_24     =   0
         varVarType_24   =   2
         var_Val_24      =   53
         Array[8]_15     =   25
         ClassName_25    =   "CCWEnumElt"
         opts_25         =   1
         Name_25         =   "Vel_tX"
         DrawList_25     =   0
         varVarType_25   =   2
         var_Val_25      =   52
         Array[9]_15     =   26
         ClassName_26    =   "CCWEnumElt"
         opts_26         =   1
         Name_26         =   "Vel_Z"
         DrawList_26     =   0
         varVarType_26   =   2
         var_Val_26      =   51
         Array[10]_15    =   27
         ClassName_27    =   "CCWEnumElt"
         opts_27         =   1
         Name_27         =   "Vel_tZ"
         DrawList_27     =   0
         varVarType_27   =   2
         var_Val_27      =   50
         Array[11]_15    =   28
         ClassName_28    =   "CCWEnumElt"
         opts_28         =   1
         Name_28         =   "Vel_Y"
         DrawList_28     =   0
         varVarType_28   =   2
         var_Val_28      =   49
         Array[12]_15    =   29
         ClassName_29    =   "CCWEnumElt"
         opts_29         =   1
         Name_29         =   "Vel_X"
         DrawList_29     =   0
         varVarType_29   =   2
         var_Val_29      =   48
         Array[13]_15    =   30
         ClassName_30    =   "CCWEnumElt"
         opts_30         =   1
         Name_30         =   "Motor 4Z"
         DrawList_30     =   0
         varVarType_30   =   2
         var_Val_30      =   39
         Array[14]_15    =   31
         ClassName_31    =   "CCWEnumElt"
         opts_31         =   1
         Name_31         =   "Motor 3Z"
         DrawList_31     =   0
         varVarType_31   =   2
         var_Val_31      =   38
         Array[15]_15    =   32
         ClassName_32    =   "CCWEnumElt"
         opts_32         =   1
         Name_32         =   "Motor 2Z"
         DrawList_32     =   0
         varVarType_32   =   2
         var_Val_32      =   37
         Array[16]_15    =   33
         ClassName_33    =   "CCWEnumElt"
         opts_33         =   1
         Name_33         =   "Motor 1Z"
         DrawList_33     =   0
         varVarType_33   =   2
         var_Val_33      =   36
         Array[17]_15    =   34
         ClassName_34    =   "CCWEnumElt"
         opts_34         =   1
         Name_34         =   "Motor 4H"
         DrawList_34     =   0
         varVarType_34   =   2
         var_Val_34      =   43
         Array[18]_15    =   35
         ClassName_35    =   "CCWEnumElt"
         opts_35         =   1
         Name_35         =   "Motor 3H"
         DrawList_35     =   0
         varVarType_35   =   2
         var_Val_35      =   42
         Array[19]_15    =   36
         ClassName_36    =   "CCWEnumElt"
         opts_36         =   1
         Name_36         =   "Motor 2H"
         DrawList_36     =   0
         varVarType_36   =   2
         var_Val_36      =   41
         Array[20]_15    =   37
         ClassName_37    =   "CCWEnumElt"
         opts_37         =   1
         Name_37         =   "Motor 1H"
         DrawList_37     =   0
         varVarType_37   =   2
         var_Val_37      =   40
         Array[21]_15    =   38
         ClassName_38    =   "CCWEnumElt"
         opts_38         =   1
         Name_38         =   "Geo 3Z "
         DrawList_38     =   0
         varVarType_38   =   2
         var_Val_38      =   29
         Array[22]_15    =   39
         ClassName_39    =   "CCWEnumElt"
         opts_39         =   1
         Name_39         =   "Geo 2Z "
         DrawList_39     =   0
         varVarType_39   =   2
         var_Val_39      =   28
         Array[23]_15    =   40
         ClassName_40    =   "CCWEnumElt"
         opts_40         =   1
         Name_40         =   "Geo 1Z "
         DrawList_40     =   0
         varVarType_40   =   2
         var_Val_40      =   27
         Array[24]_15    =   41
         ClassName_41    =   "CCWEnumElt"
         opts_41         =   1
         Name_41         =   "Geo 3H"
         DrawList_41     =   0
         varVarType_41   =   2
         var_Val_41      =   26
         Array[25]_15    =   42
         ClassName_42    =   "CCWEnumElt"
         opts_42         =   1
         Name_42         =   "Geo 2H "
         DrawList_42     =   0
         varVarType_42   =   2
         var_Val_42      =   25
         Array[26]_15    =   43
         ClassName_43    =   "CCWEnumElt"
         opts_43         =   1
         Name_43         =   "Geo 1H "
         DrawList_43     =   0
         varVarType_43   =   2
         var_Val_43      =   24
         Array[27]_15    =   44
         ClassName_44    =   "CCWEnumElt"
         opts_44         =   1
         Name_44         =   "StFF_Ypos"
         DrawList_44     =   0
         varVarType_44   =   2
         var_Val_44      =   17
         Array[28]_15    =   45
         ClassName_45    =   "CCWEnumElt"
         opts_45         =   1
         Name_45         =   "StFF_Xpos "
         DrawList_45     =   0
         varVarType_45   =   2
         var_Val_45      =   16
         Array[29]_15    =   46
         ClassName_46    =   "CCWEnumElt"
         opts_46         =   1
         Name_46         =   "StFF_Yacc "
         DrawList_46     =   0
         varVarType_46   =   2
         var_Val_46      =   19
         Array[30]_15    =   47
         ClassName_47    =   "CCWEnumElt"
         opts_47         =   1
         Name_47         =   "StFF_Xacc "
         DrawList_47     =   0
         varVarType_47   =   2
         var_Val_47      =   18
         Array[31]_15    =   48
         ClassName_48    =   "CCWEnumElt"
         opts_48         =   1
         Name_48         =   "Prx_3Z_um"
         DrawList_48     =   0
         varVarType_48   =   2
         var_Val_48      =   6
         Array[32]_15    =   49
         ClassName_49    =   "CCWEnumElt"
         opts_49         =   1
         Name_49         =   "Prx_2Z_um"
         DrawList_49     =   0
         varVarType_49   =   2
         var_Val_49      =   5
         Array[33]_15    =   50
         ClassName_50    =   "CCWEnumElt"
         opts_50         =   1
         Name_50         =   "Prx_1Z_um"
         DrawList_50     =   0
         varVarType_50   =   2
         var_Val_50      =   4
         Array[34]_15    =   51
         ClassName_51    =   "CCWEnumElt"
         opts_51         =   1
         Name_51         =   "Prx_3H_um"
         DrawList_51     =   0
         varVarType_51   =   2
         var_Val_51      =   3
         Array[35]_15    =   52
         ClassName_52    =   "CCWEnumElt"
         opts_52         =   1
         Name_52         =   "Prx_2H_um"
         DrawList_52     =   0
         varVarType_52   =   2
         var_Val_52      =   2
         Array[36]_15    =   53
         ClassName_53    =   "CCWEnumElt"
         opts_53         =   1
         Name_53         =   "Prx_1H_um"
         DrawList_53     =   0
         varVarType_53   =   2
         var_Val_53      =   1
         Array[37]_15    =   54
         ClassName_54    =   "CCWEnumElt"
         opts_54         =   1
         Name_54         =   "--other--"
         DrawList_54     =   0
         varVarType_54   =   2
         var_Val_54      =   -1
         Font_12         =   55
         ClassName_55    =   "CCWFont"
         bFont_55        =   -1  'True
         BeginProperty Font_55 {0BE35203-8F91-11CE-9DE3-00AA004BB851}
            Name            =   "Arial"
            Size            =   9
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         tickopts_12     =   2962
         Caption_12      =   56
         ClassName_56    =   "CCWDrawObj"
         opts_56         =   62
         C[0]_56         =   -2147483640
         Image_56        =   57
         ClassName_57    =   "CCWTextImage"
         font_57         =   0
         Animator_56     =   0
         Blinker_56      =   0
         DrawLst_1       =   58
         ClassName_58    =   "CDrawList"
         count_58        =   10
         list[10]_58     =   8
         list[9]_58      =   59
         ClassName_59    =   "CCWThumb"
         opts_59         =   1048631
         Name_59         =   "Pointer-1"
         SclRef_59.l     =   2
         SclRef_59.t     =   17
         SclRef_59.r     =   132
         SclRef_59.b     =   492
         Scl_59.l        =   11
         Scl_59.t        =   111
         Scl_59.r        =   57
         Scl_59.b        =   124
         C[0]_59         =   255
         C[2]_59         =   -2147483635
         Image_59        =   60
         ClassName_60    =   "CCWPictImage"
         opts_60         =   1280
         Rows_60         =   1
         Cols_60         =   1
         Pict_60         =   94
         F_60            =   255
         B_60            =   -2147483633
         ColorReplaceWith_60=   8421504
         ColorReplace_60 =   8421504
         Tolerance_60    =   2
         Animator_59     =   0
         Blinker_59      =   0
         style_59        =   3
         Value_59        =   4
         Fill_59         =   61
         ClassName_61    =   "CCWDrawObj"
         opts_61         =   62
         Image_61        =   62
         ClassName_62    =   "CCWPictImage"
         opts_62         =   1280
         Rows_62         =   1
         Cols_62         =   1
         Pict_62         =   286
         F_62            =   -2147483633
         B_62            =   -2147483633
         ColorReplaceWith_62=   8421504
         ColorReplace_62 =   8421504
         Tolerance_62    =   2
         Animator_61     =   0
         Blinker_61      =   0
         list[8]_58      =   12
         list[7]_58      =   6
         list[6]_58      =   11
         list[5]_58      =   4
         list[4]_58      =   63
         ClassName_63    =   "CCWDrawObj"
         opts_63         =   54
         SclRef_63.l     =   2
         SclRef_63.t     =   2
         SclRef_63.r     =   132
         SclRef_63.b     =   498
         Scl_63.l        =   20
         Scl_63.t        =   22
         Scl_63.r        =   47
         Scl_63.b        =   486
         C[0]_63         =   12632256
         C[1]_63         =   12632256
         Image_63        =   64
         ClassName_64    =   "CCWPictImage"
         opts_64         =   1280
         Rows_64         =   1
         Cols_64         =   1
         Pict_64         =   7
         F_64            =   12632256
         B_64            =   12632256
         ColorReplaceWith_64=   8421504
         ColorReplace_64 =   8421504
         Tolerance_64    =   2
         Animator_63     =   0
         Blinker_63      =   0
         list[3]_58      =   65
         ClassName_65    =   "CCWDrawObj"
         opts_65         =   54
         SclRef_65.l     =   2
         SclRef_65.t     =   2
         SclRef_65.r     =   132
         SclRef_65.b     =   498
         Scl_65.l        =   15
         Scl_65.t        =   488
         Scl_65.r        =   52
         Scl_65.b        =   503
         C[0]_65         =   16776960
         C[1]_65         =   12632256
         Image_65        =   66
         ClassName_66    =   "CCWPictImage"
         opts_66         =   1280
         Rows_66         =   1
         Cols_66         =   1
         Pict_66         =   96
         F_66            =   16776960
         B_66            =   12632256
         ColorReplaceWith_66=   8421504
         ColorReplace_66 =   8421504
         Tolerance_66    =   2
         Animator_65     =   0
         Blinker_65      =   0
         list[2]_58      =   67
         ClassName_67    =   "CCWDrawObj"
         opts_67         =   54
         SclRef_67.l     =   2
         SclRef_67.t     =   2
         SclRef_67.r     =   132
         SclRef_67.b     =   498
         Scl_67.l        =   14
         Scl_67.t        =   6
         Scl_67.r        =   50
         Scl_67.b        =   21
         C[0]_67         =   16711935
         C[1]_67         =   12632256
         Image_67        =   68
         ClassName_68    =   "CCWPictImage"
         opts_68         =   1280
         Rows_68         =   1
         Cols_68         =   1
         Pict_68         =   95
         F_68            =   16711935
         B_68            =   12632256
         ColorReplaceWith_68=   8421504
         ColorReplace_68 =   8421504
         Tolerance_68    =   2
         Animator_67     =   0
         Blinker_67      =   0
         list[1]_58      =   2
         Ptrs_1          =   69
         ClassName_69    =   "CCWPointerArray"
         Array_69        =   1
         Editor_69       =   70
         ClassName_70    =   "CCWPointerArrayEditor"
         Owner_70        =   1
         Array[0]_69     =   59
         Bindings_1      =   71
         ClassName_71    =   "CCWBindingHolderArray"
         Editor_71       =   72
         ClassName_72    =   "CCWBindingHolderArrayEditor"
         Owner_72        =   1
         Stats_1         =   73
         ClassName_73    =   "CCWStats"
         doInc_1         =   67
         doDec_1         =   65
         doFrame_1       =   63
      End
      Begin VB.Label LabelBNCframe
         Appearance      =   0  'Flat
         BackColor       =   &H00FFE0E0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   7575
         Index           =   0
         Left            =   0
         TabIndex        =   120
         Top             =   0
         Width           =   1905
      End
   End
   Begin VB.Frame FrameBNC1
      BorderStyle     =   0  'None
      Height          =   7575
      Left            =   8100
      TabIndex        =   122
      Top             =   1125
      Visible         =   0   'False
      Width           =   1930
      Begin CWUIControlsLib.CWSlide cwBNC1select
         Height          =   7530
         Left            =   10
         TabIndex        =   123
         ToolTipText     =   "Changes BNC Diag.1 signal"
         Top             =   10
         Width           =   1875
         _Version        =   524288
         _ExtentX        =   3307
         _ExtentY        =   13282
         _StockProps     =   68
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Reset_0         =   0   'False
         CompatibleVers_0=   524288
         Slider_0        =   1
         ClassName_1     =   "CCWSlider"
         opts_1          =   2102
         SclRef_1.l      =   2
         SclRef_1.t      =   2
         SclRef_1.r      =   132
         SclRef_1.b      =   484
         Scl_1.l         =   2
         Scl_1.t         =   2
         Scl_1.r         =   132
         Scl_1.b         =   478
         C[0]_1          =   16777215
         BGImg_1         =   2
         ClassName_2     =   "CCWDrawObj"
         opts_2          =   54
         SclRef_2.l      =   2
         SclRef_2.t      =   2
         SclRef_2.r      =   132
         SclRef_2.b      =   502
         Scl_2.l         =   2
         Scl_2.t         =   2
         Scl_2.r         =   133
         Scl_2.b         =   502
         Image_2         =   3
         ClassName_3     =   "CCWPictImage"
         opts_3          =   1280
         Rows_3          =   1
         Cols_3          =   1
         Pict_3          =   286
         F_3             =   -2147483633
         B_3             =   -2147483633
         ColorReplaceWith_3=   8421504
         ColorReplace_3  =   8421504
         Tolerance_3     =   2
         Animator_2      =   0
         Blinker_2       =   0
         BFImg_1         =   4
         ClassName_4     =   "CCWDrawObj"
         opts_4          =   62
         Image_4         =   5
         ClassName_5     =   "CCWPictImage"
         opts_5          =   1280
         Rows_5          =   1
         Cols_5          =   1
         Pict_5          =   286
         F_5             =   -2147483633
         B_5             =   -2147483633
         ColorReplaceWith_5=   8421504
         ColorReplace_5  =   8421504
         Tolerance_5     =   2
         Animator_4      =   0
         Blinker_4       =   0
         style_1         =   4
         Label_1         =   6
         ClassName_6     =   "CCWDrawObj"
         opts_6          =   54
         SclRef_6.l      =   2
         SclRef_6.t      =   2
         SclRef_6.r      =   132
         SclRef_6.b      =   504
         Scl_6.l         =   58
         Scl_6.t         =   2
         Scl_6.r         =   99
         Scl_6.b         =   18
         C[0]_6          =   -2147483640
         Image_6         =   7
         ClassName_7     =   "CCWTextImage"
         szText_7        =   "BNC_1"
         font_7          =   0
         Animator_6      =   0
         Blinker_6       =   0
         Border_1        =   8
         ClassName_8     =   "CCWDrawObj"
         opts_8          =   60
         Image_8         =   9
         ClassName_9     =   "CCWPictImage"
         opts_9          =   1280
         Rows_9          =   1
         Cols_9          =   1
         Pict_9          =   25
         F_9             =   -2147483633
         B_9             =   -2147483633
         ColorReplaceWith_9=   8421504
         ColorReplace_9  =   8421504
         Tolerance_9     =   2
         Animator_8      =   0
         Blinker_8       =   0
         FillBound_1     =   10
         ClassName_10    =   "CCWGuiObject"
         opts_10         =   60
         FillTok_1       =   11
         ClassName_11    =   "CCWGuiObject"
         opts_11         =   62
         Axis_1          =   12
         ClassName_12    =   "CCWAxis"
         opts_12         =   1079
         Name_12         =   "Axis"
         SclRef_12.l     =   2
         SclRef_12.t     =   2
         SclRef_12.r     =   132
         SclRef_12.b     =   498
         Scl_12.l        =   2
         Scl_12.t        =   27
         Scl_12.r        =   123
         Scl_12.b        =   481
         Orientation_12  =   133172
         format_12       =   13
         ClassName_13    =   "CCWFormat"
         Scale_12        =   14
         ClassName_14    =   "CCWScale"
         opts_14         =   57344
         rMin_14         =   25
         rMax_14         =   478
         dMin_14         =   1
         dMax_14         =   38
         discInterval_14 =   1
         Radial_12       =   0
         Enum_12         =   15
         ClassName_15    =   "CCWEnum"
         Array_15        =   38
         Editor_15       =   16
         ClassName_16    =   "CCWEnumArrayEditor"
         Owner_16        =   12
         Array[0]_15     =   17
         ClassName_17    =   "CCWEnumElt"
         opts_17         =   1
         Name_17         =   "Excitation"
         DrawList_17     =   0
         varVarType_17   =   2
         var_Val_17      =   120
         Array[1]_15     =   18
         ClassName_18    =   "CCWEnumElt"
         opts_18         =   1
         Name_18         =   "Pos_tY"
         DrawList_18     =   0
         varVarType_18   =   2
         var_Val_18      =   59
         Array[2]_15     =   19
         ClassName_19    =   "CCWEnumElt"
         opts_19         =   1
         Name_19         =   "Pos_tX"
         DrawList_19     =   0
         varVarType_19   =   2
         var_Val_19      =   58
         Array[3]_15     =   20
         ClassName_20    =   "CCWEnumElt"
         opts_20         =   1
         Name_20         =   "Pos_Z"
         DrawList_20     =   0
         varVarType_20   =   2
         var_Val_20      =   57
         Array[4]_15     =   21
         ClassName_21    =   "CCWEnumElt"
         opts_21         =   1
         Name_21         =   "Pos_tZ"
         DrawList_21     =   0
         varVarType_21   =   2
         var_Val_21      =   56
         Array[5]_15     =   22
         ClassName_22    =   "CCWEnumElt"
         opts_22         =   1
         Name_22         =   "Pos_Y"
         DrawList_22     =   0
         varVarType_22   =   2
         var_Val_22      =   55
         Array[6]_15     =   23
         ClassName_23    =   "CCWEnumElt"
         opts_23         =   1
         Name_23         =   "Pos_X"
         DrawList_23     =   0
         varVarType_23   =   2
         var_Val_23      =   54
         Array[7]_15     =   24
         ClassName_24    =   "CCWEnumElt"
         opts_24         =   1
         Name_24         =   "Vel_tY"
         DrawList_24     =   0
         varVarType_24   =   2
         var_Val_24      =   53
         Array[8]_15     =   25
         ClassName_25    =   "CCWEnumElt"
         opts_25         =   1
         Name_25         =   "Vel_tX"
         DrawList_25     =   0
         varVarType_25   =   2
         var_Val_25      =   52
         Array[9]_15     =   26
         ClassName_26    =   "CCWEnumElt"
         opts_26         =   1
         Name_26         =   "Vel_Z"
         DrawList_26     =   0
         varVarType_26   =   2
         var_Val_26      =   51
         Array[10]_15    =   27
         ClassName_27    =   "CCWEnumElt"
         opts_27         =   1
         Name_27         =   "Vel_tZ"
         DrawList_27     =   0
         varVarType_27   =   2
         var_Val_27      =   50
         Array[11]_15    =   28
         ClassName_28    =   "CCWEnumElt"
         opts_28         =   1
         Name_28         =   "Vel_Y"
         DrawList_28     =   0
         varVarType_28   =   2
         var_Val_28      =   49
         Array[12]_15    =   29
         ClassName_29    =   "CCWEnumElt"
         opts_29         =   1
         Name_29         =   "Vel_X"
         DrawList_29     =   0
         varVarType_29   =   2
         var_Val_29      =   48
         Array[13]_15    =   30
         ClassName_30    =   "CCWEnumElt"
         opts_30         =   1
         Name_30         =   "Motor 4Z"
         DrawList_30     =   0
         varVarType_30   =   2
         var_Val_30      =   39
         Array[14]_15    =   31
         ClassName_31    =   "CCWEnumElt"
         opts_31         =   1
         Name_31         =   "Motor 3Z"
         DrawList_31     =   0
         varVarType_31   =   2
         var_Val_31      =   38
         Array[15]_15    =   32
         ClassName_32    =   "CCWEnumElt"
         opts_32         =   1
         Name_32         =   "Motor 2Z"
         DrawList_32     =   0
         varVarType_32   =   2
         var_Val_32      =   37
         Array[16]_15    =   33
         ClassName_33    =   "CCWEnumElt"
         opts_33         =   1
         Name_33         =   "Motor 1Z"
         DrawList_33     =   0
         varVarType_33   =   2
         var_Val_33      =   36
         Array[17]_15    =   34
         ClassName_34    =   "CCWEnumElt"
         opts_34         =   1
         Name_34         =   "Motor 4H"
         DrawList_34     =   0
         varVarType_34   =   2
         var_Val_34      =   43
         Array[18]_15    =   35
         ClassName_35    =   "CCWEnumElt"
         opts_35         =   1
         Name_35         =   "Motor 3H"
         DrawList_35     =   0
         varVarType_35   =   2
         var_Val_35      =   42
         Array[19]_15    =   36
         ClassName_36    =   "CCWEnumElt"
         opts_36         =   1
         Name_36         =   "Motor 2H"
         DrawList_36     =   0
         varVarType_36   =   2
         var_Val_36      =   41
         Array[20]_15    =   37
         ClassName_37    =   "CCWEnumElt"
         opts_37         =   1
         Name_37         =   "Motor 1H"
         DrawList_37     =   0
         varVarType_37   =   2
         var_Val_37      =   40
         Array[21]_15    =   38
         ClassName_38    =   "CCWEnumElt"
         opts_38         =   1
         Name_38         =   "Geo 3Z "
         DrawList_38     =   0
         varVarType_38   =   2
         var_Val_38      =   29
         Array[22]_15    =   39
         ClassName_39    =   "CCWEnumElt"
         opts_39         =   1
         Name_39         =   "Geo 2Z "
         DrawList_39     =   0
         varVarType_39   =   2
         var_Val_39      =   28
         Array[23]_15    =   40
         ClassName_40    =   "CCWEnumElt"
         opts_40         =   1
         Name_40         =   "Geo 1Z "
         DrawList_40     =   0
         varVarType_40   =   2
         var_Val_40      =   27
         Array[24]_15    =   41
         ClassName_41    =   "CCWEnumElt"
         opts_41         =   1
         Name_41         =   "Geo 3H"
         DrawList_41     =   0
         varVarType_41   =   2
         var_Val_41      =   26
         Array[25]_15    =   42
         ClassName_42    =   "CCWEnumElt"
         opts_42         =   1
         Name_42         =   "Geo 2H "
         DrawList_42     =   0
         varVarType_42   =   2
         var_Val_42      =   25
         Array[26]_15    =   43
         ClassName_43    =   "CCWEnumElt"
         opts_43         =   1
         Name_43         =   "Geo 1H "
         DrawList_43     =   0
         varVarType_43   =   2
         var_Val_43      =   24
         Array[27]_15    =   44
         ClassName_44    =   "CCWEnumElt"
         opts_44         =   1
         Name_44         =   "StFF_Ypos"
         DrawList_44     =   0
         varVarType_44   =   2
         var_Val_44      =   17
         Array[28]_15    =   45
         ClassName_45    =   "CCWEnumElt"
         opts_45         =   1
         Name_45         =   "StFF_Xpos "
         DrawList_45     =   0
         varVarType_45   =   2
         var_Val_45      =   16
         Array[29]_15    =   46
         ClassName_46    =   "CCWEnumElt"
         opts_46         =   1
         Name_46         =   "StFF_Yacc "
         DrawList_46     =   0
         varVarType_46   =   2
         var_Val_46      =   19
         Array[30]_15    =   47
         ClassName_47    =   "CCWEnumElt"
         opts_47         =   1
         Name_47         =   "StFF_Xacc "
         DrawList_47     =   0
         varVarType_47   =   2
         var_Val_47      =   18
         Array[31]_15    =   48
         ClassName_48    =   "CCWEnumElt"
         opts_48         =   1
         Name_48         =   "Prx_3Z_um"
         DrawList_48     =   0
         varVarType_48   =   2
         var_Val_48      =   6
         Array[32]_15    =   49
         ClassName_49    =   "CCWEnumElt"
         opts_49         =   1
         Name_49         =   "Prx_2Z_um"
         DrawList_49     =   0
         varVarType_49   =   2
         var_Val_49      =   5
         Array[33]_15    =   50
         ClassName_50    =   "CCWEnumElt"
         opts_50         =   1
         Name_50         =   "Prx_1Z_um"
         DrawList_50     =   0
         varVarType_50   =   2
         var_Val_50      =   4
         Array[34]_15    =   51
         ClassName_51    =   "CCWEnumElt"
         opts_51         =   1
         Name_51         =   "Prx_3H_um"
         DrawList_51     =   0
         varVarType_51   =   2
         var_Val_51      =   3
         Array[35]_15    =   52
         ClassName_52    =   "CCWEnumElt"
         opts_52         =   1
         Name_52         =   "Prx_2H_um"
         DrawList_52     =   0
         varVarType_52   =   2
         var_Val_52      =   2
         Array[36]_15    =   53
         ClassName_53    =   "CCWEnumElt"
         opts_53         =   1
         Name_53         =   "Prx_1H_um"
         DrawList_53     =   0
         varVarType_53   =   2
         var_Val_53      =   1
         Array[37]_15    =   54
         ClassName_54    =   "CCWEnumElt"
         opts_54         =   1
         Name_54         =   "--other--"
         DrawList_54     =   0
         varVarType_54   =   2
         var_Val_54      =   -1
         Font_12         =   55
         ClassName_55    =   "CCWFont"
         bFont_55        =   -1  'True
         BeginProperty Font_55 {0BE35203-8F91-11CE-9DE3-00AA004BB851}
            Name            =   "Arial"
            Size            =   9
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         tickopts_12     =   2962
         Caption_12      =   56
         ClassName_56    =   "CCWDrawObj"
         opts_56         =   62
         C[0]_56         =   -2147483640
         Image_56        =   57
         ClassName_57    =   "CCWTextImage"
         font_57         =   0
         Animator_56     =   0
         Blinker_56      =   0
         DrawLst_1       =   58
         ClassName_58    =   "CDrawList"
         count_58        =   10
         list[10]_58     =   8
         list[9]_58      =   59
         ClassName_59    =   "CCWThumb"
         opts_59         =   1048631
         Name_59         =   "Pointer-1"
         SclRef_59.l     =   2
         SclRef_59.t     =   17
         SclRef_59.r     =   132
         SclRef_59.b     =   492
         Scl_59.l        =   11
         Scl_59.t        =   111
         Scl_59.r        =   57
         Scl_59.b        =   124
         C[0]_59         =   255
         C[2]_59         =   -2147483635
         Image_59        =   60
         ClassName_60    =   "CCWPictImage"
         opts_60         =   1280
         Rows_60         =   1
         Cols_60         =   1
         Pict_60         =   94
         F_60            =   255
         B_60            =   -2147483633
         ColorReplaceWith_60=   8421504
         ColorReplace_60 =   8421504
         Tolerance_60    =   2
         Animator_59     =   0
         Blinker_59      =   0
         style_59        =   3
         Value_59        =   4
         Fill_59         =   61
         ClassName_61    =   "CCWDrawObj"
         opts_61         =   62
         Image_61        =   62
         ClassName_62    =   "CCWPictImage"
         opts_62         =   1280
         Rows_62         =   1
         Cols_62         =   1
         Pict_62         =   286
         F_62            =   -2147483633
         B_62            =   -2147483633
         ColorReplaceWith_62=   8421504
         ColorReplace_62 =   8421504
         Tolerance_62    =   2
         Animator_61     =   0
         Blinker_61      =   0
         list[8]_58      =   12
         list[7]_58      =   6
         list[6]_58      =   11
         list[5]_58      =   4
         list[4]_58      =   63
         ClassName_63    =   "CCWDrawObj"
         opts_63         =   54
         SclRef_63.l     =   2
         SclRef_63.t     =   2
         SclRef_63.r     =   132
         SclRef_63.b     =   498
         Scl_63.l        =   20
         Scl_63.t        =   22
         Scl_63.r        =   47
         Scl_63.b        =   486
         C[0]_63         =   12632256
         C[1]_63         =   12632256
         Image_63        =   64
         ClassName_64    =   "CCWPictImage"
         opts_64         =   1280
         Rows_64         =   1
         Cols_64         =   1
         Pict_64         =   7
         F_64            =   12632256
         B_64            =   12632256
         ColorReplaceWith_64=   8421504
         ColorReplace_64 =   8421504
         Tolerance_64    =   2
         Animator_63     =   0
         Blinker_63      =   0
         list[3]_58      =   65
         ClassName_65    =   "CCWDrawObj"
         opts_65         =   54
         SclRef_65.l     =   2
         SclRef_65.t     =   2
         SclRef_65.r     =   132
         SclRef_65.b     =   498
         Scl_65.l        =   15
         Scl_65.t        =   488
         Scl_65.r        =   52
         Scl_65.b        =   503
         C[0]_65         =   16776960
         C[1]_65         =   12632256
         Image_65        =   66
         ClassName_66    =   "CCWPictImage"
         opts_66         =   1280
         Rows_66         =   1
         Cols_66         =   1
         Pict_66         =   96
         F_66            =   16776960
         B_66            =   12632256
         ColorReplaceWith_66=   8421504
         ColorReplace_66 =   8421504
         Tolerance_66    =   2
         Animator_65     =   0
         Blinker_65      =   0
         list[2]_58      =   67
         ClassName_67    =   "CCWDrawObj"
         opts_67         =   54
         SclRef_67.l     =   2
         SclRef_67.t     =   2
         SclRef_67.r     =   132
         SclRef_67.b     =   498
         Scl_67.l        =   14
         Scl_67.t        =   6
         Scl_67.r        =   50
         Scl_67.b        =   21
         C[0]_67         =   16711935
         C[1]_67         =   12632256
         Image_67        =   68
         ClassName_68    =   "CCWPictImage"
         opts_68         =   1280
         Rows_68         =   1
         Cols_68         =   1
         Pict_68         =   95
         F_68            =   16711935
         B_68            =   12632256
         ColorReplaceWith_68=   8421504
         ColorReplace_68 =   8421504
         Tolerance_68    =   2
         Animator_67     =   0
         Blinker_67      =   0
         list[1]_58      =   2
         Ptrs_1          =   69
         ClassName_69    =   "CCWPointerArray"
         Array_69        =   1
         Editor_69       =   70
         ClassName_70    =   "CCWPointerArrayEditor"
         Owner_70        =   1
         Array[0]_69     =   59
         Bindings_1      =   71
         ClassName_71    =   "CCWBindingHolderArray"
         Editor_71       =   72
         ClassName_72    =   "CCWBindingHolderArrayEditor"
         Owner_72        =   1
         Stats_1         =   73
         ClassName_73    =   "CCWStats"
         doInc_1         =   67
         doDec_1         =   65
         doFrame_1       =   63
      End
      Begin VB.Label LabelBNCframe
         Appearance      =   0  'Flat
         BackColor       =   &H00FFE0E0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   7575
         Index           =   1
         Left            =   0
         TabIndex        =   124
         Top             =   0
         Width           =   1905
      End
   End
   Begin VB.Frame FrameSystem
      BackColor       =   &H00C0FFFF&
      Caption         =   "System Control"
      BeginProperty Font
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5970
      Left            =   2745
      TabIndex        =   3
      Top             =   1661
      Width           =   11756
      Begin VB.CheckBox ChkAirIsolatorsPresent
         BackColor       =   &H000080FF&
         Caption         =   "NO Air Isolators"
         BeginProperty Font
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   810
         Style           =   1  'Graphical
         TabIndex        =   312
         ToolTipText     =   "Manual override of auto-detected air isolators presence "
         Top             =   765
         Value           =   1  'Checked
         Width           =   1410
      End
      Begin VB.Frame FramePneumFB
         BackColor       =   &H00E0E0E0&
         Height          =   4635
         Left            =   7290
         TabIndex        =   276
         Top             =   1170
         Width           =   4350
         Begin VB.CommandButton CmdSetPressurePoints
            BackColor       =   &H0000FFFF&
            Caption         =   "Update Setpoints"
            Height          =   330
            Left            =   2790
            Style           =   1  'Graphical
            TabIndex        =   277
            Tag             =   "piso>set"
            ToolTipText     =   "Update isolaotrs' pressure setpoints when payload floats normally. Umportant for proper dock-float."
            Top             =   1440
            Width           =   1455
         End
         Begin CWUIControlsLib.CWNumEdit cwNumFBgain
            Height          =   360
            Index           =   3
            Left            =   630
            TabIndex        =   278
            Tag             =   "gain03"
            ToolTipText     =   "Gain of the axis"
            Top             =   3340
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
            C[0]_1          =   16777215
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
            Value_Val_1     =   0.71
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
         Begin CWUIControlsLib.CWButton cwBut_AuxDamping
            Height          =   450
            Left            =   45
            TabIndex        =   279
            Tag             =   "loop_FBD"
            ToolTipText     =   "Auxilary pneumatic damping (might increase vert noise if gain is too high)"
            Top             =   2475
            Width           =   2895
            _Version        =   524288
            _ExtentX        =   5106
            _ExtentY        =   785
            _StockProps     =   68
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}
               Name            =   "MS Sans Serif"
               Size            =   8.16
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Reset_0         =   0   'False
            CompatibleVers_0=   524288
            Boolean_0       =   1
            ClassName_1     =   "CCWBoolean"
            opts_1          =   2614
            SclRef_1.l      =   2
            SclRef_1.t      =   2
            SclRef_1.r      =   195
            SclRef_1.b      =   31
            Scl_1.l         =   6
            Scl_1.t         =   4
            Scl_1.r         =   87
            Scl_1.b         =   28
            C[0]_1          =   16777152
            Enum_1          =   2
            ClassName_2     =   "CCWEnum"
            Array_2         =   2
            Editor_2        =   0
            Array[0]_2      =   3
            ClassName_3     =   "CCWEnumElt"
            opts_3          =   1
            Name_3          =   "Off"
            frame_3         =   286
            DrawList_3      =   4
            ClassName_4     =   "CDrawList"
            count_4         =   4
            list[4]_4       =   5
            ClassName_5     =   "CCWDrawObj"
            opts_5          =   54
            SclRef_5.l      =   6
            SclRef_5.t      =   4
            SclRef_5.r      =   87
            SclRef_5.b      =   28
            Scl_5.l         =   97
            Scl_5.t         =   8
            Scl_5.r         =   189
            Scl_5.b         =   26
            C[0]_5          =   -2147483640
            C[1]_5          =   -2147483640
            Image_5         =   6
            ClassName_6     =   "CCWTextImage"
            szText_6        =   "Aux Damping"
            font_6          =   7
            ClassName_7     =   "CCWFont"
            bFont_7         =   -1  'True
            BeginProperty Font_7 {0BE35203-8F91-11CE-9DE3-00AA004BB851}
               Name            =   "Arial"
               Size            =   11.13
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Animator_5      =   0
            Blinker_5       =   0
            list[3]_4       =   8
            ClassName_8     =   "CCWDrawObj"
            opts_8          =   60
            C[0]_8          =   -2147483640
            C[1]_8          =   -2147483640
            Image_8         =   9
            ClassName_9     =   "CCWTextImage"
            font_9          =   0
            Animator_8      =   0
            Blinker_8       =   0
            list[2]_4       =   10
            ClassName_10    =   "CCWDrawObj"
            opts_10         =   60
            C[0]_10         =   -2147483640
            C[1]_10         =   -2147483640
            Image_10        =   11
            ClassName_11    =   "CCWTextImage"
            font_11         =   0
            Animator_10     =   0
            Blinker_10      =   0
            list[1]_4       =   12
            ClassName_12    =   "CCWDrawObj"
            opts_12         =   62
            C[0]_12         =   32768
            C[1]_12         =   65535
            Image_12        =   13
            ClassName_13    =   "CCWPictImage"
            opts_13         =   256
            Rows_13         =   1
            Cols_13         =   1
            URL_13          =   "\\tsclient\C\In_Out\Software\DC2020_WEB_Analyzer\OFF_small_switch.bmp"
            Pict_13         =   "Analyzer.frx":102B18
            ColorReplaceWith_13=   8421504
            ColorReplace_13 =   8421504
            Tolerance_13    =   2
            Animator_12     =   0
            Blinker_12      =   0
            varVarType_3    =   5
            Array[1]_2      =   14
            ClassName_14    =   "CCWEnumElt"
            opts_14         =   1
            Name_14         =   "On"
            frame_14        =   286
            DrawList_14     =   15
            ClassName_15    =   "CDrawList"
            count_15        =   4
            list[4]_15      =   16
            ClassName_16    =   "CCWDrawObj"
            opts_16         =   54
            SclRef_16.l     =   6
            SclRef_16.t     =   4
            SclRef_16.r     =   87
            SclRef_16.b     =   28
            Scl_16.l        =   97
            Scl_16.t        =   8
            Scl_16.r        =   189
            Scl_16.b        =   26
            C[0]_16         =   -2147483640
            C[1]_16         =   -2147483640
            Image_16        =   6
            Animator_16     =   0
            Blinker_16      =   0
            list[3]_15      =   17
            ClassName_17    =   "CCWDrawObj"
            opts_17         =   60
            C[0]_17         =   -2147483640
            C[1]_17         =   -2147483640
            Image_17        =   9
            Animator_17     =   0
            Blinker_17      =   0
            list[2]_15      =   18
            ClassName_18    =   "CCWDrawObj"
            opts_18         =   60
            C[0]_18         =   -2147483640
            C[1]_18         =   -2147483640
            Image_18        =   11
            Animator_18     =   0
            Blinker_18      =   0
            list[1]_15      =   19
            ClassName_19    =   "CCWDrawObj"
            opts_19         =   62
            C[0]_19         =   32768
            C[1]_19         =   32768
            Image_19        =   20
            ClassName_20    =   "CCWPictImage"
            opts_20         =   256
            Rows_20         =   1
            Cols_20         =   1
            URL_20          =   "\\tsclient\C\In_Out\Software\DC2020_WEB_Analyzer\ON_small_switch.bmp"
            Pict_20         =   "Analyzer.frx":104EEA
            ColorReplaceWith_20=   8421504
            ColorReplace_20 =   8421504
            Tolerance_20    =   2
            Animator_19     =   0
            Blinker_19      =   0
            varVarType_14   =   5
            Bindings_1      =   21
            ClassName_21    =   "CCWBindingHolderArray"
            Editor_21       =   22
            ClassName_22    =   "CCWBindingHolderArrayEditor"
            Owner_22        =   1
            Style_1         =   9
            mechAction_1    =   1
            BGImg_1         =   23
            ClassName_23    =   "CCWDrawObj"
            opts_23         =   62
            Image_23        =   24
            ClassName_24    =   "CCWPictImage"
            opts_24         =   1280
            Rows_24         =   1
            Cols_24         =   1
            Pict_24         =   286
            F_24            =   -2147483633
            B_24            =   -2147483633
            ColorReplaceWith_24=   8421504
            ColorReplace_24 =   8421504
            Tolerance_24    =   2
            Animator_23     =   0
            Blinker_23      =   0
            Array_1         =   6
            Editor_1        =   0
            Array[0]_1      =   12
            Array[1]_1      =   19
            Array[2]_1      =   0
            Array[3]_1      =   0
            Array[4]_1      =   25
            ClassName_25    =   "CCWDrawObj"
            opts_25         =   62
            Image_25        =   9
            Animator_25     =   0
            Blinker_25      =   0
            Array[5]_1      =   26
            ClassName_26    =   "CCWDrawObj"
            opts_26         =   62
            Image_26        =   11
            Animator_26     =   0
            Blinker_26      =   0
            Label_1         =   27
            ClassName_27    =   "CCWDrawObj"
            opts_27         =   62
            C[0]_27         =   -2147483640
            Image_27        =   6
            Animator_27     =   0
            Blinker_27      =   0
         End
         Begin CWUIControlsLib.CWNumEdit cwNumFBgain
            Height          =   360
            Index           =   4
            Left            =   630
            TabIndex        =   280
            Tag             =   "gain04"
            ToolTipText     =   "Gain of the axis"
            Top             =   3760
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
            C[0]_1          =   16777215
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
            Value_Val_1     =   0.71
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
         Begin CWUIControlsLib.CWNumEdit cwNumFBgain
            Height          =   360
            Index           =   5
            Left            =   630
            TabIndex        =   281
            Tag             =   "gain05"
            ToolTipText     =   "Gain of the axis"
            Top             =   4185
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
            C[0]_1          =   16777215
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
            Value_Val_1     =   0.71
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
         Begin CWUIControlsLib.CWNumEdit cwNumFBgain
            Height          =   360
            Index           =   12
            Left            =   1845
            TabIndex        =   282
            Tag             =   "gain0C"
            ToolTipText     =   "Auxiliary Damping Gain of the axis"
            Top             =   3340
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
         Begin CWUIControlsLib.CWNumEdit cwNumFBgain
            Height          =   360
            Index           =   13
            Left            =   1845
            TabIndex        =   283
            Tag             =   "gain0D"
            ToolTipText     =   "Auxiliary Damping Gain of the axis"
            Top             =   3760
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
            Value_Val_1     =   2.22
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
         Begin CWUIControlsLib.CWNumEdit cwNumFBgain
            Height          =   360
            Index           =   14
            Left            =   1845
            TabIndex        =   284
            Tag             =   "gain0E"
            ToolTipText     =   "Auxiliary Damping Gain of the axis"
            Top             =   4185
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
            Value_Val_1     =   3.33
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
         Begin CWUIControlsLib.CWButton cw_Status_Bit_Indicator
            Height          =   420
            Index           =   0
            Left            =   2250
            TabIndex        =   285
            ToolTipText     =   "Shows vertical status of the payload"
            Top             =   900
            Width           =   930
            _Version        =   524288
            _ExtentX        =   1640
            _ExtentY        =   741
            _StockProps     =   68
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Reset_0         =   0   'False
            CompatibleVers_0=   524288
            Boolean_0       =   1
            ClassName_1     =   "CCWBoolean"
            opts_1          =   2614
            SclRef_1.l      =   2
            SclRef_1.t      =   2
            SclRef_1.r      =   64
            SclRef_1.b      =   30
            Scl_1.l         =   8
            Scl_1.t         =   2
            Scl_1.r         =   64
            Scl_1.b         =   30
            C[0]_1          =   16777152
            Enum_1          =   2
            ClassName_2     =   "CCWEnum"
            Array_2         =   2
            Editor_2        =   0
            Array[0]_2      =   3
            ClassName_3     =   "CCWEnumElt"
            opts_3          =   1
            Name_3          =   "Off"
            frame_3         =   286
            DrawList_3      =   4
            ClassName_4     =   "CDrawList"
            count_4         =   4
            list[4]_4       =   5
            ClassName_5     =   "CCWDrawObj"
            opts_5          =   62
            C[0]_5          =   -2147483640
            C[1]_5          =   -2147483640
            Image_5         =   6
            ClassName_6     =   "CCWTextImage"
            font_6          =   0
            Animator_5      =   0
            Blinker_5       =   0
            list[3]_4       =   7
            ClassName_7     =   "CCWDrawObj"
            opts_7          =   60
            C[0]_7          =   -2147483640
            C[1]_7          =   -2147483640
            Image_7         =   8
            ClassName_8     =   "CCWTextImage"
            szText_8        =   "DOCKED"
            font_8          =   0
            Animator_7      =   0
            Blinker_7       =   0
            list[2]_4       =   9
            ClassName_9     =   "CCWDrawObj"
            opts_9          =   62
            C[0]_9          =   -2147483640
            C[1]_9          =   -2147483640
            Image_9         =   10
            ClassName_10    =   "CCWTextImage"
            szText_10       =   "Floating"
            font_10         =   0
            Animator_9      =   0
            Blinker_9       =   0
            list[1]_4       =   11
            ClassName_11    =   "CCWDrawObj"
            opts_11         =   62
            C[0]_11         =   65280
            C[1]_11         =   65280
            Image_11        =   12
            ClassName_12    =   "CCWPictImage"
            opts_12         =   1280
            Rows_12         =   1
            Cols_12         =   1
            Pict_12         =   73
            F_12            =   65280
            B_12            =   65280
            ColorReplaceWith_12=   8421504
            ColorReplace_12 =   8421504
            Tolerance_12    =   2
            Animator_11     =   0
            Blinker_11      =   0
            varVarType_3    =   5
            Array[1]_2      =   13
            ClassName_13    =   "CCWEnumElt"
            opts_13         =   1
            Name_13         =   "On"
            frame_13        =   286
            DrawList_13     =   14
            ClassName_14    =   "CDrawList"
            count_14        =   4
            list[4]_14      =   15
            ClassName_15    =   "CCWDrawObj"
            opts_15         =   62
            C[0]_15         =   -2147483640
            C[1]_15         =   -2147483640
            Image_15        =   6
            Animator_15     =   0
            Blinker_15      =   0
            list[3]_14      =   16
            ClassName_16    =   "CCWDrawObj"
            opts_16         =   62
            C[0]_16         =   -2147483640
            C[1]_16         =   -2147483640
            Image_16        =   8
            Animator_16     =   0
            Blinker_16      =   0
            list[2]_14      =   17
            ClassName_17    =   "CCWDrawObj"
            opts_17         =   60
            C[0]_17         =   -2147483640
            C[1]_17         =   -2147483640
            Image_17        =   10
            Animator_17     =   0
            Blinker_17      =   0
            list[1]_14      =   18
            ClassName_18    =   "CCWDrawObj"
            opts_18         =   62
            C[0]_18         =   255
            C[1]_18         =   255
            Image_18        =   19
            ClassName_19    =   "CCWPictImage"
            opts_19         =   1280
            Rows_19         =   1
            Cols_19         =   1
            Pict_19         =   74
            F_19            =   255
            B_19            =   255
            ColorReplaceWith_19=   8421504
            ColorReplace_19 =   8421504
            Tolerance_19    =   2
            Animator_18     =   0
            Blinker_18      =   20
            ClassName_20    =   "CCWBlinker"
            Interval_20     =   600
            EventID_20      =   7
            Obj_20          =   18
            varVarType_13   =   5
            Bindings_1      =   21
            ClassName_21    =   "CCWBindingHolderArray"
            Editor_21       =   22
            ClassName_22    =   "CCWBindingHolderArrayEditor"
            Owner_22        =   1
            Style_1         =   4
            mechAction_1    =   3
            BGImg_1         =   23
            ClassName_23    =   "CCWDrawObj"
            opts_23         =   62
            Image_23        =   24
            ClassName_24    =   "CCWPictImage"
            opts_24         =   1280
            Rows_24         =   1
            Cols_24         =   1
            Pict_24         =   286
            F_24            =   -2147483633
            B_24            =   -2147483633
            ColorReplaceWith_24=   8421504
            ColorReplace_24 =   8421504
            Tolerance_24    =   2
            Animator_23     =   0
            Blinker_23      =   0
            Array_1         =   6
            Editor_1        =   0
            Array[0]_1      =   11
            Array[1]_1      =   18
            Array[2]_1      =   0
            Array[3]_1      =   0
            Array[4]_1      =   25
            ClassName_25    =   "CCWDrawObj"
            opts_25         =   62
            Image_25        =   8
            Animator_25     =   0
            Blinker_25      =   0
            Array[5]_1      =   26
            ClassName_26    =   "CCWDrawObj"
            opts_26         =   62
            Image_26        =   10
            Animator_26     =   0
            Blinker_26      =   0
            Label_1         =   27
            ClassName_27    =   "CCWDrawObj"
            opts_27         =   62
            C[0]_27         =   -2147483640
            Image_27        =   6
            Animator_27     =   0
            Blinker_27      =   0
         End
         Begin CWUIControlsLib.CWButton cw_Status_Bit_Indicator
            Height          =   420
            Index           =   4
            Left            =   2250
            TabIndex        =   286
            ToolTipText     =   "Shows the dither status of the main valves and optional drift control valves"
            Top             =   405
            Width           =   930
            _Version        =   524288
            _ExtentX        =   1640
            _ExtentY        =   741
            _StockProps     =   68
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Reset_0         =   0   'False
            CompatibleVers_0=   524288
            Boolean_0       =   1
            ClassName_1     =   "CCWBoolean"
            opts_1          =   2614
            SclRef_1.l      =   2
            SclRef_1.t      =   2
            SclRef_1.r      =   64
            SclRef_1.b      =   30
            Scl_1.l         =   8
            Scl_1.t         =   2
            Scl_1.r         =   64
            Scl_1.b         =   30
            C[0]_1          =   16761024
            Enum_1          =   2
            ClassName_2     =   "CCWEnum"
            Array_2         =   2
            Editor_2        =   0
            Array[0]_2      =   3
            ClassName_3     =   "CCWEnumElt"
            opts_3          =   1
            Name_3          =   "Off"
            frame_3         =   286
            DrawList_3      =   4
            ClassName_4     =   "CDrawList"
            count_4         =   4
            list[4]_4       =   5
            ClassName_5     =   "CCWDrawObj"
            opts_5          =   54
            SclRef_5.l      =   45
            SclRef_5.t      =   3
            SclRef_5.r      =   91
            SclRef_5.b      =   30
            Scl_5.l         =   5
            Scl_5.t         =   10
            Scl_5.r         =   40
            Scl_5.b         =   24
            C[0]_5          =   -2147483640
            C[1]_5          =   -2147483640
            Image_5         =   6
            ClassName_6     =   "CCWTextImage"
            font_6          =   0
            Animator_5      =   0
            Blinker_5       =   0
            list[3]_4       =   7
            ClassName_7     =   "CCWDrawObj"
            opts_7          =   62
            C[0]_7          =   -2147483640
            C[1]_7          =   -2147483640
            Image_7         =   8
            ClassName_8     =   "CCWTextImage"
            szText_8        =   "ON"
            font_8          =   0
            Animator_7      =   0
            Blinker_7       =   0
            list[2]_4       =   9
            ClassName_9     =   "CCWDrawObj"
            opts_9          =   60
            C[0]_9          =   -2147483640
            C[1]_9          =   -2147483640
            Image_9         =   10
            ClassName_10    =   "CCWTextImage"
            szText_10       =   "Disabled"
            font_10         =   0
            Animator_9      =   0
            Blinker_9       =   0
            list[1]_4       =   11
            ClassName_11    =   "CCWDrawObj"
            opts_11         =   62
            C[0]_11         =   65280
            C[1]_11         =   65280
            Image_11        =   12
            ClassName_12    =   "CCWPictImage"
            opts_12         =   1280
            Rows_12         =   1
            Cols_12         =   1
            Pict_12         =   75
            F_12            =   65280
            B_12            =   65280
            ColorReplaceWith_12=   8421504
            ColorReplace_12 =   8421504
            Tolerance_12    =   2
            Animator_11     =   0
            Blinker_11      =   0
            varVarType_3    =   5
            Array[1]_2      =   13
            ClassName_13    =   "CCWEnumElt"
            opts_13         =   1
            Name_13         =   "On"
            frame_13        =   286
            DrawList_13     =   14
            ClassName_14    =   "CDrawList"
            count_14        =   4
            list[4]_14      =   15
            ClassName_15    =   "CCWDrawObj"
            opts_15         =   54
            SclRef_15.l     =   45
            SclRef_15.t     =   3
            SclRef_15.r     =   91
            SclRef_15.b     =   30
            Scl_15.l        =   5
            Scl_15.t        =   10
            Scl_15.r        =   40
            Scl_15.b        =   24
            C[0]_15         =   -2147483640
            C[1]_15         =   -2147483640
            Image_15        =   6
            Animator_15     =   0
            Blinker_15      =   0
            list[3]_14      =   16
            ClassName_16    =   "CCWDrawObj"
            opts_16         =   60
            C[0]_16         =   -2147483640
            C[1]_16         =   -2147483640
            Image_16        =   8
            Animator_16     =   0
            Blinker_16      =   0
            list[2]_14      =   17
            ClassName_17    =   "CCWDrawObj"
            opts_17         =   62
            C[0]_17         =   -2147483640
            C[1]_17         =   -2147483640
            Image_17        =   10
            Animator_17     =   0
            Blinker_17      =   0
            list[1]_14      =   18
            ClassName_18    =   "CCWDrawObj"
            opts_18         =   62
            C[0]_18         =   65535
            C[1]_18         =   65535
            Image_18        =   19
            ClassName_19    =   "CCWPictImage"
            opts_19         =   1280
            Rows_19         =   1
            Cols_19         =   1
            Pict_19         =   76
            F_19            =   65535
            B_19            =   65535
            ColorReplaceWith_19=   8421504
            ColorReplace_19 =   8421504
            Tolerance_19    =   2
            Animator_18     =   0
            Blinker_18      =   20
            ClassName_20    =   "CCWBlinker"
            Interval_20     =   600
            EventID_20      =   7
            Obj_20          =   18
            varVarType_13   =   5
            Bindings_1      =   21
            ClassName_21    =   "CCWBindingHolderArray"
            Editor_21       =   22
            ClassName_22    =   "CCWBindingHolderArrayEditor"
            Owner_22        =   1
            Style_1         =   5
            mechAction_1    =   3
            BGImg_1         =   23
            ClassName_23    =   "CCWDrawObj"
            opts_23         =   62
            Image_23        =   24
            ClassName_24    =   "CCWPictImage"
            opts_24         =   1280
            Rows_24         =   1
            Cols_24         =   1
            Pict_24         =   286
            F_24            =   -2147483633
            B_24            =   -2147483633
            ColorReplaceWith_24=   8421504
            ColorReplace_24 =   8421504
            Tolerance_24    =   2
            Animator_23     =   0
            Blinker_23      =   0
            Array_1         =   6
            Editor_1        =   0
            Array[0]_1      =   11
            Array[1]_1      =   18
            Array[2]_1      =   0
            Array[3]_1      =   0
            Array[4]_1      =   25
            ClassName_25    =   "CCWDrawObj"
            opts_25         =   62
            Image_25        =   8
            Animator_25     =   0
            Blinker_25      =   0
            Array[5]_1      =   26
            ClassName_26    =   "CCWDrawObj"
            opts_26         =   62
            Image_26        =   10
            Animator_26     =   0
            Blinker_26      =   0
            Label_1         =   27
            ClassName_27    =   "CCWDrawObj"
            opts_27         =   62
            C[0]_27         =   -2147483640
            Image_27        =   6
            Animator_27     =   0
            Blinker_27      =   0
         End
         Begin CWUIControlsLib.CWButton cwBut_BasicCMD
            Height          =   420
            Index           =   4
            Left            =   45
            TabIndex        =   287
            ToolTipText     =   "Controls the dither of the main valves and optional drift control valves"
            Top             =   405
            Width           =   2220
            _Version        =   524288
            _ExtentX        =   3916
            _ExtentY        =   741
            _StockProps     =   68
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Reset_0         =   0   'False
            CompatibleVers_0=   524288
            Boolean_0       =   1
            ClassName_1     =   "CCWBoolean"
            opts_1          =   2614
            SclRef_1.l      =   2
            SclRef_1.t      =   2
            SclRef_1.r      =   150
            SclRef_1.b      =   29
            Scl_1.l         =   4
            Scl_1.t         =   4
            Scl_1.r         =   68
            Scl_1.b         =   27
            C[0]_1          =   16761024
            Enum_1          =   2
            ClassName_2     =   "CCWEnum"
            Array_2         =   2
            Editor_2        =   0
            Array[0]_2      =   3
            ClassName_3     =   "CCWEnumElt"
            opts_3          =   1
            Name_3          =   "Off"
            frame_3         =   286
            DrawList_3      =   4
            ClassName_4     =   "CDrawList"
            count_4         =   4
            list[4]_4       =   5
            ClassName_5     =   "CCWDrawObj"
            opts_5          =   54
            SclRef_5.l      =   4
            SclRef_5.t      =   4
            SclRef_5.r      =   68
            SclRef_5.b      =   27
            Scl_5.l         =   70
            Scl_5.t         =   1
            Scl_5.r         =   148
            Scl_5.b         =   17
            C[0]_5          =   -2147483640
            C[1]_5          =   -2147483640
            Image_5         =   6
            ClassName_6     =   "CCWTextImage"
            szText_6        =   "Valve Dither"
            font_6          =   7
            ClassName_7     =   "CCWFont"
            bFont_7         =   -1  'True
            BeginProperty Font_7 {0BE35203-8F91-11CE-9DE3-00AA004BB851}
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Animator_5      =   0
            Blinker_5       =   0
            list[3]_4       =   8
            ClassName_8     =   "CCWDrawObj"
            opts_8          =   52
            SclRef_8.l      =   4
            SclRef_8.t      =   4
            SclRef_8.r      =   68
            SclRef_8.b      =   28
            Scl_8.l         =   73
            Scl_8.t         =   16
            Scl_8.r         =   144
            Scl_8.b         =   29
            C[0]_8          =   -2147483640
            C[1]_8          =   -2147483640
            Image_8         =   9
            ClassName_9     =   "CCWTextImage"
            szText_9        =   "Click to disable"
            font_9          =   0
            Animator_8      =   0
            Blinker_8       =   0
            list[2]_4       =   10
            ClassName_10    =   "CCWDrawObj"
            opts_10         =   54
            SclRef_10.l     =   4
            SclRef_10.t     =   4
            SclRef_10.r     =   68
            SclRef_10.b     =   28
            Scl_10.l        =   73
            Scl_10.t        =   16
            Scl_10.r        =   143
            Scl_10.b        =   29
            C[0]_10         =   -2147483640
            C[1]_10         =   -2147483640
            Image_10        =   11
            ClassName_11    =   "CCWTextImage"
            szText_11       =   "Click to enable"
            font_11         =   0
            Animator_10     =   0
            Blinker_10      =   0
            list[1]_4       =   12
            ClassName_12    =   "CCWDrawObj"
            opts_12         =   62
            C[0]_12         =   32768
            C[1]_12         =   65535
            Image_12        =   13
            ClassName_13    =   "CCWPictImage"
            opts_13         =   256
            Rows_13         =   1
            Cols_13         =   1
            URL_13          =   "\\tsclient\C\In_Out\Software\DC2020_WEB_Analyzer\OFF_small_switch.bmp"
            Pict_13         =   "Analyzer.frx":1072BC
            ColorReplaceWith_13=   8421504
            ColorReplace_13 =   8421504
            Tolerance_13    =   2
            Animator_12     =   0
            Blinker_12      =   0
            varVarType_3    =   5
            Array[1]_2      =   14
            ClassName_14    =   "CCWEnumElt"
            opts_14         =   1
            Name_14         =   "On"
            frame_14        =   286
            DrawList_14     =   15
            ClassName_15    =   "CDrawList"
            count_15        =   4
            list[4]_15      =   16
            ClassName_16    =   "CCWDrawObj"
            opts_16         =   54
            SclRef_16.l     =   4
            SclRef_16.t     =   4
            SclRef_16.r     =   68
            SclRef_16.b     =   27
            Scl_16.l        =   70
            Scl_16.t        =   1
            Scl_16.r        =   148
            Scl_16.b        =   17
            C[0]_16         =   -2147483640
            C[1]_16         =   -2147483640
            Image_16        =   6
            Animator_16     =   0
            Blinker_16      =   0
            list[3]_15      =   17
            ClassName_17    =   "CCWDrawObj"
            opts_17         =   54
            SclRef_17.l     =   4
            SclRef_17.t     =   4
            SclRef_17.r     =   68
            SclRef_17.b     =   28
            Scl_17.l        =   73
            Scl_17.t        =   16
            Scl_17.r        =   144
            Scl_17.b        =   29
            C[0]_17         =   -2147483640
            C[1]_17         =   -2147483640
            Image_17        =   9
            Animator_17     =   0
            Blinker_17      =   0
            list[2]_15      =   18
            ClassName_18    =   "CCWDrawObj"
            opts_18         =   52
            SclRef_18.l     =   4
            SclRef_18.t     =   4
            SclRef_18.r     =   68
            SclRef_18.b     =   28
            Scl_18.l        =   73
            Scl_18.t        =   16
            Scl_18.r        =   143
            Scl_18.b        =   29
            C[0]_18         =   -2147483640
            C[1]_18         =   -2147483640
            Image_18        =   11
            Animator_18     =   0
            Blinker_18      =   0
            list[1]_15      =   19
            ClassName_19    =   "CCWDrawObj"
            opts_19         =   62
            C[0]_19         =   32768
            C[1]_19         =   32768
            Image_19        =   20
            ClassName_20    =   "CCWPictImage"
            opts_20         =   256
            Rows_20         =   1
            Cols_20         =   1
            URL_20          =   "\\tsclient\C\In_Out\Software\DC2020_WEB_Analyzer\ON_small_switch.bmp"
            Pict_20         =   "Analyzer.frx":10968E
            ColorReplaceWith_20=   8421504
            ColorReplace_20 =   8421504
            Tolerance_20    =   2
            Animator_19     =   0
            Blinker_19      =   0
            varVarType_14   =   5
            Bindings_1      =   21
            ClassName_21    =   "CCWBindingHolderArray"
            Editor_21       =   22
            ClassName_22    =   "CCWBindingHolderArrayEditor"
            Owner_22        =   1
            Style_1         =   9
            drawState_1     =   1
            mechAction_1    =   1
            BGImg_1         =   23
            ClassName_23    =   "CCWDrawObj"
            opts_23         =   62
            Image_23        =   24
            ClassName_24    =   "CCWPictImage"
            opts_24         =   1280
            Rows_24         =   1
            Cols_24         =   1
            Pict_24         =   286
            F_24            =   -2147483633
            B_24            =   -2147483633
            ColorReplaceWith_24=   8421504
            ColorReplace_24 =   8421504
            Tolerance_24    =   2
            Animator_23     =   0
            Blinker_23      =   0
            Array_1         =   6
            Editor_1        =   0
            Array[0]_1      =   12
            Array[1]_1      =   19
            Array[2]_1      =   0
            Array[3]_1      =   0
            Array[4]_1      =   25
            ClassName_25    =   "CCWDrawObj"
            opts_25         =   62
            Image_25        =   9
            Animator_25     =   0
            Blinker_25      =   0
            Array[5]_1      =   26
            ClassName_26    =   "CCWDrawObj"
            opts_26         =   62
            Image_26        =   11
            Animator_26     =   0
            Blinker_26      =   0
            Label_1         =   27
            ClassName_27    =   "CCWDrawObj"
            opts_27         =   62
            C[0]_27         =   -2147483640
            Image_27        =   6
            Animator_27     =   0
            Blinker_27      =   0
         End
         Begin CWUIControlsLib.CWButton cwBut_BasicCMD
            Height          =   420
            Index           =   0
            Left            =   45
            TabIndex        =   288
            ToolTipText     =   "Controls vertical status of the payload"
            Top             =   900
            Width           =   2220
            _Version        =   524288
            _ExtentX        =   3916
            _ExtentY        =   741
            _StockProps     =   68
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Reset_0         =   0   'False
            CompatibleVers_0=   524288
            Boolean_0       =   1
            ClassName_1     =   "CCWBoolean"
            opts_1          =   2614
            SclRef_1.l      =   2
            SclRef_1.t      =   2
            SclRef_1.r      =   150
            SclRef_1.b      =   29
            Scl_1.l         =   5
            Scl_1.t         =   4
            Scl_1.r         =   68
            Scl_1.b         =   27
            C[0]_1          =   16777152
            Enum_1          =   2
            ClassName_2     =   "CCWEnum"
            Array_2         =   2
            Editor_2        =   0
            Array[0]_2      =   3
            ClassName_3     =   "CCWEnumElt"
            opts_3          =   1
            Name_3          =   "Off"
            frame_3         =   286
            DrawList_3      =   4
            ClassName_4     =   "CDrawList"
            count_4         =   4
            list[4]_4       =   5
            ClassName_5     =   "CCWDrawObj"
            opts_5          =   54
            SclRef_5.l      =   4
            SclRef_5.t      =   4
            SclRef_5.r      =   69
            SclRef_5.b      =   27
            Scl_5.l         =   76
            Scl_5.t         =   1
            Scl_5.r         =   145
            Scl_5.b         =   19
            C[0]_5          =   -2147483640
            C[1]_5          =   -2147483640
            Image_5         =   6
            ClassName_6     =   "CCWTextImage"
            szText_6        =   "Vert Dock"
            font_6          =   7
            ClassName_7     =   "CCWFont"
            bFont_7         =   -1  'True
            BeginProperty Font_7 {0BE35203-8F91-11CE-9DE3-00AA004BB851}
               Name            =   "Arial"
               Size            =   11.26
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Animator_5      =   0
            Blinker_5       =   0
            list[3]_4       =   8
            ClassName_8     =   "CCWDrawObj"
            opts_8          =   52
            SclRef_8.l      =   5
            SclRef_8.t      =   4
            SclRef_8.r      =   68
            SclRef_8.b      =   28
            Scl_8.l         =   81
            Scl_8.t         =   17
            Scl_8.r         =   139
            Scl_8.b         =   30
            C[0]_8          =   -2147483640
            C[1]_8          =   -2147483640
            Image_8         =   9
            ClassName_9     =   "CCWTextImage"
            szText_9        =   "Click to float"
            font_9          =   0
            Animator_8      =   0
            Blinker_8       =   0
            list[2]_4       =   10
            ClassName_10    =   "CCWDrawObj"
            opts_10         =   54
            SclRef_10.l     =   5
            SclRef_10.t     =   4
            SclRef_10.r     =   68
            SclRef_10.b     =   28
            Scl_10.l        =   81
            Scl_10.t        =   17
            Scl_10.r        =   143
            Scl_10.b        =   30
            C[0]_10         =   -2147483640
            C[1]_10         =   -2147483640
            Image_10        =   11
            ClassName_11    =   "CCWTextImage"
            szText_11       =   "Click to dock"
            font_11         =   0
            Animator_10     =   0
            Blinker_10      =   0
            list[1]_4       =   12
            ClassName_12    =   "CCWDrawObj"
            opts_12         =   62
            C[0]_12         =   32768
            C[1]_12         =   65535
            Image_12        =   13
            ClassName_13    =   "CCWPictImage"
            opts_13         =   256
            Rows_13         =   1
            Cols_13         =   1
            URL_13          =   "C:\In_Out\Software\DC-2020_ED_Analyzer_VB6\Resources\OFF_small_green switch.bmp"
            Pict_13         =   "Analyzer.frx":10BA60
            ColorReplaceWith_13=   8421504
            ColorReplace_13 =   8421504
            Tolerance_13    =   2
            Animator_12     =   0
            Blinker_12      =   0
            varVarType_3    =   5
            Array[1]_2      =   14
            ClassName_14    =   "CCWEnumElt"
            opts_14         =   1
            Name_14         =   "On"
            frame_14        =   286
            DrawList_14     =   15
            ClassName_15    =   "CDrawList"
            count_15        =   4
            list[4]_15      =   16
            ClassName_16    =   "CCWDrawObj"
            opts_16         =   54
            SclRef_16.l     =   4
            SclRef_16.t     =   4
            SclRef_16.r     =   69
            SclRef_16.b     =   27
            Scl_16.l        =   76
            Scl_16.t        =   1
            Scl_16.r        =   145
            Scl_16.b        =   19
            C[0]_16         =   -2147483640
            C[1]_16         =   -2147483640
            Image_16        =   6
            Animator_16     =   0
            Blinker_16      =   0
            list[3]_15      =   17
            ClassName_17    =   "CCWDrawObj"
            opts_17         =   54
            SclRef_17.l     =   5
            SclRef_17.t     =   4
            SclRef_17.r     =   68
            SclRef_17.b     =   28
            Scl_17.l        =   81
            Scl_17.t        =   17
            Scl_17.r        =   139
            Scl_17.b        =   30
            C[0]_17         =   -2147483640
            C[1]_17         =   -2147483640
            Image_17        =   9
            Animator_17     =   0
            Blinker_17      =   0
            list[2]_15      =   18
            ClassName_18    =   "CCWDrawObj"
            opts_18         =   52
            SclRef_18.l     =   5
            SclRef_18.t     =   4
            SclRef_18.r     =   68
            SclRef_18.b     =   28
            Scl_18.l        =   81
            Scl_18.t        =   17
            Scl_18.r        =   143
            Scl_18.b        =   30
            C[0]_18         =   -2147483640
            C[1]_18         =   -2147483640
            Image_18        =   11
            Animator_18     =   0
            Blinker_18      =   0
            list[1]_15      =   19
            ClassName_19    =   "CCWDrawObj"
            opts_19         =   62
            C[0]_19         =   32768
            C[1]_19         =   32768
            Image_19        =   20
            ClassName_20    =   "CCWPictImage"
            opts_20         =   256
            Rows_20         =   1
            Cols_20         =   1
            URL_20          =   "C:\In_Out\Software\DC-2020_ED_Analyzer_VB6\Resources\ON_orange_small_switch.bmp"
            Pict_20         =   "Analyzer.frx":10DE32
            ColorReplaceWith_20=   8421504
            ColorReplace_20 =   8421504
            Tolerance_20    =   2
            Animator_19     =   0
            Blinker_19      =   0
            varVarType_14   =   5
            Bindings_1      =   21
            ClassName_21    =   "CCWBindingHolderArray"
            Editor_21       =   22
            ClassName_22    =   "CCWBindingHolderArrayEditor"
            Owner_22        =   1
            Style_1         =   9
            drawState_1     =   1
            mechAction_1    =   1
            BGImg_1         =   23
            ClassName_23    =   "CCWDrawObj"
            opts_23         =   62
            Image_23        =   24
            ClassName_24    =   "CCWPictImage"
            opts_24         =   1280
            Rows_24         =   1
            Cols_24         =   1
            Pict_24         =   286
            F_24            =   -2147483633
            B_24            =   -2147483633
            ColorReplaceWith_24=   8421504
            ColorReplace_24 =   8421504
            Tolerance_24    =   2
            Animator_23     =   0
            Blinker_23      =   0
            Array_1         =   6
            Editor_1        =   0
            Array[0]_1      =   12
            Array[1]_1      =   19
            Array[2]_1      =   0
            Array[3]_1      =   0
            Array[4]_1      =   25
            ClassName_25    =   "CCWDrawObj"
            opts_25         =   62
            Image_25        =   9
            Animator_25     =   0
            Blinker_25      =   0
            Array[5]_1      =   26
            ClassName_26    =   "CCWDrawObj"
            opts_26         =   62
            Image_26        =   11
            Animator_26     =   0
            Blinker_26      =   0
            Label_1         =   27
            ClassName_27    =   "CCWDrawObj"
            opts_27         =   62
            C[0]_27         =   -2147483640
            Image_27        =   6
            Animator_27     =   0
            Blinker_27      =   0
         End
         Begin CWUIControlsLib.CWNumEdit cwNum_PressureSetPoint
            Height          =   360
            Index           =   0
            Left            =   75
            TabIndex        =   289
            Tag             =   "gain31"
            ToolTipText     =   "Iso 1 last 'good' floating pressure, range 0 to 90 psi, USE AutoSet in terminal menu->Pressure Control"
            Top             =   1890
            Width           =   1020
            _Version        =   524288
            _ExtentX        =   1799
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
            opts_1          =   393278
            C[0]_1          =   16777215
            C[1]_1          =   0
            BorderStyle_1   =   1
            TextAlignment_1 =   2
            format_1        =   2
            ClassName_2     =   "CCWFormat"
            Format_2        =   ".#0"
            scale_1         =   3
            ClassName_3     =   "CCWScale"
            dMax_3          =   10
            discInterval_3  =   0.01
            ValueVarType_1  =   5
            Value_Val_1     =   33.3
            IncValueVarType_1=   5
            IncValue_Val_1  =   0.05
            AccelIncVarType_1=   5
            AccelInc_Val_1  =   0.5
            RangeMinVarType_1=   5
            RangeMaxVarType_1=   5
            RangeMax_Val_1  =   90
            ButtonStyle_1   =   0
            Bindings_1      =   4
            ClassName_4     =   "CCWBindingHolderArray"
            Editor_4        =   5
            ClassName_5     =   "CCWBindingHolderArrayEditor"
            Owner_5         =   1
         End
         Begin CWUIControlsLib.CWNumEdit cwNum_PressureSetPoint
            Height          =   360
            Index           =   1
            Left            =   1140
            TabIndex        =   290
            Tag             =   "gain32"
            ToolTipText     =   "Iso 2 last 'good' floating pressure, range 0 to 90 psi, USE AutoSet in terminal menu->Pressure Control"
            Top             =   1890
            Width           =   1020
            _Version        =   524288
            _ExtentX        =   1799
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
            C[0]_1          =   16777215
            C[1]_1          =   0
            BorderStyle_1   =   1
            TextAlignment_1 =   2
            format_1        =   2
            ClassName_2     =   "CCWFormat"
            Format_2        =   ".#0"
            scale_1         =   3
            ClassName_3     =   "CCWScale"
            dMax_3          =   10
            discInterval_3  =   0.01
            ValueVarType_1  =   5
            Value_Val_1     =   44.44
            IncValueVarType_1=   5
            IncValue_Val_1  =   0.05
            AccelIncVarType_1=   5
            AccelInc_Val_1  =   0.5
            RangeMinVarType_1=   5
            RangeMaxVarType_1=   5
            RangeMax_Val_1  =   90
            ButtonStyle_1   =   0
            Bindings_1      =   4
            ClassName_4     =   "CCWBindingHolderArray"
            Editor_4        =   5
            ClassName_5     =   "CCWBindingHolderArrayEditor"
            Owner_5         =   1
         End
         Begin CWUIControlsLib.CWNumEdit cwNum_PressureSetPoint
            Height          =   360
            Index           =   2
            Left            =   2205
            TabIndex        =   291
            Tag             =   "gain33"
            ToolTipText     =   "Iso 3 last 'good' floating pressure, range 0 to 90 psi, USE AutoSet in terminal menu->Pressure Control"
            Top             =   1890
            Width           =   1020
            _Version        =   524288
            _ExtentX        =   1799
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
            C[0]_1          =   16777215
            C[1]_1          =   0
            BorderStyle_1   =   1
            TextAlignment_1 =   2
            format_1        =   2
            ClassName_2     =   "CCWFormat"
            Format_2        =   ".#0"
            scale_1         =   3
            ClassName_3     =   "CCWScale"
            dMax_3          =   10
            discInterval_3  =   0.01
            ValueVarType_1  =   5
            Value_Val_1     =   55.55
            IncValueVarType_1=   5
            IncValue_Val_1  =   0.05
            AccelIncVarType_1=   5
            AccelInc_Val_1  =   0.5
            RangeMinVarType_1=   5
            RangeMaxVarType_1=   5
            RangeMax_Val_1  =   90
            ButtonStyle_1   =   0
            Bindings_1      =   4
            ClassName_4     =   "CCWBindingHolderArray"
            Editor_4        =   5
            ClassName_5     =   "CCWBindingHolderArrayEditor"
            Owner_5         =   1
         End
         Begin CWUIControlsLib.CWNumEdit cwNum_PressureSetPoint
            Height          =   360
            Index           =   3
            Left            =   3270
            TabIndex        =   292
            Tag             =   "gain34"
            ToolTipText     =   "Iso 4 last 'good' floating pressure, range 0 to 90 psi, USE AutoSet in terminal menu->Pressure Control"
            Top             =   1890
            Width           =   1020
            _Version        =   524288
            _ExtentX        =   1799
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
            C[0]_1          =   16777215
            C[1]_1          =   0
            BorderStyle_1   =   1
            TextAlignment_1 =   2
            format_1        =   2
            ClassName_2     =   "CCWFormat"
            Format_2        =   ".#0"
            scale_1         =   3
            ClassName_3     =   "CCWScale"
            dMax_3          =   10
            discInterval_3  =   0.01
            ValueVarType_1  =   5
            Value_Val_1     =   66.66
            IncValueVarType_1=   5
            IncValue_Val_1  =   0.05
            AccelIncVarType_1=   5
            AccelInc_Val_1  =   0.5
            RangeMinVarType_1=   5
            RangeMaxVarType_1=   5
            RangeMax_Val_1  =   90
            ButtonStyle_1   =   0
            Bindings_1      =   4
            ClassName_4     =   "CCWBindingHolderArray"
            Editor_4        =   5
            ClassName_5     =   "CCWBindingHolderArrayEditor"
            Owner_5         =   1
         End
         Begin CWUIControlsLib.CWNumEdit cwNum_DockedPos
            Height          =   360
            Left            =   3285
            TabIndex        =   293
            Tag             =   "dkpo"
            ToolTipText     =   $"Analyzer.frx":110204
            Top             =   945
            Width           =   975
            _Version        =   524288
            _ExtentX        =   1720
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
            C[0]_1          =   16777215
            C[1]_1          =   0
            BorderStyle_1   =   1
            TextAlignment_1 =   2
            format_1        =   2
            ClassName_2     =   "CCWFormat"
            scale_1         =   3
            ClassName_3     =   "CCWScale"
            dMax_3          =   10
            discInterval_3  =   1
            ValueVarType_1  =   5
            Value_Val_1     =   3000
            IncValueVarType_1=   5
            IncValue_Val_1  =   10
            AccelIncVarType_1=   5
            AccelInc_Val_1  =   100
            RangeMinVarType_1=   5
            RangeMaxVarType_1=   5
            RangeMax_Val_1  =   8000
            ButtonStyle_1   =   0
            Bindings_1      =   4
            ClassName_4     =   "CCWBindingHolderArray"
            Editor_4        =   5
            ClassName_5     =   "CCWBindingHolderArrayEditor"
            Owner_5         =   1
         End
         Begin CWUIControlsLib.CWNumEdit cwNumFBgain
            Height          =   360
            Index           =   15
            Left            =   3150
            TabIndex        =   294
            Tag             =   "gain41"
            ToolTipText     =   "Multiplier to regular gain when stage motion is detected, 0.1 to 10"
            Top             =   3340
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
            C[0]_1          =   8438015
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
         Begin CWUIControlsLib.CWNumEdit cwNumFBgain
            Height          =   360
            Index           =   16
            Left            =   3150
            TabIndex        =   295
            Tag             =   "gain42"
            ToolTipText     =   "Multiplier to regular gain when stage motion is detected, 0.1 to 10"
            Top             =   3760
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
            C[0]_1          =   8438015
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
            Value_Val_1     =   2.22
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
         Begin CWUIControlsLib.CWNumEdit cwNumFBgain
            Height          =   360
            Index           =   17
            Left            =   3150
            TabIndex        =   296
            Tag             =   "gain43"
            ToolTipText     =   "Multiplier to regular gain when stage motion is detected, 0.1 to 10"
            Top             =   4185
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
            C[0]_1          =   8438015
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
            Value_Val_1     =   3.33
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
         Begin CWUIControlsLib.CWNumEdit cwNumFBgain
            Height          =   360
            Index           =   18
            Left            =   3150
            TabIndex        =   297
            Tag             =   "gain40"
            ToolTipText     =   "Multiplier to regular gain when Linear Motor feedback is enabled, 0.1 to 10"
            Top             =   2630
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
            C[0]_1          =   12648384
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
         Begin VB.Label Label1
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            Caption         =   " Floating pressures, psi"
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
            Height          =   950
            Index           =   37
            Left            =   45
            TabIndex        =   311
            ToolTipText     =   "Use AutoSet in terminal menu->Pressure Control"
            Top             =   1395
            Width           =   4265
         End
         Begin VB.Label Label1
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            Caption         =   "Damping"
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
            Index           =   17
            Left            =   1845
            TabIndex        =   310
            Top             =   3030
            Width           =   1095
         End
         Begin VB.Label Label1
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            Caption         =   "tY"
            BeginProperty Font
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   12
            Left            =   90
            TabIndex        =   309
            Top             =   4230
            Width           =   630
         End
         Begin VB.Label Label1
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            Caption         =   "tX"
            BeginProperty Font
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   11
            Left            =   90
            TabIndex        =   308
            Top             =   3780
            Width           =   630
         End
         Begin VB.Label Label1
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            Caption         =   "Z"
            BeginProperty Font
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   9
            Left            =   90
            TabIndex        =   307
            Top             =   3330
            Width           =   630
         End
         Begin VB.Label Label1
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "FB gain"
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
            Index           =   10
            Left            =   855
            TabIndex        =   306
            Top             =   3015
            Width           =   900
         End
         Begin VB.Label lblAxis
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Pneumatic"
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
            Height          =   330
            Index           =   0
            Left            =   0
            TabIndex        =   305
            ToolTipText     =   "Float payload by using vertical proximity sensors and valves controlling pressure in isolators"
            Top             =   0
            Width           =   4335
         End
         Begin VB.Label lbValveNum
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            Caption         =   "Iso 1"
            BeginProperty Font
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   270
            Index           =   12
            Left            =   360
            TabIndex        =   304
            Top             =   1710
            Width           =   645
         End
         Begin VB.Label lbValveNum
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            Caption         =   "Iso 2"
            BeginProperty Font
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   270
            Index           =   11
            Left            =   1395
            TabIndex        =   303
            Top             =   1710
            Width           =   645
         End
         Begin VB.Label lbValveNum
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            Caption         =   "Iso 3"
            BeginProperty Font
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   270
            Index           =   10
            Left            =   2520
            TabIndex        =   302
            Top             =   1710
            Width           =   645
         End
         Begin VB.Label lbValveNum
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            Caption         =   "Iso 4"
            BeginProperty Font
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   270
            Index           =   9
            Left            =   3555
            TabIndex        =   301
            Top             =   1710
            Width           =   645
         End
         Begin VB.Label lbValveNum
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Caption         =   "Docked Pos"
            BeginProperty Font
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   675
            Index           =   8
            Left            =   3285
            TabIndex        =   300
            ToolTipText     =   $"Analyzer.frx":1102B5
            Top             =   405
            Width           =   960
         End
         Begin VB.Label Label1
            Appearance      =   0  'Flat
            BackColor       =   &H0080C0FF&
            Caption         =   "FF boost"
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
            Index           =   8
            Left            =   3150
            TabIndex        =   299
            Top             =   3060
            Width           =   1080
         End
         Begin VB.Label Label1
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "LM on boost"
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
            Index           =   38
            Left            =   2970
            TabIndex        =   298
            Top             =   2350
            Width           =   1305
         End
      End
      Begin VB.Frame FrameVelocityFB
         BackColor       =   &H00E0E0E0&
         Height          =   4635
         Left            =   4005
         TabIndex        =   259
         Top             =   1170
         Width           =   3135
         Begin VB.CommandButton CmdResetGeoOffsets
            BackColor       =   &H0080C0FF&
            Caption         =   "Reset Input offsets"
            BeginProperty Font
               Name            =   "Arial"
               Size            =   9
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   1080
            Style           =   1  'Graphical
            TabIndex        =   313
            ToolTipText     =   "Removes offset from ALL ADC channels"
            Top             =   30
            Width           =   1995
         End
         Begin CWUIControlsLib.CWNumEdit cwNumFBgain
            Height          =   360
            Index           =   8
            Left            =   1755
            TabIndex        =   260
            Tag             =   "gain08"
            ToolTipText     =   "Gain of the axis"
            Top             =   2610
            Width           =   1200
            _Version        =   524288
            _ExtentX        =   2117
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
            C[0]_1          =   12648447
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
            Value_Val_1     =   2.81
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
         Begin CWUIControlsLib.CWNumEdit cwNumFBgain
            Height          =   360
            Index           =   6
            Left            =   1755
            TabIndex        =   261
            Tag             =   "gain06"
            ToolTipText     =   "Gain of the axis"
            Top             =   1620
            Width           =   1200
            _Version        =   524288
            _ExtentX        =   2117
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
            C[0]_1          =   12648447
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
            Value_Val_1     =   1.01
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
         Begin CWUIControlsLib.CWNumEdit cwNumFBgain
            Height          =   360
            Index           =   7
            Left            =   1755
            TabIndex        =   262
            Tag             =   "gain07"
            ToolTipText     =   "Gain of the axis"
            Top             =   2115
            Width           =   1200
            _Version        =   524288
            _ExtentX        =   2117
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
            C[0]_1          =   12648447
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
            Value_Val_1     =   0.71
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
         Begin CWUIControlsLib.CWButton cwBut_AxisEn
            Height          =   360
            Index           =   2
            Left            =   180
            TabIndex        =   263
            Tag             =   "loop_FB_7"
            ToolTipText     =   "Individual axis ON/OFF, when ALL AXES switch is ON"
            Top             =   2115
            Width           =   1455
            _Version        =   524288
            _ExtentX        =   2566
            _ExtentY        =   635
            _StockProps     =   68
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Reset_0         =   0   'False
            CompatibleVers_0=   524288
            Boolean_0       =   1
            ClassName_1     =   "CCWBoolean"
            opts_1          =   2614
            SclRef_1.l      =   2
            SclRef_1.t      =   2
            SclRef_1.r      =   99
            SclRef_1.b      =   26
            Scl_1.l         =   4
            Scl_1.t         =   4
            Scl_1.r         =   69
            Scl_1.b         =   25
            C[0]_1          =   12648447
            Enum_1          =   2
            ClassName_2     =   "CCWEnum"
            Array_2         =   2
            Editor_2        =   0
            Array[0]_2      =   3
            ClassName_3     =   "CCWEnumElt"
            opts_3          =   1
            Name_3          =   "Off"
            frame_3         =   286
            DrawList_3      =   4
            ClassName_4     =   "CDrawList"
            count_4         =   4
            list[4]_4       =   5
            ClassName_5     =   "CCWDrawObj"
            opts_5          =   54
            SclRef_5.l      =   4
            SclRef_5.t      =   4
            SclRef_5.r      =   69
            SclRef_5.b      =   24
            Scl_5.l         =   73
            Scl_5.t         =   5
            Scl_5.r         =   93
            Scl_5.b         =   24
            C[0]_5          =   -2147483640
            C[1]_5          =   -2147483640
            Image_5         =   6
            ClassName_6     =   "CCWTextImage"
            szText_6        =   " Y"
            font_6          =   7
            ClassName_7     =   "CCWFont"
            bFont_7         =   -1  'True
            BeginProperty Font_7 {0BE35203-8F91-11CE-9DE3-00AA004BB851}
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Animator_5      =   0
            Blinker_5       =   0
            list[3]_4       =   8
            ClassName_8     =   "CCWDrawObj"
            opts_8          =   60
            C[0]_8          =   -2147483640
            C[1]_8          =   -2147483640
            Image_8         =   9
            ClassName_9     =   "CCWTextImage"
            font_9          =   0
            Animator_8      =   0
            Blinker_8       =   0
            list[2]_4       =   10
            ClassName_10    =   "CCWDrawObj"
            opts_10         =   60
            C[0]_10         =   -2147483640
            C[1]_10         =   -2147483640
            Image_10        =   11
            ClassName_11    =   "CCWTextImage"
            font_11         =   0
            Animator_10     =   0
            Blinker_10      =   0
            list[1]_4       =   12
            ClassName_12    =   "CCWDrawObj"
            opts_12         =   62
            C[0]_12         =   32768
            C[1]_12         =   65535
            Image_12        =   13
            ClassName_13    =   "CCWPictImage"
            opts_13         =   256
            Rows_13         =   1
            Cols_13         =   1
            URL_13          =   "\\tsclient\C\In_Out\Software\DC2020_WEB_Analyzer\OFF_small_switch.bmp"
            Pict_13         =   "Analyzer.frx":110366
            ColorReplaceWith_13=   8421504
            ColorReplace_13 =   8421504
            Tolerance_13    =   2
            Animator_12     =   0
            Blinker_12      =   0
            varVarType_3    =   5
            Array[1]_2      =   14
            ClassName_14    =   "CCWEnumElt"
            opts_14         =   1
            Name_14         =   "On"
            frame_14        =   286
            DrawList_14     =   15
            ClassName_15    =   "CDrawList"
            count_15        =   4
            list[4]_15      =   16
            ClassName_16    =   "CCWDrawObj"
            opts_16         =   54
            SclRef_16.l     =   4
            SclRef_16.t     =   4
            SclRef_16.r     =   69
            SclRef_16.b     =   24
            Scl_16.l        =   73
            Scl_16.t        =   5
            Scl_16.r        =   93
            Scl_16.b        =   24
            C[0]_16         =   -2147483640
            C[1]_16         =   -2147483640
            Image_16        =   6
            Animator_16     =   0
            Blinker_16      =   0
            list[3]_15      =   17
            ClassName_17    =   "CCWDrawObj"
            opts_17         =   60
            C[0]_17         =   -2147483640
            C[1]_17         =   -2147483640
            Image_17        =   9
            Animator_17     =   0
            Blinker_17      =   0
            list[2]_15      =   18
            ClassName_18    =   "CCWDrawObj"
            opts_18         =   60
            C[0]_18         =   -2147483640
            C[1]_18         =   -2147483640
            Image_18        =   11
            Animator_18     =   0
            Blinker_18      =   0
            list[1]_15      =   19
            ClassName_19    =   "CCWDrawObj"
            opts_19         =   62
            C[0]_19         =   32768
            C[1]_19         =   32768
            Image_19        =   20
            ClassName_20    =   "CCWPictImage"
            opts_20         =   256
            Rows_20         =   1
            Cols_20         =   1
            URL_20          =   "\\tsclient\C\In_Out\Software\DC2020_WEB_Analyzer\ON_small_switch.bmp"
            Pict_20         =   "Analyzer.frx":112738
            ColorReplaceWith_20=   8421504
            ColorReplace_20 =   8421504
            Tolerance_20    =   2
            Animator_19     =   0
            Blinker_19      =   0
            varVarType_14   =   5
            Bindings_1      =   21
            ClassName_21    =   "CCWBindingHolderArray"
            Editor_21       =   22
            ClassName_22    =   "CCWBindingHolderArrayEditor"
            Owner_22        =   1
            Style_1         =   9
            mechAction_1    =   1
            BGImg_1         =   23
            ClassName_23    =   "CCWDrawObj"
            opts_23         =   62
            Image_23        =   24
            ClassName_24    =   "CCWPictImage"
            opts_24         =   1280
            Rows_24         =   1
            Cols_24         =   1
            Pict_24         =   286
            F_24            =   -2147483633
            B_24            =   -2147483633
            ColorReplaceWith_24=   8421504
            ColorReplace_24 =   8421504
            Tolerance_24    =   2
            Animator_23     =   0
            Blinker_23      =   0
            Array_1         =   6
            Editor_1        =   0
            Array[0]_1      =   12
            Array[1]_1      =   19
            Array[2]_1      =   0
            Array[3]_1      =   0
            Array[4]_1      =   25
            ClassName_25    =   "CCWDrawObj"
            opts_25         =   62
            Image_25        =   9
            Animator_25     =   0
            Blinker_25      =   0
            Array[5]_1      =   26
            ClassName_26    =   "CCWDrawObj"
            opts_26         =   62
            Image_26        =   11
            Animator_26     =   0
            Blinker_26      =   0
            Label_1         =   27
            ClassName_27    =   "CCWDrawObj"
            opts_27         =   62
            C[0]_27         =   -2147483640
            Image_27        =   6
            Animator_27     =   0
            Blinker_27      =   0
         End
         Begin CWUIControlsLib.CWButton cwBut_AxisEn
            Height          =   360
            Index           =   3
            Left            =   180
            TabIndex        =   264
            Tag             =   "loop_FB_8"
            ToolTipText     =   "Individual axis ON/OFF, when ALL AXES switch is ON"
            Top             =   2610
            Width           =   1455
            _Version        =   524288
            _ExtentX        =   2566
            _ExtentY        =   635
            _StockProps     =   68
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Reset_0         =   0   'False
            CompatibleVers_0=   524288
            Boolean_0       =   1
            ClassName_1     =   "CCWBoolean"
            opts_1          =   2614
            SclRef_1.l      =   2
            SclRef_1.t      =   2
            SclRef_1.r      =   99
            SclRef_1.b      =   26
            Scl_1.l         =   4
            Scl_1.t         =   4
            Scl_1.r         =   69
            Scl_1.b         =   25
            C[0]_1          =   12648447
            Enum_1          =   2
            ClassName_2     =   "CCWEnum"
            Array_2         =   2
            Editor_2        =   0
            Array[0]_2      =   3
            ClassName_3     =   "CCWEnumElt"
            opts_3          =   1
            Name_3          =   "Off"
            frame_3         =   286
            DrawList_3      =   4
            ClassName_4     =   "CDrawList"
            count_4         =   4
            list[4]_4       =   5
            ClassName_5     =   "CCWDrawObj"
            opts_5          =   54
            SclRef_5.l      =   4
            SclRef_5.t      =   4
            SclRef_5.r      =   69
            SclRef_5.b      =   24
            Scl_5.l         =   73
            Scl_5.t         =   5
            Scl_5.r         =   93
            Scl_5.b         =   24
            C[0]_5          =   -2147483640
            C[1]_5          =   -2147483640
            Image_5         =   6
            ClassName_6     =   "CCWTextImage"
            szText_6        =   "tZ"
            font_6          =   7
            ClassName_7     =   "CCWFont"
            bFont_7         =   -1  'True
            BeginProperty Font_7 {0BE35203-8F91-11CE-9DE3-00AA004BB851}
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Animator_5      =   0
            Blinker_5       =   0
            list[3]_4       =   8
            ClassName_8     =   "CCWDrawObj"
            opts_8          =   60
            C[0]_8          =   -2147483640
            C[1]_8          =   -2147483640
            Image_8         =   9
            ClassName_9     =   "CCWTextImage"
            font_9          =   0
            Animator_8      =   0
            Blinker_8       =   0
            list[2]_4       =   10
            ClassName_10    =   "CCWDrawObj"
            opts_10         =   60
            C[0]_10         =   -2147483640
            C[1]_10         =   -2147483640
            Image_10        =   11
            ClassName_11    =   "CCWTextImage"
            font_11         =   0
            Animator_10     =   0
            Blinker_10      =   0
            list[1]_4       =   12
            ClassName_12    =   "CCWDrawObj"
            opts_12         =   62
            C[0]_12         =   32768
            C[1]_12         =   65535
            Image_12        =   13
            ClassName_13    =   "CCWPictImage"
            opts_13         =   256
            Rows_13         =   1
            Cols_13         =   1
            URL_13          =   "\\tsclient\C\In_Out\Software\DC2020_WEB_Analyzer\OFF_small_switch.bmp"
            Pict_13         =   "Analyzer.frx":114B0A
            ColorReplaceWith_13=   8421504
            ColorReplace_13 =   8421504
            Tolerance_13    =   2
            Animator_12     =   0
            Blinker_12      =   0
            varVarType_3    =   5
            Array[1]_2      =   14
            ClassName_14    =   "CCWEnumElt"
            opts_14         =   1
            Name_14         =   "On"
            frame_14        =   286
            DrawList_14     =   15
            ClassName_15    =   "CDrawList"
            count_15        =   4
            list[4]_15      =   16
            ClassName_16    =   "CCWDrawObj"
            opts_16         =   54
            SclRef_16.l     =   4
            SclRef_16.t     =   4
            SclRef_16.r     =   69
            SclRef_16.b     =   24
            Scl_16.l        =   73
            Scl_16.t        =   5
            Scl_16.r        =   93
            Scl_16.b        =   24
            C[0]_16         =   -2147483640
            C[1]_16         =   -2147483640
            Image_16        =   6
            Animator_16     =   0
            Blinker_16      =   0
            list[3]_15      =   17
            ClassName_17    =   "CCWDrawObj"
            opts_17         =   60
            C[0]_17         =   -2147483640
            C[1]_17         =   -2147483640
            Image_17        =   9
            Animator_17     =   0
            Blinker_17      =   0
            list[2]_15      =   18
            ClassName_18    =   "CCWDrawObj"
            opts_18         =   60
            C[0]_18         =   -2147483640
            C[1]_18         =   -2147483640
            Image_18        =   11
            Animator_18     =   0
            Blinker_18      =   0
            list[1]_15      =   19
            ClassName_19    =   "CCWDrawObj"
            opts_19         =   62
            C[0]_19         =   32768
            C[1]_19         =   32768
            Image_19        =   20
            ClassName_20    =   "CCWPictImage"
            opts_20         =   256
            Rows_20         =   1
            Cols_20         =   1
            URL_20          =   "\\tsclient\C\In_Out\Software\DC2020_WEB_Analyzer\ON_small_switch.bmp"
            Pict_20         =   "Analyzer.frx":116EDC
            ColorReplaceWith_20=   8421504
            ColorReplace_20 =   8421504
            Tolerance_20    =   2
            Animator_19     =   0
            Blinker_19      =   0
            varVarType_14   =   5
            Bindings_1      =   21
            ClassName_21    =   "CCWBindingHolderArray"
            Editor_21       =   22
            ClassName_22    =   "CCWBindingHolderArrayEditor"
            Owner_22        =   1
            Style_1         =   9
            drawState_1     =   1
            mechAction_1    =   1
            BGImg_1         =   23
            ClassName_23    =   "CCWDrawObj"
            opts_23         =   62
            Image_23        =   24
            ClassName_24    =   "CCWPictImage"
            opts_24         =   1280
            Rows_24         =   1
            Cols_24         =   1
            Pict_24         =   286
            F_24            =   -2147483633
            B_24            =   -2147483633
            ColorReplaceWith_24=   8421504
            ColorReplace_24 =   8421504
            Tolerance_24    =   2
            Animator_23     =   0
            Blinker_23      =   0
            Array_1         =   6
            Editor_1        =   0
            Array[0]_1      =   12
            Array[1]_1      =   19
            Array[2]_1      =   0
            Array[3]_1      =   0
            Array[4]_1      =   25
            ClassName_25    =   "CCWDrawObj"
            opts_25         =   62
            Image_25        =   9
            Animator_25     =   0
            Blinker_25      =   0
            Array[5]_1      =   26
            ClassName_26    =   "CCWDrawObj"
            opts_26         =   62
            Image_26        =   11
            Animator_26     =   0
            Blinker_26      =   0
            Label_1         =   27
            ClassName_27    =   "CCWDrawObj"
            opts_27         =   62
            C[0]_27         =   -2147483640
            Image_27        =   6
            Animator_27     =   0
            Blinker_27      =   0
         End
         Begin CWUIControlsLib.CWButton cwBut_AxisEn
            Height          =   360
            Index           =   1
            Left            =   180
            TabIndex        =   265
            Tag             =   "loop_FB_6"
            ToolTipText     =   "Individual axis ON/OFF, when ALL AXES switch is ON"
            Top             =   1620
            Width           =   1455
            _Version        =   524288
            _ExtentX        =   2566
            _ExtentY        =   635
            _StockProps     =   68
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Reset_0         =   0   'False
            CompatibleVers_0=   524288
            Boolean_0       =   1
            ClassName_1     =   "CCWBoolean"
            opts_1          =   2614
            SclRef_1.l      =   2
            SclRef_1.t      =   2
            SclRef_1.r      =   99
            SclRef_1.b      =   26
            Scl_1.l         =   4
            Scl_1.t         =   4
            Scl_1.r         =   69
            Scl_1.b         =   25
            C[0]_1          =   12648447
            Enum_1          =   2
            ClassName_2     =   "CCWEnum"
            Array_2         =   2
            Editor_2        =   0
            Array[0]_2      =   3
            ClassName_3     =   "CCWEnumElt"
            opts_3          =   1
            Name_3          =   "Off"
            frame_3         =   286
            DrawList_3      =   4
            ClassName_4     =   "CDrawList"
            count_4         =   4
            list[4]_4       =   5
            ClassName_5     =   "CCWDrawObj"
            opts_5          =   54
            SclRef_5.l      =   4
            SclRef_5.t      =   4
            SclRef_5.r      =   69
            SclRef_5.b      =   24
            Scl_5.l         =   73
            Scl_5.t         =   5
            Scl_5.r         =   93
            Scl_5.b         =   24
            C[0]_5          =   -2147483640
            C[1]_5          =   -2147483640
            Image_5         =   6
            ClassName_6     =   "CCWTextImage"
            szText_6        =   " X"
            font_6          =   7
            ClassName_7     =   "CCWFont"
            bFont_7         =   -1  'True
            BeginProperty Font_7 {0BE35203-8F91-11CE-9DE3-00AA004BB851}
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Animator_5      =   0
            Blinker_5       =   0
            list[3]_4       =   8
            ClassName_8     =   "CCWDrawObj"
            opts_8          =   60
            C[0]_8          =   -2147483640
            C[1]_8          =   -2147483640
            Image_8         =   9
            ClassName_9     =   "CCWTextImage"
            font_9          =   0
            Animator_8      =   0
            Blinker_8       =   0
            list[2]_4       =   10
            ClassName_10    =   "CCWDrawObj"
            opts_10         =   60
            C[0]_10         =   -2147483640
            C[1]_10         =   -2147483640
            Image_10        =   11
            ClassName_11    =   "CCWTextImage"
            font_11         =   0
            Animator_10     =   0
            Blinker_10      =   0
            list[1]_4       =   12
            ClassName_12    =   "CCWDrawObj"
            opts_12         =   62
            C[0]_12         =   32768
            C[1]_12         =   65535
            Image_12        =   13
            ClassName_13    =   "CCWPictImage"
            opts_13         =   256
            Rows_13         =   1
            Cols_13         =   1
            URL_13          =   "\\tsclient\C\In_Out\Software\DC2020_WEB_Analyzer\OFF_small_switch.bmp"
            Pict_13         =   "Analyzer.frx":1192AE
            ColorReplaceWith_13=   8421504
            ColorReplace_13 =   8421504
            Tolerance_13    =   2
            Animator_12     =   0
            Blinker_12      =   0
            varVarType_3    =   5
            Array[1]_2      =   14
            ClassName_14    =   "CCWEnumElt"
            opts_14         =   1
            Name_14         =   "On"
            frame_14        =   286
            DrawList_14     =   15
            ClassName_15    =   "CDrawList"
            count_15        =   4
            list[4]_15      =   16
            ClassName_16    =   "CCWDrawObj"
            opts_16         =   54
            SclRef_16.l     =   4
            SclRef_16.t     =   4
            SclRef_16.r     =   69
            SclRef_16.b     =   24
            Scl_16.l        =   73
            Scl_16.t        =   5
            Scl_16.r        =   93
            Scl_16.b        =   24
            C[0]_16         =   -2147483640
            C[1]_16         =   -2147483640
            Image_16        =   6
            Animator_16     =   0
            Blinker_16      =   0
            list[3]_15      =   17
            ClassName_17    =   "CCWDrawObj"
            opts_17         =   60
            C[0]_17         =   -2147483640
            C[1]_17         =   -2147483640
            Image_17        =   9
            Animator_17     =   0
            Blinker_17      =   0
            list[2]_15      =   18
            ClassName_18    =   "CCWDrawObj"
            opts_18         =   60
            C[0]_18         =   -2147483640
            C[1]_18         =   -2147483640
            Image_18        =   11
            Animator_18     =   0
            Blinker_18      =   0
            list[1]_15      =   19
            ClassName_19    =   "CCWDrawObj"
            opts_19         =   62
            C[0]_19         =   32768
            C[1]_19         =   32768
            Image_19        =   20
            ClassName_20    =   "CCWPictImage"
            opts_20         =   256
            Rows_20         =   1
            Cols_20         =   1
            URL_20          =   "\\tsclient\C\In_Out\Software\DC2020_WEB_Analyzer\ON_small_switch.bmp"
            Pict_20         =   "Analyzer.frx":11B680
            ColorReplaceWith_20=   8421504
            ColorReplace_20 =   8421504
            Tolerance_20    =   2
            Animator_19     =   0
            Blinker_19      =   0
            varVarType_14   =   5
            Bindings_1      =   21
            ClassName_21    =   "CCWBindingHolderArray"
            Editor_21       =   22
            ClassName_22    =   "CCWBindingHolderArrayEditor"
            Owner_22        =   1
            Style_1         =   9
            mechAction_1    =   1
            BGImg_1         =   23
            ClassName_23    =   "CCWDrawObj"
            opts_23         =   62
            Image_23        =   24
            ClassName_24    =   "CCWPictImage"
            opts_24         =   1280
            Rows_24         =   1
            Cols_24         =   1
            Pict_24         =   286
            F_24            =   -2147483633
            B_24            =   -2147483633
            ColorReplaceWith_24=   8421504
            ColorReplace_24 =   8421504
            Tolerance_24    =   2
            Animator_23     =   0
            Blinker_23      =   0
            Array_1         =   6
            Editor_1        =   0
            Array[0]_1      =   12
            Array[1]_1      =   19
            Array[2]_1      =   0
            Array[3]_1      =   0
            Array[4]_1      =   25
            ClassName_25    =   "CCWDrawObj"
            opts_25         =   62
            Image_25        =   9
            Animator_25     =   0
            Blinker_25      =   0
            Array[5]_1      =   26
            ClassName_26    =   "CCWDrawObj"
            opts_26         =   62
            Image_26        =   11
            Animator_26     =   0
            Blinker_26      =   0
            Label_1         =   27
            ClassName_27    =   "CCWDrawObj"
            opts_27         =   62
            C[0]_27         =   -2147483640
            Image_27        =   6
            Animator_27     =   0
            Blinker_27      =   0
         End
         Begin CWUIControlsLib.CWNumEdit cwNumFBgain
            Height          =   360
            Index           =   9
            Left            =   1755
            TabIndex        =   266
            Tag             =   "gain09"
            ToolTipText     =   "Gain of the axis"
            Top             =   3105
            Width           =   1200
            _Version        =   524288
            _ExtentX        =   2117
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
            C[0]_1          =   12648384
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
            Value_Val_1     =   0.71
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
         Begin CWUIControlsLib.CWButton cwBut_AxisEn
            Height          =   360
            Index           =   4
            Left            =   180
            TabIndex        =   267
            Tag             =   "loop_FB_9"
            ToolTipText     =   "Individual axis ON/OFF, when ALL AXES switch is ON"
            Top             =   3105
            Width           =   1455
            _Version        =   524288
            _ExtentX        =   2566
            _ExtentY        =   635
            _StockProps     =   68
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Reset_0         =   0   'False
            CompatibleVers_0=   524288
            Boolean_0       =   1
            ClassName_1     =   "CCWBoolean"
            opts_1          =   2614
            SclRef_1.l      =   2
            SclRef_1.t      =   2
            SclRef_1.r      =   99
            SclRef_1.b      =   26
            Scl_1.l         =   4
            Scl_1.t         =   4
            Scl_1.r         =   69
            Scl_1.b         =   25
            C[0]_1          =   12648384
            Enum_1          =   2
            ClassName_2     =   "CCWEnum"
            Array_2         =   2
            Editor_2        =   0
            Array[0]_2      =   3
            ClassName_3     =   "CCWEnumElt"
            opts_3          =   1
            Name_3          =   "Off"
            frame_3         =   286
            DrawList_3      =   4
            ClassName_4     =   "CDrawList"
            count_4         =   4
            list[4]_4       =   5
            ClassName_5     =   "CCWDrawObj"
            opts_5          =   54
            SclRef_5.l      =   4
            SclRef_5.t      =   4
            SclRef_5.r      =   69
            SclRef_5.b      =   24
            Scl_5.l         =   73
            Scl_5.t         =   5
            Scl_5.r         =   93
            Scl_5.b         =   24
            C[0]_5          =   -2147483640
            C[1]_5          =   -2147483640
            Image_5         =   6
            ClassName_6     =   "CCWTextImage"
            szText_6        =   " Z"
            font_6          =   7
            ClassName_7     =   "CCWFont"
            bFont_7         =   -1  'True
            BeginProperty Font_7 {0BE35203-8F91-11CE-9DE3-00AA004BB851}
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Animator_5      =   0
            Blinker_5       =   0
            list[3]_4       =   8
            ClassName_8     =   "CCWDrawObj"
            opts_8          =   60
            C[0]_8          =   -2147483640
            C[1]_8          =   -2147483640
            Image_8         =   9
            ClassName_9     =   "CCWTextImage"
            font_9          =   0
            Animator_8      =   0
            Blinker_8       =   0
            list[2]_4       =   10
            ClassName_10    =   "CCWDrawObj"
            opts_10         =   60
            C[0]_10         =   -2147483640
            C[1]_10         =   -2147483640
            Image_10        =   11
            ClassName_11    =   "CCWTextImage"
            font_11         =   0
            Animator_10     =   0
            Blinker_10      =   0
            list[1]_4       =   12
            ClassName_12    =   "CCWDrawObj"
            opts_12         =   62
            C[0]_12         =   32768
            C[1]_12         =   65535
            Image_12        =   13
            ClassName_13    =   "CCWPictImage"
            opts_13         =   256
            Rows_13         =   1
            Cols_13         =   1
            URL_13          =   "\\tsclient\C\In_Out\Software\DC2020_WEB_Analyzer\OFF_small_switch.bmp"
            Pict_13         =   "Analyzer.frx":11DA52
            ColorReplaceWith_13=   8421504
            ColorReplace_13 =   8421504
            Tolerance_13    =   2
            Animator_12     =   0
            Blinker_12      =   0
            varVarType_3    =   5
            Array[1]_2      =   14
            ClassName_14    =   "CCWEnumElt"
            opts_14         =   1
            Name_14         =   "On"
            frame_14        =   286
            DrawList_14     =   15
            ClassName_15    =   "CDrawList"
            count_15        =   4
            list[4]_15      =   16
            ClassName_16    =   "CCWDrawObj"
            opts_16         =   54
            SclRef_16.l     =   4
            SclRef_16.t     =   4
            SclRef_16.r     =   69
            SclRef_16.b     =   24
            Scl_16.l        =   73
            Scl_16.t        =   5
            Scl_16.r        =   93
            Scl_16.b        =   24
            C[0]_16         =   -2147483640
            C[1]_16         =   -2147483640
            Image_16        =   6
            Animator_16     =   0
            Blinker_16      =   0
            list[3]_15      =   17
            ClassName_17    =   "CCWDrawObj"
            opts_17         =   60
            C[0]_17         =   -2147483640
            C[1]_17         =   -2147483640
            Image_17        =   9
            Animator_17     =   0
            Blinker_17      =   0
            list[2]_15      =   18
            ClassName_18    =   "CCWDrawObj"
            opts_18         =   60
            C[0]_18         =   -2147483640
            C[1]_18         =   -2147483640
            Image_18        =   11
            Animator_18     =   0
            Blinker_18      =   0
            list[1]_15      =   19
            ClassName_19    =   "CCWDrawObj"
            opts_19         =   62
            C[0]_19         =   32768
            C[1]_19         =   32768
            Image_19        =   20
            ClassName_20    =   "CCWPictImage"
            opts_20         =   256
            Rows_20         =   1
            Cols_20         =   1
            URL_20          =   "\\tsclient\C\In_Out\Software\DC2020_WEB_Analyzer\ON_small_switch.bmp"
            Pict_20         =   "Analyzer.frx":11FE24
            ColorReplaceWith_20=   8421504
            ColorReplace_20 =   8421504
            Tolerance_20    =   2
            Animator_19     =   0
            Blinker_19      =   0
            varVarType_14   =   5
            Bindings_1      =   21
            ClassName_21    =   "CCWBindingHolderArray"
            Editor_21       =   22
            ClassName_22    =   "CCWBindingHolderArrayEditor"
            Owner_22        =   1
            Style_1         =   9
            mechAction_1    =   1
            BGImg_1         =   23
            ClassName_23    =   "CCWDrawObj"
            opts_23         =   62
            Image_23        =   24
            ClassName_24    =   "CCWPictImage"
            opts_24         =   1280
            Rows_24         =   1
            Cols_24         =   1
            Pict_24         =   286
            F_24            =   -2147483633
            B_24            =   -2147483633
            ColorReplaceWith_24=   8421504
            ColorReplace_24 =   8421504
            Tolerance_24    =   2
            Animator_23     =   0
            Blinker_23      =   0
            Array_1         =   6
            Editor_1        =   0
            Array[0]_1      =   12
            Array[1]_1      =   19
            Array[2]_1      =   0
            Array[3]_1      =   0
            Array[4]_1      =   25
            ClassName_25    =   "CCWDrawObj"
            opts_25         =   62
            Image_25        =   9
            Animator_25     =   0
            Blinker_25      =   0
            Array[5]_1      =   26
            ClassName_26    =   "CCWDrawObj"
            opts_26         =   62
            Image_26        =   11
            Animator_26     =   0
            Blinker_26      =   0
            Label_1         =   27
            ClassName_27    =   "CCWDrawObj"
            opts_27         =   62
            C[0]_27         =   -2147483640
            Image_27        =   6
            Animator_27     =   0
            Blinker_27      =   0
         End
         Begin CWUIControlsLib.CWNumEdit cwNumFBgain
            Height          =   360
            Index           =   10
            Left            =   1755
            TabIndex        =   268
            Tag             =   "gain0A"
            ToolTipText     =   "Gain of the axis"
            Top             =   3600
            Width           =   1200
            _Version        =   524288
            _ExtentX        =   2117
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
            C[0]_1          =   12648384
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
            Value_Val_1     =   0.71
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
         Begin CWUIControlsLib.CWButton cwBut_AxisEn
            Height          =   360
            Index           =   5
            Left            =   180
            TabIndex        =   269
            Tag             =   "loop_FB_A"
            ToolTipText     =   "Individual axis ON/OFF, when ALL AXES switch is ON"
            Top             =   3600
            Width           =   1455
            _Version        =   524288
            _ExtentX        =   2566
            _ExtentY        =   635
            _StockProps     =   68
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Reset_0         =   0   'False
            CompatibleVers_0=   524288
            Boolean_0       =   1
            ClassName_1     =   "CCWBoolean"
            opts_1          =   2614
            SclRef_1.l      =   2
            SclRef_1.t      =   2
            SclRef_1.r      =   99
            SclRef_1.b      =   26
            Scl_1.l         =   4
            Scl_1.t         =   4
            Scl_1.r         =   69
            Scl_1.b         =   25
            C[0]_1          =   12648384
            Enum_1          =   2
            ClassName_2     =   "CCWEnum"
            Array_2         =   2
            Editor_2        =   0
            Array[0]_2      =   3
            ClassName_3     =   "CCWEnumElt"
            opts_3          =   1
            Name_3          =   "Off"
            frame_3         =   286
            DrawList_3      =   4
            ClassName_4     =   "CDrawList"
            count_4         =   4
            list[4]_4       =   5
            ClassName_5     =   "CCWDrawObj"
            opts_5          =   54
            SclRef_5.l      =   4
            SclRef_5.t      =   4
            SclRef_5.r      =   69
            SclRef_5.b      =   24
            Scl_5.l         =   73
            Scl_5.t         =   5
            Scl_5.r         =   93
            Scl_5.b         =   24
            C[0]_5          =   -2147483640
            C[1]_5          =   -2147483640
            Image_5         =   6
            ClassName_6     =   "CCWTextImage"
            szText_6        =   "tX"
            font_6          =   7
            ClassName_7     =   "CCWFont"
            bFont_7         =   -1  'True
            BeginProperty Font_7 {0BE35203-8F91-11CE-9DE3-00AA004BB851}
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Animator_5      =   0
            Blinker_5       =   0
            list[3]_4       =   8
            ClassName_8     =   "CCWDrawObj"
            opts_8          =   60
            C[0]_8          =   -2147483640
            C[1]_8          =   -2147483640
            Image_8         =   9
            ClassName_9     =   "CCWTextImage"
            font_9          =   0
            Animator_8      =   0
            Blinker_8       =   0
            list[2]_4       =   10
            ClassName_10    =   "CCWDrawObj"
            opts_10         =   60
            C[0]_10         =   -2147483640
            C[1]_10         =   -2147483640
            Image_10        =   11
            ClassName_11    =   "CCWTextImage"
            font_11         =   0
            Animator_10     =   0
            Blinker_10      =   0
            list[1]_4       =   12
            ClassName_12    =   "CCWDrawObj"
            opts_12         =   62
            C[0]_12         =   32768
            C[1]_12         =   65535
            Image_12        =   13
            ClassName_13    =   "CCWPictImage"
            opts_13         =   256
            Rows_13         =   1
            Cols_13         =   1
            URL_13          =   "\\tsclient\C\In_Out\Software\DC2020_WEB_Analyzer\OFF_small_switch.bmp"
            Pict_13         =   "Analyzer.frx":1221F6
            ColorReplaceWith_13=   8421504
            ColorReplace_13 =   8421504
            Tolerance_13    =   2
            Animator_12     =   0
            Blinker_12      =   0
            varVarType_3    =   5
            Array[1]_2      =   14
            ClassName_14    =   "CCWEnumElt"
            opts_14         =   1
            Name_14         =   "On"
            frame_14        =   286
            DrawList_14     =   15
            ClassName_15    =   "CDrawList"
            count_15        =   4
            list[4]_15      =   16
            ClassName_16    =   "CCWDrawObj"
            opts_16         =   54
            SclRef_16.l     =   4
            SclRef_16.t     =   4
            SclRef_16.r     =   69
            SclRef_16.b     =   24
            Scl_16.l        =   73
            Scl_16.t        =   5
            Scl_16.r        =   93
            Scl_16.b        =   24
            C[0]_16         =   -2147483640
            C[1]_16         =   -2147483640
            Image_16        =   6
            Animator_16     =   0
            Blinker_16      =   0
            list[3]_15      =   17
            ClassName_17    =   "CCWDrawObj"
            opts_17         =   60
            C[0]_17         =   -2147483640
            C[1]_17         =   -2147483640
            Image_17        =   9
            Animator_17     =   0
            Blinker_17      =   0
            list[2]_15      =   18
            ClassName_18    =   "CCWDrawObj"
            opts_18         =   60
            C[0]_18         =   -2147483640
            C[1]_18         =   -2147483640
            Image_18        =   11
            Animator_18     =   0
            Blinker_18      =   0
            list[1]_15      =   19
            ClassName_19    =   "CCWDrawObj"
            opts_19         =   62
            C[0]_19         =   32768
            C[1]_19         =   32768
            Image_19        =   20
            ClassName_20    =   "CCWPictImage"
            opts_20         =   256
            Rows_20         =   1
            Cols_20         =   1
            URL_20          =   "\\tsclient\C\In_Out\Software\DC2020_WEB_Analyzer\ON_small_switch.bmp"
            Pict_20         =   "Analyzer.frx":1245C8
            ColorReplaceWith_20=   8421504
            ColorReplace_20 =   8421504
            Tolerance_20    =   2
            Animator_19     =   0
            Blinker_19      =   0
            varVarType_14   =   5
            Bindings_1      =   21
            ClassName_21    =   "CCWBindingHolderArray"
            Editor_21       =   22
            ClassName_22    =   "CCWBindingHolderArrayEditor"
            Owner_22        =   1
            Style_1         =   9
            mechAction_1    =   1
            BGImg_1         =   23
            ClassName_23    =   "CCWDrawObj"
            opts_23         =   62
            Image_23        =   24
            ClassName_24    =   "CCWPictImage"
            opts_24         =   1280
            Rows_24         =   1
            Cols_24         =   1
            Pict_24         =   286
            F_24            =   -2147483633
            B_24            =   -2147483633
            ColorReplaceWith_24=   8421504
            ColorReplace_24 =   8421504
            Tolerance_24    =   2
            Animator_23     =   0
            Blinker_23      =   0
            Array_1         =   6
            Editor_1        =   0
            Array[0]_1      =   12
            Array[1]_1      =   19
            Array[2]_1      =   0
            Array[3]_1      =   0
            Array[4]_1      =   25
            ClassName_25    =   "CCWDrawObj"
            opts_25         =   62
            Image_25        =   9
            Animator_25     =   0
            Blinker_25      =   0
            Array[5]_1      =   26
            ClassName_26    =   "CCWDrawObj"
            opts_26         =   62
            Image_26        =   11
            Animator_26     =   0
            Blinker_26      =   0
            Label_1         =   27
            ClassName_27    =   "CCWDrawObj"
            opts_27         =   62
            C[0]_27         =   -2147483640
            Image_27        =   6
            Animator_27     =   0
            Blinker_27      =   0
         End
         Begin CWUIControlsLib.CWNumEdit cwNumFBgain
            Height          =   360
            Index           =   11
            Left            =   1755
            TabIndex        =   270
            Tag             =   "gain0B"
            ToolTipText     =   "Gain of the axis"
            Top             =   4095
            Width           =   1200
            _Version        =   524288
            _ExtentX        =   2117
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
            C[0]_1          =   12648384
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
            Value_Val_1     =   0.71
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
         Begin CWUIControlsLib.CWButton cwBut_AxisEn
            Height          =   360
            Index           =   6
            Left            =   180
            TabIndex        =   271
            Tag             =   "loop_FB_B"
            ToolTipText     =   "Individual axis ON/OFF, when ALL AXES switch is ON"
            Top             =   4095
            Width           =   1455
            _Version        =   524288
            _ExtentX        =   2566
            _ExtentY        =   635
            _StockProps     =   68
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Reset_0         =   0   'False
            CompatibleVers_0=   524288
            Boolean_0       =   1
            ClassName_1     =   "CCWBoolean"
            opts_1          =   2614
            SclRef_1.l      =   2
            SclRef_1.t      =   2
            SclRef_1.r      =   99
            SclRef_1.b      =   26
            Scl_1.l         =   4
            Scl_1.t         =   4
            Scl_1.r         =   69
            Scl_1.b         =   25
            C[0]_1          =   12648384
            Enum_1          =   2
            ClassName_2     =   "CCWEnum"
            Array_2         =   2
            Editor_2        =   0
            Array[0]_2      =   3
            ClassName_3     =   "CCWEnumElt"
            opts_3          =   1
            Name_3          =   "Off"
            frame_3         =   286
            DrawList_3      =   4
            ClassName_4     =   "CDrawList"
            count_4         =   4
            list[4]_4       =   5
            ClassName_5     =   "CCWDrawObj"
            opts_5          =   54
            SclRef_5.l      =   4
            SclRef_5.t      =   4
            SclRef_5.r      =   69
            SclRef_5.b      =   24
            Scl_5.l         =   73
            Scl_5.t         =   5
            Scl_5.r         =   93
            Scl_5.b         =   24
            C[0]_5          =   -2147483640
            C[1]_5          =   -2147483640
            Image_5         =   6
            ClassName_6     =   "CCWTextImage"
            szText_6        =   "tY"
            font_6          =   7
            ClassName_7     =   "CCWFont"
            bFont_7         =   -1  'True
            BeginProperty Font_7 {0BE35203-8F91-11CE-9DE3-00AA004BB851}
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Animator_5      =   0
            Blinker_5       =   0
            list[3]_4       =   8
            ClassName_8     =   "CCWDrawObj"
            opts_8          =   60
            C[0]_8          =   -2147483640
            C[1]_8          =   -2147483640
            Image_8         =   9
            ClassName_9     =   "CCWTextImage"
            font_9          =   0
            Animator_8      =   0
            Blinker_8       =   0
            list[2]_4       =   10
            ClassName_10    =   "CCWDrawObj"
            opts_10         =   60
            C[0]_10         =   -2147483640
            C[1]_10         =   -2147483640
            Image_10        =   11
            ClassName_11    =   "CCWTextImage"
            font_11         =   0
            Animator_10     =   0
            Blinker_10      =   0
            list[1]_4       =   12
            ClassName_12    =   "CCWDrawObj"
            opts_12         =   62
            C[0]_12         =   32768
            C[1]_12         =   65535
            Image_12        =   13
            ClassName_13    =   "CCWPictImage"
            opts_13         =   256
            Rows_13         =   1
            Cols_13         =   1
            URL_13          =   "\\tsclient\C\In_Out\Software\DC2020_WEB_Analyzer\OFF_small_switch.bmp"
            Pict_13         =   "Analyzer.frx":12699A
            ColorReplaceWith_13=   8421504
            ColorReplace_13 =   8421504
            Tolerance_13    =   2
            Animator_12     =   0
            Blinker_12      =   0
            varVarType_3    =   5
            Array[1]_2      =   14
            ClassName_14    =   "CCWEnumElt"
            opts_14         =   1
            Name_14         =   "On"
            frame_14        =   286
            DrawList_14     =   15
            ClassName_15    =   "CDrawList"
            count_15        =   4
            list[4]_15      =   16
            ClassName_16    =   "CCWDrawObj"
            opts_16         =   54
            SclRef_16.l     =   4
            SclRef_16.t     =   4
            SclRef_16.r     =   69
            SclRef_16.b     =   24
            Scl_16.l        =   73
            Scl_16.t        =   5
            Scl_16.r        =   93
            Scl_16.b        =   24
            C[0]_16         =   -2147483640
            C[1]_16         =   -2147483640
            Image_16        =   6
            Animator_16     =   0
            Blinker_16      =   0
            list[3]_15      =   17
            ClassName_17    =   "CCWDrawObj"
            opts_17         =   60
            C[0]_17         =   -2147483640
            C[1]_17         =   -2147483640
            Image_17        =   9
            Animator_17     =   0
            Blinker_17      =   0
            list[2]_15      =   18
            ClassName_18    =   "CCWDrawObj"
            opts_18         =   60
            C[0]_18         =   -2147483640
            C[1]_18         =   -2147483640
            Image_18        =   11
            Animator_18     =   0
            Blinker_18      =   0
            list[1]_15      =   19
            ClassName_19    =   "CCWDrawObj"
            opts_19         =   62
            C[0]_19         =   32768
            C[1]_19         =   32768
            Image_19        =   20
            ClassName_20    =   "CCWPictImage"
            opts_20         =   256
            Rows_20         =   1
            Cols_20         =   1
            URL_20          =   "\\tsclient\C\In_Out\Software\DC2020_WEB_Analyzer\ON_small_switch.bmp"
            Pict_20         =   "Analyzer.frx":128D6C
            ColorReplaceWith_20=   8421504
            ColorReplace_20 =   8421504
            Tolerance_20    =   2
            Animator_19     =   0
            Blinker_19      =   0
            varVarType_14   =   5
            Bindings_1      =   21
            ClassName_21    =   "CCWBindingHolderArray"
            Editor_21       =   22
            ClassName_22    =   "CCWBindingHolderArrayEditor"
            Owner_22        =   1
            Style_1         =   9
            mechAction_1    =   1
            BGImg_1         =   23
            ClassName_23    =   "CCWDrawObj"
            opts_23         =   62
            Image_23        =   24
            ClassName_24    =   "CCWPictImage"
            opts_24         =   1280
            Rows_24         =   1
            Cols_24         =   1
            Pict_24         =   286
            F_24            =   -2147483633
            B_24            =   -2147483633
            ColorReplaceWith_24=   8421504
            ColorReplace_24 =   8421504
            Tolerance_24    =   2
            Animator_23     =   0
            Blinker_23      =   0
            Array_1         =   6
            Editor_1        =   0
            Array[0]_1      =   12
            Array[1]_1      =   19
            Array[2]_1      =   0
            Array[3]_1      =   0
            Array[4]_1      =   25
            ClassName_25    =   "CCWDrawObj"
            opts_25         =   62
            Image_25        =   9
            Animator_25     =   0
            Blinker_25      =   0
            Array[5]_1      =   26
            ClassName_26    =   "CCWDrawObj"
            opts_26         =   62
            Image_26        =   11
            Animator_26     =   0
            Blinker_26      =   0
            Label_1         =   27
            ClassName_27    =   "CCWDrawObj"
            opts_27         =   62
            C[0]_27         =   -2147483640
            Image_27        =   6
            Animator_27     =   0
            Blinker_27      =   0
         End
         Begin CWUIControlsLib.CWButton cwBut_AxisEn
            Height          =   600
            Index           =   0
            Left            =   180
            TabIndex        =   272
            Tag             =   "loop_fba"
            ToolTipText     =   "ALL Velocity FB loops Disable / Enable. If enabled, individual switches control each loop "
            Top             =   450
            Width           =   2310
            _Version        =   524288
            _ExtentX        =   4075
            _ExtentY        =   1058
            _StockProps     =   68
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Reset_0         =   0   'False
            CompatibleVers_0=   524288
            Boolean_0       =   1
            ClassName_1     =   "CCWBoolean"
            opts_1          =   2614
            SclRef_1.l      =   2
            SclRef_1.t      =   2
            SclRef_1.r      =   156
            SclRef_1.b      =   42
            Scl_1.l         =   2
            Scl_1.t         =   5
            Scl_1.r         =   80
            Scl_1.b         =   40
            C[0]_1          =   -2147483643
            Enum_1          =   2
            ClassName_2     =   "CCWEnum"
            Array_2         =   2
            Editor_2        =   0
            Array[0]_2      =   3
            ClassName_3     =   "CCWEnumElt"
            opts_3          =   1
            Name_3          =   "Off"
            frame_3         =   288
            DrawList_3      =   4
            ClassName_4     =   "CDrawList"
            count_4         =   4
            list[4]_4       =   5
            ClassName_5     =   "CCWDrawObj"
            opts_5          =   54
            SclRef_5.l      =   5
            SclRef_5.t      =   6
            SclRef_5.r      =   76
            SclRef_5.b      =   33
            Scl_5.l         =   84
            Scl_5.t         =   12
            Scl_5.r         =   154
            Scl_5.b         =   30
            C[0]_5          =   -2147483640
            C[1]_5          =   -2147483640
            Image_5         =   6
            ClassName_6     =   "CCWTextImage"
            szText_6        =   "ALL AXES"
            font_6          =   0
            Animator_5      =   0
            Blinker_5       =   0
            list[3]_4       =   7
            ClassName_7     =   "CCWDrawObj"
            opts_7          =   60
            C[0]_7          =   -2147483640
            C[1]_7          =   -2147483640
            Image_7         =   8
            ClassName_8     =   "CCWTextImage"
            font_8          =   0
            Animator_7      =   0
            Blinker_7       =   0
            list[2]_4       =   9
            ClassName_9     =   "CCWDrawObj"
            opts_9          =   62
            C[0]_9          =   -2147483640
            C[1]_9          =   -2147483640
            Image_9         =   10
            ClassName_10    =   "CCWTextImage"
            font_10         =   0
            Animator_9      =   0
            Blinker_9       =   0
            list[1]_4       =   11
            ClassName_11    =   "CCWDrawObj"
            opts_11         =   54
            SclRef_11.l     =   6
            SclRef_11.t     =   6
            SclRef_11.r     =   102
            SclRef_11.b     =   51
            Scl_11.l        =   4
            Scl_11.t        =   6
            Scl_11.r        =   102
            Scl_11.b        =   51
            Image_11        =   12
            ClassName_12    =   "CCWPictImage"
            Rows_12         =   1
            Cols_12         =   1
            URL_12          =   "\\tsclient\C\In_Out\Software\DC2020_WEB_Analyzer\OFF_gray_slider.bmp"
            Pict_12         =   "Analyzer.frx":12B13E
            ColorReplaceWith_12=   8421504
            ColorReplace_12 =   8421504
            Tolerance_12    =   2
            Animator_11     =   0
            Blinker_11      =   0
            varVarType_3    =   5
            Array[1]_2      =   13
            ClassName_13    =   "CCWEnumElt"
            opts_13         =   1
            Name_13         =   "On"
            frame_13        =   287
            DrawList_13     =   14
            ClassName_14    =   "CDrawList"
            count_14        =   4
            list[4]_14      =   15
            ClassName_15    =   "CCWDrawObj"
            opts_15         =   54
            SclRef_15.l     =   5
            SclRef_15.t     =   6
            SclRef_15.r     =   76
            SclRef_15.b     =   33
            Scl_15.l        =   84
            Scl_15.t        =   12
            Scl_15.r        =   154
            Scl_15.b        =   30
            C[0]_15         =   -2147483640
            C[1]_15         =   -2147483640
            Image_15        =   6
            Animator_15     =   0
            Blinker_15      =   0
            list[3]_14      =   16
            ClassName_16    =   "CCWDrawObj"
            opts_16         =   62
            C[0]_16         =   -2147483640
            C[1]_16         =   -2147483640
            Image_16        =   8
            Animator_16     =   0
            Blinker_16      =   0
            list[2]_14      =   17
            ClassName_17    =   "CCWDrawObj"
            opts_17         =   60
            C[0]_17         =   -2147483640
            C[1]_17         =   -2147483640
            Image_17        =   10
            Animator_17     =   0
            Blinker_17      =   0
            list[1]_14      =   18
            ClassName_18    =   "CCWDrawObj"
            opts_18         =   54
            SclRef_18.l     =   6
            SclRef_18.t     =   6
            SclRef_18.r     =   102
            SclRef_18.b     =   51
            Scl_18.l        =   4
            Scl_18.t        =   6
            Scl_18.r        =   102
            Scl_18.b        =   51
            Image_18        =   19
            ClassName_19    =   "CCWPictImage"
            Rows_19         =   1
            Cols_19         =   1
            URL_19          =   "\\tsclient\C\In_Out\Software\DC2020_WEB_Analyzer\ON_blue_slider.bmp"
            Pict_19         =   "Analyzer.frx":13EE88
            ColorReplaceWith_19=   8421504
            ColorReplace_19 =   8421504
            Tolerance_19    =   2
            Animator_18     =   0
            Blinker_18      =   0
            varVarType_13   =   5
            Bindings_1      =   20
            ClassName_20    =   "CCWBindingHolderArray"
            Editor_20       =   21
            ClassName_21    =   "CCWBindingHolderArrayEditor"
            Owner_21        =   1
            Style_1         =   14
            drawState_1     =   1
            frameStyle_1    =   1
            mechAction_1    =   1
            BGImg_1         =   22
            ClassName_22    =   "CCWDrawObj"
            opts_22         =   62
            Image_22        =   23
            ClassName_23    =   "CCWPictImage"
            opts_23         =   1280
            Rows_23         =   1
            Cols_23         =   1
            Pict_23         =   286
            F_23            =   -2147483633
            B_23            =   -2147483633
            ColorReplaceWith_23=   8421504
            ColorReplace_23 =   8421504
            Tolerance_23    =   2
            Animator_22     =   0
            Blinker_22      =   0
            Array_1         =   6
            Editor_1        =   0
            Array[0]_1      =   11
            Array[1]_1      =   18
            Array[2]_1      =   0
            Array[3]_1      =   0
            Array[4]_1      =   24
            ClassName_24    =   "CCWDrawObj"
            opts_24         =   62
            Image_24        =   8
            Animator_24     =   0
            Blinker_24      =   0
            Array[5]_1      =   25
            ClassName_25    =   "CCWDrawObj"
            opts_25         =   62
            Image_25        =   10
            Animator_25     =   0
            Blinker_25      =   0
            Label_1         =   26
            ClassName_26    =   "CCWDrawObj"
            opts_26         =   62
            C[0]_26         =   -2147483640
            Image_26        =   6
            Animator_26     =   0
            Blinker_26      =   0
         End
         Begin VB.Label lblAxis
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Velocity"
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
            Height          =   360
            Index           =   1
            Left            =   0
            TabIndex        =   275
            ToolTipText     =   "Suppress payload motion by using geophones (velocity sensors) and linear motors"
            Top             =   0
            Width           =   3120
         End
         Begin VB.Label Label1
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Status                  Gain"
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
            Index           =   13
            Left            =   315
            TabIndex        =   274
            Top             =   1350
            Width           =   2610
         End
         Begin VB.Label Label1
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Feedback Loops"
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
            Index           =   36
            Left            =   405
            TabIndex        =   273
            Top             =   1080
            Width           =   2610
         End
      End
      Begin VB.Frame FramePositionFeedBack
         BackColor       =   &H00E0E0E0&
         Height          =   4635
         Left            =   90
         TabIndex        =   224
         Top             =   1170
         Width           =   3765
         Begin VB.CommandButton CmdZeroHorizProx
            BackColor       =   &H0080C0FF&
            Caption         =   "Reset Horiz Pos"
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
            Left            =   1980
            Style           =   1  'Graphical
            TabIndex        =   232
            ToolTipText     =   "Removes offset from horizontal proximity readings"
            Top             =   45
            Width           =   1680
         End
         Begin VB.Frame FrmVertPosToMotGains
            BackColor       =   &H00FFC0FF&
            BorderStyle     =   0  'None
            Height          =   700
            Left            =   45
            TabIndex        =   225
            Top             =   1935
            Visible         =   0   'False
            Width           =   3660
            Begin CWUIControlsLib.CWNumEdit cwNumFBgain
               Height          =   360
               Index           =   19
               Left            =   45
               TabIndex        =   226
               Tag             =   "gain03"
               ToolTipText     =   "Gain Z prox to Linear Motors"
               Top             =   270
               Width           =   1110
               _Version        =   524288
               _ExtentX        =   1958
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
               C[0]_1          =   16777215
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
               Value_Val_1     =   -10.71
               IncValueVarType_1=   5
               IncValue_Val_1  =   0.1
               AccelIncVarType_1=   5
               AccelInc_Val_1  =   0.5
               RangeMinVarType_1=   5
               RangeMin_Val_1  =   -50
               RangeMaxVarType_1=   5
               RangeMax_Val_1  =   50
               ButtonStyle_1   =   0
               Bindings_1      =   4
               ClassName_4     =   "CCWBindingHolderArray"
               Editor_4        =   5
               ClassName_5     =   "CCWBindingHolderArrayEditor"
               Owner_5         =   1
            End
            Begin CWUIControlsLib.CWNumEdit cwNumFBgain
               Height          =   360
               Index           =   20
               Left            =   1260
               TabIndex        =   227
               Tag             =   "gain04"
               ToolTipText     =   "Gain tX prox to Linear Motors"
               Top             =   270
               Width           =   1110
               _Version        =   524288
               _ExtentX        =   1958
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
               C[0]_1          =   16777215
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
               Value_Val_1     =   49.71
               IncValueVarType_1=   5
               IncValue_Val_1  =   0.1
               AccelIncVarType_1=   5
               AccelInc_Val_1  =   0.5
               RangeMinVarType_1=   5
               RangeMin_Val_1  =   -50
               RangeMaxVarType_1=   5
               RangeMax_Val_1  =   50
               ButtonStyle_1   =   0
               Bindings_1      =   4
               ClassName_4     =   "CCWBindingHolderArray"
               Editor_4        =   5
               ClassName_5     =   "CCWBindingHolderArrayEditor"
               Owner_5         =   1
            End
            Begin CWUIControlsLib.CWNumEdit cwNumFBgain
               Height          =   360
               Index           =   21
               Left            =   2475
               TabIndex        =   228
               Tag             =   "gain05"
               ToolTipText     =   "Gain tY prox to Linear Motors"
               Top             =   270
               Width           =   1110
               _Version        =   524288
               _ExtentX        =   1958
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
               C[0]_1          =   16777215
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
               Value_Val_1     =   0.78
               IncValueVarType_1=   5
               IncValue_Val_1  =   0.1
               AccelIncVarType_1=   5
               AccelInc_Val_1  =   0.5
               RangeMinVarType_1=   5
               RangeMin_Val_1  =   -50
               RangeMaxVarType_1=   5
               RangeMax_Val_1  =   50
               ButtonStyle_1   =   0
               Bindings_1      =   4
               ClassName_4     =   "CCWBindingHolderArray"
               Editor_4        =   5
               ClassName_5     =   "CCWBindingHolderArrayEditor"
               Owner_5         =   1
            End
            Begin VB.Label Label1
               Alignment       =   2  'Center
               BackColor       =   &H00E0E0E0&
               BackStyle       =   0  'Transparent
               Caption         =   "Z"
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
               Index           =   42
               Left            =   315
               TabIndex        =   231
               Top             =   45
               Width           =   720
            End
            Begin VB.Label Label1
               Alignment       =   2  'Center
               BackColor       =   &H00E0E0E0&
               BackStyle       =   0  'Transparent
               Caption         =   "tX"
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
               Index           =   43
               Left            =   1440
               TabIndex        =   230
               Top             =   45
               Width           =   720
            End
            Begin VB.Label Label1
               Alignment       =   2  'Center
               BackColor       =   &H00E0E0E0&
               BackStyle       =   0  'Transparent
               Caption         =   "tY"
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
               Index           =   44
               Left            =   2610
               TabIndex        =   229
               Top             =   45
               Width           =   720
            End
         End
         Begin CWUIControlsLib.CWNumEdit cwNumFBgain
            Height          =   360
            Index           =   0
            Left            =   90
            TabIndex        =   233
            Tag             =   "gain00"
            ToolTipText     =   "Gain of the axis"
            Top             =   1530
            Width           =   1110
            _Version        =   524288
            _ExtentX        =   1958
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
            C[0]_1          =   16777215
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
            Value_Val_1     =   -10.71
            IncValueVarType_1=   5
            IncValue_Val_1  =   0.1
            AccelIncVarType_1=   5
            AccelInc_Val_1  =   0.5
            RangeMinVarType_1=   5
            RangeMin_Val_1  =   -50
            RangeMaxVarType_1=   5
            RangeMax_Val_1  =   50
            ButtonStyle_1   =   0
            Bindings_1      =   4
            ClassName_4     =   "CCWBindingHolderArray"
            Editor_4        =   5
            ClassName_5     =   "CCWBindingHolderArrayEditor"
            Owner_5         =   1
         End
         Begin CWUIControlsLib.CWNumEdit cwNumFBgain
            Height          =   360
            Index           =   1
            Left            =   1305
            TabIndex        =   234
            Tag             =   "gain01"
            ToolTipText     =   "Gain of the axis"
            Top             =   1530
            Width           =   1110
            _Version        =   524288
            _ExtentX        =   1958
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
            C[0]_1          =   16777215
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
            Value_Val_1     =   49.71
            IncValueVarType_1=   5
            IncValue_Val_1  =   0.1
            AccelIncVarType_1=   5
            AccelInc_Val_1  =   0.5
            RangeMinVarType_1=   5
            RangeMin_Val_1  =   -50
            RangeMaxVarType_1=   5
            RangeMax_Val_1  =   50
            ButtonStyle_1   =   0
            Bindings_1      =   4
            ClassName_4     =   "CCWBindingHolderArray"
            Editor_4        =   5
            ClassName_5     =   "CCWBindingHolderArrayEditor"
            Owner_5         =   1
         End
         Begin CWUIControlsLib.CWNumEdit cwNumFBgain
            Height          =   360
            Index           =   2
            Left            =   2520
            TabIndex        =   235
            Tag             =   "gain02"
            ToolTipText     =   "Gain of the axis"
            Top             =   1530
            Width           =   1110
            _Version        =   524288
            _ExtentX        =   1958
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
            C[0]_1          =   16777215
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
            Value_Val_1     =   0.78
            IncValueVarType_1=   5
            IncValue_Val_1  =   0.1
            AccelIncVarType_1=   5
            AccelInc_Val_1  =   0.5
            RangeMinVarType_1=   5
            RangeMin_Val_1  =   -50
            RangeMaxVarType_1=   5
            RangeMax_Val_1  =   50
            ButtonStyle_1   =   0
            Bindings_1      =   4
            ClassName_4     =   "CCWBindingHolderArray"
            Editor_4        =   5
            ClassName_5     =   "CCWBindingHolderArrayEditor"
            Owner_5         =   1
         End
         Begin CWUIControlsLib.CWButton cw_Status_Bit_Indicator
            Height          =   600
            Index           =   5
            Left            =   2700
            TabIndex        =   236
            ToolTipText     =   "Shows status of Horizontal Holding (""Soft Dock"") by linear motors"
            Top             =   405
            Width           =   975
            _Version        =   524288
            _ExtentX        =   1720
            _ExtentY        =   1058
            _StockProps     =   68
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Reset_0         =   0   'False
            CompatibleVers_0=   524288
            Boolean_0       =   1
            ClassName_1     =   "CCWBoolean"
            opts_1          =   2622
            C[0]_1          =   16761087
            Enum_1          =   2
            ClassName_2     =   "CCWEnum"
            Array_2         =   2
            Editor_2        =   0
            Array[0]_2      =   3
            ClassName_3     =   "CCWEnumElt"
            opts_3          =   1
            Name_3          =   "Off"
            frame_3         =   286
            DrawList_3      =   4
            ClassName_4     =   "CDrawList"
            count_4         =   4
            list[4]_4       =   5
            ClassName_5     =   "CCWDrawObj"
            opts_5          =   62
            C[0]_5          =   -2147483640
            C[1]_5          =   -2147483640
            Image_5         =   6
            ClassName_6     =   "CCWTextImage"
            szText_6        =   "Horiz Pos"
            font_6          =   0
            Animator_5      =   0
            Blinker_5       =   0
            list[3]_4       =   7
            ClassName_7     =   "CCWDrawObj"
            opts_7          =   60
            C[0]_7          =   -2147483640
            C[1]_7          =   -2147483640
            Image_7         =   8
            ClassName_8     =   "CCWTextImage"
            szText_8        =   "Holding"
            font_8          =   0
            Animator_7      =   0
            Blinker_7       =   0
            list[2]_4       =   9
            ClassName_9     =   "CCWDrawObj"
            opts_9          =   62
            C[0]_9          =   -2147483640
            C[1]_9          =   -2147483640
            Image_9         =   10
            ClassName_10    =   "CCWTextImage"
            szText_10       =   "FREE"
            font_10         =   0
            Animator_9      =   0
            Blinker_9       =   0
            list[1]_4       =   11
            ClassName_11    =   "CCWDrawObj"
            opts_11         =   62
            C[0]_11         =   65280
            C[1]_11         =   65280
            Image_11        =   12
            ClassName_12    =   "CCWPictImage"
            opts_12         =   1280
            Rows_12         =   1
            Cols_12         =   1
            Pict_12         =   73
            F_12            =   65280
            B_12            =   65280
            ColorReplaceWith_12=   8421504
            ColorReplace_12 =   8421504
            Tolerance_12    =   2
            Animator_11     =   0
            Blinker_11      =   0
            varVarType_3    =   5
            Array[1]_2      =   13
            ClassName_13    =   "CCWEnumElt"
            opts_13         =   1
            Name_13         =   "On"
            frame_13        =   286
            DrawList_13     =   14
            ClassName_14    =   "CDrawList"
            count_14        =   4
            list[4]_14      =   15
            ClassName_15    =   "CCWDrawObj"
            opts_15         =   62
            C[0]_15         =   -2147483640
            C[1]_15         =   -2147483640
            Image_15        =   6
            Animator_15     =   0
            Blinker_15      =   0
            list[3]_14      =   16
            ClassName_16    =   "CCWDrawObj"
            opts_16         =   62
            C[0]_16         =   -2147483640
            C[1]_16         =   -2147483640
            Image_16        =   8
            Animator_16     =   0
            Blinker_16      =   0
            list[2]_14      =   17
            ClassName_17    =   "CCWDrawObj"
            opts_17         =   60
            C[0]_17         =   -2147483640
            C[1]_17         =   -2147483640
            Image_17        =   10
            Animator_17     =   0
            Blinker_17      =   0
            list[1]_14      =   18
            ClassName_18    =   "CCWDrawObj"
            opts_18         =   62
            C[0]_18         =   16711935
            C[1]_18         =   16711935
            Image_18        =   19
            ClassName_19    =   "CCWPictImage"
            opts_19         =   1280
            Rows_19         =   1
            Cols_19         =   1
            Pict_19         =   74
            F_19            =   16711935
            B_19            =   16711935
            ColorReplaceWith_19=   8421504
            ColorReplace_19 =   8421504
            Tolerance_19    =   2
            Animator_18     =   0
            Blinker_18      =   20
            ClassName_20    =   "CCWBlinker"
            Interval_20     =   600
            EventID_20      =   7
            Obj_20          =   18
            varVarType_13   =   5
            Bindings_1      =   21
            ClassName_21    =   "CCWBindingHolderArray"
            Editor_21       =   22
            ClassName_22    =   "CCWBindingHolderArrayEditor"
            Owner_22        =   1
            Style_1         =   4
            mechAction_1    =   3
            BGImg_1         =   23
            ClassName_23    =   "CCWDrawObj"
            opts_23         =   62
            Image_23        =   24
            ClassName_24    =   "CCWPictImage"
            opts_24         =   1280
            Rows_24         =   1
            Cols_24         =   1
            Pict_24         =   286
            F_24            =   -2147483633
            B_24            =   -2147483633
            ColorReplaceWith_24=   8421504
            ColorReplace_24 =   8421504
            Tolerance_24    =   2
            Animator_23     =   0
            Blinker_23      =   0
            Array_1         =   6
            Editor_1        =   0
            Array[0]_1      =   11
            Array[1]_1      =   18
            Array[2]_1      =   0
            Array[3]_1      =   0
            Array[4]_1      =   25
            ClassName_25    =   "CCWDrawObj"
            opts_25         =   62
            Image_25        =   8
            Animator_25     =   0
            Blinker_25      =   0
            Array[5]_1      =   26
            ClassName_26    =   "CCWDrawObj"
            opts_26         =   62
            Image_26        =   10
            Animator_26     =   0
            Blinker_26      =   0
            Label_1         =   27
            ClassName_27    =   "CCWDrawObj"
            opts_27         =   62
            C[0]_27         =   -2147483640
            Image_27        =   6
            Animator_27     =   0
            Blinker_27      =   0
         End
         Begin CWUIControlsLib.CWButton cwBut_SoftDock
            Height          =   600
            Left            =   45
            TabIndex        =   237
            Tag             =   "horz"
            ToolTipText     =   "Soft dock: Hold payload by using horizontal proximity sensors and linear motors"
            Top             =   405
            Width           =   2400
            _Version        =   524288
            _ExtentX        =   4233
            _ExtentY        =   1058
            _StockProps     =   68
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Reset_0         =   0   'False
            CompatibleVers_0=   524288
            Boolean_0       =   1
            ClassName_1     =   "CCWBoolean"
            opts_1          =   2614
            SclRef_1.l      =   2
            SclRef_1.t      =   2
            SclRef_1.r      =   162
            SclRef_1.b      =   42
            Scl_1.l         =   5
            Scl_1.t         =   6
            Scl_1.r         =   80
            Scl_1.b         =   39
            C[0]_1          =   16761087
            Enum_1          =   2
            ClassName_2     =   "CCWEnum"
            Array_2         =   2
            Editor_2        =   0
            Array[0]_2      =   3
            ClassName_3     =   "CCWEnumElt"
            opts_3          =   1
            Name_3          =   "Off"
            frame_3         =   288
            DrawList_3      =   4
            ClassName_4     =   "CDrawList"
            count_4         =   4
            list[4]_4       =   5
            ClassName_5     =   "CCWDrawObj"
            opts_5          =   54
            SclRef_5.l      =   6
            SclRef_5.t      =   8
            SclRef_5.r      =   74
            SclRef_5.b      =   34
            Scl_5.l         =   85
            Scl_5.t         =   6
            Scl_5.r         =   154
            Scl_5.b         =   24
            C[0]_5          =   -2147483640
            C[1]_5          =   -2147483640
            Image_5         =   6
            ClassName_6     =   "CCWTextImage"
            szText_6        =   "Soft Dock"
            font_6          =   7
            ClassName_7     =   "CCWFont"
            bFont_7         =   -1  'True
            BeginProperty Font_7 {0BE35203-8F91-11CE-9DE3-00AA004BB851}
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Animator_5      =   0
            Blinker_5       =   0
            list[3]_4       =   8
            ClassName_8     =   "CCWDrawObj"
            opts_8          =   52
            SclRef_8.l      =   7
            SclRef_8.t      =   9
            SclRef_8.r      =   75
            SclRef_8.b      =   35
            Scl_8.l         =   85
            Scl_8.t         =   25
            Scl_8.r         =   157
            Scl_8.b         =   38
            C[0]_8          =   -2147483640
            C[1]_8          =   -2147483640
            Image_8         =   9
            ClassName_9     =   "CCWTextImage"
            szText_9        =   "Click to release"
            font_9          =   0
            Animator_8      =   0
            Blinker_8       =   0
            list[2]_4       =   10
            ClassName_10    =   "CCWDrawObj"
            opts_10         =   54
            SclRef_10.l     =   6
            SclRef_10.t     =   8
            SclRef_10.r     =   74
            SclRef_10.b     =   34
            Scl_10.l        =   85
            Scl_10.t        =   25
            Scl_10.r        =   158
            Scl_10.b        =   38
            C[0]_10         =   -2147483640
            C[1]_10         =   -2147483640
            Image_10        =   11
            ClassName_11    =   "CCWTextImage"
            szText_11       =   "Click to engage"
            font_11         =   0
            Animator_10     =   0
            Blinker_10      =   0
            list[1]_4       =   12
            ClassName_12    =   "CCWDrawObj"
            opts_12         =   54
            SclRef_12.l     =   6
            SclRef_12.t     =   6
            SclRef_12.r     =   102
            SclRef_12.b     =   51
            Scl_12.l        =   4
            Scl_12.t        =   6
            Scl_12.r        =   102
            Scl_12.b        =   51
            Image_12        =   13
            ClassName_13    =   "CCWPictImage"
            Rows_13         =   1
            Cols_13         =   1
            URL_13          =   "\\tsclient\C\In_Out\Software\DC2020_WEB_Analyzer\OFF_gray_slider.bmp"
            Pict_13         =   "Analyzer.frx":152EE2
            ColorReplaceWith_13=   8421504
            ColorReplace_13 =   8421504
            Tolerance_13    =   2
            Animator_12     =   0
            Blinker_12      =   0
            varVarType_3    =   5
            Array[1]_2      =   14
            ClassName_14    =   "CCWEnumElt"
            opts_14         =   1
            Name_14         =   "On"
            frame_14        =   287
            DrawList_14     =   15
            ClassName_15    =   "CDrawList"
            count_15        =   4
            list[4]_15      =   16
            ClassName_16    =   "CCWDrawObj"
            opts_16         =   54
            SclRef_16.l     =   6
            SclRef_16.t     =   8
            SclRef_16.r     =   74
            SclRef_16.b     =   34
            Scl_16.l        =   85
            Scl_16.t        =   6
            Scl_16.r        =   154
            Scl_16.b        =   24
            C[0]_16         =   -2147483640
            C[1]_16         =   -2147483640
            Image_16        =   6
            Animator_16     =   0
            Blinker_16      =   0
            list[3]_15      =   17
            ClassName_17    =   "CCWDrawObj"
            opts_17         =   54
            SclRef_17.l     =   7
            SclRef_17.t     =   9
            SclRef_17.r     =   75
            SclRef_17.b     =   35
            Scl_17.l        =   85
            Scl_17.t        =   25
            Scl_17.r        =   157
            Scl_17.b        =   38
            C[0]_17         =   -2147483640
            C[1]_17         =   -2147483640
            Image_17        =   9
            Animator_17     =   0
            Blinker_17      =   0
            list[2]_15      =   18
            ClassName_18    =   "CCWDrawObj"
            opts_18         =   52
            SclRef_18.l     =   6
            SclRef_18.t     =   8
            SclRef_18.r     =   74
            SclRef_18.b     =   34
            Scl_18.l        =   85
            Scl_18.t        =   25
            Scl_18.r        =   158
            Scl_18.b        =   38
            C[0]_18         =   -2147483640
            C[1]_18         =   -2147483640
            Image_18        =   11
            Animator_18     =   0
            Blinker_18      =   0
            list[1]_15      =   19
            ClassName_19    =   "CCWDrawObj"
            opts_19         =   54
            SclRef_19.l     =   6
            SclRef_19.t     =   6
            SclRef_19.r     =   102
            SclRef_19.b     =   51
            Scl_19.l        =   4
            Scl_19.t        =   6
            Scl_19.r        =   102
            Scl_19.b        =   51
            Image_19        =   20
            ClassName_20    =   "CCWPictImage"
            Rows_20         =   1
            Cols_20         =   1
            URL_20          =   "\\tsclient\C\In_Out\Software\DC2020_WEB_Analyzer\ON_small_switch.bmp"
            Pict_20         =   "Analyzer.frx":166C2C
            ColorReplaceWith_20=   8421504
            ColorReplace_20 =   8421504
            Tolerance_20    =   2
            Animator_19     =   0
            Blinker_19      =   0
            varVarType_14   =   5
            Bindings_1      =   21
            ClassName_21    =   "CCWBindingHolderArray"
            Editor_21       =   22
            ClassName_22    =   "CCWBindingHolderArrayEditor"
            Owner_22        =   1
            Style_1         =   14
            frameStyle_1    =   1
            mechAction_1    =   1
            BGImg_1         =   23
            ClassName_23    =   "CCWDrawObj"
            opts_23         =   62
            Image_23        =   24
            ClassName_24    =   "CCWPictImage"
            opts_24         =   1280
            Rows_24         =   1
            Cols_24         =   1
            Pict_24         =   286
            F_24            =   -2147483633
            B_24            =   -2147483633
            ColorReplaceWith_24=   8421504
            ColorReplace_24 =   8421504
            Tolerance_24    =   2
            Animator_23     =   0
            Blinker_23      =   0
            Array_1         =   6
            Editor_1        =   0
            Array[0]_1      =   12
            Array[1]_1      =   19
            Array[2]_1      =   0
            Array[3]_1      =   0
            Array[4]_1      =   25
            ClassName_25    =   "CCWDrawObj"
            opts_25         =   62
            Image_25        =   9
            Animator_25     =   0
            Blinker_25      =   0
            Array[5]_1      =   26
            ClassName_26    =   "CCWDrawObj"
            opts_26         =   62
            Image_26        =   11
            Animator_26     =   0
            Blinker_26      =   0
            Label_1         =   27
            ClassName_27    =   "CCWDrawObj"
            opts_27         =   62
            C[0]_27         =   -2147483640
            Image_27        =   6
            Animator_27     =   0
            Blinker_27      =   0
         End
         Begin CWUIControlsLib.CWNumEdit cwNumHeightAdj
            Height          =   360
            Index           =   0
            Left            =   135
            TabIndex        =   238
            Tag             =   "ofs1"
            ToolTipText     =   "Vertical correction of Iso #1 (default 0, range +-1000 for Turck sensors, +-4000 for AutomationDirect sensors)"
            Top             =   4185
            Width           =   1020
            _Version        =   524288
            _ExtentX        =   1799
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
            C[0]_1          =   16777215
            C[1]_1          =   0
            BorderStyle_1   =   1
            TextAlignment_1 =   1
            format_1        =   2
            ClassName_2     =   "CCWFormat"
            scale_1         =   3
            ClassName_3     =   "CCWScale"
            opts_3          =   65536
            dMax_3          =   10
            discInterval_3  =   1
            ValueVarType_1  =   5
            Value_Val_1     =   -3234
            IncValueVarType_1=   5
            IncValue_Val_1  =   10
            AccelIncVarType_1=   5
            AccelInc_Val_1  =   100
            RangeMinVarType_1=   5
            RangeMin_Val_1  =   -4000
            RangeMaxVarType_1=   5
            RangeMax_Val_1  =   4000
            ButtonStyle_1   =   0
            Bindings_1      =   4
            ClassName_4     =   "CCWBindingHolderArray"
            Editor_4        =   5
            ClassName_5     =   "CCWBindingHolderArrayEditor"
            Owner_5         =   1
         End
         Begin CWUIControlsLib.CWNumEdit cwNumHeightAdj
            Height          =   360
            Index           =   1
            Left            =   1305
            TabIndex        =   239
            Tag             =   "ofs2"
            ToolTipText     =   "Vertical correction of Iso #2 (default 0, range +-1000 for Turck sensors, +-4000 for AutomationDirect sensors)"
            Top             =   4185
            Width           =   1020
            _Version        =   524288
            _ExtentX        =   1799
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
            C[0]_1          =   16777215
            C[1]_1          =   0
            BorderStyle_1   =   1
            TextAlignment_1 =   1
            format_1        =   2
            ClassName_2     =   "CCWFormat"
            scale_1         =   3
            ClassName_3     =   "CCWScale"
            opts_3          =   65536
            dMax_3          =   10
            discInterval_3  =   1
            ValueVarType_1  =   5
            IncValueVarType_1=   5
            IncValue_Val_1  =   10
            AccelIncVarType_1=   5
            AccelInc_Val_1  =   100
            RangeMinVarType_1=   5
            RangeMin_Val_1  =   -4000
            RangeMaxVarType_1=   5
            RangeMax_Val_1  =   4000
            ButtonStyle_1   =   0
            Bindings_1      =   4
            ClassName_4     =   "CCWBindingHolderArray"
            Editor_4        =   5
            ClassName_5     =   "CCWBindingHolderArrayEditor"
            Owner_5         =   1
         End
         Begin CWUIControlsLib.CWNumEdit cwNumHeightAdj
            Height          =   360
            Index           =   2
            Left            =   2475
            TabIndex        =   240
            Tag             =   "ofs3"
            ToolTipText     =   "Vertical correction of Iso #3 (default 0, range +-1000 for Turck sensors, +-4000 for AutomationDirect sensors)"
            Top             =   4185
            Width           =   1020
            _Version        =   524288
            _ExtentX        =   1799
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
            C[0]_1          =   16777215
            C[1]_1          =   0
            BorderStyle_1   =   1
            TextAlignment_1 =   1
            format_1        =   2
            ClassName_2     =   "CCWFormat"
            scale_1         =   3
            ClassName_3     =   "CCWScale"
            opts_3          =   65536
            dMax_3          =   10
            discInterval_3  =   1
            ValueVarType_1  =   5
            Value_Val_1     =   2345
            IncValueVarType_1=   5
            IncValue_Val_1  =   10
            AccelIncVarType_1=   5
            AccelInc_Val_1  =   100
            RangeMinVarType_1=   5
            RangeMin_Val_1  =   -4000
            RangeMaxVarType_1=   5
            RangeMax_Val_1  =   4000
            ButtonStyle_1   =   0
            Bindings_1      =   4
            ClassName_4     =   "CCWBindingHolderArray"
            Editor_4        =   5
            ClassName_5     =   "CCWBindingHolderArrayEditor"
            Owner_5         =   1
         End
         Begin CWUIControlsLib.CWNumEdit cwNum_Go_NoGO_window
            Height          =   360
            Index           =   0
            Left            =   75
            TabIndex        =   241
            Tag             =   "gngx"
            ToolTipText     =   "X direction Go-NoGo window, range 0 to 1000 um"
            Top             =   3150
            Width           =   885
            _Version        =   524288
            _ExtentX        =   1561
            _ExtentY        =   635
            _StockProps     =   4
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}
               Name            =   "Arial"
               Size            =   9.75
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
            BorderStyle_1   =   1
            TextAlignment_1 =   2
            format_1        =   2
            ClassName_2     =   "CCWFormat"
            scale_1         =   3
            ClassName_3     =   "CCWScale"
            dMax_3          =   10
            discInterval_3  =   1
            ValueVarType_1  =   5
            Value_Val_1     =   1000
            IncValueVarType_1=   5
            IncValue_Val_1  =   10
            AccelIncVarType_1=   5
            AccelInc_Val_1  =   100
            RangeMinVarType_1=   5
            RangeMaxVarType_1=   5
            RangeMax_Val_1  =   1000
            ButtonStyle_1   =   0
            Bindings_1      =   4
            ClassName_4     =   "CCWBindingHolderArray"
            Editor_4        =   5
            ClassName_5     =   "CCWBindingHolderArrayEditor"
            Owner_5         =   1
         End
         Begin CWUIControlsLib.CWNumEdit cwNum_Go_NoGO_window
            Height          =   360
            Index           =   1
            Left            =   990
            TabIndex        =   242
            Tag             =   "gngy"
            ToolTipText     =   "Y direction Go-NoGo window, range 0 to 1000 um"
            Top             =   3150
            Width           =   885
            _Version        =   524288
            _ExtentX        =   1561
            _ExtentY        =   635
            _StockProps     =   4
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}
               Name            =   "Arial"
               Size            =   9.75
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
            BorderStyle_1   =   1
            TextAlignment_1 =   2
            format_1        =   2
            ClassName_2     =   "CCWFormat"
            scale_1         =   3
            ClassName_3     =   "CCWScale"
            dMax_3          =   10
            discInterval_3  =   1
            ValueVarType_1  =   5
            Value_Val_1     =   99
            IncValueVarType_1=   5
            IncValue_Val_1  =   10
            AccelIncVarType_1=   5
            AccelInc_Val_1  =   100
            RangeMinVarType_1=   5
            RangeMaxVarType_1=   5
            RangeMax_Val_1  =   1000
            ButtonStyle_1   =   0
            Bindings_1      =   4
            ClassName_4     =   "CCWBindingHolderArray"
            Editor_4        =   5
            ClassName_5     =   "CCWBindingHolderArrayEditor"
            Owner_5         =   1
         End
         Begin CWUIControlsLib.CWNumEdit cwNum_Go_NoGO_window
            Height          =   360
            Index           =   2
            Left            =   1905
            TabIndex        =   243
            Tag             =   "gngz"
            ToolTipText     =   "Vertical Go-NoGo window, range 0 to 1000 um"
            Top             =   3150
            Width           =   885
            _Version        =   524288
            _ExtentX        =   1561
            _ExtentY        =   635
            _StockProps     =   4
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}
               Name            =   "Arial"
               Size            =   9.75
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
            BorderStyle_1   =   1
            TextAlignment_1 =   2
            format_1        =   2
            ClassName_2     =   "CCWFormat"
            scale_1         =   3
            ClassName_3     =   "CCWScale"
            dMax_3          =   10
            discInterval_3  =   1
            ValueVarType_1  =   5
            Value_Val_1     =   122
            IncValueVarType_1=   5
            IncValue_Val_1  =   10
            AccelIncVarType_1=   5
            AccelInc_Val_1  =   100
            RangeMinVarType_1=   5
            RangeMaxVarType_1=   5
            RangeMax_Val_1  =   1000
            ButtonStyle_1   =   0
            Bindings_1      =   4
            ClassName_4     =   "CCWBindingHolderArray"
            Editor_4        =   5
            ClassName_5     =   "CCWBindingHolderArrayEditor"
            Owner_5         =   1
         End
         Begin CWUIControlsLib.CWNumEdit cwNum_Go_NoGO_window
            Height          =   360
            Index           =   3
            Left            =   2820
            TabIndex        =   244
            Tag             =   "gngd"
            ToolTipText     =   "Time Go-NoGo window, range 0 to 10 s"
            Top             =   3150
            Width           =   885
            _Version        =   524288
            _ExtentX        =   1561
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
            C[0]_1          =   16777215
            C[1]_1          =   0
            BorderStyle_1   =   1
            TextAlignment_1 =   2
            format_1        =   2
            ClassName_2     =   "CCWFormat"
            scale_1         =   3
            ClassName_3     =   "CCWScale"
            dMax_3          =   10
            discInterval_3  =   0.01
            ValueVarType_1  =   5
            Value_Val_1     =   0.75
            IncValueVarType_1=   5
            IncValue_Val_1  =   0.05
            AccelIncVarType_1=   5
            AccelInc_Val_1  =   0.5
            RangeMinVarType_1=   5
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
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "tZ"
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
            Index           =   14
            Left            =   2655
            TabIndex        =   257
            Top             =   1305
            Width           =   720
         End
         Begin VB.Label Label1
            Alignment       =   2  'Center
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Y"
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
            Index           =   15
            Left            =   1485
            TabIndex        =   256
            Top             =   1305
            Width           =   720
         End
         Begin VB.Label Label1
            Alignment       =   2  'Center
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "X"
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
            Index           =   16
            Left            =   360
            TabIndex        =   255
            Top             =   1305
            Width           =   720
         End
         Begin VB.Label lblAxis
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "         Position"
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
            Height          =   400
            Index           =   2
            Left            =   0
            TabIndex        =   254
            Top             =   0
            Width           =   3750
         End
         Begin VB.Label Label1
            Alignment       =   2  'Center
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Isolators Vertical offset, um"
            BeginProperty Font
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   25
            Left            =   45
            TabIndex        =   253
            Top             =   3645
            Width           =   3600
         End
         Begin VB.Label Label1
            Alignment       =   2  'Center
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Iso 1"
            BeginProperty Font
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   26
            Left            =   315
            TabIndex        =   252
            Top             =   3915
            Width           =   720
         End
         Begin VB.Label Label1
            Alignment       =   2  'Center
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Iso 2"
            BeginProperty Font
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   29
            Left            =   1575
            TabIndex        =   251
            Top             =   3915
            Width           =   720
         End
         Begin VB.Label Label1
            Alignment       =   2  'Center
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Iso 3"
            BeginProperty Font
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   30
            Left            =   2745
            TabIndex        =   250
            Top             =   3915
            Width           =   720
         End
         Begin VB.Label Label1
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFC0&
            Caption         =   "Go-NoGo window"
            BeginProperty Font
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   885
            Index           =   31
            Left            =   45
            TabIndex        =   249
            Top             =   2700
            Width           =   3675
         End
         Begin VB.Label Label1
            Alignment       =   2  'Center
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "X um"
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
            Index           =   32
            Left            =   135
            TabIndex        =   248
            Top             =   2925
            Width           =   720
         End
         Begin VB.Label Label1
            Alignment       =   2  'Center
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Y um"
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
            Index           =   33
            Left            =   1080
            TabIndex        =   247
            Top             =   2925
            Width           =   720
         End
         Begin VB.Label Label1
            Alignment       =   2  'Center
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Z um"
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
            Index           =   34
            Left            =   1980
            TabIndex        =   246
            Top             =   2925
            Width           =   720
         End
         Begin VB.Label Label1
            Alignment       =   2  'Center
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Time s"
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
            Index           =   35
            Left            =   2880
            TabIndex        =   245
            Top             =   2925
            Width           =   720
         End
         Begin VB.Label Label1
            Alignment       =   2  'Center
            BackColor       =   &H00FFC0FF&
            Caption         =   "Soft Dock Gains"
            BeginProperty Font
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   885
            Index           =   19
            Left            =   45
            TabIndex        =   258
            Top             =   1080
            Width           =   3645
         End
      End
      Begin VB.TextBox txtBootLoader_Rev
         Appearance      =   0  'Flat
         BackColor       =   &H00E0FFFF&
         BeginProperty Font
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000012&
         Height          =   330
         Left            =   10215
         TabIndex        =   221
         Text            =   " NO BootLoader"
         ToolTipText     =   "Bootloader revision, if installed"
         Top             =   705
         Width           =   1410
      End
      Begin VB.TextBox TxtSystemID
         BeginProperty Font
            Name            =   "Courier New"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3975
         TabIndex        =   213
         Text            =   "Enter System ID67890"
         Top             =   675
         Width           =   3180
      End
      Begin VB.CommandButton CmdUpdateSysID
         BackColor       =   &H00C0FFC0&
         Caption         =   "Update System ID"
         Height          =   355
         Left            =   7290
         Style           =   1  'Graphical
         TabIndex        =   212
         ToolTipText     =   "Sends System ID (from the left) to controller"
         Top             =   675
         Width           =   1500
      End
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
         ForeColor       =   &H80000012&
         Height          =   300
         Left            =   3015
         TabIndex        =   4
         Text            =   "?FW version"
         ToolTipText     =   "Firmware Version"
         Top             =   270
         Width           =   8625
      End
      Begin VB.CommandButton cmdRefresh
         BackColor       =   &H00FF80FF&
         Caption         =   "REFRESH"
         BeginProperty Font
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   800
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   270
         Width           =   1100
      End
      Begin VB.Label Label1
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0FFFF&
         Caption         =   "System Identifier"
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
         Index           =   41
         Left            =   2115
         TabIndex        =   223
         Top             =   720
         Width           =   1740
      End
      Begin VB.Label Label1
         BackColor       =   &H00E0FFFF&
         Caption         =   "Bootloader Rev"
         BeginProperty Font
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   40
         Left            =   8910
         TabIndex        =   222
         Top             =   720
         Width           =   1380
      End
      Begin VB.Image ImageSettings
         Height          =   600
         Left            =   120
         Picture         =   "Analyzer.frx":168FFE
         Stretch         =   -1  'True
         Top             =   315
         Width           =   600
      End
      Begin VB.Label Label1
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
         Left            =   2000
         TabIndex        =   5
         Top             =   270
         Width           =   1155
      End
   End
   Begin VB.Frame FrmConnection
      BackColor       =   &H00C0FFC0&
      Caption         =   "Communication"
      BeginProperty Font
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5970
      Left            =   2745
      TabIndex        =   2
      Top             =   1661
      Width           =   11756
      Begin VB.CheckBox ChkCommIssues
         Caption         =   "Show Comm Issues"
         Height          =   600
         Left            =   135
         Style           =   1  'Graphical
         TabIndex        =   195
         ToolTipText     =   "This list is updated if command takes more than 80 ms (typical response 10 ms)"
         Top             =   3015
         Width           =   1335
      End
      Begin VB.ListBox ListCommIssues
         BeginProperty Font
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2085
         ItemData        =   "Analyzer.frx":18B480
         Left            =   135
         List            =   "Analyzer.frx":18B482
         TabIndex        =   194
         ToolTipText     =   "This list is updated if command takes more than 80 ms (typical response 10 ms)"
         Top             =   3780
         Visible         =   0   'False
         Width           =   11520
      End
      Begin VB.CommandButton cmdCommSETUP
         BackColor       =   &H00C0E0FF&
         Caption         =   "Serial port setup"
         Height          =   600
         Left            =   135
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Access to Login screen, selection of COM port setup or Demo Mode"
         Top             =   2115
         Width           =   1335
      End
      Begin CWUIControlsLib.CWButton cwCOMdisconnect
         Height          =   660
         Left            =   135
         TabIndex        =   18
         ToolTipText     =   "Temporary disconnects COM port"
         Top             =   1170
         Width           =   1320
         _Version        =   524288
         _ExtentX        =   2328
         _ExtentY        =   1164
         _StockProps     =   68
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Reset_0         =   0   'False
         CompatibleVers_0=   524288
         Boolean_0       =   1
         ClassName_1     =   "CCWBoolean"
         opts_1          =   2614
         SclRef_1.l      =   2
         SclRef_1.t      =   2
         SclRef_1.r      =   63
         SclRef_1.b      =   43
         Scl_1.l         =   2
         Scl_1.t         =   17
         Scl_1.r         =   63
         Scl_1.b         =   44
         C[0]_1          =   14737632
         Enum_1          =   2
         ClassName_2     =   "CCWEnum"
         Array_2         =   2
         Editor_2        =   0
         Array[0]_2      =   3
         ClassName_3     =   "CCWEnumElt"
         opts_3          =   1
         Name_3          =   "Off"
         frame_3         =   286
         DrawList_3      =   4
         ClassName_4     =   "CDrawList"
         count_4         =   4
         list[4]_4       =   5
         ClassName_5     =   "CCWDrawObj"
         opts_5          =   62
         C[0]_5          =   -2147483641
         C[1]_5          =   -2147483641
         Image_5         =   6
         ClassName_6     =   "CCWTextImage"
         szText_6        =   "Serial port"
         font_6          =   0
         Animator_5      =   0
         Blinker_5       =   0
         list[3]_4       =   7
         ClassName_7     =   "CCWDrawObj"
         opts_7          =   60
         C[0]_7          =   0
         C[1]_7          =   0
         Image_7         =   8
         ClassName_8     =   "CCWTextImage"
         szText_8        =   "Connected"
         font_8          =   0
         Animator_7      =   0
         Blinker_7       =   0
         list[2]_4       =   9
         ClassName_9     =   "CCWDrawObj"
         opts_9          =   54
         SclRef_9.l      =   2
         SclRef_9.t      =   17
         SclRef_9.r      =   72
         SclRef_9.b      =   44
         Scl_9.l         =   10
         Scl_9.t         =   26
         Scl_9.r         =   63
         Scl_9.b         =   40
         C[0]_9          =   16777215
         C[1]_9          =   16777215
         Image_9         =   10
         ClassName_10    =   "CCWTextImage"
         szText_10       =   "DISconnected"
         font_10         =   0
         Animator_9      =   0
         Blinker_9       =   0
         list[1]_4       =   11
         ClassName_11    =   "CCWDrawObj"
         opts_11         =   54
         SclRef_11.l     =   2
         SclRef_11.t     =   17
         SclRef_11.r     =   63
         SclRef_11.b     =   44
         Scl_11.l        =   2
         Scl_11.t        =   15
         Scl_11.r        =   63
         Scl_11.b        =   42
         C[0]_11         =   255
         C[1]_11         =   255
         Image_11        =   12
         ClassName_12    =   "CCWPictImage"
         opts_12         =   1280
         Rows_12         =   1
         Cols_12         =   1
         Pict_12         =   44
         F_12            =   255
         B_12            =   255
         ColorReplaceWith_12=   8421504
         ColorReplace_12 =   8421504
         Tolerance_12    =   2
         Animator_11     =   0
         Blinker_11      =   0
         varVarType_3    =   5
         Array[1]_2      =   13
         ClassName_13    =   "CCWEnumElt"
         opts_13         =   1
         Name_13         =   "On"
         frame_13        =   286
         DrawList_13     =   14
         ClassName_14    =   "CDrawList"
         count_14        =   4
         list[4]_14      =   15
         ClassName_15    =   "CCWDrawObj"
         opts_15         =   62
         C[0]_15         =   -2147483641
         C[1]_15         =   -2147483641
         Image_15        =   6
         Animator_15     =   0
         Blinker_15      =   0
         list[3]_14      =   16
         ClassName_16    =   "CCWDrawObj"
         opts_16         =   62
         C[0]_16         =   0
         C[1]_16         =   0
         Image_16        =   8
         Animator_16     =   0
         Blinker_16      =   0
         list[2]_14      =   17
         ClassName_17    =   "CCWDrawObj"
         opts_17         =   52
         SclRef_17.l     =   2
         SclRef_17.t     =   17
         SclRef_17.r     =   72
         SclRef_17.b     =   44
         Scl_17.l        =   10
         Scl_17.t        =   26
         Scl_17.r        =   63
         Scl_17.b        =   40
         C[0]_17         =   16777215
         C[1]_17         =   16777215
         Image_17        =   10
         Animator_17     =   0
         Blinker_17      =   0
         list[1]_14      =   18
         ClassName_18    =   "CCWDrawObj"
         opts_18         =   54
         SclRef_18.l     =   2
         SclRef_18.t     =   17
         SclRef_18.r     =   63
         SclRef_18.b     =   44
         Scl_18.l        =   2
         Scl_18.t        =   15
         Scl_18.r        =   63
         Scl_18.b        =   42
         C[0]_18         =   65280
         C[1]_18         =   65280
         Image_18        =   19
         ClassName_19    =   "CCWPictImage"
         opts_19         =   1280
         Rows_19         =   1
         Cols_19         =   1
         Pict_19         =   42
         F_19            =   65280
         B_19            =   65280
         ColorReplaceWith_19=   8421504
         ColorReplace_19 =   8421504
         Tolerance_19    =   2
         Animator_18     =   0
         Blinker_18      =   0
         varVarType_13   =   5
         Bindings_1      =   20
         ClassName_20    =   "CCWBindingHolderArray"
         Editor_20       =   21
         ClassName_21    =   "CCWBindingHolderArrayEditor"
         Owner_21        =   1
         Style_1         =   7
         mechAction_1    =   1
         BGImg_1         =   22
         ClassName_22    =   "CCWDrawObj"
         opts_22         =   62
         Image_22        =   23
         ClassName_23    =   "CCWPictImage"
         opts_23         =   1280
         Rows_23         =   1
         Cols_23         =   1
         Pict_23         =   286
         F_23            =   -2147483633
         B_23            =   -2147483633
         ColorReplaceWith_23=   8421504
         ColorReplace_23 =   8421504
         Tolerance_23    =   2
         Animator_22     =   0
         Blinker_22      =   0
         Array_1         =   6
         Editor_1        =   0
         Array[0]_1      =   11
         Array[1]_1      =   18
         Array[2]_1      =   0
         Array[3]_1      =   0
         Array[4]_1      =   24
         ClassName_24    =   "CCWDrawObj"
         opts_24         =   62
         Image_24        =   8
         Animator_24     =   0
         Blinker_24      =   0
         Array[5]_1      =   25
         ClassName_25    =   "CCWDrawObj"
         opts_25         =   62
         Image_25        =   10
         Animator_25     =   0
         Blinker_25      =   0
         Label_1         =   26
         ClassName_26    =   "CCWDrawObj"
         opts_26         =   62
         C[0]_26         =   -2147483640
         Image_26        =   6
         Animator_26     =   0
         Blinker_26      =   0
      End
      Begin VB.Frame frameDebug
         BackColor       =   &H00C0E0FF&
         Caption         =   "Communication History"
         BeginProperty Font
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5580
         Left            =   1710
         TabIndex        =   20
         Top             =   225
         Width           =   9885
         Begin VB.CommandButton CmdCopyClipboard
            BackColor       =   &H00C0FFC0&
            Caption         =   "Copy Ctrl-C"
            BeginProperty Font
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   1530
            Style           =   1  'Graphical
            TabIndex        =   190
            ToolTipText     =   "Copies into Clipboard text shown on screen"
            Top             =   315
            Width           =   1110
         End
         Begin VB.CommandButton CmdSaveResponsesToPC
            Caption         =   "Save to PC"
            BeginProperty Font
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   180
            TabIndex        =   189
            ToolTipText     =   "Saves into file text shown on screen"
            Top             =   315
            Width           =   1065
         End
         Begin VB.CommandButton cmdClearTerminal
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
            Height          =   420
            Left            =   2925
            Style           =   1  'Graphical
            TabIndex        =   23
            ToolTipText     =   "clears history list"
            Top             =   315
            Width           =   1200
         End
         Begin VB.CheckBox Chk_AutoClear
            BackColor       =   &H0080FFFF&
            Caption         =   "Auto clear each 2000 lines"
            Height          =   285
            Left            =   4410
            TabIndex        =   21
            ToolTipText     =   "If checked, list will be Auto cleared when 1000 lines are recorded and started over to prevent slow-down"
            Top             =   405
            Value           =   1  'Checked
            Width           =   2265
         End
         Begin VB.ListBox lstResponse
            BeginProperty Font
               Name            =   "Courier New"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   4560
            ItemData        =   "Analyzer.frx":18B484
            Left            =   100
            List            =   "Analyzer.frx":18B486
            TabIndex        =   22
            Top             =   855
            Width           =   9540
         End
      End
      Begin VB.Image ImageConnection
         Height          =   600
         Left            =   135
         Picture         =   "Analyzer.frx":18B488
         Stretch         =   -1  'True
         Top             =   315
         Width           =   705
      End
   End
   Begin VB.Frame FrameAlarms
      BackColor       =   &H00C0C0FF&
      Caption         =   "Alarm Log"
      BeginProperty Font
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5970
      Left            =   2727
      TabIndex        =   10
      Top             =   1661
      Visible         =   0   'False
      Width           =   11756
      Begin VB.CommandButton cmdClearErrorLog
         BackColor       =   &H00C0E0FF&
         Caption         =   "Clear Alarm List (in this session)"
         BeginProperty Font
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   990
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Clears window below. DOES NOT clear Alarm History file. If file needs to be cleared, use any editor like ""Notepad"""
         Top             =   810
         Width           =   3300
      End
      Begin VB.TextBox txtAlarmListPath
         BackColor       =   &H00FFFFC0&
         BeginProperty Font
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   990
         Locked          =   -1  'True
         TabIndex        =   13
         Text            =   "Alarm Events are saved in: .\Alarm.log"
         ToolTipText     =   "You can copy file path from here to find it in Explorer"
         Top             =   315
         Width           =   10455
      End
      Begin VB.CommandButton cmdCloseAlarmLog
         BackColor       =   &H00C0FFFF&
         Caption         =   "Close Alarm Log File"
         BeginProperty Font
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   8145
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Closes the history file ""Alarm.log"""
         Top             =   810
         Width           =   3300
      End
      Begin VB.CommandButton cmdViewAlarmLog
         BackColor       =   &H00C0FFC0&
         Caption         =   "View Alarm Log File"
         BeginProperty Font
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   4815
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Shows Alarm History saved in the ""Alarm.log"" file. If file needs to be cleared, use any editor like ""Notepad"""
         Top             =   810
         Width           =   2805
      End
      Begin VB.ListBox lstAlarms
         BeginProperty Font
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2370
         ItemData        =   "Analyzer.frx":19F50A
         Left            =   225
         List            =   "Analyzer.frx":19F511
         TabIndex        =   16
         ToolTipText     =   "Alarms discovered during run of this session"
         Top             =   1395
         Width           =   11235
      End
      Begin VB.ListBox lstAlarmFile
         BackColor       =   &H00C0FFFF&
         BeginProperty Font
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1740
         ItemData        =   "Analyzer.frx":19F520
         Left            =   225
         List            =   "Analyzer.frx":19F527
         TabIndex        =   15
         ToolTipText     =   "Alarm Log File, all alarms are recorded here"
         Top             =   3915
         Visible         =   0   'False
         Width           =   11235
      End
      Begin VB.Image ImageAlarm
         Height          =   600
         Left            =   180
         Picture         =   "Analyzer.frx":19F539
         Stretch         =   -1  'True
         Top             =   315
         Width           =   600
      End
   End
   Begin VB.Frame frmFWupgrade
      BackColor       =   &H0080E0FF&
      Caption         =   "Firmware Parameters"
      Enabled         =   0   'False
      BeginProperty Font
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6022
      Left            =   2727
      TabIndex        =   6
      Top             =   1627
      Visible         =   0   'False
      Width           =   11756
      Begin VB.CommandButton cmdRestoreParameters
         BackColor       =   &H00C0FFC0&
         Caption         =   "Read && Restore Parameters"
         BeginProperty Font
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   180
         Style           =   1  'Graphical
         TabIndex        =   117
         ToolTipText     =   "Read file with saved parameters and send parameters to controller "
         Top             =   3780
         Width           =   1665
      End
      Begin VB.CommandButton CmdSaveParamList
         Caption         =   "Save Shown Params to PC"
         BeginProperty Font
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   180
         TabIndex        =   35
         ToolTipText     =   "Saves into file parameters shown on screen"
         Top             =   5310
         Width           =   1650
      End
      Begin VB.Frame FrameSaveParamsIntoFile
         BackColor       =   &H0080E0FF&
         Height          =   1680
         Left            =   90
         TabIndex        =   28
         Top             =   900
         Width           =   1860
         Begin VB.CommandButton cmdSaveParameters
            BackColor       =   &H00FFFFC0&
            Caption         =   "Retrieve && Save Parameters"
            BeginProperty Font
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Left            =   90
            Style           =   1  'Graphical
            TabIndex        =   36
            ToolTipText     =   "Read parameters from controller and save them into file"
            Top             =   180
            Width           =   1665
         End
         Begin VB.CheckBox ChkVerboseParams
            BackColor       =   &H00FFFFC0&
            Caption         =   "Include Comments"
            Height          =   375
            Left            =   90
            TabIndex        =   31
            ToolTipText     =   "Param file includes comments when checked"
            Top             =   1170
            Value           =   1  'Checked
            Width           =   1680
         End
         Begin VB.OptionButton Opt_ALL_params_or_Essential
            BackColor       =   &H00C0C0FF&
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
            Left            =   1140
            TabIndex        =   30
            ToolTipText     =   "Output ALL parameters (can take about 20 sec!)"
            Top             =   810
            Width           =   620
         End
         Begin VB.OptionButton Opt_ALL_params_or_Essential
            BackColor       =   &H00C0FFC0&
            Caption         =   "Essential"
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
            Left            =   90
            TabIndex        =   29
            ToolTipText     =   "Output only Essential parameters like gains  (quick)"
            Top             =   810
            Value           =   -1  'True
            Width           =   980
         End
      End
      Begin VB.ListBox ListParamsErrors
         BeginProperty Font
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1740
         ItemData        =   "Analyzer.frx":1A75BB
         Left            =   1980
         List            =   "Analyzer.frx":1A75C2
         TabIndex        =   27
         Top             =   4095
         Width           =   9630
      End
      Begin VB.CommandButton cmdUpdateFirmware
         BackColor       =   &H00C0C0FF&
         Caption         =   "Update Firmware"
         BeginProperty Font
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   180
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Instructions how to update firmware"
         Top             =   4545
         Width           =   1650
      End
      Begin VB.CommandButton CmdClearParamList
         Caption         =   "Clear Screen"
         BeginProperty Font
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   900
         TabIndex        =   7
         ToolTipText     =   "Clears this window, parameters in controller unaffected"
         Top             =   315
         Width           =   930
      End
      Begin CWUIControlsLib.CWButton cwLabelAction
         Height          =   405
         Left            =   225
         TabIndex        =   37
         Top             =   2655
         Visible         =   0   'False
         Width           =   1590
         _Version        =   524288
         _ExtentX        =   2796
         _ExtentY        =   714
         _StockProps     =   68
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}
            Name            =   "Arial"
            Size            =   9.74
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Reset_0         =   0   'False
         CompatibleVers_0=   524288
         Boolean_0       =   1
         ClassName_1     =   "CCWBoolean"
         opts_1          =   2614
         SclRef_1.l      =   2
         SclRef_1.t      =   2
         SclRef_1.r      =   131
         SclRef_1.b      =   35
         Scl_1.r         =   133
         Scl_1.b         =   42
         C[0]_1          =   12648447
         Enum_1          =   2
         ClassName_2     =   "CCWEnum"
         Array_2         =   2
         Editor_2        =   0
         Array[0]_2      =   3
         ClassName_3     =   "CCWEnumElt"
         opts_3          =   1
         Name_3          =   "Off"
         frame_3         =   288
         DrawList_3      =   4
         ClassName_4     =   "CDrawList"
         count_4         =   4
         list[4]_4       =   5
         ClassName_5     =   "CCWDrawObj"
         opts_5          =   62
         C[0]_5          =   0
         C[1]_5          =   0
         Image_5         =   6
         ClassName_6     =   "CCWTextImage"
         font_6          =   0
         Animator_5      =   0
         Blinker_5       =   0
         list[3]_4       =   7
         ClassName_7     =   "CCWDrawObj"
         opts_7          =   60
         C[0]_7          =   255
         C[1]_7          =   255
         Image_7         =   8
         ClassName_8     =   "CCWTextImage"
         szText_8        =   "READING FILE"
         style_8         =   48598384
         font_8          =   0
         Animator_7      =   0
         Blinker_7       =   9
         ClassName_9     =   "CCWBlinker"
         Interval_9      =   900
         EventID_9       =   7
         Obj_9           =   7
         list[2]_4       =   10
         ClassName_10    =   "CCWDrawObj"
         opts_10         =   62
         C[0]_10         =   255
         C[1]_10         =   255
         Image_10        =   11
         ClassName_11    =   "CCWTextImage"
         szText_11       =   "RETRIEVING"
         font_11         =   0
         Animator_10     =   0
         Blinker_10      =   0
         list[1]_4       =   12
         ClassName_12    =   "CCWDrawObj"
         opts_12         =   62
         C[0]_12         =   16777152
         C[1]_12         =   16777152
         Image_12        =   13
         ClassName_13    =   "CCWPictImage"
         opts_13         =   1280
         Rows_13         =   1
         Cols_13         =   1
         Pict_13         =   2
         F_13            =   16777152
         B_13            =   16777152
         ColorReplaceWith_13=   8421504
         ColorReplace_13 =   8421504
         Tolerance_13    =   2
         Animator_12     =   0
         Blinker_12      =   14
         ClassName_14    =   "CCWBlinker"
         Interval_14     =   900
         EventID_14      =   7
         Obj_14          =   12
         varVarType_3    =   5
         Array[1]_2      =   15
         ClassName_15    =   "CCWEnumElt"
         opts_15         =   1
         Name_15         =   "On"
         frame_15        =   287
         DrawList_15     =   16
         ClassName_16    =   "CDrawList"
         count_16        =   4
         list[4]_16      =   17
         ClassName_17    =   "CCWDrawObj"
         opts_17         =   62
         C[0]_17         =   0
         C[1]_17         =   0
         Image_17        =   6
         Animator_17     =   0
         Blinker_17      =   0
         list[3]_16      =   18
         ClassName_18    =   "CCWDrawObj"
         opts_18         =   62
         C[0]_18         =   255
         C[1]_18         =   255
         Image_18        =   8
         Animator_18     =   0
         Blinker_18      =   0
         list[2]_16      =   19
         ClassName_19    =   "CCWDrawObj"
         opts_19         =   60
         C[0]_19         =   255
         C[1]_19         =   255
         Image_19        =   11
         Animator_19     =   0
         Blinker_19      =   20
         ClassName_20    =   "CCWBlinker"
         Interval_20     =   600
         Enable_20       =   -1  'True
         EventID_20      =   7
         Obj_20          =   19
         list[1]_16      =   21
         ClassName_21    =   "CCWDrawObj"
         opts_21         =   62
         C[0]_21         =   12648384
         C[1]_21         =   12648384
         Image_21        =   22
         ClassName_22    =   "CCWPictImage"
         opts_22         =   1280
         Rows_22         =   1
         Cols_22         =   1
         Pict_22         =   2
         F_22            =   12648384
         B_22            =   12648384
         ColorReplaceWith_22=   8421504
         ColorReplace_22 =   8421504
         Tolerance_22    =   2
         Animator_21     =   0
         Blinker_21      =   23
         ClassName_23    =   "CCWBlinker"
         Interval_23     =   900
         Enable_23       =   -1  'True
         EventID_23      =   7
         Obj_23          =   21
         varVarType_15   =   5
         Bindings_1      =   24
         ClassName_24    =   "CCWBindingHolderArray"
         Editor_24       =   25
         ClassName_25    =   "CCWBindingHolderArrayEditor"
         Owner_25        =   1
         Style_1         =   12
         drawState_1     =   1
         frameStyle_1    =   1
         mechAction_1    =   3
         BGImg_1         =   26
         ClassName_26    =   "CCWDrawObj"
         opts_26         =   62
         Image_26        =   27
         ClassName_27    =   "CCWPictImage"
         opts_27         =   1280
         Rows_27         =   1
         Cols_27         =   1
         Pict_27         =   286
         F_27            =   -2147483633
         B_27            =   -2147483633
         ColorReplaceWith_27=   8421504
         ColorReplace_27 =   8421504
         Tolerance_27    =   2
         Animator_26     =   0
         Blinker_26      =   0
         Array_1         =   6
         Editor_1        =   0
         Array[0]_1      =   12
         Array[1]_1      =   21
         Array[2]_1      =   0
         Array[3]_1      =   0
         Array[4]_1      =   28
         ClassName_28    =   "CCWDrawObj"
         opts_28         =   62
         Image_28        =   8
         Animator_28     =   0
         Blinker_28      =   0
         Array[5]_1      =   29
         ClassName_29    =   "CCWDrawObj"
         opts_29         =   62
         Image_29        =   11
         Animator_29     =   0
         Blinker_29      =   0
         Label_1         =   30
         ClassName_30    =   "CCWDrawObj"
         opts_30         =   62
         C[0]_30         =   -2147483640
         Image_30        =   6
         Animator_30     =   0
         Blinker_30      =   0
      End
      Begin VB.ListBox ListParams
         BeginProperty Font
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3630
         ItemData        =   "Analyzer.frx":1A75DA
         Left            =   1980
         List            =   "Analyzer.frx":1A75DC
         TabIndex        =   8
         Top             =   270
         Width           =   9630
      End
      Begin VB.ListBox ListShadow
         Enabled         =   0   'False
         BeginProperty Font
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3630
         ItemData        =   "Analyzer.frx":1A75DE
         Left            =   1980
         List            =   "Analyzer.frx":1A75E5
         TabIndex        =   32
         Top             =   270
         Width           =   9630
      End
      Begin VB.Label LabelParamLines
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "1024"
         BeginProperty Font
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   945
         TabIndex        =   33
         Top             =   3195
         Width           =   870
      End
      Begin VB.Label Label1
         Alignment       =   2  'Center
         BackColor       =   &H0080E0FF&
         Caption         =   "Param count"
         BeginProperty Font
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Index           =   4
         Left            =   180
         TabIndex        =   34
         Top             =   3105
         Width           =   825
      End
      Begin VB.Image ImageParam
         Height          =   600
         Left            =   180
         Picture         =   "Analyzer.frx":1A75F1
         Stretch         =   -1  'True
         Top             =   315
         Width           =   600
      End
   End
   Begin VB.Frame frmDiagnostic
      BackColor       =   &H00FFE0E0&
      Caption         =   "Diagnostic && tuning"
      Enabled         =   0   'False
      BeginProperty Font
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5970
      Left            =   2745
      TabIndex        =   26
      Top             =   1661
      Visible         =   0   'False
      Width           =   11756
      Begin VB.CommandButton CommandDiagRefresh
         BackColor       =   &H00FF80FF&
         Caption         =   "REFRESH"
         BeginProperty Font
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   1395
         Style           =   1  'Graphical
         TabIndex        =   69
         Top             =   360
         Width           =   1050
      End
      Begin VB.CommandButton cmdPulse
         BackColor       =   &H0080FFFF&
         Caption         =   "Apply Pulse"
         BeginProperty Font
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   780
         Left            =   4905
         Picture         =   "Analyzer.frx":1C7673
         Style           =   1  'Graphical
         TabIndex        =   58
         ToolTipText     =   "Applies pulse to chosen axis (If excitation was ON, it is disabled after pulse)"
         Top             =   270
         Width           =   1095
      End
      Begin VB.OptionButton optExitation
         BackColor       =   &H00C0FFC0&
         Caption         =   "Sine"
         BeginProperty Font
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   0
         Left            =   7110
         TabIndex        =   57
         ToolTipText     =   "Enables Sine exitation (Frequency MUST BE NOT ZERO)"
         Top             =   735
         Value           =   -1  'True
         Visible         =   0   'False
         Width           =   870
      End
      Begin VB.OptionButton optExitation
         BackColor       =   &H00FFC0FF&
         Caption         =   "Noise"
         BeginProperty Font
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   1
         Left            =   7110
         TabIndex        =   56
         ToolTipText     =   "Enables noise exitation (Frequency is set to 0 Hz means niose)"
         Top             =   285
         Visible         =   0   'False
         Width           =   870
      End
      Begin VB.ComboBox cmb_ExcitAxis
         BeginProperty Font
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         ItemData        =   "Analyzer.frx":1C90B9
         Left            =   9540
         List            =   "Analyzer.frx":1C910A
         Style           =   2  'Dropdown List
         TabIndex        =   55
         ToolTipText     =   "Choose axis OUTPUT for excitation"
         Top             =   270
         Width           =   2055
      End
      Begin CWUIControlsLib.CWButton cwBut_Excitation
         Height          =   930
         Left            =   6120
         TabIndex        =   59
         ToolTipText     =   "Enable / Disable sine excitation on chosen excitation output"
         Top             =   160
         Width           =   915
         _Version        =   524288
         _ExtentX        =   1614
         _ExtentY        =   1640
         _StockProps     =   68
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Reset_0         =   0   'False
         CompatibleVers_0=   524288
         Boolean_0       =   1
         ClassName_1     =   "CCWBoolean"
         opts_1          =   2614
         SclRef_1.l      =   2
         SclRef_1.t      =   2
         SclRef_1.r      =   63
         SclRef_1.b      =   61
         Scl_1.l         =   6
         Scl_1.t         =   14
         Scl_1.r         =   59
         Scl_1.b         =   61
         C[0]_1          =   16769248
         Enum_1          =   2
         ClassName_2     =   "CCWEnum"
         Array_2         =   2
         Editor_2        =   0
         Array[0]_2      =   3
         ClassName_3     =   "CCWEnumElt"
         opts_3          =   1
         Name_3          =   "Off"
         frame_3         =   286
         DrawList_3      =   4
         ClassName_4     =   "CDrawList"
         count_4         =   4
         list[4]_4       =   5
         ClassName_5     =   "CCWDrawObj"
         opts_5          =   54
         SclRef_5.l      =   6
         SclRef_5.t      =   14
         SclRef_5.r      =   59
         SclRef_5.b      =   61
         Scl_5.l         =   3
         Scl_5.r         =   64
         Scl_5.b         =   14
         C[0]_5          =   -2147483640
         C[1]_5          =   -2147483640
         Image_5         =   6
         ClassName_6     =   "CCWTextImage"
         szText_6        =   "excitation"
         font_6          =   7
         ClassName_7     =   "CCWFont"
         bFont_7         =   -1  'True
         BeginProperty Font_7 {0BE35203-8F91-11CE-9DE3-00AA004BB851}
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Animator_5      =   0
         Blinker_5       =   0
         list[3]_4       =   8
         ClassName_8     =   "CCWDrawObj"
         opts_8          =   60
         C[0]_8          =   -2147483640
         C[1]_8          =   -2147483640
         Image_8         =   9
         ClassName_9     =   "CCWTextImage"
         font_9          =   10
         ClassName_10    =   "CCWFont"
         bFont_10        =   -1  'True
         BeginProperty Font_10 {0BE35203-8F91-11CE-9DE3-00AA004BB851}
            Name            =   "Arial"
            Size            =   12
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Animator_8      =   0
         Blinker_8       =   0
         list[2]_4       =   11
         ClassName_11    =   "CCWDrawObj"
         opts_11         =   62
         C[0]_11         =   -2147483640
         C[1]_11         =   -2147483640
         Image_11        =   12
         ClassName_12    =   "CCWTextImage"
         style_12        =   1076101120
         font_12         =   13
         ClassName_13    =   "CCWFont"
         bFont_13        =   -1  'True
         BeginProperty Font_13 {0BE35203-8F91-11CE-9DE3-00AA004BB851}
            Name            =   "Arial"
            Size            =   12
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Animator_11     =   0
         Blinker_11      =   0
         list[1]_4       =   14
         ClassName_14    =   "CCWDrawObj"
         opts_14         =   62
         C[0]_14         =   8454016
         C[1]_14         =   8454016
         Image_14        =   15
         ClassName_15    =   "CCWPictImage"
         opts_15         =   256
         Rows_15         =   1
         Cols_15         =   1
         URL_15          =   "C:\in_out\Software\DC-2020_ED_Analyzer_VB6\Resources\SwitchOffGreen.bmp"
         Pict_15         =   "Analyzer.frx":1C9233
         ColorReplaceWith_15=   8421504
         ColorReplace_15 =   8421504
         Tolerance_15    =   2
         Animator_14     =   0
         Blinker_14      =   0
         varVarType_3    =   5
         Array[1]_2      =   16
         ClassName_16    =   "CCWEnumElt"
         opts_16         =   1
         Name_16         =   "On"
         frame_16        =   286
         DrawList_16     =   17
         ClassName_17    =   "CDrawList"
         count_17        =   4
         list[4]_17      =   18
         ClassName_18    =   "CCWDrawObj"
         opts_18         =   54
         SclRef_18.l     =   6
         SclRef_18.t     =   14
         SclRef_18.r     =   59
         SclRef_18.b     =   61
         Scl_18.l        =   3
         Scl_18.r        =   64
         Scl_18.b        =   14
         C[0]_18         =   -2147483640
         C[1]_18         =   -2147483640
         Image_18        =   6
         Animator_18     =   0
         Blinker_18      =   0
         list[3]_17      =   19
         ClassName_19    =   "CCWDrawObj"
         opts_19         =   62
         C[0]_19         =   -2147483640
         C[1]_19         =   -2147483640
         Image_19        =   9
         Animator_19     =   0
         Blinker_19      =   0
         list[2]_17      =   20
         ClassName_20    =   "CCWDrawObj"
         opts_20         =   60
         C[0]_20         =   -2147483640
         C[1]_20         =   -2147483640
         Image_20        =   12
         Animator_20     =   0
         Blinker_20      =   0
         list[1]_17      =   21
         ClassName_21    =   "CCWDrawObj"
         opts_21         =   62
         C[0]_21         =   8421631
         C[1]_21         =   8421631
         Image_21        =   22
         ClassName_22    =   "CCWPictImage"
         opts_22         =   256
         Rows_22         =   1
         Cols_22         =   1
         URL_22          =   "C:\in_out\Software\DC-2020_ED_Analyzer_VB6\Resources\SwitchOnYellow.bmp"
         Pict_22         =   "Analyzer.frx":1E2965
         ColorReplaceWith_22=   8421504
         ColorReplace_22 =   8421504
         Tolerance_22    =   2
         Animator_21     =   0
         Blinker_21      =   23
         ClassName_23    =   "CCWBlinker"
         Interval_23     =   600
         EventID_23      =   7
         Obj_23          =   21
         varVarType_16   =   5
         Bindings_1      =   24
         ClassName_24    =   "CCWBindingHolderArray"
         Editor_24       =   25
         ClassName_25    =   "CCWBindingHolderArrayEditor"
         Owner_25        =   1
         mechAction_1    =   1
         BGImg_1         =   26
         ClassName_26    =   "CCWDrawObj"
         opts_26         =   62
         Image_26        =   27
         ClassName_27    =   "CCWPictImage"
         opts_27         =   1280
         Rows_27         =   1
         Cols_27         =   1
         Pict_27         =   286
         F_27            =   -2147483633
         B_27            =   -2147483633
         ColorReplaceWith_27=   8421504
         ColorReplace_27 =   8421504
         Tolerance_27    =   2
         Animator_26     =   0
         Blinker_26      =   0
         Array_1         =   6
         Editor_1        =   0
         Array[0]_1      =   14
         Array[1]_1      =   21
         Array[2]_1      =   0
         Array[3]_1      =   0
         Array[4]_1      =   28
         ClassName_28    =   "CCWDrawObj"
         opts_28         =   62
         Image_28        =   9
         Animator_28     =   0
         Blinker_28      =   0
         Array[5]_1      =   29
         ClassName_29    =   "CCWDrawObj"
         opts_29         =   62
         Image_29        =   12
         Animator_29     =   0
         Blinker_29      =   0
         Label_1         =   30
         ClassName_30    =   "CCWDrawObj"
         opts_30         =   62
         C[0]_30         =   -2147483640
         Image_30        =   6
         Animator_30     =   0
         Blinker_30      =   0
      End
      Begin CWUIControlsLib.CWNumEdit cwNumExcitFreq
         Height          =   360
         Index           =   1
         Left            =   8595
         TabIndex        =   60
         Tag             =   "freq"
         ToolTipText     =   "Driving frequency, Hz, 0.01...1000 Hz (default = 5 Hz)"
         Top             =   720
         Width           =   1110
         _Version        =   524288
         _ExtentX        =   1958
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
         C[0]_1          =   12648384
         C[1]_1          =   0
         BorderStyle_1   =   1
         TextAlignment_1 =   2
         format_1        =   2
         ClassName_2     =   "CCWFormat"
         Format_2        =   ".0#"
         scale_1         =   3
         ClassName_3     =   "CCWScale"
         opts_3          =   65536
         dMax_3          =   10
         discInterval_3  =   1
         ValueVarType_1  =   5
         Value_Val_1     =   1950
         IncValueVarType_1=   5
         IncValue_Val_1  =   1
         AccelIncVarType_1=   5
         AccelInc_Val_1  =   10
         RangeMinVarType_1=   5
         RangeMaxVarType_1=   5
         RangeMax_Val_1  =   2000
         ButtonStyle_1   =   0
         Bindings_1      =   4
         ClassName_4     =   "CCWBindingHolderArray"
         Editor_4        =   5
         ClassName_5     =   "CCWBindingHolderArrayEditor"
         Owner_5         =   1
      End
      Begin CWUIControlsLib.CWNumEdit cwNumExcitAmpl
         Height          =   360
         Index           =   1
         Left            =   10395
         TabIndex        =   61
         Tag             =   "ampl"
         ToolTipText     =   "Driving Amplitude, -30000...30000 counts (default 3000)"
         Top             =   720
         Width           =   1200
         _Version        =   524288
         _ExtentX        =   2117
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
         C[0]_1          =   12648447
         C[1]_1          =   0
         BorderStyle_1   =   1
         TextAlignment_1 =   2
         format_1        =   2
         ClassName_2     =   "CCWFormat"
         Format_2        =   ".###"
         scale_1         =   3
         ClassName_3     =   "CCWScale"
         opts_3          =   65536
         dMax_3          =   10
         discInterval_3  =   1
         ValueVarType_1  =   5
         Value_Val_1     =   -30000
         IncValueVarType_1=   5
         IncValue_Val_1  =   10
         AccelIncVarType_1=   5
         AccelInc_Val_1  =   10
         RangeMinVarType_1=   5
         RangeMin_Val_1  =   -30000
         RangeMaxVarType_1=   5
         RangeMax_Val_1  =   30000
         ButtonStyle_1   =   0
         Bindings_1      =   4
         ClassName_4     =   "CCWBindingHolderArray"
         Editor_4        =   5
         ClassName_5     =   "CCWBindingHolderArrayEditor"
         Owner_5         =   1
      End
      Begin CWUIControlsLib.CWNumEdit cwNumPulsePosSlope
         Height          =   360
         Left            =   3780
         TabIndex        =   62
         Tag             =   "prup"
         ToolTipText     =   "Pulse positive front slope, counts per loop. Larger number=sharper front. Range 0.01 to 100, default 3"
         Top             =   270
         Width           =   1020
         _Version        =   524288
         _ExtentX        =   1799
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
         C[0]_1          =   12632319
         C[1]_1          =   0
         BorderStyle_1   =   1
         TextAlignment_1 =   2
         format_1        =   2
         ClassName_2     =   "CCWFormat"
         Format_2        =   ".0#"
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
         AccelInc_Val_1  =   1
         RangeMinVarType_1=   5
         RangeMin_Val_1  =   0.01
         RangeMaxVarType_1=   5
         RangeMax_Val_1  =   100
         ButtonStyle_1   =   0
         Bindings_1      =   4
         ClassName_4     =   "CCWBindingHolderArray"
         Editor_4        =   5
         ClassName_5     =   "CCWBindingHolderArrayEditor"
         Owner_5         =   1
      End
      Begin CWUIControlsLib.CWNumEdit cwNumPulseNegSlope
         Height          =   360
         Left            =   3780
         TabIndex        =   63
         Tag             =   "prdn"
         ToolTipText     =   "Pulse negative front slope, counts per loop. Larger number=sharper front. Range 0.01 to 100, default 3"
         Top             =   675
         Width           =   1020
         _Version        =   524288
         _ExtentX        =   1799
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
         Format_2        =   ".0#"
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
         AccelInc_Val_1  =   1
         RangeMinVarType_1=   5
         RangeMin_Val_1  =   0.01
         RangeMaxVarType_1=   5
         RangeMax_Val_1  =   100
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
         BackColor       =   &H00FFFFFF&
         Caption         =   "Excitation output"
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
         Height          =   330
         Index           =   18
         Left            =   8055
         TabIndex        =   68
         ToolTipText     =   "Choose axis OUTPUT for excitation"
         Top             =   285
         Width           =   1545
      End
      Begin VB.Label Label1
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Caption         =   "Ampl"
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
         Height          =   330
         Index           =   21
         Left            =   9810
         TabIndex        =   67
         ToolTipText     =   "Driving Amplitude, -30000...+-30000 (default = 3000)"
         Top             =   720
         Width           =   600
      End
      Begin VB.Label Label1
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         Caption         =   "Freq"
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
         Height          =   330
         Index           =   23
         Left            =   8055
         TabIndex        =   66
         ToolTipText     =   "Driving frequency, Hz, 0.01...1000 Hz (default = 5 Hz)"
         Top             =   735
         Width           =   555
      End
      Begin VB.Label Label1
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
         Caption         =   "pulse '/' slope"
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
         Height          =   300
         Index           =   22
         Left            =   2565
         TabIndex        =   65
         ToolTipText     =   "Pulse positive front slope, counts per loop. Larger number=sharper front. Range 0.01 to 100, default 3"
         Top             =   285
         Width           =   1215
      End
      Begin VB.Label Label1
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Caption         =   "pulse '\' slope"
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
         Height          =   300
         Index           =   24
         Left            =   2565
         TabIndex        =   64
         ToolTipText     =   "Pulse negative front slope, counts per loop. Larger number=sharper front. Range 0.01 to 100, default 3"
         Top             =   720
         Width           =   1215
      End
      Begin VB.Image ImageDiagTuning
         Height          =   600
         Index           =   1
         Left            =   720
         Picture         =   "Analyzer.frx":1FC5B3
         Stretch         =   -1  'True
         Top             =   360
         Width           =   600
      End
      Begin VB.Image ImageDiagTuning
         Height          =   600
         Index           =   0
         Left            =   135
         Picture         =   "Analyzer.frx":210DF5
         Stretch         =   -1  'True
         Top             =   360
         Width           =   600
      End
   End
   Begin VB.Label LblFB_FF_Gains
      BackColor       =   &H00C0E0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "FB && FF Gains"
      BeginProperty Font
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   285
      Left            =   720
      TabIndex        =   150
      Top             =   7965
      Width           =   1600
   End
   Begin VB.Label LblFilters
      BackColor       =   &H00C0E0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Filters"
      BeginProperty Font
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   285
      Left            =   720
      TabIndex        =   149
      Top             =   7380
      Width           =   1605
   End
   Begin VB.Image ImageFilters
      Height          =   485
      Left            =   45
      Picture         =   "Analyzer.frx":230E77
      Stretch         =   -1  'True
      Top             =   7268
      Width           =   540
   End
   Begin VB.Label lbl_Version_Date
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
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
      Left            =   0
      TabIndex        =   53
      ToolTipText     =   "TMC part number of this software"
      Top             =   2229
      Width           =   2715
   End
   Begin VB.Image Image1
      Height          =   2025
      Left            =   0
      Picture         =   "Analyzer.frx":2326F9
      Top             =   0
      Width           =   2715
   End
   Begin VB.Image ImageScope
      Height          =   480
      Left            =   45
      Picture         =   "Analyzer.frx":24461B
      Stretch         =   -1  'True
      Top             =   5610
      Width           =   540
   End
   Begin VB.Label lbl_Part_Number
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "PN 95-38961-02 rev.G"
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
      Left            =   0
      TabIndex        =   54
      ToolTipText     =   "TMC part number of this software"
      Top             =   1993
      Width           =   2715
   End
   Begin VB.Image ImageMonitoring
      Height          =   480
      Left            =   45
      Picture         =   "Analyzer.frx":254A5D
      Stretch         =   -1  'True
      Top             =   6165
      Width           =   540
   End
   Begin VB.Image ImageTerminal
      Height          =   485
      Left            =   45
      Picture         =   "Analyzer.frx":274ADF
      Stretch         =   -1  'True
      Top             =   6705
      Width           =   540
   End
   Begin VB.Image ImageGains
      Height          =   485
      Left            =   45
      Picture         =   "Analyzer.frx":294B61
      Stretch         =   -1  'True
      Top             =   7830
      Width           =   540
   End
   Begin VB.Label LblCONNparams
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " COM15 closed"
      BeginProperty Font
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   284
      Left            =   810
      TabIndex        =   19
      Top             =   8415
      Width           =   1860
   End
   Begin VB.Image ImageDisconnected
      BorderStyle     =   1  'Fixed Single
      Height          =   360
      Left            =   90
      Picture         =   "Analyzer.frx":2A93A3
      ToolTipText     =   "Click to connect"
      Top             =   8370
      Width           =   660
   End
   Begin VB.Image ImageConnected
      BorderStyle     =   1  'Fixed Single
      Height          =   360
      Left            =   90
      Picture         =   "Analyzer.frx":2A9D45
      ToolTipText     =   "Click to DISconnect"
      Top             =   8370
      Width           =   660
   End
   Begin VB.Label lblSystemDate
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "MM/DD/YYYY"
      BeginProperty Font
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   284
      Left            =   12195
      TabIndex        =   1
      ToolTipText     =   "Do not check time - WORK !"
      Top             =   8415
      Width           =   1350
   End
   Begin VB.Label LblTerminalWindow
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Terminal"
      BeginProperty Font
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   285
      Left            =   720
      TabIndex        =   51
      Top             =   6795
      Width           =   1600
   End
   Begin VB.Label LblPOSmonitoring
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "Test && Recording"
      BeginProperty Font
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   285
      Left            =   720
      TabIndex        =   52
      Top             =   6255
      Width           =   2040
   End
   Begin VB.Label LblScope
      BackColor       =   &H00FFC0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Scope"
      BeginProperty Font
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   285
      Left            =   720
      TabIndex        =   148
      Top             =   5715
      Visible         =   0   'False
      Width           =   1605
   End
   Begin VB.Label lblSystemTime
      Alignment       =   2  'Center
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "01:00:00"
      BeginProperty Font
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   13590
      TabIndex        =   0
      ToolTipText     =   "Do not check time - WORK !"
      Top             =   8415
      Width           =   900
   End
End
Attribute VB_Name = "Analyzer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''=-=-*******-=-='''''''''''''''
Option Explicit
'On Error Resume Next ' cannot be here
'****** communication ************************************
'the Str and Val functions always assume a period is the decimal separator.
'In a majority of locales, this assumption is not valid. Instead, use the CStr, CDbl, CSng,
'CInt, and CLng functions to provide international conversions from any other data type
'to the data type you need. These functions use the system locale to determine the decimal separator.

    Public do_clean_buffer As Boolean   'true when called first time
    Public echo_removed As Boolean
    Public SendStrLen As Long
    Public ASCIIDataStr As String
    Public Application_Dir As String
    Public IsolatorNumber As String
    Public IsolatorAxis As String
    Public LTF_command As String
    Public saved_freq As Double

    Const DefaultTimeOut As Single = 2
    Const DefaultLTFtimeOut As Single = 20 'sec, it takes long time for 0.1 Hz
    Const end_of_TFstring_token As String = "#" & vbCrLf ' transfer function data
    Const end_of_TFtransmission As String = "*" & vbCrLf ' transfer function data END
    Const end_of_ScanString_token As String = "." & vbCrLf ' position monitoring scan data

'*********************************************************
    Dim Freq_data(TEST_LTF_ARRAY_LENGTH) As Double      ' data arrays
    Dim Prev_Freq_data(TEST_LTF_ARRAY_LENGTH) As Double ' data arrays
    Dim Reference_Gain_data(TEST_LTF_ARRAY_LENGTH) As Double
    Dim Predicted_Mag(TEST_LTF_ARRAY_LENGTH) As Double  ' this is reference plot * Axis_Mag
    Dim Gain_data(TEST_LTF_ARRAY_LENGTH) As Double      ' blue (current) plot
    Dim Prev_Gain_data(TEST_LTF_ARRAY_LENGTH) As Double ' green (previous) plot data arrays

    Dim Freq_Lim_data(TEST_LTF_ARRAY_LENGTH) As Double ' Limits on plot
    Dim Gain_Low_data(TEST_LTF_ARRAY_LENGTH) As Double
    Dim Gain_High_data(TEST_LTF_ARRAY_LENGTH) As Double
    Dim Phase_Low_data(TEST_LTF_ARRAY_LENGTH) As Double
    Dim Phase_High_data(TEST_LTF_ARRAY_LENGTH) As Double

    Dim Reference_Phase_data(TEST_LTF_ARRAY_LENGTH) As Double
    Dim Predicted_Phase(TEST_LTF_ARRAY_LENGTH) As Double
    Dim Phase_data(TEST_LTF_ARRAY_LENGTH) As Double      ' blue (current) plot
    Dim Prev_Phase_data(TEST_LTF_ARRAY_LENGTH) As Double ' green (previous) plot
    Dim Prev_Phase As Double ' from the privious line to detect jump
    Dim NegPhase_data(TEST_LTF_ARRAY_LENGTH) As Double
    Dim NegPrev_Phase_data(TEST_LTF_ARRAY_LENGTH) As Double
    Dim IsCoherenceAvailable As Boolean
    Dim Coherence_data(TEST_LTF_ARRAY_LENGTH) As Double
    Dim Prev_Coherence_data(TEST_LTF_ARRAY_LENGTH) As Double

    Dim line_counter As Long
    Dim printer_x As Long
    'Dim resonance as Double
    'Dim damping as Double
    Public in_time As String, in_date As String
    Public test_timer_ready As Boolean
    Dim StartTime As Single

    Dim Function_requested_TX As String 'diagnostic: name of the function which tries to send string

    Dim ParamLINEScount As Long
    Dim paramCHARcount As Long

    Const Transfer_Function_file As Integer = 1
    Const Alarm_file As Integer = 2
    Const Params_file As Integer = 3
    Dim tst As Long
    Dim retVal As String
    Dim percent_divider As Long ' number of LTF lines received from controller
    Dim ActiveTabIndex As Long ' index of current tab
'//define indexes for float_iir.par[index]
'    Const freq_1 As Long = 1
'    Const freq_2 As Long = 2
'    Const Q_f_1  As Long = 3
'    Const Q_f_2  As Long = 4
'    Const FilterGain   As Long = 0

    Const Connection_FrameIndex  As Long = 0
    Const SystemParams_FrameIndex As Long = 1
    Const Performance_FrameIndex As Long = 2
    Const Firmware_FrameIndex  As Long = 4
    Const Alarms_FrameIndex As Long = 5
    Const Diagnostic_FrameIndex As Long = 3
'    Const Tuning_FrameIndex  As Long = 6

    Const MAX_Prox_Sensors As Long = 6         '  number of proximity Sensors
    Dim OneScanProxValues(MAX_Prox_Sensors)    '  channel values decoded from record
'    Dim OneScanValues(MAX_Prox_Sensors)        '  channel values decoded from record
    Dim WaitTimeout As Single
    Public UserSelectedCriteriaPath As String

    Const WRONG_VALUE As Double = 77777.7
Private Declare Function WriteProfileString Lib "kernel32" _
() 'Alias "WriteProfileStringA" _
(ByVal lpszSection As String, _
ByVal lpszKeyName As String, _
ByVal lpszString As String) As Long

Private Declare Function SendMessage Lib "user32" _
() 'Alias "SendMessageA" _
(ByVal hwnd As Long, _
ByVal wMsg As Long, _
ByVal wParam As Long, _
lparam As String) As Long

Private Declare Function GetProfileString Lib "kernel32" _
() 'Alias "GetProfileStringA" _
(ByVal lpAppName As String, _
ByVal lpKeyName As String, _
ByVal lpDefault As String, _
ByVal lpReturnedString As String, _
ByVal nSize As Long) As Long

Const HWND_BROADCAST As Integer = &HFFFF
Const WM_WININICHANGE As Integer = &H1A

Dim BNC_ptr

Sub AssignChNames()
    Dim ch As Long
    ch = -1
    ch = ch + 1: ChName(ch) = "----"    '0
    ch = ch + 1: ChName(ch) = "Prx_1X_um"
    ch = ch + 1: ChName(ch) = "Prx_2Y_um"
    ch = ch + 1: ChName(ch) = "Prx_3X_um"
    ch = ch + 1: ChName(ch) = "Prx_1Z_um"
    ch = ch + 1: ChName(ch) = "Prx_2Z_um"
    ch = ch + 1: ChName(ch) = "Prx_3Z_um"
    ch = ch + 1: ChName(ch) = "prx1Hraw"
    ch = ch + 1: ChName(ch) = "prx2Hraw"
    ch = ch + 1: ChName(ch) = "prx3Hraw"
    ch = ch + 1: ChName(ch) = "pr1Hcnts"
    ch = ch + 1: ChName(ch) = "pr2Hcnts"
    ch = ch + 1: ChName(ch) = "pr3Hcnts"
    ch = ch + 1: ChName(ch) = "Prx1Zraw"
    ch = ch + 1: ChName(ch) = "Prx2Zraw"
    ch = ch + 1: ChName(ch) = "Prx3Zraw"
    ch = ch + 1: ChName(ch) = "FF_Xstg POS" '16
    ch = ch + 1: ChName(ch) = "FF_Ystg POS"
    ch = ch + 1: ChName(ch) = "FF_Xstg ACC"
    ch = ch + 1: ChName(ch) = "FF_Ystg ACC"
    ch = ch + 1: ChName(ch) = "P_iso1_cnts"
    ch = ch + 1: ChName(ch) = "P_iso2_cnts"
    ch = ch + 1: ChName(ch) = "P_iso3_cnts"
    ch = ch + 1: ChName(ch) = "P_iso4_cnts"
    ch = ch + 1: ChName(ch) = "Geo_1H"
    ch = ch + 1: ChName(ch) = "Geo_2H"
    ch = ch + 1: ChName(ch) = "Geo_3H"
    ch = ch + 1: ChName(ch) = "Geo_1Z"
    ch = ch + 1: ChName(ch) = "Geo_2Z"        '28
    ch = ch + 1: ChName(ch) = "Geo_3Z"
    ch = ch + 1: ChName(ch) = "Pressure_IN_cnts"
    ch = ch + 1: ChName(ch) = "Diag_In1"
    ch = ch + 1: ChName(ch) = "Valve_1"       '32
    ch = ch + 1: ChName(ch) = "Valve_2"
    ch = ch + 1: ChName(ch) = "Valve_3"
    ch = ch + 1: ChName(ch) = "Valve_4"
    ch = ch + 1: ChName(ch) = "Motor1Z"
    ch = ch + 1: ChName(ch) = "Motor2Z"
    ch = ch + 1: ChName(ch) = "Motor3Z"
    ch = ch + 1: ChName(ch) = "Motor4Z"
    ch = ch + 1: ChName(ch) = "Motor1H"
    ch = ch + 1: ChName(ch) = "Motor2H"
    ch = ch + 1: ChName(ch) = "Motor3H"
    ch = ch + 1: ChName(ch) = "Motor4H"
    ch = ch + 1: ChName(ch) = "DAC_ch6"
    ch = ch + 1: ChName(ch) = "DAC_ch7"
    ch = ch + 1: ChName(ch) = "BNCdiag0"
    ch = ch + 1: ChName(ch) = "BNCdiag1"
    ch = ch + 1: ChName(ch) = "X_vel"
    ch = ch + 1: ChName(ch) = "Y_vel"
    ch = ch + 1: ChName(ch) = "tZ_vel"
    ch = ch + 1: ChName(ch) = "Z_vel"
    ch = ch + 1: ChName(ch) = "tX_vel"
    ch = ch + 1: ChName(ch) = "tY_vel"
    ch = ch + 1: ChName(ch) = "X_Pos_BNC"
    ch = ch + 1: ChName(ch) = "Y_Pos_BNC"
    ch = ch + 1: ChName(ch) = "tZ_Pos_BNC"
    ch = ch + 1: ChName(ch) = "Z_Pos_BNC"
    ch = ch + 1: ChName(ch) = "tX_Pos_BNC"
    ch = ch + 1: ChName(ch) = "tY_Pos_BNC"
    ch = ch + 1: ChName(ch) = "BNC 0 Input"
    ch = ch + 1: ChName(ch) = "BNC 1 Input"
    ch = ch + 1: ChName(ch) = "---"
    ch = ch + 1: ChName(ch) = "---"
    ch = ch + 1: ChName(ch) = "ADC_00_16b" ' 64
    ch = ch + 1: ChName(ch) = "ADC_01_16b"
    ch = ch + 1: ChName(ch) = "ADC_02_16b"
    ch = ch + 1: ChName(ch) = "ADC_03_16b"
    ch = ch + 1: ChName(ch) = "ADC_04_16b"
    ch = ch + 1: ChName(ch) = "ADC_05_16b"
    ch = ch + 1: ChName(ch) = "ADC_06_16b"
    ch = ch + 1: ChName(ch) = "ADC_07_16b"
    ch = ch + 1: ChName(ch) = "ADC_08_16b"
    ch = ch + 1: ChName(ch) = "ADC_09_16b"
    ch = ch + 1: ChName(ch) = "ADC_10_16b"
    ch = ch + 1: ChName(ch) = "ADC_11_16b"
    ch = ch + 1: ChName(ch) = "ADC_12_16b"
    ch = ch + 1: ChName(ch) = "ADC_13_16b"
    ch = ch + 1: ChName(ch) = "ADC_14_16b"
    ch = ch + 1: ChName(ch) = "ADC_15_16b"
    ch = ch + 1: ChName(ch) = "ADC_16_12b"
    ch = ch + 1: ChName(ch) = "ADC_17_12b"
    ch = ch + 1: ChName(ch) = "ADC_18_12b"
    ch = ch + 1: ChName(ch) = "ADC_19_12b"
    ch = ch + 1: ChName(ch) = "ADC_20_12b"
    ch = ch + 1: ChName(ch) = "ADC_21_12b"
    ch = ch + 1: ChName(ch) = "ADC_22_12b"
    ch = ch + 1: ChName(ch) = "ADC_23_12b"
    ch = ch + 1: ChName(ch) = "ADC_24_12b"
    ch = ch + 1: ChName(ch) = "ADC_25_12b"
    ch = ch + 1: ChName(ch) = "ADC_26_12b"
    ch = ch + 1: ChName(ch) = "ADC_27_12b"
    ch = ch + 1: ChName(ch) = "ADC_28_12b"
    ch = ch + 1: ChName(ch) = "ADC_29_12b"
    ch = ch + 1: ChName(ch) = "ADC_30_12b"
    ch = ch + 1: ChName(ch) = "ADC_31_12b"
    ch = ch + 1: ChName(ch) = "DAC_00"          '96
    ch = ch + 1: ChName(ch) = "DAC_01"
    ch = ch + 1: ChName(ch) = "DAC_02"
    ch = ch + 1: ChName(ch) = "DAC_03"
    ch = ch + 1: ChName(ch) = "DAC_04"
    ch = ch + 1: ChName(ch) = "DAC_05"
    ch = ch + 1: ChName(ch) = "DAC_06"
    ch = ch + 1: ChName(ch) = "DAC_07"
    ch = ch + 1: ChName(ch) = "DAC_08"
    ch = ch + 1: ChName(ch) = "DAC_09"
    ch = ch + 1: ChName(ch) = "DAC_10"
    ch = ch + 1: ChName(ch) = "DAC_11"
    ch = ch + 1: ChName(ch) = "DAC_12"
    ch = ch + 1: ChName(ch) = "DAC_13"
    ch = ch + 1: ChName(ch) = "DAC_14"
    ch = ch + 1: ChName(ch) = "DAC_15"
    ch = ch + 1: ChName(ch) = "DAC_16"
    ch = ch + 1: ChName(ch) = "DAC_17"
    ch = ch + 1: ChName(ch) = "DAC_18"
    ch = ch + 1: ChName(ch) = "DAC_19"
    ch = ch + 1: ChName(ch) = "DAC_20"
    ch = ch + 1: ChName(ch) = "DAC_21"
    ch = ch + 1: ChName(ch) = "DAC_22"
    ch = ch + 1: ChName(ch) = "DAC_23"
    ch = ch + 1: ChName(ch) = "Excitation" '120
    ch = ch + 1: ChName(ch) = "zero"
    ch = ch + 1: ChName(ch) = "zero"
    ch = ch + 1: ChName(ch) = "XaccDetect"
    ch = ch + 1: ChName(ch) = "YaccDetect"
    ch = ch + 1: ChName(ch) = "XposDetect"
    ch = ch + 1: ChName(ch) = "YposDetect"
    ch = ch + 1: ChName(ch) = "zero"
    ch = ch + 1: ChName(ch) = "Xv-Filt0"
    ch = ch + 1: ChName(ch) = "Xv-Filt1"
    ch = ch + 1: ChName(ch) = "Xv-Filt2"
    ch = ch + 1: ChName(ch) = "Xv-Filt3"
    ch = ch + 1: ChName(ch) = "Xv-Filt4"
    ch = ch + 1: ChName(ch) = "Xv-Filt5"
    ch = ch + 1: ChName(ch) = "Yv-Filt0"
    ch = ch + 1: ChName(ch) = "Yv-Filt1"
    ch = ch + 1: ChName(ch) = "Yv-Filt2"
    ch = ch + 1: ChName(ch) = "Yv-Filt3"
    ch = ch + 1: ChName(ch) = "Yv-Filt4"
    ch = ch + 1: ChName(ch) = "Yv-Filt5"
    ch = ch + 1: ChName(ch) = "tZv-Filt0"
    ch = ch + 1: ChName(ch) = "tZv-Filt1"
    ch = ch + 1: ChName(ch) = "tZv-Filt2"
    ch = ch + 1: ChName(ch) = "tZv-Filt3"
    ch = ch + 1: ChName(ch) = "tZv-Filt4"
    ch = ch + 1: ChName(ch) = "tZv-Filt5"
    ch = ch + 1: ChName(ch) = "Zv-Filt0"
    ch = ch + 1: ChName(ch) = "Zv-Filt1"
    ch = ch + 1: ChName(ch) = "Zv-Filt2"
    ch = ch + 1: ChName(ch) = "Zv-Filt3"
    ch = ch + 1: ChName(ch) = "Zv-Filt4"
    ch = ch + 1: ChName(ch) = "Zv-Filt5"
    ch = ch + 1: ChName(ch) = "tXv-Filt0"        '152
    ch = ch + 1: ChName(ch) = "tXv-Filt1"
    ch = ch + 1: ChName(ch) = "tXv-Filt2"
    ch = ch + 1: ChName(ch) = "tXv-Filt3"
    ch = ch + 1: ChName(ch) = "tXv-Filt4"
    ch = ch + 1: ChName(ch) = "tXv-Filt5"
    ch = ch + 1: ChName(ch) = "tYv-Filt0"
    ch = ch + 1: ChName(ch) = "tYv-Filt1"
    ch = ch + 1: ChName(ch) = "tYv-Filt2"
    ch = ch + 1: ChName(ch) = "tYv-Filt3"
    ch = ch + 1: ChName(ch) = "tYv-Filt4"
    ch = ch + 1: ChName(ch) = "tYv-Filt5"
    ch = ch + 1: ChName(ch) = "Xp-Filt0"
    ch = ch + 1: ChName(ch) = "Xp-Filt1"
    ch = ch + 1: ChName(ch) = "Xp-Filt2"
    ch = ch + 1: ChName(ch) = "Xp-Filt3"
    ch = ch + 1: ChName(ch) = "Xp-Filt4"
    ch = ch + 1: ChName(ch) = "Xp-Filt5"
    ch = ch + 1: ChName(ch) = "Yp-Filt0"
    ch = ch + 1: ChName(ch) = "Yp-Filt1"
    ch = ch + 1: ChName(ch) = "Yp-Filt2"
    ch = ch + 1: ChName(ch) = "Yp-Filt3"
    ch = ch + 1: ChName(ch) = "Yp-Filt4"
    ch = ch + 1: ChName(ch) = "Yp-Filt5"
    ch = ch + 1: ChName(ch) = "tZp-Filt0"        '176
    ch = ch + 1: ChName(ch) = "tZp-Filt1"
    ch = ch + 1: ChName(ch) = "tZp-Filt2"
    ch = ch + 1: ChName(ch) = "tZp-Filt3"
    ch = ch + 1: ChName(ch) = "tZp-Filt4"
    ch = ch + 1: ChName(ch) = "tZp-Filt5"
    ch = ch + 1: ChName(ch) = "Zp-Filt0"
    ch = ch + 1: ChName(ch) = "Zp-Filt1"
    ch = ch + 1: ChName(ch) = "Zp-Filt2"
    ch = ch + 1: ChName(ch) = "Zp-Filt3"
    ch = ch + 1: ChName(ch) = "Zp-Filt4"
    ch = ch + 1: ChName(ch) = "Zp-Filt5"
    ch = ch + 1: ChName(ch) = "tXp-Filt0"
    ch = ch + 1: ChName(ch) = "tXp-Filt1"
    ch = ch + 1: ChName(ch) = "tXp-Filt2"
    ch = ch + 1: ChName(ch) = "tXp-Filt3"
    ch = ch + 1: ChName(ch) = "tXp-Filt4"
    ch = ch + 1: ChName(ch) = "tXp-Filt5"
    ch = ch + 1: ChName(ch) = "tYp-Filt0"
    ch = ch + 1: ChName(ch) = "tYp-Filt1"
    ch = ch + 1: ChName(ch) = "tYp-Filt2"
    ch = ch + 1: ChName(ch) = "tYp-Filt3"
    ch = ch + 1: ChName(ch) = "tYp-Filt4"
    ch = ch + 1: ChName(ch) = "tYp-Filt5"        '199
    ch = ch + 1: ChName(ch) = "XaFF-Filt0"       '200
    ch = ch + 1: ChName(ch) = "XaFF-Filt1"
    ch = ch + 1: ChName(ch) = "XaFF-Filt2"
    ch = ch + 1: ChName(ch) = "XaFF-Filt3"
    ch = ch + 1: ChName(ch) = "XaFF-Filt4"
    ch = ch + 1: ChName(ch) = "XaFF-Filt5"
    ch = ch + 1: ChName(ch) = "YaFF-Filt0"
    ch = ch + 1: ChName(ch) = "YaFF-Filt1"
    ch = ch + 1: ChName(ch) = "YaFF-Filt2"
    ch = ch + 1: ChName(ch) = "YaFF-Filt3"
    ch = ch + 1: ChName(ch) = "YaFF-Filt4"
    ch = ch + 1: ChName(ch) = "YaFF-Filt5"
    ch = ch + 1: ChName(ch) = "XpFF-Filt0"       '212
    ch = ch + 1: ChName(ch) = "XpFF-Filt1"
    ch = ch + 1: ChName(ch) = "XpFF-Filt2"
    ch = ch + 1: ChName(ch) = "XpFF-Filt3"
    ch = ch + 1: ChName(ch) = "XpFF-Filt4"
    ch = ch + 1: ChName(ch) = "XpFF-Filt5"
    ch = ch + 1: ChName(ch) = "YpFF-Filt0"
    ch = ch + 1: ChName(ch) = "YpFF-Filt1"
    ch = ch + 1: ChName(ch) = "YpFF-Filt2"
    ch = ch + 1: ChName(ch) = "YpFF-Filt3"
    ch = ch + 1: ChName(ch) = "YpFF-Filt4"
    ch = ch + 1: ChName(ch) = "YpFF-Filt5"       '223
    ch = ch + 1: ChName(ch) = "PSI_I1fil0"        '224
    ch = ch + 1: ChName(ch) = "PSI_I1fil1"
    ch = ch + 1: ChName(ch) = "PSI_I1fil2"
    ch = ch + 1: ChName(ch) = "PSI_I2fil0"
    ch = ch + 1: ChName(ch) = "PSI_I2fil1"
    ch = ch + 1: ChName(ch) = "PSI_I2fil2"
    ch = ch + 1: ChName(ch) = "PSI_I3fil0"
    ch = ch + 1: ChName(ch) = "PSI_I3fil1"
    ch = ch + 1: ChName(ch) = "PSI_I3fil2"
    ch = ch + 1: ChName(ch) = "PSI_I4fil0"
    ch = ch + 1: ChName(ch) = "PSI_I4fil1"
    ch = ch + 1: ChName(ch) = "PSI_I4fil2"
    ch = ch + 1: ChName(ch) = "Xpos-->Mot"
    ch = ch + 1: ChName(ch) = "Ypos-->Mot"
    ch = ch + 1: ChName(ch) = "tZpos->Mot"
    ch = ch + 1: ChName(ch) = "Zpos-Valve"
    ch = ch + 1: ChName(ch) = "tXpos-Valve"
    ch = ch + 1: ChName(ch) = "tYpos-Valve"
    ch = ch + 1: ChName(ch) = "Xvel-->Mot"
    ch = ch + 1: ChName(ch) = "Yvel-->Mot"
    ch = ch + 1: ChName(ch) = "tZvel->Mot"
    ch = ch + 1: ChName(ch) = "Zvel-->Mot"
    ch = ch + 1: ChName(ch) = "tXvel->Mot"
    ch = ch + 1: ChName(ch) = "tYvel->Mot"
    ch = ch + 1: ChName(ch) = "FFout->tXpos"     '248
    ch = ch + 1: ChName(ch) = "FFout->tYpos"
    ch = ch + 1: ChName(ch) = "FFout->X_vel"
    ch = ch + 1: ChName(ch) = "FFout->Y_vel"
    ch = ch + 1: ChName(ch) = "FFout->tZvel"
    ch = ch + 1: ChName(ch) = "FFout->tXvel"
    ch = ch + 1: ChName(ch) = "FFout->tYvel"
    ch = ch + 1: ChName(ch) = "zero"            '255
    CmbBNC0.Clear
    CmbBNC1.Clear
    For ch = 0 To UBound(ChName)
        CmbBNC0.AddItem (ChName(ch))
        CmbBNC1.AddItem (ChName(ch))
    Next ch
End Sub

Private Sub ChkAirIsolatorsPresent_Click()
    If (ChkAirIsolatorsPresent.Value = Checked) Then
        AirIsoPresent = True
        ChkAirIsolatorsPresent.BackColor = &HFFFFC0 ' light blue
        ChkAirIsolatorsPresent.Caption = "Air Isolators"
        FramePneumFB.Visible = True
        FrmVertPosToMotGains.Visible = False
        If IsFormLoaded("frmFilters") Then
            frmFilters.FramePneumatic.Caption = "Air Isolators"     'IK20251112 this loads form!
            frmFilters.FramePneumatic.Enabled = True
        End If
    Else
        AirIsoPresent = False
        FramePneumFB.Visible = False
        FrmVertPosToMotGains.Visible = True
        If IsFormLoaded("frmFilters") Then
            frmFilters.FramePneumatic.Caption = "NO AIR Isolators"  'IK20251112 this loads form!
            frmFilters.FramePneumatic.Enabled = False               'IK20251112 this loads form!
        End If
        ChkAirIsolatorsPresent.BackColor = &H80FF&  ' light brown
        ChkAirIsolatorsPresent.Caption = "NO Air Isolators"
    End If
End Sub

Private Sub ChkShowCoherencePlot_Click()
    If (ChkShowCoherencePlot.Value = Checked) Then
        If (ChkShowCurrentPlot.Value = Checked) Then CWGraphPhase.Plots(Coherence_TF_Num).Visible = True
        If (ChkShowPrevPlot.Value = Checked) Then CWGraphPhase.Plots(PrevCoherence_TF_Num).Visible = True
    Else
        CWGraphPhase.Plots(Coherence_TF_Num).Visible = False
        CWGraphPhase.Plots(PrevCoherence_TF_Num).Visible = False
    End If
End Sub

Private Sub ChkShowLimits_Click()
    If (ChkShowLimits.Value = Checked) Then
'        Call Check_Selected_Actuator_Sensor ' this causes to re-read limits each time they show
        CWGraph.Plots(Lower_Lim_TF_Num).Visible = True
        CWGraph.Plots(Upper_Lim_TF_Num).Visible = True
        CWGraphPhase.Plots(Lower_Lim_TF_Num).Visible = True
        CWGraphPhase.Plots(Upper_Lim_TF_Num).Visible = True
    Else
        CWGraph.Plots(Lower_Lim_TF_Num).Visible = False
        CWGraph.Plots(Upper_Lim_TF_Num).Visible = False
        CWGraphPhase.Plots(Lower_Lim_TF_Num).Visible = False
        CWGraphPhase.Plots(Upper_Lim_TF_Num).Visible = False
    End If
End Sub

Private Sub Cmd_STOP_test_Click()
    response = Analyzer.GetSend("test>stop", True) 'stop test
    'Call Read_Alarm_Status
End Sub


Private Sub ChkCommIssues_Click()
    If (ChkCommIssues.Value = Checked) Then
        lstResponse.Height = 2535
        ListCommIssues.Visible = True
    Else
        lstResponse.Height = 4560
        ListCommIssues.Visible = False
    End If
End Sub


Private Sub CmbBNC0_Click()
  Dim command As String
'  Dim BNC_index As Long
    If Init_sys_data = True Then Exit Sub
    BNC_0_index = CmbBNC0.ListIndex
    command = CmbBNC0.Tag & CStr(BNC_0_index)
    cmd_sent = False
    waitresponse = True
    response = GetSend(command, True)
    command = Find_Index_Match_with_BNC_SLIDER(0, BNC_0_index)
End Sub


Sub clear_arrays_and_plots()
    Dim i As Long
    For i = 1 To TEST_LTF_ARRAY_LENGTH
        Gain_data(i) = 0
        Phase_data(i) = 0
        Coherence_data(i) = 0
        Prev_Gain_data(i) = 0
        Prev_Phase_data(i) = 0
        Prev_Coherence_data(i) = 0
        NegPhase_data(i) = Make180phaseShift(0)
        Prev_Freq_data(i) = 0
    Next i
    CWGraph.ClearData
    CWGraphPhase.ClearData
End Sub

Private Sub CmdClearPlots_Click()
    Call clear_arrays_and_plots
End Sub

Private Sub CmdCopyClipboard_Click()
    Call Copy_response_to_clipboard
End Sub


Private Sub CmdSaveResponsesToPC_Click()
    Call Save_Responses_to_PC
End Sub

Private Sub CmdReadLimits_Click()
    UserSelectedCriteriaPath = "" ' clear path to force re-selection file of folder
    Call Check_Selected_Actuator_Sensor
End Sub

Private Sub lstResponse_KeyPress(KeyAscii As Integer)
   Dim char As String
    char = Chr(KeyAscii)
'    sym = char ' for testing purpuses this sub is changing sym
    str_Help = "KeyAscii, Char = " & KeyAscii & ", " & char
    'KeyAscii = Asc(UCase(char))
    Select Case KeyAscii
        Case 13: 'Enter 'entering new variable value in the menu
'        KeyAscii = InStr(1, ScreenLine(line_Y_index), ". ", 0) - 1 'binary search & compare
'        If KeyAscii > 0 Then
'           TerminalSend(Mid(ScreenLine(line_Y_index), KeyAscii, 1))
'              lblHelp.Caption = "symbol sent: " & Mid(ScreenLine(line_Y_index), KeyAscii, 1)
'        End If
        Case 3:  'CTRL-C
            Call Copy_response_to_clipboard
            Exit Sub
    End Select
End Sub

Public Sub Copy_response_to_clipboard()
    Dim N As Long
    Dim tmp_str As String
    For N = 0 To lstResponse.ListCount
        tmp_str = tmp_str & lstResponse.List(N) & vbCrLf
    Next N
    Clipboard.Clear
    Clipboard.SetText tmp_str
End Sub

Sub Save_Responses_to_PC()
    On Error Resume Next
    'user wants to save a test file
    'First get the filename
    Dim DefaultFilename As String
    Dim SaveFileName As String
    Dim FileName As String
    Dim i As Long
    Dim colon_pos As Long
    Dim param_line As String


   DefaultFilename = "TMC communication " & Date$ & " " & Left$(Time$, 2) & "-" & Mid$(Time$, 4, 2) & ".txt"
    dlgCommonDialog.CancelError = True
    dlgCommonDialog.DialogTitle = "Choose name parameters file"
    ' dlgCommonDialog.FileName = DefaultFilename
    dlgCommonDialog.InitDir = "c:\test"
    dlgCommonDialog.DefaultExt = "txt"
    dlgCommonDialog.Filter = "File (*.txt)|*.param|All files (*.*)|*.*"
    dlgCommonDialog.FileName = DefaultFilename 'if enabled,  FILE NAME SHOULD NOT HAVE ':'
    dlgCommonDialog.ShowSave
    If Err = 32755 Then 'Cancel is selected
        Exit Sub
    Else
        txtStatus.Text = "File(s)=" & dlgCommonDialog.FileName
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
    Open SaveFileName For Output As Params_file 'parameters
    For i = 0 To lstResponse.ListCount  ' or -3 to remove \r\n
        param_line = lstResponse.List(i)
        Print #Params_file, param_line
    Next i
    Close #Params_file
End Sub

Private Sub cmdSetControllerTime_Click()
    Call GetSend("Time=" & in_time, True)   'set time
End Sub

Private Sub cwBut_AuxDamping_ValueChanged(ByVal Value As Boolean)
    If Init_sys_data = True Then Exit Sub 'avoid firing control event subroutine during initialization when values assigned
    If cwBut_AuxDamping.Value = True Then
        Call GetSend(cwBut_AuxDamping.Tag & ">active", True)
    Else
        Call GetSend(cwBut_AuxDamping.Tag & ">passive", True)
    End If
    Call read_RT_diagnostic
End Sub

Private Sub cwBut_BasicCMD_ValueChanged(Index As Integer, ByVal Value As Boolean)
    Dim cmd   As String
    Dim retVal As String
    cmd = ""
    If Index = 0 Then 'floating - docking
        If Value = False Then
            cmd = "u"  'float UP
        Else
            cmd = "l" 'Lowering to dock
        End If
    End If
'    If Index = 1 Then ' do test
'        If Value = True Then
'            cmd = "excit>stop" 'quit excitation test
'        Else
'            cmd = "excit>start"  'do excitation TEST
'        End If
'    End If
'    If Index = 2 Then ' no index 2
'    End If
'    If Index = 3 Then ' Motor FB
'        If Value = True Then
'            cmd = "o"  'enable all
'        Else
'            cmd = "x" 'disable all
'        End If
'    End If
    If Index = 4 Then ' Dither stop Hold
        If Value = True Then
            cmd = "c"  'dither ON
        Else
            cmd = "h" 'dither HOLD
        End If
    End If
'    If Index = 5 Then ' Horizontal Hold
'        If Value = True Then
'            cmd = "f" 'release soft dock
'        Else
'            cmd = "d"  'soft DOCKING
'        End If
'    End If
'    If Index = 6 Then ' no index 6
'    End If
'    If Index = 7 Then ' Feed Forward
'        If Value = True Then
'            cmd = "e"  'FF enabled
'        Else
'            cmd = "k" 'FF disabled
'        End If
'    End If
    If Init_sys_data = False Then
        Call Analyzer.GetSend(cmd, True)
    End If
    Call read_RT_diagnostic
End Sub

Sub ShowSysStatus()
    retVal = GetSend("s", True)  'send command
    Decode_SysStatus (retVal)
End Sub

Function get_ID() As Boolean
    Dim temp_str As String
    Dim pos As Long
    get_ID = False
    temp_str = GetSend("vers?", True) 'send and wait for response
    If temp_str <> "Timeout" Then
        If InStr(temp_str, "95-") <> 0 Then 'valid string with pn 95-xxxxx
            txtFW_Version.Text = temp_str
            get_ID = True
        Else 'wrong response
            StartTime = Timer
            Do 'small delay
                DoEvents
            Loop Until (Timer - StartTime > 0.05)
            temp_str = GetSend("vers?", True) 'try to send second time and wait for response
            If InStr(temp_str, "95-") <> 0 Then 'valid string with pn 95-xxxxx
    '        If Len(temp_str) > 9 Then 'valid string
                txtFW_Version.Text = temp_str
                get_ID = True
            Else
                get_ID = False
                Exit Function
            End If
        End If
    End If
    If get_ID = True Then
'        pos = InStr(temp_str, vbCrLf)
        txtFW_Version.Text = temp_str
    Else
        txtFW_Version.Text = "No or not valid response"
    End If
    SystemFWversion = txtFW_Version.Text

    temp_str = GetSend("boot_rev", True) 'send and wait for response
    temp_str = UCase(temp_str)
    If InStr(temp_str, ">~U") <> 0 Then '>~Unrecognized, older FW
        txtBootLoader_Rev.Text = " N/A, OLD FW"
    ElseIf InStr(temp_str, ">~E") <> 0 Then
        txtBootLoader_Rev.Text = " NO BootLoader"
    Else
        txtBootLoader_Rev.Text = " " & temp_str ' 'A' or 'B', 'C', etc.
    End If

    temp_str = GetSend("`vers?", True) 'send new preffix check if it is accepted
    temp_str = LCase(temp_str)
    If InStr(temp_str, ">~un") <> 0 Then 'firmware does not support preffix? try once more
        StartTime = Timer
        Do 'small delay
            DoEvents
        Loop Until (Timer - StartTime > 0.05)
        temp_str = GetSend("`vers?", True) 'try to send second time and wait for response
        temp_str = LCase(temp_str)
        If InStr(temp_str, ">~un") <> 0 Then 'firmware does not support preffix
            FW_supports_menu_preffix = False
        Else
            FW_supports_menu_preffix = True
        End If
    Else 'correct response
            FW_supports_menu_preffix = True
    End If
    If (InStr(UCase(SystemFWversion), "PONG") <> 0) Then
        HasFloorFF = True
        ChkAirIsolatorsPresent.Value = Unchecked    ' IK20251112 NO AIR auto set based on the presence of "PONG" in version string
    Else
        HasFloorFF = False
        ChkAirIsolatorsPresent.Value = Checked    ' IK20251112 NO AIR auto set based on the presence of "PONG" in version string
    End If
End Function


Private Sub cwBut_FF_ALL_ValueChanged(ByVal Value As Boolean)
    If Init_sys_data = True Then Exit Sub 'avoid firing control event subroutine during initialization when values assigned
    If cwBut_FF_ALL.Value = True Then
        Call GetSend(cwBut_FF_ALL.Tag & ">active", True)
    Else
        Call GetSend(cwBut_FF_ALL.Tag & ">passive", True)
    End If
    Call read_RT_diagnostic
End Sub

'==================== POSITION FRAME ==================================================================================================
Private Sub cwBut_SoftDock_ValueChanged(ByVal Value As Boolean)
    If Init_sys_data = True Then Exit Sub 'avoid firing control event subroutine during initialization when values assigned
    If cwBut_SoftDock.Value = True Then
        Call GetSend(cwBut_SoftDock.Tag & ">hold", True)
    Else
        Call GetSend(cwBut_SoftDock.Tag & ">free", True)
    End If
    Call read_RT_diagnostic
End Sub

Private Sub CmdZeroHorizProx_Click()
    Call GetSend("z", True)   'send zero proximity command, updates proximity offsets
End Sub

Private Sub CmdResetGeoOffsets_Click() 'IK20251217 added
    Call GetSend("zero", True)   'send zero all - updates SysData.bias_to_zero[i] = C28v.ADC_buff[i];
End Sub

Private Sub cwNum_DockedPos_ValueChanged(Value As Variant, PreviousValue As Variant, ByVal OutOfRange As Boolean)
'"dkpo=XXXX" - will set VERTICAL docking position, query "dkpo?\r" will respond with "dkpo=XXXX\r\n"; XXXX - integer microns, range 0-9999
'Positive distance down from "floating" till "docked" position, microns; range 0 to 8000. Real depends on setup: for Turck sensors use <=2000, for AutomationDirect use <=8000
    Dim cmd As String
    If Init_sys_data = True Then Exit Sub
    cmd = cwNum_DockedPos.Tag & "=" & cwNum_DockedPos.Text
    retVal = Analyzer.GetSend(cmd, True)
End Sub

Private Sub cwNum_Go_NoGO_window_ValueChanged(Index As Integer, Value As Variant, PreviousValue As Variant, ByVal OutOfRange As Boolean)
    Dim cmd As String
    If Init_sys_data = True Then Exit Sub
    cmd = cwNum_Go_NoGO_window(Index).Tag & "=" & cwNum_Go_NoGO_window(Index).Text
    retVal = Analyzer.GetSend(cmd, True)
End Sub

Private Sub cwNum_PressureSetPoint_ValueChanged(Index As Integer, Value As Variant, PreviousValue As Variant, ByVal OutOfRange As Boolean)
  On Error Resume Next
'gain31=+34.300 // Iso1 last 'good' floating pressure
'gain32=+37.800 // Iso2 last 'good' floating pressure
'gain33=+40.100 // Iso3 last 'good' floating pressure
'gain34=+41.900 // Iso4 last 'good' floating pressure
    Dim cmd As String
    If Init_sys_data = True Then Exit Sub
    cmd = cwNum_PressureSetPoint(Index).Tag & "=" & cwNum_PressureSetPoint(Index).Text
    retVal = Analyzer.GetSend(cmd, True)
End Sub

Private Sub cwNumFBgain_ValueChanged(Index As Integer, Value As Variant, PreviousValue As Variant, ByVal OutOfRange As Boolean)
    Dim gain_cmd As String
    If Init_sys_data = True Then Exit Sub
    gain_cmd = cwNumFBgain(Index).Tag & "=" & cwNumFBgain(Index).Text
    retVal = Analyzer.GetSend(gain_cmd, True)
End Sub

Private Sub cwNumHeightAdj_ValueChanged(Index As Integer, Value As Variant, PreviousValue As Variant, ByVal OutOfRange As Boolean)
    Dim offset_cmd As String
    If Init_sys_data = True Then Exit Sub
    offset_cmd = cwNumHeightAdj(Index).Tag & "=" & cwNumHeightAdj(Index).Text
    retVal = Analyzer.GetSend(offset_cmd, True)
End Sub

'==================== END POSITION FRAME ==================================================================================================

'=================================================================================================================================
Sub ReadBNCchannelsFromSlider()
    Dim ch_index As Long
    Dim index_array() As Long
    Dim slider_channels As Long
    Dim MAX_Channels As Long
    MAX_Channels = cwBNC1select.axis.Maximum '.ValuePairs.Count
    ReDim index_array(MAX_Channels)
    For slider_channels = 1 To MAX_Channels
        ch_index = cwBNC1select.axis.ValuePairs(slider_channels).Value
        index_array(slider_channels - 1) = ch_index
        BNC_ptr(slider_channels - 1) = ch_index
    Next slider_channels
End Sub

Private Sub Form_Initialize() ' happens very first time when an instance of the form is created in your application
    percent_divider = 141 ' approximate number of LTF lines received from controller
    waitresponse = True

    If strTelnet_IP_ADR = "" Then strTelnet_IP_ADR = "169.254.20.20"
    If strTelnet_PORT = "" Then strTelnet_PORT = "2020"
    ID_valid = False
End Sub

Private Sub Form_Load() 'happens after form_initalize
    Dim zero_Xline(2) As Double
    Dim zero_Yline(2) As Double
    Dim temp_str As String

    zero_Xline(0) = 0.1
    zero_Xline(1) = 5000
    zero_Yline(0) = 0
    zero_Yline(1) = 0
    CWGraph.Plots(PreviousPlot_Num).PlotXvsY zero_Xline, zero_Yline
    CWGraphPhase.Plots(PreviousPlot_Num).PlotXvsY zero_Xline, zero_Yline

    BNC_ptr = Array(0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0)  ' last "minus one" means "other", not on slider control
' should be copied from slider control, here is manual copying
'    BNC_ptr = Array(120, 59, 58, 57, 56, 55, 54, 53, 52, 51, 50, 49, 48, 39, 38, 37, 36, 43, 42, 41, 40, 29, 28, 27, 26, 25, 24, 17, 16, 19, 18, 6, 5, 4, 3, 2, 1, -1)  ' last one means "other", not on slider control
    Call ReadBNCchannelsFromSlider ' automatic update from BNC Slider
    Call AssignChNames 'DEFAULT NAMES if it is called from Form_Initialize() it forces to call Form_Load()

    If (ConnectionType = ConnDEMO) Then
        cwCOMdisconnect.Caption = strPortNum & strPortSetting 'if it is called from Form_Initialize() it forces to call Form_Load()
        cwCOMdisconnect.Value = False
    End If
    If (ConnectionType = ConnSimDLL) Then
        cwCOMdisconnect.Caption = strPortNum
        cwCOMdisconnect.Value = True
        If (DLL_is_Initialized = False) Then Init_DLL
    End If
    If (ConnectionType = ConnCOMport Or ConnectionType = ConnTelnet) Then
        cwCOMdisconnect.Caption = strPortNum
        cwCOMdisconnect.Value = True
    End If
' here cwCOMdisconnect.Value gets changed, but envent ValueChanged is not triggered?
    If cwCOMdisconnect.Value = True Then 'open
        ImageConnected.Visible = True
        ImageDisconnected.Visible = False
        LblCONNparams.Caption = cwCOMdisconnect.Caption & " open "
    Else 'closed
        ImageDisconnected.Visible = True
        ImageConnected.Visible = False
        LblCONNparams.Caption = cwCOMdisconnect.Caption & " closed"
    End If

    Me.Show

    DiagnosticAxis = 6 ' 6 is the index of "X axis Vel gain"
' if uncommented, terminal form will be loaded because it calls its content
    'frmTerminal.ReadPortSetup
'    frmTerminal.Caption = "initializing" 'call some function to activate
'    If strPortStatus = "" Then
'        InitCommPort (1) 'initialization of COM4
'    End If
    ActiveTabIndex = Connection_FrameIndex 'SystemParams_FrameIndex
    Call bringControlsToFront_True_toBack_False(False)
        FrameSystem.Visible = False
        FramePerformance.Visible = False
        frmDiagnostic.Visible = False
        frmFWupgrade.Visible = False
        FrameAlarms.Visible = False
        FrmConnection.Visible = True
    Me.Show
    If strPortStatus = " open" Then
        Init_sys_data = True

        ActiveTabIndex = SystemParams_FrameIndex
        FrmConnection.Visible = False
        FrameSystem.Visible = True
        Call bringControlsToFront_True_toBack_False(True)

        GUI_doing_command = False 'enable "oltf>stop" command or it will wait and give error

        Call frmTerminal.QuitMenu ' menu mode exit just in case, to get parameters

        Call GetSend("echo>enab", False) 'disable verbose mode - it messes with settings update. send and wait between chars leave advanced menu, if it is there
        ID_valid = get_ID() 'this sends 'vers' and looks for "95-" gets and shows system type, corrects input and output names for SEM_base vs STACIS (H1 vs X1, V1 vs Z1)
        If ID_valid = False Then
            Init_sys_data = False
            Exit Sub
        End If
        ' names were updated by call get_ID(), reset names to force user to select output and input
        LTFtestOutputName = "" 'cmdChooseExcitation.Caption
        LTFtestInputName = "" 'CmdChooseInputSignal.Caption '

        Call ShowDialogWindow     ' calls GetSysData, which calls frmTerminal.QuitMenu, frmTerminal.TerminalSend, Module1.Send_GUI_command
        Call ReadChannelNamesFromController   ' READ real names from controller
        in_time = Format(Time, "HH:MM:SS")
        Call GetSend("Time=" & in_time, True)   'set controller time from PC time
        Call read_RT_diagnostic
    Else 'means strPortStatus <> " open"
        Init_sys_data = False
    End If
    Init_sys_data = False ' init success
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If DLL_is_Initialized Or DLL_runs_in_separate_thread Then
'        Cmd_Stop_DLL ' it will be done if user decided to leave in Form_Unload
'        DoEvents
    End If
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Dim i As Long
    Dim CmdStartTime As Single
    Dim WaitTime As Single

    i = MsgBox("Close Test program - are you sure?", vbYesNo + vbExclamation + vbDefaultButton2, "TMC Analyzer CLOSING")
    If i = vbNo Then
        Cancel = True
        Exit Sub 'it exits right here if "Cancel" is not set True
    End If
    Close ' all open files

    If DLL_is_Initialized Or DLL_runs_in_separate_thread Then
        Call frmTerminal.QuitMenu ' or VB6 IDE will crush if DLL in menu mode
        Cmd_Stop_DLL 'wait a few ms
        CmdStartTime = Timer
        Do 'small delay
            DoEvents '  here another call of GetSend might happen
'this does frmTerminal.Receive if GUI_doing_command if true
            WaitTime = Timer - CmdStartTime
        Loop Until (WaitTime > 0.01)
    End If

    For i = Forms.Count - 1 To 0 Step -1
'            Forms(i).Show
'            MsgBox ("form #" & i & ", " & Forms(i).Caption)
        Unload (Forms(i))
    Next i
    End
End Sub

'Private Sub Form_Terminate()
'    If Is_DLL_running Then
'        Cmd_Stop_DLL
'        DoEvents
'    End If
'    End
'End Sub

Function GetAlarmWord() As Boolean
'// "alarm>word";        printf("%08X",M3v.AlarmWord);
'  cputs(      "// n_a n_a n_a n_a n_a n_a n_a n_a n_a n_a Z3  H3  Z2  H2  Z1  H1\r\n    ");
'  cputs("\r\n\n// n_a n_a z3H h3H z2H h2H z1H h1H n_a n_a z3L h3L z2L h2L z1L h1L\r\n    ");
    Dim b As Long
    Dim mask_bit As Long

    cmd_sent = False
    response = Analyzer.GetSend("alarm>word", True) 'gets back HEX 00XXXXXX
    b = InStr(response, "//")
    If b <> 0 Then response = Left(response, b - 1)
    response = LTrim(response)
    response = RTrim(response)

    AlarmWord = CLng("&H" & response)

    'Call Show_Alarm_flags
'    mask_bit = 1
'    For b = 0 To 23
'        If (AlarmWord And mask_bit) <> 0 Then
'            CheckAlarmBit(b).Value = Checked   'Alarm bit
'        Else
'            CheckAlarmBit(b).Value = Unchecked 'OK
'        End If
'        mask_bit = mask_bit * 2
'    Next b
End Function


Private Sub ImageConnected_Click()
    If (cwCOMdisconnect.Value = True) Then
        cwCOMdisconnect.Value = False
'        ImageConnected.Visible = False
'        ImageDisconnected.Visible = True
    Else
'        ImageDisconnected.Visible = False
'        ImageConnected.Visible = True
    End If
End Sub

Private Sub CommandDiagRefresh_Click()
    Call GetSysData
    Call read_RT_diagnostic
End Sub

Private Sub ImageDiagTuning_Click(Index As Integer)
    Call GetSysData
End Sub

Private Sub ImageDisconnected_Click()
    If (cwCOMdisconnect.Value = False) Then
        cwCOMdisconnect.Value = True
'        ImageDisconnected.Visible = False
'        ImageConnected.Visible = True
    Else
'        ImageConnected.Visible = False
'        ImageDisconnected.Visible = True
    End If
End Sub

Sub SCOPE_start()
    Dim appPath As String
    If Windows_is_64bit Then
        appPath = "C:\Program Files (x86)\TMC\DC 2020 Scope\"
    Else
        appPath = "C:\Program Files\TMC\DC 2020 Scope\"
    End If
    appPath = appPath + "DC-2020_Scope.exe"
    If FileExists(appPath) Then
'        LblScope.ForeColor = &HC000C0   'dark magenta
'        LblScope.Font.Size = 11     ' increase font size
        LblScope.BackStyle = 1      ' opaque
        Call Shell(appPath, vbNormalFocus)
    Else
'        LblScope.ForeColor = &H404040   'dark gray
'        LblScope.Font.Size = 9      ' decrease font size
        LblScope.BackStyle = 0      ' transparent
        Call MsgBox("TMC DC-2020 scope is not found at path" & vbCrLf & appPath, vbCritical, "Scope not available")
    End If
End Sub

Sub Toggle_FB_FFgains_Controls()
    If FormGainsVisible = False And frmStageFFgains.Visible = False And frmFloorFF.Visible = False Then
        LblFB_FF_Gains.Font.Size = 11
        LblFB_FF_Gains.ForeColor = &HFF0000 ' dark blue
        LblFB_FF_Gains.BackStyle = 1      ' opaque
        FormGainsVisible = True
        If HasFloorFF = True Then
            frmFloorFF.Show
        Else
            frmStageFFgains.Show
        End If
    Else
        frmStageFFgains.Hide
        frmFloorFF.Hide
        FormGainsVisible = False
        LblFB_FF_Gains.Font.Size = 9
        LblFB_FF_Gains.ForeColor = vbBlack
        LblFB_FF_Gains.BackStyle = 0      ' transparent
    End If
End Sub

Private Sub ImageGains_Click()
    Call Toggle_FB_FFgains_Controls
End Sub

Private Sub LblFB_FF_Gains_Click()
    Call Toggle_FB_FFgains_Controls
End Sub

Private Sub ImageScope_Click()
    Call SCOPE_start
End Sub

Private Sub CmdScope_Click()
    Call SCOPE_start
End Sub

Private Sub LblScope_Click()
    Call SCOPE_start
End Sub

Sub ToggleFiltersControls()
    If FormFiltersVisible = False And frmFilters.Visible = False Then
        LblFilters.Font.Size = 11
        LblFilters.ForeColor = &H40C0 ' dark red
        LblFilters.BackStyle = 1      ' opaque
        FormFiltersVisible = True
        frmFilters.Show
    Else
        frmFilters.Hide
        FormFiltersVisible = False
        LblFilters.Font.Size = 9
        LblFilters.ForeColor = vbBlack
        LblFilters.BackStyle = 0      ' transparent
    End If
End Sub

Private Sub ImageFilters_Click()
    Call ToggleFiltersControls
End Sub

Private Sub LblFilters_Click()
    Call ToggleFiltersControls
End Sub

Sub ToggleTerminalControls()
    If FormTerminalVisible = False And frmTerminal.Visible = False Then
        LblTerminalWindow.Font.Size = 11
        LblTerminalWindow.ForeColor = vbBlack
        LblTerminalWindow.BackStyle = 1      ' opaque
        FormTerminalVisible = True
        frmTerminal.Show
    Else
        frmTerminal.Hide
        FormTerminalVisible = False
        LblTerminalWindow.Font.Size = 9
        LblTerminalWindow.ForeColor = &H404040
        LblTerminalWindow.BackStyle = 0      ' transparent
    End If
End Sub

Private Sub ImageTerminal_Click()
    Call ToggleTerminalControls
End Sub

Private Sub LblTerminalWindow_Click()
    Call ToggleTerminalControls
End Sub

Sub TogglePOSmonitoringControls()
    If FormPOSmonitoringVisble = False And frmMonitoring.Visible = False Then
        FormPOSmonitoringVisble = True
        LblPOSmonitoring.ForeColor = &H8000&    ' dark green
        LblPOSmonitoring.Font.Size = 11         ' increase font size
        LblPOSmonitoring.BackStyle = 1          ' opaque
        frmMonitoring.Show
    Else
        frmMonitoring.Hide
        FormPOSmonitoringVisble = False
        LblPOSmonitoring.ForeColor = &H404040   ' dark gray
        LblPOSmonitoring.Font.Size = 9          ' decrease font size
        LblPOSmonitoring.BackStyle = 0          ' transparent
    End If
End Sub

Private Sub ImageMonitoring_Click()
    Call TogglePOSmonitoringControls
End Sub

Private Sub LblPOSmonitoring_Click()
    Call TogglePOSmonitoringControls
End Sub

Private Sub ImagePictogram_Click(Index As Integer)
    ActiveTabIndex = Index
    Call ShowDialogWindow
End Sub

Private Sub LabelTabSelect_Click(Index As Integer)
    ActiveTabIndex = Index
    Call ShowDialogWindow
End Sub


Private Sub lbl_ProxSensor_Click(Index As Integer)
    If cwBut_RT_Monitoring.Value = False Then Read_Prox_Sensors
End Sub


Private Sub optDialogWindow_Click(Index As Integer)
    ActiveTabIndex = Index
    Call ShowDialogWindow
End Sub

Sub en_dis_frames()
    If ConnectionType <> ConnDEMO Then   ' not demo - enable all screens
        'chkTerminalWindow.Enabled = True
        LblTerminalWindow.Enabled = True
        FrameSystem.Enabled = True
        cwButStartStopLTF.Enabled = True
        FrameAlarms.Enabled = True
        frmDiagnostic.Enabled = True
        frmFWupgrade.Enabled = True
    Else ' user or demo
        'chkTerminalWindow.Enabled = False
        LblTerminalWindow.Enabled = False
        FrameSystem.Enabled = False
        cwButStartStopLTF.Enabled = False
        FrameAlarms.Enabled = False
        frmDiagnostic.Enabled = False
        frmFWupgrade.Enabled = False
    End If
End Sub

Sub bringControlsToFront_True_toBack_False(to_front As Boolean)
' controls are put on the main window. when new tab is selected, they either brought up front or send back
    If (to_front) Then
        cwBNC0select.ZOrder (vbBringToFront)
        cwBNC1select.ZOrder (vbBringToFront)
        FrameVelocityFB.ZOrder (vbBringToFront)
        FramePneumFB.ZOrder (vbBringToFront)
        FramePositionFeedBack.ZOrder (vbBringToFront)
        FrameRTmon.ZOrder (vbBringToFront)
        FrameButtons.ZOrder (vbBringToFront)
        'cwBut_FF_ALL.ZOrder (vbBringToFront)
        'CmdRestorefromFlash.ZOrder (vbBringToFront)
        'cmdLoadDefaults(0).ZOrder (vbBringToFront)
        'cmdSaveParamsIntoFLASH(0).ZOrder (vbBringToFront)
    Else
        cwBNC0select.ZOrder (vbSendToBack)
        cwBNC1select.ZOrder (vbSendToBack)
        FrameVelocityFB.ZOrder (vbSendToBack)
        FramePneumFB.ZOrder (vbSendToBack)
        FramePositionFeedBack.ZOrder (vbSendToBack)
        FrameRTmon.ZOrder (vbSendToBack)
        FrameButtons.ZOrder (vbSendToBack)
        'cwBut_FF_ALL.ZOrder (vbSendToBack)
        'CmdRestorefromFlash.ZOrder (vbSendToBack)
        'cmdLoadDefaults(0).ZOrder (vbSendToBack)
        'cmdSaveParamsIntoFLASH(0).ZOrder (vbSendToBack)
    End If

End Sub

Sub ShowDialogWindow()
    Dim tmp As Long
    If ActiveTabIndex > 5 Then ActiveTabIndex = 5
    Call en_dis_frames
    For tmp = 0 To 5  'hide everything except current label used as colored rectangle
        LabelTabSelect(tmp).Visible = False
    Next
    LabelTabSelect(ActiveTabIndex).Visible = True

    If ActiveTabIndex = Connection_FrameIndex Then 'connection=0
      Call bringControlsToFront_True_toBack_False(False)
      FrmConnection.Visible = True
        FrameSystem.Visible = False
        FramePerformance.Visible = False
        frmDiagnostic.Visible = False
        frmFWupgrade.Visible = False
        FrameAlarms.Visible = False
    End If

    If ActiveTabIndex = SystemParams_FrameIndex Then 'sys data=1
      FrameSystem.Visible = True
      Call bringControlsToFront_True_toBack_False(True)
        FramePerformance.Visible = False
        frmDiagnostic.Visible = False
        frmFWupgrade.Visible = False
        FrameAlarms.Visible = False
        FrmConnection.Visible = False
        If (Test_in_progress <> True) Then
            Call GetSysData
        End If
    End If

    If ActiveTabIndex = Performance_FrameIndex Then  'performance=2
      Call bringControlsToFront_True_toBack_False(False)
        FrameSystem.Visible = False
      FramePerformance.Visible = True
        frmDiagnostic.Visible = False
        frmFWupgrade.Visible = False
        FrameAlarms.Visible = False
        FrmConnection.Visible = False
        If (Test_in_progress <> True) Then
            Call GetSysData
            Call Get_diag_params
        End If
    End If

    If ActiveTabIndex = Firmware_FrameIndex Then 'firmware=4
      Call bringControlsToFront_True_toBack_False(False)
        FrameSystem.Visible = False
        FramePerformance.Visible = False
        frmDiagnostic.Visible = False
      frmFWupgrade.Visible = True
        FrameAlarms.Visible = False
        FrmConnection.Visible = False
    End If

    If ActiveTabIndex = Alarms_FrameIndex Then 'alarms=5
      Call bringControlsToFront_True_toBack_False(False)
        FrameSystem.Visible = False
        FramePerformance.Visible = False
        frmDiagnostic.Visible = False
        frmFWupgrade.Visible = False
      FrameAlarms.Visible = True
        FrmConnection.Visible = False
    End If

    If ActiveTabIndex = Diagnostic_FrameIndex Then 'diagnostic=3
        Call GetDiagData
        FrameSystem.Visible = False
        FramePerformance.Visible = False
      frmDiagnostic.Visible = True
      Call bringControlsToFront_True_toBack_False(True)
        FrameAlarms.Visible = False
        frmFWupgrade.Visible = False
        FrmConnection.Visible = False
    End If
'    If ActiveTabIndex = Tuning_FrameIndex Then 'tuning=6 (moved to separate window)
'        FrameSystem.Visible = False
'        FramePerformance.Visible = False
'        frmDiagnostic.Visible = False
'        frmFWupgrade.Visible = False
'        FrameAlarms.Visible = False
'        FrmConnection.Visible = False
'    End If
End Sub

Private Sub optExitation_Click(Index As Integer)
'//   Set: freq=xxx.xx, Get: freq?, where xxx.xx - float
'    Dim controller_command As String
'    If Init_sys_data = False Then 'avoid firing control event subroutine during initialization when values assigned
'        TestFrequency = cwNumExcitFreq(Index).Value
'        If Index = 0 Then ' sine
'            controller_command = "freq=" & TestFrequency
'        Else ' noise
'            controller_command = "freq=0"
'        End If
'        cmd_sent = False
'        waitresponse = True
'        response = Analyzer.GetSend(controller_command, True)
'    End If
End Sub

Private Sub aTimerMonitoring_Timer()
    If cwBut_RT_Monitoring.Value = True Then
        If Menu_Active = True Then Call frmTerminal.QuitMenu
        Call ShowSysStatus      ' "s" command
        Call Read_Alarm_Status  ' "alarm" command
        Call read_RT_diagnostic ' "brif" command
    End If
End Sub

Private Sub aTimerTestProgram_Timer()
    test_timer_ready = True
End Sub

Private Sub timeTimer_Timer() ' interval is set to 0.1 sec
    in_time = Format(Time, "HH:MM:SS")
    lblSystemTime = in_time
    in_date = Format(Date, "MM/DD/YYYY")
    lblSystemDate = in_date
    If Test_in_progress Then lblTestTime.Caption = Format(Timer - StartTime, "#0.0")
    If Chk_AutoClear.Value = Checked And lstResponse.ListCount > 2000 Then
        lstResponse.Clear
        ListCommIssues.Clear
    End If
    If MotorSensorTest_active = True And MotorSensorTest_READY = True Then
        MotorSensorTest_READY = False
        Call Measure_MotSens_at_freq
    End If
End Sub

Sub ShowNoCommunication()
    MsgBox ("No communication over serial interface. Check connections")
End Sub

Sub Remove_echo(ByRef Buf As String, Send_Str As String)
On Error Resume Next
    Dim start_respond As Long
    start_respond = InStr(Buf, Send_Str)
    If (start_respond <> 0) And (echo_removed = False) Then
       Buf = Mid(Buf, start_respond + Len(Send_Str) + 1)
       echo_removed = True
    End If
End Sub

Public Function Wait_sec(seconds As Single)
    Dim CmdStartTime As Single
    CmdStartTime = Timer
    Do 'delay
        DoEvents
    Loop Until (Timer - CmdStartTime) > seconds
End Function

Public Function GetSend(message As String, Wait As Boolean) As String
'    On Error GoTo port_problems
    Dim CmdStartTime As Single
    Dim WaitTime As Single
    Dim Cmd_String As String ' local copy in order not to collide from different calls
    Dim CR_pos As Long
    Dim pos_OK As Long, pos_response As Long
    Dim got_response As Long
    Dim temp_str As String
    Dim TMP_BUF As String * 10000  ' zero-filled, defined length (not resizable) for DLL.
    Dim LocalBUF As String    ' local buffer to work with; Buffer can be updated any time, this is time-snip copy
    Dim totalTime As Single
    ' VB6 bug? Buffer is limited in length. even if Len(Buffer)= 320, the watch window shows ~256 chars,
    ' and when copied to output string only first 250 chars are copied?
    ' TMP_BUF = Space(10000)

    '!!!! >>>> Replace(TMP_BUF, Chr(0), " ") TAKES ~0.6 sec if DEBUG UNDER Visual Studio 2010 !!!!
    'TMP_BUF = Replace(TMP_BUF, Chr(0), " ") ' replace  0 with space.

    CmdStartTime = Timer
    GetSend = ""
' for some locals, replace ',' with '.'
    message = Replace(message, ",", ".")
    message = Replace(Trim(message), vbTab, "")   ' IK 20220129 remove TABs
    message = LTrim(message)   ' IK20230117 Replace creates "oltfstop" instead of "oltf stop", bad! IK 20220130 remove spaces Replace(Trim(message), " ", "")

    txtManualCMD.Text = message ' diagnostic
    Cmd_String = message & vbCr ' Lf
    SendStrLen = Len(Cmd_String)
    pos_response = 1 ' or InStr(pos_response, Buffer, vbLf) gives error

    If ConnectionType = ConnSimDLL Then
'        Call Is_DLL_available ' it sets "DLL_is_available" which is global boolean
        If DLL_is_available Then
            If (DLL_is_Initialized = False) Then Init_DLL
            Buffer = ""
            CmdStartTime = Timer
            TMP_BUF = Send_DLL_cmd(Cmd_String) 'fills TMP_BUF with response
            GoTo get_response
        Else
            ConnectionType = ConnDEMO
            GetSend = "DEMO MODE"
            GoTo finish_function
        End If
    End If
send_cmd: ' program is here if serial port was available, and ConnType <> ConnSimDLL
    Call open_port_if_closed 'if port not avilable, it will set ConnType = ConnSimDLL or ConnDEMO
    If (ConnectionType = ConnSimDLL Or ConnectionType = ConnDEMO) Then
        If (ConnectionType = ConnCOMport) Then 'COM port
            GetSend = "Port not avail"
        End If
        If (ConnectionType = ConnTelnet) Then  'winsock
            GetSend = "Telnet not avail"
        End If
        If (ConnectionType = ConnDEMO) Then  'demo
            GetSend = "DEMO MODE"
        End If
        cwCOMdisconnect.Caption = Demo_Simul_string
        GoTo finish_function 'port cannot be open
    End If
    Call UpdateConnectControls(True)
    txtManualCMD.Text = message ' diagnostic
' Send the  command
    Call Clean_COM_Buffer ''clean the Comm port buffer
    txtStatus.Text = Cmd_String ' to debug
'   txtStatus.Refresh 'this causes flickering of the screen
'    CmdStartTime = Timer
    Call Send_GUI_command(Cmd_String) ' >>>>>>>>>
    pos_response = 1
    got_response = 0
    If Wait Then
         ' Wait for data to come back to the serial port.
      Do
        DoEvents
'        If Len(Cmd_String) < 4 Then 'basic command
'          Do 'small delay
'            DoEvents '  here another call of GetSend might happen
''this does frmTerminal.Receive if GUI_doing_command if true
'            WaitTime = Timer - CmdStartTime
'          Loop Until (WaitTime > 0.01)
get_response:
          If ConnectionType = ConnSimDLL Then
              pos_response = InStr(TMP_BUF, ">~")
              If pos_response > 0 Then
                LocalBUF = Left(TMP_BUF, pos_response + 3) ' DLL writes to fixed size buffer TMP_BUF
              End If
         'TMP_BUF = Buffer ' local copy
              If LocalBUF = "" Then
                GetSend = "Invalid answer"
                GoTo finish_function
              End If
         Else
            LocalBUF = Buffer  ' IK20201211 still produces all spaces array, does not work
         End If
'        ' TMP_BUF = Buffer ' local copy
'         pos_OK = InStr(LocalBUF, Cmd_String) 'too short command 'e' or 'f' getting response like "e" & vbCrLf & "Feedforward ENABLED...: 0f" & vbCrLf
''         If (pos_OK < Len(Cmd_String) + 2) And (pos_OK > 0) Then ' If (pos_OK <>0) and it finds first command itself, then response is cut and leaves nothing
'         If (pos_OK = 1) Then ' pos_OK =1 when controller sends echo
'            LocalBUF = Mid(LocalBUF, pos_OK + Len(Cmd_String)) 'cut echo = beginning of LocalBUF
'            ' cut LfCr
'            If Left(LocalBUF, 1) = Chr(13) Or Left(LocalBUF, 1) = Chr(10) Then
'                LocalBUF = Right(LocalBUF, Len(LocalBUF) - 1)
'            End If
'            If Left(LocalBUF, 1) = Chr(13) Or Left(LocalBUF, 1) = Chr(10) Then
'                LocalBUF = Right(LocalBUF, Len(LocalBUF) - 1)
'            End If
'          End If
'        Else ' commands longer than 2 bytes+CrLf require return >~OK or >~ERR
          LocalBUF = Buffer ' local copy
          pos_OK = InStr(LocalBUF, Acknowledgment_token) ' ">~"
          If (pos_OK <> 0) Then
            got_response = pos_OK
            'next, wait for final CrLf
            pos_OK = InStr(got_response, LocalBUF, vbCrLf) '13,10;\r\n
            If (pos_OK <> 0) Then
                pos_response = pos_OK
                got_response = pos_OK
            Else
                pos_response = 1
                got_response = 0
            End If
          Else 'Timeout
            pos_OK = InStr(LocalBUF, "Timeout")
            If pos_OK Then
                GetSend = Trim(LocalBUF)
                GoTo finish_function
            End If
          End If
'        End If
         WaitTime = Timer - CmdStartTime
' 1-letter command is too short and miss end of response. LOOK for ":" in response
' actual response: ASC(10) & "e" & ASC(13) & ASC(10) & "Feedforward ENABLED...". IT MUST BE IN ADDITION ":8e"
      Loop Until (((got_response <> 0) And (InStr(pos_response, Buffer, vbLf)) > got_response) Or WaitTime > DefaultTimeOut)
      lblCMDtime.Caption = Format(WaitTime, "#0.####")
      If WaitTime > DefaultTimeOut Then
            Buffer = "Timeout"
            GetSend = Buffer
            GoTo finish_function
      End If
    Else ' just show response
      Do
        DoEvents
'        call Get_Serial_Input
        LocalBUF = Buffer
         WaitTime = Timer - CmdStartTime
      Loop Until Len(LocalBUF) >= SendStrLen Or WaitTime > 0.2 'safety delay 0.2 sec, if there is no transmission it can stack here
'--        lblRawResponse.Caption = Buffer
    End If
'parse response
    pos_OK = InStr(LocalBUF, Acknowledgment_token) ' ">~"
    If pos_OK <> 0 Then
        Call Remove_echo(LocalBUF, Cmd_String)
        If (echo_removed = True) Then echo_removed = False ' prepare for the next command
        pos_OK = InStr(LocalBUF, vbCrLf & Acknowledgment_token & "OK" & vbCr) ' "\r\n>~OK\r"
        If pos_OK <> 0 Then 'last part of the response, like freq?\r\nfreq=5.0\r\n>~OK\r\n
            LocalBUF = Left(LocalBUF, pos_OK - 1)
        Else 'first part of response: s\r\ns\r\n>~OK The System Status is: 9e\r\n
            got_response = InStr(pos_response + 5, LocalBUF, vbCrLf) ' "\r\n>~OK\r" - search Cr Lf after >~OK
            If got_response <> 0 Then pos_response = got_response
            If pos_response <> 0 Then
                temp_str = Left(LocalBUF, pos_response - 1) ' cut either beginning or tail of response
            Else
                temp_str = (LocalBUF) '  response
            End If
            LocalBUF = temp_str
        End If
    Else
        ' 20201211 this might cause infinite loop
        'GoTo get_response
    End If
    LocalBUF = Trim(LocalBUF) ' if size is defined (As String * 10000), trim does not reduce size!
    If Right(LocalBUF, 2) = vbCrLf Then
        LocalBUF = Left(LocalBUF, Len(LocalBUF) - 2)
    End If
    If Left(LocalBUF, 1) = Chr(27) Then 'ESC
        LocalBUF = Right(LocalBUF, Len(LocalBUF) - 1)
    End If

    If Left(LocalBUF, 2) = vbCrLf Then
        LocalBUF = Right(LocalBUF, Len(LocalBUF) - 2)
    End If
    If Right(LocalBUF, 2) = vbCrLf Then
        LocalBUF = Left(LocalBUF, Len(LocalBUF) - 2)
    End If
    If Left(LocalBUF, 1) = vbCr Then
        LocalBUF = Right(LocalBUF, Len(LocalBUF) - 1)
    End If
    If Left(LocalBUF, 1) = vbCr Then
        LocalBUF = Right(LocalBUF, Len(LocalBUF) - 1)
    End If
    If Left(LocalBUF, 1) = vbLf Then
        LocalBUF = Right(LocalBUF, Len(LocalBUF) - 1)
    End If
    If Left(LocalBUF, 1) = vbLf Then
        LocalBUF = Right(LocalBUF, Len(LocalBUF) - 1)
    End If
    If Left(LocalBUF, 1) = vbCr Then
        LocalBUF = Right(LocalBUF, Len(LocalBUF) - 1)
    End If

    lstResponse.AddItem ("-->" & message) ' add command to the list
    Cmd_String = Trim(LocalBUF) ' <<<<<<<<<<<<  this is the return of the function, trim ending spaces when copied
    GetSend = Cmd_String
    If (message = "oltf stop") Then
        GetSend = "*" ' because "oltf stop" only returns ">~OK" it gets trimmed, so function return is "" empty string, needs * to confirm stop
        Exit Function
    End If
    Do  'update list
      DoEvents
        If Cmd_String = "Timeout" Then
            GoTo finish_function
        End If
        CR_pos = InStr(Cmd_String, vbLf)
        If CR_pos = 1 Then ' there is a LF in the response on first position
            Cmd_String = Mid(Cmd_String, CR_pos + 1)
        End If
        CR_pos = InStr(Cmd_String, vbCr)
        If CR_pos = 1 Then ' there is a LF in the response on first position
            Cmd_String = Mid(Cmd_String, CR_pos + 1)
        End If
        CR_pos = InStr(Cmd_String, vbCr)
        If CR_pos = 0 Then  ' there is no CR in the response
          CR_pos = InStr(Cmd_String, vbLf) ' try LF in the response
        End If
        If CR_pos Then ' there is a LF in the response
            lstResponse.AddItem (Mid(Cmd_String, 1, CR_pos - 1))
            Cmd_String = Mid(Cmd_String, CR_pos + 1)
            CR_pos = InStr(Cmd_String, vbLf)
            If CR_pos = 1 Then ' there is a LF in the response on first position
                Cmd_String = Mid(Cmd_String, CR_pos + 1)
            End If
        Else ' no CR or LF left in string, add time taken
'            totalTime = Timer - CmdStartTime
            temp_str = Format(WaitTime * 1000, "0.0## ms")
'            lstResponse.AddItem (RTrim(Cmd_String) & " // " & temp_str & " TotalT" & Format(totalTime * 1000, " 0.000 ms"))
            lstResponse.AddItem (RTrim(Cmd_String) & " // " & temp_str)
            Cmd_String = LCase(Cmd_String) ' because checking ">~u")
            If (((WaitTime * 1000) > 80) Or (InStr(Cmd_String, ">~u"))) Then
                ListCommIssues.AddItem ("cmd=" & message & " elapsed=" & temp_str & "; rsp=" & (GetSend))
            End If
            GoTo finish_function
        End If
   Loop Until (Len(Cmd_String) < 1) ' end adding to communication window
finish_function:
    GUI_doing_command = False ' release com port for next request
    Exit Function

port_problems:
    Call portproblems
End Function










'============================================================================================================================
'''''''''''''''''''''''''''''''''''''' CONNECTION   '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'============================================================================================================================

'================= PORT SETUP =========================================================================================

Private Sub cmdCommSETUP_Click()
'    frmsplash.cmdHide.Visible = True
    frmSplash.Show
End Sub

Sub UpdateConnectControls(Value As Boolean)
    cwCOMdisconnect.Value = Value
    If Value = True Then 'open
        ImageConnected.Visible = True
        ImageDisconnected.Visible = False
        LblCONNparams = cwCOMdisconnect.Caption & " open "
    Else 'close
        ImageDisconnected.Visible = True
        ImageConnected.Visible = False
        LblCONNparams = cwCOMdisconnect.Caption & " closed"
    End If

End Sub

Private Sub cwCOMdisconnect_ValueChanged(ByVal Value As Boolean) ' works on click?
    On Error GoTo port_problems
    If Init_sys_data = True Then GoTo change_pic

    If cwCOMdisconnect.Value = True Then 'open
        Call open_port_if_closed
        Call frmTerminal.Update_Form_Caption
    Else 'close
        Call Close_Port
    End If
change_pic:
    If cwCOMdisconnect.Value = True Then 'open
        ImageConnected.Visible = True
        ImageDisconnected.Visible = False
        LblCONNparams = cwCOMdisconnect.Caption & " open "
    Else 'close
        ImageDisconnected.Visible = True
        ImageConnected.Visible = False
        LblCONNparams = cwCOMdisconnect.Caption & " closed"
    End If
    Exit Sub
port_problems:
  Call portproblems      ' port not available
End Sub

Sub portproblems(Optional prog_called As String = "")
    Dim Msg, Style, Title, answer
    GUI_doing_command = False ' release com port for next request

    If prog_called <> "" Then
        Title = prog_called   ' Define title.
    Else 'no argument string
        Title = Function_requested_TX   ' Define title.
    End If

    Msg = "PC port: COM" & frmSplash.SerialComm.CommPort & " is not available, probably used by other application."
'    FrmDebug.Add_Debug_Message_to_File ("caller: " & Title & " " & Msg)

    Title = Title & " COM port is not available"   ' Define title.
    Msg = Msg & vbCrLf & "Retry again / try other port? 'Cancel' will switch to DEMO mode"
    Style = vbRetryCancel + vbDefaultButton1 + vbExclamation ' Define buttons.
        answer = MsgBox(Msg, Style, Title)
        If answer = vbRetry Then   ' User chose Retry.
            FrmConnection.Caption = ""
            Exit Sub   ' Perform some action.
        Else   ' User chose No.
            If (DLL_is_available And DLL_is_Initialized = False) Then
                Init_DLL 'DLL checks if it is already initialized and ignores second attempt
                ConnectionType = ConnSimDLL
            Else
                ConnectionType = ConnDEMO ' port not available
            End If
            Call UpdateConnectControls(False)
            FrmConnection.Caption = "DEMO MODE"
        End If
End Sub

'========================  E N D  O F  ======================================================================================
'''''''''''''''''''''''''' CONNECTION and PORT SETUP ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''



'============================================================================================================================
'''''''''''''''''''''''''''''''''''''' FORM SYSTEM CONTROL  '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'============================================================================================================================

'=====================================  E N D  O F  =========================================================================
'''''''''''''''''''''''''''''''''''''' FORM SYSTEM CONTROL  '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''






'============================================================================================================================
'''''''''''''''''''''''''''''''''''''' PERFORMANCE  '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'============================================================================================================================

'================= TEST MOTOR_SENSOR PAIR =========================================================================================

Private Sub ImageTuning_Click()
    frmFilters.Show
End Sub

Private Sub LblTuning_Click()
    frmFilters.Show
End Sub

Private Sub CmdFilters_Click(Index As Integer)
    frmFilters.Show
End Sub

Private Sub cwButStartStopMotSenTest_ValueChanged(ByVal Value As Boolean)
    Dim decoding_str As String
    If cwButStartStopMotSenTest.Value = True Then
        If LTFtestOutputName = "" Or LTFtestInputName = "" Then
            response = MsgBox("Please choose Output and Input first", vbOKOnly + vbExclamation, "Parameters Not Defined")
            cwButStartStopMotSenTest.Value = False
            MotorSensorTest_active = False
            Exit Sub
        End If
        decoding_str = "echo"
        response = GetSend(decoding_str, True) 'update freq
        If InStr(LCase(response), "verb") Then
            response = GetSend("echo>enable", True) 'disable VERBOSE echo
        End If
        decoding_str = "freq=" & cwNumExcitFreq(0).Value
        response = GetSend(decoding_str, True) 'update freq
        decoding_str = "ampl=" & cwNumExcitAmpl(0).Value
        response = GetSend(decoding_str, True) 'update test amplitude
    Else
    'If cwButStartStopMotSenTest.Value = False Then
        MotorSensorTest_active = False
        MotorSensorTest_READY = False 'allow to start in Timer
        Exit Sub
    End If
    MotorSensorTest_active = True
    MotorSensorTest_READY = True 'allow to start in Timer
End Sub

Private Sub cw_Status_Bit_Indicator_ValueChanged(Index As Integer, ByVal Value As Boolean)
'    If Index = 1 Then showSysStatus ' click on go-NoGo
End Sub

Private Sub cw_Status_Bit_Indicator_Click(Index As Integer)
    'If Index = 0 Or Index = 1 Or Index = 2 Or Index = 4 Or Index = 6 Then
    Call ShowSysStatus ' click on go-NoGo
    'End If
End Sub

Private Sub cwNumExcitAmpl_ValueChanged(Index As Integer, Value As Variant, PreviousValue As Variant, ByVal OutOfRange As Boolean)
'//   Set: ampl=xxx.xx, Get: ampl?, where xxx.xx - float
    Dim controller_command As String
    If Init_sys_data = False Then 'avoid firing control event subroutine during initialization when values assigned
        TestGain = cwNumExcitAmpl(Index).Value
        controller_command = "ampl=" & TestGain
        cmd_sent = False
        waitresponse = True
        response = Analyzer.GetSend(controller_command, True)
    End If
End Sub

Private Sub cwNumExcitFreq_ValueChanged(Index As Integer, Value As Variant, PreviousValue As Variant, ByVal OutOfRange As Boolean)
'//   Set: freq=xxx.xx, Get: freq?, where xxx.xx - float
    Dim controller_command As String
    If Init_sys_data = True Then Exit Sub 'avoid firing control event subroutine during initialization when values assigned
    TestFrequency = cwNumExcitFreq(Index).Value
    controller_command = "freq=" & TestFrequency
    cmd_sent = False
    waitresponse = True
    response = Analyzer.GetSend(controller_command, True)

    Init_sys_data = True
    ' there are two controls on Analyzer form: one on diagnostic, another on performace
    If (Value <= 10) Then cwNumExcitFreq(Index).IncDecValue = 0.1
    If (Value >= 10.1) Then
        Value = 11
        TestFrequency = Value
        cwNumExcitFreq(Index).IncDecValue = 1
    End If
    ' update increment of other control
    cwNumExcitFreq(1 - Index).IncDecValue = cwNumExcitFreq(Index).IncDecValue
    If (Index = 0) Then
        cwNumExcitFreq(1).Value = TestFrequency
    Else
        cwNumExcitFreq(0).Value = TestFrequency
    End If
    'update control on monitoring form
    If IsFormLoaded("frmMonitoring") Then
        frmMonitoring.cwNumExcitFreq.Value = TestFrequency
        frmMonitoring.cwNumExcitFreq.IncDecValue = cwNumExcitFreq(Index).IncDecValue
    End If
    Init_sys_data = False
End Sub

Private Sub cmdPulse_Click()
    Dim controller_command As String
    If p_cmd_as_PN = False Then
        controller_command = "p"
    Else
        controller_command = "puls"
    End If
    cmd_sent = False
    waitresponse = True

    response = Analyzer.GetSend(controller_command, True)
    If InStr(response, "95") Then 'Bruker Insight AFP 95-43009-01r.At3dbg @26-Oct-16;compiled Oct 28 2016 13:48:19
        p_cmd_as_PN = True
    'Else
    '    p_cmd_as_PN = False
    End If
    Init_sys_data = True
    Excitation_status = False  'after pulse excitation stops
    Analyzer.cwBut_Excitation.Value = Excitation_status
    If IsFormLoaded("frmMonitoring") Then frmMonitoring.cwBut_Excitation.Value = Excitation_status
'    frmStageFFgains.cmdPulse.Value = Excitation_status
    Init_sys_data = False
End Sub

Private Sub cmdChooseExcitation_Click()
    frmDrivingOutput.Show
End Sub

Private Sub CmdChooseInputSignal_Click()
    SecondInputSelection = False
    frmSensorInput.Show
End Sub

Private Sub CmdChoose2ndInputSignal_Click()
    SecondInputSelection = True
    frmSensorInput.Show
End Sub

Sub Measure_MotSens_at_freq()
    Dim decoding_str As String
    Dim MotSen_testData As String
    Dim pos_separator As Long
    Dim measVal As Double

    cwBut_RT_Monitoring.Value = False ' cancel RT monitoring
    Call create_LTF_command
    waitresponse = True
    cmd_sent = False
    '    CmdStartTime = Timer
    '    Do
    '         WaitTime = Timer - CmdStartTime

    '    clean buffer
    Call Clean_COM_Buffer
    If Len(Buffer) Then
        Buffer = ""
    End If

   ' Send the  command
   If LTF_command <> "" And cmd_sent = False Then
        echo_removed = False
        SendString = LTF_command
        lstResponse.AddItem ("=>" & LTF_command)
        SendString = SendString + vbCrLf
        SendStrLen = Len(SendString)
        Call open_port_if_closed
        Call Send_GUI_command(SendString) 'it will add to history
        cmd_sent = True
    Else
        SendStrLen = 0
    End If
    ' Set up for response if query
    Call Get_Serial_Input
    MotSen_testData = GetDataString(">~ER", ">~OK" & vbCrLf, (4 / cwNumExcitFreq(0).Value + 3)) 'wait time: = 4 / cwNumExcitFreq.Value + 3
    pos_separator = InStr(MotSen_testData, ">~OK")
'    If (pos_separator) Then
'    MotSen_testData = Mid(MotSen_testData, 1, Len(">~Doing CMD") + 2) ' cut doing command
'    End If
    If (pos_separator And (Len(MotSen_testData) > pos_separator + 2)) Then ' can be '>~Doing CMD\r\n', len=11
        MotSen_testData = Left(MotSen_testData, pos_separator - 3)
    End If
    lstResponse.AddItem (MotSen_testData)
    'return -mm.mm, -deg.deg \r\n >~OK
'mtst ma_x pa_x

'-22.21, -163.7
'>~OK
'echo>verb

'>~OK
'mtst ma_x pa_x
'wait 2.80sec......
'Ratio: -22.21dB[0.078x], -163.7deg; @ freg=5.0Hz, ampl= 1000
'>~OK
    pos_separator = InStr(MotSen_testData, ",")
    If pos_separator = 0 Then
        txtStatus = ("Motor-Sensor decoding error")
        MotorSensorTest_READY = True
        Exit Sub
    End If
    If IsNumber(Left(MotSen_testData, pos_separator)) Then
        decoding_str = Left(MotSen_testData, pos_separator - 1) ' separate magnitude
    Else
        If (pos_separator > 7) Then
            decoding_str = Mid(MotSen_testData, pos_separator - 7, 7) ' separate magnitude
        End If
    End If

    If IsNumber(decoding_str) Then
        txtMag.Text = Val(decoding_str)
    End If
    decoding_str = Mid(MotSen_testData, pos_separator + 1) ' separate phase
    If IsNumber(decoding_str) Then
        txtPhase.Text = Val(decoding_str)
    End If
    MotorSensorTest_READY = True
End Sub
'====  E N D  O F  ==== TEST MOTOR_SENSOR PAIR ============================================================================





'================= OLTF and P_TF  =========================================================================================

Private Sub cwNumMaxFreq_ValueChanged(Value As Variant, PreviousValue As Variant, ByVal OutOfRange As Boolean)
    If (PreviousValue = 1000) And (Value = 1500) Then
        Value = 2000
        cwNumMaxFreq.IncDecValue = 1000
    End If
    If (PreviousValue = 1000) And (Value <= 100) Then 'from above 1000 Hz down
        Value = 500
        cwNumMaxFreq.IncDecValue = 100
    End If

    If (Value = 500) Then cwNumMaxFreq.IncDecValue = 100
    If (Value = 600) Then
        Value = 1000
        cwNumMaxFreq.IncDecValue = 1000
    End If
    CWGraph.Axes(1).Maximum = Value
    CWGraphPhase.Axes(1).Maximum = Value
End Sub

Private Sub cwMax_dB_ValueChanged(Value As Variant, PreviousValue As Variant, ByVal OutOfRange As Boolean)
    CWGraph.Axes.Item(2).Maximum = cwMax_dB.Value
End Sub

Private Sub cwMin_dB_ValueChanged(Value As Variant, PreviousValue As Variant, ByVal OutOfRange As Boolean)
    CWGraph.Axes.Item(2).Minimum = cwMin_dB.Value
End Sub

Private Sub ChkReversePhase_Click()
    If ChkReversePhase.Value = Checked Then
        PhaseReversed = True
    Else
        PhaseReversed = False
    End If
    Call Update_prev_plots
    Call update_plot(line_counter)
    'Call Calc_prediction
End Sub

Private Sub optTestInput_Click(Index As Integer)
    Call create_LTF_command
End Sub

Private Sub CmdReadScan_Click()
    Call ReadMeasurementPlot(0) ' zero means standard read; replace: current->previous; this->current
End Sub

Private Sub SetDefaultPrinter(ByVal PrinterName As String, _
    ByVal DriverName As String, ByVal PrinterPort As String)
    Dim DeviceLine As String
    Dim r As Long
    Dim l As Long
    DeviceLine = PrinterName & "," & DriverName & "," & PrinterPort
    ' Store the new printer information in the [WINDOWS] section of
    ' the WIN.INI file for the DEVICE= item
    r = WriteProfileString("windows", "Device", DeviceLine)
    ' Cause all applications to reload the INI file:
    l = SendMessage(HWND_BROADCAST, WM_WININICHANGE, 0, "windows")
End Sub

Private Sub cmdPrint_Click()
    On Error GoTo port_problems

    Dim PrinterName As String

    Set frmPrint.picGain.Picture = CWGraph.ControlImage
    Set frmPrint.picPhase.Picture = CWGraphPhase.ControlImage
'    Set frmPrint.FrameInfo = frameAxis
'    Set frmPrint.PicInfo.Picture = frmMonitoring NOT WORKING
'    Set frmPrint.PicInfo.Picture = frameAxis NOT WORKING
    frmPrint.lblAxisN.Caption = IsolatorNumber 'cwIsolatorNumber.Value
'    frmPrint.lblAxisOrt.Caption = IsolatorAxis
    frmPrint.lblDate.Caption = lblSystemDate.Caption
    frmPrint.lblTime.Caption = lblSystemTime.Caption
'    frmPrint.optExcitOutput(0) = optExcitOutput(0)
'    frmPrint.optExcitOutput(1) = optExcitOutput(1)
'    frmPrint.optAxis(0) = optAxis(0)
'    frmPrint.optAxis(1) = optAxis(1)
'    frmPrint.optAxis(2) = optAxis(2)
'    frmPrint.optAxis(3) = optAxis(3)
'    frmPrint.optAxis(4) = optAxis(4)
'    frmPrint.optAxis(5) = optAxis(5)
'    If (optTestInput(0).Value = True) Then
'        frmPrint.LblTestType.Caption = " " + optTestInput(0).Caption
'    ElseIf (optTestInput(1).Value = True) Then
'        frmPrint.LblTestType.Caption = " " + optTestInput(1).Caption
'    ElseIf (optTestInput(2).Value = True) Then
'        frmPrint.LblTestType.Caption = " " + optTestInput(2).Caption
'    End If
    frmPrint.txtFW_Version = txtFW_Version
    frmPrint.LblCOMport = strPortNum
    frmPrint.Lbl_BAUD = strPortSetting
    dlgCommonDialog.CancelError = True
    dlgCommonDialog.DialogTitle = "Choose printer"
    dlgCommonDialog.Action = 5
    PrinterName = Printer.DeviceName ' save default system printer name
    Printer.TrackDefault = True
    Printer.Orientation = vbPRORLandscape

    If Err = 32755 Then 'Cancel is selected
        Exit Sub
    Else
        frmPrint.PrintForm  ' print Form TO THE DEFAULT SYSTEM PRINTER
    End If

    Call WinNTSetDefaultPrinter(PrinterName)
    Exit Sub

port_problems:
    Call MsgBox("Error '" & Error & "' " & vbCrLf & "Check Printer connection & Driver", vbCritical, "General Operation System Error")
End Sub

Public Function GetLTFdata(message As String, Optional Wait_Time As Single = -1) As String
    On Error GoTo port_problems
    Dim CmdStartTime As Single
    Dim cmdWaitTime As Single
    Dim WaitingTime As Single

    Dim start_respond As Long
    Dim pos_response As Long

    If (Wait_Time = -1) Then ' argument is missing
        WaitingTime = DefaultLTFtimeOut '2 sec
    Else
        WaitingTime = Wait_Time
    End If
    If WaitingTime = 0 Then WaitingTime = DefaultLTFtimeOut '2 sec
send_cmd:

   ' Send the  command
   If message <> "" And cmd_sent = False Then
        echo_removed = False
        SendString = message
        lstResponse.AddItem ("=>" & message)
        SendString = SendString + vbCrLf
        SendStrLen = Len(SendString)
        Call open_port_if_closed
        Call Send_GUI_command(SendString)
        'show last command sent
        txtManualCMD.Text = message

        cmd_sent = True
   Else
        SendStrLen = 0
   End If
    ' Set up for response if query
   If Len(Buffer) Then
   End If
    GetLTFdata = GetDataString(end_of_TFtransmission, end_of_TFstring_token, Wait_Time)

get_response:
    If (waitresponse = False) Then
            GetLTFdata = "UserStop"
    End If

    If Right(GetLTFdata, 2) = vbCrLf Then
        GetLTFdata = Left(GetLTFdata, Len(GetLTFdata) - 2)
    End If
    If Left(GetLTFdata, 2) = vbCrLf Then
        GetLTFdata = Mid(GetLTFdata, 3)
    End If
'    lstResponse.AddItem GetLTFdata
    Exit Function

port_problems:
    Call portproblems
End Function

Private Sub ChkAdvTF_Click()
    If (ChkAdvTF.Value = Checked) Then 'Advanced TF test
        LTFadvanced = True
        ' S E N S O R S
        CmdChooseInputSignal.Width = 1100
        CmdChoose2ndInputSignal.Visible = True
'        frmSensorInput.optSensorInput(23).Visible = True
        ' D R I V I N G   O U T P U T
        frmDrivingOutput.optExcitOutput(25).Visible = True
        cmdChooseExcitation.Width = 1100
        FrameLTFaveraging.Visible = True
    Else
        LTFadvanced = False
        ' S E N S O R S
        CmdChooseInputSignal.Width = 2265
        CmdChoose2ndInputSignal.Visible = False
'        frmSensorInput.optSensorInput(23).Visible = False
        ' D R I V I N G   O U T P U T
        frmDrivingOutput.optExcitOutput(25).Visible = False
        cmdChooseExcitation.Width = 2265
        FrameLTFaveraging.Visible = False
    End If
End Sub

Function create_LTF_command() As String
    Dim TestCMD As String
    If optTestInput(0).Value = True Then 'mot-sen test
        TestCMD = optTestInput(0).Tag
        FrameLTF.Visible = False
        FrameMeasureAtFreq.Visible = True
    End If
    If optTestInput(1).Value = True Then 'OLTF test
        TestCMD = optTestInput(1).Tag
        FrameMeasureAtFreq.Visible = False
        FrameLTF.Visible = True
    End If
    If optTestInput(2).Value = True Then 'plant TF test
        TestCMD = optTestInput(2).Tag
        FrameMeasureAtFreq.Visible = False
        FrameLTF.Visible = True
    End If

    LTF_command = TestCMD 'IsolatorNumber & IsolatorAxis 'CStr(cwIsolatorNumber.Value) & IsolatorAxis
    LTF_command = LTF_command & LTFtestOutputName & " " & LTFtestInputName

    If (LTFadvanced = True) Then
        LTF_command = LTF_command & " " & LTFtest2ndInputName
    End If

    If (cwLTF_time_scale.Value <> 1) Then
        LTF_command = LTF_command & Str(cwLTF_time_scale.Value)  ' IK 20220304 fixed inability to increase average on German computers: "," decimal separator ignored
    End If
    If (ChkUseAsStartTFfreq.Value = Checked) Then ' IK 20210414
        Dim startFreq As Long
        startFreq = cwNumMaxFreq.Value
        If startFreq > 1000 Then ' IK 20210414 Firmware accepts up to 4050 Hz, plot max is 5000
            startFreq = 1000
        End If
        If startFreq >= 500 Then
            startFreq = startFreq + 50
        ElseIf startFreq >= 200 Then
            startFreq = startFreq + 25
        ElseIf startFreq >= 100 Then
            startFreq = startFreq + 10
        Else:
            startFreq = startFreq
        End If
        LTF_command = LTF_command & " " & startFreq
    End If

    Call Check_Selected_Actuator_Sensor
    txtManualCMD.Text = LTF_command
End Function

Private Sub cwButStartStopLTF_ValueChanged(ByVal Value As Boolean)
    Dim controller_command As String
    Call Analyzer.GetSend("zero", False)  ' IK20251217 added: zero geophone offset command
    If LTFtestOutputName = "" Or LTFtestInputName = "" Then
        If cwButStartStopLTF.Value = False Then
            Exit Sub
        Else
            response = MsgBox("Please choose Output and Input first", vbOKOnly + vbExclamation, "Parameters Not Defined")
            cwButStartStopLTF.Value = False
        Exit Sub
        End If
    End If
    If Value = True Then ' start
        FormTerminalVisible = False
'        frmTerminal.Hide
        Call create_LTF_command
        controller_command = txtManualCMD.Text
        If InStr(txtManualCMD.Text, "mtst") Then
            MsgBox ("please choose ether 'Frequency Sweep Plant TF' or 'Open Loop TF'")
            Value = False
            Exit Sub
        End If

        TestFrequency = RectifyValueResponse("freq?") '"+5.000 // excitation frequency" <<<<< with verbose response

        Call Update_prev_plots
        Call clear_arrays
'        response = GetSend("run>ltf", True)

        waitresponse = True
    '    CmdStartTime = Timer
    '    Do
    '         WaitTime = Timer - CmdStartTime

    '    Loop Until (WaitTime > 0.2) '20 ms delay
        saved_freq = TestFrequency ' save current frequency because it will finish as 0.1 Hz

        response = DoTest(controller_command) ' program pointer is inside function until test finished
        ' IK200407 DoTest returns negative (user cancel) or number of frequencies measured. for ElDamp 141
        If (response < 0) Then 'stop run LTF by controller
            response = GetSend("oltf>stop", True)
        ElseIf (response >= 1) Then ' LTF finished and sent '*'
            Test_in_progress = False
        End If
    Else 'value is to STOP
        If InStr(txtManualCMD.Text, "mtst") Then
            Exit Sub
        End If
        waitresponse = False
        cmd_sent = False
        If Test_in_progress = True Then
            response = GetSend("oltf>stop", True) ' send stop, wait for response
            Wait_sec (0.3)
        End If
        response = response & Buffer

        If (((InStr(response, "*") <> 0) And (Test_in_progress = True)) Or (InStr(response, "Unknown Command") <> 0)) Or (InStr(response, ">~OK") <> 0) Then '
            cwSldProgress.Caption = "! User Stop !"
            Test_in_progress = False
        Else
            If Test_in_progress = True Then MsgBox ("command 'stop' to controller failed, requests for data stopped")
        End If
    End If
    Test_in_progress = False
    response = GetSend("freq=" & saved_freq, True) ' restore saved frequecy, wait for response
End Sub

Private Sub cmdClearTerminal_Click()
    lstResponse.Clear
    ListCommIssues.Clear
End Sub

Function CoherenceIsPresent() As Boolean ' looks if array has not zeroes
    Dim test_sum As Double
    Dim indx As Long
    test_sum = 0
    For indx = 0 To UBound(Coherence_data)
        test_sum = test_sum + Coherence_data(indx)
    Next
    If test_sum <> 0 Then
        CoherenceIsPresent = True
    Else
        CoherenceIsPresent = False
    End If
End Function

' IK200407 DoTest returns negative (user cancel) or number of frequencies measured. for ElDamp 141
Public Function DoTest(ScanCmd As String) As Long
' this function performs LTFs
On Error Resume Next
    Dim str_frequency As String
    Dim str_gain As String
    Dim str_phase As String
    Dim Frequency As Double
    Dim gain As Double
    Dim Phase As Double
    Dim Coherence As Double
    Dim separator As Long   ' position of symbols ', '
    Dim temp_str As String
    Dim scan_buf As String
    Dim str_len As Long
'    Dim switch_freq_line As Long
    Dim indx As Long
    Dim WaitTime As Single

    cwBut_RT_Monitoring.Value = False ' cancel RT monitoring

    WaitTimeout = 60 * cwLTF_time_scale.Value '6 periods of 0.1 Hz
    If (WaitTimeout < 60) Then WaitTimeout = 60

    Test_in_progress = True
    line_counter = 0
    lstResponse.Clear
    Buffer = ""
    cmd_sent = False
    StartTime = Timer
another_line: '    Do
    Do
        WaitTime = Timer - StartTime
        DoEvents
    Loop Until (WaitTime > 0.1) '100 ms delay
read_again:
        scan_buf = GetLTFdata(ScanCmd, WaitTimeout) 'here it waits
'        here first time it receives ">~OK" and exits
        separator = InStr(scan_buf, "OK")
        If separator <> 0 Then
            GoTo read_again
        Else
            separator = InStr(scan_buf, "*")
            If separator <> 0 Then
                GoTo show_plot
            End If
        End If
    If scan_buf = "UserStop" Then
        DoTest = -1
        Exit Function
    End If

    separator = InStr(scan_buf, "#")  'find '#'
    ASCIIDataStr = Left(scan_buf, separator)
    separator = InStr(UCase(scan_buf), "CMD")
    If separator <> 0 Then ' \r\n>~Doing CMD\r\n
        temp_str = Left(scan_buf, separator + 3)
        lstResponse.AddItem (temp_str)
        line_counter = line_counter + 1
        GoTo another_line
    End If

    LabelCurrentLTFline.Caption = (Format(line_counter, "00#") & "->" & ASCIIDataStr)
    lstResponse.AddItem (LabelCurrentLTFline.Caption)
    If InStr(scan_buf, "*") <> 0 Or ASCIIDataStr = "" Then
        GoTo show_plot ' END of transmission
    End If
    If ASCIIDataStr = "Timeout" Then  ' something bad, reinitialize stage
        lstResponse.AddItem ("Timeout")
        Call ShowNoCommunication
        Test_in_progress = False
        DoTest = 0
        Exit Function
    End If

    'we've got the measurement string, format
    'fff.ff, ggg.gg, ppp.pp #\r\n
    'fff.ff - frequency
    'ggg.gg - gain in dB
    'ppp.pp - phase in degrees wrapped

    If Len(ASCIIDataStr) = 0 Then ' bad data line
        Call ShowNoCommunication
        GoTo show_plot
    End If
    If Len(ASCIIDataStr) > 10 Then ' maybe valid format, lets confirm
        separator = InStr(ASCIIDataStr, "#") ' find end of line
        If (separator = 0) Then       ' not a valid line
           GoTo another_line ' for test
        End If
        temp_str = ASCIIDataStr

        If line_counter = 30 Then
           indx = 0 'test trap
        End If
        str_len = Len(temp_str)

        ' convert frequency
        separator = InStr(temp_str, ",") ' find
        If (separator > 1) Then
            str_frequency = Left(temp_str, separator - 1)
        Else
            str_frequency = "n/a"  'error
        End If
        If (IsNumber(str_frequency)) Then
            Frequency = Val(str_frequency) 'Val ALWAYS look for "." as fraction separator, not depending on locale
         ' trim frequency - we've converted it
            temp_str = Mid(temp_str, separator + 2)
        Else
            GoTo another_line ' for test
        End If

        ' convert gain
        separator = InStr(temp_str, ",") ' find
        If (separator > 1) Then
            str_gain = Left(temp_str, separator - 1)
        Else
            str_gain = "n/a"  'error
        End If
        If (IsNumber(str_gain)) Then
            gain = Val(str_gain)
            ' trim gain - we've converted it
            temp_str = Mid(temp_str, separator + 2)
        Else
            GoTo another_line ' for test
        End If

        ' convert phase
        separator = InStr(temp_str, ",") ' find phase
        If (separator > 1) Then
            str_phase = Left(temp_str, separator - 1)
        Else ' could be older firmware
            separator = InStr(temp_str, "#") ' find end of line
            If (separator > 1) Then
                str_phase = Left(temp_str, separator - 1)
            Else
                str_phase = "n/a"  'error
            End If
        End If
        If (IsNumber(str_phase)) Then
           Phase = Val(str_phase)
            ' trim phase - we've converted it
            temp_str = Mid(temp_str, separator + 2)
        Else
            GoTo another_line ' for test
        End If

        ' convert coherence
        separator = InStr(temp_str, ",") ' find coherence
        If (separator > 1) Then
            temp_str = Left(temp_str, separator - 1)
        Else ' could be older firmware
            separator = InStr(temp_str, "#") ' find end of line
            If (separator > 1) Then
                temp_str = Left(temp_str, separator - 1)
            Else ' old firmware without coherence
                temp_str = "n/a"
            End If
        End If
        If (IsNumber(temp_str)) Then
            Coherence = Val(temp_str)
        Else
            Coherence = 0
        End If

     Else   ' the length is less than expected
        GoTo another_line ' for test
     End If
'    Else 'for program test - without getting real data from controller
'        percent_divider = 500 ' approximate number of lines received from controller
'        switch_freq_line = 100
'        If line_counter >= percent_divider Or (waitresponse = False) Then
'            GoTo show_plot
'        Else
'        switch_line = percent_divider - switch_freq_line '400
'         If line_counter < switch_line Then '<400
'            frequency = (switch_line - line_counter) + switch_freq_line / 10 'switch freq = 10 Hz
'            ' (400-line_counter) + 100/10 = 10+(400-line_counter)
'            If line_counter > switch_line - 3 Then 'test trap
'                lblTestTime.Caption = Format(Timer - StartTime, "#0.0")
'            End If
'         Else
'            frequency = (percent_divider - line_counter) / (switch_freq_line / 10)
'            ' (500 - 400-5)/(100/10) = 90/10 = 9
'         End If
'            If frequency < 0.1 Then frequency = 0.1
'            aTimerTestProgram.Interval = 500 / frequency
'            If aTimerTestProgram.Interval < 5 Then aTimerTestProgram.Interval = 5 ' 5 millisecond min interval in test mode
'    '        aTimerTestProgram.Interval = 10
'            test_timer_ready = False
'            aTimerTestProgram.Enabled = True
'
'            Do
'                DoEvents
'            Loop While (test_timer_ready = False)
'
'            gain = calculate_test_transfer_function(frequency)
'            phase = calculate_test_phase(frequency)
'        End If
'    End If
     ' assign values to the array
        updateProgress (line_counter / (percent_divider + 1))
        lblTestTime.Caption = Format(Timer - StartTime, "#0.0")
        Freq_data(line_counter) = Frequency
        Gain_data(line_counter) = gain
        Phase_data(line_counter) = Phase

        NegPhase_data(line_counter) = Make180phaseShift(Phase)
        Coherence_data(line_counter) = Coherence * 100
        If Coherence <> 0 Then
            IsCoherenceAvailable = True
        Else
            IsCoherenceAvailable = False
        End If

        If line_counter >= 105 Then 'test trap
            Call update_plot(line_counter)
        End If

        Call update_plot(line_counter)
        CWGraph.Caption = "Freq = " & Format(Frequency, "#0.0") & ";  Gain = " & Format(gain, "0#.0")
        CWGraphPhase.Caption = "Freq = " & Format(Frequency, "#0.0") & ";  Phase = " & Format(Phase, "0##.#")
        line_counter = line_counter + 1

        GoTo another_line
show_plot:
        IsCoherenceAvailable = CoherenceIsPresent()
        percent_divider = line_counter ' exact number of lines received from controller
        Test_in_progress = False
        cwButStartStopLTF.Value = False
        cwSldProgress.Value = 100
        cwSldProgress.Caption = "! DONE !"
        DoTest = line_counter
        Call find_unity_gain_phase(line_counter)
End Function

Sub Update_prev_plots()
    Dim i As Long
    'update previous plots
    For i = 0 To TEST_LTF_ARRAY_LENGTH
        Prev_Freq_data(i) = Freq_data(i)
        Prev_Phase_data(i) = Phase_data(i)
        NegPrev_Phase_data(i) = Make180phaseShift(Phase_data(i))
        Prev_Gain_data(i) = Gain_data(i)
        Prev_Coherence_data(i) = Coherence_data(i)
    Next i
    CWGraph.Plots(PreviousPlot_Num).PlotXvsY Prev_Freq_data, Prev_Gain_data

    If PhaseReversed = False Then
        CWGraphPhase.Plots(PreviousPlot_Num).PlotXvsY Prev_Freq_data, Prev_Phase_data
    Else
        CWGraphPhase.Plots(PreviousPlot_Num).PlotXvsY Prev_Freq_data, NegPrev_Phase_data
    End If
    If (IsCoherenceAvailable) Then
        CWGraphPhase.Plots(PrevCoherence_TF_Num).PlotXvsY Prev_Freq_data, Prev_Coherence_data
    End If
End Sub

Sub find_unity_gain_phase(start_index As Long)
    Dim indx As Long
    Dim min_indx As Long
    Dim max_indx As Long
    Dim min_str As String
    Dim max_str As String
    Dim resonance_str As String
    min_str = ""
    max_str = ""
    Dim min_Unity_Phase_present As Boolean
    Dim MinIndx As Long, minVal As Double, MaxIndx As Long, maxVal As Double
    Dim ApproxFreq As Double

    min_indx = UBound(Freq_data)
    ' find rasing gain cross by going from the end of array (where freq = 0.1 Hz) towards high freq, beginning of array
    max_indx = 0
    minVal = -100
    For indx = (start_index) To 1 Step -1  ' find zero-cross, array is in reverse, first index with max freq, start_index when freq=0.1
        If ((Gain_data(indx)) >= 0) And (minVal < 0) Then ' previous phase was negative, now positive -> crossed zero, increasing
            ApproxFreq = Freq_data(indx) + (Freq_data(indx + 1) - Freq_data(indx)) * Gain_data(indx) / ((Gain_data(indx)) - Gain_data(indx + 1))
            min_str = Format(ApproxFreq, "#0.00Hz ")
            min_indx = indx
            Exit For
        End If
        minVal = Gain_data(indx)
    Next indx

    ' find dropping gain cross by going from the rasing cross freq  towards high freq, beginning of array
    minVal = 100
    For indx = (min_indx - 1) To 1 Step -1
'        If (Gain_data(indx) < 0.2) And (max_indx = 0) Then
        If ((Gain_data(indx)) <= 0) And (minVal > 0) Then ' previous phase was negative, now positive -> crossed zero, increasing
            ApproxFreq = Freq_data(indx) + (Freq_data(indx + 1) - Freq_data(indx)) * Gain_data(indx + 1) / ((Gain_data(indx + 1)) - Gain_data(indx))
            max_str = Format(ApproxFreq, "#0.0Hz ")
            max_indx = indx
            Exit For
        End If
        minVal = Gain_data(indx)
    Next indx

    indx = FindMaxMin(Gain_data, 0, maxVal, minVal, MaxIndx, MinIndx) 'ignore zeroes = false
    If (indx = 1) Then 'found max
        resonance_str = "Max @ " & Format(Freq_data(MaxIndx), "#0.0Hz") & " (" & Format(maxVal, "#0.0dB") & ")"
    Else
        resonance_str = "peak not foud"
    End If
    min_str = "        Unity Gain @ " & min_str & "(~0dB, " & Format(Phase_data(min_indx), "##0.0deg") & ")"
    max_str = "Unity Gain @ " & max_str & "(~0dB, " & Format(Phase_data(max_indx), "##0.0deg") & ")"
    CWGraph.Caption = min_str & " <= " & resonance_str & " => " & max_str

'calculate Unity Phase

        min_indx = UBound(Freq_data)
        max_indx = 0
        min_Unity_Phase_present = False
        minVal = 180
        For indx = (MaxIndx) To start_index ' find lower than resonance zero-cross, array is in reverse, first index with max freq
            If ((Phase_data(indx)) >= 0) And (minVal < 0) Then ' previous phase was negative, now positive -> crossed zero, increasing
                ApproxFreq = Freq_data(indx - 1) + (Freq_data(indx - 1) - Freq_data(indx)) * Phase_data(indx - 1) / ((Phase_data(indx)) - Phase_data(indx - 1))
                min_str = Format(ApproxFreq, "#0.00Hz ")
                min_str = "  Zero Phase at: " & min_str & "(" & Format(Gain_data(indx), "##0.0 dB)") 'separate & "(" & is needed because Format creates -(14.3dB) instead of (-14.3dB)
                min_Unity_Phase_present = True
                Exit For
            Else
                min_str = "No Zero Phase below 10 Hz"
            End If
            minVal = Phase_data(indx)
        Next indx

        minVal = -180
        For indx = (MaxIndx) To 0 Step -1 ' find lower than resonance zero-cross, array is in reverse, first index with max freq
            If ((Phase_data(indx)) <= 0) And (minVal > 0) Then ' previous phase was positive, now negative -> crossed zero, decreasing
                ApproxFreq = Freq_data(indx) + (Freq_data(indx + 1) - Freq_data(indx)) * Phase_data(indx) / ((Phase_data(indx)) - Phase_data(indx + 1))
                max_str = Format(ApproxFreq, "#0.0Hz")
                max_indx = indx - 1
                Exit For
            End If
            minVal = Phase_data(indx)
        Next indx
        max_str = max_str & " (" & Format(Gain_data(max_indx), " ##0.0 dB)")

        CWGraphPhase.Caption = min_str & "  <= =>  " & max_str
End Sub

Sub updateProgress(percent As Double)
    DoEvents
    cwSldProgress.Value = (percent * 100) Mod 100
    cwSldProgress.Refresh
    cwSldProgress.Caption = "Progress " & Format(cwSldProgress.Value / 100, "#0%")
End Sub

Sub clear_arrays()
    Dim i As Long
    For i = 1 To TEST_LTF_ARRAY_LENGTH
        Gain_data(i) = 0
        Phase_data(i) = 0
        Coherence_data(i) = 0
        NegPhase_data(i) = Make180phaseShift(0)
        Freq_data(i) = 0
    Next i
End Sub
'Sub fill_arrays()
'    Dim i As Long
'    Dim displacement_x As Double
'    Dim displacement_u As Double
'    resonance = 2.5
'    damping = 0.25
'    For i = 1 To TEST_LTF_ARRAY_LENGTH
'       Freq_data(i) = i / 10
'       Gain_data(i) = calculate_test_transfer_function(Val(Freq_data(i)))
'       Phase_data(i) = calculate_test_phase(Val(Freq_data(i)))
'       NegPhase_data(i) = Make180phaseShift(Phase_data(i) )
'    Next i
'
'End Sub
'
'Function calculate_test_transfer_function(frequency As Double) As Double
'    Dim displacement_x As Double
'    Dim displacement_u As Double
'        resonance = 2.5
'        damping = 0.25
'
'       displacement_x = 1 + (2 * damping * frequency / resonance) ^ 2
'       displacement_u = Abs((1 - (frequency ^ 2 / resonance ^ 2) ^ 2 + (2 * damping * frequency / resonance) ^ 2))
'       calculate_test_transfer_function = 5 * Log((displacement_x / displacement_u) ^ 0.5)
'
'End Function
'Function calculate_test_phase(frequency as Double) As Double
'    calculate_test_phase = (-180 + 2 * frequency)
'    calculate_test_phase = ((calculate_test_phase + 180) Mod 360) - 180
'End Function
'
'Private Sub cmdTest_Click()
'    fill_arrays
'End Sub

Sub copy_arrays_to_plot(Xdat() As Double, Ydat() As Double, Xplot() As Double, Yplot() As Double)
    Dim i As Long
    Dim Arr_size As Long
    Arr_size = UBound(Xplot())
    For i = 0 To Arr_size
        Xplot(i) = Xdat(i)
        Yplot(i) = Ydat(i)
    Next i
End Sub

Private Sub CmdPlotReference_Click()
    Call ReadMeasurementPlot(1) '1 means save plot as reference
End Sub

Public Sub update_plot(Arr_size As Long)
    Dim X_data() As Double
    Dim Y_data() As Double
    ReDim X_data(Arr_size)
    ReDim Y_data(Arr_size)
'
    If Arr_size > 0 Then
        If PhaseReversed = False Then
'           Call CWGraphPhase.Plots(CurrentPlot_Num).PlotXvsY(Freq_data, Phase_data)  'this plot - shows line to zero
            Call copy_arrays_to_plot(Freq_data, Phase_data, X_data, Y_data) ' when copied to resizable array, does not show line to (0,0)
            Call CWGraphPhase.Plots(CurrentPlot_Num).PlotXvsY(X_data, Y_data)  'this plot
        Else
            Call copy_arrays_to_plot(Freq_data, NegPhase_data, X_data, Y_data)
            Call CWGraphPhase.Plots(CurrentPlot_Num).PlotXvsY(X_data, Y_data)  'this plot phase
        End If

        If (IsCoherenceAvailable) Then
            Call copy_arrays_to_plot(Freq_data, Coherence_data, X_data, Y_data) ' when copied to resizable array, does not show line to (0,0)
            Call CWGraphPhase.Plots(Coherence_TF_Num).PlotXvsY(X_data, Y_data)  'this plot coherence
        End If
        ' magnitude plots
        Call copy_arrays_to_plot(Freq_data, Gain_data, X_data, Y_data)
        Call CWGraph.Plots(CurrentPlot_Num).PlotXvsY(X_data, Y_data)
    End If
End Sub

'============ GRAPH MANIPULATIONS ===========================================================================================
Public Sub update_reference_plot()
    CWGraph.Plots(RefPlot_Num).LineColor = &H40C0&
    CWGraphPhase.Plots(RefPlot_Num).LineColor = &H40C0&
    CWGraph.Plots(RefPlot_Num).PlotXvsY Reference_Freq_data, Reference_Gain_data
    CWGraphPhase.Plots(RefPlot_Num).PlotXvsY Reference_Freq_data, Reference_Phase_data
End Sub

Private Sub OptGraphMode_Click(Index As Integer)
    '2-moves cursor
    '3-moves in X and Y (pan)
    '4-moves in X only
    '5-moves in Y only
    '6-cursor-selected zoom in
    '7-cursor-selected X zoom in
    '8-cursor-selected Y zoom in
    If Index = 0 Then
        CWGraph.TrackMode = 2  'cursor
        CWGraphPhase.TrackMode = 2  'cursor
    End If
    If Index = 1 Then
        CWGraph.TrackMode = 3  'pan
        CWGraphPhase.TrackMode = 3  'pan
    End If
    If Index = 2 Then
        CWGraph.TrackMode = 6  'zoom window
        CWGraphPhase.TrackMode = 6  'zoom window
    End If
'    If Index = 3 Then
'    'Reset the axes.
'    '    cwGraphDataAcq.Axes(1).AutoScaleNow
'        cwGraphDataAcq.Axes(2).AutoScaleNow
'        cwGraphDataAcq.Axes(3).AutoScaleNow
'    End If

End Sub

Private Sub cwAutoscale_ValueChanged(ByVal Value As Boolean)
    Const XminScale As Double = 0.1
    Const XmaxScale As Double = 1000  ' set  X = frequency scale
    Const YminScale As Double = -30
    Const YmaxScale As Double = 50
    Dim Xmax As Double
    Dim UserYminScale As Double
    Dim UserYmaxScale As Double

    Xmax = cwNumMaxFreq.Value

    'CWGraph.Axes.Item(2).AutoScale = True
        CWGraphPhase.Axes.Item(1).SetMinMax 0.1, Xmax
        CWGraphPhase.Axes.Item(2).SetMinMax -180, 180
    If cwAutoscale.Value = True Then
'        cwMax_dB.Enabled = False
'        cwMin_dB.Enabled = False
        CWGraph.Axes.Item(1).SetMinMax 0.1, Xmax
        CWGraph.Axes.Item(2).AutoScaleNow
    Else    'FULL SCALE
'        cwMax_dB.Enabled = True
'        cwMin_dB.Enabled = True
        CWGraph.Axes.Item(1).SetMinMax 0.1, Xmax
        UserYmaxScale = cwMax_dB.Value
        UserYminScale = cwMin_dB.Value
        CWGraph.Axes.Item(2).SetMinMax UserYminScale, UserYmaxScale
    End If
    CWGraph.Refresh
End Sub

Private Sub ChkShowCurrentPlot_Click()
    If ChkShowCurrentPlot.Value = Checked Then
        CWGraph.Plots(CurrentPlot_Num).Visible = True
        CWGraphPhase.Plots(CurrentPlot_Num).Visible = True
        If ((ChkShowCoherencePlot.Value = Checked) And (ChkShowCoherencePlot.Value = Checked)) Then
            CWGraphPhase.Plots(Coherence_TF_Num).Visible = True
        End If
    Else
        CWGraph.Plots(CurrentPlot_Num).Visible = False
        CWGraphPhase.Plots(CurrentPlot_Num).Visible = False
        CWGraphPhase.Plots(Coherence_TF_Num).Visible = False
    End If
End Sub

Private Sub ChkShowPrevPlot_Click()
    If ChkShowPrevPlot.Value = Checked Then
        CWGraph.Plots(PreviousPlot_Num).Visible = True
        CWGraphPhase.Plots(PreviousPlot_Num).Visible = True
        If ((ChkShowCoherencePlot.Value = Checked) And (ChkShowCoherencePlot.Value = Checked)) Then
            CWGraphPhase.Plots(PrevCoherence_TF_Num).Visible = True
        End If
    Else ' hide
        CWGraph.Plots(PreviousPlot_Num).Visible = False
        CWGraphPhase.Plots(PreviousPlot_Num).Visible = False
        CWGraphPhase.Plots(PrevCoherence_TF_Num).Visible = False
    End If
End Sub

Private Sub ChkShowRefPlot_Click()
    If ChkShowRefPlot.Value = Checked Then
        CWGraph.Plots(OneFilterTF_Num).Visible = True
        CWGraph.Plots(Axis_TF_Num).Visible = True
        CWGraph.Plots(RefPlot_Num).Visible = True
        CWGraphPhase.Plots(OneFilterTF_Num).Visible = True
        CWGraphPhase.Plots(Axis_TF_Num).Visible = True
        CWGraphPhase.Plots(RefPlot_Num).Visible = True
    Else
        CWGraph.Plots(OneFilterTF_Num).Visible = False
        CWGraph.Plots(Axis_TF_Num).Visible = False
        CWGraph.Plots(RefPlot_Num).Visible = False
        CWGraphPhase.Plots(OneFilterTF_Num).Visible = False
        CWGraphPhase.Plots(Axis_TF_Num).Visible = False
        CWGraphPhase.Plots(RefPlot_Num).Visible = False
    End If
End Sub

Sub Calc_prediction()
    Dim freq_pt As Long
    Dim prev_freq As Double
    Dim freq_sum_for_test As Double
    freq_sum_for_test = 0
    For freq_pt = 0 To TEST_LTF_ARRAY_LENGTH  '0 to 200
        If Reference_Freq_data(freq_pt) > 0 Then 'prepare test array for axis OLTF calculation
            prev_freq = Reference_Freq_data(freq_pt)         'prepare test array for axis OLTF calculation
'            Freq_points(freq_pt) = prev_freq
'        Else
'            Freq_points(freq_pt) = prev_freq ' fill up to the end with the last valid frequency
        End If
        Freq_points(freq_pt) = prev_freq ' fill up to the end with the last valid frequency
        freq_sum_for_test = freq_sum_for_test + prev_freq
    Next freq_pt
    freq_pt = freq_pt - 1
    If freq_sum_for_test = 0 Then
        If (ChkShowRefPlot.Value = Checked) Then
            retVal = MsgBox("Need reference for prediction" + vbCrLf + "Please load Reference Transfer Function" + vbCrLf + "Use 'Get Ref Plot' button", vbOKOnly + vbExclamation, "Reference plot not loaded")
        End If
        Exit Sub
    End If
    ' recalculate axis TF based on real frequecy points
    Call frmFilters.CalculateAxis_TF(PhaseReversed)
    For freq_pt = 0 To TEST_LTF_ARRAY_LENGTH '0 to 200
    ' this is reference plot * Axis_Mag
    ' IK20251218 added experimantal prediction correction for floor FF
        Predicted_Mag(freq_pt) = frmFilters.PredictionCorrectionFloorFF(Reference_Gain_data(freq_pt) + Axis_Mag(freq_pt)) 'in dB already

        Predicted_Phase(freq_pt) = frmFilters.Phase_Limit_TF_plus_minus_180(Reference_Phase_data(freq_pt) + Axis_Phase(freq_pt))
    Next freq_pt
    freq_pt = freq_pt - 1 '
    CWGraph.Plots(Prediction_TF_Num).PlotXvsY Freq_points, Predicted_Mag
    CWGraphPhase.Plots(Prediction_TF_Num).PlotXvsY Freq_points, Predicted_Phase

    CWGraph.Plots(Prediction_TF_Num).Visible = True
    CWGraphPhase.Plots(Prediction_TF_Num).Visible = True

End Sub

Private Sub ChkShowPredictionPlot_Click()
    If ChkShowPredictionPlot.Value = Checked Then
        Call Calc_prediction
    Else
        CWGraph.Plots(Prediction_TF_Num).Visible = False
        CWGraphPhase.Plots(Prediction_TF_Num).Visible = False
    End If
    CWGraph.Refresh
End Sub

Private Sub cwGraph_CursorChange(CursorIndex As Long, Xpos As Variant, Ypos As Variant, bTracking As Boolean)
     CWGraph.Caption = "Freq = " & Format(Xpos, "#000.0") & ";  Gain = " & Format(Ypos, "0#.0")
End Sub

Private Sub CWGraph_Click()
    CWGraph.Cursors(1).Visible = True
End Sub

Private Sub CWGraph_dblClick()
    CWGraph.Cursors(1).Visible = False
End Sub

Public Function ArrayPosition(ByVal FindValue As Variant, arrSearch As Variant) As Long
    Dim lngLoop As Long
    ArrayPosition = -1  'Set default value of "not found"
    On Error GoTo LocalError
    If Not IsArray(arrSearch) Then Exit Function
    For lngLoop = LBound(arrSearch) To UBound(arrSearch)
      If (arrSearch(lngLoop)) = FindValue Then
        ArrayPosition = lngLoop
        Exit Function
      End If
    Next lngLoop
Exit Function
LocalError:
    'Nothing
End Function

Private Sub cwGraphPhase_CursorChange(CursorIndex As Long, Xpos As Variant, Ypos As Variant, bTracking As Boolean)
     Dim plot_caption As String
     plot_caption = "Freq = " & Format(Xpos, "#000.0") & ";  Phase = " & Format(Ypos, "0#.#")
     If IsCoherenceAvailable = True Then
        Dim indx As Long
        indx = ArrayPosition(Xpos, Freq_data)
        plot_caption = plot_caption & "; Coherence = " & Format(Coherence_data(indx), "0.####") & "%"
     End If
     CWGraphPhase.Caption = plot_caption
End Sub
Private Sub CWGraphPhase_Click()
    CWGraphPhase.Cursors(1).Visible = True
End Sub

Private Sub CWGraphPhase_dblClick()
    CWGraphPhase.Cursors(1).Visible = False
End Sub

Private Sub cwSaveScan_ValueChanged(ByVal Value As Boolean)
    'one measurement
        If cwSaveScan.Value = True Then
            Call SaveRawMeasurementPlot(True) 'Real_true_Prediction_false
            cwSaveScan.Value = False
        End If
End Sub

Private Sub Cmd_SavePrediction_Click()
    If (Predicted_Mag(0) <> 0) Then ' check if prediction was made
        Call SaveRawMeasurementPlot(False) 'Real_true_Prediction_false
    Else
        MsgBox ("No prediction data")
    End If
End Sub

Sub SaveRawMeasurementPlot(Real_true_Prediction_false As Boolean)
    On Error Resume Next
    'user wants to save a test file
    'First get the filename
    Dim DefaultFilename As String
    Dim SaveFileName As String
    Dim FileName As String
    Dim i  As Long
    Dim chGain As String
    Dim chPower As String
    Dim chAxis As String
    Dim chData As String
    Dim chFreq As Double

    cwBut_RT_Monitoring.Value = False ' cancel RT monitoring
    DiagnosticAxis = InputChosen - 32 ' 32 is the index of "X axis Pos" choice on the frmSensorInput
    chGain = Format(cwNumFBgain(DiagnosticAxis).Value, "#0.00")
    'If Len(LTF_command) < 6 Then
    If LTFtestOutputName = "" Or LTFtestInputName = "" Then
        response = MsgBox("Please choose Output and Input first", vbOKOnly + vbExclamation, "Parameters Not Defined")
        Exit Sub
    Else
        If (LTF_command = "") Then Call create_LTF_command
    End If

    DefaultFilename = DefaultFilename & LTF_command & " " & Date$ & " " & Left$(Time$, 2) & "-" & Mid$(Time$, 4, 2) & _
    ", gain=" & chGain & " excit_ampl=" & TestGain & ".csv"
    dlgCommonDialog.CancelError = True
    dlgCommonDialog.DialogTitle = "Choose name for the measurement"
    ' dlgCommonDialog.FileName = DefaultFilename
    dlgCommonDialog.InitDir = "c:\test"
    dlgCommonDialog.DefaultExt = "csv"
    dlgCommonDialog.Filter = "Data File (*.csv)|*.csv|All files (*.*)|*.*"
    dlgCommonDialog.FileName = DefaultFilename 'if enabled,  FILE NAME SHOULD NOT HAVE ':'
    dlgCommonDialog.ShowSave
    If Err = 32755 Then 'Cancel is selected
        Exit Sub
    Else
        txtStatus.Text = "Saving measurement into file '" & dlgCommonDialog.FileName & " ....."
    End If
    'SaveFileName becomes "C:\test\my_test.txt"

'    SaveFileName = Left(dlgCommonDialog.FileName, Len(dlgCommonDialog.FileName) - 4) ' remove ".txt" from full file path
    SaveFileName = dlgCommonDialog.FileName
    'SaveFileName becomes "C:\test\my_test"
    i = InStrRev(SaveFileName, "\")
    FileName = Right(SaveFileName, Len(SaveFileName) - i) 'i.e. "my_test"

    SaveFileName = dlgCommonDialog.FileName
    'now lets open the file and write the data
    Open SaveFileName For Output As #Transfer_Function_file 'Values
        Print #Transfer_Function_file, "File name,"; FileName
        Print #Transfer_Function_file, "Date,"; Date$
        Print #Transfer_Function_file, "Time,"; Time$
        Print #Transfer_Function_file, "Electro_Damp-Type System ", SystemFWversion
        Print #Transfer_Function_file, "Analyzer P/N,"; lbl_Part_Number.Caption; ",Verson,"; lbl_Version_Date
        Print #Transfer_Function_file, "Excit Ampl,"; Format(TestGain, "#0")
    If Real_true_Prediction_false = False Then ' predicted OLTF plot
        Print #Transfer_Function_file, "Frequency Hz,Predicted dB,Predicted Ph Deg"
    Else ' real plot
        Print #Transfer_Function_file, "Frequency Hz,Gain dB,Phase Deg,Coherence"
    End If
    For i = 0 To UBound(Gain_data)
        If Real_true_Prediction_false = False Then ' predicted OLTF plot
            If (Reference_Freq_data(i) > 0) Then
                'chFreq = Reference_Freq_data(i) ' do not print empty data
                'If chFreq * 100 = 0 Then Exit For
                chData = Str(Reference_Freq_data(i)) ' Str always converts decimal separator to "." reverse function to Val()
                chFreq = chData
                If (chData < 10) Then
                    'chFreq = Round(chData, 1)
                    Print #Transfer_Function_file, Format(chFreq, "#0.0"); ", "; ' frequency
                Else
                    Print #Transfer_Function_file, Format(chFreq, "#0"); ", ";
                End If
                chData = Str(Predicted_Mag(i))
                Print #Transfer_Function_file, Format(chData, "#0.00"); ", ";   'magnitude
                chData = Str(Predicted_Phase(i))
                Print #Transfer_Function_file, Format(chData, "#0.000"); ","; vbTab; "#" ' phase
            Else ' no prediction made
                Exit For
            End If
        Else ' real plot
            If (Freq_data(i) > 0) Then
                'chFreq = Reference_Freq_data(i) ' do not print empty data
                'If chFreq * 100 = 0 Then Exit For
                chData = Str(Freq_data(i)) ' Str always converts decimal separator to "." reverse function to Val()
                chFreq = chData
                If (chData < 10) Then
                    'chFreq = Round(chData, 1)
                    Print #Transfer_Function_file, Format(chFreq, "#0.0"); ", ";
                Else
                    Print #Transfer_Function_file, Format(chFreq, "#0"); ", ";
                End If
                chData = Str(Gain_data(i))
                Print #Transfer_Function_file, Format(chData, "#0.00"); ", ";
                chData = Str(Phase_data(i))
                If (IsCoherenceAvailable) Then
                    Print #Transfer_Function_file, Format(chData, "#0.00"); ", ";
                    chData = Str(Coherence_data(i)) * 0.01
                End If
                Print #Transfer_Function_file, Format(chData, "#0.000"); ","; vbTab; "#"
            Else ' no prediction made
                Exit For
            End If
        End If
    Next i
    Close #Transfer_Function_file
    cwSaveScan.ToolTipText = "Saves current plots from screen into data file, last: '" + FileName + "'"
    txtStatus.Text = "File saved: " & dlgCommonDialog.FileName
End Sub

Sub ReadMeasurementPlot(Optional ref_plot As Integer = 0)
    On Error Resume Next
    'user wants to read a test file
    'First get the filename
    Dim DefaultFilename As String
    Dim DataFileName As String
    Dim FileName As String
    Dim i As Long
    Dim record_str As String
    Dim temp_str As String
    Dim sep_pos As Long
    Dim meas_line_identified As Long
    Dim line_counter As Long
    Dim TFtype_from_file As String
    Dim sensor_from_file As String
    Dim actuator_from_file As String

'    frmTerminal.Hide ' or it gives message "can't show non-modal form when modal form is displayed

    cwBut_RT_Monitoring.Value = False ' cancel RT monitoring

   DefaultFilename = Date$ & " " & Left$(Time$, 5) & "_tranasfer_function.csv"
    dlgCommonDialog.CancelError = True
    dlgCommonDialog.DialogTitle = "Choose data file name to read and show"
    ' dlgCommonDialog.FileName = DefaultFilename
    dlgCommonDialog.InitDir = "c:\test"
    dlgCommonDialog.DefaultExt = "csv"
    dlgCommonDialog.Filter = "Data File (*.csv)|*.csv|All files (*.*)|*.*"
    dlgCommonDialog.Action = 1
    'dlgCommonDialog.ShowSave
    If Err = 32755 Then 'Cancel is selected
        Exit Sub
    Else
        txtStatus.Text = "Reading measurement from file '" & dlgCommonDialog.FileName & " ....."
    End If
    'SaveFileName becomes "C:\test\my_test.txt"

'    SaveFileName = Left(dlgCommonDialog.FileName, Len(dlgCommonDialog.FileName) - 4) ' remove ".txt" from full file path
    DataFileName = dlgCommonDialog.FileName
    i = InStrRev(DataFileName, "\")
    FileName = Right(DataFileName, Len(DataFileName) - i) 'i.e. "my_test"

    'now lets open the file and read the data

    line_counter = 0
    Open DataFileName For Input As #Transfer_Function_file 'Values

    Line Input #Transfer_Function_file, record_str '    "File name," FileName
    ' decode line
    sep_pos = InStr(record_str, ",")
    If LCase(Left(record_str, sep_pos - 1)) <> "file name" Then
        If (IsNumber(Left(record_str, 3))) Then
            line_counter = -1 ' to start from index 0
            Call clear_arrays
            GoTo matlab
        Else
            GoTo error_format
        End If
    Else
        If (ref_plot = 0) Then
            Call Update_prev_plots
            Call clear_arrays
            txtStatus.Text = Mid(record_str, sep_pos + 1)
            CmdReadScan.ToolTipText = "Reads saved data and shows plots, now '" + FileName + "'"
        End If
    End If
    line_counter = 1 ' got first line already
    Do
        Line Input #Transfer_Function_file, record_str 'Print #Transfer_Function_file,
        line_counter = line_counter + 1
        meas_line_identified = InStr(LCase(record_str), "freq")
    Loop Until ((meas_line_identified > 0) Or (line_counter > 10))

    line_counter = -1 ' to start from index 0
read_again:
    Do While (Not EOF(Transfer_Function_file))  ' Loop until end of file.

        Line Input #Transfer_Function_file, record_str 'Print #Transfer_Function_file,
matlab:
        temp_str = record_str ' record_str keeps copy for debug
        If (InStr(temp_str, "*")) Then ' end of TF might be written into file
            GoTo end_reading
        End If
        sep_pos = InStr(temp_str, ",%")
        If (sep_pos <> 0) Then 'NOT valid TF line, it is Real-Time line
            Close #Transfer_Function_file
            MsgBox ("This is Real Time File, choose Transfer Function File")
            Exit Sub
        End If
        sep_pos = InStr(temp_str, ",#")
        If (sep_pos <> 0) Or (meas_line_identified) Then 'valid TF line
            If (sep_pos = 0) Then sep_pos = Len(temp_str)
            temp_str = Left(temp_str, sep_pos - 1) ' remove comma and #
            line_counter = line_counter + 1


'Note   In Visual Basic, the Str and Val functions always assume a period is the decimal separator.
'In a majority of locales, this assumption is not valid. Instead, use the CStr, CDbl, CSng, CInt, and CLng functions
'to provide international conversions from any other data type to the data type you need.
'These functions use the system locale to determine the decimal separator.

            ' decode line
            sep_pos = InStr(temp_str, ",") ' Print #Transfer_Function_file, Format(Freq_data(i), "#0.00,");
            If IsNumber(Left(temp_str, sep_pos - 1)) Then
                Freq_data(line_counter) = Val(Left(temp_str, sep_pos - 1)) ' the Str and Val functions always assume a period is the decimal separator.
                temp_str = Mid(temp_str, sep_pos + 1)
            Else
                GoTo error_format
            End If
                'Print #Transfer_Function_file, Format(Gain_data(i), "#0.000,");
            sep_pos = InStr(temp_str, ",")
            If IsNumber(Left(temp_str, sep_pos - 1)) Then
                Gain_data(line_counter) = Val(Left(temp_str, sep_pos - 1))
                temp_str = Mid(temp_str, sep_pos + 1)
            Else
                GoTo error_format
            End If

                'Print #Transfer_Function_file, Format(Phase_data(i), "#0.0")
            If IsNumber(temp_str) Then
                'Phase_data(line_counter) = Phase_Limit_plus_minus_180(CSng(temp_str)) ' Csng() is locale - dependent
                Phase_data(line_counter) = (Val(temp_str))                             ' Va() is always use dot as decimal point
                NegPhase_data(line_counter) = Make180phaseShift(Phase_data(line_counter))
            Else
                GoTo error_format
            End If
            sep_pos = InStr(temp_str, ",")
            If sep_pos > 0 Then
                If IsNumber(Left(temp_str, sep_pos - 1)) Then
                    temp_str = Mid(temp_str, sep_pos + 1) ' separate POSSIBLE coherence
                End If
            Else
                temp_str = ""
            End If
            If IsNumber(temp_str) Then
                Coherence_data(line_counter) = (Val(temp_str)) * 100
            Else
                Coherence_data(line_counter) = 0
            End If

            'lblLines.Caption = line_counter
         If (Freq_data(line_counter) < 0.1) Then Exit Do
        Else 'not valid line without #, skip
        End If
    Loop
end_reading:

    Close #Transfer_Function_file
    IsCoherenceAvailable = CoherenceIsPresent()

    If Freq_data(0) = 0 Then Freq_data(0) = 1000
    If (ref_plot <> 0) Then
        'update reference plots ONLY
        For i = 0 To TEST_LTF_ARRAY_LENGTH
            Reference_Freq_data(i) = Freq_data(i)
            Freq_data(i) = 0
            Reference_Gain_data(i) = Gain_data(i)
            Gain_data(i) = 0
            Reference_Phase_data(i) = Phase_data(i)
            Phase_data(i) = 0
        Next i
        CmdPlotReference.ToolTipText = "Reads saved data and shows persistent Reference plot, now '" + FileName + "'"
        Call update_reference_plot
    Else 'standard test read = replace: current->previous; this->current
        Call update_plot(line_counter)
    End If
    Call find_unity_gain_phase(line_counter)
    If cwAutoscale.Value = True Then
        Call CWGraph.Axes(2).AutoScaleNow ' adjust Y-scale, plot might be not visible
        ' update plot Y axis controls
        cwMax_dB.Value = CWGraph.Axes.Item(2).Maximum
        cwMin_dB.Value = CWGraph.Axes.Item(2).Minimum
    End If
    txtStatus.Text = "Measurement file '" & dlgCommonDialog.FileName & "' read successfully"

    TFtype_from_file = Mid(FileName, 1, 4)
    actuator_from_file = Mid(FileName, 6, 4)
    sensor_from_file = Mid(FileName, 11, 4)
    Call ReadGoNoGoPlot(TFtype_from_file, actuator_from_file, sensor_from_file)
    Exit Sub
error_format:
    Close #Transfer_Function_file
    Call MsgBox("file '" & FileName & "'" & vbCrLf & "folder " & UserSelectedCriteriaPath, vbOKOnly + vbExclamation, "Unrecognizable LTF data file format")
    MsgBox ("Unrecognizable file format")
End Sub

Public Function GetDataString(end_of_data_token As String, end_of_string_token As String, WaitingTime As Single) As String ' string of ASCII data, ending with vbLf
    On Error GoTo port_problems
    Dim CmdStartTime As Single
    Dim WaitTime As Single
    Dim pos_end_of_string As Long
    Dim DataString As String
    Dim buffer_size As Long
    Dim pos_CrLf As Long
    Dim token_len As Long
   'clean the Comm port buffer if called the first time
    If do_clean_buffer = True Then
        Call Clean_COM_Buffer
        do_clean_buffer = False
    End If

    If WaitingTime = 0 Then
         WaitingTime = 0.5 ' read buffer for data to come back to the serial port during wating time seconds.
    End If

'initialize the timer
    CmdStartTime = Timer
    pos_end_of_string = 1

    Do
          Do 'small delay
            DoEvents
            WaitTime = Timer - CmdStartTime
          Loop Until (WaitTime > 0.002)
        Call Get_Serial_Input

        'remove echo only once
        If (echo_removed = False) Then Call Remove_echo(Buffer, SendString)
        If (echo_removed = True) Then echo_removed = False ' prepare for the next command

        token_len = Len(end_of_string_token)
        If Left(Buffer, 1) = vbCrLf Then Buffer = Mid(Buffer, 3)
        If Left(Buffer, 1) = vbLf Then Buffer = Mid(Buffer, 2)
        If Left(Buffer, 1) = vbCr Then Buffer = Mid(Buffer, 2)
        If Left(Buffer, 1) = vbLf Then Buffer = Mid(Buffer, 2)
        If Left(Buffer, 1) = vbCr Then Buffer = Mid(Buffer, 2)
        If (InStr(Buffer, "*") And Len(Buffer) < 12) Then
        ' typical last string has length about 22 without coherence and 30 with coherence
        '  0.1, -125.30,  40.16, 0.9832, #
        '*
            GetDataString = "*"
            Exit Function
        End If
    '**** here real data starts from second line, data string has header and first line
        pos_CrLf = InStr(LCase(Buffer), ">~doing") ' header ending ">~Doing CMD"\r\n
        If pos_CrLf <> 0 Then 'there is header ending "Doing CMD"\r\n
            Buffer = Mid(Buffer, pos_CrLf + Len(">~Doing CMD") + 2) ' remove "Doing CMD\r\n" isolate one line
        End If
        pos_end_of_string = InStr(Buffer, end_of_string_token)
        If pos_end_of_string <> 0 Then 'the complete data line is received
           DataString = Left(Buffer, pos_end_of_string + token_len - 1) ' -1 to be before ">" or "#\r\n", in data line ending with "\r\n>~OK\r\n or #\r\n
           Buffer = Mid(Buffer, pos_end_of_string + token_len)  ' and trim this line from buffer
    ' **** here buffer starts from second line, data string has header and first line
    '         pos_end_of_string = InStr(DataString, "CMD") ' header ending "Doing CMD"\r\n
    '         If pos_end_of_string <> 0 Then 'the complete data line is received
    '             DataString = Right(DataString, Len(DataString) - pos_end_of_string - 2) ' isolate one line
    '         End If
           If DataString = "" Then
    '     '        MsgBox ("empty string")
           End If
           GoTo got_data_string
        End If

        token_len = Len(end_of_data_token)
        pos_end_of_string = InStr(Buffer, end_of_data_token) ' check for "*\r\n" - end of test
        If (pos_end_of_string <> 0) Then 'Or (InStr(Buffer, "OK")) Then         'it is the last line
            DataString = Mid(Buffer, pos_end_of_string)
            If DataString = "" Then
        '        MsgBox ("empty string")
            End If
            GoTo got_data_string
        End If
        pos_CrLf = InStr(Buffer, vbCrLf)
        If pos_CrLf <> 0 Then 'the complete data line is received
    '**** here real data starts from second line, data string has header and first line
            pos_end_of_string = InStr(Buffer, ">~Doing CMD") ' header ending "Doing CMD"\r\n
            If pos_end_of_string <> 0 Then '">~Doing CMD" received
               Buffer = Mid(Buffer, pos_end_of_string + Len(">~Doing CMD") + 2) ' remove "doing cmd" isolate one line
            Else ' no ">~Doing CMD" received
                DataString = Left(Buffer, pos_CrLf - 1) ' isolate one line
                Buffer = Mid(Buffer, pos_CrLf) ' and trim this line from buffer
                GoTo got_data_string
            End If
        End If
    Loop Until WaitTime > WaitingTime Or InStr(Buffer, vbLf) Or (waitresponse = False)
    If DataString = "" Then
'        MsgBox ("empty string")
    End If

    If WaitTime > WaitingTime Then
'        DataString = "Timeout"
'        MsgBox ("GetDataString timeout")
    End If
got_data_string:
    WaitTime = Timer - CmdStartTime
    If DataString = "" Then
'        MsgBox ("DataString empty")
    End If
    GetDataString = DataString
'--    lblRawResponse.Caption = DataString
    Exit Function

port_problems:
    Call portproblems
End Function


Public Sub SelectPrinter(NewPrinter As String)
    Dim Prt As Printer

    For Each Prt In Printers
        If Prt.DeviceName = NewPrinter Then
            Set Printer = Prt
        Exit For
        End If
    Next
End Sub

Private Sub WinNTSetDefaultPrinter(PrinterName As String)
    Dim PrintBuffer As String
    Dim DeviceName As String
    Dim DriverName As String
    Dim PrinterPort As String
    Dim r As Long
'    If ListCommIssues.ListIndex > -1 Then
        ' Get the printer information for the currently selected
        ' printer in the list. The information is taken from the
        ' WIN.INI file.
        PrintBuffer = Space(1024)
'        PrinterName = ListCommIssues.Text
        r = GetProfileString("PrinterPorts", PrinterName, "", _
            PrintBuffer, Len(PrintBuffer))

        ' Parse the driver name and port name out of the buffer
        GetDriverAndPort PrintBuffer, DriverName, PrinterPort

           If DriverName <> "" And PrinterPort <> "" Then
               SetDefaultPrinter PrinterName, DriverName, PrinterPort
               If Printer.DeviceName <> PrinterName Then
               ' Make sure Printer object is set to the new printer
                  SelectPrinter (PrinterName)
               End If
           End If
 '      End If
End Sub

Private Sub GetDriverAndPort(ByVal Buffer As String, DriverName As _
    String, PrinterPort As String)

    Dim iDriver As Long
    Dim iPort As Long
    DriverName = ""
    PrinterPort = ""

    ' The driver name is first in the string terminated by a comma
    iDriver = InStr(Buffer, ",")
    If iDriver > 0 Then

         ' Strip out the driver name
        DriverName = Left(Buffer, iDriver - 1)

        ' The port name is the second entry after the driver name
        ' separated by commas.
        iPort = InStr(iDriver + 1, Buffer, ",")

        If iPort > 0 Then
            ' Strip out the port name
            PrinterPort = Mid(Buffer, iDriver + 1, _
            iPort - iDriver - 1)
        End If
    End If
End Sub

Sub GetSysData()
  On Error Resume Next
    Dim temp_str As String
    Dim indx As Long
    Dim cont_field_monitoring As Boolean
    Dim paramValue As Double
    Dim bnc_index_found As Boolean
    Dim ControllerCommand As String
    Init_sys_data = True
'    lstResponse.Clear
'    cont_field_monitoring = frmMonitoring.aTimerContMonitoring.Enabled ' save status
'    frmMonitoring.aTimerContMonitoring.Enabled = False ' disable update - not to interfere with these com port requests
    Call frmTerminal.QuitMenu

    If (M3v_Diagnostic_running) Then
        temp_str = GetSend("oltf>stop", False) 'cancel oltf just in case
    End If
'    If ID_valid = False Then
    If ID_valid = False Then ID_valid = get_ID() ' try again, if fails THIS TIME, exit
    If ID_valid = False Then Exit Sub 'do not try if there is no proper response on "p" or "vers"

    ControllerCommand = GetSend("sysid", True)
    temp_str = (Mid(ControllerCommand, 9, 20)) 'sisID is 20 chars long
    TxtSystemID.Text = temp_str

    For indx = 0 To cwNum_Go_NoGO_window.UBound
    ' go - no go window settings
        paramValue = RectifyValueResponse(cwNum_Go_NoGO_window(indx).Tag & "?") 'gngx?
        If (paramValue <> WRONG_VALUE) Then
            cwNum_Go_NoGO_window(indx).Value = CInt(paramValue)
        End If
    Next indx

    For indx = 0 To cwNumHeightAdj.UBound
    ' iso offset settings
        paramValue = RectifyValueResponse(cwNumHeightAdj(indx).Tag & "?") 'ofs1? ofs2? ofs3?
        If (paramValue <> WRONG_VALUE) Then
            cwNumHeightAdj(indx).Value = CInt(paramValue)
        End If
    Next indx

    For indx = 0 To 3
    ' Pressure set points settings
        paramValue = RectifyValueResponse(cwNum_PressureSetPoint(indx).Tag & "?") 'gain31? gain32? gain33? gain34?
        If (paramValue <> WRONG_VALUE) Then
            cwNum_PressureSetPoint(indx).Value = CSng(paramValue)
        End If
    Next indx

    paramValue = RectifyValueResponse(cwNum_DockedPos.Tag & "?") 'dkpo?
    '"dkpo?": if firmware before 2016-nov-16 this command was commented out
    'so it returns ">~Unrecognized Cmd, Error Code =1", then RectifyValueResponse returnd 77777.7
    If (paramValue = WRONG_VALUE) Then ' old firmware - set color orange
        cwNum_DockedPos.BackColorText = &H80C0FF
        cwNum_DockedPos.Value = 0
        cwNum_DockedPos.Enabled = False
        cwNum_DockedPos.ToolTipText = "not available in this version of firmware, use Terminal menu->Pressure control to access"
    Else ' command recognized
        cwNum_DockedPos.BackColorText = &HFFFFFF
        cwNum_DockedPos.Value = CInt(paramValue)
        cwNum_DockedPos.Enabled = True
    End If

    ControllerCommand = "bncd0"
    bnc_index_found = False
'    cmd_sent = False
'    waitresponse = True
    paramValue = RectifyValueResponse(ControllerCommand)
    If (paramValue <> WRONG_VALUE) Then
        ' update drop-down lists
        BNC_0_index = CInt(paramValue)
        CmbBNC0.ListIndex = BNC_0_index

        ' find if index matches with slider control
        bnc_index_found = Find_Index_Match_with_BNC_SLIDER(0, BNC_0_index)

        If bnc_index_found = False Then 'index is not in the limited set of slider known indexes
            cwBNC0select.ValuePairIndex = 38 ' other 57 ' Pos Z
        End If
    Else
        cwBNC0select.ValuePairIndex = 38 ' other 57 ' Pos Z
    End If

    ControllerCommand = "bncd1"
    bnc_index_found = False
    cmd_sent = False
    waitresponse = True
    paramValue = RectifyValueResponse(ControllerCommand)
    If (paramValue <> WRONG_VALUE) Then            ' update drop-down lists
        BNC_1_index = CInt(paramValue)
        CmbBNC1.ListIndex = BNC_1_index

        ' find if index matches with slider control
        bnc_index_found = Find_Index_Match_with_BNC_SLIDER(1, BNC_1_index)

        If bnc_index_found = False Then 'index is not in the limited set of slider known indexes
            cwBNC1select.ValuePairIndex = 38 ' other 5 ' ProxZ2
        End If
    Else
        cwBNC1select.ValuePairIndex = 38 ' other 5 ' ProxZ2
    End If

    If ActiveTabIndex <> SystemParams_FrameIndex Then
        Call Get_diag_params
    End If
    Call Refresh_gains

    Call ShowSysStatus      ' "s" command
'    Init_sys_data = False
'   frmMonitoring.aTimerContMonitoring.Enabled = cont_field_monitoring  ' restore status
end_get_sys_data:
    Init_sys_data = False
End Sub

' Function GetSysStatus() As String
'    Dim temp_str As String
'    If ID_valid = False Then ID_valid = get_ID()
''vers
'' 95-34577-03 rev.At3 @ 14-Apr-2011
''>~OK
'    If ID_valid = False Then
'        GetSysStatus = "No Controller Found"
'        Exit Function
'    End If
'    temp_str = GetSend("s", True) 'get status byte as hex
''   "s"
''   Front port response "The System Status is: 8e"
''   Rear port response "8e"
'    GetSysStatus = Decode_SysStatus(temp_str)
'
'    GetSysStatus = temp_str
'End Function


Function Decode_SysStatus(stat_str As String) As String
  On Error Resume Next ' some cwBut_BasicCMD(index) are removed
    Dim retVal As String
    Dim token_pos As Long
    Dim bits_str As String
    Dim sys_stat_int As Long
    Dim bit_N As Long
'//********** sys_status definitions ********************** Status byte for system
'//#define BIT_ERROR_EQ_1               Bit_0  //Any error - set this bit
'#define BIT_FLOAT_UP_EQ_0_DOWN_EQ_1  Bit_0  //float: bit=0, VERTICALLY docked: bit=1
'#define BIT_OUTSIDE_GO_NOGO_EQ_1     Bit_1  //Inside Go-NoGo - bit=0; outside Go-NoGo - set this bit
'#define BIT_LMsafety_EQ_0            Bit_2  //if excessive oscillation, turn LM off and set this bit
'#define BIT_LM_on_EQ_0               Bit_3  //LM are enabled:- bit=0
'#define BIT_DITHER_on_EQ_0           Bit_4  //dither enabled - bit=0
'#define BIT_FLOAT_EQ_0_DOCK_EQ_1     Bit_5  //Horizontally free: bit=0, HELD = SOFT-docked: bit=1
'#define BIT_TestInProgress_EQ_1      Bit_6  // "Data collection started: " or "3Hz test" = set bit 6 on
'#define BIT_FF_Disabled_EQ_1         Bit_7  //FF enabled: - bit=0
    If Len(stat_str) > 2 Then 'front port
        token_pos = InStr(stat_str, ": ")
        If ((token_pos <> 0) And (Len(stat_str) >= token_pos + 2)) Then
            retVal = Mid(stat_str, token_pos + 2, 2) '"8e"
        Else
            Decode_SysStatus = ""
            Exit Function
        End If
    Else 'rear port
        retVal = stat_str '"8e"
    End If
    sys_stat_int = Val("&H" & retVal)
    bits_str = BIN(CByte(sys_stat_int))
    txtStatus.Text = "Status 0x" & retVal & "; " & CStr(sys_stat_int) & "D; b" & bits_str
    Init_sys_data = True
    For bit_N = 0 To 7
'        If (sys_stat_int And (2 ^ bit_N)) Then
'            cw_Status_Bit_Indicator(bit_N).Value = True
'        Else
'            cw_Status_Bit_Indicator(bit_N).Value = False
'        End If
        If (Mid(bits_str, 16 - 2 * bit_N, 1)) = "1" Then
'-!-            cwBut_BasicCMD(bit_N).Value = False '-!- need correctly show switch but dock - float status does not change at once - it takes time.
            cw_Status_Bit_Indicator(bit_N).Value = True
        Else
'-!-        cwBut_BasicCMD(bit_N).Value = True '-!- need correctly show switch but dock - float status does not change at once - it takes time.
'-!- So switch immediately returns to existing position
            cw_Status_Bit_Indicator(bit_N).Value = False
        End If
    Next bit_N
    Init_sys_data = False
    Decode_SysStatus = retVal
End Function

'====================================  E N D   O F  =========================================================================
''''''''''''''''''''''''''''''''''''''' SYSTEM ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


'============================================================================================================================
'''''''''''''''''''''''''''''''''''''' SYSTEM GAINS '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'============================================================================================================================

Function create_test_brif(test As Integer) As String

If test = 1 Then
    create_test_brif = _
"Prox:  Iso#1 Iso#2 Iso#3  Lin_Motor_FB_Axes: SafetyOFF TooLowInpPSI" + vbCrLf + _
"Vrt um -7746 -7746 -7746  FF_is ON  X- Y-          tZ  Iso1 Iso2 Iso3 Iso4 Inp" + vbCrLf + _
"Hrz um +6937 +6937 -6934 Excit=NONE   MotPWR   0%  PSI  0.6  0.6  0.5  0.5  0.7" + vbCrLf + _
"                                 BncO=Zpos_mm       mA  8.0  8.0  8.0  8.0" + vbCrLf
End If

If test = 2 Then
    create_test_brif = _
"Prox:  Iso#1 Iso#2 Iso#3  Lin_Motor_FB_Axes: ACTIVE      ~~floating~~ OK" + vbCrLf + _
"Vrt um +9    +9    +7     FF_is ON                 tZ  Iso1 Iso2 Iso3 Iso4 Inp" + vbCrLf + _
"Hrz um -7    +14   -6    Excit=Pulse  MotPWR 100%  PSI 11.3 12.0 11.2 11.6 36.3" + vbCrLf + _
"                                 BncO=GATZ          mA 74.1 75.6 78.9 77.5" + vbCrLf
End If


If test = 3 Then
    create_test_brif = _
"Prox:  Iso#1 Iso#2 Iso#3  Lin_Motor_FB_Axes: SafetyOFF TooLowInpPSI" + vbCrLf + _
"Vrt um -46   -29   +71    FF_is ON  X- Y- Z- tX tY tZ    Rel_dB = +21.31" + vbCrLf + _
"Hrz um +37   +6    -93   Excit=Sine   MotPWR   0%        Rel_deg= +90.5" + vbCrLf + _
"                                 BncO=Zpos_mm            Bnc1=p2Z_10V/mm" + vbCrLf
End If

End Function

Sub read_RT_diagnostic()
' decodes 4 lines from the 'brif' command
'Prox:  Iso#1 Iso#2 Iso#3  Lin_Motor_FB_Axes: SafetyOFF TooLowInpPSI
'Vrt um -7746 -7746 -7746  FF_is ON  X- Y-          tZ  Iso1 Iso2 Iso3 Iso4 Inp
'Hrz um +6937 +6937 -6934 Excit=NONE   MotPWR   0%  PSI  0.6  0.6  0.5  0.5  0.7
'                                 BncO=Zpos_mm       mA  8.0  8.0  8.0  8.0

' or when floating normally
'Prox:  Iso#1 Iso#2 Iso#3  Lin_Motor_FB_Axes: ACTIVE      ~~floating~~ OK
'Vrt um +9    +9    +7     FF_is ON                 tZ  Iso1 Iso2 Iso3 Iso4 Inp
'Hrz um -7    +14   -6    Excit=Pulse  MotPWR 100%  PSI 11.3 12.0 11.2 11.6 36.3
'                                 BncO=GATZ          mA 74.1 75.6 78.9 77.5

' or when excitation is working
'Prox:  Iso#1 Iso#2 Iso#3  Lin_Motor_FB_Axes: SafetyOFF TooLowInpPSI
'Vrt um -7746 -2323 -3333  FF_is ON  X- Y- Z- tX tY tZ    Rel_dB = +21.31
'Hrz um +6937 +6937 -6933 Excit=Sine   MotPWR   0%        Rel_deg= +90.5
'                                 BncO=Zpos_mm            Bnc1=p2Z_10V/mm
On Error Resume Next ' it often cannot decode - part of strings are missing
    Dim ch As Long
    Dim tmp_GetSysData As Boolean
    Dim token_pos As Long
    Dim brif_msg As String
    Dim RCIcommand As String

    Dim tmpDataStr As String
    Dim tmp_str As String
    Dim separator As Long   ' position of symbol ','
    Dim str_len As Long
    Dim excitation_is_working As Boolean
    RCIcommand = "brif"
    cmd_sent = False
    waitresponse = True

    If ConnectionType = ConnDEMO Then
        If Excitation_status = True Then
            brif_msg = create_test_brif(3)
        Else
            brif_msg = create_test_brif(2)
        End If
    Else
        GUI_doing_command = True
        brif_msg = Analyzer.GetSend(RCIcommand, True)
    End If
    If (InStr(brif_msg, "meout")) Then Exit Sub '"Timeout" sometimes seen as "imeout"

    token_pos = InStr(brif_msg, "dB")
    If token_pos <> 0 Then
        excitation_is_working = True
        ' magnitude
        Frame_RT_Mag_Ph.Visible = True
        tmpDataStr = Mid(brif_msg, token_pos + 4, 7)
        If (IsNumber(tmpDataStr)) Then lbl_Mag_dB.Caption = Trim(tmpDataStr)
        ' degrees
        token_pos = InStr(brif_msg, "deg=")
        If token_pos <> 0 Then
            tmpDataStr = Mid(brif_msg, token_pos + 4, 6)
            If (IsNumber(tmpDataStr)) Then lbl_Phase_diff.Caption = Trim(tmpDataStr)
        End If
        token_pos = InStr(brif_msg, "Bnc0=")
        If token_pos = 0 Then token_pos = InStr(brif_msg, "BncO=")
        If token_pos <> 0 Then
            separator = InStr(token_pos + 5, brif_msg, "Bnc1=")
            If (separator <> 0) Then
                tmp_str = Mid(brif_msg, token_pos + 5, separator - token_pos - 5)
                lbl_BNC_0.Caption = Trim(tmp_str)
                'token_pos = InStr(separator, brif_msg, vbCr)
                tmp_str = Mid(brif_msg, separator + 5) ', token_pos)
                lbl_BNC_1.Caption = Trim(tmp_str)
            End If
        End If
    Else
        token_pos = InStr(brif_msg, "Bnc0=")
        If token_pos = 0 Then token_pos = InStr(brif_msg, "BncO=")
        If token_pos <> 0 Then
            separator = InStr(token_pos + 5, brif_msg, "mA")
            tmp_str = Mid(brif_msg, token_pos + 5, separator - token_pos - 5)
            lbl_BNC_0.Caption = Trim(tmp_str)
        End If
        excitation_is_working = False
        Frame_RT_Mag_Ph.Visible = False
    End If
    tmp_GetSysData = Init_sys_data
    Init_sys_data = True
    cwBut_Excitation.Value = excitation_is_working ' update excitation button without sending command
    Init_sys_data = tmp_GetSysData


    str_len = InStr(brif_msg, vbCr) ' find end of the FIRST string
    separator = InStr(brif_msg, "xes:") 'Lin_Motor_FB_Axes:
    If separator <> 0 Then
        tmpDataStr = Mid(brif_msg, separator + 4, str_len - separator - 4)
        tmp_str = LCase(tmpDataStr)
        If (InStr(tmp_str, "safetyoff")) Then
            cw_Status_Bit_Indicator(3).OnText = "Velocity FB OFF by Safety" ' reversed to be in sync with "s" command - zero means active
            cw_Status_Bit_Indicator(3).Value = True ' reversed to be in sync with "s" command
        Else
          If InStr(tmp_str, "pass") Then
            cw_Status_Bit_Indicator(3).Value = True ' reversed to be in sync with "s" command
            cw_Status_Bit_Indicator(3).OnText = "Velocity FB OFF" ' reversed to be in sync with "s" command
          Else
            If InStr(tmp_str, "acti") Then
            cw_Status_Bit_Indicator(3).OffText = "Velocity FB ON"
              cw_Status_Bit_Indicator(3).Value = False
            Else
              cw_Status_Bit_Indicator(3).OnText = "Vel FB unknown"
              cw_Status_Bit_Indicator(3).OffText = "Vel FB unknown"
            End If
          End If
        End If
        token_pos = InStr(3, tmpDataStr, " ") 'find next space
        tmp_str = Mid(tmpDataStr, token_pos + 1, str_len - token_pos - 1)
        lbl_RTstatus.Caption = Trim(tmp_str)
    End If
    brif_msg = Mid(brif_msg, str_len + 2) ' cut first line
' second line
    str_len = InStr(brif_msg, vbCr) ' find end of string
    ' vertical proximity readings
    separator = InStr(brif_msg, "um") 'Vert um
    tmpDataStr = Mid(brif_msg, separator + 3, 5) ' separate first Vprx -7746
    If (IsNumber(tmpDataStr)) Then lbl_ProxSensor(3).Caption = Trim(tmpDataStr)
    tmpDataStr = Mid(brif_msg, 14, 5)  ' separate second Vprx -2323
    If (IsNumber(tmpDataStr)) Then lbl_ProxSensor(4).Caption = Trim(tmpDataStr)
    tmpDataStr = Mid(brif_msg, 20, 5)  ' separate third Vprx -3333
    If (IsNumber(tmpDataStr)) Then lbl_ProxSensor(5).Caption = Trim(tmpDataStr)

    separator = InStr(brif_msg, "FF_is")
    tmp_str = Mid(brif_msg, separator + 6, 3) ' separate "FF_is ON" or "FF_is OFF"
    If (InStr(tmp_str, "N")) Then
        cw_Status_Bit_Indicator(7).Value = False ' reversed to be in sync with "s" command - zero means active
    Else
        cw_Status_Bit_Indicator(7).Value = True ' reversed to be in sync with "s" command - zero means active
    End If
    ' ! VB6, not DLL or COM returns incomplete string and here str_len = 0
    ' VB sends "Run-time rror 5: Invalid procedure call or argument"
    tmp_str = Mid(brif_msg, 37, str_len - 37)  ' cut second line up to axes status
    'axes status
    If (InStr(tmp_str, "X-")) Then
        cwAxisFBstatus(0).Value = True
    Else
        cwAxisFBstatus(0).Value = False
    End If
    If (InStr(tmp_str, "Y-")) Then
        cwAxisFBstatus(1).Value = True
    Else
        cwAxisFBstatus(1).Value = False
    End If
    If (InStr(tmp_str, "Z-")) Then
        cwAxisFBstatus(2).Value = True
    Else
        cwAxisFBstatus(2).Value = False
    End If
    If (InStr(tmp_str, "tX")) Then
        cwAxisFBstatus(3).Value = True
    Else
        cwAxisFBstatus(3).Value = False
    End If
    If (InStr(tmp_str, "tY")) Then
        cwAxisFBstatus(4).Value = True
    Else
        cwAxisFBstatus(4).Value = False
    End If
    If (InStr(tmp_str, "tZ")) Then
        cwAxisFBstatus(5).Value = True
    Else
        cwAxisFBstatus(5).Value = False
    End If
    brif_msg = Mid(brif_msg, str_len + 2) ' cut second line
'third line
    str_len = InStr(brif_msg, vbCr) ' find end of string
    ' horizontal proximity readings Hrz um -7    +14   -6
    separator = InStr(brif_msg, "um") 'Hrz um
    tmpDataStr = Mid(brif_msg, separator + 3, 5) ' separate first Vprx -7
    If (IsNumber(tmpDataStr)) Then lbl_ProxSensor(0).Caption = Trim(tmpDataStr)
    tmpDataStr = Mid(brif_msg, 14, 5)  ' separate second Vprx +14
    If (IsNumber(tmpDataStr)) Then lbl_ProxSensor(1).Caption = Trim(tmpDataStr)
    tmpDataStr = Mid(brif_msg, 20, 5)  ' separate third Vprx -6
    If (IsNumber(tmpDataStr)) Then lbl_ProxSensor(2).Caption = Trim(tmpDataStr)
    'motor power
    token_pos = InStr(3, brif_msg, "%") 'find percent char
    If token_pos <> 0 Then
        tmpDataStr = Mid(brif_msg, token_pos - 3, 4) ' power percent
        lbl_MotPower.Caption = tmpDataStr
    End If
    ' iso pressure (when excitation is disabled)
    If excitation_is_working = False Then
        token_pos = InStr(3, brif_msg, "PSI") 'PSI 50.2 53.3 40.3 42.7 75.9
        If (token_pos <> 0) Then 'airless systems do not have this reading
            lbl_PSI.Visible = True
            lbl_IsoPressure(0).Visible = True
            lbl_IsoPressure(1).Visible = True
            lbl_IsoPressure(2).Visible = True
            lbl_IsoPressure(3).Visible = True
            lbl_IsoPressure(4).Visible = True
            For ch = 0 To 4
                tmpDataStr = Mid(brif_msg, token_pos + 4 + ch * 5, 4)
                If (IsNumber(tmpDataStr)) Then lbl_IsoPressure(ch).Caption = (tmpDataStr)
            Next ch
        Else 'airless systems do not have this reading
            lbl_PSI.Visible = False
            lbl_IsoPressure(0).Visible = False
            lbl_IsoPressure(1).Visible = False
            lbl_IsoPressure(2).Visible = False
            lbl_IsoPressure(3).Visible = False
            lbl_IsoPressure(4).Visible = False
        End If
    Else 'excitation_is_working = True

    End If

    brif_msg = Mid(brif_msg, str_len + 2) ' cut third line
'forth line
    ' iso pressure (when excitation is disabled)
    If excitation_is_working = False Then
        token_pos = InStr(3, brif_msg, "mA") ' mA 95.1 89.6 96.3 93.4
        If (token_pos <> 0) Then 'airless systems do not have this reading
            lbl_mA.Visible = True
            lbl_ValveCurrent(0).Visible = True
            lbl_ValveCurrent(1).Visible = True
            lbl_ValveCurrent(2).Visible = True
            lbl_ValveCurrent(3).Visible = True
            For ch = 0 To 3
                tmpDataStr = Mid(brif_msg, token_pos + 3 + ch * 5, 4)
                If (IsNumber(tmpDataStr)) Then lbl_ValveCurrent(ch).Caption = (tmpDataStr)
            Next ch
        Else 'airless systems do not have this reading
            lbl_mA.Visible = False
            lbl_ValveCurrent(0).Visible = False
            lbl_ValveCurrent(1).Visible = False
            lbl_ValveCurrent(2).Visible = False
            lbl_ValveCurrent(3).Visible = False
        End If
    Else 'excitation_is_working = True
        ' there is frame on the top of these controls
    End If
'    Call ShowSysStatus
End Sub

Sub Read_Prox_Sensors()
    Dim scan_line As Long
    Dim load_sens As String
    Dim RCIcommand As String

    RCIcommand = "relm"
    cmd_sent = False
    waitresponse = True

    load_sens = Analyzer.GetSend(RCIcommand, True)

    scan_line = ExtractLoadDataFromScanLine(load_sens)

    Call Show_load_values
End Sub

Function Read_Alarm_Status() As Long
On Error GoTo read_err
' * RCI command: "exst"
' * Emits extended status in ascii-hex, to stdout.
' * Output: 11111111,22222222,33333333,4
' *      where   11111111 is 32-bit M3v.Cmd_Word
' *              22222222 is 32-bit M3v.AlarmWord
' *              33333333 is 32-bit Prod_Test_BITS
' *              4        is M3v.Diagnostic_running (non-zero when test selected)
' updates M3v_Diagnostic_running =
'     NO_TESTS                   = 0 // no test
'    ,LOGICAL_MotSens_TEST_Table = 1 // motor-sensor test: inject into one LOGICAL output control_signal[], see InGeo[] or InProx[],  PTR_FLOAT
'    ,DAC_TO_ADC_loopback_TEST   = 2 // LOOP-BACK DB-37 SOCKETS REQUIRED! DAC-to-ADC loop back test, index the same, i.e. drive DAC[0], look at ADC[0], PTR_INT16
'    ,PHYSICAL_PAIR_TEST         = 3 // DAC[d_index] to ADC[a_index] index to different index, i.e. drive DAC[6], look at ADC[15], PTR_INT16
'    ,MotorSensorTEST_List       = 4 // for "test pair" command, output as a list
'    ,EXCITATION_TO_INT16 = 5

' * RCI command: "exst!" sets M3v.measurement_done = TRUE
' * RCI command: "exst=0x12345678"   sets M3v.AlarmWord = 0x12345678
    Dim alarm_status As String
    Dim separator As Long

    Dim tmpDataStr As String
    Dim binFormat As String

        Read_Alarm_Status = 0
        tmpDataStr = "exst"
        cmd_sent = False
        waitresponse = True
        alarm_status = Analyzer.GetSend(tmpDataStr, True) '"0x120F0008,0x00000000,0x00000000,0>~Doing CMD"
        separator = InStr(alarm_status, ",") - 1
        If (separator > 0) Then
            tmpDataStr = Left(alarm_status, separator) ' 32-bit M3v.Cmd_Word, "0x120F0008"
            tmpDataStr = "&H" & Mid(tmpDataStr, 3) ' 32-bit M3v.Cmd_Word, "120F0008"
            If (IsNumeric(tmpDataStr)) Then ' CAN USE IsNumeric, because 'tmpDataStr' is integer
                M3v_Cmd_Word = Val(tmpDataStr) '
            End If
            tmpDataStr = Mid(alarm_status, separator + 4)  ' remove "0x" and the piece of string from which we've converted information
        Else
            Read_Alarm_Status = Read_Alarm_Status - 1 ' adding -1 means error
        End If

        alarm_status = tmpDataStr
        separator = InStr(alarm_status, ",") - 1
        If (separator > 0) Then
            tmpDataStr = "&H" & Left(alarm_status, separator) ' 32-bit M3v.AlarmWord
            If (IsNumeric(tmpDataStr)) Then ' CAN USE IsNumeric, because 'tmpDataStr' is integer
                AlarmWord = Val(tmpDataStr) 'As Long
                self_test_word = AlarmWord
                binFormat = BIN(self_test_word And &HFF)
                self_test_word = self_test_word / 256
                binFormat = BIN(self_test_word And &HFF) + binFormat
                self_test_word = self_test_word / 256
                binFormat = BIN(self_test_word And &HFF) + binFormat
                self_test_word = self_test_word / 256
                binFormat = BIN(self_test_word And &HFF) + binFormat
                txtStatus.Text = "AlarmWord = " & tmpDataStr & " " & binFormat
'                Call Show_Alarm_flags
'                If AlarmWord = 0 Then
'                    cwIndicatorAlarm.Value = False
'                Else
'                    cwIndicatorAlarm.Value = True
'                End If
            End If
            tmpDataStr = Mid(alarm_status, separator + 4) ' remove "0x" and the piece of string from which we've converted information
        Else
            Read_Alarm_Status = Read_Alarm_Status - 1 ' adding -1 means error
        End If

        alarm_status = tmpDataStr
        separator = InStr(alarm_status, ",") - 1
        If (separator > 0) Then
            tmpDataStr = "&H" & Left(alarm_status, separator) ' 32-bit Prod_Test_BITS
            If (IsNumeric(tmpDataStr)) Then ' CAN USE IsNumeric, because 'tmpDataStr' is integer
                self_test_word = Val(tmpDataStr) 'As Long
                tmpDataStr = Mid(alarm_status, separator + 2) ' remove the piece of string from which we've converted information
            Else
                self_test_word = 0
            End If
        Else
            Read_Alarm_Status = Read_Alarm_Status - 1 ' adding -1 means error
        End If

read_err:
    Init_sys_data = False
End Function


Private Sub cmdRefresh_Click()
    Call Refresh_gains
    ID_valid = get_ID()
End Sub

Sub Refresh_gains()
On Error Resume Next
    Dim i As Long
    Dim sep_pos As Long
    Dim controller_param As String
    Dim paramValue As Double
    Dim ControllerCommand As String
    Dim max_indx As Long
    Dim scan_line As Long
    Dim contr_response As String

    Init_sys_data = True
    ControllerCommand = "echo>enab" ' disable verbose
    cmd_sent = False
    waitresponse = True
    contr_response = Analyzer.GetSend(ControllerCommand, True)

    ControllerCommand = "relm"
    cmd_sent = False
    waitresponse = True
    contr_response = Analyzer.GetSend(ControllerCommand, True)
    scan_line = ExtractLoadDataFromScanLine(contr_response)
    Call Show_load_values

' read FB gains
    max_indx = cwNumFBgain.UBound
    For i = 0 To max_indx
      If (cwNumFBgain(i).Visible = True) Then
        paramValue = RectifyValueResponse(cwNumFBgain(i).Tag)
        If paramValue <> WRONG_VALUE Then
            cwNumFBgain(i).Value = paramValue
        End If
      End If 'if visible
    Next i

'enable.Tag
    max_indx = cwBut_AxisEn.UBound
    For i = 0 To max_indx
      If (cwBut_AxisEn(i).Visible = True) Then
        controller_param = Analyzer.GetSend(cwBut_AxisEn(i).Tag, True) 'loop_FB_xx
        sep_pos = InStr(controller_param, ">")
        If (sep_pos <> 0) Then
            If (LCase(Mid(controller_param, sep_pos + 1, 4)) = "acti") Then 'acti
               cwBut_AxisEn(i).Value = True
            End If
            If (LCase(Mid(controller_param, sep_pos + 1, 4)) = "pass") Then 'pass
               cwBut_AxisEn(i).Value = False
            End If
        End If
      End If 'if visible
    Next i

    ' Auxilary Pneumatic damping
    controller_param = Analyzer.GetSend(cwBut_AuxDamping.Tag, True) 'loop_FBD
    sep_pos = InStr(controller_param, ">")
    If (sep_pos <> 0) Then
        If (LCase(Mid(controller_param, sep_pos + 1, 4)) = "acti") Then 'acti
           cwBut_AuxDamping.Value = True
        End If
        If (LCase(Mid(controller_param, sep_pos + 1, 4)) = "pass") Then 'pass
           cwBut_AuxDamping.Value = False
        End If
    End If
    ' Soft dock
    controller_param = Analyzer.GetSend(cwBut_SoftDock.Tag, True) 'horz
    sep_pos = InStr(controller_param, ">")
    If (sep_pos <> 0) Then
        If (LCase(Mid(controller_param, sep_pos + 1, 4)) = "hold") Then 'soft dock
           cwBut_SoftDock.Value = True
        End If
        If (LCase(Mid(controller_param, sep_pos + 1, 4)) = "free") Then 'payload is free
           cwBut_SoftDock.Value = False
        End If
    End If
    ' FF all
    controller_param = Analyzer.GetSend(cwBut_FF_ALL.Tag, True) 'loop_ffa
    sep_pos = InStr(controller_param, ">")
    If (sep_pos <> 0) Then
        If (LCase(Mid(controller_param, sep_pos + 1, 4)) = "acti") Then 'FF is on
           cwBut_FF_ALL.Value = True
        End If
        If (LCase(Mid(controller_param, sep_pos + 1, 4)) = "pass") Then 'FF is off
           cwBut_FF_ALL.Value = False
        End If
    End If

'    ControllerCommand = "ison"
'    cmd_sent = False
'    waitresponse = True
'    load_sens = Analyzer.GetSend(ControllerCommand, True)
'    If IsNumber(load_sens) Then
'        cwNumOfISO.Value = CInt(load_sens)
'    Else
'        cwNumOfISO.Value = 4
'    End If
'    ShowHideIso (cwNumOfISO.Value)
    Init_sys_data = False
End Sub

'Private Sub cwBut_AxisEn_ValueChanged(Index As Integer, ByVal Value As Boolean)
'    If Init_sys_data = True Then Exit Sub 'avoid firing control event subroutine during initialization when values assigned
'    If cwBut_AxisEn(Index).Value = True Then
'        Call Analyzer.GetSend(cwBut_AxisEn(Index).Tag & ">active", True)
'    Else
'        Call Analyzer.GetSend(cwBut_AxisEn(Index).Tag & ">passive", True)
'    End If
'End Sub

Private Sub cwBut_AxisEn_ValueChanged(Index As Integer, ByVal Value As Boolean)
    If Init_sys_data = True Then Exit Sub 'avoid firing control event subroutine during initialization when values assigned
    If cwBut_AxisEn(Index).Value = True Then
        Call Analyzer.GetSend(cwBut_AxisEn(Index).Tag & ">active", True)
    Else
        Call Analyzer.GetSend(cwBut_AxisEn(Index).Tag & ">passive", True)
    End If
    Call read_RT_diagnostic
End Sub

Private Sub CmdRestorefromFlash_Click()
    cmd_sent = False
    Call GetSend("rsto", True)      'restore params from flash cmd
    Call GetSysData
End Sub

Private Sub cmdLoadDefaults_Click(Index As Integer)
    cmd_sent = False
    Call Analyzer.GetSend("dflt", True)      'default cmd
End Sub
Private Sub cmdSaveParamsIntoFLASH_Click(Index As Integer)
    cmd_sent = False
    Call Analyzer.GetSend("save", True)      'save cmd
End Sub

'====================================  E N D   O F  =========================================================================
''''''''''''''''''''''''''''''''''''''' GAINS ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''




'============================================================================================================================
''''''''''''''''''''''''''''''''''''''  DIAGNOSTIC  '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'============================================================================================================================

Private Sub cwBut_Excitation_ValueChanged(ByVal Value As Boolean)
'// "excitx=start"  "excitx=stop" excitation.
    If Init_sys_data = False Then 'avoid firing control event subroutine during initialization when values assigned
        cmd_sent = False
        If cwBut_Excitation.Value = False Then
            Call Analyzer.GetSend("excit>stop", True) 'none (cancel)
            Excitation_status = False  ' global excitation status for monitoring and main forms
        Else
            If optExitation(1).Value = True Then
                Call Analyzer.GetSend("freq=0", True) 'enable noise excitation
            End If
            Call Analyzer.GetSend("excit>start", True) 'enable excitation
            Excitation_status = True  ' global excitation status for monitoring and main forms
        End If

        'update control on frmMonitoring form
        Init_sys_data = True
        If IsFormLoaded("frmMonitoring") Then frmMonitoring.cwBut_Excitation.Value = Excitation_status
        Init_sys_data = False
    End If
    Call read_RT_diagnostic
End Sub

Private Sub cmb_ExcitAxis_Click()
  Dim ControllerCommand As String
    If Init_sys_data = True Then Exit Sub
    ExcitationAxis = cmb_ExcitAxis.ListIndex
    ControllerCommand = "drvo=" & CStr(ExcitationAxis)
    cmd_sent = False
    waitresponse = True
    response = Analyzer.GetSend(ControllerCommand, True)
    'update control on monitoring form
    Init_sys_data = True
    If IsFormLoaded("frmMonitoring") Then frmMonitoring.cmb_ExcitAxis.ListIndex = ExcitationAxis
    Init_sys_data = False
End Sub


Private Sub CmbBNC1_Click()
  Dim command As String
'  Dim BNC_index As Long
    If Init_sys_data = True Then Exit Sub
    BNC_1_index = CmbBNC1.ListIndex
    command = CmbBNC1.Tag & CStr(BNC_1_index)
    cmd_sent = False
    waitresponse = True
    response = GetSend(command, True)
    command = Find_Index_Match_with_BNC_SLIDER(1, BNC_1_index)
End Sub


Private Sub cwBNCsliders_ValueChanged(ByVal Value As Boolean)
    FrameBNC0.Visible = cwBNCsliders.Value
    FrameBNC1.Visible = cwBNCsliders.Value
    If Value = True Then
        FrameBNC0.ZOrder (vbBringToFront)
        FrameBNC1.ZOrder (vbBringToFront)
    Else
        FrameBNC0.ZOrder (vbSendToBack)
        FrameBNC1.ZOrder (vbSendToBack)
    End If
End Sub

Private Sub cwBNC0select_PointerValueChanged(ByVal Pointer As Long, Value As Variant)
    Dim cmd As String
    If Init_sys_data = True Then Exit Sub
    BNC_0_index = Value
    retVal = cwBNC0select.axis.ValuePairs(cwBNC0select.ValuePairIndex).Name
    If cwBNC0select.axis.ValuePairs(cwBNC0select.ValuePairIndex).Value >= 0 Then ' -other-- has ValuePairs.Value -1, do not send command
        cmd = "bncd0=" & CStr(Value)
        Call Analyzer.GetSend(cmd, True)
        Init_sys_data = True ' prevent sending second command when change drop down control
          CmbBNC0.ListIndex = BNC_0_index
        Init_sys_data = False
    End If
End Sub

Private Sub cwBNC1select_PointerValueChanged(ByVal Pointer As Long, Value As Variant)
    Dim cmd As String
    If Init_sys_data = True Then Exit Sub
    BNC_1_index = Value
    retVal = cwBNC1select.axis.ValuePairs(cwBNC1select.ValuePairIndex).Name
    If cwBNC1select.axis.ValuePairs(cwBNC1select.ValuePairIndex).Value >= 0 Then ' -other-- has ValuePairs.Value -1, do not send command
        cmd = "bncd1=" & CStr(Value)
        Call Analyzer.GetSend(cmd, True)
        Init_sys_data = True ' prevent sending second command when change drop down control
          CmbBNC1.ListIndex = BNC_1_index
        Init_sys_data = False
    End If
End Sub

Private Sub cwNumPulseNegSlope_ValueChanged(Value As Variant, PreviousValue As Variant, ByVal OutOfRange As Boolean)
'//   Set: prdn=xxx.xx, Get: prdn?, where xxx.xx - float
    Dim controller_command As String
    If Init_sys_data = False Then 'avoid firing control event subroutine during initialization when values assigned
        TestPulseNegSlope = cwNumPulseNegSlope.Value
        controller_command = "prdn=" & TestPulseNegSlope
        cmd_sent = False
        waitresponse = True
        response = Analyzer.GetSend(controller_command, True)
    End If
End Sub

Private Sub cwNumPulsePosSlope_ValueChanged(Value As Variant, PreviousValue As Variant, ByVal OutOfRange As Boolean)
'//   Set: prup=xxx.xx, Get: prup?, where xxx.xx - float
    Dim controller_command As String
    If Init_sys_data = False Then 'avoid firing control event subroutine during initialization when values assigned
        TestPulsePosSlope = cwNumPulsePosSlope.Value
        controller_command = "prup=" & TestPulsePosSlope
        cmd_sent = False
        waitresponse = True
        response = Analyzer.GetSend(controller_command, True)
    End If
End Sub

Function Find_Index_Match_with_BNC_SLIDER(BNC_num As Long, paramValue As Long) As Boolean
    Dim indx As Long
    Dim tmp_init_sys_data As Boolean
    tmp_init_sys_data = Init_sys_data
    Find_Index_Match_with_BNC_SLIDER = False
    ' find if index matches with slider control
    For indx = 0 To UBound(BNC_ptr)
        If BNC_ptr(indx) = paramValue Then
            Init_sys_data = True ' prevent sending command bncdX, we are just searching
            If (BNC_num = 0) Then
                cwBNC0select.ValuePairIndex = CInt(indx + 1)
            Else
                cwBNC1select.ValuePairIndex = CInt(indx + 1)
            End If
            Find_Index_Match_with_BNC_SLIDER = True
            Init_sys_data = tmp_init_sys_data
            Exit Function
        End If
    Next indx
    If (BNC_num = 0) Then
        cwBNC0select.ValuePairIndex = cwBNC0select.axis.Maximum ' --other--
    Else
        cwBNC1select.ValuePairIndex = cwBNC1select.axis.Maximum ' --other--
    End If
    Init_sys_data = tmp_init_sys_data
End Function

Function RectifyValueResponse(cmd As String) As Double
    Dim temp_str As String
    Dim retVal As String
    Dim token_pos As Long
    RectifyValueResponse = WRONG_VALUE ' some strange unlikely response
    temp_str = GetSend(cmd, True) 'command string
' retval = GetSend("prup")  returns "=+3.000" because echo removed
' retval = GetSend("prup?") returns "prup=+3.000" because echo WHICH IS "prup?" is not removed
'">~ER drvo Param, Error Code =2  drvo=0"
' 92.00,  11.80, -64.60  #
'->freq
'
'freq = 90#
'>~Doing CMD
' 90.00,  11.90, -64.70  #

    token_pos = InStr(temp_str, ">~U") ' ">~Unrecognized Cmd, Error Code =1"
    If (token_pos <> 0) Then Exit Function
' error in older firmware response on "drvo?" command
    token_pos = InStr(temp_str, ">~E") ' ">~ER drvo Param, Error Code =2  drvo=0"
    If (token_pos <> 0) Then
        token_pos = InStr(temp_str, "drvo=") ' ">~ER drvo Param, Error Code =2  drvo=0"
        If (token_pos <> 0) Then
            temp_str = Mid(temp_str, token_pos)
        Else
            Exit Function
        End If
    End If
    token_pos = InStr(temp_str, ">~D") ' "freq=+2.5\r\n>~Doing CMD\r\n
    If (token_pos <> 0) Then
        temp_str = Left(temp_str, token_pos - 1) ' "freq=+2.5\r\n
    End If
    token_pos = InStr(temp_str, "=") 'find value which starts after "="
    If (token_pos <> 0) Then
        retVal = Mid(temp_str, token_pos + 1)
        token_pos = InStr(retVal, "//") ''if verbose response
        If (token_pos <> 0) Then
            retVal = Left(retVal, token_pos - 1)
        End If
        If IsNumber(retVal) Then
            RectifyValueResponse = Val(retVal)
        End If
    End If
End Function

Sub Get_diag_params()
    Dim temp_str As String
    Dim retVal As String
    Dim paramValue As Double

    Dim indx As Long
    Dim token_pos As Long
    Dim command As String

    Init_sys_data = True

    paramValue = RectifyValueResponse("freq?") '"+5.000 // excitation frequency" <<<<< with verbose response
    If (paramValue <> WRONG_VALUE) Then
        TestFrequency = paramValue
        If IsFormLoaded("frmMonitoring") Then frmMonitoring.cwNumExcitFreq.Value = TestFrequency
        Analyzer.cwNumExcitFreq(0).Value = TestFrequency
        Analyzer.cwNumExcitFreq(1).Value = TestFrequency
    End If

    paramValue = RectifyValueResponse("ampl?") '"+1.20E+03 // excitation amplitude" <<<<< with verbose response
    If (paramValue <> WRONG_VALUE) Then
        TestGain = paramValue
        If IsFormLoaded("frmMonitoring") Then frmMonitoring.cwNumExcitAmpl.Value = TestGain
        Analyzer.cwNumExcitAmpl(0).Value = TestGain
        Analyzer.cwNumExcitAmpl(1).Value = TestGain
    End If

    paramValue = RectifyValueResponse(cwNumPulsePosSlope.Tag + "?")
    '"prup" GetSend returns "=+3.000" because echo removed
    '"prup?" GetSend returns "prup=+3.000" because echo WHICH IS "prup?" is not removed
    If (paramValue <> WRONG_VALUE) Then
        TestPulsePosSlope = paramValue ' "+3.0 // Excitation pulse \ slope" <<<<< with verbose response
        cwNumPulsePosSlope.Value = TestPulsePosSlope
    End If

    '"prdn" GetSend returns "=+3.000" because echo removed
    '"prdn?" GetSend returns "prdn=+3.000" because echo WHICH IS "prdn?" is not removed
    paramValue = RectifyValueResponse(cwNumPulseNegSlope.Tag + "?")
    If (paramValue <> WRONG_VALUE) Then
        TestPulseNegSlope = paramValue 'USE Val() "+3.0 // Excitation pulse \ slope" <<<<< with verbose response
        cwNumPulseNegSlope.Value = TestPulseNegSlope
    End If

'    temp_str = GetSend("vbxo1", True)
'    retval = Mid(temp_str, 7)
'    If IsNumber(retval) Then
'        frmMonitoring.CwOffsetCurrent(0).Value = Val(retval)
'    End If
'   temp_str = GetSend("vbxo2", True)
'    retval = Mid(temp_str, 7)
'    If IsNumber(retval) Then
'        frmMonitoring.CwOffsetCurrent(1).Value = Val(retval)
'    End If
'
'   temp_str = GetSend("vbxo3", True)
'    retval = Mid(temp_str, 7)
'    If IsNumber(retval) Then
'        frmMonitoring.CwOffsetCurrent(2).Value = Val(retval)
'    End If
'
'   temp_str = GetSend("vbxo4", True)
'    retval = Mid(temp_str, 7)
'    If IsNumber(retval) Then
'        frmMonitoring.CwOffsetCurrent(3).Value = Val(retval)
'    End If


    paramValue = RectifyValueResponse("drvo?")
    If (paramValue <> WRONG_VALUE) Then
        ExcitationAxis = CInt(paramValue)
        If (ExcitationAxis < cmb_ExcitAxis.ListCount) Then ' if accidently connect to ED controller excitation index 19 but ListCount=12
            If IsFormLoaded("frmMonitoring") Then frmMonitoring.cmb_ExcitAxis.ListIndex = ExcitationAxis
            Analyzer.cmb_ExcitAxis.ListIndex = ExcitationAxis
        End If
    End If

    temp_str = GetSend("excit", True)
    retVal = LCase(Mid(temp_str, 7))
    If InStr(retVal, "star") Then
        Excitation_status = True  ' global excitation status for monitoring and main forms
    Else
        Excitation_status = False  ' global excitation status for monitoring and main forms
    End If
    cwBut_Excitation.Value = Excitation_status

    Init_sys_data = False
End Sub

Sub GetDiagData()
    Dim temp_str As String
    Dim retVal As String
    Dim monitor_stat As Boolean
    Dim cont_field_monitoring As Boolean
    Init_sys_data = True

    Wait_sec (0.1) 'let transmission, if any , finish

    temp_str = GetSend("freq?", 2)
    If temp_str = "Timeout" Then 'error, port buzy or something
        Call portproblems("getting sys data")
    End If
    retVal = Mid(temp_str, 6)
    If IsNumber(retVal) Then 'returned string "10.000", but for German/French it is not valid: "10,000" is ten Hertz
        'Locale_tmp_number = CSng(retval) 'CStr converted to value 10000 when German locale
        Locale_tmp_number = Val(retVal) 'Val converted as 10.000
        TestFrequency = Locale_tmp_number
        cwNumExcitFreq(0).Value = TestFrequency
        cwNumExcitFreq(1).Value = TestFrequency
    End If
    temp_str = GetSend("ampl?", 2) 'returned string "0.200", but for German/French it must be: "0,2"
    retVal = Mid(temp_str, 6)
    If IsNumber(retVal) Then
        TestGain = Val(retVal) 'Val ALWAYS converted using "."
        cwNumExcitAmpl(0).Value = TestGain 'variable converted as 0.2 or 0,2 depends on locale
        cwNumExcitAmpl(1).Value = TestGain
    End If
    temp_str = GetSend("excit?", 2)
    If InStr(temp_str, "star") Then ' firmware returns "excitx=start"
        Excitation_status = True  ' global excitation status for monitoring and main forms
    Else
        Excitation_status = False  ' global excitation status for monitoring and main forms
    End If
    Analyzer.cwBut_Excitation.Value = Excitation_status
    If IsFormLoaded("frmMonitoring") Then frmMonitoring.cwBut_Excitation.Value = Excitation_status

    Init_sys_data = False
End Sub

'Private Sub cmdGetLoadSens_Click()
'    Call Read_Load_Sensors
'End Sub

Function ExtractLoadDataFromScanLine(data_line As String) As Long
On Error GoTo format_error 'returns zero if OK or -1 if error
    Dim tmpDataStr As String
    Dim tmp_str As String
    Dim ch As Long
    Dim separator As Long   ' position of symbol ','
    Dim str_len As Long

'    separator = InStr(data_line, ",")
'    tmp_str = Left(data_line, separator - 1) ' index of line
'    If (IsNumber(tmp_str)) Then
'        ExtractLoadDataFromScanLine = Val(tmp_str)  'line count starts from 0000
'    Else
'        Exit Function
'    End If
'

process_data:
    For ch = 0 To MAX_Prox_Sensors - 1 'find separator ',' 6 times
        str_len = Len(data_line)
        If (str_len = 0) Then
            ExtractLoadDataFromScanLine = -9999 ' -9999 means error
            Exit Function
        Else
            OneScanProxValues(ch) = 0
        End If
        separator = InStr(data_line, ",")
        tmpDataStr = Left(data_line, separator - 1) ' data
        If (IsNumber(tmpDataStr)) Then
            OneScanProxValues(ch) = Val(tmpDataStr) '
        End If
        data_line = Right(data_line, str_len - separator)  ' remove the piece of string from which we've converted information
'        separator = InStr(data_line, ",")
'        tmpDataStr = Left(data_line, separator - 1) ' index of line
    Next ch
    Exit Function

format_error:
'    MsgBox ("Unrecognized data format")
End Function

Private Sub Show_load_values()
    Dim Y As Double
    Dim ch As Long

    For ch = 0 To MAX_Prox_Sensors - 1 'find separator ',' 10 times
        Y = OneScanProxValues(ch)
       lbl_ProxSensor(ch).Caption = Format(Y, "")
    Next ch
End Sub


'********************************************************************************************************************
Private Sub ReadChannelNamesFromController()
  On Error Resume Next
    Dim ch As Integer
    Dim resp As String
    Dim response As String
    Dim cmdStr As String
    Dim ch0_indx As Integer
    Dim spos As Long
    Dim dname_command As Boolean
    Dim NameOfVariable As String
    Dim StartTime As Single
    Dim Time_Now As Single

    StartTime = Timer
    response = GetSend("dnam0", True) '// "dNam###[?]" - will print the name of the variable
    If InStr(LCase(response), "unrec") Then
        dname_command = False
        response = GetSend("echo>verb", True)
        'preserve actual index / name of ch0
        cmdStr = "frame0?"
        ch0_indx = GetIndexOfChannel(0)
        For ch = 0 To UBound(ChName) - 1 'Ubaund(ChName) is 256
            ' "frameN=XX" - will replace pointer to the variable in frame "N"
            'cmd: frame0
            'resp: frame0=4 // SPI data Prx_1Z_um ->frame #0(0)
            cmdStr = "frame0=" + CStr(ch)
            response = GetSend(cmdStr, True)
            ' Thread.Sleep(5);
            cmdStr = "frame0?"
            'cmd: frame0
            response = GetSend(cmdStr, True)
            resp = response
            'keep original for debug
            'resp: frame0=4 // SPI data Prx_1Z_um ->frame #0(0)
            spos = InStr(resp, "data")
            NameOfVariable = "N/A"
            If spos > 0 Then
                NameOfVariable = Mid(resp, spos + 5)
                'result: Prx_1Z_um ->frame #0(0)
                spos = InStr(NameOfVariable, " ->")
                If spos > 0 Then
                    NameOfVariable = Left(NameOfVariable, spos)
                    'result: Prx_1Z_um
                    ChName(ch) = NameOfVariable
                Else    'error
                    NameOfVariable = "error" + ch
                End If
            Else    'error
                NameOfVariable = "error" + ch
            End If
            txtStatus.Text = Format((Timer - StartTime), " 0.000 s") + " Reading Ch Name " + Format(ch, " 000 ") + NameOfVariable
            txtManualCMD.Text = cmdStr + ":" + NameOfVariable
        Next ch

        ' restore original frame
        cmdStr = "frame0=" + CStr(ch0_indx)
        response = GetSend(cmdStr, True)
    Else
        dname_command = True
        For ch = 0 To UBound(ChName) - 1 'Ubaund is 256
            'cmd: dnamXXX
            'resp: dnam0=Prx_1Z_um
            cmdStr = "dnam" + CStr(ch)
'            Time_Now = Timer
            NameOfVariable = GetSend(cmdStr, True)
'            CmdTime = Timer - Time_Now
'            DoEvents
            If (NameOfVariable = "zero") Then NameOfVariable = "--zero--"
            ChName(ch) = NameOfVariable
            txtManualCMD.Text = cmdStr + ":" + NameOfVariable
            txtStatus.Text = Format((Timer - StartTime), "#0.000 s") + " Reading Ch Name " + Format(ch, " 000 ") + NameOfVariable
        Next ch
    End If
    'response = GetSend("echo>enab", False) ' disable verbose to speed up transmission
    'replace names
    Call CmbBNC0.Clear
    Call CmbBNC1.Clear
    For ch = 0 To UBound(ChName)
        CmbBNC0.AddItem (ChName(ch))
        CmbBNC1.AddItem (ChName(ch))
    Next ch
End Sub

Private Function GetIndexOfChannel(chnl As Integer) As Integer
    ' "frameN=XX" - will replace pointer to the variable in frame "N"
    'response
    Dim resp As String
    Dim cmdStr As String
    Dim spos As Long
    Dim indx_int As Integer
    Dim indx_str As String

    cmdStr = "frame" + Hex(chnl) + "?"
    ' Emit frame index in hex (:X) format
    resp = GetSend(cmdStr, True)
    'resp: frameC=4 // SPI data Prx_1Z_um ->frame #0(0)
    spos = InStr(resp, "=")
    If spos > 0 Then
        indx_str = Mid(resp, spos + 1)
        spos = InStr(indx_str, "//")
        'trim verbose comment
        If spos > 1 Then
            indx_str = Left(indx_str, spos - 1)
        End If
        spos = InStr(indx_str, ">~")
        If (IsNumber(indx_str)) Then
            indx_int = Int(indx_str)
        Else 'error
            indx_int = -1
        End If
    Else   'error
        indx_int = -1
    End If
    GetIndexOfChannel = indx_int ' return
End Function

'====================================  E N D   O F  =========================================================================
''''''''''''''''''''''''''''''''''''''  DIAGNOSTIC  '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


'============================================================================================================================
'''''''''''''''''''''''''''''''''''''' ALARMS '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'============================================================================================================================

Private Sub cmdClearErrorLog_Click()
    Dim LogMessage As String
    lstAlarms.Clear
    LogMessage = Me.in_time & " " & Me.in_date & ", Alarm List Cleared"
'    frmSplash.AddLogMessage (LogMessage)
End Sub

Private Sub cmdCloseAlarmLog_Click()
    lstAlarmFile.Clear
    lstAlarmFile.Visible = False
    lstAlarms.Height = 6570
End Sub

Private Sub cmdViewAlarmLog_Click()
  On Error GoTo error_no_file
    Dim rec As Long
    Dim temp_str As String
    rec = 0
    lstAlarms.Height = 3420 'roll up to free space
    lstAlarmFile.Clear
    lstAlarmFile.Visible = True

    Open Application_Dir & "\Alarm.log" For Input As Alarm_file 'log line
    Do While (Not EOF(Alarm_file))  ' Loop until end of file.
        Line Input #Alarm_file, temp_str
        rec = rec + 1
        lstAlarmFile.AddItem temp_str
    Loop

    Close #Alarm_file
    txtStatus.Text = Application_Dir & "\Alarm.log has " & rec & " records"
    Exit Sub
error_no_file:
    rec = MsgBox("File  might be missing or deleted, create it?", vbYesNo + vbDefaultButton1 + vbExclamation, "Error opening file 'Alarm.log'")
    If rec = vbYes Then   ' User chose yes, create.
        Open Application_Dir & "\Alarm.log" For Output As Alarm_file 'log line
        Print #Alarm_file, "File Alarm.log"
        Close #Alarm_file
    End If
End Sub

'====================================  E N D   O F  =========================================================================
''''''''''''''''''''''''''''''''''''''' ALARMS ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''






'============================================================================================================================
''''''''''''''''''''''''''''''''''''' FIRMWARE PARAMS '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'============================================================================================================================

Private Sub cmdUpdateFirmware_Click()
    Call Update_Help
End Sub

Sub Update_Help()
    Dim appPath As String
    Dim Msg, Style, Title, answer
    Dim IsBootloader As Boolean
    Call UpdateConnectControls(False)
    If Windows_is_64bit Then
        appPath = "C:\Program Files (x86)\C2Prog\"
    Else
        appPath = "C:\Program Files\C2Prog\"
    End If
    appPath = appPath + "C2Prog.exe"
    If FileExists(appPath) = False Then
        Call MsgBox("CodeSkin Firmware updater 'C2Prog' is not found at" & vbCrLf & appPath & vbCrLf & "Please install from www.CodeSkin.com")
        Exit Sub
    End If
    IsBootloader = (InStr(SystemFWversion, "bld") <> 0) Or (InStr(SystemFWversion, "-b") <> 0)
    Title = "How To Update Firmware - READ ALL before dismissing!"
    Msg = "You should 'Retrieve and Save' Parameters before upgrading," & vbCrLf & "if not - press 'Cancel' and save." & vbCrLf & vbCrLf
        If (IsBootloader = True) Then ' bootloader is present
            Msg = Msg + "Upgrade is possible via Ethernet, Front, and Rear USB type A ports." & vbCrLf & "DO NOT unplug USB cable!" & vbCrLf
        Else ' NO bootloader, upgrade via micro USB
            Msg = Msg + "Upgrade is possible via Front micro USB." & vbCrLf & "If micro USB cable is plugged, unplug it!" & vbCrLf
        End If
        Msg = Msg & "If you click OK, C2Prog.exe will start"
        Style = vbOKCancel + vbInformation ' Define buttons.
        answer = MsgBox(Msg, Style, Title)
        If answer = vbOK Then
            Call Shell(appPath, vbNormalFocus) ' START C2Prog
        Else
            Exit Sub
        End If
        Title = "How To Update Firmware - setting up C2Prog"
        Msg = "C2Prog.exe has been started." & vbCrLf & vbCrLf & "Select new firmware image '1st_Upgr*.ehx' file" & vbCrLf & vbCrLf
        If (ConnectionType = ConnTelnet) Then
        ' message for Telnet port
        Msg = Msg & "To load via Ethernet: type 'socket:controller_IP_adr:2020' (no spaces!)," & vbCrLf & "for example:" & vbCrLf & _
                    "'socket:169.254.20.20:2020' on the bottom 'Port selection' line" & vbCrLf & vbCrLf & _
                    "When you press 'OK' Ethernet port becomes 'DISconnected' to prepare for upgrade," & vbCrLf

        Else '            If (ConnectionType = ConnCOMport) Then
        Msg = Msg & "On 'C2Prog' screen press 'Configure Ports', in new window press 'Scan Ports', " & vbCrLf & _
                    "On 'C2Prog' from drop down list select COM port number Analyzer was connected to" & vbCrLf & _
                    "On 'C2Prog' hit 'OK' to return to C2Prog main window" & vbCrLf & _
                    "On 'C2Prog' choose 1st firmware image '1st_Upgr*.ehx' file " & vbCrLf & _
                    "When you press 'OK' COM port becomes 'DISconnected' to prepare for upgrade," & vbCrLf
        End If
        Msg = Msg & "and DC-2020 will be restarted. PERFORM NEXT STEP IN LESS THAN 15 sec!" & vbCrLf & _
                     "When you hit 'OK' here, IMMEDIATELY switch to 'C2Prog' and click 'Program' " & vbCrLf & _
                     "Now press 'OK' to restart DC-2020 and continue update process or 'Cancel' to quit"

    Style = vbOKCancel + vbInformation ' Define buttons.
    answer = MsgBox(Msg, Style, Title)
    If answer = vbOK Then
        Title = "How To Update Firmware - programming by using C2Prog"
        If (IsBootloader = True) Then ' bootloader is present
            cwCOMdisconnect.Value = False ' disconnect from the port
            response = GetSend("bootload", False) ' 16 seconds
            Msg = "1. DC-2020 has been restarted. Perform steps 2-4 in less than 15 sec!" & vbCrLf & _
                  "2. 'C2Prog' should be already set: port and 'ehx' file are selected" & vbCrLf & _
                  "3. On 'C2Prog', press the 'Program' button" & vbCrLf & _
                  "4. The 'C2Prog' should start programming, DC-2020 LED will blink yellow" & vbCrLf & _
                  "4a. If not: after 30 sec trying C2Prog throws an error." & vbCrLf & _
                  " b. On DC-2020, turn off POWER switch, and on 'C2Prog' press the 'Program' button" & vbCrLf & _
                  " c. On DC-2020, turn ON the POWER switch. LED might blink RED then blinks yellow" & vbCrLf & _
                  "5. Wait while 'C2Prog' informs : 'erasing'; 'downloading'; 'verifying'" & vbCrLf & _
                  "6. When programming is done, choose 2nd firmware image '2nd_Upgr*.ehx' file" & vbCrLf & _
                  "7. On 'C2Prog' screen, press the 'Program' button" & vbCrLf & _
                  "8. When programing has been finished, controller should restart automatically" & vbCrLf & _
                  "9. After pressing 'OK' on this dialog, then press 'Read and Restore Parameters'" & vbCrLf & _
                  "10. Parameters will be restored in RAM, do not forget to save them to FLASH" & vbCrLf & _
                  "11. Now press 'OK' to dismiss this screen"

        Else ' NO bootloader, upgrade via micro USB
            response = GetSend("restart", False) '
            cwCOMdisconnect.Value = False ' disconnect from the port
            Msg = "1. C2Prog.exe' updater had started, select new firmware image - '*1st.ehx' file" & vbCrLf & _
                  "2. On DC-2020 turn off POWER switch, press and HOLD the UPGRADE button" & vbCrLf & _
                  "3. While HOLDING DOWN the UPGRADE button, turn on DC-2020" & vbCrLf & _
                  "4. RELEASE the UPGRADE button, connect micro-USB cable" & vbCrLf & _
                  "5. Using Device Manager find new 'FTDI' COM port note COM number #" & vbCrLf & _
                  "6. On 'C2Prog' screen press 'Configure Ports', in new window press 'Scan Ports', " & vbCrLf & _
                  "7. Select that COM # from drop down list. Hit 'OK' to return to C2Prog window" & vbCrLf & _
                  "8. On 'C2Prog' screen, press the 'Program' button" & vbCrLf & _
                  "9. Wait while 'C2Prog' informs you: 'erasing'; 'downloading'; 'verifying'" & vbCrLf & _
                  "10. When 1st file programming is done, choose firmware image - '*2nd.ehx' file" & vbCrLf & _
                  "11. Press the 'Program' button" & vbCrLf & _
                  "12. When programing has been finished, DC-2020 should restart automatically" & vbCrLf & _
                  "13. After pressing 'OK' on this dialog, then press 'Read and Restore Parameters'" & vbCrLf & _
                  "14. Parameters will be restored in RAM, do not forget to save them to FLASH" & vbCrLf & _
                  "15. Now press 'OK' to dismiss this screen" ' 9 chars left in available length
        End If
        Style = vbOKOnly + vbInformation ' Define buttons.
        answer = MsgBox(Msg, Style, Title)
    End If
End Sub

Private Sub CmdClearParamList_Click()
'-   textParams = ""
'-   textParams.Visible = True
'-   textParams.Refresh
    ParamLINEScount = 0
    paramCHARcount = 0
    ListParams.Clear

    ListParamsErrors.Clear

    ListShadow.Visible = False
    ListParams.Visible = True
    ListParams.Refresh
    LabelParamLines.Caption = ""

End Sub


Private Sub cmdRestoreParameters_Click()
    Call Read_Parameters_File
End Sub

Function ReadLineAndAddToList(ByRef temp_str As String, ByRef lineNum As Long) As String
    Dim line_No_str As String
    Line Input #Params_file, temp_str
    lineNum = lineNum + 1
    line_No_str = Format(lineNum, "000#: ")
    AddToParameters (line_No_str & temp_str)
    ReadLineAndAddToList = temp_str
    If ParamLINEScount > 40 Then ' show shadow list / hide real one to speed up
        ListShadow.Visible = True
        ListParams.Visible = False
    Else ' hide shadow / show real list
        ListParams.Visible = True
        ListShadow.Visible = False
        ListShadow.AddItem (line_No_str & temp_str) 'ADDING LINE TO SHADOW LIST
    End If
End Function

Sub Read_Parameters_File()
    On Error Resume Next
    Dim cont_monitoring As Boolean
    Dim DefaultFilename As String
    Dim DataFileName As String
    Dim FileName As String
    Dim i As Long
    Dim temp_str As String
    Dim sep_pos As Long
    Dim line_counter As Long
    Dim line_No_str As String
    Dim line_error_number As Long
    Dim Msg As String, Style As Long, Title As String, answer As Long
    Dim FW_version As String
    Dim Param_FW_version As String
    Dim start_time As Single
    cwBut_RT_Monitoring.Value = False ' cancel RT monitoring

   DefaultFilename = Date$ & "DC-2020 parameters.param"
    dlgCommonDialog.CancelError = True
    dlgCommonDialog.DialogTitle = "Choose DC-2020 parameters file name to load"
    ' dlgCommonDialog.FileName = DefaultFilename
    dlgCommonDialog.InitDir = "c:\test"
    dlgCommonDialog.DefaultExt = "param"
    dlgCommonDialog.Filter = "Data File (*.param)|*.param|All files (*.*)|*.*"
    dlgCommonDialog.Action = 1
    'dlgCommonDialog.ShowSave
    If Err = 32755 Then 'Cancel is selected
        Exit Sub
    Else
        txtStatus.Text = "Reading Saved Parameters file '" & dlgCommonDialog.FileName & " ....."
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

    line_counter = 0
'    textParams = ""
    ListParams.Clear
    ListShadow.Clear
    ListParamsErrors.Clear
    cwLabelAction.Value = True
    cwLabelAction.Visible = True
    ParamLINEScount = 0
    paramCHARcount = 0

    Open DataFileName For Input As #Params_file 'Values
    start_time = Timer
    Line Input #Params_file, temp_str 'Print #Params_file, "File name, "; FileName
    ' decode line
    sep_pos = InStr(temp_str, ",")
    If InStr(LCase(temp_str), "file name") = 0 Then
        GoTo error_format
    Else
        txtStatus = Mid(temp_str, sep_pos + 1)
    End If
' next line is "DC-2020 parameters"
    temp_str = ReadLineAndAddToList(temp_str, line_counter) '    Print #Params_file, "DC-2020 parameters"
    ' decode line
    sep_pos = InStr(LCase(temp_str), "param")
    If sep_pos = 0 Then
        GoTo error_format
    End If

' next line is FW version
find_version:
    temp_str = ReadLineAndAddToList(temp_str, line_counter) '    Print #Params_file, "Date, "; Date$
    'Line Input #Params_file, temp_str   '    Print #Params_file, "Time, "; Time$
    'Line Input #Params_file, temp_str '//// Controller internal Time=00:58:49.21
        sep_pos = InStr(LCase(temp_str), "version")
        If sep_pos <> 0 Then
            FW_version = temp_str
            Param_FW_version = Mid(FW_version, 14)
        Else
            If InStr(temp_str, "//") Then
                GoTo find_version
            Else ' end of header
                FW_version = "version string is not found"
            End If
        End If
        sep_pos = InStr(Param_FW_version, SystemFWversion) 'compare saved and current versions
        If sep_pos = 0 Then 'versions different
            Title = "Firmware and Parameters file have different versions"
            Msg = "   Firmware: " & SystemFWversion & vbCrLf & _
                "Parameters: " & Param_FW_version & vbCrLf & _
                "Continue?" & vbCrLf & _
                "'Yes' will update only recognized parameters; 'No' will not change parameters"

            Style = vbYesNo + vbDefaultButton1 + vbExclamation ' Define buttons.
            answer = MsgBox(Msg, Style, Title)
            If answer = vbNo Then   ' User chose NO.
                Close #Params_file
                Exit Sub   ' Perform some action.
            End If
        End If

' load parameters
    Msg = ""
    If (IsFormLoaded("frmMonitoring")) Then ' DO NOT USE frmMonitoring.Name - this will load the form to get its name. instead, use string im quotas "frmMonitoring"
        cont_monitoring = frmMonitoring.aTimerContMonitoring.Enabled ' save status
        frmMonitoring.aTimerContMonitoring.Enabled = False ' disable update - not to interfere with these com port requests
    End If
    Wait_sec (0.1) 'let transmission, if any, finish

    temp_str = GetSend("echo>enable", True) 'set echo - it might be disabled



read_again:
    Do While (Not EOF(Params_file))  ' Loop until end of file.
        Msg = ""
        temp_str = ReadLineAndAddToList(temp_str, line_counter)  '   read settings line
        If (temp_str = "") Then GoTo read_again

        sep_pos = InStr(temp_str, "//") 'find comment
        ' it can be comment WITHOUT command, like //// version
        ' or comment AFTER command
        If sep_pos <> 0 Then 'comment
            temp_str = Left(temp_str, sep_pos - 1) 'cut comment - controller does not need it
        End If
        temp_str = Replace(Trim(temp_str), vbTab, "")   ' IK 20220129 remove TABs
        temp_str = Replace(temp_str, " ", "")   ' IK 20220130 remove spaces
        If (Len(temp_str) > 0) Then 'it was command with comment
            Msg = GetSend(temp_str, True) '>>>>>> SENDING TO CONTROLLER
        End If
        If (InStr(LCase(Msg), ">~un") <> 0) Then
            ListParamsErrors.AddItem "! Error on Line: " & line_counter & " '" & Msg & " '"
            ListParamsErrors.ItemData(ListParamsErrors.NewIndex) = line_counter
            line_error_number = line_error_number + 1
        ElseIf (InStr(LCase(Msg), ">~er") <> 0) Then
            ListParamsErrors.AddItem "! Error on Line: " & line_counter & " '" & Msg & " '"
            ListParamsErrors.ItemData(ListParamsErrors.NewIndex) = line_counter
            line_error_number = line_error_number + 1
        ElseIf (InStr(LCase(Msg), "timeout") <> 0) Then
            ListParamsErrors.AddItem "! Error on Line: " & line_counter & " '" & Msg & " '"
            ListParamsErrors.ItemData(ListParamsErrors.NewIndex) = line_counter
            line_error_number = line_error_number + 1
        ElseIf (InStr(LCase(Msg), "<!") <> 0) Then ' warning
            ListParamsErrors.AddItem "Warning on Line: " & line_counter & " '" & Msg & " '"
            ListParamsErrors.ItemData(ListParamsErrors.NewIndex) = line_counter
        ElseIf (Msg <> "") Then ' if response is not OK, warning
            ListParamsErrors.AddItem "Warning on Line: " & line_counter & " '" & Msg & " '"
            ListParamsErrors.ItemData(ListParamsErrors.NewIndex) = line_counter
            AddToParameters (Msg)
        End If
    Loop
    Close #Params_file
    start_time = Timer - start_time

'    cwLabelAction.Value = True
    cwLabelAction.Visible = False

    ' hide shadow / show real list
    ListShadow.Visible = False
    ListParams.Visible = True
    ListParamsErrors.AddItem "Received " & line_counter & " lines" & ", took " & start_time & " sec"
    ListParamsErrors.AddItem "Errors " & line_error_number

    If (IsFormLoaded("frmMonitoring")) Then ' DO NOT USE frmMonitoring.Name - this will load the form to get its name. instead, use string im quotas "frmMonitoring"
        frmMonitoring.aTimerContMonitoring.Enabled = cont_monitoring  ' restore status
    End If
    Exit Sub
error_format:
    Close #Params_file
    cwLabelAction.Visible = False
'    FrmDebug.Add_Debug_Message_to_File ("Read parameters file line #" & line_counter & "'" & temp_str & "' unrecognized")
    MsgBox ("Unrecognizable file format")
'    Init_sys_data = False
End Sub


Private Sub cmdSaveParameters_Click()
    Call get_sys_parameters
End Sub

Public Function get_sys_parameters()
'    On Error GoTo end_retrieving

    Dim tmp_init As Boolean
    Dim monitor_stat, cont_field_monitoring As Boolean
    Dim temp_str As String

    Dim CmdStartTime As Single
    Dim WaitTime As Single
    Dim Cmd_String As String
    Dim sentCmd As String
    Dim CR_pos As Long
    Dim pos_OK As Long, pos_response As Long
    Dim got_response As Integer

    tmp_init = Init_sys_data
    Init_sys_data = True

'-    textParams = ""
'-    textParams.Visible = True
    frmTerminal.Hide
    ListParams.Visible = True
    ListShadow.Visible = False
    ListShadow.Clear
    ListParams.Clear
    ListParamsErrors.Clear
    ParamLINEScount = 0
    paramCHARcount = 0

    If IsFormLoaded("frmMonitoring") Then
        cwBut_RT_Monitoring.Value = False ' disable RT
        cont_field_monitoring = frmMonitoring.aTimerContMonitoring.Enabled ' save status
        frmMonitoring.aTimerContMonitoring.Enabled = False ' disable update - not to interfere with these com port requests
    End If
    Wait_sec (0.1) 'let transmission, if any , finish

    Call frmTerminal.QuitMenu ' menu mode exit to get parameters

    cwLabelAction.Value = False
    cwLabelAction.Visible = True

'start recording

'''''''''''''''' capture and save lines
    sentCmd = "expo"
    If Opt_ALL_params_or_Essential(1).Value = True Then sentCmd = sentCmd + " esse"
    If ChkVerboseParams.Value = Checked Then sentCmd = sentCmd + " verb"

send_cmd:
    Call open_port_if_closed 'if port not available, it will set ConnectionType = ConnDEMO
    If ConnectionType = ConnDEMO Or ConnectionType = ConnSimDLL Then
        cwCOMdisconnect.Caption = Demo_Simul_string
        If (DLL_is_available = False) Then '
            Exit Function 'port cannot be open, and DLL not found
         End If
    End If
    Call UpdateConnectControls(True) '    If Value = True Then 'open
    txtManualCMD.Text = sentCmd ' diagnostic
' Send the  command
    Cmd_String = sentCmd & vbCr 'Lf
    SendStrLen = Len(Cmd_String)
    Call Clean_COM_Buffer ''clean the Comm port buffer
    CmdStartTime = Timer
    Call Send_GUI_command(Cmd_String) ' >>>>>>>>> "expo"
    pos_response = 1
    got_response = 0
    Wait_sec (0.1) 'let echo transmission  finish
    'cwReceiving.Visible = True

    Do
        DoEvents
read_line:
        ' if it is DLL, it gets there when buffer of 5000 chars is loaded because buffer is limited to 5000
        ' Call VB6controller(cmd, resp, 5000) ' response is returned into Buffer by DLLhasCHAR_CallBackFunction

        pos_OK = InStr(Buffer, vbCrLf) 'wait for line end 13,10;\r\n
        If (pos_OK <> 0) Then
            temp_str = Left(Buffer, pos_OK - 1) 'separate one line
'            DoEvents
            If (InStr(temp_str, sentCmd)) Then
                Buffer = Mid(Buffer, pos_OK + 2)  ' trim buffer of echo of command
                GoTo read_line
            End If
            Cmd_String = temp_str ' make a copy as PREVIOUS string

            Buffer = Mid(Buffer, pos_OK + 2)  ' trim buffer

            ' check end of transmission
            got_response = InStr(Cmd_String, Acknowledgment_token & "OK")  ' ">~"
            If (got_response <> 0) Then
                ListShadow.Visible = False
                ListParams.Visible = True
                ListParams.Refresh
                ListParamsErrors.AddItem "Got Lines = " & ParamLINEScount
                ListParamsErrors.AddItem "Got Chars = " & paramCHARcount
                ListParamsErrors.AddItem "Time taken " & Format(WaitTime, "0.0## s")
            Else ' ending and writing statistics

                AddToParameters (temp_str) 'ADDING LINE TO LIST

'                cwReceiving.Caption = CStr(ParamLINEScount) & " Params"
                DoEvents

                If ParamLINEScount > 40 Then
                    ListShadow.Visible = True
'-                    textParams.Visible = False
                    ListParams.Visible = False
                Else
                    ListShadow.Visible = False
'-                    textParams.Visible = True
                    ListParams.Visible = True
                    ListShadow.AddItem temp_str 'ADDING LINE TO LIST
                End If
            End If
'            DoEvents
        End If
        WaitTime = Timer - CmdStartTime
' 1-letter command is too short and miss end of response. LOOK for ":" in response
' actual response: ASC(10) & "e" & ASC(13) & ASC(10) & "Feedforward ENABLED...". IT MUST BE IN ADDITION ":8e"
        'DoEvents
    Loop Until (got_response > 0) Or (WaitTime > 30)

end_retrieving: ' from controller
    lblCMDtime.Caption = Format(WaitTime, "#0.####")
    Wait_sec (0.1) 'let transmission, if any , finish
    Buffer = "" ' clear buffer just in case

    ListShadow.Visible = False
    ListParams.Visible = True
    ListParams.Refresh
    cwLabelAction.Visible = False

    Call Save_Parameters_to_PC
    If (IsFormLoaded("frmMonitoring")) Then ' DO NOT USE frmMonitoring.Name - this will load the form to get its name. instead, use string im quotas "frmMonitoring"
        frmMonitoring.aTimerContMonitoring.Enabled = cont_field_monitoring  ' restore status
    End If
    Init_sys_data = tmp_init
End Function

Sub Save_Parameters_to_PC()
    On Error Resume Next
    'user wants to save a test file
    'First get the filename
    Dim DefaultFilename As String
    Dim SaveFileName As String
    Dim FileName As String
    Dim i As Long
    Dim colon_pos As Long
    Dim param_line As String

    cwBut_RT_Monitoring.Value = False ' cancel RT monitoring

   DefaultFilename = "DC-2020 parameters " & Date$ & " " & Left$(Time$, 2) & "-" & Mid$(Time$, 4, 2) & ".param"
    dlgCommonDialog.CancelError = True
    dlgCommonDialog.DialogTitle = "Choose parameters file name"
    ' dlgCommonDialog.FileName = DefaultFilename
    dlgCommonDialog.InitDir = "c:\test"
    dlgCommonDialog.DefaultExt = "csv"
    dlgCommonDialog.Filter = "Data File (*.param)|*.param|All files (*.*)|*.*"
    dlgCommonDialog.FileName = DefaultFilename 'if enabled,  FILE NAME SHOULD NOT HAVE ':'
    dlgCommonDialog.ShowSave
    If Err = 32755 Then 'Cancel is selected
        Exit Sub
    Else
        txtStatus.Text = "File(s)=" & dlgCommonDialog.FileName
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
    Open SaveFileName For Output As Params_file 'parameters
    Print #Params_file, "//// File name, "; FileName
    Print #Params_file, "//// DC-2020 parameters"
    Print #Params_file, "//// Date, "; Date$
    Print #Params_file, "//// Time, "; Time$

'-    Print #Params_file, textParams.Text
    For i = 0 To ParamLINEScount 'textParams.ListCount
        param_line = ListParams.List(i)
        colon_pos = InStr(param_line, ": ")
        If (colon_pos < 6 And colon_pos > 0) Then
            param_line = Mid(param_line, colon_pos + 2)
        End If
        Print #Params_file, param_line
    Next i
    Close #Params_file
End Sub


Public Sub AddToParameters(msg_str As String)
    ListParams.AddItem (msg_str)
    DoEvents
    ParamLINEScount = ParamLINEScount + 1
    paramCHARcount = paramCHARcount + Len(msg_str) 'LfCr is not included
    LabelParamLines.Caption = ParamLINEScount
End Sub

Private Sub ListParamsErrors_Click()
    Dim selectedIndex As Long
    selectedIndex = ListParamsErrors.ItemData(ListParamsErrors.ListIndex)
    ListParams.ListIndex = selectedIndex
    If (selectedIndex > 3) Then
        ListParams.TopIndex = selectedIndex - 3
    End If
End Sub

Private Sub CmdSaveParamList_Click()
    Call Save_Parameters_to_PC
End Sub

'====================================  E N D   O F  =========================================================================
'''''''''''''''''''''''''''''''''''''' FIRMWARE PARAMS ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''



'============================================================================================================================
'''''''''''''''''''''''''''''''''''''' HIDE SHOW EXPERT CONTROLS ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'============================================================================================================================
Sub HideShowExpertControls()
    LblTerminalWindow.Visible = ShowExpertControls
    ImageTerminal.Visible = ShowExpertControls
'    FrameCallTuning.Visible = ShowExpertControls
    ChkShowPredictionPlot.Visible = ShowExpertControls
    'chkPOSmonitoring.Visible = ShowExpertControls
    LblPOSmonitoring.Visible = ShowExpertControls
    ImageMonitoring.Visible = ShowExpertControls
    ChkVerboseParams.Visible = ShowExpertControls
    Opt_ALL_params_or_Essential(0).Visible = ShowExpertControls
    Opt_ALL_params_or_Essential(1).Visible = ShowExpertControls
    If (ShowExpertControls = True) Then ' on parameters tab, blocks visible "all parameters" and "include comments"
        FrameSaveParamsIntoFile.Height = 1770
    Else
        FrameSaveParamsIntoFile.Height = 825
    End If
End Sub
'====================================  E N D   O F  =========================================================================
'''''''''''''''''''''''''''''''''''''' HIDE SHOW EXPERT CONTROLS ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''



Private Sub txtSerialNumber_Change()
End Sub

Private Sub TxtSystemID_KeyPress(KeyAscii As Integer)
   Dim charPressed As String
    charPressed = Chr(KeyAscii)
    If (charPressed = vbCr) Then
       response = GetSend("sysid=" & TxtSystemID.Text, True) ' send command with wait, vbCr detected
    End If
End Sub

Private Sub CmdUpdateSysID_Click()
    response = GetSend("sysid=" & TxtSystemID.Text, True) ' send command with wait, vbCr detected
End Sub


Sub ReadGoNoGoPlot(TestType As String, actuator As String, sensor As String)
    On Error GoTo error_noFile
    Dim DataFileName As String
    Dim axis As String
    Dim FileName As String
    Dim i  As Long
    Dim temp_str As String
    Dim sep_pos As Long
    Dim line_counter As Long
    Dim limits_type As Boolean
'    Dim FullPath As String, ParentFolder As String, l() As String
'    FullPath = "" '... Write here the path from ComDlg
'    l = Split(FullPath, "\")
    If (ChkShowLimits.Value <> Checked) Then
        Exit Sub 'IK20251218 disabled always read Go-NoGo plot, now checkmark controls it
    End If

    TestType = UCase(TestType)
    actuator = UCase(actuator)
    sensor = UCase(sensor)
    If ((TestType <> "OLTF") And (TestType <> "P_TF")) Then
        GoTo invalid_mot_sens_combination
    End If

    axis = Mid(actuator, 3, 2)
    If (axis <> "_X" And axis <> "_Y" And axis <> "_Z") _
      And (axis <> "TX" And axis <> "TY" And axis <> "TZ") _
      Or (optTestInput(0).Value = True) Then  'or Measure at one freq
        GoTo invalid_mot_sens_combination
    End If

    If (Left(actuator, 1) = "V" And (axis = "_X" Or axis = "_Y" Or axis = "TZ")) Then   'Valves cannot drive horizontal axes
        GoTo invalid_mot_sens_combination
    End If

    ' source of limits from User selected folder/file
    ' need to read test limits file
    FileName = ""
    dlgCommonDialog.FileName = "Dbl-click on file or click OPEN to select folder"
    If (UserSelectedCriteriaPath = "") Then
'            MsgBox ("Select Path to Criteria Files")
        With dlgCommonDialog
            .CancelError = True
            .DialogTitle = "Dbl-click on test limits file OR click OPEN - criteria auto-selects from this folder"
            .DefaultExt = "csv"
            .FileName = "Dbl-click on file or click OPEN to select folder"
            .Filter = "Data File (*.txt)|*.txt|All files (*.*)|*.*"
            .Flags = cdlOFNLongNames + cdlOFNPathMustExist
            .InitDir = "C:\test"

            .ShowOpen
            If Err = 32755 Then 'Cancel is selected
                Exit Sub
            End If
        End With

        DataFileName = dlgCommonDialog.FileName 'i.e. "C:\In_Out\Software\DC-2020_ED_Analyzer_VB6\Test\TMC_Win-450_Test_System\X_OLTF_Go_NoGo.csv"
        i = InStrRev(DataFileName, "\")
        FileName = Right(DataFileName, Len(DataFileName) - i) 'i.e. "X_OLTF_Go_NoGo.csv"
        UserSelectedCriteriaPath = Left(DataFileName, i) 'i.e. "C:\In_Out\Software\DC-2020_ED_Analyzer_VB6\Test\TMC_Win-450_Test_System\
    End If

    If (InStr(dlgCommonDialog.FileName, "Dbl-click")) Then ' if user selected folder, user did not select file, means criteria auto selection
        FileName = TestType & " " & actuator & " " & sensor & " Go_NoGo.txt"
        DataFileName = UserSelectedCriteriaPath & FileName
    Else ' user selected particular file
        DataFileName = dlgCommonDialog.FileName
        i = InStrRev(DataFileName, "\")
        FileName = Right(DataFileName, Len(DataFileName) - i) 'i.e. "X_OLTF_Go_NoGo.csv"
    End If

    'now lets open the file and read the data
    txtStatus.Text = "Reading Limits from file '" & DataFileName & " ....."


    Close #Transfer_Function_file
    Open DataFileName For Input As #Transfer_Function_file 'Values
    Line Input #Transfer_Function_file, temp_str '    "File name,"; vbTab; FileName
    ' decode line
    sep_pos = InStr(LCase(temp_str), "go_nogo.txt")
    If (sep_pos = 0) Then
        GoTo error_format
    Else
        If sep_pos > 25 Then txtStatus.Text = Mid(temp_str, sep_pos - 15)
    End If

    For i = 0 To TEST_LTF_ARRAY_LENGTH - 1
        Freq_Lim_data(i) = 0 ' Limits on plot
        Gain_Low_data(i) = 0
        Gain_High_data(i) = 0
        Phase_Low_data(i) = 0
        Phase_High_data(i) = 0
    Next i

    Line Input #Transfer_Function_file, temp_str 'Print #Transfer_Function_file, "Frequency, Hz"; vbTab; "Gain, dB"; vbTab; "Phase, deg"
    '.......add code ................
    line_counter = -1 ' to start from zero
    Do While Not EOF(Transfer_Function_file)   ' Loop until end of file.
        line_counter = line_counter + 1

        Line Input #Transfer_Function_file, temp_str 'Print #Transfer_Function_file,
        ' decode line
        sep_pos = InStr(temp_str, vbTab) ' Print #Transfer_Function_file, Format(Freq_data(i), "#0.00"); vbTab;
        If IsNumber(Left(temp_str, sep_pos - 1)) Then
            Freq_Lim_data(line_counter) = Val(Left(temp_str, sep_pos - 1))
            temp_str = Mid(temp_str, sep_pos + 1)
        Else
            GoTo error_format
        End If
            'Print #Transfer_Function_file, Format(Gain_Low_data(i), "#0.000"); vbTab;
        sep_pos = InStr(temp_str, vbTab)
        If IsNumber(Left(temp_str, sep_pos - 1)) Then
            Gain_Low_data(line_counter) = Val(Left(temp_str, sep_pos - 1))
            temp_str = Mid(temp_str, sep_pos + 1)
        Else
            GoTo error_format
        End If

            'Print #Transfer_Function_file, Format(Phase_Low_data(i), "#0.0")
        sep_pos = InStr(temp_str, vbTab)
        If IsNumber(Left(temp_str, sep_pos - 1)) Then
            Phase_Low_data(line_counter) = Val(Left(temp_str, sep_pos - 1))
            temp_str = Mid(temp_str, sep_pos + 1)
        Else
            GoTo error_format
        End If

             'Print #Transfer_Function_file, Format(Gain_High_data(i), "#0.000"); vbTab;
        sep_pos = InStr(temp_str, vbTab)
        If IsNumber(Left(temp_str, sep_pos - 1)) Then
            Gain_High_data(line_counter) = Val(Left(temp_str, sep_pos - 1))
            temp_str = Mid(temp_str, sep_pos + 1)
        Else
            GoTo error_format
        End If

            'Print #Transfer_Function_file, Format(Phase_High_data(i), "#0.0")
        If IsNumber(temp_str) Then
            Phase_High_data(line_counter) = Val(temp_str)
        Else
            GoTo error_format
        End If

       ' lblLines.Caption = line_counter
    Loop

    Close #Transfer_Function_file
    'End If ' read limits from Criteria folder

    Call update_go_nogo_plot
    txtStatus.Text = "Limits file '" & DataFileName & "' read successfully"
    Exit Sub
invalid_mot_sens_combination:
    CWGraph.Plots(Lower_Lim_TF_Num).Visible = False
    CWGraph.Plots(Upper_Lim_TF_Num).Visible = False
    CWGraphPhase.Plots(Lower_Lim_TF_Num).Visible = False
    CWGraphPhase.Plots(Upper_Lim_TF_Num).Visible = False
'        response = MsgBox("Axes not selected correctly: " + LTFtestOutputName + " -> " + LTFtestInputName + vbCr + _
    "Select X-X, Y-Y, or Z-Z piezo-geo within the same isolator", vbOKOnly + vbExclamation, "Cannot load go_nogo file")
    txtStatus.Text = "File with Go-NoGO Limits not found, Select 'OLTF', 'P_TF' and X-X, Y-Y, or Z-Z out-in axis or actuator-sensor combination"
    Exit Sub
error_format:
    Close #Transfer_Function_file
    Call MsgBox("file '" & FileName & "'" & vbCrLf & "folder " & UserSelectedCriteriaPath, vbOKOnly + vbExclamation, "Unrecognizable limits file format")
    Exit Sub
error_noFile:
    txtStatus.Text = "File with Go-NoGO Limits not found: " & DataFileName
End Sub

Public Sub update_go_nogo_plot()
    Dim plot_num As Integer
'    plot_num = CWGraph.Plots.Count
    'AutoScaleGraph (Gain_data)
    'CWGraph.PlotY (Gain_data)
    CWGraph.Plots(Lower_Lim_TF_Num).ClearData
    CWGraph.Plots(Upper_Lim_TF_Num).ClearData
    CWGraphPhase.Plots(Lower_Lim_TF_Num).ClearData
    CWGraphPhase.Plots(Upper_Lim_TF_Num).ClearData

    CWGraph.Plots(Lower_Lim_TF_Num).PlotXvsY Freq_Lim_data, Gain_Low_data
    CWGraph.Plots(Upper_Lim_TF_Num).PlotXvsY Freq_Lim_data, Gain_High_data
    CWGraphPhase.Plots(Lower_Lim_TF_Num).PlotXvsY Freq_Lim_data, Phase_Low_data
    CWGraphPhase.Plots(Upper_Lim_TF_Num).PlotXvsY Freq_Lim_data, Phase_High_data
    If (ChkShowLimits.Value = Checked) Then
        CWGraph.Plots(Lower_Lim_TF_Num).Visible = True
        CWGraph.Plots(Upper_Lim_TF_Num).Visible = True
        CWGraphPhase.Plots(Lower_Lim_TF_Num).Visible = True
        CWGraphPhase.Plots(Upper_Lim_TF_Num).Visible = True
    Else
        CWGraph.Plots(Lower_Lim_TF_Num).Visible = False
        CWGraph.Plots(Upper_Lim_TF_Num).Visible = False
        CWGraphPhase.Plots(Lower_Lim_TF_Num).Visible = False
        CWGraphPhase.Plots(Upper_Lim_TF_Num).Visible = False
    End If
End Sub

Public Sub Check_Selected_Actuator_Sensor()
    Dim actuator_name As String
    Dim test_name As String
    Dim sensor_name As String
    Dim inp_name_2nd As String

    If (ChkShowLimits.Value = Unchecked) Then Exit Sub

    actuator_name = UCase(LTFtestOutputName) ' valid first letter 'V':VA_Z, VATX, VATY; 'M':MA_X, MA_Y, MA_Z, MATX, MATY, MATZ
    sensor_name = UCase(LTFtestInputName)    ' valid first letter 'P':PA_X, PA_Y, PATZ, PA_Z, PATX, PATY; 'G':GA_X, GA_Y, GA_Z, GATX, GATY, GATZ
    inp_name_2nd = UCase(LTFtest2ndInputName)
    If actuator_name = "" And sensor_name = "" Then GoTo In_out_is_not_selected ' user have not selected piezo
    If actuator_name = "" And sensor_name <> "" Then
        txtStatus.Text = "Actuator is not selected"
        cmdChooseExcitation.SetFocus
        Exit Sub ' user have not selected actuator
    End If
    If actuator_name <> "" And sensor_name = "" Then
        txtStatus.Text = "Sensor is not selected"
        CmdChooseInputSignal.SetFocus
        Exit Sub ' user have not selected geo
    End If
    If inp_name_2nd <> "" Then
        Call ReadGoNoGoPlot("-", "-", "-") 'this will cause hiding limit plots and message
        Exit Sub  ' user selected 2nd input for advanced diagnostic
    End If

    If (Mid(actuator_name, 2, 3) = Mid(sensor_name, 2, 3)) Then 'MA_X  GA_X, VA_Z, PA_Z,
        If optTestInput(1).Value = True Then 'OLTF test
            test_name = "OLTF"
        End If
        If optTestInput(2).Value = True Then 'plant TF test
            test_name = "P_TF"
        End If
        Call ReadGoNoGoPlot(test_name, actuator_name, sensor_name)
    Else ' axes names are different, like M_Z1, G_Y2
        Call ReadGoNoGoPlot("-", "-", "-") 'this will cause hiding limit plots and message
    End If
    Exit Sub
In_out_is_not_selected:
    Call MsgBox("Please choose Output and Input first", vbOKOnly + vbExclamation, "Parameters Not Defined")
End Sub
