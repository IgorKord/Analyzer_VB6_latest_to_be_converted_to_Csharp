VERSION 5.00
Object = "{D940E4E4-6079-11CE-88CB-0020AF6845F6}#1.6#0"; "cwui.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmMonitoring 
   Caption         =   "Payload Monitoring and Excitation"
   ClientHeight    =   9285
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11670
   Icon            =   "frmMonitoring.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9285
   ScaleWidth      =   11670
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer aTimerContMonitoring 
      Enabled         =   0   'False
      Left            =   6660
      Top             =   90
   End
   Begin VB.Timer aTimerTestProgram 
      Interval        =   100
      Left            =   10395
      Top             =   8640
   End
   Begin VB.TextBox txtStatus 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      Height          =   285
      Left            =   90
      TabIndex        =   82
      Text            =   "Position Monitoring, Recording and Excitation"
      Top             =   0
      Width           =   11400
   End
   Begin VB.Frame FramePosScan 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   8520
      Left            =   0
      TabIndex        =   0
      Top             =   300
      Width           =   11610
      Begin VB.Frame FrameTest 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Test type"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1770
         Left            =   90
         TabIndex        =   86
         Top             =   -45
         Width           =   2445
         Begin VB.ComboBox cmb_TestType 
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
            ItemData        =   "frmMonitoring.frx":030A
            Left            =   90
            List            =   "frmMonitoring.frx":0317
            Style           =   2  'Dropdown List
            TabIndex        =   105
            ToolTipText     =   "Choose axis OUTPUT for excitation"
            Top             =   225
            Width           =   2280
         End
         Begin VB.CommandButton cmdReadPositionFile 
            BackColor       =   &H00FFC0FF&
            Caption         =   "Read Test"
            Height          =   375
            Left            =   1215
            Style           =   1  'Graphical
            TabIndex        =   87
            ToolTipText     =   "Reads saved position data and shows plots"
            Top             =   1260
            Width           =   915
         End
         Begin CWUIControlsLib.CWButton cwSaveScan 
            Height          =   495
            Left            =   225
            TabIndex        =   106
            ToolTipText     =   "If plots exist, saves plots into data file. If there is no plots, starts recording, plotting, and saving into file."
            Top             =   585
            Width           =   1860
            _Version        =   196609
            _ExtentX        =   3281
            _ExtentY        =   873
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
            CompatibleVers_0=   196609
            Boolean_0       =   1
            ClassName_1     =   "CCWBoolean"
            opts_1          =   2614
            SclRef_1.l      =   2
            SclRef_1.t      =   2
            SclRef_1.r      =   120
            SclRef_1.b      =   40
            Scl_1.l         =   6
            Scl_1.t         =   7
            Scl_1.r         =   115
            Scl_1.b         =   36
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
            C[0]_7          =   -2147483640
            C[1]_7          =   -2147483640
            Image_7         =   8
            ClassName_8     =   "CCWTextImage"
            szText_8        =   "Saving, hit to STOP"
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
            szText_11       =   "Save as File(s)"
            font_11         =   0
            Animator_10     =   0
            Blinker_10      =   0
            list[1]_4       =   12
            ClassName_12    =   "CCWDrawObj"
            opts_12         =   62
            C[0]_12         =   57344
            C[1]_12         =   57344
            Image_12        =   13
            ClassName_13    =   "CCWPictImage"
            opts_13         =   1280
            Rows_13         =   1
            Cols_13         =   1
            Pict_13         =   2
            F_13            =   57344
            B_13            =   57344
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
            opts_17         =   62
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
            C[0]_19         =   16750335
            C[1]_19         =   16750335
            Image_19        =   20
            ClassName_20    =   "CCWPictImage"
            opts_20         =   1280
            Rows_20         =   1
            Cols_20         =   1
            Pict_20         =   2
            F_20            =   16750335
            B_20            =   16750335
            ColorReplaceWith_20=   8421504
            ColorReplace_20 =   8421504
            Tolerance_20    =   2
            Animator_19     =   0
            Blinker_19      =   21
            ClassName_21    =   "CCWBlinker"
            Interval_21     =   900
            EventID_21      =   7
            Obj_21          =   19
            varVarType_14   =   5
            Bindings_1      =   22
            ClassName_22    =   "CCWBindingHolderArray"
            Editor_22       =   23
            ClassName_23    =   "CCWBindingHolderArrayEditor"
            Owner_23        =   1
            Style_1         =   12
            frameStyle_1    =   1
            mechAction_1    =   1
            BGImg_1         =   24
            ClassName_24    =   "CCWDrawObj"
            opts_24         =   62
            Image_24        =   25
            ClassName_25    =   "CCWPictImage"
            opts_25         =   1280
            Rows_25         =   1
            Cols_25         =   1
            Pict_25         =   286
            F_25            =   -2147483633
            B_25            =   -2147483633
            ColorReplaceWith_25=   8421504
            ColorReplace_25 =   8421504
            Tolerance_25    =   2
            Animator_24     =   0
            Blinker_24      =   0
            Array_1         =   6
            Editor_1        =   0
            Array[0]_1      =   12
            Array[1]_1      =   19
            Array[2]_1      =   0
            Array[3]_1      =   0
            Array[4]_1      =   26
            ClassName_26    =   "CCWDrawObj"
            opts_26         =   62
            Image_26        =   8
            Animator_26     =   0
            Blinker_26      =   0
            Array[5]_1      =   27
            ClassName_27    =   "CCWDrawObj"
            opts_27         =   62
            Image_27        =   11
            Animator_27     =   0
            Blinker_27      =   0
            Label_1         =   28
            ClassName_28    =   "CCWDrawObj"
            opts_28         =   62
            C[0]_28         =   -2147483640
            Image_28        =   6
            Animator_28     =   0
            Blinker_28      =   0
         End
         Begin VB.Label lbLineNum 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "9000"
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
            Left            =   135
            TabIndex        =   88
            ToolTipText     =   "current  line number of position data"
            Top             =   1350
            Width           =   825
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Caption         =   "data line # "
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   17
            Left            =   135
            TabIndex        =   89
            Top             =   1125
            Width           =   825
         End
      End
      Begin VB.Frame framExitation 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Exitation"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1770
         Left            =   6840
         TabIndex        =   57
         ToolTipText     =   $"frmMonitoring.frx":035C
         Top             =   -45
         Width           =   4650
         Begin VB.CommandButton cmdPulse 
            BackColor       =   &H0000FFFF&
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
            Height          =   510
            Left            =   1035
            Style           =   1  'Graphical
            TabIndex        =   61
            ToolTipText     =   "Applies pulse to chosen axis (If excitation was ON, it is disabled after pulse)"
            Top             =   1125
            Width           =   870
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
            Height          =   300
            Index           =   0
            Left            =   1035
            TabIndex        =   60
            ToolTipText     =   "Enables Sine exitation (Frequency MUST BE NOT ZERO)"
            Top             =   270
            Value           =   -1  'True
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
            Height          =   300
            Index           =   1
            Left            =   1035
            TabIndex        =   59
            ToolTipText     =   "Enables noise exitation (Frequency is set to 0 Hz means niose)"
            Top             =   720
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
            ItemData        =   "frmMonitoring.frx":03F9
            Left            =   2655
            List            =   "frmMonitoring.frx":044A
            Style           =   2  'Dropdown List
            TabIndex        =   58
            ToolTipText     =   "Choose axis OUTPUT for excitation"
            Top             =   225
            Width           =   1875
         End
         Begin CWUIControlsLib.CWButton cwBut_Excitation 
            Height          =   1245
            Left            =   90
            TabIndex        =   62
            ToolTipText     =   "Enable / Disable excitation on chosen axis output"
            Top             =   360
            Width           =   870
            _Version        =   196609
            _ExtentX        =   1535
            _ExtentY        =   2196
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
            CompatibleVers_0=   196609
            Boolean_0       =   1
            ClassName_1     =   "CCWBoolean"
            opts_1          =   2614
            SclRef_1.l      =   2
            SclRef_1.t      =   2
            SclRef_1.r      =   60
            SclRef_1.b      =   85
            Scl_1.l         =   5
            Scl_1.t         =   12
            Scl_1.r         =   57
            Scl_1.b         =   75
            C[0]_1          =   8438015
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
            opts_7          =   52
            SclRef_7.l      =   7
            SclRef_7.t      =   19
            SclRef_7.r      =   66
            SclRef_7.b      =   84
            Scl_7.l         =   26
            Scl_7.t         =   26
            Scl_7.r         =   48
            Scl_7.b         =   45
            C[0]_7          =   -2147483640
            C[1]_7          =   -2147483640
            Image_7         =   8
            ClassName_8     =   "CCWTextImage"
            font_8          =   9
            ClassName_9     =   "CCWFont"
            bFont_9         =   -1  'True
            BeginProperty Font_9 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   12
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Animator_7      =   0
            Blinker_7       =   0
            list[2]_4       =   10
            ClassName_10    =   "CCWDrawObj"
            opts_10         =   54
            SclRef_10.l     =   7
            SclRef_10.t     =   19
            SclRef_10.r     =   66
            SclRef_10.b     =   84
            Scl_10.l        =   25
            Scl_10.t        =   58
            Scl_10.r        =   47
            Scl_10.b        =   77
            C[0]_10         =   -2147483640
            C[1]_10         =   -2147483640
            Image_10        =   11
            ClassName_11    =   "CCWTextImage"
            style_11        =   1076101120
            font_11         =   12
            ClassName_12    =   "CCWFont"
            bFont_12        =   -1  'True
            BeginProperty Font_12 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   12
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Animator_10     =   0
            Blinker_10      =   0
            list[1]_4       =   13
            ClassName_13    =   "CCWDrawObj"
            opts_13         =   62
            C[0]_13         =   8454016
            C[1]_13         =   8454016
            Image_13        =   14
            ClassName_14    =   "CCWPictImage"
            opts_14         =   256
            Rows_14         =   1
            Cols_14         =   1
            URL_14          =   "C:\in_out\Software\DC-2020_ED_Analyzer_VB6\Resources\SwitchOffGreen.bmp"
            Pict_14         =   "frmMonitoring.frx":0573
            ColorReplaceWith_14=   8421504
            ColorReplace_14 =   8421504
            Tolerance_14    =   2
            Animator_13     =   0
            Blinker_13      =   0
            varVarType_3    =   5
            Array[1]_2      =   15
            ClassName_15    =   "CCWEnumElt"
            opts_15         =   1
            Name_15         =   "On"
            frame_15        =   286
            DrawList_15     =   16
            ClassName_16    =   "CDrawList"
            count_16        =   4
            list[4]_16      =   17
            ClassName_17    =   "CCWDrawObj"
            opts_17         =   62
            C[0]_17         =   -2147483640
            C[1]_17         =   -2147483640
            Image_17        =   6
            Animator_17     =   0
            Blinker_17      =   0
            list[3]_16      =   18
            ClassName_18    =   "CCWDrawObj"
            opts_18         =   54
            SclRef_18.l     =   7
            SclRef_18.t     =   19
            SclRef_18.r     =   66
            SclRef_18.b     =   84
            Scl_18.l        =   26
            Scl_18.t        =   26
            Scl_18.r        =   48
            Scl_18.b        =   45
            C[0]_18         =   -2147483640
            C[1]_18         =   -2147483640
            Image_18        =   8
            Animator_18     =   0
            Blinker_18      =   0
            list[2]_16      =   19
            ClassName_19    =   "CCWDrawObj"
            opts_19         =   52
            SclRef_19.l     =   7
            SclRef_19.t     =   19
            SclRef_19.r     =   66
            SclRef_19.b     =   84
            Scl_19.l        =   25
            Scl_19.t        =   58
            Scl_19.r        =   47
            Scl_19.b        =   77
            C[0]_19         =   -2147483640
            C[1]_19         =   -2147483640
            Image_19        =   11
            Animator_19     =   0
            Blinker_19      =   0
            list[1]_16      =   20
            ClassName_20    =   "CCWDrawObj"
            opts_20         =   62
            C[0]_20         =   8421631
            C[1]_20         =   8421631
            Image_20        =   21
            ClassName_21    =   "CCWPictImage"
            opts_21         =   256
            Rows_21         =   1
            Cols_21         =   1
            URL_21          =   "C:\in_out\Software\DC-2020_ED_Analyzer_VB6\Resources\SwitchOnYellow.bmp"
            Pict_21         =   "frmMonitoring.frx":19CA5
            ColorReplaceWith_21=   8421504
            ColorReplace_21 =   8421504
            Tolerance_21    =   2
            Animator_20     =   0
            Blinker_20      =   22
            ClassName_22    =   "CCWBlinker"
            Interval_22     =   600
            EventID_22      =   7
            Obj_22          =   20
            varVarType_15   =   5
            Bindings_1      =   23
            ClassName_23    =   "CCWBindingHolderArray"
            Editor_23       =   24
            ClassName_24    =   "CCWBindingHolderArrayEditor"
            Owner_24        =   1
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
            Array[0]_1      =   13
            Array[1]_1      =   20
            Array[2]_1      =   0
            Array[3]_1      =   0
            Array[4]_1      =   27
            ClassName_27    =   "CCWDrawObj"
            opts_27         =   62
            Image_27        =   8
            Animator_27     =   0
            Blinker_27      =   0
            Array[5]_1      =   28
            ClassName_28    =   "CCWDrawObj"
            opts_28         =   62
            Image_28        =   11
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
         Begin CWUIControlsLib.CWNumEdit cwNumExcitFreq 
            Height          =   360
            Left            =   3330
            TabIndex        =   63
            ToolTipText     =   "Driving frequency, Hz, 0.01...1000 Hz (default = 5 Hz)"
            Top             =   720
            Width           =   1200
            _Version        =   196609
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
            CompatibleVers_0=   196609
            NumEdit_0       =   1
            ClassName_1     =   "CCWNumEdit"
            opts_1          =   458814
            C[0]_1          =   12648384
            C[1]_1          =   0
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
            Value_Val_1     =   200
            IncValueVarType_1=   5
            IncValue_Val_1  =   1
            AccelIncVarType_1=   5
            AccelInc_Val_1  =   10
            RangeMinVarType_1=   5
            RangeMaxVarType_1=   5
            RangeMax_Val_1  =   200
            Bindings_1      =   4
            ClassName_4     =   "CCWBindingHolderArray"
            Editor_4        =   5
            ClassName_5     =   "CCWBindingHolderArrayEditor"
            Owner_5         =   1
         End
         Begin CWUIControlsLib.CWNumEdit cwNumExcitAmpl 
            Height          =   360
            Left            =   3195
            TabIndex        =   64
            ToolTipText     =   "Driving Amplitude, 0...30000 (default = 300)"
            Top             =   1260
            Width           =   1335
            _Version        =   196609
            _ExtentX        =   2355
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
            CompatibleVers_0=   196609
            NumEdit_0       =   1
            ClassName_1     =   "CCWNumEdit"
            opts_1          =   458814
            C[0]_1          =   16777215
            C[1]_1          =   0
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
            IncValue_Val_1  =   10
            AccelIncVarType_1=   5
            AccelInc_Val_1  =   10
            RangeMinVarType_1=   5
            RangeMin_Val_1  =   -30000
            RangeMaxVarType_1=   5
            RangeMax_Val_1  =   30000
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
            Caption         =   "Output"
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
            Index           =   4
            Left            =   1980
            TabIndex        =   67
            ToolTipText     =   "Choose axis OUTPUT for excitation"
            Top             =   270
            Width           =   690
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Caption         =   "Amplitude"
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
            Index           =   5
            Left            =   1980
            TabIndex        =   66
            ToolTipText     =   "Driving Amplitude, 0...30000 (default = 300)"
            Top             =   1260
            Width           =   1335
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFC0&
            Caption         =   " Sine Freq, Hz"
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
            Index           =   6
            Left            =   1980
            TabIndex        =   65
            ToolTipText     =   "Driving frequency, Hz, 0.01...1000 Hz (default = 5 Hz)"
            Top             =   720
            Width           =   1380
         End
      End
      Begin CWUIControlsLib.CWGraph cwGraphDataAcq 
         Height          =   4905
         Left            =   2565
         TabIndex        =   29
         ToolTipText     =   "Click - show cursor, dblClick - hide cursor"
         Top             =   1755
         Width           =   8925
         _Version        =   196609
         _ExtentX        =   15743
         _ExtentY        =   8652
         _StockProps     =   71
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Reset_0         =   0   'False
         CompatibleVers_0=   196609
         Graph_0         =   1
         ClassName_1     =   "CCWGraphFrame"
         opts_1          =   46
         C[0]_1          =   16777215
         C[1]_1          =   13684944
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
         Name_5          =   "V1"
         C[0]_5          =   16755728
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
         opts_8          =   90112
         rMin_8          =   45
         rMax_8          =   566
         dMax_8          =   1000
         discInterval_8  =   1
         discBase_8      =   100
         Radial_6        =   0
         Enum_6          =   9
         ClassName_9     =   "CCWEnum"
         Editor_9        =   10
         ClassName_10    =   "CCWEnumArrayEditor"
         Owner_10        =   6
         Font_6          =   11
         ClassName_11    =   "CCWFont"
         bFont_11        =   -1  'True
         BeginProperty Font_11 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         tickopts_6      =   2751
         major_6         =   200
         minor_6         =   100
         Caption_6       =   12
         ClassName_12    =   "CCWDrawObj"
         opts_12         =   62
         C[0]_12         =   -2147483640
         Image_12        =   13
         ClassName_13    =   "CCWTextImage"
         style_13        =   1076101120
         font_13         =   0
         Animator_12     =   0
         Blinker_12      =   0
         Y_5             =   14
         ClassName_14    =   "CCWAxis"
         opts_14         =   1599
         Name_14         =   "Y"
         C[2]_14         =   16711680
         C[3]_14         =   0
         C[4]_14         =   0
         Orientation_14  =   2451
         LabelWidth_14   =   3
         format_14       =   15
         ClassName_15    =   "CCWFormat"
         Scale_14        =   16
         ClassName_16    =   "CCWScale"
         opts_16         =   122880
         rMin_16         =   26
         rMax_16         =   299
         dMax_16         =   100
         discInterval_16 =   1
         Radial_14       =   0
         Enum_14         =   17
         ClassName_17    =   "CCWEnum"
         Editor_17       =   18
         ClassName_18    =   "CCWEnumArrayEditor"
         Owner_18        =   14
         Font_14         =   19
         ClassName_19    =   "CCWFont"
         bFont_19        =   -1  'True
         BeginProperty Font_19 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         tickopts_14     =   2743
         major_14        =   10
         minor_14        =   5
         Caption_14      =   20
         ClassName_20    =   "CCWDrawObj"
         opts_20         =   62
         C[0]_20         =   16711680
         Image_20        =   21
         ClassName_21    =   "CCWTextImage"
         szText_21       =   "Pos, cts"
         style_21        =   1076101120
         font_21         =   0
         Animator_20     =   0
         Blinker_20      =   0
         LineStyle_5     =   1
         LineWidth_5     =   3
         BasePlot_5      =   0
         DefaultXInc_5   =   1
         DefaultPlotPerRow_5=   -1  'True
         Array[1]_3      =   22
         ClassName_22    =   "CCWDataPlot"
         opts_22         =   4194367
         Name_22         =   "V2"
         C[0]_22         =   4243456
         C[1]_22         =   255
         C[2]_22         =   16711680
         C[3]_22         =   16776960
         Event_22        =   2
         X_22            =   6
         Y_22            =   14
         LineStyle_22    =   1
         LineWidth_22    =   1
         BasePlot_22     =   0
         DefaultXInc_22  =   1
         DefaultPlotPerRow_22=   -1  'True
         Array[2]_3      =   23
         ClassName_23    =   "CCWDataPlot"
         opts_23         =   4194367
         Name_23         =   "V3"
         C[0]_23         =   33023
         C[1]_23         =   255
         C[2]_23         =   16711680
         C[3]_23         =   16776960
         Event_23        =   2
         X_23            =   6
         Y_23            =   14
         LineStyle_23    =   1
         LineWidth_23    =   1
         BasePlot_23     =   0
         DefaultXInc_23  =   1
         DefaultPlotPerRow_23=   -1  'True
         Array[3]_3      =   24
         ClassName_24    =   "CCWDataPlot"
         opts_24         =   4194367
         Name_24         =   "H1"
         C[0]_24         =   32896
         C[1]_24         =   255
         C[2]_24         =   16711680
         C[3]_24         =   16776960
         Event_24        =   2
         X_24            =   6
         Y_24            =   14
         LineStyle_24    =   1
         LineWidth_24    =   1
         BasePlot_24     =   0
         DefaultXInc_24  =   1
         DefaultPlotPerRow_24=   -1  'True
         Array[4]_3      =   25
         ClassName_25    =   "CCWDataPlot"
         opts_25         =   4194367
         Name_25         =   "H2"
         C[0]_25         =   16711808
         C[1]_25         =   255
         C[2]_25         =   16711680
         C[3]_25         =   16776960
         Event_25        =   2
         X_25            =   6
         Y_25            =   14
         LineStyle_25    =   1
         LineWidth_25    =   1
         BasePlot_25     =   0
         DefaultXInc_25  =   1
         DefaultPlotPerRow_25=   -1  'True
         Array[5]_3      =   26
         ClassName_26    =   "CCWDataPlot"
         opts_26         =   4194367
         Name_26         =   "H3"
         C[0]_26         =   16711935
         C[1]_26         =   255
         C[2]_26         =   16711680
         C[3]_26         =   16776960
         Event_26        =   2
         X_26            =   6
         Y_26            =   14
         LineStyle_26    =   1
         LineWidth_26    =   1
         BasePlot_26     =   0
         DefaultXInc_26  =   1
         DefaultPlotPerRow_26=   -1  'True
         Array[6]_3      =   27
         ClassName_27    =   "CCWDataPlot"
         opts_27         =   4194367
         Name_27         =   "FF_XP"
         C[0]_27         =   255
         C[1]_27         =   255
         C[2]_27         =   16711680
         C[3]_27         =   16776960
         Event_27        =   2
         X_27            =   6
         Y_27            =   14
         LineStyle_27    =   1
         LineWidth_27    =   1
         BasePlot_27     =   0
         DefaultXInc_27  =   1
         DefaultPlotPerRow_27=   -1  'True
         Array[7]_3      =   28
         ClassName_28    =   "CCWDataPlot"
         opts_28         =   4194367
         Name_28         =   "FF_YP"
         C[0]_28         =   8421376
         C[1]_28         =   255
         C[2]_28         =   16711680
         C[3]_28         =   16776960
         Event_28        =   2
         X_28            =   6
         Y_28            =   14
         LineStyle_28    =   1
         LineWidth_28    =   1
         BasePlot_28     =   0
         DefaultXInc_28  =   1
         DefaultPlotPerRow_28=   -1  'True
         Array[8]_3      =   29
         ClassName_29    =   "CCWDataPlot"
         opts_29         =   4194367
         Name_29         =   "FF_XA"
         C[0]_29         =   16711680
         C[1]_29         =   255
         C[2]_29         =   16711680
         C[3]_29         =   16776960
         Event_29        =   2
         X_29            =   6
         Y_29            =   14
         LineStyle_29    =   1
         LineWidth_29    =   1
         BasePlot_29     =   0
         DefaultXInc_29  =   1
         DefaultPlotPerRow_29=   -1  'True
         Array[9]_3      =   30
         ClassName_30    =   "CCWDataPlot"
         opts_30         =   4194367
         Name_30         =   "FF_YA"
         C[0]_30         =   12583104
         C[1]_30         =   255
         C[2]_30         =   16711680
         C[3]_30         =   16776960
         Event_30        =   2
         X_30            =   6
         Y_30            =   14
         LineStyle_30    =   1
         LineWidth_30    =   1
         BasePlot_30     =   0
         DefaultXInc_30  =   1
         DefaultPlotPerRow_30=   -1  'True
         Axes_1          =   31
         ClassName_31    =   "CCWAxes"
         Array_31        =   2
         Editor_31       =   32
         ClassName_32    =   "CCWGFAxisArrayEditor"
         Owner_32        =   1
         Array[0]_31     =   6
         Array[1]_31     =   14
         DefaultPlot_1   =   33
         ClassName_33    =   "CCWDataPlot"
         opts_33         =   4194365
         Name_33         =   "[Template]"
         C[0]_33         =   0
         C[1]_33         =   255
         C[2]_33         =   16711680
         C[3]_33         =   16776960
         Event_33        =   2
         X_33            =   6
         Y_33            =   14
         LineStyle_33    =   1
         LineWidth_33    =   1
         BasePlot_33     =   0
         DefaultXInc_33  =   1
         DefaultPlotPerRow_33=   -1  'True
         Cursors_1       =   34
         ClassName_34    =   "CCWCursors"
         Array_34        =   1
         Editor_34       =   35
         ClassName_35    =   "CCWGFCursorArrayEditor"
         Owner_35        =   1
         Array[0]_34     =   36
         ClassName_36    =   "CCWCursor"
         opts_36         =   61
         Name_36         =   "Cursor-1"
         C[0]_36         =   255
         Event_36        =   2
         X_36            =   6
         Y_36            =   14
         SnapMode_36     =   2
         ChrosshairStyle_36=   8
         PointStyle_36   =   9
         LockPlot_36     =   23
         TrackMode_1     =   2
         GraphBackground_1=   0
         GraphFrame_1    =   37
         ClassName_37    =   "CCWDrawObj"
         opts_37         =   62
         C[0]_37         =   13684944
         C[1]_37         =   13684944
         Image_37        =   38
         ClassName_38    =   "CCWPictImage"
         opts_38         =   1280
         Rows_38         =   1
         Cols_38         =   1
         F_38            =   13684944
         B_38            =   13684944
         ColorReplaceWith_38=   8421504
         ColorReplace_38 =   8421504
         Tolerance_38    =   2
         Animator_37     =   0
         Blinker_37      =   0
         PlotFrame_1     =   39
         ClassName_39    =   "CCWDrawObj"
         opts_39         =   62
         C[0]_39         =   13684944
         C[1]_39         =   16777215
         Image_39        =   40
         ClassName_40    =   "CCWPictImage"
         opts_40         =   1280
         Rows_40         =   1
         Cols_40         =   1
         Pict_40         =   1
         F_40            =   13684944
         B_40            =   16777215
         ColorReplaceWith_40=   8421504
         ColorReplace_40 =   8421504
         Tolerance_40    =   2
         Animator_39     =   0
         Blinker_39      =   0
         Caption_1       =   41
         ClassName_41    =   "CCWDrawObj"
         opts_41         =   62
         C[0]_41         =   -2147483630
         Image_41        =   42
         ClassName_42    =   "CCWTextImage"
         szText_42       =   "Position readings"
         style_42        =   1076101120
         font_42         =   0
         Animator_41     =   0
         Blinker_41      =   0
         DefaultXInc_1   =   1
         DefaultPlotPerRow_1=   -1  'True
         Bindings_1      =   43
         ClassName_43    =   "CCWBindingHolderArray"
         Editor_43       =   44
         ClassName_44    =   "CCWBindingHolderArrayEditor"
         Owner_44        =   1
      End
      Begin VB.Frame FrameTimeControl 
         BackColor       =   &H00D0D0D0&
         Height          =   1140
         Left            =   2565
         TabIndex        =   54
         Top             =   6570
         Width           =   8945
         Begin VB.TextBox TextXunits 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   1485
            TabIndex        =   101
            Text            =   " sec"
            Top             =   810
            Width           =   1590
         End
         Begin VB.ComboBox Cmb_X_scale 
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
            ItemData        =   "frmMonitoring.frx":338F3
            Left            =   45
            List            =   "frmMonitoring.frx":33909
            Style           =   2  'Dropdown List
            TabIndex        =   100
            ToolTipText     =   "Choose X axis column and scale format"
            Top             =   720
            Width           =   1425
         End
         Begin CWUIControlsLib.CWSlide cwSlideDynamicZoom 
            Height          =   900
            Left            =   1485
            TabIndex        =   55
            ToolTipText     =   "When 'X scale' in 'Manual' state, adjust X coordinates to zoom in and out"
            Top             =   150
            Width           =   7320
            _Version        =   196609
            _ExtentX        =   12912
            _ExtentY        =   1587
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
            Enabled         =   0   'False
            Reset_0         =   0   'False
            CompatibleVers_0=   196609
            Slider_0        =   1
            ClassName_1     =   "CCWSlider"
            opts_1          =   2102
            SclRef_1.l      =   2
            SclRef_1.t      =   2
            SclRef_1.r      =   490
            SclRef_1.b      =   62
            Scl_1.l         =   6
            Scl_1.t         =   2
            Scl_1.r         =   488
            Scl_1.b         =   58
            C[0]_1          =   13684944
            BGImg_1         =   2
            ClassName_2     =   "CCWDrawObj"
            opts_2          =   62
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
            style_1         =   3
            Label_1         =   6
            ClassName_6     =   "CCWDrawObj"
            opts_6          =   54
            SclRef_6.l      =   2
            SclRef_6.t      =   2
            SclRef_6.r      =   490
            SclRef_6.b      =   62
            Scl_6.l         =   202
            Scl_6.t         =   45
            Scl_6.r         =   323
            Scl_6.b         =   61
            C[0]_6          =   -2147483640
            Image_6         =   7
            ClassName_7     =   "CCWTextImage"
            szText_7        =   "Valve current, mA"
            style_7         =   29
            font_7          =   8
            ClassName_8     =   "CCWFont"
            bFont_8         =   -1  'True
            BeginProperty Font_8 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Animator_6      =   0
            Blinker_6       =   0
            Border_1        =   9
            ClassName_9     =   "CCWDrawObj"
            opts_9          =   60
            Image_9         =   10
            ClassName_10    =   "CCWPictImage"
            opts_10         =   1280
            Rows_10         =   1
            Cols_10         =   1
            Pict_10         =   25
            F_10            =   -2147483633
            B_10            =   -2147483633
            ColorReplaceWith_10=   8421504
            ColorReplace_10 =   8421504
            Tolerance_10    =   2
            Animator_9      =   0
            Blinker_9       =   0
            FillBound_1     =   11
            ClassName_11    =   "CCWGuiObject"
            opts_11         =   60
            FillTok_1       =   12
            ClassName_12    =   "CCWGuiObject"
            opts_12         =   62
            Axis_1          =   13
            ClassName_13    =   "CCWAxis"
            opts_13         =   567
            Name_13         =   "Axis"
            SclRef_13.l     =   3
            SclRef_13.t     =   7
            SclRef_13.r     =   502
            SclRef_13.b     =   63
            Scl_13.l        =   18
            Scl_13.t        =   2
            Scl_13.r        =   487
            Scl_13.b        =   54
            C[3]_13         =   0
            Orientation_13  =   548
            format_13       =   14
            ClassName_14    =   "CCWFormat"
            Scale_13        =   15
            ClassName_15    =   "CCWScale"
            opts_15         =   90112
            rMin_15         =   18
            rMax_15         =   470
            dMax_15         =   10
            discInterval_15 =   1
            Radial_13       =   0
            Enum_13         =   16
            ClassName_16    =   "CCWEnum"
            Editor_16       =   17
            ClassName_17    =   "CCWEnumArrayEditor"
            Owner_17        =   13
            Font_13         =   0
            tickopts_13     =   2738
            major_13        =   1
            minor_13        =   0.5
            Caption_13      =   18
            ClassName_18    =   "CCWDrawObj"
            opts_18         =   62
            C[0]_18         =   -2147483640
            Image_18        =   19
            ClassName_19    =   "CCWTextImage"
            style_19        =   10
            font_19         =   0
            Animator_18     =   0
            Blinker_18      =   0
            DrawLst_1       =   20
            ClassName_20    =   "CDrawList"
            count_20        =   11
            list[11]_20     =   9
            list[10]_20     =   21
            ClassName_21    =   "CCWThumb"
            opts_21         =   55
            Name_21         =   "Pointer-1"
            SclRef_21.l     =   2
            SclRef_21.t     =   15
            SclRef_21.r     =   642
            SclRef_21.b     =   72
            Scl_21.l        =   162
            Scl_21.t        =   45
            Scl_21.r        =   182
            Scl_21.b        =   67
            C[0]_21         =   16776960
            C[1]_21         =   16776960
            C[2]_21         =   -2147483635
            Image_21        =   22
            ClassName_22    =   "CCWPictImage"
            opts_22         =   1280
            Rows_22         =   1
            Cols_22         =   1
            Pict_22         =   95
            F_22            =   16776960
            B_22            =   16776960
            ColorReplaceWith_22=   8421504
            ColorReplace_22 =   8421504
            Tolerance_22    =   2
            Animator_21     =   0
            Blinker_21      =   0
            style_21        =   3
            Value_21        =   1
            FillStyle_21    =   3
            list[9]_20      =   23
            ClassName_23    =   "CCWThumb"
            opts_23         =   1048631
            Name_23         =   "Pointer-2"
            SclRef_23.l     =   2
            SclRef_23.t     =   15
            SclRef_23.r     =   642
            SclRef_23.b     =   72
            Scl_23.l        =   318
            Scl_23.t        =   22
            Scl_23.r        =   338
            Scl_23.b        =   47
            C[0]_23         =   16711935
            C[1]_23         =   16711935
            C[2]_23         =   -2147483635
            Image_23        =   24
            ClassName_24    =   "CCWPictImage"
            opts_24         =   1280
            Rows_24         =   1
            Cols_24         =   1
            Pict_24         =   96
            F_24            =   16711935
            B_24            =   16711935
            ColorReplaceWith_24=   8421504
            ColorReplace_24 =   8421504
            Tolerance_24    =   2
            Animator_23     =   0
            Blinker_23      =   0
            style_23        =   2
            Value_23        =   5
            FillStyle_23    =   4
            list[8]_20      =   13
            list[7]_20      =   6
            list[6]_20      =   12
            list[5]_20      =   4
            list[4]_20      =   25
            ClassName_25    =   "CCWDrawObj"
            opts_25         =   54
            SclRef_25.l     =   3
            SclRef_25.t     =   7
            SclRef_25.r     =   502
            SclRef_25.b     =   63
            Scl_25.l        =   18
            Scl_25.t        =   23
            Scl_25.r        =   487
            Scl_25.b        =   43
            C[0]_25         =   16777160
            Image_25        =   26
            ClassName_26    =   "CCWPictImage"
            opts_26         =   1280
            Rows_26         =   1
            Cols_26         =   1
            Pict_26         =   7
            F_26            =   16777160
            B_26            =   -2147483633
            ColorReplaceWith_26=   8421504
            ColorReplace_26 =   8421504
            Tolerance_26    =   2
            Animator_25     =   0
            Blinker_25      =   0
            list[3]_20      =   27
            ClassName_27    =   "CCWDrawObj"
            opts_27         =   54
            SclRef_27.l     =   3
            SclRef_27.t     =   7
            SclRef_27.r     =   508
            SclRef_27.b     =   57
            Scl_27.l        =   5
            Scl_27.t        =   17
            Scl_27.r        =   18
            Scl_27.b        =   45
            C[0]_27         =   16776960
            Image_27        =   28
            ClassName_28    =   "CCWPictImage"
            opts_28         =   1280
            Rows_28         =   1
            Cols_28         =   1
            Pict_28         =   93
            F_28            =   16776960
            B_28            =   -2147483633
            ColorReplaceWith_28=   8421504
            ColorReplace_28 =   8421504
            Tolerance_28    =   2
            Animator_27     =   0
            Blinker_27      =   0
            list[2]_20      =   29
            ClassName_29    =   "CCWDrawObj"
            opts_29         =   54
            SclRef_29.l     =   3
            SclRef_29.t     =   7
            SclRef_29.r     =   508
            SclRef_29.b     =   57
            Scl_29.l        =   493
            Scl_29.t        =   17
            Scl_29.r        =   507
            Scl_29.b        =   47
            C[0]_29         =   255
            Image_29        =   30
            ClassName_30    =   "CCWPictImage"
            opts_30         =   1280
            Rows_30         =   1
            Cols_30         =   1
            Pict_30         =   94
            F_30            =   255
            B_30            =   -2147483633
            ColorReplaceWith_30=   8421504
            ColorReplace_30 =   8421504
            Tolerance_30    =   2
            Animator_29     =   0
            Blinker_29      =   0
            list[1]_20      =   2
            Ptrs_1          =   31
            ClassName_31    =   "CCWPointerArray"
            Array_31        =   2
            Editor_31       =   32
            ClassName_32    =   "CCWPointerArrayEditor"
            Owner_32        =   1
            Array[0]_31     =   21
            Array[1]_31     =   23
            Bindings_1      =   33
            ClassName_33    =   "CCWBindingHolderArray"
            Editor_33       =   34
            ClassName_34    =   "CCWBindingHolderArrayEditor"
            Owner_34        =   1
            Stats_1         =   35
            ClassName_35    =   "CCWStats"
            doInc_1         =   29
            doDec_1         =   27
            doFrame_1       =   25
         End
         Begin CWUIControlsLib.CWButton cwBut_Xautoscale 
            Height          =   615
            Left            =   45
            TabIndex        =   56
            ToolTipText     =   "Horizontal Scale toggle"
            Top             =   135
            Width           =   1230
            _Version        =   196609
            _ExtentX        =   2170
            _ExtentY        =   1085
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
            CompatibleVers_0=   196609
            Boolean_0       =   1
            ClassName_1     =   "CCWBoolean"
            opts_1          =   2614
            SclRef_1.l      =   2
            SclRef_1.t      =   2
            SclRef_1.r      =   84
            SclRef_1.b      =   43
            Scl_1.l         =   7
            Scl_1.t         =   16
            Scl_1.r         =   77
            Scl_1.b         =   44
            C[0]_1          =   13684944
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
            SclRef_5.l      =   7
            SclRef_5.t      =   16
            SclRef_5.r      =   77
            SclRef_5.b      =   44
            Scl_5.l         =   21
            Scl_5.r         =   67
            Scl_5.b         =   16
            C[0]_5          =   -2147483641
            C[1]_5          =   -2147483641
            Image_5         =   6
            ClassName_6     =   "CCWTextImage"
            szText_6        =   "X scale"
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
            szText_8        =   "Auto"
            font_8          =   0
            Animator_7      =   0
            Blinker_7       =   0
            list[2]_4       =   9
            ClassName_9     =   "CCWDrawObj"
            opts_9          =   54
            SclRef_9.l      =   7
            SclRef_9.t      =   16
            SclRef_9.r      =   77
            SclRef_9.b      =   41
            Scl_9.l         =   19
            Scl_9.t         =   24
            Scl_9.r         =   66
            Scl_9.b         =   40
            C[0]_9          =   16777215
            C[1]_9          =   16777215
            Image_9         =   10
            ClassName_10    =   "CCWTextImage"
            szText_10       =   "Manual"
            font_10         =   0
            Animator_9      =   0
            Blinker_9       =   0
            list[1]_4       =   11
            ClassName_11    =   "CCWDrawObj"
            opts_11         =   62
            C[0]_11         =   32768
            C[1]_11         =   32768
            Image_11        =   12
            ClassName_12    =   "CCWPictImage"
            opts_12         =   1280
            Rows_12         =   1
            Cols_12         =   1
            Pict_12         =   44
            F_12            =   32768
            B_12            =   32768
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
            SclRef_15.l     =   7
            SclRef_15.t     =   16
            SclRef_15.r     =   77
            SclRef_15.b     =   44
            Scl_15.l        =   21
            Scl_15.r        =   67
            Scl_15.b        =   16
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
            SclRef_17.l     =   7
            SclRef_17.t     =   16
            SclRef_17.r     =   77
            SclRef_17.b     =   41
            Scl_17.l        =   19
            Scl_17.t        =   24
            Scl_17.r        =   66
            Scl_17.b        =   40
            C[0]_17         =   16777215
            C[1]_17         =   16777215
            Image_17        =   10
            Animator_17     =   0
            Blinker_17      =   0
            list[1]_14      =   18
            ClassName_18    =   "CCWDrawObj"
            opts_18         =   62
            C[0]_18         =   16761087
            C[1]_18         =   16761087
            Image_18        =   19
            ClassName_19    =   "CCWPictImage"
            opts_19         =   1280
            Rows_19         =   1
            Cols_19         =   1
            Pict_19         =   42
            F_19            =   16761087
            B_19            =   16761087
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
      Begin VB.OptionButton optChooseCh 
         BackColor       =   &H00FF00FF&
         Caption         =   "FF Yacc"
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   10
         Left            =   135
         Style           =   1  'Graphical
         TabIndex        =   53
         ToolTipText     =   "Click to highlight this channel"
         Top             =   7560
         Width           =   800
      End
      Begin VB.OptionButton optChooseCh 
         BackColor       =   &H00FF0000&
         Caption         =   "FF Xacc"
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   9
         Left            =   135
         Style           =   1  'Graphical
         TabIndex        =   52
         ToolTipText     =   "Click to highlight this channel"
         Top             =   7200
         Width           =   800
      End
      Begin VB.OptionButton optChooseCh 
         BackColor       =   &H0000C000&
         Caption         =   "FF Ypos"
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   8
         Left            =   135
         Style           =   1  'Graphical
         TabIndex        =   51
         ToolTipText     =   "Click to highlight this channel"
         Top             =   6840
         Width           =   800
      End
      Begin VB.OptionButton optChooseCh 
         BackColor       =   &H000000FF&
         Caption         =   "FF Xpos"
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   7
         Left            =   135
         Style           =   1  'Graphical
         TabIndex        =   50
         ToolTipText     =   "Click to highlight this channel"
         Top             =   6480
         Width           =   800
      End
      Begin VB.OptionButton optChooseCh 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Prox V1"
         Height          =   285
         Index           =   1
         Left            =   135
         Style           =   1  'Graphical
         TabIndex        =   49
         ToolTipText     =   "Click to highlight this channel"
         Top             =   4140
         Value           =   -1  'True
         Width           =   800
      End
      Begin VB.OptionButton optChooseCh 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Prox V2"
         Height          =   285
         Index           =   2
         Left            =   135
         Style           =   1  'Graphical
         TabIndex        =   48
         ToolTipText     =   "Click to highlight this channel"
         Top             =   4500
         Width           =   800
      End
      Begin VB.OptionButton optChooseCh 
         BackColor       =   &H000080FF&
         Caption         =   "Prox V3"
         Height          =   285
         Index           =   3
         Left            =   135
         Style           =   1  'Graphical
         TabIndex        =   47
         ToolTipText     =   "Click to highlight this channel"
         Top             =   4860
         Width           =   800
      End
      Begin VB.OptionButton optChooseCh 
         BackColor       =   &H0000C080&
         Caption         =   "Prox H1"
         Height          =   285
         Index           =   4
         Left            =   135
         Style           =   1  'Graphical
         TabIndex        =   46
         ToolTipText     =   "Click to highlight this channel"
         Top             =   5310
         Width           =   800
      End
      Begin VB.OptionButton optChooseCh 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Prox H2"
         Height          =   285
         Index           =   5
         Left            =   135
         Style           =   1  'Graphical
         TabIndex        =   45
         ToolTipText     =   "Click to highlight this channel"
         Top             =   5670
         Width           =   800
      End
      Begin VB.OptionButton optChooseCh 
         BackColor       =   &H00FFC0FF&
         Caption         =   "Prox H3"
         Height          =   285
         Index           =   6
         Left            =   135
         Style           =   1  'Graphical
         TabIndex        =   44
         ToolTipText     =   "Click to highlight this channel"
         Top             =   6030
         Width           =   800
      End
      Begin VB.CheckBox chkPOSvisible 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Index           =   0
         Left            =   855
         TabIndex        =   28
         ToolTipText     =   "Check to plot this sensor"
         Top             =   4140
         Value           =   1  'Checked
         Width           =   535
      End
      Begin VB.CheckBox chkPOSvisible 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFC0&
         Height          =   285
         Index           =   1
         Left            =   855
         TabIndex        =   27
         ToolTipText     =   "Check to plot this sensor"
         Top             =   4500
         Value           =   1  'Checked
         Width           =   435
      End
      Begin VB.CheckBox chkPOSvisible 
         Alignment       =   1  'Right Justify
         BackColor       =   &H000080FF&
         Height          =   285
         Index           =   2
         Left            =   855
         TabIndex        =   26
         ToolTipText     =   "Check to plot this sensor"
         Top             =   4860
         Value           =   1  'Checked
         Width           =   435
      End
      Begin VB.CheckBox chkPOSvisible 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0000C080&
         Height          =   285
         Index           =   3
         Left            =   855
         TabIndex        =   25
         ToolTipText     =   "Check to plot this sensor"
         Top             =   5310
         Value           =   1  'Checked
         Width           =   435
      End
      Begin VB.CheckBox chkPOSvisible 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFC0C0&
         Height          =   285
         Index           =   4
         Left            =   855
         TabIndex        =   24
         ToolTipText     =   "Check to plot this sensor"
         Top             =   5670
         Value           =   1  'Checked
         Width           =   435
      End
      Begin VB.CheckBox chkPOSvisible 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFC0FF&
         Height          =   285
         Index           =   5
         Left            =   855
         TabIndex        =   23
         ToolTipText     =   "Check to plot this sensor"
         Top             =   6030
         Value           =   1  'Checked
         Width           =   435
      End
      Begin VB.CheckBox chkPOSvisible 
         Alignment       =   1  'Right Justify
         BackColor       =   &H000000FF&
         Height          =   285
         Index           =   6
         Left            =   855
         TabIndex        =   22
         ToolTipText     =   "Check to plot this sensor"
         Top             =   6480
         Value           =   1  'Checked
         Width           =   435
      End
      Begin VB.CheckBox chkPOSvisible 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0000C000&
         Height          =   285
         Index           =   7
         Left            =   855
         TabIndex        =   21
         ToolTipText     =   "Check to plot this sensor"
         Top             =   6840
         Value           =   1  'Checked
         Width           =   435
      End
      Begin VB.CheckBox chkPOSvisible 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FF0000&
         Height          =   285
         Index           =   8
         Left            =   855
         TabIndex        =   20
         ToolTipText     =   "Check to plot this sensor"
         Top             =   7200
         Value           =   1  'Checked
         Width           =   435
      End
      Begin VB.CheckBox chkPOSvisible 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FF00FF&
         Height          =   285
         Index           =   9
         Left            =   855
         TabIndex        =   19
         ToolTipText     =   "Check to plot this sensor"
         Top             =   7560
         Value           =   1  'Checked
         Width           =   435
      End
      Begin VB.PictureBox picStatistics 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         FillColor       =   &H80000004&
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   735
         Left            =   2580
         ScaleHeight     =   705
         ScaleWidth      =   8910
         TabIndex        =   2
         ToolTipText     =   "Calculated basic statistical results of the chosen channel"
         Top             =   7695
         Width           =   8945
         Begin VB.TextBox txtPk_Pk 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   7830
            TabIndex        =   10
            Text            =   "N/A"
            ToolTipText     =   "= Max Value - Min Value"
            Top             =   315
            Width           =   1005
         End
         Begin VB.TextBox txtMinSignal 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Height          =   315
            Left            =   1080
            TabIndex        =   9
            Text            =   "N/A"
            ToolTipText     =   "Minimum Value"
            Top             =   315
            Width           =   810
         End
         Begin VB.TextBox txtMaxSignal 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C0FF&
            Height          =   315
            Left            =   3060
            TabIndex        =   8
            Text            =   "N/A"
            ToolTipText     =   "Maximum Value"
            Top             =   315
            Width           =   810
         End
         Begin VB.TextBox txtMaxSignal_lineNum 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C0FF&
            Height          =   315
            Left            =   2115
            TabIndex        =   7
            Text            =   "N/A"
            ToolTipText     =   "The Index of the maximum value in the chosen channel"
            Top             =   315
            Width           =   765
         End
         Begin VB.TextBox txtMinSignal_lineNum 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Height          =   315
            Left            =   135
            TabIndex        =   6
            Text            =   "N/A"
            ToolTipText     =   "The Index of the minimum value in the chosen channel"
            Top             =   315
            Width           =   765
         End
         Begin VB.TextBox measAvg 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   4050
            TabIndex        =   5
            Text            =   "N/A"
            ToolTipText     =   "The Average of the active channel"
            Top             =   315
            Width           =   1140
         End
         Begin VB.TextBox measSNR 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   6570
            TabIndex        =   4
            Text            =   "undefined"
            ToolTipText     =   "= Average / St.Dev."
            Top             =   315
            Width           =   1140
         End
         Begin VB.TextBox measError 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   5310
            TabIndex        =   3
            Text            =   "N/A"
            ToolTipText     =   "The Error (standard deviation) of the chosen channel"
            Top             =   315
            Width           =   1140
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            Caption         =   "Pk-Pk"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   24
            Left            =   7920
            TabIndex        =   18
            ToolTipText     =   "= Max Value - Min Value"
            Top             =   90
            Width           =   855
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Caption         =   "MinVal"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   11
            Left            =   1080
            TabIndex        =   17
            ToolTipText     =   "Minimum Value"
            Top             =   45
            Width           =   795
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            Caption         =   "Max @ index"
            Height          =   255
            Index           =   7
            Left            =   2025
            TabIndex        =   16
            ToolTipText     =   "The Index of the maximum value in the chosen channel"
            Top             =   45
            Width           =   930
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C0FF&
            Caption         =   "MaxVal"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   20
            Left            =   3060
            TabIndex        =   15
            ToolTipText     =   "Maximum Value"
            Top             =   45
            Width           =   795
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H80000005&
            Caption         =   "Min @ index"
            Height          =   255
            Index           =   8
            Left            =   45
            TabIndex        =   14
            ToolTipText     =   "The Index of the minimum value in the chosen channel"
            Top             =   45
            Width           =   930
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            Caption         =   "Average"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   21
            Left            =   4230
            TabIndex        =   13
            ToolTipText     =   "The Average of the active channel"
            Top             =   45
            Width           =   855
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            Caption         =   "SNR"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   23
            Left            =   6705
            TabIndex        =   12
            ToolTipText     =   "=Average/St.Dev."
            Top             =   45
            Width           =   855
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            Caption         =   "St.Dev."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   22
            Left            =   5490
            TabIndex        =   11
            ToolTipText     =   "The Error (standard deviation) of the chosen channel"
            Top             =   45
            Width           =   855
         End
      End
      Begin VB.CommandButton cmdGetCurrentPosition 
         BackColor       =   &H0080FF80&
         Caption         =   "Get current position once"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Reads momentary proximity sensors signals"
         Top             =   7920
         Width           =   2220
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00D0D0D0&
         Height          =   2175
         Left            =   90
         TabIndex        =   72
         ToolTipText     =   "Plot controls"
         Top             =   1755
         Width           =   2895
         Begin VB.OptionButton OptGraphMode 
            Height          =   600
            Index           =   0
            Left            =   90
            Picture         =   "frmMonitoring.frx":33950
            Style           =   1  'Graphical
            TabIndex        =   77
            ToolTipText     =   "Snap cursor to chosen plot"
            Top             =   1440
            Width           =   600
         End
         Begin VB.OptionButton OptGraphMode 
            Height          =   600
            Index           =   1
            Left            =   900
            Picture         =   "frmMonitoring.frx":34ACA
            Style           =   1  'Graphical
            TabIndex        =   76
            ToolTipText     =   "Pan window"
            Top             =   1440
            Width           =   600
         End
         Begin VB.OptionButton OptGraphMode 
            Height          =   600
            Index           =   2
            Left            =   1710
            Picture         =   "frmMonitoring.frx":35B3C
            Style           =   1  'Graphical
            TabIndex        =   75
            ToolTipText     =   "Zoom window"
            Top             =   1440
            Width           =   600
         End
         Begin VB.CommandButton cmdZoom2x 
            Appearance      =   0  'Flat
            Height          =   600
            Left            =   90
            Picture         =   "frmMonitoring.frx":36BAE
            Style           =   1  'Graphical
            TabIndex        =   74
            ToolTipText     =   "When ""Y scale"" in ""Manual"" mode, Zoom IN Y scale"
            Top             =   765
            Width           =   600
         End
         Begin VB.CommandButton cmdZoom_05x 
            BackColor       =   &H00C0FFC0&
            Height          =   600
            Left            =   1710
            Picture         =   "frmMonitoring.frx":37C20
            Style           =   1  'Graphical
            TabIndex        =   73
            ToolTipText     =   "When ""Y scale"" in ""Manual"" mode, Zoom OUT Y scale"
            Top             =   765
            Width           =   600
         End
         Begin CWUIControlsLib.CWButton cwBut_Yautoscale 
            Height          =   645
            Left            =   765
            TabIndex        =   78
            ToolTipText     =   "Vertical Scale toggle"
            Top             =   765
            Width           =   900
            _Version        =   196609
            _ExtentX        =   1587
            _ExtentY        =   1138
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
            CompatibleVers_0=   196609
            Boolean_0       =   1
            ClassName_1     =   "CCWBoolean"
            opts_1          =   2614
            SclRef_1.l      =   2
            SclRef_1.t      =   2
            SclRef_1.r      =   62
            SclRef_1.b      =   45
            Scl_1.l         =   7
            Scl_1.t         =   18
            Scl_1.r         =   57
            Scl_1.b         =   45
            C[0]_1          =   13684944
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
            SclRef_5.l      =   7
            SclRef_5.t      =   18
            SclRef_5.r      =   57
            SclRef_5.b      =   43
            Scl_5.l         =   6
            Scl_5.r         =   58
            Scl_5.b         =   16
            C[0]_5          =   -2147483641
            C[1]_5          =   -2147483641
            Image_5         =   6
            ClassName_6     =   "CCWTextImage"
            szText_6        =   "Y scale"
            font_6          =   7
            ClassName_7     =   "CCWFont"
            bFont_7         =   -1  'True
            BeginProperty Font_7 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
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
            opts_8          =   60
            C[0]_8          =   0
            C[1]_8          =   0
            Image_8         =   9
            ClassName_9     =   "CCWTextImage"
            szText_9        =   "Auto"
            font_9          =   0
            Animator_8      =   0
            Blinker_8       =   0
            list[2]_4       =   10
            ClassName_10    =   "CCWDrawObj"
            opts_10         =   54
            SclRef_10.l     =   7
            SclRef_10.t     =   8
            SclRef_10.r     =   56
            SclRef_10.b     =   39
            Scl_10.l        =   15
            Scl_10.t        =   20
            Scl_10.r        =   49
            Scl_10.b        =   34
            C[0]_10         =   16777215
            C[1]_10         =   16777215
            Image_10        =   11
            ClassName_11    =   "CCWTextImage"
            szText_11       =   "Manual"
            font_11         =   0
            Animator_10     =   0
            Blinker_10      =   0
            list[1]_4       =   12
            ClassName_12    =   "CCWDrawObj"
            opts_12         =   62
            C[0]_12         =   8388608
            C[1]_12         =   8388608
            Image_12        =   13
            ClassName_13    =   "CCWPictImage"
            opts_13         =   1280
            Rows_13         =   1
            Cols_13         =   1
            Pict_13         =   44
            F_13            =   8388608
            B_13            =   8388608
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
            SclRef_16.l     =   7
            SclRef_16.t     =   18
            SclRef_16.r     =   57
            SclRef_16.b     =   43
            Scl_16.l        =   6
            Scl_16.r        =   58
            Scl_16.b        =   16
            C[0]_16         =   -2147483641
            C[1]_16         =   -2147483641
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
            opts_18         =   52
            SclRef_18.l     =   7
            SclRef_18.t     =   8
            SclRef_18.r     =   56
            SclRef_18.b     =   39
            Scl_18.l        =   15
            Scl_18.t        =   20
            Scl_18.r        =   49
            Scl_18.b        =   34
            C[0]_18         =   16777215
            C[1]_18         =   16777215
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
            Pict_20         =   42
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
            Style_1         =   7
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
         Begin CWUIControlsLib.CWNumEdit cwSpreadPlotsByY 
            Height          =   390
            Left            =   1350
            TabIndex        =   80
            ToolTipText     =   "When plotting relative format, traces spread vertically by this value"
            Top             =   225
            Width           =   1005
            _Version        =   196609
            _ExtentX        =   1773
            _ExtentY        =   688
            _StockProps     =   4
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
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
            TextAlignment_1 =   2
            format_1        =   2
            ClassName_2     =   "CCWFormat"
            scale_1         =   3
            ClassName_3     =   "CCWScale"
            dMax_3          =   10
            discInterval_3  =   100
            ValueVarType_1  =   5
            Value_Val_1     =   100
            IncValueVarType_1=   5
            IncValue_Val_1  =   100
            AccelIncVarType_1=   5
            AccelInc_Val_1  =   100
            RangeMinVarType_1=   5
            RangeMaxVarType_1=   5
            RangeMax_Val_1  =   10000
            Bindings_1      =   4
            ClassName_4     =   "CCWBindingHolderArray"
            Editor_4        =   5
            ClassName_5     =   "CCWBindingHolderArrayEditor"
            Owner_5         =   1
         End
         Begin CWUIControlsLib.CWButton cwRelAbsPosition 
            Height          =   645
            Left            =   45
            TabIndex        =   79
            ToolTipText     =   "Remove Offset of the RAW signal - subtract the FIRST reading from each consecutive measurement"
            Top             =   0
            Width           =   1230
            _Version        =   196609
            _ExtentX        =   2170
            _ExtentY        =   1138
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
            CompatibleVers_0=   196609
            Boolean_0       =   1
            ClassName_1     =   "CCWBoolean"
            opts_1          =   2614
            SclRef_1.l      =   2
            SclRef_1.t      =   2
            SclRef_1.r      =   84
            SclRef_1.b      =   45
            Scl_1.l         =   4
            Scl_1.t         =   15
            Scl_1.r         =   83
            Scl_1.b         =   43
            C[0]_1          =   16777215
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
            szText_6        =   "Plot Shown"
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
            szText_8        =   "RELATIVE"
            font_8          =   0
            Animator_7      =   0
            Blinker_7       =   0
            list[2]_4       =   9
            ClassName_9     =   "CCWDrawObj"
            opts_9          =   54
            SclRef_9.l      =   4
            SclRef_9.t      =   15
            SclRef_9.r      =   83
            SclRef_9.b      =   43
            Scl_9.l         =   16
            Scl_9.t         =   28
            Scl_9.r         =   71
            Scl_9.b         =   42
            C[0]_9          =   0
            C[1]_9          =   0
            Image_9         =   10
            ClassName_10    =   "CCWTextImage"
            szText_10       =   "ABSOLUTE"
            font_10         =   0
            Animator_9      =   0
            Blinker_9       =   0
            list[1]_4       =   11
            ClassName_11    =   "CCWDrawObj"
            opts_11         =   54
            SclRef_11.l     =   4
            SclRef_11.t     =   15
            SclRef_11.r     =   83
            SclRef_11.b     =   43
            Scl_11.l        =   7
            Scl_11.t        =   15
            Scl_11.r        =   80
            Scl_11.b        =   43
            C[0]_11         =   49344
            C[1]_11         =   49344
            Image_11        =   12
            ClassName_12    =   "CCWPictImage"
            opts_12         =   1280
            Rows_12         =   1
            Cols_12         =   1
            Pict_12         =   44
            F_12            =   49344
            B_12            =   49344
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
            C[0]_15         =   0
            C[1]_15         =   0
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
            SclRef_17.l     =   4
            SclRef_17.t     =   15
            SclRef_17.r     =   83
            SclRef_17.b     =   43
            Scl_17.l        =   16
            Scl_17.t        =   28
            Scl_17.r        =   71
            Scl_17.b        =   42
            C[0]_17         =   0
            C[1]_17         =   0
            Image_17        =   10
            Animator_17     =   0
            Blinker_17      =   0
            list[1]_14      =   18
            ClassName_18    =   "CCWDrawObj"
            opts_18         =   54
            SclRef_18.l     =   4
            SclRef_18.t     =   15
            SclRef_18.r     =   83
            SclRef_18.b     =   43
            Scl_18.l        =   7
            Scl_18.t        =   15
            Scl_18.r        =   80
            Scl_18.b        =   43
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
         Begin VB.Label lblSpreadPlots 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            Caption         =   "Spread plots by"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   645
            Left            =   1215
            TabIndex        =   81
            Top             =   0
            Width           =   1365
         End
      End
      Begin VB.Frame frameMonitoring 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Monitoring"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1770
         Left            =   2565
         TabIndex        =   68
         ToolTipText     =   "Position recording can be internal, by controller, up to 10 sec, and external, by PC, up to 1000 lines (100 - 1000 sec)"
         Top             =   -45
         Width           =   4245
         Begin VB.CheckBox ChkUntilSTOP 
            BackColor       =   &H00FF80FF&
            Caption         =   "Saving Until Stop"
            Height          =   315
            Left            =   1350
            TabIndex        =   104
            ToolTipText     =   "If CHECKED, number of records is not limited"
            Top             =   1170
            Value           =   1  'Checked
            Width           =   1590
         End
         Begin VB.CheckBox ChkExcelFriendly 
            BackColor       =   &H00C0FFC0&
            Caption         =   "Excel friendly"
            Height          =   300
            Left            =   1350
            TabIndex        =   103
            ToolTipText     =   $"frmMonitoring.frx":38C92
            Top             =   720
            Value           =   1  'Checked
            Width           =   1605
         End
         Begin CWUIControlsLib.CWNumEdit cwContCMDperiod 
            Height          =   360
            Left            =   135
            TabIndex        =   69
            ToolTipText     =   "Wait interval between repeated position requests, sec, 0.1-10"
            Top             =   585
            Width           =   990
            _Version        =   196609
            _ExtentX        =   1746
            _ExtentY        =   635
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
            Value_Val_1     =   0.2
            IncValueVarType_1=   5
            IncValue_Val_1  =   0.05
            AccelIncVarType_1=   5
            AccelInc_Val_1  =   5
            RangeMinVarType_1=   5
            RangeMin_Val_1  =   0.1
            RangeMaxVarType_1=   5
            RangeMax_Val_1  =   10
            Bindings_1      =   4
            ClassName_4     =   "CCWBindingHolderArray"
            Editor_4        =   5
            ClassName_5     =   "CCWBindingHolderArrayEditor"
            Owner_5         =   1
         End
         Begin CWUIControlsLib.CWButton cwButContinuesSend 
            Height          =   555
            Left            =   90
            TabIndex        =   70
            ToolTipText     =   "Initiates / stops continuous position monitoring"
            Top             =   1035
            Width           =   1140
            _Version        =   196609
            _ExtentX        =   2011
            _ExtentY        =   979
            _StockProps     =   68
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Reset_0         =   0   'False
            CompatibleVers_0=   196609
            Boolean_0       =   1
            ClassName_1     =   "CCWBoolean"
            opts_1          =   2614
            SclRef_1.l      =   2
            SclRef_1.t      =   2
            SclRef_1.r      =   63
            SclRef_1.b      =   36
            Scl_1.l         =   4
            Scl_1.t         =   5
            Scl_1.r         =   58
            Scl_1.b         =   36
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
            opts_7          =   52
            SclRef_7.l      =   4
            SclRef_7.t      =   5
            SclRef_7.r      =   66
            SclRef_7.b      =   36
            Scl_7.l         =   19
            Scl_7.t         =   11
            Scl_7.r         =   54
            Scl_7.b         =   27
            C[0]_7          =   0
            C[1]_7          =   0
            Image_7         =   8
            ClassName_8     =   "CCWTextImage"
            szText_8        =   "STOP"
            font_8          =   0
            Animator_7      =   0
            Blinker_7       =   0
            list[2]_4       =   9
            ClassName_9     =   "CCWDrawObj"
            opts_9          =   54
            SclRef_9.l      =   4
            SclRef_9.t      =   5
            SclRef_9.r      =   66
            SclRef_9.b      =   36
            Scl_9.l         =   16
            Scl_9.t         =   15
            Scl_9.r         =   57
            Scl_9.b         =   31
            C[0]_9          =   0
            C[1]_9          =   0
            Image_9         =   10
            ClassName_10    =   "CCWTextImage"
            szText_10       =   "START"
            font_10         =   0
            Animator_9      =   0
            Blinker_9       =   0
            list[1]_4       =   11
            ClassName_11    =   "CCWDrawObj"
            opts_11         =   54
            SclRef_11.l     =   12
            SclRef_11.t     =   15
            SclRef_11.r     =   103
            SclRef_11.b     =   43
            Scl_11.l        =   13
            Scl_11.t        =   15
            Scl_11.r        =   104
            Scl_11.b        =   39
            C[0]_11         =   49152
            C[1]_11         =   49152
            Image_11        =   12
            ClassName_12    =   "CCWPictImage"
            opts_12         =   1280
            Rows_12         =   1
            Cols_12         =   1
            Pict_12         =   44
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
            frame_13        =   286
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
            opts_16         =   54
            SclRef_16.l     =   4
            SclRef_16.t     =   5
            SclRef_16.r     =   66
            SclRef_16.b     =   36
            Scl_16.l        =   19
            Scl_16.t        =   11
            Scl_16.r        =   54
            Scl_16.b        =   27
            C[0]_16         =   0
            C[1]_16         =   0
            Image_16        =   8
            Animator_16     =   0
            Blinker_16      =   0
            list[2]_14      =   17
            ClassName_17    =   "CCWDrawObj"
            opts_17         =   52
            SclRef_17.l     =   4
            SclRef_17.t     =   5
            SclRef_17.r     =   66
            SclRef_17.b     =   36
            Scl_17.l        =   16
            Scl_17.t        =   15
            Scl_17.r        =   57
            Scl_17.b        =   31
            C[0]_17         =   0
            C[1]_17         =   0
            Image_17        =   10
            Animator_17     =   0
            Blinker_17      =   0
            list[1]_14      =   18
            ClassName_18    =   "CCWDrawObj"
            opts_18         =   54
            SclRef_18.l     =   12
            SclRef_18.t     =   15
            SclRef_18.r     =   103
            SclRef_18.b     =   43
            Scl_18.l        =   13
            Scl_18.t        =   15
            Scl_18.r        =   104
            Scl_18.b        =   39
            C[0]_18         =   6316287
            C[1]_18         =   6316287
            Image_18        =   19
            ClassName_19    =   "CCWPictImage"
            opts_19         =   1280
            Rows_19         =   1
            Cols_19         =   1
            Pict_19         =   42
            F_19            =   6316287
            B_19            =   6316287
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
            opts_22         =   54
            SclRef_22.l     =   2
            SclRef_22.t     =   2
            SclRef_22.r     =   107
            SclRef_22.b     =   43
            Scl_22.l        =   11
            Scl_22.t        =   2
            Scl_22.r        =   101
            Scl_22.b        =   40
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
         Begin CWUIControlsLib.CWNumEdit cwNumFileRecords 
            Height          =   360
            Left            =   1755
            TabIndex        =   108
            ToolTipText     =   $"frmMonitoring.frx":38D27
            Top             =   270
            Width           =   1200
            _Version        =   196609
            _ExtentX        =   2117
            _ExtentY        =   635
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
            C[0]_1          =   12648384
            C[1]_1          =   0
            TextAlignment_1 =   1
            Appearance_1    =   0
            format_1        =   2
            ClassName_2     =   "CCWFormat"
            scale_1         =   3
            ClassName_3     =   "CCWScale"
            opts_3          =   65536
            dMax_3          =   10
            discInterval_3  =   1
            ValueVarType_1  =   5
            Value_Val_1     =   80320
            IncValueVarType_1=   5
            IncValue_Val_1  =   10
            AccelIncVarType_1=   5
            AccelInc_Val_1  =   60
            RangeMinVarType_1=   5
            RangeMaxVarType_1=   5
            RangeMax_Val_1  =   100000
            Bindings_1      =   4
            ClassName_4     =   "CCWBindingHolderArray"
            Editor_4        =   5
            ClassName_5     =   "CCWBindingHolderArrayEditor"
            Owner_5         =   1
         End
         Begin VB.Label lbMissingLines 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0E0FF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
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
            Left            =   3105
            TabIndex        =   136
            ToolTipText     =   "Decrease sample interval or switch to Ethernet if lines are missing"
            Top             =   1170
            Width           =   915
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0E0FF&
            Caption         =   "missing lines"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   450
            Index           =   32
            Left            =   3105
            TabIndex        =   135
            Top             =   720
            Width           =   915
         End
         Begin VB.Label Label1 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Interval, sec"
            Height          =   255
            Index           =   33
            Left            =   135
            TabIndex        =   71
            Top             =   270
            Width           =   915
         End
      End
      Begin VB.Frame FrameInternalRecording 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Buffer Recording"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1635
         Left            =   2565
         TabIndex        =   83
         Top             =   -45
         Width           =   4245
         Begin VB.CommandButton cmdStartMotorTest 
            BackColor       =   &H008080FF&
            Caption         =   "Start motor test"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Left            =   3060
            Style           =   1  'Graphical
            TabIndex        =   102
            ToolTipText     =   "Initiates position recording into Controller and sends motor test command"
            Top             =   585
            Width           =   1005
         End
         Begin VB.CommandButton cmdGetRecordFromBuffer 
            BackColor       =   &H0000FFFF&
            Caption         =   "Get Record"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Left            =   1620
            Style           =   1  'Graphical
            TabIndex        =   85
            ToolTipText     =   "Retrieves Position monitoring data from controller (up to 200 lines)"
            Top             =   585
            Width           =   1005
         End
         Begin VB.CommandButton cmdStartPosMonitoring 
            BackColor       =   &H00C0E0FF&
            Caption         =   "Start  Recording"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Left            =   135
            Style           =   1  'Graphical
            TabIndex        =   84
            ToolTipText     =   "Initiates position recording into Controller (up to 2 seconds)"
            Top             =   585
            Width           =   1005
         End
      End
      Begin VB.Frame FrameValveTest 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Valve Test"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1770
         Left            =   2565
         TabIndex        =   90
         ToolTipText     =   "Driving current is incremented, pressure is recorded. Need input pressure ans closed volumes!"
         Top             =   -45
         Visible         =   0   'False
         Width           =   8925
         Begin VB.TextBox txtV4sn 
            Height          =   285
            Left            =   3600
            TabIndex        =   125
            Text            =   "4444"
            ToolTipText     =   "Enter Valve 4 Serial  Number"
            Top             =   1395
            Width           =   555
         End
         Begin VB.TextBox txtV3sn 
            Height          =   285
            Left            =   3600
            TabIndex        =   124
            Text            =   "3333"
            ToolTipText     =   "Enter Valve 3 Serial  Number"
            Top             =   990
            Width           =   555
         End
         Begin VB.TextBox txtV2sn 
            Height          =   285
            Left            =   3600
            TabIndex        =   123
            Text            =   "2222"
            ToolTipText     =   "Enter Valve 2 Serial  Number"
            Top             =   585
            Width           =   555
         End
         Begin VB.TextBox txtV1sn 
            Height          =   285
            Left            =   3600
            TabIndex        =   122
            Text            =   "1111"
            ToolTipText     =   "Enter Valve 1 Serial  Number"
            Top             =   180
            Width           =   555
         End
         Begin VB.CommandButton CmdShowAnalysis 
            Caption         =   "Show Valves Performance Window"
            Height          =   735
            Left            =   90
            TabIndex        =   121
            Top             =   945
            Width           =   1140
         End
         Begin CWUIControlsLib.CWButton cwButValveTest 
            Height          =   510
            Left            =   135
            TabIndex        =   91
            ToolTipText     =   "Initiates / stops valve test: Driving current is incremented, pressure is recorded. Need input pressure ans closed volumes!"
            Top             =   270
            Width           =   1050
            _Version        =   196609
            _ExtentX        =   1852
            _ExtentY        =   900
            _StockProps     =   68
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Reset_0         =   0   'False
            CompatibleVers_0=   196609
            Boolean_0       =   1
            ClassName_1     =   "CCWBoolean"
            opts_1          =   2614
            SclRef_1.l      =   2
            SclRef_1.t      =   2
            SclRef_1.r      =   63
            SclRef_1.b      =   36
            Scl_1.l         =   4
            Scl_1.t         =   5
            Scl_1.r         =   58
            Scl_1.b         =   36
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
            opts_7          =   52
            SclRef_7.l      =   4
            SclRef_7.t      =   5
            SclRef_7.r      =   66
            SclRef_7.b      =   36
            Scl_7.l         =   19
            Scl_7.t         =   11
            Scl_7.r         =   54
            Scl_7.b         =   27
            C[0]_7          =   0
            C[1]_7          =   0
            Image_7         =   8
            ClassName_8     =   "CCWTextImage"
            szText_8        =   "STOP"
            font_8          =   0
            Animator_7      =   0
            Blinker_7       =   0
            list[2]_4       =   9
            ClassName_9     =   "CCWDrawObj"
            opts_9          =   54
            SclRef_9.l      =   4
            SclRef_9.t      =   5
            SclRef_9.r      =   66
            SclRef_9.b      =   36
            Scl_9.l         =   16
            Scl_9.t         =   15
            Scl_9.r         =   57
            Scl_9.b         =   31
            C[0]_9          =   0
            C[1]_9          =   0
            Image_9         =   10
            ClassName_10    =   "CCWTextImage"
            szText_10       =   "START"
            font_10         =   0
            Animator_9      =   0
            Blinker_9       =   0
            list[1]_4       =   11
            ClassName_11    =   "CCWDrawObj"
            opts_11         =   54
            SclRef_11.l     =   12
            SclRef_11.t     =   15
            SclRef_11.r     =   103
            SclRef_11.b     =   43
            Scl_11.l        =   13
            Scl_11.t        =   15
            Scl_11.r        =   104
            Scl_11.b        =   39
            C[0]_11         =   49152
            C[1]_11         =   49152
            Image_11        =   12
            ClassName_12    =   "CCWPictImage"
            opts_12         =   1280
            Rows_12         =   1
            Cols_12         =   1
            Pict_12         =   44
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
            frame_13        =   286
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
            opts_16         =   54
            SclRef_16.l     =   4
            SclRef_16.t     =   5
            SclRef_16.r     =   66
            SclRef_16.b     =   36
            Scl_16.l        =   19
            Scl_16.t        =   11
            Scl_16.r        =   54
            Scl_16.b        =   27
            C[0]_16         =   0
            C[1]_16         =   0
            Image_16        =   8
            Animator_16     =   0
            Blinker_16      =   0
            list[2]_14      =   17
            ClassName_17    =   "CCWDrawObj"
            opts_17         =   52
            SclRef_17.l     =   4
            SclRef_17.t     =   5
            SclRef_17.r     =   66
            SclRef_17.b     =   36
            Scl_17.l        =   16
            Scl_17.t        =   15
            Scl_17.r        =   57
            Scl_17.b        =   31
            C[0]_17         =   0
            C[1]_17         =   0
            Image_17        =   10
            Animator_17     =   0
            Blinker_17      =   0
            list[1]_14      =   18
            ClassName_18    =   "CCWDrawObj"
            opts_18         =   54
            SclRef_18.l     =   12
            SclRef_18.t     =   15
            SclRef_18.r     =   103
            SclRef_18.b     =   43
            Scl_18.l        =   13
            Scl_18.t        =   15
            Scl_18.r        =   104
            Scl_18.b        =   39
            C[0]_18         =   6316287
            C[1]_18         =   6316287
            Image_18        =   19
            ClassName_19    =   "CCWPictImage"
            opts_19         =   1280
            Rows_19         =   1
            Cols_19         =   1
            Pict_19         =   42
            F_19            =   6316287
            B_19            =   6316287
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
            opts_22         =   54
            SclRef_22.l     =   2
            SclRef_22.t     =   2
            SclRef_22.r     =   107
            SclRef_22.b     =   43
            Scl_22.l        =   11
            Scl_22.t        =   2
            Scl_22.r        =   101
            Scl_22.b        =   40
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
         Begin CWUIControlsLib.CWNumEdit cwValveMeasTime 
            Height          =   300
            Left            =   2070
            TabIndex        =   92
            ToolTipText     =   "Wait interval to let pressure stabilize, sec, 0.2-2"
            Top             =   990
            Width           =   900
            _Version        =   196609
            _ExtentX        =   1587
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
            AccelInc_Val_1  =   5
            RangeMinVarType_1=   5
            RangeMin_Val_1  =   0.2
            RangeMaxVarType_1=   5
            RangeMax_Val_1  =   5
            Bindings_1      =   4
            ClassName_4     =   "CCWBindingHolderArray"
            Editor_4        =   5
            ClassName_5     =   "CCWBindingHolderArrayEditor"
            Owner_5         =   1
         End
         Begin CWUIControlsLib.CWNumEdit cwValveStartCurrent 
            Height          =   300
            Left            =   1935
            TabIndex        =   93
            ToolTipText     =   "Starting driving mA, default 70"
            Top             =   180
            Width           =   1035
            _Version        =   196609
            _ExtentX        =   1826
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
            Value_Val_1     =   70
            IncValueVarType_1=   5
            IncValue_Val_1  =   1
            AccelIncVarType_1=   5
            AccelInc_Val_1  =   2
            RangeMinVarType_1=   5
            RangeMaxVarType_1=   5
            RangeMax_Val_1  =   100
            Bindings_1      =   4
            ClassName_4     =   "CCWBindingHolderArray"
            Editor_4        =   5
            ClassName_5     =   "CCWBindingHolderArrayEditor"
            Owner_5         =   1
         End
         Begin CWUIControlsLib.CWNumEdit cwValveEndCurrent 
            Height          =   300
            Left            =   1935
            TabIndex        =   94
            ToolTipText     =   "Driving mA to stop, default +120"
            Top             =   1395
            Width           =   1035
            _Version        =   196609
            _ExtentX        =   1826
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
            Value_Val_1     =   140
            IncValueVarType_1=   5
            IncValue_Val_1  =   1
            AccelIncVarType_1=   5
            AccelInc_Val_1  =   5
            RangeMinVarType_1=   5
            RangeMaxVarType_1=   5
            RangeMax_Val_1  =   150
            Bindings_1      =   4
            ClassName_4     =   "CCWBindingHolderArray"
            Editor_4        =   5
            ClassName_5     =   "CCWBindingHolderArrayEditor"
            Owner_5         =   1
         End
         Begin CWUIControlsLib.CWNumEdit cwValveIncrementCurrent 
            Height          =   300
            Left            =   2070
            TabIndex        =   95
            ToolTipText     =   "Increment mA, default 0.5"
            Top             =   585
            Width           =   900
            _Version        =   196609
            _ExtentX        =   1587
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
            TextAlignment_1 =   1
            Appearance_1    =   0
            format_1        =   2
            ClassName_2     =   "CCWFormat"
            Format_2        =   ".#0"
            scale_1         =   3
            ClassName_3     =   "CCWScale"
            opts_3          =   65536
            dMax_3          =   10
            discInterval_3  =   1
            ValueVarType_1  =   5
            Value_Val_1     =   0.5
            IncValueVarType_1=   5
            IncValue_Val_1  =   0.1
            AccelIncVarType_1=   5
            AccelInc_Val_1  =   1
            RangeMinVarType_1=   5
            RangeMaxVarType_1=   5
            RangeMax_Val_1  =   2
            Bindings_1      =   4
            ClassName_4     =   "CCWBindingHolderArray"
            Editor_4        =   5
            ClassName_5     =   "CCWBindingHolderArrayEditor"
            Owner_5         =   1
         End
         Begin CWUIControlsLib.CWNumEdit CwUpperPSIthreshold 
            Height          =   300
            Left            =   4500
            TabIndex        =   130
            ToolTipText     =   "Upper pressure threshold to calculate slope psi. Default 40 psi"
            Top             =   630
            Width           =   900
            _Version        =   196609
            _ExtentX        =   1587
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
            Value_Val_1     =   40
            IncValueVarType_1=   5
            IncValue_Val_1  =   1
            AccelIncVarType_1=   5
            AccelInc_Val_1  =   2
            RangeMinVarType_1=   5
            RangeMin_Val_1  =   10
            RangeMaxVarType_1=   5
            RangeMax_Val_1  =   60
            Bindings_1      =   4
            ClassName_4     =   "CCWBindingHolderArray"
            Editor_4        =   5
            ClassName_5     =   "CCWBindingHolderArrayEditor"
            Owner_5         =   1
         End
         Begin CWUIControlsLib.CWNumEdit CwLowerPSIthreshold 
            Height          =   300
            Left            =   4500
            TabIndex        =   132
            ToolTipText     =   "Lower pressure threshold to calculate slope, psi. Default 20 psi"
            Top             =   1350
            Width           =   900
            _Version        =   196609
            _ExtentX        =   1587
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
            Value_Val_1     =   20
            IncValueVarType_1=   5
            IncValue_Val_1  =   1
            AccelIncVarType_1=   5
            AccelInc_Val_1  =   2
            RangeMinVarType_1=   5
            RangeMin_Val_1  =   5
            RangeMaxVarType_1=   5
            RangeMax_Val_1  =   40
            Bindings_1      =   4
            ClassName_4     =   "CCWBindingHolderArray"
            Editor_4        =   5
            ClassName_5     =   "CCWBindingHolderArrayEditor"
            Owner_5         =   1
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Upper psi"
            Height          =   210
            Index           =   31
            Left            =   4500
            TabIndex        =   134
            Top             =   405
            Width           =   915
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Lower psi"
            Height          =   210
            Index           =   30
            Left            =   4500
            TabIndex        =   133
            Top             =   1125
            Width           =   915
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Slope Analysis"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   29
            Left            =   4230
            TabIndex        =   131
            Top             =   180
            Width           =   1365
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "V4 SN"
            Height          =   210
            Index           =   28
            Left            =   3015
            TabIndex        =   129
            Top             =   1440
            Width           =   555
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "V3 SN"
            Height          =   210
            Index           =   27
            Left            =   3015
            TabIndex        =   128
            Top             =   1035
            Width           =   555
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "V2 SN"
            Height          =   210
            Index           =   26
            Left            =   3015
            TabIndex        =   127
            Top             =   630
            Width           =   555
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "V1 SN"
            Height          =   210
            Index           =   25
            Left            =   3015
            TabIndex        =   126
            Top             =   225
            Width           =   555
         End
         Begin VB.Label Label1 
            BackColor       =   &H00FFFFC0&
            Caption         =   "Interval, s"
            Height          =   255
            Index           =   2
            Left            =   1305
            TabIndex        =   99
            ToolTipText     =   "Settling time before measuring pressure"
            Top             =   990
            Width           =   735
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Caption         =   "Start mA"
            Height          =   255
            Index           =   3
            Left            =   1215
            TabIndex        =   98
            Top             =   225
            Width           =   690
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Caption         =   "Increment"
            Height          =   255
            Index           =   9
            Left            =   1260
            TabIndex        =   97
            Top             =   630
            Width           =   735
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Caption         =   "End mA"
            Height          =   255
            Index           =   10
            Left            =   1170
            TabIndex        =   96
            Top             =   1440
            Width           =   735
         End
      End
      Begin VB.Label lbPosition 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "65000"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   285
         Index           =   6
         Left            =   1575
         TabIndex        =   42
         ToolTipText     =   "Measured Stage Feed Forward X-position Signal"
         Top             =   6480
         Width           =   780
      End
      Begin VB.Label lbPosition 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "65000"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   285
         Index           =   7
         Left            =   1575
         TabIndex        =   41
         ToolTipText     =   "Measured Stage Feed Forward Y-position Signal"
         Top             =   6840
         Width           =   780
      End
      Begin VB.Label lbPosition 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "65000"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Index           =   8
         Left            =   1575
         TabIndex        =   40
         ToolTipText     =   "Measured Stage Feed Forward X-acceleration Signal"
         Top             =   7200
         Width           =   780
      End
      Begin VB.Label lbPosition 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "65000"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C000C0&
         Height          =   285
         Index           =   9
         Left            =   1575
         TabIndex        =   39
         ToolTipText     =   "Measured Stage Feed Forward Y-acceleration Signal"
         Top             =   7560
         Width           =   780
      End
      Begin VB.Label lbPosition 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "65000"
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
         Left            =   1575
         TabIndex        =   38
         ToolTipText     =   "Measured Horizontal #3 Prox Sensor Raw Signal"
         Top             =   6030
         Width           =   780
      End
      Begin VB.Label lbPosition 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "65000"
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
         Left            =   1575
         TabIndex        =   37
         ToolTipText     =   "Measured Horizontal #2 Prox Sensor Raw Signal"
         Top             =   5670
         Width           =   780
      End
      Begin VB.Label lbPosition 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "65000"
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
         TabIndex        =   36
         ToolTipText     =   "Measured Horizontal #1 Prox Sensor Raw Signal"
         Top             =   5310
         Width           =   780
      End
      Begin VB.Label lbPosition 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "65000"
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
         Left            =   1575
         TabIndex        =   35
         ToolTipText     =   "Measured Vertical #3 Prox Sensor Raw Signal"
         Top             =   4860
         Width           =   780
      End
      Begin VB.Label lbPosition 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "65000"
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
         Left            =   1575
         TabIndex        =   34
         ToolTipText     =   "Measured Vertical #2 Prox Sensor Raw Signal"
         Top             =   4500
         Width           =   780
      End
      Begin VB.Label lbPosition 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "65000"
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
         TabIndex        =   33
         ToolTipText     =   "Measured Vertical #1 Prox Sensor Raw Signal"
         Top             =   4140
         Width           =   780
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "Data"
         Height          =   240
         Index           =   0
         Left            =   1575
         TabIndex        =   32
         Top             =   3915
         Width           =   780
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "Signal"
         Height          =   240
         Index           =   18
         Left            =   135
         TabIndex        =   31
         Top             =   3915
         Width           =   690
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "Plot?"
         Height          =   240
         Index           =   1
         Left            =   900
         TabIndex        =   30
         Top             =   3915
         Width           =   420
      End
   End
   Begin MSComDlg.CommonDialog dlgCommonDialog 
      Left            =   11070
      Top             =   8640
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin CWUIControlsLib.CWSlide cwBNC0channel 
      Height          =   9195
      Left            =   11745
      TabIndex        =   107
      ToolTipText     =   "change channel to see signal description"
      Top             =   0
      Width           =   1590
      _Version        =   196609
      _ExtentX        =   2805
      _ExtentY        =   16219
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
      CompatibleVers_0=   196609
      Slider_0        =   1
      ClassName_1     =   "CCWSlider"
      opts_1          =   2102
      SclRef_1.l      =   2
      SclRef_1.t      =   2
      SclRef_1.r      =   144
      SclRef_1.b      =   555
      Scl_1.l         =   2
      Scl_1.t         =   9
      Scl_1.r         =   144
      Scl_1.b         =   545
      C[0]_1          =   16777215
      BGImg_1         =   2
      ClassName_2     =   "CCWDrawObj"
      opts_2          =   62
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
      opts_4          =   54
      SclRef_4.l      =   24
      SclRef_4.t      =   28
      SclRef_4.r      =   52
      SclRef_4.b      =   550
      Scl_4.l         =   24
      Scl_4.t         =   28
      Scl_4.r         =   52
      Scl_4.b         =   540
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
      SclRef_6.r      =   108
      SclRef_6.b      =   612
      Scl_6.l         =   20
      Scl_6.t         =   2
      Scl_6.r         =   61
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
      SclRef_12.t     =   9
      SclRef_12.r     =   108
      SclRef_12.b     =   545
      Scl_12.l        =   2
      Scl_12.t        =   24
      Scl_12.r        =   100
      Scl_12.b        =   538
      Orientation_12  =   133172
      format_12       =   13
      ClassName_13    =   "CCWFormat"
      Scale_12        =   14
      ClassName_14    =   "CCWScale"
      opts_14         =   24576
      rMin_14         =   25
      rMax_14         =   594
      dMin_14         =   1
      dMax_14         =   32
      discInterval_14 =   1
      Radial_12       =   0
      Enum_12         =   15
      ClassName_15    =   "CCWEnum"
      Array_15        =   32
      Editor_15       =   16
      ClassName_16    =   "CCWEnumArrayEditor"
      Owner_16        =   12
      Array[0]_15     =   17
      ClassName_17    =   "CCWEnumElt"
      opts_17         =   1
      Name_17         =   "ADC0"
      DrawList_17     =   0
      varVarType_17   =   5
      Array[1]_15     =   18
      ClassName_18    =   "CCWEnumElt"
      opts_18         =   1
      Name_18         =   "ADC1"
      DrawList_18     =   0
      varVarType_18   =   2
      var_Val_18      =   1
      Array[2]_15     =   19
      ClassName_19    =   "CCWEnumElt"
      opts_19         =   1
      Name_19         =   "ADC2"
      DrawList_19     =   0
      varVarType_19   =   2
      var_Val_19      =   2
      Array[3]_15     =   20
      ClassName_20    =   "CCWEnumElt"
      opts_20         =   1
      Name_20         =   "ADC3"
      DrawList_20     =   0
      varVarType_20   =   2
      var_Val_20      =   3
      Array[4]_15     =   21
      ClassName_21    =   "CCWEnumElt"
      opts_21         =   1
      Name_21         =   "ADC4"
      DrawList_21     =   0
      varVarType_21   =   2
      var_Val_21      =   4
      Array[5]_15     =   22
      ClassName_22    =   "CCWEnumElt"
      opts_22         =   1
      Name_22         =   "ADC5"
      DrawList_22     =   0
      varVarType_22   =   2
      var_Val_22      =   5
      Array[6]_15     =   23
      ClassName_23    =   "CCWEnumElt"
      opts_23         =   1
      Name_23         =   "ADC6"
      DrawList_23     =   0
      varVarType_23   =   2
      var_Val_23      =   6
      Array[7]_15     =   24
      ClassName_24    =   "CCWEnumElt"
      opts_24         =   1
      Name_24         =   "ADC7"
      DrawList_24     =   0
      varVarType_24   =   2
      var_Val_24      =   7
      Array[8]_15     =   25
      ClassName_25    =   "CCWEnumElt"
      opts_25         =   1
      Name_25         =   "ADC8"
      DrawList_25     =   0
      varVarType_25   =   2
      var_Val_25      =   8
      Array[9]_15     =   26
      ClassName_26    =   "CCWEnumElt"
      opts_26         =   1
      Name_26         =   "ADC9"
      DrawList_26     =   0
      varVarType_26   =   2
      var_Val_26      =   9
      Array[10]_15    =   27
      ClassName_27    =   "CCWEnumElt"
      opts_27         =   1
      Name_27         =   "ADC10"
      DrawList_27     =   0
      varVarType_27   =   2
      var_Val_27      =   10
      Array[11]_15    =   28
      ClassName_28    =   "CCWEnumElt"
      opts_28         =   1
      Name_28         =   "ADC11"
      DrawList_28     =   0
      varVarType_28   =   2
      var_Val_28      =   11
      Array[12]_15    =   29
      ClassName_29    =   "CCWEnumElt"
      opts_29         =   1
      Name_29         =   "ADC12"
      DrawList_29     =   0
      varVarType_29   =   2
      var_Val_29      =   12
      Array[13]_15    =   30
      ClassName_30    =   "CCWEnumElt"
      opts_30         =   1
      Name_30         =   "ADC13"
      DrawList_30     =   0
      varVarType_30   =   2
      var_Val_30      =   13
      Array[14]_15    =   31
      ClassName_31    =   "CCWEnumElt"
      opts_31         =   1
      Name_31         =   "ADC14"
      DrawList_31     =   0
      varVarType_31   =   2
      var_Val_31      =   14
      Array[15]_15    =   32
      ClassName_32    =   "CCWEnumElt"
      opts_32         =   1
      Name_32         =   "ADC15"
      DrawList_32     =   0
      varVarType_32   =   2
      var_Val_32      =   15
      Array[16]_15    =   33
      ClassName_33    =   "CCWEnumElt"
      opts_33         =   1
      Name_33         =   "ADC16"
      DrawList_33     =   0
      varVarType_33   =   2
      var_Val_33      =   16
      Array[17]_15    =   34
      ClassName_34    =   "CCWEnumElt"
      opts_34         =   1
      Name_34         =   "ADC17"
      DrawList_34     =   0
      varVarType_34   =   2
      var_Val_34      =   17
      Array[18]_15    =   35
      ClassName_35    =   "CCWEnumElt"
      opts_35         =   1
      Name_35         =   "ADC18"
      DrawList_35     =   0
      varVarType_35   =   2
      var_Val_35      =   18
      Array[19]_15    =   36
      ClassName_36    =   "CCWEnumElt"
      opts_36         =   1
      Name_36         =   "ADC19"
      DrawList_36     =   0
      varVarType_36   =   2
      var_Val_36      =   19
      Array[20]_15    =   37
      ClassName_37    =   "CCWEnumElt"
      opts_37         =   1
      Name_37         =   "ADC20"
      DrawList_37     =   0
      varVarType_37   =   2
      var_Val_37      =   20
      Array[21]_15    =   38
      ClassName_38    =   "CCWEnumElt"
      opts_38         =   1
      Name_38         =   "ADC21"
      DrawList_38     =   0
      varVarType_38   =   2
      var_Val_38      =   21
      Array[22]_15    =   39
      ClassName_39    =   "CCWEnumElt"
      opts_39         =   1
      Name_39         =   "ADC22"
      DrawList_39     =   0
      varVarType_39   =   2
      var_Val_39      =   22
      Array[23]_15    =   40
      ClassName_40    =   "CCWEnumElt"
      opts_40         =   1
      Name_40         =   "ADC23"
      DrawList_40     =   0
      varVarType_40   =   2
      var_Val_40      =   23
      Array[24]_15    =   41
      ClassName_41    =   "CCWEnumElt"
      opts_41         =   1
      Name_41         =   "ADC24"
      DrawList_41     =   0
      varVarType_41   =   2
      var_Val_41      =   24
      Array[25]_15    =   42
      ClassName_42    =   "CCWEnumElt"
      opts_42         =   1
      Name_42         =   "ADC25"
      DrawList_42     =   0
      varVarType_42   =   2
      var_Val_42      =   25
      Array[26]_15    =   43
      ClassName_43    =   "CCWEnumElt"
      opts_43         =   1
      Name_43         =   "ADC26"
      DrawList_43     =   0
      varVarType_43   =   2
      var_Val_43      =   26
      Array[27]_15    =   44
      ClassName_44    =   "CCWEnumElt"
      opts_44         =   1
      Name_44         =   "ADC27"
      DrawList_44     =   0
      varVarType_44   =   2
      var_Val_44      =   27
      Array[28]_15    =   45
      ClassName_45    =   "CCWEnumElt"
      opts_45         =   1
      Name_45         =   "ADC28"
      DrawList_45     =   0
      varVarType_45   =   2
      var_Val_45      =   28
      Array[29]_15    =   46
      ClassName_46    =   "CCWEnumElt"
      opts_46         =   1
      Name_46         =   "ADC29"
      DrawList_46     =   0
      varVarType_46   =   2
      var_Val_46      =   29
      Array[30]_15    =   47
      ClassName_47    =   "CCWEnumElt"
      opts_47         =   1
      Name_47         =   "ADC30"
      DrawList_47     =   0
      varVarType_47   =   2
      var_Val_47      =   30
      Array[31]_15    =   48
      ClassName_48    =   "CCWEnumElt"
      opts_48         =   1
      Name_48         =   "ADC31"
      DrawList_48     =   0
      varVarType_48   =   2
      var_Val_48      =   31
      Font_12         =   0
      tickopts_12     =   2962
      Caption_12      =   49
      ClassName_49    =   "CCWDrawObj"
      opts_49         =   62
      C[0]_49         =   -2147483640
      Image_49        =   50
      ClassName_50    =   "CCWTextImage"
      font_50         =   0
      Animator_49     =   0
      Blinker_49      =   0
      DrawLst_1       =   51
      ClassName_51    =   "CDrawList"
      count_51        =   10
      list[10]_51     =   8
      list[9]_51      =   52
      ClassName_52    =   "CCWThumb"
      opts_52         =   1048631
      Name_52         =   "Pointer-1"
      SclRef_52.l     =   2
      SclRef_52.t     =   9
      SclRef_52.r     =   108
      SclRef_52.b     =   545
      Scl_52.l        =   15
      Scl_52.t        =   48
      Scl_52.r        =   55
      Scl_52.b        =   67
      C[0]_52         =   255
      C[2]_52         =   -2147483635
      Image_52        =   53
      ClassName_53    =   "CCWPictImage"
      opts_53         =   1280
      Rows_53         =   1
      Cols_53         =   1
      Pict_53         =   94
      F_53            =   255
      B_53            =   -2147483633
      ColorReplaceWith_53=   8421504
      ColorReplace_53 =   8421504
      Tolerance_53    =   2
      Animator_52     =   0
      Blinker_52      =   0
      style_52        =   3
      Value_52        =   3
      list[8]_51      =   12
      list[7]_51      =   6
      list[6]_51      =   11
      list[5]_51      =   4
      list[4]_51      =   54
      ClassName_54    =   "CCWDrawObj"
      opts_54         =   54
      SclRef_54.l     =   2
      SclRef_54.t     =   10
      SclRef_54.r     =   108
      SclRef_54.b     =   601
      Scl_54.l        =   17
      Scl_54.t        =   29
      Scl_54.r        =   49
      Scl_54.b        =   594
      C[0]_54         =   12632256
      C[1]_54         =   12632256
      Image_54        =   55
      ClassName_55    =   "CCWPictImage"
      opts_55         =   1280
      Rows_55         =   1
      Cols_55         =   1
      Pict_55         =   7
      F_55            =   12632256
      B_55            =   12632256
      ColorReplaceWith_55=   8421504
      ColorReplace_55 =   8421504
      Tolerance_55    =   2
      Animator_54     =   0
      Blinker_54      =   0
      list[3]_51      =   56
      ClassName_56    =   "CCWDrawObj"
      opts_56         =   54
      SclRef_56.l     =   2
      SclRef_56.t     =   10
      SclRef_56.r     =   108
      SclRef_56.b     =   601
      Scl_56.l        =   12
      Scl_56.t        =   596
      Scl_56.r        =   55
      Scl_56.b        =   611
      C[0]_56         =   16776960
      C[1]_56         =   12632256
      Image_56        =   57
      ClassName_57    =   "CCWPictImage"
      opts_57         =   1280
      Rows_57         =   1
      Cols_57         =   1
      Pict_57         =   96
      F_57            =   16776960
      B_57            =   12632256
      ColorReplaceWith_57=   8421504
      ColorReplace_57 =   8421504
      Tolerance_57    =   2
      Animator_56     =   0
      Blinker_56      =   0
      list[2]_51      =   58
      ClassName_58    =   "CCWDrawObj"
      opts_58         =   54
      SclRef_58.l     =   2
      SclRef_58.t     =   10
      SclRef_58.r     =   108
      SclRef_58.b     =   601
      Scl_58.l        =   11
      Scl_58.t        =   16
      Scl_58.r        =   53
      Scl_58.b        =   28
      C[0]_58         =   16711935
      C[1]_58         =   12632256
      Image_58        =   59
      ClassName_59    =   "CCWPictImage"
      opts_59         =   1280
      Rows_59         =   1
      Cols_59         =   1
      Pict_59         =   95
      F_59            =   16711935
      B_59            =   12632256
      ColorReplaceWith_59=   8421504
      ColorReplace_59 =   8421504
      Tolerance_59    =   2
      Animator_58     =   0
      Blinker_58      =   0
      list[1]_51      =   2
      Ptrs_1          =   60
      ClassName_60    =   "CCWPointerArray"
      Array_60        =   1
      Editor_60       =   61
      ClassName_61    =   "CCWPointerArrayEditor"
      Owner_61        =   1
      Array[0]_60     =   52
      Bindings_1      =   62
      ClassName_62    =   "CCWBindingHolderArray"
      Editor_62       =   63
      ClassName_63    =   "CCWBindingHolderArrayEditor"
      Owner_63        =   1
      Stats_1         =   64
      ClassName_64    =   "CCWStats"
      doInc_1         =   58
      doDec_1         =   56
      doFrame_1       =   54
   End
   Begin CWUIControlsLib.CWSlide cwBNC1channel 
      Height          =   9195
      Left            =   13365
      TabIndex        =   109
      ToolTipText     =   "change channel to see signal description"
      Top             =   0
      Width           =   1950
      _Version        =   196609
      _ExtentX        =   3440
      _ExtentY        =   16219
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
      CompatibleVers_0=   196609
      Slider_0        =   1
      ClassName_1     =   "CCWSlider"
      opts_1          =   2102
      SclRef_1.l      =   2
      SclRef_1.t      =   2
      SclRef_1.r      =   144
      SclRef_1.b      =   555
      Scl_1.l         =   2
      Scl_1.t         =   9
      Scl_1.r         =   144
      Scl_1.b         =   545
      C[0]_1          =   16777215
      BGImg_1         =   2
      ClassName_2     =   "CCWDrawObj"
      opts_2          =   62
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
      opts_4          =   54
      SclRef_4.l      =   24
      SclRef_4.t      =   28
      SclRef_4.r      =   52
      SclRef_4.b      =   550
      Scl_4.l         =   24
      Scl_4.t         =   28
      Scl_4.r         =   52
      Scl_4.b         =   540
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
      SclRef_6.b      =   612
      Scl_6.l         =   73
      Scl_6.t         =   1
      Scl_6.r         =   114
      Scl_6.b         =   17
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
      SclRef_12.t     =   9
      SclRef_12.r     =   108
      SclRef_12.b     =   545
      Scl_12.l        =   2
      Scl_12.t        =   24
      Scl_12.r        =   100
      Scl_12.b        =   538
      Orientation_12  =   133265
      format_12       =   13
      ClassName_13    =   "CCWFormat"
      Scale_12        =   14
      ClassName_14    =   "CCWScale"
      opts_14         =   24576
      rMin_14         =   25
      rMax_14         =   594
      dMin_14         =   1
      dMax_14         =   32
      discInterval_14 =   1
      Radial_12       =   0
      Enum_12         =   15
      ClassName_15    =   "CCWEnum"
      Array_15        =   32
      Editor_15       =   16
      ClassName_16    =   "CCWEnumArrayEditor"
      Owner_16        =   12
      Array[0]_15     =   17
      ClassName_17    =   "CCWEnumElt"
      opts_17         =   1
      Name_17         =   "=Geo_3Z=="
      DrawList_17     =   0
      varVarType_17   =   5
      Array[1]_15     =   18
      ClassName_18    =   "CCWEnumElt"
      opts_18         =   1
      Name_18         =   "=Geo_3H=="
      DrawList_18     =   0
      varVarType_18   =   2
      var_Val_18      =   1
      Array[2]_15     =   19
      ClassName_19    =   "CCWEnumElt"
      opts_19         =   1
      Name_19         =   "=Prx1H-lega"
      DrawList_19     =   0
      varVarType_19   =   2
      var_Val_19      =   2
      Array[3]_15     =   20
      ClassName_20    =   "CCWEnumElt"
      opts_20         =   1
      Name_20         =   "=Prx2H-lega"
      DrawList_20     =   0
      varVarType_20   =   2
      var_Val_20      =   3
      Array[4]_15     =   21
      ClassName_21    =   "CCWEnumElt"
      opts_21         =   1
      Name_21         =   "=BNC0_inp="
      DrawList_21     =   0
      varVarType_21   =   2
      var_Val_21      =   4
      Array[5]_15     =   22
      ClassName_22    =   "CCWEnumElt"
      opts_22         =   1
      Name_22         =   "===ADC_5=="
      DrawList_22     =   0
      varVarType_22   =   2
      var_Val_22      =   5
      Array[6]_15     =   23
      ClassName_23    =   "CCWEnumElt"
      opts_23         =   1
      Name_23         =   "===ADC_6=="
      DrawList_23     =   0
      varVarType_23   =   2
      var_Val_23      =   6
      Array[7]_15     =   24
      ClassName_24    =   "CCWEnumElt"
      opts_24         =   1
      Name_24         =   "=Prx3H-lega"
      DrawList_24     =   0
      varVarType_24   =   2
      var_Val_24      =   7
      Array[8]_15     =   25
      ClassName_25    =   "CCWEnumElt"
      opts_25         =   1
      Name_25         =   "=Geo_1Z=="
      DrawList_25     =   0
      varVarType_25   =   2
      var_Val_25      =   8
      Array[9]_15     =   26
      ClassName_26    =   "CCWEnumElt"
      opts_26         =   1
      Name_26         =   "=Geo_1H=="
      DrawList_26     =   0
      varVarType_26   =   2
      var_Val_26      =   9
      Array[10]_15    =   27
      ClassName_27    =   "CCWEnumElt"
      opts_27         =   1
      Name_27         =   "=Geo_2Z=="
      DrawList_27     =   0
      varVarType_27   =   2
      var_Val_27      =   10
      Array[11]_15    =   28
      ClassName_28    =   "CCWEnumElt"
      opts_28         =   1
      Name_28         =   "=Geo_2H=="
      DrawList_28     =   0
      varVarType_28   =   2
      var_Val_28      =   11
      Array[12]_15    =   29
      ClassName_29    =   "CCWEnumElt"
      opts_29         =   1
      Name_29         =   "=Prx1Z-lega"
      DrawList_29     =   0
      varVarType_29   =   2
      var_Val_29      =   12
      Array[13]_15    =   30
      ClassName_30    =   "CCWEnumElt"
      opts_30         =   1
      Name_30         =   "=Prx2Z-lega"
      DrawList_30     =   0
      varVarType_30   =   2
      var_Val_30      =   13
      Array[14]_15    =   31
      ClassName_31    =   "CCWEnumElt"
      opts_31         =   1
      Name_31         =   "=Prx3Z-lega"
      DrawList_31     =   0
      varVarType_31   =   2
      var_Val_31      =   14
      Array[15]_15    =   32
      ClassName_32    =   "CCWEnumElt"
      opts_32         =   1
      Name_32         =   "=BNC_1_inp"
      DrawList_32     =   0
      varVarType_32   =   2
      var_Val_32      =   15
      Array[16]_15    =   33
      ClassName_33    =   "CCWEnumElt"
      opts_33         =   1
      Name_33         =   "=StFF-XAcc"
      DrawList_33     =   0
      varVarType_33   =   2
      var_Val_33      =   16
      Array[17]_15    =   34
      ClassName_34    =   "CCWEnumElt"
      opts_34         =   1
      Name_34         =   "=StFF-YAcc"
      DrawList_34     =   0
      varVarType_34   =   2
      var_Val_34      =   17
      Array[18]_15    =   35
      ClassName_35    =   "CCWEnumElt"
      opts_35         =   1
      Name_35         =   "=StFF-XPos"
      DrawList_35     =   0
      varVarType_35   =   2
      var_Val_35      =   18
      Array[19]_15    =   36
      ClassName_36    =   "CCWEnumElt"
      opts_36         =   1
      Name_36         =   "=StFF-YPos"
      DrawList_36     =   0
      varVarType_36   =   2
      var_Val_36      =   19
      Array[20]_15    =   37
      ClassName_37    =   "CCWEnumElt"
      opts_37         =   1
      Name_37         =   "=Pressure_1"
      DrawList_37     =   0
      varVarType_37   =   2
      var_Val_37      =   20
      Array[21]_15    =   38
      ClassName_38    =   "CCWEnumElt"
      opts_38         =   1
      Name_38         =   "=Pressure_2"
      DrawList_38     =   0
      varVarType_38   =   2
      var_Val_38      =   21
      Array[22]_15    =   39
      ClassName_39    =   "CCWEnumElt"
      opts_39         =   1
      Name_39         =   "=Pressure_3"
      DrawList_39     =   0
      varVarType_39   =   2
      var_Val_39      =   22
      Array[23]_15    =   40
      ClassName_40    =   "CCWEnumElt"
      opts_40         =   1
      Name_40         =   "=Pressure_4"
      DrawList_40     =   0
      varVarType_40   =   2
      var_Val_40      =   23
      Array[24]_15    =   41
      ClassName_41    =   "CCWEnumElt"
      opts_41         =   1
      Name_41         =   "=In-Pressure"
      DrawList_41     =   0
      varVarType_41   =   2
      var_Val_41      =   24
      Array[25]_15    =   42
      ClassName_42    =   "CCWEnumElt"
      opts_42         =   1
      Name_42         =   "Temperature"
      DrawList_42     =   0
      varVarType_42   =   2
      var_Val_42      =   25
      Array[26]_15    =   43
      ClassName_43    =   "CCWEnumElt"
      opts_43         =   1
      Name_43         =   "=Prox_3Z=="
      DrawList_43     =   0
      varVarType_43   =   2
      var_Val_43      =   26
      Array[27]_15    =   44
      ClassName_44    =   "CCWEnumElt"
      opts_44         =   1
      Name_44         =   "=Prox_2Z=="
      DrawList_44     =   0
      varVarType_44   =   2
      var_Val_44      =   27
      Array[28]_15    =   45
      ClassName_45    =   "CCWEnumElt"
      opts_45         =   1
      Name_45         =   "=Prox_1Z=="
      DrawList_45     =   0
      varVarType_45   =   2
      var_Val_45      =   28
      Array[29]_15    =   46
      ClassName_46    =   "CCWEnumElt"
      opts_46         =   1
      Name_46         =   "=Prox_1H=="
      DrawList_46     =   0
      varVarType_46   =   2
      var_Val_46      =   29
      Array[30]_15    =   47
      ClassName_47    =   "CCWEnumElt"
      opts_47         =   1
      Name_47         =   "=Prox_2H=="
      DrawList_47     =   0
      varVarType_47   =   2
      var_Val_47      =   30
      Array[31]_15    =   48
      ClassName_48    =   "CCWEnumElt"
      opts_48         =   1
      Name_48         =   "=Prox_2H=="
      DrawList_48     =   0
      varVarType_48   =   2
      var_Val_48      =   31
      Font_12         =   0
      tickopts_12     =   2962
      Caption_12      =   49
      ClassName_49    =   "CCWDrawObj"
      opts_49         =   62
      C[0]_49         =   -2147483640
      Image_49        =   50
      ClassName_50    =   "CCWTextImage"
      font_50         =   0
      Animator_49     =   0
      Blinker_49      =   0
      DrawLst_1       =   51
      ClassName_51    =   "CDrawList"
      count_51        =   10
      list[10]_51     =   8
      list[9]_51      =   52
      ClassName_52    =   "CCWThumb"
      opts_52         =   1048631
      Name_52         =   "Pointer-1"
      SclRef_52.l     =   2
      SclRef_52.t     =   10
      SclRef_52.r     =   132
      SclRef_52.b     =   601
      Scl_52.l        =   88
      Scl_52.t        =   71
      Scl_52.r        =   128
      Scl_52.b        =   93
      C[0]_52         =   255
      C[2]_52         =   -2147483635
      Image_52        =   53
      ClassName_53    =   "CCWPictImage"
      opts_53         =   1280
      Rows_53         =   1
      Cols_53         =   1
      Pict_53         =   93
      F_53            =   255
      B_53            =   -2147483633
      ColorReplaceWith_53=   8421504
      ColorReplace_53 =   8421504
      Tolerance_53    =   2
      Animator_52     =   0
      Blinker_52      =   0
      style_52        =   2
      Value_52        =   9
      list[8]_51      =   12
      list[7]_51      =   6
      list[6]_51      =   11
      list[5]_51      =   4
      list[4]_51      =   54
      ClassName_54    =   "CCWDrawObj"
      opts_54         =   54
      SclRef_54.l     =   2
      SclRef_54.t     =   10
      SclRef_54.r     =   132
      SclRef_54.b     =   601
      Scl_54.l        =   92
      Scl_54.t        =   29
      Scl_54.r        =   124
      Scl_54.b        =   594
      C[0]_54         =   12632256
      C[1]_54         =   12632256
      Image_54        =   55
      ClassName_55    =   "CCWPictImage"
      opts_55         =   1280
      Rows_55         =   1
      Cols_55         =   1
      Pict_55         =   7
      F_55            =   12632256
      B_55            =   12632256
      ColorReplaceWith_55=   8421504
      ColorReplace_55 =   8421504
      Tolerance_55    =   2
      Animator_54     =   0
      Blinker_54      =   0
      list[3]_51      =   56
      ClassName_56    =   "CCWDrawObj"
      opts_56         =   54
      SclRef_56.l     =   2
      SclRef_56.t     =   10
      SclRef_56.r     =   132
      SclRef_56.b     =   601
      Scl_56.l        =   85
      Scl_56.t        =   597
      Scl_56.r        =   131
      Scl_56.b        =   612
      C[0]_56         =   16776960
      C[1]_56         =   12632256
      Image_56        =   57
      ClassName_57    =   "CCWPictImage"
      opts_57         =   1280
      Rows_57         =   1
      Cols_57         =   1
      Pict_57         =   96
      F_57            =   16776960
      B_57            =   12632256
      ColorReplaceWith_57=   8421504
      ColorReplace_57 =   8421504
      Tolerance_57    =   2
      Animator_56     =   0
      Blinker_56      =   0
      list[2]_51      =   58
      ClassName_58    =   "CCWDrawObj"
      opts_58         =   54
      SclRef_58.l     =   2
      SclRef_58.t     =   10
      SclRef_58.r     =   132
      SclRef_58.b     =   601
      Scl_58.l        =   83
      Scl_58.t        =   15
      Scl_58.r        =   130
      Scl_58.b        =   28
      C[0]_58         =   16711935
      C[1]_58         =   12632256
      Image_58        =   59
      ClassName_59    =   "CCWPictImage"
      opts_59         =   1280
      Rows_59         =   1
      Cols_59         =   1
      Pict_59         =   95
      F_59            =   16711935
      B_59            =   12632256
      ColorReplaceWith_59=   8421504
      ColorReplace_59 =   8421504
      Tolerance_59    =   2
      Animator_58     =   0
      Blinker_58      =   0
      list[1]_51      =   2
      Ptrs_1          =   60
      ClassName_60    =   "CCWPointerArray"
      Array_60        =   1
      Editor_60       =   61
      ClassName_61    =   "CCWPointerArrayEditor"
      Owner_61        =   1
      Array[0]_60     =   52
      Bindings_1      =   62
      ClassName_62    =   "CCWBindingHolderArray"
      Editor_62       =   63
      ClassName_63    =   "CCWBindingHolderArrayEditor"
      Owner_63        =   1
      Stats_1         =   64
      ClassName_64    =   "CCWStats"
      doInc_1         =   58
      doDec_1         =   56
      doFrame_1       =   54
   End
   Begin CWUIControlsLib.CWNumEdit cwValveHoldCurrent 
      Height          =   300
      Left            =   15975
      TabIndex        =   110
      ToolTipText     =   "ANOTHER VALVE SET driving mA, default +60"
      Top             =   5310
      Width           =   945
      _Version        =   196609
      _ExtentX        =   1667
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
      Value_Val_1     =   60
      IncValueVarType_1=   5
      IncValue_Val_1  =   1
      AccelIncVarType_1=   5
      AccelInc_Val_1  =   5
      RangeMinVarType_1=   5
      RangeMaxVarType_1=   5
      RangeMax_Val_1  =   150
      Bindings_1      =   4
      ClassName_4     =   "CCWBindingHolderArray"
      Editor_4        =   5
      ClassName_5     =   "CCWBindingHolderArrayEditor"
      Owner_5         =   1
   End
   Begin CWUIControlsLib.CWNumEdit CwOffsetCurrent 
      Height          =   300
      Index           =   0
      Left            =   16020
      TabIndex        =   111
      ToolTipText     =   "ANOTHER VALVE SET driving mA, default +60"
      Top             =   3915
      Visible         =   0   'False
      Width           =   855
      _Version        =   196609
      _ExtentX        =   1508
      _ExtentY        =   529
      _StockProps     =   4
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
      CompatibleVers_0=   196609
      NumEdit_0       =   1
      ClassName_1     =   "CCWNumEdit"
      opts_1          =   196670
      C[0]_1          =   16777215
      C[1]_1          =   0
      TextAlignment_1 =   1
      Appearance_1    =   0
      format_1        =   2
      ClassName_2     =   "CCWFormat"
      Format_2        =   ".#0"
      scale_1         =   3
      ClassName_3     =   "CCWScale"
      opts_3          =   65536
      dMax_3          =   10
      discInterval_3  =   1
      ValueVarType_1  =   5
      Value_Val_1     =   60
      IncValueVarType_1=   5
      IncValue_Val_1  =   1
      AccelIncVarType_1=   5
      AccelInc_Val_1  =   5
      RangeMinVarType_1=   5
      RangeMaxVarType_1=   5
      RangeMax_Val_1  =   150
      Bindings_1      =   4
      ClassName_4     =   "CCWBindingHolderArray"
      Editor_4        =   5
      ClassName_5     =   "CCWBindingHolderArrayEditor"
      Owner_5         =   1
   End
   Begin CWUIControlsLib.CWNumEdit CwOffsetCurrent 
      Height          =   300
      Index           =   1
      Left            =   16020
      TabIndex        =   112
      ToolTipText     =   "ANOTHER VALVE SET driving mA, default +60"
      Top             =   4230
      Visible         =   0   'False
      Width           =   855
      _Version        =   196609
      _ExtentX        =   1508
      _ExtentY        =   529
      _StockProps     =   4
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
      CompatibleVers_0=   196609
      NumEdit_0       =   1
      ClassName_1     =   "CCWNumEdit"
      opts_1          =   196670
      C[0]_1          =   16777215
      C[1]_1          =   0
      TextAlignment_1 =   1
      Appearance_1    =   0
      format_1        =   2
      ClassName_2     =   "CCWFormat"
      Format_2        =   ".#0"
      scale_1         =   3
      ClassName_3     =   "CCWScale"
      opts_3          =   65536
      dMax_3          =   10
      discInterval_3  =   1
      ValueVarType_1  =   5
      Value_Val_1     =   60
      IncValueVarType_1=   5
      IncValue_Val_1  =   1
      AccelIncVarType_1=   5
      AccelInc_Val_1  =   5
      RangeMinVarType_1=   5
      RangeMaxVarType_1=   5
      RangeMax_Val_1  =   150
      Bindings_1      =   4
      ClassName_4     =   "CCWBindingHolderArray"
      Editor_4        =   5
      ClassName_5     =   "CCWBindingHolderArrayEditor"
      Owner_5         =   1
   End
   Begin CWUIControlsLib.CWNumEdit CwOffsetCurrent 
      Height          =   300
      Index           =   2
      Left            =   16020
      TabIndex        =   113
      ToolTipText     =   "ANOTHER VALVE SET driving mA, default +60"
      Top             =   4590
      Visible         =   0   'False
      Width           =   855
      _Version        =   196609
      _ExtentX        =   1508
      _ExtentY        =   529
      _StockProps     =   4
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
      CompatibleVers_0=   196609
      NumEdit_0       =   1
      ClassName_1     =   "CCWNumEdit"
      opts_1          =   196670
      C[0]_1          =   16777215
      C[1]_1          =   0
      TextAlignment_1 =   1
      Appearance_1    =   0
      format_1        =   2
      ClassName_2     =   "CCWFormat"
      Format_2        =   ".#0"
      scale_1         =   3
      ClassName_3     =   "CCWScale"
      opts_3          =   65536
      dMax_3          =   10
      discInterval_3  =   1
      ValueVarType_1  =   5
      Value_Val_1     =   60
      IncValueVarType_1=   5
      IncValue_Val_1  =   1
      AccelIncVarType_1=   5
      AccelInc_Val_1  =   5
      RangeMinVarType_1=   5
      RangeMaxVarType_1=   5
      RangeMax_Val_1  =   150
      Bindings_1      =   4
      ClassName_4     =   "CCWBindingHolderArray"
      Editor_4        =   5
      ClassName_5     =   "CCWBindingHolderArrayEditor"
      Owner_5         =   1
   End
   Begin CWUIControlsLib.CWNumEdit CwOffsetCurrent 
      Height          =   300
      Index           =   3
      Left            =   16020
      TabIndex        =   114
      ToolTipText     =   "ANOTHER VALVE SET driving mA, default +60"
      Top             =   4950
      Visible         =   0   'False
      Width           =   855
      _Version        =   196609
      _ExtentX        =   1508
      _ExtentY        =   529
      _StockProps     =   4
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
      CompatibleVers_0=   196609
      NumEdit_0       =   1
      ClassName_1     =   "CCWNumEdit"
      opts_1          =   196670
      C[0]_1          =   16777215
      C[1]_1          =   0
      TextAlignment_1 =   1
      Appearance_1    =   0
      format_1        =   2
      ClassName_2     =   "CCWFormat"
      Format_2        =   ".#0"
      scale_1         =   3
      ClassName_3     =   "CCWScale"
      opts_3          =   65536
      dMax_3          =   10
      discInterval_3  =   1
      ValueVarType_1  =   5
      Value_Val_1     =   60
      IncValueVarType_1=   5
      IncValue_Val_1  =   1
      AccelIncVarType_1=   5
      AccelInc_Val_1  =   5
      RangeMinVarType_1=   5
      RangeMaxVarType_1=   5
      RangeMax_Val_1  =   150
      Bindings_1      =   4
      ClassName_4     =   "CCWBindingHolderArray"
      Editor_4        =   5
      ClassName_5     =   "CCWBindingHolderArrayEditor"
      Owner_5         =   1
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFC0&
      Caption         =   "Hold mA"
      Height          =   255
      Index           =   12
      Left            =   15210
      TabIndex        =   120
      ToolTipText     =   "ANOTHER VALVE SET driving mA, default +60"
      Top             =   5310
      Width           =   735
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFC0&
      Caption         =   "Deflate Offset mA"
      Height          =   255
      Index           =   13
      Left            =   15525
      TabIndex        =   119
      ToolTipText     =   "ANOTHER VALVE SET driving mA, default +60"
      Top             =   3645
      Visible         =   0   'False
      Width           =   1320
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFC0&
      Caption         =   "V1set"
      Height          =   210
      Index           =   14
      Left            =   15345
      TabIndex        =   118
      ToolTipText     =   "ANOTHER VALVE SET driving mA, default +60"
      Top             =   3915
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFC0&
      Caption         =   "V2set"
      Height          =   210
      Index           =   15
      Left            =   15345
      TabIndex        =   117
      ToolTipText     =   "ANOTHER VALVE SET driving mA, default +60"
      Top             =   4275
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFC0&
      Caption         =   "V3set"
      Height          =   210
      Index           =   16
      Left            =   15345
      TabIndex        =   116
      ToolTipText     =   "ANOTHER VALVE SET driving mA, default +60"
      Top             =   4635
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFC0&
      Caption         =   "V4set"
      Height          =   210
      Index           =   19
      Left            =   15345
      TabIndex        =   115
      ToolTipText     =   "ANOTHER VALVE SET driving mA, default +60"
      Top             =   4995
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.Label lblRawResponse 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "raw response"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   0
      TabIndex        =   43
      ToolTipText     =   "The last controller's response"
      Top             =   8865
      Width           =   11640
   End
End
Attribute VB_Name = "frmMonitoring"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Const EXCEL_SCAN_ARRAY_SIZE = 60000 'Excel can open only 65535 lines
'    Const MAX_BUFFER_ARRAY_SIZE As Long = 202
    Const MAX_SCAN_ARRAY_SIZE As Long = 100000 '86400 for 1 day, 2678400 for 31 days
    Const MAX_CH As Long = 9
    Const DataAcq_file As Long = 3
    Const PositionMonitoring_file As Long = 2
    Const end_of_ScanString_token As String = "%" & vbCrLf ' position monitoring scan data
    Const end_of_ScanTransmission As String = "OK" '  & vbCrLf position monitoring scan data END
    Const Valve_Counts_to_mA As Single = 0.00252
    Const Valve_mA_Offset As Single = 90.72

    Dim Xsize As Long ' resized form size
    Dim Ysize As Long ' resized form size
    Dim XaxisCaption(6) As String

    Dim PC_RealTime As Boolean
    Dim StartTime As Single
    Dim delayTime As Single
    Dim line_counter As Long
    Public MAXmeasINDEX As Long

    Public ASCIIDataStr As String
    Dim DataHeaderStr(10) As String
'    Public response As String
'    Dim Buffer As String
    Dim Status As Long
    Dim test_timer_ready As Boolean

    Dim CMDperiod As Double 'ususlly = cwContCMDperiod.Value
    Dim start_saving_time As String
    Private tmp_array(MAX_SCAN_ARRAY_SIZE) As Single     '  scan lines, temp array
    Dim OneScanValues(MAX_CH)    '  channel values decoded from record
    Private ScanValues(MAX_CH, MAX_SCAN_ARRAY_SIZE) As Single           ' 100000 scan lines, 10 channels array
    Dim Recorded_SCAN_ARRAY_SIZE As Long
    Private ScanTimeStamps(MAX_SCAN_ARRAY_SIZE) As Double     ' 100000 scan lines, time
    Private ScanSTARTtimeStamp As Double   ' start time of the scan, to show X TIME of the plot starting from zero
    Private MaxScanTimeStamp As Double     ' maximum duration (time) of the scan, to properly set X scale of the plot and X scale slider
    'Private TestValues(MAX_CH, MAX_SCAN_ARRAY_SIZE) As Single           ' 100000 scan lines, 10 channels array to show valve test
    Private RelativeScanValues(MAX_CH, MAX_SCAN_ARRAY_SIZE) As Single   ' 100000 scan lines, 10 channels array. The first reading is subtracted from each consecutive measurement
    Private plt_X_coord() As Double ' keeps time stamps
    Private ScanValuesToPlot() As Single   ' 10 channels array, 10000 scan lines .
    Dim ScanStDev(MAX_CH) As Double  ' to calculate how vell it holds postion
    Dim ScanAverage(MAX_CH) As Double  ' to calculate how vell it holds postion
    Dim ScanSNR(MAX_CH) As Double  ' to calculate how vell it holds postion
    Dim ScanPk_Pk(MAX_CH) As Double  ' to calculate how vell it holds postion
    Dim MinValue(MAX_CH) As Single
    Dim MaxValue(MAX_CH) As Single
    Dim MinIndex(MAX_CH) As Long
    Dim MaxIndex(MAX_CH) As Long

    ' analyze valve plot
    Dim Input_pressure_at_valves_OFF As Single
    Dim deflation_slope(0 To 3) As Single
    Dim shutdown_pressure(0 To 3) As Single
    Dim Valve_opening_mA(0 To 3) As Single
    Dim inflation_slope(0 To 3) As Single
    Dim Slope_Correction(0 To 3) As Single
    Dim fully_open_pressure(0 To 3) As Single
    Dim Input_pressure_at_valves_FULL_ON As Single
    Dim avg_slope As Single

    Public XaxisUnits As Long
    Public XminScale As Single
    Public XmaxScale As Single
    Public YminScale As Single
    Public YmaxScale As Single
    Public ScanNumber As Long ' # of scans when doing quazi-real scan
    Dim PreviousCh As Long
    Dim ChosenCh As Long
    Dim Number_of_Records As Long
    Dim Total_Number_of_Records As Long

    Dim SaveExsistingData As Boolean
    Dim SaveFileName As String
    Dim DefaultDirectory As String
    Dim Time_Start_Saving As Single  'for multiple files savings
    Dim Date_Start_Saving As Variant 'for multiple files savings
    Dim File_Saving As Boolean       'for multiple files savings
    Dim File_Reading As Boolean      'true during file reading to read time stamps instead of add time stamps
    Dim start_cmd As Single      'to monitor time it takes
    Dim interval_cmd As Single   'to monitor time it takes
    Dim User_has_been_Warned As Boolean ' about multiple files
    Dim MissingLines As Long
Public Enum TestType
    Monitor_Pos = 0
    Monitor_Pressure = 1
    Valves_2020 = 2
'    Valve_box = 3
'    Buffer_rec = 4
End Enum

Public Enum ValveTestResults
    NO_ERROR = 0
    ERR_INP_PRESSURE_TOO_LOW
    ERR_DEFL_PRESSURE_DID_NOT_CROSS_40
    ERR_DEFL_PRESSURE_DID_NOT_CROSS_20
    ERR_VALVES_WERE_NOT_COMPLITELY_SHUT
    ERR_INFL_PRESSURE_DID_NOT_CROSS_20
    ERR_INFL_PRESSURE_DID_NOT_CROSS_40
End Enum

    'program begins

Private Sub ChkExcelFriendly_Click()
    If ChkExcelFriendly = Checked Then
        If Recorded_SCAN_ARRAY_SIZE > EXCEL_SCAN_ARRAY_SIZE Then
            Recorded_SCAN_ARRAY_SIZE = EXCEL_SCAN_ARRAY_SIZE
        End If
    Else
        If Recorded_SCAN_ARRAY_SIZE > MAX_SCAN_ARRAY_SIZE Then
            Recorded_SCAN_ARRAY_SIZE = MAX_SCAN_ARRAY_SIZE
        End If
    End If
    cwNumFileRecords.Value = Recorded_SCAN_ARRAY_SIZE
End Sub

Function Conv_mAtoCounts(mA As Single) As Long
'ave_i_mA[ch] = +0.00252f * (ave_i_temp[ch]) + 90.72f;//90.72= 36000/0.00252 calibrated for Internal Valve board 09-Jul-2013
'ave_i_temp[ch] = (ave_i_mA[ch] - 90.72f)/+0.00252f  ;//90.72= 36000/0.00252 calibrated for Internal Valve board 09-Jul-2013
    Dim tfloat As Single
    tfloat = (mA - Valve_mA_Offset) / Valve_Counts_to_mA
    Conv_mAtoCounts = CLng(tfloat)
End Function


Private Sub cmb_TestType_Click()
    MonitoringTestType = cmb_TestType.ListIndex
    ChangeTestType (MonitoringTestType)
End Sub


Private Sub Cmb_X_scale_Change()
    'Line Number
    'Time
    '2 nd Column
    '3 rd Column
    '4 th Column
    '5 th Column
    XaxisUnits = Cmb_X_scale.ListIndex 'column #
    If (XaxisUnits = 0) Then TextXunits.Text = " line #"
    If (XaxisUnits = 1) Then TextXunits.Text = " sec"
    If (XaxisUnits = 2) Then TextXunits.Text = " Inflation_Cnt"
    If (XaxisUnits = 3) Then TextXunits.Text = " Ehxhaust_Cnt"
    If (XaxisUnits = 4) Then TextXunits.Text = " Inflation_mA"
    If (XaxisUnits = 5) Then TextXunits.Text = " Exhaust_mA"

    cwSlideDynamicZoom.Caption = Cmb_X_scale.List(XaxisUnits) 'XaxisCaption(XaxisUnits)
    Call UpdatePositionPlot(-1) ' -1 means update whole plot, not a particular point
End Sub

Function ADC_help(Value As Variant)
Dim TipText As String
    TipText = "change channel to see signal description"
    Select Case (Value)
    Case 0:
        TipText = "ADC0, Socket 'Isolator 1', pin 16"
    Case 1:
        TipText = "ADC1, Socket 'Isolator 1', pin 17"
    Case 2:
        TipText = "ADC2, Socket 'Isolator 1', pin 18"
    Case 3:
        TipText = "ADC3, W4 jumper 1-2 (default): BNC 0 input; W4 jumper 2-3: Socket 'Isolator 1', pin 19"
    Case 4:
        TipText = "ADC4, W5 jumper 1-2 (default): Socket 'Isolator 2', pin 16; W5 jumper 2-3: BNC 0 input"
    Case 5:
        TipText = "ADC5, Socket 'Isolator 2', pin 17"
    Case 6:
        TipText = "ADC6, Socket 'Isolator 2', pin 18"
    Case 7:
        TipText = "ADC7, W6 jumper 1-2 (default): BNC 1 input; W6 jumper 2-3: Socket 'Isolator 2', pin 19"
    Case 8:
        TipText = "ADC8, Socket 'Isolator 3', pin 16"
    Case 9:
        TipText = "ADC9, Socket 'Isolator 3', pin 17"
    Case 10:
        TipText = "ADC10, Socket 'Isolator 3', pin 18"
    Case 11:
        TipText = "ADC11, Socket 'Isolator 3', pin 19"
    Case 12:
        TipText = "ADC12, Socket 'Isolator 4', pin 16"
    Case 13:
        TipText = "ADC13, Socket 'Isolator 4', pin 17"
    Case 14:
        TipText = "ADC14, Socket 'Isolator 4', pin 18"
    Case 15:
        TipText = "ADC15, W7 jumper 1-2 (default): Socket 'Isolator 4', pin 19; W7 jumper 2-3: BNC 1 input"

    Case 16:
        TipText = "ADC16, Socket 'FeedForward+I/O', pin 2"
    Case 17:
        TipText = "ADC17, Socket 'FeedForward+I/O', pin 5"
    Case 18:
        TipText = "ADC18, Socket 'FeedForward+I/O', pin 8"
    Case 19:
        TipText = "ADC19, Socket 'FeedForward+I/O', pin 11"
    Case 20:
        TipText = "ADC20, internal pressure sensor 'Isolator 1'"
    Case 21:
        TipText = "ADC21, internal pressure sensor 'Isolator 2'"
    Case 22:
        TipText = "ADC22, internal pressure sensor 'Isolator 3'"
    Case 23:
        TipText = "ADC23, internal pressure sensor 'Isolator 4'"
    Case 24:
        TipText = "ADC24, internal pressure sensor 'Input Pressure'"
    Case 25:
        TipText = "ADC25, internal Temperature sensor"
    Case 26:
        TipText = "ADC26, 'Vert Prox Green Socket', pin 2 from the right"
    Case 27:
        TipText = "ADC27, 'Vert Prox Green Socket', pin 4 from the right"
    Case 28:
        TipText = "ADC28, 'Vert Prox Green Socket', pin 6 from the right"
    Case 29:
        TipText = "ADC29, 'Horiz Prox Black Socket', pin 2(current) or 3(voltage) from the LEFT"
    Case 30:
        TipText = "ADC30, 'Horiz Prox Black Socket', pin 6(current) or 7(voltage) from the LEFT"
    Case 31:
        TipText = "ADC31, 'Horiz Prox Black Socket', pin 10(current) or 11(voltage) from the LEFT"
    End Select
    ADC_help = TipText
End Function


Private Sub cwBNC0channel_PointerValueChanged(ByVal Pointer As Long, Value As Variant)
   Dim cmd As String
   cwBNC0channel.ToolTipText = ADC_help(Value)
   cmd = "bnc0=adc" & Value
   Call Analyzer.GetSend(cmd, True)
End Sub

Private Sub cwBNC1channel_PointerValueChanged(ByVal Pointer As Long, Value As Variant)
   Dim cmd As String
   cwBNC1channel.ToolTipText = ADC_help(Value)
   cmd = "bnc1=adc" & Value
   Call Analyzer.GetSend(cmd, True)
End Sub

Private Sub cwButValveTest_ValueChanged(ByVal Value As Boolean)
    Dim command As String
    Dim capture_txt As String
    Dim scan_line As Long
    If cwButValveTest.Value = True Then
        command = "getpa"
        cmd_sent = False
        waitresponse = True
        command = Analyzer.GetSend("getpa", True) ' pressure monitoring vs time: "getpa"
        'Win-450 rev B returns: " 58.12, 57.70, 59.78, 54.18, 87.21,*"
        scan_line = ExtractDataFromScanLine("0," & command, 0) 'add fake index
        If (ScanValues(4, 0) < 80) Then
            scan_line = MsgBox("Input pressure is " & Format(ScanValues(4, 0), "0.#0") & " psi. It is recommended to" & vbCrLf & _
                               "increase input pressure > 80 psi (standard 90 psi)" & vbCrLf & _
                               "and repeat the test, continue?", _
            vbExclamation + vbYesNo, "Input pressure < 80 psi")
            If scan_line = vbNo Then 'save existing data selected
                cwButValveTest.Value = False
                Exit Sub
            End If
        End If
        cmdReadPositionFile.Enabled = False

        Call prepare_for_plotting
        cwRelAbsPosition.Value = False 'absolute position for valve test
        If (MonitoringTestType = Valves_2020) Then 'DC2020 valve test
             command = "vtst " & CStr(cwValveStartCurrent.Value) & " " & CStr(cwValveIncrementCurrent.Value) _
             & " " & CStr(cwValveEndCurrent.Value) & " " & CStr(cwValveMeasTime.Value) & " " & CStr(cwValveHoldCurrent.Value) _
             & " " & CStr(CwLowerPSIthreshold.Value) & " " & CStr(CwUpperPSIthreshold.Value)
        End If
'        If (MonitoringTestType = Valve_box) Then 'valve BOX test
'             command = "vbxt " & CStr(cwValveStartCurrent.Value) & " " & CStr(cwValveIncrementCurrent.Value) _
'             & " " & CStr(cwValveEndCurrent.Value) & " " & CStr(cwValveMeasTime.Value) & " " & CStr(cwValveHoldCurrent.Value)
'        End If
'        cmd_sent = False
    '    frmTerminal.Hide
        Analyzer.txtManualCMD.Text = command
    '    cwSaveScan.Visible = False
        'cmdReadPositionFile.Visible = False
        ScanNumber = GetPosScan(command, (cwValveMeasTime.Value + 0.1)) 'program is here during test and plotting, not during saving
        If ScanNumber > 10 Then
            cwButValveTest.Value = False
        End If
        cwSaveScan.Visible = True
        'cmdReadPositionFile.Visible = True

    Else
        cmdReadPositionFile.Enabled = True
        command = "st"
        cmd_sent = False
        Call Analyzer.GetSend(command, True)
    End If
End Sub

Private Sub ChkUntilSTOP_Click()
'    If ChkUntilSTOP.Value = Checked Then
'        cwSavingTimeMin.Enabled = False
'        cwSavingTimeHr.Enabled = False
'        cwSavingTimeDays.Enabled = False
'        cwSavingRecords.Enabled = False
'    Else
'        cwSavingTimeMin.Enabled = True
'        cwSavingTimeHr.Enabled = True
'        cwSavingTimeDays.Enabled = True
'        cwSavingRecords.Enabled = True
'    End If
End Sub

Private Sub cmdZoom_05x_Click()
 ' 0.5x zoom - reduction
    YminScale = YminScale * 2
    YmaxScale = YmaxScale * 2
    cwGraphDataAcq.Axes.Item(2).SetMinMax CSng(YminScale), CSng(YmaxScale)
End Sub

Private Sub cmdZoom2x_Click()
 ' 2x zoom - magnification
    YminScale = YminScale / 2
    YmaxScale = YmaxScale / 2
    cwGraphDataAcq.Axes.Item(2).SetMinMax CSng(YminScale), CSng(YmaxScale)
End Sub

'Private Sub cwButStartStopMotSenTest_ValueChanged(ByVal Value As Boolean)
'    Dim controller_command As String
'    controller_command = "p"
'    cmd_sent = False
'    waitresponse = True
'    response = Analyzer.GetSend(controller_command, True)
'    Init_sys_data = True
'    cwBut_Excitation.Value = False 'after pulse excitation stops
'    Init_sys_data = False
'End Sub

Private Sub cwGraphDataAcq_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If ((cwGraphDataAcq.TrackMode = 3) Or (cwGraphDataAcq.TrackMode = 6)) Then 'pan & zoom
        YminScale = CSng(CLng(cwGraphDataAcq.Axes.Item(2).Minimum / 10) * 10)
        YmaxScale = CSng(CLng(cwGraphDataAcq.Axes.Item(2).Maximum / 10) * 10)
        cwGraphDataAcq.Axes.Item(2).SetMinMax CSng(YminScale), CSng(YmaxScale)
    End If

End Sub

Private Sub CwOffsetCurrent_ValueChanged(Index As Integer, Value As Variant, PreviousValue As Variant, ByVal OutOfRange As Boolean)
'//   Set: freq=xxx.xx, Get: freq?, where xxx.xx - float
    Dim controller_command As String
    If Init_sys_data = False Then 'avoid firing control event subroutine during initialization when values assigned
        DeflationOffset_mA(Index) = CwOffsetCurrent(Index).Value
        controller_command = "vbxo" & CStr(Index + 1) & "=" & DeflationOffset_mA(Index)
        cmd_sent = False
        waitresponse = True
        response = Analyzer.GetSend(controller_command, True)
    End If
End Sub

    'program begins
Private Sub cwNumFileRecords_ValueChanged(Value As Variant, PreviousValue As Variant, ByVal OutOfRange As Boolean)
    Recorded_SCAN_ARRAY_SIZE = cwNumFileRecords.Value
End Sub

Private Sub cwValveMeasTime_ValueChanged(Value As Variant, PreviousValue As Variant, ByVal OutOfRange As Boolean)
    CMDperiod = Value '
End Sub

Private Sub Form_Load()
    Init_sys_data = True
    User_has_been_Warned = False
    XminScale = 0
    XmaxScale = (100#)   ' set  X-time scale in sec
    PC_RealTime = False
    YminScale = -60000
    YmaxScale = 60000
    cwGraphDataAcq.Axes.Item(2).AutoScale = True
    cwGraphDataAcq.Axes.Item(2).AutoScaleNow
'    cmb_ExcitAxis.Text = "Pos Z  Valves"
    Cmb_X_scale.ListIndex = 0
    cmb_TestType.ListIndex = Monitor_Pos
    XaxisUnits = Cmb_X_scale.ListIndex
    XaxisCaption(0) = "INDEX (sample #) of measurement"
    XaxisCaption(1) = "1st Column"
    XaxisCaption(2) = "2nd Column"
    XaxisCaption(3) = "3rd Column"
    XaxisCaption(4) = "4th Column"
    XaxisCaption(5) = "Time, sec"
    cwSlideDynamicZoom.Caption = XaxisCaption(XaxisUnits)

    MAXmeasINDEX = MAX_SCAN_ARRAY_SIZE
    Me.Left = Analyzer.Left + 5000
    Me.Top = Analyzer.Top + 2500
    If FormPOSmonitoringVisble = True Then
        Call Analyzer.TogglePOSmonitoringControls
        Me.Show
    End If
    cmd_sent = False
    waitresponse = True
' instead, update controls from gloabal variables
    frmMonitoring.cwNumExcitAmpl.Value = TestGain
    frmMonitoring.cwNumExcitFreq.Value = TestFrequency
    If (ExcitationAxis < cmb_ExcitAxis.ListCount) Then
        cmb_ExcitAxis.ListIndex = ExcitationAxis
    End If
    frmMonitoring.cwBut_Excitation.Value = Excitation_status

    Init_sys_data = False
    Recorded_SCAN_ARRAY_SIZE = cwNumFileRecords.Value
'    If ChkExcelFriendly = Checked Then
'        If Recorded_SCAN_ARRAY_SIZE > EXCEL_SCAN_ARRAY_SIZE Then
'            Recorded_SCAN_ARRAY_SIZE = EXCEL_SCAN_ARRAY_SIZE
'        End If
'    Else
        If Recorded_SCAN_ARRAY_SIZE > MAX_SCAN_ARRAY_SIZE Then
            Recorded_SCAN_ARRAY_SIZE = MAX_SCAN_ARRAY_SIZE
        End If
'    End If
    cwNumFileRecords.Value = Recorded_SCAN_ARRAY_SIZE
End Sub

Private Sub Form_Resize()
    Xsize = Me.Width
    Ysize = Me.Height
    Const FormXsize As Long = 11700
    Const FormYsize As Long = 9800
    If (Xsize > FormXsize) Then
        FramePosScan.Width = Xsize - FormXsize + 12450
        cwGraphDataAcq.Width = Xsize - FormXsize + 8900
        'txtStatus.Width = Xsize - FormXsize + 11900
        'lblRawResponse.Width = Xsize - FormXsize + 11900
    End If
    If (Ysize > FormYsize) Then
        lblRawResponse.Top = Ysize - FormYsize + 8900
        picStatistics.Top = Ysize - FormYsize + 7775
        FrameTimeControl.Top = Ysize - FormYsize + 6570
        FramePosScan.Height = Ysize - FormYsize + 8600
        cwGraphDataAcq.Height = Ysize - FormYsize + 4920
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Dim tmp_stat As Boolean
  tmp_stat = Init_sys_data
    Cancel = -1
    Call Analyzer.TogglePOSmonitoringControls
    DoEvents
    Me.Hide
  Init_sys_data = tmp_stat
End Sub

Private Sub cmb_ExcitAxis_Click()
  Dim command As String
    If Init_sys_data = True Then Exit Sub
    ExcitationAxis = cmb_ExcitAxis.ListIndex
    command = "drvo=" & CStr(ExcitationAxis)
    cmd_sent = False
    waitresponse = True
    response = Analyzer.GetSend(command, True)
End Sub

Private Sub cmdGetRecordFromBuffer_Click() '*pCommandString == 'g' - get acquared data
    Dim command As String
    Dim capture_txt As String
    Dim scan_line As Long
    command = "g"
    cmd_sent = False
    cmdGetRecordFromBuffer.Enabled = False
    capture_txt = cmdGetRecordFromBuffer.Caption
    cmdGetRecordFromBuffer.Caption = "Wait for " & MAXmeasINDEX & " lines"
    If FormTerminalVisible = True Then Call Analyzer.ToggleTerminalControls
    ScanNumber = GetPosScan(command, 0.2)
    cmdGetRecordFromBuffer.Caption = capture_txt
    cmdGetRecordFromBuffer.Enabled = True
End Sub

Private Sub cmdPulse_Click()
    Dim controller_command As String
    controller_command = "p"
    cmd_sent = False
    waitresponse = True
    response = Analyzer.GetSend(controller_command, True)
    Init_sys_data = True
    Excitation_status = False  ' global excitation status for monitoring and main forms
    cwBut_Excitation.Value = Excitation_status 'after pulse excitation stops
    Analyzer.cwBut_Excitation.Value = Excitation_status

    If (ExcitationAxis < Analyzer.cmb_ExcitAxis.ListCount) Then
        Analyzer.cmb_ExcitAxis.ListIndex = ExcitationAxis
    End If
    Init_sys_data = False
End Sub


Private Sub cwBut_Excitation_ValueChanged(ByVal Value As Boolean)
'// "excitx=start"  "excitx=stop" excitation.
    If Init_sys_data = True Then Exit Sub 'avoid firing control event subroutine during initialization when values assigned
    cmd_sent = False
    If cwBut_Excitation.Value = False Then
        response = Analyzer.GetSend("excit>stop", True) 'none (cancel)
        Excitation_status = False  ' global excitation status for monitoring and main forms
    Else
        response = Analyzer.GetSend("excit>start", True) 'enable custom field excitation
        Excitation_status = True  ' global excitation status for monitoring and main forms
    End If

    'update control on Analyzer form
    Init_sys_data = True
    Analyzer.cwBut_Excitation.Value = Excitation_status
    Init_sys_data = False
End Sub

Private Sub cwNumExcitAmpl_ValueChanged(Value As Variant, PreviousValue As Variant, ByVal OutOfRange As Boolean)
'//   Set: ampl=xxx.xx, Get: ampl?, where xxx.xx - float
    Dim controller_command As String
    If Init_sys_data = True Then Exit Sub 'avoid firing control event subroutine during initialization when values assigned
    TestGain = cwNumExcitAmpl.Value
    controller_command = "ampl=" & TestGain
    cmd_sent = False
    waitresponse = True
    response = Analyzer.GetSend(controller_command, True)

    'update control on Analyzer form
    Init_sys_data = True
    ' there are two controls on Analyzer form: one on diagnostic, another on performace
    Analyzer.cwNumExcitAmpl(0).Value = TestGain
    Analyzer.cwNumExcitAmpl(1).Value = TestGain
    Init_sys_data = False
End Sub

Sub CalcNumber_ofRecords()
    Dim retval As Long
'    If File_Saving = False Then
'        cwSavingRecords.Value = Total_Number_of_Records
'    End If
    If cwSaveScan.Value = True Then 'recording into file
        If ChkUntilSTOP = Unchecked Then
            Number_of_Records = Number_of_Records - ScanNumber ' reduce recorded lines
        Else 'infinite saving
            If ScanNumber = 0 Then
                Number_of_Records = 0
            Else
                Number_of_Records = Number_of_Records + ScanNumber + 1 ' increase recorded lines
            End If
        End If
        ScanNumber = 0
    End If
End Sub

'Private Sub cwSavingTimeDays_ValueChanged(Value As Variant, PreviousValue As Variant, ByVal OutOfRange As Boolean)
'    Call CalcNumber_ofRecords
'End Sub
'Private Sub cwSavingTimeHr_ValueChanged(Value As Variant, PreviousValue As Variant, ByVal OutOfRange As Boolean)
'    Call CalcNumber_ofRecords
'End Sub
'
'Private Sub cwSavingTimeMin_ValueChanged(Value As Variant, PreviousValue As Variant, ByVal OutOfRange As Boolean)
'    Call CalcNumber_ofRecords
'End Sub

Private Sub cwNumExcitFreq_ValueChanged(Value As Variant, PreviousValue As Variant, ByVal OutOfRange As Boolean)
'//   Set: freq=xxx.xx, Get: freq?, where xxx.xx - float
    Dim controller_command As String
    If Init_sys_data = False Then 'avoid firing control event subroutine during initialization when values assigned
        TestFrequency = cwNumExcitFreq.Value
        controller_command = "freq=" & TestFrequency
        cmd_sent = False
        waitresponse = True
        response = Analyzer.GetSend(controller_command, True)
    End If

    'update control on Analyzer form
    Init_sys_data = True
    ' there are two controls on Analyzer form: one on diagnostic, another on performace
    Analyzer.cwNumExcitFreq(0).Value = TestFrequency
    Analyzer.cwNumExcitFreq(1).Value = TestFrequency
    Init_sys_data = False
End Sub

Private Sub optExitation_Click(Index As Integer)
'// "excitx=start"  "excitx=stop" excitation.
    If Init_sys_data = False Then 'avoid firing control event subroutine during initialization when values assigned
        cmd_sent = False
        If Index = 0 Then
            response = Analyzer.GetSend("freq=" & cwNumExcitFreq.Value, True)
        End If
        If Index = 1 Then
            response = Analyzer.GetSend("freq=0", True) ' start injecting noise
        End If
        lblRawResponse.Caption = response
    End If
End Sub

Private Sub cwSlideDynamicZoom_PointerValueChanged(ByVal Pointer As Long, Value As Variant)
    If cwBut_Xautoscale.Value = True Then
        Exit Sub
    Else
        If cwSlideDynamicZoom.Pointers(1).Value < cwSlideDynamicZoom.Pointers(2).Value Then
            cwSlideDynamicZoom.BackColor = &HC0FFC0      'OK, green color
            XminScale = cwSlideDynamicZoom.Pointers(1).Value ' index, for sec: Value/ cwContCMDperiod.Value
            cwGraphDataAcq.Axes.Item(1).Minimum = XminScale
            XmaxScale = cwSlideDynamicZoom.Pointers(2).Value ' index, for sec: Value / cwContCMDperiod.Value
            cwGraphDataAcq.Axes.Item(1).Maximum = XmaxScale
            cwGraphDataAcq.Refresh
        Else
            cwSlideDynamicZoom.BackColor = &H8080FF 'pink - error
        End If
    End If
End Sub

Private Sub optChooseCh_Click(Index As Integer)
    ChosenCh = Index
    If PreviousCh < 1 Then PreviousCh = 1
    Call HighlightChannel(ChosenCh, PreviousCh)
    PreviousCh = ChosenCh
End Sub

Private Sub OptGraphMode_Click(Index As Integer)
    '2-moves cursor
    '3-moves in X and Y (pan)
    '4-moves in X only
    '5-moves in Y only
    '6-cursor-selected zoom in
    '7-cursor-selected X zoom in
    '8-cursor-selected Y zoom in
    If Index = 0 Then cwGraphDataAcq.TrackMode = 2  'cursor
    If Index = 1 Then cwGraphDataAcq.TrackMode = 3  'pan
    If Index = 2 Then cwGraphDataAcq.TrackMode = 6  'zoom window
'    If Index = 3 Then
'    'Reset the axes.
'    '    cwGraphDataAcq.Axes(1).AutoScaleNow
'        cwGraphDataAcq.Axes(2).AutoScaleNow
'        cwGraphDataAcq.Axes(3).AutoScaleNow
'    End If
End Sub

Private Sub ChangeTestType(Index As Integer)
    If Index = Monitor_Pressure Or Index = Valves_2020 Then  '1, 2
'0 Monitor Pos
'1 Monitor Pressure
'2 2020 valves
'removed 3 Buffer Rec
'removed 4 Valve Box

' hide save and read buttons until test is done
'        If ScanNumber < 1 Then
'            If Index = Valves_2020 Then
'                cwSaveScan.Visible = False
'            '    cmdReadPositionFile.Visible = False
'            End If
'        End If
        optChooseCh(1).Caption = "Iso1_psi"
        optChooseCh(2).Caption = "Iso2_psi"
        optChooseCh(3).Caption = "Iso3_psi"
        optChooseCh(4).Caption = "Iso4_psi"
        optChooseCh(5).Caption = "Inp_PSI"
        optChooseCh(6).Caption = "Prox H3"
        optChooseCh(7).Caption = "FF Xpos"
        optChooseCh(8).Caption = "FF Ypos"
        optChooseCh(9).Caption = "FF Xacc"
        optChooseCh(10).Caption = "FF Yacc"
        optChooseCh(6).Visible = False
        optChooseCh(7).Visible = False
        optChooseCh(8).Visible = False
        optChooseCh(9).Visible = False
        optChooseCh(10).Visible = False
        chkPOSvisible(5).Value = Unchecked
        chkPOSvisible(6).Value = Unchecked
        chkPOSvisible(7).Value = Unchecked
        chkPOSvisible(8).Value = Unchecked
        chkPOSvisible(9).Value = Unchecked
        chkPOSvisible(5).Visible = False
        chkPOSvisible(6).Visible = False
        chkPOSvisible(7).Visible = False
        chkPOSvisible(8).Visible = False
        chkPOSvisible(9).Visible = False
        lbPosition(5).Visible = False
        lbPosition(6).Visible = False
        lbPosition(7).Visible = False
        lbPosition(8).Visible = False
        lbPosition(9).Visible = False
        cmdGetCurrentPosition.Visible = False
        cwButContinuesSend.Enabled = True
        cwButContinuesSend.Visible = True
        cwNumFileRecords.Enabled = True
        ChkExcelFriendly.Enabled = True 'allow changing array size after stop
        framExitation.Visible = True
        If Index = Monitor_Pressure Then ''pressure monitoring
            FrameValveTest.Visible = False
            frameMonitoring.Visible = True
            FrameInternalRecording.Visible = False
            Cmb_X_scale.ListIndex = 1 ''show plot versus time
            XaxisUnits = Cmb_X_scale.ListIndex
            TextXunits.Text = " sec"
        Else 'index 2 3 valve test
            FrameValveTest.Visible = True
            FrameInternalRecording.Visible = False
            frameMonitoring.Visible = False
            framExitation.Visible = False
            Cmb_X_scale.ListIndex = 4        ''show plot versus Inflation mA
            XaxisUnits = Cmb_X_scale.ListIndex
            If (XaxisUnits = 4) Then TextXunits.Text = " Inflation_mA"
        End If
    Else 'index 0 or 1, 0=Monitor Pos; 1=Buffer Rec
' show save and read buttons until test is done
        cwSaveScan.Visible = True
'        cmdReadPositionFile.Visible = True
        If Index = Monitor_Pos Then ' position or pressure monitoring
            FrameValveTest.Visible = False
            frameMonitoring.Visible = True
            FrameInternalRecording.Visible = False
            Cmb_X_scale.ListIndex = 1 ''show plot versus time
            XaxisUnits = Cmb_X_scale.ListIndex
            TextXunits.Text = " sec"
        End If
        If Index = 1 Then ' data from the buffer
            FrameValveTest.Visible = False
            frameMonitoring.Visible = False
            FrameInternalRecording.Visible = True
            Cmb_X_scale.ListIndex = 1 ''show plot versus time
            XaxisUnits = Cmb_X_scale.ListIndex
            TextXunits.Text = " sec"
        End If
        optChooseCh(1).Caption = "Prox V1"
        optChooseCh(2).Caption = "Prox V2"
        optChooseCh(3).Caption = "Prox V3"
        optChooseCh(4).Caption = "Prox H1"
        optChooseCh(5).Caption = "Prox H2"
        optChooseCh(6).Caption = "Prox H3"
        optChooseCh(7).Caption = "FF Xpos"
        optChooseCh(8).Caption = "FF Ypos"
        optChooseCh(9).Caption = "FF Xacc"
        optChooseCh(10).Caption = "FF Yacc"
        chkPOSvisible(5).Value = Checked
        chkPOSvisible(6).Value = Checked
        chkPOSvisible(7).Value = Checked
        chkPOSvisible(8).Value = Checked
        chkPOSvisible(9).Value = Checked
        optChooseCh(6).Visible = True
        optChooseCh(7).Visible = True
        optChooseCh(8).Visible = True
        optChooseCh(9).Visible = True
        optChooseCh(10).Visible = True
        chkPOSvisible(5).Visible = True
        chkPOSvisible(6).Visible = True
        chkPOSvisible(7).Visible = True
        chkPOSvisible(8).Visible = True
        chkPOSvisible(9).Visible = True
        lbPosition(5).Visible = True
        lbPosition(6).Visible = True
        lbPosition(7).Visible = True
        lbPosition(8).Visible = True
        lbPosition(9).Visible = True
        cmdGetCurrentPosition.Visible = True
        Cmb_X_scale.ListIndex = 0 ''show plot versus line#
        XaxisUnits = Cmb_X_scale.ListIndex
        If (XaxisUnits = 0) Then TextXunits.Text = " line#"
    End If
'? not here ?    cwGraphDataAcq.Axes.Item(2).AutoScaleNow
End Sub

Private Sub aTimerTestProgram_Timer()
    test_timer_ready = True
'    cwGraphDataAcq.Axes.Item(1).AutoScaleNow
'    cwGraphDataAcq.Axes.Item(2).AutoScaleNow
End Sub

Function Fill_Data_Array(ch As Long, def_value As Single) As Long
Dim lineNum As Long
    For lineNum = 0 To MAX_SCAN_ARRAY_SIZE
        ScanValues(ch, lineNum) = def_value
        RelativeScanValues(ch, lineNum) = cwSpreadPlotsByY.Value * (5 - ch)
    Next lineNum
End Function

Function ExtractDataFromScanLine(data_line As String, Optional scan_num As Long = 0) As Long
On Error GoTo format_error
' function returns number of decoded data line,
'it updates either decoded data index (for "g" -" GetArray of scans" command)
' or "scan_num" index in array (for quazi-real time continuous monitoring using "b" command)
' time is recorded instead of index
' index     V1        V2      V3      H1      H2      H3    FF_Xpos     FF_Ypos     FF_Xacc     FF_Yacc
' '12:45:51.45    -2094   -2094   -2094   -4284   -4293   -4309   4240    4176    4176    4272    %

'char * format_line = "//index,  V1,    V2,    V3,    H1,    H2,    H3,  ffXP,  ffYP,  ffXA,  ffYA%\r\n";
'0000 ,+19304,+35087,+24176,+33354,+45331,+27864,+30547,-40255,+00002,-00045.
'0001 ,+19298,+35082,+24170,+33359,+45309,+27854,+30533,-40258,+00006,-00041.
'0002 ,+19289,+35099,+24162,+33358,+45286,+27838,+30516,-40246,+00020,-00048.
'0003 ,+19245,+35104,+24155,+33384,+45255,+27816,+30501,-40244,+00036,-00043.
' or with valve test data
'line#,InfCnt,EhxCnt, Inflat_mA,Exhaust_mA, PSI_1,PSI_2,PSI_3,PSI_4,PSI_in,\r\n"
'0019 ,  -500, -6000,  90.7,     43.7,         0.2,  0.1, -0.1, -0.0,  -0.1,%
'0020 ,     0,-6000,91.2,43.7,0.2,0.1,-0.1,-0.0,-0.1,%
'0021 ,  +500,-6000,91.7,43.7,0.2,0.1,-0.1,-0.0,-0.1,%
'Test finished: *
'>~OK
' pressure monitoring vs time: "getpa"
'"Win-450 rev B" returns: " 58.12, 57.70, 59.78, 54.18, 87.21,*"

    Dim tmpDataStr As String
    Dim tmp_str As String
    Dim ch As Long
    Dim separator As Long   ' position of symbol ','
    Dim pt_separator As Long   ' position of symbol '.'
    Dim str_len As Long
    Dim tmp_time As Double

    ExtractDataFromScanLine = -2 ' -2 means error
    If (scan_num < 1000) Then
        lblRawResponse = data_line ' display current line
    Else
        If ((scan_num Mod 10) = 0) Then ' display each tenth line if there are more than 1000
            lblRawResponse = data_line
        End If
    End If
    separator = InStr(data_line, "%") ' end_of_ScanString_token)  'find "%CrLf", but CrLf got cut by GetSend
    If separator = 0 Then separator = InStr(data_line, "*")
    If (data_line <> "Timeout" And data_line <> "" And separator <> 0) Then
        tmpDataStr = Left(data_line, separator - 2)   'one line with index, 0000-9999, without last ",%"
        lbLineNum.BackColor = &HC0FFC0 'green
    Else
        If InStr(data_line, "PC-save") Then '   ", - End of PC-saved file"
            ExtractDataFromScanLine = -1 'return: not a valid line
            Exit Function
        Else
            lbLineNum.BackColor = &H8080FF 'red
            GoTo format_error
        End If
    End If
    If Len(tmpDataStr) < 2 Then
        ExtractDataFromScanLine = -1 'return: not a valid line
        Exit Function 'error
    End If
    If scan_num > -1 Then 'test mode or single reading - save time of measurement
        If File_Reading = True Then 'reading line like this "'12:45:51.457,-2094,-2094,-2094,-4284,-4293,-4309,4240,4176,4176,4272,%"
' or for valve test "0 ,-11999 , 2226 , 59.94 , 59.97 , 56.92 , 56.15 , 58.12 , 52.43 , 91.49 ,'19:20:48.133,%"
' or older format position vs time: " 1 , 191 , 304 ,-63 , 548 ,-408 ,-897 , 12528 ,-240 ,-80 ,-144 ,%"
            pt_separator = InStr(data_line, ":") ' find first semicolon in time stamp
            If (pt_separator > 0) Then ' there is time stamp
                separator = InStr(pt_separator, data_line, ",") ' find first comma AFTER ":"
            Else
                separator = InStr(data_line, ",")  ' find first comma
            End If
            If pt_separator <> 0 Then ' there is time stamp in the file
                tmp_str = Mid(data_line, pt_separator - 2, separator - pt_separator + 3) ' cut "'12:45:51.45"
                tmp_time = ConvertTimeStringToSingle(tmp_str) ' seconds from midnight
                If scan_num > 0 Then
                    If tmp_time = 0 Then  'time string '00:00:00.00
                        tmp_time = ScanTimeStamps(scan_num - 1) + CMDperiod
                    Else
                        If scan_num = 3 Then
                            If MonitoringTestType = Valves_2020 Then
                                CMDperiod = cwValveMeasTime.Value
                            Else
                                CMDperiod = ScanTimeStamps(2) - ScanTimeStamps(1) ' =cwContCMDperiod.Value
                                CMDperiod = Round(CMDperiod, 3)
                                If (cwContCMDperiod.Value <> 0) Then
                                    If ((Abs(CMDperiod - cwContCMDperiod.Value) / cwContCMDperiod.Value) < 0.1) Then
                                        CMDperiod = cwContCMDperiod.Value
                                    End If
                                End If
                                If CMDperiod < 0.02 Then PC_RealTime = False
                            End If
                        End If
                    End If
                End If
                If (scan_num = 0) Then ScanSTARTtimeStamp = (tmp_time) ' start time of the scan, to show X TIME of the plot starting from zero
                ScanTimeStamps(scan_num) = tmp_time
                If MonitoringTestType = Valves_2020 Then ' Or MonitoringTestType = Valve_box Then ' valve test
                    If (ScanSTARTtimeStamp = 0) And (scan_num > 0) Then
                        ScanTimeStamps(scan_num) = ScanTimeStamps(scan_num - 1) + CMDperiod
                    End If
                    If (Left(data_line, 1) <> "'") Then ' format like: " 0 ,-7999 , 2226 , 69.96 , 59.97 , 53.41 , 53.29 , 54.1 , 49.08 , 91.15 ,'15:29:11.902,%"
                        tmpDataStr = Left(data_line, pt_separator - 5)
                    Else 'format like: "'17:03:57.22,-02094,-02094,-02088,-04365,-02791,-02775,+04192,+04192,+04192,+04192,%"
                        tmpDataStr = Mid(data_line, separator + 1)
                        tmpDataStr = "0000," & tmpDataStr ' add fake index
                    End If
'                    tmp_str = Mid(data_line, pt_separator - 3)
'                    tmp_str = Mid(tmp_str, Len(tmp_str) - 2) ' cut ,%
                Else
                    ' for pos / pressure vs time cut begining: "'12:45:51.457,-2094,-2094,-2094,-4284,-4293,-4309,4240,4176,4176,4272,%"
                    tmp_str = Mid(data_line, separator + 1) ' cut "'12:45:51.45"
                    tmpDataStr = "0000," & tmp_str ' add fake index "0000, 58.15, 57.82, 59.74, 54.31, 87.23,*"
                End If
            Else ' no time stamp
                If (scan_num) = 0 Then
                    ScanTimeStamps(scan_num) = 0  '
                    ScanSTARTtimeStamp = 0
                Else
                    ScanTimeStamps(scan_num) = ScanTimeStamps(scan_num - 1) + CMDperiod ' create time Stamps
                End If
                If MonitoringTestType <> Monitor_Pos Then
                    tmpDataStr = Mid(data_line, separator + 1)
                    tmpDataStr = "0000," & tmpDataStr ' add fake index
                End If
            End If
        Else ' not reading - collecting data - add time stamp
            tmp_time = Timer
            If ((scan_num = 0) And (line_counter = 0)) Then ScanSTARTtimeStamp = (tmp_time) ' start time of the scan, to show X TIME of the plot starting from zero
            If SaveExsistingData = False Then ' starting new recording
            ' position data has fake index 0000 already
                If (MonitoringTestType = Monitor_Pressure) Then ' it does not have index in the first field
                    tmpDataStr = Format(ScanNumber, "0#") & "," & data_line ' add fake index
                End If
                'ScanTimeStamps(ScanNumber) = tmp_time ' ScanNumber increases when saving data DIRECTLY into file
            Else
                ScanTimeStamps(line_counter) = tmp_time 'line_counter increases when saving EXISTING data into file
            End If
        End If
    End If
    separator = InStr(tmpDataStr, ",")
    tmp_str = Left(tmpDataStr, separator - 1) ' index of line
    If (IsNumber(tmp_str)) Then
        ExtractDataFromScanLine = Val(tmp_str)  'line count starts from 0000
    Else
        ExtractDataFromScanLine = -1 'return: not a valid line
        Exit Function
    End If

    tmpDataStr = Mid(tmpDataStr, separator + 1) ' data itself, without index

process_data:
    For ch = 0 To MAX_CH 'find separator ',' 10 times
        str_len = Len(tmpDataStr) ' it can be just last "*" which will be saved as 0 in last channels
        If (str_len = 0) Then
            ExtractDataFromScanLine = -1 ' -1 means error, return: not a valid line
            Exit Function
        Else
            OneScanValues(ch) = 0 'reset value - we will write there
        End If
        separator = InStr(tmpDataStr, ",")
        If (separator <> 0) Then
            tmp_str = Left(tmpDataStr, separator - 1)
        Else
            tmp_str = tmpDataStr 'last value
        End If
        If (IsNumber(tmp_str)) Then
            OneScanValues(ch) = Val(tmp_str) '
          If scan_num = 0 Then 'if optional parameter is missing, update data array line number taken from controller (for "g" command)
            ScanValues(ch, ExtractDataFromScanLine) = Val(tmp_str)
            RelativeScanValues(ch, ExtractDataFromScanLine) = Val(tmp_str) - ScanValues(ch, 0) + cwSpreadPlotsByY.Value * (5 - ch) ' subtract the first measurement
          Else
            ScanValues(ch, scan_num) = Val(tmp_str)
            RelativeScanValues(ch, scan_num) = Val(tmp_str) - ScanValues(ch, 0) + cwSpreadPlotsByY.Value * (5 - ch) ' subtract the first measurement
          End If
        End If
        tmpDataStr = Right(tmpDataStr, str_len - separator)  ' remove the piece of string from which we've converted information
    Next ch

    ExtractDataFromScanLine = scan_num 'return:>=0 valid line
    Exit Function

format_error:
    ExtractDataFromScanLine = -1 'return: not a valid line
'    MsgBox ("Unrecognized data format")
End Function

Sub GetOneScanLine() '*pCommandString == 'b' - get one position measurement
    Dim command As String
    Dim ch As Long
    Dim scan_line As Long
    Dim response As String
    command = "b"
    cmd_sent = False
    waitresponse = True
'    If Analyzer.chkTestProgram <> Checked Then
        response = Analyzer.GetSend(command, True)
'    Else
'        response = "0000, +19304,+35087,+24176,+33354,+45331,+27864,+30547,-40255,+00002,-00045." & vbCrLf
'    End If

'    lstResponse.AddItem "=>" & response

    scan_line = ExtractDataFromScanLine(response)
'    If scan_line >= -1 Then
        Call Show_raw_values
'    End If
End Sub

Private Sub cmdGetCurrentPosition_Click()
    GetOneScanLine
End Sub

Function InitiateInternslRec(cmd As String) As String
    Dim command As String
    Dim CmdStartTime As Single
    Dim cmdAcqTime As Single
    Dim status_chk As Long
    Dim temp_str As String

    status_chk = 0
    CmdStartTime = Timer

    command = cmd
    cmd_sent = False
    waitresponse = True
    response = Analyzer.GetSend(command, True)
    command = "s" ' get status, response is " The System Status is: XX"
    waitresponse = True
    Do
        DoEvents

        aTimerTestProgram.Interval = 400 '  millisecond interval getting status
        test_timer_ready = False
        aTimerTestProgram.Enabled = True
        Do
            DoEvents
        Loop While (test_timer_ready = False)

        cmd_sent = False
        response = Analyzer.GetSend(command, True)
        Status = InStr(response, ":")
        If IsNumeric("&H" & response) Then
            temp_str = "&H" & response
            Status = CInt(temp_str)
        End If
        If ((Status <> 0) And (Len(response) = Status + 3)) Then
            temp_str = "&H" & (Right(response, 2))
            Status = CInt(temp_str)
        End If
        If ((Status And 64) = 0) Then Exit Do
        cmdAcqTime = Timer - CmdStartTime
        status_chk = status_chk + 1
        If response = "Timeout" Then Exit Do
    Loop While (cmdAcqTime < 11) '64 = 0x40, // set bit 6 ON - data not ready

    lblRawResponse = "acquisition time = " & Format(cmdAcqTime, "#.00") & " sec; status checked " & status_chk & " times"
    InitiateInternslRec = response
End Function

Private Sub cmdStartPosMonitoring_Click() '*pCommandString == 'a' - start acquisition
    cmdStartPosMonitoring.Enabled = False
    cmdStartMotorTest.Enabled = False
    InitiateInternslRec ("a")
    cmdStartPosMonitoring.Enabled = True
    cmdStartMotorTest.Enabled = True
End Sub

Private Sub cmdStartMotorTest_Click()
  Dim ret_str As String
  Dim separator As Long
    cmdStartPosMonitoring.Enabled = False
    cmdStartMotorTest.Enabled = False
    PC_RealTime = False
    Test_in_progress = True
    Analyzer.lstResponse.Clear
    Buffer = ""
    cmd_sent = False
    StartTime = Timer
    cwGraphDataAcq.Caption = "motor test"
    txtStatus.Text = "Performing motor test...."
    ret_str = Analyzer.GetSend("m", True)
          Do 'small delay
            DoEvents
            delayTime = Timer - StartTime
          Loop Until (delayTime > 2.2)
    separator = InStr(Buffer, "PASS")
    If separator <> 0 Then txtStatus.Text = Mid(Buffer, separator)
    separator = InStr(Buffer, "FAIL")
    If separator <> 0 Then txtStatus.Text = Mid(Buffer, separator)

    cmdStartPosMonitoring.Enabled = True
    cmdStartMotorTest.Enabled = True
End Sub

Public Function GetPosScan(cmd As String, WaitingTime As Single) As Long
' function is used during "valve test" and "get from buffer"
    Dim temp_str As String
    Dim scan_buf As String
    Dim str_len As Long
    Dim separator As Long
    Dim percent_divider As Long
    Dim lines_received As Long, ch As Long
'    Dim tmp_array(MAX_SCAN_ARRAY_SIZE) As Long

    PC_RealTime = False
    Test_in_progress = True
    line_counter = 0
    Analyzer.lstResponse.Clear
    Buffer = ""
    cmd_sent = False
    percent_divider = MAXmeasINDEX ' approximate number of lines received from controller
    If WaitingTime = 0 Then WaitingTime = 0.3
    ' first send command and get header back
'    If Analyzer.chkTestProgram = Unchecked Then
'        scan_buf = Analyzer.GetSend(cmd, True)
    Call Send_GUI_command(cmd & vbCr)
    StartTime = Timer
    If (MonitoringTestType = Valves_2020) Then 'Or (MonitoringTestType = Valve_box) Then ' valve test
      txtStatus.Text = "Testing valve board...."
      cwGraphDataAcq.Caption = "Iso Pressure vs. Valve Current"
      StartTime = StartTime + 1.2
      CMDperiod = cwValveMeasTime.Value
    End If
    If (MonitoringTestType = Monitor_Pressure) Then
      cwGraphDataAcq.Caption = "Pressure Readings"
      txtStatus.Text = "Getting pressure monitoring data from controller...."
      StartTime = StartTime + 0.2
    End If
    If (MonitoringTestType = Monitor_Pos) Then
      cwGraphDataAcq.Caption = "Position Readings"
      txtStatus.Text = "Getting position monitoring data from controller...."
      StartTime = StartTime + 0.2
    End If
'    End If
          Do 'small delay
            DoEvents
            delayTime = Timer - StartTime
          Loop Until (delayTime > WaitingTime)

another_line:
'    If Analyzer.chkTestProgram = Unchecked Then

        scan_buf = Analyzer.GetDataString(end_of_ScanTransmission, end_of_ScanString_token, WaitingTime + 0.5)
        If scan_buf = "UserStop" Then
            GetPosScan = -1
            Test_in_progress = False
            Exit Function
        End If
        scan_buf = LTrim(scan_buf)
'        separator = InStr(scan_buf, (Mid(cmd, 2) & vbCr)) 'find echo & vbCrLf
'        If (separator <> 0) Then
'            scan_buf = Mid(scan_buf, separator + 2)  'remove command string
'        End If
        separator = InStr(scan_buf, end_of_ScanString_token)  'find "%" & vbCrLf
        If separator = 0 Then 'header?
            If Len(scan_buf) > 0 Then
                Analyzer.lstResponse.AddItem Format(line_counter, "00#") & "->" & scan_buf
                If (File_Saving) Then
        '            Call write_one_line
                    Print #PositionMonitoring_file, scan_buf
                End If
            Else 'buffer empty
                Analyzer.lstResponse.AddItem "Empty string"
                GoTo another_line 'PROBLEM it cannot read line in time?
'                Call Analyzer.ShowNoCommunication
'                Test_in_progress = False
'                GetPosScan = 0
'                Exit Function
            End If
            ASCIIDataStr = scan_buf
        Else 'found "%" & vbCrLf - data line
            ASCIIDataStr = Left(scan_buf, separator)
            separator = InStr(ASCIIDataStr, vbCr) 'is there a header received?
            If separator <> 0 Then 'first line is a header terminated by CR
' line from controller:  "line#,   InfCnt,  EhxCnt,   Inflat_mA,  Exhaust_mA,   PSI_1,  PSI_2,   PSI_3,   PSI_4,   PSI_in"
' but we created and printed it already, no need to print
                'If (File_Saving) Then
                '    Print #PositionMonitoring_file, Left(ASCIIDataStr, separator - 1) ' without ";" at the end Print automatically adds LfCr
                'End If
                ASCIIDataStr = Mid(ASCIIDataStr, separator)
            End If
            If ASCIIDataStr <> "" Then
                lines_received = ExtractDataFromScanLine(ASCIIDataStr)  'returns real scan line number
                ScanTimeStamps(line_counter) = Timer
                If (File_Saving) Then
                    Call write_one_line(line_counter)
                End If
            End If
            Analyzer.lstResponse.AddItem Format(lines_received, "00#") & "->" & ASCIIDataStr
        End If

        If InStr(scan_buf, "OK") <> 0 Or ASCIIDataStr = "" Then
            GoTo show_plot ' END of transmission
        End If
        If ASCIIDataStr = "Timeout" Then  ' something bad
            Analyzer.lstResponse.AddItem "Timeout"
            Call Analyzer.ShowNoCommunication
            Test_in_progress = False
            GetPosScan = 0
            Exit Function
        End If

        'we've got the measurement strings, the FIRST line is header (IGNORED),
        'the rest of lines are DATA, format
'//index,  V1,    V2,    V3,    H1,    H2,    H3,  ffXP,  ffYP,  ffXA,  ffYA%\r\n
'0000, +19304,+35087,+24176,+33354,+45331,+27864,+30547,-40255,+00002,-00045%\r\n
'0001, +19298,+35082,+24170,+33359,+45309,+27854,+30533,-40258,+00006,-00041%\r\n
'0002, +19289,+35099,+24162,+33358,+45286,+27838,+30516,-40246,+00020,-00048%\r\n
'0003, +19245,+35104,+24155,+33384,+45255,+27816,+30501,-40244,+00036,-00043%\r\n

        If Len(ASCIIDataStr) = 0 Then ' bad data line
            Call Analyzer.ShowNoCommunication
            GoTo show_plot
        End If
        If Len(ASCIIDataStr) > 50 Then 'valid format
           temp_str = ASCIIDataStr
'           If line_counter = 0 Then
'                GoTo another_line ' skip header
'           End If
           If line_counter = 30 Then
               Status = 0 'test trap
           End If
        Else   ' the length is less than expected
           GoTo another_line ' for test
        End If
''**********************************************************************
'    Else 'for program test - without getting real data from controller
''**********************************************************************
'        percent_divider = MAX_SCAN_ARRAY_SIZE ' approximate number of lines received from controller
'        If line_counter >= percent_divider Or (waitresponse = False) Then
'            GoTo show_plot
'        Else
'            aTimerTestProgram.Interval = 2 ' 20 millisecond min interval in test mode
'            test_timer_ready = False
'            aTimerTestProgram.Enabled = True
'
'            Do
'                DoEvents
'            Loop While (test_timer_ready = False)
'            temp_str = Format(line_counter, "000#") & ", "
'            If (line_counter Mod 5) = 0 Then lines_received = ExtractDataFromScanLine(temp_str & "+19304,+35087,+24176,+33354,+45331,+27864,+30547,-40255,+00002,-00045." & vbCrLf)
'            If (line_counter Mod 5) = 1 Then lines_received = ExtractDataFromScanLine(temp_str & "+19298,+35082,+24170,+33359,+45309,+27854,+30533,-40258,+00006,-00041." & vbCrLf)
'            If (line_counter Mod 5) = 2 Then lines_received = ExtractDataFromScanLine(temp_str & "+19289,+35099,+24162,+33358,+45286,+27838,+30516,-40246,+00020,-00048." & vbCrLf)
'            If (line_counter Mod 5) = 3 Then lines_received = ExtractDataFromScanLine(temp_str & "+19245,+35104,+24155,+33384,+45255,+27816,+30501,-40244,+00036,-00043." & vbCrLf)
'            If (line_counter Mod 5) = 4 Then lines_received = ExtractDataFromScanLine(temp_str & "+19205,+35122,+24133,+33390,+45230,+27790,+30490,-40240,+00042,-00040." & vbCrLf)
'        End If
'    End If


'        If (line_counter = 0) Then
'        ' fill arrays with the first measurement for better plot scaling
'            For ch = 0 To MAX_CH
'                Status = SetPositionArray(ch, ScanValues(ch, 0))
'            Next ch
'        End If
        Call UpdatePositionPlot(line_counter) ' add / update a particular point
        line_counter = line_counter + 1
'        Analyzer.updateProgress (line_counter / (percent_divider + 1))

        lbLineNum.Caption = line_counter

'        CWGraph.Caption = "Freq = " & Format(frequency, "#0.0") & ";  Gain = " & Format(gain, "0#.0")
'        CWGraphPhase.Caption = "Freq = " & Format(frequency, "#0.0") & ";  Phase = " & Format(phase, "0##.#")

        GoTo another_line
show_plot:
        If (File_Saving) Then
            cwSaveScan.Value = False
        End If

        If line_counter < 0 Then Exit Function ' increments during real scan
        If lines_received < 0 Then Exit Function ' increments when reading file
        ' fill end of array with the last valid data for better auto scaling
        If (lines_received > 1 Or line_counter > 1) Then

          ScanNumber = line_counter ' because UpdatePositionPlot(-1) checks If SaveExsistingData = False and assigns to max points ScanNumber
          Call UpdatePositionPlot(-1) ' -1 means update whole plot, not a particular point

          If (MonitoringTestType = Valves_2020) Then
            Call Analyze_Valve_plot
          Else
            For ch = 0 To MAX_CH
              For line_counter = 0 To lines_received
                  tmp_array(line_counter) = ScanValues(ch, line_counter)
              Next line_counter
              separator = RefreshNumDisplay(tmp_array, 1, MinIndex(ch), MinValue(ch), MaxIndex(ch), MaxValue(ch), 0, 0, 0, 0, ch)
              'show average values
              lbPosition(ch).Caption = Format(ScanAverage(ch), "0000#")
            Next ch
            Call St_Dev(lines_received)
          End If
        End If

        Test_in_progress = False
        Analyzer.cwSldProgress.Value = 100
        Analyzer.cwSldProgress.Caption = "! DONE !"
        If (lines_received > line_counter) Then
            GetPosScan = lines_received ' increments when reading file
        Else
            GetPosScan = line_counter ' increments during real scan
        End If
End Function



Private Sub chkPOSvisible_Click(Index As Integer)
    If (chkPOSvisible(Index).Value = Checked) Then
        cwGraphDataAcq.Plots.Item(Index + 1).Visible = True
    Else
        cwGraphDataAcq.Plots.Item(Index + 1).Visible = False
    End If

End Sub

Sub UpdatePositionPlot(Optional scan_num As Long = -2) 'minus 2 when parameter omitted
  Dim i As Long
  Dim j As Long
  Dim smpl As Long
'-!-  Dim temp_array_Y() As Double
  If MonitoringTestType = Valves_2020 Then ' Or MonitoringTestType = Valve_box Then ' valve test
'"line#,DrvCnt,Curr_mA1,_mA2,_mA3,_mA4, PSI_1,PSI_2,PSI_3,PSI_4,PSI_in,Hold_mA\r\n"
'0019 ,  -500, 90.7, 90.7, 90.7, 90.7, 0.2,  0.1, -0.1, -0.0, -0.1, 77.00%
'0020 ,    +0, 90.7, 90.7, 90.7, 90.7, 0.1, -0.0, -0.1, -0.0, -0.1, 77.00%
'0021 ,  +500, 90.7, 90.7, 90.7, 90.7, 0.1,  0.1, -0.1, -0.0, -0.1, 77.00%
' line#,InfCnt,EhxCnt, Inflat_mA,Exhaust_mA, PSI_1,PSI_2,PSI_3,PSI_4,PSI_in,
'0000, -11999, 2226,   59.94,    59.97,      49.58,51.76,50.49,51.22,80.15,%
'0001, -11799, 2226,   60.44,    59.97,      49.25,51.62,50.08,51.12,80.11,%
'Test finished: *
'>~OK
    ' absolute
'    make option "plot XY", where X either mA of the Valve, or holding current mA, or index, or sec.
    'Line Number =0
    '2 nd Column =1
    '3 rd Column =2
    '4 th Column =3
    'Time        =4
    cwGraphDataAcq.Axes.Item(1).Caption = ""
    cwGraphDataAcq.Axes.Item(2).Caption = "Pressure, PSI"
    TextXunits.Visible = True

    XaxisUnits = Cmb_X_scale.ListIndex 'column # in the csv file
    If (XaxisUnits = 0) Then TextXunits.Text = " line #"
    If (XaxisUnits = 1) Then TextXunits.Text = " sec"
    If (XaxisUnits = 2) Then TextXunits.Text = " Inflation_Cnt"
    If (XaxisUnits = 3) Then TextXunits.Text = " Ehxhaust_Cnt"
    If (XaxisUnits = 4) Then TextXunits.Text = " Inflation_mA"
    If (XaxisUnits = 5) Then TextXunits.Text = " Exhaust_mA"

    If SaveExsistingData = False Then ' starting new recording
        If (scan_num > -1) Then
            j = scan_num
        Else
            j = ScanNumber - 1 ' minus one because ScanNumber has been already increased
        End If
    Else
        j = line_counter - 1 ' minus one because ScanNumber has been already increased
    End If
    If j <= 0 Then j = 0
    If j < MAX_SCAN_ARRAY_SIZE - 1 Then
        ReDim Preserve ScanValuesToPlot(MAX_CH, j)
        ReDim Preserve plt_X_coord(j)
'-!- cannot do ReDim if "array already dimensioned", i.e defined NOT as arr(), but arr(10,3)
'-!-        ReDim Preserve tmp_array(j)
'-!-        ReDim Preserve ScanValues(MAX_CH, j)
'-!-        ReDim Preserve RelativeScanValues(MAX_CH, j)
    End If

    If (XaxisUnits = 0) Or (XaxisUnits = 1) Then 'plot with X as line number; or with X as time: format X scale
        TextXunits.Visible = True
        If XaxisUnits = 0 Then
            cwGraphDataAcq.Axes.Item(1).FormatString = ""
        Else
            cwGraphDataAcq.Axes.Item(1).FormatString = "*" & CStr(cwValveMeasTime.Value)
        End If

        If (scan_num = -1) Then 'recalculate data to plot
            For smpl = 0 To j
                If (XaxisUnits = 0) Then
                    plt_X_coord(smpl) = smpl 'create X axis in index
                Else
                    plt_X_coord(j) = ScanTimeStamps(j) - ScanTimeStamps(0)  'create X time scale: subtract first time stamp
                End If
                For i = 0 To MAX_CH - 4 'add one measurement
                    If cwRelAbsPosition.Value = True Then 'relative
                        ScanValuesToPlot(i, smpl) = RelativeScanValues(i + 4, smpl) 'add one (latest) measurement
                    Else ' absolute
                        ScanValuesToPlot(i, smpl) = ScanValues(i + 4, smpl) 'add one (latest) measurement
                    End If
                Next i
            Next smpl
        Else 'scan_num > -1, do not recalculate whole array
            plt_X_coord(j) = ScanTimeStamps(j) - ScanTimeStamps(0)  'create X time scale: subtract first time stamp
            For i = 0 To MAX_CH - 4 'add one (latest) measurement
                If cwRelAbsPosition.Value = True Then 'relative
                    ScanValuesToPlot(i, j) = RelativeScanValues(i + 4, j) 'add one (latest) measurement
                Else ' absolute
                    ScanValuesToPlot(i, j) = ScanValues(i + 4, j) 'add one (latest) measurement
                End If
            Next i
        End If
        XminScale = plt_X_coord(0)
        XmaxScale = plt_X_coord(UBound(plt_X_coord))
        cwGraphDataAcq.Axes.Item(1).SetMinMax XminScale, XmaxScale '-!- if called often, this causes sharp increase in processor load
        cwGraphDataAcq.PlotXvsY plt_X_coord, ScanValuesToPlot ' do plot not so often
        cwGraphDataAcq.Axes.Item(2).AutoScaleNow
    End If

    If (XaxisUnits > 1) Then 'plot with X as 2nd column
        If (scan_num = -1) Then 'recalculate data to plot
            For smpl = 0 To j
                plt_X_coord(smpl) = ScanValues(2, smpl) 'create X axis in mA:

                For i = 0 To MAX_CH - 4 'add one measurement
                    If cwRelAbsPosition.Value = True Then 'relative
                        ScanValuesToPlot(i, smpl) = RelativeScanValues(i + 4, smpl) 'add one (latest) measurement
                    Else ' absolute
                        ScanValuesToPlot(i, smpl) = ScanValues(i + 4, smpl) 'add one (latest) measurement
                    End If
                Next i
            Next smpl
        Else 'scan_num > -1, do not recalculate whole array
            plt_X_coord(j) = ScanValues(2, j) 'create X axis in mA:
            For i = 0 To MAX_CH - 4 'add one (latest) measurement
                If cwRelAbsPosition.Value = True Then 'relative
                    ScanValuesToPlot(i, j) = RelativeScanValues(i + 4, j) 'add one (latest) measurement
                Else ' absolute
                    ScanValuesToPlot(i, j) = ScanValues(i + 4, j) 'add one (latest) measurement
                End If
            Next i
        End If
        XminScale = plt_X_coord(0)
        XmaxScale = plt_X_coord(UBound(plt_X_coord))
        cwGraphDataAcq.Axes.Item(1).SetMinMax XminScale, XmaxScale '-!- if called often, this causes sharp increase in processor load
        cwGraphDataAcq.PlotXvsY plt_X_coord, ScanValuesToPlot ' do plot
        cwGraphDataAcq.Axes.Item(2).AutoScaleNow
    End If

  Else 'position test
    If (MonitoringTestType = Monitor_Pos) Or (MonitoringTestType = Monitor_Pressure) Then
    ' continuous monitoring position or pressure
      If (MonitoringTestType = Monitor_Pos) Then cwGraphDataAcq.Axes.Item(2).Caption = "counts"
      If (MonitoringTestType = Monitor_Pressure) Then cwGraphDataAcq.Axes.Item(2).Caption = "PSI"
      Call update_Graph_time_scale
    End If

'    If (cmdGetRecordFromBuffer.Value = True) Then
'    If MonitoringTestType = Buffer_rec Then
'        TextXunits.Text = "sample #"
'        cwSlideDynamicZoom.Axis.Maximum = line_counter
'        cwGraphDataAcq.Axes.Item(1).FormatString = ""
'        cwGraphDataAcq.Axes.Item(1).SetMinMax 0, line_counter
'        If cwRelAbsPosition.Value = True Then 'relative
'            cwGraphDataAcq.PlotY RelativeScanValues
'        Else ' absolute
'            cwGraphDataAcq.PlotY ScanValues
'        End If
'    Else 'seconds
        TextXunits.Visible = True
        TextXunits.Text = "sec"
        If SaveExsistingData = False Then ' starting new recording
            If (scan_num > -1) Then
                j = scan_num
            Else
                j = ScanNumber - 1 ' minus one because ScanNumber has been already increased
            End If
        Else
            j = line_counter - 1 ' minus one because ScanNumber has been already increased
        End If
        If j <= 0 Then j = 0
        If j < MAX_SCAN_ARRAY_SIZE Then
            ReDim Preserve ScanValuesToPlot(MAX_CH, j)
            ReDim Preserve plt_X_coord(j)
'-!- cannot do ReDim if "array already dimensioned", i.e defined NOT as arr(), but arr(10,3)
'-!-        ReDim Preserve tmp_array(j)
'-!-        ReDim Preserve ScanValues(MAX_CH, j)
'-!-        ReDim Preserve RelativeScanValues(MAX_CH, j)
        End If

        plt_X_coord(j) = ScanTimeStamps(j) - ScanTimeStamps(0)  'create X time scale: subtract first time stamp
        MaxScanTimeStamp = plt_X_coord(j) ' save max time - duration of the scan

        cwSlideDynamicZoom.axis.Minimum = 0
        cwSlideDynamicZoom.axis.Maximum = plt_X_coord(j)

'-!- vvvvvvvvvvvvvvvvvvvvvvvvvvvv this causes sharp increase in processor load
'        cwGraphDataAcq.Axes.Item(1).SetMinMax 0, MaxScanTimeStamp
'-!- ^^^^^^^^^^^^^^^^^^^^^^^^^^^^ this causes sharp increase in processor load
        cwGraphDataAcq.Axes.Item(1).FormatString = "*1"
        If (scan_num = -1) Then 'recalculate data to plot
            For smpl = 0 To j
                plt_X_coord(smpl) = ScanTimeStamps(smpl) - ScanTimeStamps(0)  'create X time scale: subtract first time stamp
                MaxScanTimeStamp = plt_X_coord(smpl) ' save max time - duration of the scan

                For i = 0 To MAX_CH - 1 'add one (latest) measurement
                    If cwRelAbsPosition.Value = True Then 'relative
                        ScanValuesToPlot(i, smpl) = RelativeScanValues(i, smpl)  'add one (latest) measurement
                    Else ' absolute
                        ScanValuesToPlot(i, smpl) = ScanValues(i, smpl)  'add one (latest) measurement
                    End If
                Next i
            Next smpl
            XminScale = plt_X_coord(0)
            XmaxScale = plt_X_coord(UBound(plt_X_coord))
            cwGraphDataAcq.Axes.Item(1).SetMinMax XminScale, XmaxScale '-!- if called often, this causes sharp increase in processor load
            cwGraphDataAcq.PlotXvsY plt_X_coord, ScanValuesToPlot ' do plot not so often
            cwGraphDataAcq.Axes.Item(2).AutoScaleNow
        Else 'scan_num > -1, do not recalculate whole array
            For i = 0 To MAX_CH - 1 'add one (latest) measurement
                If cwRelAbsPosition.Value = True Then 'relative
                    ScanValuesToPlot(i, j) = RelativeScanValues(i, j)  'add one (latest) measurement
                Else ' absolute
                    ScanValuesToPlot(i, j) = ScanValues(i, j)  'add one (latest) measurement
                End If
            Next i
'            cwGraphDataAcq.PlotXvsY plt_X_coord, ScanValuesToPlot ' do plot not so often
'            cwGraphDataAcq.Axes.Item(2).AutoScaleNow
        End If
        XminScale = plt_X_coord(0)
        XmaxScale = plt_X_coord(UBound(plt_X_coord))
        If (j < 1000) Then
            cwGraphDataAcq.Axes.Item(1).SetMinMax XminScale, XmaxScale '-!- if called often, this causes sharp increase in processor load
            cwGraphDataAcq.PlotXvsY plt_X_coord, ScanValuesToPlot ' plot all data
            cwGraphDataAcq.Axes.Item(2).AutoScaleNow
        Else
            If ((j > 1000) And (j Mod 10) = 0) Then ' do plot not so often
'                cwGraphDataAcq.Axes.Item(1).Maximum = MaxScanTimeStamp '-!- if called often, this causes sharp increase in processor load
                cwGraphDataAcq.Axes.Item(1).SetMinMax XminScale, XmaxScale '-!- if called often, this causes sharp increase in processor load
'                cwGraphDataAcq.Axes.Item(1).AutoScaleNow 'gives loose bigger max scale
                cwGraphDataAcq.PlotXvsY plt_X_coord, ScanValuesToPlot ' do plot not so often
                cwGraphDataAcq.Axes.Item(2).AutoScaleNow
            End If
        End If
  End If
'-!-            Exit Sub
End Sub

Private Sub cwBut_Yautoscale_ValueChanged(ByVal Value As Boolean)

'    cwGraphDataAcq.Axes.Item(2).AutoScale = True
    If Value = True Then
        cwGraphDataAcq.Axes.Item(2).AutoScaleNow
    Else    'FULL SCALE
        cwGraphDataAcq.Axes.Item(2).SetMinMax CSng(YminScale), CSng(YmaxScale)
    End If
    cwGraphDataAcq.Refresh
End Sub

Sub update_Graph_time_scale()
    Dim i_tmp As Long
    Cmb_X_scale.ListIndex = 1 'time
    XaxisUnits = Cmb_X_scale.ListIndex
    i_tmp = CInt(cwContCMDperiod.Value * 10)
'    If MonitoringTestType = Buffer_rec Then
''    If PC_RealTime = True Then
'        cwGraphDataAcq.Axes.Item(1).FormatString = "*0.01"
'    Else
        If Cmb_X_scale.ListIndex = 1 Then ' time, sec
            cwGraphDataAcq.Axes.Item(1).FormatString = "*1"
        Else
            cwGraphDataAcq.Axes.Item(1).FormatString = "*0." & CStr(i_tmp)
        End If
'    End If
End Sub

Private Sub cwBut_Xautoscale_ValueChanged(ByVal Value As Boolean)
    XminScale = plt_X_coord(0)
    XmaxScale = plt_X_coord(UBound(plt_X_coord))
    cwSlideDynamicZoom.axis.Minimum = XminScale
    cwSlideDynamicZoom.axis.Maximum = XmaxScale
'    cwGraphDataAcq.Axes.Item(1).SetMinMax XminScale, XmaxScale

    If Value = True Then 'MAX FULL SCALE
        cwSlideDynamicZoom.Enabled = False
        cwSlideDynamicZoom.BackColor = &HD0D0D0 'gray
    Else    'ENABLE SCALE ADJUSTMENT
        cwSlideDynamicZoom.Enabled = True
        cwSlideDynamicZoom.BackColor = &HC0FFC0 'light green
        cwSlideDynamicZoom.Pointers(2).Value = XminScale + 0.85 * (XmaxScale - XminScale)
        cwSlideDynamicZoom.Pointers(1).Value = XminScale + 0.15 * (XmaxScale - XminScale)
        XminScale = cwSlideDynamicZoom.Pointers(1).Value ' index, for sec: Value/ cwContCMDperiod.Value
        XmaxScale = cwSlideDynamicZoom.Pointers(2).Value ' index, for sec: Value / cwContCMDperiod.Value
    End If
    cwGraphDataAcq.Axes.Item(1).SetMinMax CSng(XminScale), CSng(XmaxScale)
    Call UpdatePositionPlot(-1) ' -1 means update whole plot, not a particular point
End Sub

'Private Sub txtMaxXScale_Change()
'    If cwBut_Xautoscale.Value = True Then
'        Exit Sub
'    Else
'        If (IsNumeric(txtMaxXScale) = True) Then
'            If CSng(txtMaxXScale.Text) > CSng(txtMinXScale.Text) Then
'                txtMaxXScale.BackColor = &HC0FFC0   'OK
'                If PC_RealTime = False Then  'saved scan , index change = 0.01 sec
'                    XmaxScale = CVar(txtMaxXScale) * 100
'                Else
'                    XmaxScale = CVar(txtMaxXScale) / cwContCMDperiod.Value
'                End If
'                cwGraphDataAcq.Axes.Item(1).Maximum = XmaxScale
'                cwGraphDataAcq.Refresh
'            Else
'                txtMaxXScale.BackColor = &H8080FF 'pink - error
'            End If
'        Else
'            txtMaxXScale.BackColor = &H8080FF 'pink - error
'        End If
'    End If
'End Sub
'
'Private Sub txtMinXScale_Change()
'    If cwBut_Xautoscale.Value = True Then
'        Exit Sub
'    Else
'        If (IsNumeric(txtMinXScale) = True) Then
'            If CSng(txtMinXScale.Text) < CSng(txtMaxXScale.Text) Then
'                txtMinXScale.BackColor = &HC0FFC0   'OK
'                If PC_RealTime = False Then  'saved scan , index change = 0.01 sec
'                    XminScale = CVar(txtMinXScale) * 100
'                Else
'                    XminScale = CVar(txtMinXScale) / cwContCMDperiod.Value
'                End If
'                cwGraphDataAcq.Axes.Item(1).Minimum = XminScale
'                cwGraphDataAcq.Refresh
'            Else
'                txtMinXScale.BackColor = &H8080FF 'pink - error
'            End If
'        Else
'            txtMinXScale.BackColor = &H8080FF 'pink - error
'        End If
'    End If
'End Sub
'
'Private Sub txtMaxYScale_Change()
'    If cwBut_Yautoscale.Value = True Then
'        Exit Sub
'    Else
'        If (IsNumeric(txtMaxYScale) = True) Then
'            If CSng(txtMaxYScale.Text) > CSng(txtMinYScale.Text) Then
'                txtMaxYScale.BackColor = &HC0FFC0   'OK
'                YmaxScale = CVar(txtMaxYScale)
'                cwGraphDataAcq.Axes.Item(2).Maximum = YmaxScale
'                cwGraphDataAcq.Refresh
'            Else
'                txtMaxYScale.BackColor = &H8080FF 'pink - error
'            End If
'        Else
'            txtMaxYScale.BackColor = &H8080FF 'pink - error
'        End If
'    End If
'End Sub
'
'Private Sub txtMinYScale_Change()
'    If cwBut_Yautoscale.Value = True Then
'        Exit Sub
'    Else
'        If (IsNumeric(txtMinYScale) = True) Then
'            If CSng(txtMinYScale.Text) < CSng(txtMaxYScale.Text) Then
'                txtMinYScale.BackColor = &HC0FFC0   'OK
'                YminScale = CVar(txtMinYScale)
'                cwGraphDataAcq.Axes.Item(2).Minimum = YminScale
'                cwGraphDataAcq.Refresh
'            Else
'                txtMinYScale.BackColor = &H8080FF 'pink - error
'            End If
'        Else
'            txtMinYScale.BackColor = &H8080FF 'pink - error
'        End If
'    End If
'
'End Sub
Private Sub cwRelAbsPosition_ValueChanged(ByVal Value As Boolean)
    If MonitoringTestType = Valves_2020 Then 'Or MonitoringTestType = Valve_box Then ' valve test
        cwSpreadPlotsByY.Visible = False
        lblSpreadPlots.Caption = ""
    Else '''''''''''''''''  positon test
        If cwRelAbsPosition.Value = True Then 'relative
            cwSpreadPlotsByY.Visible = True
            lblSpreadPlots.Caption = "Spread plots by"
'            cwGraphDataAcq.PlotY RelativeScanValues
        Else ' absolute
            cwSpreadPlotsByY.Visible = False
            lblSpreadPlots.Caption = ""
'            cwGraphDataAcq.PlotY ScanValues
        End If
    End If
    Call UpdatePositionPlot(-1) ' -1 means update whole plot, not a particular point
End Sub

Sub Re_Calculate_Relative_Data_line(indx As Long)
    Dim ch As Long
    For ch = 0 To MAX_CH ' - 1
       RelativeScanValues(ch, indx) = ScanValues(ch, indx) - ScanValues(ch, 0) + cwSpreadPlotsByY.Value * (5 - ch)
    Next ch
End Sub

Sub recalculate_Rel_plot()
    Dim indx As Long
    indx = ScanNumber
    indx = line_counter
    For indx = 0 To MAX_SCAN_ARRAY_SIZE
        Re_Calculate_Relative_Data_line (indx)
    Next indx
End Sub


Private Sub cwSpreadPlotsByY_ValueChanged(Value As Variant, PreviousValue As Variant, ByVal OutOfRange As Boolean)
    Call recalculate_Rel_plot
    Call UpdatePositionPlot(-1) ' -1 means update whole plot, not a particular point
'    If cwRelAbsPosition.Value = True Then 'relative
'        cwGraphDataAcq.PlotY RelativeScanValues
'    Else ' absolute
'        cwGraphDataAcq.PlotY ScanValues 'SHOWN_values
'    End If
End Sub

Sub Update_labels()
    Dim X, Y As Variant
    Dim tmp_f As Single
    Dim array_index As Long
    Dim ch As Long
    Dim time_str As String
    Dim data_string As String

' update graph label and channel readings
    If MonitoringTestType = Valves_2020 Then ' valve test
        array_index = cwGraphDataAcq.Cursors(1).PointIndex ' this is the index
        X = cwGraphDataAcq.Cursors(1).XPosition ' X is the mA, NOT INDEX
        Y = ScanValuesToPlot(ChosenCh - 1, array_index)
    Else
        array_index = cwGraphDataAcq.Cursors(1).PointIndex ' this is the index
        X = cwGraphDataAcq.Cursors(1).XPosition ' X is the time, NOT INDEX
        Y = ScanValues(ChosenCh - 1, array_index)
    End If
    If cwButContinuesSend.Value = False Then
      If MonitoringTestType = Valves_2020 Then ' valve test
        If (Abs(Y) > 100) Then
            data_string = " Pressuse " & Left(optChooseCh(ChosenCh).Caption, 4) & "=" & Format(Y, "#0.0") & "psi @ Current=" & Format(X, "#0.00") & "mA"
        Else 'less than 100
            data_string = " Pressuse " & Left(optChooseCh(ChosenCh).Caption, 4) & "=" & Format(Y, "#0.00") & "psi @ Current=" & Format(X, "#0.00") & "mA"
        End If
        cwGraphDataAcq.Caption = "Index = " & Format(array_index, "#0000") & "; " & data_string
        Call Show_Scan_raw_values(array_index)
      Else  'position test
        If (Abs(Y) > 100) Then
            data_string = optChooseCh(ChosenCh).Caption & " Value = " & Format(Y, "00#")
        Else 'less than 100
            data_string = optChooseCh(ChosenCh).Caption & " Value = " & Format(Y, "#0.00")
        End If
        time_str = ConvertSnglToTime(CDbl(X)) 'Format(X, "0#:0#:00.00#")
        cwGraphDataAcq.Caption = "Index = " & Format(array_index, "#0000") & "; Time = " & time_str & "; " & data_string
        Call Show_Scan_raw_values(array_index)
      End If
    Else ' during acquisition or file saving
       cwGraphDataAcq.Caption = "elapsed time, sec = " & Format(X * CMDperiod, "#000.0")
    End If
'''''''''
End Sub

Sub HighlightChannel(Value As Variant, PreviousValue As Variant)
    Dim X, Y As Variant
    Dim array_index As Long
    If Value = 0 Then Value = 1
    If PreviousValue = 0 Then PreviousValue = 2
    If PreviousValue > 10 Then PreviousValue = 10
    cwGraphDataAcq.Plots(PreviousValue).LineWidth = 1 'PreviousValue and Value can be same, reduce width first
    cwGraphDataAcq.Plots(Value).LineWidth = 3           ' then increase width
    chkPOSvisible(PreviousValue - 1).FontBold = False
    chkPOSvisible(PreviousValue - 1).Width = 435
    chkPOSvisible(Value - 1).FontBold = True
    chkPOSvisible(Value - 1).Width = 535


'    txtMinSignal_lineNum.Text = CStr(MinIndex(Value - 1))
'    txtMinSignal.Text = Format(MinValue(Value - 1), "####0")
'    txtMaxSignal_lineNum.Text = CStr(MaxIndex(Value - 1))
'    txtMaxSignal.Text = Format(MaxValue(Value - 1), "####0")
'    measAvg.Text = Format(ScanAverage(Value - 1), "####0.00")
'
' snap cursor on the active plot - does not work!
'    X = cwGraphDataAcq.Cursors(1).XPosition
'    Y = cwGraphDataAcq.Cursors(1).YPosition
'    cwGraphDataAcq.Cursors.Item(1).SnapMode = cwCSnapFloating
'
'    If cwRelAbsPosition = True Then
'        cwGraphDataAcq.Cursors(1).YPosition = RelativeScanValues((Value - 1), X)
'    Else
'        cwGraphDataAcq.Cursors(1).YPosition = SHOWN_values((Value - 1), X)
'    End If
'    cwGraphDataAcq.Cursors.Item(1).SnapMode = cwCSnapPointsOnPlot

    Set cwGraphDataAcq.Cursors(1).Plot = cwGraphDataAcq.Plots(Value)

    Call Update_labels
    If Test_in_progress = True Then Exit Sub

    For array_index = 0 To MAX_SCAN_ARRAY_SIZE ''isolate one ch
        tmp_array(array_index) = ScanValues(ChosenCh - 1, array_index)
    Next
    array_index = ChosenCh - 1
    picStatistics.BackColor = optChooseCh(ChosenCh).BackColor
    Y = RefreshNumDisplay(tmp_array, 1, MinIndex(array_index), MinValue(array_index), MaxIndex(array_index), MaxValue(array_index), 0, 0, 0, 0, array_index)
'
'    txtMinSignal_Time.Text = ConvertSnglToTime(ScanTimeStamps(MinIndex(array_index)))
'    txtMaxSignal_Time.Text = ConvertSnglToTime(ScanTimeStamps(MaxIndex(array_index)))
'
'    Y = ScanStDev(array_index)
'    If (Y) > 0 Then
'        If Y > 0.01 And Y < 100 Then
'            measError.Text = Format(Y, "0.0000")
'        Else
'            measError.Text = Format(Y, "Scientific")
'        End If
'    Else
'        measError.Text = "N/A"
'    End If
'
'    Y = ScanSNR(array_index)
'    If Y <> -1 Then
'        If Y > 1 And Y < 10000 Then
'            measSNR.Text = Format(Y, "0.0")
'        Else
'            measSNR.Text = Format(Y, "Scientific")
'        End If
'    Else
'            measSNR.Text = "undefined"
'    End If
'
'    Y = ScanPk_Pk(array_index)
'    If Y > 0 Then
'       txtPk_Pk.Text = CStr(Y)
'    Else
'       txtPk_Pk.Text = "N/A"
'    End If

End Sub

' -1 as argument tells NOT to calculate this parameter
Function RefreshNumDisplay(inArray As Variant, start_index As Long, _
MinIndx As Long, minVal As Single, MaxIndx As Long, maxVal As Single, _
avrg As Single, St_Dev_Value As Double, snr As Single, max_min As Single, ch As Long) As Long

    Dim strMaxMinValue As String
    Dim array_size As Long
    On Error Resume Next
    If MinIndx <> -1 And MaxIndx <> -1 And minVal <> -1 And maxVal <> -1 Then

        Status = FindMaxMin(inArray, start_index, maxVal, minVal, MaxIndx, MinIndx) 'ignore zeroes = false

        MinIndex(ch) = MinIndx
        MinValue(ch) = minVal

        MaxIndex(ch) = MaxIndx
        MaxValue(ch) = maxVal

        txtMinSignal_lineNum.Text = CStr(MinIndex(ch))
        txtMinSignal.Text = Format(MinValue(ch), "####0")
        txtMaxSignal_lineNum.Text = CStr(MaxIndex(ch))
        txtMaxSignal.Text = Format(MaxValue(ch), "####0")

    End If
            array_size = MAX_SCAN_ARRAY_SIZE
    If avrg <> -1 Then
        avrg = Avg(inArray, array_size)
        ScanAverage(ch) = avrg

'        If avrg > 0.01 And avrg < 100 Then
        measAvg.Text = Format(ScanAverage(ch), "####0.00")
'        Else
'            measAvg.Text = Format(CStr(avrg), "Scientific")
'        End If
    Else
        measAvg.Text = " N/A "
        ScanAverage(ch) = -1
    End If

    If St_Dev_Value <> -1 Then ' calculate std dev
'       St_Dev_Value = St_Dev(inArray, array_size, True) 'DO ignore zeroes
        St_Dev_Value = St_Dev(inArray, array_size, False) 'do not ignore zeroes
        If St_Dev_Value > 0.01 And St_Dev_Value < 100 Then
            measError.Text = Format(St_Dev_Value, "0.0000")
        Else
            measError.Text = Format(St_Dev_Value, "Scientific")
        End If
        ScanStDev(ch) = St_Dev_Value
    Else
        measError.Text = " N/A "
        ScanStDev(ch) = -1
    End If

    If St_Dev_Value <> -1 And avrg <> -1 Then
        If St_Dev_Value <> 0 Then
            snr = Abs(avrg / St_Dev_Value)
                If snr > 1 And snr < 10000 Then
                    measSNR.Text = Format(CStr(snr), "0.0")
                Else
                    measSNR.Text = Format(CStr(snr), "Scientific")
                End If
                ScanSNR(ch) = snr
        Else
            measSNR.Text = "undefined"
            ScanSNR(ch) = -1
        End If
    Else
            measSNR.Text = " N/A "
            ScanSNR(ch) = -1
    End If
    If max_min <> -1 Then
        txtPk_Pk.Text = CStr(Round((maxVal - minVal), 3))
        ScanPk_Pk(ch) = (maxVal - minVal)
     Else
        txtPk_Pk.Text = "N/A"
        ScanPk_Pk(ch) = -1

     End If
RefreshNumDisplay = ch
End Function

Private Sub aTimerContMonitoring_Timer()
  On Error GoTo saving_problems
    'turn on continuous command sending
    Dim command As String
    Dim time_ms As String
    Dim ch As Long
    Dim scan_line As Long
    Dim exact_time As Double
    If Init_sys_data = True Then Exit Sub

    If (MonitoringTestType = Monitor_Pos) Then '0 => position recording
        command = "b" '"j"
    End If
'    If (MonitoringTestType = Valves_2020) Then '2=> DC2020 valve test
'        command = "vtst " & CStr(cwValveStartCurrent.Value) & " " & CStr(cwValveIncrementCurrent.Value) _
'        & " " & CStr(cwValveEndCurrent.Value) & " " & CStr(cwValveMeasTime.Value) & " " & CStr(cwValveHoldCurrent.Value)
'    End If
'    If (MonitoringTestType = Valve_box) Then '3=> valve BOX test
'        command = "vbxt " & CStr(cwValveStartCurrent.Value) & " " & CStr(cwValveIncrementCurrent.Value) _
'        & " " & CStr(cwValveEndCurrent.Value) & " " & CStr(cwValveMeasTime.Value) & " " & CStr(cwValveHoldCurrent.Value)
'    End If

    If (MonitoringTestType = Monitor_Pressure) Then '4=> pressure recording
        command = "getpa"
    End If

    cmd_sent = False
    waitresponse = True
'    If Analyzer.chkTestProgram <> Checked Then
        response = Analyzer.GetSend(command, True) ' pressure monitoring vs time: "getpa"
        'Win-450 rev B returns: " 58.12, 57.70, 59.78, 54.18, 87.21,*"
'    Else
'        If (ScanNumber Mod 5) = 0 Then response = "0000, +19304,+35087,+24176,+33354,+45331,+27864,+30547,-40255,+00002,-00045%" & vbCrLf
'        If (ScanNumber Mod 5) = 1 Then response = "0000, +19298,+35082,+24170,+33359,+45309,+27854,+30533,-40258,+00006,-00041%" & vbCrLf
'        If (ScanNumber Mod 5) = 2 Then response = "0000, +19289,+35099,+24162,+33358,+45286,+27838,+30516,-40246,+00020,-00048%" & vbCrLf
'        If (ScanNumber Mod 5) = 3 Then response = "0000, +19245,+35104,+24155,+33384,+45255,+27816,+30501,-40244,+00036,-00043%" & vbCrLf
'        If (ScanNumber Mod 5) = 4 Then response = "0000, +19205,+35122,+24133,+33390,+45230,+27690,+30490,-40240,+00042,-00040%" & vbCrLf
'    End If

'    Analyzer.lstResponse.AddItem "=>" & response
''>>>>vvvvvvvvvvv THIS DOES NOT WORK WELL, easier to hide save button until test is done vvvvvvvvvvvvvvvvvvvvvv
'    If (MonitoringTestType = Valves_2020) Or (MonitoringTestType = Valve_box) Then '2=> DC2020 valve test
'        ScanNumber = GetPosScan(command, (cwValveMeasTime.Value + 0.1)) 'program is here during test and plotting
'        Close #PositionMonitoring_file
'        Exit Sub
'    End If
'>>>>^^^^^^^^^^^^^ THIS DOES NOT WORK WELL ^^^^^^^^^^^^^^^^^^^^^
    scan_line = ExtractDataFromScanLine(response, ScanNumber)
    If (scan_line < 0) Then ' IK20201211 response = ">~Unrecognized (EMPTY) Cmd, Error Code =1", returns -1
       MissingLines = MissingLines + 1
       lbMissingLines.Caption = MissingLines
       Exit Sub
    End If

    ScanTimeStamps(ScanNumber) = Timer ' create time stamp
    If (cwSaveScan.Value = True) Then 'lines are saving
    ' find out if we reached planned number of lines

        If ((Date_Start_Saving + 1 = Date) And (Timer > Time_Start_Saving)) Then 'next day, same time - close this file, reopen next
            Close #PositionMonitoring_file
            ch = SaveDataAcqFile(True) 'Next_Multiple_file As Boolean, it updates Date_Start_Saving and Time_Start_Saving
'            File_Saving = True
        End If

        If (ScanNumber <= Number_of_Records) Or ChkUntilSTOP = Checked Then  'save line

            ch = InStr(response, ",") 'find first ','
            If ch >= 5 And ch < 10 Then 'cut controller time, insert PC time
                command = Mid(response, ch + 1)
'               If InStr(command, "*alarm") = 0 Then 'no alarm
'                   ch = InStr(command, ",") 'find second ','
'                   command = Mid(command, ch)
'                   command = Format(Time, "hh:mm:ss,") & CStr(scan_line) & command
'               Else 'catch alarm string
                exact_time = Timer
'                    ScanTimeStamps(ScanNumber) = exact_time
'                    time_ms = Format(exact_time, "#.0##")
'                    time_ms = Mid(time_ms, InStr(time_ms, "."))
                command = "'" & ConvertSnglToTime(exact_time) & "," & command
                Print #PositionMonitoring_file, command 'print time in the first column
                'read buffer again just in case not to miss *alarm message (it is not synchronized)
                'if len(buffer)>xxx, new message - process
                ' recognize
                'print into file
'                    End If
            Else    'format is wrong, time is not recognized
                Print #PositionMonitoring_file, response
            End If
        Else    ' all lines are recorded, close file
            cwSaveScan.Value = False
        End If
    End If

    For ch = 0 To MAX_CH 'show current values
        lbPosition(ch).Caption = Format(ScanValues(ch, ScanNumber), "0000#")
    Next ch
    Call UpdatePositionPlot(ScanNumber) ' add / update a particular point
    ScanNumber = ScanNumber + 1
    lbLineNum.Caption = Format(ScanNumber, "0000")
' auto X- scaling
'    If cwBut_Xautoscale.Value = True Then
'        Call updatePositionPlot

'        cwGraphDataAcq.Axes.Item(1).Maximum = ScanNumber + 1
'        cwGraphDataAcq.Axes.Item(1).Minimum = 1
'        cwGraphDataAcq.Axes.Item(1).SetMinMax 0, ScanNumber - 1
'        cwGraphDataAcq.Axes.Item(1).AutoScale = True
'        cwGraphDataAcq.Axes.Item(1).AutoScaleNow
'        cwGraphDataAcq.Refresh

'    End If
' prevent array overflow - start from the beginning
        If (ScanNumber > (Recorded_SCAN_ARRAY_SIZE - 1)) Then 'next file, - close this file, reopen next
            ScanNumber = ScanNumber - 1 ' corrected number of saved records, open here next file
            If (cwSaveScan.Value = True) Then
'                If (ChkUntilSTOP = Unchecked) Then
'                    cwSavingRecords.Value = cwSavingRecords.Value - Recorded_SCAN_ARRAY_SIZE ' reduce recorded lines
'                Else
'                    'cwSavingRecords.Value = cwSavingRecords.Value + Recorded_SCAN_ARRAY_SIZE ' increase recorded lines
'                End If
                Close #PositionMonitoring_file
                ch = SaveDataAcqFile(True) 'pass "Next_Multiple_file" As Boolean, if "True", it updates Date_Start_Saving and Time_Start_Saving
            End If
        End If
'        ScanNumber = 0
    Exit Sub

saving_problems:
    MsgBox ("File '" & SaveFileName & "' is not open for writing")
End Sub

Sub prepare_for_plotting()
    Call frmTerminal.QuitMenu
    frmTerminal.Visible = False
    Analyzer.lstResponse.Clear
    Analyzer.cwBut_RT_Monitoring.Value = False 'stop RT monitoring
    lbLineNum.Caption = "0"
    Call clear_DataAcqArrays
End Sub

Private Sub cwButContinuesSend_ValueChanged(ByVal Value As Boolean)
    If cwButContinuesSend = True Then
        '*********** set up for continuous position monitoring
       MissingLines = 0
       lbMissingLines.Caption = MissingLines

        PC_RealTime = True ' set flag, it resets ONLY when getting scanned array from controller by "g" command
        ScanNumber = 0
        Call prepare_for_plotting
        If cwContCMDperiod.Value = 0 Then
            CMDperiod = 0.1 '0.1 sec
        Else
            CMDperiod = cwContCMDperiod.Value
        End If
        aTimerContMonitoring.Interval = CMDperiod * 1000
        '1000 diagnostic gives 1.03-1.04, up to 1.21 interval at *1000
        '960 gives interval 0.99925
        '460-454 gives interval 0.49925 (0.484-0.51)
        '258 gives interval 0.296-0.421
        '158 gives interval 0.20275
        '80 gives interval 0.125
        '63 gives interval 0.0109
        '62 gives interval 0.09225-0.099
        ' Interval    Timer setting
        ' 1.035           1000
        ' 0.99925         960
        ' 0.49925         458
        ' 0.296           258
        ' 0.20275         158
        ' 0.125           80
        ' 0.109           63
        ' 0.09925         62
        'y = 1006.2x - 43.432
        cmdReadPositionFile.Enabled = False

        If MonitoringTestType = Monitor_Pressure Then ''pressure monitoring
            cwRelAbsPosition.Value = False
        End If
        aTimerContMonitoring.Enabled = True
        cwContCMDperiod.Enabled = False
        If cwSaveScan.Value = False Then
'            File_Saving = False
            SaveExsistingData = False
        Else
'            File_Saving = True
        End If
    Else 'cwButContinuesSend = False - stop
        aTimerContMonitoring.Enabled = False
        cwContCMDperiod.Enabled = True
        cmdReadPositionFile.Enabled = True
    '************* turn off  position monitoring
    End If
End Sub

Private Sub cwContCMDperiod_ValueChanged(Value As Variant, PreviousValue As Variant, ByVal OutOfRange As Boolean)
    Dim CMD_period As Single
    If cwContCMDperiod.Value = 0 Then
        CMD_period = 0.1 '0.05 sec
    Else
        CMD_period = cwContCMDperiod.Value
    End If
    aTimerContMonitoring.Interval = CMD_period * 1000
    CMDperiod = CMD_period 'update global
    Call CalcNumber_ofRecords
    Call update_Graph_time_scale
End Sub

Private Sub cwGraphDataAcq_Click()
    Dim dummy As Long
    If ((cwGraphDataAcq.Cursors(1).XPosition <= cwGraphDataAcq.Axes(1).Minimum) Or (cwGraphDataAcq.Cursors(1).XPosition >= cwGraphDataAcq.Axes(1).Maximum)) Then
        cwGraphDataAcq.Cursors(1).XPosition = (cwGraphDataAcq.Axes(1).Minimum + cwGraphDataAcq.Axes(1).Maximum) / 2
    End If
    cwGraphDataAcq.Cursors(1).Visible = True
    If PreviousCh = ChosenCh Then
        If ChosenCh < 9 Then dummy = 9
        If ChosenCh = 9 Then dummy = 1
        If ChosenCh = 0 Then
            ChosenCh = 1
            PreviousCh = 1
        End If
    End If
    Call HighlightChannel(ChosenCh, dummy)
End Sub

Private Sub cwGraphDataAcq_dblClick()
    cwGraphDataAcq.Cursors(1).Visible = False
End Sub

Sub Show_Scan_raw_values(indx As Variant)
    Dim Y As Variant, ch As Long
'    lbLineNum.Caption = Format(indx, "0000")
    If MonitoringTestType = Valves_2020 Then  ' valve test
      For ch = 0 To MAX_CH - 2 'show current values
        Y = ScanValuesToPlot(ch, indx)
        If (Abs(Y) > 100) Then
            lbPosition(ch).Caption = Format(Y, "00#")
        Else 'less than 100
            lbPosition(ch).Caption = Format(Y, "#0.00")
        End If
      Next ch
    Else 'position plot
      For ch = 0 To MAX_CH 'show position values
        Y = ScanValues(ch, indx)
        If (Abs(Y) > 100) Then
            lbPosition(ch).Caption = Format(Y, "#0.0#")
        Else 'less than 100
            lbPosition(ch).Caption = Format(Y, "#0.00")
        End If
      Next ch
    End If
End Sub

Sub Show_raw_values()
    Dim Y, ch As Variant
'    lbLineNum.Caption = Format(scan_line, "0000")
    lbLineNum.Caption = Format(ScanNumber, "0###")
    For ch = 0 To MAX_CH - 1 'show current values
        Y = OneScanValues(ch)
        If (Abs(Y) > 100) Then
            lbPosition(ch).Caption = Format(Y, "00#")
        Else 'less than 100
            lbPosition(ch).Caption = Format(Y, "#0.00")
        End If
    Next ch
End Sub
Private Sub cwGraphDataAcq_CursorChange(CursorIndex As Long, Xpos As Variant, Ypos As Variant, bTracking As Boolean)
    Call Update_labels
End Sub

Function ConvertSnglToTime(time_stamp As Double) As String
    Dim t_str As String
    Dim tmp_val As Double
    Dim t_time As Double
    Dim t_hours As Integer

    tmp_val = (time_stamp - 0.4999) Mod 60 'seconds. -0.4999 sec because time 52.51 Mod 60 gives 53 sec!
    t_str = Format(tmp_val, "0#")

    t_time = time_stamp - tmp_val
    tmp_val = (t_time - 0.4999) Mod 3600 'minutes
    tmp_val = tmp_val / 60
    t_str = Format(tmp_val, "0#") + ":" + t_str

    t_time = t_time - tmp_val
    tmp_val = (t_time - 0.4999) Mod 86400 'hours
    t_hours = CInt((tmp_val / 3600) - 0.4999)
    t_str = Format(t_hours, "0#") + ":" + t_str
    t_time = t_time - tmp_val
    t_str = t_str & Format(t_time, ".000")
    ConvertSnglToTime = t_str 'add ' so excel does not convert time to its format
End Function

Function ConvertTimeStringToSingle(time_string As String) As Double
    ' returns number of seconds passed from midnight
    '  "'12:45:51.45" is converted to Timer() format - single
    Dim t_str As String
    Dim cut_str As String
    Dim tmp_val As Double
    Dim t_time As Double
    Dim sep As Long 'separator

    t_str = time_string
    sep = InStr(t_str, "'")
    If (sep <> 0) Then t_str = Mid(t_str, 2)

    sep = InStr(t_str, ":")
    If (sep <> 0) Then
        cut_str = Left(t_str, sep - 1)
        t_time = Val(cut_str) * 3600 'seconds in hour
        t_str = Mid(t_str, sep + 1) ' minutes left: "45:51.45"
    End If
    sep = InStr(t_str, ":")
    If (sep <> 0) Then
        cut_str = Left(t_str, sep - 1)
        t_time = t_time + Val(cut_str) * 60 'seconds in minute
        t_str = Mid(t_str, sep + 1) ' seconds left: "51.45"
    End If
    t_time = t_time + Val(t_str) 'seconds and fractions
    ConvertTimeStringToSingle = t_time
End Function

Private Sub cwSaveScan_ValueChanged(ByVal Value As Boolean)
    Static retval As Long
    If cwSaveScan.Value = True Then 'START SAVING
        cmdReadPositionFile.Enabled = False
        retval = SaveDataAcqFile(False) 'false means first file in multiple files
        If retval = 32755 Then 'Cancel is selected
            cwSaveScan.Value = False
            File_Saving = False
        Else
            MissingLines = 0
            lbMissingLines.Caption = MissingLines
            If SaveExsistingData = False Then ' NOT saving existing data
                ChkUntilSTOP.Enabled = False
                cwNumFileRecords.Enabled = False
                cwButContinuesSend.Enabled = False
            End If
        End If
    Else 'false, STOP SAVING
        cmdReadPositionFile.Enabled = True
'        If SaveExsistingData = False Then ' NOT saving existing data
'            ScanNumber = 0 'prepare for next click
'        End If
        Close #PositionMonitoring_file
        cwButContinuesSend.Value = False 'this will disable timer
        cwButContinuesSend.Enabled = True
        cwNumFileRecords.Enabled = True
        ChkExcelFriendly.Enabled = True 'allow changing array size after stop
        ChkUntilSTOP.Enabled = True

        File_Saving = False
'        SaveExsistingData = False
'        SaveExsistingData = True
        If retval <> 32755 Then
            txtStatus.Text = "Measurement file '" & SaveFileName & "' Saved"
            Analyzer.txtStatus.Text = "File(s)=" & SaveFileName & "' Saved"
        End If
    End If
End Sub

Function SaveDataAcqFile(Next_Multiple_file As Boolean) As Long
    'Sub SavePositionPlot()
    On Error Resume Next
    'user wants to save a test file
    'First get the filename
    Dim DefaultFilename As String
    'Dim SaveFileName As String
    Dim start_saving_time As String
    Dim FileName As String
    Dim i As Long
    Dim ch As Long
    Dim m_p As Single ' measurement period

    If MonitoringTestType = Monitor_Pos Then '  position vs time test
        DefaultFilename = Date$ & "_" & Left$(Time$, 2) & "-" & Mid$(Time$, 4, 2) & " Real_time.csv"
    End If
    If MonitoringTestType = Monitor_Pressure Then ' pressure vs time test
        DefaultFilename = Date$ & "_" & Left$(Time$, 2) & "-" & Mid$(Time$, 4, 2) & " Pressure_vs_time.csv"
    End If
    If MonitoringTestType = Valves_2020 Then ' valve test
        DefaultFilename = Date$ & "_" & Left$(Time$, 2) & "-" & Mid$(Time$, 4, 2) & " Test_VlvBrd"
        If (txtV1sn.Text <> "") Then
            DefaultFilename = DefaultFilename & ",v1-" & Trim(txtV1sn.Text)
        End If
        If (txtV2sn.Text <> "") Then
            DefaultFilename = DefaultFilename & ",v2-" & Trim(txtV2sn.Text)
        End If
        If (txtV3sn.Text <> "") Then
            DefaultFilename = DefaultFilename & ",v3-" & Trim(txtV3sn.Text)
        End If
        If (txtV4sn.Text <> "") Then
            DefaultFilename = DefaultFilename & ",v4-" & Trim(txtV4sn.Text)
        End If

        DefaultFilename = DefaultFilename & ".csv"
    End If

    If ScanNumber > 1 Then
        i = MsgBox("There is data on the plot: " & ScanNumber & " data record lines." & vbCrLf _
        & "Press 'YES' to save THIS data or" & vbCrLf & "'NO' to clear, RECORD and SAVE NEW data", _
        vbExclamation + vbYesNo, "Save Exsisting Data ?")
        If i = vbYes Then 'save existing data selected
            SaveExsistingData = True ' saving existing data
        Else
            SaveExsistingData = False ' starting new recording
            response = Analyzer.GetSend("echo>enab", True) 'reset verbose responce
        End If
    End If


    If (Next_Multiple_file = False) Then
        User_has_been_Warned = False
        dlgCommonDialog.CancelError = True
        dlgCommonDialog.DialogTitle = "Choose name for the measurement"
'        dlgCommonDialog.InitDir = "c:\test"
        dlgCommonDialog.DefaultExt = "csv"
        dlgCommonDialog.Filter = "Data File (*.csv)|*.csv|All files (*.*)|*.*"
        dlgCommonDialog.FileName = DefaultFilename 'if enabled, dialog does not show - just goes to save IT WAS A BUG FILE NAME SHOULD NOT HAVE ':'
        dlgCommonDialog.ShowSave
        If Err = 32755 Then 'Cancel is selected
            SaveDataAcqFile = 32755
            File_Saving = False
            Exit Function
        Else

            txtStatus.Text = "Saving data into file '" & dlgCommonDialog.FileName & " ....."
            SaveFileName = dlgCommonDialog.FileName  ' add ".txt" from full file path
            i = InStrRev(dlgCommonDialog.FileName, "\")
            DefaultDirectory = Left(dlgCommonDialog.FileName, i - 1)
            dlgCommonDialog.InitDir = DefaultDirectory
            If SaveExsistingData = False Then ' saving new data
                cwButContinuesSend.Enabled = False 'so user can't press it to stop - saving must be stopped first
                ChkExcelFriendly.Enabled = False 'prevent changing array size during records
                Call CalcNumber_ofRecords
                Number_of_Records = Total_Number_of_Records
                Call clear_DataAcqArrays
            End If
        End If
    Else 'next file in multiple files
        i = InStrRev(dlgCommonDialog.FileName, "\")
        dlgCommonDialog.InitDir = Left(dlgCommonDialog.FileName, i - 1)
        SaveFileName = dlgCommonDialog.InitDir & "\" & DefaultFilename 'this gets new date
    End If 'end single-multiple files

    'SaveFileName becomes "C:\test\my_test.csv"
    Analyzer.txtStatus.Text = "Saving data into file(s) " & SaveFileName & " ....."
    txtStatus.Text = "Saving data into file(s) " & SaveFileName & " ....."
    'now lets open the file and write the data
    i = InStrRev(SaveFileName, "\")
    FileName = Right(SaveFileName, Len(SaveFileName) - i) 'i.e. "my_test"
    File_Saving = True

    Open SaveFileName For Output As #PositionMonitoring_file 'Values
        Print #PositionMonitoring_file, "File name, "; FileName
        Print #PositionMonitoring_file, ",Date,"; Date$
        start_saving_time = Time$
        Print #PositionMonitoring_file, ",Time,"; start_saving_time
        If (MonitoringTestType = Valves_2020) Then
            m_p = cwValveMeasTime.Value
        Else ' position or pressure vs time
            m_p = cwContCMDperiod.Value
        End If
        CMDperiod = cwContCMDperiod.Value
        'If PC_RealTime = True Then 'if True saved time is PC time. If False time is (usually) 10 ms increments from controller acquisituin buffer

            If (MonitoringTestType = Valves_2020) Then 'Or (MonitoringTestType = Valve_box) Then ' valve test
                Print #PositionMonitoring_file, "Scan Type:, Valve Test"
                Print #PositionMonitoring_file, "command:,";
                Print #PositionMonitoring_file, "vtst ";
                'If (MonitoringTestType = Valve_box) Then Print #PositionMonitoring_file, "vbxt ";
                Print #PositionMonitoring_file, CStr(cwValveStartCurrent.Value); " "; CStr(cwValveIncrementCurrent.Value); " ";
                Print #PositionMonitoring_file, CStr(cwValveEndCurrent.Value); " "; CStr(cwValveMeasTime.Value); " ";
                Print #PositionMonitoring_file, CStr(cwValveHoldCurrent.Value)
                Print #PositionMonitoring_file, " period:,"; cwValveMeasTime.Value; ",sec"
            End If
            If MonitoringTestType = Monitor_Pressure Then ' pressure vs time test
                Print #PositionMonitoring_file, "Scan Type:, Pressure vs time"
                Print #PositionMonitoring_file, "request period, "; m_p; ",<<<" 'the token of quasi-real measurement is 'request period'
            End If
            If MonitoringTestType = Monitor_Pos Then ' position vs time test
                Print #PositionMonitoring_file, "Scan Type:, Position vs time"
                Print #PositionMonitoring_file, "request period, "; m_p; ",<<<" 'the token of quasi-real measurement is 'request period'
            End If
        'End If
        If SaveExsistingData = False Then ' starting new recording
            line_counter = 0
            Print #PositionMonitoring_file, " lines:, collecting in real time";
        Else 'existing data
            Print #PositionMonitoring_file, " lines:, "; ScanNumber;
        End If

        If (MonitoringTestType = Valves_2020) Then 'Or (MonitoringTestType = Valve_box) Then ' valve test
           Print #PositionMonitoring_file, ",,Valve Serial Number,,"; txtV1sn.Text; ","; txtV2sn.Text; ","; txtV3sn.Text; ","; txtV4sn.Text
           Print #PositionMonitoring_file, " line#,InfCnt,EhxCnt, Inflat_mA,Exhaust_mA, PSI_1,PSI_2,PSI_3,PSI_4,PSI_in, time"
        End If

        If MonitoringTestType = Monitor_Pos Then ' position vs time test
           Print #PositionMonitoring_file, vbCrLf; "time,  V1,    V2,    V3,    H1,    H2,    H3,  FF_Xpos,  FF_Ypos,  FF_Xacc,  FF_Yacc"
        End If

        If MonitoringTestType = Monitor_Pressure Then ' pressure vs time test
           Print #PositionMonitoring_file, vbCrLf; "time, PSI_1, PSI_2, PSI_3, PSI_4, PSI_in, "
        End If

' HEADER IS DONE
    Time_Start_Saving = Timer
    Date_Start_Saving = Date

'now Data lines
'//// here insert check either of scan type (If Index = Monitor_Pos Then ) or check mark "save existing scan"
       If SaveExsistingData = True Then ' saving existing data
'       If line_counter <> 0 Then 'if cwSaveScan.Value = False And records acquired first, then saved
'            If (MonitoringTestType = Monitor_Pos) Or (MonitoringTestType = Monitor_Pressure) Then ' it was periodic retrieval by PC timer
                For i = 0 To ScanNumber - 1 'it was periodic retrieval by PC timer
                    write_one_line (i)
                Next i
'            Else
'                If (MonitoringTestType = Valves_2020) Then ' script works inside DC-2020 and sends information, we capture it in the while loop
                ''''''''''''''
'                End If
'            End If
            Print #PositionMonitoring_file, ", - End of PC-saved file"    'end of file
'            cwSavingRecords.Enabled = True
            cwSaveScan.Value = False 'it will Close #PositionMonitoring_file
        Else ' acquiring and saving
            If (MonitoringTestType = Valves_2020) Then ' script works inside DC-2020 and sends information, we capture it in the while loop
                cwButValveTest.Value = True
            Else ' periodic retrieval by timer
                Call CalcNumber_ofRecords
                cwButContinuesSend.Value = True 'this will enable timer
            End If
            txtStatus.Text = "Saving data into file(s)'" & SaveFileName & " ....."
        End If
    SaveDataAcqFile = 0 'return success
End Function

Sub write_one_line(i As Long) '-!-
    Dim ch As Long
    Dim active_channels As Long
    Dim t_str As String
    t_str = "'" & ConvertSnglToTime(ScanTimeStamps(i))

    If MonitoringTestType = Monitor_Pos Then ' position vs time test
        active_channels = MAX_CH
        'If SaveExsistingData = False Then ' NOT saving existing data
            Print #PositionMonitoring_file, t_str; ","; ' print time stamp
        'End If
    End If

    If MonitoringTestType = Monitor_Pressure Then ' pressure vs time test
'       Print #PositionMonitoring_file, "index, PSI_1, PSI_2, PSI_3, PSI_4, PSI_in, time"
        active_channels = 4
'        If SaveExsistingData = False Then ' NOT saving existing data
            Print #PositionMonitoring_file, t_str; ","; ' print time stamp
'        End If
    End If

    If (MonitoringTestType = Valves_2020) Then 'Or (MonitoringTestType = Valve_box) Then ' valve test
        Print #PositionMonitoring_file, i; ",";
        active_channels = MAX_CH - 1
    End If

    For ch = 0 To active_channels
        Print #PositionMonitoring_file, ScanValues(ch, i); ",";
    Next ch

    If (MonitoringTestType = Valves_2020) Then 'Or (MonitoringTestType = Valve_box) Then ' valve test
        Print #PositionMonitoring_file, t_str; ","; ' print time stamp at the end
    End If
    Print #PositionMonitoring_file, "%"    'end of line
End Sub

Sub clear_DataAcqArrays()
    On Error Resume Next ' dynamic arrays plt_X_coord() and ScanValuesToPlot() might not be resized yet
    Dim i As Long
    Dim c As Long
    Dim max_i As Long
    Dim max_c As Long
    For i = 0 To MAX_SCAN_ARRAY_SIZE
      For c = 0 To MAX_CH - 1
        ScanValues(c, i) = 0  ' MAX_SCAN_ARRAY_SIZE scan lines,
      Next c
      ScanTimeStamps(i) = 0
    Next i
    max_i = UBound(plt_X_coord)
    For i = 0 To max_i
        plt_X_coord(i) = 0
        For c = 0 To MAX_CH - 1
            ScanValuesToPlot(c, i) = 0
        Next c
    Next i
End Sub

Private Sub cmdReadPositionFile_Click()
   Call ReadDataAcqFile
   Call UpdatePositionPlot(-1) ' -1 means update whole plot, not a particular point
   If MonitoringTestType = Valves_2020 Then ' Or MonitoringTestType = Valve_box Then ' valve test
        Call Analyze_Valve_plot
   End If
End Sub

Sub ReadDataAcqFile()
    On Error Resume Next
    'user wants to read a test file
    'First get the filename
    Dim DefaultFilename As String
    Dim DataFileName As String
    Dim FileName As String
    Dim i  As Long
    Dim temp_str As String
    Dim t_str As String
    Dim sep_pos As Long
    Dim line_read As Long
    Dim total_lines As Long
    Dim MaxSizeReached As Boolean 'to read file after array size is reached

   'DefaultFilename = Date$ & " " & Left$(Time$, 5) & "_Mag-NetX.csv"
    dlgCommonDialog.CancelError = True
    dlgCommonDialog.DialogTitle = "Choose data file name to read and show"
    dlgCommonDialog.FileName = "" 'OTHERWISE Dialog will not OPEN
'    dlgCommonDialog.InitDir = "c:\test"
    dlgCommonDialog.DefaultExt = "csv"
    dlgCommonDialog.Filter = "Data File (*.csv)|*.csv|All files (*.*)|*.*"
    dlgCommonDialog.Action = 1
    'dlgCommonDialog.ShowSave
    If Err = 32755 Then 'Cancel is selected
        Exit Sub
    Else
        txtStatus.Text = "Reading measurement from file '" & dlgCommonDialog.FileName & " ....."
'        Analyzer.txtStatus.Text = "File(s)=" & dlgCommonDialog.FileName
    End If
    'SaveFileName becomes "C:\test\my_test.txt"

'    SaveFileName = Left(dlgCommonDialog.FileName, Len(dlgCommonDialog.FileName) - 4) ' remove ".txt" from full file path
    DataFileName = dlgCommonDialog.FileName
    'SaveFileName becomes "C:\test\my_test"
    i = InStrRev(DataFileName, "\")
    DefaultDirectory = Left(DataFileName, i - 1)
    dlgCommonDialog.InitDir = DefaultDirectory
    FileName = Right(DataFileName, Len(DataFileName) - i) 'i.e. "my_test"

    DataFileName = dlgCommonDialog.FileName
    i = InStrRev(DataFileName, "\")
    FileName = Right(DataFileName, Len(DataFileName) - i) 'i.e. "my_test"
    'now lets open the file and write the data

    line_read = 0
    total_lines = 0
    MaxSizeReached = False
    File_Reading = True
    Open DataFileName For Input As #DataAcq_file 'Values
read_again:
    Line Input #DataAcq_file, temp_str '    "File name," FileName
    ' decode line
    If (temp_str = "") Then GoTo read_again

    sep_pos = InStr(temp_str, ",") 'look for coma
    If sep_pos = 0 Then GoTo read_again 'no coma - read again
    If Left(temp_str, sep_pos - 1) <> "File name" Then
        If (IsNumeric(Left(temp_str, 3))) Then
            line_read = 1
            clear_DataAcqArrays
            GoTo matlab
        Else
            GoTo error_format
        End If
    Else
'        Analyzer.txtStatus = Mid(temp_str, sep_pos + 1)
    End If


        Do '' skip header for now'.......add code ................
    '        Line Input #DataAcq_file, temp_str 'Print #DataAcq_file, "Date"; vbTab; Date$
    '        ' decode line
    '        sep_pos = InStr(temp_str, ",")
    '        '.......add code ................
    '        Line Input #DataAcq_file, temp_str 'Print #DataAcq_file, "Time"; vbTab; Time$
    '        ' decode line
    '        sep_pos = InStr(temp_str, ",")
    '        '.......add code ................
    '
    '
    '        Line Input #DataAcq_file, temp_str 'Print #DataAcq_file,
    '        ' decode line
    '        sep_pos = InStr(temp_str, ",")
    '        '.......add code ................
    '
    '        'cheating.............add code ................
            If EOF(DataAcq_file) Then Exit Sub
            Line Input #DataAcq_file, temp_str
            DoEvents
            If MaxSizeReached = True Then Exit Do
            sep_pos = InStr(temp_str, "Scan Type:") 'catch type of file
            If sep_pos <> 0 Then
                sep_pos = InStr(temp_str, "Valve Test") 'catch type of file
                If sep_pos > 6 Then
                    ChangeMonitoringTestType (Valves_2020)
                    Cmb_X_scale.ListIndex = 2 '2nd column == current mA
                Else ' not valve test, position or pressure vs time
                    Cmb_X_scale.ListIndex = 1 '1st column == time
                    sep_pos = InStr(temp_str, "Pressure") 'catch type of file
                    If sep_pos > 6 Then
                        ChangeMonitoringTestType (Monitor_Pressure)
                    Else
                        ChangeMonitoringTestType (Monitor_Pos)
                    End If
                End If
            End If

            sep_pos = InStr(temp_str, "command:") 'catch command
            If sep_pos <> 0 Then
                sep_pos = InStr(temp_str, "vtst") 'catch DC2020 valve test
                If sep_pos > 6 Then
                    ChangeMonitoringTestType (Valves_2020)
                    PC_RealTime = True
                Else
'                    sep_pos = InStr(temp_str, "vbxt") 'catch valveBOX test
'                    If sep_pos > 6 Then
'                        ChangeMonitoringTestType (Valve_box)
'                        Cmb_X_scale.ListIndex = 2 '2nd column == current mA
'                    End If
                End If
                If sep_pos > 6 Then '"vtst" decode command line "vtst 60 .5 120 3 60" start mA inc mA stop mA wait sec hold mA
' valve test >>>>>>>vvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvv
                    t_str = Mid(temp_str, sep_pos + 4)
                    temp_str = LTrim(t_str) & " " 'adding space to catch the last parameter
                    sep_pos = InStr(temp_str, " ") 'catch next space
                    t_str = Mid(temp_str, 1, sep_pos)
                    If IsNumber(t_str) Then cwValveStartCurrent.Value = Val(t_str)     ' update valve start current

                    t_str = Mid(temp_str, sep_pos)
                    temp_str = LTrim(t_str)
                    sep_pos = InStr(temp_str, " ") 'catch next space
                    t_str = Mid(temp_str, 1, sep_pos)
                    If IsNumber(t_str) Then cwValveIncrementCurrent.Value = Val(t_str)     ' update valve current increment

                    t_str = Mid(temp_str, sep_pos)
                    temp_str = LTrim(t_str)
                    sep_pos = InStr(temp_str, " ") 'catch next space
                    t_str = Mid(temp_str, 1, sep_pos)
                    If IsNumber(t_str) Then cwValveEndCurrent.Value = Val(t_str)     ' update valve end current

                    t_str = Mid(temp_str, sep_pos)
                    temp_str = LTrim(t_str)
                    sep_pos = InStr(temp_str, " ") 'catch next space
                    t_str = Mid(temp_str, 1, sep_pos)
                    If IsNumber(t_str) Then
                        cwValveMeasTime.Value = Val(t_str)     ' update valve delay time interval, CMDperiod updates when cwValveMeasTime.Value changes
                        CMDperiod = Val(t_str) ' need if time is not specified on older records and for analyzing curve
                    End If
                    t_str = Mid(temp_str, sep_pos)
                    temp_str = LTrim(t_str)
                    sep_pos = InStr(temp_str, " ") 'catch next space
                    t_str = Mid(temp_str, 1, sep_pos)
                    If IsNumber(t_str) Then cwValveHoldCurrent.Value = Val(t_str)  ' update valve hold current (for push-pull valve config)
' valve test >>>>>>> ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
                End If
            End If 'command

            sep_pos = InStr(temp_str, ",#") 'catch if we read LTF instead of real time
            If sep_pos <> 0 Then
                Close #DataAcq_file
                MsgBox ("This is Transfer Function File, choose Real Time File")
                Exit Sub
            End If

            If (MonitoringTestType = Valves_2020) Then 'Or MonitoringTestType = Valve_box) Then
' valve test >>>>>>>vvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvv
            '" lines:,160,space, space,Valve SN, 11111,22222,33333,44444"
                sep_pos = InStr(temp_str, "lines:") 'catch header
                If sep_pos <> 0 Then
read_SN:
                    sep_pos = InStr(temp_str, ",") 'catch lines:,
                    i = Len(temp_str) ' lines:, 159
                    If (sep_pos > 0) Then temp_str = Mid(temp_str, sep_pos + 1) ' remains 159 and possibly valve numbers
                    sep_pos = InStr(temp_str, ",") 'catch lines:,
                    If (sep_pos > 0) Then  ' there are valve numbers, comma detected
                        temp_str = Mid(temp_str, sep_pos + 1) ' cut number of recorder lines

                    sep_pos = InStr(temp_str, "r,,") 'catch "r,," in line,"Valve Serial Number,,7626,7643,7659,7625,,," ,44444 last serial number V4

                    t_str = Mid(temp_str, sep_pos + 3) '7626,7643,7659,7625,,,"
                    temp_str = t_str

                    sep_pos = InStr(temp_str, ",") 'catch "," in line,"7626,7643,7659,7625,,," ,44444 last serial number V4
                    If (sep_pos > 1) Then t_str = Left(temp_str, sep_pos - 1) '="7626"
                    If (Len(t_str) > 1) Then
                        txtV1sn.Text = Trim(t_str)
                    'Else
                    '    temp_str = Left(temp_str, sep_pos - 1) '=44444
                    '    GoTo read_SN
                    End If

                    t_str = Mid(temp_str, sep_pos + 1) 'without 7626
                    temp_str = t_str

                    sep_pos = InStr(temp_str, ",") 'catch ,2222 V2 serial number
                    If (sep_pos > 1) Then t_str = Left(temp_str, sep_pos - 1) '=33333
                    txtV2sn.Text = Trim(t_str)
                    t_str = Mid(temp_str, sep_pos + 1) 'without 2222
                    temp_str = t_str

                    sep_pos = InStr(temp_str, ",") 'catch ,3333 V3 serial number
                    If (sep_pos > 1) Then t_str = Left(temp_str, sep_pos - 1) '=3333
                    txtV3sn.Text = Trim(t_str)
                    t_str = Mid(temp_str, sep_pos + 1) 'without 2222
                    temp_str = t_str

                    sep_pos = InStr(temp_str, ",") 'catch ,4444 V4 serial number
                    If (sep_pos > 1) Then t_str = Left(temp_str, sep_pos - 1) '=4444
                    txtV4sn.Text = Trim(t_str)
                    Else ' older record file without SN
                        txtV1sn.Text = ""
                        txtV2sn.Text = ""
                        txtV3sn.Text = ""
                        txtV4sn.Text = ""
                    End If
                End If

            '" line#,InfCnt,EhxCnt, Inflat_mA,Exhaust_mA, PSI_1,PSI_2,PSI_3,PSI_4,PSI_in,time"
                sep_pos = InStr(temp_str, "line#") 'catch header
                If sep_pos <> 0 Then
                    For i = 0 To 9
                        sep_pos = InStr(temp_str, ",") 'catch ','
                        If sep_pos <> 0 Then
                            DataHeaderStr(i) = Left(temp_str, sep_pos - 1)
                            DataHeaderStr(i) = LTrim(DataHeaderStr(i))
                            temp_str = Mid(temp_str, sep_pos + 1)
                        Else
                            Exit For
                        End If
                    Next i
                End If
' valve test >>>>>>> ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
            Else ' not valve test
                sep_pos = InStr(temp_str, "wait time") 'catch header
                If sep_pos <> 0 Then
                ' >>>>>> there is "wait time" in the header
                    sep_pos = InStr(temp_str, ",") 'catch ','
                    If sep_pos <> 0 Then
                        t_str = Mid(temp_str, sep_pos + 1)
                        sep_pos = InStr(t_str, "sec") 'catch 'sec'
                        t_str = Mid(t_str, 1, sep_pos - 1) ' actually t_str = " 0.5,"
                        If IsNumber(t_str) Then
                            If Val(t_str) = 0.01 Then 'real time
                                PC_RealTime = False
                                CMDperiod = 0.01
                            Else
                                PC_RealTime = True
                                cwContCMDperiod.Value = Val(t_str)
                                CMDperiod = cwContCMDperiod.Value
                            End If
                        End If
                    End If
                ' end of >>>>>> there is "wait time" in the header
                End If

                sep_pos = InStr(temp_str, "request period") 'catch quasi-real header when saved time == PC time
                If sep_pos <> 0 Then
                    sep_pos = InStr(temp_str, ",") 'catch ','
                    If sep_pos <> 0 Then '0.1
                        t_str = Mid(temp_str, sep_pos + 1)
                        sep_pos = InStr(t_str, ",") 'catch possible second ',' like ', <<<'
                        If sep_pos Then t_str = Left(t_str, sep_pos - 1)
                        If IsNumber(t_str) Then
                            If Val(t_str) = 0.01 Then 'real controller time with 10 ms increments
                                PC_RealTime = False
                            Else
                                PC_RealTime = True 'when saved time == PC time
                                cwContCMDperiod.Value = Val(t_str)
                                CMDperiod = cwContCMDperiod.Value
                            End If
                        End If
                    End If
                End If
            End If
            DoEvents
            If EOF(DataAcq_file) Then Exit Do
        Loop Until ((InStr(temp_str, "%") <> 0 And InStr(temp_str, "%") = Len(temp_str)) Or (InStr(temp_str, "End") <> 0) Or (InStr(temp_str, "*") <> 0)) ' last char is %
        ' decode line
matlab:
        Call clear_DataAcqArrays ' if program is here, it is not a LT file and header is processed

    Do While (Not EOF(DataAcq_file))  ' Loop until end of file.
' here is reading of line
        DoEvents

'14:14:20,65000,15720,-28065,-54104,5,1,3,%
'14:14:21,65001,15720,-28065,-54179,1,2,5,%
'14:14:22,65002,15495,-28065,-53241,2,0,3,%

'////////////////// array size vs file size
        If line_read > MAX_SCAN_ARRAY_SIZE - 1 Then
            MaxSizeReached = True
        End If
' skip bad strings
        If Len(temp_str) < 4 Then GoTo read_data_line
        If MaxSizeReached = False Then
'            If (Not EOF(DataAcq_file)) Then
                sep_pos = ExtractDataFromScanLine(temp_str, line_read)
'                Call UpdatePositionPlot(line_read) ' add / update a particular point
                If sep_pos >= 0 Then line_read = line_read + 1
'            End If
        End If
        total_lines = total_lines + 1

        If ((line_read > 200) And (line_read Mod 10) = 0) Then lbLineNum.Caption = line_read ' display current line
'        If ((total_lines > 200) And (total_lines Mod 100) = 0) Then cwSavingRecords.Value = total_lines
read_data_line:
        Line Input #DataAcq_file, temp_str
    Loop ' end of read loop

'    Exit Do
' =========================================================================================
    ScanNumber = line_read
    lbLineNum.Caption = line_read ' display current line
'    cwSavingRecords.Value = total_lines

    If MaxSizeReached = True Then
        sep_pos = MsgBox("This version only reads " & MAX_SCAN_ARRAY_SIZE & " data record lines." & vbCrLf _
        & "File has " & total_lines & " record lines. The rest of lines is skipped", _
        vbExclamation, "Data Size LIMIT")
    End If

    Close #DataAcq_file
    File_Reading = False

' here all lines are red, fill up to the end to allow auto scaling
'    For i = line_read To MAX_SCAN_ARRAY_SIZE
'        For sep_pos = 0 To MAX_CH
'            DoEvents
'            ScanValues(sep_pos, i) = ScanValues(sep_pos, line_read - 1)
'            RelativeScanValues(sep_pos, i) = RelativeScanValues(sep_pos, line_read - 1)
'        Next sep_pos
''        If i = 10000 Then
''            ScanNumber = line_read
''        End If
'    Next i
    Call HighlightChannel(ChosenCh, PreviousCh)

'    cwGraphDataAcq.ClearData
'''========= POSITION =====================================================
    If MonitoringTestType = Monitor_Pos Then 'prox sensors vs time plot
        XminScale = 0
        XmaxScale = CSng(ScanTimeStamps(ScanNumber - 1) - ScanTimeStamps(0))
        cwSlideDynamicZoom.Pointers(1).Value = 0
        cwSlideDynamicZoom.Pointers(2).Value = XmaxScale * 0.75
        cwSlideDynamicZoom.Pointers(1).Value = XmaxScale * 0.25
        cwSlideDynamicZoom.Caption = "Time, sec"
        TextXunits.Text = "sec"
    End If

'''========= PRESSURE =====================================================
    If MonitoringTestType = Monitor_Pressure Then 'pressure sensors vs time plot
        XminScale = 0
        XmaxScale = CSng((ScanNumber - 1) - ScanTimeStamps(0))
        cwSlideDynamicZoom.Pointers(1).Value = ScanNumber * 0.25
        cwSlideDynamicZoom.Pointers(2).Value = ScanNumber * 0.75
        cwRelAbsPosition.Value = False ' this will call UpdatePositionPlot(-1) if relative values are shown
        TextXunits.Text = "sec"
        cwSlideDynamicZoom.Caption = "Time, sec"
'        Call update_Graph_time_scale
    End If

'''========= VALVES =====================================================
    If MonitoringTestType = Valves_2020 Then 'valve plot
        XminScale = ScanValues(2, 0)
        XmaxScale = ScanValues(2, ScanNumber - 1)
        If (XminScale > XmaxScale) Then ''reverse
          delayTime = XminScale
          XminScale = XmaxScale
          XmaxScale = delayTime
        End If
        cwSlideDynamicZoom.axis.Minimum = XminScale
        cwSlideDynamicZoom.axis.Maximum = XmaxScale
        cwSlideDynamicZoom.Pointers(2).Value = XminScale + 0.85 * (XmaxScale - XminScale)
        cwSlideDynamicZoom.Pointers(1).Value = XminScale + 0.15 * (XmaxScale - XminScale)
        TextXunits.Text = "mA"
        cwSlideDynamicZoom.Caption = "Valve current, mA"
        cwGraphDataAcq.Axes.Item(1).FormatString = ""
        cwRelAbsPosition.Value = False
    End If
    line_counter = ScanNumber
'    cwGraphDataAcq.Axes.Item(2).AutoScaleNow
'    cwGraphDataAcq.Axes.Item(1).SetMinMax XminScale, XmaxScale

    If PC_RealTime = True Then
      If MonitoringTestType = Valves_2020 Then 'valve plot
        temp_str = "recorded by Analyzer each " & cwValveMeasTime.Text & " sec"
      Else
        temp_str = "recorded by Analyzer each " & cwContCMDperiod.Text & " sec"
      End If
    Else
        temp_str = "CONTROLLER recorded data, Analyzer saved"
    End If
    temp_str = "Loaded file '" & FileName & "'; " & ScanNumber - 1 & " data lines; " & temp_str
    txtStatus.Text = temp_str
    Exit Sub
error_format:
    Close #DataAcq_file
    File_Reading = False
    MsgBox ("Unrecognizable file format")

End Sub

Sub ChangeMonitoringTestType(test_type As Integer)
'line#,InfCnt,EhxCnt, Inflat_mA,Exhaust_mA, PSI_1,PSI_2,PSI_3,PSI_4,PSI_in,
    MonitoringTestType = test_type
    cmb_TestType.ListIndex = test_type
    If MonitoringTestType = Monitor_Pos Then Cmb_X_scale.ListIndex = 1    '1st column == time
'    If MonitoringTestType = Buffer_rec Then Cmb_X_scale.ListIndex = 0     'line number
    If MonitoringTestType = Valves_2020 Then Cmb_X_scale.ListIndex = 4    'show plot versus Inflation mA '4th column == current mA'= 2
'    If MonitoringTestType = Valve_box Then Cmb_X_scale.ListIndex = 2      '2nd column == current mA'= 3
    If MonitoringTestType = Monitor_Pressure Then Cmb_X_scale.ListIndex = 1 '1st column == time
End Sub

Sub Calc_SlopeCorrection()
    Dim ch As Integer
    Dim wanted_slope
    wanted_slope = FrmValveTestResult.CwStandardSlope.Value
    avg_slope = 0
    For ch = 0 To 3
        avg_slope = avg_slope + inflation_slope(ch) * 0.25
    Next ch
    'avg_slope can be used if we need to bring all slopes to same common slope
    For ch = 0 To 3
        If ((inflation_slope(ch) <> -99) And (inflation_slope(ch) <> 0)) Then
            Slope_Correction(ch) = wanted_slope / inflation_slope(ch)
            FrmValveTestResult.lbl_SlopeCorrection(ch) = Format(Slope_Correction(ch), "0.#0")
        Else
            FrmValveTestResult.lbl_SlopeCorrection(ch) = "N/A"
        End If
    Next
End Sub


'    Dim Input_pressure_at_valves_OFF As Single
'    Dim deflation_slope(0 To 3) As Single
'    Dim shutdown_pressure(0 To 3) As Single
'    Dim inflation_slope(0 To 3) As Single
'    Dim fully_open_pressure(0 To 3) As Single
'    Dim Input_pressure_at_valves_FULL_ON As Single

Function Analyze_Valve_plot() As Long '
    On Error Resume Next
    'Public Enum ValveTestResults
    '    NO_ERROR = 0
    '    ERR_INP_PRESSURE_TOO_LOW
    '    ERR_DEFL_PRESSURE_DID_NOT_CROSS_40
    '    ERR_DEFL_PRESSURE_DID_NOT_CROSS_20
    '    ERR_VALVES_WERE_NOT_COMPLITELY_SHUT
    '    ERR_INFL_PRESSURE_DID_NOT_CROSS_20
    '    ERR_INFL_PRESSURE_DID_NOT_CROSS_40
    'End Enum
    Dim error_code As Long
    Dim time_at_40psi(0 To 7) As Double ' first 4 when pressure just got below criteria, last 4 when pressure was just above criteria to calculate more precise crossing time
    Dim time_at_20psi(0 To 7) As Double
    Dim time_down_at5psi(0 To 3) As Double
    Dim time_goingup_at5psi(0 To 3) As Double
    Dim mA_at_20psi(0 To 7) As Single
    Dim mA_at_40psi(0 To 7) As Single

    Dim index_at_40psi(0 To 7) As Long ' first 4 when pressure just got below criteria, last 4 when pressure was just above criteria to calculate more precise crossing time
    Dim index_at_20psi(0 To 7) As Long '
    Dim index_down_at5psi(0 To 3) As Long
    Dim index_goingup_at5psi(0 To 3) As Long
    Dim index_mA_at_20psi(0 To 7) As Long
    Dim index_mA_at_40psi(0 To 7) As Long

    Dim pressure As Single
    Dim mA As Single

    Dim delta_t As Double
    Dim time_correction As Double
    Dim delta_p As Single
    Dim delta_mA As Single
    Dim mA_correction As Double

    Dim ch As Long
    Dim i As Long
    Dim tmp_indx As Long
    Dim min_pressure As Double
    Dim min_pressure_indx As Long
    Dim max_i As Long

    error_code = NO_ERROR
    Analyze_Valve_plot = error_code
    max_i = UBound(plt_X_coord)
    If (max_i = 0) Then GoTo show_results
    Call FrmValveTestResult.setDefaults
    For ch = 0 To 3
        For i = 0 To max_i
            pressure = ScanValuesToPlot(ch, i)
            If pressure > CwUpperPSIthreshold.Value Then
                time_at_40psi(ch) = ScanTimeStamps(i) 'time stamp will be updated until pressure just above 40 psi
                time_at_40psi(ch + 4) = ScanTimeStamps(i + 1) 'time stamp assuming mext sample pressure will be below 40 psi
                index_at_40psi(ch) = i
                index_at_40psi(ch + 4) = i + 1
            Else
             If pressure > CwLowerPSIthreshold.Value Then
                time_at_20psi(ch) = ScanTimeStamps(i) 'time stamp will be updated until pressure drops below 20 psi
                time_at_20psi(ch + 4) = ScanTimeStamps(i + 1) 'time stamp assuming mext sample pressure will be below 20 psi
                index_at_20psi(ch) = i
                index_at_20psi(ch + 4) = i + 1
             Else
              If pressure > 5 Then
                time_down_at5psi(ch) = ScanTimeStamps(i) 'time stamp will be updated until pressure drops below 5 psi
                index_down_at5psi(ch) = i
              Else
                Exit For
              End If
             End If
            End If
        Next i
 ' protection if params are not right and pressure does not cross 40, 20, and 5 psi
        If (index_at_20psi(ch) > index_down_at5psi(ch)) Or (index_at_40psi(ch) > index_down_at5psi(ch)) Then
            error_code = ERR_VALVES_WERE_NOT_COMPLITELY_SHUT 'error - pressure has not been below 5 psi
            Analyze_Valve_plot = error_code
        Else
' this would not happen if pressure even was not below 5 psi
'            If (index_at_40psi(ch) >= max_i) Then
'                error_code = ERR_DEFL_PRESSURE_DID_NOT_CROSS_40 'error - pressure has not been below 40 psi
'                Analyze_Valve_plot = error_code
'    '            index_at_40psi(ch) = max_i - 1
'    '            index_at_40psi(ch + 4) = max_i - 2
'            Else
'                If (index_at_20psi(ch) >= max_i) Then
'                    error_code = ERR_DEFL_PRESSURE_DID_NOT_CROSS_20 'error - pressure has not been below 20 psi
'                    Analyze_Valve_plot = error_code
'        '            index_at_20psi(ch) = max_i - 1
'        '            index_at_20psi(ch + 4) = max_i - 2
'                End If
'            End If
        End If
'        If error_code <> NO_ERROR Then
'            GoTo show_results
'        End If

 ' here we found all downword parameters, find upword parameters

 ' find minimum pressure index
        min_pressure = 100
        For tmp_indx = 0 To max_i
            If ScanValuesToPlot(ch, tmp_indx) < min_pressure Then
                min_pressure = ScanValuesToPlot(ch, tmp_indx)
                min_pressure_indx = tmp_indx
            End If
        Next tmp_indx

' start from min pressure
        For i = min_pressure_indx To max_i
            pressure = ScanValuesToPlot(ch, i)
            mA = plt_X_coord(i)
            If pressure < 5 Then
                time_goingup_at5psi(ch) = ScanTimeStamps(i) 'time stamp will be updated until pressure increases above 5 psi
                index_goingup_at5psi(ch) = i
            Else
             If pressure < CwLowerPSIthreshold.Value Then
                mA_at_20psi(ch) = mA 'mA will be updated until pressure increases above 20 psi, so the last recoded will be just below 20 mA
                mA_at_20psi(ch + 4) = plt_X_coord(i + 1) 'mA assuming next pressure will be above 20 psi
                index_mA_at_20psi(ch) = i
                index_mA_at_20psi(ch + 4) = i + 1
             Else
              If pressure < CwUpperPSIthreshold.Value Then
                mA_at_40psi(ch) = mA 'mA will be updated until pressure increases above 40 psi, so the last recoded will be just below 40 mA
                mA_at_40psi(ch + 4) = plt_X_coord(i + 1) 'mA assuming mext pressure will be above 40 psi
                index_mA_at_40psi(ch) = i
                index_mA_at_40psi(ch + 4) = i + 1
              Else
                Exit For
              End If
             End If
            End If
        Next i
 ' protection if params are not right and pressure does not cross 40, 20, and 5 psi
        If (index_goingup_at5psi(ch) < index_down_at5psi(ch)) Or (index_at_40psi(ch) > index_down_at5psi(ch)) Then
            error_code = ERR_VALVES_WERE_NOT_COMPLITELY_SHUT 'error - pressure has not been below 5 psi
            Analyze_Valve_plot = error_code
        End If
        If (index_mA_at_20psi(ch) >= max_i) Then
            error_code = ERR_INFL_PRESSURE_DID_NOT_CROSS_20 'error - pressure has not been below 40 psi
            Analyze_Valve_plot = error_code
'            index_at_40psi(ch) = max_i - 1
'            index_at_40psi(ch + 4) = max_i - 2
        End If
        If (index_mA_at_40psi(ch) >= max_i) Then
            error_code = ERR_INFL_PRESSURE_DID_NOT_CROSS_40 'error - pressure has not been below 20 psi
            Analyze_Valve_plot = error_code
'            index_at_20psi(ch) = max_i - 1
'            index_at_20psi(ch + 4) = max_i - 2
        End If
'        If error_code <> NO_ERROR Then
'            GoTo show_results
'        End If
    Next ch

    Input_pressure_at_valves_OFF = ScanValuesToPlot(4, 0)
    FrmValveTestResult.lbl_Inp_PSI_valves_OFF.Caption = Format(Input_pressure_at_valves_OFF, "#.#0")

    Input_pressure_at_valves_FULL_ON = ScanValuesToPlot(4, max_i)
    FrmValveTestResult.lbl_Inp_PSI_valves_FULL_ON.Caption = Format(Input_pressure_at_valves_FULL_ON, "#.#0") ' needs to update first before pressure fully open

' now calculate more precise time moments and mA when pressure crosses test criteria 20 and 40 psi
' update this delay first, because FrmValveTestResult uses it as criterion
    FrmValveTestResult.LblDelay.Caption = Format(CMDperiod, "0.0")
    For ch = 0 To 3
        If (index_at_40psi(ch + 4) <= max_i) And (index_at_20psi(ch + 4) <= max_i) Then
            delta_p = ScanValuesToPlot(ch, index_at_40psi(ch + 4)) - ScanValuesToPlot(ch, index_at_40psi(ch))
            delta_t = time_at_40psi(ch + 4) - time_at_40psi(ch)
            time_correction = (CwUpperPSIthreshold.Value - ScanValuesToPlot(ch, index_at_40psi(ch))) / delta_p  'must be between zero and 1
            time_at_40psi(ch) = time_at_40psi(ch) + delta_t * time_correction

            delta_p = ScanValuesToPlot(ch, index_at_20psi(ch + 4)) - ScanValuesToPlot(ch, index_at_20psi(ch))
            delta_t = time_at_20psi(ch + 4) - time_at_20psi(ch)
            time_correction = (CwLowerPSIthreshold.Value - ScanValuesToPlot(ch, index_at_20psi(ch))) / delta_p 'must be between zero and 1
            time_at_20psi(ch) = time_at_20psi(ch) + delta_t * time_correction
            If (time_at_40psi(ch) - time_at_20psi(ch)) Then
                deflation_slope(ch) = (CwUpperPSIthreshold.Value - CwLowerPSIthreshold.Value) / (time_at_40psi(ch) - time_at_20psi(ch)) ' in psi/sec
                FrmValveTestResult.lbl_DeflationSlope_PSI_per_sec(ch).Caption = Format(deflation_slope(ch), "0.#0")
            Else
                deflation_slope(ch) = -99 '-99 means error
                FrmValveTestResult.lbl_DeflationSlope_PSI_per_sec(ch).Caption = "noTimeInfo"
            End If
        Else 'deflation did not cross 40 or 20 psi
            deflation_slope(ch) = -99 '-99 means error
            FrmValveTestResult.lbl_DeflationSlope_PSI_per_sec(ch).Caption = "--N/A--"
        End If
        'shutdown_pressure is measured at middle index between pressure goes down below 5 psi and pressure goes up above 5 psi
        shutdown_pressure(ch) = ScanValuesToPlot(ch, (index_down_at5psi(ch) + index_goingup_at5psi(ch)) / 2)
        FrmValveTestResult.lbl_PSI_deflated(ch).Caption = Format(shutdown_pressure(ch), "0.#0")

        Valve_opening_mA(ch) = plt_X_coord(index_goingup_at5psi(ch))
        FrmValveTestResult.lbl_Opening_Current(ch).Caption = Format(Valve_opening_mA(ch), "#0.#0")

        delta_mA = mA_at_20psi(ch + 4) - mA_at_20psi(ch)
        'delta_t = index_mA_at_20psi(ch + 4) - index_mA_at_20psi(ch)
        delta_p = ScanValuesToPlot(ch, index_mA_at_20psi(ch + 4)) - ScanValuesToPlot(ch, index_mA_at_20psi(ch))
        If (delta_p <> 0) Then
            mA_correction = (20 - ScanValuesToPlot(ch, index_mA_at_20psi(ch))) / delta_p 'must be between zero and 1
        Else
            mA_correction = 0
        End If
        mA_at_20psi(ch) = mA_at_20psi(ch) + delta_mA * mA_correction

        delta_mA = mA_at_40psi(ch + 4) - mA_at_40psi(ch)
        'delta_t = index_mA_at_40psi(ch + 4) - index_mA_at_40psi(ch)
        delta_p = ScanValuesToPlot(ch, index_mA_at_40psi(ch + 4)) - ScanValuesToPlot(ch, index_mA_at_40psi(ch))
        mA_correction = (40 - ScanValuesToPlot(ch, index_mA_at_40psi(ch))) / delta_p 'must be between zero and 1
        mA_at_40psi(ch) = mA_at_40psi(ch) + delta_mA * mA_correction

        If ((mA_at_40psi(ch) - mA_at_20psi(ch)) <> 0) And (mA_at_40psi(ch) <> 0) And (mA_at_20psi(ch) <> 0) Then
            inflation_slope(ch) = (CwUpperPSIthreshold.Value - CwLowerPSIthreshold.Value) / (mA_at_40psi(ch) - mA_at_20psi(ch)) ' in psi/mA
            FrmValveTestResult.lbl_InflationSlope_PSI_per_mA(ch).Caption = Format(inflation_slope(ch), "0.#0")
        Else
            inflation_slope(ch) = -99 'means error
            FrmValveTestResult.lbl_InflationSlope_PSI_per_mA(ch).Caption = "--N/A--"
        End If

        fully_open_pressure(ch) = ScanValuesToPlot(ch, max_i)
        FrmValveTestResult.lbl_PSI_at_FULLY_OPEN(ch).Caption = Format(fully_open_pressure(ch), "#.#0")
    Next ch
    Call Calc_SlopeCorrection
show_results:

    FrmValveTestResult.Show
    Select Case error_code
        Case ERR_INP_PRESSURE_TOO_LOW
        i = MsgBox("Input pressure is " & Format(ScanValues(4, 0), "0.#0") & " psi. It is recommended to" & vbCrLf & _
                   "increase input pressure > 80 psi (standard 90 psi)" & vbCrLf & _
                   "and repeat the test, continue?", _
                    vbExclamation, "Input pressure < 80 psi")

        Case ERR_DEFL_PRESSURE_DID_NOT_CROSS_40
        i = MsgBox("While deflating, pressure did not cross " & CwUpperPSIthreshold.Text & " psi" & vbCrLf & _
                   "SOME RESULTS WERE NOT RECALCULATED!", _
                   vbExclamation, "RESULTS MIGHT BE WRONG!")

        Case ERR_DEFL_PRESSURE_DID_NOT_CROSS_20
        i = MsgBox("While deflating, pressure did not cross " & CwLowerPSIthreshold.Text & " psi" & vbCrLf & _
                   "SOME RESULTS WERE NOT RECALCULATED!", _
                   vbExclamation, "RESULTS MIGHT BE WRONG!")

        Case ERR_VALVES_WERE_NOT_COMPLITELY_SHUT
        i = MsgBox("While deflating, pressure did not cross 5 psi" & vbCrLf & _
                   "Valves were not complitely shut!" & vbCrLf & _
                   "Recommended to increase 'Interval, sec' " & vbCrLf & _
                   "and / or decrease 'Start, mA' than repeat test", _
                   vbExclamation, "RESULTS MIGHT BE WRONG!")

        Case ERR_INFL_PRESSURE_DID_NOT_CROSS_20
        i = MsgBox("While inflating, pressure did not cross 20 psi" & vbCrLf & _
                   "Maybe not enough input pressure" & vbCrLf & _
                   "Check pressure and tubing connections than repeat test", _
                   vbExclamation, "RESULTS MIGHT BE WRONG!")

        Case ERR_INFL_PRESSURE_DID_NOT_CROSS_40
        i = MsgBox("While inflating, pressure did not cross 40 psi" & vbCrLf & _
                   "Maybe not enough input pressure" & vbCrLf & _
                   "Check pressure and tubing connections than repeat test", _
                   vbExclamation, "RESULTS MIGHT BE WRONG!")

        Case Else   ' Case NO_ERROR: Other values.
    End Select

End Function
Private Sub CmdShowAnalysis_Click()
    FrmValveTestResult.Show
End Sub

Private Sub txtV1sn_Change()
  If (txtV1sn.Text = "") Then
    txtV1sn.BackColor = vbWhite
    FrmValveTestResult.Lbl_V1sn = "N/A"
  Else
    If (IsNumber(txtV1sn.Text)) Then
        txtV1sn.BackColor = vbWhite
        FrmValveTestResult.Lbl_V1sn = txtV1sn.Text
    Else
        txtV1sn.BackColor = vbYellow
    End If
  End If
End Sub

Private Sub txtV2sn_Change()
  If (txtV2sn.Text = "") Then
    txtV2sn.BackColor = vbWhite
    FrmValveTestResult.Lbl_V2sn = "N/A"
  Else
    If (IsNumber(txtV2sn.Text)) Then
      txtV2sn.BackColor = vbWhite
      FrmValveTestResult.Lbl_V2sn = txtV2sn.Text
    Else
      txtV2sn.BackColor = vbYellow
    End If
  End If
End Sub

Private Sub txtV3sn_Change()
  If (txtV3sn.Text = "") Then
    txtV3sn.BackColor = vbWhite
    FrmValveTestResult.Lbl_V3sn = "N/A"
  Else
    If (IsNumber(txtV3sn.Text)) Then
       txtV3sn.BackColor = vbWhite
       FrmValveTestResult.Lbl_V3sn = txtV3sn.Text
    Else
       txtV3sn.BackColor = vbYellow
    End If
  End If
End Sub

Private Sub txtV4sn_Change()
  If (txtV4sn.Text = "") Then
    txtV4sn.BackColor = vbWhite
    FrmValveTestResult.Lbl_V4sn = "N/A"
  Else
    If (IsNumber(txtV4sn.Text)) Then
      txtV4sn.BackColor = vbWhite
      FrmValveTestResult.Lbl_V4sn = txtV4sn.Text
    Else
      txtV4sn.BackColor = vbYellow
    End If
  End If
End Sub

