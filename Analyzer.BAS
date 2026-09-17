Attribute VB_Name = "Module1"
Option Explicit
Public Buffer As String 'SIZE IS LIMITED BY FrmTerminal.Timer1_Timer -
'it checks and limits size: If Len(Buffer) > 32000 Then Buffer = Mid(Buffer, 32000) 'cut the beginning of buffer if it gets too long

Public response As String
Public cmd_sent As Boolean  'need this to send only one command like "ltf>x"
Public p_cmd_as_PN As Boolean  'p command responds as version = true, actual pulse = false
Public SendString As String
Public waitresponse As Boolean
Public GUI_doing_command As Boolean ' =true when request is sent, but OK is not received yet. Prevent getting response from one command to the other
Public CONTROLLER_doing_CMD As Boolean 'when controller sends >~Doing CMD
Public Init_sys_data As Boolean  'initial read and update of controls should not trigger reaction (i.e. sending command) on their change
Public Test_in_progress As Boolean
Public MotorSensorTest_active As Boolean 'when initiated driver motor and measure sensor
Public MotorSensorTest_READY As Boolean 'when initiated driver motor and measure sensor

Public FormFiltersVisible As Boolean
Public FormGainsVisible As Boolean
Public FormTerminalVisible As Boolean
Public AnalyzerVisible As Boolean
Public FormPOSmonitoringVisble As Boolean
Public Menu_Active As Boolean 'when controller shows Advanced menu on Terminal
Public Const PI As Double = 3.14159265358979

' MY COMM declarations
Public str_MouseX As String
Public str_MouseY As String
Public str_Help As String
Public strPortNum As String
Public strPortStatus As String
Public strPortSetting As String
Public strBaud As String
Public strParity As String * 1
Public strBitnum As String * 1
Public strBitstop As String
Public Port_chosen As Boolean
Public FW_supports_menu_preffix As Boolean
Public HasFloorFF As Boolean    'IK20251029 added expansion for floor FF

'Any non-zero value is considered as true. In VB 6,
'True has a numeric value of -1. False has a numeric value of 0.
'The reason for this is because the Boolean data type is stored as a 16-bit signed integer.
Public Enum ConnType
    ConnTelnet = 0  ' it is "False" (was boolean type)
    ConnCOMport = 1 ' it is "True"
    ConnSimDLL = 2
    ConnDEMO = 4
End Enum
Public Windows_is_64bit As Boolean
Public ConnectionType As Integer '=true (or NOT zero) if serial port, =false (or 0) if telnet
Public strTelnet_IP_ADR As String
Public strTelnet_PORT As String

Public Const DLL_Name As String = "ElectroDampSim"
Public DLL_version As String
Public DLL_is_available As Boolean
Public DLL_is_Initialized As Boolean
Public DLL_runs_in_separate_thread As Boolean
Public DLL_CallBackInitialized As Boolean
Public Demo_Simul_string As String
Public DLL_CharArrived As Boolean
Public CharFromDLL As String ' it is actually Integer

Public ID_valid As Boolean '=true if vers returns a string containing "95-"
Public Port As Long
Public idComDev As Long
Public Application_Dir As String
Public DataFileName As String
Public MonitoringTestType As Long   'on frmMonitoring
Public ExcitationAxis As Long       'where to apply excitation
Public Excitation_status As Boolean ' global excitation status for monitoring and main forms
Public BNC_0_index As Long          'index of bncd0=BNC_0_index command
Public BNC_1_index As Long          'index of bncd0=BNC_1_index command

' plot number assignment
Public Const CurrentPlot_Num As Integer = 1
Public Const PreviousPlot_Num As Integer = 2
Public Const RefPlot_Num As Integer = 3
Public Const OneFilterTF_Num As Integer = 4
Public Const Axis_TF_Num As Integer = 5
Public Const Prediction_TF_Num As Integer = 6
Public Const Lower_Lim_TF_Num As Integer = 7
Public Const Upper_Lim_TF_Num As Integer = 8
Public Const Coherence_TF_Num As Integer = 9
Public Const PrevCoherence_TF_Num As Integer = 10

Public ChName(256) As String
'filter frame variables
Public Const LOOP_FREQUENCY As Double = 5000  'for ElectroDamp
Public PI_x2_OVER_FREQUENCY As Double ' = PI * 2 / LOOP_FREQUENCY
Public UseSingleCalculations As Boolean ' IK20221007 old way of calculating
Public FilterAxis As Long           'filter axis to change parameters
Public FilterNumberInChain As Long  'filter number in the axis to change parameters
Public FilterParamNumber As Long    '0 to 10, filter user parameter number 0-5, 0(Filter Type, integer), 1-5 params, floats,6-10 filter coefficients
Public FilterTypeNumber As Long     'filter number in the axis to change parameters
Public TestGain As Double
Public TestFrequency As Double
Public TestPulsePosSlope As Double
Public TestPulseNegSlope As Double
Public LTFtestOutputName As String
Public LTFtestInputName As String
Public LTFadvanced As Boolean
Public LTFtest2ndInputName As String
Public SecondInputSelection As Boolean
Public BufferLEN As Long
Public Locale_tmp_number As Double
Public LastStringSent As String
Public LastStringReceived As String
Public SystemFWversion As String
Public LastKeyPressed As Long
Public SystemType  As Long 'gets value from vers string if it has "STACIS" and assigns value from SysType enum
Public SystemName  As String 'gets value from vers string
Public PhaseReversed As Boolean '=true if 180 added, =false if direct (default)
Public Const TEST_LTF_ARRAY_LENGTH As Long = 200
    '*********************************************************
Public Freq_Reference(TEST_LTF_ARRAY_LENGTH) As Double

Public OneFilter_Mag(TEST_LTF_ARRAY_LENGTH) As Double  ' data arrays
Public OneFilter_Phase(TEST_LTF_ARRAY_LENGTH) As Double

Public Freq_points(TEST_LTF_ARRAY_LENGTH) As Double ' used in prediction, based on Freq_Reference
Public Axis_Mag(TEST_LTF_ARRAY_LENGTH) As Double
Public Axis_Phase(TEST_LTF_ARRAY_LENGTH) As Double
Public M3v_Cmd_Word As Long
Public AlarmWord As Long
Public DeflationOffset_mA(4) As Double
Public PneumFilterChain As Boolean
Public ProdTestStatus As Boolean
Public self_test_word As Long 'is 32-bit Prod_Test_BITS
Public ShowExpertControls As Boolean
Public M3v_Diagnostic_running As Long
Public M3v_measurement_done As Long
Public InputChosen As Long
Public Input2ndChosen As Long
Public OutputChosen As Long
Public DiagnosticAxis As Long 'X pos, Y pos, tZ pos, Z pos, tx pos, tY pos, X vel, Y vel, tZ vel, Z vel, tX vel, tY vel

Public Const CMD_LEN As Integer = 4
Public Const Acknowledgment_token As String = ">~"     ' not basic (more than 2 bytes) command ends with ">~OK xxxxx" & vbCrLf or ">~ERR xxxxx" & vbCrLf line
Public Const end_of_ScanTransmission As String = "OK" & vbCrLf     ' position monitoring scan data END
Public Const OK_message As String = ">~OK" & vbCrLf

Public Const FILTER_PARAMS_TOTAL   As Long = 11 'filt type, 5 params, 5 coefficients
'#define PID_P_gain FilterGain
'#define PID_I_gain freq_1
'#define PID_D_gain freq_2
Public Const MAX_FILTERS_IN_AXIS   As Long = 6
'7 filters (0 to 6) in Axis chain, 10 parameters (0 to 5)-user, 6=a1,7=a2,8=b0,9=b1,0xA=b2
Public Const FB_FILTERS_IN_AXIS As Long = MAX_FILTERS_IN_AXIS - 2 '0 to 4 in ElectroDamp, the last filter is used for stage FF adaptive algorithm
Public Const FF_ADAPTIVE_FILT_NUM As Long = MAX_FILTERS_IN_AXIS - 1 '#5

Public EnablePrecisionLog As Boolean
Public PrecisionLogFile As String

Public AirIsoPresent As Boolean
'Public Const NUM_POS_FB_AXES As Long = 6
'Public Const NUM_VEL_FB_AXES As Long = 6
'Public Const NUM_FF_AXES As Long = 4
Dim gISIDE As Boolean
'Public SimulatedFiltersArray(NUM_POS_FB_AXES + NUM_VEL_FB_AXES + NUM_FF_AXES - 3, MAX_FILTERS_IN_AXIS - 1) As float_iir ' (0-5, 0-9) red from controller filters

Private Const OF_EXIST         As Long = &H4000
Private Const OFS_MAXPATHNAME  As Long = 128
Private Const HFILE_ERROR      As Long = -1


Private Type OFSTRUCT
    cBytes As Byte
    fFixedDisk As Byte
    nErrCode As Integer
    Reserved1 As Integer
    Reserved2 As Integer
    szPathName(OFS_MAXPATHNAME) As Byte
End Type

Public Type tComplex
    Re As Double
    Im As Double
End Type

Public Type tF_coeff
    a1 As Double
    a2 As Double
    b0 As Double
    b1 As Double
    b2 As Double
End Type

Public Type tFreqHz_MagTimes_PhaseDeg
    mag As Double
    Phase As Double
    Freq_Hz As Double
End Type

Public Const USER_PARAM_NUMBER As Long = 5
Type float_iir
    ftyp As Long
    par(USER_PARAM_NUMBER - 1) As Double 'IK20221007 changed from Single for Low SampleRate
    a1 As Double 'IK20221007 changed from Single for Low SampleRate
    a2 As Double
    b0 As Double
    b1 As Double
    b2 As Double
    'coefficients are rarely change
'    Private HistBuff As p_iir_history
End Type

'//define indexes for float_iir.par[index]
'    Const freq_1 As Long = 1
'    Const freq_2 As Long = 2
'    Const Q_f_1  As Long = 3
'    Const Q_f_2  As Long = 4
'    Const FilterGain   As Long = 0
Public Enum filt_par
    FilterGain = 0
    freq_1 = 1
    freq_2 = 2
    Q_f_1 = 3
    Q_f_2 = 4
End Enum
Const PID_P_gain As Long = FilterGain
Const PID_I_gain As Long = freq_1
Const PID_D_gain As Long = freq_2

Const POSITIVE_NEAR_ZERO_FLOATING_NUMBER As Single = (0.0000001)

Enum filt_types
    GAIN_ONLY = 0
    LPF_1stOr = 1        ' STACIS (normalized)
    HPF_1stOr = 2        ' STACIS (normalized)
    LEADLAG_1stOr = 3    ' STACIS !NOT normalized! LEAD or LAG DOES depend on F1,F2: if F1>F2 - it is "-\_LAG, HiFreq, LoFreq"; if F1<F2 - it is "_/-LEAD, LoFreq, HiFreq";
    LEADLAG_2ndOr_wPk = 4 ' STACIS !NOT normalized! LEAD or LAG DOES depend on F1,F2: if F1>F2 - it is "-\_LAG, HiFreq, LoFreq"; if F1<F2 - it is "_/-LEAD, LoFreq, HiFreq";
'ADDITIONAL filters for ELDAMP
    LEAD_1stOrNorm = 5  ' __/-- (normalized) does not depend on F1<F2 or F1>F2, the lowest of F1,F2 is low corner freq, the highest of F1,F2 - high corner freq
    LAG_1stOr = 6       '--\__ NOT normalized! potentially high gain at low freq! does not depend on F1<F2 or F1>F2, the lowest of F1,F2 is low corner freq, the highest of F1,F2 - high corner freq
    LPF_2ndOr_wPk = 7   '(normalized)
    HPF_2ndOr_wPk = 8   '(normalized)
    PID = 9             '(Leaky)
    PEAK_on_FLAT = 10   ' ____/\____ peak regulates by Q1, smooth phase transition
    NOTCH = 11          ' ----\/---- depth regulates by Q1, smooth phase transition
'      second column in the menu
    LEAD_2ndOr_wPk ' __/-- (normalized) does not depend on F1<F2 or F1>F2, the lowest of F1,F2 is low corner freq, the highest of F1,F2 - high corner freq
    LAG_2ndOr_wPk  ' --\__ (normalized) does not depend on F1<F2 or F1>F2, the lowest of F1,F2 is low corner freq, the highest of F1,F2 - high corner freq
    NOTCH_Sharp    ' ----\/---- Very deep (~ -40 dB), width regulates by Q1, sharp phase transition
    BPF            '"roof shape BPF" +20dB/dec-peak-20dB/dec; F1=peak freq, Q1= peak magnitude
    VELOCITY       ' approximate system velocity transfer function
    Position       ' approximate system position transfer function
    COMBINED_LAG1_LEAD1 ', // combines LAG and LEAD 1st order into single 2nd order --\____/--;  --\__ determined by F1,Q1 (float frequency); __/-- determined by F2,Q2
    DIRECT_COEFF ',       // directly set a, b factors:
' NOT IN USE does not depend on F1<F2 or F1>F2, the lowest of F1,F2 is low corner freq, the highest of F1,F2 - high corner freq
'    LEAD_1stOrSemiNorm '__/-- (SemiNormalized) potentially high gain at high freq! Gain=1.0 @ (F1+F2)/2.
'    LAG_1stOrNorm  ' --\__ (normalized) does not depend on F1<F2 or F1>F2, the lowest of F1,F2 is low corner freq, the highest of F1,F2 - high corner freq
'    LAG_1stOrSemiNorm  '--\__ Semi Normalized! potentially high gain at low freq! Gain=1.0 @ (F1+F2)/2. Does not depend on F1<F2 or F1>F2, the lowest of F1,F2 is low corner freq, the highest of F1,F2 - high corner freq
'    GEO_OPT        '(geo optimization) gain = 1 when f > F2, INTENDED USE: F1=0.2, F2=0.5; it is LEDLAG --2nd order with Q1=Q2~0.1
    MaxFilterTyp
End Enum

Public Enum SysType
    UNKNOWN_system = 0
    WIN450_system = 1
    ECP2_system = 2
    EXIGE_system = 3
    LOTUS_system = 4
    CDSAX_system = 5
    EUCLID_system = 6
    CONDOR_system = 7
End Enum

' IK20250813 attempt to create User Defined type as either Single or Double
' see comment at 'Set_float_iir()
#Const UseDouble = True
#If UseDouble Then
    Public Type t_filt_var
        Value As Double
    End Type
#Else
    Public Type t_filt_var
        Value As Single
    End Type
#End If

Private Declare Function OpenFile Lib "kernel32" (ByVal lpFileName As String, _
                        lpReOpenBuff As OFSTRUCT, ByVal wStyle As Long) As Long

'============= M3 DLL functions ======================================
'' DLL displays in MsgBox string sent ByRef
'Private Declare Sub DisplayStringByRef Lib "ElectroDampSim.dll" _
'         (sMyString As String)
'
'' DLL displays in MsgBox string sent ByVal
'Private Declare Sub DisplayStringByVal Lib "ElectroDampSim.dll" _
'         (ByVal sMyString As String)
'
'Private Declare Sub FillString Lib "ElectroDampSim.dll" _
'         (ByVal sMyString As String, ByVal cBufferSize As Long)
'
'Private Declare Sub AlterString Lib "ElectroDampSim.dll" _
'         (sMyString As String)
'
''BSTR __stdcall ReturnBstr()
'Private Declare Function ReturnBstr Lib "ElectroDampSim.dll" _
'         () As String

Private Declare Sub DLL_Init Lib "ElectroDampSim.dll" ()

Private Declare Function DLL_IsRunning Lib "ElectroDampSim.dll" () As Integer

Private Declare Sub DLL_Stop Lib "ElectroDampSim.dll" ()

' returns char array converted to BSTR ar parameter
'Private Declare Sub DLL_version Lib "ElectroDampSim.dll" _
'       (ByVal sMyString As String, ByVal cBufferSize As Long)

' returns BSTR (string of chars with length info)
Private Declare Function VB6_DLL_version Lib "ElectroDampSim.dll" () As String

' returns char * to global array char DLL_response_buf[5000], accepts char * commands
Private Declare Function Controller Lib "ElectroDampSim.dll" _
         (ByVal sMyString As String) As String

' returns BSTR ; accepts LPCSTR command
Private Declare Sub VB6controller Lib "ElectroDampSim.dll" _
         (ByVal cmd As String, ByVal rsp As String, ByVal rspSize As Long)

' reads internal DLL output buffer, returns BSTR ; clears internal DLL output buffer
Private Declare Sub DLL_ReadBuf Lib "ElectroDampSim.dll" _
         (ByVal rsp As String, ByVal rspSize As Long)

' send command to DLL and exits without waiting for response
Private Declare Sub DLL_SendCMD Lib "ElectroDampSim.dll" _
         (ByVal cmd As String)

Private Declare Sub InitStringCallBackFuncPtr Lib "ElectroDampSim.dll" _
        (ByVal lpfnMyFunc As Long)

Private Declare Sub InitCharCallBackFuncPtr Lib "ElectroDampSim.dll" _
        (ByVal lpfnMyFunc As Long)

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    ByRef Destination As Any, _
    ByRef Source As Any, _
    ByVal Length As Long)


Public Function IsIDE() As Boolean '
        IsIDE = False
        'This line is only executed if running in the IDE and then returns True
        Debug.Assert CheckIDE
        If gISIDE Then
            IsIDE = True
        End If
End Function

Private Function CheckIDE() As Boolean ' this is a helper function for Public Function IsIDE()
        gISIDE = True 'set global flag
        CheckIDE = True
End Function

'==================================================================
Public Function Is_DLL_available() As Boolean
    Dim exists As Boolean
    Dim start_dir As String
    start_dir = App.Path
    exists = FileExists(start_dir + "\" + DLL_Name & ".dll")
    If exists Then
        exists = FileExists(start_dir + "\" + DLL_Name & ".lib")
    End If
'    Buffer = CurDir ' just to check
    DLL_is_available = exists ' global boolean
    Is_DLL_available = exists ' function return
End Function

' >>>>>>>>> DO NOT USE StringCallBack_to_DLL <<<<<<<<<<<<<<<
'Public Sub Init_StringCallBack_to_DLL() ' >>> BUG>>> after few commands VB6 throws " out of String Space (Error 14)". I use char-by-char transfer instead
''if call initialization, output from DLL happens via CallBack and after few commands
'' VB6 throws  "out of string space (error 14)"
'    InitStringCallBackFuncPtr AddressOf DLLhasString_CallBackFunction
'    DLL_CallBackInitialized = True
'End Sub
' >>>>>>>>> DO NOT USE StringCallBack_to_DLL <<<<<<<<<<<<<<<

Public Sub Init_CharCallBack_to_DLL()
    InitCharCallBackFuncPtr AddressOf DLLhasCHAR_CallBackFunction
    DLL_CallBackInitialized = True
End Sub

Public Sub Init_DLL()
    If Is_DLL_available Then
        Call DLL_Init
        DLL_version = VB6_DLL_version
        Call Init_CharCallBack_to_DLL
        DoEvents
        DLL_is_Initialized = True
    Else
        DLL_is_Initialized = False
    End If
End Sub

Public Sub Cmd_Stop_DLL()
    DLL_is_Initialized = False
    Call DLL_Stop
End Sub

Public Function Is_DLL_running() As Boolean
    Dim running_1_not_0 As Integer
     running_1_not_0 = DLL_IsRunning
     If running_1_not_0 = 1 Then
        Is_DLL_running = True
     Else
        Is_DLL_running = False
     End If
     DLL_runs_in_separate_thread = Is_DLL_running
End Function

Public Function Get_DLL_version() As String
    Dim DLLvers As String
    If Is_DLL_available Then
        DLLvers = VB6_DLL_version
    Else
        DLLvers = "ElectroDampSim DLL or LIB not found"
    End If
'    MsgBox DLLvers, vbInformation, "DLL version as BSTR"
    Get_DLL_version = DLLvers
End Function

Public Function Send_DLL_cmd(cmd As String) As String
    Dim resp As String ' wchar_t wide char 16 bits
    If Is_DLL_available Then
        resp = Space$(65520) ' was 5000; increasing up to max size does not help.... still about 5100 chars
    '    Dim Cmd As String
    '    Cmd = TextCMD.Text
        Call VB6controller(cmd, resp, 65520) '  was 100000; response is returned into Buffer by DLLhasCHAR_CallBackFunction VB6 has size limit 65526
    '    TextResp.Text = resp
    '    MsgBox resp, vbInformation, "VB got response"
    Else
        resp = "ElectroDampSim DLL or LIB not found"
    End If
    If (DLL_CallBackInitialized = True) Then
        Send_DLL_cmd = Buffer
    Else
        Send_DLL_cmd = resp
    End If
    resp = ""
End Function

Public Sub Send_CMD_toDLLwoWaitForResponse(cmd As String)
    Call DLL_SendCMD(cmd) ' send command to DLL and exits without waiting for response
End Sub

Public Function Read_DLL_whole_buffer() As String
    Dim rsp_len As Long
    Dim resp As String ' wchar_t wide char 16 bits
    If DLL_is_available Then
        resp = Space$(65520) ' was 10000; increasing up to max size does not help.... still about 5100 chars
        Call DLL_ReadBuf(resp, 65520) ' was 10000; increasing up to max size does not help.... still about 5100 chars
        '    TextResp.Text = resp
        '    MsgBox resp, vbInformation, "VB got response"
    Else
        '    resp = "ElectroDampSim DLL or LIB not found"
    End If
    rsp_len = Len(resp)
    If rsp_len <> 0 Then
        rsp_len = InStr(resp, Chr(0))
        If rsp_len <> 0 Then ' test if DLL trasfers zeroes.
            resp = Replace(resp, Chr(0), " ") ' break point here
        End If
    End If
    resp = RTrim(resp) ' cut end of long buffer containing spaces
    resp = Replace(resp, Chr(0), " ") ' replace possible 0 on the end with space
    resp = RTrim(resp) ' cut end of string which was 0 and now it is space
    Read_DLL_whole_buffer = resp
End Function

' >>>>>>>>> DO NOT USE StringCallBack_to_DLL <<<<<<<<<<<<<<<
'Function DLLhasString_CallBackFunction(ByVal DLLstring As String, ByVal strlen As Long) As Long
'' after few commands VB6 throws " out of String Space (Error 14)"
'    Dim resp As String ' wchar_t wide char 16 bits
'   Dim tpos As Long
'    resp = Space$(500)
'    Call DLL_ReadBuf(resp, Len(resp)) 'it includes EoS 0x00, cut resp up to it
'    tpos = InStr(resp, Chr(0))
'    If tpos <> 0 Then resp = Left(resp, tpos - 1)
'    If frmTerminal.Visible = True Then 'Print sym
'        frmTerminal.Receive (resp)
'    Else ' add to Buffer without printing
'        Buffer = Buffer + resp ' do not use trim, resp might have valid space
'    End If
'    DLLhasString_CallBackFunction = strlen
'    resp = ""
'End Function

Sub DLLhasCHAR_CallBackFunction(ByVal NewChar As Long)
    Dim resp As String ' wchar_t wide char 16 bits
    resp = Chr(NewChar)
'    If frmTerminal.Visible = True Then 'Print sym
'        ' if DLL has separate thread, this thread comes here and calls non static member
'        ' and gives: Unhandled exception at 0x72949197 in TMC Analyzer.exe: 0xC0000005: Access violation reading location 0x00000018.
'        ' it actually needs Invoke method
'        frmTerminal.Receive (resp)
'    Else ' add to Buffer without printing
'HOWEVER if DLL main is running in separate thread, THIS FUNCTION MUST NOT be called from VB6 thread,
'because of deadlock using DLL_CharArrived


'    If FormTerminalVisible = True Then
'        If DLL_CharArrived = False Then
'            CharFromDLL = resp ' copy to global var
'            DLL_CharArrived = True ' signal to frmTerminal
'            Do
'              DoEvents ' wait for frmTerminal to consume char
'            Loop Until (DLL_CharArrived = True)
'        End If

    If frmTerminal.Visible = True Then 'Print sym
        frmTerminal.Receive (resp)
    Else ' add to Buffer without printing
        Buffer = Buffer + resp ' do not use trim, resp might have valid space
    End If
End Sub

Public Function FileExists(ByVal Fname As String) As Boolean
    Dim lRetVal As Long
    Dim OfSt As OFSTRUCT

    lRetVal = OpenFile(Fname, OfSt, OF_EXIST)
    If lRetVal <> HFILE_ERROR Then
        FileExists = True
    Else
        FileExists = False
    End If
End Function

Public Sub FFT(FR() As Double, FI() As Double, LnN As Long)

  'THE FAST FOURIER TRANSFORM
  'copyright 1997-1999 by California Technical Publishing
  'published with  permission from Steven W Smith, www.dspguide.com
  'GUI by logix4u , www.logix4u.net
  'modified by logix4u, www.logix4.net
  'Upon entry, N% contains the number of points in the DFT, FR( ) and
  'FI( ) contain the real and imaginary parts of the input. Upon return,
  'FR( ) and FI( ) contain the DFT output. All signals run from 0 to N%-1.
Dim ND2 As Long
Dim i As Long, j As Long, k As Long, l As Long
Dim le As Long, le2 As Long
Dim ip As Long, NP_FFT As Long

Dim s As Double, ur As Double, ur1 As Double, ui As Double, wr As Double, wi As Double
Dim tr As Double, ti As Double, DivN As Double

  NP_FFT = 2 ^ LnN  'Number of points for use in FFT

  ND2 = NP_FFT / 2  'Half points
  j = 1

  For i = 1 To NP_FFT - 1

    If i < j Then
      s = FR(i)
      FR(i) = FR(j)
      FR(j) = s
      s = FI(i)
      FI(i) = FI(j)
      FI(j) = s
    End If

    k = ND2

    While k < j
      j = j - k
      k = k / 2
    Wend

  j = j + k
  Next i

  For l = 1 To LnN 'Loop for each stage

    le = 2 ^ l
    le2 = le / 2
    ur = 1
    ui = 0
    wr = Cos(PI / le2) 'Calculate sine & cosine values
    wi = -Sin(PI / le2)
    For j = 1 To le2 'Loop for each sub DFT
      i = j
      While (i <= NP_FFT)
          ip = i + le2
          tr = FR(ip) * ur - FI(ip) * ui 'Butterfly calculation
          ti = FR(ip) * ui + FI(ip) * ur
          FR(ip) = FR(i) - tr
          FI(ip) = FI(i) - ti
          FR(i) = FR(i) + tr
          FI(i) = FI(i) + ti
          i = i + le
      Wend
      ur1 = ur * wr - ui * wi
      ui = ur * wi + ui * wr
      ur = ur1
      Next j
    Next l  '{l - l>kke}

End Sub

Public Function Avg(Aray As Variant, Optional Size As Long = -1) As Double
    On Error Resume Next
    Dim DataSum As Double
    Dim counter As Long
    If UBound(Aray) < 1 Then ReDim Aray(2)
    If Size = -1 Then Size = UBound(Aray) 'argument is missing
    DataSum = 0
    For counter = 0 To Size - 1
        DataSum = DataSum + Aray(counter)
    Next counter
    Avg = DataSum / Size
End Function

Public Function FindMaxMin(inArray As Variant, StartIndex As Long, maxV As Variant, minV As Variant, MaxIndex As Long, MinIndex As Long, Optional IgnoreZeroes As Boolean = False) As Long
Dim i As Long
Dim max As Double
Dim min As Double
Dim ArrSize As Long
max = -1E+308   'set max and min possible values
min = 1E+308
ArrSize = UBound(inArray)

If StartIndex < ArrSize Then
    For i = StartIndex To ArrSize - 1
      If i > ArrSize Then Exit For
        If max < inArray(i) Then
            max = inArray(i)
            MaxIndex = i
        End If
        If IgnoreZeroes = True And inArray(i) = 0 Then
            i = i + 1
        Else
            If min > inArray(i) Then
                min = inArray(i)
                MinIndex = i
            End If
        End If
     Next i
     minV = min
     maxV = max
        FindMaxMin = 1  'sucsess
 Else
        FindMaxMin = 0
 End If
End Function

Public Function St_Dev(Aray As Variant, Optional Size As Long = -1, Optional IgnoreZeroes As Boolean = False) As Double
    On Error Resume Next
    Dim sum_of_squares As Double
    Dim array_sum As Double
    Dim diff As Double
    Dim array_index As Long
    Dim final_size As Long
    If UBound(Aray) < 1 Then ReDim Aray(2)

    If Size = -1 Then Size = UBound(Aray) 'argument is missing ****
    final_size = Size
    sum_of_squares = 0
    array_sum = 0
    For array_index = 0 To Size - 1
        array_sum = array_sum + Aray(array_index)
        sum_of_squares = sum_of_squares + (Aray(array_index) ^ 2)
        If IgnoreZeroes = True And Aray(array_index) = 0 Then
            final_size = final_size - 1
        End If
    Next array_index
    If final_size < 2 Then final_size = 2 'prevent calc errors
    diff = (final_size * sum_of_squares - array_sum ^ 2)
    If diff > 0 Then    'because of limited accuracy, sometimes diff can be negative
        St_Dev = (diff / (final_size * (final_size - 1))) ^ 0.5
    Else
        St_Dev = 0
    End If
End Function

Public Function BIN(inbyte As Long) As String
' subprogram translates byte to BIN$ format, 127 = 01111111
    Dim m As String, b As String
    Dim j As Long, k As Long, bt As Long

     bt = inbyte
     m = ""
     For j = 0 To 7
        If bt < 2 Then GoTo end_conv
        b = Str(bt - 2 * (Int(bt / 2)))
        bt = Int(bt / 2)
        m = b + m
     Next j
end_conv:
    m = Str(bt) + m
    For k = j To 6
        m = " 0" + m
    Next k
    BIN = m
End Function

Sub Get_Serial_Input() 'read everything into Buffer
    If (ConnectionType = ConnCOMport) Then
        Buffer = Buffer & frmSplash.SerialComm.Input
    End If
    If (ConnectionType = ConnTelnet) Then
        If frmSplash.Winsock.State = sckConnected Then
            Dim strData As String
            frmSplash.Winsock.GetData strData, vbString
            Buffer = Buffer & strData
        End If
    End If
    If (ConnectionType = ConnSimDLL) Then
'        Buffer = Buffer & Read_DLL_whole_buffer 'Buffer is updated it happens automatically - no need here
    End If
End Sub

Public Sub Clean_COM_Buffer()
    Call Get_Serial_Input 'it reads everything into Buffer
    Buffer = "" 'clean the Comm port RxBuffer
End Sub

Public Sub Send_GUI_command(cmd As String) 'sets doing_cmd=True and sends cmd to port
    GUI_doing_command = True ' this must be cleaned when command finishes
    SendString = cmd
    If (ConnectionType = ConnCOMport) Then 'COM port
        If frmSplash.SerialComm.PortOpen = True Then frmSplash.SerialComm.Output = cmd 'send command
    End If
    If (ConnectionType = ConnTelnet) Then
        If frmSplash.Winsock.State = sckConnected Then
            frmSplash.Winsock.SendData (cmd)
        Else
            Analyzer.txtStatus.Text = "Telnet is not ready"
        End If
    End If
    If (ConnectionType = ConnSimDLL) Then
        Call Send_CMD_toDLLwoWaitForResponse(cmd) ' -!- >>>>>
    End If

    cmd_sent = True
    LastStringSent = LastStringSent + cmd 'for test
    If InStr(LastStringSent, "menu" & vbCr) <> 0 Then
        Menu_Active = True 'sending "menu" command
    End If
    If InStr(LastStringSent, Chr(27) & "q" & vbCr) <> 0 Then
        Menu_Active = False 'sending esc+q (leave "menu") command
    End If
    If cmd = vbCr Then
        If ((Menu_Active = False) And (FrmHistory.ChkLocalEcho = 1)) Then
            FrmHistory.AddToHistory ("->" & LastStringSent & "<ENTER>")
        End If
        LastStringSent = ""
    End If
End Sub

 Sub open_port_if_closed()
    On Error Resume Next
    Dim StartTime As Single
    Dim WaitTime As Single
    If (ConnectionType = ConnCOMport) Then 'COM port
        If frmSplash.SerialComm.PortOpen = False Then
            'MsgBox ("Port COM " & Port & ": closed - opening...")
            frmSplash.SerialComm.CommPort = Port
            frmSplash.SerialComm.PortOpen = True
            If frmSplash.SerialComm.PortOpen = False Then ' port cannot be open
                strPortStatus = " COM" & Port & " not avail"
            Else
                strPortStatus = " open"
            End If
        End If
    End If ' ConnCOMport
    If (ConnectionType = ConnTelnet) Then 'telnet
        If frmSplash.Winsock.State = sckClosed Then
            frmSplash.Winsock.RemoteHost = strTelnet_IP_ADR
            frmSplash.Winsock.RemotePort = strTelnet_PORT
            frmSplash.Winsock.Connect
            StartTime = Timer
            Do 'small delay
                DoEvents
                WaitTime = Timer - StartTime
                DoEvents
            Loop Until (WaitTime > 0.05)

            If frmSplash.Winsock.State <> sckConnecting And frmSplash.Winsock.State <> sckConnected And frmSplash.Winsock.State <> sckConnectionPending Then
                MsgBox ("Telnet connection error: " & (frmSplash.Winsock.State))
                strPortStatus = " Telnet not avail"
'updated once in frmSplash.search_COM_ports Analyzer.cwCOMdisconnect.Caption = Demo_Simul_string
                Analyzer.cwCOMdisconnect.Caption = Demo_Simul_string
            Else
                strPortStatus = " open"
            End If
        End If
    End If ' ConnTelnet

    If (ConnectionType = ConnSimDLL) Then 'simulation DLL
        If (DLL_is_Initialized = False) Then Init_DLL 'DLL checks if it is already initialized and ignores second attempt
    End If
End Sub

Public Sub Close_Port()
    If (ConnectionType = ConnCOMport) Then 'COM port
        If (frmSplash.SerialComm.PortOpen = True) Then
            frmSplash.SerialComm.PortOpen = False        ' close the port to change settings
    '            FramePort.Caption = "COM" & frmSplash.SerialComm.CommPort & " DISconnected"
        End If
    Else 'telnet
        If frmSplash.Winsock.State <> sckClosed Then ' close the port to change settings
            frmSplash.Winsock.Close
        End If
    End If
    If (ConnectionType = ConnSimDLL) Then 'simulation DLL
        If (DLL_is_Initialized = True) Then Cmd_Stop_DLL 'DLL works in real time on separate thread. This stops that thread
    End If
    strPortStatus = " closed"
    Call frmTerminal.Update_Form_Caption
End Sub

Sub portproblems()
    If DLL_is_available = True Then
        If (ConnectionType = ConnCOMport) Then
          response = MsgBox("Switch to Simulation mode?", vbYesNo, strPortNum & " port is not available, probably used by other application")
        Else
          response = MsgBox("Switch to Simulation mode?", vbYesNo, "IP address: " & strTelnet_IP_ADR & ", port" & strTelnet_PORT & " is not available")
        End If
        If response = vbYes Then
            ConnectionType = ConnSimDLL
            If (DLL_is_Initialized = False) Then Init_DLL 'DLL checks if it is already initialized and ignores second attempt
            Analyzer.cwCOMdisconnect.Caption = Demo_Simul_string
            Analyzer.UpdateConnectControls (False)
        End If
    Else
        If (ConnectionType = ConnCOMport) Then
          response = MsgBox("Switch to DEMO mode?", vbYesNo, strPortNum & " port is not available, probably used by other application")
        Else
          response = MsgBox("Switch to DEMO mode?", vbYesNo, "IP address: " & strTelnet_IP_ADR & ", port" & strTelnet_PORT & " is not available")
        End If
        If response = vbYes Then
            ConnectionType = ConnDEMO
            Analyzer.cwCOMdisconnect.Caption = Demo_Simul_string
            Analyzer.UpdateConnectControls (False)
        End If
    End If
End Sub

Function ReadPortSetup(Optional SkipDialog As Boolean = False) As Boolean 'DataFileName must be set in advance if argument is true
  On Error Resume Next
'user wants to read a test file
'First get the filename
    Dim DefaultFilename As String
    Dim FileName As String
    Dim i As Long
    Dim temp_str As String

    ReadPortSetup = False ' prepare for the worse - file eroor, not recognized, etc
    If SkipDialog = False Then
        DefaultFilename = "TMC.port"
        frmSplash.dlgCommonDialog.CancelError = True
        frmSplash.dlgCommonDialog.DialogTitle = "Choose the Port Setup file to load"
        ' frmSplash.dlgCommonDialog.FileName = DefaultFilename
        frmSplash.dlgCommonDialog.InitDir = Application_Dir
        frmSplash.dlgCommonDialog.DefaultExt = "port"
        frmSplash.dlgCommonDialog.Filter = "Data File (*.port)|*.port|All files (*.*)|*.*"
        frmSplash.dlgCommonDialog.Action = 1
        'dlgCommonDialog.ShowSave
        If Err = 32755 Then 'Cancel is selected
            Exit Function
        Else
    '        txtFileName.Text = "Reading measurement from file '" & dlgCommonDialog.FileName & " ....."
        End If
        'SaveFileName becomes "C:\test\my_test.txt"

    '    SaveFileName = Left(dlgCommonDialog.FileName, Len(dlgCommonDialog.FileName) - 4) ' remove ".txt" from full file path
        DataFileName = frmSplash.dlgCommonDialog.FileName
        'SaveFileName becomes "C:\test\my_test"
    Else
        If FileExists(".\TMC.port") Then DataFileName = "TMC.port" ' ?? changes from "C:\In_Out\Software\DC-2020_ED_Analyzer_VB6\TMC.port" to "TMC.port"
    End If 'SkipDialog
    If DataFileName = "" Then
        Call MsgBox("File name is empty", vbCritical, "File 'TMC.port' not found or corrupt")
        Exit Function
    End If
    i = InStrRev(DataFileName, "\")
    FileName = Right(DataFileName, Len(DataFileName) - i) 'i.e. "my_test"

    'now lets open the file and read the data
    Open DataFileName For Input As #1

    Line Input #1, temp_str '    "TMC PORT SETUP " FileName
    ' decode line
    i = InStr(UCase(temp_str), "TMC PORT SETUP")
    If i <> 0 Then
'        Print #1, "PORT COM"' Port
'        Print #1, strPortSetting
        Line Input #1, temp_str '    "PORT COM"; Port OR "TELNET IP: 192.168.1.37, port 2020"
        i = InStrRev(UCase(temp_str), "M")
        If i <> 0 Then 'string PORT COM
            If (IsNumber(Mid(temp_str, i + 1))) Then  'port number
                Port = Mid(temp_str, i + 1)
            Else
                i = MsgBox("Port Number 'COM" & temp_str & "' is not recognized in the file" & vbCrLf & DataFileName, , "Wrong Port Settings in the file")
                Close #1
                Exit Function
            End If

            Line Input #1, temp_str '    "115200,N,8,1"
            If (IsNumeric(Left(temp_str, 3))) Then
                strPortSetting = temp_str
                strPortNum = "COM" & LTrim(Str(Port))
                ConnectionType = ConnCOMport
            Else
                i = MsgBox("Port Settings Line '" & temp_str & "' is not recognized in the file" & vbCrLf & DataFileName, , "Wrong Port Settings in the file")
                Close #1
                Exit Function
            End If
        Else ' no 'M'
            i = InStrRev(temp_str, "IP:")
            If i <> 0 Then 'string TELNET IP:
                strPortNum = LTrim(Mid(temp_str, i + 3))
                strTelnet_IP_ADR = strPortNum
            Else
                i = MsgBox("IP Address '" & temp_str & "' is not recognized in the file" & vbCrLf & DataFileName, , "Wrong Port Settings in the file")
                Close #1
                Exit Function
            End If

            Line Input #1, temp_str 'TELNET PORT: 2020
            i = InStrRev(UCase(temp_str), "RT:")
            If i <> 0 Then 'string TELNET PORT:
                If (IsNumeric(Mid(temp_str, i + 4))) Then
                    strPortSetting = LTrim(Mid(temp_str, i + 3)) 'telnet port
                    strTelnet_PORT = strPortSetting
                    ConnectionType = ConnTelnet
                End If
            Else
                i = MsgBox("Settings Line '" & temp_str & "' is not recognized in the file" & vbCrLf & DataFileName, , "Wrong Port Settings file")
                Close #1
                Exit Function
            End If
        End If
    Else
        i = MsgBox("File Signature '" & temp_str & "' is not recognized", , "Wrong Port Settings in the file")
        Close #1
        Exit Function
    End If
    Close #1
    ReadPortSetup = True ' success
'open_port_if_closed sets port name according to the TMC.port, but user might want to change it
'    Call open_port_if_closed
End Function


Sub SavePortSetup()
    On Error Resume Next
    'user wants to save a test file
    'First get the filename
    Dim DefaultFilename As String
    Dim SaveFileName As String
    Dim FileName As String
    Dim i As Long

    DefaultFilename = "TMC.port"
    frmSplash.dlgCommonDialog.CancelError = True
    frmSplash.dlgCommonDialog.DialogTitle = "Choose name for the Port Setup"
    frmSplash.dlgCommonDialog.InitDir = Application_Dir
    frmSplash.dlgCommonDialog.DefaultExt = "port"
    frmSplash.dlgCommonDialog.Filter = "Port Setup File (*.port)|*.port|All files (*.*)|*.*"
    frmSplash.dlgCommonDialog.FileName = DefaultFilename 'if enabled,  FILE NAME SHOULD NOT HAVE ':'
    frmSplash.dlgCommonDialog.ShowSave
    If Err = 32755 Then 'Cancel is selected
        Exit Sub
    Else
'        txtFileName.Text = "File(s)=" & dlgCommonDialog.FileName
    End If
    'SaveFileName becomes "C:\test\my_test.txt"

'    SaveFileName = Left(dlgCommonDialog.FileName, Len(dlgCommonDialog.FileName) - 4) ' remove ".txt" from full file path
    SaveFileName = frmSplash.dlgCommonDialog.FileName
    'SaveFileName becomes "C:\test\my_test"
    i = InStrRev(SaveFileName, "\")
    FileName = Right(SaveFileName, Len(SaveFileName) - i) 'i.e. "my_test"

'    SaveFileName = dlgCommonDialog.FileName
    'now lets open the file and write the data
    Open SaveFileName For Output As #1 'Values
        Print #1, "TMC PORT SETUP "; FileName
        If (ConnectionType = ConnCOMport) Then
            strPortSetting = strBaud + "," + strParity + "," + strBitnum + "," + strBitstop
            strPortNum = frmSplash.s_port.Text
            Print #1, "PORT "; strPortNum
            Print #1, strPortSetting
        Else
            Print #1, "TELNET IP: "; strPortNum
            Print #1, "TELNET PORT: "; strTelnet_PORT
        End If
        Print #1, "Date: "; Date$
        Print #1, "Time: "; Time$
    Close #1
End Sub

Function Make180phaseShift(phaseIN As Double) As Double
    Dim phaseOut As Double
    phaseOut = phaseIN + 180
    If phaseOut > 180 Then
        phaseOut = phaseOut - 360
    Else '<=0
'        phaseOut = phaseIN - 180
'        If phaseOut < -180 Then phaseOut = phaseOut + 180
    End If
    Make180phaseShift = phaseOut
End Function

'-----------------------------------------------
' Complex Number Math Functions
'-----------------------------------------------

Public Function Complex(ByVal dReal As Double, ByVal dImag As Double) As tComplex
'Set the real and imaginary numbers in a complex number type
    Complex.Re = dReal
    Complex.Im = dImag
End Function

Public Function IM_NEG(ByRef arg1 As tComplex) As tComplex
'Compute unary - operation on a single complex number,
'returns a complex number
    IM_NEG = Complex(-(arg1.Re), -(arg1.Im))
End Function

Public Function IM_ADD(ByRef arg1 As tComplex, ByRef arg2 As tComplex) As tComplex
'Addition of two complex numbers, returns a complex number
    IM_ADD = Complex(arg1.Re + arg2.Re, arg1.Im + arg2.Im)
End Function

Public Function IM_SUB(ByRef arg1 As tComplex, ByRef arg2 As tComplex) As tComplex
'Subtraction of two complex numbers, returns a complex number
    IM_SUB = Complex(arg1.Re - arg2.Re, arg1.Im - arg2.Im)
End Function

Public Function IM_MULT(ByRef arg1 As tComplex, ByRef arg2 As tComplex) As tComplex
'Multiplication of two complex numbers, returns a complex number
    IM_MULT = Complex((arg1.Re * arg2.Re) - (arg1.Im * arg2.Im), (arg1.Re * arg2.Im) + (arg2.Re * arg1.Im))
End Function

Public Function IM_DIV(ByRef arg1 As tComplex, ByRef arg2 As tComplex) As tComplex
'division of two complex numbers:
' multiply the numerator and denominator by the complex conjugate of the denominator,
'for example, with z_1=a+bi and z_2=c+di, z=z_1/z_2 is given by
'numerator and denomerator  two complex numbers, returns a complex number
    Dim Denom_Conj As tComplex
    Dim Denom_Mag As Double
    Denom_Conj = Complex(arg2.Re, -arg2.Im)
    Denom_Mag = arg2.Re * arg2.Re + arg2.Im * arg2.Im
    IM_DIV = Complex(((arg1.Re * arg2.Re) + (arg1.Im * arg2.Im)) / Denom_Mag, ((arg1.Im * arg2.Re) - (arg1.Re * arg2.Im)) / Denom_Mag)
End Function
' returns radians, -Pi... +Pi
Function Atan2(X As Double, Y As Double) As Double
    Atan2 = 0
    'now if x and y are positive
    If (X > 0) Then  ' 1st and 4th quadrant
        Atan2 = Atn(Y / X)
    Else 'left side of circle, X<0
        If (Y >= 0) Then ' 2nd quadrant
            Atan2 = (PI + Atn(Y / X))
        Else            ' Y<0, 3rd quadrant
            Atan2 = (-PI + Atn(Y / X))
        End If
    End If
End Function

'===============================================================================================================================
' this function calculates magnitude-phase prediction based on filter coefficients,
' magnitude returns as "times", not dB, phase returns in degrees
'===============================================================================================================================
Public Function Mag_Ph_vs_Freq(Coeff As float_iir, freq As Double, LoopFreq As Double) As tFreqHz_MagTimes_PhaseDeg
    On Error Resume Next ' if frequency = 0 it gives division by zero
    Dim Omega As Double
    Dim H As tComplex
    Dim Num As tComplex
    Dim Den As tComplex

    Omega = freq * 2 * PI / LoopFreq
    Num = Numerator_Complex(Coeff, Omega)
    Den = Denominator_Complex(Coeff, Omega)
    H = IM_DIV(Num, Den)
    Mag_Ph_vs_Freq.mag = Sqr(H.Re * H.Re + H.Im * H.Im)
    If (Mag_Ph_vs_Freq.mag = 0) Then Mag_Ph_vs_Freq.mag = 0.00001 'we take a Log and it cannot be a zero' 1e-5 means it is an error
    If (H.Re <> 0) Then
        Mag_Ph_vs_Freq.Phase = -Atan2(H.Re, H.Im) * 180 / PI  ' returns between -180 and +180
        'Mag_Ph_vs_Freq.Phase = -Atn(H.Im / H.Re) * 180 / Pi   ' returns between -90 and +90
    Else
        Mag_Ph_vs_Freq.Phase = -360 'this means it is an error
    End If
End Function

Public Function Numerator_Complex(Coeff As float_iir, Omega As Double) As tComplex
    Dim NumRe As Double
    Dim NumIm As Double
    NumRe = Coeff.b0 + Coeff.b1 * Cos(Omega) + Coeff.b2 * Cos(2 * Omega)
    NumIm = Coeff.b1 * Sin(Omega) + Coeff.b2 * Sin(2 * Omega)
    Numerator_Complex = Complex(NumRe, NumIm)
End Function

Public Function Denominator_Complex(Coeff As float_iir, Omega As Double) As tComplex
    Dim DenRe As Double
    Dim DenIm As Double
    DenRe = 1 - Coeff.a1 * Cos(Omega) - Coeff.a2 * Cos(2 * Omega)
    DenIm = -Coeff.a1 * Sin(Omega) - Coeff.a2 * Sin(2 * Omega)
    Denominator_Complex = Complex(DenRe, DenIm)
End Function
''''''''''''''' End of complex math '''''''''''''''''''''''''''''''''''''''''''''''

''================FILTER SETTING ++++++++++++++++++++++++++

' -------------------------
' Internal: Append to log file
' -------------------------
Private Sub LogPrecisionLoss(ByVal opName As String, ByVal input1 As Double, _
                             Optional ByVal input2 As Variant, _
                             Optional ByVal dblResult As Double, _
                             Optional ByVal sngResult As Double)
    On Error Resume Next
    If EnablePrecisionLog And LenB(PrecisionLogFile) > 0 Then
        Dim diff As Double
        Dim PPMerror As Double
        diff = dblResult - sngResult
        PPMerror = diff * 1000000#
        Dim f As Integer
        f = FreeFile
        Open PrecisionLogFile For Append As #f
        If IsMissing(input2) Then
            Print #f, opName & "," & Format$(input1, "0.###############") & _
                      ",," & Format$(dblResult, "0.###############") & _
                      "," & Format$(sngResult, "0.###############") & _
                      "," & Format$(diff, "0.###############")
        Else
            Print #f, opName & "," & Format$(input1, "0.###############") & _
                      "," & Format$(CDbl(input2), "0.###############") & _
                      "," & Format$(dblResult, "0.###############") & _
                      "," & Format$(sngResult, "0.###############") & _
                      "," & Format$(PPMerror, "0.###############")
        End If
        Close #f
    End If
End Sub

' -------------------------
' Precision control
' -------------------------
Public Function CFPrecision(ByVal input_val As Double) As Double
    If (UseSingleCalculations) Then ' Convert to Single, assign to Double
        CFPrecision = CSng(input_val)
    Else ' use Doubles
        CFPrecision = CDbl(input_val)
    End If
End Function


' ============================================================
' Math wrappers with logging
' ============================================================

Public Function FSin(ByVal X As Double) As Double
    Dim dblVal As Double, sngVal As Single
    dblVal = Sin(X)
    sngVal = CSng(dblVal)
    FSin = IIf(UseSingleCalculations, sngVal, dblVal)
    Call LogPrecisionLoss("Sin", X, , dblVal, sngVal)
End Function

Public Function FCos(ByVal X As Double) As Double
    Dim dblVal As Double, sngVal As Single
    dblVal = Cos(X)
    sngVal = CSng(dblVal)
    FCos = IIf(UseSingleCalculations, sngVal, dblVal)
    Call LogPrecisionLoss("Cos", X, , dblVal, sngVal)
End Function

Public Function FTan(ByVal X As Double) As Double
    Dim dblVal As Double, sngVal As Single
    dblVal = Tan(X)
    sngVal = CSng(dblVal)
    FTan = IIf(UseSingleCalculations, sngVal, dblVal)
    Call LogPrecisionLoss("Tan", X, , dblVal, sngVal)
End Function

Public Function FSqr(ByVal X As Double) As Double
    Dim dblVal As Double, sngVal As Single
    dblVal = Sqr(X)
    sngVal = CSng(dblVal)
    FSqr = IIf(UseSingleCalculations, sngVal, dblVal)
    Call LogPrecisionLoss("Sqr", X, , dblVal, sngVal)
End Function

' ============================================================
' Robust Atan2 with logging
' ============================================================
Public Function FAtan2(ByVal Y As Double, ByVal X As Double) As Double
    Const PI As Double = 3.14159265358979
    Dim dblVal As Double, sngVal As Single

    ' Full quadrant handling
    If X > 0 Then
        dblVal = Atn(Y / X)
    ElseIf X < 0 Then
        If Y >= 0 Then
            dblVal = Atn(Y / X) + PI
        Else
            dblVal = Atn(Y / X) - PI
        End If
    Else ' x = 0
        If Y > 0 Then
            dblVal = PI / 2
        ElseIf Y < 0 Then
            dblVal = -PI / 2
        Else
            dblVal = 0 ' undefined, both zero
        End If
    End If

    ' Simulate single precision
    sngVal = CSng(dblVal)
    FAtan2 = IIf(UseSingleCalculations, sngVal, dblVal)

    ' Log precision loss
    Call LogPrecisionLoss("Atan2", Y, X, dblVal, sngVal)
End Function

' ============================================================
' Arithmetic wrappers with logging
' ============================================================
Public Function FAdd(ByVal a As Double, ByVal b As Double) As Double
    Dim dblVal As Double, sngVal As Single
    dblVal = a + b
    sngVal = CSng(dblVal)
    FAdd = IIf(UseSingleCalculations, sngVal, dblVal)
    Call LogPrecisionLoss("Add", a, b, dblVal, sngVal)
End Function

Public Function FSub(ByVal a As Double, ByVal b As Double) As Double
    Dim dblVal As Double, sngVal As Single
    dblVal = a - b
    sngVal = CSng(dblVal)
    FSub = IIf(UseSingleCalculations, sngVal, dblVal)
    Call LogPrecisionLoss("Sub", a, b, dblVal, sngVal)
End Function

Public Function FMul(ByVal a As Double, ByVal b As Double) As Double
    Dim dblVal As Double, sngVal As Single
    dblVal = a * b
    sngVal = CSng(dblVal)
    FMul = IIf(UseSingleCalculations, sngVal, dblVal)
    Call LogPrecisionLoss("Mul", a, b, dblVal, sngVal)
End Function

Public Function FDiv(ByVal a As Double, ByVal b As Double) As Double
    Dim dblVal As Double, sngVal As Single
    dblVal = a / b
    sngVal = CSng(dblVal)
    FDiv = IIf(UseSingleCalculations, sngVal, dblVal)
    Call LogPrecisionLoss("Div", a, b, dblVal, sngVal)
End Function

'How to use logging
'
'
'
'Private Sub Form_Load()
'    ' Enable simulation and logging
'    UseSingleCalculations = True
'    EnablePrecisionLog = True
'    PrecisionLogFile = App.Path & "\PrecisionLog.csv"
'
'    ' Optional: write CSV header
'    Dim f As Integer
'    f = FreeFile
'    Open PrecisionLogFile For Output As #f
'    Print #f, "Operation,Input1,Input2,DoubleResult,SingleResult,Error"
'    Close #f
'
'    ' Example DSP calculation
'    Dim fc As Double, fs As Double, omega As Double
'    fc = 1000#
'    fs = 48000#
'    omega = FMul(FTwoPI(), FDiv(fc, fs))
'    Debug.Print "Omega = "; omega
'End Sub
'>> What you get in PrecisionLog.csv
'Example Output:
'
'Operation , input1, input2, DoubleResult, SingleResult, Error
'Div,1000,48000,0.0208333333333333,0.02083333,

Public Sub set_float_iir(ByRef fp As float_iir)
    On Error Resume Next ' CAUTION !! it can miss divide by zero


'===============================================================
' IK20250813 I simply redefined local variables as Single to simulate TI FPU32
' Now fully routed through CFPrecision / F* wrappers so UseSingleCalculations
' correctly emulates FPU32 single-precision behaviour.
'===============================================================
    '#define omega_1pole omega_1pole 'alias name: omega_Numenator
    '#define omega_2zero omega_2zero 'alias name: omega_DEnominator
    'Dim fp As float_iir
    Dim h1 As Double
    Dim h2 As Double
    Dim t_gain As Double
    Dim s_freq_gain_corr As Double
    Dim cotan_of_half_omega_1 As Double
    Dim freq_1_correction As Double
    Dim freq_2_correction As Double
    Dim tiny_positive_f32_number As Double
    Dim omega_1pole As Double
    Dim omega_2zero As Double
    Dim testVar As Double

    Dim FilterType As Long
    Dim f1 As Double
    Dim f2 As Double
    Dim q1 As Double
    Dim q2 As Double

    f1 = fp.par(freq_1)
    f2 = fp.par(freq_2)
    q1 = fp.par(Q_f_1)
    q2 = fp.par(Q_f_2)
    tiny_positive_f32_number = POSITIVE_NEAR_ZERO_FLOATING_NUMBER

    FilterType = fp.ftyp

    ' Fully wrapped constant
    PI_x2_OVER_FREQUENCY = FDiv(FMul(CFPrecision(PI), 2#), CFPrecision(LOOP_FREQUENCY))

    omega_1pole = CFPrecision(FMul(PI_x2_OVER_FREQUENCY, f1))   'par[1] is filter corner frequency
    omega_2zero = CFPrecision(FMul(PI_x2_OVER_FREQUENCY, f2))

    Call ensure_float_number_is_not_zero(q1, tiny_positive_f32_number) 'If q1 = 0 Then q1 = 0.5 ' prevent division by zero
    Call ensure_float_number_is_not_zero(q2, tiny_positive_f32_number) 'If q2 = 0 Then q2 = 0.5 ' prevent division by zero

    testVar = FMul(q1, q1)
    freq_1_correction = FDiv(1#, FSqr(FAdd(1#, FDiv(0.25, testVar))))   'correction for 2nd order …

    testVar = FMul(q2, q2)   ' IK20250814 fixed bug, was Q2 * Q1
    freq_2_correction = FDiv(1#, FSqr(FAdd(1#, FDiv(0.25, testVar))))

    If (omega_1pole > omega_2zero) Then   'choose bigger of two frequencies for gain correction
        s_freq_gain_corr = FDiv(FAdd(CFPrecision(LOOP_FREQUENCY), f1), CFPrecision(LOOP_FREQUENCY))
    Else
        s_freq_gain_corr = FDiv(FAdd(CFPrecision(LOOP_FREQUENCY), f2), CFPrecision(LOOP_FREQUENCY))
    End If

    h1 = FTan(FMul(omega_1pole, 0.5))
    Call ensure_float_number_is_not_zero(h1, tiny_positive_f32_number)
    cotan_of_half_omega_1 = FDiv(1#, h1)

    t_gain = fp.par(FilterGain)

    ' set default parameters as GAIN_ONLY for start (needed gains will be replaced)
    fp.b2 = 0#
    fp.b1 = 0#
    fp.b0 = t_gain
    fp.a2 = 0#
    fp.a1 = 0#

 Select Case (FilterType)

' ============== STACIS FILTERS: only 4 needed ============================================================================================
  Case LPF_1stOr:           'case 1: sharp drop at high freq, notch at Nyquist, phase = -90
    testVar = FAdd(1#, cotan_of_half_omega_1)  '//IK20250818 tan() and cotan() can be =-1 at 135 and 315 deg
    Call ensure_float_number_is_not_zero(testVar, tiny_positive_f32_number)
    h1 = FDiv(1#, testVar)
    fp.b1 = CFPrecision(FMul(t_gain, h1))
    fp.b0 = fp.b1
    fp.a1 = CFPrecision(FMul(FSub(1#, cotan_of_half_omega_1), h1))
     ' Break '

  Case HPF_1stOr:           'case 2: normalized
    testVar = FAdd(1#, cotan_of_half_omega_1)
    Call ensure_float_number_is_not_zero(testVar, tiny_positive_f32_number)
    h1 = FDiv(1#, testVar)
    fp.a1 = CFPrecision(FMul(FSub(1#, cotan_of_half_omega_1), h1))
    fp.b1 = CFPrecision(FMul(FMul(FSub(0#, cotan_of_half_omega_1), h1), t_gain))
    fp.b0 = FSub(0#, fp.b1) ' == -fp.b1
     ' Break '

  Case LEADLAG_1stOr:      'case 3: normalized maximum unity gain at band pass frequency
'  if (omega_1pole <= omega_2zero) t_gain = omega_1pole / omega_2zero * t_gain'
    fp.a1 = CFPrecision(FSub(omega_2zero, 1#))
'  fp.a2 = 0  'set to 0 above "switch(fp.ftyp){}"
    fp.b0 = t_gain '1# '
    fp.b1 = CFPrecision(FMul(FSub(omega_1pole, 1#), t_gain))
'  fp.b2 = 0  'set to 0 above "switch(fp.ftyp){}"
     ' Break '

  Case LEADLAG_2ndOr_wPk: 'case 4: general form with Q factors (NOT normalized) LEAD or LAG DOES depend on F1,F2: if F1>F2 - it is "-\_LAG, HiFreq, LoFreq"' if F1<F2 - it is "_/-LEAD, LoFreq, HiFreq"'
' Max gain can be bigger than 1 if any of Q is bigger than 0.5, generally max gain approx = Q.
' *****!!***** BE CAREFUL WITH Q!!  IT CAN OVERFLOW IN FIXED POINT CALCULATIONS
'  Normalized to unity gain at frequencies much bigger/smaller than F1 or F2.
'  omega_1pole *= freq_1_correction'  'correction for 2nd order LPF, HPF, Lead, Lag to have 90 degree phase at any Q_1
'  omega_2zero *= freq_2_correction'  'correction for 2nd order LPF, HPF, Lead, Lag to have 90 degree phase at any Q_1

' //IK20250818 checked at the beginning ensure_float_number_is_not_zero(&q2, tiny_positive_f32_number);
    h2 = FSub(1#, FDiv(omega_2zero, FMul(2#, q2)))
    fp.a1 = CFPrecision(FMul(FMul(-2#, h2), FCos(omega_2zero)))
    fp.a2 = CFPrecision(FMul(h2, h2))

    h1 = FSub(1#, FDiv(omega_1pole, FMul(2#, q1)))
    fp.b1 = CFPrecision(FMul(FMul(FMul(-2#, h1), FCos(omega_1pole)), t_gain))
    fp.b2 = CFPrecision(FMul(FMul(h1, h1), t_gain))
     ' Break '
  Case DIRECT_COEFF:  ' directly enter 2nd order coeffs (b0+b1*z^-1+b2*z^-2)/(1+a1*z^-1+a2*z^-2)
    fp.b0 = t_gain
    fp.b1 = f1
    fp.b2 = f2

    fp.a1 = FSub(0#, q1)   'IK 20210301 reversed sign because it is reversed at the end of function
    fp.a2 = FSub(0#, q2)

    ' FilterGain  Q_f_1  Q_f_2 freq_1 freq_2
    ' break '
' ============== STACIS FILTERS END only 4 needed ============================================================================================

' ============== ADDITIONAL FILTERS FOR ELECTRODUMP ==========================================================================================

  Case LEAD_1stOrNorm ', LEAD_1stOrSemiNorm     'case 5: __/-- F1<F2' HPF with max gain = param4 above F2, !! CAREFUL with FIX POINT !!gain at F1 can be much bigger than 1!!
'  Case LEAD_1stOrSemiNorm:  'case 18: --\__ Semi Normalized! potentially high gain at high freq! Gain=1 @ (F1+F2)/2.
'  Does not depend on F1<F2 or F1>F2, the lowest of F1,F2 is low corner freq, the highest of F1,F2 - high corner freq
    h1 = t_gain
    If (omega_1pole >= omega_2zero) Then
        fp.a1 = CFPrecision(FSub(omega_1pole, 1#))
        fp.b1 = CFPrecision(FSub(omega_2zero, 1#))
'        If (fp.ftyp = LEAD_1stOrSemiNorm) Then h1 = t_gain * Sqr(omega_1pole / omega_2zero) '
    Else
        fp.a1 = CFPrecision(FSub(omega_2zero, 1#))
        fp.b1 = CFPrecision(FSub(omega_1pole, 1#))
'        If (fp.ftyp = LEAD_1stOrSemiNorm) Then h1 = t_gain * Sqr(omega_2zero / omega_1pole) '
    End If
    fp.b0 = CFPrecision(h1)
    fp.b1 = CFPrecision(FMul(fp.b1, h1))
     ' Break '

  Case LAG_1stOr: ', LAG_1stOrSemiNorm          'case 6: --\__  NOT normalized gain at lowest frequency, at highest frequency gain = 1
'  Case LAG_1stOrSemiNorm:   'case 20:--\__ Semi Normalized! potentially high gain at low freq! Gain=1 @ (F1+F2)/2.
'  Does not depend on F1<F2 or F1>F2, the lowest of F1,F2 is low corner freq, the highest of F1,F2 - high corner freq
    h1 = t_gain
    If (omega_1pole <= omega_2zero) Then
        fp.a1 = CFPrecision(FSub(omega_1pole, 1#))
        fp.b1 = CFPrecision(FSub(omega_2zero, 1#))
'        If (fp.ftyp = LAG_1stOrSemiNorm) Then h1 = t_gain * Sqr(omega_1pole / omega_2zero) '
    Else
        fp.a1 = CFPrecision(FSub(omega_2zero, 1#))
        fp.b1 = CFPrecision(FSub(omega_1pole, 1#))
'        If (fp.ftyp = LAG_1stOrSemiNorm) Then h1 = t_gain * Sqr(omega_2zero / omega_1pole) '
    End If
    fp.b0 = CFPrecision(h1)
    fp.b1 = CFPrecision(FMul(fp.b1, h1))
     ' Break '

  Case COMBINED_LAG1_LEAD1: '// combines LAG and LEAD 1st order into single 2nd order --\____/--;  --\__ determined by F1,Q1 (float frequency); __/-- determined by F2,Q2
'      omega_1pole = PI_x2_OVER_FREQUENCY * f1; //par[1] is filter corner frequency
'      omega_2zero = PI_x2_OVER_FREQUENCY * f2;

        Dim omega_D_tmp As Double, a11 As Double, b11 As Double ';// temp a11, b11 for the LEG, b01 is gain ==1 by default
        Dim omega_N_tmp As Double, a12 As Double, b12 As Double ';// temp a11, b11 for the LEG, b02 is gain ==1 by default

        omega_D_tmp = FMul(PI_x2_OVER_FREQUENCY, q1)
        a11 = CFPrecision(FSub(omega_1pole, 1#))   '// normally, and here as well, comes from F1
        b11 = CFPrecision(FSub(omega_D_tmp, 1#))   '// from Q1

'//already known         omega_2zero = PI_x2_OVER_FREQUENCY * f2;

        omega_N_tmp = FMul(PI_x2_OVER_FREQUENCY, q2)
        b12 = CFPrecision(FSub(omega_2zero, 1#))
        a12 = CFPrecision(FSub(omega_N_tmp, 1#))

        fp.b2 = CFPrecision(FMul(FMul(b11, b12), t_gain))
        fp.b1 = CFPrecision(FMul(FAdd(b11, b12), t_gain))
        fp.b0 = CFPrecision(t_gain)
        fp.a2 = CFPrecision(FMul(a11, a12))
        fp.a1 = CFPrecision(FAdd(a11, a12))

  Case LPF_2ndOr_wPk:   'case 7: par1 = F1 = corner freq, par2,4 not used, par3 = Q factor Q=0.5 - classic, Q=2, peak approximately 2 times over unity
    h2 = FMul(FMul(omega_1pole, omega_1pole), 0.25) '       temp use (omega_1pole/2)^2
    h1 = FAdd(FDiv(0.25, FMul(q1, q1)), 1#) '   temp use (0.25/Q_1^2 + 1)
    h2 = FMul(FMul(FMul(h2, h1), t_gain), freq_1_correction)  'overall gain

    h1 = FSub(1#, FDiv(omega_1pole, FMul(2#, q1)))

    fp.a1 = CFPrecision(FMul(FMul(-2#, h1), FCos(omega_1pole)))
    fp.a2 = CFPrecision(FMul(h1, h1))
    fp.b0 = CFPrecision(h2)
    fp.b1 = CFPrecision(FMul(2#, h2))
    fp.b2 = CFPrecision(h2)
     ' Break '

  Case HPF_2ndOr_wPk:   'case 8:
    omega_1pole = FMul(omega_1pole, freq_1_correction)
    h1 = FSub(1#, FDiv(omega_1pole, FMul(2#, q1)))

    fp.a1 = CFPrecision(FMul(FMul(-2#, h1), FCos(omega_1pole)))
    fp.a2 = CFPrecision(FMul(h1, h1))

    fp.b0 = CFPrecision(t_gain)
    fp.b1 = CFPrecision(FMul(-2#, t_gain))
    fp.b2 = CFPrecision(t_gain)
     ' Break '

  Case LEAD_2ndOr_wPk:  'case 12:  __/--- , normalized (max gain = Q, can be > 1)
    omega_1pole = FMul(omega_1pole, freq_1_correction) 'correction for 2nd order LPF, HPF, Lead, Lag to have 90 degree phase at any Q_1
    omega_2zero = FMul(omega_2zero, freq_2_correction) 'correction for 2nd order LPF, HPF, Lead, Lag to have 90 degree phase at any Q_2

    Call ensure_float_number_is_not_zero(s_freq_gain_corr, tiny_positive_f32_number)

    If (omega_1pole > omega_2zero) Then ' gain_corr = s_freq_gain_corr * (omega_1pole/omega_2zero)^2
        h1 = FSub(1#, FDiv(omega_1pole, FMul(2#, q1)))
        h2 = FSub(1#, FDiv(omega_2zero, FMul(2#, q2)))
        fp.a1 = CFPrecision(FMul(FMul(-2#, h1), FCos(omega_1pole)))
        fp.a2 = CFPrecision(FMul(h1, h1))

        t_gain = FDiv(t_gain, s_freq_gain_corr)
        h1 = FCos(omega_2zero)    'temp use for b1 calculation
    Else 'omega_1pole <= omega_2zero
        h1 = FSub(1#, FDiv(omega_2zero, FMul(2#, q2)))
        h2 = FSub(1#, FDiv(omega_1pole, FMul(2#, q1)))
        fp.a1 = CFPrecision(FMul(FMul(-2#, h1), FCos(omega_2zero)))
        fp.a2 = CFPrecision(FMul(h1, h1))

        t_gain = FDiv(t_gain, s_freq_gain_corr)
        h1 = FCos(omega_1pole)    'temp use for b1 calculation
    End If
    fp.b0 = CFPrecision(t_gain)
    fp.b1 = CFPrecision(FMul(FMul(FMul(-2#, h2), h1), t_gain))
    fp.b2 = CFPrecision(FMul(FMul(h2, h2), t_gain))
     ' Break '

  Case LAG_2ndOr_wPk:   'case 13:--\__  normalized (max gain = Q, can be > 1)
' Q-factor gain correction: =(1+1/(2*Q_1)^2)/(1+1/(2*Q_2)^2)
    omega_1pole = FMul(omega_1pole, freq_1_correction) 'correction for 2nd order LPF, HPF, Lead, Lag to have 90 degree phase at any Q_1
    omega_2zero = FMul(omega_2zero, freq_2_correction) 'correction for 2nd order LPF, HPF, Lead, Lag to have 90 degree phase at any Q_2

    Call ensure_float_number_is_not_zero(omega_2zero, tiny_positive_f32_number)
    h1 = FDiv(omega_1pole, omega_2zero)
    h1 = FMul(h1, h1)
    h2 = FAdd(1#, FDiv(0.25, FMul(q2, q2)))
    h2 = FDiv(FAdd(1#, FDiv(0.25, FMul(q1, q1))), h2)
    h1 = FMul(h1, h2)   ' h1 includes F2/F1, Q2 and Q1 gain correction

    'reverse F1 and F2 to maintain LEAD shape
    If (omega_1pole <= omega_2zero) Then ' gain_corr = s_freq_gain_corr * (omega_1pole/omega_2zero)^2
        t_gain = FMul(FMul(s_freq_gain_corr, t_gain), h1)  ' h1 includes F2/F1, Q2 and Q1 gain correction
        h1 = FSub(1#, FDiv(omega_1pole, FMul(2#, q1)))
        h2 = FSub(1#, FDiv(omega_2zero, FMul(2#, q2)))
        fp.a1 = CFPrecision(FMul(FMul(-2#, h1), FCos(omega_1pole)))
        fp.a2 = CFPrecision(FMul(h1, h1))
        h1 = FCos(omega_2zero)    'temp use for b1 calculation
    Else 'omega_1pole > omega_2zero
        t_gain = FDiv(FMul(s_freq_gain_corr, t_gain), h1) ' h1 includes F2/F1, Q2 and Q1 gain correction

        h1 = FSub(1#, FDiv(omega_2zero, FMul(2#, q2)))
        h2 = FSub(1#, FDiv(omega_1pole, FMul(2#, q1)))
        fp.a1 = CFPrecision(FMul(FMul(-2#, h1), FCos(omega_2zero)))
        fp.a2 = CFPrecision(FMul(h1, h1))
        h1 = FCos(omega_1pole)    'temp use for b1 calculation
    End If
    fp.b0 = CFPrecision(t_gain)
    fp.b1 = CFPrecision(FMul(FMul(FMul(-2#, h2), h1), t_gain))
    fp.b2 = CFPrecision(FMul(FMul(h2, h2), t_gain))
      ' Break '

  Case PEAK_on_FLAT: ' made from leadlag
    h2 = FAdd(0.89, FDiv(0.064, q1))   '' Experimental automatic second frequency calculation to have unity gain from both sides of peak
    omega_2zero = FMul(omega_1pole, h2)
    h1 = FSub(1#, FDiv(omega_1pole, FMul(2#, q1)))
    h2 = FSub(1#, FDiv(omega_2zero, 2#))
    fp.a1 = CFPrecision(FMul(FMul(-2#, h1), FCos(omega_1pole)))
    fp.a2 = CFPrecision(FMul(h1, h1))
    t_gain = FMul(t_gain, s_freq_gain_corr)
    fp.b0 = CFPrecision(t_gain)
    fp.b1 = CFPrecision(FMul(FMul(FMul(-2#, h2), FCos(omega_2zero)), t_gain))
    fp.b2 = CFPrecision(FMul(FMul(h2, h2), t_gain))
     ' Break '

  Case NOTCH: ' made from leadlag
'  h1  = (q1 - 1)'' automatic second frequency calculation to have unity gain from both sides of notch
'  h2  = 1 / (q1 + 1)'
'  h2  = h2 * h2'
'  omega_2zero = omega_1pole * (1 - h1 * h2)' = (1 - (Q - 1)/((Q + 1)^2))

'    h2 = 0.89 + 0.064 / q1 '' Experimental automatic second frequency calculation to have unity gain from both sides of peak
    h2 = FAdd(FAdd(FAdd(0.8208, FDiv(0.1673, q1)), FMul(0.009105, q1)), FMul(FMul(0.00001825, q1), f1))
    omega_2zero = FMul(omega_1pole, h2)
    h1 = FSub(1#, FDiv(omega_1pole, FMul(2#, q1)))
    h2 = FSub(1#, FMul(omega_2zero, 0.5))
    fp.a1 = CFPrecision(FMul(FMul(-2#, h2), FCos(omega_2zero)))
    fp.a2 = CFPrecision(FMul(h2, h2))

    t_gain = FMul(t_gain, s_freq_gain_corr)

    fp.b0 = CFPrecision(t_gain)
    fp.b1 = CFPrecision(FMul(FMul(FMul(-2#, h1), FCos(omega_1pole)), t_gain))
    fp.b2 = CFPrecision(FMul(FMul(h1, h1), t_gain))
     ' Break '

  Case NOTCH_Sharp: 'Q regulates "width" of a notch
    testVar = FAdd(FAdd(1#, FDiv(cotan_of_half_omega_1, q1)), FMul(cotan_of_half_omega_1, cotan_of_half_omega_1))
    Call ensure_float_number_is_not_zero(testVar, tiny_positive_f32_number)
    h1 = FDiv(1#, testVar)
    h2 = FMul(FMul(2#, FSub(1#, FMul(cotan_of_half_omega_1, cotan_of_half_omega_1))), h1)

    fp.a1 = CFPrecision(h2)
    fp.a2 = CFPrecision(FMul(FAdd(FSub(1#, FDiv(cotan_of_half_omega_1, q1)), FMul(cotan_of_half_omega_1, cotan_of_half_omega_1)), h1))
    fp.b0 = CFPrecision(FMul(FMul(FAdd(1#, FMul(cotan_of_half_omega_1, cotan_of_half_omega_1)), h1), t_gain))
    fp.b1 = CFPrecision(FMul(h2, t_gain))   ' IK 20210422 fix: added * t_gain; otherwise TF is correct only when filter gain == 1
    fp.b2 = fp.b0
     ' Break '

  Case BPF: '"roof shape BPF" +20dB/dec-peak-20dB/dec' par1=peak freq, par3=Q1=peak mag
    testVar = FAdd(FAdd(1#, FDiv(cotan_of_half_omega_1, q1)), FMul(cotan_of_half_omega_1, cotan_of_half_omega_1))
    Call ensure_float_number_is_not_zero(testVar, tiny_positive_f32_number)
    h1 = FDiv(1#, testVar)
    h2 = FDiv(FMul(FMul(t_gain, cotan_of_half_omega_1), h1), q1)

    fp.a1 = CFPrecision(FMul(FMul(2#, FSub(1#, FMul(cotan_of_half_omega_1, cotan_of_half_omega_1))), h1))
    fp.a2 = CFPrecision(FMul(FAdd(FSub(1#, FDiv(cotan_of_half_omega_1, q1)), FMul(cotan_of_half_omega_1, cotan_of_half_omega_1)), h1))
    fp.b0 = CFPrecision(h2)
    fp.b1 = 0#
    fp.b2 = CFPrecision(FSub(0#, h2))
     ' Break '

  Case PID: '"leaky" integral
    h1 = FDiv(FMul(fp.par(PID_I_gain), 1#), LOOP_FREQUENCY)          ' Integral
    h2 = FMul(FMul(fp.par(PID_D_gain), LOOP_FREQUENCY), 0.001)       ' differential, /1000 because D gain is too big, should be in the range 0.001 then mag bend up happens at 100 Hz
    'par(3) is filter Q, here used as frequency of leak
    fp.a1 = FSub(FMul(PI_x2_OVER_FREQUENCY, q1), 1#)                 'slightly more, then -1.0f - "leaky" integral - ALLOWS GRADUALLY PURGE delay buffers
' IK in fixed point math (inherited from PEPS-2 Arm 7TDMA processor ADuC7026/7126)
' experimentally found 0.9999 cannot efficiently remove offset at low gain if OP is way too low,
' but too much 0.99998 can overflow filter internal delay buffers and flip over at start, especially if input pressure only 10 psi above max iso pressure.
'IT DEPENDS ON SAMPLE FREQUENCY -same a1, higher freq gives faster purge!!!
    fp.b0 = CFPrecision(FAdd(FAdd(fp.par(PID_P_gain), h1), h2))
    fp.b1 = CFPrecision(FSub(0#, FAdd(fp.par(PID_P_gain), FMul(2#, h2))))
    fp.b2 = CFPrecision(h2)
     ' Break '

'  Case GEO_OPT: 'gain = 1 when f > F2, INTENDED USE: F1=0.2, F2=0.5' it is LEDLAG --2nd order with Q1=Q2~0.1
'      '**** CAUTION WITH FIX POINT !!, VERY HIGH GAIN AT LOW FREQUENCY!!!
'    h1 = Tan(omega_2zero / 2) 'use as temporary to calculate h2
'    h2 = (1 - h1) / (1 + h1) '         ' = (1 - tan(PI*f2/fs) / (1 + tan(PI*f2/fs) approx = exp(-f2* PI_x2_OVER_FREQUENCY), better than 1% when fs/f0 > 63
'' Note: (1 - tan(x)) / (1 + tan(x)) = ((cotan(x) - 1) /  ((cotan(x) + 1)
'    h1 = (cotan_of_half_omega_1 - 1) / (cotan_of_half_omega_1 + 1) ' = (1 - tan(PI*f1/fs) / (1 + tan(PI*f1/fs)~= exp(-f1* PI_x2_OVER_FREQUENCY)
'
'    fp.a1 = CFPrecision(-2 * h2) '
'    fp.a2 = CFPrecision(h2 * h2) '
'    fp.b0 = CFPrecision(t_gain) '
'    fp.b1 = CFPrecision(-2 * h1 * t_gain) '
'    fp.b2 = CFPrecision(h1 * h1 * t_gain) '
'     ' Break '

  Case VELOCITY:
    h1 = FTan(FDiv(omega_2zero, 2#))  'use as temporary to calculate h2
    testVar = FAdd(1#, h1)
    Call ensure_float_number_is_not_zero(testVar, tiny_positive_f32_number)
    h2 = FDiv(FSub(1#, h1), testVar)
' Note: (1 - tan(x)) / (1 + tan(x)) = ((cotan(x) - 1) /  ((cotan(x) + 1)
    testVar = FAdd(cotan_of_half_omega_1, 1#)
    Call ensure_float_number_is_not_zero(testVar, tiny_positive_f32_number)
    h1 = FDiv(FSub(cotan_of_half_omega_1, 1#), testVar)

    fp.a1 = CFPrecision(FSub(0#, FAdd(h1, h2)))
    fp.a2 = CFPrecision(FMul(h1, h2))

    If (omega_1pole > omega_2zero) Then
        Call ensure_float_number_is_not_zero(f2, tiny_positive_f32_number)
        t_gain = FMul(FMul(FDiv(t_gain, f2), LOOP_FREQUENCY), 0.08) ' when F2 > F1
    Else
        Call ensure_float_number_is_not_zero(f1, tiny_positive_f32_number)
        t_gain = FMul(FMul(FDiv(t_gain, f1), LOOP_FREQUENCY), 0.08) ' when F2 < F1
    End If
    fp.b0 = CFPrecision(FMul(FAdd(FAdd(1#, fp.a1), fp.a2), t_gain))  'gain correction '-!- PHASE CORRECTED 09-25-09
    fp.b2 = FSub(0#, fp.b0)
     ' Break '

  Case Position:
    h1 = FTan(FDiv(omega_2zero, 2#)) 'use as temporary to calculate h2
    testVar = FAdd(1#, h1)
    Call ensure_float_number_is_not_zero(testVar, tiny_positive_f32_number)
    h2 = FDiv(FSub(1#, h1), testVar) ' = (1 - tan(PI*f2/fs) / (1 + tan(PI*f2/fs) approx = exp(-f2* PI_x2_OVER_FREQUENCY), better than 1% when fs/f0 > 63
' Note: (1 - tan(x)) / (1 + tan(x)) = ((cotan(x) - 1) /  ((cotan(x) + 1)
    testVar = FAdd(cotan_of_half_omega_1, 1#)
    Call ensure_float_number_is_not_zero(testVar, tiny_positive_f32_number)
    h1 = FDiv(FSub(cotan_of_half_omega_1, 1#), testVar)   '= (1 - tan(PI*f1/fs) / (1 + tan(PI*f1/fs)~= exp(-f1* PI_x2_OVER_FREQUENCY)
    fp.a1 = CFPrecision(FSub(0#, FAdd(h1, h2)))
    fp.a2 = CFPrecision(FMul(h1, h2))
    fp.b0 = CFPrecision(FMul(FMul(0.5, FAdd(FAdd(1#, fp.a1), fp.a2)), t_gain)) '        'fp.b0= -(1.0 + fp.a2 + fp.a1)/2''-!- PHASE CORRECTED 09-25-09
    fp.b1 = fp.b0
     ' Break '

  Case GAIN_ONLY:
    ' already set above
  Case Else:
    ' already set above
 End Select

''Reverse a1, a2 to use only ADD (summing) and do not use SUB (no subtraction)
    fp.a1 = FSub(0#, fp.a1)
    fp.a2 = FSub(0#, fp.a2)
End Sub

' 20250806 Simon Li
' if incoming floating number is close to zero, this function will change input floating number to
' a very small number to avoid divide by zero problem
' near_zero_positive_float_number is used to allow using different small positive floating number
Sub ensure_float_number_is_not_zero(ByRef var_ptr As Double, near_zero_positive_float_number As Double)
    If ((var_ptr >= 0#) And (var_ptr < near_zero_positive_float_number)) Then
        var_ptr = near_zero_positive_float_number
    Else
        If ((var_ptr < 0#) And (var_ptr > -near_zero_positive_float_number)) Then
            var_ptr = -near_zero_positive_float_number
        End If
    End If
End Sub

Function ExtendedHEX(int_in As Long) As String
' returns 0-9, A-F, G-Z. if error, returns '-'
    Dim ConvertedCHAR As String
    If int_in > (9 + 27) Then 'error, return '='
        ConvertedCHAR = "-"
        ExtendedHEX = ConvertedCHAR
        Exit Function
    End If

    If (int_in < 10) Then
        ConvertedCHAR = Hex(int_in)
    Else
        ConvertedCHAR = Chr(Asc("A") + int_in - 10)
    End If
    ExtendedHEX = ConvertedCHAR
End Function

Public Function IsNumber(ByVal Expression As String) As Boolean
    'ChatGPT & IK20251112
    ' Locale- and format-safe numeric validator for VB6
    ' Supports decimal, scientific, and hexadecimal (&Hxxxx) forms.
    ' Returns True if Expression can be converted safely to a number.

    Dim Negative As Boolean
    Dim Exponent As Boolean
    Dim NumberFound As Boolean
    Dim Period As Boolean
    Dim Positive As Boolean
    Dim Hexadecimal As Boolean
    Dim ch As String
    Dim i As Long

    On Error GoTo NotNumber
    IsNumber = False
    Expression = Trim$(Expression)
    If Len(Expression) = 0 Then Exit Function

    ' Check for hex form first
    If UCase$(Left$(Expression, 2)) = "&H" Then
        ' Must be valid hex digits after &H
        For i = 3 To Len(Expression)
            ch = Mid$(Expression, i, 1)
            Select Case ch
                Case "0" To "9", "A" To "F", "a" To "f"
                    ' ok
                Case Else
                    Exit Function
            End Select
        Next i
        IsNumber = True
        Exit Function
    End If

    ' Regular decimal or exponential form
    For i = 1 To Len(Expression)
        ch = Mid$(Expression, i, 1)
        Select Case ch
            Case "0" To "9"
                NumberFound = True

            Case "."
                If Period Then Exit Function
                Period = True

            Case "E", "e"
                If Exponent Or Not NumberFound Then Exit Function
                Exponent = True
                NumberFound = False
                Period = False
                Negative = False
                Positive = False

            Case "+"
                If i > 1 And Not Exponent Then Exit Function
                If Positive Then Exit Function
                Positive = True

            Case "-"
                If i > 1 And Not Exponent Then Exit Function
                If Negative Then Exit Function
                Negative = True

            Case " ", vbTab
                ' spaces only allowed before or after the number
                If NumberFound Then Exit Function

            Case "," ' if it is in the last position after some numbers, Val will convert it
                If NumberFound Then
                    IsNumber = NumberFound
                    Exit Function ' catch separator inside string
                End If
            Case Else
                Exit Function
        End Select
    Next i

    ' Valid only if number portion exists and exponent not left incomplete
    If Not NumberFound Then Exit Function

    IsNumber = True
    Exit Function

NotNumber:
    IsNumber = False
End Function


'Function set_get_fpar(message As String) As String
''fpar&$#? returns "fpar&$#=GGG.gg" scientific format displayed; fpar&$#=ggg.gg sets filter parameter
''& = axis [0-F]== first index of SysData.filters[6][max_filt_in_channel]; SysData.PROXfilt[6][]; SysData.FF_Filt[4][],
''    BECAUSE THESE FILTERS ARE LOCATED in SysData ONE AFTER ANOTHER,
''&=0 to 5 will access SysData.filters[6];
''&=6 to 11(B-hex) will access SysData.PROXfilt[6];
''&=12(C-hex) to 15(F-hex) will access SysData.FF_Filt[4]
''&=16(G-hex) to 19(J-hex) will access SysData.PressureFilter[4][3] for valve
''&=20(K-hex) or 21(L-hex) will access SysData.AuxFilter[1][12] as AuxFilter[2][6]
''
''$= filter number in the axis [0,1,2,3,4,5] = = second index in SysData.filters[NUM_FILT_CHANs][max_filt_in_channel];
''
''#= filter parameter number [0,1,2,3,4,5] == long ftyp; float par[5];
'Dim axis As Integer
'Dim filter_num As Integer
'Dim param As Integer
'Dim c   As Byte
'    c = Mid(message, CMD_LEN + 1, 1) '    c = temp_Inp_str[CMD_LEN];
'    c = LCase(c)
'    If ((c >= "g") And (c <= "l")) Then
'        axis = Asc(c - "a" + 10) ';//pressure filters
'    Else
'        axis = Asc(c) - Asc("a")     ';//decode @ =axis; ASCIIToHexChar returns = -1 and sets ErrorStatus =BAD_VALUE_ERR if conversion unsuccessful
'    End If
'
'    c = Mid(message, CMD_LEN + 2, 1) ' *filter_num = ASCIIToHexChar(temp_Inp_str[CMD_LEN + 1]);
'    filter_num = Asc(c)
'
'    c = Mid(message, CMD_LEN + 3, 1) ' *param = ASCIIToHexChar(temp_Inp_str[CMD_LEN + 2]);
'    param = Asc(c)
'
'    set_get_fpar = 1.1
'End Function

'Function set_get_gain(message As String) As String
''// "gain<index1><index2>=ggg.gg" - set/get will define gain for a particular loop, gain 3 takes to 2 arguments
''"gain#$=GGG.GGG" Set/get Feed Back Loop Gain.
''Parameter "#": 0...3, gain for the: 0=Axes, 1=Test_gains, 2=Test_phases, 3=Pressure.
''Parameter "$": index of axis : position Xp=0,Yp=1,tZp=2,Zp=3,tXp=4,tYp=5; velocity Xv=6,Yv=7,tZv=8,Zv=9,tXv=A,tYv=B; damping Zd=C,tXd=D,tYd=E
''Value "ggg.ggg"= feedback gain can be regular or scientific notation "+/-m.mmE+/-pp"; range: +/-100.0
''Get gain "gain$#?": returns ->"gain@#=ggg.gg
'Dim axis As Integer
'Dim index1 As Integer
'Dim param As Integer
'Dim parStr As String
'Dim c   As Byte
''//decode '$' parameter - axis type
'    parStr = Mid(message, CMD_LEN + 1, 1) '    c = temp_Inp_str[CMD_LEN];
'    parStr = LCase(parStr)
'    c = Asc(parStr) - Asc("0")
'    If ((c >= 0) And (c <= 3)) Then
'        axis = c  '
'    Else
'        set_get_gain = ">~ERR BAD param1, valid 0=Axes, 1=Test_gains, 2=Test_phases, 3=Pressure" & vbCrLf & OK_message
'        Exit Function
'    End If
''//decode '#' parameter - particular gain
'    parStr = LCase(Mid(message, CMD_LEN + 2, 1)) ' *filter_num = ASCIIToHexChar(temp_Inp_str[CMD_LEN + 1]);
'    c = Asc(parStr) - Asc("0")
'    If c > 10 Then c = c + Asc("0") - Asc("a") + 10
'    If (c >= 0) And (c < 16) Then
'        index1 = (c)
'    Else
'        set_get_gain = ">~ERR BAD param2, valid Xp=0,Yp=1,tZp=2,Zp=3,tXp=4,tYp=5,Xv=6,Yv=7,tZv=8,Zv=9,tXv=A,tYv=B,Zd=C,tXd=D,tYd=E, bd=F" & vbCrLf & OK_message
'        Exit Function
'    End If
'    If (axis = 0) Then '//position - velocity - Aux_damping
'    'SetGet_param(2+SHOW_float, -500.0f, +500.0f, (void*)&SysData.gain[index1], (char*)Gain_Name[index1]);
'        If (Mid(message, CMD_LEN + 3, 1) = "=") Then ' set command
'            FBgain(index1) = Val(Mid(message, CMD_LEN + 4))
'            set_get_gain = vbCrLf & OK_message
'        Else 'get command
'            set_get_gain = Left(message, CMD_LEN + 2) & "=" & FBgain(index1) & vbCrLf & OK_message
'        End If
'    Else
'        set_get_gain = ">~ERR " & vbCrLf & OK_message
'    End If
'End Function

Function IsDirectory(ByVal sFile As String) As Boolean
    On Error GoTo dir_search_error 'Resume Next
    IsDirectory = False
    IsDirectory = ((GetAttr(sFile) And vbDirectory) = vbDirectory)
    Exit Function
dir_search_error:
    IsDirectory = False
End Function

Function IsWindows64() As Boolean
    Dim Dir_str As String
    If FileExists("C:\WINDOWS\SysWOW64\format.com") Then
    'Dir_str = Dir("C:\Program Files (x86)", vbDirectory)
    'If Dir_str <> "" Then
        IsWindows64 = True
    Else
        IsWindows64 = False
    End If
End Function

 ' DO NOT USE frmMonitoring.Name - this will load the form to get its name. instead, use string im quotas "frmMonitoring"
 Public Function IsFormLoaded(FormName As String) As Boolean
    Dim frm As Form
    For Each frm In Forms
        If frm.Name = FormName Then 'The form is loaded
            IsFormLoaded = True
            Exit Function
        End If
    Next frm
    IsFormLoaded = False
End Function

Public Function FORMAT_G(ByVal Value As Single, Optional ByVal UseLocale As Boolean = True) As String
    '' ChatGPT & IK20251112
    ' Emulate C printf("%.9g") for 32-bit float
    ' Locality-aware: decimal separator follows system settings when UseLocale=True

    Const SIGDIGITS As Integer = 9
    Dim AbsVal As Double
    Dim Exp As Integer
    Dim SciStr As String
    Dim FixedStr As String
    Dim Fmt As String
    Dim dScale As Double
    Dim DecimalSep As String

    On Error Resume Next
    If IsNull(Value) Or IsEmpty(Value) Then FORMAT_G = "": Exit Function
    If Value = 0! Then FORMAT_G = "0": Exit Function
    On Error GoTo 0

    ' Detect system decimal separator
    DecimalSep = Mid$(Format$(0.1, "0.0"), 2, 1)

    AbsVal = Abs(CDbl(Value))
    Exp = Int(Log(AbsVal) / Log(10#))

    ' Round to 9 significant digits
    dScale = 10# ^ (SIGDIGITS - Exp - 1)
    Value = CSng(Round(Value * dScale) / dScale)

    ' Build formats
    Fmt = "0." & String$(SIGDIGITS - 1, "#")
    FixedStr = Format$(Value, Fmt)
    SciStr = Format$(Value, Fmt & "E+00")

    ' Choose notation
    If Exp < -4 Or Exp >= SIGDIGITS Then
        FORMAT_G = SciStr
    Else
        FORMAT_G = FixedStr
    End If

    ' Trim trailing zeros / point
    Do While Right$(FORMAT_G, 1) = "0"
        FORMAT_G = Left$(FORMAT_G, Len(FORMAT_G) - 1)
    Loop
    If Right$(FORMAT_G, 1) = "." Then
        FORMAT_G = Left$(FORMAT_G, Len(FORMAT_G) - 1)
    End If

    ' Replace decimal point if locale uses comma
    If UseLocale And DecimalSep <> "." Then
        FORMAT_G = Replace(FORMAT_G, ".", DecimalSep)
    End If
End Function

'public function Modulo(a/b) = a - (b * Fix(a / b))
Public Function Modulo(a As Double, b As Double) As Double
Modulo = a - (b * Fix(a / b))
End Function


'Note: BitsToSingle() is not a numeric conversion - it directly reinterprets the bit pattern (identical to a C-style union or type pun).
'For example:
'BitsToSingle(&H3F800000) = 1#
'BitsToSingle(&HC0000000) = -2#
Public Function BitsToSingle(ByVal l As Long) As Single
    ' Reinterpret a 32-bit integer value as IEEE-754 Single
    CopyMemory BitsToSingle, l, 4&
End Function

Public Function SingleToHex(ByVal v As Single) As String
    Dim l As Long
    CopyMemory l, v, 4&
    SingleToHex = "&H" & Right$("00000000" & Hex$(l), 8)
End Function

' --- Main normalization function ---
' ChatGPT & IK20251112
' This function developer-friendly and DSP-debugger-ready - while remaining totally safe for normal users.
' They'll never know the hex path exists unless they paste one in.
Public Function NormalizeFloatStringEx(ByVal s As String, ByRef IsConverted As Boolean) As String
    Dim DecimalSep As String
    Dim ThousandSep As String
    Dim Result As String
    Dim vSingle As Single

    IsConverted = False
    Result = Trim$(s)

    ' Get system decimal and thousand separators
    DecimalSep = Mid$(Format$(0.1, "0.0"), 2, 1)
    ThousandSep = Mid$(Format$(1000, "#,###"), 2, 1)

    ' --- 1. Hexadecimal 32-bit float handling ---
    If UCase$(Left$(Result, 2)) = "&H" Then
        Dim hexVal As Long
        On Error Resume Next
        hexVal = CLng(Result)  ' convert hex literal to Long
        If Err.Number = 0 Then
            vSingle = BitsToSingle(hexVal)
            NormalizeFloatStringEx = FORMAT_G(vSingle, False)  ' decimal string
            IsConverted = True
        Else
            NormalizeFloatStringEx = ""
        End If
        Err.Clear
        Exit Function
    End If

    ' --- 2. Remove thousand separators ---
    If ThousandSep <> "" Then Result = Replace(Result, ThousandSep, "")

    ' --- 3. Replace locale decimal separator with dot ---
    If DecimalSep <> "." Then Result = Replace(Result, DecimalSep, ".")

    ' --- 4. Validate numeric input ---
    If IsNumber(Result) Then
        IsConverted = True
        NormalizeFloatStringEx = Result
    Else
        IsConverted = False
        NormalizeFloatStringEx = ""
    End If
End Function

