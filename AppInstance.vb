Option Strict Off
'Option Explicit Off

Imports System.Windows.Forms
Imports System.Globalization
Imports System.Drawing
Imports System.Threading
Imports System.Reflection

Public Class AppInstanceClass
    Inherits Logger.ErrorClass

    Public Shared DefInstance As AppInstanceClass

    Public ResMan As Resources.ResourceManager

    Public gbPleaseLogOn As Boolean = True
    Public SystemLocale As CultureInfo
    Public MQS As MQSeries.MQSeries
    Public MQ_Server As String
    Public MQ_ReadQueue As String
    Public MQ_WriteQueue As String
    Public Unit As String
    Public UnitName As String
    Public Logger As Logger.LoggerClass
    Public NumEng As CNEP0100.NumEngClass
    Public IsLoggedOn As Boolean
    Public isShuttingDown As Boolean = False
    Public LatestRevision As String
    Public gfrmMain As Object 'Windows.Forms.Form
    Public gfrmPrint As Object 'Windows.Forms.Form
    Public gLastErrorStr As String = ""
    Public Statistics As Statistics_struct
    Public gWaitTaT As Integer = 100
    Public LastMsgID As String
    Public MQSyncObject As New Object
    Public IceDBParams As IceDBParams_Struct
    Public nShowToolBar As Integer = -1

    Public MQErrors As New Collection

    Public PrintThreadReference As Integer = 0
    'Public gFoxRunning As Integer = 0
    '''Public gIcePrinting As Integer = 0
    '''Public gIcePrintThread As Threading.Thread
    Public gAppBusy As Boolean = False
    Public gLoggingOff As Boolean = False

    Private m_LoggedUser As String = ""
    Private m_AbortGetMessage As Integer = 0

    Private tmrStatusText As System.Windows.Forms.Timer
    Private LastErrQue As Queue


    Private IcoRed As Icon
    Private IcoGry As Icon
    Private IcoGrn As Icon



    Public PrintSync As New Object
    Public PrintThreads As New Collection

    Private m_CstDB As SIBL0100.SQL.SQLOleDB

    Public ReadOnly Property CstDB() As SIBL0100.SQL.SQLOleDB
        Get
            If m_CstDB Is Nothing Then
                m_CstDB = New SIBL0100.SQL.SQLOleDB(getCstDBConnection)
            End If
            Return m_CstDB
        End Get
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''     Gets the db connection.
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Y093sahu]	dd/mm/yyyy	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Function getCstDBConnection() As OleDb.OleDbConnection
        Dim sCon As String = String.Empty

        sCon = "ice"
        sCon += "ice"
        Dim DBname As String = ICEI0100.IcePaths.DefInstance.DatabasePath & "\CstDB_" & Me.Unit & ".mdb"
                sCon = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & DBname & ";Jet OLEDB:Database Password=" & sCon
                sCon &= "baby"
        Return New OleDb.OleDbConnection(sCon)
    End Function

    Public Structure PrintThreadObject_stuct
        Dim ObjKey As String
        Dim CreateStamp As Date
        Dim PrintObject As Object
        Dim PrintThread As Threading.Thread
    End Structure

    Public Structure MQError_Item_struct
        Dim MsgID As String
        Dim ErrStr As String
    End Structure

    Private m_AllowExit As Boolean = True
    Public Property AllowExit() As Boolean
        Get
            Return m_AllowExit Or isShuttingDown
        End Get
        Set(ByVal Value As Boolean)
            m_AllowExit = Value
        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Function to syncrhonize the printer queue in ICE
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[u910krzi]	04-Apr-2009	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Sub EnterPrintBlock()
        Interlocked.Increment(PrintThreadReference)
        Dim Res As Boolean = Monitor.TryEnter(PrintSync)
        If Not Res Then
            ModalMsgBox(clsPrintForm.PrintingQueued, MsgBoxStyle.Information, "Print Queue")
            Monitor.Enter(PrintSync)
        End If
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Function to syncrhonize the printer queue in ICE
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[u910krzi]	04-Apr-2009	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Sub ExitPrintBlock()
        Interlocked.Decrement(PrintThreadReference)
        Monitor.Exit(PrintSync)
    End Sub

#Region "WinAPI Declarations"

    Public Const DefArbSa As Long = &H401
    Public Const DefEngUs As Long = &H409

    Public Structure SYSTEMTIME
        Dim wYear As Short 'Integer
        Dim wMonth As Short 'Integer
        Dim wDayOfWeek As Short 'Integer
        Dim wDay As Short 'Integer
        Dim wHour As Short 'Integer
        Dim wMinute As Short 'Integer
        Dim wSecond As Short 'Integer
        Dim wMilliseconds As Short 'Integer
    End Structure

    Public Declare Function MoveFileEx Lib "kernel32" Alias "MoveFileExA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As Integer, ByVal dwFlags As Integer) As Integer
    'Public Declare Function GetDateFormat Lib "kernel32" Alias "GetDateFormatA" (ByVal Locale As Long, ByVal dwFlags As Long, ByVal lpDate As SYSTEMTIME, ByVal lpFormat As String, ByVal lpDateStr As String, ByVal cchDate As Long) As Long
    Public Declare Function GetDateFormat Lib "kernel32" Alias "GetDateFormatA" (ByVal Locale As Integer, ByVal dwFlags As Integer, ByRef lpDate As SYSTEMTIME, ByVal lpFormat As String, ByVal lpDateStr As String, ByVal cchDate As Integer) As Integer
    Public Declare Function SetEnvironmentVariable Lib "kernel32" Alias "SetEnvironmentVariableA" (ByVal lpName As String, ByVal lpName As String) As Integer
    Public Const MOVEFILE_DELAY_UNTIL_REBOOT As Short = &H4S


    Public Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As IntPtr) As Integer
    Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As IntPtr, ByVal nCmdShow As Integer) As Integer
    Public Declare Function ShowWindowAsync Lib "user32" (ByVal hwnd As IntPtr, ByVal nCmdShow As Integer) As Integer
    Public Declare Function IsIconic Lib "user32" (ByVal hwnd As IntPtr) As Integer
    Public Const SW_HIDE As Short = 0
    Public Const SW_NORMAL As Short = 1
    Public Const SW_SHOWMINIMIZED As Short = 2
    Public Const SW_SHOWMAXIMIZED As Short = 3
    Public Const SW_SHOWNOACTIVATE As Short = 4
    Public Const SW_RESTORE As Short = 9
    Public Const SW_SHOWDEFAULT As Short = 10

    Public Declare Function SetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal hwnd As IntPtr, ByVal lpString As String) As Integer

    Public Declare Function SendMessageTimeout Lib "user32" Alias "SendMessageTimeoutA" (ByVal hwnd As Integer, ByVal msg As Integer, ByVal wParam As Short, ByVal lParam As String, ByVal fuFlags As Integer, ByVal uTimeout As Integer, ByRef lpdwResult As Integer) As Integer
    Public Const SMTO_NORMAL As Integer = &H0
    Public Const SMTO_BLOCK As Integer = &H1
    Public Const SMTO_ABORTIFHUNG As Integer = &H2
    Public Const HWND_BROADCAST As Integer = &HFFFF&
    Public Const WM_SETTINGCHANGED As Integer = &H1A
    Public Const WM_FONTCHANGE As Integer = &H1D

    Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As IntPtr, ByVal lMsg As Integer, ByVal wParam As Short, ByRef lParam As LVHITTESTINFO) As Integer
    Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Integer, ByVal lMsg As Integer, ByVal wParam As Short, ByVal lParam As Integer) As Integer
    '
    Public Const LVM_FIRST As Integer = &H1000
    Public Const LVM_SUBITEMHITTEST As Decimal = (LVM_FIRST + 57)
    '
    Public Const LVHT_NOWHERE As Short = &H1S
    Public Const LVHT_ONITEMICON As Short = &H2S
    Public Const LVHT_ONITEMLABEL As Short = &H4S
    Public Const LVHT_ONITEMSTATEICON As Short = &H8S
    Public Const LVHT_ONITEM As Boolean = (LVHT_ONITEMICON Or LVHT_ONITEMLABEL Or LVHT_ONITEMSTATEICON)
    '
    Public Structure POINTAPI
        Dim x As Integer
        Dim y As Integer
    End Structure
    '
    Public Structure LVHITTESTINFO
        Dim pt As POINTAPI
        Dim lFlags As Integer
        Dim lItem As Integer
        Dim lSubItem As Integer
    End Structure
    ''''' start sofyan digital substiuation 13/06/2011
    Public Const DS_CONTEXT As Integer = 0
    Public Const DS_NONE As Integer = 1
    Public Const DS_NATIONAL As Integer = 2

    Public Const wdNumeralArabic As Integer = 0
    Public Const wdNumeralHindi As Integer = 1
    Public Const wdNumeralContext As Integer = 2
    Public Const wdNumeralSystem As Integer = 3

    Public Const LOCALE_SNATIVEDIGITS As Integer = &H13
    Public Const LOCALE_IDIGITSUBSTITUTION As Integer = &H1014

    Public Const WM_SETTINGCHANGE As Integer = &H1A

    'API declaration
    Public Declare Function GetLocaleInfo Lib "kernel32" Alias "GetLocaleInfoA" (ByVal Locale As Integer, ByVal LCType As Integer, ByRef lpLCData As Short, ByVal iLen As Integer) As Integer
    Public Declare Function SetLocaleInfo Lib "kernel32" Alias "SetLocaleInfoA" (ByVal Locale As Integer, ByVal LCType As Integer, ByVal lpLCData As String) As Integer
    Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Integer, ByVal wMsg As Integer, ByVal wParam As Integer, ByVal lParam As Integer) As Integer
    Public Declare Function GetSystemDefaultLCID Lib "kernel32" () As Integer
    Public Declare Function GetUserDefaultLCID Lib "kernel32" () As Integer

    Public Function GetDigitSubstitution() As Integer
        Dim dwLCID As Integer
        Dim dwRCDE As Integer
        Dim dwDSVL As Short 'Integer

        GetDigitSubstitution = -1
        dwDSVL = -1 'Len(dwDSVL)

        dwLCID = GetUserDefaultLCID() 'GetSystemDefaultLCID()
        dwRCDE = GetLocaleInfo(dwLCID, LOCALE_IDIGITSUBSTITUTION, dwDSVL, Len(dwDSVL)) 'Len(dwDSVL))
        If dwRCDE = 0 Then
            Return 0
            Exit Function
        End If
        Return dwDSVL

    End Function

    Public Function SetDigitSubstitution(ByVal iSbsCode As Integer) As Boolean
        Dim dwLCID As Integer
        Dim dwRCDE As Integer

        dwLCID = GetUserDefaultLCID() 'GetSystemDefaultLCID()
        dwRCDE = SetLocaleInfo(dwLCID, LOCALE_IDIGITSUBSTITUTION, iSbsCode)
        If dwRCDE = 0 Then
            Return False
            Exit Function
        End If
        PostMessage(HWND_BROADCAST, WM_SETTINGCHANGE, 0, 0)
        Return True

    End Function

    ''''' End sofyan 13/06/2011
    ''''' End sofyan 13/06/2011
    ''''' Start Sofyan 25/06/2011
    ''''' API Declaration for font installation

    Public Declare Function AddFontResource Lib "gdi32" Alias "AddFontResourceA" (ByVal lpFileName As String) As Long

    Public Function InstallFont(ByVal p_FontFullPath As String) As Boolean

        Dim RetCde As Long

        If p_FontFullPath.Trim = "" Then Return False
        Logger.LogInfo(0, "Before calling Windows API [AddFontResource].", 3)
        If (AddFontResource(p_FontFullPath) > 0) Then
            Logger.LogInfo(0, "After calling Windows API [AddFontResource].", 3)
            Logger.LogInfo(0, "Before calling Windows API [SendMessage].", 3)
            RetCde = SendMessage(HWND_BROADCAST, WM_FONTCHANGE, 0, 0)
            Logger.LogInfo(0, "After calling Windows API [SendMessage].", 3)
            Return True
        Else

            Return False

        End If

    End Function

    ''''' End
#End Region

    Public Enum enumActivityType As Integer
        None = 0
        Write = 1
        Wait = 2
        Read = 3
    End Enum

    Public Structure IceDBParams_Struct
        Dim CheckDates As Boolean
        Dim NumOfDays As Integer
    End Structure

    Public Structure LastErr_Struct
        Dim ErrTme As Date
        Dim ErrDsc As String
    End Structure

    Public Structure Statistics_struct
        Dim MessagesSent As Integer
        Dim MessagesRecv As Integer
        Dim AverageTat As Double
        Dim AverageCnt As Integer
    End Structure


    Public Property LoggedUser() As String
        Get
            LoggedUser = m_LoggedUser
        End Get
        Set(ByVal Value As String)
            m_LoggedUser = Value
            Try
                If Not (MQS Is Nothing) Then MQS.LoggedUser = Value
            Catch
                'do nothing
            End Try
        End Set
    End Property

    Public ReadOnly Property Workstation() As String
        Get
            Return Environment.GetEnvironmentVariable("COMPUTERNAME").ToUpper
        End Get
    End Property

    Public Sub AbortMessage()
        Thread.VolatileWrite(m_AbortGetMessage, 1)
    End Sub

    Public Sub New()
        LoggedUser = (Environment.UserName).ToUpper
        LastErrQue = New Queue
        tmrStatusText = New System.Windows.Forms.Timer
        AddHandler tmrStatusText.Tick, AddressOf TimerEventProcessor
        tmrStatusText.Interval = 5000
        DefInstance = Me
        ResMan = New Resources.ResourceManager("ICEI0100.ErrorCodes", Me.GetType.Assembly)
    End Sub

    Public Sub Initialize(ByVal StartUpPath As String)
        Logger = New Logger.LoggerClass("ICEP0100", ICEI0100.IcePaths.DefInstance.LogFilePath, LoggedUser, "Ice Log")
        MQS = New MQSeries.MQSeries
        NumEng = New CNEP0100.NumEngClass(StartUpPath & "\CNEF0100.ICE")
        If NumEng.ErrNum <> 0 Then
            m_ErrNum = NumEng.ErrNum
            m_errdsc = "Error Initializing Numbering Engine: " & NumEng.ErrDsc
        End If
    End Sub

    Protected Overrides Sub Finalize()
        Try
            If Not (Logger Is Nothing) Then
                Logger = Nothing
                GC.Collect()
            End If
            tmrStatusText.Stop()
            tmrStatusText.Dispose()
        Catch
        End Try
        MyBase.Finalize()
    End Sub

    <Runtime.InteropServices.DllImport("user32.dll")> _
    Private Shared Function GetForegroundWindow() As IntPtr
    End Function

    Public Function GetActiveWindow() As Windows.Forms.Form
        Dim Handle As IntPtr = GetForegroundWindow()
        Dim Ctl As Control = Control.FromHandle(Handle)
        Return Ctl
    End Function

    Public Function GetErrStr(ByVal ErrStr As String) As String
        Dim st As String = ResMan.GetString(ErrStr)
        Return CStr(IIf(st = "", ErrStr, st))
    End Function

    Public Function ModalMsgBox(ByVal Prompt As Object, Optional ByVal Style As MsgBoxStyle = MsgBoxStyle.OKOnly, Optional ByVal Title As Object = Nothing) As MsgBoxResult
        'Dim Btns As MessageBoxButtons = Style And &HF
        'Dim DefBtn As MessageBoxDefaultButton = Style And &HF00
        'Dim Icn As MessageBoxIcon = Style And &HF0
        Return CType(ModalMessageBox(Prompt, Style And &HF, Style And &HF00, Style And &HF0, Title), MsgBoxResult)
    End Function

    Public Function ModalMessageBox(ByVal Prompt As Object, Optional ByVal Btns As MessageBoxButtons = MessageBoxButtons.OK, _
                                    Optional ByVal DefBtn As MessageBoxDefaultButton = MessageBoxDefaultButton.Button1, _
                                    Optional ByVal Icn As MessageBoxIcon = MessageBoxIcon.None, Optional ByVal Title As String = "") As DialogResult
        Dim ctl As Control = GetActiveWindow()
        If ctl Is Nothing Then ctl = gfrmMain

        Try
            If Not ctl Is Nothing Then
                Return MessageBox.Show(ctl, Prompt, Title, Btns, Icn, DefBtn)
            Else
                Return MessageBox.Show(Prompt, Title, Btns, Icn, DefBtn)
            End If
        Catch ex As Exception 'There might be an exception when trying to raise a msgbox with its parent being disposed, so this will raise it without a parent.
            Return MessageBox.Show(Prompt, Title, Btns, Icn, DefBtn)
        End Try

    End Function

    Public Sub IceLogOn()
        gfrmMain.menuFile_Login_Click(New System.Object, System.EventArgs.Empty)
    End Sub

    Public Sub IceLogOff()
        gfrmMain.menuFile_Logout_Click(New System.Object, System.EventArgs.Empty)
    End Sub

    'Function returns true when the supplied account number or customer number represents an internal account
    Public Function IsInternalAccount(ByVal AccNum As String) As Boolean
        If AccNum Is Nothing Then Return False
        AccNum = AccNum.Trim
        If (AccNum.Length <> 6) And (AccNum.Length <> 13) Then Return False
        Dim st As String = AccNum
        If AccNum.Length = 13 Then st = AccNum.Substring(4, 6)
        'Return (st > "799999")
        Return (st.StartsWith("8") Or st.StartsWith("9"))
    End Function

    Private Sub SetPanelIcon(ByVal DestPanel As StatusBarPanel, ByVal Ico As Icon)
        'DestPanel.Icon = Ico

        Dim tt As New SetFromThread(gfrmMain, DestPanel, Ico)
        tt.SetIcon()

    End Sub


    Public Sub SetControlText_Thread(ByRef pForm As Form, ByVal ctl As Control, Text As String)
        Dim tt As New SetFromThread(gfrmMain, ctl, Text)
        tt.SetText()
    End Sub

    Public Sub SetPanelText_Thread(ByVal Pnl As StatusBarPanel, Text As String)
        Dim tt As New SetFromThread(gfrmMain, Pnl, Text)
        tt.SetPanel()
    End Sub

    Public Sub SetPanelIcon(ByVal DestPanel As StatusBarPanel, ByVal ImgList As ImageList, ByVal Index As Int32)
        Try
            ''gfrmMain.MainStatusBar.Visible = False
            Dim bmp As Drawing.Bitmap = New Drawing.Bitmap(ImgList.Images(Index))
            Dim ico As Drawing.Icon = Drawing.Icon.FromHandle(bmp.GetHicon())
            DestPanel.Icon = ico
            bmp.Dispose()
            bmp = Nothing
        Catch
            DestPanel.Icon = Nothing
        Finally
            ''gfrmMain.MainStatusBar.Visible = True
        End Try
    End Sub

    Public Sub SetPanelIcon(ByVal DestPanel As StatusBarPanel, ByVal IconFile As String)
        Try
            Dim ico As New Drawing.Icon(IconFile)
            DestPanel.Icon = ico
            ico.Dispose()
            ico = Nothing
        Catch
            DestPanel.Icon = Nothing
        End Try
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Reset the control supplied to its original state
    ''' </summary>
    ''' <param name="p_Control">The control which needs its contents reset</param>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[u910krzi]	04-Apr-2009	Created
    '''     ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Shared Sub ClearControlData(ByRef p_Control As Control)
        If TypeOf p_Control Is Label Then
            Return
        ElseIf TypeOf p_Control Is LinkLabel Then
            Return
        ElseIf TypeOf p_Control Is GroupBox Then
            Return
        ElseIf TypeOf p_Control Is Panel Then
            Return
        ElseIf TypeOf p_Control Is TabControl Then
            Return
        ElseIf TypeOf p_Control Is PictureBox Then
            Return
        ElseIf TypeOf p_Control Is CheckBox Then
            CType(p_Control, CheckBox).Checked = False
            Return
        ElseIf TypeOf p_Control Is RadioButton Then
            'CType(p_Control, RadioButton).AutoCheck = Not (isReadOnly)
            Return
        ElseIf TypeOf p_Control Is Button Then
            'p_Control.Enabled = Not (isReadOnly)
            Return
        ElseIf TypeOf p_Control Is TextBox Then
            CType(p_Control, TextBox).Text = ""
        ElseIf TypeOf p_Control Is ComboBox Then
            If CType(p_Control, ComboBox).Items.Count > 0 Then CType(p_Control, ComboBox).SelectedIndex = 0
        ElseIf TypeOf p_Control Is CheckedListBox Then
            CType(p_Control, CheckedListBox).Items.Clear()
        ElseIf TypeOf p_Control Is ListBox Then
            CType(p_Control, ListBox).Items.Clear()
        ElseIf TypeOf p_Control Is ListView Then
            CType(p_Control, ListView).Items.Clear()
        ElseIf TypeOf p_Control Is NumericUpDown Then
            'CType(p_Control, NumericUpDown).Text = "0"
        ElseIf TypeOf p_Control Is TrackBar Then
            CType(p_Control, TrackBar).Value = CType(p_Control, TrackBar).Minimum
        ElseIf TypeOf p_Control Is ProgressBar Then
            CType(p_Control, ProgressBar).Value = CType(p_Control, ProgressBar).Minimum
        End If
    End Sub

    Public Sub SetControlState(ByRef p_Control As Control, ByVal isReadOnly As Boolean)
        If TypeOf p_Control Is Label Then
            Return
        ElseIf TypeOf p_Control Is LinkLabel Then
            Return
        ElseIf TypeOf p_Control Is GroupBox Then
            Return
        ElseIf TypeOf p_Control Is Panel Then
            Return
        ElseIf TypeOf p_Control Is TabControl Then
            Return
        ElseIf TypeOf p_Control Is PictureBox Then
            Return
        ElseIf TypeOf p_Control Is CheckBox Then
            CType(p_Control, CheckBox).AutoCheck = Not (isReadOnly)
            Return
        ElseIf TypeOf p_Control Is RadioButton Then
            CType(p_Control, RadioButton).AutoCheck = Not (isReadOnly)
            Return
        ElseIf TypeOf p_Control Is Button Then
            p_Control.Enabled = Not (isReadOnly)
            Return
        ElseIf TypeOf p_Control Is TextBox Then
            CType(p_Control, TextBox).ReadOnly = isReadOnly
        ElseIf TypeOf p_Control Is ComboBox Then
            p_Control.Enabled = Not (isReadOnly)
        ElseIf TypeOf p_Control Is CheckedListBox Then
            p_Control.Enabled = Not (isReadOnly)
        ElseIf TypeOf p_Control Is TrackBar Then
            p_Control.Enabled = Not (isReadOnly)
        ElseIf TypeOf p_Control Is ProgressBar Then
            p_Control.Enabled = Not (isReadOnly)
        ElseIf TypeOf p_Control Is NumericUpDown Then
            CType(p_Control, NumericUpDown).Enabled = Not (isReadOnly)
        Else
            Return 'Prevent custom controls from changing their bakcground color
        End If

        If isReadOnly Then
            p_Control.BackColor = Color.FromArgb(255, 255, 192)
        Else
            p_Control.BackColor = SystemColors.Window
        End If
    End Sub


    Public Class SetFromThread
        Public Frm As Form
        Public Ctl As Control
        Public Pnl As StatusBarPanel
        Public Text As String
        Public mIcon As Icon

        Public Sub New(ByRef pForm As Form, ByRef pControl As Control, pText As String)
            Frm = pForm
            Ctl = pControl
            Text = pText
        End Sub

        Public Sub New(ByRef pForm As Form, ByRef pControl As StatusBarPanel, pText As String)
            Frm = pForm
            Pnl = pControl
            Text = pText
        End Sub

        Public Sub New(ByRef pForm As Form, ByRef pControl As StatusBarPanel, pIcon As Icon)
            Frm = pForm
            Pnl = pControl
            mIcon = pIcon
        End Sub

        Public Sub SetText()
            If Frm.InvokeRequired Then
                Try : Frm.Invoke(New MethodInvoker(AddressOf SetText)) : Catch : End Try
            Else
                Ctl.Text = Text
            End If
        End Sub

        Public Sub SetPanel()
            If Frm.InvokeRequired Then
                Try : Frm.Invoke(New MethodInvoker(AddressOf SetPanel)) : Catch : End Try
            Else
                Pnl.Text = Text
            End If
        End Sub

        Public Sub SetIcon()
            If Frm.InvokeRequired Then
                Try : Frm.Invoke(New MethodInvoker(AddressOf SetIcon)) : Catch : End Try
            Else
                Pnl.Icon = mIcon
            End If
        End Sub

    End Class
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Reset all the fields under the root control supplied to their original state
    ''' </summary>
    ''' <param name="p_Root">A root control (e.g. panel, form, or group box)</param>
    ''' <param name="ExceptionList">An array of controls that are to be excluded from the process</param>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[u910krzi]	04-Apr-2009	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Sub ClearAllControlsData(ByRef p_Root As Control, Optional ByRef ExceptionList() As Control = Nothing)
        Dim ExcpArray As System.Array = ExceptionList
        Dim skip As Boolean = False
        If p_Root.Controls.Count = 0 Then Return
        For Each ctl As Control In p_Root.Controls
            skip = False
            If Not (ExceptionList Is Nothing) Then
                If Array.IndexOf(ExceptionList, ctl) >= 0 Then skip = True
            End If
            If Not (skip) Then
                ClearControlData(ctl)
                ClearAllControlsData(ctl, ExceptionList)
            End If
        Next
    End Sub

    Sub SetAllControlsState(ByRef p_Root As System.Windows.Forms.Control, ByVal isReadOnly As Boolean, Optional ByRef ExceptionControl As System.Windows.Forms.Control = Nothing)
        Call SetAllControlsState(p_Root, isReadOnly, New Control() {ExceptionControl})
    End Sub

    Sub SetAllControlsState(ByRef p_objRoot As Object, ByVal isReadOnly As Boolean, ByRef ExceptionList() As System.Windows.Forms.Control)
        Dim p_Root As Control = CType(p_objRoot, Control)
        Dim ExcpArray As System.Array = ExceptionList
        Dim skip As Boolean = False

        If p_Root.Controls.Count = 0 Then Return
        For Each ctl As Control In p_Root.Controls
            skip = False
            If Not (ExceptionList Is Nothing) Then
                If Array.IndexOf(ExceptionList, ctl) >= 0 Then skip = True
            End If
            If Not (skip) Then
                SetControlState(ctl, isReadOnly)
                SetAllControlsState(ctl, isReadOnly, ExceptionList)
            End If
        Next
    End Sub

    Public Sub ShowBusyIcon2(Optional ByVal BusyState As Boolean = False)
        'Dim PreCur As Cursor = gfrmMain.Cursor
        ' ESS Start 29-08-2004
        '     ESS End 29-08-2004
        'gfrmMain.Enabled = Not (BusyState)
        gAppBusy = BusyState
        If IcoGrn Is Nothing Then
            Dim bmp As Drawing.Bitmap = New Drawing.Bitmap(CType(gfrmMain.imglstReady.Images(0), Image))
            IcoGrn = Drawing.Icon.FromHandle(bmp.GetHicon())
            bmp.Dispose()
            bmp = Nothing
        End If
        SetPanelIcon(gfrmMain.statusPanel_Ready, CType(IIf(BusyState, IcoRed, IcoGrn), Icon))
        'gfrmMain.Cursor.Current = Cursors.WaitCursor
        'gfrmMain.Enabled = Not (BusyState)
        'gfrmMain.Cursor = PreCur 'IIf(BusyState, Cursors.WaitCursor, Cursors.Default)

        gfrmMain.MenuFile.Enabled = Not (BusyState)
        gfrmMain.menuServices.Enabled = ((Not (BusyState)) And (IsLoggedOn))
        gfrmMain.menuManage.Enabled = ((Not (BusyState)) And (IsLoggedOn))
        gfrmMain.menuTools.Enabled = Not (BusyState)
        gfrmMain.menuAddsIns.Enabled = ((Not (BusyState)) And (IsLoggedOn))
        gfrmMain.MenuWindow.Enabled = Not (BusyState)
        gfrmMain.menuHelp.Enabled = Not (BusyState)
        gfrmMain.MenuItemDTS.Enabled = (Not (BusyState) And (IsLoggedOn))


        Application.DoEvents()
        gfrmMain.Cursor = IIf(BusyState, Cursors.WaitCursor, Cursors.Default)
        Cursor.Current = gfrmMain.Cursor
    End Sub


    Private Sub AccessControl(Optional ByVal BusyState As Boolean = False)
        Try
            If gfrmMain.InvokeRequired Then
                gfrmMain.Invoke(New MethodInvoker(AddressOf AccessControl))
            Else
                gAppBusy = BusyState
                If IcoGrn Is Nothing Then
                    Dim bmp As Drawing.Bitmap = New Drawing.Bitmap(CType(gfrmMain.imglstReady.Images(0), Image))
                    IcoGrn = Drawing.Icon.FromHandle(bmp.GetHicon())
                    bmp.Dispose()
                    bmp = Nothing
                End If
                SetPanelIcon(gfrmMain.statusPanel_Ready, CType(IIf(BusyState, IcoRed, IcoGrn), Icon))
                'gfrmMain.Cursor.Current = Cursors.WaitCursor
                'gfrmMain.Enabled = Not (BusyState)
                'gfrmMain.Cursor = PreCur 'IIf(BusyState, Cursors.WaitCursor, Cursors.Default)

                gfrmMain.MenuFile.Enabled = Not (BusyState)
                gfrmMain.menuServices.Enabled = ((Not (BusyState)) And (IsLoggedOn))
                gfrmMain.menuManage.Enabled = ((Not (BusyState)) And (IsLoggedOn))
                gfrmMain.menuTools.Enabled = Not (BusyState)
                'gfrmMain.menuAddsIns.Enabled = ((Not (BusyState)) And (IsLoggedOn))
                gfrmMain.MenuWindow.Enabled = Not (BusyState)
                gfrmMain.menuHelp.Enabled = Not (BusyState)
                gfrmMain.MenuItemDTS.Enabled = (Not (BusyState) And (IsLoggedOn))


                Application.DoEvents()
                gfrmMain.Cursor = IIf(BusyState, Cursors.WaitCursor, Cursors.Default)
                Cursor.Current = gfrmMain.Cursor
            End If
        Catch ex As ObjectDisposedException
            'Do Nothing
        End Try

    End Sub

    Sub MyThread1(Optional ByVal BusyState As Boolean = False)
        ' Working code
        ' Working code
        ' Working code
        ' Working code
        ' Working code
        ' Working code

        AccessControl(BusyState)

    End Sub

    Public Sub ShowBusyIcon(Optional ByVal BusyState As Boolean = False)
        Dim Strt As System.Threading.Thread
        Strt = New System.Threading.Thread(AddressOf MyThread1)
        Strt.Start()

    End Sub

    Public Sub ShowConnectedIcon(Optional ByVal ConnectState As Boolean = True)
        SetPanelIcon(gfrmMain.statusPanel_Online, gfrmMain.imglstOnline, IIf(ConnectState, 1, 0))
    End Sub

    Public Sub ShowActivityIcon(Optional ByVal Chx As enumActivityType = enumActivityType.None)
        If IcoGry Is Nothing Then
            Dim bmp As Drawing.Bitmap = New Drawing.Bitmap(CType(gfrmMain.imglstActivity.Images(0), Image))
            IcoGry = Drawing.Icon.FromHandle(bmp.GetHicon())
            bmp.Dispose()
            bmp = Nothing
        End If
        If IcoRed Is Nothing Then
            Dim bmp As Drawing.Bitmap = New Drawing.Bitmap(CType(gfrmMain.imglstActivity.Images(3), Image))
            IcoRed = Drawing.Icon.FromHandle(bmp.GetHicon())
            bmp.Dispose()
            bmp = Nothing
        End If
        Dim IcoWrite As Icon = IcoGry
        Dim IcoWait As Icon = IcoGry
        Dim IcoRead As Icon = IcoGry
        Select Case Chx
            Case enumActivityType.None
                gfrmMain.statusPanel_Text.Text = ""
            Case enumActivityType.Write
                IcoWrite = IcoRed
                gfrmMain.statusPanel_Text.Text = "Writing Message to Queue..."
            Case enumActivityType.Wait
                IcoWait = IcoRed
                gfrmMain.statusPanel_Text.Text = "Waiting for a Message in Queue..."
            Case enumActivityType.Read
                IcoRead = IcoRed
                gfrmMain.statusPanel_Text.Text = "Reading a Message from Queue..."
        End Select
        SetPanelIcon(gfrmMain.statusPanel_Activity_Write, IcoWrite)
        SetPanelIcon(gfrmMain.statusPanel_Activity_Wait, IcoWait)
        SetPanelIcon(gfrmMain.statusPanel_Activity_Read, IcoRead)
        Application.DoEvents()
    End Sub

    Public Sub ShowActivityIconOld(Optional ByVal Chx As enumActivityType = enumActivityType.None)
        Dim chxWrite As Integer = 0
        Dim chxWait As Integer = 1
        Dim chxRead As Integer = 2
        Select Case Chx
            Case enumActivityType.None
                gfrmMain.statusPanel_Text.Text = ""
            Case enumActivityType.Write
                chxWrite = 3
                gfrmMain.statusPanel_Text.Text = "Writing Message to Queue..."
            Case enumActivityType.Wait
                chxWait = 4
                gfrmMain.statusPanel_Text.Text = "Waiting for a Message in Queue..."
            Case enumActivityType.Read
                chxRead = 5
                gfrmMain.statusPanel_Text.Text = "Reading a Message from Queue..."
        End Select
        SetPanelIcon(gfrmMain.statusPanel_Activity_Write, gfrmMain.imglstActivity, chxWrite)
        SetPanelIcon(gfrmMain.statusPanel_Activity_Wait, gfrmMain.imglstActivity, chxWait)
        SetPanelIcon(gfrmMain.statusPanel_Activity_Read, gfrmMain.imglstActivity, chxRead)
        Application.DoEvents()
    End Sub

    Public Sub ShowStatusText(Optional ByVal msgText As String = Nothing)
        'Try block to maintain capabilities with older ICEP0100.EXE files
        Try
            gfrmMain.gStatusText = msgText
            gfrmMain.MainStatusBar.Invalidate()
        Catch
            'Do Nothing
        End Try
        Try
            'Application.DoEvents()
            gfrmMain.statusPanel_Text.Text = msgText 'String.Empty
            If (Not (msgText Is Nothing)) OrElse (msgText <> "") Then
                tmrStatusText.Stop()
                tmrStatusText.Start()
            End If
        Catch ex As Exception
            Debug.WriteLine("Appinstance.ShowStatusText: " & "err")
        End Try
    End Sub

    Public Sub HideAllPanels(ByRef Container As Control)
        'Dim o As Panel
        For i As Integer = 0 To Container.Controls.Count - 1
            If Container.Controls(i).GetType.Name = "Panel" Then
                Container.Controls(i).Hide()
            End If
        Next
    End Sub

    Public Sub HideAllPanels(ByRef Container As Control, ByVal Exception As Panel)
        'Dim o As Panel
        For i As Integer = 0 To Container.Controls.Count - 1
            If Container.Controls(i).GetType.Name = "Panel" AndAlso Not (Container.Controls(i) Is Exception) Then
                Container.Controls(i).Hide()
            End If
        Next
    End Sub

    Public Sub ShowPanel(ByRef pnl As Windows.Forms.Panel, ByRef frm As Form)
        For i As Integer = 0 To frm.Controls.Count - 1
            If frm.Controls(i).GetType.Equals(pnl.GetType) Then
                If Not (pnl Is frm.Controls(i)) Then
                    frm.Controls(i).Hide()
                End If
            End If
        Next
        pnl.Dock = DockStyle.Fill
        pnl.Show()
        pnl.BringToFront()
    End Sub

    Public Sub ShowPanel(ByRef ctl As Control, ByRef Container As Control)
        For i As Integer = 0 To Container.Controls.Count - 1
            If Container.Controls(i).GetType.Equals(ctl.GetType) Then
                If Not (ctl Is Container.Controls(i)) Then
                    Container.Controls(i).Hide()
                End If
            End If
        Next
        ctl.Dock = DockStyle.Fill
        ctl.Show()
        ctl.BringToFront()
    End Sub

    Public Sub MenuErrosItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim MnuItm As MenuItem
        Dim QueItm As LastErr_Struct
        MnuItm = CType(sender, MenuItem)
        QueItm = CType(LastErrQue.ToArray(LastErrQue.Count - 1 - MnuItm.Index), LastErr_Struct)
        ModalMsgBox("Error at " & QueItm.ErrTme & "!" & vbCrLf & QueItm.ErrDsc, MsgBoxStyle.Exclamation)
    End Sub


    Public Sub HandleError(ByVal ErrorNum As Integer, ByVal ErrStr As String, Optional ByVal ErrTtl As String = "", Optional ByRef p_exception As Exception = Nothing)
        Dim st As String
        Dim QueItm As LastErr_Struct
        Dim QueArr() As Object
        Dim i As Integer

        ShowBusyIcon()

#If DEBUG Then
        If Not (p_exception Is Nothing) Then
            Debug.WriteLine("AppInstance.HandleError: " & String.Format("{0}:{1}{2} {3}", ErrorNum, ErrTtl, vbCrLf, p_exception.ToString))
            Debug.WriteLine("AppInstance.HandleError: " & p_exception.StackTrace)
        End If
#End If
        Try
            st = ErrStr
            If Not (ErrStr Is Nothing) Then ErrStr = ErrStr.Replace(vbCrLf, " ")
            gLastErrorStr = ErrStr
            'Add error to Last 10 erros queue
            If LastErrQue.Count >= 10 Then LastErrQue.Dequeue()
            QueItm.ErrTme = Now
            QueItm.ErrDsc = ErrStr
            LastErrQue.Enqueue(QueItm)
            QueArr = LastErrQue.ToArray
            gfrmMain.menuErrors.MenuItems.Clear()
            Dim mu As Menu = gfrmMain.menuErrors

            For i = LastErrQue.Count - 1 To 0 Step -1
                mu.MenuItems.Add(QueArr(i).ErrTme & " : " & SafeSubString(QueArr(i).ErrDsc, 0, 60), _
                AddressOf MenuErrosItem_Click)
            Next

            Logger.LogError(ErrorNum, ErrStr)
            ShowStatusText(ErrStr)
            SetPanelIcon(gfrmMain.statusPanel_Error, gfrmMain.imglstError, 0)
            gfrmMain.tmrErrorIconClear.Start()
            If Trim(ErrTtl) <> "" Then
                ModalMessageBox(st, , , MessageBoxIcon.Error, CStr(IIf(ErrorNum < &H80000000, ErrorNum, Format(ErrorNum, "x") & ": " & ErrTtl)))
            End If
        Catch
            'Nothing, if error handling failed, then the program must not be affected.
        End Try

    End Sub

    Public ReadOnly Property IceEqnBrn() As String
        Get
            Return getICEPProperty("IceEqnBrn")
        End Get
    End Property

    Public ReadOnly Property IceEqnRol() As String
        Get
            Return getICEPProperty("IceEqnRol")
        End Get
    End Property

    Public ReadOnly Property IceEqnTit() As String
        Get
            Return getICEPProperty("IceEqnTit")
        End Get
    End Property

    Private Function getICEPProperty(ByVal p_property As String) As String
        If (p_property Is Nothing) OrElse (p_property.Trim = "") Then Return ""

        Dim oAssem As System.Reflection.Assembly
        'TMPL0100.extractInternalResource(TemplateName, tmpDot)
        Try
            'If oAssem Is Nothing Then
            oAssem = System.Reflection.Assembly.GetEntryAssembly
            'End If

            Dim modtyp As Type = oAssem.GetType("ICEP0100.ICEComm")
            If modtyp Is Nothing Then Throw New VersionNotFoundException
            Dim rmi As PropertyInfo = modtyp.GetProperty(p_property)
            If rmi Is Nothing Then Throw New VersionNotFoundException
            Return rmi.GetValue(Nothing, New Object() {})
        Catch ex As VersionNotFoundException
            'Return FormatValString(CcyAmt, 2)
        End Try
        Return ""
    End Function

    Public Function getICEArray(ByVal p_property As String) As Array
        If (p_property Is Nothing) OrElse (p_property.Trim = "") Then Return Nothing

        Dim oAssem As System.Reflection.Assembly
        Try
            'If oAssem Is Nothing Then
            oAssem = System.Reflection.Assembly.GetEntryAssembly
            'End If

            Dim modtyp As Type = oAssem.GetType("ICEP0100.ICEComm")
            If modtyp Is Nothing Then Return Nothing
            Dim rmi As FieldInfo = modtyp.GetField(p_property)
            If rmi Is Nothing Then Return Nothing
            Return rmi.GetValue(Nothing)
        Catch ex As Exception
            Return Nothing
        End Try
    End Function

    Public Function getFullDumb(ByVal ErrorNum As Integer, ByVal ErrorStr As String, Optional ByVal pLogLvl As Integer = 1, Optional ByVal exp_ As Exception = Nothing) As String
        Dim oFrame As StackFrame
        Dim sDumb As String = ""
        Dim sTab As String = vbTab
        Dim sAllOtherProcesses As String = ""
        Dim OldCultInfo, NewCultInfo As CultureInfo
        OldCultInfo = System.Threading.Thread.CurrentThread.CurrentCulture
        NewCultInfo = New CultureInfo("en-US", False)
        System.Threading.Thread.CurrentThread.CurrentCulture = NewCultInfo

        Try

            oFrame = New StackTrace(True).GetFrame(1) 'This will get the caller StackFrame

            sDumb &= SIBL0100.Util.UString.repeat(150, "*") & vbCrLf

            sDumb &= sTab & Date.Now.ToString("yyyy/MM/dd hh:mm:ss") & vbCrLf
            sDumb &= sTab & "Error#: " & ErrorNum & vbCrLf
            sDumb &= sTab & "Error Message: " & ErrorStr & vbCrLf

            sDumb &= vbCrLf & sTab & "-------------------Error-------------------" & vbCrLf
            With oFrame
                sDumb &= sTab & "Error In File: " & .GetFileName & vbCrLf
                sDumb &= sTab & "Error In Method: " & SIBL0100.Util.Debug.getStackFrame(oFrame) & vbCrLf
                sDumb &= sTab & "Error Line: " & .GetFileLineNumber & vbCrLf
            End With

            sDumb &= vbCrLf & sTab & "-------------------Environment-------------------" & vbCrLf
            'With Environment
            sDumb &= sTab & "CommandLine: " & Environment.CommandLine & vbCrLf
            sDumb &= sTab & "CommandLine Args: " & Environment.GetCommandLineArgs.ToString & vbCrLf
            sDumb &= sTab & "Current Directory: " & Environment.CurrentDirectory & vbCrLf
            'sDumb &= sTab & "Environment Variables: " & Environment.GetEnvironmentVariables & vbCrLf
            sDumb &= sTab & "Machine Name: " & Environment.MachineName & vbCrLf
            sDumb &= sTab & "IP: " & getIP() & vbCrLf
            sDumb &= sTab & "OS Version: " & Environment.OSVersion.Version.ToString & vbCrLf
            sDumb &= sTab & "User Domain Name: " & Environment.UserDomainName & vbCrLf
            sDumb &= sTab & ".Net Version: " & Environment.Version.ToString & vbCrLf


            ' Find all printers installed
            Dim iCountPrinters As Integer
            Dim sPrinters As String
            sPrinters = getPrinters(iCountPrinters)

            If sPrinters = "" Then
                sDumb &= sTab & "Installed Printers: (0) Printers installed." & vbCrLf
            Else
                sDumb &= sTab & "Installed Printers: (" & iCountPrinters & ")" & sPrinters & vbCrLf
            End If

            sDumb &= vbCrLf & sTab & "-------------------ICE Memory info-------------------" & vbCrLf

            Dim Proc As Process = Process.GetCurrentProcess
            Proc.Refresh()
            'sDumb &= sTab & "Process PrivateMemorySize: " & String.Format("{0:0,00} k", Proc.PrivateMemorySize / 1024) & vbCrLf
            'sDumb &= sTab & "Process NonpagedSystemMemorySize: " & String.Format("{0:0,00} k", Proc.NonpagedSystemMemorySize / 1024) & vbCrLf
            'sDumb &= sTab & "Process PagedMemorySize: " & String.Format("{0:0,00} k", Proc.PagedMemorySize / 1024) & vbCrLf
            'sDumb &= sTab & "Process PagedSystemMemorySize: " & String.Format("{0:0,00} k", Proc.PagedSystemMemorySize / 1024) & vbCrLf
            'sDumb &= sTab & "Process PeakVirtualMemorySize: " & String.Format("{0:0,00} k", Proc.PeakVirtualMemorySize / 1024) & vbCrLf
            sDumb &= sTab & "Process StartTime: " & Proc.StartTime & vbCrLf
            sDumb &= sTab & "Process EndTime: " & Date.Now.ToString & vbCrLf
            sDumb &= sTab & "Process Threads: " & Proc.Threads.Count & vbCrLf

            sDumb &= vbCrLf & sTab & "-------------------ICE info-------------------" & vbCrLf

            sDumb &= sTab & "Logged User: " & Me.LoggedUser & vbCrLf
            sDumb &= sTab & "Equation Branch: " & IceEqnBrn & vbCrLf
            sDumb &= sTab & "Unit Name: " & Me.Unit & ", " & Me.UnitName & vbCrLf
            sDumb &= sTab & "Equation Role: " & IceEqnRol & vbCrLf
            sDumb &= sTab & "Equation Title: " & IceEqnTit & vbCrLf

            Dim sAsmOne As String = ""
            Dim sAsm As String = ""
            Dim sIceBuild As String = ""
            Dim sIceVersion As String = ""
            For Each asm As System.Reflection.Assembly In System.AppDomain.CurrentDomain.GetAssemblies()
                sAsmOne = asm.GetName.Name
                'sIceBuild = CStr(System.Diagnostics.FileVersionInfo.GetVersionInfo(asm.Location).FilePrivatePart)
                sIceVersion = System.Diagnostics.FileVersionInfo.GetVersionInfo(asm.Location).FileMajorPart & "." & _
                                  System.Diagnostics.FileVersionInfo.GetVersionInfo(asm.Location).FileMinorPart & "." & _
                                  System.Diagnostics.FileVersionInfo.GetVersionInfo(asm.Location).FileBuildPart
                sAsm &= ", " & sAsmOne & " v" & sIceVersion & sIceBuild
            Next

            sDumb &= sTab & "Loaded Assemblies: " & sAsm.Substring(1).Trim & vbCrLf

            'To get local address

            'Dim iCountProcs As Integer = 0
            'Dim iMemory As Integer = 0
            'For Each Proc In Process.GetProcesses
            '    Proc.Refresh()
            '    If Not Proc.ProcessName Is Nothing AndAlso Proc.ProcessName.Trim <> "" Then
            '        sAllOtherProcesses &= ", " & Proc.ProcessName.Trim
            '    End If
            '    iMemory += Proc.PagedMemorySize
            '    iCountProcs += 1
            'Next
            'sDumb &= sTab & "Process: (" & iCountProcs & ") -->" & sAllOtherProcesses.Substring(1).Replace(vbTab, "") & vbCrLf
            'sDumb &= sTab & "Processes (Sum) PagedMemorySize: " & String.Format("{0:0,00} k", iMemory / 1024) & vbCrLf


            sDumb &= vbCrLf & sTab & "-------------------Exception-------------------" & vbCrLf
            If Not exp_ Is Nothing Then
                With exp_
                    sDumb &= sTab & "Exception Message: " & .Message & vbCrLf
                    sDumb &= sTab & "Exception Source: " & .Source & vbCrLf
                    sDumb &= sTab & "Exception Full Data: " & .ToString & vbCrLf
                    sDumb &= sTab & "Exception Full Data: " & .StackTrace & vbCrLf
                End With
            End If

            sDumb &= SIBL0100.Util.UString.repeat(150, "*") & vbCrLf

            Return sDumb

        Catch ex As Exception
            ModalMsgBox("Error:" & sAllOtherProcesses)
        End Try
        System.Threading.Thread.CurrentThread.CurrentCulture = OldCultInfo
        NewCultInfo = Nothing
        GC.Collect()
        Return sDumb
    End Function

    Private Function getPrinters(ByRef iCountPrinters As Integer) As String
        ' Use the ObjectQuery to get the list of configured printers
        Dim sInstalledPrinters As String = ""
        iCountPrinters = 0
        ' Find all printers installed
        For Each pkInstalledPrinters As String In System.Drawing.Printing.PrinterSettings.InstalledPrinters
            sInstalledPrinters &= ", " & pkInstalledPrinters
            iCountPrinters += 1
        Next

        If sInstalledPrinters.Length > 0 Then
            Return sInstalledPrinters.Substring(1)
        End If
        Return sInstalledPrinters
    End Function

    Private Function getIP() As String
        Dim sIP As String = ""
        Dim sHostName As String
        Dim i As Integer
        sHostName = System.Net.Dns.GetHostName()
        Dim ipE As System.Net.IPHostEntry = System.Net.Dns.GetHostEntry(sHostName)
        Dim IpA() As System.Net.IPAddress = ipE.AddressList
        For i = 0 To IpA.GetUpperBound(0)
            Console.Write("IP Address {0}: {1} ", i, IpA(i).ToString)
            sIP = IpA(i).ToString
        Next
        Return sIP
    End Function

    Public Function GetUserBranch() As String
        Dim st As String
        st = Workstation
        If st.Chars(0) = "B" Or st.Chars(0) = "N" Then
            Return "0" & SafeSubString(st, 1, 3)
        ElseIf "SAIB" = SafeSubString(st, 0, 4) Then
            Return "0101"
        End If
        Return Nothing
    End Function

    Public Function PackString(ByVal prmStr As String, ByVal prmLen As Integer) As String
        Dim RetStr As Char()
        Dim i As Integer

        If prmStr Is Nothing Then Return Space(prmLen)
        If prmStr.Length >= prmLen Then Return Left(prmStr, prmLen)

        ReDim RetStr(prmLen - 1)
        For i = 0 To prmStr.Length - 1
            RetStr(i) = prmStr.Chars(i)
        Next
        For i = prmStr.Length To prmLen - 1
            RetStr(i) = Chr(Keys.Space)
        Next
        Return RetStr
    End Function

    Public Function strip(ByRef strMsg As String, ByVal nChars As Integer) As String
        Dim RetStr As String
        If strMsg Is Nothing Then Return ""
        If strMsg.Length < 1 Then Return ""
        'If Len(strMsg) < nChars Then strip = "": Exit Function
        If strMsg.Length <= nChars Then
            RetStr = strMsg
            strMsg = ""
        Else
            RetStr = Left(strMsg, nChars)
            strMsg = Right(strMsg, strMsg.Length - nChars)
        End If
        Return RetStr
    End Function

    Public Shared Function GetListViewItemAt(ByVal Ctl As ListView, ByVal x As Integer, ByVal y As Integer, ByRef Item As Integer, ByRef SubItem As Integer) As Integer
        '
        Dim tHitTest As LVHITTESTINFO
        Dim lRet As Integer

        With tHitTest
            .lFlags = 0
            .lItem = 0
            .lSubItem = 0
            .pt.x = x
            .pt.y = y
        End With
        '
        ' Return the filled Structure to the routine
        '
        lRet = SendMessage(Ctl.Handle, LVM_SUBITEMHITTEST, 0, tHitTest)
        Try
            If (tHitTest.lFlags = LVHT_NOWHERE) Then         ' empty space clicked
                Return 0
            ElseIf (tHitTest.lFlags = LVHT_ONITEMICON) Or _
                   (tHitTest.lFlags = LVHT_ONITEMLABEL) Or _
                   (tHitTest.lFlags = LVHT_ONITEMSTATEICON) Then
                Item = tHitTest.lItem
                SubItem = tHitTest.lSubItem
                Return 1
            Else
                Return -1
            End If
        Catch
        End Try
        Return -1
        '
    End Function

    Public Function FormatStpDat(ByVal StpWho As String, ByVal StpDte As String) As String
        If StpWho Is Nothing OrElse StpDte Is Nothing Then Return String.Empty
        Dim st As String = UnpackDate(StpDte, True) & " by " & StpWho.Trim 'msg 615000 'fixed FIR14711 for showing bad formatted date
        If st.Trim = "by" Then st = String.Empty
        Return st
    End Function

    Public Function FormatUsrStpDat(ByVal StpStr As String) As String
        If (StpStr Is Nothing) OrElse (StpStr.Trim = "") OrElse (StpStr.Length <> 30) Then
            'Abort processing, return the string as received
            Return StpStr
        End If
        Dim st As String = StpStr.Substring(3, 3).Trim
        Select Case st
            Case "NEW"
                st = "Created by "
            Case "MOD"
                st = "Modified by "
            Case "DEL"
                st = "Deleted by "
            Case Else
                st &= " by "
        End Select

        Select Case StpStr.Substring(0, 3).Trim
            Case "IBK" : st &= "internet user "
            Case "IVR" : st &= "phone banking user "
            Case "WAP" : st &= "mobile banking user "
            Case "USR" : st &= "bank user "
            Case "SMS" : st &= "SMS banking user "
            Case Else : st &= StpStr.Substring(0, 3).Trim & " user "
        End Select
        st &= StpStr.Substring(6, 10).Trim & " at "
        st &= UnpackDate(StpStr.Substring(16), True)
        Return st
    End Function

    Public Function SelectPrinterName(ByVal p_PrevSelectedPrinter As String) As String
        Dim l_SelectedPrinter As String
        Dim dlgPrn As New System.Windows.Forms.PrintDialog
        Try
            Dim objPrint As New System.Drawing.Printing.PrinterSettings
            Dim res As DialogResult
            dlgPrn.AllowPrintToFile = False
            dlgPrn.PrinterSettings = objPrint
            If p_PrevSelectedPrinter <> "" Then
                dlgPrn.PrinterSettings.PrinterName = p_PrevSelectedPrinter
            End If
            res = dlgPrn.ShowDialog()
            If res = DialogResult.OK Then
                l_SelectedPrinter = dlgPrn.PrinterSettings.PrinterName
            Else
                Return p_PrevSelectedPrinter
            End If
        Catch
            Return p_PrevSelectedPrinter
        End Try
        Return l_SelectedPrinter
    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Formats the user ID (by smartly comining the Corporate and User IDE
    ''' </summary>
    ''' <param name="CorIde">The user's corporate ID</param>
    ''' <param name="UsrIDe">The user's user ID</param>
    ''' <returns>The formatted value string is returned</returns>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[u910krzi]	01-Feb-2010	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Function FormatUserID(ByVal CorIde As String, ByVal UsrIde As String) As String
        If UsrIde Is Nothing Then UsrIde = ""
        If CorIde Is Nothing Then CorIde = ""

        Return CStr(IIf(UsrIde = CorIde, UsrIde.Trim, CorIde.Trim & ";" & UsrIde.Trim))
    End Function


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Formats a NS string into a value string based on how many decimal digits after the floating point.
    ''' </summary>
    ''' <param name="InStr">The NS string to be formatted</param>
    ''' <param name="exp">The number of decimal places</param>
    ''' <param name="bCommaSep">IF set to true, a thousands separator will be used</param>
    ''' <returns>The formatted value string is returned</returns>
    ''' <remarks>
    ''' Example: 0001500000+ with exp=2 will be turned into 15,000.00
    ''' </remarks>
    ''' <history>
    ''' 	[u910krzi]	04-Apr-2009	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Function FormatValString(ByVal InStr As String, ByVal exp As Integer, Optional ByVal bCommaSep As Boolean = True) As String
        Dim str As String = String.Empty
        Dim dd As Double
        Dim LeftPad As String = "0000"
        Dim Sign As String = ""
        If ((InStr Is Nothing) OrElse InStr.Trim = "") Then Return Nothing
        Try
            'If the input string is a signed number (such as NS string (with a trailing sign)), then remove the sign,
            'process it, then add the sign as a leading sign
            If InStr.IndexOf("+") >= 0 Then
                Sign = "+"
                InStr = InStr.Replace("+", "")
            End If
            If InStr.IndexOf("-") >= 0 Then
                Sign = "-"
                InStr = InStr.Replace("-", "")
            End If

            If InStr.Length < exp Then
                'InStr &= ".0"
                InStr = LeftPad.Substring(0, exp - InStr.Length) & InStr
            End If
            InStr = InStr.Insert(InStr.Length - exp, ".")
            dd = Val(InStr)
            If bCommaSep Then
                Select Case exp
                    Case 0 : str += Format(dd, "#,#0")
                    Case 1 : str += Format(dd, "#,#0.0")
                    Case 2 : str += Format(dd, "#,#0.00")
                    Case 3 : str += Format(dd, "#,#0.000")
                    Case 4 : str += Format(dd, "#,#0.0000")
                End Select
            Else
                Select Case exp
                    Case 0 : str += Format(dd, "#0")
                    Case 1 : str += Format(dd, "#0.0")
                    Case 2 : str += Format(dd, "#0.00")
                    Case 3 : str += Format(dd, "#0.000")
                    Case 4 : str += Format(dd, "#0.0000")
                End Select
            End If
            Return Sign & str 'Add the sign, if any, as a leading character

        Catch ex As Exception
            HandleError(&H80004003, "FormatValString:" & ex.Message)
            Return Nothing
        End Try
    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Format an amount string into an NS format (e.g. 0001500+ for 15.00)
    ''' </summary>
    ''' <param name="InStr">The input string; it must be an amount string (e.g. 15,000.00)</param>
    ''' <param name="NSLen">The required length, default is 15 characters (including the sign)</param>
    ''' <param name="AddSign">If set to true, then a sign would be appended at the end of the string</param>
    ''' <returns>The formatted string is returned</returns>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[u910krzi]	04-Apr-2009	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Function PackNSString(ByVal InStr As String, Optional ByVal NSLen As Integer = 15, Optional ByVal AddSign As Boolean = True) As String
        Dim PadStr As String = "000000000000000000000000000000000"
        Dim RetStr As String
        Dim idx As Integer

        RetStr = InStr
        If (RetStr Is Nothing) OrElse (RetStr.Trim = "") Then
            RetStr = "0" 'PadStr.Substring(0, NSLen)
            'Return RetStr
        End If
        RetStr = RetStr.Replace(",", "").Replace(".", "")
        If AddSign Then
            idx = RetStr.IndexOf("-")
            If (idx >= 0) Then
                RetStr = RetStr.Remove(idx, 1)
                RetStr &= "-"
            Else
                idx = RetStr.IndexOf("+")
                If (idx >= 0) Then
                    RetStr = RetStr.Remove(idx, 1)
                End If
                RetStr &= "+"
            End If
        End If

        RetStr = PadStr.Substring(0, NSLen - RetStr.Length) & RetStr
        Return RetStr
    End Function

    Public Function PackNString(ByVal InStr As String, Optional ByVal NSLen As Integer = 15) As String
        Return PackNSString(InStr, NSLen, False)
    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Functions that shows a specified currency by selecting its entry within a combo box
    ''' </summary>
    ''' <param name="cmb">The combo box containing the currencies</param>
    ''' <param name="CcyCde">The specified currency code</param>
    ''' <remarks>
    ''' This function assumes that the currency code is t
    ''' </remarks>
    ''' <history>
    ''' 	[u910krzi]	07-Jun-2009	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Sub ShowSelectedCcy(ByRef cmb As ComboBox, ByVal CcyCde As String)
        If (cmb Is Nothing) OrElse (CcyCde Is Nothing) OrElse (CcyCde.Trim = "") Then Return
        If cmb.Items.Count = 0 Then Return
        For i As Integer = 0 To cmb.Items.Count - 1
            If SafeSubString(CStr(cmb.Items(i)), 0, 3).Trim.ToUpper = CcyCde Then
                cmb.SelectedIndex = i
                Return
            End If
        Next
        cmb.SelectedIndex = -1
    End Sub

    Public Sub FillCcyCombo(ByRef cmb As ComboBox, Optional ByVal bTradedOnly As Boolean = True)
        'Refelection: This method is called from ICEP0100
        Dim l_Assm As [Assembly] = [Assembly].GetEntryAssembly
        Dim modtyp As Type = l_Assm.GetType("ICEP0100.ICEDataBase")
        Dim rmi As MethodInfo = modtyp.GetMethod("FillCcyCombo")
        rmi.Invoke(Nothing, New Object() {cmb, bTradedOnly})
    End Sub


    Public Function ReverseFormatNSString(ByVal InStr As String, ByVal exp As Integer, Optional ByVal bCommaSep As Boolean = True) As String
        Dim str As String
        If ((InStr Is Nothing) Or InStr = "") Then Return Nothing
        Try
            str = Left(InStr, 1)
            InStr = InStr.Substring(1, InStr.Length - 1)
            str &= FormatValString(InStr, exp, bCommaSep)
            Return str

        Catch ex As Exception
            HandleError(&H80004004, "ReverseFormatNSString:" & ex.Message)
            Return Nothing
        End Try
    End Function

    Public Function ReverseNSString(ByVal InStr As String) As String
        Dim str As String
        If ((InStr Is Nothing) Or InStr = "") Then Return Nothing
        Try
            str = Right(InStr, 1)
            InStr = InStr.Substring(0, InStr.Length - 1)
            str = str & InStr
            Return str
        Catch ex As Exception
            HandleError(&H80004004, "ReverseNSString:" & ex.Message)
            Return Nothing
        End Try
    End Function

    Public Function FormatNSString(ByVal InStr As String, ByVal exp As Integer, Optional ByVal bCommaSep As Boolean = True) As String
        Dim str As String
        Dim StrZero As String = "000000000000000000000000000000000000000000000000000000000000000000000000000"
        'Dim dd As Double
        If ((InStr Is Nothing) Or InStr = "") Then Return Nothing
        Try
            str = Right(InStr, 1)
            InStr = InStr.Substring(0, InStr.Length - 1)
            If StrZero.IndexOf(InStr) < 0 Then 'True when InStr is not all zeros, if so, add the sign
                str &= FormatValString(InStr, exp, bCommaSep)
            Else
                str = FormatValString(InStr, exp, bCommaSep)
            End If
            Return str

        Catch ex As Exception
            HandleError(&H80004004, "FormatNSString:" & ex.Message)
            Return Nothing
        End Try
    End Function

    Public Function FormatDBCRString(ByVal InStr As String) As String
        Dim t As Char
        Dim RetStr As String

        If (InStr Is Nothing) Then Return Nothing
        t = InStr.Chars(0)
        'If value is 0, don't add a CR/DB
        'If ExVal(InStr) = 0.0 Then Return InStr.Substring(1)
        If ExVal(InStr) = 0.0 Then
            If InStr.Chars(0) = CChar("0") Then
                Return InStr
            Else
                Return InStr.Substring(1)
            End If
        End If
        'else Replace - with DB and + with CR
        If (t = "-") Then
            RetStr = InStr.Substring(1)
            RetStr += " DB"
        Else
            RetStr = IIf((t = "+"), InStr.Substring(1), InStr)
            RetStr += " CR"
        End If
        Return RetStr
    End Function

    Public Sub FormatMQSError(ByVal MsgSig As String, ByVal ErrNum As Integer, ByVal ErrStr As String)
        If ErrNum = 2033 Then
            HandleError(&H80002001, MsgSig & ":" & MQS.ErrNum & ":" & MQS.ErrDsc)
            ModalMsgBox("Error 80002001: No Response was received from the Host!" & vbCrLf & "Diagnostics Information: " & _
                    MQS.ErrNum & ": " & MQS.ErrDsc, MsgBoxStyle.Exclamation, "ICE Communications")
        Else
            HandleError(&H80002002, MsgSig & ":" & MQS.ErrNum & ":" & MQS.ErrDsc, "ICE Communications")
        End If
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Function to format a date string into ICE standard format.
    ''' </summary>
    ''' <param name="DteStr">The date string to be formated; input format is YYYYDDMM with HHmmSS optional</param>
    ''' <param name="AddTime">IF true, the function will append the time to the formatted date</param>
    ''' <returns>The formatted date string</returns>
    ''' <remarks>
    ''' Higri dates are automatically detected and formatted appropriatly.
    ''' </remarks>
    ''' <history>
    ''' 	[u910krzi]	04-Apr-2009	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Function UnpackDate(ByVal DteStr As String, Optional ByVal AddTime As Boolean = False) As String
        Dim strDate As String
        Dim dt As DateTime
        Dim ZeroString As String = "00000000000000000000000000000000"

        If DteStr Is Nothing Then Return ""

        DteStr = Trim(DteStr)

        If DteStr.Length < 8 Then Return ""
        If (DteStr = SafeSubString(ZeroString, 0, DteStr.Length)) Then
            Return ""
        End If
        Try
            'Try to determine is the date is Higri or Gregorian, it will assume that years 1000 to 1699 belongs to the Hijri range
            Dim st As String = DteStr.Substring(0, 4)
            If (st >= "1000") And (st < "1700") Then 'It is assumed as Higri date
                'dt = New DateTime(DteStr.Substring(0, 4), DteStr.Substring(4, 2), DteStr.Substring(6, 2), New System.Globalization.HijriCalendar)
                strDate = DteStr.Substring(0, 4) & "/" & DteStr.Substring(4, 2) & "/" & DteStr.Substring(6, 2)
            Else
                dt = New DateTime(DteStr.Substring(0, 4), DteStr.Substring(4, 2), DteStr.Substring(6, 2), New System.Globalization.GregorianCalendar)
                'dt = DateSerial(CInt(DteStr.Substring(0, 4)), CInt(DteStr.Substring(4, 2)), CInt(DteStr.Substring(6, 2)))
                '    dt = DteStr.Substring(6, 2) & "/" & DteStr.Substring(4, 2) & "/" & DteStr.Substring(0, 4)
                strDate = IceFormatDate(dt, "dd-MMM-yyyy")  'dt.ToShortDateString
            End If

        Catch ex As Exception
            HandleError(&H80004005, "FormatDate:Invalid date format.")
            Return Nothing
        End Try

        If AddTime Then
            If DteStr.Trim.Length <> 14 Then
                strDate &= " 00:00:00"
            Else
                strDate += " "
                strDate += SafeSubString(DteStr, 8, 2)
                strDate += ":"
                strDate += SafeSubString(DteStr, 10, 2)
                strDate += ":"
                strDate += SafeSubString(DteStr, 12, 2)
            End If
        End If
        Return strDate
    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Function that will format the date (and time) according to the supplied format.
    ''' </summary>
    ''' <param name="DteStr">The date string to be formated; input format is YYYYDDMM with HHmmSS optional</param>
    ''' <param name="DateFormat">The date format desired</param>
    ''' <param name="TimeFormat">The time format deisred</param>
    ''' <returns>The formatted string</returns>
    ''' <remarks>
    ''' Set TimeFormat to none if time is not to be included in the formatted string
    ''' </remarks>
    ''' <history>
    ''' 	[u910krzi]	04-Apr-2009	Created
    ''' 	[Y093sahu]	14-Feb-2010	modified to allow the return of the time without the date
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Function UnpackDateEx(ByVal DteStr As String, ByVal DateFormat As String, ByVal TimeFormat As String) As String
        Dim strDate As String
        Dim dt As DateTime
        Dim ZeroString As String = "00000000000000000000000000000000"

        If DteStr Is Nothing Then Return ""

        DteStr = Trim(DteStr)

        If DteStr.Length < 8 Then Return ""
        If (DteStr = SafeSubString(ZeroString, 0, DteStr.Length)) Then
            Return ""
        End If
        Try
            If (TimeFormat Is Nothing) OrElse TimeFormat.Trim = "" Then
                dt = New DateTime(DteStr.Substring(0, 4), DteStr.Substring(4, 2), DteStr.Substring(6, 2), _
                       New System.Globalization.GregorianCalendar)
                strDate = IceFormatDate(dt, DateFormat)
            Else
                dt = New DateTime(DteStr.Substring(0, 4), DteStr.Substring(4, 2), DteStr.Substring(6, 2), _
                        SafeSubString(DteStr, 8, 2), SafeSubString(DteStr, 10, 2), SafeSubString(DteStr, 12, 2), _
                        New System.Globalization.GregorianCalendar)
                'dt = DateSerial(CInt(DteStr.Substring(0, 4)), CInt(DteStr.Substring(4, 2)), CInt(DteStr.Substring(6, 2)))
                '    dt = DteStr.Substring(6, 2) & "/" & DteStr.Substring(4, 2) & "/" & DteStr.Substring(0, 4)
                If (DateFormat Is Nothing) OrElse DateFormat.Trim = "" Then
                    strDate = Format(dt, TimeFormat)   'dt.ToShortDateString
                Else
                    strDate = IceFormatDate(dt, DateFormat) & Format(dt, TimeFormat) 'dt.ToShortDateString
                End If

            End If

        Catch ex As Exception
            HandleError(&H80004005, "UnpackDateEx:Invalid date format.")
            Return Nothing
        End Try
        Return strDate
    End Function

    Public Function FormatHijriDate(ByVal DteStr As String, Optional ByVal AddTime As Boolean = False) As String
        Dim strDate As String
        'Dim dt As DateTime
        Dim ZeroString As String = "00000000000000000000000000000000"

        If DteStr Is Nothing Then Return ""

        DteStr = Trim(DteStr)

        If DteStr.Length < 8 Then Return ""
        If (DteStr = SafeSubString(ZeroString, 0, DteStr.Length)) Then
            Return ""
        End If
        Try
            'dt = New DateTime(DteStr.Substring(0, 4), DteStr.Substring(4, 2), DteStr.Substring(6, 2), New System.Globalization.HijriCalendar)
            strDate = DteStr.Substring(6, 2) & "/" & DteStr.Substring(4, 2) & "/" & DteStr.Substring(0, 4)
        Catch ex As Exception
            HandleError(&H80004005, "FormatHijriDate:Invalid date format.")
            Return Nothing
        End Try

        If AddTime Then
            strDate += " "
            strDate += SafeSubString(DteStr, 8, 2)
            strDate += ":"
            strDate += SafeSubString(DteStr, 10, 2)
            strDate += ":"
            strDate += SafeSubString(DteStr, 12, 2)
        End If
        Return strDate
    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Function to format SAIB ATM card number as nnnnnn-nnnnnnnn-nnnn.
    ''' </summary>
    ''' <param name="AtmNum">ATM card number to be formatted</param>
    ''' <returns>Formatted account number.</returns>
    ''' <remarks>
    ''' If the provided ATM card number is invalid, the input string is returned as is
    ''' </remarks>
    ''' <history>
    ''' 	[u910krzi]	04-Apr-2009	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Function FormatATMCard(ByVal AtmNum As String, Optional ByVal ShowClearValue As Boolean = False) As String
        Dim RetStr As String
        AtmNum = Trim(AtmNum)
        If (AtmNum Is Nothing) Then Return Nothing
        If AtmNum.Length <> 16 Then
            HandleError(&H80004006, "FormatATMCard:Wrong Length:" & AtmNum)
            Return AtmNum
        End If
        If ShowClearValue Then
            RetStr = AtmNum.Substring(0, 6) & "-" & AtmNum.Substring(6, 6) & "-" & AtmNum.Substring(12)
        Else
            RetStr = AtmNum.Substring(0, 6) & "-******-" & AtmNum.Substring(12)
        End If
        Return RetStr
    End Function

    Public Function FormatLoan(ByVal LoaStr As String) As String
        If LoaStr Is Nothing Then Return ""
        If LoaStr = "" Then Return ""
        If LoaStr.Length < 8 Then Return LoaStr
        Dim st As String
        st = strip(LoaStr, 4) & "-"
        st &= strip(LoaStr, 3) & "-"
        st &= LoaStr
        Return st
    End Function

    Public Function FormatQueryCount(ByVal p_ItemCount As Integer, Optional ByVal p_Unit As String = "record") As String
        If p_ItemCount = 0 Then
            Return "No " & p_Unit & "s found"
        ElseIf p_ItemCount = 1 Then
            Return "Showing 1 " & p_Unit
        Else
            Return "Showing " & p_ItemCount & " " & p_Unit & "s"
        End If
    End Function

    Public Function FormatQueryCount(ByVal p_ReturnedCount As Integer, ByVal p_TotalRecords As Integer, Optional ByVal p_Unit As String = "record") As String
        If p_ReturnedCount = 0 Then
            Return "No " & p_Unit & "s found."
        ElseIf p_ReturnedCount = 1 Then
            Return "Showing 1 " & p_Unit
        Else
            If p_ReturnedCount = p_TotalRecords Then
                Return "Showing " & Format(p_ReturnedCount, "#,#") & " " & p_Unit & "s."
            Else
                Return "Showing " & Format(p_ReturnedCount, "#,#") & " of " & Format(p_TotalRecords, "#,#") & " " & p_Unit & "s."
            End If
        End If
    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Function to format an ASR number as nnnn-nnnn-nnnn-nnnn.
    ''' </summary>
    ''' <param name="AsrVal">ASR number to be formatted</param>
    ''' <returns>Formatted account number.</returns>
    ''' <remarks>
    ''' If the provided ASR number is invalid, the input string is returned as is
    ''' </remarks>
    ''' <history>
    ''' 	[u910krzi]	04-Apr-2009	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Function FormatASR(ByVal AsrVal As String) As String
        Dim RetStr As String
        If (AsrVal Is Nothing) Then Return Nothing
        AsrVal = CompactString(AsrVal)
        If AsrVal.Length <> 16 Then
            HandleError(&H80004022, "FormatASR:Wrong Length:" & AsrVal)
            Return AsrVal
        End If
        RetStr = AsrVal.Insert(12, "-")
        RetStr = RetStr.Insert(8, "-")
        RetStr = RetStr.Insert(4, "-")

        Return RetStr
    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Function to format SAIB account as nnnn-nnnnnn-nnn.
    ''' </summary>
    ''' <param name="AccNum">Account number to be formatted</param>
    ''' <returns>Formatted account number.</returns>
    ''' <remarks>
    ''' If the provided account number is invalid, the input string is returned as is
    ''' </remarks>
    ''' <history>
    ''' 	[u910krzi]	31-Mar-2009	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Function FormatAccount(ByVal AccNum As String) As String
        Dim RetStr As String = AccNum
        If (RetStr Is Nothing) OrElse (RetStr.Trim = "") Then Return ""
        RetStr = RetStr.Replace(" ", "")
        If AccNum.Trim.Length <> 13 Then
            'HandleError(&H80004007, "FormatAccount:Wrong Length:" & AccNum)
            Return AccNum
        End If
        RetStr = RetStr.Substring(0, 4) & "-" & RetStr.Substring(4, 6) & "-" & RetStr.Substring(10)
        Return RetStr
    End Function


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Format an amount based on the supplied currency
    ''' </summary>
    ''' <param name="CcyAmt">Amount to be formated</param>
    ''' <param name="CcyCde">Currency of the amount (3 letter code)</param>
    ''' <returns>Formatted amount</returns>
    ''' <remarks>
    ''' The currency code supplied determines the format of the amount (number of decimal places)
    ''' </remarks>
    ''' <history>
    ''' 	[u910krzi]	31-Mar-2009	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Function FormatCcyAmt(ByVal CcyAmt As String, ByVal CcyCde As String) As String
        If (CcyAmt Is Nothing) OrElse (CcyAmt.Trim = "") Then Return ""
        If (CcyCde Is Nothing) OrElse (CcyCde.Trim = "") Then Return ""
        Dim l_Assm As [Assembly] = [Assembly].GetEntryAssembly
        Try
            Dim modtyp As Type = l_Assm.GetType("ICEP0100.ICEDataBase")
            If modtyp Is Nothing Then Throw New VersionNotFoundException
            Dim rmi As MethodInfo = modtyp.GetMethod("FormatCcyAmt")
            If rmi Is Nothing Then Throw New VersionNotFoundException
            Return rmi.Invoke(Nothing, New Object() {CcyAmt, CcyCde}).ToString
        Catch ex As VersionNotFoundException
            Return FormatValString(CcyAmt, 2)
        End Try
    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Format an amount based on the supplied currency and appends a sign to it.
    ''' </summary>
    ''' <param name="CcyAmt">Signed amount to be formated</param>
    ''' <param name="CcyCde">Currency of the amount (3 letter code)</param>
    ''' <returns>Formatted amount</returns>
    ''' <remarks>
    ''' The currency code supplied determines the format of the amount (number of decimal places)
    ''' </remarks>
    ''' <history>
    ''' 	[u910krzi]	31-Mar-2009	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Function FormatCcyAmtNS(ByVal CcyAmt As String, ByVal CcyCde As String) As String
        If (CcyAmt Is Nothing) OrElse (CcyAmt.Trim = "") Then Return ""
        If (CcyCde Is Nothing) OrElse (CcyCde.Trim = "") Then Return ""
        Dim sSign As String
        sSign = CcyAmt.Substring(CcyAmt.Length - 1)
        Return FormatCcyAmt(CcyAmt.Substring(0, CcyAmt.Length - 1), CcyCde) & sSign
    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Format an amount based on the supplied currency and prefix it with a CR/DR.
    ''' </summary>
    ''' <param name="CcyAmt">Signed amount to be formated</param>
    ''' <param name="CcyCde">Currency of the amount (3 letter code)</param>
    ''' <returns>Formatted amount</returns>
    ''' <remarks>
    ''' The currency code supplied determines the format of the amount (number of decimal places)
    ''' </remarks>
    ''' <history>
    ''' 	[u910krzi]	31-Mar-2009	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Function FormatCcyAmtCRDR(ByVal CcyAmt As String, ByVal CcyCde As String) As String
        If (CcyAmt Is Nothing) OrElse (CcyAmt.Trim = "") Then Return ""
        If (CcyCde Is Nothing) OrElse (CcyCde.Trim = "") Then Return ""
        Dim sSign As String
        sSign = CcyAmt.Substring(CcyAmt.Length - 1)
        If (sSign = "-") Then
            sSign = " DR"
        Else
            sSign = " CR"
        End If
        Return FormatCcyAmt(CcyAmt.Substring(0, CcyAmt.Length - 1), CcyCde) & sSign
    End Function

    Function formatSlipAmount(ByVal p_number As String, ByVal p_Ccy As String) As String
        Dim sVal As String
        sVal = ICEI0100.AppInstanceClass.DefInstance.FormatCcyAmt(p_number, p_Ccy)
        sVal = sVal.Replace("+", String.Empty)
        sVal = sVal.Replace("-", String.Empty)

        'Return SIBL0100.Util.UString.repeat(15 - sVal.Length, "*") & sVal
        Return sVal.PadLeft(15, CChar("*"))
    End Function

    Function formatSlipAmount(ByVal p_number As String) As String
        Dim sVal As String
        If p_number Is Nothing Then
            p_number = String.Empty
        End If
        sVal = p_number
        'Return SIBL0100.Util.UString.repeat(15 - sVal.Length, "*") & sVal
        Return sVal.PadLeft(15, CChar("*"))
    End Function

    Public Function ToShortDate(ByVal InStr As String) As String
        Dim dt As Date
        Dim RetStr As String
        RetStr = InStr

        Dim OldCultInfo, NewCultInfo As CultureInfo
        OldCultInfo = System.Threading.Thread.CurrentThread.CurrentCulture
        NewCultInfo = New CultureInfo("en-US", False)
        System.Threading.Thread.CurrentThread.CurrentCulture = NewCultInfo

        If (RetStr Is Nothing) Then Return Nothing
        Try
            dt = InStr
            RetStr = IceFormatDate(dt, "ddMMMyy")
        Catch ex As Exception
            HandleError(&H80004008, "ToShortDate:Could not format date:" & ex.Message)
            RetStr = "Error! " & ex.Message
        End Try
        System.Threading.Thread.CurrentThread.CurrentCulture = OldCultInfo
        NewCultInfo = Nothing
        GC.Collect()
        Return RetStr
    End Function

    Public Function PackDate(ByVal Dte As Date) As String
        Dim RetStr As String
        'Dim OldCultInfo, NewCultInfo As CultureInfo
        'OldCultInfo = System.Threading.Thread.CurrentThread.CurrentCulture
        'NewCultInfo = New CultureInfo("en-US", False)
        'System.Threading.Thread.CurrentThread.CurrentCulture = NewCultInfo

        Try
            RetStr = ""
            RetStr += Trim(Str(Dte.Year))
            RetStr += Format(Dte.Month, "#00")
            RetStr += Format(Dte.Day, "#00")
        Catch
            RetStr = "00000000"
        End Try
        'System.Threading.Thread.CurrentThread.CurrentCulture = OldCultInfo
        'NewCultInfo = Nothing
        Return RetStr
    End Function

    Public Function PackDate(ByVal DateStr As String) As String
        Dim dte As Date
        Dim RetStr As String
        Dim OldCultInfo, NewCultInfo As CultureInfo
        OldCultInfo = System.Threading.Thread.CurrentThread.CurrentCulture
        NewCultInfo = New CultureInfo("en-US", False)
        System.Threading.Thread.CurrentThread.CurrentCulture = NewCultInfo

        If (DateStr Is Nothing) OrElse (DateStr.Trim = "") Then
            RetStr = "00000000"
        Else
            Try
                dte = DateStr
                RetStr = ""
                RetStr += Trim(Str(dte.Year))
                RetStr += Format(dte.Month, "#00")
                RetStr += Format(dte.Day, "#00")
            Catch
                RetStr = "00000000"
            End Try
        End If
        System.Threading.Thread.CurrentThread.CurrentCulture = OldCultInfo
        NewCultInfo = Nothing
        Return RetStr
    End Function

    Public Function PackTime(ByVal TimeStr As String) As String
        Dim dte As Date
        Dim RetStr As String
        Dim OldCultInfo, NewCultInfo As CultureInfo
        OldCultInfo = System.Threading.Thread.CurrentThread.CurrentCulture
        NewCultInfo = New CultureInfo("en-US", False)
        System.Threading.Thread.CurrentThread.CurrentCulture = NewCultInfo

        If (TimeStr Is Nothing) OrElse (TimeStr.Trim = "") Then
            RetStr = "000000"
        Else
            Try
                dte = "06-06-06 " & TimeStr
                RetStr = ""
                RetStr += Format(dte.Hour, "#00")
                RetStr += Format(dte.Minute, "#00")
                RetStr += Format(dte.Second, "#00")
            Catch
                RetStr = "000000"
            End Try
        End If
        System.Threading.Thread.CurrentThread.CurrentCulture = OldCultInfo
        NewCultInfo = Nothing
        GC.Collect()
        Return RetStr
    End Function

    Public Function PackDateTime(ByVal p_DateStr As String, ByVal p_TimeStr As String) As String
        Return PackDate(p_DateStr) & PackTime(p_TimeStr)
    End Function

    Public Function PackDateTime(ByVal p_Date As Date) As String
        Dim RetStr As String
        Dim OldCultInfo, NewCultInfo As CultureInfo
        OldCultInfo = System.Threading.Thread.CurrentThread.CurrentCulture
        NewCultInfo = New CultureInfo("en-US", False)
        System.Threading.Thread.CurrentThread.CurrentCulture = NewCultInfo
        RetStr = PackDate(Format(p_Date, "MM/dd/yyyy")) & PackTime(Format(p_Date, "HH:mm:ss"))
        System.Threading.Thread.CurrentThread.CurrentCulture = OldCultInfo
        NewCultInfo = Nothing
        GC.Collect()
        Return RetStr
    End Function

    Public Function PackDateTime(ByVal p_DateStr As String) As String
        'Dim RetStr As String

        If (p_DateStr Is Nothing) OrElse (p_DateStr.Trim = "") Then
            Return "000000000000"
        End If
        If p_DateStr.Trim.Length <= 6 Then
            Return PackDate(p_DateStr) & "000000"
        End If

        Return PackDate(p_DateStr) & PackTime(p_DateStr)
    End Function

    'As Requested by ziad to move this to sibl0100
    'Public Function Str2Date(ByVal DateStr As String) As Date

    '    Dim dte As Date
    '    Dim OldCultInfo, NewCultInfo As CultureInfo
    '    OldCultInfo = System.Threading.Thread.CurrentThread.CurrentCulture
    '    NewCultInfo = New CultureInfo("en-US", False)
    '    System.Threading.Thread.CurrentThread.CurrentCulture = NewCultInfo

    '    dte = DateStr

    '    System.Threading.Thread.CurrentThread.CurrentCulture = OldCultInfo

    '    Return dte
    'End Function

    Public Function SafeSubString(ByVal Src As String, ByVal BegIdx As Integer, ByVal lenStr As Integer) As String
        If Src Is Nothing Then Return String.Empty
        If lenStr <= 0 Then Return Src
        If Src.Length < BegIdx Then Return String.Empty
        If Src.Length < BegIdx + lenStr Then Return Src.Substring(BegIdx)
        Return Src.Substring(BegIdx, lenStr)
    End Function

    Public Function SafeSubString(ByVal Src As String, ByVal BegIdx As Integer) As String
        If Src Is Nothing Then Return String.Empty
        If Src.Length < BegIdx Then Return String.Empty
        Return Src.Substring(BegIdx)
    End Function


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' A Function to get a part of the string starting from the left specified by BegIdx and for the length of lenStr. BegIdx with value zero specifies the last character in the string.
    ''' </summary>
    ''' <returns>A sub string</returns>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[u910krzi/y093sahu]	12-Mar-2011	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Function LeftString(ByVal Src As String, ByVal BegIdx As Integer, ByVal lenStr As Integer) As String
        If Src Is Nothing Then Return String.Empty
        If lenStr <= 0 Then Return Src
        If Src.Length < BegIdx Then Return String.Empty
        If Src.Length < BegIdx + lenStr Then Return Src.Substring(0, BegIdx)
        'Return Src.Substring(Src.Length - lenStr - BegIdx + 1, lenStr)
        Dim ch() As Char = Src.ToCharArray
        Array.Reverse(ch)
        Dim st As String = CStr(ch)
        st = st.Substring(BegIdx, lenStr)
        ch = st.ToCharArray
        Array.Reverse(ch)
        Return CStr(ch)
    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' A Function to perform an Extended Int, i.e., it will transform "9,999.99" to 10000, "9,999.2" to 9999
    ''' </summary>
    ''' <param name="ValStr">A string value to be converted to int</param>
    ''' <returns>A rounded integer</returns>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[u910krzi]	04-Apr-2009	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Function ExInt(ByVal ValStr As String) As Integer
        If (ValStr Is Nothing) Then Return 0
        Try
            Return CInt(CDbl(ValStr))
        Catch
            Return CInt(Val(ValStr))
        End Try
    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Function to perform an Extended Val, i.e., it will transform "9,999.99" to 9999.99
    ''' </summary>
    ''' <param name="ValStr">A string value to be converted to double</param>
    ''' <returns></returns>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[u910krzi]	04-Apr-2009	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Function ExVal(ByVal ValStr As String) As Double
        Dim st As String
        Try
            If ValStr Is Nothing Then Return 0.0
            If Trim(ValStr) = "" Then Return 0.0

            st = UCase(ValStr)
            If st.IndexOf("CR") >= 0 Then
                st = st.Replace("CR", "")
                st = "+" & st
            End If
            If st.IndexOf("DB") >= 0 Then
                st = st.Replace("DB", "")
                st = "-" & st
            ElseIf st.IndexOf("DR") >= 0 Then
                st = st.Replace("DR", "")
                st = "-" & st
            End If
            Return CDbl(st)
        Catch
            Return 0.0
        End Try
    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Displays a confirmation message to perform host data update. This is to standarize the message across all ICE windows.
    ''' </summary>
    ''' <returns>True if the user clicks the "Yes" button, False otherwise</returns>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[u910krzi]	04-Apr-2009	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Function ConfirmHostUpdate() As Boolean
        Dim dlgRes As DialogResult
        dlgRes = MessageBox.Show("About to send the update changes to the host." & vbCrLf & _
                "Do you want to continue?", "ICE Host Update", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2)
        Return (dlgRes = DialogResult.Yes)
    End Function

    Public Function SimahQuery(ByVal number As String, ByVal type As String, ByRef score As String, ByRef reference As String) As Boolean
        Dim result As Boolean = False

        Try
            If (Not IsNothing(number)) AndAlso (number.Trim <> "") Then
                If (Not IsNothing(type)) AndAlso (type.Trim <> "") Then
                    Dim entryAssembly As [Assembly] = [Assembly].GetEntryAssembly
                    Dim methodClass As Type = entryAssembly.GetType("ICEP0100.SIMAH.clsSimah")
                    If methodClass Is Nothing Then Throw New VersionNotFoundException

                    Dim methodInfo As MethodInfo = methodClass.GetMethod("GetQueryResults")
                    If methodInfo Is Nothing Then Throw New VersionNotFoundException

                    Dim params() As Object = New Object() {number, type, score, reference}
                    Dim response As Object = methodInfo.Invoke(Nothing, params)
                    If Not (response Is Nothing) Then
                        result = CBool(response)
                        score = params(2)
                        reference = params(3)
                    End If
                End If
            End If
        Catch ex As VersionNotFoundException
            HandleError(&H80009099, "SmahQuery() - exception:" & ex.Message)
        End Try

        Return result
    End Function

    ''' Process the ShowStatusText text-clear timer. When timer fires, the text will be cleared.
    Private Sub TimerEventProcessor(ByVal myObject As Object, ByVal myEventArgs As EventArgs)
        tmrStatusText.Stop()
        ShowStatusText("")
        gfrmMain.statusPanel_Text.Style = IIf(Me.gbPleaseLogOn, StatusBarPanelStyle.OwnerDraw, StatusBarPanelStyle.Text)
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Function to disable a textbox according to ICE standard
    ''' </summary>
    ''' <param name="txt">Textbox control to be disabled</param>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[u910krzi]	04-Apr-2009	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Sub DisableText(ByVal txt As TextBox)
        txt.BackColor = SystemColors.Control
        txt.Enabled = False
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Function to enable a textbox according to the ICE standard.
    ''' </summary>
    ''' <param name="txt">Textbox control to be enabled</param>
    ''' <param name="bReadOnly">Enable as readonly or read-write</param>
    ''' <remarks>
    ''' If the control is in readonly mode, then it will be colored with a yellow-tip background.
    ''' </remarks>
    ''' <history>
    ''' 	[u910krzi]	04-Apr-2009	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Sub EnableText(ByVal txt As TextBox, Optional ByVal bReadOnly As Boolean = True)
        txt.Enabled = True
        If bReadOnly Then
            txt.BackColor = Color.FromArgb(255, 255, 192)
        Else
            txt.BackColor = SystemColors.Window
        End If
        txt.ReadOnly = bReadOnly
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Removes extra spaces in a string (leading spaces, trailing spaces, and inner double spaces)
    ''' </summary>
    ''' <param name="FullString">String to be compacted</param>
    ''' <returns>Compacted string</returns>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[u910krzi]	04-Apr-2009	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Function CompactString(ByVal FullString As String) As String
        Dim stSrc, stDst As String
        Dim src(), dst() As Char
        Dim i, idx As Integer
        Dim sp As Char = Chr(Keys.Space)

        If (FullString Is Nothing) Then Return ""
        stSrc = Trim(FullString)
        If (stSrc.Length < 1) Then Return ""

        stDst = Space(stSrc.Length)
        src = stSrc.ToCharArray
        dst = stDst.ToCharArray

        idx = 1
        dst(0) = src(0)
        For i = 1 To src.Length - 1
            If ((src(i) <> sp) Or (src(i) = sp And src(i - 1) <> sp)) Then
                dst(idx) = src(i)
                idx += 1
            End If
        Next

        stDst = dst
        Return Trim(stDst)
    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' A function that compact miltuple lines by removing empty lines from between lines
    ''' </summary>
    ''' <param name="FullLines">A string array that contains the lines to be compacted</param>
    ''' <param name="Separator">The line seprator string or charachter. Default is vbCrLF</param>
    ''' <returns>A string containing all the lines minus the empty lines, speparted by the supplied Separator</returns>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[u910krzi]	04-Apr-2009	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Function CompactLines(ByRef FullLines() As String, Optional ByVal Separator As String = vbCrLf) As String
        Dim CompLines As String = ""
        Dim st As String
        If FullLines Is Nothing Then Return False
        If FullLines.Length = 0 Then Return False
        For i As Integer = 0 To FullLines.Length - 1
            st = FullLines(i)
            If (Not (st Is Nothing)) Then
                st = st.Replace(Chr(9), " ")
                If Trim(st) <> "" Then
                    If CompLines.Length > 0 Then
                        CompLines = CompLines & Separator & FullLines(i)
                    Else
                        CompLines = FullLines(i)
                    End If
                End If
            End If
        Next
        Return CompLines
    End Function

    '* Formats the time part as requested in pFormatString
    Public Function IceFormatTime(ByVal pDtpDate As Date, ByVal pFormatString As String) As String
        'Dim MyResult As String
        Try
            '* Use time part only
            IceFormatTime = Format(TimeSerial(Hour(pDtpDate), Minute(pDtpDate), Second(pDtpDate)), pFormatString)
        Catch ex As Exception
            HandleError(&H80004009, "ICE generated an error while formatting time! " & ex.Message)
            IceFormatTime = ""
        End Try
        Exit Function

    End Function

    '* Formats the date part as requested in pFormatString
    Public Function IceFormatDate(ByVal pDtpDate As DateTime, ByVal pFormatString As String) As String
        Dim MyResult As String = String.Empty

        '* Use English US locale
        Try
            IceFormatDate = FormatDateForLocale(DefEngUs, pDtpDate, pFormatString, MyResult)

        Catch ex As Exception
            HandleError(&H80004010, "ICE generated an error while formatting date!" & ex.Message)
            IceFormatDate = ""
        End Try

    End Function

    '* Formats the date part as requested in pFormatString
    Public Function IceFormatDate(ByVal pDtpDate As String, ByVal pFormatString As String) As String
        Dim MyResult As String
        '* Use English US locale
        Dim OldCultInfo, NewCultInfo As CultureInfo
        OldCultInfo = System.Threading.Thread.CurrentThread.CurrentCulture
        NewCultInfo = New CultureInfo("en-US", False)
        System.Threading.Thread.CurrentThread.CurrentCulture = NewCultInfo
        Try
            MyResult = IceFormatDate(CDate(pDtpDate), pFormatString)
        Catch ex As Exception
            MyResult = ""
        End Try
        System.Threading.Thread.CurrentThread.CurrentCulture = OldCultInfo
        NewCultInfo = Nothing
        Return MyResult
    End Function


    '    If (RetStr Is Nothing) Then Return Nothing
    '    Try
    '        dt = InStr
    '        RetStr = IceFormatDate(dt, "ddMMMyy")
    '    Catch ex As Exception
    '        HandleError(&H80004008, "ToShortDate:Could not format date:" & ex.Message)
    '        RetStr = "Error! " & ex.Message
    '    End Try


    'd      : Day of month as digits with no leading zero for single-digit days.
    'dd     : Day of month as digits with leading zero for single-digit days.
    'ddd    : Day of week as a three-letter abbreviation. The function uses the LOCALE_SABBREVDAYNAME value associated with the specified locale.
    'dddd   : Day of week as its full name. The function uses the LOCALE_SDAYNAME value associated with the specified locale.
    'M      : Month as digits with no leading zero for single-digit months.
    'MM     : Month as digits with leading zero for single-digit months.
    'MMM    : Month as a three-letter abbreviation. The function uses the LOCALE_SABBREVMONTHNAME value associated with the specified locale.
    'MMMM   : Month as its full name. The function uses the LOCALE_SMONTHNAME value associated with the specified locale.
    'y      : Year as last two digits, but with no leading zero for years less than 10.
    'yy     : Year as last two digits, but with leading zero for years less than 10.
    'yyyy   : Year represented by full four digits.
    'gg     : Period/era string. The function uses the CAL_SERASTRING value associated with the specified locale.
    '         This element is ignored if the date to be formatted does not have an associated era or period string.


    '*** *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** ***
    Public Function FormatDateForLocale(ByVal TheLocale As Integer, ByVal TheDate As DateTime, ByVal TheFormat As String, ByVal TheResult As String) As String
        '*** *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** ***
        Dim RetCode As Long
        Dim ResultStr As String
        Dim MyDate As SYSTEMTIME
        'dim muCal as 
        Dim Ret As String = String.Empty

        Try

            With MyDate
                'Microsoft.VisualBasic. 
                'Dim x As Object
                '                System.Globalization.DateTimeFormatInfo.CurrentInfo.Calendar = System.Globalization.GregorianCalendar.CurrentEra
                'dt = DateSerial(CInt(DteStr.Substring(0, 4)), CInt(DteStr.Substring(4, 2)), CInt(DteStr.Substring(6, 2)))

                'Dim myDT As New DateTime(Year(TheDate), Month(TheDate), Microsoft.VisualBasic.DateAndTime.Day(TheDate), New System.Globalization.GregorianCalendar)
                'Dim mtim As New DateTime(DateSerial(Year(TheDate)), DateSerial(Month(TheDate)), DateSerial(Microsoft.VisualBasic.DateAndTime.Day(TheDate))
                .wDay = TheDate.Day
                .wMonth = TheDate.Month
                .wYear = TheDate.Year

                '.wDay = Microsoft.VisualBasic.DateAndTime.Day(TheDate)
                '.wMonth = Microsoft.VisualBasic.DateAndTime.Month(TheDate)
                '.wYear = Microsoft.VisualBasic.DateAndTime.Year(TheDate)

                .wHour = 0
                .wMinute = 0
                .wSecond = 0
                .wMilliseconds = 0
            End With

            TheFormat = Replace(TheFormat, "m", "M")
            TheFormat = Replace(TheFormat, "D", "d")
            TheFormat = Replace(TheFormat, "Y", "y")
            TheFormat = Replace(TheFormat, "G", "g")

            ResultStr = Space(100)
            RetCode = GetDateFormat(TheLocale, 0, MyDate, TheFormat, ResultStr, 50)
            Ret = Left(ResultStr, RetCode - 1)


        Catch ex As Exception
            HandleError(&H80004011, "ICE generated an error while formatting date!" & ex.Message)
        End Try
        Return Ret
    End Function

    Public Sub ShowHostError(ByVal MsgCde As String, ByVal MsgID As String, ByVal ResCde As String, ByVal ActCde As String)
        'Refelection: This method is called from ICEP0100
        Dim l_Assm As [Assembly] = [Assembly].GetEntryAssembly
        Dim modtyp As Type = l_Assm.GetType("ICEP0100.ICEComm")
        Dim rmi As MethodInfo = modtyp.GetMethod("ShowHostError_Core")
        rmi.Invoke(Nothing, New Object() {MsgCde, MsgID, ResCde, ActCde})
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Function that will check the current user authority list against the supplied bit numer
    ''' </summary>
    ''' <param name="BitNum">The bit number required (possible values are from 0 to 511)</param>
    ''' <returns>True if the user has the bit set to true in his role, otheriwse, it returns false</returns>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[u910krzi]	04-Apr-2009	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Function CheckUserAuthority(ByVal BitNum As Integer) As Boolean
        Dim RetObj As Object
        Dim RetVal As Boolean = False
        Dim l_Assm As [Assembly] = [Assembly].GetEntryAssembly
        Dim modtyp As Type = l_Assm.GetType("ICEP0100.ICEComm")
        Dim rmi As MethodInfo = modtyp.GetMethod("CheckUserAuthority_Core")
        RetObj = rmi.Invoke(Nothing, New Object() {BitNum})
        If Not RetObj Is Nothing Then
            RetVal = CBool(RetObj)
        End If
        Return RetVal
    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Function that will validate the MID returned from a SendMessage using the MQ system
    ''' </summary>
    ''' <param name="MsgID">The Message ID to be checked</param>
    ''' <returns>If the MID is valid, then TRUE is returned, otheriwse, FALSE is returned</returns>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[u910krzi]	31-May-2009	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Function IsValidMid(ByVal MsgID As String) As Boolean
        Dim RetObj As Object
        Dim RetVal As Boolean = False
        Dim l_Assm As [Assembly] = [Assembly].GetEntryAssembly
        Dim modtyp As Type = l_Assm.GetType("ICEP0100.ICEComm")
        Dim rmi As MethodInfo = modtyp.GetMethod("IsValidMid")
        RetObj = rmi.Invoke(Nothing, New Object() {MsgID})
        If Not RetObj Is Nothing Then
            RetVal = CBool(RetObj)
        End If
        Return RetVal
    End Function


#Region "Form State Preservation"
    Public Function LoadLastPosition(ByVal FormID As String) As Point
        Dim pp As Point
        Dim ini As INIClass.IniFileClass
        pp.X = 0
        pp.Y = 0
        Try
            ini = New INIClass.IniFileClass(ICEI0100.IcePaths.DefInstance.IceUsrIni)
            If Not (ini.FileExist) Then
                HandleError(&H80004012, "Could not locate IceUsr.ini! Form position will be defaulted.")
                Return pp
            End If
            pp.X = CInt(ini.GetValue("XPos", FormID, "0"))
            pp.Y = CInt(ini.GetValue("YPos", FormID, "0"))
        Catch ex As Exception
            HandleError(&H80004013, "Could not Load Form position: " & ex.Message, "ICE Config")
        End Try
        Return pp
    End Function

    Public Function LoadLastPosition(ByVal FormID As String, ByRef state As FormWindowState, Optional ByVal bClip As Boolean = True) As Rectangle
        Dim rec As Rectangle
        Dim ini As INIClass.IniFileClass
        rec.X = 0
        rec.Y = 0
        rec.Width = 0
        rec.Height = 0
        Try
            ini = New INIClass.IniFileClass(ICEI0100.IcePaths.DefInstance.IceUsrIni)
            If Not (ini.FileExist) Then
                HandleError(&H80004014, "Could not locate IceUsr.ini! Form position will be defaulted.")
                Return rec
            End If
            rec.X = CInt(ini.GetValue("XPos", FormID, "0"))
            rec.Y = CInt(ini.GetValue("YPos", FormID, "0"))
            rec.Width = CInt(ini.GetValue("Width", FormID, "0"))
            rec.Height = CInt(ini.GetValue("Height", FormID, "0"))
            state = CInt(ini.GetValue("State", FormID, FormWindowState.Normal))
            If state = FormWindowState.Minimized Then
                'we will not allow loading of forms in minized state
                state = FormWindowState.Normal
            End If
            If (bClip And (FormID <> "FORM_MAIN")) Then
                Dim ClientRec As Rectangle
                ClientRec = gfrmMain.ClientRectangle
                If rec.X > ClientRec.Width Then rec.X = 0
                If rec.Y > ClientRec.Height Then rec.Y = 0
            Else
                If (rec.X < 0) Or (rec.X > Screen.PrimaryScreen.WorkingArea.Width) Then rec.X = 0
                If (rec.Y < 0) Or (rec.Y > Screen.PrimaryScreen.WorkingArea.Height) Then rec.Y = 0
            End If
        Catch ex As Exception
            HandleError(&H80004015, "Could not Load Form position: " & ex.Message, "ICE Config")
        End Try
        Return rec
    End Function

    Public Sub LoadLastPosition(ByVal p_form As System.Windows.Forms.Form, Optional ByVal p_form_name As String = "")
        'Dim rec As Rectangle
        Dim st As FormWindowState

        Dim stFormName As String = p_form_name
        If (stFormName Is Nothing) OrElse (stFormName.Trim = "") Then stFormName = p_form.Name

        If p_form Is Nothing Then Return
        Dim pp As Rectangle = LoadLastPosition(stFormName, st, (Not p_form.MdiParent Is Nothing))
        If pp.Height = 0 Then
            p_form.StartPosition = CType(IIf(p_form.Parent Is Nothing, FormStartPosition.CenterScreen, FormStartPosition.CenterParent), FormStartPosition)
            Return
        End If
        If st = FormWindowState.Normal Then
            p_form.Location = New Point(pp.X, pp.Y)
            Dim sz As New Size(pp.Width, pp.Height)
            If ((sz.Width > 0) And (sz.Height > 0)) Then
                p_form.Size = sz
            End If
        Else
            p_form.WindowState = st
        End If

    End Sub

    Public Function SaveLastPosition(ByVal p_form As System.Windows.Forms.Form, Optional ByVal p_form_name As String = "") As Boolean
        If p_form Is Nothing Then Return False
        Dim stFormName As String = p_form_name
        If (stFormName Is Nothing) OrElse (stFormName.Trim = "") Then stFormName = p_form.Name
        Return SaveLastPosition(stFormName, p_form.WindowState, New Rectangle(p_form.Location, p_form.Size))
    End Function

    Public Function SaveLastPosition(ByVal FormID As String, ByVal LastPos As Point) As Boolean
        Dim ini As INIClass.IniFileClass
        Try
            ini = New INIClass.IniFileClass(ICEI0100.IcePaths.DefInstance.IceUsrIni)
            If Not (ini.FileExist) Then
                HandleError(&H80004016, "Could not locate IceUsr.ini! Form position will not be recorded.")
                Return False
            End If
            ini.SetValue("XPos", LastPos.X.ToString, FormID)
            ini.SetValue("YPos", LastPos.Y.ToString, FormID)
        Catch ex As Exception
            HandleError(&H80004017, "Could not Save Form position: " & ex.Message, "ICE Config")
            Return False
        End Try
        Return True
    End Function

    Public Function SaveLastPosition(ByVal FormID As String, ByVal State As FormWindowState, ByVal LastPos As Rectangle) As Boolean
        Dim ini As INIClass.IniFileClass
        Try
            If State = FormWindowState.Minimized Then
                Logger.LogInfo(0, "Form " & FormID & " is minimized; therefore, its state will not be recorded.", 2)
                Exit Try
            End If
            ini = New INIClass.IniFileClass(ICEI0100.IcePaths.DefInstance.IceUsrIni)
            If Not (ini.FileExist) Then
                HandleError(&H80004018, "Could not locate IceUsr.ini! Form position will not be recorded.")
                Return False
            End If
            If Not (State = FormWindowState.Maximized) Then
                ini.SetValue("XPos", LastPos.X.ToString, FormID)
                ini.SetValue("YPos", LastPos.Y.ToString, FormID)
                ini.SetValue("Width", LastPos.Width.ToString, FormID)
                ini.SetValue("Height", LastPos.Height.ToString, FormID)
            Else
                'If state is maximized, then the dimensions will not be recorded as the restoration dimensions would be lost.
            End If
            ini.SetValue("State", State, FormID)
        Catch ex As Exception
            HandleError(&H80004019, "Could not Save Form position: " & ex.Message, "ICE Config")
            Return False
        End Try
        Return True
    End Function

    Public Function SaveLastPosition(ByVal FormID As String, ByVal LastXPos As Integer, ByVal LastYPos As Integer) As Boolean
        Dim pp As Point
        pp.X = LastXPos
        pp.Y = LastYPos
        Return SaveLastPosition(FormID, pp)
    End Function
#End Region


    Public Class clsPrintForm
        Private Shared m_Assm As [Assembly] = [Assembly].GetEntryAssembly
        Private Shared m_FrmPrintType As Type = m_Assm.GetType("ICEP0100.frmPrint")
        Private Shared m_Status As PropertyInfo = m_FrmPrintType.GetProperty("Status")
        Private Shared m_Printer As PropertyInfo = m_FrmPrintType.GetProperty("Printer")
        Private Shared m_FullName As PropertyInfo = m_FrmPrintType.GetProperty("FullName")
        Private Shared m_Title As PropertyInfo = m_FrmPrintType.GetProperty("Title")
        Private Shared m_PrintingQueued As PropertyInfo = m_FrmPrintType.GetProperty("PrintingQueued")

        Public Shared WriteOnly Property Status() As String
            Set(ByVal Value As String)
                If DefInstance.gfrmPrint Is Nothing Then Return
                m_Status.SetValue(DefInstance.gfrmPrint, Value, Nothing)
            End Set
        End Property

        Public Shared WriteOnly Property Printer() As String
            Set(ByVal Value As String)
                If DefInstance.gfrmPrint Is Nothing Then Return
                m_Printer.SetValue(DefInstance.gfrmPrint, Value, Nothing)
            End Set
        End Property

        Public Shared WriteOnly Property FullName() As String
            Set(ByVal Value As String)
                If DefInstance.gfrmPrint Is Nothing Then Return
                m_FullName.SetValue(DefInstance.gfrmPrint, Value, Nothing)
            End Set
        End Property

        Public Shared WriteOnly Property Title() As String
            Set(ByVal Value As String)
                If DefInstance.gfrmPrint Is Nothing Then Return
                m_Title.SetValue(DefInstance.gfrmPrint, Value, Nothing)
            End Set
        End Property

        Public Shared ReadOnly Property PrintingQueued() As String
            Get
                If DefInstance.gfrmPrint Is Nothing Then Return ""
                Return CStr(m_PrintingQueued.GetValue(DefInstance.gfrmPrint, Nothing))
            End Get
        End Property

        Public Shared Sub ShowPrintForm()
            Dim modtyp As Type = m_Assm.GetType("ICEP0100.Data")
            Dim rmi As MethodInfo = modtyp.GetMethod("ShowPrintForm")
            rmi.Invoke(Nothing, Nothing)
        End Sub

        Public Shared Sub CreatePrintForm()
            Dim modtyp As Type = m_Assm.GetType("ICEP0100.Data")
            Dim rmi As MethodInfo = modtyp.GetMethod("CreatePrintForm")
            rmi.Invoke(Nothing, Nothing)
        End Sub

        Public Shared Sub ClosePrintForm()
            Dim modtyp As Type = m_Assm.GetType("ICEP0100.Data")
            Dim rmi As MethodInfo = modtyp.GetMethod("ClosePrintForm")
            rmi.Invoke(Nothing, Nothing)
        End Sub

        Public Shared Sub ShowPrintFormEx(ByRef PrnObj As Object)
            Dim modtyp As Type = m_Assm.GetType("ICEP0100.Data")
            Dim rmi As MethodInfo = modtyp.GetMethod("ShowPrintFormEx")
            rmi.Invoke(Nothing, New Object() {PrnObj})
        End Sub

        Public Shared Sub HidePrintForm()
            Dim modtyp As Type = m_Assm.GetType("ICEP0100.Data")
            Dim rmi As MethodInfo = modtyp.GetMethod("HidePrintForm")
            rmi.Invoke(Nothing, Nothing)
        End Sub
    End Class


#Region "MQS Messages - Visual"

    Private MQBusy As Integer = 0

    Public Function GetNextMsgID() As String
        Dim MsgID As String = ""
        Try
            Dim NumEngVal As Integer
            ' IsLoggedOn = True
            If Not IsLoggedOn Then
                IceLogOn()
                If Not IsLoggedOn Then
                    ModalMsgBox("The last operation failed because it requires a logged in user." & vbCrLf & _
                           "Please retry this operation and login when prompted.", MsgBoxStyle.Exclamation, "ICE Login required")
                    MsgID = Nothing
                    Exit Try
                End If
            End If
            'ShowBusyIcon(True)
            NumEngVal = NumEng.NumEngGiv()
            If NumEngVal = 0 Then
                HandleError(&H81000021, "Error in Numbering Engine: " & NumEng.ErrDsc, "ICE Critical Error")
                MsgID = ""
                Exit Try
            End If
            Dim stWksNam As String = Environment.MachineName
            If stWksNam.Length > 8 Then stWksNam = stWksNam.Substring(stWksNam.Length - 8)
            stWksNam = PackString(stWksNam, 8)
            MsgID = stWksNam & NumEng.ConvertToBase64(NumEngVal)
        Catch ex As Exception
            HandleError(&H81000140, "Error in GetNextMsgID: " & ex.Message, "Prepare MQSeries Message")
            MsgID = Nothing
        End Try
        Return MsgID
    End Function

    Public Function GetMessage_Common(ByVal MsgCde As String, ByVal MsgNam As String, ByVal MsgID As String, ByVal ErrNum As Integer, ByRef resp As String, _
                                  Optional ByVal bShowAlert As Boolean = True, Optional ByVal BypassError As Boolean = False) As Boolean
        'Refelection: This method is called from ICEP0100
        Dim RetObj As Object
        Dim RetVal As Boolean = False
        Dim l_Assm As [Assembly] = [Assembly].GetEntryAssembly
        Dim modtyp As Type = l_Assm.GetType("ICEP0100.ICEComm")
        Dim rmi As MethodInfo = modtyp.GetMethod("GetMessage_Common")
        Dim PrmArr() As Object
        PrmArr = New Object() {MsgCde, MsgNam, MsgID, ErrNum, resp, bShowAlert, BypassError}
        RetObj = rmi.Invoke(Nothing, PrmArr)
        If Not (RetObj Is Nothing) Then
            RetVal = CBool(RetObj)
            resp = PrmArr(4)
        End If
        Return RetVal
    End Function


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Function that prepares an MQ message to be sent, but does not send it, it
    ''' simply builds the message structure
    ''' </summary>
    ''' <param name="MsgCode"></param>
    ''' <param name="MsgData"></param>
    ''' <param name="MsgCls"></param>
    ''' <param name="CusNum"></param>
    ''' <param name="AccNum"></param>
    ''' <param name="CrdNum"></param>
    ''' <returns></returns>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[u910krzi]	03/03/2007	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Function PrepareMQMessage(ByVal MsgID As String, ByVal MsgCode As String, ByVal MsgData As String, ByVal MsgCls As String, _
                                        ByVal CusNum As String, ByVal AccNum As String, ByVal CrdNum As String, _
                                        ByVal IceTryFin As String, ByVal AutUsrIde As String, ByVal AutPacVal As String, _
                                        ByRef PreparedMsg As String, Optional ByVal AskMidRef As String = "") As String
        'Dim MsgId As String = Nothing
        Try
            Dim retVal As String
            If Not (IsValidMid(MsgID)) Then
                MsgID = GetNextMsgID()
            End If
            LastMsgID = MsgID

            retVal = MQS.PrepareMessage(MsgCode, MsgData, MsgID, MsgCls, CusNum, AccNum, CrdNum, "000", "00", "000", IceTryFin, AutUsrIde, AutPacVal, AskMidRef)
            If (retVal Is Nothing) Then
                Logger.LogError(&H81000022, "Message Preparing Failed! " & MQS.ErrNum & ": " & MQS.ErrDsc)
            End If
            MsgID = CStr(IIf(retVal Is Nothing, Nothing, MsgID))
            PreparedMsg = retVal
            'ShowBusyIcon()
        Catch ex As Exception
            HandleError(&H81000023, "Error in PrepareMQMessage: " & ex.Message, "Prepare MQSeries Message")
            MsgID = Nothing
        End Try
        Return MsgID
    End Function

    Public Function SendMessagePreparedVisually(ByVal MsgID As String, ByVal MsgCode As String, ByVal MsgBody As String) As String
        SyncLock GetType(AppInstanceClass) 'MQSyncObject
            Try
                Dim retVal As Boolean
                ShowActivityIcon(enumActivityType.Write)
                Dim st As String
                st = String.Format("Sending Message: Code({0}:{1}), ID({2}), Data({3})", MsgCode, gfrmMain.DecodeMsgCde(MsgCode), MsgID, MsgBody)
                Logger.LogInfo(&H80000031, st, 2)
                Dim nRetries As Integer = MQS.SendRetries
                retVal = False
                While ((nRetries > 0) And (Not retVal))
                    nRetries -= 1
                    retVal = MQS.PutMessagePrepared(MsgID, MsgBody)
                    Statistics.MessagesSent += 1
                    If (Not retVal) Then
                        Logger.LogError(&H81000024, "Message Sending Failed! [Attemp " & MQS.SendRetries - nRetries & "]: " & MQS.ErrNum & ": " & MQS.ErrDsc)
                        If (MQS.ErrNum = MQSeries.MQSeries.MQRC_NOT_OPEN) Then
                            Dim reconnect_res As Boolean
                            ShowStatusText("Attempting to reconnect...")
                            Logger.LogInfo(&H80000000, "Attempting to reconnect")
                            reconnect_res = MQS.ReConnect()
                            If reconnect_res Then
                                ShowStatusText("Reconnect Succeeded.")
                                Logger.LogInfo(&H80000100, "Reconnect Succeeded.")
                            Else
                                ShowStatusText("Reconnect Failed!")
                                Logger.LogError(&H80000100, "Reconnect Failed!")
                            End If
                        End If
                        Threading.Thread.Sleep(MQS.ResendDelay)
                    End If
                End While
                MsgID = CStr(IIf(retVal, MsgID, Nothing))
                ShowActivityIcon()
                'ShowBusyIcon()
            Catch ex As Exception
                HandleError(&H81000025, "Error in SendMessagePreparedVisually: " & ex.Message, "Send MQSeries Message")
                MsgID = ""
            End Try
            Return MsgID
        End SyncLock
    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Overloaded function to maintain compatibility with older code
    ''' </summary>
    ''' <param name="MsgCode"></param>
    ''' <param name="MsgData"></param>
    ''' <param name="MsgCls"></param>
    ''' <returns></returns>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[u910krzi]	03/03/2007	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Function SendMessageVisually(ByVal MsgCode As String, Optional ByVal MsgData As String = "", _
                                        Optional ByVal MsgCls As String = "GEN") As String
        Return SendMessageVisuallyEx(MsgCode, MsgData, MsgCls, "", "", "")
    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Overloaded function to maintain compatibility with older code
    ''' </summary>
    ''' <param name="MsgCode"></param>
    ''' <param name="MsgData"></param>
    ''' <param name="MsgCls"></param>
    ''' <param name="CusNum"></param>
    ''' <param name="AccNum"></param>
    ''' <param name="CrdNum"></param>
    ''' <returns></returns>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[u910krzi]	03/03/2007	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Function SendMessageVisuallyEx(ByVal MsgCode As String, ByVal MsgData As String, ByVal MsgCls As String, _
                                        ByVal CusNum As String, ByVal AccNum As String, ByVal CrdNum As String) As String
        Dim dummy As String = String.Empty
        ShowBusyIcon(True)
        Return SendMessageVisuallyEx(MsgCode, MsgData, MsgCls, CusNum, AccNum, CrdNum, dummy)
    End Function


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Sends an MQ Message to the host while activating all the visual cues on the UI
    ''' in regards to the current status of sending the message.
    ''' </summary>
    ''' <param name="MsgCode"></param>
    ''' <param name="MsgData"></param>
    ''' <param name="MsgCls"></param>
    ''' <param name="CusNum"></param>
    ''' <param name="AccNum"></param>
    ''' <param name="CrdNum"></param>
    ''' <param name="MsgDump"></param>
    ''' <returns></returns>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[u910krzi]	03/03/2007	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Function SendMessageVisuallyEx(ByVal MsgCode As String, ByVal MsgData As String, ByVal MsgCls As String, _
                                        ByVal CusNum As String, ByVal AccNum As String, ByVal CrdNum As String, ByRef MsgDump As String) As String
        Dim MsgId As String = Nothing
        'Dim m_AppInstance As AppInstanceClass = Me
        SyncLock GetType(AppInstanceClass) 'MQSyncObject
            Try
                Dim retVal As Boolean
                Dim NumEngVal As Integer
                ' IsLoggedOn = True
                If Not IsLoggedOn Then
                    IceLogOn()
                    If Not IsLoggedOn Then
                        ModalMsgBox("The last operation failed because it requires a logged in user." & vbCrLf & _
                               "Please retry this operation and login when prompted.", MsgBoxStyle.Exclamation, "ICE Login required")
                        MsgId = Nothing
                        Exit Try
                    End If
                End If
                'ShowBusyIcon(True)
                NumEngVal = NumEng.NumEngGiv()
                If NumEngVal = 0 Then
                    HandleError(&H80002074, "Error in Numbering Engine: " & NumEng.ErrDsc, "ICE Critical Error")
                    MsgId = ""
                    Exit Try
                End If
                Dim stWksNam As String = Environment.MachineName
                If stWksNam.Length > 8 Then stWksNam = stWksNam.Substring(stWksNam.Length - 8)
                stWksNam = PackString(stWksNam, 8)
                MsgId = stWksNam & NumEng.ConvertToBase64(NumEngVal)
                LastMsgID = MsgId
                ShowActivityIcon(enumActivityType.Write)
                Dim st As String
                st = String.Format("Sending Message: Code({0}:{1}), ID({2}), Data({3})", MsgCode, gfrmMain.DecodeMsgCde(MsgCode), MsgId, MsgData)
                Logger.LogInfo(&H80000031, st, 2)
                Dim nRetries As Integer = MQS.SendRetries
                retVal = False
                'm_AppInstance.Logger.LogInfo(&H80000030, "Message " & MsgId & " was successuflly read.", 3)
                While ((nRetries > 0) And (Not retVal))
                    nRetries -= 1

                    retVal = MQS.PutMessageEx(MsgCode, MsgData, MsgId, MsgCls, CusNum, AccNum, CrdNum, MsgDump, "000", "00", "000")
                    Statistics.MessagesSent += 1
                    If (Not retVal) Then
                        Logger.LogError(&H80000100, "Message Sending Failed! [Attemp " & MQS.SendRetries - nRetries & "]: " & MQS.ErrNum & ": " & MQS.ErrDsc)
                        If (MQS.ErrNum = MQSeries.MQSeries.MQRC_NOT_OPEN) Then
                            Dim reconnect_res As Boolean
                            ShowStatusText("Attempting to reconnect...")
                            Logger.LogInfo(&H80000000, "Attempting to reconnect")
                            reconnect_res = MQS.ReConnect()
                            If reconnect_res Then
                                ShowStatusText("Reconnect Succeeded.")
                                Logger.LogInfo(&H80000100, "Reconnect Succeeded.")
                            Else
                                ShowStatusText("Reconnect Failed!")
                                Logger.LogError(&H80000100, "Reconnect Failed!")
                            End If
                        End If
                        Threading.Thread.Sleep(MQS.ResendDelay)
                    End If
                End While
                MsgId = CStr(IIf(retVal, MsgId, Nothing))
                ShowActivityIcon()
                'ShowBusyIcon()
            Catch ex As Exception
                HandleError(&H80002093, "Error in SendMessageVisually: " & ex.Message, "Send MQSeries Message")
                MsgId = ""
            End Try
            Return MsgId
        End SyncLock
    End Function

    Public Structure OverrideBlock_Struct
        Dim Filled As Boolean
        Dim ErrBlkMrk As String '5x
        Dim ErrBlkTyp As String '3x
        'ErrChn001	31x
        Dim RefReaCde() As String '	10*3x
        Dim RefReaOvf As Boolean
    End Structure

    Public Function MQGetMessage(ByVal MsgID As String, Optional ByVal BypassError As Boolean = False) As String
        'ShowBusyIcon(True)
        Dim iii As Integer = CInt(Rnd() * 10000)
        Dim m_AppInstance As AppInstanceClass = Me
        Dim ReturnString As String = Nothing
        Dim MsgObj As Object = Nothing
        Dim ErrNum As Integer

        SyncLock GetType(AppInstanceClass)
            'While System.Threading.Thread.VolatileRead(MQBusy) > 0
            '    Application.DoEvents()
            '    'System.Threading.Thread.Sleep(1000)
            'End While
            'System.Threading.Thread.VolatileWrite(MQBusy, 1)
            Debug.WriteLine("AppInstance.MQGetMessage: >>> Entering MQGetMessage " & iii)
            m_AppInstance.ShowActivityIcon(enumActivityType.Wait)

            Try
                Dim TimerTicks As Integer
                Dim dt1, dt2 As Date
                Dim ts As TimeSpan
                TimerTicks = 0
                dt1 = Now
                Threading.Thread.VolatileWrite(m_AbortGetMessage, 0)
                m_AppInstance.Logger.LogInfo(&H80000032, "Getting Message with ID:" & MsgID, 3)
                Do
                    m_AppInstance.gfrmMain.ProgressBarValue = CType((((m_AppInstance.MQS.TmeOut - TimerTicks) * 100) / m_AppInstance.MQS.TmeOut), Integer)
                    'secs = CInt(Int((MQS.TmeOut - TimerTicks) / 1000))
                    'If TimerTicks > 0 Then
                    'ticks = CType((MQS.TmeOut - TimerTicks - secs * 1000) / 1000 * 60, Integer)
                    'End If
                    'gfrmMain.ProgressBarString = secs.ToString & ":" & ticks.ToString("#00")
                    m_AppInstance.gfrmMain.ProgressBarString = "Timeout in " & ((m_AppInstance.MQS.TmeOut - TimerTicks) / 1000).ToString("0") & " Sec"
                    m_AppInstance.gfrmMain.MainStatusBar.Invalidate()
                    ErrNum = m_AppInstance.MQS.GetMessage(MsgObj, MsgID)
                    Threading.Thread.Sleep(m_AppInstance.gWaitTaT)
                    dt2 = Now
                    ts = dt2.Subtract(dt1)
                    TimerTicks = CInt(ts.TotalMilliseconds)
                    'TimerTicks += gWaitTaT
                    Application.DoEvents()
                Loop While (TimerTicks < m_AppInstance.MQS.TmeOut) And (ErrNum = 2033) And (Threading.Thread.VolatileRead(m_AbortGetMessage) = 0)
                Threading.Thread.VolatileWrite(m_AbortGetMessage, 0)
                If ErrNum = 0 Then
                    m_AppInstance.Statistics.AverageTat *= m_AppInstance.Statistics.AverageCnt
                    m_AppInstance.Statistics.AverageCnt += 1
                    m_AppInstance.Statistics.AverageTat += (TimerTicks - m_AppInstance.gWaitTaT)
                    m_AppInstance.Statistics.AverageTat /= m_AppInstance.Statistics.AverageCnt
                End If
            Catch ex As Exception
                m_AppInstance.HandleError(&H80002075, "Error in MQGetMessage: " & ex.Message, "Get MQSeries Message")
            Finally
                m_AppInstance.gfrmMain.ProgressBarValue = 0
                m_AppInstance.gfrmMain.ProgressBarString = ""
                m_AppInstance.gfrmMain.MainStatusBar.Invalidate()
            End Try

            If ErrNum = 0 Then
                m_AppInstance.ShowActivityIcon(enumActivityType.Read)
                m_AppInstance.ShowBusyIcon(True)
                ReturnString = m_AppInstance.MQS.ReadMessageFast(MsgObj)
                m_AppInstance.Statistics.MessagesRecv += 1
                System.Threading.Thread.Sleep(200)
                m_AppInstance.Logger.LogInfo(&H80000030, "Message " & MsgID & " was successuflly read.", 3)
            Else
                m_AppInstance.Logger.LogError(&H80002076, "Error while reading message " & MsgID & ":" & m_AppInstance.MQS.ErrNum & ":" & m_AppInstance.MQS.ErrDsc, 3)
            End If
            m_AppInstance.ShowActivityIcon()

            'ShowBusyIcon()

            If (Not (ReturnString Is Nothing)) Then
                ReturnString = ReturnString.Replace(Chr(0), " ")
                Dim st As String = m_AppInstance.SafeSubString(ReturnString, 83, 3)
                If st <> "000" Then
                    If st = "104" Then
                        'We received a Duplicate MID error, using the resulting value to reseed the numbering engine
                        Dim lstMsgID As String = Trim(m_AppInstance.SafeSubString(ReturnString, 206, 12))
                        If lstMsgID.Length = 12 Then
                            m_AppInstance.Logger.LogInfo(&H80000000, "As received from host: Last MID sent is " & lstMsgID)
                            Dim num64 As String = lstMsgID.Substring(8, 4)
                            Dim NewSeed As Integer = m_AppInstance.NumEng.Base64toDecimal(lstMsgID.Substring(8, 4))
                            m_AppInstance.NumEng.SetSeed(NewSeed + 1)
                            m_AppInstance.Logger.LogInfo(&H80000000, "Numbering enging is set to " & m_AppInstance.NumEng.ConvertToBase64((NewSeed + 1)))
                        End If
                    End If
                    Dim MsgResInd As String = m_AppInstance.SafeSubString(ReturnString, 92, 1)
                    If MsgResInd = "S" Or MsgResInd = "T" Then 'check for MsgResInd
                        'Clean up old collection entries. max allowed depth of collection is 100 items
                        'when it reaches the treahshold, the oldest 10 are removed
                        If MQErrors.Count >= 100 Then
                            For colidx As Integer = 0 To 9
                                MQErrors.Remove(0)
                            Next
                        End If
                        'Now add the error entry to the collection
                        Dim rs As String = ReturnString
                        Dim ln As String = 132
                        Dim MQError_Item As MQError_Item_struct
                        MQError_Item.MsgID = MsgID
                        MQError_Item.ErrStr = m_AppInstance.SafeSubString(rs, 256, 132)
                        MQErrors.Add(MQError_Item, MsgID)
                        ln = ExInt(m_AppInstance.SafeSubString(rs, 256 + 132, 4))
                        ReturnString = m_AppInstance.SafeSubString(rs, 0, 256) & m_AppInstance.SafeSubString(rs, 256 + 132 + 4 + ln)
                    End If
                    If MsgResInd = "V" Then 'check for MsgResInd
                        'Clean up old collection entries. max allowed depth of collection is 100 items
                        'when it reaches the treahshold, the oldest 10 are removed
                        If MQErrors.Count >= 100 Then
                            For colidx As Integer = 0 To 9
                                MQErrors.Remove(0)
                            Next
                        End If
                        'Now add the error entry to the collection
                        Dim rs As String = ReturnString
                        Dim MQError_Item As MQError_Item_struct
                        MQError_Item.MsgID = MsgID
                        MQError_Item.ErrStr = m_AppInstance.SafeSubString(rs, 256, 64)
                        MQErrors.Add(MQError_Item, MsgID)
                        ReturnString = m_AppInstance.SafeSubString(rs, 0, 256) & m_AppInstance.SafeSubString(rs, 256 + 64)
                    End If
                    If BypassError Then
                        'ReturnString = m_AppInstance.SafeSubString(ReturnString, 83, 10)
                        'Keep the return string, flag it with *ERROR*, the caller should handle the error.
                        ReturnString = "*ERROR*" & ReturnString
                        m_AppInstance.Logger.LogWarn(&H80002077, "Objective for message " & MsgID & " failed:" & ReturnString, 3)
                    Else
                        ReturnString = m_AppInstance.SafeSubString(ReturnString, 83, 10)
                        m_AppInstance.Logger.LogWarn(&H80002077, "Objective for message " & MsgID & " failed:" & ReturnString, 3)
                    End If
                End If
            End If
            Debug.WriteLine("AppInstance.MQGetMessage: <<< Exiting MQGetMessage " & iii)
            'System.Threading.Thread.VolatileWrite(MQBusy, 0)
        End SyncLock
        'Insert "CST=" version table checking code in here
        Return ReturnString
    End Function

    Public Function MQGetMessageEx(ByVal MsgID As String, ByRef OverrideBlock As OverrideBlock_Struct, Optional ByVal BypassError As Boolean = False) As String
        'ShowBusyIcon(True)
        Dim iii As Integer = CInt(Rnd() * 10000)
        Dim m_AppInstance As AppInstanceClass = Me
        Dim ReturnString As String = Nothing
        Dim MsgObj As Object = Nothing
        Dim ErrNum As Integer
        OverrideBlock.Filled = False

        SyncLock GetType(AppInstanceClass)
            'While System.Threading.Thread.VolatileRead(MQBusy) > 0
            '    Application.DoEvents()
            '    'System.Threading.Thread.Sleep(1000)
            'End While
            'System.Threading.Thread.VolatileWrite(MQBusy, 1)
            Debug.WriteLine("AppInstance.MQGetMessage: >>> Entering MQGetMessage " & iii)
            m_AppInstance.ShowActivityIcon(enumActivityType.Wait)

            Try
                Dim TimerTicks As Integer
                Dim dt1, dt2 As Date
                Dim ts As TimeSpan
                TimerTicks = 0
                dt1 = Now
                Threading.Thread.VolatileWrite(m_AbortGetMessage, 0)
                m_AppInstance.Logger.LogInfo(&H80000032, "Getting Message with ID:" & MsgID, 3)
                Do
                    m_AppInstance.gfrmMain.ProgressBarValue = CType((((m_AppInstance.MQS.TmeOut - TimerTicks) * 100) / m_AppInstance.MQS.TmeOut), Integer)
                    'secs = CInt(Int((MQS.TmeOut - TimerTicks) / 1000))
                    'If TimerTicks > 0 Then
                    'ticks = CType((MQS.TmeOut - TimerTicks - secs * 1000) / 1000 * 60, Integer)
                    'End If
                    'gfrmMain.ProgressBarString = secs.ToString & ":" & ticks.ToString("#00")
                    m_AppInstance.gfrmMain.ProgressBarString = "Timeout in " & ((m_AppInstance.MQS.TmeOut - TimerTicks) / 1000).ToString("0") & " Sec"
                    m_AppInstance.gfrmMain.MainStatusBar.Invalidate()
                    ErrNum = m_AppInstance.MQS.GetMessage(MsgObj, MsgID)
                    Threading.Thread.Sleep(m_AppInstance.gWaitTaT)
                    dt2 = Now
                    ts = dt2.Subtract(dt1)
                    TimerTicks = CInt(ts.TotalMilliseconds)
                    'TimerTicks += gWaitTaT
                    Application.DoEvents()
                Loop While (TimerTicks < m_AppInstance.MQS.TmeOut) And (ErrNum = 2033) And (Threading.Thread.VolatileRead(m_AbortGetMessage) = 0)
                Threading.Thread.VolatileWrite(m_AbortGetMessage, 0)
                If ErrNum = 0 Then
                    m_AppInstance.Statistics.AverageTat *= m_AppInstance.Statistics.AverageCnt
                    m_AppInstance.Statistics.AverageCnt += 1
                    m_AppInstance.Statistics.AverageTat += (TimerTicks - m_AppInstance.gWaitTaT)
                    m_AppInstance.Statistics.AverageTat /= m_AppInstance.Statistics.AverageCnt
                End If
            Catch ex As Exception
                m_AppInstance.HandleError(&H80002075, "Error in MQGetMessage: " & ex.Message, "Get MQSeries Message")
            Finally
                m_AppInstance.gfrmMain.ProgressBarValue = 0
                m_AppInstance.gfrmMain.ProgressBarString = ""
                m_AppInstance.gfrmMain.MainStatusBar.Invalidate()
            End Try

            If ErrNum = 0 Then
                m_AppInstance.ShowActivityIcon(enumActivityType.Read)
                m_AppInstance.ShowBusyIcon(True)
                ReturnString = m_AppInstance.MQS.ReadMessageFast(MsgObj)
                m_AppInstance.Statistics.MessagesRecv += 1
                System.Threading.Thread.Sleep(200)
                m_AppInstance.Logger.LogInfo(&H80000030, "Message " & MsgID & " was successuflly read.", 3)
            Else
                m_AppInstance.Logger.LogError(&H80002076, "Error while reading message " & MsgID & ":" & m_AppInstance.MQS.ErrNum & ":" & m_AppInstance.MQS.ErrDsc, 3)
            End If
            m_AppInstance.ShowActivityIcon()

            'ShowBusyIcon()

            If (Not (ReturnString Is Nothing)) Then
                ReturnString = ReturnString.Replace(Chr(0), " ")
                Dim st As String = m_AppInstance.SafeSubString(ReturnString, 83, 3)
                'Extract the override block, if available
                Dim MsgResInd As String = m_AppInstance.SafeSubString(ReturnString, 92, 1)
                If MsgResInd = "S" Or MsgResInd = "T" Then 'check for MsgResInd
                    Dim rs As String = ReturnString
                    Dim ln As String = 132
                    Dim OvrStr As String = m_AppInstance.SafeSubString(rs, 256, 39) '64)
                    With OverrideBlock
                        .ErrBlkMrk = strip(OvrStr, 5)
                        .ErrBlkTyp = strip(OvrStr, 3)
                        .Filled = (OvrStr.Trim <> "") 'True
                        ReDim .RefReaCde(9)
                        For i As Integer = 0 To 9 : .RefReaCde(i) = strip(OvrStr, 3) : Next
                        .RefReaOvf = (strip(OvrStr, 1) = "+")
                    End With
                    ln = ExInt(m_AppInstance.SafeSubString(rs, 256 + 132, 4))
                    ReturnString = m_AppInstance.SafeSubString(rs, 0, 256) & m_AppInstance.SafeSubString(rs, 256 + 132 + 4 + ln)
                ElseIf MsgResInd = "V" Then 'check for MsgResInd
                    Dim rs As String = ReturnString
                    Dim OvrStr As String = m_AppInstance.SafeSubString(rs, 256, 39) '64)
                    With OverrideBlock
                        .ErrBlkMrk = strip(OvrStr, 5)
                        .ErrBlkTyp = strip(OvrStr, 3)
                        .Filled = (OvrStr.Trim <> "") 'True
                        ReDim .RefReaCde(9)
                        For i As Integer = 0 To 9 : .RefReaCde(i) = strip(OvrStr, 3) : Next
                        .RefReaOvf = (strip(OvrStr, 1) = "+")
                    End With
                    ReturnString = m_AppInstance.SafeSubString(rs, 0, 256) & m_AppInstance.SafeSubString(rs, 256 + 64)
                End If
                'Check when there is an error
                If st <> "000" Then
                    If st = "104" Then
                        'We received a Duplicate MID error, using the resulting value to reseed the numbering engine
                        Dim lstMsgID As String = Trim(m_AppInstance.SafeSubString(ReturnString, 206, 12))
                        If lstMsgID.Length = 12 Then
                            m_AppInstance.Logger.LogInfo(&H80000000, "As received from host: Last MID sent is " & lstMsgID)
                            Dim num64 As String = lstMsgID.Substring(8, 4)
                            Dim NewSeed As Integer = m_AppInstance.NumEng.Base64toDecimal(lstMsgID.Substring(8, 4))
                            m_AppInstance.NumEng.SetSeed(NewSeed + 1)
                            m_AppInstance.Logger.LogInfo(&H80000000, "Numbering enging is set to " & m_AppInstance.NumEng.ConvertToBase64((NewSeed + 1)))
                        End If
                    End If
                    If BypassError Then
                        'ReturnString = m_AppInstance.SafeSubString(ReturnString, 83, 10)
                        'Keep the return string, flag it with *ERROR*, the caller should handle the error.
                        ReturnString = "*ERROR*" & ReturnString
                        m_AppInstance.Logger.LogWarn(&H80002077, "Objective for message " & MsgID & " failed:" & ReturnString, 3)
                    Else
                        ReturnString = m_AppInstance.SafeSubString(ReturnString, 83, 10)
                        m_AppInstance.Logger.LogWarn(&H80002077, "Objective for message " & MsgID & " failed:" & ReturnString, 3)
                    End If
                End If
            End If
            Debug.WriteLine("AppInstance.MQGetMessage: <<< Exiting MQGetMessage " & iii)
            'System.Threading.Thread.VolatileWrite(MQBusy, 0)
        End SyncLock
        'Insert "CST=" version table checking code in here
        Return ReturnString
    End Function


#End Region

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''     Saves the Form Last Location and Size.
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Y093sahu]	dd/mm/yyyy	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Sub getFormLastPosition(ByRef p_frm As Form)
        Dim oWindowState As FormWindowState
        Dim rect As Rectangle
        Dim sName As String

        'sName = "[(" & p_frm.GetType().Assembly.CodeBase & ").(" & p_frm.GetType().Assembly.FullName & ")]." & p_frm.Name
        sName = p_frm.Name
        sName = sName.Replace("frm", "FORM_")
        rect = Me.LoadLastPosition(sName, oWindowState)

        If oWindowState = FormWindowState.Normal Then
            p_frm.Location = New Point(rect.X, rect.Y)
            Dim oSize As New Size(rect.Width, rect.Height)
            If ((oSize.Width > 0) And (oSize.Height > 0)) Then
                p_frm.Size = oSize
            End If
        Else
            p_frm.WindowState = oWindowState
        End If
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''     Gets called on closing the Form and it saves the position of the window and before closing it
    '''      will check if the window is updated and if its is it will give a question to discard them and close or not.
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Y093sahu]	dd/mm/yyyy	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Function setFormLastPosition(ByRef p_frm As Form, ByRef p_isFormUpdated As Boolean) As Boolean
        If (Me.gAppBusy And Not (Me.gLoggingOff)) Then Return True

        Dim sName As String

        'sName = "[(" & p_frm.GetType().Assembly.CodeBase & ").(" & p_frm.GetType().Assembly.FullName & ")]." & p_frm.Name
        sName = p_frm.Name
        sName = sName.Replace("frm", "FORM_")
        Me.SaveLastPosition(sName, p_frm.WindowState, New Rectangle(p_frm.Location, p_frm.Size))

        If p_isFormUpdated Then
            Dim Res As MsgBoxResult
            Res = Me.ModalMsgBox("You have made some changes that are not yet submitted to the host" & vbCrLf & _
                                        "Do you want to continue and discard those changes?", _
                                        MsgBoxStyle.YesNo Or MsgBoxStyle.DefaultButton2 Or MsgBoxStyle.Information, "Invalidate " & p_frm.Text)
            If Res = MsgBoxResult.No Then Return True
        End If
        p_isFormUpdated = False
        Return False
    End Function

    Public Function LocateBrnCfgRow(ByVal BrnNum As String) As SortedList
        Dim l_Assm As [Assembly] = [Assembly].GetEntryAssembly
        Dim modtyp As Type = l_Assm.GetType("ICEP0100.ICEDataBase")
        Dim rmi As MethodInfo = modtyp.GetMethod("LocateBrnCfgRow")
        Return DirectCast(rmi.Invoke(Nothing, New Object() {BrnNum}), SortedList)
    End Function

    Public Function GregToHijriDate(ByVal DateStr As String, ByRef OutStr As String) As Boolean
        Dim l_Assm As [Assembly] = [Assembly].GetEntryAssembly
        Dim modtyp As Type = l_Assm.GetType("ICEP0100.ICEDataBase")
        Dim rmi As MethodInfo = modtyp.GetMethod("GregToHijriDate", New Type() {GetType(String), Type.GetType("System.String&")})
        Dim PrmObj() As Object = {DateStr, OutStr}
        Dim Res As Boolean = DirectCast(rmi.Invoke(Nothing, PrmObj), Boolean)
        OutStr = PrmObj(1)
        Return Res
    End Function

    Public Function HijriToGregDate(ByVal DateStr As String, ByRef OutStr As String) As Boolean
        Dim l_Assm As [Assembly] = [Assembly].GetEntryAssembly
        Dim modtyp As Type = l_Assm.GetType("ICEP0100.ICEDataBase")
        Dim rmi As MethodInfo = modtyp.GetMethod("HijriToGregDate", New Type() {GetType(String), Type.GetType("System.String&")})
        Dim PrmObj() As Object = {DateStr, OutStr}
        Dim Res As Boolean = DirectCast(rmi.Invoke(Nothing, PrmObj), Boolean)
        OutStr = PrmObj(1)
        Return Res
    End Function

    Public Function GregToHijriDate(ByVal DateStr As String, ByRef ErrStr As String, ByVal ShowErrors As Boolean) As String
        Dim l_Assm As [Assembly] = [Assembly].GetEntryAssembly
        Dim modtyp As Type = l_Assm.GetType("ICEP0100.ICEDataBase")
        Dim rmi As MethodInfo = modtyp.GetMethod("GregToHijriDate", New Type() {GetType(String), Type.GetType("System.String&"), GetType(Boolean)})
        Dim PrmObj() As Object = {DateStr, ErrStr, ShowErrors}
        Dim Res As String = DirectCast(rmi.Invoke(Nothing, PrmObj), String)
        ErrStr = PrmObj(1)
        Return Res
    End Function

    Public Function HijriToGregDate(ByVal DateStr As String, ByRef ErrStr As String, ByVal ShowErrors As Boolean) As String
        Dim l_Assm As [Assembly] = [Assembly].GetEntryAssembly
        Dim modtyp As Type = l_Assm.GetType("ICEP0100.ICEDataBase")
        Dim rmi As MethodInfo = modtyp.GetMethod("HijriToGregDate", New Type() {GetType(String), Type.GetType("System.String&"), GetType(Boolean)})
        Dim PrmObj() As Object = {DateStr, ErrStr, ShowErrors}
        Dim Res As String = DirectCast(rmi.Invoke(Nothing, PrmObj), String)
        ErrStr = PrmObj(1)
        Return Res
    End Function

    Public Function getBrnTabTableCollection() As SortedList
        Dim l_Assm As [Assembly] = [Assembly].GetEntryAssembly
        Dim modtyp As Type = l_Assm.GetType("ICEP0100.ICEDataBase")
        Dim rmi As MethodInfo = modtyp.GetMethod("getBrnTabTableCollection")
        Return DirectCast(rmi.Invoke(Nothing, Nothing), SortedList)
    End Function

    Public Function getDTSDBServer() As SortedList
        Dim l_Assm As [Assembly] = [Assembly].GetEntryAssembly
        Dim modtyp As Type = l_Assm.GetType("ICEP0100.ICEDataBase")
        Dim rmi As MethodInfo = modtyp.GetMethod("getDTSDBServer")
        Return DirectCast(rmi.Invoke(Nothing, Nothing), SortedList)
    End Function

    Public Function getBrnTabTable() As String()
        Dim l_Assm As [Assembly] = [Assembly].GetEntryAssembly
        Dim modtyp As Type = l_Assm.GetType("ICEP0100.ICEDataBase")
        Dim rmi As MethodInfo = modtyp.GetMethod("getBrnTabTable")
        Return DirectCast(rmi.Invoke(Nothing, Nothing), String())
    End Function
End Class

''Remember that u can't log info inside the functions of this class
Public Class IcePaths

#Region " Fields "

    Private Shared myLock As New Object
    Private Shared m_defInstance As IcePaths = Nothing

    Private m_ICE As String = String.Empty
    Private m_TemplatePath As String = String.Empty
    Private m_DatabasePath As String = String.Empty
    Private m_LogFilePath As String = String.Empty
    Private m_VersionsPath As String = String.Empty


    Public Const C_IceCfg_INI As String = "IceCfg.INI"
    Public Const C_IceUsr_INI As String = "IceUsr.INI"


#End Region

#Region " Properties "

    Public Shared ReadOnly Property DefInstance() As IcePaths
        Get
            If m_defInstance Is Nothing Then
                SyncLock myLock
                    m_defInstance = New IcePaths
                End SyncLock
            End If
            Return m_defInstance
        End Get
    End Property

    Public ReadOnly Property TemplatePath() As String
        Get
            Return m_TemplatePath
        End Get
    End Property

    Public ReadOnly Property DatabasePath() As String
        Get
            Return m_DatabasePath
        End Get
    End Property

    Public ReadOnly Property LogFilePath() As String
        Get
            Return m_LogFilePath
        End Get
    End Property

    Public ReadOnly Property VersionsPath() As String
        Get
            Return m_VersionsPath
        End Get
    End Property

    Public ReadOnly Property ICE() As String
        Get
            Return m_ICE
        End Get
    End Property

    Public ReadOnly Property IceCfgIni() As String
        Get
            Return Me.ICE & "\" & C_IceCfg_INI
        End Get
    End Property

    Public ReadOnly Property IceUsrIni() As String
        Get
            Return Me.ICE & "\" & C_IceUsr_INI
        End Get
    End Property

#End Region

    Private Sub New()
        load()
    End Sub

    'Load paths
    Public Sub load()
        Dim oOS As New SIBL0100.COS

        'If oOS.isVistaPlus() Then 'Or m_isSimulateVista
        '    settingVistaPaths()
        'Else
            Me.m_ICE = Application.StartupPath
        'End If

        loadIni()

    End Sub

    Public Sub settingVistaPaths()
        Try
            With Me
                'Creating paths
                Dim oOS As New SIBL0100.COS
                Dim sUsersApplicationDataPath As String = String.Empty
                Dim sICEPath As String = String.Empty

                sUsersApplicationDataPath = oOS.getAllUsersApplicationData()

                sICEPath = sUsersApplicationDataPath & "\SAIB"
                SIBL0100.UFolder.createFolder(sICEPath)
                sICEPath &= "\ICE"
                SIBL0100.UFolder.createFolder(sICEPath)

                .m_ICE = sICEPath
                .m_TemplatePath = .m_ICE & "\Templates"
                .m_DatabasePath = .m_ICE & "\Data"
                .m_LogFilePath = .m_ICE & "\Logs"
                .m_VersionsPath = .m_ICE & "\Versions"

                SIBL0100.UFolder.createFolder(.m_TemplatePath)
                SIBL0100.UFolder.createFolder(.m_DatabasePath)
                SIBL0100.UFolder.createFolder(.m_LogFilePath)
                SIBL0100.UFolder.createFolder(.m_VersionsPath)

                'Copying ini
                If Not IO.File.Exists(Me.IceCfgIni) Then
                    IO.File.Copy(Application.StartupPath & "\" & C_IceCfg_INI, Me.IceCfgIni)
                End If

                If Not IO.File.Exists(Me.IceUsrIni) Then
                    IO.File.Copy(Application.StartupPath & "\" & C_IceUsr_INI, Me.IceUsrIni)
                End If

                'Coping the files in dir m_DatabasePath
                SIBL0100.UFolder.CopyFolder(Application.StartupPath & "\Data", .m_DatabasePath)

                SIBL0100.UFolder.FolderReadOnly(.ICE, False)
                SIBL0100.UFolder.FolderReadOnly(.DatabasePath, False)

                Dim iniUsr As New INIClass.IniFileClass(Me.IceUsrIni)

                'Resetting the paths in the user.ini
                iniUsr.SectionName = "PATHS"
                iniUsr.SetValue("TemplatePath", .m_TemplatePath)
                iniUsr.SetValue("DatabasePath", .m_DatabasePath)
                iniUsr.SetValue("LogFilePath", .m_LogFilePath)
                iniUsr.SetValue("VersionPath", .m_VersionsPath)

            End With

        Catch ex As Exception
            MessageBox.Show("Error in settingVistaPaths. " & vbCrLf & ex.Message, "ICE Critical Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Public Sub loadIni()
        Try

            Dim iniCfg As New INIClass.IniFileClass(Me.IceCfgIni)
            Dim iniUsr As New INIClass.IniFileClass(Me.IceUsrIni)

            '  MessageBox.Show("is iniUsr nothing: " & (iniUsr Is Nothing).ToString, "test", MessageBoxButtons.OK, MessageBoxIcon.Information)
            With Me
                iniUsr.SectionName = "PATHS"
                .m_TemplatePath = iniUsr.GetValue("TemplatePath").Trim
                .m_DatabasePath = iniUsr.GetValue("DatabasePath").Trim
                .m_LogFilePath = iniUsr.GetValue("LogFilePath").Trim
                .m_VersionsPath = iniUsr.GetValue("VersionPath").Trim

                Dim sMsgErr As String = String.Empty
                sMsgErr = "800000003: ICE Application generated an error and cannot continue!" & vbCrLf & "{0} [{1}] " & vbCrLf & "was not found."

                If Not IO.Directory.Exists(.m_ICE) Then
                    AppInstanceClass.DefInstance.ModalMsgBox(String.Format(sMsgErr, "ICEPath", .m_ICE), MsgBoxStyle.Critical, "ICE Critical Error")
                End If

                If Not IO.Directory.Exists(.m_TemplatePath) Then
                    AppInstanceClass.DefInstance.ModalMsgBox(String.Format(sMsgErr, "TemplatePath", .m_TemplatePath), MsgBoxStyle.Critical, "ICE Critical Error")
                End If
                If Not IO.Directory.Exists(.m_LogFilePath) Then
                    AppInstanceClass.DefInstance.ModalMsgBox(String.Format(sMsgErr, "LogFilePath", .m_LogFilePath), MsgBoxStyle.Critical, "ICE Critical Error")
                End If
                If Not IO.Directory.Exists(.m_VersionsPath) Then
                    AppInstanceClass.DefInstance.ModalMsgBox(String.Format(sMsgErr, "VersionsPath", .m_VersionsPath), MsgBoxStyle.Critical, "ICE Critical Error")
                End If
                If Not IO.Directory.Exists(.m_DatabasePath) Then
                    AppInstanceClass.DefInstance.ModalMsgBox(String.Format(sMsgErr, "DatabasePath", .m_DatabasePath), MsgBoxStyle.Critical, "ICE Critical Error")
                End If
            End With

        Catch ex As Exception
            MessageBox.Show("Error in loading ini paths. " & vbCrLf & ex.Message, "ICE Critical Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    
End Class
