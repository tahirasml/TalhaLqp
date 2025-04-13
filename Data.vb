Option Strict Off

Imports System.Runtime.InteropServices
Imports System.Security
Imports System.Diagnostics
Imports Microsoft.Win32
Imports System.Threading
Imports System.Globalization
Imports System.Reflection
Imports win = System.Windows.Forms
Imports sys = System.Diagnostics
Imports dtsh = DTS.Helper


'Imports win32 = Microsoft.Win32
'Imports System.Globalization.GregorianCalendar



Public Module Data

    Public Const WM_SIZE As Integer = 5
    Public Const SIZE_MINIMIZED As Integer = 1
    Public Const SIZE_MAXIMIZED As Integer = 2

    Public Const STRING_NUMBERS As String = "0123456789"
    Public Const STRING_ARAPUNC As String = " +()،؟؛,-./"
    Public Const STRING_ENGPUNC As String = " +(),-./"
    Public Const STRING_ARAALPHA As String = "ضصثقفغعهخحجدطكمنتالبيسشئءؤرلاىةوزظذإلإألألآ"
    Public Const STRING_ENGALPHA As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz"
    Public Const STRING_ENGALPHA_UPPER As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
    Public Const STRING_ENGALPHA_LOWER As String = "abcdefghijklmnopqrstuvwxyz"

    Public m_hitLimitHigh As Integer = 90
    Public m_hitLimitLow As Integer = 60
    Public m_MailerType As String = ""
    Public m_useANTSNAAN As Boolean = False
    ''' RASD fields validation
    'Public m_RASDValidateNat As Boolean = False
    'Public m_RASDValidateNIN As Boolean = False
    ''' End
    ''' HR posting variables
    Public G_LastRunDate As String
    Public G_ProcessingStep As Integer
    Public G_sSrcDirSAIB As String
    Public G_sSrcDirEMS As String
    Public G_TargetLibrary As String
    'Public G_TargetConn As OleDb.OleDbConnection
    Public G_TargetConn As Odbc.OdbcConnection
    Public G_TargetDSN As String
    Public G_SCMthPostFile As String ' The AS/400 single-currency table name
    Public G_SCDlyPostFile As String

    '''' End
    Public m_EnableSigCap As Boolean = True
    Public m_DoubtfulCheckFreqency As Long = 300000

    Public gCultureInfoEnAU As CultureInfo = New CultureInfo("en-AU")
    Public gCultureInfoArSA As CultureInfo = New CultureInfo("ar-SA")
    Public gbPrintingLang As enumReceiptLanguage = enumReceiptLanguage.English

    Public Enum enumReceiptLanguage
        Arabic
        English
    End Enum

    Public Enum SignatureSystem
        SigCap
        VeriPark
    End Enum

    Public ReadOnly Property HitLimitHigh() As Integer
        Get
            Return m_hitLimitHigh
        End Get
    End Property

    Public ReadOnly Property HitLimitLow() As Integer
        Get
            Return m_hitLimitLow
        End Get
    End Property

    Public ReadOnly Property UseANTSNAAN() As Boolean
        Get
            Return m_useANTSNAAN
        End Get
    End Property


    Public Class ExDataGrid
        Inherits DataGrid

        Private m_NextControl As Control

        Protected Overrides Sub OnMouseUp(ByVal e As System.Windows.Forms.MouseEventArgs)
            MyBase.OnMouseUp(e)
            Dim ht As DataGrid.HitTestInfo
            ht = HitTest(e.X, e.Y)
            If (Not ht.Equals(ht.Nowhere)) Then
                If (ht.Type = DataGrid.HitTestType.ColumnHeader) Then
                    AutoSizeGrid(Me)
                End If
            End If
        End Sub

        Protected Overrides Function ProcessCmdKey(ByRef msg As System.Windows.Forms.Message, ByVal keyData As System.Windows.Forms.Keys) As Boolean
            If msg.WParam.ToInt32() = CInt(Keys.Tab) Then
                'SendKeys.Send("{Tab}")
                Try
                    If (Not (m_NextControl Is Nothing)) Then
                        m_NextControl.Focus()
                        Return True
                    End If
                    Me.Parent.Focus()
                    Return True
                Catch
                    'Nothing, let the base class handle the event.
                End Try
            End If
            Return MyBase.ProcessCmdKey(msg, keyData)
        End Function

        Public Sub New()
            MyBase.New()
        End Sub

        Public Sub New(ByVal NextControl As Control)
            MyBase.New()
            m_NextControl = NextControl
        End Sub
    End Class

    Public Class ListViewStringItemComparer
        Implements IComparer

        Private col As Integer
        Private order As Boolean

        Public Sub New()
            col = 0
        End Sub

        Public Sub New(ByVal column As Integer, ByVal SortOrder As Boolean)
            col = column
            order = SortOrder
        End Sub

        Public Function Compare(ByVal x As Object, ByVal y As Object) As Integer _
           Implements IComparer.Compare
            If order Then
                Return [String].Compare(CType(y, ListViewItem).SubItems(col).Text, CType(x, ListViewItem).SubItems(col).Text)
            Else
                Return [String].Compare(CType(x, ListViewItem).SubItems(col).Text, CType(y, ListViewItem).SubItems(col).Text)
            End If
        End Function
    End Class

    Public Class ListViewDateItemComparer
        Implements IComparer

        Private col As Integer
        Private order As Boolean
        Private m_Len As Integer

        Public Sub New()
            col = 0
        End Sub

        Public Sub New(ByVal column As Integer, ByVal SortOrder As Boolean, Optional ByVal Len As Integer = -1)
            col = column
            order = SortOrder
            m_Len = Len
        End Sub

        Public Function Compare(ByVal x As Object, ByVal y As Object) As Integer _
           Implements IComparer.Compare
            Dim Dt1, Dt2 As Date
            Try
                'Check if dates are empty, this will speed up the comparision as no exception will be raised
                If CType(x, ListViewItem).SubItems(col).Text.Trim = "" Then Return CInt(IIf(order, 1, -1))
                If CType(y, ListViewItem).SubItems(col).Text.Trim = "" Then Return CInt(IIf(order, -1, 1))

                If m_Len > 0 Then
                    Dt1 = SIBL0100.Util.UDate.Str2Date(AppInstance.SafeSubString(CType(x, ListViewItem).SubItems(col).Text, 0, m_Len))
                Else
                    Dt1 = SIBL0100.Util.UDate.Str2Date(CType(x, ListViewItem).SubItems(col).Text)
                End If

            Catch
                Dt1 = Now
            End Try
            Try
                If m_Len > 0 Then
                    Dt2 = SIBL0100.Util.UDate.Str2Date(AppInstance.SafeSubString(CType(y, ListViewItem).SubItems(col).Text, 0, m_Len))
                Else
                    Dt2 = SIBL0100.Util.UDate.Str2Date(CType(y, ListViewItem).SubItems(col).Text)
                End If

            Catch
                Dt2 = Now
            End Try
            If order Then
                If Dt2 < Dt1 Then Return -1
                If Dt2 > Dt1 Then Return 1
                Return 0
            Else
                If Dt1 < Dt2 Then Return -1
                If Dt1 > Dt2 Then Return 1
                Return 0
            End If
        End Function
    End Class

    Public Class ListViewDoubleItemComparer
        Implements IComparer

        Private col As Integer
        Private order As Boolean

        Public Sub New()
            col = 0
        End Sub

        Public Sub New(ByVal column As Integer, ByVal SortOrder As Boolean)
            col = column
            order = SortOrder
        End Sub

        Public Function Compare(ByVal x As Object, ByVal y As Object) As Integer _
           Implements IComparer.Compare
            Dim d1 As Double = AppInstance.ExVal(CType(x, ListViewItem).SubItems(col).Text)
            Dim d2 As Double = AppInstance.ExVal(CType(y, ListViewItem).SubItems(col).Text)
            If order Then
                If d2 < d1 Then Return -1
                If d2 > d1 Then Return 1
                Return 0
            Else
                If d1 < d2 Then Return -1
                If d1 > d2 Then Return 1
                Return 0
            End If
        End Function
    End Class

    Public Class ServiceTotals_Class
        Public CCY As String '3x
        Public CshInwAmt As String '14n Net Cash amount
        Public CshInwCnt As Integer
        Public CshOutAmt As String '14n Net Cash amount
        Public CshOutCnt As Integer
        Public CshNetAmt As String '14n Net Cash amount
        Public InsInwAmt As String '14n Instruments In amount
        Public InsInwCnt As Integer
        Public InsOutAmt As String '14n Instruments Out amount
        Public InsOutCnt As Integer
        Public InsNetAmt As String '14n Instruments Out amount
    End Class

    Public ServiceTotals As New Collection

#Region "Declarations"
    Private Structure RollBack_Struct
        Dim OrgFilSrc As String
        Dim NewFileSrc As String
        Dim FileVer As String
        Dim FileName As String
    End Structure
#End Region

#Region "Globals - Variables"
    Public AppInstance As ICEI0100.AppInstanceClass
    Public AppLoaded As Boolean
    'Public enumActivity As enumActivityType
    Public gIceUnitList() As String
    Public gIceVersion As String
    Public gIceBuild As String
#End Region

#Region "Globals - Forms"
    Public gfrmSplash As frmSplash
    Public gfrmMain As frmMain
    Public gfrmCustomer As frmCustomer
    Public gfrmPrintForms As frmForms
    Public gfrmManageAutDsc As frmIceAutDsc
    Public gfrmManageIceRol As frmIceRole
    Public gfrmError As frmError
    Public gfrmRevealPwd As frmRevealPwd
    Public gfrmConfig As frmConfig
    Public gfrmStatistics As frmStatistics
    Public gfrmPwMailer As frmPwMailer
    Public gfrmFoxRates As frmFoxRates
    Public gfrmPrint As frmPrint
    Public gfrmTPP As frmTPP
    Public gfrmOlClearCache As frmOlClearCache
    Public gfrmJournal As frmJournal
    Public gfrmNumEngine As frmNumEngine
    'Public gfrmAnyMailer As frmAnyMailer
    'Public gfrmAnyMailer_Invalidate As frmAnyMailer_Invalidate
    'Public gfrmAnyMailer_Allocate As frmAnyMailer_Allocate
    'Public gfrmUsbKey_Init As frmUsbKey_Init
    'Public gfrmUsbKey_Manage As frmUsbKey_Manage

    Public gfrmLogViewer As frmLogViewer
    Public gfrmLogViewerMQ As frmLogViewerMQ
    Public gfrmInternaAccounts As frmInternalAccounts

#End Region

    Public Function CControl(ByVal ctl As Object) As Control
        Return DirectCast(ctl, Control) 'CType(ctl, Control)
    End Function

    Public Sub test123()
        'Dim wooo As New WOSA0110.WRM_WFSPTRPRINTFORM_Class
        'wooo.AddField("test=123")
        'wooo.AddField("test=456")
        'wooo.AddField("test=789")
        'wooo.AddField("test=555")
        'Dim st As String = (wooo.PackFields)
        'AppInstance.ModalMsgBox(st)
    End Sub

    Public Function IsEmpty(ByVal str As String) As Boolean
        Return (str Is Nothing) OrElse (str.Trim = "")
    End Function

    Public Sub Swap(ByRef Obj1 As Object, ByRef Obj2 As Object)
        Dim ObjTemp As Object
        ObjTemp = Obj2
        Obj2 = Obj1
        Obj1 = ObjTemp
    End Sub

    Public Sub Sleep(ByVal dwMillisec As Int32)
        System.Threading.Thread.Sleep(dwMillisec)
    End Sub

    Public Function SwitchLocale(ByVal LocaleID As String) As CultureInfo
        Dim OldCultInfo, NewCultInfo As CultureInfo
        OldCultInfo = System.Threading.Thread.CurrentThread.CurrentCulture
        NewCultInfo = New CultureInfo(LocaleID, False)
        System.Threading.Thread.CurrentThread.CurrentCulture = NewCultInfo
        Return OldCultInfo
    End Function

    Public Sub RestoreLocale(ByVal LocaleInfo As CultureInfo)
        System.Threading.Thread.CurrentThread.CurrentUICulture = LocaleInfo
    End Sub

    Private Sub LoadOptions(ByVal LoadStep As Integer)
        'Load paths
        'Dim sInisializeFolders As String 'This call just to init the folders and paths for ice
        'sInisializeFolders = ICEI0100.IcePaths.DefInstance.DatabasePath

        Dim iniCfg As New INIClass.IniFileClass(ICEI0100.IcePaths.DefInstance.IceCfgIni)
        Dim iniUsr As New INIClass.IniFileClass(ICEI0100.IcePaths.DefInstance.IceUsrIni)
        If Not (iniCfg.FileExist) Then
            ModalMsgBox("80000002: ICE Application generated an error and cannot continue!" & vbCrLf & ICEI0100.IcePaths.DefInstance.IceCfgIni & " was not found.", _
                                MsgBoxStyle.Critical, "ICE Critical Error")
            End
        End If
        If Not (iniUsr.FileExist) Then
            ModalMsgBox("800000003: ICE Application generated an error and cannot continue!" & vbCrLf & ICEI0100.IcePaths.DefInstance.IceUsrIni & " was not found.", _
                                MsgBoxStyle.Critical, "ICE Critical Error")
            End
        End If
        Select Case LoadStep
            'Case 1 'Load paths
            '    Dim sInisializeFolders As String
            '    sInisializeFolders = ICEI0100.IcePaths.DefInstance.DatabasePath
            'With AppInstance.IcePaths
            '    iniUsr.SectionName = "PATHS"
            '    .TemplatePath = iniUsr.GetValue("TemplatePath").Trim
            '    .DatabasePath = iniUsr.GetValue("DatabasePath").Trim
            '    .LogFilePath = iniUsr.GetValue("LogFilePath").Trim
            '    .VersionsPath = iniUsr.GetValue("VersionPath").Trim
            '    '*** Check for Legacy INI file, where these values where in iniCfg instead
            '    If (.TemplatePath = iniUsr.KeyWordNotFound) Then
            '        iniCfg.SectionName = "PATHS"
            '        .TemplatePath = iniCfg.GetValue("TemplatePath", , Application.StartupPath & "\Templates").Trim
            '        .DatabasePath = iniCfg.GetValue("DatabasePath", , Application.StartupPath & "\Data").Trim
            '        .LogFilePath = iniCfg.GetValue("LogFilePath", , Application.StartupPath & "\Logs").Trim
            '        .VersionsPath = iniCfg.GetValue("VersionPath", , Application.StartupPath & "\Versions").Trim
            '    End If
            '    If Not IO.Directory.Exists(.TemplatePath) Then
            '        ModalMsgBox("800000003: ICE Application generated an error and cannot continue!" & vbCrLf & .TemplatePath & " was not found.", _
            '                            MsgBoxStyle.Critical, "ICE Critical Error")
            '        End
            '    End If
            '    If Not IO.Directory.Exists(.LogFilePath) Then
            '        ModalMsgBox("800000003: ICE Application generated an error and cannot continue!" & vbCrLf & .LogFilePath & " was not found.", _
            '                            MsgBoxStyle.Critical, "ICE Critical Error")
            '        End
            '    End If
            '    If Not IO.Directory.Exists(.VersionsPath) Then
            '        ModalMsgBox("800000003: ICE Application generated an error and cannot continue!" & vbCrLf & .VersionsPath & " was not found.", _
            '                            MsgBoxStyle.Critical, "ICE Critical Error")
            '        End
            '    End If
            '    If Not IO.Directory.Exists(.DatabasePath) Then
            '        ModalMsgBox("800000003: ICE Application generated an error and cannot continue!" & vbCrLf & .DatabasePath & " was not found.", _
            '                            MsgBoxStyle.Critical, "ICE Critical Error")
            '        End
            '    End If
            'End With
            Case 2 'Load Log file options
                AppInstance.nShowToolBar = CInt(iniUsr.GetValue("ShowToolBar", "Global", -1).Trim)
                With AppInstance.Logger
                    iniUsr.SectionName = "Global"
                    .LogLevel = iniUsr.GetValue("LogLevel").Trim
                    .WriteToEventViewer = iniUsr.GetValue("LogToEventViewer", , "False").Trim
                    .WriteToLogFile = iniUsr.GetValue("LogToFile", , "True").Trim
                    '*** Check for Legacy INI file, where these values where in iniCfg instead
                    If (.LogLevel.ToString = iniUsr.KeyWordNotFound) Then
                        iniCfg.SectionName = "Globals"
                        .LogLevel = iniCfg.GetValue("LogLevel", , "3").Trim
                        .WriteToEventViewer = iniCfg.GetValue("LogToEventViewer", , "False").Trim
                        .WriteToLogFile = iniCfg.GetValue("LogToFile", , "True").Trim
                    End If
                    .LogPath = ICEI0100.IcePaths.DefInstance.LogFilePath
                    .MaxCount = iniCfg.GetValue("MaxCount", "Logger", 0).Trim
                    .MaxSize = iniCfg.GetValue("MaxSize", "Logger", 128000).Trim
                    .MaxDays = iniCfg.GetValue("MaxDays", "Logger", 90).Trim
                End With

            Case 3 'Load MQSeries Settings
                iniCfg.SectionName = "TST1"
                'ZAK: Should the unit be loaded here?
                AppInstance.Unit = "" 'iniCfg.GetValue("Unit")
                AppInstance.MQ_Server = iniCfg.GetValue("Server").Trim
                AppInstance.MQ_ReadQueue = iniCfg.GetValue("ReadQueue").Trim
                AppInstance.MQ_WriteQueue = iniCfg.GetValue("WriteQueue").Trim

            Case 4 'Load MQSeries options and other options
                iniCfg.SectionName = "MQSeries"
                Dim sVal As String = ""

                sVal = iniCfg.GetValue("WaitTaT", , "200").Trim
                If sVal Is Nothing OrElse sVal.Trim = String.Empty OrElse Val(sVal) < 50 Then
                    sVal = "50"
                End If
                AppInstance.gWaitTaT = CInt(sVal)

                sVal = iniCfg.GetValue("TimeOut", , "20000").Trim
                If sVal Is Nothing OrElse sVal.Trim = String.Empty OrElse Val(sVal) < 1000 Then
                    sVal = "1000"
                End If
                AppInstance.MQS.TmeOut = CInt(sVal)
                AppInstance.MQS.SendRetries = iniCfg.GetValue("SendRetries", , "3").Trim
                AppInstance.MQS.GetRetries = iniCfg.GetValue("GetRetries", , "3").Trim

                sVal = iniCfg.GetValue("ReSendDelay", , "1000").Trim
                If sVal Is Nothing OrElse sVal.Trim = String.Empty OrElse Val(sVal) < 1000 Then
                    sVal = "1000"
                End If
                AppInstance.MQS.ResendDelay = CInt(sVal)
                m_useANTSNAAN = CBool(iniCfg.GetValue("UseANTSNAAN", "Globals", "False").Trim)
                ''' RASD fields validations
                'm_RASDValidateNat = CBool(iniCfg.GetValue("RASDValidateNat", "Globals", "False").Trim)
                'm_RASDValidateNIN = CBool(iniCfg.GetValue("RASDValidateNIN", "Globals", "False").Trim)

                '''

                ''' DTS INI variables
                'dtsh.ConfigHelper.DTSConnectionString = iniCfg.GetValue("DTSConnectionString", "DTS", "").Trim
                dtsh.ConfigHelper.DTSLogFilePath = ICEI0100.IcePaths.DefInstance.LogFilePath
                dtsh.ConfigHelper.DTSTemplatePath = ICEI0100.IcePaths.DefInstance.TemplatePath & "\\"
                dtsh.ConfigHelper.DTSDownloadPath = ICEI0100.IcePaths.DefInstance.TemplatePath & "\\"


                ''
#If HR Then
                iniCfg.SectionName = "HOST"
                'G_sSrcDir = ICEI0100.IcePaths.DefInstance.LogFilePath
                G_sSrcDirSAIB = iniCfg.GetValue("PostFilePath", "", " ").Trim
                G_sSrcDirEMS = iniCfg.GetValue("PostFilePathEMS", "", " ").Trim
                If G_sSrcDirSAIB.Trim = "" Then G_sSrcDirSAIB = ICEI0100.IcePaths.DefInstance.LogFilePath
                If G_sSrcDirEMS.Trim = "" Then
                    If G_sSrcDirSAIB.Trim <> "" Then
                        If Not G_sSrcDirSAIB.Trim.EndsWith("\") Then G_sSrcDirSAIB = G_sSrcDirSAIB & "\"
                        G_sSrcDirEMS = G_sSrcDirSAIB
                    Else
                        G_sSrcDirEMS = ICEI0100.IcePaths.DefInstance.LogFilePath
                    End If

                End If

                'G_sSrcDir = "D:\HRDP\Source\"

                sVal = iniUsr.GetValue("PostFilePath", "HOST", " ").Trim
                If sVal.Trim = "" Then
                    iniUsr.SetValue("PostFilePath", G_sSrcDirSAIB, "HOST")
                Else
                    G_sSrcDirSAIB = sVal
                End If

                sVal = iniUsr.GetValue("PostFilePathEMS", "HOST", " ").Trim
                If sVal.Trim = "" Then

                    iniUsr.SetValue("PostFilePathEMS", G_sSrcDirEMS, "HOST")
                Else
                    G_sSrcDirEMS = sVal
                End If


                G_TargetLibrary = "UFIL" & AppInstance.Unit.Trim
                'If AppInstance.Unit = "PRD" Then
                '    G_TargetLibrary = "UFILPRD"
                '    G_TargetLibrary = iniCfg.GetValue("TargetLibrary", "", "UFILPRD").Trim

                'ElseIf AppInstance.Unit = "QRD" Then
                '    G_TargetLibrary = "UFILQRD"
                '    G_TargetLibrary = iniCfg.GetValue("TargetLibrary", "", "UFILQRD").Trim
                'Else
                '    G_TargetLibrary = "UFILEQN"
                '    G_TargetLibrary = iniCfg.GetValue("TargetLibrary", "", "UFILEQN").Trim

                'End If

                'TargetDSN = AS400EQN
                'G_TargetDSN = iniCfg.GetValue("TargetDSN", "", "AS400EQN").Trim
                G_TargetDSN = "AS400" & AppInstance.Unit.Trim
                G_SCMthPostFile = iniCfg.GetValue("MthPostFile", "", "HRDF400").Trim
                G_SCDlyPostFile = iniCfg.GetValue("DlyPostFile", "", "HRDF440").Trim

#End If


                sVal = iniUsr.GetValue("DoubtfulCheckFrequency", "Global", "300000").Trim
                If sVal Is Nothing OrElse sVal.Trim = String.Empty OrElse Val(sVal) > 0 Then
                    m_DoubtfulCheckFreqency = CLng(sVal)
                    'gfrmMain.tmrDoubtful.Interval = m_DoubtfulCheckFreqency
                Else
                    m_DoubtfulCheckFreqency = 0
                End If
                CountFormatFlag = iniCfg.GetValue("PayrollCountFormat", "Globals", "X")
            Case 5
                'load default printing language
                iniUsr.SectionName = "Global"
                Dim str As String

                str = iniUsr.GetValue("PrintLang", , "English")
                If str.ToUpper.StartsWith("A") Then
                    gbPrintingLang = enumReceiptLanguage.Arabic
                End If
                If str.Trim.ToUpper.StartsWith("E") Then
                    gbPrintingLang = enumReceiptLanguage.English
                End If
                iniUsr.SetValue("PrintLang", str, )
        End Select

    End Sub

    Public Function ValidateBranch() As Boolean
        ICEI0100.AppInstanceClass.DefInstance.Logger.LogInfo(0, "Validate Branch --> Started.", 4)
        '''hhhh: Testing
#If DEBUG Then
        'Skip ValidateBranch if command line contains /NoUpdate
        If Environment.CommandLine.ToUpper.IndexOf("/NO_ValidateBranch".ToUpper) >= 0 Then
            ICEI0100.AppInstanceClass.DefInstance.Logger.LogInfo(0, String.Format("Validate Branch --> Debug Mode and flag /NO_ValidateBranch is found so bypassing ValidateBranch function.", IceUserAut.IceEqnBrn), 1)
            Return True
        End If
#End If
        'Get the IP prefix for the local machine
        ICEI0100.AppInstanceClass.DefInstance.Logger.LogInfo(0, String.Format("Validate Branch --> Started for branch [{0}]...", IceUserAut.IceEqnBrn), 1)
        If Not (IceUserAut.IceAutLst(40) Or IceUserAut.IceAutLst(42) Or IceUserAut.IceAutLst(18)) Then
            ICEI0100.AppInstanceClass.DefInstance.Logger.LogInfo(0, String.Format("Validate Branch --> User is not authorized for financial access.", IceUserAut.IceEqnBrn), 1)
            ICEI0100.AppInstanceClass.DefInstance.Logger.LogInfo(0, String.Format("Validate Branch --> Not needed, Exited.", IceUserAut.IceEqnBrn), 1)
            Return False
        End If

        'derive the branch from machine name
        Dim BrnPfx As String = AppInstance.Workstation.Substring(0, 4).ToUpper
        If BrnPfx = "SAIB" Then BrnPfx = "0101"
        Dim ch() As Char = BrnPfx.ToCharArray()
        ch(0) = CChar("0")
        BrnPfx = CStr(ch)

        'Locate the branch
        'Return True
        If BrnPfx <> "0101" Then
            ICEI0100.AppInstanceClass.DefInstance.Logger.LogInfo(0, String.Format("Validate Branch --> Teller services are not required for branch [{0}], only HO branch.", IceUserAut.IceEqnBrn), 1)
            Return False
        End If
        Dim idx As Integer = LocateBrnCfg(IceUserAut.IceEqnBrn)
        If idx < 0 Then
            ICEI0100.AppInstanceClass.DefInstance.Logger.LogInfo(0, String.Format("Validate Branch --> Error: Could not find configurations for branch [{0}].", IceUserAut.IceEqnBrn), 1)
            Return False
        End If
        ICEI0100.AppInstanceClass.DefInstance.Logger.LogInfo(0, String.Format("Validate Branch --> branch [{0}] configuration Found.", IceUserAut.IceEqnBrn), 3)

        If BrnCfgTable(idx).IceBrnSrv.Trim = "" Or BrnCfgTable(idx).BrnSubNet.Trim = "" Then
            ICEI0100.AppInstanceClass.DefInstance.Logger.LogInfo(0, String.Format("Validate Branch --> Branch subnet/server are not defined for branch [{0}], teller services will be disabled.", IceUserAut.IceEqnBrn), 1)
            Return False
        End If
        'do the checking
        'Check for Machine name to belong to the user's branch
        If IceUserAut.IceEqnBrn <> BrnPfx Then
            ICEI0100.AppInstanceClass.DefInstance.Logger.LogInfo(0, String.Format("Validate Branch --> Error: User branch [{0}] is not allowed in the current branch [{1}].", IceUserAut.IceEqnBrn, BrnPfx), 1)
            Return False
        End If
        ICEI0100.AppInstanceClass.DefInstance.Logger.LogInfo(0, String.Format("Validate Branch --> Machine name belongs to the user's branch [{0}].", IceUserAut.IceEqnBrn), 3)

        'check the user's branch to be in the proper subnet
        Dim str_Host As String = System.Net.Dns.GetHostName
        If System.Net.Dns.GetHostByName(str_Host).AddressList.Length = 0 Then
            ICEI0100.AppInstanceClass.DefInstance.Logger.LogInfo(0, "Validate Branch --> Error: Could not obtain IP address.", 1)
            Return False
        End If
        Dim bProperSubNet As Boolean = False
        Dim str_IPPrefix As String
        For ip_cnt As Integer = 0 To System.Net.Dns.GetHostByName(str_Host).AddressList.Length - 1
            Dim str_IP As String = System.Net.Dns.GetHostByName(str_Host).AddressList(ip_cnt).ToString
            str_IPPrefix = str_IP.Split(CChar("."))(0)
            If BrnCfgTable(idx).BrnSubNet.Split(CChar("."))(0) = str_IPPrefix Then
                bProperSubNet = True
                Exit For
            End If
        Next
        If bProperSubNet Then
            ICEI0100.AppInstanceClass.DefInstance.Logger.LogInfo(0, String.Format("Validate Branch --> User's branch [{0}] is in the proper subnet.", IceUserAut.IceEqnBrn), 3)
        Else
            ICEI0100.AppInstanceClass.DefInstance.Logger.LogInfo(0, String.Format("Validate Branch --> Error: Workstation [{0}] does not belong to the current branch [{1}].", AppInstance.Workstation, IceUserAut.IceEqnBrn), 1)
            Return False
        End If
        '''If (str_IP Is Nothing) OrElse (str_IP.Trim = "") Then
        '''    ICEI0100.AppInstanceClass.DefInstance.Logger.LogInfo(0, String.Format("Validate Branch --> Failed for branch [{0}] for not having an IP.", IceUserAut.IceEqnBrn), 1)
        '''    Return False
        '''End If

        'check on Branch Connection & the IceCfgTab has the correct IceEqnBrn
        Try

            'Dim branch As New ICED0100.BranchDB(getBrnTabTableCollection(), False)
            Dim m_BranchDB As ICED0100.BranchDB = FinancialTransaction.clsTransaction.BranchDB

            'Dim row As DataRow
            If m_BranchDB.isConnected = False Then
                ICEI0100.AppInstanceClass.DefInstance.Logger.LogInfo(0, String.Format("Validate Branch --> Error: Workstation is not connected to branch [{0}] database.", IceUserAut.IceEqnBrn), 1)
                Return False
            End If
        Catch ex As Exception
            HandleError(&H81000134, "Validate Branch --> Error in connecting to branch database server." & vbCrLf & ex.Message, "Branch Database")
            Return False
        End Try
        ICEI0100.AppInstanceClass.DefInstance.Logger.LogInfo(0, "Validate Branch --> User successfully connected to branch database.", 2)

        Return True
    End Function

    Public Function ValidateCIBBranch1() As Boolean
        ICEI0100.AppInstanceClass.DefInstance.Logger.LogInfo(0, "Validate CIB Branch --> Started.", 4)
        '''hhhh: Testing

        'check on Branch Connection & the IceCfgTab has the correct IceEqnBrn
        Try

            'Dim branch As New ICED0100.BranchDB(getBrnTabTableCollection(), False)
            Dim m_BranchDB As ICED0100.CIBDB = MakerChecker.BranchDB

            'Dim row As DataRow
            If m_BranchDB.isConnected = False Then
                ICEI0100.AppInstanceClass.DefInstance.Logger.LogInfo(0, String.Format("Validate Branch --> Error: Workstation is not connected to branch [{0}] database.", IceUserAut.IceEqnBrn), 1)
                Return False
            End If
        Catch ex As Exception
            HandleError(&H81000134, "Validate Branch --> Error in connecting to branch database server." & vbCrLf & ex.Message, "Branch Database")
            Return False
        End Try
        ICEI0100.AppInstanceClass.DefInstance.Logger.LogInfo(0, "Validate Branch --> User successfully connected to branch database.", 2)

        Return True
    End Function

    Public Function CreateUniqueFileName() As String
        Static Sequence As Integer
        Sequence += 1
        Dim FileName As String
        Dim dt As Date
        dt = Now
        FileName = "ICE" & Format(dt.Year, "#0000") & Format(dt.Month, "#00") & Format(dt.Day, "#00") & _
                    Format(dt.Hour, "#00") & Format(dt.Minute, "#00") & Format(dt.Second, "#00") & "." & _
                    Format(Sequence, "#000")
        Return FileName
    End Function

    Public Function CreateUniqueFileName(ByVal p_fileName_ As String) As String
        Static Sequence As Integer
        Sequence += 1
        Dim FileName As String
        Dim dt As Date
        dt = Now
        FileName = "ICE" & "_" & p_fileName_ & "_" & Format(dt.Year, "#0000") & Format(dt.Month, "#00") & Format(dt.Day, "#00") & _
                    Format(dt.Hour, "#00") & Format(dt.Minute, "#00") & Format(dt.Second, "#00") & "." & _
                    Format(Sequence, "#000")
        Return FileName
    End Function

    Public Function MarkFileForDeletion(ByVal FileName As String) As Boolean
        ' Function: Mark file for deletion upon reboot. Used mainly for AutoUpdate feature
        ' Input:    FileName: full path of the file name
        If Not IO.File.Exists(FileName) Then Return True
        Dim ret As Integer
        ret = AppInstance.MoveFileEx(FileName, 0, AppInstance.MOVEFILE_DELAY_UNTIL_REBOOT)
        If ret <> 0 Then Return True
        Return False
    End Function

    Private Function GetUpdateServer(ByRef ini As INIClass.IniFileClass, ByVal UnitName As String) As String
        Dim st, pth, srv As String
        Dim WksPrefix As String
        Dim SrvPostFix As String

        'Get the machine name
        pth = ini.GetValue("UpdatePath", UnitName, "\G900.AutoUpdate").Trim
        st = AppInstance.Workstation.ToUpper
        WksPrefix = AppInstance.SafeSubString(st, 0, 4)
        'Use the machine name to deduce update server name from the IceCfg.ini
        srv = ini.GetValue(WksPrefix, "RC_Servers", "xXxxXxx").Trim
        'If failed to locate the machine name prefix within the RC_Servers section, default to ini setting
        'instead of hard code the server postfix S010 or S079 we will read it from Ini settings
        SrvPostFix = "S079"
        SrvPostFix = ini.GetValue("FixedUpdateServer", "Globals", "S079").Trim
        If srv = "xXxxXxx" Then
            srv = "\\" & WksPrefix & SrvPostFix
            'srv = "\\" & WksPrefix & "S010"
            If Dir(srv & pth, FileAttribute.Directory) = "" Then
                srv = ini.GetValue("UpdateServer", UnitName, "G:").Trim
            End If
        End If
        'Special case for TST1 and TST2 and QUA1
        'Select Case UnitName
        '    Case "TST1", "TST2", "QUA1"
        '        srv = ini.GetValue("UpdateServer", UnitName, "G:")
        'End Select
        Return srv & pth
    End Function

    Private Function CheckForUpdate(ByVal UnitName As String, ByRef bRestart As Boolean, ByRef bSkipUpdate As Boolean) As Boolean
        Dim ini As New INIClass.IniFileClass(ICEI0100.IcePaths.DefInstance.IceUsrIni)
        Dim iniCfg As New INIClass.IniFileClass(ICEI0100.IcePaths.DefInstance.IceCfgIni)
        Dim upd As INIClass.IniFileClass
        Dim CurVer As String
        Dim NewVer As String
        Dim OrgFileVer As String
        Dim NewFileVer As String
        Dim ServerPath As String
        Dim CopyBlockList, CopyLineList, CopyFileList, RunFileList As String
        Dim AutoUpdate, AdminUpdate As Boolean
        Dim frm As frmUpdate
        Dim tmp As String
        Dim RollBackCount As Integer = 0
        Dim RollBack() As RollBack_Struct
        Dim bDoVisuals, bUpdated As Boolean

        bRestart = False
        bSkipUpdate = True
        'Return False
        'Skip autoupdate is command line contains /NoUpdate
        If Environment.CommandLine.ToUpper.IndexOf("/NOUPDATE") >= 0 Then
            bSkipUpdate = True
            Return False
        End If

        Try
            AppInstance.Logger.LogInfo(&H80000015, "Checking for update...", 1)

            'iniCfg.SectionName = UnitName
            '''ServerPath = iniCfg.GetValue("UpdateServer", UnitName, "G:\G900.AutoUpdate\ICE")
            ServerPath = GetUpdateServer(iniCfg, UnitName)


            'ini.SectionName = "Versions"
            With System.Diagnostics.FileVersionInfo.GetVersionInfo(System.Reflection.Assembly.GetExecutingAssembly.Location)
                CurVer = ini.GetValue("Current", "Versions", .FileMajorPart & "." & .FileMinorPart & "." & .FileBuildPart).Trim
            End With


            AppInstance.Logger.LogInfo(&H80000016, "Accessing Update Server: " & ServerPath, 1)

            upd = New INIClass.IniFileClass(ServerPath & "\Latest.ini")


            'upd.SectionName = "Image"

            If Not (upd.FileExist) Then
                AppInstance.Logger.LogError(&H80000024, "Could not access update server, folder, or the file 'Latest.ini'", 1)
                'ModalMsgBox("Update server could not be accessed." & vbCrLf & "ICE is unable to determine if any updates are available." & _
                '         vbCrLf & "ICE will continue running using the local version.", _
                '         MsgBoxStyle.Exclamation, "ICE Auto Update")
                ModalMsgBox("Update server could not be accessed." & vbCrLf & "ICE is unable to determine if any updates are available.", MsgBoxStyle.Exclamation, "ICE Auto Update")
                bSkipUpdate = False
                Return False
            End If

            tmp = upd.GetValue("ReDir", "Image").Trim
            If ((tmp <> upd.KeyWordNotFound) And (Trim(tmp) <> "")) Then
                ServerPath = tmp
                upd = New INIClass.IniFileClass(ServerPath & "\Latest.ini")
                AppInstance.Logger.LogInfo(&H80000017, "Redirected to a new update server: " & ServerPath, 1)
                iniCfg.SetValue("UpdateServer", ServerPath, UnitName)
                'upd.SectionName = "Image"
            End If

            If Not (upd.FileExist) Then
                AppInstance.Logger.LogError(&H80000024, "Could not access update server, folder, or the file 'Latest.ini'", 1)
                'ModalMsgBox("Update server could not be accessed." & vbCrLf & "ICE is unable to determine if any updates are available." & _
                '         vbCrLf & "ICE will continue running using the local version.", _
                '         MsgBoxStyle.Exclamation, "ICE Auto Update")
                ModalMsgBox("Update server could not be accessed." & vbCrLf & "ICE is unable to determine if any updates are available.", MsgBoxStyle.Exclamation, "ICE Auto Update")
                bSkipUpdate = False
                Return False
            End If


            bRestart = (AppInstance.SafeSubString(upd.GetValue("Restart", "Image", "False"), 0, 1).ToUpper = "T")
            NewVer = upd.GetValue("Folder", "Image", "").Trim
            If NewVer = upd.KeyWordNotFound Then
                AppInstance.Logger.LogError(&H80000025, "Could not determine the new version on the update server.'", 1)
                'ModalMsgBox("Version on update server could not be determined." & vbCrLf & "ICE is unable to perform any updates on your computer." & _
                '        vbCrLf & "ICE will continue running using the local version.", _
                '        MsgBoxStyle.Exclamation, "ICE Auto Update")
                ModalMsgBox("Version on update server could not be determined." & vbCrLf & "ICE is unable to perform any updates on your computer.", MsgBoxStyle.Exclamation, "ICE Auto Update")
                bSkipUpdate = False
                Return False
            End If

            If NewVer <> CurVer Then
                AppInstance.Logger.LogInfo(&H80000018, "New binary update detected: Current:" & CurVer & ", New:" & NewVer, 2)
                frm = New frmUpdate
                frm.pnlInfoBeg.Show()
                frm.pnlInfoBeg.BringToFront()
                frm.lblInfoCurBeg.Text = CurVer
                frm.lblInfoNewBeg.Text = NewVer
                frm.lblDscBeg.Text = "Updates are available for your computer (related to unit " & AppInstance.Unit & " on " & AppInstance.SafeSubString(AppInstance.MQ_Server, 3, 4) & "). Press OK to apply these changes."
                frm.ShowDialog()
                'If frm.DialogResult = DialogResult.Ignore Then
                '    AppInstance.Logger.LogWarn(&H90000001, "User Cancelled AutoUpdate from" & CurVer & " to New:" & NewVer, 2)
                '    bSkipUpdate = True
                '    Return False
                'End If

                If frm.DialogResult <> DialogResult.OK Then
                    AppInstance.Logger.LogWarn(&H90000001, "User Cancelled AutoUpdate from" & CurVer & " to New:" & NewVer, 2)
                    bSkipUpdate = False
                    Return False
                End If
                frm.pnlProgress.Show()
                frm.pnlProgress.BringToFront()
                frm.Show()
                bDoVisuals = True
            Else
                AppInstance.Logger.LogInfo(&H80000023, "No major binary updates found.", 1)
            End If

            AutoUpdate = upd.GetValue("Auto", "Image", "True").Trim
            AdminUpdate = upd.GetValue("Admin", "Image", "False").Trim
            If AutoUpdate And Not (AdminUpdate) Then
                'Do the Update
                'frm = New frmUpdate
                If bDoVisuals Then
                    frm.lblCurVer.Text = CurVer
                    frm.lblNewVer.Text = NewVer
                End If
                '********
                '* Run the updater programs before copy
                '********
                ' IH only do this if the version has changed (not every time)
                If NewVer <> CurVer Then
                    RunFileList = upd.GetValue("RunBefore", "Image", "").Trim
                    If ((RunFileList.Trim <> String.Empty) And (RunFileList.ToUpper <> "Keyword not found".ToUpper)) Then
                        Dim FileList() As String = RunFileList.Split(",")
                        For i As Integer = 0 To FileList.Length - 1
                            Try
                                AppInstance.Logger.LogInfo(&H80000019, "Copying and executing file " & filelist(i), 3)
                                FileCopy(ServerPath & "\" & NewVer & "\" & FileList(i), Application.StartupPath & "\" & FileList(i))
                                Shell(Application.StartupPath & "\" & FileList(i), AppWinStyle.MinimizedNoFocus, True, -1)
                                Kill(Application.StartupPath & "\" & FileList(i))
                            Catch ex As Exception
                                ' IH Only log message as advised by KM
                                AppInstance.Logger.LogInfo(&H80000020, "Could not run update utility [" & FileList(i) & "]" & vbCrLf & _
                                       "Error:" & ex.Message, 3)
                                'ModalMsgBox("Could not run update utility [" & FileList(i) & "]" & vbCrLf & _
                                '       "Error:" & ex.Message & vbCrLf & "Please contact Technical Support before attempting to use ICE again.", _
                                '       MsgBoxStyle.Exclamation, "ICE Update")
                            End Try
                        Next
                    End If
                End If

                '********
                '* Copy updated files from the server to the local machine
                '********

                CopyBlockList = upd.GetValue("CopyBlock", "Image", " ").Trim
                If (CopyBlockList.Trim <> String.Empty) Then
                    Dim BlockList() As String = CopyBlockList.Split(",")

                    For j As Integer = 0 To BlockList.Length - 1
                        Dim DefaultDir As String
                        DefaultDir = upd.GetValue("DefaultDir", BlockList(j), " ").Trim
                        '* Create the target directory if it does not exist
                        Try
                            If DefaultDir.Trim <> String.Empty Then
                                If Not System.IO.Directory.Exists(Application.StartupPath & "\" & DefaultDir) Then
                                    System.IO.Directory.CreateDirectory(Application.StartupPath & "\" & DefaultDir)
                                End If
                            End If
                        Catch
                        End Try

                        CopyLineList = upd.GetValue("CopyLines", BlockList(j), " ").Trim
                        If Trim(CopyLineList) <> "" Then
                            Dim LineList() As String = CopyLineList.Split(",")
                            For k As Integer = 0 To LineList.Length - 1

                                CopyFileList = upd.GetValue(LineList(k), BlockList(j), " ").Trim
                                If Trim(CopyFileList) <> "" Then
                                    Dim FileList() As String = CopyFileList.Split(",")
                                    Dim UnqFile, FleNam As String
                                    If bDoVisuals Then
                                        frm.barProgress.Maximum = FileList.Length
                                        frm.barProgress.Value = 0
                                        Application.DoEvents()
                                    End If
                                    For i As Integer = 0 To FileList.Length - 1
                                        FleNam = DefaultDir & FileList(i)
                                        If bDoVisuals Then
                                            frm.lblSrcFil.Text = FleNam
                                            frm.lblDstFil.Text = Application.StartupPath & "\" & FleNam
                                        End If
                                        tmp = Application.StartupPath & "\" & FleNam
                                        Try
                                            NewFileVer = upd.GetValue(filelist(i), "Versions", "0.0.0").Trim
                                            OrgFileVer = ini.GetValue(filelist(i), "Versions", " ").Trim
                                            If NewFileVer = ini.GetValue(filelist(i), "Versions", "x.x.x").Trim Then
                                                'AppInstance.Logger.LogInfo(&H80000020, "Skipping file " & FleNam & ", no update needed.", 3)
                                                Exit Try
                                            End If
                                            Try
                                                If Dir(tmp) <> "" Then
                                                    ReDim Preserve RollBack(RollBackCount)
                                                    UnqFile = ""
                                                    UnqFile = Application.StartupPath & "\Versions\" & CreateUniqueFileName()
                                                    RollBack(RollBackCount).OrgFilSrc = tmp
                                                    RollBack(RollBackCount).NewFileSrc = unqfile
                                                    RollBack(RollBackCount).FileVer = OrgFileVer
                                                    RollBack(RollBackCount).FileName = filelist(i)
                                                    Rename(tmp, UnqFile)
                                                    MarkFileForDeletion(UnqFile)
                                                    RollBackCount += 1
                                                End If
                                                AppInstance.Logger.LogInfo(&H80000020, "Updating file " & FleNam & " from version " & OrgFileVer & " to  version " & NewFileVer, 3)
                                                AppInstance.ShowStatusText("Updating file " & FleNam & " from version " & OrgFileVer & " to  version " & NewFileVer)
                                                FileCopy(ServerPath & "\" & NewVer & "\" & FleNam, Application.StartupPath & "\" & FleNam)
                                                ini.SetValue(filelist(i), NewFileVer, "Versions")
                                                bUpdated = True
                                            Catch ex1 As Exception
                                                AppInstance.Logger.LogError(&H80000026, "Failed while copying file " & FleNam, 1)
                                                AppInstance.Logger.LogError(&H80000026, "Rolling back...", 1)
                                                ModalMsgBox("Could not update file [" & Application.StartupPath & "\" & FleNam & "]" & vbCrLf & _
                                                    "Error:" & ex1.Message & vbCrLf & vbCrLf & "ICE will now roll back the installed updates. You will continue with your current version of ICE", _
                                                    MsgBoxStyle.Exclamation, "ICE Update")
                                                'Rollback files
                                                If bDoVisuals Then
                                                    frm.lblNewVer.Text = "Rolling back..."
                                                    frm.barProgress.Maximum = RollBackCount
                                                    frm.barProgress.Value = 0
                                                    Application.DoEvents()
                                                End If
                                                Try
                                                    For n As Integer = 0 To RollBackCount
                                                        AppInstance.Logger.LogError(&H80000026, "Rolling back " & RollBack(n).OrgFilSrc)
                                                        If bDoVisuals Then
                                                            frm.lblDstFil.Text = RollBack(n).OrgFilSrc
                                                            Application.DoEvents()
                                                            Sleep(300)
                                                            frm.barProgress.Value = RollBackCount - n
                                                        End If
                                                        FileCopy(RollBack(n).NewFileSrc, RollBack(n).OrgFilSrc)
                                                        ini.SetValue(RollBack(n).FileName, RollBack(n).FileVer, "Versions")
                                                    Next
                                                Catch 'nothing, ignore error
                                                End Try
                                                If bDoVisuals Then
                                                    frm.Close()
                                                    frm = Nothing
                                                    ModalMsgBox("ICE was unable to apply the new upates." & vbCrLf & _
                                                                       "You will continue to use your current version of ICE," & vbCrLf & _
                                                                       "but you will not have access to the latest fixes and functionalities." & vbCrLf & _
                                                                       vbCrLf & _
                                                                       "Please contact technical support to resolve the issue.", MsgBoxStyle.Exclamation, _
                                                                       "ICE Update")
                                                End If
                                                bSkipUpdate = False
                                                Return False
                                                'If UnqFile <> "" Then
                                                '    FileCopy(UnqFile, tmp)
                                                'End If
                                            End Try
                                        Catch ex As Exception
                                            AppInstance.Logger.LogError(&H80000026, "Failed while copying file " & FleNam, 1)
                                            ModalMsgBox("Could not update file [" & Application.StartupPath & "\" & FleNam & "]" & vbCrLf & _
                                                "Error:" & ex.Message & vbCrLf & "Please contact Technical Support before attempting to use ICE again.", _
                                                MsgBoxStyle.Exclamation, "ICE Update")
                                            Return True
                                        End Try
                                        If bDoVisuals Then
                                            frm.barProgress.Value = i + 1
                                            Application.DoEvents()
                                            Sleep(100)
                                        End If
                                    Next
                                End If
                            Next k
                        End If

                    Next j
                Else
                    'File might not have a CopyBlock entry (an old image)
                    CopyFileList = upd.GetValue("Copy", "Image", "").Trim
                    If (CopyFileList.Trim <> String.Empty) Then
                        Dim FileList() As String = CopyFileList.Split(",")
                        Dim UnqFile As String
                        If bDoVisuals Then
                            frm.barProgress.Maximum = FileList.Length
                            frm.barProgress.Value = 0
                        End If
                        Application.DoEvents()
                        For i As Integer = 0 To FileList.Length - 1
                            If bDoVisuals Then
                                frm.lblSrcFil.Text = FileList(i)
                                frm.lblDstFil.Text = Application.StartupPath & "\" & FileList(i)
                            End If
                            Try
                                If Dir(Application.StartupPath & "\" & FileList(i)) <> "" Then
                                    UnqFile = Application.StartupPath & "\Versions\" & CreateUniqueFileName()
                                    Rename(Application.StartupPath & "\" & FileList(i), UnqFile)
                                    MarkFileForDeletion(UnqFile)
                                End If
                                AppInstance.Logger.LogInfo(&H80000023, "Copying file " & filelist(i), 3)
                                FileCopy(ServerPath & "\" & NewVer & "\" & FileList(i), Application.StartupPath & "\" & FileList(i))
                            Catch ex As Exception
                                ModalMsgBox("Could not update file [" & Application.StartupPath & "\" & FileList(i) & "]" & vbCrLf & _
                                       "Error:" & ex.Message & vbCrLf & "Please contact Technical Support before attempting to use ICE again.", _
                                       MsgBoxStyle.Exclamation, "ICE Update")
                            End Try
                            If bDoVisuals Then
                                frm.barProgress.Value = i + 1
                                Application.DoEvents()
                                Sleep(100)
                            End If
                        Next
                    End If
                End If


                '********
                '* Run the CopyReg Block
                '********
                'If NewVer <> CurVer Then
                CopyRegistryFiles(upd, ini, ServerPath, NewVer, frm)
                'End If


                '********
                '* Run the updater programs after copy
                '********
                RunFileList = upd.GetValue("RunAfter", "Image", "").Trim
                If ((RunFileList.Trim <> String.Empty) And (RunFileList.ToUpper <> "Keyword not found".ToUpper)) Then
                    Dim FileList() As String = RunFileList.Split(",")
                    For i As Integer = 0 To FileList.Length - 1
                        Try
                            AppInstance.Logger.LogInfo(&H80000021, "Copying and executing file " & filelist(i), 3)
                            FileCopy(ServerPath & "\" & NewVer & "\" & FileList(i), Application.StartupPath & "\" & FileList(i))
                            Shell(Application.StartupPath & "\" & FileList(i), AppWinStyle.NormalFocus, True, -1)
                            Kill(Application.StartupPath & "\" & FileList(i))
                        Catch ex As Exception
                            ModalMsgBox("Could not run update utility [" & FileList(i) & "]" & vbCrLf & _
                                        "Error:" & ex.Message & vbCrLf & "Please contact Technical Support before attempting to use ICE again.", _
                                        MsgBoxStyle.Exclamation, "ICE Update")
                        End Try
                    Next
                End If

                '********
                '* Update the version stamp on the local machine
                '********
                If bUpdated Then
                    AppInstance.Logger.LogInfo(&H80000022, "Update done", 1)
                    AppInstance.ShowStatusText("Update done")
                End If
                If bDoVisuals Then
                    ini.SetValue("Current", NewVer, "Versions")
                    If bRestart Then
                        Dim RestartForm As New frmRestart
                        RestartForm.Location = frm.Location
                        frm.Close()
                        frm = Nothing
                        RestartForm.TopMost = True
                        RestartForm.ShowDialog()
                    Else
                        frm.Hide()
                        frm.pnlInfoEnd.Show()
                        frm.pnlInfoEnd.BringToFront()
                        frm.lblInfoCurEnd.Text = CurVer
                        frm.lblInfoNewEnd.Text = NewVer
                        frm.lblDscEnd.Text = "ICE Auto Update Completed. ICE application must be restarted for changes to take effect. ICE will now restart. "
                        frm.ShowDialog()
                        frm = Nothing
                    End If
                    Return True
                Else
                    Return False
                End If
            ElseIf AdminUpdate Then
                Dim Ans As MsgBoxResult
                Ans = ModalMsgBox("A new update is available, but it requires an Administrator to install it." & vbCrLf & _
                        "Please contact your Techinical Support to perform this upgrade." & vbCrLf & _
                        "The software might not run properly without the update, continue anyway?", MsgBoxStyle.YesNo)
                If Ans = MsgBoxResult.No Then Return True
            Else 'Manual user update
                Dim Ans As MsgBoxResult
                Ans = ModalMsgBox("A new update is available, but it requires authorization to install it." & vbCrLf & _
                        "Please contact your Techinical Support for authorization code." & vbCrLf & _
                        "Continue with update?", MsgBoxStyle.YesNo)
                If Ans = MsgBoxResult.Yes Then
                    ModalMsgBox("A dialog will popup here asking your for the authorization code <NYI>")
                End If
            End If
        Catch ex_1 As Exception
            HandleError(&H80004020, "ICE auto update couldn't continue: " & ex_1.Message, "ICE Auto Update")
            If Not (frm Is Nothing) Then
                frm.Hide()
                frm = Nothing
            End If
        End Try
        Return False
    End Function

    Private Sub CopyRegistryFiles(ByVal upd As INIClass.IniFileClass, ByVal ini As INIClass.IniFileClass, ByVal ServerPath As String, ByVal NewVer As String, ByVal frm As frmUpdate)
        Dim sRegKeys As String
        Dim arrRegKeys() As String
        Dim sRegKey As String
        Dim sUPDTempSection As String
        Dim sIniTempSection As String
        Dim RollBack(0) As RollBack_Struct

        'If NewVer <> CurVer Then
        Try
            'AppInstance.Logger.LogInfo(&H80000000, "Proccessing ini block [CopyReg] --> Entered...")
            sUPDTempSection = upd.SectionName
            sIniTempSection = ini.SectionName
            upd.SectionName = "CopyReg"
            sRegKeys = upd.GetValue("CopyFiles").Trim
            If sRegKeys = String.Empty OrElse sRegKeys = upd.KeyWordNotFound Then
                AppInstance.Logger.LogInfo(&H80000000, "Proccessing ini block [CopyReg] --> Exiting with nothing to proccess.")
                Exit Sub
            End If
            arrRegKeys = sRegKeys.Split(",".ToCharArray)
            'AppInstance.Logger.LogInfo(&H80000000, String.Format("Proccessing ini block [CopyReg] --> Started for [{0}] registry keys...", arrRegKeys.Length))
            Dim iCount As Integer = 0
            Dim NewFileVer As String
            Dim OrgFileVer As String

            For Each sFile As String In arrRegKeys
                NewFileVer = upd.GetValue(sFile, "Versions", "0.0.0").Trim
                OrgFileVer = ini.GetValue(sFile, "Versions", " ").Trim

                If NewFileVer <> OrgFileVer Then
                    sRegKey = upd.GetValue(sFile).Trim
                    '  If sRegKey <> String.Empty And sRegKey <> upd.KeyWordNotFound Then
                    iCount += 1
                    'AppInstance.Logger.LogInfo(&H80000000, String.Format("Proccessing ini block [CopyReg] --> Proccessing key #({0}) [{1}]...", iCount, sRegKey))
                    CopyRegistryFile(upd, ini, ServerPath, NewVer, sFile, sRegKey, frm, RollBack)
                    'AppInstance.Logger.LogInfo(&H80000000, String.Format("Proccessing ini block [CopyReg] --> Proccessing finished for key #({0}) [{1}]...", iCount, sRegKey))
                    'Else
                    '    AppInstance.Logger.LogError(&H80000000, String.Format("Proccessing ini block [CopyReg] --> File [{0}] did not have any regkey in the .ini file.", sFile))
                    'End If
                End If
            Next
            'AppInstance.Logger.LogInfo(&H80000000, "Proccessing ini block [CopyReg] --> End Successfully.")
            'Throw New Exception("") 'test if an exception happened
        Catch ex As Exception
            upd.SectionName = sUPDTempSection
            'Debug.WriteLine(  "CopyRegistryFiles: "   & ": " & ex.ToString)
            AppInstance.Logger.LogError(&H81000191, "Could not process ini block [CopyReg]." & vbCrLf & ex.Message, 1)

            Try
                For n As Integer = 0 To RollBack.Length - 1
                    AppInstance.Logger.LogError(&H81000192, "Rolling back " & RollBack(n).OrgFilSrc)
                    FileCopy(RollBack(n).NewFileSrc, RollBack(n).OrgFilSrc)
                    ini.SetValue(RollBack(n).FileName, RollBack(n).FileVer, "Versions")
                Next
            Catch 'nothing, ignore error
            End Try
        End Try
        upd.SectionName = sUPDTempSection
        ini.SectionName = sIniTempSection
    End Sub

    Private Sub CopyRegistryFile(ByVal upd As INIClass.IniFileClass, ByVal ini As INIClass.IniFileClass, ByVal ServerPath As String, ByVal NewVer As String, ByVal sFile As String, ByVal sRegKey As String, ByVal frm As frmUpdate, ByRef RollBack() As RollBack_Struct)
        Dim RegKeyMainNode As Microsoft.Win32.RegistryKey
        Dim RegKey As Microsoft.Win32.RegistryKey
        Dim sDestinationPath As String
        Dim arrKeyValuePair() As String
        Dim sKey As String
        Dim sSubKey As String
        Dim NewFileVer As String
        Dim OrgFileVer As String
        'Dim RollBackCount As Integer = 0
        Dim isRegKeyFound As Boolean = False


        Dim sSourcePath As String = ServerPath & "\" & NewVer & "\" & sFile 'this should be specified

        Try
            'AppInstance.Logger.LogInfo(&H81000193, String.Format("Proccessing ini block [CopyReg] --> Started file [{0}]...", sSourcePath))
            arrKeyValuePair = sRegKey.Split(",".ToCharArray)
            sKey = nz(arrKeyValuePair(0)).Trim
            sSubKey = nz(arrKeyValuePair(1)).Trim

            isRegKeyFound = SIBL0100.URegistry.getRegKey(sKey, sSubKey, sDestinationPath)

            If isRegKeyFound = False OrElse sDestinationPath Is Nothing OrElse sDestinationPath = String.Empty Then
                AppInstance.Logger.LogError(&H81000194, String.Format("Registry key [{0}] not found. File [{1}] will be copyed to local directory.", sKey & sSubKey, sSourcePath))
            End If

            If (sDestinationPath Is Nothing) OrElse sDestinationPath.Trim = String.Empty Then
                sDestinationPath = System.Windows.Forms.Application.StartupPath & "\"
                AppInstance.Logger.LogInfo(&H80000000, String.Format("Registry key [{0}] was not found or empty, so copying to application path [{1}].", sRegKey, sDestinationPath))
            Else
                'AppInstance.Logger.LogInfo(&H80000000, String.Format("Registry key found [{0}] with value [{1}].", sRegKey, sDestinationPath))
                If (sDestinationPath.ToUpper.IndexOf(sFile.ToUpper) < 0) Then
                    AppInstance.Logger.LogInfo(&H80000000, String.Format("File string [{0}] was not found in registry key [{1}].", sFile, sRegKey))
                End If
            End If
            sDestinationPath = sDestinationPath.Substring(0, sDestinationPath.LastIndexOf("\"))
            sDestinationPath = sDestinationPath & "\" & sFile

            If Not (frm Is Nothing) AndAlso frm.Visible = True Then
                frm.lblSrcFil.Text = sSourcePath
                frm.lblDstFil.Text = sDestinationPath
                Application.DoEvents()
            End If

            NewFileVer = upd.GetValue(sFile, "Versions", "0.0.0").Trim
            OrgFileVer = ini.GetValue(sFile, "Versions", " ").Trim
            If IO.File.Exists(sDestinationPath) Then  'Dir(sDestinationPath) <> "" Then
                If RollBack(RollBack.Length - 1).NewFileSrc Is Nothing OrElse RollBack(RollBack.Length - 1).NewFileSrc.Trim = String.Empty Then
                    ReDim RollBack(0)
                Else
                    ReDim Preserve RollBack(RollBack.Length)
                End If
                RollBack(RollBack.Length - 1).NewFileSrc = Application.StartupPath & "\Versions\" & CreateUniqueFileName(sFile)
                RollBack(RollBack.Length - 1).OrgFilSrc = sDestinationPath
                RollBack(RollBack.Length - 1).FileVer = OrgFileVer
                RollBack(RollBack.Length - 1).FileName = sFile
                Rename(sDestinationPath, RollBack(RollBack.Length - 1).NewFileSrc)
                'System.IO.File.Move(sDestinationPath, RollBack(RollBack.Length - 1).NewFileSrc)
                MarkFileForDeletion(RollBack(RollBack.Length - 1).NewFileSrc)
            Else
                ''
            End If

            AppInstance.Logger.LogInfo(&H80000000, String.Format("File string [{0}] in IceUsr.ini has been set from version#[{1}] --> to version#[{2}].", sFile, OrgFileVer, NewFileVer))

            'Dim sTmp As String
            'sTmp = sSourcePath.Substring(0, 1)
            'sSourcePath = sTmp & sSourcePath.Substring(1).Replace("\\", "\")

            'sTmp = sDestinationPath.Substring(0, 1)
            'sDestinationPath = sTmp & sDestinationPath.Substring(1).Replace("\\", "\")

            Microsoft.VisualBasic.FileSystem.FileCopy(sSourcePath, sDestinationPath)

            'System.IO.File.Copy(sSourcePath, sDestinationPath, True)
            AppInstance.Logger.LogInfo(&H80000000, String.Format("Proccessing ini block [CopyReg] --> Copyed Successfully from [{0}] to [{1}].", sSourcePath, sDestinationPath))
            ini.SetValue(sFile, NewFileVer, "Versions")

        Catch ex As Exception
            AppInstance.Logger.LogError(&H81000195, String.Format("Proccessing ini block [CopyReg] -->  Failed to copy file [{0}] to [{1}].{2}", sSourcePath, sDestinationPath, ex.Message))
            Throw  'Exception need to be propagated to caller in order to perform the rollback at the caller's level
        End Try
    End Sub

    Public Function HideCustomerFinData_old_Progressive(ByRef CusDat As frmCustomer.CusDta_Struct, Optional ByVal StartLevel As String = "0") As Boolean
        'Return ((CusDat.CusDtl.CusInf.CusSnsLev >= "3") And Not (IceUserAut.IceAutLst(293))) Or _
        '           ((CusDat.CusDtl.CusInf.CusSnsLev = "2") And Not (IceUserAut.IceAutLst(293) Or IceUserAut.IceAutLst(292))) Or _
        '           ((CusDat.CusDtl.CusInf.CusSnsLev = "1") And Not (IceUserAut.IceAutLst(293) Or IceUserAut.IceAutLst(292) Or IceUserAut.IceAutLst(291)))

        'If customer sensitivity level is below the threshold, then don't hide any data
        If CusDat.CusDtl.CusInf.CusSnsLev < StartLevel Then Return False

        'Bit 294 overrides the security restrictions
        If (IceUserAut.IceAutLst(294)) Then Return False

        Select Case CusDat.CusDtl.CusInf.CusSnsLev
            Case "3"
                'bit 293 is needed to allow seeing full customer information at level 3
                If Not (IceUserAut.IceAutLst(293)) Then Return True
            Case "2"
                'bit 293 or 292 are needed to allow seeing full customer information at level 2 or at an unspecified level
                If Not (IceUserAut.IceAutLst(293) Or IceUserAut.IceAutLst(292)) Then Return True
            Case "1"
                'bit 293, 292, or 291 are needed to allow seeing full customer information at level 1
                If Not (IceUserAut.IceAutLst(293) Or IceUserAut.IceAutLst(292) Or IceUserAut.IceAutLst(291)) Then Return True
            Case "", "0"
                'no bits are needed to allow seeing full customer information at level 0 (default level)
                Return False
            Case Else 'Case "2"
                'if customer is at an unspecified level, then it is considered to be level 2
                If Not (IceUserAut.IceAutLst(293) Or IceUserAut.IceAutLst(292)) Then Return True
        End Select
        Return False
    End Function

    Public Function HideCustomerFinData(ByRef CusDat As frmCustomer.CusDta_Struct, Optional ByVal SpecificLevel As String = "") As Boolean
        Return HideCustomerFinData(CusDat.CusDtl.CusInf.CusSnsLev, SpecificLevel)
    End Function

    Public Function HideCustomerFinData(ByRef CusSnsLev As String, Optional ByVal SpecificLevel As String = "") As Boolean
        'Bit 294 overrides the security restrictions
        If (IceUserAut.IceAutLst(294)) Then Return False

        'If customer sensitivity level is below the threshold, then don't hide any data
        If (SpecificLevel <> "") AndAlso (CusSnsLev <> SpecificLevel) Then Return False

        Select Case CusSnsLev.Trim
            Case "3"
                If Not (IceUserAut.IceAutLst(293)) Then Return True
            Case "2"
                If Not (IceUserAut.IceAutLst(292)) Then Return True
            Case "1"
                If Not (IceUserAut.IceAutLst(291)) Then Return True
            Case "", "0"
                'no bits are needed to allow seeing full customer information at level 0 (default level)
                Return False
            Case Else 'Case "2"
                Return True
        End Select
        Return False
    End Function
    '''
    'Added By: Y309ABSO
    'Date: 2013-11-27
    'Purpose: To hide customer balance field on SER account transfer and account SADAD payment (pre, post and MOI)
    '''
    Public Function HideCustomerBalance(ByVal AccNum As String, ByVal CusNum As String) As Boolean
        Dim CusSearchResults() As IceCusItm_Struct
        Dim i As Integer
        Dim frm As frmCusSelect
        Dim LocItm As New structLocCusInf
        Dim Result As Boolean


        ShowBusyIcon(True)

        LocItm.AccNum = "" 'AccNum
        LocItm.CusNum = AppInstance.PackString(CusNum, 6)
        Erase CusSearchResults

        Dim MsgId As String = SendMessage_LocCusInf(LocItm)
        If IsValidMid(MsgId) Then
            Result = GetMessage_LocCusInf(MsgId, CusSearchResults)
        End If

        If Not (Result) Then 'A problem occurred
            ShowBusyIcon()
            Return False
        End If

        If CusSearchResults Is Nothing OrElse CusSearchResults.Length <= 0 Then
            ShowStatusText("No Records Found!")
            ModalMessageBox("Your search did not return any customer matching [" & CusNum & "]", , , MessageBoxIcon.Information, "Locate Customer")
            ShowBusyIcon()
            Return False
        End If
        ShowBusyIcon()
        Dim HideFinData As Boolean = False

        If Not CusSearchResults(0).CusSnsLev Is Nothing Then
            HideFinData = HideCustomerFinData(CusSearchResults(0).CusSnsLev.Trim)
            If HideFinData Then
                Return True
            Else
                Return False
            End If
        Else
            Return True
        End If

        ShowBusyIcon()
        Return False
    End Function

    Public Sub ShowPrintForm()
        If Not (gfrmPrint Is Nothing) Then
            gfrmPrint.Close()
            gfrmPrint.Dispose()
        End If
        gfrmPrint = New frmPrint
        AppInstance.gfrmPrint = gfrmPrint
        gfrmPrint.Owner = gfrmMain
        gfrmPrint.Show()
        gfrmPrint.TopLevel = True
    End Sub

    Public Sub CreatePrintForm()
        If (gfrmPrint Is Nothing) Then
            gfrmPrint = New frmPrint
            AppInstance.gfrmPrint = gfrmPrint
            gfrmPrint.Owner = gfrmMain
            gfrmPrint.Show()
            gfrmPrint.TopLevel = True
        End If
        gfrmMain.tmrPrintDialog.Start()
    End Sub

    Public Sub ClosePrintForm()
        If Not (gfrmPrint Is Nothing) Then
            gfrmPrint.Close()
        End If
        gfrmPrint = Nothing
        AppInstance.gfrmPrint = gfrmPrint
    End Sub

    Public Sub ShowPrintFormEx(ByRef PrnObj As Object)
        Try
            gfrmPrint.m_PrintThread = PrnObj.PrintThread
            gfrmPrint.Show()
            gfrmPrint.TopLevel = True
        Catch ex As Exception
            Throw
        Finally
            ShowBusyIcon()
        End Try
    End Sub

    Public Sub HidePrintForm()
        If Not (gfrmPrint Is Nothing) Then
            gfrmPrint.m_PrintThread = Nothing
            gfrmPrint.Hide()
        End If
        gfrmMain.tmrPrintDialog.Stop()

        'Dim a As New SetFromThread(gfrmPrint, gfrmMain.tmrPrintDialog)
        'a.StopTimer()
    End Sub


    Public Class SetFromThread
        Public Frm As Form
        Private _gfrmPrint As frmPrint
        Private _timer As Windows.Forms.Timer

        Sub New(gfrmPrint As frmPrint, timer As Windows.Forms.Timer)
            ' TODO: Complete member initialization 
            _gfrmPrint = gfrmPrint
            _timer = timer
        End Sub


        Public Sub StopTimer()
            If Frm.InvokeRequired Then
                Try : Frm.Invoke(New MethodInvoker(AddressOf StopTimer)) : Catch : End Try
            Else
                _timer.Stop()
            End If
        End Sub

    End Class

    Public Function PrevInstance(ByVal bBringToFront As Boolean, ByVal bWaitForExit As Boolean) As Boolean
        Try
            Dim Proc As String = Process.GetCurrentProcess.ProcessName
            Dim Processes() As Process = Process.GetProcessesByName(Proc)
            If Processes.Length > 1 Then
                Dim p As Process = Process.GetCurrentProcess
                Dim n As Integer = 0    'assume the other process is at index 0
                If (Processes(0).Id = p.Id) Then n = 1 'then the other process is at index 1
                If bBringToFront Then
                    Dim hWnd As IntPtr = Processes(n).MainWindowHandle
                    If (AppInstance.IsIconic(hWnd)) Then AppInstance.ShowWindowAsync(hWnd, AppInstance.SW_RESTORE)
                    AppInstance.SetForegroundWindow(hWnd)
                End If
                If bWaitForExit Then
                    Return Not (Processes(n).WaitForExit(32100))
                End If
                Return True 'Mutiple instances
            Else
                Return False 'This is the first isntance
            End If
        Catch ex As Exception
            ModalMsgBox("Unrecoverable error occurred while checking for ICE Previous Instance!" & vbCrLf & ex.Message & _
                    vbCrLf & "Application Cannot continue and must shut down.", MsgBoxStyle.Critical)
            Return True
        End Try
    End Function

    Public Function IsMQSeriesInstalled() As Boolean
        Dim RegKey As RegistryKey
        Try
            RegKey = Registry.LocalMachine.OpenSubKey("Software\IBM\MQSeries\CurrentVersion")
            If RegKey Is Nothing Then
                RegKey = Registry.LocalMachine.OpenSubKey("SOFTWARE\IBM\WebSphere MQ\Installation\Installation1")
            End If
            'RegKey = Registry.LocalMachine.OpenSubKey("Software\IBM\MQSeries\CurrentVersion")
            'HKEY_LOCAL_MACHINE\SOFTWARE\IBM\WebSphere MQ\Installation\Installation1
            RegKey.GetValue("FilePath")
        Catch
            Return False
        End Try
        Return True
    End Function

    Public Function ShowMQErrorLog() As Boolean
        'HKEY_LOCAL_MACHINE\SOFTWARE\IBM\MQSeries\CurrentVersion
        Dim RegKey As RegistryKey
        Dim FilePath As String
        Try
            RegKey = Registry.LocalMachine.OpenSubKey("Software\IBM\MQSeries\CurrentVersion")
            FilePath = RegKey.GetValue("FilePath")
            If Not FileExist(FilePath & "\Errors\AMQERR01.LOG") Then
                ModalMsgBox("Could not locate the Log file " & FilePath & "\Errors\AMQERR01.LOG", MsgBoxStyle.Exclamation, "MQSeries Log")
            Else
                Shell("notepad " & FilePath & "\Errors\AMQERR01.LOG", AppWinStyle.NormalFocus, False)
                'Sleep(1000)
                SendKeys.Send("^{END}")
            End If
        Catch ex As Exception
            HandleError(&H80004021, "could not show MQSeries Error Log: " & ex.Message, "MQ Series Error Log")
            Return False
        End Try
        Return True
    End Function

    'Public Sub HandleError(ByVal ErrorNum As Integer, ByVal ErrStr As String, Optional ByVal ErrTtl As String = "")
    '    AppInstance.HandleError(ErrorNum, ErrStr, ErrTtl)
    'End Sub

    Public Sub HandleError(ByVal ErrorNum As Integer, ByVal ErrStr As String, Optional ByVal ErrTtl As String = "", Optional ByVal ex_ As Exception = Nothing)
        If ErrTtl Is Nothing Then ErrTtl = ""
        If ErrTtl = "Error" Then ErrTtl = "" 'Special handling for hussein's change to all HandleError calls where the "Error" was added to title

        If Not ex_ Is Nothing Then
            If (Not ErrStr Is Nothing) Then
                If (ErrStr.IndexOf(ex_.Message) < 0) Then ErrStr &= ": " & ex_.Message
            Else
                ErrStr = ex_.Message
            End If
        End If

        Dim oFram As StackFrame
        Dim DbgStr As String
        If ErrTtl.Trim = "" Then
            oFram = New StackTrace(True).GetFrame(1)
            DbgStr = SIBL0100.Util.Debug.getStackFrame(oFram)
        End If
#If DEBUG Then
        If Not (DbgStr Is Nothing) AndAlso (DbgStr.Trim <> "") Then DbgStr = ":" & DbgStr
        AppInstance.HandleError(ErrorNum, ErrStr & DbgStr, ErrTtl, ex_)
#Else
        AppInstance.HandleError(ErrorNum, ErrStr, ErrTtl, ex_)
        AppInstance.Logger.LogError(ErrorNum, "..." & DbgStr)
#End If
    End Sub

    Public Sub ShowBusyIcon(Optional ByVal BusyState As Boolean = False)
        'Dim frm As Form
        'Dim ctl As Control
        'Dim Crs As Cursor

        AppInstance.ShowBusyIcon(BusyState)

        'Crs = CType(IIf(BusyState, Cursors.WaitCursor, Cursors.Default), Cursor)
        'Cursor.Current = Crs
        'gfrmMain.Cursor = Crs
        'gfrmMain.Cursor.Current = Crs
        'For Each frm In gfrmMain.MdiChildren()
        '    frm.Cursor = Crs
        '    frm.Cursor.Current = Crs
        '    For Each ctl In frm.Controls
        '        ctl.Cursor = Crs
        '    Next
        'Next

        'If Not (gfrmCustomer Is Nothing) Then gfrmCustomer.lstATM.Cursor = Crs
    End Sub

    Public Function GetErrStr(ByVal ErrStr As String) As String
        Return AppInstance.GetErrStr(ErrStr)
    End Function

    Public Sub ShowStatusText(Optional ByVal msgText As String = Nothing)
        AppInstance.ShowStatusText(msgText)
    End Sub

    Public Sub ICE_LogOn()

        Dim MyForm As New LoginForm
        Dim MsgID As String
        Dim bTablesLoaded As Boolean
        Dim bUserAutLoaded As Boolean
        Dim skipUpdate As Boolean = False
        ShowStatusText("Logging in...")
        AppInstance.Logger.LogInfo(&H80000001, "-------------" & SIBL0100.Util.UDate.Gregorian(Date.Now, "yyyy-MM-dd hh:mm:ss"))
        AppInstance.Logger.LogInfo(&H80000001, "Logging in...")


        Debug.WriteLine("f")
        Try
            'Auto logon if command line contains /LogonPass=
            Dim idx As Integer = Environment.CommandLine.ToUpper.IndexOf("/LOGONPASS=")
            If (idx >= 0) Then
                Dim pwdLen As Integer = CInt(Environment.CommandLine.Substring(idx + 11, 2))
                If Environment.CommandLine.Substring(idx + 13, 1) <> "," Then Throw New Exception("Invalid command arugment in LogonPass")
                Dim pwdStr As String = Environment.CommandLine.Substring(idx + 14, pwdLen)
                MyForm.Show()
                MyForm.txtPassword.Text = pwdStr
                MyForm.btnLogin_Click(gfrmMain, System.EventArgs.Empty)
            Else
                MyForm.ShowDialog()
            End If
        Catch ex As Exception
            HandleError(&H89999999, "Could not auto logon: " & ex.Message)
            MyForm.ShowDialog()
        End Try

        ShowBusyIcon(True)
        If Not (AppInstance.IsLoggedOn) Then
            ShowStatusText()
            If MyForm.DialogResult <> DialogResult.Cancel Then
                ShowBusyIcon()
                HandleError(&H80000102, "Login Operation Failed!", "ICE Login")
            End If
            AppInstance.Unit = ""
        Else
            gfrmMain.menuFile_Login.Enabled = False

            Dim bRestart As Boolean = False

            'skipUpdate = CheckForUpdate(AppInstance.UnitName, bRestart)
            'If Environment.CommandLine.ToUpper.IndexOf("/NOUPDATE") >= 0 Then
            '    skipUpdate = True
            'End If
            If CheckForUpdate(AppInstance.UnitName, bRestart, skipUpdate) Then
                If Not bRestart Then Process.Start(Application.ExecutablePath, "WAIT_FOR_EXIT")
                Application.Exit()
                Application.DoEvents()
                Return
            Else
                If Not skipUpdate Then
                    Dim Msg As String = "ICE cannot go online." & vbCrLf & _
                                        "The following error occurred: " & vbCrLf & "ICE update server is not reachanble and cannot verify this version of ICE" & vbCrLf & _
                                        "Please contact your Network Support Personnel for assistance."
                    ModalMsgBox(Msg, MsgBoxStyle.Critical, "ICEP0100")
                    ShowBusyIcon(False)
                    Application.Exit()
                    Application.DoEvents()
                    Return
                End If

            End If

            AppInstance.Logger.LogInfo(&H80000000, "Connecting to MQ Server (" & AppInstance.MQ_Server & ", " & _
                                        AppInstance.MQ_WriteQueue & ", " & AppInstance.MQ_ReadQueue & ")")

            'If Not (AppInstance.MQS.Connect(AppInstance.MQ_Server, AppInstance.MQ_WriteQueue, AppInstance.MQ_ReadQueue)) Then
            'Dim Msg As String = "Could not connect to MQ Server. ICE cannot go online." & vbCrLf & _
            '                    "The following error occurred:" & AppInstance.MQS.ErrDsc & vbCrLf & _
            '                    "Please contact your Network Support Personnel for assistance." & vbCrLf & _
            '                    "Would you like to see the MQSeries Log File?"
            'Dim Resp As MsgBoxResult
            ''AppInstance.Logger.LogError(Msg)
            'AppInstance.Logger.LogError(&H80000101, "Could not connect to MQ Server. " & AppInstance.MQS.ErrNum & ":" & AppInstance.MQS.ErrDsc)
            'Resp = ModalMsgBox(Msg, MsgBoxStyle.YesNo Or MsgBoxStyle.Critical, "ICEP0100")
            'If Resp = MsgBoxResult.Yes Then ShowMQErrorLog()
            'AppInstance.IsLoggedOn = False
            'AppInstance.Unit = ""
            'End If

            If Not (AppInstance.MQS.Connect(AppInstance.MQ_Server, AppInstance.MQ_WriteQueue, AppInstance.MQ_ReadQueue)) Then
                'If True Then
                AppInstance.Logger.LogError(&H80000101, "Could not connect to MQ Server. " & AppInstance.MQS.ErrNum & ":" & AppInstance.MQS.ErrDsc)
                AppInstance.IsLoggedOn = False
                AppInstance.Unit = ""
                Dim frmMsg As New frmMsgOpenLog
                frmMsg.Tag = AppInstance.MQS.ErrDsc
                'frmMsg.MdiParent = gfrmMain
                frmMsg.ShowDialog(gfrmMain)

            End If

            If AppInstance.IsLoggedOn Then

                AppInstance.Logger.LogInfo(&H80000002, "Connected to MQ Server")

                LoadOptions(4)

                With AppInstance.MQS
                    .SrcEnvName = AppInstance.Unit
                    .SrcSysName = "ICE"
                    .DstEnvName = AppInstance.Unit
                    .DstSysName = "SAIB"
                    .LoggedUser = AppInstance.LoggedUser
                    '.TmeOut = 2000
                    .BrnCde = AppInstance.GetUserBranch()
                    .DoWait = False
                End With
                bTablesLoaded = SyncTables()
                If bTablesLoaded Then
                    MsgID = SendMessage_GetIceUsrAut(AppInstance.LoggedUser)
                    If IsValidMid(MsgID) Then
                        bUserAutLoaded = GetMessage_GetIceUsrAut(MsgID, IceUserAut)
                    End If
                End If
                If bTablesLoaded And bUserAutLoaded Then
                    'Fill the unit tooltip data
                    gfrmMain.statusPanelUnit.ToolTipText = AppInstance.LoggedUser & "; " & IceUserAut.IceEqnBrn & "; " & IceUserAut.IceEqnTit.Trim & " (" & IceUserAut.IceEqnRol.Trim & "); " & AppInstance.MQ_Server
                    AppInstance.Logger.LogInfo(0, "User info: " & gfrmMain.statusPanelUnit.ToolTipText)
                    'gfrmMain.menuFile_Logout.Enabled = True
                    'gfrmMain.menuFile_Login.Enabled = False
                    gfrmMain.menuServices.Enabled = True
                    gfrmMain.menuManage.Enabled = True
                    gfrmMain.menuTools_Config.Enabled = True
                    gfrmMain.menuTools_Security.Enabled = True
                    gfrmMain.menuTools_Stats.Enabled = True
                    gfrmMain.menuTools_Journal.Enabled = True
                    gfrmMain.menuTools_CashierJournal.Enabled = True
                    'gfrmMain.menuAddsIns.Enabled = True
                    gfrmMain.MenuItemDTS.Enabled = True
                    'gfrmMain.MenuItemDTSCreatePackage.Enabled = True
                    'gfrmMain.MenuItemDTSUpdateStatus.Enabled = True
                    'gfrmMain.MenuItemDTSTracking.Enabled = True
                    'gfrmMain.MenuItemDTSReports.Enabled = True
                    'gfrmMain.MenuItemDTSSettings.Enabled = True
                    ''HACK:Must check for Authority for FOX
                    'gfrmMain.menuAddsIns_FoxMonitor.Enabled = True
                    ShowStatusText("User Successfully Logged in")
                    AppInstance.Logger.LogInfo(&H80000003, "Login successful")

                    'If ValidateBranch() = True Then
                    '    gfrmMain.menuService_Financial.Enabled = True
                    '    'AppInstance.HandleError(1, "Could not validate the branch workstation.", "Financial Services")
                    'End If

                Else
                    AppInstance.IsLoggedOn = False
                    AppInstance.Unit = ""
                    'Clear the unit tooltip data
                    gfrmMain.statusPanelUnit.ToolTipText = ""
                    If Not (bTablesLoaded) Then
                        HandleError(&H80004001, "Unable to Login. Could not access host's database.")
                    Else
                        HandleError(&H80004002, "Unable to Login. Could not obtain user authorities.")
                    End If
                    'ZAK: gfrmMain.statusPanelUnit.Text = ""
                End If
            Else
                'ZAK: gfrmMain.statusPanelUnit.Text = ""
                AppInstance.Unit = ""
            End If
        End If
        If AppInstance.IsLoggedOn Then
            If AppInstance.Unit = "PRD" Then
                DTS.Helper.ConfigHelper.DTSConnectionString = EncryptionHelper.Decrypt("F5o4Y/Jnk7D3aY1piSxM/d8654Msi4xgx/cStAlMWiAviPzX5Zu8sFuvcWmGDSJ3OU320XiJlHfcH5VX/Q7YQ8AQLeEXYR64zLNHxHE3MPWG5uSJcCLb2VgLzb6D0c3/YCZsrBfIooYkFc9ccwAHKMge4wXN1eRhrlmvIMS6lKc=")
            ElseIf AppInstance.Unit = "QRD" Then
                DTS.Helper.ConfigHelper.DTSConnectionString = EncryptionHelper.Decrypt("F5o4Y/Jnk7D3aY1piSxM/cvU9VhwDX1XsyXQ1O8NOe0+5VduLWL5UzqFca2S/+w2RZXykFROXt22Y8ru/d3+HaDuoYRO+FXmSNz96OO5lvRoqVV5eIog1S2GtP0JNHSIUe4B2XlQHvCFpXjLPZrCvvzTU0hrAAyjqN7UYv0frnU=")
            Else
                DTS.Helper.ConfigHelper.DTSConnectionString = EncryptionHelper.Decrypt("F5o4Y/Jnk7D3aY1piSxM/d5RLald5MVB6tDTaP40KxOWURcP8m5Hr/u6IDmcKwpM09cRBRvimV9GsJx0nnwm4LkmhTacy5mQfXxuMqqfXxeTEe+7l+0ilNbnrv/vPnhn")
            End If
            'DTS.Helper.ConfigHelper.OpenDB()

        End If

        MyForm.Dispose()
        MyForm = Nothing
        gfrmMain.menuFile_Logout.Enabled = AppInstance.IsLoggedOn
        gfrmMain.menuFile_Login.Enabled = Not (AppInstance.IsLoggedOn)

        AppInstance.ShowConnectedIcon(AppInstance.IsLoggedOn)
        If AppInstance.IsLoggedOn Then


#If DEBUG Then
            'gfrmMain.menuTools_security_NumberEngine.Visible = True
            'gfrmMain.MenuItem4.Visible = True
            If AppInstance.nShowToolBar >= 0 Then
                'gfrmMain.toolBarMain.Dock = DockStyle.Top
                gfrmMain.toolBarMain.Dock = AppInstance.nShowToolBar 'DockStyle.Left=3
                'gfrmMain.toolBarMain.Show()
                gfrmMain.Refresh()
            End If

#Else
            'HACK: Enumerate SADAD MOI billers always until included in CstVerTab
            If (IceUserAut.IceAutLst(54) Or IceUserAut.IceAutLst(42) Or IceUserAut.IceAutLst(40) Or IceUserAut.IceAutLst(18)) Then
                Dim m_Business As New clsTransaction_MOIFeeUI
                Dim st As String = Financial_MOICodes.MOIList.BlrActCnt 'this will force the billers to be enumerated from the host
                Dim st2 As String = Financial_MOICodes.MOIPaymentTypes.PmtItmCnt 'this will force Pre-Paid payment types to be enumerated also
                Dim st3 As String = Financial_MOICodes.SADADPrePaidParameters.PrmItmCnt 'this will force Pre-Paid parameters to be enumerated also
            End If
#End If



            AppInstance.ShowStatusText("Welcome, ICE services are online for unit " & AppInstance.Unit)
            AppInstance.gbPleaseLogOn = False
            '''gfrmWelcome = New frmWelcome(AppInstance.gfrmMain)
            '''gfrmWelcome.Owner = gfrmMain
            '''gfrmWelcome.StartPosition = FormStartPosition.Manual
            '''gfrmWelcome.lblUnit.Text = AppInstance.Unit
            '''Select Case AppInstance.Unit
            '''    Case "PRD"
            '''        gfrmWelcome.BackColor = Color.Cyan
            '''    Case ""
            '''        gfrmWelcome.BackColor = SystemColors.Control
            '''    Case Else
            '''        gfrmWelcome.BackColor = Color.FromArgb(255, 255, 192)
            '''End Select
            ''''ZAK:Show current unit as a status line at the bottom of a form
            ''''''gfrmWelcome.Show()
            ''''''gfrmWelcome.Animate_Vertical()
            '''gfrmWelcome.MdiParent = gfrmMain
            '''gfrmWelcome.Show()
            '''gfrmWelcome.DockBottom()
        End If

        ShowBusyIcon()

    End Sub

    Public Sub ICE_LogOff()
        Try
            Dim frm As Form
            AppInstance.gLoggingOff = True
            ShowStatusText("Logging out...")
            ShowBusyIcon(True)
            gfrmMain.toolBarMain.Hide()
            AppInstance.Logger.LogInfo(&H80000004, "Logging out...")
            Sleep(1000)
            For Each frm In gfrmMain.MdiChildren()
                frm.Close()
            Next
            For Each frm In gfrmMain.OwnedForms
                frm.Close()
            Next
            gfrmMain.menuFile_Logout.Enabled = False
            gfrmMain.menuFile_Login.Enabled = True
            gfrmMain.menuServices.Enabled = False
            gfrmMain.menuManage.Enabled = False
            gfrmMain.menuTools_Config.Enabled = False
            gfrmMain.menuTools_Journal.Enabled = False
            gfrmMain.menuTools_CashierJournal.Enabled = False
            gfrmMain.menuTools_Security.Enabled = False
            gfrmMain.menuTools_Stats.Enabled = False
            'gfrmMain.menuAddsIns.Enabled = False
            gfrmMain.MenuItemDTS.Enabled = False
            'gfrmMain.MenuItemDTSCreatePackage.Enabled = False
            'gfrmMain.MenuItemDTSUpdateStatus.Enabled = False
            'gfrmMain.MenuItemDTSTracking.Enabled = False
            'gfrmMain.MenuItemDTSReports.Enabled = False
            'gfrmMain.MenuItemDTSSettings.Enabled = False
            With AppInstance
                .IsLoggedOn = False
                .Unit = ""
                .UnitName = ""
            End With
            AppInstance.ShowConnectedIcon(AppInstance.IsLoggedOn)
            ShowStatusText("User Logged Out")
            'ZAK: gfrmMain.statusPanelUnit.Text = ""
            ModalMessageBox("Logout successful.", , , MessageBoxIcon.Information, "ICE Logout")
            AppInstance.Logger.LogInfo(&H80000005, "Logout successful")
            gfrmMain.statusPanel_Text.Style = StatusBarPanelStyle.OwnerDraw

            gfrmMain.statusPanelUnit.ToolTipText = ""

            ICED0100.BranchDB.Close()
            'ICED0100.Printer_Close()

        Catch ex As Exception
            Debug.WriteLine("ICE_LogOff: " & SIBL0100.Util.Debug.getStackFrame(New StackTrace(True).GetFrame(0)) & ex.ToString)
        End Try
        AppInstance.gLoggingOff = False
        AppInstance.gbPleaseLogOn = True
        ShowBusyIcon()
    End Sub

    Public Sub AutoSizeGrid(ByVal dGrid As ExDataGrid)
        ''DataGrid should be bound to a DataTable for this part to work. 
        Dim numRows As Integer
        Dim numCols As Integer
        Dim g As Graphics
        Dim sf As StringFormat
        Dim size, sizeTemp As SizeF
        '' Since DataGridRows[] is not exposed directly by the DataGrid 
        '' we use reflection to hack internally to it.. There is actually 
        '' a method get_DataGridRows that returns the collection of rows 
        '' that is what we are doing here, and casting it to a System.Array 
        Dim mi As MethodInfo
        Dim dgra As System.Array
        '' Convert this to an ArrayList, little bit easier to deal with 
        '' that way, plus we can strip out the newrow row. 
        Dim DataGridRows As ArrayList

        Try

            numRows = (DirectCast(dGrid.DataSource, DataTable)).Rows.Count
            'numCols = (DirectCast(dGrid.DataSource, DataTable)).Columns.Count
            numCols = 0
            For Each aa As System.Windows.Forms.DataGridColumnStyle In dGrid.TableStyles(0).GridColumnStyles
                numCols += 1
            Next

            g = Graphics.FromHwnd(dGrid.Handle)
            sf = New StringFormat(StringFormat.GenericTypographic)

            mi = (New DataGrid).GetType().GetMethod("get_DataGridRows", _
                                                        BindingFlags.FlattenHierarchy Or BindingFlags.IgnoreCase Or BindingFlags.Instance Or _
                                                        BindingFlags.NonPublic Or BindingFlags.Public Or BindingFlags.Static)
            dgra = DirectCast(mi.Invoke(dGrid, Nothing), System.Array)
            DataGridRows = New ArrayList

            For Each dgrr As Object In dgra
                If (dgrr.ToString().EndsWith("DataGridRelationshipRow")) Then
                    DataGridRows.Add(dgrr)
                End If
            Next

            '' Now loop through all the rows in the grid 

            'for ( i as integer = 0; i < numRows; ++i) 
            For i As Integer = 0 To numRows - 1
                '' Here we are telling it that the column width is set to 
                '' 1280.. so size will contain the Height it needs to be. 
                size.Height = 8 '20 '8
                For c As Integer = 0 To numCols - 1
                    sizeTemp = g.MeasureString(dGrid(i, c).ToString(), dGrid.Font, 1280, sf)
                    If sizeTemp.Height > size.Height Then
                        size = sizeTemp
                    End If
                Next c

                Dim h As Integer = Convert.ToInt32(size.Height) + 8 '' Little extra cellpadding space  (+8)

                '' Now we pick that row out of the DataGridRows[] Array 
                '' that we have and set it's Height property to what we 
                '' think it should be. 
                Dim pi As PropertyInfo = DataGridRows(i).GetType().GetProperty("Height")
                pi.SetValue(DataGridRows(i), h, Nothing)

                '' I have read here that after you set the Height in this manner that you should 
                '' Call the DataGrid Invalidate() method, but I haven't seen any prob with not calling it.. 
                dGrid.Invalidate()
            Next i
            'dGrid.Invalidate()
            'dGrid.Refresh()
            'Application.DoEvents()

            g.Dispose()
        Catch ex As Exception
            HandleError(&H81000640, "AutoSizeGrid --> Error in resizing grid." & vbCrLf & ex.Message, "Grid Layout")
        End Try
    End Sub


    '* Checks the existance of a file, returns false if it doesn't exist
    Public Function FileExist(ByVal FileName As String) As Boolean
        Dim FileStmp As DateTime
        Try
            FileStmp = FileDateTime(FileName)
        Catch ex As System.IO.FileNotFoundException
            Return False
        Catch ex As Exception
            ModalMsgBox("Error " & ex.ToString & "!", MsgBoxStyle.Exclamation, "INI - File Exist")
        End Try
        Return True
    End Function

    Sub SetMQServerEnvironment(ByVal ChnNam As String)
        Dim RegKey As RegistryKey
        Dim lpResult, Res As Integer
        RegKey = Registry.CurrentUser.OpenSubKey("Environment", True)
        RegKey.SetValue("MQSERVER", ChnNam)
        Dim tt As String = "Environment"
        Res = AppInstance.SendMessageTimeout(AppInstance.HWND_BROADCAST, AppInstance.WM_SETTINGCHANGED, _
                                             0, tt, AppInstance.SMTO_NORMAL, 3000, lpResult)
        If Res = 0 Then
            ModalMsgBox("Could not set MQSERVER Environment Variable!")
        End If
    End Sub

    Public Function ModalMsgBox(ByVal Prompt As String, Optional ByVal Style As MsgBoxStyle = MsgBoxStyle.OKOnly, Optional ByVal Title As String = Nothing) As MsgBoxResult
        Return AppInstance.ModalMsgBox(Prompt, Style, Title)
    End Function

    Public Function ModalMsgBox(ByVal Owner As IWin32Window, ByVal Prompt As String, Optional ByVal Style As MsgBoxStyle = MsgBoxStyle.OKOnly, Optional ByVal Title As String = Nothing) As MsgBoxResult
        Return CType(MessageBox.Show(Owner, Prompt, Title, CType(Style And &HF, MessageBoxButtons), CType(Style And &HF0, MessageBoxIcon), CType(Style And &HF00, MessageBoxDefaultButton)), MsgBoxResult)
    End Function

    Public Function ModalMessageBox(ByVal Prompt As Object, Optional ByVal Btns As MessageBoxButtons = MessageBoxButtons.OK, _
                                    Optional ByVal DefBtn As MessageBoxDefaultButton = MessageBoxDefaultButton.Button1, _
                                    Optional ByVal Icn As MessageBoxIcon = MessageBoxIcon.None, Optional ByVal Title As String = "", Optional ByVal Owner As IWin32Window = Nothing) As DialogResult
        'If Owner Is Nothing Then Owner = gfrmMain
        'Return MessageBox.Show(Owner, Prompt, Title, Btns, Icn, DefBtn)
        If Owner Is Nothing Then
            Return AppInstance.ModalMessageBox(Prompt, Btns, DefBtn, Icn, Title)
        Else
            Return MessageBox.Show(Owner, Prompt, Title, Btns, Icn, DefBtn)
        End If
    End Function

    Public Function ModalMessageBox(ByVal Owner As IWin32Window, ByVal Prompt As Object, Optional ByVal Btns As MessageBoxButtons = MessageBoxButtons.OK, _
                                    Optional ByVal DefBtn As MessageBoxDefaultButton = MessageBoxDefaultButton.Button1, _
                                    Optional ByVal Icn As MessageBoxIcon = MessageBoxIcon.None, Optional ByVal Title As String = "") As DialogResult
        Return MessageBox.Show(Owner, Prompt, Title, Btns, Icn, DefBtn)
    End Function

    Public Function ReadStringFromBlock(ByRef Block() As Byte, ByVal BegIdx As Integer, ByVal StrLen As Integer) As String
        Dim st As String = ""
        If Block Is Nothing Then Return ""
        If Block.Length <= 0 Then Return ""
        If ((BegIdx + StrLen) > Block.Length) Then Return ""

        For i As Integer = 0 To StrLen - 1
            st &= Chr(Block(BegIdx + i))
        Next
        Return st
    End Function

    Public Function WriteStringToBlock(ByVal StrVal As String, ByRef Block() As Byte, ByVal BegIdx As Integer, ByVal StrLen As Integer) As Integer
        Dim st As String = ""
        Dim ch() As Char

        If Block Is Nothing Then Return -10
        If Block.Length <= 0 Then Return -20
        'If Trim(StrVal) = "" Then Return -30
        StrVal = AppInstance.PackString(StrVal, StrLen)
        If ((BegIdx + StrLen) > Block.Length) Then Return -40

        ch = StrVal.ToCharArray
        For i As Integer = 0 To StrLen - 1
            'st &= Chr(Block(BegIdx + i))
            Block(BegIdx + i) = CByte(Asc(ch(i)))
        Next
        Return 0
    End Function


    Private Sub TelTest()
        Dim st() As String = { _
         "8001248000", _
         "800-1248000", _
         "056277129", _
         "05-6277129", _
         "05.6277129", _
         "0506277129", _
         "050-6277129", _
         "050.6277129", _
         "05-06277129", _
         "05.06277129", _
         "+966509979942", _
         "+966056277219", _
         "+9660506277129", _
         "+9660556277129", _
         "+96602.1234567", _
         "+9668001248000", _
         "+966014778433+1348", _
         "+966 01 4778433+1348", _
         "+966 050-6277129", _
         "966-056277129", _
         "966056277129", _
         "00966056277129", _
         "096244569", _
         "123456789", _
         "1234567890", _
         "8956632"}

        Dim oLnd, oAre, oNum, oExt, oStr As String
        Dim res As Integer

        Dim cs As New TELI0100.Phone
        For i As Integer = 0 To st.Length - 1
            oLnd = Space(100)
            oAre = Space(100)
            oNum = Space(100)
            oExt = Space(100)
            oStr = Space(100)
            res = cs.CheckPhone(st(i), "Y".Chars(0), "Y".Chars(0), oLnd, oAre, oNum, oExt, oStr)
            Console.WriteLine("In:" & st(i) & Space(5) & "Out:" & oStr & Space(5) & "Result:" & res)
        Next
        'int res, i;
        'i=0;
        'while(strcmp(TestSet[i],"end"))
        '{
        '	res=CheckPhone(TestSet[i],'Y','Y',oLnd,oAre,oNum,oExt,oStr);
        '	printf("In:%32s, Out:%25s, Result:%d\n",TestSet[i],oStr,res);
        '	i++;
        '}
    End Sub


    Public Sub OnShuttingdown(ByVal sender As Object, ByVal e As SessionEndingEventArgs)
        Console.WriteLine("Shutting down - Reason is " & e.Reason)
        Thread.VolatileWrite(AppInstance.isShuttingDown, True)
    End Sub

    Public Sub OnShutdown(ByVal sender As Object, ByVal e As SessionEndedEventArgs)
        Console.WriteLine("Shutdown - Reason is " & e.Reason)
        'TODO:Add shutdown code
        'Ensure that you do not return from this method until all ICE shutdown routine has been performed
        'Maybe implement a time out in here
    End Sub

    Public Function nz(ByVal objStr As String, Optional ByVal defaultValue As String = "") As String
        Return SIBL0100.Util.UString.nz(objStr, defaultValue)
    End Function

    Public Function nz(ByVal objStr As Object, Optional ByVal defaultValue As String = "") As String
        Return SIBL0100.Util.UString.nz(objStr, defaultValue)
    End Function

    Public Function nzt(ByVal objStr As Object, Optional ByVal defaultValue As String = "") As String
        Return CStr(SIBL0100.Util.UString.nz(objStr, defaultValue)).Trim
    End Function

    Public Sub main()
        TELI0100.Phone.SetCompatibilityMode("B")
        Try
            With (New EnableThemingInScope(True))
                'System.Windows.Forms.Application.EnableVisualStyles()

                'Load Splash Screen

                gfrmSplash = New frmSplash
                gfrmSplash.Starting = True
                gfrmSplash.Cursor = Cursors.WaitCursor

                gfrmSplash.ShowInTaskbar = True
                gfrmSplash.Show()
                Application.DoEvents()

                '''AppInstance.ModalMsgBox(Environ$("MQSERVER"))

                'First, Check for a previous instance
                Dim bb As Boolean = (Environment.CommandLine.IndexOf("WAIT_FOR_EXIT") > 0)
                Try
                    If PrevInstance(False, bb) Then
                        MessageBox.Show("Warning 90000002!" & vbCrLf & _
                                "ICE is already running on your computer." & vbCrLf & "Multiple Instances cannot be started." & _
                                vbCrLf & "Please use the currently active ICE session.", "ICE", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                        PrevInstance(True, False)
                        Return
                    End If
                Catch ex As Exception
                    MessageBox.Show("80000000:Could not check for running instances of ICE, ICE will proceed." & ex.Message, "ICE Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                End Try

                Try
                    SetAttr(ICEI0100.IcePaths.DefInstance.IceUsrIni, FileAttribute.Archive)
                    SetAttr(ICEI0100.IcePaths.DefInstance.IceCfgIni, FileAttribute.Archive)
                Catch ex As Exception
                    MessageBox.Show("80000000: Could not set attributes for ICE's INI files! " & ex.Message, "ICE Critical Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End Try

                Dim OldCultInfo, NewCultInfo As CultureInfo
                OldCultInfo = System.Threading.Thread.CurrentThread.CurrentCulture
                NewCultInfo = New CultureInfo("ar-SA", False)
                'NewCultInfo = New CultureInfo("en-AU", False)
                System.Threading.Thread.CurrentThread.CurrentCulture = NewCultInfo
                'System.Threading.Thread.CurrentThread.CurrentUICulture = NewCultInfo


                Dim DlgAns As DialogResult
                Dim bOffline As Boolean = False

                '"BQ.GOLF.SAIBW300/tcp/Golf"
                AppInstance = New ICEI0100.AppInstanceClass
                AppInstance.SystemLocale = OldCultInfo
                SIBL0100.WinWord.WinWordError.ErrorHandler = AddressOf AppInstance.HandleError
                AddHandler SystemEvents.SessionEnding, AddressOf OnShuttingdown
                AddHandler SystemEvents.SessionEnded, AddressOf OnShutdown

                LoadOptions(1)
                AppInstance.Initialize(ICEI0100.IcePaths.DefInstance.ICE)
                If AppInstance.ErrNum <> 0 Then
                    MessageBox.Show(AppInstance.ErrDsc & vbCrLf & "ICE cannot continue!", "ICE Critical Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Return
                End If
                LoadOptions(2)
                AppInstance.Logger.OpenLog()
                AppInstance.Logger.LogInfo(&H80000000, "___________________________________________________________________________________________________________________")
                AppInstance.Logger.LogInfo(&H80000000, "___________________________________________________________________________________________________________________")
                AppInstance.Logger.LogInfo(&H80000006, "Starting ICE (version " & gIceVersion & ", Build " & gIceBuild & ")...")
                AppInstance.Logger.LogInfo(&H80000011, "Application Instantiated.", 2)
                AppInstance.Logger.LogInfo(&H80000007, "Logging Services Started.")
                AppInstance.Logger.LogInfo(&H80000012, "Global Options Loaded.", 2)


                If Not (IsMQSeriesInstalled()) Then
                    DlgAns = MessageBox.Show("MQSeries software appears not to be isntalled." & vbCrLf & "This is required for ICE to function Online." & _
                            vbCrLf & "Would you like to continue offline?", "ICE Critical Error", MessageBoxButtons.YesNo, MessageBoxIcon.Error, MessageBoxDefaultButton.Button2)

                    'DlgAns = ModalMsgBox("MQSeries software appears not to be isntalled." & vbCrLf & "This is required for ICE to function Online." & _
                    '        vbCrLf & "Would you like to continue offline?", MsgBoxStyle.Exclamation Or MsgBoxStyle.YesNo Or MsgBoxStyle.DefaultButton2)
                    If DlgAns = DialogResult.No Then Return
                    bOffline = True
                End If

                If (Environment.GetEnvironmentVariable("MQSERVER") = "") And (Not bOffline) Then
                    DlgAns = MessageBox.Show("ICE has detected that some parameters required for communicating with Host have not been set." & vbCrLf & _
                                              "Setting up these parameters will require you to restart ICE." & vbCrLf & _
                                              "If you continue without setting these parameters you will be running offline." & vbCrLf & _
                                              "Would you like ICE to set these parameters now?", "ICE Critical Error", MessageBoxButtons.YesNo, MessageBoxIcon.Error)
                    'DlgAns = ModalMsgBox("ICE has detected that some parameters required for communicating with Host have not been set." & vbCrLf & _
                    '                "Setting up these parameters will require you to restart ICE." & vbCrLf & _
                    '                "If you continue without setting these parameters you will be running offline." & vbCrLf & _
                    '                "Would you like ICE to set these parameters now?", MsgBoxStyle.Exclamation Or MsgBoxStyle.YesNo)
                    If DlgAns = DialogResult.Yes Then
                        Dim st As String
                        Dim ini As INIClass.IniFileClass
                        ini = New INIClass.IniFileClass(ICEI0100.IcePaths.DefInstance.IceUsrIni)
                        If ini.FileExist Then
                            st = ini.GetValue("LastUsed", "Units", "PRD1").Trim
                            ini = New INIClass.IniFileClass(ICEI0100.IcePaths.DefInstance.IceCfgIni)
                            st = ini.GetValue("Channel", st, "BQ.ECHO.SAIBW832/TCP/ECHO").Trim
                            SetMQServerEnvironment(st)
                        Else
                            MessageBox.Show("Could not locate INI file IceUsr.ini!", "ICE Critical Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                            'ModalMsgBox("Could not locate INI file IceUsr.ini!", MsgBoxStyle.Critical, "ICE Critical Error")
                        End If
                        Return
                    End If
                    bOffline = True
                End If

                'AddressOf Sort.CompareValues

                ICEComm.Initialize()
                AppInstance.Logger.LogInfo(&H80000013, "ICEComm Module Initialized", 1)

                LoadOptions(3)
#If ArabicWosa Then
                LoadOptions(5)
                AppInstance.Logger.LogInfo(&H80000014, "Loading Default Printing Language", 1)
#End If
                AppInstance.Logger.LogInfo(&H80000014, "MQSeries Options Loaded", 1)

                'Sleep(1000)

                gfrmMain = New frmMain
                AppInstance.gfrmMain = gfrmMain
                gfrmMain.menuFile_Login.Enabled = Not bOffline
                Plugins.LoadPlugin(Application.StartupPath, "ICEP0100.Plugin", CType(gfrmMain, Form), "ANT*.dll")
                gfrmSplash.Close()
                gfrmSplash = Nothing

                gfrmPrint = New frmPrint
                AppInstance.gfrmPrint = gfrmPrint
                ICEI0100.AppInstanceClass.clsPrintForm.Status = "as"
                ICEI0100.AppInstanceClass.clsPrintForm.ShowPrintForm()

                gfrmPrint.Close()
                gfrmPrint = Nothing

                gfrmMain.Show()
                Application.Run()
                AppInstance.gfrmMain = Nothing
                AppInstance.Logger.LogInfo(&H80000007, "Main Application Form Unloaded")

                AppInstance.Logger.LogInfo(&H80000008, "Performing Cleanup...")
                AppInstance.Logger.LogInfo(&H80000009, "ICE Ended.")

                System.Threading.Thread.CurrentThread.CurrentCulture = OldCultInfo
                NewCultInfo = Nothing


                GC.Collect()
            End With

        Catch ex As Exception
            Debug.WriteLine("main: " & New StackFrame(True).GetMethod.Name & ": " & ex.ToString)

            'ModalMsgBox("80000001: ICE Application generated an error and cannot continue!" _
            '& vbCrLf & ex.ToString, MsgBoxStyle.Critical, "ICE Critical Error")
            MessageBox.Show("80000001: ICE Application generated an error and cannot continue!" _
                & vbCrLf & ex.ToString, "ICE Critical Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Public Class GetChannelWordTable_Class


        Public Structure Req_Struct '8,224x, App #761, Msg #
            Dim TabRowVer As String '12x                                  Desc: Table row version
            Dim Rsv140 As String '20x                            Desc: Reserved, space filled

            Public Function PackTable() As String
                Dim st As String = String.Empty
                st &= AppInstance.PackString(TabRowVer, 12)
                st &= AppInstance.PackString(Rsv140, 20)
                Return st
            End Function

            Public Sub ReadTable(ByRef DataStream As String)
                TabRowVer = AppInstance.strip(DataStream, 12)
                Rsv140 = AppInstance.strip(DataStream, 20)
            End Sub

            Public Sub Clear()
                TabRowVer = String.Empty
                Rsv140 = String.Empty
            End Sub

        End Structure

        Public Structure Chnwrd_Struct '08x, App #761, Msg #805130
            Dim ChnWrdEng As String '4x                                   Desc: Channel enrolment word in English for English (lower case)
            Dim ChnWrdAra As String '4x                                   Desc: Channel enrolment word in English for Arabic (lower case)

            Public Function PackTable() As String
                Dim st As String = String.Empty
                st &= AppInstance.PackString(ChnWrdEng, 4)
                st &= AppInstance.PackString(ChnWrdAra, 4)
                Return st
            End Function

            Public Sub ReadTable(ByRef DataStream As String)
                ChnWrdEng = AppInstance.strip(DataStream, 4)
                ChnWrdAra = AppInstance.strip(DataStream, 4)
            End Sub

            Public Sub Clear()
                ChnWrdEng = String.Empty
                ChnWrdAra = String.Empty
            End Sub
        End Structure

        Public Structure Rsp_Struct '8,224x, App #761, Msg #805130

            Public TabRowVer As String '12n                                                    Desc: Table row version
            Public Rsv100 As String '16x                                                              Desc: Reserved, space filled
            Public ChnWrdCnt As String '4n                                                      Desc: Number of word sets that follow, constant 1024
            Public ChnWrd() As Chnwrd_Struct '8 * 1024                                                            Desc:Channel word sets

            Public Function PackTable() As String
                Dim st As String = String.Empty
                st &= AppInstance.PackNString(TabRowVer, 12)
                st &= AppInstance.PackString(Rsv100, 16)
                st &= AppInstance.PackNString(ChnWrdCnt, 4)
                For ind As Integer = 0 To ChnWrd.Length - 1 : st &= ChnWrd(ind).PackTable() : Next
                Return st
            End Function

            Public Sub ReadTable(ByRef DataStream As String)
                TabRowVer = AppInstance.strip(DataStream, 12)
                Rsv100 = AppInstance.strip(DataStream, 16)
                ChnWrdCnt = AppInstance.strip(DataStream, 4)
                ReDim ChnWrd(CInt(ChnWrdCnt) - 1)
                For ind As Integer = 0 To ChnWrd.Length - 1 : ChnWrd(ind).ReadTable(DataStream) : Next
            End Sub

            Public Sub Clear()
                TabRowVer = String.Empty
                Rsv100 = String.Empty
                ChnWrdCnt = String.Empty
                ChnWrd = Nothing
            End Sub

        End Structure

        Dim Req As Req_Struct
        Dim Rsp As Rsp_Struct
        Dim MsgID As String

        Public Function SendMessage() As String
            Dim st As String = String.Empty
            st &= Req.PackTable() '8224x
            MsgID = AppInstance.SendMessageVisuallyEx("805130", st, "GEN", "", "", "")
            Return MsgID
        End Function

        Public Function GetMessage(ByVal MsgId As String, ByRef mADO As SIBL0100.AdoAccess) As Boolean
            If MsgId Is Nothing OrElse MsgId.Trim = "" Then
                MsgId = Me.MsgID
            End If

            Dim resp As String = String.Empty
            Dim verStr As String
            Dim i, count As Integer
            Dim verRec As CoreSys_VerItm
            Dim row As DataRow

            'App #761
            If Not (GetMessage_Common("805130", New StackFrame().GetMethod.Name, MsgId, &H80099999, resp)) Then Return False
            Try
                AppInstance.strip(resp, 256)
                Rsp.ReadTable(resp) '8224x
                If Rsp.ChnWrd Is Nothing OrElse Rsp.ChnWrd.Length = 0 Then
                    'HandleError(&H80002060, SIBL0100.Util.Debug.getStackFrame(New StackTrace(True).GetFrame(0)) & ": No data received", "ICE Communications")
                    'Return False
                    Erase ChnWrdTable
                    mADO.DeleteTable("ChnWrdTab")
                    Return True
                End If
                count = Rsp.ChnWrd.Length
                ReDim ChnWrdTable(count - 1)
                mADO.DeleteTable("ChnWrdTab")
                For i = 0 To ChnWrdTable.Length - 1
                    With ChnWrdTable(i)
                        .ChnWrdEng = Rsp.ChnWrd(i).ChnWrdEng
                        .ChnWrdAra = Rsp.ChnWrd(i).ChnWrdAra
                        mADO.AddRowToTable("ChnWrdTab", New String() {.ChnWrdEng, .ChnWrdAra})
                    End With
                Next
            Catch ex As Exception
                HandleError(&H80099999, New StackFrame().GetMethod.Name & ":" & ex.Message, , ex)
                Return False
            End Try
            Return True
        End Function
    End Class
End Module


<SuppressUnmanagedCodeSecurityAttribute()> _
Public Class EnableThemingInScope
    Implements IDisposable
    ' Private data
    Private cookie As Integer
    Private Shared enableThemingActivationContext As ACTCTX
    Private Shared hActCtx As IntPtr
    Private Shared contextCreationSucceeded As Boolean = False

    Public Sub New(ByVal enable As Boolean)
        cookie = 0
        If enable AndAlso OSFeature.Feature.IsPresent(OSFeature.Themes) Then
            If EnsureActivateContextCreated() Then
                If Not ActivateActCtx(hActCtx, cookie) Then
                    ' Be sure cookie always zero if activation failed
                    cookie = 0
                End If
            End If
        End If
    End Sub

    Protected Overrides Sub Finalize()
        Try
            Dispose(False)
        Finally
            MyBase.Finalize()
        End Try
    End Sub

    Private Sub Dispose() Implements IDisposable.Dispose
        Dispose(True)
    End Sub

    Private Sub Dispose(ByVal disposing As Boolean)
        If cookie <> 0 Then
            Try
                If DeactivateActCtx(0, cookie) Then
                    ' deactivation succeeded...
                    cookie = 0
                End If
            Catch
                cookie = 0
            End Try
        End If
    End Sub

    Private Function EnsureActivateContextCreated() As Boolean
        SyncLock GetType(EnableThemingInScope)
            If Not contextCreationSucceeded Then
                ' Pull manifest from the .NET Framework install
                ' directory

                Dim assemblyLoc As String = Nothing

                Dim fiop As New Permissions.FileIOPermission(Permissions.PermissionState.None)
                fiop.AllFiles = Permissions.FileIOPermissionAccess.PathDiscovery
                fiop.Assert()
                Try
                    assemblyLoc = GetType([Object]).Assembly.Location
                Finally
                    CodeAccessPermission.RevertAssert()
                End Try

                Dim manifestLoc As String = Nothing
                Dim installDir As String = Nothing
                If Not (assemblyLoc Is Nothing) Then
                    installDir = IO.Path.GetDirectoryName(assemblyLoc)
                    Const manifestName As String = "XPThemes.manifest"
                    manifestLoc = IO.Path.Combine(installDir, manifestName)
                End If

                If Not (manifestLoc Is Nothing) AndAlso Not (installDir Is Nothing) Then
                    enableThemingActivationContext = New ACTCTX
                    enableThemingActivationContext.cbSize = Marshal.SizeOf(GetType(ACTCTX))
                    enableThemingActivationContext.lpSource = manifestLoc

                    ' Set the lpAssemblyDirectory to the install
                    ' directory to prevent Win32 Side by Side from
                    ' looking for comctl32 in the application
                    ' directory, which could cause a bogus dll to be
                    ' placed there and open a security hole.
                    enableThemingActivationContext.lpAssemblyDirectory = installDir
                    enableThemingActivationContext.dwFlags = ACTCTX_FLAG_ASSEMBLY_DIRECTORY_VALID

                    ' Note this will fail gracefully if file specified
                    ' by manifestLoc doesn't exist.
                    hActCtx = CreateActCtx(enableThemingActivationContext)
                    contextCreationSucceeded = Not hActCtx.Equals(New IntPtr(-1))
                End If
            End If

            ' If we return false, we'll try again on the next call into
            ' EnsureActivateContextCreated(), which is fine.
            Return contextCreationSucceeded
        End SyncLock
    End Function

    ' All the pinvoke goo...
    <DllImport("Kernel32.dll")> _
    Private Shared Function CreateActCtx(ByRef actctx As ACTCTX) As IntPtr
    End Function
    <DllImport("Kernel32.dll")> _
    Private Shared Function ActivateActCtx(ByVal hActCtx As IntPtr, ByRef lpCookie As Integer) As Boolean
    End Function
    <DllImport("Kernel32.dll")> _
    Private Shared Function DeactivateActCtx(ByVal dwFlags As Integer, ByVal lpCookie As Integer) As Boolean
    End Function

    Private Const ACTCTX_FLAG_ASSEMBLY_DIRECTORY_VALID As Integer = &H4

    Private Structure ACTCTX
        Public cbSize As Integer
        Public dwFlags As Integer
        Public lpSource As String
        Public wProcessorArchitecture As Short
        Public wLangId As Short
        Public lpAssemblyDirectory As String
        Public lpResourceName As String
        Public lpApplicationName As String
    End Structure
End Class

Module ObsoleteFunctions
    Public Sub SetApartmentState_MTA(ByVal PrintThread As Threading.Thread)
        PrintThread.ApartmentState = ApartmentState.MTA
    End Sub

End Module








