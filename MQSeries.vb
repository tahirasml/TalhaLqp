Option Strict Off

Imports System.Globalization

Public Class MQSeries
    Inherits ErrorClass

    Shared Function strip(ByRef strMsg As String, ByVal nChars As Integer) As String
        Dim RetStr As String
        If strMsg Is Nothing Then Return Nothing
        If strMsg.Length < 1 Then Return Nothing
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

    Shared Function IsValidMid(ByVal MsgID As String) As Boolean
        If MsgID Is Nothing Then Return False
        If MsgID.Trim = "" Then Return False
        If MsgID.Length <> 12 Then Return False
        Return True
    End Function


#Region "Declaration Section"

    Public Enum enumWaitTime As Integer
        TENTH = 100
        HALF = 500
        ONE = 1000
        TWO = 2000
        THREE = 3000
        FOUR = 4000
        FIVE = 5000
    End Enum

    'Private Const MQ_ABEND_CODE_LENGTH = 4
    'Private Const MQ_ACCOUNTING_TOKEN_LENGTH = 32
    'Private Const MQ_APPL_IDENTITY_DATA_LENGTH = 32
    'Private Const MQ_APPL_NAME_LENGTH = 28
    'Private Const MQ_APPL_ORIGIN_DATA_LENGTH = 4
    'Private Const MQ_ATTENTION_ID_LENGTH = 4
    'Private Const MQ_AUTHENTICATOR_LENGTH = 8
    'Private Const MQ_BRIDGE_NAME_LENGTH = 24
    'Private Const MQ_CANCEL_CODE_LENGTH = 4
    'Private Const MQ_CHANNEL_DATE_LENGTH = 12
    'Private Const MQ_CF_STRUC_NAME_LENGTH = 12
    'Private Const MQ_CHANNEL_DESC_LENGTH = 64
    'Private Const MQ_CHANNEL_NAME_LENGTH = 20
    'Private Const MQ_CHANNEL_TIME_LENGTH = 8
    'Private Const MQ_CLUSTER_NAME_LENGTH = 48
    'Private Const MQ_CONN_NAME_LENGTH = 264
    'Private Const MQ_CONN_TAG_LENGTH = 128
    'Private Const MQ_CORREL_ID_LENGTH = 24
    'Private Const MQ_CREATION_DATE_LENGTH = 12
    'Private Const MQ_CREATION_TIME_LENGTH = 8
    'Private Const MQ_DATE_LENGTH = 12
    'Private Const MQ_EXIT_DATA_LENGTH = 32
    'Private Const MQ_EXIT_NAME_LENGTH = 128
    'Private Const MQ_EXIT_USER_AREA_LENGTH = 16
    'Private Const MQ_FACILITY_LENGTH = 8
    'Private Const MQ_FACILITY_LIKE_LENGTH = 4
    'Private Const MQ_FORMAT_LENGTH = 8
    'Private Const MQ_FUNCTION_LENGTH = 4
    'Private Const MQ_GROUP_ID_LENGTH = 24
    'Private Const MQ_LTERM_OVERRIDE_LENGTH = 8
    'Private Const MQ_LUWID_LENGTH = 16
    'Private Const MQ_MAX_EXIT_NAME_LENGTH = 128
    'Private Const MQ_MAX_MCA_USER_ID_LENGTH = 64
    'Private Const MQ_MCA_JOB_NAME_LENGTH = 28
    'Private Const MQ_MCA_NAME_LENGTH = 20
    'Private Const MQ_MCA_USER_ID_LENGTH = 64
    'Private Const MQ_MFS_MAP_NAME_LENGTH = 8
    'Private Const MQ_MODE_NAME_LENGTH = 8
    'Private Const MQ_MSG_HEADER_LENGTH = 4000
    'Private Const MQ_MSG_ID_LENGTH = 24
    'Private Const MQ_MSG_TOKEN_LENGTH = 16
    'Private Const MQ_NAMELIST_DESC_LENGTH = 64
    'Private Const MQ_NAMELIST_NAME_LENGTH = 48
    'Private Const MQ_OBJECT_INSTANCE_ID_LENGTH = 24
    'Private Const MQ_OBJECT_NAME_LENGTH = 48
    'Private Const MQ_PASSWORD_LENGTH = 12
    'Private Const MQ_PROCESS_APPL_ID_LENGTH = 256
    'Private Const MQ_PROCESS_DESC_LENGTH = 64
    'Private Const MQ_PROCESS_ENV_DATA_LENGTH = 128
    'Private Const MQ_PROCESS_NAME_LENGTH = 48
    'Private Const MQ_PROCESS_USER_DATA_LENGTH = 128
    'Private Const MQ_PUT_APPL_NAME_LENGTH = 28
    'Private Const MQ_PUT_DATE_LENGTH = 8
    'Private Const MQ_PUT_TIME_LENGTH = 8
    'Private Const MQ_Q_DESC_LENGTH = 64
    'Private Const MQ_Q_MGR_DESC_LENGTH = 64
    'Private Const MQ_Q_MGR_IDENTIFIER_LENGTH = 48
    'Private Const MQ_Q_MGR_NAME_LENGTH = 48
    'Private Const MQ_Q_NAME_LENGTH = 48
    'Private Const MQ_QSG_NAME_LENGTH = 4
    'Private Const MQ_REMOTE_SYS_ID_LENGTH = 4
    'Private Const MQ_SECURITY_ID_LENGTH = 40
    'Private Const MQ_SERVICE_NAME_LENGTH = 32
    'Private Const MQ_SERVICE_STEP_LENGTH = 8
    'Private Const MQ_SHORT_CONN_NAME_LENGTH = 20
    'Private Const MQ_START_CODE_LENGTH = 4
    'Private Const MQ_STORAGE_CLASS_LENGTH = 8
    'Private Const MQ_TIME_LENGTH = 8
    'Private Const MQ_TOTAL_EXIT_DATA_LENGTH = 999
    'Private Const MQ_TOTAL_EXIT_NAME_LENGTH = 999
    'Private Const MQ_TP_NAME_LENGTH = 64
    'Private Const MQ_TRANSACTION_ID_LENGTH = 4
    'Private Const MQ_TRAN_INSTANCE_ID_LENGTH = 16
    'Private Const MQ_TRIGGER_DATA_LENGTH = 64
    'Private Const MQ_USER_ID_LENGTH = 12
    'Private Const MQACT_FORCE_REMOVE = 1
    'Private Const MQAR_ALL = 1
    'Private Const MQAR_NONE = 0
    'Private Const MQAT_UNKNOWN = &HFFFFFFFF
    'Private Const MQAT_NO_CONTEXT = 0
    'Private Const MQAT_CICS = 1
    'Private Const MQAT_MVS = 2
    'Private Const MQAT_OS390 = 2
    'Private Const MQAT_IMS = 3
    'Private Const MQAT_OS2 = 4
    'Private Const MQAT_DOS = 5
    'Private Const MQAT_AIX = 6
    'Private Const MQAT_UNIX = 6
    'Private Const MQAT_QMGR = 7
    'Private Const MQAT_OS400 = 8
    'Private Const MQAT_WINDOWS = 9
    'Private Const MQAT_CICS_VSE = 10
    'Private Const MQAT_DEFAULT = 11
    'Private Const MQAT_WINDOWS_NT = 11
    'Private Const MQAT_VMS = 12
    'Private Const MQAT_GUARDIAN = 13
    'Private Const MQAT_NSK = 13
    'Private Const MQAT_VOS = 14
    'Private Const MQAT_IMS_BRIDGE = 19
    'Private Const MQAT_XCF = 20
    'Private Const MQAT_CICS_BRIDGE = 21
    'Private Const MQAT_NOTES_AGENT = 22
    'Private Const MQAT_BROKER = 26
    'Private Const MQAT_JAVA = 28
    'Private Const MQAT_DQM = 29
    'Private Const MQAT_USER_FIRST = &H10000
    'Private Const MQAT_USER_LAST = &H3B9AC9FF
    'Private Const MQBL_NULL_TERMINATED = &HFFFFFFFF
    'Private Const MQBND_BIND_NOT_FIXED = 1
    'Private Const MQBND_BIND_ON_OPEN = 0
    'Private Const MQBO_CURRENT_VERSION = 1
    'Private Const MQBO_NONE = 0
    'Private Const MQBO_VERSION_1 = 1
    'Private Const MQBT_OTMA = 1
    'Private Const MQCA_ALTERATION_DATE = 2027
    'Private Const MQCA_ALTERATION_TIME = 2028
    'Private Const MQCA_APPL_ID = 2001
    'Private Const MQCA_BACKOUT_REQ_Q_NAME = 2019
    'Private Const MQCA_BASE_Q_NAME = 2002
    'Private Const MQCA_CF_STRUC_NAME = 2039
    'Private Const MQCA_CHANNEL_AUTO_DEF_EXIT = 2026
    'Private Const MQCA_CLUSTER_DATE = 2037
    'Private Const MQCA_CLUSTER_NAME = 2029
    'Private Const MQCA_CLUSTER_NAMELIST = 2030
    'Private Const MQCA_CLUSTER_Q_MGR_NAME = 2031
    'Private Const MQCA_CLUSTER_TIME = 2038
    'Private Const MQCA_CLUSTER_WORKLOAD_DATA = 2034
    'Private Const MQCA_CLUSTER_WORKLOAD_EXIT = 2033
    'Private Const MQCA_COMMAND_INPUT_Q_NAME = 2003
    'Private Const MQCA_CREATION_DATE = 2004
    'Private Const MQCA_CREATION_TIME = 2005
    'Private Const MQCA_DEAD_LETTER_Q_NAME = 2006
    'Private Const MQCA_DEF_XMIT_Q_NAME = 2025
    'Private Const MQCA_ENV_DATA = 2007
    'Private Const MQCA_FIRST = 2001
    'Private Const MQCA_IGQ_USER_ID = 2041
    'Private Const MQCA_INITIATION_Q_NAME = 2008
    'Private Const MQCA_LAST = 4000
    'Private Const MQCA_LAST_USED = 2041
    'Private Const MQCA_NAMELIST_DESC = 2009
    'Private Const MQCA_NAMELIST_NAME = 2010
    'Private Const MQCA_NAMES = 2020
    'Private Const MQCA_PROCESS_DESC = 2011
    'Private Const MQCA_PROCESS_NAME = 2012
    'Private Const MQCA_Q_DESC = 2013
    'Private Const MQCA_Q_MGR_DESC = 2014
    'Private Const MQCA_Q_MGR_IDENTIFIER = 2032
    'Private Const MQCA_Q_MGR_NAME = 2015
    'Private Const MQCA_Q_NAME = 2016
    'Private Const MQCA_QSG_NAME = 2040
    'Private Const MQCA_REMOTE_Q_MGR_NAME = 2017
    'Private Const MQCA_REMOTE_Q_NAME = 2018
    'Private Const MQCA_REPOSITORY_NAME = 2035
    'Private Const MQCA_REPOSITORY_NAMELIST = 2036
    'Private Const MQCA_STORAGE_CLASS = 2022
    'Private Const MQCA_TRIGGER_DATA = 2023
    'Private Const MQCA_USER_DATA = 2021
    'Private Const MQCA_USER_LIST = 4000
    'Private Const MQCA_XMIT_Q_NAME = 2024
    'Private Const MQCACF_FIRST = 3001
    'Private Const MQCACF_FROM_Q_NAME = 3001
    'Private Const MQCACF_TO_Q_NAME = 3002
    'Private Const MQCACF_FROM_PROCESS_NAME = 3003
    'Private Const MQCACF_TO_PROCESS_NAME = 3004
    'Private Const MQCACF_FROM_NAMELIST_NAME = 3005
    'Private Const MQCACF_TO_NAMELIST_NAME = 3006
    'Private Const MQCACF_FROM_CHANNEL_NAME = 3007
    'Private Const MQCACF_TO_CHANNEL_NAME = 3008
    'Private Const MQCACF_Q_NAMES = 3011
    'Private Const MQCACF_PROCESS_NAMES = 3012
    'Private Const MQCACF_NAMELIST_NAMES = 3013
    'Private Const MQCACF_ESCAPE_TEXT = 3014
    'Private Const MQCACF_LOCAL_Q_NAMES = 3015
    'Private Const MQCACF_MODEL_Q_NAMES = 3016
    'Private Const MQCACF_ALIAS_Q_NAMES = 3017
    'Private Const MQCACF_REMOTE_Q_NAMES = 3018
    'Private Const MQCACF_SENDER_CHANNEL_NAMES = 3019
    'Private Const MQCACF_SERVER_CHANNEL_NAMES = 3020
    'Private Const MQCACF_REQUESTER_CHANNEL_NAMES = 3021
    'Private Const MQCACF_RECEIVER_CHANNEL_NAMES = 3022
    'Private Const MQCACF_OBJECT_Q_MGR_NAME = 3023
    'Private Const MQCACF_APPL_NAME = 3024
    'Private Const MQCACF_USER_IDENTIFIER = 3025
    'Private Const MQCACF_AUX_ERROR_DATA_STR_1 = 3026
    'Private Const MQCACF_AUX_ERROR_DATA_STR_2 = 3027
    'Private Const MQCACF_AUX_ERROR_DATA_STR_3 = 3028
    'Private Const MQCACF_BRIDGE_NAME = 3029
    'Private Const MQCACF_STREAM_NAME = 3030
    'Private Const MQCACF_TOPIC = 3031
    'Private Const MQCACF_PARENT_Q_MGR_NAME = 3032
    'Private Const MQCACF_PUBLISH_TIMESTAMP = 3034
    'Private Const MQCACF_STRING_DATA = 3035
    'Private Const MQCACF_SUPPORTED_STREAM_NAME = 3036
    'Private Const MQCACF_REG_TOPIC = 3037
    'Private Const MQCACF_REG_TIME = 3038
    'Private Const MQCACF_REG_USER_ID = 3039
    'Private Const MQCACF_CHILD_Q_MGR_NAME = 3040
    'Private Const MQCACF_REG_STREAM_NAME = 3041
    'Private Const MQCACF_REG_Q_MGR_NAME = 3042
    'Private Const MQCACF_REG_Q_NAME = 3043
    'Private Const MQCACF_REG_CORREL_ID = 3044
    'Private Const MQCACF_LAST_USED = 3044
    'Private Const MQCACH_FIRST = 3501
    'Private Const MQCACH_CHANNEL_NAME = 3501
    'Private Const MQCACH_DESC = 3502
    'Private Const MQCACH_MODE_NAME = 3503
    'Private Const MQCACH_TP_NAME = 3504
    'Private Const MQCACH_XMIT_Q_NAME = 3505
    'Private Const MQCACH_CONNECTION_NAME = 3506
    'Private Const MQCACH_MCA_NAME = 3507
    'Private Const MQCACH_SEC_EXIT_NAME = 3508
    'Private Const MQCACH_MSG_EXIT_NAME = 3509
    'Private Const MQCACH_SEND_EXIT_NAME = 3510
    'Private Const MQCACH_RCV_EXIT_NAME = 3511
    'Private Const MQCACH_CHANNEL_NAMES = 3512
    'Private Const MQCACH_SEC_EXIT_USER_DATA = 3513
    'Private Const MQCACH_MSG_EXIT_USER_DATA = 3514
    'Private Const MQCACH_SEND_EXIT_USER_DATA = 3515
    'Private Const MQCACH_RCV_EXIT_USER_DATA = 3516
    'Private Const MQCACH_USER_ID = 3517
    'Private Const MQCACH_PASSWORD = 3518
    'Private Const MQCACH_LAST_MSG_TIME = 3524
    'Private Const MQCACH_LAST_MSG_DATE = 3525
    'Private Const MQCACH_MCA_USER_ID = 3527
    'Private Const MQCACH_CHANNEL_START_TIME = 3528
    'Private Const MQCACH_CHANNEL_START_DATE = 3529
    'Private Const MQCACH_MCA_JOB_NAME = 3530
    'Private Const MQCACH_LAST_LUWID = 3531
    'Private Const MQCACH_CURRENT_LUWID = 3532
    'Private Const MQCACH_FORMAT_NAME = 3533
    'Private Const MQCACH_MR_EXIT_NAME = 3534
    'Private Const MQCACH_MR_EXIT_USER_DATA = 3535
    'Private Const MQCACH_LAST_USED = 3535
    'Private Const MQCADSD_NONE = 0
    'Private Const MQCADSD_SEND = 1
    'Private Const MQCADSD_RECV = 16
    'Private Const MQCADSD_MSGFORMAT = 256
    'Private Const MQCBO_NONE = 0
    'Private Const MQCBO_USER_BAG = 0
    'Private Const MQCBO_ADMIN_BAG = 1
    'Private Const MQCBO_COMMAND_BAG = 16
    'Private Const MQCBO_SYSTEM_BAG = 32
    'Private Const MQCBO_LIST_FORM_ALLOWED = 2
    'Private Const MQCBO_LIST_FORM_INHIBITED = 0
    'Private Const MQCBO_REORDER_AS_REQUIRED = 4
    'Private Const MQCBO_DO_NOT_REORDER = 0
    'Private Const MQCBO_CHECK_SELECTORS = 8
    'Private Const MQCBO_DO_NOT_CHECK_SELECTORS = 0
    Public Const MQCC_OK As Integer = 0
    'Private Const MQCC_WARNING = 1
    'Private Const MQCC_FAILED = 2
    'Private Const MQCC_UNKNOWN = &HFFFFFFFF
    'Private Const MQCCSI_DEFAULT = 0
    'Private Const MQCCSI_EMBEDDED = &HFFFFFFFF
    'Private Const MQCCSI_Q_MGR = 0
    'Private Const MQCCT_YES = 1
    'Private Const MQCCT_NO = 0
    'Private Const MQCFC_LAST = 1
    'Private Const MQCFC_NOT_LAST = 0
    'Private Const MQCFH_CURRENT_VERSION = 1
    'Private Const MQCFH_STRUC_LENGTH = 36
    'Private Const MQCFH_VERSION_1 = 1
    'Private Const MQCFIL_STRUC_LENGTH_FIXED = 16
    'Private Const MQCFIN_STRUC_LENGTH = 16
    'Private Const MQCFSL_STRUC_LENGTH_FIXED = 24
    'Private Const MQCFST_STRUC_LENGTH_FIXED = 20
    'Private Const MQCFT_COMMAND = 1
    'Private Const MQCFT_RESPONSE = 2
    'Private Const MQCFT_INTEGER = 3
    'Private Const MQCFT_STRING = 4
    'Private Const MQCFT_INTEGER_LIST = 5
    'Private Const MQCFT_STRING_LIST = 6
    'Private Const MQCFT_EVENT = 7
    'Private Const MQCFT_USER = 8
    'Private Const MQCGWI_DEFAULT = &HFFFFFFFE
    'Private Const MQCHAD_DISABLED = 0
    'Private Const MQCHAD_ENABLED = 1
    'Private Const MQCHIDS_INDOUBT = 1
    'Private Const MQCHIDS_NOT_INDOUBT = 0
    'Private Const MQCHS_INACTIVE = 0
    'Private Const MQCHS_BINDING = 1
    'Private Const MQCHS_STARTING = 2
    'Private Const MQCHS_RUNNING = 3
    'Private Const MQCHS_STOPPING = 4
    'Private Const MQCHS_RETRYING = 5
    'Private Const MQCHS_STOPPED = 6
    'Private Const MQCHS_REQUESTING = 7
    'Private Const MQCHS_PAUSED = 8
    'Private Const MQCHS_INITIALIZING = 13
    'Private Const MQCHSR_STOP_NOT_REQUESTED = 0
    'Private Const MQCHSR_STOP_REQUESTED = 1
    'Private Const MQCHTAB_Q_MGR = 1
    'Private Const MQCHTAB_CLNTCONN = 2
    'Private Const MQCIH_VERSION_1 = 1
    'Private Const MQCIH_VERSION_2 = 2
    'Private Const MQCIH_CURRENT_VERSION = 2
    'Private Const MQCIH_LENGTH_1 = 164
    'Private Const MQCIH_LENGTH_2 = 180
    'Private Const MQCIH_CURRENT_LENGTH = 180
    'Private Const MQCIH_NONE = 0
    'Private Const MQCLT_PROGRAM = 1
    'Private Const MQCLT_TRANSACTION = 2
    'Private Const MQCMD_CHANGE_Q_MGR = 1
    'Private Const MQCMD_INQUIRE_Q_MGR = 2
    'Private Const MQCMD_CHANGE_PROCESS = 3
    'Private Const MQCMD_COPY_PROCESS = 4
    'Private Const MQCMD_CREATE_PROCESS = 5
    'Private Const MQCMD_DELETE_PROCESS = 6
    'Private Const MQCMD_INQUIRE_PROCESS = 7
    'Private Const MQCMD_CHANGE_Q = 8
    'Private Const MQCMD_CLEAR_Q = 9
    'Private Const MQCMD_COPY_Q = 10
    'Private Const MQCMD_CREATE_Q = 11
    'Private Const MQCMD_DELETE_Q = 12
    'Private Const MQCMD_INQUIRE_Q = 13
    'Private Const MQCMD_RESET_Q_STATS = 17
    'Private Const MQCMD_INQUIRE_Q_NAMES = 18
    'Private Const MQCMD_INQUIRE_PROCESS_NAMES = 19
    'Private Const MQCMD_INQUIRE_CHANNEL_NAMES = 20
    'Private Const MQCMD_CHANGE_CHANNEL = 21
    'Private Const MQCMD_COPY_CHANNEL = 22
    'Private Const MQCMD_CREATE_CHANNEL = 23
    'Private Const MQCMD_DELETE_CHANNEL = 24
    'Private Const MQCMD_INQUIRE_CHANNEL = 25
    'Private Const MQCMD_PING_CHANNEL = 26
    'Private Const MQCMD_RESET_CHANNEL = 27
    'Private Const MQCMD_START_CHANNEL = 28
    'Private Const MQCMD_STOP_CHANNEL = 29
    'Private Const MQCMD_START_CHANNEL_INIT = 30
    'Private Const MQCMD_START_CHANNEL_LISTENER = 31
    'Private Const MQCMD_CHANGE_NAMELIST = 32
    'Private Const MQCMD_COPY_NAMELIST = 33
    'Private Const MQCMD_CREATE_NAMELIST = 34
    'Private Const MQCMD_DELETE_NAMELIST = 35
    'Private Const MQCMD_INQUIRE_NAMELIST = 36
    'Private Const MQCMD_INQUIRE_NAMELIST_NAMES = 37
    'Private Const MQCMD_ESCAPE = 38
    'Private Const MQCMD_RESOLVE_CHANNEL = 39
    'Private Const MQCMD_PING_Q_MGR = 40
    'Private Const MQCMD_INQUIRE_CHANNEL_STATUS = 42
    'Private Const MQCMD_Q_MGR_EVENT = 44
    'Private Const MQCMD_PERFM_EVENT = 45
    'Private Const MQCMD_CHANNEL_EVENT = 46
    'Private Const MQCMD_DELETE_PUBLICATION = 60
    'Private Const MQCMD_DEREGISTER_PUBLISHER = 61
    'Private Const MQCMD_DEREGISTER_SUBSCRIBER = 62
    'Private Const MQCMD_PUBLISH = 63
    'Private Const MQCMD_REGISTER_PUBLISHER = 64
    'Private Const MQCMD_REGISTER_SUBSCRIBER = 65
    'Private Const MQCMD_REQUEST_UPDATE = 66
    'Private Const MQCMD_BROKER_INTERNAL = 67
    'Private Const MQCMD_INQUIRE_CLUSTER_Q_MGR = 70
    'Private Const MQCMD_RESUME_Q_MGR_CLUSTER = 71
    'Private Const MQCMD_SUSPEND_Q_MGR_CLUSTER = 72
    'Private Const MQCMD_REFRESH_CLUSTER = 73
    'Private Const MQCMD_RESET_CLUSTER = 74
    'Private Const MQCMD_REFRESH_SECURITY = 78
    'Private Const MQCMD_NONE = 0
    'Private Const MQCMDL_LEVEL_1 = 100
    'Private Const MQCMDL_LEVEL_101 = 101
    'Private Const MQCMDL_LEVEL_110 = 110
    'Private Const MQCMDL_LEVEL_114 = 114
    'Private Const MQCMDL_LEVEL_120 = 120
    'Private Const MQCMDL_LEVEL_200 = 200
    'Private Const MQCMDL_LEVEL_201 = 201
    'Private Const MQCMDL_LEVEL_210 = 210
    'Private Const MQCMDL_LEVEL_220 = 220
    'Private Const MQCMDL_LEVEL_221 = 221
    'Private Const MQCMDL_LEVEL_320 = 320
    'Private Const MQCMDL_LEVEL_420 = 420
    'Private Const MQCMDL_LEVEL_500 = 500
    'Private Const MQCMDL_LEVEL_510 = 510
    'Private Const MQCMDL_LEVEL_520 = 520
    'Private Const MQCNO_CURRENT_VERSION = 2
    'Private Const MQCNO_FASTPATH_BINDING = 1
    'Private Const MQCNO_NONE = 0
    'Private Const MQCNO_STANDARD_BINDING = 0
    'Private Const MQCNO_VERSION_1 = 1
    'Private Const MQCNO_VERSION_2 = 2
    'Private Const MQCO_DELETE = 1
    'Private Const MQCO_DELETE_PURGE = 2
    'Private Const MQCO_NONE = 0
    'Private Const MQCODL_AS_INPUT = &HFFFFFFFF
    'Private Const MQCQT_ALIAS_Q = 2
    'Private Const MQCQT_LOCAL_Q = 1
    'Private Const MQCQT_Q_MGR_ALIAS = 4
    'Private Const MQCQT_REMOTE_Q = 3
    'Private Const MQCRC_OK = 0
    'Private Const MQCRC_CICS_EXEC_ERROR = 1
    'Private Const MQCRC_MQ_API_ERROR = 2
    'Private Const MQCRC_BRIDGE_ERROR = 3
    'Private Const MQCRC_BRIDGE_ABEND = 4
    'Private Const MQCRC_APPLICATION_ABEND = 5
    'Private Const MQCRC_SECURITY_ERROR = 6
    'Private Const MQCRC_PROGRAM_NOT_AVAILABLE = 7
    'Private Const MQCRC_BRIDGE_TIMEOUT = 8
    'Private Const MQCRC_TRANSID_NOT_AVAILABLE = 9
    'Private Const MQCTES_NOSYNC = 0
    'Private Const MQCTES_COMMIT = 256
    'Private Const MQCTES_BACKOUT = 4352
    'Private Const MQCTES_ENDTASK = &H10000
    'Private Const MQCUOWC_ONLY = 273
    'Private Const MQCUOWC_CONTINUE = &H10000
    'Private Const MQCUOWC_FIRST = 17
    'Private Const MQCUOWC_MIDDLE = 16
    'Private Const MQCUOWC_LAST = 272
    'Private Const MQCUOWC_COMMIT = 256
    'Private Const MQCUOWC_BACKOUT = 4352
    'Private Const MQDH_CURRENT_VERSION = 1
    'Private Const MQDH_VERSION_1 = 1
    'Private Const MQDHF_NEW_MSG_IDS = 1
    'Private Const MQDHF_NONE = 0
    'Private Const MQDL_NOT_SUPPORTED = 0
    'Private Const MQDL_SUPPORTED = 1
    'Private Const MQDLH_CURRENT_VERSION = 1
    'Private Const MQDLH_VERSION_1 = 1
    'Private Const MQEI_UNLIMITED = &HFFFFFFFF
    'Private Const MQEC_MSG_ARRIVED = 2
    'Private Const MQEC_WAIT_INTERVAL_EXPIRED = 3
    'Private Const MQEC_WAIT_CANCELED = 4
    'Private Const MQEC_Q_MGR_QUIESCING = 5
    'Private Const MQEC_CONNECTION_QUIESCING = 6
    'Private Const MQENC_NATIVE = 546
    'Private Const MQENC_INTEGER_MASK = 15
    'Private Const MQENC_DECIMAL_MASK = 240
    'Private Const MQENC_FLOAT_MASK = 3840
    'Private Const MQENC_RESERVED_MASK = &HFFFFF000
    'Private Const MQENC_INTEGER_UNDEFINED = 0
    'Private Const MQENC_INTEGER_NORMAL = 1
    'Private Const MQENC_INTEGER_REVERSED = 2
    'Private Const MQENC_DECIMAL_UNDEFINED = 0
    'Private Const MQENC_DECIMAL_NORMAL = 16
    'Private Const MQENC_DECIMAL_REVERSED = 32
    'Private Const MQENC_FLOAT_UNDEFINED = 0
    'Private Const MQENC_FLOAT_IEEE_NORMAL = 256
    'Private Const MQENC_FLOAT_IEEE_REVERSED = 512
    'Private Const MQENC_FLOAT_S390 = 768
    'Private Const MQET_MQSC = 1
    'Private Const MQFB_APPL_CANNOT_BE_STARTED = 265
    'Private Const MQFB_APPL_FIRST = &H10000
    'Private Const MQFB_APPL_LAST = &H3B9AC9FF
    'Private Const MQFB_APPL_TYPE_ERROR = 267
    'Private Const MQFB_BUFFER_OVERFLOW = 294
    'Private Const MQFB_CHANNEL_COMPLETED = 262
    'Private Const MQFB_CHANNEL_FAIL = 264
    'Private Const MQFB_CHANNEL_FAIL_RETRY = 263
    'Private Const MQFB_CICS_INTERNAL_ERROR = 401
    'Private Const MQFB_CICS_NOT_AUTHORIZED = 402
    'Private Const MQFB_CICS_BRIDGE_FAILURE = 403
    'Private Const MQFB_CICS_CORREL_ID_ERROR = 404
    'Private Const MQFB_CICS_CCSID_ERROR = 405
    'Private Const MQFB_CICS_ENCODING_ERROR = 406
    'Private Const MQFB_CICS_CIH_ERROR = 407
    'Private Const MQFB_CICS_UOW_ERROR = 408
    'Private Const MQFB_CICS_COMMAREA_ERROR = 409
    'Private Const MQFB_CICS_APPL_NOT_STARTED = 410
    'Private Const MQFB_CICS_APPL_ABENDED = 411
    'Private Const MQFB_CICS_DLQ_ERROR = 412
    'Private Const MQFB_CICS_UOW_BACKED_OUT = 413
    'Private Const MQFB_COA = 259
    'Private Const MQFB_COD = 260
    'Private Const MQFB_DATA_LENGTH_NEGATIVE = 292
    'Private Const MQFB_DATA_LENGTH_TOO_BIG = 293
    'Private Const MQFB_DATA_LENGTH_ZERO = 291
    'Private Const MQFB_EXPIRATION = 258
    'Private Const MQFB_IIH_ERROR = 296
    'Private Const MQFB_IMS_ERROR = 300
    'Private Const MQFB_IMS_FIRST = 301
    'Private Const MQFB_IMS_LAST = 399
    'Private Const MQFB_LENGTH_OFF_BY_ONE = 295
    'Private Const MQFB_NAN = 276
    'Private Const MQFB_NONE = 0
    'Private Const MQFB_NOT_A_REPOSITORY_MSG = 280
    'Private Const MQFB_NOT_AUTHORIZED_FOR_IMS = 298
    'Private Const MQFB_PAN = 275
    'Private Const MQFB_QUIT = 256
    'Private Const MQFB_STOPPED_BY_CHAD_EXIT = 277
    'Private Const MQFB_STOPPED_BY_MSG_EXIT = 268
    'Private Const MQFB_SYSTEM_FIRST = 1
    'Private Const MQFB_SYSTEM_LAST = 65535
    'Private Const MQFB_TM_ERROR = 266
    'Private Const MQFB_XMIT_Q_MSG_ERROR = 271
    'Private Const MQFC_NO = 0
    'Private Const MQFC_YES = 1
    'Private Const MQGMO_VERSION_1 = 1
    'Private Const MQGMO_VERSION_2 = 2
    'Private Const MQGMO_VERSION_3 = 3
    'Private Const MQGMO_CURRENT_VERSION = 3
    Public Const MQGMO_WAIT As Integer = 1
    'Private Const MQGMO_NO_WAIT = 0
    'Private Const MQGMO_SYNCPOINT = 2
    'Private Const MQGMO_SYNCPOINT_IF_PERSISTENT = 4096
    Public Const MQGMO_NO_SYNCPOINT As Integer = 4
    'Private Const MQGMO_MARK_SKIP_BACKOUT = 128
    'Private Const MQGMO_BROWSE_FIRST = 16
    'Private Const MQGMO_BROWSE_NEXT = 32
    'Private Const MQGMO_BROWSE_MSG_UNDER_CURSOR = 2048
    'Private Const MQGMO_MSG_UNDER_CURSOR = 256
    'Private Const MQGMO_LOCK = 512
    'Private Const MQGMO_UNLOCK = 1024
    'Private Const MQGMO_ACCEPT_TRUNCATED_MSG = 64
    'Private Const MQGMO_SET_SIGNAL = 8
    'Private Const MQGMO_FAIL_IF_QUIESCING = 8192
    'Private Const MQGMO_CONVERT = 16384
    'Private Const MQGMO_LOGICAL_ORDER = 32768
    'Private Const MQGMO_COMPLETE_MSG = &H10000
    'Private Const MQGMO_ALL_MSGS_AVAILABLE = &H20000
    'Private Const MQGMO_ALL_SEGMENTS_AVAILABLE = &H40000
    'Private Const MQGMO_NONE = 0
    'Private Const MQHA_BAG_HANDLE = 4001
    'Private Const MQHA_FIRST = 4001
    'Private Const MQHA_LAST = 6000
    'Private Const MQHA_LAST_USED = 4001
    'Private Const MQHB_NONE = &HFFFFFFFE
    'Private Const MQHB_UNUSABLE_HBAG = &HFFFFFFFF
    'Private Const MQHC_DEF_HCONN = 0
    'Private Const MQHC_UNUSABLE_HCONN = &HFFFFFFFF
    'Private Const MQHO_NONE = 0
    'Private Const MQHO_UNUSABLE_HOBJ = &HFFFFFFFF
    'Private Const MQIA_APPL_TYPE = 1
    'Private Const MQIA_ARCHIVE = 60
    'Private Const MQIA_AUTHORITY_EVENT = 47
    'Private Const MQIA_BACKOUT_THRESHOLD = 22
    'Private Const MQIA_CHANNEL_AUTO_DEF = 55
    'Private Const MQIA_CHANNEL_AUTO_DEF_EVENT = 56
    'Private Const MQIA_CLUSTER_Q_TYPE = 59
    'Private Const MQIA_CLUSTER_WORKLOAD_LENGTH = 58
    'Private Const MQIA_CODED_CHAR_SET_ID = 2
    'Private Const MQIA_COMMAND_LEVEL = 31
    'Private Const MQIA_CURRENT_Q_DEPTH = 3
    'Private Const MQIA_DEF_BIND = 61
    'Private Const MQIA_DEF_INPUT_OPEN_OPTION = 4
    'Private Const MQIA_DEF_PERSISTENCE = 5
    'Private Const MQIA_DEF_PRIORITY = 6
    'Private Const MQIA_DEFINITION_TYPE = 7
    'Private Const MQIA_DIST_LISTS = 34
    'Private Const MQIA_FIRST = 1
    'Private Const MQIA_HARDEN_GET_BACKOUT = 8
    'Private Const MQIA_HIGH_Q_DEPTH = 36
    'Private Const MQIA_IGQ_PUT_AUTHORITY = 65
    'Private Const MQIA_INDEX_TYPE = 57
    'Private Const MQIA_INHIBIT_EVENT = 48
    'Private Const MQIA_INHIBIT_GET = 9
    'Private Const MQIA_INHIBIT_PUT = 10
    'Private Const MQIA_INTRA_GROUP_QUEUING = 64
    'Private Const MQIA_LAST = 2000
    'Private Const MQIA_LAST_USED = 65
    'Private Const MQIA_LOCAL_EVENT = 49
    'Private Const MQIA_MAX_HANDLES = 11
    'Private Const MQIA_MAX_MSG_LENGTH = 13
    'Private Const MQIA_MAX_PRIORITY = 14
    'Private Const MQIA_MAX_Q_DEPTH = 15
    'Private Const MQIA_MAX_UNCOMMITTED_MSGS = 33
    'Private Const MQIA_MSG_DELIVERY_SEQUENCE = 16
    'Private Const MQIA_MSG_DEQ_COUNT = 38
    'Private Const MQIA_MSG_ENQ_COUNT = 37
    'Private Const MQIA_NAME_COUNT = 19
    'Private Const MQIA_OPEN_INPUT_COUNT = 17
    'Private Const MQIA_OPEN_OUTPUT_COUNT = 18
    'Private Const MQIA_PERFORMANCE_EVENT = 53
    'Private Const MQIA_PLATFORM = 32
    'Private Const MQIA_Q_DEPTH_HIGH_EVENT = 43
    'Private Const MQIA_Q_DEPTH_HIGH_LIMIT = 40
    'Private Const MQIA_Q_DEPTH_LOW_EVENT = 44
    'Private Const MQIA_Q_DEPTH_LOW_LIMIT = 41
    'Private Const MQIA_Q_DEPTH_MAX_EVENT = 42
    'Private Const MQIA_Q_SERVICE_INTERVAL = 54
    'Private Const MQIA_Q_SERVICE_INTERVAL_EVENT = 46
    'Private Const MQIA_Q_TYPE = 20
    'Private Const MQIA_QSG_DISP = 63
    'Private Const MQIA_REMOTE_EVENT = 50
    'Private Const MQIA_RETENTION_INTERVAL = 21
    'Private Const MQIA_SCOPE = 45
    'Private Const MQIA_SHAREABILITY = 23
    'Private Const MQIA_START_STOP_EVENT = 52
    'Private Const MQIA_SYNCPOINT = 30
    'Private Const MQIA_TIME_SINCE_RESET = 35
    'Private Const MQIA_TRIGGER_CONTROL = 24
    'Private Const MQIA_TRIGGER_DEPTH = 29
    'Private Const MQIA_TRIGGER_INTERVAL = 25
    'Private Const MQIA_TRIGGER_MSG_PRIORITY = 26
    'Private Const MQIA_TRIGGER_TYPE = 28
    'Private Const MQIA_USAGE = 12
    'Private Const MQIA_USER_LIST = 2000
    'Private Const MQIACF_FIRST = 1001
    'Private Const MQIACF_Q_MGR_ATTRS = 1001
    'Private Const MQIACF_Q_ATTRS = 1002
    'Private Const MQIACF_PROCESS_ATTRS = 1003
    'Private Const MQIACF_NAMELIST_ATTRS = 1004
    'Private Const MQIACF_FORCE = 1005
    'Private Const MQIACF_REPLACE = 1006
    'Private Const MQIACF_PURGE = 1007
    'Private Const MQIACF_QUIESCE = 1008
    'Private Const MQIACF_ALL = 1009
    'Private Const MQIACF_PARAMETER_ID = 1012
    'Private Const MQIACF_ERROR_ID = 1013
    'Private Const MQIACF_ERROR_IDENTIFIER = 1013
    'Private Const MQIACF_SELECTOR = 1014
    'Private Const MQIACF_CHANNEL_ATTRS = 1015
    'Private Const MQIACF_ESCAPE_TYPE = 1017
    'Private Const MQIACF_ERROR_OFFSET = 1018
    'Private Const MQIACF_REASON_QUALIFIER = 1020
    'Private Const MQIACF_COMMAND = 1021
    'Private Const MQIACF_OPEN_OPTIONS = 1022
    'Private Const MQIACF_AUX_ERROR_DATA_INT_1 = 1070
    'Private Const MQIACF_AUX_ERROR_DATA_INT_2 = 1071
    'Private Const MQIACF_CONV_REASON_CODE = 1072
    'Private Const MQIACF_BRIDGE_TYPE = 1073
    'Private Const MQIACF_INQUIRY = 1074
    'Private Const MQIACF_WAIT_INTERVAL = 1075
    'Private Const MQIACF_OPTIONS = 1076
    'Private Const MQIACF_BROKER_OPTIONS = 1077
    'Private Const MQIACF_SEQUENCE_NUMBER = 1079
    'Private Const MQIACF_INTEGER_DATA = 1080
    'Private Const MQIACF_REGISTRATION_OPTIONS = 1081
    'Private Const MQIACF_PUBLICATION_OPTIONS = 1082
    'Private Const MQIACF_CLUSTER_INFO = 1083
    'Private Const MQIACF_Q_MGR_DEFINITION_TYPE = 1084
    'Private Const MQIACF_Q_MGR_TYPE = 1085
    'Private Const MQIACF_ACTION = 1086
    'Private Const MQIACF_SUSPEND = 1087
    'Private Const MQIACF_BROKER_COUNT = 1088
    'Private Const MQIACF_APPL_COUNT = 1089
    'Private Const MQIACF_ANONYMOUS_COUNT = 1090
    'Private Const MQIACF_REG_REG_OPTIONS = 1091
    'Private Const MQIACF_DELETE_OPTIONS = 1092
    'Private Const MQIACF_CLUSTER_Q_MGR_ATTRS = 1093
    'Private Const MQIACF_LAST_USED = 1093
    'Private Const MQIACH_FIRST = 1501
    'Private Const MQIACH_XMIT_PROTOCOL_TYPE = 1501
    'Private Const MQIACH_BATCH_SIZE = 1502
    'Private Const MQIACH_DISC_INTERVAL = 1503
    'Private Const MQIACH_SHORT_TIMER = 1504
    'Private Const MQIACH_SHORT_RETRY = 1505
    'Private Const MQIACH_LONG_TIMER = 1506
    'Private Const MQIACH_LONG_RETRY = 1507
    'Private Const MQIACH_PUT_AUTHORITY = 1508
    'Private Const MQIACH_SEQUENCE_NUMBER_WRAP = 1509
    'Private Const MQIACH_MAX_MSG_LENGTH = 1510
    'Private Const MQIACH_CHANNEL_TYPE = 1511
    'Private Const MQIACH_DATA_COUNT = 1512
    'Private Const MQIACH_MSG_SEQUENCE_NUMBER = 1514
    'Private Const MQIACH_DATA_CONVERSION = 1515
    'Private Const MQIACH_IN_DOUBT = 1516
    'Private Const MQIACH_MCA_TYPE = 1517
    'Private Const MQIACH_CHANNEL_INSTANCE_TYPE = 1523
    'Private Const MQIACH_CHANNEL_INSTANCE_ATTRS = 1524
    'Private Const MQIACH_CHANNEL_ERROR_DATA = 1525
    'Private Const MQIACH_CHANNEL_TABLE = 1526
    'Private Const MQIACH_CHANNEL_STATUS = 1527
    'Private Const MQIACH_INDOUBT_STATUS = 1528
    'Private Const MQIACH_LAST_SEQ_NUMBER = 1529
    'Private Const MQIACH_LAST_SEQUENCE_NUMBER = 1529
    'Private Const MQIACH_CURRENT_MSGS = 1531
    'Private Const MQIACH_CURRENT_SEQ_NUMBER = 1532
    'Private Const MQIACH_CURRENT_SEQUENCE_NUMBER = 1532
    'Private Const MQIACH_MSGS = 1534
    'Private Const MQIACH_BYTES_SENT = 1535
    'Private Const MQIACH_BYTES_RCVD = 1536
    'Private Const MQIACH_BYTES_RECEIVED = 1536
    'Private Const MQIACH_BATCHES = 1537
    'Private Const MQIACH_BUFFERS_SENT = 1538
    'Private Const MQIACH_BUFFERS_RCVD = 1539
    'Private Const MQIACH_BUFFERS_RECEIVED = 1539
    'Private Const MQIACH_LONG_RETRIES_LEFT = 1540
    'Private Const MQIACH_SHORT_RETRIES_LEFT = 1541
    'Private Const MQIACH_MCA_STATUS = 1542
    'Private Const MQIACH_STOP_REQUESTED = 1543
    'Private Const MQIACH_MR_COUNT = 1544
    'Private Const MQIACH_MR_INTERVAL = 1545
    'Private Const MQIACH_NPM_SPEED = 1562
    'Private Const MQIACH_HB_INTERVAL = 1563
    'Private Const MQIACH_BATCH_INTERVAL = 1564
    'Private Const MQIACH_NETWORK_PRIORITY = 1565
    'Private Const MQIACH_LAST_USED = 1565
    'Private Const MQIASY_BAG_OPTIONS = &HFFFFFFF8
    'Private Const MQIASY_CODED_CHAR_SET_ID = &HFFFFFFFF
    'Private Const MQIASY_COMMAND = &HFFFFFFFD
    'Private Const MQIASY_COMP_CODE = &HFFFFFFFA
    'Private Const MQIASY_CONTROL = &HFFFFFFFB
    'Private Const MQIASY_FIRST = &HFFFFFFFF
    'Private Const MQIASY_LAST = &HFFFFF830
    'Private Const MQIASY_LAST_USED = &HFFFFFFF8
    'Private Const MQIASY_MSG_SEQ_NUMBER = &HFFFFFFFC
    'Private Const MQIASY_REASON = &HFFFFFFF9
    'Private Const MQIASY_TYPE = &HFFFFFFFE
    'Private Const MQIAV_NOT_APPLICABLE = &HFFFFFFFF
    'Private Const MQIAV_UNDEFINED = &HFFFFFFFE
    'Private Const MQIDO_BACKOUT = 2
    'Private Const MQIDO_COMMIT = 1
    'Private Const MQIIH_CURRENT_VERSION = 1
    'Private Const MQIIH_LENGTH_1 = 84
    'Private Const MQIIH_NONE = 0
    'Private Const MQIIH_VERSION_1 = 1
    'Private Const MQIND_NONE = &HFFFFFFFF
    'Private Const MQIND_ALL = &HFFFFFFFE
    'Private Const MQIT_NONE = 0
    'Private Const MQIT_INTEGER = 1
    'Private Const MQIT_MSG_ID = 1
    'Private Const MQIT_STRING = 2
    'Private Const MQIT_CORREL_ID = 2
    'Private Const MQIT_BAG = 3
    'Private Const MQIT_MSG_TOKEN = 4
    'Private Const MQMCAS_RUNNING = 3
    'Private Const MQMCAS_STOPPED = 0
    'Private Const MQMD_CURRENT_VERSION = 2
    'Private Const MQMD_VERSION_1 = 1
    'Private Const MQMD_VERSION_2 = 2
    'Private Const MQMDE_CURRENT_VERSION = 2
    'Private Const MQMDE_LENGTH_2 = 72
    'Private Const MQMDE_VERSION_2 = 2
    'Private Const MQMDEF_NONE = 0
    'Private Const MQMDS_FIFO = 1
    'Private Const MQMDS_PRIORITY = 0
    'Private Const MQMF_ACCEPT_UNSUP_IF_XMIT_MASK = &HFF000
    'Private Const MQMF_ACCEPT_UNSUP_MASK = &HFFF00000
    'Private Const MQMF_LAST_MSG_IN_GROUP = 16
    'Private Const MQMF_LAST_SEGMENT = 4
    'Private Const MQMF_MSG_IN_GROUP = 8
    'Private Const MQMF_NONE = 0
    'Private Const MQMF_REJECT_UNSUP_MASK = 4095
    'Private Const MQMF_SEGMENT = 2
    'Private Const MQMF_SEGMENTATION_ALLOWED = 1
    'Private Const MQMF_SEGMENTATION_INHIBITED = 0
    Public Const MQMO_MATCH_MSG_ID As Integer = 1
    Public Const MQMO_MATCH_CORREL_ID As Integer = 2
    'Private Const MQMO_MATCH_GROUP_ID = 4
    'Private Const MQMO_MATCH_MSG_SEQ_NUMBER = 8
    'Private Const MQMO_MATCH_OFFSET = 16
    'Private Const MQMO_MATCH_MSG_TOKEN = 32
    'Private Const MQMO_NONE = 0
    'Private Const MQMT_SYSTEM_FIRST = 1
    'Private Const MQMT_REQUEST = 1
    'Private Const MQMT_REPLY = 2
    'Private Const MQMT_DATAGRAM = 8
    'Private Const MQMT_REPORT = 4
    'Private Const MQMT_MQE_FIELDS_FROM_MQE = 112
    'Private Const MQMT_MQE_FIELDS = 113
    'Private Const MQMT_SYSTEM_LAST = 65535
    'Private Const MQMT_APPL_FIRST = &H10000
    'Private Const MQMT_APPL_LAST = &H3B9AC9FF
    'Private Const MQNC_MAX_NAMELIST_NAME_COUNT = 256
    'Private Const MQOA_FIRST = 1
    'Private Const MQOA_LAST = 6000
    'Private Const MQOD_VERSION_1 = 1
    'Private Const MQOD_VERSION_2 = 2
    'Private Const MQOD_VERSION_3 = 3
    'Private Const MQOD_CURRENT_VERSION = 3
    'Private Const MQOD_CURRENT_LENGTH = 336
    'Private Const MQOL_UNDEFINED = &HFFFFFFFF
    'Private Const MQOO_ALTERNATE_USER_AUTHORITY = 4096
    'Private Const MQOO_BIND_AS_Q_DEF = 0
    'Private Const MQOO_BIND_NOT_FIXED = 32768
    'Private Const MQOO_BIND_ON_OPEN = 16384
    'Private Const MQOO_BROWSE = 8
    'Private Const MQOO_FAIL_IF_QUIESCING = 8192
    Public Const MQOO_INPUT_AS_Q_DEF As Integer = 1
    'Private Const MQOO_INPUT_EXCLUSIVE = 4
    'Private Const MQOO_INPUT_SHARED = 2
    'Private Const MQOO_INQUIRE = 32
    Public Const MQOO_OUTPUT As Integer = 16
    'Private Const MQOO_PASS_ALL_CONTEXT = 512
    'Private Const MQOO_PASS_IDENTITY_CONTEXT = 256
    'Private Const MQOO_SAVE_ALL_CONTEXT = 128
    'Private Const MQOO_SET = 64
    'Private Const MQOO_SET_ALL_CONTEXT = 2048
    'Private Const MQOO_SET_IDENTITY_CONTEXT = 1024
    'Private Const MQOT_ALIAS_Q = 1002
    'Private Const MQOT_ALL = 1001
    'Private Const MQOT_CHANNEL = 6
    'Private Const MQOT_CLNTCONN_CHANNEL = 1014
    'Private Const MQOT_CURRENT_CHANNEL = 1011
    'Private Const MQOT_LOCAL_Q = 1004
    'Private Const MQOT_MODEL_Q = 1003
    'Private Const MQOT_NAMELIST = 2
    'Private Const MQOT_PROCESS = 3
    'Private Const MQOT_Q = 1
    'Private Const MQOT_Q_MGR = 5
    'Private Const MQOT_RECEIVER_CHANNEL = 1010
    'Private Const MQOT_REMOTE_Q = 1005
    'Private Const MQOT_REQUESTER_CHANNEL = 1009
    'Private Const MQOT_RESERVED_1 = 7
    'Private Const MQOO_RESOLVE_NAMES = &H10000
    'Private Const MQOT_SAVED_CHANNEL = 1012
    'Private Const MQOT_SENDER_CHANNEL = 1007
    'Private Const MQOT_SERVER_CHANNEL = 1008
    'Private Const MQOT_SVRCONN_CHANNEL = 1013
    'Private Const MQPER_NOT_PERSISTENT = 0
    'Private Const MQPER_PERSISTENCE_AS_Q_DEF = 2
    'Private Const MQPER_PERSISTENT = 1
    'Private Const MQPL_AIX = 3
    'Private Const MQPL_MVS = 1
    'Private Const MQPL_NSK = 13
    'Private Const MQPL_OS2 = 2
    'Private Const MQPL_OS390 = 1
    'Private Const MQPL_OS400 = 4
    'Private Const MQPL_UNIX = 3
    'Private Const MQPL_VMS = 12
    'Private Const MQPL_WINDOWS = 5
    'Private Const MQPL_WINDOWS_NT = 11
    'Private Const MQPMO_ALTERNATE_USER_AUTHORITY = 4096
    'Private Const MQPMO_CURRENT_LENGTH = 152
    'Private Const MQPMO_CURRENT_VERSION = 2
    'Private Const MQPMO_DEFAULT_CONTEXT = 32
    'Private Const MQPMO_FAIL_IF_QUIESCING = 8192
    'Private Const MQPMO_LOGICAL_ORDER = 32768
    'Private Const MQPMO_NEW_CORREL_ID = 128
    'Private Const MQPMO_NEW_MSG_ID = 64
    'Private Const MQPMO_NO_CONTEXT = 16384
    Public Const MQPMO_NO_SYNCPOINT As Integer = 4
    'Private Const MQPMO_NONE = 0
    'Private Const MQPMO_PASS_ALL_CONTEXT = 512
    'Private Const MQPMO_PASS_IDENTITY_CONTEXT = 256
    'Private Const MQPMO_SET_ALL_CONTEXT = 2048
    'Private Const MQPMO_SET_IDENTITY_CONTEXT = 1024
    'Private Const MQPMO_SYNCPOINT = 2
    'Private Const MQPMO_VERSION_1 = 1
    'Private Const MQPMO_VERSION_2 = 2
    'Private Const MQPMRF_ACCOUNTING_TOKEN = 16
    'Private Const MQPMRF_CORREL_ID = 2
    'Private Const MQPMRF_FEEDBACK = 8
    'Private Const MQPMRF_GROUP_ID = 4
    'Private Const MQPMRF_MSG_ID = 1
    'Private Const MQPMRF_NONE = 0
    'Private Const MQPO_NO = 0
    'Private Const MQPO_YES = 1
    'Private Const MQPRI_PRIORITY_AS_Q_DEF = &HFFFFFFFF
    'Private Const MQQA_BACKOUT_HARDENED = 1
    'Private Const MQQA_BACKOUT_NOT_HARDENED = 0
    'Private Const MQQA_GET_ALLOWED = 0
    'Private Const MQQA_GET_INHIBITED = 1
    'Private Const MQQA_NOT_SHAREABLE = 0
    'Private Const MQQA_PUT_ALLOWED = 0
    'Private Const MQQA_PUT_INHIBITED = 1
    'Private Const MQQA_SHAREABLE = 1
    'Private Const MQQDT_PERMANENT_DYNAMIC = 2
    'Private Const MQQDT_PREDEFINED = 1
    'Private Const MQQDT_TEMPORARY_DYNAMIC = 3
    'Private Const MQQDT_SHARED_DYNAMIC = 4
    'Private Const MQQMDT_AUTO_CLUSTER_SENDER = 2
    'Private Const MQQMDT_AUTO_EXP_CLUSTER_SENDER = 4
    'Private Const MQQMDT_CLUSTER_RECEIVER = 3
    'Private Const MQQMDT_EXPLICIT_CLUSTER_SENDER = 1
    'Private Const MQQMT_NORMAL = 0
    'Private Const MQQMT_REPOSITORY = 1
    'Private Const MQQO_NO = 0
    'Private Const MQQO_YES = 1
    'Private Const MQQT_ALIAS = 3
    'Private Const MQQT_ALL = 1001
    'Private Const MQQT_CLUSTER = 7
    'Private Const MQQT_LOCAL = 1
    'Private Const MQQT_MODEL = 2
    'Private Const MQQT_REMOTE = 6
    'Private Const MQRC_NONE = 0
    'Private Const MQRC_ALIAS_BASE_Q_TYPE_ERROR = 2001
    'Private Const MQRC_ALREADY_CONNECTED = 2002
    'Private Const MQRC_BACKED_OUT = 2003
    'Private Const MQRC_BUFFER_ERROR = 2004
    'Private Const MQRC_BUFFER_LENGTH_ERROR = 2005
    'Private Const MQRC_CHAR_ATTR_LENGTH_ERROR = 2006
    'Private Const MQRC_CHAR_ATTRS_ERROR = 2007
    'Private Const MQRC_CHAR_ATTRS_TOO_SHORT = 2008
    Public Const MQRC_CONNECTION_BROKEN As Integer = 2009
    'Private Const MQRC_DATA_LENGTH_ERROR = 2010
    'Private Const MQRC_DYNAMIC_Q_NAME_ERROR = 2011
    'Private Const MQRC_ENVIRONMENT_ERROR = 2012
    'Private Const MQRC_EXPIRY_ERROR = 2013
    'Private Const MQRC_FEEDBACK_ERROR = 2014
    'Private Const MQRC_GET_INHIBITED = 2016
    'Private Const MQRC_HANDLE_NOT_AVAILABLE = 2017
    'Private Const MQRC_HCONN_ERROR = 2018
    'Private Const MQRC_HOBJ_ERROR = 2019
    'Private Const MQRC_INHIBIT_VALUE_ERROR = 2020
    'Private Const MQRC_INT_ATTR_COUNT_ERROR = 2021
    'Private Const MQRC_INT_ATTR_COUNT_TOO_SMALL = 2022
    'Private Const MQRC_INT_ATTRS_ARRAY_ERROR = 2023
    'Private Const MQRC_SYNCPOINT_LIMIT_REACHED = 2024
    'Private Const MQRC_MAX_CONNS_LIMIT_REACHED = 2025
    'Private Const MQRC_MD_ERROR = 2026
    'Private Const MQRC_MISSING_REPLY_TO_Q = 2027
    'Private Const MQRC_MSG_TYPE_ERROR = 2029
    'Private Const MQRC_MSG_TOO_BIG_FOR_Q = 2030
    'Private Const MQRC_MSG_TOO_BIG_FOR_Q_MGR = 2031
    Public Const MQRC_NO_MSG_AVAILABLE As Integer = 2033
    'Private Const MQRC_NO_MSG_UNDER_CURSOR = 2034
    'Private Const MQRC_NOT_AUTHORIZED = 2035
    'Private Const MQRC_NOT_OPEN_FOR_BROWSE = 2036
    'Private Const MQRC_NOT_OPEN_FOR_INPUT = 2037
    'Private Const MQRC_NOT_OPEN_FOR_INQUIRE = 2038
    'Private Const MQRC_NOT_OPEN_FOR_OUTPUT = 2039
    'Private Const MQRC_NOT_OPEN_FOR_SET = 2040
    'Private Const MQRC_OBJECT_CHANGED = 2041
    'Private Const MQRC_OBJECT_IN_USE = 2042
    'Private Const MQRC_OBJECT_TYPE_ERROR = 2043
    'Private Const MQRC_OD_ERROR = 2044
    'Private Const MQRC_OPTION_NOT_VALID_FOR_TYPE = 2045
    'Private Const MQRC_OPTIONS_ERROR = 2046
    'Private Const MQRC_PERSISTENCE_ERROR = 2047
    'Private Const MQRC_PERSISTENT_NOT_ALLOWED = 2048
    'Private Const MQRC_PRIORITY_EXCEEDS_MAXIMUM = 2049
    'Private Const MQRC_PRIORITY_ERROR = 2050
    'Private Const MQRC_PUT_INHIBITED = 2051
    'Private Const MQRC_Q_DELETED = 2052
    'Private Const MQRC_Q_FULL = 2053
    'Private Const MQRC_Q_NOT_EMPTY = 2055
    'Private Const MQRC_Q_SPACE_NOT_AVAILABLE = 2056
    'Private Const MQRC_Q_TYPE_ERROR = 2057
    'Private Const MQRC_Q_MGR_NAME_ERROR = 2058
    'Private Const MQRC_Q_MGR_NOT_AVAILABLE = 2059
    'Private Const MQRC_REPORT_OPTIONS_ERROR = 2061
    'Private Const MQRC_SECOND_MARK_NOT_ALLOWED = 2062
    'Private Const MQRC_SECURITY_ERROR = 2063
    'Private Const MQRC_SELECTOR_COUNT_ERROR = 2065
    'Private Const MQRC_SELECTOR_LIMIT_EXCEEDED = 2066
    'Private Const MQRC_SELECTOR_ERROR = 2067
    'Private Const MQRC_SELECTOR_NOT_FOR_TYPE = 2068
    'Private Const MQRC_SIGNAL_OUTSTANDING = 2069
    'Private Const MQRC_SIGNAL_REQUEST_ACCEPTED = 2070
    'Private Const MQRC_STORAGE_NOT_AVAILABLE = 2071
    'Private Const MQRC_SYNCPOINT_NOT_AVAILABLE = 2072
    'Private Const MQRC_TRIGGER_CONTROL_ERROR = 2075
    'Private Const MQRC_TRIGGER_DEPTH_ERROR = 2076
    'Private Const MQRC_TRIGGER_MSG_PRIORITY_ERR = 2077
    'Private Const MQRC_TRIGGER_TYPE_ERROR = 2078
    'Private Const MQRC_TRUNCATED_MSG_ACCEPTED = 2079
    'Private Const MQRC_TRUNCATED_MSG_FAILED = 2080
    'Private Const MQRC_UNKNOWN_ALIAS_BASE_Q = 2082
    'Private Const MQRC_UNKNOWN_OBJECT_NAME = 2085
    'Private Const MQRC_UNKNOWN_OBJECT_Q_MGR = 2086
    'Private Const MQRC_UNKNOWN_REMOTE_Q_MGR = 2087
    'Private Const MQRC_WAIT_INTERVAL_ERROR = 2090
    'Private Const MQRC_XMIT_Q_TYPE_ERROR = 2091
    'Private Const MQRC_XMIT_Q_USAGE_ERROR = 2092
    'Private Const MQRC_NOT_OPEN_FOR_PASS_ALL = 2093
    'Private Const MQRC_NOT_OPEN_FOR_PASS_IDENT = 2094
    'Private Const MQRC_NOT_OPEN_FOR_SET_ALL = 2095
    'Private Const MQRC_NOT_OPEN_FOR_SET_IDENT = 2096
    'Private Const MQRC_CONTEXT_HANDLE_ERROR = 2097
    'Private Const MQRC_CONTEXT_NOT_AVAILABLE = 2098
    'Private Const MQRC_SIGNAL1_ERROR = 2099
    'Private Const MQRC_OBJECT_ALREADY_EXISTS = 2100
    'Private Const MQRC_OBJECT_DAMAGED = 2101
    'Private Const MQRC_RESOURCE_PROBLEM = 2102
    'Private Const MQRC_ANOTHER_Q_MGR_CONNECTED = 2103
    'Private Const MQRC_UNKNOWN_REPORT_OPTION = 2104
    'Private Const MQRC_STORAGE_CLASS_ERROR = 2105
    'Private Const MQRC_COD_NOT_VALID_FOR_XCF_Q = 2106
    'Private Const MQRC_XWAIT_CANCELED = 2107
    'Private Const MQRC_XWAIT_ERROR = 2108
    'Private Const MQRC_SUPPRESSED_BY_EXIT = 2109
    'Private Const MQRC_FORMAT_ERROR = 2110
    'Private Const MQRC_SOURCE_CCSID_ERROR = 2111
    'Private Const MQRC_SOURCE_INTEGER_ENC_ERROR = 2112
    'Private Const MQRC_SOURCE_DECIMAL_ENC_ERROR = 2113
    'Private Const MQRC_SOURCE_FLOAT_ENC_ERROR = 2114
    'Private Const MQRC_TARGET_CCSID_ERROR = 2115
    'Private Const MQRC_TARGET_INTEGER_ENC_ERROR = 2116
    'Private Const MQRC_TARGET_DECIMAL_ENC_ERROR = 2117
    'Private Const MQRC_TARGET_FLOAT_ENC_ERROR = 2118
    'Private Const MQRC_NOT_CONVERTED = 2119
    'Private Const MQRC_CONVERTED_MSG_TOO_BIG = 2120
    'Private Const MQRC_TRUNCATED = 2120
    'Private Const MQRC_NO_EXTERNAL_PARTICIPANTS = 2121
    'Private Const MQRC_PARTICIPANT_NOT_AVAILABLE = 2122
    'Private Const MQRC_OUTCOME_MIXED = 2123
    'Private Const MQRC_OUTCOME_PENDING = 2124
    'Private Const MQRC_BRIDGE_STARTED = 2125
    'Private Const MQRC_BRIDGE_STOPPED = 2126
    'Private Const MQRC_ADAPTER_STORAGE_SHORTAGE = 2127
    'Private Const MQRC_UOW_IN_PROGRESS = 2128
    'Private Const MQRC_ADAPTER_CONN_LOAD_ERROR = 2129
    'Private Const MQRC_ADAPTER_SERV_LOAD_ERROR = 2130
    'Private Const MQRC_ADAPTER_DEFS_ERROR = 2131
    'Private Const MQRC_ADAPTER_DEFS_LOAD_ERROR = 2132
    'Private Const MQRC_ADAPTER_CONV_LOAD_ERROR = 2133
    'Private Const MQRC_BO_ERROR = 2134
    'Private Const MQRC_DH_ERROR = 2135
    'Private Const MQRC_MULTIPLE_REASONS = 2136
    'Private Const MQRC_OPEN_FAILED = 2137
    'Private Const MQRC_ADAPTER_DISC_LOAD_ERROR = 2138
    'Private Const MQRC_CNO_ERROR = 2139
    'Private Const MQRC_CICS_WAIT_FAILED = 2140
    'Private Const MQRC_DLH_ERROR = 2141
    'Private Const MQRC_HEADER_ERROR = 2142
    'Private Const MQRC_SOURCE_LENGTH_ERROR = 2143
    'Private Const MQRC_TARGET_LENGTH_ERROR = 2144
    'Private Const MQRC_SOURCE_BUFFER_ERROR = 2145
    'Private Const MQRC_TARGET_BUFFER_ERROR = 2146
    'Private Const MQRC_IIH_ERROR = 2148
    'Private Const MQRC_PCF_ERROR = 2149
    'Private Const MQRC_DBCS_ERROR = 2150
    'Private Const MQRC_OBJECT_NAME_ERROR = 2152
    'Private Const MQRC_OBJECT_Q_MGR_NAME_ERROR = 2153
    'Private Const MQRC_RECS_PRESENT_ERROR = 2154
    'Private Const MQRC_OBJECT_RECORDS_ERROR = 2155
    'Private Const MQRC_RESPONSE_RECORDS_ERROR = 2156
    'Private Const MQRC_ASID_MISMATCH = 2157
    'Private Const MQRC_PMO_RECORD_FLAGS_ERROR = 2158
    'Private Const MQRC_PUT_MSG_RECORDS_ERROR = 2159
    'Private Const MQRC_CONN_ID_IN_USE = 2160
    'Private Const MQRC_Q_MGR_QUIESCING = 2161
    'Private Const MQRC_Q_MGR_STOPPING = 2162
    'Private Const MQRC_DUPLICATE_RECOV_COORD = 2163
    'Private Const MQRC_PMO_ERROR = 2173
    'Private Const MQRC_API_EXIT_NOT_FOUND = 2182
    'Private Const MQRC_API_EXIT_LOAD_ERROR = 2183
    'Private Const MQRC_REMOTE_Q_NAME_ERROR = 2184
    'Private Const MQRC_INCONSISTENT_PERSISTENCE = 2185
    'Private Const MQRC_GMO_ERROR = 2186
    'Private Const MQRC_CICS_BRIDGE_RESTRICTION = 2187
    'Private Const MQRC_STOPPED_BY_CLUSTER_EXIT = 2188
    'Private Const MQRC_CLUSTER_RESOLUTION_ERROR = 2189
    'Private Const MQRC_CONVERTED_STRING_TOO_BIG = 2190
    'Private Const MQRC_TMC_ERROR = 2191
    'Private Const MQRC_PAGESET_FULL = 2192
    'Private Const MQRC_STORAGE_MEDIUM_FULL = 2192
    'Private Const MQRC_PAGESET_ERROR = 2193
    'Private Const MQRC_NAME_NOT_VALID_FOR_TYPE = 2194
    'Private Const MQRC_UNEXPECTED_ERROR = 2195
    'Private Const MQRC_UNKNOWN_XMIT_Q = 2196
    'Private Const MQRC_UNKNOWN_DEF_XMIT_Q = 2197
    'Private Const MQRC_DEF_XMIT_Q_TYPE_ERROR = 2198
    'Private Const MQRC_DEF_XMIT_Q_USAGE_ERROR = 2199
    'Private Const MQRC_NAME_IN_USE = 2201
    'Private Const MQRC_CONNECTION_QUIESCING = 2202
    'Private Const MQRC_CONNECTION_STOPPING = 2203
    'Private Const MQRC_ADAPTER_NOT_AVAILABLE = 2204
    'Private Const MQRC_MSG_ID_ERROR = 2206
    'Private Const MQRC_CORREL_ID_ERROR = 2207
    'Private Const MQRC_FILE_SYSTEM_ERROR = 2208
    'Private Const MQRC_NO_MSG_LOCKED = 2209
    'Private Const MQRC_FILE_NOT_AUDITED = 2216
    'Private Const MQRC_CONNECTION_NOT_AUTHORIZED = 2217
    'Private Const MQRC_MSG_TOO_BIG_FOR_CHANNEL = 2218
    'Private Const MQRC_CALL_IN_PROGRESS = 2219
    'Private Const MQRC_RMH_ERROR = 2220
    'Private Const MQRC_Q_MGR_ACTIVE = 2222
    'Private Const MQRC_Q_MGR_NOT_ACTIVE = 2223
    'Private Const MQRC_Q_DEPTH_HIGH = 2224
    'Private Const MQRC_Q_DEPTH_LOW = 2225
    'Private Const MQRC_Q_SERVICE_INTERVAL_HIGH = 2226
    'Private Const MQRC_Q_SERVICE_INTERVAL_OK = 2227
    'Private Const MQRC_UNIT_OF_WORK_NOT_STARTED = 2232
    'Private Const MQRC_CHANNEL_AUTO_DEF_OK = 2233
    'Private Const MQRC_CHANNEL_AUTO_DEF_ERROR = 2234
    'Private Const MQRC_CFH_ERROR = 2235
    'Private Const MQRC_CFIL_ERROR = 2236
    'Private Const MQRC_CFIN_ERROR = 2237
    'Private Const MQRC_CFSL_ERROR = 2238
    'Private Const MQRC_CFST_ERROR = 2239
    'Private Const MQRC_INCOMPLETE_GROUP = 2241
    'Private Const MQRC_INCOMPLETE_MSG = 2242
    'Private Const MQRC_INCONSISTENT_CCSIDS = 2243
    'Private Const MQRC_INCONSISTENT_ENCODINGS = 2244
    'Private Const MQRC_INCONSISTENT_UOW = 2245
    'Private Const MQRC_INVALID_MSG_UNDER_CURSOR = 2246
    'Private Const MQRC_MATCH_OPTIONS_ERROR = 2247
    'Private Const MQRC_MDE_ERROR = 2248
    'Private Const MQRC_MSG_FLAGS_ERROR = 2249
    'Private Const MQRC_MSG_SEQ_NUMBER_ERROR = 2250
    'Private Const MQRC_OFFSET_ERROR = 2251
    'Private Const MQRC_ORIGINAL_LENGTH_ERROR = 2252
    'Private Const MQRC_SEGMENT_LENGTH_ZERO = 2253
    'Private Const MQRC_UOW_NOT_AVAILABLE = 2255
    'Private Const MQRC_WRONG_GMO_VERSION = 2256
    'Private Const MQRC_WRONG_MD_VERSION = 2257
    'Private Const MQRC_GROUP_ID_ERROR = 2258
    'Private Const MQRC_INCONSISTENT_BROWSE = 2259
    'Private Const MQRC_XQH_ERROR = 2260
    'Private Const MQRC_SRC_ENV_ERROR = 2261
    'Private Const MQRC_SRC_NAME_ERROR = 2262
    'Private Const MQRC_DEST_ENV_ERROR = 2263
    'Private Const MQRC_DEST_NAME_ERROR = 2264
    'Private Const MQRC_TM_ERROR = 2265
    'Private Const MQRC_CLUSTER_EXIT_ERROR = 2266
    'Private Const MQRC_CLUSTER_EXIT_LOAD_ERROR = 2267
    'Private Const MQRC_CLUSTER_PUT_INHIBITED = 2268
    'Private Const MQRC_CLUSTER_RESOURCE_ERROR = 2269
    'Private Const MQRC_NO_DESTINATIONS_AVAILABLE = 2270
    'Private Const MQRC_CONN_TAG_IN_USE = 2271
    'Private Const MQRC_PARTIALLY_CONVERTED = 2272
    'Private Const MQRC_CONNECTION_ERROR = 2273
    'Private Const MQRC_OPTION_ENVIRONMENT_ERROR = 2274
    'Private Const MQRC_CD_ERROR = 2277
    'Private Const MQRC_CLIENT_CONN_ERROR = 2278
    'Private Const MQRC_CHANNEL_STOPPED_BY_USER = 2279
    'Private Const MQRC_HCONFIG_ERROR = 2280
    'Private Const MQRC_FUNCTION_ERROR = 2281
    'Private Const MQRC_CHANNEL_STARTED = 2282
    'Private Const MQRC_CHANNEL_STOPPED = 2283
    'Private Const MQRC_CHANNEL_CONV_ERROR = 2284
    'Private Const MQRC_SERVICE_NOT_AVAILABLE = 2285
    'Private Const MQRC_INITIALIZATION_FAILED = 2286
    'Private Const MQRC_TERMINATION_FAILED = 2287
    'Private Const MQRC_UNKNOWN_Q_NAME = 2288
    'Private Const MQRC_SERVICE_ERROR = 2289
    'Private Const MQRC_Q_ALREADY_EXISTS = 2290
    'Private Const MQRC_USER_ID_NOT_AVAILABLE = 2291
    'Private Const MQRC_UNKNOWN_ENTITY = 2292
    'Private Const MQRC_UNKNOWN_AUTH_ENTITY = 2293
    'Private Const MQRC_UNKNOWN_REF_OBJECT = 2294
    'Private Const MQRC_CHANNEL_ACTIVATED = 2295
    'Private Const MQRC_CHANNEL_NOT_ACTIVATED = 2296
    'Private Const MQRC_UOW_CANCELED = 2297
    'Private Const MQRC_FUNCTION_NOT_SUPPORTED = 2298
    'Private Const MQRC_SELECTOR_TYPE_ERROR = 2299
    'Private Const MQRC_COMMAND_TYPE_ERROR = 2300
    'Private Const MQRC_MULTIPLE_INSTANCE_ERROR = 2301
    'Private Const MQRC_SYSTEM_ITEM_NOT_ALTERABLE = 2302
    'Private Const MQRC_BAG_CONVERSION_ERROR = 2303
    'Private Const MQRC_SELECTOR_OUT_OF_RANGE = 2304
    'Private Const MQRC_SELECTOR_NOT_UNIQUE = 2305
    'Private Const MQRC_INDEX_NOT_PRESENT = 2306
    'Private Const MQRC_STRING_ERROR = 2307
    'Private Const MQRC_ENCODING_NOT_SUPPORTED = 2308
    'Private Const MQRC_SELECTOR_NOT_PRESENT = 2309
    'Private Const MQRC_OUT_SELECTOR_ERROR = 2310
    'Private Const MQRC_STRING_TRUNCATED = 2311
    'Private Const MQRC_SELECTOR_WRONG_TYPE = 2312
    'Private Const MQRC_INCONSISTENT_ITEM_TYPE = 2313
    'Private Const MQRC_INDEX_ERROR = 2314
    'Private Const MQRC_SYSTEM_BAG_NOT_ALTERABLE = 2315
    'Private Const MQRC_ITEM_COUNT_ERROR = 2316
    'Private Const MQRC_FORMAT_NOT_SUPPORTED = 2317
    'Private Const MQRC_SELECTOR_NOT_SUPPORTED = 2318
    'Private Const MQRC_ITEM_VALUE_ERROR = 2319
    'Private Const MQRC_HBAG_ERROR = 2320
    'Private Const MQRC_PARAMETER_MISSING = 2321
    'Private Const MQRC_CMD_SERVER_NOT_AVAILABLE = 2322
    'Private Const MQRC_STRING_LENGTH_ERROR = 2323
    'Private Const MQRC_INQUIRY_COMMAND_ERROR = 2324
    'Private Const MQRC_NESTED_BAG_NOT_SUPPORTED = 2325
    'Private Const MQRC_BAG_WRONG_TYPE = 2326
    'Private Const MQRC_ITEM_TYPE_ERROR = 2327
    'Private Const MQRC_SYSTEM_BAG_NOT_DELETABLE = 2328
    'Private Const MQRC_SYSTEM_ITEM_NOT_DELETABLE = 2329
    'Private Const MQRC_CODED_CHAR_SET_ID_ERROR = 2330
    'Private Const MQRC_MSG_TOKEN_ERROR = 2331
    'Private Const MQRC_MISSING_WIH = 2332
    'Private Const MQRC_WIH_ERROR = 2333
    'Private Const MQRC_RFH_ERROR = 2334
    'Private Const MQRC_RFH_STRING_ERROR = 2335
    'Private Const MQRC_RFH_COMMAND_ERROR = 2336
    'Private Const MQRC_RFH_PARM_ERROR = 2337
    'Private Const MQRC_RFH_DUPLICATE_PARM = 2338
    'Private Const MQRC_RFH_PARM_MISSING = 2339
    'Private Const MQRC_CHAR_CONVERSION_ERROR = 2340
    'Private Const MQRC_UCS2_CONVERSION_ERROR = 2341
    'Private Const MQRC_DB2_NOT_AVAILABLE = 2342
    'Private Const MQRC_OBJECT_NOT_UNIQUE = 2343
    'Private Const MQRC_CONN_TAG_NOT_RELEASED = 2344
    'Private Const MQRC_CF_NOT_AVAILABLE = 2345
    'Private Const MQRC_CF_STRUC_IN_USE = 2346
    'Private Const MQRC_CF_STRUC_LIST_HDR_IN_USE = 2347
    'Private Const MQRC_CF_STRUC_AUTH_FAILED = 2348
    'Private Const MQRC_CF_STRUC_ERROR = 2349
    'Private Const MQRC_CONN_TAG_NOT_USABLE = 2350
    'Private Const MQRC_GLOBAL_UOW_CONFLICT = 2351
    'Private Const MQRC_LOCAL_UOW_CONFLICT = 2352
    'Private Const MQRC_HANDLE_IN_USE_FOR_UOW = 2353
    'Private Const MQRC_UOW_ENLISTMENT_ERROR = 2354
    'Private Const MQRC_UOW_MIX_NOT_SUPPORTED = 2355
    'Private Const MQRC_WXP_ERROR = 2356
    'Private Const MQRC_CURRENT_RECORD_ERROR = 2357
    'Private Const MQRC_NEXT_OFFSET_ERROR = 2358
    'Private Const MQRC_NO_RECORD_AVAILABLE = 2359
    'Private Const MQRC_OBJECT_LEVEL_INCOMPATIBLE = 2360
    'Private Const MQRC_NEXT_RECORD_ERROR = 2361
    'Private Const MQRC_APPL_FIRST = 900
    'Private Const MQRC_APPL_LAST = 999
    'Private Const MQRCCF_CFH_TYPE_ERROR = 3001
    'Private Const MQRCCF_CFH_LENGTH_ERROR = 3002
    'Private Const MQRCCF_CFH_VERSION_ERROR = 3003
    'Private Const MQRCCF_CFH_MSG_SEQ_NUMBER_ERR = 3004
    'Private Const MQRCCF_CFH_CONTROL_ERROR = 3005
    'Private Const MQRCCF_CFH_PARM_COUNT_ERROR = 3006
    'Private Const MQRCCF_CFH_COMMAND_ERROR = 3007
    'Private Const MQRCCF_COMMAND_FAILED = 3008
    'Private Const MQRCCF_CFIN_LENGTH_ERROR = 3009
    'Private Const MQRCCF_CFST_LENGTH_ERROR = 3010
    'Private Const MQRCCF_CFST_STRING_LENGTH_ERR = 3011
    'Private Const MQRCCF_FORCE_VALUE_ERROR = 3012
    'Private Const MQRCCF_STRUCTURE_TYPE_ERROR = 3013
    'Private Const MQRCCF_CFIN_PARM_ID_ERROR = 3014
    'Private Const MQRCCF_CFST_PARM_ID_ERROR = 3015
    'Private Const MQRCCF_MSG_LENGTH_ERROR = 3016
    'Private Const MQRCCF_CFIN_DUPLICATE_PARM = 3017
    'Private Const MQRCCF_CFST_DUPLICATE_PARM = 3018
    'Private Const MQRCCF_PARM_COUNT_TOO_SMALL = 3019
    'Private Const MQRCCF_PARM_COUNT_TOO_BIG = 3020
    'Private Const MQRCCF_Q_ALREADY_IN_CELL = 3021
    'Private Const MQRCCF_Q_TYPE_ERROR = 3022
    'Private Const MQRCCF_MD_FORMAT_ERROR = 3023
    'Private Const MQRCCF_CFSL_LENGTH_ERROR = 3024
    'Private Const MQRCCF_REPLACE_VALUE_ERROR = 3025
    'Private Const MQRCCF_CFIL_DUPLICATE_VALUE = 3026
    'Private Const MQRCCF_CFIL_COUNT_ERROR = 3027
    'Private Const MQRCCF_CFIL_LENGTH_ERROR = 3028
    'Private Const MQRCCF_QUIESCE_VALUE_ERROR = 3029
    'Private Const MQRCCF_MSG_SEQ_NUMBER_ERROR = 3030
    'Private Const MQRCCF_PING_DATA_COUNT_ERROR = 3031
    'Private Const MQRCCF_PING_DATA_COMPARE_ERROR = 3032
    'Private Const MQRCCF_CFSL_PARM_ID_ERROR = 3033
    'Private Const MQRCCF_CHANNEL_TYPE_ERROR = 3034
    'Private Const MQRCCF_PARM_SEQUENCE_ERROR = 3035
    'Private Const MQRCCF_XMIT_PROTOCOL_TYPE_ERR = 3036
    'Private Const MQRCCF_BATCH_SIZE_ERROR = 3037
    'Private Const MQRCCF_DISC_INT_ERROR = 3038
    'Private Const MQRCCF_SHORT_RETRY_ERROR = 3039
    'Private Const MQRCCF_SHORT_TIMER_ERROR = 3040
    'Private Const MQRCCF_LONG_RETRY_ERROR = 3041
    'Private Const MQRCCF_LONG_TIMER_ERROR = 3042
    'Private Const MQRCCF_SEQ_NUMBER_WRAP_ERROR = 3043
    'Private Const MQRCCF_MAX_MSG_LENGTH_ERROR = 3044
    'Private Const MQRCCF_PUT_AUTH_ERROR = 3045
    'Private Const MQRCCF_PURGE_VALUE_ERROR = 3046
    'Private Const MQRCCF_CFIL_PARM_ID_ERROR = 3047
    'Private Const MQRCCF_MSG_TRUNCATED = 3048
    'Private Const MQRCCF_CCSID_ERROR = 3049
    'Private Const MQRCCF_ENCODING_ERROR = 3050
    'Private Const MQRCCF_DATA_CONV_VALUE_ERROR = 3052
    'Private Const MQRCCF_INDOUBT_VALUE_ERROR = 3053
    'Private Const MQRCCF_ESCAPE_TYPE_ERROR = 3054
    'Private Const MQRCCF_CHANNEL_TABLE_ERROR = 3062
    'Private Const MQRCCF_MCA_TYPE_ERROR = 3063
    'Private Const MQRCCF_CHL_INST_TYPE_ERROR = 3064
    'Private Const MQRCCF_CHL_STATUS_NOT_FOUND = 3065
    'Private Const MQRCCF_CFSL_DUPLICATE_PARM = 3066
    'Private Const MQRCCF_CFSL_TOTAL_LENGTH_ERROR = 3067
    'Private Const MQRCCF_CFSL_COUNT_ERROR = 3068
    'Private Const MQRCCF_CFSL_STRING_LENGTH_ERR = 3069
    'Private Const MQRCCF_BROKER_DELETED = 3070
    'Private Const MQRCCF_STREAM_ERROR = 3071
    'Private Const MQRCCF_TOPIC_ERROR = 3072
    'Private Const MQRCCF_NOT_REGISTERED = 3073
    'Private Const MQRCCF_Q_MGR_NAME_ERROR = 3074
    'Private Const MQRCCF_INCORRECT_STREAM = 3075
    'Private Const MQRCCF_Q_NAME_ERROR = 3076
    'Private Const MQRCCF_NO_RETAINED_MSG = 3077
    'Private Const MQRCCF_DUPLICATE_IDENTITY = 3078
    'Private Const MQRCCF_INCORRECT_Q = 3079
    'Private Const MQRCCF_CORREL_ID_ERROR = 3080
    'Private Const MQRCCF_NOT_AUTHORIZED = 3081
    'Private Const MQRCCF_UNKNOWN_STREAM = 3082
    'Private Const MQRCCF_REG_OPTIONS_ERROR = 3083
    'Private Const MQRCCF_PUB_OPTIONS_ERROR = 3084
    'Private Const MQRCCF_UNKNOWN_BROKER = 3085
    'Private Const MQRCCF_Q_MGR_CCSID_ERROR = 3086
    'Private Const MQRCCF_DEL_OPTIONS_ERROR = 3087
    'Private Const MQRCCF_CLUSTER_NAME_CONFLICT = 3088
    'Private Const MQRCCF_REPOS_NAME_CONFLICT = 3089
    'Private Const MQRCCF_CLUSTER_Q_USAGE_ERROR = 3090
    'Private Const MQRCCF_ACTION_VALUE_ERROR = 3091
    'Private Const MQRCCF_COMMS_LIBRARY_ERROR = 3092
    'Private Const MQRCCF_NETBIOS_NAME_ERROR = 3093
    'Private Const MQRCCF_BROKER_COMMAND_FAILED = 3094
    'Private Const MQRCCF_OBJECT_ALREADY_EXISTS = 4001
    'Private Const MQRCCF_OBJECT_WRONG_TYPE = 4002
    'Private Const MQRCCF_LIKE_OBJECT_WRONG_TYPE = 4003
    'Private Const MQRCCF_OBJECT_OPEN = 4004
    'Private Const MQRCCF_ATTR_VALUE_ERROR = 4005
    'Private Const MQRCCF_UNKNOWN_Q_MGR = 4006
    'Private Const MQRCCF_Q_WRONG_TYPE = 4007
    'Private Const MQRCCF_OBJECT_NAME_ERROR = 4008
    'Private Const MQRCCF_ALLOCATE_FAILED = 4009
    'Private Const MQRCCF_HOST_NOT_AVAILABLE = 4010
    'Private Const MQRCCF_CONFIGURATION_ERROR = 4011
    'Private Const MQRCCF_CONNECTION_REFUSED = 4012
    'Private Const MQRCCF_ENTRY_ERROR = 4013
    'Private Const MQRCCF_SEND_FAILED = 4014
    'Private Const MQRCCF_RECEIVED_DATA_ERROR = 4015
    'Private Const MQRCCF_RECEIVE_FAILED = 4016
    'Private Const MQRCCF_CONNECTION_CLOSED = 4017
    'Private Const MQRCCF_NO_STORAGE = 4018
    'Private Const MQRCCF_NO_COMMS_MANAGER = 4019
    'Private Const MQRCCF_LISTENER_NOT_STARTED = 4020
    'Private Const MQRCCF_BIND_FAILED = 4024
    'Private Const MQRCCF_CHANNEL_INDOUBT = 4025
    'Private Const MQRCCF_MQCONN_FAILED = 4026
    'Private Const MQRCCF_MQOPEN_FAILED = 4027
    'Private Const MQRCCF_MQGET_FAILED = 4028
    'Private Const MQRCCF_MQPUT_FAILED = 4029
    'Private Const MQRCCF_PING_ERROR = 4030
    'Private Const MQRCCF_CHANNEL_IN_USE = 4031
    'Private Const MQRCCF_CHANNEL_NOT_FOUND = 4032
    'Private Const MQRCCF_UNKNOWN_REMOTE_CHANNEL = 4033
    'Private Const MQRCCF_REMOTE_QM_UNAVAILABLE = 4034
    'Private Const MQRCCF_REMOTE_QM_TERMINATING = 4035
    'Private Const MQRCCF_MQINQ_FAILED = 4036
    'Private Const MQRCCF_NOT_XMIT_Q = 4037
    'Private Const MQRCCF_CHANNEL_DISABLED = 4038
    'Private Const MQRCCF_USER_EXIT_NOT_AVAILABLE = 4039
    'Private Const MQRCCF_COMMIT_FAILED = 4040
    'Private Const MQRCCF_CHANNEL_ALREADY_EXISTS = 4042
    'Private Const MQRCCF_DATA_TOO_LARGE = 4043
    'Private Const MQRCCF_CHANNEL_NAME_ERROR = 4044
    'Private Const MQRCCF_XMIT_Q_NAME_ERROR = 4045
    'Private Const MQRCCF_MCA_NAME_ERROR = 4047
    'Private Const MQRCCF_SEND_EXIT_NAME_ERROR = 4048
    'Private Const MQRCCF_SEC_EXIT_NAME_ERROR = 4049
    'Private Const MQRCCF_MSG_EXIT_NAME_ERROR = 4050
    'Private Const MQRCCF_RCV_EXIT_NAME_ERROR = 4051
    'Private Const MQRCCF_XMIT_Q_NAME_WRONG_TYPE = 4052
    'Private Const MQRCCF_MCA_NAME_WRONG_TYPE = 4053
    'Private Const MQRCCF_DISC_INT_WRONG_TYPE = 4054
    'Private Const MQRCCF_SHORT_RETRY_WRONG_TYPE = 4055
    'Private Const MQRCCF_SHORT_TIMER_WRONG_TYPE = 4056
    'Private Const MQRCCF_LONG_RETRY_WRONG_TYPE = 4057
    'Private Const MQRCCF_LONG_TIMER_WRONG_TYPE = 4058
    'Private Const MQRCCF_PUT_AUTH_WRONG_TYPE = 4059
    'Private Const MQRCCF_MISSING_CONN_NAME = 4061
    'Private Const MQRCCF_CONN_NAME_ERROR = 4062
    'Private Const MQRCCF_MQSET_FAILED = 4063
    'Private Const MQRCCF_CHANNEL_NOT_ACTIVE = 4064
    'Private Const MQRCCF_TERMINATED_BY_SEC_EXIT = 4065
    'Private Const MQRCCF_DYNAMIC_Q_SCOPE_ERROR = 4067
    'Private Const MQRCCF_CELL_DIR_NOT_AVAILABLE = 4068
    'Private Const MQRCCF_MR_COUNT_ERROR = 4069
    'Private Const MQRCCF_MR_COUNT_WRONG_TYPE = 4070
    'Private Const MQRCCF_MR_EXIT_NAME_ERROR = 4071
    'Private Const MQRCCF_MR_EXIT_NAME_WRONG_TYPE = 4072
    'Private Const MQRCCF_MR_INTERVAL_ERROR = 4073
    'Private Const MQRCCF_MR_INTERVAL_WRONG_TYPE = 4074
    'Private Const MQRCCF_NPM_SPEED_ERROR = 4075
    'Private Const MQRCCF_NPM_SPEED_WRONG_TYPE = 4076
    'Private Const MQRCCF_HB_INTERVAL_ERROR = 4077
    'Private Const MQRCCF_HB_INTERVAL_WRONG_TYPE = 4078
    'Private Const MQRCCF_CHAD_ERROR = 4079
    'Private Const MQRCCF_CHAD_WRONG_TYPE = 4080
    'Private Const MQRCCF_CHAD_EVENT_ERROR = 4081
    'Private Const MQRCCF_CHAD_EVENT_WRONG_TYPE = 4082
    'Private Const MQRCCF_CHAD_EXIT_ERROR = 4083
    'Private Const MQRCCF_CHAD_EXIT_WRONG_TYPE = 4084
    'Private Const MQRCCF_SUPPRESSED_BY_EXIT = 4085
    'Private Const MQRCCF_BATCH_INT_ERROR = 4086
    'Private Const MQRCCF_BATCH_INT_WRONG_TYPE = 4087
    'Private Const MQRCCF_NET_PRIORITY_ERROR = 4088
    'Private Const MQRCCF_NET_PRIORITY_WRONG_TYPE = 4089
    'Private Const MQRCCF_CHANNEL_CLOSED = 4090
    'Private Const MQRL_UNDEFINED = &HFFFFFFFF
    'Private Const MQRMH_CURRENT_VERSION = 1
    'Private Const MQRMH_VERSION_1 = 1
    'Private Const MQRMHF_LAST = 1
    'Private Const MQRMHF_NOT_LAST = 0
    'Private Const MQRO_EXCEPTION = &H1000000
    'Private Const MQRO_EXCEPTION_WITH_DATA = &H3000000
    'Private Const MQRO_EXCEPTION_WITH_FULL_DATA = &H7000000
    'Private Const MQRO_EXPIRATION = &H200000
    'Private Const MQRO_EXPIRATION_WITH_DATA = &H600000
    'Private Const MQRO_EXPIRATION_WITH_FULL_DATA = &HE00000
    'Private Const MQRO_COA = 256
    'Private Const MQRO_COA_WITH_DATA = 768
    'Private Const MQRO_COA_WITH_FULL_DATA = 1792
    'Private Const MQRO_COD = 2048
    'Private Const MQRO_COD_WITH_DATA = 6144
    'Private Const MQRO_COD_WITH_FULL_DATA = 14336
    'Private Const MQRO_PAN = 1
    'Private Const MQRO_NAN = 2
    'Private Const MQRO_NEW_MSG_ID = 0
    'Private Const MQRO_PASS_MSG_ID = 128
    'Private Const MQRO_COPY_MSG_ID_TO_CORREL_ID = 0
    'Private Const MQRO_PASS_CORREL_ID = 64
    'Private Const MQRO_DEAD_LETTER_Q = 0
    'Private Const MQRO_DISCARD_MSG = &H8000000
    'Private Const MQRO_NONE = 0
    'Private Const MQRO_REJECT_UNSUP_MASK = &H101C0000
    'Private Const MQRO_ACCEPT_UNSUP_MASK = &HEFE000FF
    'Private Const MQRO_ACCEPT_UNSUP_IF_XMIT_MASK = &H3FF00
    'Private Const MQRP_NO = 0
    'Private Const MQRP_YES = 1
    'Private Const MQRQ_CONN_NOT_AUTHORIZED = 1
    'Private Const MQRQ_OPEN_NOT_AUTHORIZED = 2
    'Private Const MQRQ_CLOSE_NOT_AUTHORIZED = 3
    'Private Const MQRQ_CMD_NOT_AUTHORIZED = 4
    'Private Const MQRQ_Q_MGR_STOPPING = 5
    'Private Const MQRQ_Q_MGR_QUIESCING = 6
    'Private Const MQRQ_CHANNEL_STOPPED_OK = 7
    'Private Const MQRQ_CHANNEL_STOPPED_ERROR = 8
    'Private Const MQRQ_CHANNEL_STOPPED_RETRY = 9
    'Private Const MQRQ_CHANNEL_STOPPED_DISABLED = 10
    'Private Const MQRQ_BRIDGE_STOPPED_OK = 11
    'Private Const MQRQ_BRIDGE_STOPPED_ERROR = 12
    'Private Const MQSP_AVAILABLE = 1
    'Private Const MQSP_NOT_AVAILABLE = 0
    'Private Const MQTC_OFF = 0
    'Private Const MQTC_ON = 1
    'Private Const MQTM_CURRENT_VERSION = 1
    'Private Const MQTM_VERSION_1 = 1
    'Private Const MQTT_NONE = 0
    'Private Const MQTT_EVERY = 2
    'Private Const MQTT_FIRST = 1
    'Private Const MQTT_DEPTH = 3
    'Private Const MQUS_NORMAL = 0
    'Private Const MQUS_TRANSMISSION = 1
    'Private Const MQWI_UNLIMITED = &HFFFFFFFF
    'Private Const MQXQH_CURRENT_VERSION = 1
    'Private Const MQXQH_VERSION_1 = 1
    'Private Const MQEVR_DISABLED = 0
    'Private Const MQEVR_ENABLED = 1
    'Private Const MQQSIE_HIGH = 1
    'Private Const MQQSIE_NONE = 0
    'Private Const MQQSIE_OK = 2
    'Private Const MQSCO_CELL = 2
    'Private Const MQSCO_Q_MGR = 1
    'Private Const MQRC_LIBRARY_LOAD_ERROR = 6000
    'Private Const MQRC_CLASS_LIBRARY_ERROR = 6001
    'Private Const MQRC_STRING_LENGTH_TOO_BIG = 6002
    'Private Const MQRC_WRITE_VALUE_ERROR = 6003
    'Private Const MQRC_PACKED_DECIMAL_ERROR = 6004
    'Private Const MQRC_REOPEN_EXCL_INPUT_ERROR = 6100
    'Private Const MQRC_REOPEN_INQUIRE_ERROR = 6101
    'Private Const MQRC_REOPEN_SAVED_CONTEXT_ERR = 6102
    'Private Const MQRC_REOPEN_TEMPORARY_Q_ERROR = 6103
    'Private Const MQRC_ATTRIBUTE_LOCKED = 6104
    'Private Const MQRC_CURSOR_NOT_VALID = 6105
    'Private Const MQRC_ENCODING_ERROR = 6106
    'Private Const MQRC_STRUC_ID_ERROR = 6107
    'Private Const MQRC_NULL_POINTER = 6108
    'Private Const MQRC_NO_CONNECTION_REFERENCE = 6109
    'Private Const MQRC_NO_BUFFER = 6110
    'Private Const MQRC_BINARY_DATA_LENGTH_ERROR = 6111
    'Private Const MQRC_BUFFER_NOT_AUTOMATIC = 6112
    'Private Const MQRC_INSUFFICIENT_BUFFER = 6113
    'Private Const MQRC_INSUFFICIENT_DATA = 6114
    'Private Const MQRC_DATA_TRUNCATED = 6115
    'Private Const MQRC_ZERO_LENGTH = 6116
    'Private Const MQRC_NEGATIVE_LENGTH = 6117
    'Private Const MQRC_NEGATIVE_OFFSET = 6118
    'Private Const MQRC_INCONSISTENT_FORMAT = 6119
    'Private Const MQRC_INCONSISTENT_OBJECT_STATE = 6120
    'Private Const MQRC_CONTEXT_OBJECT_NOT_VALID = 6121
    'Private Const MQRC_CONTEXT_OPEN_ERROR = 6122
    'Private Const MQRC_STRUC_LENGTH_ERROR = 6123
    'Private Const MQRC_NOT_CONNECTED = 6124
    Public Const MQRC_NOT_OPEN As Integer = 6125
    'Private Const MQSEL_ANY_SELECTOR = &HFFFF8ACF
    'Private Const MQSEL_ANY_USER_SELECTOR = &HFFFF8ACE
    'Private Const MQSEL_ANY_SYSTEM_SELECTOR = &HFFFF8ACD
    'Private Const MQSEL_ALL_SELECTORS = &HFFFF8ACF
    'Private Const MQSEL_ALL_USER_SELECTORS = &HFFFF8ACE
    'Private Const MQSEL_ALL_SYSTEM_SELECTORS = &HFFFF8ACD
    'Private Const MQSUS_NO = 0
    'Private Const MQSUS_YES = 1
    'Private Const MQWIH_VERSION_1 = 1
    'Private Const MQWIH_CURRENT_VERSION = 1
    'Private Const MQWIH_CURRENT_LENGTH = 120
    'Private Const MQWIH_LENGTH_1 = 120
    'Private Const MQWIH_NONE = 0
    'Private Const MQRFH_VERSION_1 = 1
    'Private Const MQRFH_VERSION_2 = 2
    'Private Const MQRFH_STRUC_LENGTH_FIXED = 32
    'Private Const MQRFH_STRUC_LENGTH_FIXED_2 = 36
    'Private Const MQRFH_NONE = 0

    Private Seconds As enumWaitTime

    Private Const CUSTOMERFOUND As Integer = 1
    Private Const MONTHCOUNTDEPOSITS As Integer = 12
    Private Const MONTHCOUNTWITHDRAWALS As Integer = 12
    Private Const MT_BANK_TEST As Integer = 65536
    Private Const MT_BANK_REQUEST As Integer = 65537
    Private Const MT_BANK_REPLY As Integer = 65538
    Private Const BANK_TEST_RESPONSE As String = "This is a bank test response."

    'Formats'
    Private Const MQFMT_NONE As String = "        "
    Private Const MQFMT_ADMIN As String = "MQADMIN "
    Private Const MQFMT_CHANNEL_COMPLETED As String = "MQCHCOM "
    Private Const MQFMT_CICS As String = "MQCICS  "
    Private Const MQFMT_COMMAND_1 As String = "MQCMD1  "
    Private Const MQFMT_COMMAND_2 As String = "MQCMD2  "
    Private Const MQFMT_DEAD_LETTER_HEADER As String = "MQDEAD  "
    Private Const MQFMT_DIST_HEADER As String = "MQHDIST "
    Private Const MQFMT_EVENT As String = "MQEVENT "
    Private Const MQFMT_IMS As String = "MQIMS   "
    Private Const MQFMT_IMS_VAR_STRING As String = "MQIMSVS "
    Private Const MQFMT_MD_EXTENSION As String = "MQHMDE  "
    Private Const MQFMT_PCF As String = "MQPCF   "
    Private Const MQFMT_REF_MSG_HEADER As String = "MQHREF  "
    Private Const MQFMT_RF_HEADER As String = "MQHRF   "
    Private Const MQFMT_RF_HEADER_2 As String = "MQHRF2  "
    Private Const MQFMT_STRING As String = "MQSTR   "
    Private Const MQFMT_TRIGGER As String = "MQTRIG  "
    Private Const MQFMT_WORK_INFO_HEADER As String = "MQHWIH  "
    Private Const MQFMT_XMIT_Q_HEADER As String = "MQXMIT  "

    'Encoding'
    Private Const DefEnc_OS2 As Integer = 546
    Private Const DefEnc_DOS As Integer = 546
    Private Const DefEnc_Windows As Integer = 546
    Private Const DefEnc_MicroFocusCOBOL As Integer = 17
    Private Const DefEnc_OpenVMS As Integer = 273
    Private Const DefEnc_MVSESA As Integer = 785
    Private Const DefEnc_OS400 As Integer = 273
    Private Const DefEnc_Tandem As Integer = 273
    Private Const DefEnc_UNIX As Integer = 273

    'Private m_MQSess As MQAX200.MQSession                 '* session object
    'Private m_QMgr As MQAX200.MQQueueManager              '* queue manager object
    'Private m_InputQueue As MQAX200.MQQueue               '* input queue object
    'Private m_OutputQueue As MQAX200.MQQueue              '* output queue object
    'Private m_PutMsg As MQAX200.MQMessage                 '* message object for put
    'Private m_GetMsg As MQAX200.MQMessage                 '* message object for get
    'Private m_PutOptions As MQAX200.MQPutMessageOptions   '* get message options
    'Private m_GetOptions As MQAX200.MQGetMessageOptions   '* put message options

    Private m_MQSess As Object '* session object
    Private m_QMgr As Object '* queue manager object
    Private m_InputQueue As Object          '* input queue object
    Private m_OutputQueue As Object            '* output queue object
    Private m_PutMsg As Object           '* message object for put
    Private m_GetMsg As Object                 '* message object for get
    Private m_PutOptions As Object   '* get message options
    Private m_GetOptions As Object   '* put message options

    Private m_WksID As String = "SAIBW000"
    Private m_BrnCde As String = "0101"
    Private m_MsgSrc As String
    Private m_MsgDst As String
    Private m_EnvSrc As String
    Private m_EnvDst As String
    Private m_SentMsgID As String
    Private m_SendRetries As Integer = 3
    Private m_GetRetries As Integer = 3
    Private m_ReSendDelay As Integer = 3000
    Private m_DoWait As Boolean = True
    Private m_TmeOut As Integer = 2000
    Private m_LoggedUser As String
    Private m_Connected As Boolean
    Private m_QueueManager As String
    Private m_ReadQueue As String
    Private m_WriteQueue As String
    Private m_ReplyToQueue As String = ""

    Public Structure SibHdrStruct
        'SubStruct:MsgHdr	Size:256
        Dim MsgProCde As String
        Dim MsgMid As String
        Dim MsgStp As String
        'SubStruct:MsgOrg	Size:10
        Dim MsgSrc As String
        Dim MsgSrcEnv As String
        Dim MsgSrcRsv As String
        Dim MsgUsrIde As String
        'SubStruct:MsgDst	Size:9
        Dim MsgTgt As String
        Dim MsgTgtEnv As String
        Dim MsgTgtRsv As String
        'SubStruct:MsgResCde	Size:10
        Dim MsgResVal As String
        Dim MsgActCde As String
        Dim MsgSysCon As String
        Dim Rsv1b As String
        Dim MsgResInd As String
        Dim SibResCde As String
        Dim SifRefNum As String
        'SubStruct:SibEnvLoc	Size:90
        Dim SibEnvUnt As String    '3
        Dim SibOrgCls As String     '3
        'Dim SibKeyLoc As String    '84
        Dim SibKeyCls As String '3x
        Dim SibCus As String '6n
        Dim SibAcc As String '13n
        Dim SibVar As String '20x
        '---Dim SibDea As String '20x
        '---Dim SibRetIde As String '12x
        '---Dim SibTrmIde As String '4x
        '---Dim SibCrdNum As String '16x
        Dim MsiDat As String '10x
        Dim Rsv21 As String  '10x
        Dim CstVerTab As String  '12n
        Dim SibCusSeg As String '1x
        Dim Rsv2 As String  '9x
        '**************************************************************************************************
        'SubStruct:SibNonRep	Size:50
        Dim IceBrnNum As String '4x
        Dim LstNonDupMid As String '12x
        Dim IceTryFin As String '1x
        Dim Rsv415 As String '2x
        Dim AskMidRef As String '12x	MID of the corresponding ASK message
        Dim IceCapCde As String '2x	    ICE system feature capability code
        Dim AutUsrIde As String '10x
        Dim AutPacVal As String '6x
        Dim SibNrdTyp As String '1x
        Dim Rsv3 As String   '4x

        Public Function PackHeader(ByVal BaseClass As MQSeries) As String
            Dim RetStr As String
            Try
                IceCapCde = "02"
                With BaseClass
                    RetStr = .PackString(MsgProCde, 6)
                    RetStr &= .PackString(MsgMid, 12)
                    RetStr &= .PackString(MsgStp, 14)
                    RetStr &= .PackString(MsgSrc, 4)
                    RetStr &= .PackString(MsgSrcEnv, 3)
                    RetStr &= .PackString(MsgSrcRsv, 3)
                    RetStr &= .PackString(MsgUsrIde, 32)
                    RetStr &= .PackString(MsgTgt, 4)
                    RetStr &= .PackString(MsgTgtEnv, 3)
                    RetStr &= .PackString(MsgTgtRsv, 2)
                    RetStr &= .PackString(MsgResVal, 3)
                    RetStr &= .PackString(MsgActCde, 2)
                    RetStr &= .PackString(MsgSysCon, 3)
                    RetStr &= .PackString(Rsv1b, 1)
                    RetStr &= .PackString(MsgResInd, 1)
                    RetStr &= .PackString(SibResCde, 7)
                    RetStr &= .PackString(SifRefNum, 12)
                    RetStr &= .PackString(SibEnvUnt, 3)
                    RetStr &= .PackString(SibOrgCls, 3)
                    ' SibKeyLoc, 84
                    RetStr &= .PackString(SibKeyCls, 3)
                    RetStr &= .PackString(SibCus, 6)
                    RetStr &= .PackString(SibAcc, 13)
                    RetStr &= .PackString(SibVar, 20)
                    '---RetStr &= .PackString(SibDea, 20)
                    '---RetStr &= .PackString(SibRetIde, 12)
                    '---RetStr &= .PackString(SibTrmIde, 4)
                    '---RetStr &= .PackString(SibCrdNum, 16)
                    RetStr &= .PackString(MsiDat, 10)
                    RetStr &= .PackString(Rsv21, 10)
                    RetStr &= .PackString(CstVerTab, 12)
                    RetStr &= .PackString(SibCusSeg, 1)
                    RetStr &= .PackString(Rsv2, 9)
                    'SubStruct:SibNonRep	Size:50
                    'RetStr &= .PackString("0303", 4) 
                    RetStr &= .PackString(BaseClass.m_BrnCde, 4)
                    RetStr &= .PackString(BaseClass.m_SentMsgID, 12)
                    RetStr &= .PackString(IceTryFin, 1)
                    RetStr &= .PackString(Rsv415, 2)
                    RetStr &= .PackString(AskMidRef, 12)
                    RetStr &= .PackString(IceCapCde, 2)
                    RetStr &= .PackString(AutUsrIde, 10)
                    RetStr &= .PackString(AutPacVal, 6)
                    RetStr &= .PackString("I", 1)  'RetStr &= .PackString(.SibNrdTyp, 1)
                    'rsv3
                    RetStr &= .PackString(Rsv3, 4)
                End With
            Catch
                RetStr = Space(256)
            End Try
            Return RetStr
        End Function

        Public Sub ReadHeader(ByVal HdrStr As String)
            'SibHdr	    256-->	Bank Header (includes BankAway Header)

            'BwyMsgHdr	93 -->	BankAway Message header (structure follows)
            MsgProCde = strip(HdrStr, 6)
            MsgMid = strip(HdrStr, 12)
            MsgStp = strip(HdrStr, 14)

            'MsgOrg	    10x-->	Message sender
            MsgSrc = strip(HdrStr, 4)
            MsgSrcEnv = strip(HdrStr, 3)
            MsgSrcRsv = strip(HdrStr, 3)

            MsgUsrIde = strip(HdrStr, 32)

            'MsgDst	    9x-->	Message destination
            MsgTgt = strip(HdrStr, 4)
            MsgTgtEnv = strip(HdrStr, 3)
            MsgTgtRsv = strip(HdrStr, 2)

            'MsgResCde	10-->	Result code
            MsgResVal = strip(HdrStr, 3)
            MsgActCde = strip(HdrStr, 2)
            MsgSysCon = strip(HdrStr, 3)
            Rsv1b = strip(HdrStr, 1)
            MsgResInd = strip(HdrStr, 1)

            SibResCde = strip(HdrStr, 7)
            SifRefNum = strip(HdrStr, 12)

            'SibEnvLoc	90-->	Environment and Locus (Structure follows)
            SibEnvUnt = strip(HdrStr, 3)
            SibOrgCls = strip(HdrStr, 3)

            'SibKeyLoc	84-->	Key Locus values (structure follows)
            SibKeyCls = strip(HdrStr, 3)
            SibCus = strip(HdrStr, 6)
            SibAcc = strip(HdrStr, 13)
            SibVar = strip(HdrStr, 20)
            '---SibDea = strip(HdrStr, 20)
            '---SibRetIde = strip(HdrStr, 12)
            '---SibTrmIde = strip(HdrStr, 4)
            '---SibCrdNum = strip(HdrStr, 16)
            MsiDat = strip(HdrStr, 10)
            Rsv21 = strip(HdrStr, 10)
            CstVerTab = strip(HdrStr, 12)
            SibCusSeg = strip(HdrStr, 1)
            Rsv2 = strip(HdrStr, 9)

            'SibNonRep	50x-->	Non-repudiation data (refer to the structures)
            IceBrnNum = strip(HdrStr, 4)
            LstNonDupMid = strip(HdrStr, 12)
            IceTryFin = strip(HdrStr, 1)
            Rsv415 = strip(HdrStr, 2)
            AskMidRef = strip(HdrStr, 12)
            IceCapCde = strip(HdrStr, 2)
            AutUsrIde = strip(HdrStr, 10)
            AutPacVal = strip(HdrStr, 6)
            SibNrdTyp = strip(HdrStr, 1)

            Rsv3 = strip(HdrStr, 4)
        End Sub
    End Structure

    Private Structure SibHdrStruct83
        'SubStruct:MsgHdr	Size:83
        Dim MsgProCde As String '6x
        Dim MsgMid As String '12x
        Dim MsgStp As String '14x
        Dim MsgOrg As String '10x
        Dim MsgUsrIde As String '32x
        Dim MsgDst As String '9x
    End Structure

    Private Structure SibHdrStruct93
        'SubStruct:MsgHdr	Size:93
        Dim MsgProCde As String '6x
        Dim MsgMid As String '12x
        Dim MsgStp As String '14x
        Dim MsgOrg As String '10x
        Dim MsgUsrIde As String '32x
        Dim MsgDst As String '9x
        'MsgResCde	10
        Dim MsgResVal As String '3n
        Dim MsgActCde As String '2x
        Dim MsgSysCon As String '3x
        Dim Rsv1 As String  '2x
    End Structure

#End Region

    'Public Section

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Gets or sets the currently logged on user ID for MQ Series messages    ''' 
    ''' </summary>
    ''' <remarks>
    ''' This property will not return the current Windows User! It just returns whatever was set in it before     ''' 
    ''' </remarks>
    ''' <history>
    ''' 	[U910KRZI]	05/05/2004	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Property LoggedUser() As String
        Set(ByVal Value As String)
            m_LoggedUser = Value
        End Set
        Get
            LoggedUser = m_LoggedUser
        End Get
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Returns the connection status of the MQ Series class with the QueueManager    ''' 
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[U910KRZI]	05/05/2004	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public ReadOnly Property Connected() As Boolean
        Get
            Connected = m_Connected
        End Get
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Gets or sets the source system name for the SAIB message header    ''' 
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[U910KRZI]	05/05/2004	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Property SrcSysName() As String
        Get
            SrcSysName = m_MsgSrc
        End Get
        Set(ByVal Value As String)
            m_MsgSrc = Value
        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Gets or sets the Timeout period before a message retrieval is considered failed    ''' 
    ''' </summary>
    ''' <remarks>
    ''' If DoWait is set to False, this value has no effect    ''' 
    ''' </remarks>
    ''' <history>
    ''' 	[U910KRZI]	05/05/2004	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Property TmeOut() As Integer
        Get
            TmeOut = m_TmeOut
        End Get
        Set(ByVal Value As Integer)
            m_TmeOut = Value
            m_GetOptions.WaitInterval = m_TmeOut
        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Gets or sets The user's branch code    ''' 
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[U910KRZI]	27/07/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Property BrnCde() As String
        Get
            BrnCde = m_BrnCde
        End Get
        Set(ByVal Value As String)
            m_BrnCde = Value
        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Gets or sets whether message retrieval blocks the current thread with the specified TmeOut    ''' 
    ''' </summary>
    ''' <remarks>
    ''' if Set to False, GetMessage will returns immediatly regardless if a message was found or not.
    ''' The application must provide logic to handle the message retrieval retries if desired so.   ''' 
    ''' </remarks>
    ''' <history>
    ''' 	[U910KRZI]	05/05/2004	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Property DoWait() As Boolean
        Get
            DoWait = m_DoWait
        End Get
        Set(ByVal Value As Boolean)
            m_DoWait = Value
            m_GetOptions.Options = MQGMO_NO_SYNCPOINT
            If m_DoWait Then m_GetOptions.Options += MQGMO_WAIT
        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Gets or sets the the destination system name for the SAIB message header    ''' 
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[U910KRZI]	05/05/2004	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Property DstSysName() As String
        Get
            DstSysName = m_MsgDst
        End Get
        Set(ByVal Value As String)
            m_MsgDst = Value
        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Gets or sets the the source environment name for the SAIB message header    ''' 
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[U910KRZI]	05/05/2004	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Property SrcEnvName() As String
        Get
            SrcEnvName = m_EnvSrc
        End Get
        Set(ByVal Value As String)
            m_EnvSrc = Value
        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Gets or sets the message send retry count    ''' 
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[U910KRZI]	233/04/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Property SendRetries() As Integer
        Get
            SendRetries = m_SendRetries
        End Get
        Set(ByVal Value As Integer)
            m_SendRetries = Value
        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Gets or sets the message send retry count    ''' 
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[U910KRZI]	233/04/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Property GetRetries() As Integer
        Get
            GetRetries = m_GetRetries
        End Get
        Set(ByVal Value As Integer)
            m_GetRetries = Value
        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Gets or sets the delay between each message resend    ''' 
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[U910KRZI]	23/04/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Property ResendDelay() As Integer
        Get
            ResendDelay = m_ReSendDelay
        End Get
        Set(ByVal Value As Integer)
            m_ReSendDelay = Value
        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Gets or sets the destination environment name for the SAIB message header    ''' 
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[U910KRZI]	05/05/2004	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Property DstEnvName() As String
        Get
            DstEnvName = m_EnvDst
        End Get
        Set(ByVal Value As String)
            m_EnvDst = Value
        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Gets the last MessageID (MID) sent with the SAIB message header    ''' 
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[U910KRZI]	06/09/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public ReadOnly Property SentMsgID() As String
        Get
            Return m_SentMsgID
        End Get
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Gets the last MessageID (MID) sent with the SAIB message header    ''' 
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[U910KRZI]	06/09/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Property ReplyToQueue() As String
        Get
            Return m_ReplyToQueue
        End Get
        Set(ByVal Value As String)
            m_ReplyToQueue = Value
        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Create a new instance of the MQSeries class    ''' 
    ''' </summary>
    ''' <remarks>
    ''' Before accessing any methods, Connect must be called with the desired Queue name     ''' 
    ''' </remarks>
    ''' <history>
    ''' 	[U910KRZI]	05/05/2004	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Sub New()
        Try
            'm_MQSess = New MQAX200.MQSession
            m_MQSess = CreateObject("MQAX200.MQSession")
        Catch
            If (m_MQSess Is Nothing) Then
                m_ErrNum = &H80006666
                m_ErrDsc = "Could not create an MQ Series session"
            Else
                m_ErrNum = m_MQSess.ReasonCode
                m_ErrDsc = m_MQSess.ReasonName
            End If
        End Try
        m_WksID = Trim(Environment.GetEnvironmentVariable("COMPUTERNAME"))
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Create a new instance of the MQSeries class and connects to the specified queue manager name provided    ''' 
    ''' </summary>
    ''' <param name="QueueManagerName"></param>
    ''' <remarks>
    ''' If connection failed, MQSeries.ErrNum would be a non-zero value, otherwise, it is set to zero    ''' 
    ''' </remarks>
    ''' <history>
    ''' 	[U910KRZI]	05/05/2004	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Sub New(ByVal QueueManagerName As String)
        Try
            m_QueueManager = QueueManagerName
            'm_MQSess = New MQAX200.MQSession
            m_MQSess = CreateObject("MQAX200.MQSession")
            m_QMgr = m_MQSess.AccessQueueManager(QueueManagerName)
        Catch
            If (m_MQSess Is Nothing) Then
                m_ErrNum = &H80006666
                m_ErrDsc = "Could not create an MQ Series session"
            Else
                m_ErrNum = m_MQSess.ReasonCode
                m_ErrDsc = m_MQSess.ReasonName
            End If
        End Try
        m_WksID = Trim(Environment.GetEnvironmentVariable("COMPUTERNAME"))
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Connects to the specified QueueManagerName and open the specified Write and Read Queues    ''' 
    ''' </summary>
    ''' <param name="QueueManagerName"></param>
    ''' <param name="WriteQueueName"></param>
    ''' <param name="ReadQueueName"></param>
    ''' <returns></returns>
    ''' <remarks>
    ''' If connection failed, MQSeries.ErrNum would be a non-zero value, otherwise, it is set to zero    ''' 
    ''' </remarks>
    ''' <history>
    ''' 	[U910KRZI]	05/05/2004	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Function Connect(ByVal QueueManagerName As String, ByVal WriteQueueName As String, ByVal ReadQueueName As String) As Boolean
        Try
            m_QueueManager = QueueManagerName
            'm_MQSess = New MQAX200.MQSession
            m_QMgr = m_MQSess.AccessQueueManager(QueueManagerName)
            Return Connect(WriteQueueName, ReadQueueName)
        Catch
            If (m_MQSess Is Nothing) Then
                m_ErrNum = &H80006666
                m_ErrDsc = "Could not create an MQ Series session"
            Else
                m_ErrNum = m_MQSess.ReasonCode
                m_ErrDsc = m_MQSess.ReasonName
            End If
            m_Connected = False
            Return False
        End Try
    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Connects to Queue Manager name given to New(QueueName) and open the Read and Write queues provided here   ''' 
    ''' </summary>
    ''' <param name="WriteQueueName"></param>
    ''' <param name="ReadQueueName"></param>
    ''' <returns></returns>
    ''' <remarks>
    ''' You must use the full Connect version if you initialized the MQ Series with New()    ''' 
    ''' </remarks>
    ''' <history>
    ''' 	[U910KRZI]	05/05/2004	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Function Connect(ByVal WriteQueueName As String, ByVal ReadQueueName As String) As Boolean
        Dim RetValue As Boolean
        m_ReadQueue = ReadQueueName
        m_WriteQueue = WriteQueueName
        If (m_ReplyToQueue = "") Then m_ReplyToQueue = ReadQueueName
        Try
            m_InputQueue = m_QMgr.AccessQueue(ReadQueueName, MQOO_INPUT_AS_Q_DEF)
            m_GetOptions = m_MQSess.AccessGetMessageOptions()
            m_GetOptions.Options = MQGMO_NO_SYNCPOINT
            m_GetOptions.MatchOptions = MQMO_MATCH_MSG_ID + MQMO_MATCH_CORREL_ID
            If m_DoWait Then m_GetOptions.Options += MQGMO_WAIT
            m_GetOptions.WaitInterval = m_TmeOut
            m_OutputQueue = m_QMgr.AccessQueue(WriteQueueName, MQOO_OUTPUT)
            m_PutOptions = m_MQSess.AccessPutMessageOptions()
            m_PutOptions.Options = MQPMO_NO_SYNCPOINT
            'm_PutOptions.UserID = m_LoggedUser
            RetValue = True
            m_ErrNum = 0
            m_ErrDsc = ""
        Catch ex As Exception
            If (m_MQSess Is Nothing) Then
                m_ErrNum = &H80006666
                m_ErrDsc = "Could not create an MQ Series session"
            Else
                If m_MQSess.CompletionCode <> MQCC_OK Then
                    m_ErrNum = m_MQSess.ReasonCode
                    m_ErrDsc = m_MQSess.ReasonName
                End If
            End If
            RetValue = False
        End Try

        m_Connected = RetValue
        Return RetValue

    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' ReConnects to Queue Manager name given to New(QueueName) and Connect and open the Read and Write queues provided ''' 
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks>
    ''' You must use the full Connect version if you initialized the MQ Series with New()    ''' 
    ''' </remarks>
    ''' <history>
    ''' 	[U910KRZI]	23/04/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Function ReConnect() As Boolean
        Return Connect(m_WriteQueue, m_ReadQueue)
    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Disconnect from the Queue and flushes the data    ''' 
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks>
    ''' To Reconnect, you can simply call the Connect(String,String) function    ''' 
    ''' </remarks>
    ''' <history>
    ''' 	[U910KRZI]	05/05/2004	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Function Disconnect() As Boolean
        m_Connected = False
        m_InputQueue = Nothing
        m_OutputQueue = Nothing
        m_GetOptions = Nothing
        m_PutOptions = Nothing
        Return True
    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Reads a message from the Queue    ''' 
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks>
    ''' The Message is not return to the caller, use ReadMessage to obtain the Message
    ''' Check MQSeries.ErrNum for success code. If succeeded, return value will be zero    ''' 
    ''' </remarks>
    ''' <history>
    ''' 	[U910KRZI]	05/05/2004	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Function GetMessage(ByRef MsgObj As Object, Optional ByVal MsgID As String = "", Optional ByVal bWait As Boolean = False) As Integer

        If Not m_Connected Then
            m_ErrDsc = "Not Connected to Queue!"
            m_ErrNum = &H80001234
            Return Nothing
        End If

        Dim my_GetOptions As New Object
        Dim my_MQSess As New Object
        Dim RetVal As Integer
        my_GetOptions = m_GetOptions
        my_MQSess = m_MQSess
        my_GetOptions.Options = MQGMO_NO_SYNCPOINT
        If bWait Then my_GetOptions.Options += MQGMO_WAIT

        ''DoWait = bWait

        MsgObj = my_MQSess.AccessMessage()
        'GetMsg.Encoding = 546 'DefEnc_Windows = 546
        'GetMsg.Format = MQFMT_STRING
        my_MQSess.ExceptionThreshold = 3              '* process GetMsg errors in line
        If IsValidMid(MsgID) Then
            If MsgID.Length = 12 Then
                'm_GetMsg.MessageId = "12345678901234567890ABCD"
                MsgObj.MessageId = MsgID & MsgID  '24 characters
            End If
        End If
        m_InputQueue.Get(MsgObj, my_GetOptions)

        my_MQSess.ExceptionThreshold = 2

        m_ErrNum = my_MQSess.ReasonCode
        m_ErrDsc = my_MQSess.ReasonName
        RetVal = my_MQSess.ReasonCode
        ''Dim RetStr As String
        ''If my_MQSess.ReasonCode <> MQRC_NO_MSG_AVAILABLE Then
        ''    Dim rc As Char()
        ''    ReDim rc(MsgObj.DataLength - 1)
        ''    For i As Integer = 0 To rc.GetUpperBound(0)
        ''        rc(i) = Chr(MsgObj.ReadUnsignedByte())
        ''    Next
        ''    RetStr = rc
        ''    m_ErrNum = 0
        ''    m_ErrDsc = ""
        ''    Return RetStr
        ''End If

        my_GetOptions = Nothing
        my_MQSess = Nothing

        Return RetVal
    End Function


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Returns a string represenation of the message read by MQSeries    ''' 
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[U910KRZI]	05/05/2004	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Function ReadMessageByByte(ByRef MsgObj As Object) As String
        'Old function, reads one byte at time and attempts a transaltion
        'Clients linked to use ReadMessage will automatically use the newer and faster
        'function ReadMessageFast instead of this old implementation
        Dim RetStr As String
        Dim rc As Char()
        ReDim rc(MsgObj.DataLength - 1)
        For i As Integer = 0 To rc.Length - 1
            rc(i) = Chr(MsgObj.ReadUnsignedByte())
        Next
        RetStr = rc
        m_ErrNum = 0
        m_ErrDsc = ""
        Return RetStr
        'End If
        MsgObj = Nothing
    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Returns a string represenation of the message read by MQSeries    ''' 
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[U910KRZI]	05/05/2004	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Function ReadMessage(ByRef MsgObj As Object) As String
        'Provided for compatibility only
        Return ReadMessageFast(MsgObj)
    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Returns a string represenation of the message read by MQSeries    ''' 
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[U910KRZI]	05/05/2004	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Function ReadMessageFast(ByRef MsgObj As Object) As String
        Dim RetStr As String
        'If m_MQSess.ReasonCode <> MQRC_NO_MSG_AVAILABLE Then
        MsgObj.CharacterSet = 1256
        ' If attempted to read the whole DataLength using ReadString, then an exception will be raised
        'the reason behind this behavior remains unknown at this point.
        ' So a work around is to read DataLength-1, then read a single unsigned byte (which will be the
        'remaining byte in the buffer and cast it as a character and append it to the string.
        ' This workaround proved to have insignificant impact on performance, and thus it is accepted.
        RetStr = MsgObj.ReadString(MsgObj.DataLength - 1)
        RetStr &= Chr(MsgObj.ReadUnsignedByte())
        m_ErrNum = 0
        m_ErrDsc = ""
        Return RetStr
        'End If
        MsgObj = Nothing
    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Function that builds an MQ message to be sent, but does not send it out, it is
    ''' used for logging purposes.
    ''' </summary>
    ''' <param name="MsgCode"></param>
    ''' <param name="MsgData"></param>
    ''' <param name="MsgID"></param>
    ''' <param name="SibOrgCls"></param>
    ''' <param name="CusNum"></param>
    ''' <param name="AccNum"></param>
    ''' <param name="CrdNum"></param>
    ''' <param name="ResVal"></param>
    ''' <param name="ActCde"></param>
    ''' <param name="SysCon"></param>
    ''' <returns></returns>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[u910krzi]	03/03/2007	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Function PrepareMessage(ByVal MsgCode As String, ByVal MsgData As String, ByVal MsgID As String, ByVal SibOrgCls As String, _
                                ByVal CusNum As String, ByVal AccNum As String, ByVal CrdNum As String, ByVal ResVal As String, ByVal ActCde As String, _
                                ByVal SysCon As String, ByVal IceTryFin As String, ByVal AutUsrIde As String, ByVal AutPacVal As String, _
                                ByVal AskMidRef As String) As String
        Dim TxtSndHeader, TxtMsgBody As String
        Try
            TxtSndHeader = BuildSndHeaderEx(MsgID, MsgCode, m_MsgSrc, m_EnvSrc, m_MsgDst, m_EnvDst, SibOrgCls, CusNum, AccNum, CrdNum, ResVal, ActCde, _
                                            SysCon, IceTryFin, AutUsrIde, AutPacVal, AskMidRef)
            TxtMsgBody = TxtSndHeader & MsgData
        Catch ex As Exception
            TxtMsgBody = Nothing
        End Try
        Return TxtMsgBody
    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Write a message onto the Queue    ''' 
    ''' </summary>
    ''' <param name="MsgCode"></param>
    ''' <param name="MsgData"></param>
    ''' <returns></returns>
    ''' <remarks>
    ''' Check MQSeries.ErrNum for success code. If succeeded, return value will be zero     ''' 
    ''' </remarks>
    ''' <history>
    ''' 	[U910KRZI]	05/05/2004	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Function PutMessage(ByVal MsgCode As String, ByVal MsgData As String, Optional ByVal MsgID As String = "", _
                                Optional ByVal SibOrgCls As String = "GEN") As Boolean
        Return PutMessageEx(MsgCode, MsgData, MsgID, SibOrgCls, "", "", "")
    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Overloaded function, used to maintain compatiblity with previous code
    ''' </summary>
    ''' <param name="MsgCode"></param>
    ''' <param name="MsgData"></param>
    ''' <param name="MsgID"></param>
    ''' <param name="SibOrgCls"></param>
    ''' <param name="CusNum"></param>
    ''' <param name="AccNum"></param>
    ''' <param name="CrdNum"></param>
    ''' <param name="ResVal"></param>
    ''' <param name="ActCde"></param>
    ''' <param name="SysCon"></param>
    ''' <returns></returns>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[u910krzi]	03/03/2007	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Function PutMessageEx(ByVal MsgCode As String, ByVal MsgData As String, ByVal MsgID As String, ByVal SibOrgCls As String, _
                                ByVal CusNum As String, ByVal AccNum As String, ByVal CrdNum As String, Optional ByVal ResVal As String = "000", _
                                Optional ByVal ActCde As String = "00", Optional ByVal SysCon As String = "000") As Boolean
        Dim dummy As String = Nothing
        Return PutMessageEx(MsgCode, MsgData, MsgID, SibOrgCls, CusNum, AccNum, CrdNum, dummy, ResVal, ActCde, SysCon)
    End Function


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Write a message onto the Queue    ''' 
    ''' </summary>
    ''' <param name="MsgCode"></param>
    ''' <param name="MsgData"></param>
    ''' <returns></returns>
    ''' <remarks>
    ''' Check MQSeries.ErrNum for success code. If succeeded, return value will be zero     ''' 
    ''' </remarks>
    ''' <history>
    ''' 	[U910KRZI]	05/05/2004	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Function PutMessageEx(ByVal MsgCode As String, ByVal MsgData As String, ByVal MsgID As String, ByVal SibOrgCls As String, _
                                ByVal CusNum As String, ByVal AccNum As String, ByVal CrdNum As String, ByRef MsgDump As String, ByVal ResVal As String, _
                                ByVal ActCde As String, ByVal SysCon As String) As Boolean
        Return PutMessageFinancial(MsgCode, MsgData, MsgID, SibOrgCls, CusNum, AccNum, CrdNum, MsgDump, ResVal, ActCde, SysCon, "", "", "")
    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Write a message onto the Queue    ''' 
    ''' </summary>
    ''' <param name="MsgCode"></param>
    ''' <param name="MsgData"></param>
    ''' <returns></returns>
    ''' <remarks>
    ''' Check MQSeries.ErrNum for success code. If succeeded, return value will be zero     ''' 
    ''' </remarks>
    ''' <history>
    ''' 	[U910KRZI]	05/05/2004	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Function PutMessageFinancial(ByVal MsgCode As String, ByVal MsgData As String, ByVal MsgID As String, ByVal SibOrgCls As String, _
                                ByVal CusNum As String, ByVal AccNum As String, ByVal CrdNum As String, ByRef MsgDump As String, ByVal ResVal As String, _
                                ByVal ActCde As String, ByVal SysCon As String, ByVal IceTryFin As String, ByVal AutUsrIde As String, ByVal AutPacVal As String) As Boolean
        Dim my_PutMsg As Object
        Dim RetVal As Boolean = False

        If Not m_Connected Then
            m_ErrDsc = "Not Connected to Queue!"
            m_ErrNum = MQRC_NOT_OPEN
            Return False
        End If

        Dim TxtSndHeader As String
        Dim TxtMsgBody As String

        Try

            TxtSndHeader = BuildSndHeaderEx(MsgID, MsgCode, m_MsgSrc, m_EnvSrc, m_MsgDst, m_EnvDst, SibOrgCls, _
                                            CusNum, AccNum, CrdNum, ResVal, ActCde, SysCon, IceTryFin, AutUsrIde, AutPacVal)
            TxtMsgBody = TxtSndHeader & MsgData

            my_PutMsg = m_MQSess.AccessMessage
            my_PutMsg.ReplyToQueueManagerName = m_QueueManager '"MQ.GOLF"
            my_PutMsg.ReplyToQueueName = m_ReplyToQueue
            my_PutMsg.Encoding = 420 'DefEnc_Windows = 546
            my_PutMsg.Format = MQFMT_NONE 'MQFMT_STRING
            my_PutMsg.CharacterSet = 1256
            my_PutMsg.Expiry = CInt((m_TmeOut - 1000) / 100) '' Value is set to m_TmeOut - 1 sec
            my_PutMsg.MessageId = Nothing
            If IsValidMid(MsgID) Then
                If MsgID.Length = 12 Then
                    'm_PutMsg.MessageId = "12345678901234567890ABCD"
                    my_PutMsg.MessageId = MsgID & MsgID '24 characters
                End If
            End If
            Debug.Print(TxtMsgBody)
            my_PutMsg.WriteString(TxtMsgBody)
            MsgDump = TxtMsgBody
            m_OutputQueue.Put(my_PutMsg, m_PutOptions)
            m_SentMsgID = my_PutMsg.MessageId
            RetVal = True
        Catch ex As Exception
            RetVal = False
        End Try
        m_ErrNum = m_MQSess.ReasonCode
        m_ErrDsc = m_MQSess.ReasonName
        Return RetVal
    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Function that posts a pre-prepared message (made using PrepareMessage function)
    ''' to the host.
    ''' </summary>
    ''' <param name="MsgID"></param>
    ''' <param name="MsgBody"></param>
    ''' <returns></returns>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[u910krzi]	03/03/2007	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Function PutMessagePrepared(ByVal MsgID As String, ByVal MsgBody As String) As Boolean
        Dim my_PutMsg As Object
        Dim RetVal As Boolean = False

        If Not m_Connected Then
            m_ErrDsc = "Not Connected to Queue!"
            m_ErrNum = MQRC_NOT_OPEN
            Return False
        End If


        Try
            my_PutMsg = m_MQSess.AccessMessage
            my_PutMsg.ReplyToQueueManagerName = m_QueueManager '"MQ.GOLF"
            my_PutMsg.ReplyToQueueName = m_ReplyToQueue
            my_PutMsg.Encoding = 420 'DefEnc_Windows = 546
            my_PutMsg.Format = MQFMT_NONE 'MQFMT_STRING
            my_PutMsg.CharacterSet = 1256
            my_PutMsg.Expiry = CInt((m_TmeOut - 1000) / 100) '' Value is set to m_TmeOut - 1 sec
            my_PutMsg.MessageId = Nothing
            If IsValidMid(MsgID) Then
                If MsgID.Length = 12 Then
                    'm_PutMsg.MessageId = "12345678901234567890ABCD"
                    my_PutMsg.MessageId = MsgID & MsgID '24 characters
                End If
            End If

            my_PutMsg.WriteString(MsgBody)
            m_OutputQueue.Put(my_PutMsg, m_PutOptions)
            m_SentMsgID = my_PutMsg.MessageId
            RetVal = True
        Catch ex As Exception
            RetVal = False
        End Try
        m_ErrNum = m_MQSess.ReasonCode
        m_ErrDsc = m_MQSess.ReasonName
        Return RetVal
    End Function

    Public Function PutMessageRaw1(ByVal MsgData As String, Optional ByVal MsgID As String = "", Optional ByVal ReplyQueue As String = "", Optional ByVal CorlID As String = "") As Boolean
        Dim my_PutMsg As Object
        Dim RetVal As Boolean = False

        If Not m_Connected Then
            m_ErrDsc = "Not Connected to Queue!"
            m_ErrNum = MQRC_NOT_OPEN
            Return False
        End If

        Try
            my_PutMsg = m_MQSess.AccessMessage
            my_PutMsg.ReplyToQueueManagerName = m_QueueManager '"MQ.GOLF"
            my_PutMsg.ReplyToQueueName = CStr(IIf(ReplyQueue = "", m_ReplyToQueue, ReplyQueue))
            my_PutMsg.Encoding = 420 'DefEnc_Windows = 546
            my_PutMsg.Format = MQFMT_NONE
            my_PutMsg.CharacterSet = 1256
            my_PutMsg.Expiry = CInt((m_TmeOut - 1000) / 100) '' Value is set to m_TmeOut - 1 sec
            my_PutMsg.MessageId = Nothing
            If IsValidMid(MsgID) Then
                If MsgID.Length = 12 Then
                    'm_PutMsg.MessageId = "12345678901234567890ABCD"
                    my_PutMsg.MessageId = MsgID & MsgID '24 characters
                End If
            End If

            If Trim(CorlID) = "" Then
                my_PutMsg.CorrelationId = MsgID
            Else
                my_PutMsg.CorrelationId = PackString(CorlID, 24)
            End If

            my_PutMsg.WriteString(MsgData)
            m_OutputQueue.Put(my_PutMsg, m_PutOptions)
            m_SentMsgID = my_PutMsg.MessageId
            RetVal = True
        Catch ex As Exception
            RetVal = False
        End Try
        m_ErrNum = m_MQSess.ReasonCode
        m_ErrDsc = m_MQSess.ReasonName
        Return RetVal
    End Function

    Public Function PutMessageInQueueEx(ByVal WriteQueueName As String, _
                                ByVal MsgCode As String, ByVal MsgData As String, ByVal MsgID As String, ByVal SibOrgCls As String, _
                                ByVal CusNum As String, ByVal AccNum As String, ByVal CrdNum As String, Optional ByVal ResVal As String = "000", _
                                Optional ByVal ActCde As String = "00", Optional ByVal SysCon As String = "000", _
                                Optional ByVal ReplyQueue As String = "", Optional ByVal CorlID As String = "") As Boolean
        Dim TxtSndHeader As String
        Dim TxtMsgBody As String

        TxtSndHeader = BuildSndHeaderEx(MsgID, MsgCode, m_MsgSrc, m_EnvSrc, m_MsgDst, m_EnvDst, SibOrgCls, CusNum, AccNum, CrdNum, ResVal, ActCde, SysCon)
        TxtMsgBody = TxtSndHeader & MsgData

        Return PutMessageInQueue(WriteQueueName, TxtMsgBody, MsgID, ReplyQueue, CorlID)
    End Function

    Public Function PutMessageInQueue(ByVal WriteQueueName As String, ByVal MsgData As String, _
                                                    Optional ByVal MsgID As String = "", Optional ByVal ReplyQueue As String = "", _
                                                    Optional ByVal CorlID As String = "") As Boolean
        Dim lOutQueue As Object
        Dim lPutOptions As Object
        Dim my_PutMsg As Object
        Dim RetVal As Boolean = False

        Try
            lOutQueue = m_QMgr.AccessQueue(WriteQueueName, MQOO_OUTPUT)
            lPutOptions = m_MQSess.AccessPutMessageOptions()
            lPutOptions.Options = MQPMO_NO_SYNCPOINT
            my_PutMsg = m_MQSess.AccessMessage
            my_PutMsg.ReplyToQueueManagerName = m_QueueManager '"MQ.GOLF"
            my_PutMsg.ReplyToQueueName = CStr(IIf(ReplyQueue = "", m_ReplyToQueue, ReplyQueue))
            my_PutMsg.Encoding = 420 'DefEnc_Windows = 546
            my_PutMsg.Format = MQFMT_NONE
            my_PutMsg.CharacterSet = 1256
            my_PutMsg.Expiry = CInt((m_TmeOut - 1000) / 100) '' Value is set to m_TmeOut - 1 sec
            my_PutMsg.MessageId = Nothing
            If IsValidMid(MsgID) Then
                If MsgID.Length = 12 Then
                    'm_PutMsg.MessageId = "12345678901234567890ABCD"
                    my_PutMsg.MessageId = MsgID & MsgID '24 characters
                Else
                    my_PutMsg.MessageId = PackString(MsgID, 24)
                End If
            End If

            If Trim(CorlID) = "" Then
                my_PutMsg.CorrelationId = PackString(MsgID, 24)
            Else
                my_PutMsg.CorrelationId = PackString(CorlID, 24)
            End If

            my_PutMsg.WriteString(MsgData)
            lOutQueue.Put(my_PutMsg, lPutOptions)
            m_SentMsgID = my_PutMsg.MessageId
            RetVal = True
        Catch ex As Exception
            RetVal = False
        End Try
        m_ErrNum = m_MQSess.ReasonCode
        m_ErrDsc = m_MQSess.ReasonName
        Return RetVal
    End Function

    Public Function PutMessage83(ByVal MsgCode As String, ByVal MsgData As String, ByVal MsgID As String) As Boolean

        Dim my_PutMsg As Object
        Dim RetVal As Boolean = False

        If Not m_Connected Then
            m_ErrDsc = "Not Connected to Queue!"
            m_ErrNum = MQRC_NOT_OPEN
            Return False
        End If

        Dim TxtSndHeader As String
        Dim TxtMsgBody As String

        Try
            TxtSndHeader = BuildSndHeader83(MsgID, MsgCode, m_MsgSrc, m_MsgDst)
            TxtMsgBody = TxtSndHeader & MsgData


            my_PutMsg = m_MQSess.AccessMessage
            my_PutMsg.ReplyToQueueManagerName = m_QueueManager '"MQ.GOLF"
            my_PutMsg.ReplyToQueueName = m_ReplyToQueue
            my_PutMsg.Encoding = 420 'DefEnc_Windows = 546
            my_PutMsg.Format = MQFMT_NONE
            my_PutMsg.CharacterSet = 1256
            my_PutMsg.Expiry = CInt((m_TmeOut - 1000) / 100) '' Value is set to m_TmeOut - 1 sec
            my_PutMsg.MessageId = Nothing
            If IsValidMid(MsgID) Then
                If MsgID.Length = 12 Then
                    'm_PutMsg.MessageId = "12345678901234567890ABCD"
                    my_PutMsg.MessageId = MsgID & MsgID '24 characters
                End If
            End If

            my_PutMsg.WriteString(TxtMsgBody)
            m_OutputQueue.Put(my_PutMsg, m_PutOptions)
            m_SentMsgID = my_PutMsg.MessageId
            RetVal = True
        Catch ex As Exception
            RetVal = False
        End Try
        m_ErrNum = m_MQSess.ReasonCode
        m_ErrDsc = m_MQSess.ReasonName
        Return RetVal
    End Function

    Public Function PutMessage93(ByVal MsgCode As String, ByVal MsgData As String, ByVal MsgID As String, ByVal ResVal As String, _
                                ByVal ActCde As String, ByVal SysCon As String) As Boolean

        Dim my_PutMsg As Object
        Dim RetVal As Boolean = False

        If Not m_Connected Then
            m_ErrDsc = "Not Connected to Queue!"
            m_ErrNum = MQRC_NOT_OPEN
            Return False
        End If

        Dim TxtSndHeader As String
        Dim TxtMsgBody As String

        Try
            TxtSndHeader = BuildSndHeader93(MsgID, MsgCode, m_MsgSrc, m_MsgDst, ResVal, ActCde, SysCon)
            TxtMsgBody = TxtSndHeader & MsgData


            my_PutMsg = m_MQSess.AccessMessage
            my_PutMsg.ReplyToQueueManagerName = m_QueueManager '"MQ.GOLF"
            my_PutMsg.ReplyToQueueName = m_ReplyToQueue
            my_PutMsg.Encoding = 420 'DefEnc_Windows = 546
            my_PutMsg.Format = MQFMT_NONE
            my_PutMsg.CharacterSet = 1256
            my_PutMsg.Expiry = CInt((m_TmeOut - 1000) / 100) '' Value is set to m_TmeOut - 1 sec
            my_PutMsg.MessageId = Nothing
            If IsValidMid(MsgID) Then
                If MsgID.Length = 12 Then
                    'm_PutMsg.MessageId = "12345678901234567890ABCD"
                    my_PutMsg.MessageId = MsgID & MsgID '24 characters
                End If
            End If

            my_PutMsg.WriteString(TxtMsgBody)
            m_OutputQueue.Put(my_PutMsg, m_PutOptions)
            m_SentMsgID = my_PutMsg.MessageId
            RetVal = True
        Catch ex As Exception
            RetVal = False
        End Try
        m_ErrNum = m_MQSess.ReasonCode
        m_ErrDsc = m_MQSess.ReasonName
        Return RetVal
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
            RetStr(i) = " "
        Next
        Return RetStr
    End Function

    Private Function PackHeader123(ByVal SibHeader As SibHdrStruct) As String
        Dim RetStr As String
        Try
            With SibHeader
                RetStr = PackString(.MsgProCde, 6)
                RetStr += PackString(.MsgMid, 12)
                RetStr += PackString(.MsgStp, 14)
                RetStr += PackString(.MsgSrc, 4)
                RetStr += PackString(.MsgSrcEnv, 3)
                RetStr += PackString(.MsgSrcRsv, 3)
                RetStr += PackString(.MsgUsrIde, 32)
                RetStr += PackString(.MsgTgt, 4)
                RetStr += PackString(.MsgTgtEnv, 3)
                RetStr += PackString(.MsgTgtRsv, 2)
                RetStr += PackString(.MsgResVal, 3)
                RetStr += PackString(.MsgActCde, 2)
                RetStr += PackString(.MsgSysCon, 3)
                RetStr += PackString(.Rsv1b, 1)
                RetStr += PackString(.MsgResInd, 1)
                RetStr += PackString(.SibResCde, 7)
                RetStr += PackString(.SifRefNum, 12)
                RetStr += PackString(.SibEnvUnt, 3)
                RetStr += PackString(.SibOrgCls, 3)
                ' SibKeyLoc, 84
                RetStr += PackString(.SibKeyCls, 3)
                RetStr += PackString(.SibCus, 6)
                RetStr += PackString(.SibAcc, 13)
                RetStr += PackString(.SibVar, 20)
                '---RetStr += PackString(.SibDea, 20)
                '---RetStr += PackString(.SibRetIde, 12)
                '---RetStr += PackString(.SibTrmIde, 4)
                '---RetStr += PackString(.SibCrdNum, 16)
                RetStr += PackString(.MsiDat, 10)
                RetStr += PackString(.Rsv21, 10)
                RetStr += PackString(.CstVerTab, 12)
                RetStr += PackString(.SibCusSeg, 1)
                RetStr += PackString(.Rsv2, 9)
                'SubStruct:SibNonRep	Size:50
                'RetStr += PackString("0303", 4) 
                RetStr += PackString(m_BrnCde, 4)
                RetStr += PackString(.Rsv415, 45)
                RetStr += PackString("I", 1)  'RetStr += PackString(.SibNrdTyp, 1)
                'rsv3
                RetStr += PackString(.Rsv3, 4)
            End With
        Catch
            RetStr = Space(Len(SibHeader))
        End Try
        Return RetStr

    End Function

    Private Function PackHeader83(ByVal SibHeader As SibHdrStruct83) As String
        Dim RetStr As String
        Try
            With SibHeader
                RetStr = PackString(.MsgProCde, 6)
                RetStr += PackString(.MsgMid, 12)
                RetStr += PackString(.MsgStp, 14)
                RetStr += PackString(.MsgOrg, 10)
                RetStr += PackString(.MsgUsrIde, 32)
                RetStr += PackString(.MsgDst, 9)
            End With
        Catch
            RetStr = Space(Len(SibHeader))
        End Try
        Return RetStr
    End Function

    Private Function PackHeader93(ByVal SibHeader As SibHdrStruct93) As String
        Dim RetStr As String
        Try
            With SibHeader
                RetStr = PackString(.MsgProCde, 6)
                RetStr += PackString(.MsgMid, 12)
                RetStr += PackString(.MsgStp, 14)
                RetStr += PackString(.MsgOrg, 10)
                RetStr += PackString(.MsgUsrIde, 32)
                RetStr += PackString(.MsgDst, 9)
                RetStr += PackString(.MsgResVal, 3)
                RetStr += PackString(.MsgActCde, 2)
                RetStr += PackString(.MsgSysCon, 3)
                RetStr += PackString(.Rsv1, 2)
            End With
        Catch
            RetStr = Space(Len(SibHeader))
        End Try
        Return RetStr
    End Function

    Private Function ToHexID(ByVal DecID As String) As String
        Dim tt As String = "6963653A"
        For i As Integer = 0 To DecID.Length - 1
            tt += "3" & DecID.Chars(i)
        Next
        Return tt
    End Function

    Private Function US_FormatDate(ByVal pDtpDate As DateTime, ByVal pFormatString As String) As String
        Dim MyResult As String
        Dim OldCultInfo, NewCultInfo As CultureInfo
        '* Use English US locale
        OldCultInfo = System.Threading.Thread.CurrentThread.CurrentCulture
        NewCultInfo = New CultureInfo("en-US", False)
        System.Threading.Thread.CurrentThread.CurrentCulture = NewCultInfo

        MyResult = Format(pDtpDate, pFormatString)
        System.Threading.Thread.CurrentThread.CurrentCulture = OldCultInfo
        NewCultInfo = Nothing
        Return MyResult
    End Function

    Private Function BuildSndHeader(ByVal MsgId As String, ByVal pTrxCde As String, ByVal pMsgSrc As String, _
                                    ByVal pSrcEnv As String, ByVal pMsgDst As String, ByVal pDstEnv As String, _
                                    Optional ByVal SibOrgCls As String = "GEN") As String
        Return BuildSndHeaderEx(MsgId, pTrxCde, pMsgSrc, pSrcEnv, pMsgDst, pDstEnv, SibOrgCls, "", "", "", "000", "00", "000")
    End Function

    Private Function BuildSndHeaderEx(ByVal MsgId As String, ByVal pTrxCde As String, ByVal pMsgSrc As String, _
                                    ByVal pSrcEnv As String, ByVal pMsgDst As String, ByVal pDstEnv As String, _
                                    ByVal SibOrgCls As String, ByVal CusNum As String, ByVal AccNum As String, ByVal CrdNum As String, _
                                    ByVal ResVal As String, ByVal ActCde As String, ByVal SysCon As String, _
                                    Optional ByVal IceTryFin As String = "", _
                                    Optional ByVal AutUsrIde As String = "", Optional ByVal AutPacVal As String = "", _
                                    Optional ByVal AskMidRef As String = "") As String
        Dim SibHdr As New SibHdrStruct
        'Dim TmpStr As String
        Try
            ErrClr()
            With SibHdr
                .MsgProCde = pTrxCde
                'TmpStr = "000000000000" & Replace(CStr(CLng(Now.ToOADate)), ".", "") & Replace(CStr(VB6.Format(CDbl(VB.Timer()), "#0.000")), ".", "") '& "000000000000"
                .MsgMid = MsgId 'Right(TmpStr, 12)
                .MsgStp = US_FormatDate(Now, "yyyyMMddHHmmss") '& Format(Timer - Int(Timer), ".000")
                .MsgSrc = pMsgSrc
                .MsgSrcEnv = pSrcEnv
                .MsgSrcRsv = "   "
                .MsgUsrIde = m_LoggedUser & ";" & m_LoggedUser
                .MsgTgt = pMsgDst
                .MsgTgtEnv = pDstEnv
                .MsgTgtRsv = "  "
                .MsgResVal = ResVal
                .MsgActCde = ActCde
                .MsgSysCon = SysCon
                .SibOrgCls = SibOrgCls
                .SibKeyCls = SibOrgCls
                .SibCus = CusNum
                .SibAcc = AccNum
                .SibVar = "    " & CrdNum
                '.SibCrdNum = CrdNum
                .IceTryFin = IceTryFin
                .AutUsrIde = AutUsrIde
                .AutPacVal = AutPacVal
                .AskMidRef = AskMidRef
            End With
            Return (SibHdr.PackHeader(Me))
        Catch ex As Exception
            m_ErrNum = &H80001234
            m_ErrDsc = ex.Message
        End Try
        Return (SibHdr.PackHeader(Me))
    End Function

    Private Function BuildSndHeader83(ByVal MsgId As String, ByVal pTrxCde As String, ByVal pMsgSrc As String, _
                                    ByVal pMsgDst As String) As String
        Dim SibHdr As New SibHdrStruct83
        ' Dim TmpStr As String
        Try
            ErrClr()
            With SibHdr
                .MsgProCde = pTrxCde
                'TmpStr = "000000000000" & Replace(CStr(CLng(Now.ToOADate)), ".", "") & Replace(CStr(VB6.Format(CDbl(VB.Timer()), "#0.000")), ".", "") '& "000000000000"
                .MsgMid = MsgId 'Right(TmpStr, 12)
                .MsgStp = US_FormatDate(Now, "yyyyMMddHHmmss") '& Format(Timer - Int(Timer), ".000")
                .MsgOrg = pMsgSrc
                .MsgUsrIde = m_LoggedUser & ";" & m_LoggedUser
                .MsgDst = pMsgDst
            End With
            Return (PackHeader83(SibHdr))
        Catch ex As Exception
            m_ErrNum = &H80001234
            m_ErrDsc = ex.Message
        End Try
        Return (PackHeader83(SibHdr))
    End Function

    Private Function BuildSndHeader93(ByVal MsgId As String, ByVal pTrxCde As String, ByVal pMsgSrc As String, _
                                    ByVal pMsgDst As String, ByVal ResVal As String, ByVal ActCde As String, _
                                    ByVal SysCon As String) As String
        Dim SibHdr As New SibHdrStruct93
        'Dim TmpStr As String
        Try
            ErrClr()
            With SibHdr
                .MsgProCde = pTrxCde
                'TmpStr = "000000000000" & Replace(CStr(CLng(Now.ToOADate)), ".", "") & Replace(CStr(VB6.Format(CDbl(VB.Timer()), "#0.000")), ".", "") '& "000000000000"
                .MsgMid = MsgId 'Right(TmpStr, 12)
                .MsgStp = US_FormatDate(Now, "yyyyMMddHHmmss") '& Format(Timer - Int(Timer), ".000")
                .MsgOrg = pMsgSrc
                .MsgUsrIde = m_LoggedUser & ";" & m_LoggedUser
                .MsgDst = pMsgDst
                .MsgResVal = ResVal
                .MsgActCde = ActCde
                .MsgSysCon = SysCon
                .Rsv1 = Space(100)
            End With
            Return (PackHeader93(SibHdr))
        Catch ex As Exception
            m_ErrNum = &H80001234
            m_ErrDsc = ex.Message
        End Try
        Return (PackHeader93(SibHdr))
    End Function

End Class
