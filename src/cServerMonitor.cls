VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cServerMonitor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'=========================================================================
'
'   Ucs SQL Monitor
'   Copyright (c) 2026 Unicontsoft
'
'   One connected SQL Server: its connection, the sp_who2 command or the
'   extended events reader, and the async fetch that feeds the main list
'
'=========================================================================
Option Explicit
DefObj A-Z
Private Const MODULE_NAME As String = "cServerMonitor"

'=========================================================================
' Constants and member variables
'=========================================================================

Private Const ERR_NO_MORE_RESULTS       As Long = &H40EC9
Private Const ERR_SQL_DB_CHANGED        As Long = 5701
Private Const ERR_SQL_LANG_CHANGED      As Long = 5703
Private Const STR_ODBC_NO_CURSOR        As String = "Executing SQL directly; no cursor"
Private Const LNG_TIMEOUT               As Long = 5
Private Const DBL_RETRY_MIN             As Double = 1
Private Const DBL_RETRY_MAX             As Double = 30

Private m_sServer                   As String
Private m_eMode                     As UcsMonitorMode
Private m_bSystemProcesses          As Boolean
Private m_bStatements               As Boolean
Private m_sConnectString            As String
Private m_oCmd                      As ADODB.Command
Private WithEvents m_oConn          As ADODB.Connection
Attribute m_oConn.VB_VarHelpID = -1
Private WithEvents m_rsResult       As ADODB.Recordset
Attribute m_rsResult.VB_VarHelpID = -1
Private m_oExtEvents                As cExtEvents
Private m_oOwner                    As frmMain
Private m_dblInterval               As Double
Private m_dblLastFetch              As Double
Private m_bDelayFetch               As Boolean
Private m_bClosing                  As Boolean
Private m_dblFetchStart             As Double
Private m_bReopening                As Boolean
Private m_dblRetryAt                As Double
Private m_dblRetryDelay             As Double
Private m_oPopulateCmd               As ADODB.Command
Private m_bPopulating               As Boolean

'=========================================================================
' Error handling
'=========================================================================

Private Sub PrintError(sFunc As String)
    Debug.Print MODULE_NAME & "." & sFunc & "(" & m_sServer & "): " & Error
End Sub

'=========================================================================
' Properties
'=========================================================================

Public Property Get Server() As String
    Server = m_sServer
End Property

Public Property Get Mode() As UcsMonitorMode
    Mode = m_eMode
End Property

Public Property Get SystemProcesses() As Boolean
    SystemProcesses = m_bSystemProcesses
End Property

Public Property Get ExtEvents() As cExtEvents
    Set ExtEvents = m_oExtEvents
End Property

Public Property Get Connected() As Boolean
    Connected = Not m_oConn Is Nothing
End Property

'=========================================================================
' Methods
'=========================================================================

'--- takes over the connection the Connect dialog opened, raises when it cannot be monitored
Public Sub Init( _
            oCmd As ADODB.Command, _
            ByVal eMode As UcsMonitorMode, _
            ByVal lRefreshRate As Long, _
            ByVal bSystemProcesses As Boolean, _
            ByVal bStatements As Boolean, _
            sConnectString As String, _
            sServer As String, _
            oOwner As frmMain)
    Dim dblStart        As Double

    dblStart = TraceStart()
    m_sServer = sServer
    Set m_oCmd = oCmd
    Set m_oConn = oCmd.ActiveConnection
    m_bClosing = False
    m_eMode = eMode
    m_bSystemProcesses = bSystemProcesses
    m_bStatements = bStatements
    m_sConnectString = sConnectString
    m_dblInterval = 1 / lRefreshRate
    pvStart oOwner
    TraceEnd "monitor.init", dblStart, m_sServer
End Sub

'--- opens a new connection with the settings it was first connected with, raises when that fails
Public Sub Reconnect(oOwner As frmMain)
    Dim dblStart        As Double

    dblStart = TraceStart()
    m_bClosing = False
    m_bReopening = False
    m_dblRetryDelay = 0
    Set m_oConn = New ADODB.Connection
    m_oConn.ConnectionTimeout = LNG_TIMEOUT
    m_oConn.Open m_sConnectString
    Set m_oCmd = New ADODB.Command
    Set m_oCmd.ActiveConnection = m_oConn
    m_oCmd.CommandTimeout = LNG_TIMEOUT
    If m_eMode = ucsMonSpWho2 Then
        m_oCmd.CommandText = "exec sp_who2"
    End If
    m_dblLastFetch = 0
    pvStart oOwner
    TraceEnd "monitor.reconnect", dblStart, m_sServer
End Sub

'--- called on every timer tick, fetches when this server's refresh interval has passed and
'--- opens a lost connection again once its retry delay has passed
Public Sub Tick()
    If m_oConn Is Nothing Or m_bClosing Then
        Exit Sub
    End If
    Select Case m_oConn.State
    Case adStateOpen
        If TimerEx - m_dblLastFetch >= m_dblInterval Then
            pvFetch
        End If
    Case adStateClosed
        If Not m_bReopening And TimerEx >= m_dblRetryAt Then
            pvReopen
        End If
    End Select
End Sub

Public Sub KillSession(ByVal lSpid As Long)
    m_oConn.Execute "KILL " & lSpid, , adExecuteNoRecords
End Sub

'--- releases the extended events session if this was its last monitor and closes the connection
Public Sub Shutdown()
    Const FUNC_NAME     As String = "Shutdown"
    Dim dblStart        As Double
    Dim bConnecting     As Boolean

    On Error GoTo EH
    dblStart = TraceStart()
    '--- cancelling the fetch fires FetchComplete, which must not start another one meanwhile
    m_bClosing = True
    pvCloseResult
    '--- without pooling the close ends the session, which releases its shared lock. A connection
    '--- still being opened is only released, closing it would wait for the server
    If Not m_oConn Is Nothing Then
        bConnecting = ((m_oConn.State And adStateConnecting) <> 0)
        If Not bConnecting And m_oConn.State <> adStateClosed Then
            m_oConn.Close
        End If
    End If
    '--- a server that cannot be reached right now cannot drop the session either
    If Not m_oExtEvents Is Nothing And Not bConnecting Then
        m_oExtEvents.DropIfLast m_sConnectString
    End If
    Set m_oExtEvents = Nothing
QH:
    Set m_oCmd = Nothing
    Set m_oConn = Nothing
    Set m_oOwner = Nothing
    TraceEnd "monitor.shutdown", dblStart, m_sServer
    Exit Sub
EH:
    PrintError FUNC_NAME
    Resume QH
End Sub

Private Sub pvStart(oOwner As frmMain)
    If m_eMode = ucsMonExtEvents Then
        Set m_oExtEvents = New cExtEvents
        m_oExtEvents.Init m_oConn, m_bStatements
    End If
    Set m_oOwner = oOwner
End Sub

Private Sub pvFetch()
    Const FUNC_NAME     As String = "pvFetch"
    Dim lState          As Long
    Dim oCmd            As ADODB.Command
    Dim dblStart        As Double

    On Error GoTo EH
    If m_bClosing Then
        Exit Sub
    End If
    dblStart = TraceStart()
    m_dblLastFetch = TimerEx
    If m_bPopulating Then
        m_bDelayFetch = True
        TraceEnd "fetch.busy", dblStart, m_sServer
        Exit Sub
    End If
    If m_rsResult Is Nothing Then
        Set m_rsResult = New Recordset
        m_rsResult.CursorLocation = adUseClient
    End If
    On Error Resume Next
    lState = m_rsResult.State
    If Err.Number <> 0 Then
        pvTestConn
        Exit Sub
    End If
    On Error GoTo EH
    If (lState And (adStateConnecting Or adStateExecuting Or adStateFetching)) <> 0 Then
        m_bDelayFetch = True
        TraceEnd "fetch.busy", dblStart, m_sServer
        Exit Sub
    End If
    If lState = adStateOpen Then
        m_rsResult.Close
    End If
    m_dblFetchStart = dblStart
    If m_eMode = ucsMonExtEvents Then
        '--- Extended Events fill the session's tables first, ExecuteComplete then fetches them
        Set m_oPopulateCmd = m_oExtEvents.FetchCommand(m_bSystemProcesses)
        m_bPopulating = True
        On Error Resume Next
        m_oPopulateCmd.Execute , , adAsyncExecute Or adExecuteNoRecords
        If Err.Number <> 0 Then
            m_bPopulating = False
            pvTestConn
        End If
        TraceEnd "fetch.open", dblStart, m_sServer
        Exit Sub
    End If
    Set oCmd = m_oCmd
    On Error Resume Next
    m_rsResult.Open oCmd, , adOpenStatic, adLockBatchOptimistic, adAsyncExecute Or adAsyncFetch
    If Err.Number <> 0 Then
        pvTestConn
    End If
    TraceEnd "fetch.open", dblStart, m_sServer
    Exit Sub
EH:
    PrintError FUNC_NAME
    Resume Next
End Sub

'--- a failed fetch drops the connection, Tick opens it again without waiting on the server.
'--- A dead connection could keep a probe query waiting, so none is sent
Private Sub pvTestConn()
    Const FUNC_NAME     As String = "pvTestConn"
    Dim dblStart        As Double

    On Error GoTo EH
    dblStart = TraceStart()
    pvCloseResult
    If m_oConn.State <> adStateClosed Then
        m_oConn.Close
    End If
QH:
    pvScheduleRetry
    TraceEnd "monitor.testconn", dblStart, m_sServer
    Exit Sub
EH:
    PrintError FUNC_NAME
    Resume QH
End Sub

Private Sub pvScheduleRetry()
    If m_dblRetryDelay < DBL_RETRY_MIN Then
        m_dblRetryDelay = DBL_RETRY_MIN
    Else
        m_dblRetryDelay = m_dblRetryDelay * 2
        If m_dblRetryDelay > DBL_RETRY_MAX Then
            m_dblRetryDelay = DBL_RETRY_MAX
        End If
    End If
    m_dblRetryAt = TimerEx + m_dblRetryDelay
End Sub

Private Sub pvReopen()
    Const FUNC_NAME     As String = "pvReopen"

    On Error GoTo EH
    TraceLine "monitor.reopen", 0, m_sServer & vbTab & "delay=" & m_dblRetryDelay
    m_bReopening = True
    m_oConn.Open m_sConnectString, , , adAsyncConnect
    Exit Sub
EH:
    PrintError FUNC_NAME
    m_bReopening = False
    pvScheduleRetry
End Sub

Private Sub pvCloseResult()
    On Error GoTo QH
    If m_bPopulating Then
        m_bPopulating = False
        m_oPopulateCmd.Cancel
    End If
    If m_rsResult Is Nothing Then
        Exit Sub
    End If
    If (m_rsResult.State And (adStateExecuting Or adStateFetching)) <> 0 Then
        m_rsResult.Cancel
    End If
    If (m_rsResult.State And adStateOpen) <> 0 Then
        m_rsResult.Close
    End If
QH:
    Set m_rsResult = Nothing
End Sub

'=========================================================================
' Control events
'=========================================================================

'--- outcome of an asynchronous reopen started by Tick, a dialog's open waits on its own
Private Sub m_oConn_ConnectComplete(ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pConnection As ADODB.Connection)
    Const FUNC_NAME     As String = "m_oConn_ConnectComplete"
    Dim dblStart        As Double

    On Error GoTo EH
    If Not m_bReopening Then
        Exit Sub
    End If
    m_bReopening = False
    If m_bClosing Then
        Exit Sub
    End If
    If adStatus = adStatusErrorsOccurred Or m_oConn.State <> adStateOpen Then
        pvScheduleRetry
        Exit Sub
    End If
    dblStart = TraceStart()
    Set m_oCmd.ActiveConnection = m_oConn
    If Not m_oExtEvents Is Nothing Then
        m_oExtEvents.Init m_oConn, m_bStatements
    End If
    m_dblRetryDelay = 0
    m_dblLastFetch = 0
    TraceEnd "monitor.reopened", dblStart, m_sServer
    Exit Sub
EH:
    PrintError FUNC_NAME
    '--- the session could not be set up again, drop it and try later
    pvTestConn
End Sub

'--- the populate batch is done, its results are fetched asynchronously like sp_who2's
Private Sub m_oConn_ExecuteComplete(ByVal RecordsAffected As Long, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pCommand As ADODB.Command, ByVal pRecordset As ADODB.Recordset, ByVal pConnection As ADODB.Connection)
    Const FUNC_NAME     As String = "m_oConn_ExecuteComplete"

    On Error GoTo EH
    If Not m_bPopulating Or pCommand Is Nothing Then
        Exit Sub
    End If
    If Not pCommand Is m_oPopulateCmd Then
        Exit Sub
    End If
    m_bPopulating = False
    If m_bClosing Then
        Exit Sub
    End If
    If TraceEnabled Then
        TraceLine "fetch.populate", (TimerEx - m_dblFetchStart) * 1000000#, m_sServer
    End If
    If adStatus = adStatusErrorsOccurred Then
        If Not pError Is Nothing Then
            TraceLine "fetch.populate.error", 0, m_sServer & vbTab & pError.Description
        End If
        pvTestConn
        Exit Sub
    End If
    If m_rsResult Is Nothing Then
        Set m_rsResult = New Recordset
        m_rsResult.CursorLocation = adUseClient
    ElseIf (m_rsResult.State And adStateOpen) <> 0 Then
        m_rsResult.Close
    End If
    m_rsResult.Open m_oExtEvents.ResultCommand, , adOpenStatic, adLockBatchOptimistic, adAsyncExecute Or adAsyncFetch
    Exit Sub
EH:
    PrintError FUNC_NAME
    pvTestConn
End Sub

Private Sub m_oConn_InfoMessage(ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pConnection As ADODB.Connection)
    Const FUNC_NAME     As String = "m_oConn_InfoMessage"

    On Error GoTo EH
    '--- ODBC passes on the login's changed database/language notices
    Select Case pError.NativeError
    Case ERR_SQL_DB_CHANGED, ERR_SQL_LANG_CHANGED
        Exit Sub
    End Select
    '--- and the driver's notice on a batch run without a cursor, like the populate one
    If InStr(pError.Description, STR_ODBC_NO_CURSOR) > 0 Then
        Exit Sub
    End If
    If pError.Number <> ERR_NO_MORE_RESULTS Then
        MsgBox m_sServer & ": " & pError.Description, vbExclamation
    End If
    Exit Sub
EH:
    PrintError FUNC_NAME
    Resume Next
End Sub

Private Sub m_rsResult_FetchComplete(ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    Const FUNC_NAME     As String = "m_rsResult_FetchComplete"
    Dim bDelayed        As Boolean
    Dim dblStart        As Double

    On Error GoTo EH
    If m_bClosing Then
        Exit Sub
    End If
    dblStart = TraceStart()
    If TraceEnabled And m_dblFetchStart <> 0 Then
        TraceLine "fetch.roundtrip", (dblStart - m_dblFetchStart) * 1000000#, m_sServer
    End If
    If Not pError Is Nothing Then
        Exit Sub
    End If
    bDelayed = m_bDelayFetch
    m_bDelayFetch = False
    If Not m_oOwner Is Nothing Then
        m_oOwner.frShowResults Me, m_rsResult
    End If
    If bDelayed And Not m_oConn Is Nothing Then
        pvFetch
    End If
    TraceEnd "fetch.complete", dblStart, m_sServer
    Exit Sub
EH:
    PrintError FUNC_NAME
    Resume Next
End Sub
