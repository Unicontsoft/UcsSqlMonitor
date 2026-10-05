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
Private Const LNG_TIMEOUT               As Long = 5

Private m_sServer                   As String
Private m_eMode                     As UcsMonitorMode
Private m_bSystemProcesses          As Boolean
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
            sConnectString As String, _
            sServer As String, _
            oOwner As frmMain)
    m_sServer = sServer
    Set m_oCmd = oCmd
    Set m_oConn = oCmd.ActiveConnection
    m_bClosing = False
    m_eMode = eMode
    m_bSystemProcesses = bSystemProcesses
    m_sConnectString = sConnectString
    m_dblInterval = 1 / lRefreshRate
    pvStart oOwner
End Sub

'--- opens a new connection with the settings it was first connected with, raises when that fails
Public Sub Reconnect(oOwner As frmMain)
    m_bClosing = False
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
End Sub

'--- called on every timer tick, fetches when this server's refresh interval has passed
Public Sub Tick()
    If m_oConn Is Nothing Or m_bClosing Then
        Exit Sub
    End If
    If TimerEx - m_dblLastFetch >= m_dblInterval Then
        pvFetch
    End If
End Sub

Public Sub KillSession(ByVal lSpid As Long)
    m_oConn.Execute "KILL " & lSpid, , adExecuteNoRecords
End Sub

'--- releases the extended events session if this was its last monitor and closes the connection
Public Sub Shutdown()
    Const FUNC_NAME     As String = "Shutdown"

    On Error GoTo EH
    '--- cancelling the fetch fires FetchComplete, which must not start another one meanwhile
    m_bClosing = True
    pvCloseResult
    '--- without pooling the close ends the session, which releases its shared lock
    If Not m_oConn Is Nothing Then
        If m_oConn.State <> adStateClosed Then
            m_oConn.Close
        End If
    End If
    If Not m_oExtEvents Is Nothing Then
        m_oExtEvents.DropIfLast m_sConnectString
        Set m_oExtEvents = Nothing
    End If
QH:
    Set m_oCmd = Nothing
    Set m_oConn = Nothing
    Set m_oOwner = Nothing
    Exit Sub
EH:
    PrintError FUNC_NAME
    Resume QH
End Sub

Private Sub pvStart(oOwner As frmMain)
    If m_eMode = ucsMonExtEvents Then
        Set m_oExtEvents = New cExtEvents
        m_oExtEvents.Init m_oConn
    End If
    Set m_oOwner = oOwner
End Sub

Private Sub pvFetch()
    Const FUNC_NAME     As String = "pvFetch"
    Dim lState          As Long
    Dim oCmd            As ADODB.Command

    On Error GoTo EH
    If m_bClosing Then
        Exit Sub
    End If
    m_dblLastFetch = TimerEx
    If m_rsResult Is Nothing Then
        Set m_rsResult = New Recordset
        m_rsResult.CursorLocation = adUseClient
    End If
    On Error Resume Next
    lState = m_rsResult.State
    If Err.Number <> 0 Then
        Debug.Print MODULE_NAME & ".pvFetch(" & m_sServer & ") Error="; Error, TimerEx
        pvTestConn
    End If
    On Error GoTo EH
    If (lState And (adStateConnecting Or adStateExecuting Or adStateFetching)) <> 0 Then
        m_bDelayFetch = True
        Exit Sub
    End If
    If lState = adStateOpen Then
        m_rsResult.Close
    End If
    If m_eMode = ucsMonExtEvents Then
        Set oCmd = m_oExtEvents.FetchCommand(m_bSystemProcesses)
    Else
        Set oCmd = m_oCmd
    End If
    On Error Resume Next
    m_rsResult.Open oCmd, , adOpenStatic, adLockBatchOptimistic, adAsyncExecute Or adAsyncFetch
    If Err.Number <> 0 Then
        Debug.Print MODULE_NAME & ".pvFetch(" & m_sServer & ") Error="; Error, TimerEx
        pvTestConn
    End If
    Exit Sub
EH:
    PrintError FUNC_NAME
    Resume Next
End Sub

Private Sub pvTestConn()
    On Error Resume Next
    If m_oConn.State = adStateOpen Then
        m_oConn.Execute "SELECT @@TRANCOUNT"
        If Err.Number <> 0 Then
            m_oConn.Close
        End If
    End If
    If m_oConn.State = adStateClosed Then
        m_oConn.Open m_sConnectString
        Set m_oCmd.ActiveConnection = m_oConn
        If Not m_oExtEvents Is Nothing Then
            m_oExtEvents.Init m_oConn
        End If
    End If
End Sub

Private Sub pvCloseResult()
    On Error GoTo QH
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

Private Sub m_oConn_Disconnect(adStatus As ADODB.EventStatusEnum, ByVal pConnection As ADODB.Connection)
    Debug.Print MODULE_NAME & ".m_oConn_Disconnect(" & m_sServer & ") adStatus="; adStatus, TimerEx
End Sub

Private Sub m_oConn_InfoMessage(ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pConnection As ADODB.Connection)
    Const FUNC_NAME     As String = "m_oConn_InfoMessage"

    On Error GoTo EH
    '--- ODBC passes on the login's changed database/language notices
    Select Case pError.NativeError
    Case ERR_SQL_DB_CHANGED, ERR_SQL_LANG_CHANGED
        Exit Sub
    End Select
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

    On Error GoTo EH
    If m_bClosing Then
        Exit Sub
    End If
    If Not pError Is Nothing Then
        Debug.Print MODULE_NAME & ".m_rsResult_FetchComplete(" & m_sServer & ") " & pError.Description
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
    Exit Sub
EH:
    PrintError FUNC_NAME
    Resume Next
End Sub
