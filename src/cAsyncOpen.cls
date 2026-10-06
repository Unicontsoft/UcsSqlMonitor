VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cAsyncOpen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'=========================================================================
'
'   Ucs SQL Monitor
'   Copyright (c) 2026 Unicontsoft
'
'   Opens an ADO connection asynchronously and keeps the message loop going
'   until it is open or has failed, so the window stays responsive while a
'   server is slow or gone. Callers guard against what the user can do meanwhile
'
'=========================================================================
Option Explicit
DefObj A-Z
Private Const MODULE_NAME As String = "cAsyncOpen"

'=========================================================================
' Constants and member variables
'=========================================================================

Private Const LNG_PUMP_SLEEP            As Long = 10
Private Const DBL_EVENT_GRACE           As Double = 2
Private Const DBL_TIMEOUT_GRACE         As Double = 3
Private Const STR_TIMEOUT               As String = "Timeout expired"

Private WithEvents m_oConn          As ADODB.Connection
Attribute m_oConn.VB_VarHelpID = -1
Private m_bDone                     As Boolean
Private m_sError                    As String

'=========================================================================
' Methods
'=========================================================================

'--- raises with the provider's message when the connection cannot be opened. SQLOLEDB does not
'--- keep to the connection timeout when asynchronous, so the wait ends on its own deadline and
'--- the connection still being opened is left to the caller to release, which does not block
Public Sub OpenConnection(oConn As ADODB.Connection, sConnectString As String)
    Const FUNC_NAME     As String = "OpenConnection"
    Dim dblIdle         As Double
    Dim dblDeadline     As Double

    m_bDone = False
    m_sError = vbNullString
    Set m_oConn = oConn
    dblDeadline = TimerEx + oConn.ConnectionTimeout + DBL_TIMEOUT_GRACE
    oConn.Open sConnectString, , , adAsyncConnect
    Do While Not m_bDone
        If TimerEx > dblDeadline Then
            m_sError = STR_TIMEOUT
            Exit Do
        End If
        If (oConn.State And adStateConnecting) = 0 Then
            '--- the outcome comes with the event, wait a little for it once the provider is done
            If dblIdle = 0 Then
                dblIdle = TimerEx
            ElseIf TimerEx - dblIdle > DBL_EVENT_GRACE Then
                Exit Do
            End If
        End If
        DoEvents
        Call Sleep(LNG_PUMP_SLEEP)
    Loop
    Set m_oConn = Nothing
    If oConn.State <> adStateOpen Then
        Err.Raise vbObjectError, MODULE_NAME & "." & FUNC_NAME, IIf(LenB(m_sError) <> 0, m_sError, "Cannot connect")
    End If
End Sub

'=========================================================================
' Control events
'=========================================================================

Private Sub m_oConn_ConnectComplete(ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pConnection As ADODB.Connection)
    If adStatus = adStatusErrorsOccurred Then
        If Not pError Is Nothing Then
            m_sError = pError.Description
        End If
    End If
    m_bDone = True
End Sub
