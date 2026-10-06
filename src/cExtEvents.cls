VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cExtEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'=========================================================================
'
'   Ucs SQL Monitor
'   Copyright (c) 2026 Unicontsoft
'
'   Shared UcsSqlMonitor extended events session and its incremental reader.
'   Every monitor holds a shared app lock, the last one to leave drops the session.
'   Statement events are added by the first monitor tracing them, the others skip them
'
'=========================================================================
Option Explicit
DefObj A-Z
Private Const MODULE_NAME As String = "cExtEvents"

'=========================================================================
' Constants and member variables
'=========================================================================

Private Const STR_SESSION_NAME          As String = "UcsSqlMonitor"
Private Const DBL_RESYNC_INTERVAL       As Double = 60
Private Const LNG_OBJTYPE_FUNCTION      As Long = 20038
Private Const LNG_OBJTYPE_PREPARED      As Long = 20816
Private Const ERR_SQL_XE_NOT_FOUND      As Long = 15151
Private Const ERR_SQL_XE_EVENT_EXISTS   As Long = 25625
Private Const ERR_SQL_XE_EXISTS         As Long = 25631
Private Const ERR_SQL_XE_STARTED        As Long = 25705

Private m_oConn                     As ADODB.Connection
Private m_oCmd                      As ADODB.Command
Private m_oResultCmd                As ADODB.Command
Private m_bStatements               As Boolean
Private m_sPath                     As String
Private m_sPattern                  As String
Private m_bPathStale                As Boolean
Private m_vFile                     As Variant
Private m_vOffset                   As Variant
Private m_vDropped                  As Variant
Private m_dblLastFull               As Double

'=========================================================================
' Error handling
'=========================================================================

Private Sub PrintError(sFunc As String)
    Debug.Print MODULE_NAME & "." & sFunc & ": " & Error
End Sub

'=========================================================================
' Methods
'=========================================================================

'--- takes the shared lock and creates/starts the session, raises on failure
Public Sub Init(oConn As ADODB.Connection, ByVal bStatements As Boolean)
    Dim rs              As Recordset
    Dim lResult         As Long
    Dim dblStart        As Double

    dblStart = TraceStart()
    Set m_oConn = oConn
    m_bStatements = bStatements
    Set rs = m_oConn.Execute("SET NOCOUNT ON" & vbCrLf & _
        "DECLARE @rc INT" & vbCrLf & _
        "EXEC @rc = master.sys.sp_getapplock @Resource = N'" & STR_SESSION_NAME & "', @LockMode = 'Shared', @LockOwner = 'Session', @LockTimeout = 10000" & vbCrLf & _
        "SELECT @rc")
    lResult = rs.Fields(0).Value
    '--- with ODBC an unread result makes the next command run on a hidden connection
    rs.Close
    TraceEnd "xe.applock", dblStart
    If lResult < 0 Then
        On Error GoTo 0
        Err.Raise vbObjectError, , "Cannot lock " & STR_SESSION_NAME & " event session (" & lResult & ")"
    End If
    pvEnsureSession
    pvReadTargetPath
    pvCreateResultTables
    If m_oCmd Is Nothing Then
        Set m_oCmd = New ADODB.Command
        m_oCmd.CommandText = pvGetFetchSql()
        m_oCmd.Parameters.Append m_oCmd.CreateParameter("file", adVarWChar, adParamInput, 260)
        m_oCmd.Parameters.Append pvCreateDecimalParam("offset")
        m_oCmd.Parameters.Append pvCreateDecimalParam("dropped")
        m_oCmd.Parameters.Append m_oCmd.CreateParameter("full", adInteger, adParamInput)
        m_oCmd.Parameters.Append m_oCmd.CreateParameter("system", adInteger, adParamInput)
        m_oCmd.Parameters.Append m_oCmd.CreateParameter("statements", adInteger, adParamInput)
        m_oCmd.Parameters.Append m_oCmd.CreateParameter("path", adVarWChar, adParamInput, 260)
        m_oCmd.Parameters.Append m_oCmd.CreateParameter("pattern", adVarWChar, adParamInput, 260)
        m_vFile = Null
        m_vOffset = Null
        m_vDropped = 0
    End If
    Set m_oCmd.ActiveConnection = m_oConn
    m_oCmd.CommandTimeout = 30
    If m_oResultCmd Is Nothing Then
        Set m_oResultCmd = New ADODB.Command
        m_oResultCmd.CommandText = pvGetResultSql()
    End If
    Set m_oResultCmd.ActiveConnection = m_oConn
    m_oResultCmd.CommandTimeout = 30
    TraceEnd "xe.init", dblStart
End Sub

'--- drops the session when no other monitor holds the shared lock. Runs on a connection of
'--- its own as the monitoring one may still be busy, its close already released this lock
Public Sub DropIfLast(sConnectString As String)
    Const FUNC_NAME     As String = "DropIfLast"
    Dim oConn           As ADODB.Connection
    Dim dblStart        As Double

    On Error GoTo EH
    dblStart = TraceStart()
    Set m_oCmd = Nothing
    Set m_oResultCmd = Nothing
    Set m_oConn = Nothing
    Set oConn = New ADODB.Connection
    oConn.ConnectionTimeout = 5
    oConn.Open sConnectString
    oConn.Execute "DECLARE @rc INT" & vbCrLf & _
        "EXEC @rc = master.sys.sp_getapplock @Resource = N'" & STR_SESSION_NAME & "', @LockMode = 'Exclusive', @LockOwner = 'Session', @LockTimeout = 1000" & vbCrLf & _
        "IF @rc >= 0" & vbCrLf & _
        "BEGIN" & vbCrLf & _
        "    BEGIN TRY" & vbCrLf & _
        "        DROP EVENT SESSION [" & STR_SESSION_NAME & "] ON SERVER" & vbCrLf & _
        "    END TRY" & vbCrLf & _
        "    BEGIN CATCH" & vbCrLf & _
        "        IF ERROR_NUMBER() <> " & ERR_SQL_XE_NOT_FOUND & " THROW" & vbCrLf & _
        "    END CATCH" & vbCrLf & _
        "    EXEC master.sys.sp_releaseapplock @Resource = N'" & STR_SESSION_NAME & "', @LockOwner = 'Session'" & vbCrLf & _
        "END", , adExecuteNoRecords
    oConn.Close
    TraceEnd "xe.drop", dblStart
    Exit Sub
EH:
    PrintError FUNC_NAME
End Sub

'--- the batch returning what the last populate batch left in the session's tables
Public Property Get ResultCommand() As ADODB.Command
    Set ResultCommand = m_oResultCmd
End Property

'--- the per-tick populate batch: reads new events and fills the session's tables, returns nothing
Public Function FetchCommand(ByVal bSystem As Boolean) As ADODB.Command
    Dim bFull           As Boolean

    bFull = (TimerEx - m_dblLastFull > DBL_RESYNC_INTERVAL)
    '--- the target is asked again only when the batch could not read the file it was given
    If m_bPathStale Or LenB(m_sPath) = 0 Then
        pvReadTargetPath
    End If
    m_oCmd.Parameters("file").Value = m_vFile
    m_oCmd.Parameters("offset").Value = m_vOffset
    m_oCmd.Parameters("dropped").Value = m_vDropped
    m_oCmd.Parameters("full").Value = IIf(bFull, 1, 0)
    m_oCmd.Parameters("system").Value = IIf(bSystem, 1, 0)
    m_oCmd.Parameters("statements").Value = IIf(m_bStatements, 1, 0)
    m_oCmd.Parameters("path").Value = IIf(LenB(m_sPath) <> 0, m_sPath, Null)
    m_oCmd.Parameters("pattern").Value = IIf(LenB(m_sPattern) <> 0, m_sPattern, Null)
    Set FetchCommand = m_oCmd
End Function

'--- one history entry: a comment line with time and stats, then the statement
Public Function FormatEvent(rs As Recordset) As String
    Dim sHeader         As String
    Dim sText           As String

    sHeader = "-- " & C_Str(rs!event_time.Value) & " "
    sText = Replace(Replace(C_Str(rs!sql_text.Value), vbCrLf, vbLf), vbLf, vbCrLf)
    Select Case C_Str(rs!event_name.Value)
    Case "error_reported"
        FormatEvent = sHeader & "ERROR " & C_Str(rs!error_number.Value) & " (severity " & C_Str(rs!severity.Value) & "): " & _
            C_Str(rs!message.Value) & vbCrLf
        Exit Function
    Case "attention"
        FormatEvent = sHeader & "CANCELLED after " & C_Str(rs!duration_ms.Value) & " ms" & vbCrLf
        Exit Function
    Case "sp_statement_completed"
        sHeader = sHeader & "stmt " & C_Str(rs!object_name.Value) & ":" & C_Str(rs!line_number.Value)
    Case "rpc_completed"
        sHeader = sHeader & "rpc"
    Case Else
        sHeader = sHeader & "batch"
    End Select
    sHeader = sHeader & ", " & C_Str(rs!duration_ms.Value) & " ms, cpu " & C_Str(rs!cpu_ms.Value) & " ms, reads " & _
        C_Str(rs!logical_reads.Value) & ", writes " & C_Str(rs!writes.Value) & ", rows " & C_Str(rs!row_count.Value)
    If LenB(C_Str(rs!result.Value)) <> 0 And C_Str(rs!result.Value) <> "OK" Then
        sHeader = sHeader & ", " & C_Str(rs!result.Value)
    End If
    If Right$(sText, 2) <> vbCrLf Then
        sText = sText & vbCrLf
    End If
    FormatEvent = sHeader & vbCrLf & sText & "GO" & vbCrLf
End Function

'--- the first recordset of the batch, advances the read position
Friend Function frReadInfo(rs As Recordset, bFull As Boolean, lLost As Long) As Boolean
    If rs.EOF Then
        Exit Function
    End If
    If C_Lng(rs!is_missing.Value) <> 0 Then
        '--- someone dropped the session while it was in use, start over
        pvEnsureSession
        m_bPathStale = True
        m_vFile = Null
        m_vOffset = Null
        m_vDropped = 0
        Exit Function
    End If
    m_bPathStale = (C_Lng(rs!is_stale.Value) <> 0)
    m_vFile = rs!file_name.Value
    m_vOffset = rs!file_offset.Value
    m_vDropped = rs!dropped.Value
    lLost = C_Lng(rs!lost.Value)
    bFull = (C_Lng(rs!is_full.Value) <> 0)
    If bFull Then
        m_dblLastFull = TimerEx
    End If
    frReadInfo = True
End Function

Private Sub pvEnsureSession()
    Dim cSql            As Collection
    Dim dblStart        As Double

    dblStart = TraceStart()
    Set cSql = New Collection
    cSql.Add "IF NOT EXISTS (SELECT * FROM sys.server_event_sessions WHERE name = N'" & STR_SESSION_NAME & "')"
    cSql.Add "BEGIN TRY"
    cSql.Add "CREATE EVENT SESSION [" & STR_SESSION_NAME & "] ON SERVER"
    cSql.Add "ADD EVENT sqlserver.login (SET collect_options_text = 0"
    cSql.Add "    ACTION (package0.event_sequence, sqlserver.session_id) WHERE sqlserver.is_system = 0),"
    cSql.Add "ADD EVENT sqlserver.logout ("
    cSql.Add "    ACTION (package0.event_sequence, sqlserver.session_id) WHERE sqlserver.is_system = 0),"
    cSql.Add "ADD EVENT sqlserver.sql_batch_completed ("
    cSql.Add "    ACTION (package0.event_sequence, sqlserver.session_id, sqlserver.database_name)"
    cSql.Add "    WHERE sqlserver.is_system = 0 AND " & pvNotOwnApp() & "),"
    cSql.Add "ADD EVENT sqlserver.rpc_completed ("
    cSql.Add "    ACTION (package0.event_sequence, sqlserver.session_id, sqlserver.database_name)"
    cSql.Add "    WHERE sqlserver.is_system = 0 AND " & pvNotOwnApp() & " AND object_name <> N'sp_reset_connection'),"
    cSql.Add "ADD EVENT sqlserver.attention ("
    cSql.Add "    ACTION (package0.event_sequence, sqlserver.session_id, sqlserver.sql_text)"
    cSql.Add "    WHERE sqlserver.is_system = 0 AND " & pvNotOwnApp() & "),"
    cSql.Add "ADD EVENT sqlserver.error_reported ("
    cSql.Add "    ACTION (package0.event_sequence, sqlserver.session_id, sqlserver.sql_text)"
    cSql.Add "    WHERE severity >= 11 AND sqlserver.is_system = 0 AND " & pvNotOwnApp() & ")"
    cSql.Add "ADD TARGET package0.event_file (SET filename = N'" & STR_SESSION_NAME & "', max_file_size = 100, max_rollover_files = 4)"
    cSql.Add "WITH (MAX_MEMORY = 8 MB, EVENT_RETENTION_MODE = ALLOW_SINGLE_EVENT_LOSS, MAX_DISPATCH_LATENCY = 1 SECONDS,"
    cSql.Add "    MAX_EVENT_SIZE = 8 MB, TRACK_CAUSALITY = OFF, STARTUP_STATE = OFF)"
    cSql.Add "END TRY"
    cSql.Add "BEGIN CATCH"
    '--- another monitor may have created it first
    cSql.Add "    IF ERROR_NUMBER() <> " & ERR_SQL_XE_EXISTS & " THROW"
    cSql.Add "END CATCH"
    cSql.Add "IF NOT EXISTS (SELECT * FROM sys.dm_xe_sessions WHERE name = N'" & STR_SESSION_NAME & "')"
    cSql.Add "BEGIN TRY"
    cSql.Add "    ALTER EVENT SESSION [" & STR_SESSION_NAME & "] ON SERVER STATE = START"
    cSql.Add "END TRY"
    cSql.Add "BEGIN CATCH"
    '--- another monitor may have started it first
    cSql.Add "    IF ERROR_NUMBER() <> " & ERR_SQL_XE_STARTED & " THROW"
    cSql.Add "END CATCH"
    If m_bStatements Then
        '--- added to the running session on demand, it stays until the last monitor drops it
        cSql.Add "IF NOT EXISTS (SELECT * FROM sys.dm_xe_sessions s JOIN sys.dm_xe_session_events e ON e.event_session_address = s.address"
        cSql.Add "        WHERE s.name = N'" & STR_SESSION_NAME & "' AND e.event_name = N'sp_statement_completed')"
        cSql.Add "BEGIN TRY"
        '--- scalar FUNCTION statements fire per row, PREPARED repeats the sp_executesql text
        cSql.Add "    ALTER EVENT SESSION [" & STR_SESSION_NAME & "] ON SERVER"
        cSql.Add "    ADD EVENT sqlserver.sp_statement_completed ("
        cSql.Add "        ACTION (package0.event_sequence, sqlserver.session_id, sqlserver.database_name)"
        cSql.Add "        WHERE sqlserver.is_system = 0 AND " & pvNotOwnApp() & _
            " AND object_type <> " & LNG_OBJTYPE_FUNCTION & " AND object_type <> " & LNG_OBJTYPE_PREPARED & ")"
        cSql.Add "END TRY"
        cSql.Add "BEGIN CATCH"
        '--- another monitor may have added it first
        cSql.Add "    IF ERROR_NUMBER() <> " & ERR_SQL_XE_EVENT_EXISTS & " THROW"
        cSql.Add "END CATCH"
    End If
    m_oConn.Execute ConcatCollection(cSql), , adExecuteNoRecords
    TraceEnd "xe.ensure", dblStart
End Sub

Private Sub pvReadTargetPath()
    Dim rs              As Recordset
    Dim lPos            As Long
    Dim dblStart        As Double

    dblStart = TraceStart()
    m_sPath = vbNullString
    m_sPattern = vbNullString
    m_bPathStale = False
    Set rs = m_oConn.Execute("SELECT CAST(t.target_data AS XML).value('(EventFileTarget/File/@name)[1]', 'NVARCHAR(260)')" & vbCrLf & _
        "FROM sys.dm_xe_sessions s JOIN sys.dm_xe_session_targets t ON t.event_session_address = s.address" & vbCrLf & _
        "WHERE s.name = N'" & STR_SESSION_NAME & "' AND t.target_name = N'event_file'")
    If Not rs.EOF Then
        m_sPath = C_Str(rs.Fields(0).Value)
    End If
    '--- with ODBC an unread result makes the next command run on a hidden connection
    rs.Close
    lPos = InStrRev(m_sPath, "\")
    If lPos > 0 Then
        m_sPattern = Left$(m_sPath, lPos) & STR_SESSION_NAME & "*.xel"
    End If
    TraceEnd "xe.path", dblStart
End Sub

Private Function pvGetFetchSql() As String
    Dim cSql            As Collection

    Set cSql = New Collection
    cSql.Add "SET NOCOUNT ON"
    cSql.Add "DECLARE @file NVARCHAR(260) = ?, @offset BIGINT = ?, @dropped BIGINT = ?, @full INT = ?, @system INT = ?, @statements INT = ?, @path NVARCHAR(260) = ?, @pattern NVARCHAR(260) = ?"
    cSql.Add "DECLARE @cur_dropped BIGINT, @since DATETIME2 = DATEADD(MINUTE, -5, SYSUTCDATETIME())"
    cSql.Add "DECLARE @new_file NVARCHAR(260), @new_offset BIGINT, @lost INT = 0, @stale INT = 0"
    cSql.Add "SELECT @cur_dropped = dropped_event_count FROM sys.dm_xe_sessions WHERE name = N'" & STR_SESSION_NAME & "'"
    cSql.Add "DROP TABLE IF EXISTS #raw"
    cSql.Add "CREATE TABLE #raw (file_name NVARCHAR(260), file_offset BIGINT, event_data NVARCHAR(MAX))"
    '--- reading from a position returns only the buffers after it
    cSql.Add "IF @path IS NOT NULL AND @file IS NOT NULL"
    cSql.Add "BEGIN TRY"
    cSql.Add "    INSERT #raw SELECT file_name, file_offset, event_data FROM sys.fn_xe_file_target_read_file(@pattern, NULL, @file, @offset)"
    cSql.Add "END TRY"
    cSql.Add "BEGIN CATCH"
    cSql.Add "    SELECT @file = NULL, @lost = 1"
    cSql.Add "END CATCH"
    '--- first read takes only the current file and only its recent events
    '--- a stale path is reported back, so the next fetch asks the target again
    cSql.Add "IF @path IS NOT NULL AND @file IS NULL"
    cSql.Add "BEGIN TRY"
    cSql.Add "    INSERT #raw SELECT file_name, file_offset, IIF(timestamp_utc >= @since, event_data, NULL) FROM sys.fn_xe_file_target_read_file(@path, NULL, NULL, NULL)"
    cSql.Add "    SELECT @full = 1, @dropped = @cur_dropped"
    cSql.Add "END TRY"
    cSql.Add "BEGIN CATCH"
    cSql.Add "    SET @stale = 1"
    cSql.Add "END CATCH"
    '--- parse each event once, .value() on CAST in an APPLY parses it again for every column
    cSql.Add "DROP TABLE IF EXISTS #x"
    cSql.Add "CREATE TABLE #x (file_name NVARCHAR(260), file_offset BIGINT, x XML)"
    cSql.Add "INSERT #x SELECT file_name, file_offset, CAST(event_data AS XML) FROM #raw WHERE event_data IS NOT NULL"
    cSql.Add "DROP TABLE IF EXISTS #ev"
    cSql.Add "SELECT  r.file_name, r.file_offset,"
    cSql.Add "        e.value('(@name)[1]', 'NVARCHAR(60)') AS event_name,"
    cSql.Add "        e.value('(action[@name=""event_sequence""]/value)[1]', 'BIGINT') AS seq,"
    cSql.Add "        e.value('(action[@name=""session_id""]/value)[1]', 'INT') AS session_id,"
    cSql.Add "        e.value('(@timestamp)[1]', 'DATETIME2') AS ts,"
    cSql.Add "        COALESCE(e.value('(data[@name=""batch_text""]/value)[1]', 'NVARCHAR(MAX)'),"
    cSql.Add "            e.value('(data[@name=""statement""]/value)[1]', 'NVARCHAR(MAX)'),"
    cSql.Add "            e.value('(action[@name=""sql_text""]/value)[1]', 'NVARCHAR(MAX)')) AS sql_text,"
    cSql.Add "        NULLIF(e.value('(data[@name=""object_name""]/value)[1]', 'NVARCHAR(256)'), N'') AS object_name,"
    cSql.Add "        e.value('(data[@name=""object_id""]/value)[1]', 'INT') AS object_id,"
    cSql.Add "        e.value('(data[@name=""source_database_id""]/value)[1]', 'INT') AS source_database_id,"
    cSql.Add "        e.value('(data[@name=""line_number""]/value)[1]', 'INT') AS line_number,"
    cSql.Add "        e.value('(data[@name=""duration""]/value)[1]', 'BIGINT') AS duration,"
    cSql.Add "        e.value('(data[@name=""cpu_time""]/value)[1]', 'BIGINT') AS cpu_time,"
    cSql.Add "        e.value('(data[@name=""logical_reads""]/value)[1]', 'BIGINT') AS logical_reads,"
    cSql.Add "        e.value('(data[@name=""writes""]/value)[1]', 'BIGINT') AS writes,"
    cSql.Add "        e.value('(data[@name=""row_count""]/value)[1]', 'BIGINT') AS row_count,"
    cSql.Add "        e.value('(data[@name=""result""]/text)[1]', 'NVARCHAR(20)') AS result,"
    cSql.Add "        e.value('(data[@name=""error_number""]/value)[1]', 'INT') AS error_number,"
    cSql.Add "        e.value('(data[@name=""severity""]/value)[1]', 'INT') AS severity,"
    cSql.Add "        e.value('(data[@name=""message""]/value)[1]', 'NVARCHAR(4000)') AS message"
    cSql.Add "INTO    #ev"
    cSql.Add "FROM    #x r CROSS APPLY r.x.nodes('/event') n(e)"
    '--- the shared session collects statements for whoever traces them
    cSql.Add "IF @statements = 0"
    cSql.Add "    DELETE #ev WHERE event_name = N'sp_statement_completed'"
    '--- OBJECT_NAME per event row is very slow so resolve each distinct object once
    cSql.Add "DROP TABLE IF EXISTS #obj"
    cSql.Add "SELECT DISTINCT object_id, source_database_id, CAST(NULL AS NVARCHAR(256)) AS object_name INTO #obj FROM #ev WHERE object_name IS NULL AND object_id IS NOT NULL"
    cSql.Add "UPDATE #obj SET object_name = OBJECT_NAME(object_id, source_database_id)"
    cSql.Add "UPDATE e SET object_name = o.object_name FROM #ev e JOIN #obj o ON o.object_id = e.object_id AND o.source_database_id = e.source_database_id WHERE e.object_name IS NULL"
    cSql.Add "SELECT TOP 1 @new_file = file_name, @new_offset = file_offset FROM #raw ORDER BY file_name DESC, file_offset DESC"
    cSql.Add "IF @new_file IS NULL"
    cSql.Add "    SELECT @new_file = @file, @new_offset = @offset"
    cSql.Add "IF @cur_dropped > @dropped OR EXISTS (SELECT * FROM #ev WHERE event_name IN (N'login', N'logout'))"
    cSql.Add "    SET @full = 1"
    '--- the results go to the session's tables, a second batch started when this one is done
    '--- returns them. The tables are the session's as ODBC runs this batch in a scope of its own
    cSql.Add "TRUNCATE TABLE #snap"
    '--- sessions snapshot, all on full resync otherwise only those with new events or a running request
    cSql.Add "INSERT  #snap (SPID, Status, Login, Host, DB, Program, CPU, Dsk, Last_Batch, LoginTime)"
    cSql.Add "SELECT  s.session_id, s.status, s.login_name, s.host_name, DB_NAME(s.database_id), LEFT(s.program_name, 120), s.cpu_time,"
    cSql.Add "        CAST(IIF(s.reads + s.writes > 2147483647, 2147483647, s.reads + s.writes) AS INT),"
    cSql.Add "        FORMAT(s.last_request_start_time, 'MM/dd HH:mm:ss'), s.login_time"
    cSql.Add "FROM    sys.dm_exec_sessions s"
    cSql.Add "WHERE   (@full = 1 OR s.session_id IN (SELECT session_id FROM #ev) OR s.session_id IN (SELECT session_id FROM sys.dm_exec_requests))"
    cSql.Add "        AND s.session_id <> @@SPID"
    cSql.Add "        AND (@system = 1 OR (s.is_user_process = 1 AND ISNULL(s.program_name, N'') <> N'" & STR_APP_NAME & "'))"
    cSql.Add "TRUNCATE TABLE #live"
    '--- live requests and open transactions
    cSql.Add "INSERT  #live (SPID, Status, Command, Blk, Wait, Trans, DB)"
    cSql.Add "SELECT  r.session_id, r.status, r.command, NULLIF(r.blocking_session_id, 0), r.wait_type, NULLIF(r.open_transaction_count, 0), DB_NAME(r.database_id)"
    cSql.Add "FROM    sys.dm_exec_requests r"
    cSql.Add "WHERE   r.session_id <> @@SPID AND r.session_id IN (SELECT session_id FROM sys.dm_exec_sessions WHERE @system = 1 OR is_user_process = 1)"
    cSql.Add "UNION ALL"
    cSql.Add "SELECT  s.session_id, s.status, NULL, NULL, NULL, s.open_transaction_count, DB_NAME(s.database_id)"
    cSql.Add "FROM    sys.dm_exec_sessions s"
    cSql.Add "WHERE   s.open_transaction_count > 0 AND s.session_id <> @@SPID AND (@system = 1 OR s.is_user_process = 1)"
    cSql.Add "        AND NOT EXISTS (SELECT * FROM sys.dm_exec_requests r WHERE r.session_id = s.session_id)"
    cSql.Add "TRUNCATE TABLE #out"
    '--- new events, auto-stats and empty statements are noise
    cSql.Add "INSERT  #out (file_name, file_offset, seq, SPID, event_name, event_time, sql_text, object_name, line_number,"
    cSql.Add "        duration_ms, cpu_ms, logical_reads, writes, row_count, result, error_number, severity, message)"
    cSql.Add "SELECT  file_name, file_offset, seq, session_id, event_name,"
    cSql.Add "        FORMAT(DATEADD(MINUTE, DATEDIFF(MINUTE, SYSUTCDATETIME(), SYSDATETIME()), ts), 'HH:mm:ss.fff'),"
    cSql.Add "        sql_text, object_name, line_number, duration / 1000, cpu_time / 1000, logical_reads, writes, row_count,"
    cSql.Add "        result, error_number, severity, message"
    cSql.Add "FROM    #ev"
    cSql.Add "WHERE   event_name NOT IN (N'login', N'logout')"
    cSql.Add "        AND NOT (event_name = N'sp_statement_completed' AND (ISNULL(sql_text, N'') = N'' OR sql_text LIKE N'%[_]MS[_]UPDSTATS[_]TBL%'))"
    cSql.Add "TRUNCATE TABLE #info"
    cSql.Add "INSERT  #info (file_name, file_offset, dropped, lost, is_full, is_missing, is_stale)"
    cSql.Add "SELECT  @new_file, @new_offset, ISNULL(@cur_dropped, 0), IIF(@cur_dropped > @dropped, @cur_dropped - @dropped, 0) + @lost,"
    cSql.Add "        @full, IIF(@cur_dropped IS NULL, 1, 0), @stale"
    pvGetFetchSql = ConcatCollection(cSql)
End Function

'--- the results of the last populate batch: 1. info, 2. sessions snapshot, 3. live requests,
'--- 4. new events. Bigint is DECIMAL as VB6 has no 64-bit Variant, NTEXT as MSDASQL garbles NVARCHAR(MAX)
Private Function pvGetResultSql() As String
    Dim cSql            As Collection

    Set cSql = New Collection
    cSql.Add "SELECT  file_name, file_offset, dropped, lost, is_full, is_missing, is_stale FROM #info"
    cSql.Add "SELECT  SPID, Status, Login, Host, DB, Program, CPU, Dsk, Last_Batch, LoginTime FROM #snap"
    cSql.Add "SELECT  SPID, Status, Command, Blk, Wait, Trans, DB FROM #live"
    cSql.Add "SELECT  SPID, event_name, event_time, CAST(sql_text AS NTEXT) AS sql_text, object_name, line_number,"
    cSql.Add "        duration_ms, cpu_ms, logical_reads, writes, row_count, result, error_number, severity, message"
    cSql.Add "FROM    #out"
    cSql.Add "ORDER BY file_name, file_offset, seq"
    pvGetResultSql = ConcatCollection(cSql)
End Function

'--- the tables the populate batch fills and the result batch reads, in the session's own scope
'--- so they outlive the batch
Private Sub pvCreateResultTables()
    Dim cSql            As Collection

    Set cSql = New Collection
    cSql.Add "DROP TABLE IF EXISTS #info"
    cSql.Add "CREATE TABLE #info (file_name NVARCHAR(260), file_offset DECIMAL(19, 0), dropped DECIMAL(19, 0), lost INT,"
    cSql.Add "    is_full INT, is_missing INT, is_stale INT)"
    cSql.Add "DROP TABLE IF EXISTS #snap"
    cSql.Add "CREATE TABLE #snap (SPID INT, Status NVARCHAR(30), Login NVARCHAR(128), Host NVARCHAR(128), DB NVARCHAR(128),"
    cSql.Add "    Program NVARCHAR(120), CPU INT, Dsk INT, Last_Batch NVARCHAR(20), LoginTime DATETIME)"
    cSql.Add "DROP TABLE IF EXISTS #live"
    cSql.Add "CREATE TABLE #live (SPID INT, Status NVARCHAR(30), Command NVARCHAR(32), Blk INT, Wait NVARCHAR(60), Trans INT, DB NVARCHAR(128))"
    cSql.Add "DROP TABLE IF EXISTS #out"
    cSql.Add "CREATE TABLE #out (file_name NVARCHAR(260), file_offset BIGINT, seq BIGINT, SPID INT, event_name NVARCHAR(60),"
    cSql.Add "    event_time NVARCHAR(20), sql_text NVARCHAR(MAX), object_name NVARCHAR(256), line_number INT,"
    cSql.Add "    duration_ms DECIMAL(19, 0), cpu_ms DECIMAL(19, 0), logical_reads DECIMAL(19, 0), writes DECIMAL(19, 0),"
    cSql.Add "    row_count DECIMAL(19, 0), result NVARCHAR(20), error_number INT, severity INT, message NVARCHAR(4000))"
    m_oConn.Execute ConcatCollection(cSql), , adExecuteNoRecords
End Sub

Private Function pvNotOwnApp() As String
    pvNotOwnApp = "sqlserver.client_app_name <> N'" & STR_APP_NAME & "'"
End Function

Private Function pvCreateDecimalParam(sName As String) As ADODB.Parameter
    Set pvCreateDecimalParam = m_oCmd.CreateParameter(sName, adDecimal, adParamInput)
    pvCreateDecimalParam.Precision = 19
    pvCreateDecimalParam.NumericScale = 0
End Function
