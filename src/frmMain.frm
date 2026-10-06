VERSION 5.00
Begin VB.Form frmMain
   Caption         =   "Ucs DB Monitor"
   ClientHeight    =   5988
   ClientLeft      =   192
   ClientTop       =   840
   ClientWidth     =   9300
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "frmMain"
   ScaleHeight     =   5988
   ScaleWidth      =   9300
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picSplitter
      BorderStyle     =   0  'None
      Height          =   3540
      Left            =   4032
      MousePointer    =   9  'Size W E
      ScaleHeight     =   3540
      ScaleWidth      =   96
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   924
      Width           =   96
   End
   Begin UcsSQLMonitor.ctxListView lvwMain
      Height          =   4632
      Left            =   336
      TabIndex        =   1
      Top             =   252
      Width           =   4212
      _ExtentX        =   7430
      _ExtentY        =   8170
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   7.8
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MultiSelect     =   -1  'True
      OwnerData       =   -1  'True
      SortHeaders     =   -1  'True
   End
   Begin VB.PictureBox picTreeSplitter
      BorderStyle     =   0  'None
      Height          =   3540
      Left            =   2100
      MousePointer    =   9  'Size W E
      ScaleHeight     =   3540
      ScaleWidth      =   96
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   924
      Width           =   96
   End
   Begin UcsSQLMonitor.ctxTreeView tvwServers
      Height          =   4632
      Left            =   0
      TabIndex        =   0
      Top             =   252
      Width           =   2000
      _ExtentX        =   3528
      _ExtentY        =   8170
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}
         Name            =   "Tahoma"
         Size            =   7.8
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Timer tmrFetch
      Enabled         =   0   'False
      Interval        =   200
      Left            =   6216
      Top             =   5208
   End
   Begin VB.TextBox txtInput
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   7.8
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4548
      Left            =   4956
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   2
      Top             =   336
      Width           =   4128
   End
   Begin VB.Menu mnuMain
      Caption         =   "File"
      Index           =   0
      Begin VB.Menu mnuFile
         Caption         =   "Connect"
         Index           =   0
         Shortcut        =   ^{F2}
      End
      Begin VB.Menu mnuFile
         Caption         =   "Filter"
         Index           =   1
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuFile
         Caption         =   "Statistics"
         Index           =   2
         Shortcut        =   {F6}
      End
      Begin VB.Menu mnuFile
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu mnuFile
         Caption         =   "Exit"
         Index           =   4
      End
   End
   Begin VB.Menu mnuMain
      Caption         =   "Help"
      Index           =   1
      Begin VB.Menu mnuHelp
         Caption         =   "About"
         Index           =   0
      End
   End
   Begin VB.Menu mnuMain
      Caption         =   "Popup"
      Index           =   2
      Visible         =   0   'False
      Begin VB.Menu mnuPopup
         Caption         =   "Kill"
         Index           =   0
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnuMain
      Caption         =   "TreePopup"
      Index           =   3
      Visible         =   0   'False
      Begin VB.Menu mnuTree
         Caption         =   "Connect..."
         Index           =   0
      End
      Begin VB.Menu mnuTree
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnuTree
         Caption         =   "Disconnect"
         Index           =   2
      End
      Begin VB.Menu mnuTree
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu mnuTree
         Caption         =   "Properties..."
         Index           =   4
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const MODULE_NAME As String = "frmMain"

'=========================================================================
' Constants and member variables
'=========================================================================

Private Const STR_REG_COMMON        As String = "Common"
Private Const CLR_ACTIVE            As Long = &H80FF00
Private Const MSG_CONTINUE          As String = "Do you want to continue?"
Private Const LNG_HISTORY_SIZE      As Long = 32000
Private Const DBL_TREE_MIN_WIDTH    As Double = 300
Private Const DBL_TREE_WIDTH        As Double = 2400
Private Const STR_TREE_ROOT         As String = "SQL Servers"
Private Const STR_RES_PNG           As String = "CUSTOM"
Private Const LNG_RES_SERVERS       As Long = 101
Private Const LNG_RES_SERVER        As Long = 102
Private Const LNG_ICON_SIZE         As Long = 16
Private Const LNG_TICK_INTERVAL     As Long = 40
Private Const DBL_RENDER_INTERVAL   As Double = 0.1
Private Const LNG_COMPACT_DELETED   As Long = 200

Private m_cServers          As Collection
Private m_rsList            As Recordset
Private m_cList             As Collection
Private m_rsListSort        As Recordset
Private m_sFilter           As String
Private m_oFilter           As cRowFilter
Private m_bDown            As Boolean
Private m_dblDownX          As Double
Private m_dblRatio          As Double
Private m_dblTreeWidth      As Double
Private m_hTreeImages       As LongPtr
Private m_sMenuServer       As String
Private m_rsStats           As Recordset
Private m_cStats            As Collection
Private WithEvents m_oFrmStats As frmStats
Attribute m_oFrmStats.VB_VarHelpID = -1
Private m_cSelected         As Collection
Private m_bInSet            As Boolean
Private m_aColumns()        As UcsColumnInfo
Private m_cHistory          As Collection
Private m_lDeleted          As Long
Private m_bRenderPending    As Boolean
Private m_bRenderData       As Boolean
Private m_bRenderStats      As Boolean
Private m_dblLastRender     As Double
Private m_dblLastTick       As Double
Private m_dblLastFlush      As Double
Private m_lCellCount        As Long
Private m_dblCellTime       As Double

Private Type UcsColumnInfo
    Field                   As String
    NumberFormat            As String
End Type

Private Enum UcsTreeImages
    ucsImgServers
    ucsImgServer
End Enum

Private Enum UcsMenuIndexes
    ucsMnuFileConnect = 0
    ucsMnuFileFilter = 1
    ucsMnuFileStats = 2
    ucsMnuFileExit = 4
    ucsMnuHelpAbout = 0
    ucsMnuMainPopup = 2
    ucsMnuPopupKill = 0
    ucsMnuMainTree = 3
    ucsMnuTreeConnect = 0
    ucsMnuTreeDisconnect = 2
    ucsMnuTreeProperties = 4
End Enum

'=========================================================================
' Error handling
'=========================================================================

Private Sub PrintError(sFunc As String)
    Debug.Print MODULE_NAME & "." & sFunc & ": " & Error
    MsgBox MODULE_NAME & "." & sFunc & "(" & Erl & ") : " & Error, vbCritical
End Sub

'=========================================================================
' Properties
'=========================================================================

Private Property Get pvSelectedRows() As Collection
    Const FUNC_NAME     As String = "pvSelectedRows [get]"
    Dim lRow            As Long
    Dim lIdx            As Long
    Dim sKey            As String
    Dim lCount          As Long
    Dim oServer         As ADODB.Field
    Dim oSpid           As ADODB.Field

    On Error GoTo EH
    Set pvSelectedRows = New Collection
    lRow = lvwMain.FocusedRow
    If lRow > 0 Then
        If pvMoveToRow(lRow) Then
            pvSelectedRows.Add pvGetRowKey(m_rsListSort!Server.Value, m_rsListSort!SPID.Value)
        End If
    End If
    If m_rsListSort Is Nothing Then
        Exit Property
    End If
    If lvwMain.SelectedCount > 0 And m_rsListSort.RecordCount > 0 Then
        '--- one pass in list order, positioning a row by number costs a scan of the rows before it
        Set oServer = m_rsListSort.Fields("Server")
        Set oSpid = m_rsListSort.Fields("SPID")
        lCount = lvwMain.RowCount
        m_rsListSort.MoveFirst
        Do While Not m_rsListSort.EOF And lIdx < lCount
            lIdx = lIdx + 1
            If lvwMain.RowSelected(lIdx) Then
                sKey = pvGetRowKey(oServer.Value, oSpid.Value)
                pvSelectedRows.Add sKey, sKey
            End If
            m_rsListSort.MoveNext
        Loop
    End If
    Exit Property
EH:
    PrintError FUNC_NAME
    Resume Next
End Property

Private Property Set pvSelectedRows(oValue As Collection)
    Const FUNC_NAME     As String = "pvSelectedRows [let]"
    Dim lFocus          As Long
    Dim lIdx            As Long
    Dim sKey            As String
    Dim bSelected       As Boolean
    Dim lCount          As Long
    Dim oServer         As ADODB.Field
    Dim oSpid           As ADODB.Field
    Dim sFocusKey       As String

    On Error GoTo EH
    If oValue.Count > 0 Then
        m_bInSet = True
        '--- one pass in list order, positioning a row by number costs a scan of the rows before it
        Set oServer = m_rsListSort.Fields("Server")
        Set oSpid = m_rsListSort.Fields("SPID")
        sFocusKey = oValue(1)
        lCount = lvwMain.RowCount
        If m_rsListSort.RecordCount > 0 Then
            m_rsListSort.MoveFirst
        End If
        For lIdx = 1 To lCount
            bSelected = False
            If Not m_rsListSort.EOF Then
                sKey = pvGetRowKey(oServer.Value, oSpid.Value)
                If sKey = sFocusKey Then
                    lFocus = lIdx
                End If
                bSelected = SearchCollection(oValue, sKey)
                m_rsListSort.MoveNext
            End If
            If lvwMain.RowSelected(lIdx) <> bSelected Then
                lvwMain.RowSelected(lIdx) = bSelected
            End If
        Next
        If lvwMain.FocusedRow <> lFocus Then
            lvwMain.FocusedRow = lFocus
        End If
        m_bInSet = False
    End If
    Exit Property
EH:
    PrintError FUNC_NAME
    Resume Next
End Property

'=========================================================================
' Methods
'=========================================================================

'--- results of one server's fetch, merged into the list every server shares
Friend Sub frShowResults(oServer As cServerMonitor, rs As Recordset)
    pvShowResults oServer, rs
End Sub

Private Sub pvRefreshUI()
    Const FUNC_NAME     As String = "pvRefreshUI"
    Dim dblStart        As Double

    On Error GoTo EH
    If m_rsListSort Is Nothing Then
        Exit Sub
    End If
    dblStart = TraceStart()
    lvwMain.Refresh
    TraceEnd "ui.list", dblStart
    dblStart = TraceStart()
    pvRefreshInput
    TraceEnd "ui.input", dblStart, vbTab & "len=" & Len(txtInput.Text)
    Exit Sub
EH:
    PrintError FUNC_NAME
    Resume Next
End Sub

Private Sub pvRefreshInput()
    Const FUNC_NAME     As String = "pvRefreshInput"
    Dim lRow            As Long
    Dim sInputBuffer    As String
    Dim vHistory        As Variant

    On Error GoTo EH
    If m_rsListSort Is Nothing Then
        Exit Sub
    End If
    lRow = lvwMain.FocusedRow
    If lRow > 0 Then
        If pvMoveToRow(lRow) Then
            If SearchCollection(m_cHistory, pvGetRowKey(m_rsListSort!Server.Value, m_rsListSort!SPID.Value), RetVal:=vHistory) Then
                sInputBuffer = vHistory
            End If
            If txtInput.Text <> sInputBuffer Then
                txtInput.Text = sInputBuffer
                txtInput.SelStart = Len(sInputBuffer)
            End If
        End If
    Else
        txtInput.Text = vbNullString
    End If
    Exit Sub
EH:
    PrintError FUNC_NAME
    Resume Next
End Sub

'--- parses and stores the filter text, False with a message when it is not valid
Private Function pvSetFilter(sFilter As String) As Boolean
    Dim oFilter         As cRowFilter

    On Error GoTo EH
    Set oFilter = New cRowFilter
    oFilter.Init sFilter
    If oFilter.Active Then
        Set m_oFilter = oFilter
    Else
        Set m_oFilter = Nothing
    End If
    m_sFilter = sFilter
    pvSetFilter = True
    Exit Function
EH:
    MsgBox "Invalid filter: " & Err.Description, vbExclamation
End Function

'--- ADO filters have no NOT and cannot AND groups of ORs, so rows are matched here
Private Sub pvApplyFilter()
    Dim lIter           As Long
    Dim dblStart        As Double

    If m_rsListSort Is Nothing Then
        Exit Sub
    End If
    dblStart = TraceStart()
    If m_oFilter Is Nothing Then
        m_rsListSort.Filter = vbNullString
        TraceEnd "filter.none", dblStart
        Exit Sub
    End If
    Do While MoveRecordset(m_rsList, lIter)
        m_rsList!IsMatch.Value = m_oFilter.Matches(Array( _
            LCase$(C_Str(m_rsList!Server.Value)), _
            LCase$(C_Str(m_rsList!Program.Value)), _
            LCase$(C_Str(m_rsList!DB.Value)), _
            LCase$(C_Str(m_rsList!Host.Value)), _
            LCase$(C_Str(m_rsList!Login.Value)), _
            LCase$(C_Str(m_rsList!Status.Value)), _
            LCase$(C_Str(m_rsList!Command.Value))))
    Loop
    m_rsListSort.Filter = "IsMatch = True"
    TraceEnd "filter.apply", dblStart, vbTab & "list=" & m_rsList.RecordCount
End Sub

Private Sub pvShowResults(oServer As cServerMonitor, rs As Recordset)
    Const FUNC_NAME     As String = "pvShowResults"
    Dim sServer         As String
    Dim bRefreshData    As Boolean
    Dim lIter           As Long
    Dim lIdx            As Long
    Dim oFld            As ADODB.Field
    Dim vValue          As Variant
    Dim cResult         As Collection
    Dim sKey            As String
    Dim bIsActive       As Boolean
    Dim bRefreshStats   As Boolean
    Dim dblStart        As Double
    Dim dblStep         As Double
    Dim lCount          As Long
    Dim aSource()       As ADODB.Field
    Dim aTarget()       As ADODB.Field
    Dim lStatusIdx      As Long
    Dim lCommandIdx     As Long
    Dim lWaitIdx        As Long
    Dim oSpid           As ADODB.Field
    Dim sValue          As String
    Dim cSeen           As Collection
    Dim lAdded          As Long
    Dim lWrites         As Long

    On Error GoTo EH
    If rs Is Nothing Then
        Exit Sub
    End If
    If rs.State <> adStateOpen Then
        Exit Sub
    End If
    If oServer.Mode = ucsMonExtEvents Then
        pvShowExtEvents oServer, rs
        Exit Sub
    End If
    dblStart = TraceStart()
    sServer = oServer.Server
    pvPrepareList
    dblStep = TraceStart()
    '--- sync rs, sp_who2 columns map to list fields by position up to REQUESTID
    '--- each column is looked up once, a Field follows its recordset's current row
    If rs.RecordCount > 0 Then
        For Each oFld In rs.Fields
            If oFld.Name = "REQUESTID" Then
                Exit For
            End If
            lCount = lCount + 1
        Next
    End If
    If lCount > 0 Then
        ReDim aSource(0 To lCount - 1) As ADODB.Field
        ReDim aTarget(0 To lCount - 1) As ADODB.Field
        lStatusIdx = -1
        lCommandIdx = -1
        lWaitIdx = -1
        For lIdx = 0 To lCount - 1
            Set aSource(lIdx) = rs.Fields(lIdx)
            Set aTarget(lIdx) = m_rsList.Fields(lIdx)
            Select Case aTarget(lIdx).Name
            Case "Status"
                lStatusIdx = lIdx
            Case "Command"
                lCommandIdx = lIdx
            Case "Wait"
                lWaitIdx = lIdx
            End Select
        Next
        Set oSpid = rs.Fields("SPID")
        rs.MoveFirst
    End If
    Set cSeen = New Collection
    Do While lCount > 0
        If rs.EOF Then
            Exit Do
        End If
        sKey = pvGetRowKey(sServer, Trim$(C_Str(oSpid.Value)))
        '--- rows the list does not show are skipped instead of added and deleted again on
        '--- every refresh, and of the rows of a parallel query only the first is the session
        If Not pvIsShownProcess(oServer, oSpid, aSource, lStatusIdx, lCommandIdx) Then
            GoTo NextRow
        End If
        If SearchCollection(cSeen, sKey) Then
            GoTo NextRow
        End If
        cSeen.Add True, sKey
        If Not SetBookmark(m_rsList, m_cList, sKey) Then
            m_rsList.AddNew
            m_rsList!Server.Value = sServer
            m_rsList!IsActive.Value = False
            m_rsList!LastActive.Value = False
            m_rsList!SPID.Value = Trim$(C_Str(oSpid.Value))
            RemoveCollection m_cList, sKey
            m_cList.Add m_rsList.Bookmark, sKey
            lAdded = lAdded + 1
        Else
            Debug.Assert C_Str(m_rsList!SPID.Value) = Trim$(C_Str(oSpid.Value))
        End If
        bIsActive = False
        For lIdx = 0 To lCount - 1
            sValue = Trim$(C_Str(aSource(lIdx).Value))
            If lIdx = lWaitIdx And sValue = "0" Then
                sValue = "."
            ElseIf lIdx = lCommandIdx And sValue = "AWAITING COMMAND" Then
                sValue = vbNullString
            End If
            '--- only changes are written and a Null becomes the empty string once
            vValue = aTarget(lIdx).Value
            If C_Str(vValue) <> sValue Then
                If Not IsNull(vValue) Then
                    bIsActive = True
                End If
                aTarget(lIdx).Value = sValue
                bRefreshData = True
                lWrites = lWrites + 1
            ElseIf IsNull(vValue) Then
                aTarget(lIdx).Value = sValue
                lWrites = lWrites + 1
            End If
        Next
        If m_rsList!IsActive.Value <> bIsActive Then
            m_rsList!IsActive.Value = bIsActive
        End If
NextRow:
        rs.MoveNext
    Loop
    TraceEnd "sp_who2.sync", dblStep, sServer & vbTab & "rows=" & rs.RecordCount & " shown=" & cSeen.Count & " added=" & lAdded & " writes=" & lWrites
    dblStep = TraceStart()
    If m_rsList.RecordCount <> 0 Then
        Set cResult = InitIndexCollection(rs, "SPID")
        lIter = 0
        Do While MoveRecordset(m_rsList, lIter)
            If C_Str(m_rsList!Server.Value) <> sServer Then
                GoTo LoopNext
            End If
            sKey = C_Str(m_rsList!Host.Value) & "#" & C_Str(m_rsList!Login.Value) & "#" & C_Str(m_rsList!DB.Value)
            If Not SetBookmark(m_rsStats, m_cStats, sKey) Then
                m_rsStats.AddNew Array("Host", "Login", "DB", "SPID", "Opers"), Array(C_Str(m_rsList!Host.Value), C_Str(m_rsList!Login.Value), C_Str(m_rsList!DB.Value), m_rsList!SPID.Value, 1)
                m_cStats.Add m_rsStats.Bookmark, sKey
            End If
            If (m_rsList!SPID.Value > 50 And LCase$(C_Str(m_rsList!Status.Value)) <> "background" And LCase$(C_Str(m_rsList!Command.Value)) <> "task manager") Or oServer.SystemProcesses Then
                If Not cResult Is Nothing Then
                    If Not SetBookmark(rs, cResult, "#" & C_Str(m_rsList!SPID.Value)) Then
                        pvDeleteRow pvGetRowKey(sServer, m_rsList!SPID.Value)
                        GoTo LoopNext
                    End If
                End If
                bIsActive = False
                If Trim$(LCase(C_Str(m_rsList!Command.Value))) <> "awaiting command" _
                        And Trim$(LCase(C_Str(m_rsList!Status.Value))) <> "sleeping" _
                        And Trim$(LCase(C_Str(m_rsList!Status.Value))) <> "background" _
                        And Trim$(LCase(C_Str(m_rsList!Status.Value))) <> "suspended" Then
                    bIsActive = True
                End If
                If bIsActive Then
                    If m_rsList!IsActive.Value <> bIsActive Then
                        m_rsList!IsActive.Value = bIsActive
                        bRefreshData = True
                    End If
                End If
                If m_rsList!IsActive.Value And Not m_rsList!LastActive.Value Then
                    m_rsStats!SPID.Value = m_rsList!SPID.Value
                    m_rsStats!Opers.Value = m_rsStats!Opers.Value + 1
                    bRefreshStats = True
                End If
                If m_rsList!LastActive.Value <> m_rsList!IsActive.Value Then
                    m_rsList!LastActive.Value = m_rsList!IsActive.Value
                End If
            Else
                pvDeleteRow pvGetRowKey(sServer, m_rsList!SPID.Value)
            End If
LoopNext:
        Loop
    End If
    TraceEnd "sp_who2.stats", dblStep, sServer & vbTab & "rows=" & rs.RecordCount & " list=" & m_rsList.RecordCount
    pvQueueRender bRefreshData, bRefreshStats
    TraceEnd "sp_who2.total", dblStart, sServer
    Exit Sub
EH:
    If MsgBox(Error & vbCrLf & vbCrLf & MSG_CONTINUE, vbQuestion Or vbYesNo, MODULE_NAME & "." & FUNC_NAME & "(" & Erl & ")") = vbYes Then
        Exit Sub
        Resume
    End If
    pvDisconnect oServer.Server
End Sub

'--- the sp_who2 rows the list shows, the same rule its second pass removes the others by
Private Function pvIsShownProcess( _
            oServer As cServerMonitor, _
            oSpid As ADODB.Field, _
            aSource() As ADODB.Field, _
            ByVal lStatusIdx As Long, _
            ByVal lCommandIdx As Long) As Boolean
    If oServer.SystemProcesses Then
        pvIsShownProcess = True
        Exit Function
    End If
    If C_Dbl(Trim$(C_Str(oSpid.Value))) <= 50 Then
        Exit Function
    End If
    If lStatusIdx >= 0 Then
        If LCase$(Trim$(C_Str(aSource(lStatusIdx).Value))) = "background" Then
            Exit Function
        End If
    End If
    If lCommandIdx >= 0 Then
        If LCase$(Trim$(C_Str(aSource(lCommandIdx).Value))) = "task manager" Then
            Exit Function
        End If
    End If
    pvIsShownProcess = True
End Function

'--- the list, its sorted clone and its index live for the whole session: the clone follows
'--- added, changed and deleted rows by itself. Only the deleted rows the recordset keeps
'--- are dropped now and then by a round trip through a property bag
Private Sub pvPrepareList()
    Dim sSort           As String
    Dim dblStart        As Double

    If m_rsList Is Nothing Then
        Set m_rsList = CreateRecordset( _
            "SPID", adDouble, _
            "Status", adVarWChar, 30, _
            "Login", adVarWChar, 128, _
            "Host", adVarWChar, 128, _
            "Blk", adVarWChar, 10, _
            "DB", adVarWChar, 128, _
            "Command", adVarWChar, 128, _
            "CPU", adInteger, _
            "Dsk", adInteger, _
            "Last_Batch", adVarWChar, 128, _
            "Program", adVarWChar, 120, _
            "SP2", adInteger, _
            "Wait", adVarWChar, 64, _
            "Trans", adInteger, _
            "IsActive", adBoolean, _
            "LastActive", adBoolean, _
            "LoginTime", adDate, _
            "Server", adVarWChar, 128, _
            "IsMatch", adBoolean)
        Set m_rsListSort = m_rsList.Clone
        m_rsListSort.Sort = "Server, DB, Login, Host, SPID"
        Set m_cList = New Collection
        Set m_cHistory = New Collection
        m_lDeleted = 0
        pvApplyFilter
        pvShowSortOrder
    ElseIf m_lDeleted >= LNG_COMPACT_DELETED Then
        sSort = m_rsListSort.Sort
        dblStart = TraceStart()
        With New PropertyBag
            .WriteProperty "rs", m_rsList
            Set m_rsList = .ReadProperty("rs")
        End With
        Set m_rsListSort = m_rsList.Clone
        m_rsListSort.Sort = sSort
        Set m_cList = InitIndexCollection(m_rsList, "Server", "SPID")
        m_lDeleted = 0
        pvApplyFilter
        TraceEnd "prepare.compact", dblStart, vbTab & "list=" & m_rsList.RecordCount
    End If
    If m_rsStats Is Nothing Then
        Set m_rsStats = CreateRecordset( _
            "Host", adVarWChar, 128, _
            "Login", adVarWChar, 128, _
            "DB", adVarWChar, 128, _
            "SPID", adDouble, _
            "Opers", adInteger)
        Set m_cStats = InitIndexCollection(m_rsStats, "Host", "Login", "DB")
    End If
End Sub

'--- the current row of the list leaves it together with its index entry and history
Private Sub pvDeleteRow(sKey As String)
    RemoveCollection m_cList, sKey
    RemoveCollection m_cHistory, sKey
    m_rsList.Delete
    m_lDeleted = m_lDeleted + 1
End Sub

'--- appends to a session's history, which keeps only its most recent part
Private Sub pvAppendHistory(sKey As String, sText As String)
    Dim vHistory        As Variant

    If SearchCollection(m_cHistory, sKey, RetVal:=vHistory) Then
        RemoveCollection m_cHistory, sKey
    End If
    m_cHistory.Add Right$(vHistory & sText, LNG_HISTORY_SIZE), sKey
End Sub

'--- results of every server end up in one repaint per render interval
Private Sub pvQueueRender(ByVal bRefreshData As Boolean, ByVal bRefreshStats As Boolean)
    m_bRenderPending = True
    m_bRenderData = m_bRenderData Or bRefreshData
    m_bRenderStats = m_bRenderStats Or bRefreshStats
End Sub

Private Sub pvRender()
    Dim dblStart        As Double
    Dim bRefreshData    As Boolean
    Dim bRefreshStats   As Boolean

    dblStart = TraceStart()
    bRefreshData = m_bRenderData
    bRefreshStats = m_bRenderStats
    m_bRenderPending = False
    m_bRenderData = False
    m_bRenderStats = False
    m_dblLastRender = TimerEx
    If Not m_rsListSort Is Nothing Then
        pvRefreshList bRefreshData, bRefreshStats
    End If
    TraceEnd "render", dblStart, vbTab & "data=" & -bRefreshData & " stats=" & -bRefreshStats
End Sub

'--- header arrows for the sort keys, all but the SPID tie-break
Private Sub pvShowSortOrder()
    Dim cOrder          As Collection
    Dim vSplit          As Variant
    Dim lIdx            As Long
    Dim sElem           As String
    Dim sField          As String
    Dim vOrder          As Variant

    Set cOrder = New Collection
    vSplit = Split(m_rsListSort.Sort, ",")
    For lIdx = 0 To UBound(vSplit)
        sElem = Trim$(vSplit(lIdx))
        sField = At(Split(sElem, " "), 0)
        If Not (sField = "SPID" And lIdx = UBound(vSplit) And lIdx > 0) And Not SearchCollection(cOrder, sField) Then
            cOrder.Add IIf(pvIsSortDesc(sElem), ucsSortDescending, ucsSortAscending), sField
        End If
    Next
    For lIdx = 0 To UBound(m_aColumns)
        If SearchCollection(cOrder, m_aColumns(lIdx).Field, RetVal:=vOrder) Then
            lvwMain.ColumnSortOrder(lIdx + 1) = vOrder
        Else
            lvwMain.ColumnSortOrder(lIdx + 1) = ucsSortNone
        End If
    Next
End Sub

Private Function pvIsSortDesc(sElem As String) As Boolean
    pvIsSortDesc = (UCase$(Right$(sElem, 5)) = " DESC")
End Function

Private Sub pvRefreshList(ByVal bRefreshData As Boolean, ByVal bRefreshStats As Boolean)
    Dim dblStart        As Double

    lvwMain.Redraw = False
    '--- rows added or changed since the list was prepared are matched again
    If Not m_oFilter Is Nothing Then
        pvApplyFilter
    End If
    If lvwMain.RowCount <> m_rsListSort.RecordCount Then
        dblStart = TraceStart()
        m_bInSet = True
        lvwMain.RowCount = m_rsListSort.RecordCount
        m_bInSet = False
        bRefreshData = True
        TraceEnd "refresh.rowcount", dblStart, vbTab & "rows=" & m_rsListSort.RecordCount
    End If
    If bRefreshData Then
        dblStart = TraceStart()
        Set pvSelectedRows = m_cSelected
        TraceEnd "refresh.selection", dblStart, vbTab & "selected=" & m_cSelected.Count & " rows=" & lvwMain.RowCount
        pvRefreshUI
    End If
    dblStart = TraceStart()
    lvwMain.Redraw = True
    TraceEnd "refresh.redraw", dblStart
    If bRefreshStats And Not m_oFrmStats Is Nothing Then
        dblStart = TraceStart()
        m_oFrmStats.frRefresh m_rsStats
        TraceEnd "refresh.stats", dblStart
    End If
End Sub

Private Sub pvShowExtEvents(oServer As cServerMonitor, rs As Recordset)
    Const FUNC_NAME     As String = "pvShowExtEvents"
    Dim sServer         As String
    Dim bFull           As Boolean
    Dim lLost           As Long
    Dim rsSnap          As Recordset
    Dim lIter           As Long
    Dim sKey            As String
    Dim sRowKey         As String
    Dim bRefreshData    As Boolean
    Dim cSeen           As Collection
    Dim vField          As Variant
    Dim rsLive          As Recordset
    Dim cLive           As Collection
    Dim rsEvents        As Recordset
    Dim cActive         As Collection
    Dim bIsActive       As Boolean
    Dim bRefreshStats   As Boolean
    Dim dblStart        As Double
    Dim dblStep         As Double
    Dim lEvents         As Long

    On Error GoTo EH
    dblStart = TraceStart()
    sServer = oServer.Server
    If Not oServer.ExtEvents.frReadInfo(rs, bFull, lLost) Then
        Exit Sub
    End If
    pvPrepareList
    dblStep = TraceStart()
    '--- 2. sessions snapshot, only a full one has every SPID so missing ones are gone
    Set rsSnap = rs.NextRecordset
    Set cSeen = New Collection
    Do While MoveRecordset(rsSnap, lIter)
        sKey = "#" & rsSnap!SPID.Value
        sRowKey = pvGetRowKey(sServer, rsSnap!SPID.Value)
        If Not SetBookmark(m_rsList, m_cList, sRowKey) Then
            m_rsList.AddNew
            m_rsList!Server.Value = sServer
            m_rsList!SPID.Value = rsSnap!SPID.Value
            m_rsList!IsActive.Value = False
            m_rsList!LastActive.Value = False
            RemoveCollection m_cList, sRowKey
            m_cList.Add m_rsList.Bookmark, sRowKey
            bRefreshData = True
        ElseIf C_Str(m_rsList!LoginTime.Value) <> C_Str(rsSnap!LoginTime.Value) Then
            '--- same SPID reused by a new session
            RemoveCollection m_cHistory, sRowKey
            bRefreshData = True
        End If
        For Each vField In Array("Status", "Login", "Host", "DB", "Program", "CPU", "Dsk", "Last_Batch", "LoginTime")
            If pvSetValue(m_rsList.Fields(vField), rsSnap.Fields(vField).Value) Then
                bRefreshData = True
            End If
        Next
        RemoveCollection cSeen, sKey
        cSeen.Add True, sKey
    Loop
    '--- 3. live requests and open transactions
    Set rsLive = rs.NextRecordset
    Set cLive = InitIndexCollection(rsLive, "SPID")
    '--- 4. new events go to the history of their session
    Set rsEvents = rs.NextRecordset
    Set cActive = New Collection
    lIter = 0
    Do While MoveRecordset(rsEvents, lIter)
        sKey = "#" & rsEvents!SPID.Value
        sRowKey = pvGetRowKey(sServer, rsEvents!SPID.Value)
        If SearchCollection(m_cList, sRowKey) Then
            pvAppendHistory sRowKey, oServer.ExtEvents.FormatEvent(rsEvents)
            RemoveCollection cActive, sKey
            cActive.Add True, sKey
            bRefreshData = True
        End If
        lEvents = lEvents + 1
    Loop
    lIter = 0
    Do While MoveRecordset(m_rsList, lIter)
        If C_Str(m_rsList!Server.Value) <> sServer Then
            GoTo LoopNext
        End If
        sKey = "#" & m_rsList!SPID.Value
        If bFull And Not SearchCollection(cSeen, sKey) Then
            pvDeleteRow pvGetRowKey(sServer, m_rsList!SPID.Value)
            bRefreshData = True
        Else
            If SetBookmark(rsLive, cLive, sKey) Then
                For Each vField In Array("Status", "Command", "Blk", "Wait", "Trans", "DB")
                    If pvSetValue(m_rsList.Fields(vField), rsLive.Fields(vField).Value) Then
                        bRefreshData = True
                    End If
                Next
            Else
                '--- no request so the session is idle
                For Each vField In Array("Command", "Blk", "Wait", "Trans")
                    If pvSetValue(m_rsList.Fields(vField), Null) Then
                        bRefreshData = True
                    End If
                Next
                Select Case LCase$(C_Str(m_rsList!Status.Value))
                Case "running", "runnable", "suspended"
                    m_rsList!Status.Value = "sleeping"
                    bRefreshData = True
                End Select
            End If
            bIsActive = SearchCollection(cActive, sKey)
            If bIsActive And lLost > 0 Then
                pvAppendHistory pvGetRowKey(sServer, m_rsList!SPID.Value), "-- " & lLost & " events lost" & vbCrLf
            End If
            Select Case LCase$(C_Str(m_rsList!Status.Value))
            Case "running", "runnable"
                bIsActive = True
            End Select
            If C_Lng(m_rsList!Trans.Value) > 0 Then
                bIsActive = True
            End If
            If m_rsList!IsActive.Value <> bIsActive Then
                m_rsList!IsActive.Value = bIsActive
                bRefreshData = True
            End If
            If pvUpdateStats() Then
                bRefreshStats = True
            End If
        End If
LoopNext:
    Loop
    TraceEnd "xe.merge", dblStep, sServer & vbTab & "snap=" & rsSnap.RecordCount & " live=" & rsLive.RecordCount & " events=" & lEvents & " full=" & -bFull & " list=" & m_rsList.RecordCount
    pvQueueRender bRefreshData, bRefreshStats
    TraceEnd "xe.total", dblStart, sServer
    Exit Sub
EH:
    PrintError FUNC_NAME
    Resume Next
End Sub

Private Function pvSetValue(oFld As ADODB.Field, vValue As Variant) As Boolean
    If C_Str(oFld.Value) <> C_Str(vValue) Or IsNull(oFld.Value) <> IsNull(vValue) Then
        oFld.Value = vValue
        pvSetValue = True
    End If
End Function

Private Function pvUpdateStats() As Boolean
    Dim sKey            As String

    sKey = C_Str(m_rsList!Host.Value) & "#" & C_Str(m_rsList!Login.Value) & "#" & C_Str(m_rsList!DB.Value)
    If Not SetBookmark(m_rsStats, m_cStats, sKey) Then
        m_rsStats.AddNew Array("Host", "Login", "DB", "SPID", "Opers"), Array(C_Str(m_rsList!Host.Value), C_Str(m_rsList!Login.Value), C_Str(m_rsList!DB.Value), m_rsList!SPID.Value, 0)
        m_cStats.Add m_rsStats.Bookmark, sKey
    End If
    If m_rsList!IsActive.Value And Not m_rsList!LastActive.Value Then
        m_rsStats!SPID.Value = m_rsList!SPID.Value
        m_rsStats!Opers.Value = m_rsStats!Opers.Value + 1
        pvUpdateStats = True
    End If
    If m_rsList!LastActive.Value <> m_rsList!IsActive.Value Then
        m_rsList!LastActive.Value = m_rsList!IsActive.Value
    End If
End Function

Private Function pvGetRowKey(ByVal sServer As String, ByVal vSpid As Variant) As String
    pvGetRowKey = "#" & sServer & "#" & C_Str(vSpid)
End Function

'--- the dialog opens with the profile of the given server, else with the last used one
Private Sub pvConnect(Optional ByVal sServer As String)
    Const FUNC_NAME     As String = "pvConnect"
    Dim oFrmConnect     As frmConnect
    Dim oCmd            As ADODB.Command
    Dim eMode           As UcsMonitorMode
    Dim lRefreshRate    As Long
    Dim bSystemProcesses As Boolean
    Dim bStatements     As Boolean
    Dim sConnectString  As String

    On Error GoTo EH
    Set oFrmConnect = New frmConnect
    If oFrmConnect.frInit(oCmd, eMode, lRefreshRate, bSystemProcesses, bStatements, sConnectString, sServer) Then
        pvAddServer sServer, oCmd, eMode, lRefreshRate, bSystemProcesses, bStatements, sConnectString
    End If
    Exit Sub
EH:
    PrintError FUNC_NAME
    Resume Next
End Sub

'--- a server of a saved profile, connected without the dialog
Private Sub pvConnectProfile(sServer As String)
    Const FUNC_NAME     As String = "pvConnectProfile"
    Dim oFrmConnect     As frmConnect
    Dim oCmd            As ADODB.Command
    Dim eMode           As UcsMonitorMode
    Dim lRefreshRate    As Long
    Dim bSystemProcesses As Boolean
    Dim bStatements     As Boolean
    Dim sConnectString  As String
    Dim dblStart        As Double

    On Error GoTo EH
    dblStart = TraceStart()
    Screen.MousePointer = vbHourglass
    Set oFrmConnect = New frmConnect
    If oFrmConnect.frConnectProfile(sServer, oCmd, eMode, lRefreshRate, bSystemProcesses, bStatements, sConnectString) Then
        pvAddServer sServer, oCmd, eMode, lRefreshRate, bSystemProcesses, bStatements, sConnectString
    End If
    Screen.MousePointer = vbDefault
    TraceEnd "connect.profile", dblStart, sServer
    Exit Sub
EH:
    Screen.MousePointer = vbDefault
    PrintError FUNC_NAME
    Resume Next
End Sub

Private Sub pvAddServer( _
            sServer As String, _
            oCmd As ADODB.Command, _
            ByVal eMode As UcsMonitorMode, _
            ByVal lRefreshRate As Long, _
            ByVal bSystemProcesses As Boolean, _
            ByVal bStatements As Boolean, _
            sConnectString As String)
    Const FUNC_NAME     As String = "pvAddServer"
    Dim oServer         As cServerMonitor
    Dim oOther          As cServerMonitor

    On Error GoTo EH
    Set oServer = New cServerMonitor
    If Not pvInitServer(oServer, oCmd, eMode, lRefreshRate, bSystemProcesses, bStatements, sConnectString, sServer) Then
        Exit Sub
    End If
    '--- an existing one is replaced as the dialog may have changed its settings, a connected one
    '--- is shut down after the new one is in, so a shared events session outlives the swap
    If SearchCollection(m_cServers, LCase$(oServer.Server), RetVal:=oOther) Then
        If oOther.Connected Then
            pvDisconnect oOther.Server
        End If
        m_cServers.Remove LCase$(oServer.Server)
    End If
    m_cServers.Add oServer, LCase$(oServer.Server)
    pvShowServers oServer.Server
    pvSetCaption Me
    tmrFetch.Enabled = True
    oServer.Tick
    Exit Sub
EH:
    PrintError FUNC_NAME
    Resume Next
End Sub

Private Function pvInitServer( _
            oServer As cServerMonitor, _
            oCmd As ADODB.Command, _
            ByVal eMode As UcsMonitorMode, _
            ByVal lRefreshRate As Long, _
            ByVal bSystemProcesses As Boolean, _
            ByVal bStatements As Boolean, _
            sConnectString As String, _
            sServer As String) As Boolean
    On Error GoTo EH
    oServer.Init oCmd, eMode, lRefreshRate, bSystemProcesses, bStatements, sConnectString, sServer, Me
    pvInitServer = True
    Exit Function
EH:
    MsgBox "Cannot monitor server: " & Error, vbExclamation
    oServer.Shutdown
End Function

'--- the server's rows leave the list, it stays in the tree to be reconnected
Private Sub pvDisconnect(sServer As String)
    Const FUNC_NAME     As String = "pvDisconnect"
    Dim oServer         As cServerMonitor
    Dim lIter           As Long

    On Error GoTo EH
    If Not SearchCollection(m_cServers, LCase$(sServer), RetVal:=oServer) Then
        Exit Sub
    End If
    oServer.Shutdown
    If Not m_rsList Is Nothing Then
        Do While MoveRecordset(m_rsList, lIter)
            If C_Str(m_rsList!Server.Value) = sServer Then
                pvDeleteRow pvGetRowKey(sServer, m_rsList!SPID.Value)
            End If
        Loop
        pvPrepareList
        pvRefreshList True, False
    End If
    If Not pvHasConnected() Then
        tmrFetch.Enabled = False
    End If
    pvShowServers sServer
    pvSetCaption Me
    Exit Sub
EH:
    PrintError FUNC_NAME
    Resume Next
End Sub

Private Sub pvReconnect(sServer As String)
    Const FUNC_NAME     As String = "pvReconnect"
    Dim oServer         As cServerMonitor

    On Error GoTo EH
    If Not SearchCollection(m_cServers, LCase$(sServer), RetVal:=oServer) Then
        Exit Sub
    End If
    If oServer.Connected Then
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    If Not pvReconnectServer(oServer) Then
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    Screen.MousePointer = vbDefault
    pvShowServers sServer
    pvSetCaption Me
    tmrFetch.Enabled = True
    oServer.Tick
    Exit Sub
EH:
    PrintError FUNC_NAME
    Resume Next
End Sub

Private Function pvReconnectServer(oServer As cServerMonitor) As Boolean
    On Error GoTo EH
    oServer.Reconnect Me
    pvReconnectServer = True
    Exit Function
EH:
    Screen.MousePointer = vbDefault
    MsgBox "Cannot reconnect to " & oServer.Server & ": " & Error, vbExclamation
    oServer.Shutdown
End Function

Private Function pvIsConnected(sServer As String) As Boolean
    Dim oServer         As cServerMonitor

    If SearchCollection(m_cServers, LCase$(sServer), RetVal:=oServer) Then
        pvIsConnected = oServer.Connected
    End If
End Function

Private Function pvHasConnected() As Boolean
    Dim oServer         As cServerMonitor

    For Each oServer In m_cServers
        If oServer.Connected Then
            pvHasConnected = True
            Exit Function
        End If
    Next
End Function

Private Sub pvDisconnectAll()
    Dim oServer         As cServerMonitor
    Dim dblStart        As Double

    dblStart = TraceStart()
    tmrFetch.Enabled = False
    For Each oServer In m_cServers
        oServer.Shutdown
    Next
    Set m_cServers = New Collection
    TraceEnd "disconnect.all", dblStart
End Sub

Private Sub pvShowServers(Optional Selected As String)
    Dim oFrmConnect     As frmConnect
    Dim cNames          As Collection
    Dim oServer         As cServerMonitor
    Dim hRoot           As LongPtr
    Dim vName           As Variant
    Dim hNode           As LongPtr
    Dim cModes          As Collection
    Dim vMode           As Variant
    Dim sText           As String

    Set oFrmConnect = New frmConnect
    Set cNames = oFrmConnect.frGetServers(cModes)
    '--- a monitored server shows the mode it runs in, the others the one of their profile
    For Each oServer In m_cServers
        If Not pvContainsText(cNames, oServer.Server) Then
            cNames.Add oServer.Server
        End If
        RemoveCollection cModes, LCase$(oServer.Server)
        cModes.Add oServer.Mode, LCase$(oServer.Server)
    Next
    '--- updated in place, rebuilding it inside a tree notification upsets the tree
    hRoot = tvwServers.GetRootNode()
    If hRoot = 0 Then
        hRoot = tvwServers.AddNode(0, vbNullString, STR_TREE_ROOT, Image:=ucsImgServers)
        tvwServers.SelectedNode = hRoot
    End If
    For Each vName In cNames
        If LenB(vName) <> 0 Then
            If Not SearchCollection(cModes, LCase$(vName), RetVal:=vMode) Then
                vMode = ucsMonExtEvents
            End If
            sText = vName & IIf(vMode = ucsMonSpWho2, " (sp_who2)", " (XE)")
            hNode = tvwServers.NodeByKey(vName)
            If hNode = 0 Then
                hNode = tvwServers.AddNode(hRoot, vName, sText, Image:=ucsImgServer)
            ElseIf tvwServers.NodeText(hNode) <> sText Then
                tvwServers.NodeText(hNode) = sText
            End If
            tvwServers.NodeBold(hNode) = pvIsConnected(C_Str(vName))
            If LCase$(vName) = LCase$(Selected) Then
                tvwServers.SelectedNode = hNode
            End If
        End If
    Next
    '--- by name here, the connect dialog keeps the last used first
    tvwServers.SortChildren hRoot
End Sub

Private Function pvContainsText(cItems As Collection, sText As String) As Boolean
    Dim vItem           As Variant

    For Each vItem In cItems
        If LCase$(vItem) = LCase$(sText) Then
            pvContainsText = True
            Exit Function
        End If
    Next
End Function

Private Sub pvInitTreeImages()
    Const FUNC_NAME     As String = "pvInitTreeImages"
    Dim vResId          As Variant
    Dim baData()        As Byte
    Dim hBitmap         As LongPtr

    On Error GoTo EH
    m_hTreeImages = ImageList_Create(LNG_ICON_SIZE, LNG_ICON_SIZE, ILC_COLOR32, 2, 2)
    If m_hTreeImages = 0 Then
        Exit Sub
    End If
    '--- in UcsTreeImages order
    For Each vResId In Array(LNG_RES_SERVERS, LNG_RES_SERVER)
        baData = LoadResData(vResId, STR_RES_PNG)
        hBitmap = DecodePngBitmap(baData)
        If hBitmap <> 0 Then
            Call ImageList_Add(m_hTreeImages, hBitmap, 0)
            Call DeleteObject(hBitmap)
        End If
    Next
    tvwServers.ImageList = m_hTreeImages
    Exit Sub
EH:
    PrintError FUNC_NAME
    Resume Next
End Sub

Private Sub pvSetCaption(oForm As VB.Form)
    Dim cNames          As Collection
    Dim oServer         As cServerMonitor

    Set cNames = New Collection
    For Each oServer In m_cServers
        If oServer.Connected Then
            cNames.Add oServer.Server
        End If
    Next
    oForm.Caption = IIf(LenB(m_sFilter) <> 0, m_sFilter & " - ", vbNullString) & STR_APP_NAME & _
        IIf(cNames.Count > 0, " - [" & ConcatCollection(cNames, ", ") & "]", vbNullString)
End Sub

'--- positions m_rsListSort on a list row, a no-op when already there as cells of one row repaint together
Private Function pvMoveToRow(ByVal lRow As Long) As Boolean
    If m_rsListSort Is Nothing Then
        Exit Function
    End If
    If lRow < 1 Or lRow > m_rsListSort.RecordCount Then
        Exit Function
    End If
    If m_rsListSort.AbsolutePosition = lRow Then
        '--- the current record may have been deleted since the last paint
        If (m_rsListSort.Status And adRecDeleted) = 0 Then
            pvMoveToRow = True
            Exit Function
        End If
    End If
    pvMoveToRow = SetAbsolutePosition(m_rsListSort, lRow)
End Function

Private Sub pvInitColumns()
    pvAddColumn "Server", "Server", 100
    pvAddColumn "SPID", "SPID", 48, NumberFormat:="#,##0"
    pvAddColumn "Login", "Login", 167
    pvAddColumn "Host", "Host", 83
    pvAddColumn "DB", "DB", 208
    pvAddColumn "Program", "Program", 417
    pvAddColumn "Status", "Status", 83
    pvAddColumn "Command", "Command", 83
    pvAddColumn "Blk", "Blk", 33, NumberFormat:="#,##0"
    pvAddColumn "Wait", "Wait", 125
    pvAddColumn "Trans", "Trans", 33, Align:=LVCFMT_CENTER, NumberFormat:="#,##0"
    pvAddColumn "CPU", "CPU", 62, Align:=LVCFMT_RIGHT, NumberFormat:="#,##0"
    pvAddColumn "Dsk", "Dsk", 62, Align:=LVCFMT_RIGHT, NumberFormat:="#,##0"
    pvAddColumn "Last_Batch", "LastBatch", 117
End Sub

Private Sub pvAddColumn( _
            sField As String, _
            sCaption As String, _
            ByVal lWidth As Long, _
            Optional ByVal Align As Long = LVCFMT_LEFT, _
            Optional NumberFormat As String)
    Dim lCount          As Long

    lCount = lvwMain.ColumnCount
    ReDim Preserve m_aColumns(0 To lCount) As UcsColumnInfo
    m_aColumns(lCount).Field = sField
    m_aColumns(lCount).NumberFormat = NumberFormat
    lvwMain.AddColumn sCaption, lWidth, Align:=Align
End Sub

'=========================================================================
' Control events
'=========================================================================

Private Sub Form_Resize()
    Dim dblLeft          As Double

    On Error Resume Next
    If WindowState <> vbMinimized Then
        tvwServers.Move 0, 0, m_dblTreeWidth, ScaleHeight
        dblLeft = tvwServers.Left + tvwServers.Width
        picTreeSplitter.Move dblLeft, 0, 60, ScaleHeight
        dblLeft = picTreeSplitter.Left + picTreeSplitter.Width
        '--- the list and history split what is right of the tree
        lvwMain.Move dblLeft, 0, (ScaleWidth - dblLeft) * m_dblRatio, ScaleHeight
        dblLeft = lvwMain.Left + lvwMain.Width
        picSplitter.Move dblLeft, 0, 60, ScaleHeight
        dblLeft = picSplitter.Left + picSplitter.Width
        txtInput.Move dblLeft, 0, ScaleWidth - dblLeft, ScaleHeight
    End If
End Sub

Private Sub lvwMain_ColumnClick(ByVal Col As Long)
    Const FUNC_NAME     As String = "lvwMain_ColumnClick"
    Dim sKey            As String
    Dim bCtrl           As Boolean
    Dim vSplit          As Variant
    Dim lIdx            As Long
    Dim sElem           As String
    Dim sField          As String
    Dim bDesc           As Boolean
    Dim sSort           As String

    On Error GoTo EH
    If m_rsListSort Is Nothing Then
        Exit Sub
    End If
    sKey = m_aColumns(Col - 1).Field
    bCtrl = ((GetShiftState() And vbCtrlMask) <> 0)
    '--- a new column sorts ascending, clicking the primary one flips it. Ctrl appends
    '--- the column to the keys or flips it where it is
    vSplit = Split(m_rsListSort.Sort, ",")
    For lIdx = 0 To UBound(vSplit)
        sElem = Trim$(vSplit(lIdx))
        sField = At(Split(sElem, " "), 0)
        If sField = sKey Then
            If lIdx = 0 Or bCtrl Then
                bDesc = Not pvIsSortDesc(sElem)
            End If
        ElseIf bCtrl And sField <> "SPID" Then
            sSort = sSort & sElem & ", "
        End If
    Next
    sSort = sSort & sKey & IIf(bDesc, " DESC", vbNullString)
    '--- SPID breaks ties so rows keep their places between refreshes
    If sKey <> "SPID" Then
        sSort = sSort & ", SPID"
    End If
    m_rsListSort.Sort = sSort
    pvShowSortOrder
    pvRefreshUI
    Exit Sub
EH:
    PrintError FUNC_NAME
    Resume Next
End Sub

Private Sub lvwMain_GetCellText(ByVal Row As Long, ByVal Col As Long, Text As String)
    Const FUNC_NAME     As String = "lvwMain_GetCellText"
    Dim vValue          As Variant
    Dim dblStart        As Double

    On Error GoTo EH
    dblStart = TraceStart()
    If Not pvMoveToRow(Row) Then
        Exit Sub
    End If
    With m_aColumns(Col - 1)
        vValue = m_rsListSort.Fields(.Field).Value
        If LenB(.NumberFormat) <> 0 And Not IsNull(vValue) Then
            Text = Format$(vValue, .NumberFormat)
        Else
            Text = C_Str(vValue)
        End If
    End With
    If TraceEnabled Then
        m_lCellCount = m_lCellCount + 1
        m_dblCellTime = m_dblCellTime + TimerEx - dblStart
    End If
    Exit Sub
EH:
    PrintError FUNC_NAME
    Resume Next
End Sub

Private Sub lvwMain_RowPrePaint(ByVal Row As Long, Color As OLE_COLOR, BackColor As OLE_COLOR, Bold As Boolean, Handled As Boolean)
    Const FUNC_NAME     As String = "lvwMain_RowPrePaint"

    On Error GoTo EH
    If pvMoveToRow(Row) Then
        If C_Bool(m_rsListSort!IsActive.Value) Then
            BackColor = CLR_ACTIVE
            Handled = True
        End If
    End If
    Exit Sub
EH:
    PrintError FUNC_NAME
    Resume Next
End Sub

Private Sub lvwMain_KeyDown(KeyCode As Integer, Shift As Integer)
    Const FUNC_NAME     As String = "lvwMain_KeyDown"

    On Error GoTo EH
    If Shift = vbCtrlMask And KeyCode = vbKeyC Then
        ClipCopy lvwMain, SelectedOnly:=True
    ElseIf Shift = vbCtrlMask And KeyCode = vbKeyA Then
        lvwMain.SelectAll
    End If
    Exit Sub
EH:
    PrintError FUNC_NAME
    Resume Next
End Sub

Private Sub lvwMain_RightClick(ByVal Row As Long)
    Const FUNC_NAME     As String = "lvwMain_RightClick"

    On Error GoTo EH
    If pvMoveToRow(Row) Then
        If C_Lng(m_rsListSort!SPID.Value) <> 0 Then
            PopupMenu mnuMain(ucsMnuMainPopup)
        End If
    End If
    Exit Sub
EH:
    PrintError FUNC_NAME
    Resume Next
End Sub

Private Sub lvwMain_SelectionChanged()
    Const FUNC_NAME     As String = "lvwMain_SelectionChanged"
    Dim dblStart        As Double

    On Error GoTo EH
    If Not m_bInSet Then
        dblStart = TraceStart()
        pvRefreshInput
        Set m_cSelected = pvSelectedRows
        TraceEnd "selection.changed", dblStart, vbTab & "selected=" & m_cSelected.Count
    End If
    Exit Sub
EH:
    PrintError FUNC_NAME
    Resume Next
End Sub

Private Sub m_oFrmStats_BeforeClose()
    Set m_oFrmStats = Nothing
End Sub

Private Sub mnuFile_Click(Index As Integer)
    Const FUNC_NAME     As String = "mnuFile_Click"
    Dim sFilter         As String

    On Error GoTo EH
    Select Case Index
    Case ucsMnuFileConnect
        pvConnect
    Case ucsMnuFileFilter
        sFilter = InputBox("Server, program, database, host, login, status or command. Use * for wildcards, " & _
            "AND, OR, NOT and brackets to combine, quotes for text with keywords", "Filter", m_sFilter)
        If StrPtr(sFilter) <> 0 Then
            If Not pvSetFilter(sFilter) Then
                Exit Sub
            End If
            pvSetCaption Me
            If Not m_rsListSort Is Nothing Then
                pvApplyFilter
                lvwMain.RowCount = m_rsListSort.RecordCount
                pvRefreshUI
            End If
        End If
    Case ucsMnuFileStats
        If m_oFrmStats Is Nothing Then
            Set m_oFrmStats = New frmStats
        End If
        m_oFrmStats.frInit m_rsStats, Me
    Case ucsMnuFileExit
        Unload Me
    End Select
    Exit Sub
EH:
    PrintError FUNC_NAME
    Resume Next
End Sub

Private Sub mnuHelp_Click(Index As Integer)
    Const FUNC_NAME     As String = "mnuHelp_Click"

    On Error GoTo EH
    Select Case Index
    Case ucsMnuHelpAbout
        MsgBox App.Title & " " & App.Major & "." & App.Minor & "." & App.Revision & vbCrLf & vbCrLf & _
            App.LegalCopyright & vbCrLf & vbCrLf & _
            "Contact: wqweto@gmail.com", vbExclamation, "About"
    End Select
    Exit Sub
EH:
    PrintError FUNC_NAME
    Resume Next
End Sub

Private Sub mnuPopup_Click(Index As Integer)
    Const FUNC_NAME     As String = "mnuPopup_Click"
    Dim lIdx            As Long
    Dim oServer         As cServerMonitor

    On Error GoTo EH
    Select Case Index
    Case ucsMnuPopupKill
        For lIdx = 1 To lvwMain.RowCount
            If lvwMain.RowSelected(lIdx) Then
                If pvMoveToRow(lIdx) Then
                    If SearchCollection(m_cServers, LCase$(C_Str(m_rsListSort!Server.Value)), RetVal:=oServer) Then
                        oServer.KillSession m_rsListSort!SPID.Value
                    End If
                End If
            End If
        Next
        lvwMain.SelectedRow = 0
    End Select
    Exit Sub
EH:
    If MsgBox(Error & vbCrLf & vbCrLf & MSG_CONTINUE, vbQuestion Or vbYesNo, MODULE_NAME & "." & FUNC_NAME & "(" & Erl & ")") = vbYes Then
        Exit Sub
        Resume
    End If
    Resume Next
End Sub

Private Sub mnuTree_Click(Index As Integer)
    Const FUNC_NAME     As String = "mnuTree_Click"

    On Error GoTo EH
    Select Case Index
    Case ucsMnuTreeConnect
        pvConnect
    Case ucsMnuTreeDisconnect
        pvDisconnect m_sMenuServer
    Case ucsMnuTreeProperties
        pvConnect m_sMenuServer
    End Select
    Exit Sub
EH:
    PrintError FUNC_NAME
    Resume Next
End Sub

Private Sub picSplitter_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    m_bDown = True
    m_dblDownX = X
End Sub

Private Sub picSplitter_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Const FUNC_NAME     As String = "picSplitter_MouseMove"

    On Error GoTo EH
    If Button = 0 Then
        m_bDown = False
    End If
    If m_bDown Then
        m_dblRatio = Limit((picSplitter.Left + X - m_dblDownX - lvwMain.Left) / (ScaleWidth - lvwMain.Left), 0.05, 0.95)
        Form_Resize
    End If
    Exit Sub
EH:
    PrintError FUNC_NAME
    Resume Next
End Sub

Private Sub picSplitter_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    m_bDown = False
End Sub

Private Sub picTreeSplitter_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    m_bDown = True
    m_dblDownX = X
End Sub

Private Sub picTreeSplitter_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Const FUNC_NAME     As String = "picTreeSplitter_MouseMove"

    On Error GoTo EH
    If Button = 0 Then
        m_bDown = False
    End If
    If m_bDown Then
        m_dblTreeWidth = Limit(picTreeSplitter.Left + X - m_dblDownX, DBL_TREE_MIN_WIDTH, ScaleWidth / 2)
        Form_Resize
    End If
    Exit Sub
EH:
    PrintError FUNC_NAME
    Resume Next
End Sub

Private Sub picTreeSplitter_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    m_bDown = False
End Sub

Private Sub tmrFetch_Timer()
    Const FUNC_NAME     As String = "tmrFetch_Timer"
    Dim oServer         As cServerMonitor
    Dim dblStart        As Double

    On Error GoTo EH
    dblStart = TraceStart()
    If TraceEnabled Then
        '--- a gap much over the interval is time the UI thread was busy elsewhere
        If m_dblLastTick <> 0 Then
            TraceLine "tick.gap", (dblStart - m_dblLastTick) * 1000000#
        End If
        m_dblLastTick = dblStart
        If m_lCellCount > 0 Then
            TraceLine "paint.cells", m_dblCellTime * 1000000#, CStr(m_lCellCount)
            m_lCellCount = 0
            m_dblCellTime = 0
        End If
        If dblStart - m_dblLastFlush > 1 Then
            TraceFlush
            m_dblLastFlush = dblStart
            TraceEnd "trace.flush", dblStart
        End If
    End If
    For Each oServer In m_cServers
        oServer.Tick
    Next
    If m_bRenderPending Then
        If TimerEx - m_dblLastRender >= DBL_RENDER_INTERVAL Then
            pvRender
        End If
    End If
    TraceEnd "tick", dblStart
    Exit Sub
EH:
    PrintError FUNC_NAME
    Resume Next
End Sub

Private Sub Form_Load()
    Const FUNC_NAME     As String = "Form_Load"

    On Error GoTo EH
    TraceInit
    Set m_cSelected = New Collection
    Set m_cServers = New Collection
    tmrFetch.Interval = LNG_TICK_INTERVAL
    Caption = STR_APP_NAME
    App.Title = STR_APP_NAME
    WindowState = GetSetting(STR_APP_NAME, STR_REG_COMMON, "WindowState", vbNormal)
    If Not pvSetFilter(GetSetting(STR_APP_NAME, STR_REG_COMMON, "Filter", vbNullString)) Then
        pvSetFilter vbNullString
    End If
    m_dblRatio = Limit(C_Dbl(GetSetting(STR_APP_NAME, STR_REG_COMMON, "Ratio", 0.75)), 0.05, 0.95)
    m_dblTreeWidth = Limit(C_Dbl(GetSetting(STR_APP_NAME, STR_REG_COMMON, "TreeWidth", DBL_TREE_WIDTH)), DBL_TREE_MIN_WIDTH, Screen.Width / 2)
    pvInitColumns
    InitGdiplus
    pvInitTreeImages
    pvShowServers
    txtInput.Font.Size = 8
    Call SendMessage(txtInput.hWnd, EM_SETTABSTOPS, 1, 16&)
    If txtInput.Font.Name <> "Consolas" Then
        txtInput.Font.Name = "Courier New"
    End If
    pvConnect
    Exit Sub
EH:
    PrintError FUNC_NAME
    Resume Next
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Const FUNC_NAME     As String = "Form_Unload"

    On Error GoTo EH
    pvDisconnectAll
    If WindowState <> vbMinimized Then
        Call SaveSetting(STR_APP_NAME, STR_REG_COMMON, "WindowState", WindowState)
    Else
        On Error Resume Next
        Call DeleteSetting(STR_APP_NAME, STR_REG_COMMON, "WindowState")
        On Error GoTo EH
    End If
    Call SaveSetting(STR_APP_NAME, STR_REG_COMMON, "Filter", m_sFilter)
    Call SaveSetting(STR_APP_NAME, STR_REG_COMMON, "Ratio", m_dblRatio)
    Call SaveSetting(STR_APP_NAME, STR_REG_COMMON, "TreeWidth", m_dblTreeWidth)
    If m_hTreeImages <> 0 Then
        tvwServers.ImageList = 0
        Call ImageList_Destroy(m_hTreeImages)
        m_hTreeImages = 0
    End If
    TerminateGdiplus
    TraceFlush
    Exit Sub
EH:
    PrintError FUNC_NAME
    Resume Next
End Sub

Private Sub tvwServers_BeforeCollapse(ByVal hItem As LongPtr, Cancel As Boolean)
    Const FUNC_NAME     As String = "tvwServers_BeforeCollapse"

    On Error GoTo EH
    '--- the root is static and must keep the servers in view
    Cancel = (hItem = tvwServers.GetRootNode())
    Exit Sub
EH:
    PrintError FUNC_NAME
    Resume Next
End Sub

Private Sub tvwServers_DblClick()
    Const FUNC_NAME     As String = "tvwServers_DblClick"
    Dim sServer         As String

    On Error GoTo EH
    sServer = tvwServers.NodeKey(tvwServers.SelectedNode)
    If LenB(sServer) = 0 Then
        Exit Sub
    End If
    If SearchCollection(m_cServers, LCase$(sServer)) Then
        pvReconnect sServer
    ElseIf Not pvIsConnected(sServer) Then
        pvConnectProfile sServer
    End If
    Exit Sub
EH:
    PrintError FUNC_NAME
    Resume Next
End Sub

Private Sub tvwServers_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Const FUNC_NAME     As String = "tvwServers_MouseDown"
    Dim hNode           As LongPtr

    On Error GoTo EH
    If Button <> vbRightButton Then
        Exit Sub
    End If
    hNode = tvwServers.HitTest(X, Y)
    If hNode = 0 Then
        Exit Sub
    End If
    tvwServers.SelectedNode = hNode
    m_sMenuServer = tvwServers.NodeKey(hNode)
    mnuTree(ucsMnuTreeDisconnect).Enabled = pvIsConnected(m_sMenuServer)
    mnuTree(ucsMnuTreeProperties).Enabled = (LenB(m_sMenuServer) <> 0)
    PopupMenu mnuMain(ucsMnuMainTree), DefaultMenu:=mnuTree(ucsMnuTreeConnect)
    Exit Sub
EH:
    PrintError FUNC_NAME
    Resume Next
End Sub

Private Sub txtInput_KeyDown(KeyCode As Integer, Shift As Integer)
    Const FUNC_NAME     As String = "txtInput_KeyDown"

    On Error GoTo EH
    If Shift = vbCtrlMask And KeyCode = vbKeyA Then
        Call SendMessage(txtInput.hWnd, EM_SETSEL, 0, ByVal -1&)
    End If
    Exit Sub
EH:
    PrintError FUNC_NAME
    Resume Next
End Sub

Private Sub txtInput_KeyPress(KeyAscii As Integer)
    '--- prevent beep
    If KeyAscii < 32 And KeyAscii <> 8 And KeyAscii <> 9 Then '-- 8 = backspace, 9 = tab
        KeyAscii = 0
    End If
End Sub

'=========================================================================
' Base class events
'=========================================================================

Private Sub Form_Initialize()
    InitCommonControlsVB
End Sub
