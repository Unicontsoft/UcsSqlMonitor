VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Ucs DB Monitor"
   ClientHeight    =   5988
   ClientLeft      =   192
   ClientTop       =   840
   ClientWidth     =   9300
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
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
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   924
      Width           =   96
   End
   Begin UcsSQLMonitor.ctxListView lvwMain 
      Height          =   4632
      Left            =   336
      TabIndex        =   0
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
      TabIndex        =   1
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
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const MODULE_NAME As String = "frmMain"

'=========================================================================
' API
'=========================================================================

Private Const EM_SETSEL                 As Long = &HB1
Private Const EM_SETTABSTOPS            As Long = &HCB

'=========================================================================
' Constants and member variables
'=========================================================================

Private Const STR_REG_COMMON        As String = "Common"
Private Const CLR_ACTIVE            As Long = &H80FF00
Private Const ERR_NO_MORE_RESULTS   As Long = &H40EC9
Private Const MSG_CONTINUE          As String = "Do you want to continue?"

Private m_oCmd              As ADODB.Command
Private WithEvents m_oConn  As ADODB.Connection
Attribute m_oConn.VB_VarHelpID = -1
Private WithEvents m_rsResult As Recordset
Attribute m_rsResult.VB_VarHelpID = -1
Private m_rsList            As Recordset
Private m_cList             As Collection
Private m_rsListSort        As Recordset
Private m_cFieldMap         As Collection
Private m_sServer           As String
Private m_sFilter           As String
Private m_bDown             As Boolean
Private m_dblDownX          As Double
Private m_dblRatio          As Double
Private m_bSystemProcesses  As Boolean
Private m_dblCurrentSPID    As Double
Private m_rsStats           As Recordset
Private m_cStats            As Collection
Private WithEvents m_oFrmStats As frmStats
Attribute m_oFrmStats.VB_VarHelpID = -1
Private m_cSelected         As Collection
Private m_bInSet            As Boolean
Private m_bDelayFetch       As Boolean
Private m_sPassword         As String
Private m_aColumns()        As UcsColumnInfo

Private Type UcsColumnInfo
    Field                   As String
    NumberFormat            As String
End Type

Private Enum UcsMenuIndexes
    ucsMnuFileConnect = 0
    ucsMnuFileFilter = 1
    ucsMnuFileStats = 2
    ucsMnuFileExit = 4
    ucsMnuHelpAbout = 0
    ucsMnuMainPopup = 2
    ucsMnuPopupKill = 0
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

Private Property Get pvSelectedSpids() As Collection
    Const FUNC_NAME     As String = "pvSelectedSpids [get]"
    Dim lRow            As Long
    Dim lIdx            As Long

    On Error GoTo EH
    Set pvSelectedSpids = New Collection
    lRow = lvwMain.FocusedRow
    If lRow > 0 Then
        If pvMoveToRow(lRow) Then
            pvSelectedSpids.Add m_rsListSort!SPID.Value
        End If
    End If
    If lvwMain.SelectedCount > 0 Then
        For lIdx = 1 To lvwMain.ItemCount
            If lvwMain.ItemSelected(lIdx) Then
                If pvMoveToRow(lIdx) Then
                    pvSelectedSpids.Add m_rsListSort!SPID.Value, "#" & m_rsListSort!SPID.Value
                End If
            End If
        Next
    End If
    Exit Property
EH:
    PrintError FUNC_NAME
    Resume Next
End Property

Private Property Set pvSelectedSpids(oValue As Collection)
    Const FUNC_NAME     As String = "pvSelectedSpids [let]"
    Dim lFocus          As Long
    Dim lIdx            As Long
    Dim bSelected       As Boolean

    On Error GoTo EH
    If oValue.Count > 0 Then
        m_bInSet = True
        If SearchRecordset(m_rsListSort, "SPID=" & oValue(1)) Then
            lFocus = m_rsListSort.AbsolutePosition
        End If
        For lIdx = 1 To lvwMain.ItemCount
            bSelected = False
            If pvMoveToRow(lIdx) Then
                bSelected = SearchCollection(oValue, "#" & m_rsListSort!SPID.Value)
            End If
            If lvwMain.ItemSelected(lIdx) <> bSelected Then
                lvwMain.ItemSelected(lIdx) = bSelected
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

Private Sub pvRefreshUI()
    Const FUNC_NAME     As String = "pvRefreshUI"

    On Error GoTo EH
    If m_rsListSort Is Nothing Then
        Exit Sub
    End If
    lvwMain.Refresh
    pvRefreshInput
    Exit Sub
EH:
    PrintError FUNC_NAME
    Resume Next
End Sub

Private Sub pvRefreshInput()
    Const FUNC_NAME     As String = "pvRefreshInput"
    Dim lRow            As Long
    Dim sInputBuffer    As String

    On Error GoTo EH
    If m_rsListSort Is Nothing Then
        Exit Sub
    End If
    lRow = lvwMain.FocusedRow
    If lRow > 0 Then
        If pvMoveToRow(lRow) Then
            If SearchCollection(m_rsListSort.Fields, "Input_Buffer2") And Not m_rsListSort.EOF Then
                sInputBuffer = C_Str(m_rsListSort!Input_Buffer2.Value)
                If txtInput.Text <> sInputBuffer Then
                    txtInput.Text = sInputBuffer
                    txtInput.SelStart = Len(sInputBuffer)
                End If
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

Private Function pvGetFilter() As String
    If LenB(m_sFilter) <> 0 Then
        pvGetFilter = "Program LIKE '" & Quote(m_sFilter) & "' OR DB LIKE '" & Quote(m_sFilter) & "' OR Host LIKE '" & Quote(m_sFilter) & "' OR Login LIKE '" & Quote(m_sFilter) & "' OR Status LIKE '" & Quote(m_sFilter) & "' OR Command LIKE '" & Quote(m_sFilter) & "'"
    Else
        pvGetFilter = vbNullString
    End If
End Function

Private Sub pvShowResults(rs As Recordset)
    Const FUNC_NAME     As String = "pvShowResults"
    Dim rsSess          As Recordset
    Dim rsSpids         As Recordset
    Dim oFld            As ADODB.Field
    Dim bRefreshData    As Boolean
    Dim bRefreshStats   As Boolean
    Dim bIsActive       As Boolean
    Dim lIdx            As Long
    Dim vElem           As Variant
    Dim sSort           As String
    Dim vValue          As Variant
    Dim sKey            As String
    Dim cResult         As Collection
    Dim lIter           As Long
    Dim lSessIter       As Long
    
    On Error GoTo EH
    If rs Is Nothing Then
        Exit Sub
    End If
    If rs.State <> adStateOpen Then
        Exit Sub
    End If
    On Error Resume Next
    Set rsSpids = rs.NextRecordset
    On Error GoTo EH
    '--- create list
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
            "Input_Buffer", adBSTR, _
            "Input_Buffer2", adBSTR, _
            "IsActive", adBoolean, _
            "LastActive", adBoolean)
        Set m_rsListSort = m_rsList.Clone
        m_rsListSort.Sort = "DB, Login, Host, SPID"
        m_rsListSort.Filter = pvGetFilter
        Set m_cFieldMap = New Collection
        If InStr(m_oCmd.CommandText, "sp_whoisactive") Then
            Const STR_WHOISACTIVE_MAP As String = "SPID;session_id|Status;status|Login;login_name|Host;host_name|Blk;blocking_session_id|DB;database_name|Command;percent_complete|CPU;CPU|Dsk;physical_reads|Last_Batch;start_time|Program;program_name|SP2;request_id|Wait;wait_info|Trans;open_tran_count|Input_Buffer;sql_text"
            For Each vElem In Split(STR_WHOISACTIVE_MAP, "|")
                vElem = Split(vElem, ";")
                m_cFieldMap.Add vElem(0), vElem(1)
            Next
        End If
    Else
        sSort = m_rsListSort.Sort
        With New PropertyBag
            .WriteProperty "rs", m_rsList
            Set m_rsList = .ReadProperty("rs")
        End With
        Set m_rsListSort = m_rsList.Clone
        m_rsListSort.Sort = sSort
        m_rsListSort.Filter = pvGetFilter
        Debug.Print "After sort", Timer
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
    If Not rsSpids Is Nothing And m_rsList.RecordCount > 0 Then
        Do While MoveRecordset(m_rsList, lIter)
            If m_rsList!IsActive.Value Or LCase$(C_Str(m_rsList!Status.Value)) <> "sleeping" Then
                If Not SearchRecordset(rs, "session_id=" & m_rsList!SPID.Value) Then
                    m_oCmd.Parameters("show_sleeping_spids").Value = 2
                    m_oCmd.Parameters("filter").Value = m_rsList!SPID.Value
                    Set rsSess = New Recordset
                    rsSess.CursorLocation = adUseClient
                    rsSess.Open m_oCmd, , adOpenStatic, adLockBatchOptimistic
                    If rsSess.RecordCount > 0 Then
                        Do While MoveRecordset(rsSess, lSessIter)
                            If SearchRecordset(m_rsList, "SPID=" & rsSess!session_id.Value) Then
                                If pvCopyRow(rsSess, m_rsList) Then
                                    bRefreshData = True
                                End If
                                m_rsList!IsActive.Value = False
                            End If
                        Loop
                    Else
                        m_rsList.Delete
                    End If
                End If
            Else
                m_rsList!IsActive.Value = False
            End If
        Loop
    End If
    Debug.Print "Before sync rs", Timer
    '--- sync rs
    Set m_cList = InitIndexCollection(m_rsList, "SPID")
    Do While MoveRecordset(rs, lIter)
        If SearchCollection(rs.Fields, "session_id") Then
            If Not SetBookmark(m_rsList, m_cList, "#" & Trim$(C_Str(rs!session_id.Value))) Then
                m_rsList.AddNew
                m_rsList!LastActive.Value = False
                m_rsList!SPID.Value = Trim$(C_Str(rs!session_id.Value))
                RemoveCollection m_cList, "#" & m_rsList!SPID.Value
                m_cList.Add m_rsList.Bookmark, "#" & m_rsList!SPID.Value
            End If
            If pvCopyRow(rs, m_rsList) Then
                bRefreshData = True
            End If
        Else
            If Not SetBookmark(m_rsList, m_cList, "#" & Trim$(C_Str(rs!SPID.Value))) Then
                m_rsList.AddNew
                m_rsList!IsActive.Value = False
                m_rsList!LastActive.Value = False
                m_rsList!SPID.Value = Trim$(C_Str(rs!SPID.Value))
                RemoveCollection m_cList, "#" & m_rsList!SPID.Value
                m_cList.Add m_rsList.Bookmark, "#" & m_rsList!SPID.Value
            Else
                Debug.Assert C_Str(m_rsList!SPID.Value) = Trim$(C_Str(rs!SPID.Value))
            End If
            If InStr(m_oCmd.CommandText, "sp_who2") > 0 Then
                lIdx = 0
                m_rsList!IsActive.Value = False
                For Each oFld In rs.Fields
                    If oFld.Name = "REQUESTID" Then
                        Exit For
                    End If
                    vValue = Trim$(C_Str(oFld.Value))
                    If m_rsList(lIdx).Name = "Wait" And Trim$(C_Str(vValue)) = "0" Then
                        vValue = "."
                    ElseIf m_rsList(lIdx).Name = "Command" And vValue = "AWAITING COMMAND" Then
                        vValue = vbNullString
                    End If
                    If C_Str(m_rsList(lIdx).Value) <> vValue Then
                        Debug.Assert m_rsList(lIdx).Name <> "SPID"
                        If Not IsNull(m_rsList(lIdx).Value) Then
                            m_rsList!IsActive.Value = True
                        End If
                        m_rsList(lIdx).Value = C_Str(vValue)
                        bRefreshData = True
                    Else
                        m_rsList(lIdx).Value = C_Str(m_rsList(lIdx).Value)
                    End If
                    lIdx = lIdx + 1
                Next
            Else
                If Trim$(rs!Program.Value) = "sp_who_3 Input Buffers" And Not m_bSystemProcesses Then
                    RemoveCollection m_cList, "#" & C_Str(rs!SPID.Value)
                    m_rsList.Delete
                ElseIf Trim$(rs!Program.Value) = "Ucs SQL Monitor" And Not m_bSystemProcesses Then
                    RemoveCollection m_cList, "#" & C_Str(rs!SPID.Value)
                    m_rsList.Delete
                Else
                    m_rsList!IsActive.Value = False
                    For Each oFld In rs.Fields
                        vValue = Trim$(C_Str(oFld.Value))
                        If oFld.Name = "Trans" And vValue = "." Then
                            vValue = Null
                        ElseIf oFld.Name = "Command" And vValue = "AWAITING COMMAND" Then
                            vValue = vbNullString
                        End If
                        If C_Str(m_rsList(oFld.Name).Value) <> C_Str(vValue) Then
                            If Not IsNull(m_rsList(oFld.Name).Value) Then
                                m_rsList!IsActive.Value = True
                            End If
                            If LCase(oFld.Name) = "input_buffer" And LenB(m_rsList!Input_Buffer.Value) <> 0 Then
                                m_rsList!Input_Buffer2.Value = Right(C_Str(m_rsList!Input_Buffer2.Value) & m_rsList!Input_Buffer.Value & _
                                    IIf(Right$(m_rsList!Input_Buffer.Value, 2) <> vbCrLf, vbCrLf, vbNullString) & "GO" & vbCrLf, 32000)
                            End If
                            m_rsList(oFld.Name).Value = vValue
                            bRefreshData = True
                        ElseIf m_rsList(oFld.Name).Type <> adInteger Then
                            m_rsList(oFld.Name).Value = C_Str(m_rsList(oFld.Name).Value)
                        End If
                    Next
                    If C_Lng(m_rsList!Trans.Value) > 0 Then
                        m_rsList!IsActive.Value = True
                    End If
                End If
            End If
        End If
    Loop
    Debug.Print "Before m_rsList", Timer
    If m_rsList.RecordCount <> 0 Then
        If Not SearchCollection(rs.Fields, "session_id") Then
            Set cResult = InitIndexCollection(rs, "SPID")
        End If
        lIter = 0
        Do While MoveRecordset(m_rsList, lIter)
            sKey = C_Str(m_rsList!Host.Value) & "#" & C_Str(m_rsList!Login.Value) & "#" & C_Str(m_rsList!DB.Value)
            If Not SetBookmark(m_rsStats, m_cStats, sKey) Then
                m_rsStats.AddNew Array("Host", "Login", "DB", "SPID", "Opers"), Array(C_Str(m_rsList!Host.Value), C_Str(m_rsList!Login.Value), C_Str(m_rsList!DB.Value), m_rsList!SPID.Value, 1)
                m_cStats.Add m_rsStats.Bookmark, sKey
            End If
            If (m_rsList!SPID.Value > 50 And LCase$(C_Str(m_rsList!Status.Value)) <> "background" And LCase$(C_Str(m_rsList!Command.Value)) <> "task manager") Or m_bSystemProcesses Then
                If Not cResult Is Nothing Then
                    If Not SetBookmark(rs, cResult, "#" & C_Str(m_rsList!SPID.Value)) Then
                        m_rsList.Delete
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
                m_rsList!LastActive.Value = m_rsList!IsActive.Value
            Else
                RemoveCollection m_cList, "#" & C_Str(m_rsList!SPID.Value)
                m_rsList.Delete
            End If
LoopNext:
        Loop
    End If
    lvwMain.Redraw = False
    If lvwMain.ItemCount <> m_rsListSort.RecordCount Then
        m_bInSet = True
        lvwMain.ItemCount = m_rsListSort.RecordCount
        m_bInSet = False
        bRefreshData = True
    End If
    If bRefreshData Then
        Set pvSelectedSpids = m_cSelected
        pvRefreshUI
    End If
    lvwMain.Redraw = True
    If bRefreshStats And Not m_oFrmStats Is Nothing Then
        m_oFrmStats.frRefresh m_rsStats
    End If
    Exit Sub
EH:
    If MsgBox(Error & vbCrLf & vbCrLf & MSG_CONTINUE, vbQuestion Or vbYesNo, MODULE_NAME & "." & FUNC_NAME & "(" & Erl & ")") = vbYes Then
        Exit Sub
        Resume
    End If
    tmrFetch.Enabled = False
End Sub

Private Function pvCopyRow(rs As Recordset, rsList As Recordset) As Boolean
    Const FUNC_NAME     As String = "pvCopyRow"
    Dim oFld            As ADODB.Field
    Dim sField          As String
    Dim vValue          As Variant
    
    On Error GoTo EH
    m_rsList!LastActive.Value = rsList!IsActive.Value
    rsList!IsActive.Value = False
    For Each oFld In rs.Fields
        If SearchCollection(m_cFieldMap, oFld.Name) Then
            sField = m_cFieldMap(oFld.Name)
            If IsNull(oFld.Value) Then
                vValue = Null
            Else
                Select Case rsList.Fields(sField).Type
                Case adInteger
                    If LCase(sField) = "trans" And C_Lng(oFld.Value) = 0 Then
                        vValue = Null
                    Else
                        vValue = C_Lng(oFld.Value)
                    End If
                Case adBoolean
                    vValue = C_Bool(oFld.Value)
                Case Else
                    vValue = C_Str(oFld.Value)
                End Select
            End If
            If C_Str(rsList.Fields(sField).Value) <> C_Str(vValue) Then
                If Not IsNull(vValue) Then
                    rsList!IsActive.Value = True
                End If
                rsList.Fields(sField).Value = vValue
                If LCase(sField) = "input_buffer" And LenB(rsList!Input_Buffer.Value) <> 0 Then
                    rsList!Input_Buffer2.Value = Right(C_Str(rsList!Input_Buffer2.Value) & rsList!Input_Buffer.Value & _
                        IIf(Right$(rsList!Input_Buffer.Value, 2) <> vbCrLf, vbCrLf, vbNullString) & "GO" & vbCrLf, 32000)
                End If
                pvCopyRow = True
            End If
        End If
    Next
    rsList!IsActive.Value = (LCase$(rsList!Status.Value) <> "sleeping" And LCase$(rsList!Status.Value) <> "background") Or C_Lng(rsList!Trans.Value) > 0
    Exit Function
EH:
    PrintError FUNC_NAME
    Resume Next
End Function

Private Sub pvSetCaption(oForm As VB.Form)
    oForm.Caption = IIf(LenB(m_sFilter) <> 0, m_sFilter & " - ", vbNullString) & STR_APP_NAME & " - [" & m_sServer & "]"
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
        pvMoveToRow = True
    Else
        pvMoveToRow = SetAbsolutePosition(m_rsListSort, lRow)
    End If
End Function

Private Sub pvInitColumns()
    pvAddColumn "SPID", "SPID", 40, NumberFormat:="#,##0"
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

Private Sub Form_Initialize()
    InitCommonControlsVB
End Sub

Private Sub Form_Resize()
    Dim dblLeft          As Double

    On Error Resume Next
    If WindowState <> vbMinimized Then
        lvwMain.Move 0, 0, ScaleWidth * m_dblRatio, ScaleHeight
        dblLeft = lvwMain.Left + lvwMain.Width
        picSplitter.Move dblLeft, 0, 60, ScaleHeight
        dblLeft = picSplitter.Left + picSplitter.Width
        txtInput.Move dblLeft, 0, ScaleWidth - dblLeft, ScaleHeight
    End If
End Sub

Private Sub lvwMain_ColumnClick(ByVal Col As Long)
    Const FUNC_NAME     As String = "lvwMain_ColumnClick"
    Dim sKey            As String
    Dim sPrev           As String
    Dim sDesc           As String

    On Error GoTo EH
    If m_rsListSort Is Nothing Then
        Exit Sub
    End If
    sKey = m_aColumns(Col - 1).Field
    If InStr(m_rsListSort.Sort, sKey & ",") Then
        sDesc = " DESC"
    End If
    If (GetShiftState() And vbCtrlMask) <> 0 Then
        sPrev = Replace(Replace(Replace(m_rsListSort.Sort, sKey & " DESC, ", vbNullString), sKey & ", ", vbNullString), ", SPID", vbNullString)
        If sPrev = "SPID" Then
            sPrev = vbNullString
        Else
            sPrev = sPrev & ", "
        End If
    End If
    m_rsListSort.Sort = sPrev & sKey & sDesc & ", SPID"
    pvRefreshUI
    Exit Sub
EH:
    PrintError FUNC_NAME
    Resume Next
End Sub

Private Sub lvwMain_GetItemText(ByVal Row As Long, ByVal Col As Long, Text As String)
    Const FUNC_NAME     As String = "lvwMain_GetItemText"
    Dim vValue          As Variant

    On Error GoTo EH
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
    Exit Sub
EH:
    PrintError FUNC_NAME
    Resume Next
End Sub

Private Sub lvwMain_ItemPrePaint(ByVal Row As Long, Color As OLE_COLOR, BackColor As OLE_COLOR, Bold As Boolean, Handled As Boolean)
    Const FUNC_NAME     As String = "lvwMain_ItemPrePaint"

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
            m_dblCurrentSPID = C_Lng(m_rsListSort!SPID.Value)
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

    On Error GoTo EH
    If Not m_bInSet Then
        pvRefreshInput
        Set m_cSelected = pvSelectedSpids
    End If
    Exit Sub
EH:
    PrintError FUNC_NAME
    Resume Next
End Sub

Private Sub m_oConn_Disconnect(adStatus As ADODB.EventStatusEnum, ByVal pConnection As ADODB.Connection)
    Debug.Print "m_oConn_Disconnect adStatus="; adStatus, Timer
End Sub

Private Sub m_oConn_ExecuteComplete(ByVal RecordsAffected As Long, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pCommand As ADODB.Command, ByVal pRecordset As ADODB.Recordset, ByVal pConnection As ADODB.Connection)
    Const FUNC_NAME     As String = "m_oConn_ExecuteComplete"
    
    On Error GoTo EH
    Debug.Print "m_oConn_ExecuteComplete m_bDelayFetch="; m_bDelayFetch, Timer
    If m_bDelayFetch Then
        tmrFetch.Enabled = False
        tmrFetch.Enabled = True
    End If
    Exit Sub
EH:
    PrintError FUNC_NAME
    Resume Next
End Sub

Private Sub m_oConn_InfoMessage(ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pConnection As ADODB.Connection)
    Const FUNC_NAME     As String = "m_oConn_InfoMessage"
    
    On Error GoTo EH
    If pError.Number <> ERR_NO_MORE_RESULTS Then
        If MsgBox(pError & vbCrLf & vbCrLf & MSG_CONTINUE, vbQuestion Or vbYesNo, MODULE_NAME & "." & FUNC_NAME & "(" & Erl & ")") = vbNo Then
            Exit Sub
        End If
        If Not m_bDelayFetch Then
            tmrFetch.Enabled = False
            tmrFetch.Enabled = True
        End If
    End If
    Exit Sub
EH:
    PrintError FUNC_NAME
    Resume Next
End Sub

Private Sub m_oFrmStats_BeforeClose()
    Set m_oFrmStats = Nothing
End Sub

Private Sub m_rsResult_FetchComplete(ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    Const FUNC_NAME     As String = "m_rsResult_FetchComplete"
    Dim dblTimer        As Double
    Dim rsResult        As Recordset
    Dim bTimerDelayed   As Boolean
    
    On Error GoTo EH
    dblTimer = Timer
    If Not pError Is Nothing Then
        Debug.Print "pError.Description=" & pError.Description
        Exit Sub
    End If
    Set rsResult = m_rsResult
    bTimerDelayed = m_bDelayFetch
    m_bDelayFetch = False
    pvShowResults rsResult
    If bTimerDelayed Then
        tmrFetch_Timer
    End If
    Debug.Print "m_rsResult_FetchComplete, Elapsed=" & Format$(Timer - dblTimer, "0.000"), Timer
    Exit Sub
EH:
    PrintError FUNC_NAME
    Resume Next
End Sub

Private Sub mnuFile_Click(Index As Integer)
    Const FUNC_NAME     As String = "mnuFile_Click"
    Dim oFrmConnect     As New frmConnect
    Dim sFilter         As String
    Dim lRefreshRate    As Long
    
    On Error GoTo EH
    Select Case Index
    Case ucsMnuFileConnect
        tmrFetch.Enabled = False
        m_bDelayFetch = False
        Set m_rsResult = Nothing
        Set m_oConn = Nothing
        If oFrmConnect.frInit(m_oCmd, lRefreshRate, m_bSystemProcesses, m_sPassword) Then
            Set m_oConn = m_oCmd.ActiveConnection
            With m_oConn.Execute("SELECT srvnetname FROM sysservers WHERE srvid = 0")
                If Not .EOF Then
                    m_sServer = Trim$(.Fields(0).Value)
                Else
                    m_sServer = m_oConn.Properties("Data Source").Value
                End If
            End With
            pvSetCaption Me
            Set m_rsList = Nothing
            Set m_rsListSort = Nothing
            Set m_rsStats = Nothing
            Set m_cStats = Nothing
            tmrFetch.Interval = Round(1000# / lRefreshRate)
            tmrFetch.Enabled = True
            tmrFetch_Timer
        Else
            Caption = STR_APP_NAME
        End If
        lvwMain.ItemCount = 0
    Case ucsMnuFileFilter
        sFilter = InputBox("Program Filter (use * for wildcards)", "Filter", m_sFilter)
        If StrPtr(sFilter) <> 0 Then
            If LenB(sFilter) <> 0 And InStr(sFilter, "*") = 0 And InStr(sFilter, "_") = 0 Then
                m_sFilter = "*" & sFilter & "*"
            Else
                m_sFilter = sFilter
            End If
            pvSetCaption Me
            If Not m_rsListSort Is Nothing Then
                m_rsListSort.Filter = pvGetFilter
                lvwMain.ItemCount = m_rsListSort.RecordCount
                pvRefreshUI
            End If
        End If
    Case ucsMnuFileStats
        If m_oFrmStats Is Nothing Then
            Set m_oFrmStats = New frmStats
        End If
        m_oFrmStats.frInit m_rsStats, Me
    Case ucsMnuFileExit
        tmrFetch.Enabled = False
        Set m_rsResult = Nothing
        Set m_oConn = Nothing
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
    
    On Error GoTo EH
    Select Case Index
    Case ucsMnuPopupKill
        For lIdx = 1 To lvwMain.ItemCount
            If lvwMain.ItemSelected(lIdx) Then
                If pvMoveToRow(lIdx) Then
                    m_oConn.Execute "KILL " & m_rsListSort!SPID.Value
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
        m_dblRatio = Limit((picSplitter.Left + X - m_dblDownX) / ScaleWidth, 0.05, 0.95)
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

Private Sub pvTestConn()
    On Error Resume Next
    If m_oConn.State = adStateOpen Then
        m_oConn.Execute "SELECT @@TRANCOUNT"
        If Err.Number <> 0 Then
            m_oConn.Close
        End If
    End If
    If m_oConn.State = adStateClosed Then
        m_oConn.Open m_oConn.ConnectionString, , m_sPassword
        Set m_oCmd.ActiveConnection = m_oConn
    End If
End Sub

Private Sub tmrFetch_Timer()
    Const FUNC_NAME     As String = "tmrFetch_Timer"
    Dim lState          As Long
  
    On Error GoTo EH
    If m_rsResult Is Nothing Then
        Set m_rsResult = New Recordset
        m_rsResult.CursorLocation = adUseClient
    End If
    On Error Resume Next
    lState = m_rsResult.State
    If Err.Number <> 0 Then
        Debug.Print "tmrFetch_Timer Error="; Error, Timer
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
    If SearchCollection(m_oCmd.Parameters, "show_sleeping_spids") Then
        m_oCmd.Parameters("show_sleeping_spids").Value = IIf(m_rsList Is Nothing, 2, 1)
    End If
    If SearchCollection(m_oCmd.Parameters, "filter") Then
        m_oCmd.Parameters("filter").Value = vbNullString
    End If
    On Error Resume Next
    m_rsResult.Open m_oCmd, , adOpenStatic, adLockBatchOptimistic, adAsyncExecute Or adAsyncFetch
    If Err.Number <> 0 Then
        Debug.Print "tmrFetch_Timer Error="; Error, Timer
        pvTestConn
    End If
    On Error GoTo EH
'    pvShowResults m_rsResult
    Exit Sub
EH:
    PrintError FUNC_NAME
    Resume Next
End Sub

Private Sub Form_Load()
    Const FUNC_NAME     As String = "Form_Load"
    
    On Error GoTo EH
    Set m_cSelected = New Collection
    Caption = STR_APP_NAME
    App.Title = STR_APP_NAME
    WindowState = GetSetting(STR_APP_NAME, STR_REG_COMMON, "WindowState", vbNormal)
    m_sFilter = GetSetting(STR_APP_NAME, STR_REG_COMMON, "Filter", vbNullString)
    m_dblRatio = Limit(C_Dbl(GetSetting(STR_APP_NAME, STR_REG_COMMON, "Ratio", 0.75)), 0.05, 0.95)
    pvInitColumns
    txtInput.Font.Size = 8
    Call SendMessage(txtInput.hWnd, EM_SETTABSTOPS, 1, 16&)
    If txtInput.Font.Name <> "Consolas" Then
        txtInput.Font.Name = "Courier New"
    End If
    mnuFile_Click ucsMnuFileConnect
    Exit Sub
EH:
    PrintError FUNC_NAME
    Resume Next
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Const FUNC_NAME     As String = "Form_Unload"
    
    On Error GoTo EH
    tmrFetch.Enabled = False
    If WindowState <> vbMinimized Then
        Call SaveSetting(STR_APP_NAME, STR_REG_COMMON, "WindowState", WindowState)
    Else
        On Error Resume Next
        Call DeleteSetting(STR_APP_NAME, STR_REG_COMMON, "WindowState")
        On Error GoTo EH
    End If
    Call SaveSetting(STR_APP_NAME, STR_REG_COMMON, "Filter", m_sFilter)
    Call SaveSetting(STR_APP_NAME, STR_REG_COMMON, "Ratio", m_dblRatio)
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

