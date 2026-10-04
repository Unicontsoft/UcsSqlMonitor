VERSION 5.00
Begin VB.Form frmStats
   Caption         =   "Statistics"
   ClientHeight    =   7980
   ClientLeft      =   108
   ClientTop       =   408
   ClientWidth     =   7896
   Icon            =   "frmStats.frx":0000
   LinkTopic       =   "frmStats"
   ScaleHeight     =   7980
   ScaleWidth      =   7896
   StartUpPosition =   3  'Windows Default
   Begin UcsSQLMonitor.ctxListView lvwStats
      Height          =   3288
      Left            =   84
      TabIndex        =   0
      Top             =   84
      Width           =   5304
      _ExtentX        =   9356
      _ExtentY        =   5800
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}
         Name            =   "Tahoma"
         Size            =   7.2
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
End
Attribute VB_Name = "frmStats"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
DefObj A-Z
Private Const MODULE_NAME As String = "frmStats"

'=========================================================================
' Public events
'=========================================================================

Public Event BeforeClose()

'=========================================================================
' Constants and member variables
'=========================================================================

Private Const STR_DEFAULT_SORT          As String = "Host"

Private m_rsStats                   As Recordset
Private m_aColumns()                As UcsColumnInfo

Private Type UcsColumnInfo
    Field                   As String
    NumberFormat            As String
End Type

'=========================================================================
' Error handling
'=========================================================================

Private Sub PrintError(sFunc As String)
    Debug.Print MODULE_NAME & "." & sFunc & ": " & Error
End Sub

'=========================================================================
' Methods
'=========================================================================

Friend Sub frInit(rs As Recordset, OwnerForm As Form)
    Const FUNC_NAME     As String = "frInit"

    On Error GoTo EH
    If lvwStats.ColumnCount = 0 Then
        pvAddColumn "Host", "Host", 167
        pvAddColumn "Login", "Login", 167
        pvAddColumn "DB", "DB", 167
        pvAddColumn "Opers", "Operations", 67, Align:=LVCFMT_RIGHT, NumberFormat:="#,#"
    End If
    pvSetRecordset rs
    Show vbModeless, OwnerForm
    Exit Sub
EH:
    PrintError FUNC_NAME
    Resume Next
End Sub

Friend Sub frRefresh(rs As Recordset)
    Const FUNC_NAME     As String = "frRefresh"

    On Error GoTo EH
    pvSetRecordset rs
    Exit Sub
EH:
    PrintError FUNC_NAME
    Resume Next
End Sub

Private Sub pvSetRecordset(rs As Recordset)
    Dim sSort           As String

    sSort = STR_DEFAULT_SORT
    If Not m_rsStats Is Nothing Then
        sSort = m_rsStats.Sort
    End If
    Set m_rsStats = Nothing
    If Not rs Is Nothing Then
        Set m_rsStats = rs.Clone
        m_rsStats.Sort = sSort
        lvwStats.RowCount = m_rsStats.RecordCount
    Else
        lvwStats.RowCount = 0
    End If
    lvwStats.Refresh
End Sub

Private Function pvMoveToRow(ByVal lRow As Long) As Boolean
    If m_rsStats Is Nothing Then
        Exit Function
    End If
    If lRow < 1 Or lRow > m_rsStats.RecordCount Then
        Exit Function
    End If
    If m_rsStats.AbsolutePosition = lRow Then
        pvMoveToRow = True
    Else
        pvMoveToRow = SetAbsolutePosition(m_rsStats, lRow)
    End If
End Function

Private Sub pvAddColumn( _
            sField As String, _
            sCaption As String, _
            ByVal lWidth As Long, _
            Optional ByVal Align As Long = LVCFMT_LEFT, _
            Optional NumberFormat As String)
    Dim lCount          As Long

    lCount = lvwStats.ColumnCount
    ReDim Preserve m_aColumns(0 To lCount) As UcsColumnInfo
    m_aColumns(lCount).Field = sField
    m_aColumns(lCount).NumberFormat = NumberFormat
    lvwStats.AddColumn sCaption, lWidth, Align:=Align
End Sub

'=========================================================================
' Control events
'=========================================================================

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    RaiseEvent BeforeClose
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    lvwStats.Move 0, 0, ScaleWidth, ScaleHeight
End Sub

Private Sub lvwStats_ColumnClick(ByVal Col As Long)
    Const FUNC_NAME     As String = "lvwStats_ColumnClick"
    Dim sKey            As String

    On Error GoTo EH
    If m_rsStats Is Nothing Then
        Exit Sub
    End If
    sKey = m_aColumns(Col - 1).Field
    If m_rsStats.Sort = sKey Then
        m_rsStats.Sort = sKey & " DESC"
    Else
        m_rsStats.Sort = sKey
    End If
    lvwStats.Refresh
    Exit Sub
EH:
    PrintError FUNC_NAME
    Resume Next
End Sub

Private Sub lvwStats_GetCellText(ByVal Row As Long, ByVal Col As Long, Text As String)
    Const FUNC_NAME     As String = "lvwStats_GetCellText"
    Dim vValue          As Variant

    On Error GoTo EH
    If Not pvMoveToRow(Row) Then
        Exit Sub
    End If
    With m_aColumns(Col - 1)
        vValue = m_rsStats.Fields(.Field).Value
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

Private Sub lvwStats_KeyDown(KeyCode As Integer, Shift As Integer)
    Const FUNC_NAME     As String = "lvwStats_KeyDown"

    On Error GoTo EH
    If Shift = vbCtrlMask And KeyCode = vbKeyC Then
        ClipCopy lvwStats
    ElseIf Shift = vbCtrlMask And KeyCode = vbKeyA Then
        lvwStats.SelectAll
    End If
    Exit Sub
EH:
    PrintError FUNC_NAME
    Resume Next
End Sub
