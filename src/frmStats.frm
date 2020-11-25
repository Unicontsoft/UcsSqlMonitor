VERSION 5.00
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmStats 
   Caption         =   "Statistics"
   ClientHeight    =   7980
   ClientLeft      =   108
   ClientTop       =   408
   ClientWidth     =   7896
   Icon            =   "frmStats.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7980
   ScaleWidth      =   7896
   StartUpPosition =   3  'Windows Default
   Begin GridEX20.GridEX geCtl 
      Height          =   3288
      Left            =   84
      TabIndex        =   0
      Top             =   84
      Width           =   5304
      _ExtentX        =   9356
      _ExtentY        =   5800
      Version         =   "2.0"
      AutomaticSort   =   -1  'True
      RecordNavigator =   -1  'True
      HoldSortSettings=   -1  'True
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      GridLineStyle   =   2
      GroupFooterStyle=   2
      MultiSelect     =   -1  'True
      HideSelection   =   1
      HeaderStyle     =   3
      MethodHoldFields=   -1  'True
      ContScroll      =   -1  'True
      AllowEdit       =   0   'False
      BorderStyle     =   2
      MaskColor       =   16711935
      RowHeaders      =   -1  'True
      HeaderFontName  =   "Tahoma"
      FontName        =   "Tahoma"
      FontSize        =   7.2
      ColumnHeaderHeight=   264
      IntProp1        =   0
      IntProp2        =   0
      IntProp7        =   0
      ColumnsCount    =   4
      Column(1)       =   "frmStats.frx":000C
      Column(2)       =   "frmStats.frx":0148
      Column(3)       =   "frmStats.frx":0254
      Column(4)       =   "frmStats.frx":0418
      SortKeysCount   =   1
      SortKey(1)      =   "frmStats.frx":0530
      FormatStylesCount=   6
      FormatStyle(1)  =   "frmStats.frx":0598
      FormatStyle(2)  =   "frmStats.frx":06E4
      FormatStyle(3)  =   "frmStats.frx":0794
      FormatStyle(4)  =   "frmStats.frx":0848
      FormatStyle(5)  =   "frmStats.frx":0920
      FormatStyle(6)  =   "frmStats.frx":09D8
      ImageCount      =   0
      PrinterProperties=   "frmStats.frx":0AB8
   End
End
Attribute VB_Name = "frmStats"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Event BeforeClose()

Friend Sub frInit(rs As Recordset, OwnerForm As Form)
    geCtl.HoldSortSettings = True
    geCtl.HoldFields
    Set geCtl.ADORecordset = rs.Clone
    Show vbModeless, OwnerForm
End Sub

Friend Sub frRefresh(rs As Recordset)
    Dim lFirstItem      As Long
    Dim lRow            As Long
    
    lFirstItem = geCtl.FirstItem
    lRow = geCtl.Row
    geCtl.HoldFields
    Set geCtl.ADORecordset = rs.Clone
    geCtl.FirstItem = lFirstItem
    geCtl.Row = lRow
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    RaiseEvent BeforeClose
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    geCtl.Move 0, 0, ScaleWidth, ScaleHeight
End Sub

Private Sub geCtl_BeforeGroupChange(ByVal Group As GridEX20.JSGroup, ByVal ChangeOperation As GridEX20.jgexGroupChange, ByVal GroupPosition As Integer, ByVal Cancel As GridEX20.JSRetBoolean)
    On Error Resume Next
    geCtl.Columns(Group.ColIndex).Visible = Not (ChangeOperation = jgexGroupInsert)
End Sub

Private Sub geCtl_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = vbCtrlMask And KeyCode = vbKeyC Then
        ClipCopy geCtl
    End If
End Sub
