VERSION 5.00
Begin VB.Form frmConnect 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Connect"
   ClientHeight    =   4272
   ClientLeft      =   36
   ClientTop       =   336
   ClientWidth     =   4548
   Icon            =   "frmConnect.frx":0000
   LinkTopic       =   "frmConnect"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4272
   ScaleWidth      =   4548
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton optSpWhoIsActive 
      Caption         =   "sp_whoisactive"
      Height          =   276
      Left            =   1764
      TabIndex        =   16
      Top             =   2352
      Width           =   2448
   End
   Begin VB.CheckBox chkSystemProcesses 
      Caption         =   "Show system processes"
      Height          =   264
      Left            =   588
      TabIndex        =   7
      Top             =   3192
      Width           =   3876
   End
   Begin VB.ComboBox cobServer 
      Height          =   288
      Left            =   1764
      TabIndex        =   0
      Top             =   168
      Width           =   2616
   End
   Begin VB.ComboBox cobRerfesh 
      Height          =   288
      Left            =   1764
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   2772
      Width           =   2028
   End
   Begin VB.OptionButton optSpWho2 
      Caption         =   "sp_who2"
      Height          =   276
      Left            =   3024
      TabIndex        =   5
      Top             =   1932
      Width           =   1188
   End
   Begin VB.OptionButton optSpWho3 
      Caption         =   "sp_who_3"
      Height          =   276
      Left            =   1764
      TabIndex        =   4
      Top             =   1932
      Value           =   -1  'True
      Width           =   1188
   End
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   348
      Left            =   3108
      TabIndex        =   9
      Top             =   3696
      Width           =   1272
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   348
      Left            =   1764
      TabIndex        =   8
      Top             =   3696
      Width           =   1272
   End
   Begin VB.TextBox txtDB 
      Height          =   288
      Left            =   1764
      TabIndex        =   1
      Top             =   588
      Width           =   2616
   End
   Begin VB.TextBox txtUser 
      Height          =   288
      Left            =   1764
      TabIndex        =   2
      Top             =   1008
      Width           =   2028
   End
   Begin VB.TextBox txtPass 
      Height          =   288
      IMEMode         =   3  'DISABLE
      Left            =   1764
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   1428
      Width           =   2028
   End
   Begin VB.Label Label3 
      Caption         =   "Refresh rate:"
      Height          =   264
      Left            =   588
      TabIndex        =   15
      Top             =   2772
      Width           =   1104
   End
   Begin VB.Label Label2 
      Caption         =   "Type:"
      Height          =   264
      Left            =   588
      TabIndex        =   14
      Top             =   1932
      Width           =   1104
   End
   Begin VB.Label Label1 
      Caption         =   "SQL Server:"
      Height          =   264
      Left            =   588
      TabIndex        =   13
      Top             =   168
      Width           =   1104
   End
   Begin VB.Label labDB 
      Caption         =   "SQL DB:"
      Height          =   264
      Left            =   588
      TabIndex        =   12
      Top             =   588
      Width           =   1104
   End
   Begin VB.Label labUser 
      Caption         =   "User:"
      Height          =   264
      Left            =   588
      TabIndex        =   11
      Top             =   1008
      Width           =   1104
   End
   Begin VB.Label Label4 
      Caption         =   "Pass:"
      Height          =   264
      Left            =   588
      TabIndex        =   10
      Top             =   1428
      Width           =   1104
   End
End
Attribute VB_Name = "frmConnect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const MODULE_NAME As String = "frmConnect"

Private Declare Function UpdateWindow Lib "user32" (ByVal hWnd As Long) As Long

'=========================================================================
' Constants and member variables
'=========================================================================

Private Const STR_REG_COUNT     As String = "Count"
Private Const STR_REG_PROFILE   As String = "Profile"
Private Const STR_REG_CURRENT   As String = "Current"
Private Const STR_REG_CONNECT   As String = "Connect"
Private Const STR_REFRESH_RATES As String = "25|20|15|10|5|4|3|2|1"
Private Const STR_DELIM         As String = ""

Private m_bOk               As Boolean
Private m_oCmd              As ADODB.Command
Private m_lRefreshRate      As Long
Private m_cProfiles         As Collection

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

Private Property Get pvContents() As String
    pvContents = Join(Array( _
        cobServer.Text, _
        txtDB.Text, _
        txtUser.Text, _
        txtPass.Text, _
        -optSpWho3.Value, _
        -optSpWho2.Value, _
        cobRerfesh.ListIndex, _
        chkSystemProcesses.Value, _
        -optSpWhoIsActive.Value), STR_DELIM)
End Property

Private Property Let pvContents(sValue As String)
    Dim vSplit          As Variant
    
    vSplit = Split(sValue, STR_DELIM)
    On Error Resume Next
    cobServer.Text = vSplit(0)
    txtDB.Text = vSplit(1)
    txtUser.Text = vSplit(2)
    txtPass.Text = vSplit(3)
    optSpWho3.Value = vSplit(4)
    optSpWho2.Value = vSplit(5)
    cobRerfesh.ListIndex = -1
    cobRerfesh.ListIndex = vSplit(6)
    chkSystemProcesses.Value = vSplit(7)
    If cobRerfesh.ListIndex < 0 Then
        cobRerfesh.ListIndex = 4
    End If
    optSpWhoIsActive.Value = vSplit(8)
End Property

Private Property Get pvProfile(sServer As String) As String
    pvProfile = m_cProfiles(sServer)
End Property

Private Property Let pvProfile(sServer As String, sProfile As String)
    On Error Resume Next
    If LenB(sServer) <> 0 Then
        If SearchCollection(m_cProfiles, sServer) Then
            m_cProfiles.Remove sServer
        End If
        m_cProfiles.Add sProfile, sServer
    End If
End Property

Private Property Let pvDisabled(ByVal bValue As Boolean)
    Dim oCtl            As Object
    
    On Error Resume Next
    For Each oCtl In Controls
        Select Case LCase(TypeName(oCtl))
        Case "textbox"
            oCtl.Locked = bValue
            oCtl.BackColor = IIf(bValue, vbButtonFace, vbWindowBackground)
        Case "commandbutton", "optionbutton", "checkbox"
            oCtl.Enabled = Not bValue
        Case "combobox"
            oCtl.Locked = bValue
            oCtl.BackColor = IIf(bValue, vbButtonFace, vbWindowBackground)
        End Select
    Next
    Call UpdateWindow(hWnd)
End Property

'=========================================================================
' Methods
'=========================================================================

Friend Function frInit( _
            oCmd As ADODB.Command, _
            lRefreshRate As Long, _
            bSystemProcesses As Boolean, _
            sPassword As String) As Boolean
    Const FUNC_NAME     As String = "frInit"
    Dim vElem           As Variant
    Dim lIdx            As Long
    Dim sProfile        As String
    
    On Error GoTo EH
    '--- fill static combos
    For Each vElem In Split(STR_REFRESH_RATES, "|")
        cobRerfesh.AddItem vElem & " fps"
        cobRerfesh.ItemData(cobRerfesh.NewIndex) = vElem
    Next
    cobRerfesh.ListIndex = 4
    '--- load profiles
    Set m_cProfiles = New Collection
    For lIdx = 1 To C_Lng(GetSetting(STR_APP_NAME, STR_REG_CONNECT, STR_REG_COUNT, 0))
        sProfile = GetSetting(STR_APP_NAME, STR_REG_CONNECT, STR_REG_PROFILE & lIdx, vbNullString)
        pvProfile(pvGetServer(sProfile)) = sProfile
    Next
    cobServer.Text = GetSetting(STR_APP_NAME, STR_REG_CONNECT, STR_REG_CURRENT, vbNullString)
    For lIdx = m_cProfiles.Count To 1 Step -1
        cobServer.AddItem pvGetServer(m_cProfiles(lIdx))
        If cobServer.Text = cobServer.List(cobServer.NewIndex) Then
            cobServer.ListIndex = cobServer.NewIndex
        End If
    Next
    '--- show UI
    m_bOk = False
    Show vbModal
    If m_bOk Then
        '--- persist last profile w/ successful connection
        pvProfile(cobServer.Text) = pvContents
        '--- cleanup (save only last 100 profiles)
        Do While m_cProfiles.Count > 100
            m_cProfiles.Remove 1
        Loop
        '--- save profiles
        Call SaveSetting(STR_APP_NAME, STR_REG_CONNECT, STR_REG_COUNT, m_cProfiles.Count)
        For lIdx = 1 To m_cProfiles.Count
            Call SaveSetting(STR_APP_NAME, STR_REG_CONNECT, STR_REG_PROFILE & lIdx, m_cProfiles(lIdx))
        Next
        Call SaveSetting(STR_APP_NAME, STR_REG_CONNECT, STR_REG_CURRENT, cobServer.Text)
        Set oCmd = m_oCmd
        lRefreshRate = m_lRefreshRate
        bSystemProcesses = (chkSystemProcesses.Value = vbChecked)
        sPassword = txtPass.Text
        '--- success
        frInit = True
    End If
    Unload Me
    Exit Function
EH:
    PrintError FUNC_NAME
    Resume Next
End Function

Private Sub pvInstallSpWho3(oConn As ADODB.Connection)
    Const FUNC_NAME     As String = "pvInstallSpWho3"
    Dim vElem           As Variant
    Dim sSQL            As String
    Dim lVersion        As Long
    
    On Error GoTo EH
    lVersion = oConn.Execute("SELECT @@microsoftversion / POWER(2, 24)").Fields(0).Value
    If lVersion > 8 Then
        sSQL = "exec dbo.sp_configure 'show advanced options', 1" & vbCrLf & _
               "RECONFIGURE" & vbCrLf & _
               "exec dbo.sp_configure 'OLE Automation Procedures', 1" & vbCrLf & _
               "exec dbo.sp_configure 'Ad Hoc Distributed Queries', 1" & vbCrLf & _
               "RECONFIGURE"
        oConn.Execute sSQL
    End If
    For Each vElem In Split(StrConv(LoadResData(101, "CUSTOM"), vbUnicode), vbCrLf)
        If LCase(vElem) = "go" Then
            oConn.Execute sSQL
            sSQL = vbNullString
        Else
            If Len(sSQL) Then
                sSQL = sSQL & vbCrLf
            End If
            If lVersion > 8 Then
                vElem = Replace(Replace(Replace(vElem, _
                            "system_function_schema", "dbo"), _
                            "FROM OPENROWSET(fngetsql, @sql_handle)", "FROM sys.dm_exec_sql_text(@sql_handle)"), _
                            "ELSE fn_view_input_buffer", "ELSE dbo.fn_view_input_buffer")
            End If
            sSQL = sSQL & vElem
        End If
    Next
    Exit Sub
EH:
    PrintError FUNC_NAME
    Resume Next
End Sub

Private Function pvGetServer(sProfile As String) As String
    On Error Resume Next
    pvGetServer = Split(sProfile, STR_DELIM)(0)
End Function

'=========================================================================
' Control events
'=========================================================================

Private Sub cobServer_Click()
    Const FUNC_NAME     As String = "cobServer_Click"
    
    On Error GoTo EH
    If cobServer.ListIndex >= 0 Then
        pvContents = pvProfile(cobServer.Text)
    End If
    Exit Sub
EH:
    PrintError FUNC_NAME
    Resume Next
End Sub

Private Sub Command1_Click()
    Dim oConn           As ADODB.Connection
    
    On Error GoTo EH
    pvDisabled = True
    Set oConn = New Connection
    oConn.ConnectionTimeout = 5
    oConn.Open "Provider=SQLOLEDB;Data Source=" & cobServer.Text & ";" & IIf(LenB(txtDB) <> 0, "Initial catalog=" & txtDB & ";", "") & IIf(LenB(txtUser) <> 0, "User ID=" & txtUser & ";Password=" & txtPass & ";", "Integrated security=SSPI;") & "Application Name=" & App.Title
    Set m_oCmd = New ADODB.Command
    Set m_oCmd.ActiveConnection = oConn
    m_oCmd.CommandTimeout = 5
    If optSpWhoIsActive.Value Then
        m_oCmd.CommandText = "exec sp_whoisactive @show_sleeping_spids=?, @filter=?" & IIf(chkSystemProcesses.Value = vbChecked, ", @show_system_spids=1", vbNullString) & vbCrLf & _
                             "SELECT spid FROM master..sysprocesses"
        m_oCmd.Parameters.Append m_oCmd.CreateParameter("show_sleeping_spids", adInteger)
        m_oCmd.Parameters.Append m_oCmd.CreateParameter("filter", adVarWChar, Size:=128)
    ElseIf optSpWho3.Value Then
        m_oCmd.CommandText = "exec sp_who_3 'input'" ' active
        If oConn.Execute("SELECT * FROM master.INFORMATION_SCHEMA.ROUTINES WHERE ROUTINE_NAME = 'sp_who_3'").EOF Then
            If MsgBox("sp_who_3 stored procedure not found." & vbCrLf & vbCrLf & "Do you want to install?", vbQuestion + vbYesNo) = vbYes Then
                pvInstallSpWho3 oConn
            Else
                '--- revert to sp_who2
                optSpWho2.Value = True
                m_oCmd.CommandText = "exec sp_who2"
            End If
        End If
    Else
        m_oCmd.CommandText = "exec sp_who2"
    End If
    m_lRefreshRate = cobRerfesh.ItemData(cobRerfesh.ListIndex)
    m_bOk = True
    Visible = False
QH:
    pvDisabled = False
    Exit Sub
EH:
    MsgBox Error, vbExclamation
    GoTo QH
End Sub

Private Sub Command2_Click()
    Visible = False
End Sub

