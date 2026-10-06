VERSION 5.00
Begin VB.Form frmConnect
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Connect"
   ClientHeight    =   4944
   ClientLeft      =   36
   ClientTop       =   336
   ClientWidth     =   4548
   Icon            =   "frmConnect.frx":0000
   LinkTopic       =   "frmConnect"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4944
   ScaleWidth      =   4548
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkSystemProcesses
      Caption         =   "Show system processes"
      Height          =   264
      Left            =   588
      TabIndex        =   7
      Top             =   3192
      Width           =   3876
   End
   Begin VB.CheckBox chkEncrypt
      Caption         =   "Encrypt connection"
      Height          =   264
      Left            =   588
      TabIndex        =   8
      Top             =   3528
      Width           =   3876
   End
   Begin VB.CheckBox chkStatements
      Caption         =   "Trace statements in procedures"
      Height          =   264
      Left            =   588
      TabIndex        =   9
      Top             =   3864
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
      Caption         =   "sp_who2 (SQL 2000)"
      Height          =   276
      Left            =   1764
      TabIndex        =   5
      Top             =   2352
      Width           =   2616
   End
   Begin VB.OptionButton optExtEvents
      Caption         =   "Extended Events (SQL 2019+)"
      Height          =   276
      Left            =   1764
      TabIndex        =   4
      Top             =   1932
      Value           =   -1  'True
      Width           =   2616
   End
   Begin VB.CommandButton Command2
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   348
      Left            =   3108
      TabIndex        =   11
      Top             =   4368
      Width           =   1272
   End
   Begin VB.CommandButton Command1
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   348
      Left            =   1764
      TabIndex        =   10
      Top             =   4368
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
      TabIndex        =   17
      Top             =   2772
      Width           =   1104
   End
   Begin VB.Label Label2
      Caption         =   "Type:"
      Height          =   264
      Left            =   588
      TabIndex        =   16
      Top             =   1932
      Width           =   1104
   End
   Begin VB.Label Label1
      Caption         =   "SQL Server:"
      Height          =   264
      Left            =   588
      TabIndex        =   15
      Top             =   168
      Width           =   1104
   End
   Begin VB.Label labDB
      Caption         =   "SQL DB:"
      Height          =   264
      Left            =   588
      TabIndex        =   14
      Top             =   588
      Width           =   1104
   End
   Begin VB.Label labUser
      Caption         =   "User:"
      Height          =   264
      Left            =   588
      TabIndex        =   13
      Top             =   1008
      Width           =   1104
   End
   Begin VB.Label Label4
      Caption         =   "Pass:"
      Height          =   264
      Left            =   588
      TabIndex        =   12
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
Private m_eMode             As UcsMonitorMode
Private m_oCmd              As ADODB.Command
Private m_lRefreshRate      As Long
Private m_cProfiles         As Collection
Private m_sConnectString    As String

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
        -optExtEvents.Value, _
        -optSpWho2.Value, _
        cobRerfesh.ListIndex, _
        chkSystemProcesses.Value, _
        chkEncrypt.Value, _
        chkStatements.Value), STR_DELIM)
End Property

Private Property Let pvContents(sValue As String)
    Dim vSplit          As Variant
    Dim lIdx            As Long

    vSplit = Split(sValue, STR_DELIM)
    cobServer.Text = At(vSplit, 0)
    txtDB.Text = At(vSplit, 1)
    txtUser.Text = At(vSplit, 2)
    txtPass.Text = At(vSplit, 3)
    optExtEvents.Value = C_Bool(At(vSplit, 4))
    optSpWho2.Value = C_Bool(At(vSplit, 5))
    lIdx = C_Lng(At(vSplit, 6, "4"))
    If lIdx < 0 Or lIdx >= cobRerfesh.ListCount Then
        lIdx = 4
    End If
    cobRerfesh.ListIndex = lIdx
    chkSystemProcesses.Value = C_Lng(At(vSplit, 7))
    chkEncrypt.Value = C_Lng(At(vSplit, 8))
    chkStatements.Value = C_Lng(At(vSplit, 9))
    '--- profiles saved with sp_whoisactive have neither option set
    If Not optSpWho2.Value Then
        optExtEvents.Value = True
    End If
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
            eMode As UcsMonitorMode, _
            lRefreshRate As Long, _
            bSystemProcesses As Boolean, _
            bStatements As Boolean, _
            sConnectString As String, _
            sServer As String) As Boolean
    Const FUNC_NAME     As String = "frInit"

    On Error GoTo EH
    pvLoadProfiles
    '--- a given server opens with its own profile
    If LenB(sServer) <> 0 Then
        cobServer.Text = sServer
        If SearchCollection(m_cProfiles, sServer) Then
            pvContents = pvProfile(sServer)
        End If
    End If
    '--- show UI
    m_bOk = False
    Show vbModal
    If m_bOk Then
        pvSaveProfiles
        Set oCmd = m_oCmd
        eMode = m_eMode
        lRefreshRate = m_lRefreshRate
        bSystemProcesses = (chkSystemProcesses.Value = vbChecked)
        bStatements = (chkStatements.Value = vbChecked)
        sConnectString = m_sConnectString
        sServer = cobServer.Text
        frInit = True
    End If
    Unload Me
    Exit Function
EH:
    PrintError FUNC_NAME
    Resume Next
End Function

'--- connects with the server's saved profile without showing the dialog
Friend Function frConnectProfile( _
            sServer As String, _
            oCmd As ADODB.Command, _
            eMode As UcsMonitorMode, _
            lRefreshRate As Long, _
            bSystemProcesses As Boolean, _
            bStatements As Boolean, _
            sConnectString As String) As Boolean
    Const FUNC_NAME     As String = "frConnectProfile"

    On Error GoTo EH
    pvLoadProfiles
    If SearchCollection(m_cProfiles, sServer) Then
        pvContents = pvProfile(sServer)
        If pvOpen() Then
            pvSaveProfiles
            Set oCmd = m_oCmd
            eMode = m_eMode
            lRefreshRate = m_lRefreshRate
            bSystemProcesses = (chkSystemProcesses.Value = vbChecked)
            bStatements = (chkStatements.Value = vbChecked)
            sConnectString = m_sConnectString
            frConnectProfile = True
        End If
    End If
    Unload Me
    Exit Function
EH:
    PrintError FUNC_NAME
    Resume Next
End Function

'--- the servers of the saved profiles, the last used first
Friend Function frGetServers() As Collection
    Dim lIdx            As Long

    Set frGetServers = New Collection
    For lIdx = C_Lng(GetSetting(STR_APP_NAME, STR_REG_CONNECT, STR_REG_COUNT, 0)) To 1 Step -1
        frGetServers.Add pvGetServer(GetSetting(STR_APP_NAME, STR_REG_CONNECT, STR_REG_PROFILE & lIdx, vbNullString))
    Next
    Unload Me
End Function

Private Sub pvLoadProfiles()
    Dim vElem           As Variant
    Dim lIdx            As Long
    Dim sProfile        As String

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
End Sub

'--- persists the profile of a successful connection, keeping only the last 100
Private Sub pvSaveProfiles()
    Dim lIdx            As Long

    pvProfile(cobServer.Text) = pvContents
    Do While m_cProfiles.Count > 100
        m_cProfiles.Remove 1
    Loop
    Call SaveSetting(STR_APP_NAME, STR_REG_CONNECT, STR_REG_COUNT, m_cProfiles.Count)
    For lIdx = 1 To m_cProfiles.Count
        Call SaveSetting(STR_APP_NAME, STR_REG_CONNECT, STR_REG_PROFILE & lIdx, m_cProfiles(lIdx))
    Next
    Call SaveSetting(STR_APP_NAME, STR_REG_CONNECT, STR_REG_CURRENT, cobServer.Text)
End Sub


Private Function pvOpen() As Boolean
    Const MIN_VERSION_EXT_EVENTS As Long = 15
    Dim oConn           As ADODB.Connection
    Dim sDriver         As String
    Dim dblStart        As Double

    On Error GoTo EH
    dblStart = TraceStart()
    Set oConn = New Connection
    oConn.ConnectionTimeout = 5
    '--- ODBC drivers cannot talk to SQL 2000 so sp_who2 mode keeps SQLOLEDB
    If optExtEvents.Value Then
        sDriver = pvGetOdbcDriver()
    End If
    m_sConnectString = pvGetConnectString(sDriver)
    oConn.Open m_sConnectString
    TraceEnd "connect.open", dblStart, cobServer.Text
    dblStart = TraceStart()
    Set m_oCmd = New ADODB.Command
    Set m_oCmd.ActiveConnection = oConn
    m_oCmd.CommandTimeout = 5
    If optExtEvents.Value Then
        If oConn.Execute("SELECT @@MICROSOFTVERSION / 0x1000000").Fields(0).Value < MIN_VERSION_EXT_EVENTS Then
            MsgBox "Extended Events mode needs SQL Server 2019 or later. Use sp_who2 for this server.", vbExclamation
            GoTo QH
        End If
        m_eMode = ucsMonExtEvents
    Else
        m_oCmd.CommandText = "exec sp_who2"
        m_eMode = ucsMonSpWho2
    End If
    m_lRefreshRate = cobRerfesh.ItemData(cobRerfesh.ListIndex)
    TraceEnd "connect.check", dblStart, cobServer.Text
    pvOpen = True
QH:
    Exit Function
EH:
    MsgBox cobServer.Text & ": " & Error, vbExclamation
    GoTo QH
End Function

Private Function pvGetServer(sProfile As String) As String
    pvGetServer = At(Split(sProfile, STR_DELIM), 0)
End Function

'--- newest installed "ODBC Driver NN for SQL Server", empty when there is none
Private Function pvGetOdbcDriver() As String
    Const STR_PREFIX    As String = "ODBC Driver "
    Const STR_SUFFIX    As String = " for SQL Server"
    Dim sBuffer         As String
    Dim nSize           As Integer
    Dim vElem           As Variant
    Dim lVersion        As Long
    Dim lBest           As Long

    sBuffer = String$(8192, 0)
    If SQLGetInstalledDrivers(StrPtr(sBuffer), Len(sBuffer), nSize) = 0 Then
        Exit Function
    End If
    For Each vElem In Split(Left$(sBuffer, nSize), vbNullChar)
        If Left$(vElem, Len(STR_PREFIX)) = STR_PREFIX And Right$(vElem, Len(STR_SUFFIX)) = STR_SUFFIX Then
            lVersion = C_Lng(Mid$(vElem, Len(STR_PREFIX) + 1, Len(vElem) - Len(STR_PREFIX) - Len(STR_SUFFIX)))
            If lVersion > lBest Then
                lBest = lVersion
                pvGetOdbcDriver = vElem
            End If
        End If
    Next
End Function

Private Function pvGetConnectString(sDriver As String) As String
    Dim bEncrypt        As Boolean

    bEncrypt = (chkEncrypt.Value = vbChecked)
    If LenB(sDriver) <> 0 Then
        pvGetConnectString = "Provider=MSDASQL;Driver=" & pvBraceValue(sDriver) & ";Server=" & pvBraceValue(cobServer.Text) & ";" & _
            IIf(LenB(txtDB.Text) <> 0, "Database=" & pvBraceValue(txtDB.Text) & ";", vbNullString) & _
            IIf(LenB(txtUser.Text) <> 0, "UID=" & pvBraceValue(txtUser.Text) & ";PWD=" & pvBraceValue(txtPass.Text) & ";", "Trusted_Connection=Yes;") & _
            IIf(bEncrypt, "Encrypt=Yes;TrustServerCertificate=Yes;", "Encrypt=No;") & _
            "APP=" & App.Title & ";OLE DB Services=-2"
    Else
        pvGetConnectString = "Provider=SQLOLEDB;Data Source=" & pvQuoteValue(cobServer.Text) & ";" & _
            IIf(LenB(txtDB.Text) <> 0, "Initial catalog=" & pvQuoteValue(txtDB.Text) & ";", vbNullString) & _
            IIf(LenB(txtUser.Text) <> 0, "User ID=" & pvQuoteValue(txtUser.Text) & ";Password=" & pvQuoteValue(txtPass.Text) & ";", "Integrated security=SSPI;") & _
            IIf(bEncrypt, "Use Encryption for Data=True;", vbNullString) & _
            "Application Name=" & App.Title & ";OLE DB Services=-2"
    End If
End Function

'--- ODBC value in braces with } inside doubled, MSDASQL passes it through as is
Private Function pvBraceValue(sValue As String) As String
    pvBraceValue = "{" & Replace(sValue, "}", "}}") & "}"
End Function

'--- OLE DB value in double quotes when ; quotes or edge spaces would break it, " inside doubled
Private Function pvQuoteValue(sValue As String) As String
    If InStr(sValue, ";") > 0 Or InStr(sValue, """") > 0 Or InStr(sValue, "'") > 0 _
            Or Trim$(sValue) <> sValue Then
        pvQuoteValue = """" & Replace(sValue, """", """""") & """"
    Else
        pvQuoteValue = sValue
    End If
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
    pvDisabled = True
    If pvOpen() Then
        m_bOk = True
        Visible = False
    End If
    pvDisabled = False
End Sub

Private Sub Command2_Click()
    Visible = False
End Sub
