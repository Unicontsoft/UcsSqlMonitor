VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cRowFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'=========================================================================
'
'   Ucs SQL Monitor
'   Copyright (c) 2026 Unicontsoft
'
'   A filter expression of search terms combined with AND, OR, NOT and
'   brackets. Adjacent words form one term, quotes keep keywords as text.
'   A term matches when any of the values is like it, * is a wildcard and
'   a term without one matches anywhere in the value
'
'=========================================================================
Option Explicit
DefObj A-Z
Private Const MODULE_NAME As String = "cRowFilter"

'=========================================================================
' Constants and member variables
'=========================================================================

Private Enum UcsTokenKind
    ucsTokTerm
    ucsTokAnd
    ucsTokOr
    ucsTokNot
    ucsTokOpen
    ucsTokClose
    ucsTokEnd
End Enum

Private m_aText()                   As String
Private m_aKind()                   As UcsTokenKind
Private m_lCount                    As Long
Private m_lPos                      As Long
Private m_vRoot                     As Variant

'=========================================================================
' Error handling
'=========================================================================

Private Sub RaiseError(sFunc As String, sMessage As String)
    Err.Raise vbObjectError, MODULE_NAME & "." & sFunc, sMessage
End Sub

'=========================================================================
' Properties
'=========================================================================

'--- False for an empty expression, which matches everything
Public Property Get Active() As Boolean
    Active = Not IsEmpty(m_vRoot)
End Property

'=========================================================================
' Methods
'=========================================================================

'--- parses the expression, raises with a readable message when it is not valid
Public Sub Init(sText As String)
    Const FUNC_NAME     As String = "Init"

    m_vRoot = Empty
    pvTokenize sText
    m_lPos = 0
    If m_lCount = 0 Then
        Exit Sub
    End If
    m_vRoot = pvParseOr()
    If pvPeek() <> ucsTokEnd Then
        RaiseError FUNC_NAME, "Unexpected " & m_aText(m_lPos)
    End If
End Sub

'--- values must be lower case
Public Function Matches(vValues As Variant) As Boolean
    If IsEmpty(m_vRoot) Then
        Matches = True
    Else
        Matches = pvEval(m_vRoot, vValues)
    End If
End Function

Private Sub pvTokenize(sText As String)
    Const FUNC_NAME     As String = "pvTokenize"
    Dim lIdx            As Long
    Dim lEnd            As Long
    Dim sWord           As String

    m_lCount = 0
    ReDim m_aText(0 To 7) As String
    ReDim m_aKind(0 To 7) As UcsTokenKind
    For lIdx = 1 To Len(sText)
        Select Case Mid$(sText, lIdx, 1)
        Case " ", vbTab
            pvAddWord sWord
        Case "("
            pvAddWord sWord
            pvAddToken ucsTokOpen, "("
        Case ")"
            pvAddWord sWord
            pvAddToken ucsTokClose, ")"
        Case """"
            pvAddWord sWord
            lEnd = InStr(lIdx + 1, sText, """")
            If lEnd = 0 Then
                RaiseError FUNC_NAME, "Missing closing quote"
            End If
            '--- quoted text is a term even when it reads like a keyword
            pvAddToken ucsTokTerm, Mid$(sText, lIdx + 1, lEnd - lIdx - 1)
            lIdx = lEnd
        Case Else
            sWord = sWord & Mid$(sText, lIdx, 1)
        End Select
    Next
    pvAddWord sWord
End Sub

Private Sub pvAddWord(sWord As String)
    If LenB(sWord) = 0 Then
        Exit Sub
    End If
    Select Case UCase$(sWord)
    Case "AND"
        pvAddToken ucsTokAnd, sWord
    Case "OR"
        pvAddToken ucsTokOr, sWord
    Case "NOT"
        pvAddToken ucsTokNot, sWord
    Case Else
        pvAddToken ucsTokTerm, sWord
    End Select
    sWord = vbNullString
End Sub

Private Sub pvAddToken(ByVal eKind As UcsTokenKind, sText As String)
    If m_lCount > UBound(m_aText) Then
        ReDim Preserve m_aText(0 To 2 * m_lCount) As String
        ReDim Preserve m_aKind(0 To 2 * m_lCount) As UcsTokenKind
    End If
    m_aText(m_lCount) = sText
    m_aKind(m_lCount) = eKind
    m_lCount = m_lCount + 1
End Sub

Private Function pvPeek() As UcsTokenKind
    If m_lPos >= m_lCount Then
        pvPeek = ucsTokEnd
    Else
        pvPeek = m_aKind(m_lPos)
    End If
End Function

Private Function pvParseOr() As Variant
    Dim vNode           As Variant

    vNode = pvParseAnd()
    Do While pvPeek() = ucsTokOr
        m_lPos = m_lPos + 1
        vNode = Array(ucsTokOr, vNode, pvParseAnd())
    Loop
    pvParseOr = vNode
End Function

'--- AND binds tighter than OR and may be left out before NOT or a bracket
Private Function pvParseAnd() As Variant
    Dim vNode           As Variant

    vNode = pvParseNot()
    Do
        Select Case pvPeek()
        Case ucsTokAnd
            m_lPos = m_lPos + 1
        Case ucsTokNot, ucsTokOpen, ucsTokTerm
        Case Else
            Exit Do
        End Select
        vNode = Array(ucsTokAnd, vNode, pvParseNot())
    Loop
    pvParseAnd = vNode
End Function

Private Function pvParseNot() As Variant
    If pvPeek() = ucsTokNot Then
        m_lPos = m_lPos + 1
        pvParseNot = Array(ucsTokNot, pvParseNot())
    Else
        pvParseNot = pvParsePrimary()
    End If
End Function

Private Function pvParsePrimary() As Variant
    Const FUNC_NAME     As String = "pvParsePrimary"
    Dim sTerm           As String

    Select Case pvPeek()
    Case ucsTokOpen
        m_lPos = m_lPos + 1
        pvParsePrimary = pvParseOr()
        If pvPeek() <> ucsTokClose Then
            RaiseError FUNC_NAME, "Missing )"
        End If
        m_lPos = m_lPos + 1
    Case ucsTokTerm
        '--- adjacent words are one term, so a saved "dreem 2" keeps its meaning
        sTerm = m_aText(m_lPos)
        m_lPos = m_lPos + 1
        Do While pvPeek() = ucsTokTerm
            sTerm = sTerm & " " & m_aText(m_lPos)
            m_lPos = m_lPos + 1
        Loop
        pvParsePrimary = Array(ucsTokTerm, pvGetPattern(sTerm))
    Case ucsTokEnd
        RaiseError FUNC_NAME, "Missing search term at the end"
    Case Else
        RaiseError FUNC_NAME, "Missing search term before " & m_aText(m_lPos)
    End Select
End Function

'--- a Like pattern, its own special characters taken literally
Private Function pvGetPattern(ByVal sTerm As String) As String
    sTerm = Replace(LCase$(sTerm), "[", "[[]")
    sTerm = Replace(Replace(sTerm, "?", "[?]"), "#", "[#]")
    sTerm = Replace(sTerm, "%", "*")
    If InStr(sTerm, "*") = 0 Then
        sTerm = "*" & sTerm & "*"
    End If
    pvGetPattern = sTerm
End Function

Private Function pvEval(vNode As Variant, vValues As Variant) As Boolean
    Dim vValue          As Variant

    Select Case vNode(0)
    Case ucsTokTerm
        For Each vValue In vValues
            If vValue Like vNode(1) Then
                pvEval = True
                Exit Function
            End If
        Next
    Case ucsTokNot
        pvEval = Not pvEval(vNode(1), vValues)
    Case ucsTokAnd
        If pvEval(vNode(1), vValues) Then
            pvEval = pvEval(vNode(2), vValues)
        End If
    Case ucsTokOr
        If pvEval(vNode(1), vValues) Then
            pvEval = True
        Else
            pvEval = pvEval(vNode(2), vValues)
        End If
    End Select
End Function
