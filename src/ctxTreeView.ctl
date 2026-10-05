VERSION 5.00
Begin VB.UserControl ctxTreeView
   ClientHeight    =   2880
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3840
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
End
Attribute VB_Name = "ctxTreeView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'=========================================================================
'
'   Reports Editor
'   Copyright (c) 2026 Unicontsoft
'
'   A SysTreeView32 of our own, so COMCTL32.OCX is not needed. A node is an
'   HTREEITEM and the keys it was added under are kept here
'
'=========================================================================
Option Explicit
DefObj A-Z
Private Const STR_MODULE_NAME As String = "ctxTreeView"

'=========================================================================
' Public events
'=========================================================================

Event NodeClick(ByVal hItem As LongPtr)
Event DblClick()
Event KeyDown(KeyCode As Integer, Shift As Integer)
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event PreviewKeyDown(wParam As Long, lParam As Long, Cancel As Boolean)
Event ItemPrePaint(ByVal hItem As LongPtr, Color As OLE_COLOR, Handled As Boolean)
Event BeforeCollapse(ByVal hItem As LongPtr, Cancel As Boolean)

'=========================================================================
' Constants and member variables
'=========================================================================

Private Const MAX_NODE_TEXT             As Long = 1024

Private m_uIPAO                     As IPAOHookStruct
Private m_hTree                     As LongPtr
Private m_cKeys                     As Collection
Private m_cItems                    As Collection
Private m_bNoClickEvent             As Boolean
Private m_pHostHook                 As IUnknown
Private m_pTreeHook                 As IUnknown

'=========================================================================
' Error management
'=========================================================================

Private Sub PrintError(sFunction As String)
    PopPrintError PushError, STR_MODULE_NAME, sFunction
End Sub

'=========================================================================
' Properties
'=========================================================================

Public Property Get Font() As StdFont
    Set Font = UserControl.Font
End Property

Public Property Set Font(oValue As StdFont)
    Set UserControl.Font = oValue
    pvApplyFont
    PropertyChanged "Font"
End Property

Public Property Get hWndTree() As LongPtr
    hWndTree = m_hTree
End Property

Public Property Let ImageList(ByVal hValue As LongPtr)
    If m_hTree <> 0 Then
        Call SendMessage(m_hTree, TVM_SETIMAGELIST, TVSIL_NORMAL, ByVal hValue)
    End If
End Property

Public Property Get NodeCount() As Long
    If m_hTree <> 0 Then
        NodeCount = SendMessage(m_hTree, TVM_GETCOUNT, 0, ByVal 0&)
    End If
End Property

Public Property Get SelectedNode() As LongPtr
    If m_hTree <> 0 Then
        SelectedNode = SendMessage(m_hTree, TVM_GETNEXTITEM, TVGN_CARET, ByVal 0&)
    End If
End Property

'--- a selection made here is not a click, so nothing is told about it
Public Property Let SelectedNode(ByVal hValue As LongPtr)
    If m_hTree = 0 Then
        Exit Property
    End If
    m_bNoClickEvent = True
    Call SendMessage(m_hTree, TVM_SELECTITEM, TVGN_CARET, ByVal hValue)
    m_bNoClickEvent = False
End Property

Public Property Get NodeText(ByVal hItem As LongPtr) As String
    Dim uItem           As TVITEM
    Dim sBuffer         As String

    If m_hTree = 0 Or hItem = 0 Then
        Exit Property
    End If
    sBuffer = String$(MAX_NODE_TEXT, 0)
    uItem.Mask = TVIF_TEXT
    uItem.hItem = hItem
    uItem.pszText = StrPtr(sBuffer)
    uItem.cchTextMax = MAX_NODE_TEXT
    If SendMessage(m_hTree, TVM_GETITEMW, 0, uItem) <> 0 Then
        NodeText = GetStringAt(StrPtr(sBuffer))
    End If
End Property

Public Property Let NodeText(ByVal hItem As LongPtr, ByVal sValue As String)
    Dim uItem           As TVITEM

    If m_hTree = 0 Or hItem = 0 Then
        Exit Property
    End If
    uItem.Mask = TVIF_TEXT
    uItem.hItem = hItem
    uItem.pszText = StrPtr(sValue)
    Call SendMessage(m_hTree, TVM_SETITEMW, 0, uItem)
End Property

Public Property Let NodeBold(ByVal hItem As LongPtr, ByVal bValue As Boolean)
    Dim uItem           As TVITEM

    If m_hTree = 0 Or hItem = 0 Then
        Exit Property
    End If
    uItem.Mask = TVIF_STATE
    uItem.hItem = hItem
    uItem.stateMask = TVIS_BOLD
    If bValue Then
        uItem.State = TVIS_BOLD
    End If
    Call SendMessage(m_hTree, TVM_SETITEM, 0, uItem)
End Property

Public Property Let NodeExpanded(ByVal hItem As LongPtr, ByVal bValue As Boolean)
    Dim lFlag           As Long

    If m_hTree = 0 Or hItem = 0 Then
        Exit Property
    End If
    If bValue Then
        lFlag = TVE_EXPAND
    Else
        lFlag = TVE_COLLAPSE
    End If
    Call SendMessage(m_hTree, TVM_EXPAND, lFlag, ByVal hItem)
End Property

'=========================================================================
' Methods
'=========================================================================

'--- a node under hParent, or a root when that is zero; the key is kept here
Public Function AddNode(ByVal hParent As LongPtr, ByVal sKey As String, _
            ByVal sText As String, Optional ByVal Image As Long = -1) As LongPtr
    Dim uIns            As TVINSERTSTRUCT

    If m_hTree = 0 Then
        Exit Function
    End If
    uIns.hParent = hParent
    uIns.hInsertAfter = TVI_LAST
    '--- a node arrives open: this tree is read, not navigated
    uIns.Item.Mask = TVIF_TEXT Or TVIF_STATE
    uIns.Item.State = TVIS_EXPANDED
    uIns.Item.stateMask = TVIS_EXPANDED
    uIns.Item.pszText = StrPtr(sText)
    If Image >= 0 Then
        uIns.Item.Mask = uIns.Item.Mask Or TVIF_IMAGE Or TVIF_SELECTEDIMAGE
        uIns.Item.iImage = Image
        uIns.Item.iSelectedImage = Image
    End If
    AddNode = SendMessage(m_hTree, TVM_INSERTITEMW, 0, uIns)
    If AddNode = 0 Then
        Exit Function
    End If
    If LenB(sKey) <> 0 Then
        m_cKeys.Add AddNode, LCase$(sKey)
        m_cItems.Add sKey, CStr(AddNode)
    End If
End Function

Public Sub Clear()
    Set m_cKeys = New Collection
    Set m_cItems = New Collection
    If m_hTree <> 0 Then
        Call SendMessage(m_hTree, TVM_DELETEITEM, 0, ByVal TVI_ROOT)
    End If
End Sub

'--- the keys that went in are kept, so a missing node is an answer not an error
Public Function NodeByKey(ByVal sKey As String) As LongPtr
    If SearchCollection(m_cKeys, LCase$(sKey)) Then
        NodeByKey = m_cKeys.Item(LCase$(sKey))
    End If
End Function

Public Function NodeKey(ByVal hItem As LongPtr) As String
    If SearchCollection(m_cItems, CStr(hItem)) Then
        NodeKey = m_cItems.Item(CStr(hItem))
    End If
End Function

Public Function GetRootNode() As LongPtr
    If m_hTree <> 0 Then
        GetRootNode = SendMessage(m_hTree, TVM_GETNEXTITEM, TVGN_ROOT, ByVal 0&)
    End If
End Function

Public Function GetChildNode(ByVal hItem As LongPtr) As LongPtr
    If m_hTree <> 0 Then
        GetChildNode = SendMessage(m_hTree, TVM_GETNEXTITEM, TVGN_CHILD, ByVal hItem)
    End If
End Function

Public Function GetNextNode(ByVal hItem As LongPtr) As LongPtr
    If m_hTree <> 0 Then
        GetNextNode = SendMessage(m_hTree, TVM_GETNEXTITEM, TVGN_NEXT, ByVal hItem)
    End If
End Function

Public Sub EnsureVisible(ByVal hItem As LongPtr)
    If m_hTree <> 0 And hItem <> 0 Then
        Call SendMessage(m_hTree, TVM_ENSUREVISIBLE, 0, ByVal hItem)
    End If
End Sub

'--- orders the direct children by text, case-insensitive
Public Sub SortChildren(ByVal hItem As LongPtr)
    If m_hTree <> 0 And hItem <> 0 Then
        Call SendMessage(m_hTree, TVM_SORTCHILDREN, 0, ByVal hItem)
    End If
End Sub

'--- the node under a point in pixels, or zero where there is none
Public Function HitTest(ByVal lX As Long, ByVal lY As Long) As LongPtr
    Dim uHit            As TVHITTESTINFO

    If m_hTree = 0 Then
        Exit Function
    End If
    uHit.pt.X = lX
    uHit.pt.Y = lY
    HitTest = SendMessage(m_hTree, TVM_HITTEST, 0, uHit)
End Function

Public Function SubclassProc( _
            ByVal hWnd As Long, _
            ByVal lMsg As Long, _
            ByVal lWParam As Long, _
            ByVal lParam As Long, _
            Handled As Boolean) As Long
    #If lWParam Then '--- touch args
    #End If
    Dim lRetVal         As Long

    Select Case lMsg
    Case WM_NOTIFY
        If pvCustomDraw(lParam, lRetVal) Then
            SubclassProc = lRetVal
            Handled = True
        ElseIf pvNotify(lParam, lRetVal) Then
            SubclassProc = lRetVal
            Handled = True
        End If
    Case WM_SETFOCUS
        If hWnd <> m_hTree Then
            Call SetFocusAPI(m_hTree)
        Else
            CallNextSubclassProc m_pHostHook, hWnd, lMsg, lWParam, lParam
            SetIPAO m_uIPAO, UserControl.hWnd
            Handled = True
        End If
    Case WM_LBUTTONDOWN, WM_RBUTTONDOWN
        pvRaiseMouseDown lMsg, lParam
    End Select
End Function

Friend Function frTranslateAccel(uMsg As APIMSG) As Boolean
    Select Case uMsg.lMessage
    Case WM_KEYDOWN
        RaiseEvent PreviewKeyDown(uMsg.wParam, uMsg.lParam, frTranslateAccel)
        If frTranslateAccel Then
            Exit Function
        End If
        '--- VB6 moves focus on an arrow unless the control claims it first
        Select Case uMsg.wParam
        Case vbKeyUp, vbKeyDown, vbKeyLeft, vbKeyRight, _
                vbKeyPageUp, vbKeyPageDown, vbKeyHome, vbKeyEnd
            Call SendMessage(m_hTree, uMsg.lMessage, uMsg.wParam, ByVal uMsg.lParam)
            frTranslateAccel = True
        End Select
    End Select
End Function

'= private ===============================================================

Private Sub pvCreateTree()
    Const FUNC_NAME     As String = "pvCreateTree"
    Dim lStyle          As Long

    On Error GoTo EH
    If m_hTree <> 0 Then
        Exit Sub
    End If
    '--- no lines: the Explorer theme draws chevrons where they would go
    lStyle = WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or TVS_HASBUTTONS _
        Or TVS_SHOWSELALWAYS Or TVS_TRACKSELECT
    m_hTree = CreateWindowEx(0, StrPtr(STR_CLASS_TREEVIEW), 0, lStyle, _
        0, 0, ScaleWidth, ScaleHeight, UserControl.hWnd, 0, App.hInstance, ByVal 0&)
    If m_hTree = 0 Then
        Exit Sub
    End If
    '--- Explorer is what a tree looks like on this machine
    Call SetWindowTheme(m_hTree, StrPtr(STR_THEME_EXPLORER), 0)
    Call SendMessage(m_hTree, TVM_SETEXTENDEDSTYLE, TVS_EX_FADEINOUTEXPANDOS, _
        ByVal TVS_EX_FADEINOUTEXPANDOS)
    Call SendMessage(m_hTree, TVM_SETEXTENDEDSTYLE, TVS_EX_DOUBLEBUFFER, _
        ByVal TVS_EX_DOUBLEBUFFER)
    Exit Sub
EH:
    PrintError FUNC_NAME
    Resume Next
End Sub

'--- True when the notification was custom draw, whose answer the caller returns
Private Function pvCustomDraw(ByVal lParam As Long, lReturn As Long) As Boolean
    Dim uDraw           As NMTVCUSTOMDRAW
    Dim clrText         As OLE_COLOR
    Dim bHandled        As Boolean

    Call CopyMemory(uDraw, ByVal lParam, LenB(uDraw))
    If uDraw.Nmcd.Hdr.hWndFrom <> m_hTree Or uDraw.Nmcd.Hdr.Code <> NM_CUSTOMDRAW Then
        Exit Function
    End If
    pvCustomDraw = True
    Select Case uDraw.Nmcd.dwDrawStage
    Case CDDS_PREPAINT
        lReturn = CDRF_NOTIFYITEMDRAW
    Case CDDS_ITEMPREPAINT
        lReturn = CDRF_DODEFAULT
        clrText = uDraw.clrText
        RaiseEvent ItemPrePaint(uDraw.Nmcd.dwItemSpec, clrText, bHandled)
        If bHandled And clrText <> uDraw.clrText Then
            uDraw.clrText = TranslateColor(clrText)
            Call CopyMemory(ByVal lParam, uDraw, LenB(uDraw))
        End If
    Case Else
        lReturn = CDRF_DODEFAULT
    End Select
End Function

'--- True when the notification needs an answer, which the caller returns
Private Function pvNotify(ByVal lParam As Long, lReturn As Long) As Boolean
    Dim uHdr            As NMHDR
    Dim uTree           As NMTREEVIEW
    Dim uKey            As NMLVKEYDOWN
    Dim nKeyCode        As Integer
    Dim bCancel         As Boolean

    Call CopyMemory(uHdr, ByVal lParam, LenB(uHdr))
    If uHdr.hWndFrom <> m_hTree Then
        Exit Function
    End If
    Select Case uHdr.Code
    Case TVN_ITEMEXPANDING, TVN_ITEMEXPANDINGW
        Call CopyMemory(uTree, ByVal lParam, LenB(uTree))
        If (uTree.Action And TVE_COLLAPSE) <> 0 Then
            RaiseEvent BeforeCollapse(uTree.itemNew.hItem, bCancel)
            If bCancel Then
                lReturn = 1
                pvNotify = True
            End If
        End If
    Case TVN_SELCHANGED, TVN_SELCHANGEDW
        If Not m_bNoClickEvent Then
            Call CopyMemory(uTree, ByVal lParam, LenB(uTree))
            RaiseEvent NodeClick(uTree.itemNew.hItem)
        End If
    Case NM_DBLCLK
        RaiseEvent DblClick
    Case TVN_KEYDOWN
        '--- NMTVKEYDOWN is packed, so copy only as far as the key code
        Call CopyMemory(uKey, ByVal lParam, LenB(uHdr) + 2)
        nKeyCode = uKey.wVKey
        RaiseEvent KeyDown(nKeyCode, pvGetShiftState())
    End Select
End Function

Private Sub pvRaiseMouseDown(ByVal lMsg As Long, ByVal lParam As Long)
    Dim nButton         As Integer

    If lMsg = WM_RBUTTONDOWN Then
        nButton = vbRightButton
    Else
        nButton = vbLeftButton
    End If
    RaiseEvent MouseDown(nButton, pvGetShiftState(), _
        GetXLParam(lParam), GetYLParam(lParam))
End Sub

Private Function pvGetShiftState() As Integer
    If (GetKeyState(vbKeyShift) And &H8000&) <> 0 Then
        pvGetShiftState = pvGetShiftState Or vbShiftMask
    End If
    If (GetKeyState(vbKeyControl) And &H8000&) <> 0 Then
        pvGetShiftState = pvGetShiftState Or vbCtrlMask
    End If
    If (GetKeyState(vbKeyMenu) And &H8000&) <> 0 Then
        pvGetShiftState = pvGetShiftState Or vbAltMask
    End If
End Function

Private Function pvAddressOfSubclassProc() As ctxTreeView
    Set pvAddressOfSubclassProc = InitAddressOfMethod(Me, 5)
End Function

Private Sub pvApplyFont()
    Dim pFont           As IFont

    If m_hTree = 0 Then
        Exit Sub
    End If
    Set pFont = UserControl.Font
    Call SendMessage(m_hTree, WM_SETFONT, pFont.hFont, ByVal 1&)
End Sub

Private Sub pvInit()
    Set m_cKeys = New Collection
    Set m_cItems = New Collection
    If Not Ambient.UserMode Then
        Exit Sub
    End If
    InitIPAO m_uIPAO, Me
    pvCreateTree
    pvApplyFont
    Set m_pHostHook = InitSubclassingThunk(UserControl.hWnd, Me, _
        pvAddressOfSubclassProc.SubclassProc(0, 0, 0, 0, 0))
    Set m_pTreeHook = InitSubclassingThunk(m_hTree, Me, _
        pvAddressOfSubclassProc.SubclassProc(0, 0, 0, 0, 0))
End Sub

'=========================================================================
' Control events
'=========================================================================

Private Sub UserControl_InitProperties()
    Set UserControl.Font = Ambient.Font
    pvInit
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Set UserControl.Font = PropBag.ReadProperty("Font", Ambient.Font)
    pvInit
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Font", UserControl.Font, Ambient.Font)
End Sub

Private Sub UserControl_Resize()
    If m_hTree <> 0 Then
        Call MoveWindow(m_hTree, 0, 0, ScaleWidth, ScaleHeight, 1)
    End If
End Sub

Private Sub UserControl_EnterFocus()
    Call SetFocusAPI(m_hTree)
    SetIPAO m_uIPAO, UserControl.hWnd
End Sub

'=========================================================================
' Base class events
'=========================================================================

Private Sub UserControl_Terminate()
    Set m_pTreeHook = Nothing
    Set m_pHostHook = Nothing
    TerminateIPAO m_uIPAO
    If m_hTree <> 0 Then
        Call DestroyWindow(m_hTree)
        m_hTree = 0
    End If
End Sub
