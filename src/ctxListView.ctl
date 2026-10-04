VERSION 5.00
Begin VB.UserControl ctxListView
   ClientHeight    =   2880
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3840
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
End
Attribute VB_Name = "ctxListView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'=========================================================================
'
'   Reports Editor
'   Copyright (c) 2026 Unicontsoft
'
'   A report-view SysListView32 of our own, so COMCTL32.OCX is not needed.
'   A row is a 1-based Long and there are no wrapper objects
'
'=========================================================================
Option Explicit
DefObj A-Z
Private Const STR_MODULE_NAME As String = "ctxListView"

'=========================================================================
' Public events
'=========================================================================

Event Click()
Event DblClick()
Event KeyDown(KeyCode As Integer, Shift As Integer)
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event PreviewKeyDown(wParam As Long, lParam As Long, Cancel As Boolean)
Event Scrolled()
Event ItemPrePaint(ByVal Row As Long, Color As OLE_COLOR, BackColor As OLE_COLOR, Bold As Boolean, Handled As Boolean)
Event GetItemText(ByVal Row As Long, ByVal Col As Long, Text As String)
Event ColumnClick(ByVal Col As Long)
Event SelectionChanged()
Event RightClick(ByVal Row As Long)

'=========================================================================
' API
'=========================================================================

'--- window classes
Private Const STR_CLASS_LISTVIEW            As String = "SysListView32"
Private Const STR_THEME_EXPLORER            As String = "Explorer"

'=========================================================================
' Constants and member variables
'=========================================================================

Private Const MAX_ITEM_TEXT             As Long = 1024
Private Const DEF_MULTISELECT           As Boolean = False
Private Const DEF_OWNERDATA             As Boolean = False
Private Const DEF_SORTHEADERS           As Boolean = False

Private m_uIPAO                     As IPAOHookStruct
Private m_hList                     As LongPtr
Private m_hFontBold                 As LongPtr
Private m_pHostHook                 As IUnknown
Private m_pListHook                 As IUnknown
Private m_bMultiSelect              As Boolean
Private m_bOwnerData                As Boolean
Private m_bSortHeaders              As Boolean

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

'--- the next three are styles of the list, so they take effect when it is created
Public Property Get MultiSelect() As Boolean
    MultiSelect = m_bMultiSelect
End Property

Public Property Let MultiSelect(ByVal bValue As Boolean)
    m_bMultiSelect = bValue
    PropertyChanged "MultiSelect"
End Property

Public Property Get OwnerData() As Boolean
    OwnerData = m_bOwnerData
End Property

Public Property Let OwnerData(ByVal bValue As Boolean)
    m_bOwnerData = bValue
    PropertyChanged "OwnerData"
End Property

Public Property Get SortHeaders() As Boolean
    SortHeaders = m_bSortHeaders
End Property

Public Property Let SortHeaders(ByVal bValue As Boolean)
    m_bSortHeaders = bValue
    PropertyChanged "SortHeaders"
End Property

Public Property Get hWndList() As LongPtr
    hWndList = m_hList
End Property

Public Property Get ItemCount() As Long
    If m_hList <> 0 Then
        ItemCount = SendMessage(m_hList, LVM_GETITEMCOUNT, 0, ByVal 0&)
    End If
End Property

'--- only an owner-data list has a count of its own, its rows come from GetItemText
Public Property Let ItemCount(ByVal lValue As Long)
    If m_hList <> 0 And m_bOwnerData Then
        Call SendMessage(m_hList, LVM_SETITEMCOUNT, lValue, ByVal LVSICF_NOSCROLL)
    End If
End Property

Public Property Get ColumnCount() As Long
    Dim uColumn         As LVCOLUMN

    If m_hList = 0 Then
        Exit Property
    End If
    uColumn.Mask = LVCF_WIDTH
    Do While SendMessage(m_hList, LVM_GETCOLUMN, ColumnCount, uColumn) <> 0
        ColumnCount = ColumnCount + 1
    Loop
End Property

Public Property Get SelectedRow() As Long
    If m_hList <> 0 Then
        SelectedRow = SendMessage(m_hList, LVM_GETNEXTITEM, -1, ByVal LVNI_SELECTED) + 1
    End If
End Property

Public Property Let SelectedRow(ByVal lValue As Long)
    Dim uItem           As LVITEM

    If m_hList = 0 Then
        Exit Property
    End If
    uItem.stateMask = LVIS_SELECTED Or LVIS_FOCUSED
    If lValue < 1 Then
        Call SendMessage(m_hList, LVM_SETITEMSTATE, -1, uItem)
        Exit Property
    End If
    uItem.State = LVIS_SELECTED Or LVIS_FOCUSED
    Call SendMessage(m_hList, LVM_SETITEMSTATE, lValue - 1, uItem)
End Property

Public Property Get SelectedCount() As Long
    If m_hList <> 0 Then
        SelectedCount = SendMessage(m_hList, LVM_GETSELECTEDCOUNT, 0, ByVal 0&)
    End If
End Property

Public Property Get ItemSelected(ByVal lRow As Long) As Boolean
    If m_hList <> 0 And lRow >= 1 Then
        ItemSelected = (SendMessage(m_hList, LVM_GETITEMSTATE, lRow - 1, ByVal LVIS_SELECTED) <> 0)
    End If
End Property

Public Property Let ItemSelected(ByVal lRow As Long, ByVal bValue As Boolean)
    Dim uItem           As LVITEM

    If m_hList = 0 Or lRow < 1 Then
        Exit Property
    End If
    uItem.stateMask = LVIS_SELECTED
    If bValue Then
        uItem.State = LVIS_SELECTED
    End If
    Call SendMessage(m_hList, LVM_SETITEMSTATE, lRow - 1, uItem)
End Property

'--- the row with the focus rectangle, which selecting does not have to move
Public Property Get FocusedRow() As Long
    If m_hList <> 0 Then
        FocusedRow = SendMessage(m_hList, LVM_GETNEXTITEM, -1, ByVal LVNI_FOCUSED) + 1
    End If
End Property

Public Property Let FocusedRow(ByVal lValue As Long)
    Dim uItem           As LVITEM

    If m_hList = 0 Then
        Exit Property
    End If
    uItem.stateMask = LVIS_FOCUSED
    If lValue < 1 Then
        Call SendMessage(m_hList, LVM_SETITEMSTATE, -1, uItem)
        Exit Property
    End If
    uItem.State = LVIS_FOCUSED
    Call SendMessage(m_hList, LVM_SETITEMSTATE, lValue - 1, uItem)
End Property

Public Property Let Redraw(ByVal bValue As Boolean)
    If m_hList = 0 Then
        Exit Property
    End If
    Call SendMessage(m_hList, WM_SETREDRAW, -bValue, ByVal 0&)
    If bValue Then
        Refresh
    End If
End Property

Public Property Get HasBoldFont() As Boolean
    HasBoldFont = (m_hFontBold <> 0)
End Property

Public Property Get TopRow() As Long
    If m_hList <> 0 Then
        TopRow = SendMessage(m_hList, LVM_GETTOPINDEX, 0, ByVal 0&) + 1
    End If
End Property

Public Property Get ColumnText(ByVal lCol As Long) As String
    Dim uColumn         As LVCOLUMN
    Dim sBuffer         As String

    If m_hList = 0 Then
        Exit Property
    End If
    sBuffer = String$(MAX_ITEM_TEXT, 0)
    uColumn.Mask = LVCF_TEXT
    uColumn.pszText = StrPtr(sBuffer)
    uColumn.cchTextMax = MAX_ITEM_TEXT
    If SendMessage(m_hList, LVM_GETCOLUMN, lCol - 1, uColumn) <> 0 Then
        ColumnText = GetStringAt(StrPtr(sBuffer))
    End If
End Property

Public Property Let ColumnText(ByVal lCol As Long, ByVal sValue As String)
    Dim uColumn         As LVCOLUMN

    If m_hList = 0 Then
        Exit Property
    End If
    uColumn.Mask = LVCF_TEXT
    uColumn.pszText = StrPtr(sValue)
    Call SendMessage(m_hList, LVM_SETCOLUMN, lCol - 1, uColumn)
End Property

Public Property Get ColumnWidth(ByVal lCol As Long) As Long
    If m_hList <> 0 Then
        ColumnWidth = SendMessage(m_hList, LVM_GETCOLUMNWIDTH, lCol - 1, ByVal 0&)
    End If
End Property

Public Property Let ColumnWidth(ByVal lCol As Long, ByVal lValue As Long)
    If m_hList <> 0 Then
        Call SendMessage(m_hList, LVM_SETCOLUMNWIDTH, lCol - 1, ByVal lValue)
    End If
End Property

Public Property Get ItemText(ByVal lRow As Long, ByVal lCol As Long) As String
    Dim uItem           As LVITEM
    Dim sBuffer         As String

    If m_hList = 0 Or lRow < 1 Then
        Exit Property
    End If
    sBuffer = String$(MAX_ITEM_TEXT, 0)
    uItem.iSubItem = lCol - 1
    uItem.pszText = StrPtr(sBuffer)
    uItem.cchTextMax = MAX_ITEM_TEXT
    Call SendMessage(m_hList, LVM_GETITEMTEXT, lRow - 1, uItem)
    ItemText = GetStringAt(StrPtr(sBuffer))
End Property

Public Property Let ItemText(ByVal lRow As Long, ByVal lCol As Long, ByVal sValue As String)
    Dim uItem           As LVITEM

    If m_hList = 0 Or lRow < 1 Then
        Exit Property
    End If
    uItem.iSubItem = lCol - 1
    uItem.pszText = StrPtr(sValue)
    Call SendMessage(m_hList, LVM_SETITEMTEXT, lRow - 1, uItem)
End Property

'=========================================================================
' Methods
'=========================================================================

Public Sub AddColumn(ByVal sText As String, ByVal lWidth As Long, Optional ByVal Align As Long = LVCFMT_LEFT)
    Dim uColumn         As LVCOLUMN

    If m_hList = 0 Then
        Exit Sub
    End If
    uColumn.Mask = LVCF_FMT Or LVCF_WIDTH Or LVCF_TEXT Or LVCF_SUBITEM
    uColumn.fmt = Align
    uColumn.cx = lWidth
    uColumn.pszText = StrPtr(sText)
    uColumn.iSubItem = ColumnCount
    Call SendMessage(m_hList, LVM_INSERTCOLUMN, uColumn.iSubItem, uColumn)
End Sub

Public Sub ClearColumns()
    If m_hList = 0 Then
        Exit Sub
    End If
    Do While SendMessage(m_hList, LVM_DELETECOLUMN, 0, ByVal 0&) <> 0
    Loop
End Sub

'--- the row it landed on, which is the one below every row already there
Public Function AddItem(ByVal sText As String) As Long
    Dim uItem           As LVITEM

    If m_hList = 0 Then
        Exit Function
    End If
    uItem.Mask = LVIF_TEXT
    uItem.iItem = ItemCount
    uItem.pszText = StrPtr(sText)
    AddItem = SendMessage(m_hList, LVM_INSERTITEM, 0, uItem) + 1
End Function

Public Sub Clear()
    If m_hList <> 0 Then
        Call SendMessage(m_hList, LVM_DELETEALLITEMS, 0, ByVal 0&)
    End If
End Sub

Public Sub SelectAll()
    Dim uItem           As LVITEM

    If m_hList = 0 Or Not m_bMultiSelect Then
        Exit Sub
    End If
    uItem.stateMask = LVIS_SELECTED
    uItem.State = LVIS_SELECTED
    Call SendMessage(m_hList, LVM_SETITEMSTATE, -1, uItem)
End Sub

'--- repaints every row, which is how an owner-data list shows changed data
Public Sub Refresh()
    Dim lCount          As Long

    lCount = ItemCount
    If lCount > 0 Then
        Call SendMessage(m_hList, LVM_REDRAWITEMS, 0, ByVal lCount - 1)
    End If
End Sub

Public Sub EnsureVisible(ByVal lRow As Long)
    If m_hList <> 0 And lRow >= 1 Then
        Call SendMessage(m_hList, LVM_ENSUREVISIBLE, lRow - 1, ByVal 0&)
    End If
End Sub

'--- the row under a point in pixels, or zero where there is none
Public Function HitTest(ByVal lX As Long, ByVal lY As Long) As Long
    Dim uHit            As LVHITTESTINFO

    If m_hList = 0 Then
        Exit Function
    End If
    uHit.pt.X = lX
    uHit.pt.Y = lY
    HitTest = SendMessage(m_hList, LVM_HITTEST, 0, uHit) + 1
End Function

'--- where a cell is on screen, in pixels, clipped to what the window shows.
'--- A RECT would be a public UDT out of a standard module, which VB6 refuses
Public Function GetCellRect(ByVal lRow As Long, ByVal lCol As Long, lLeft As Long, _
            lTop As Long, lRight As Long, lBottom As Long) As Boolean
    Dim uRect           As RECT
    Dim uClient         As RECT

    If m_hList = 0 Or lRow < 1 Then
        Exit Function
    End If
    uRect.Top = lCol - 1
    uRect.Left = LVIR_BOUNDS
    If SendMessage(m_hList, LVM_GETSUBITEMRECT, lRow - 1, uRect) = 0 Then
        Exit Function
    End If
    If GetClientRect(m_hList, uClient) <> 0 Then
        If uRect.Right > uClient.Right Then
            uRect.Right = uClient.Right
        End If
    End If
    lLeft = uRect.Left
    lTop = uRect.Top
    lRight = uRect.Right
    lBottom = uRect.Bottom
    GetCellRect = (uRect.Right > uRect.Left)
End Function

Public Function SubclassProc( _
            ByVal hWnd As Long, _
            ByVal lMsg As Long, _
            ByVal lWParam As Long, _
            ByVal lParam As Long, _
            Handled As Boolean) As Long
    #If lWParam Then '--- touch args
    #End If
    Const NF_QUERY      As Long = 3
    Const NFR_UNICODE   As Long = 2
    Dim lRetVal         As Long

    Select Case lMsg
    Case WM_NOTIFYFORMAT
        '--- VB6 host is an ANSI window, so claim Unicode for the list's notifications
        If hWnd <> m_hList And lParam = NF_QUERY Then
            SubclassProc = NFR_UNICODE
            Handled = True
        End If
    Case WM_NOTIFY
        If pvCustomDraw(lParam, lRetVal) Then
            SubclassProc = lRetVal
            Handled = True
        Else
            pvNotify lParam
        End If
    Case WM_SETFOCUS
        If hWnd <> m_hList Then
            Call SetFocusAPI(m_hList)
        Else
            CallNextSubclassProc m_pHostHook, hWnd, lMsg, lWParam, lParam
            SetIPAO m_uIPAO, UserControl.hWnd
            Handled = True
        End If
    Case WM_LBUTTONDOWN, WM_RBUTTONDOWN
        pvRaiseMouseDown lMsg, lParam
    Case WM_VSCROLL, WM_HSCROLL, WM_MOUSEWHEEL
        If hWnd = m_hList Then
            RaiseEvent Scrolled
        End If
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
            Call SendMessage(m_hList, uMsg.lMessage, uMsg.wParam, ByVal uMsg.lParam)
            frTranslateAccel = True
        End Select
    End Select
End Function

'= private ===============================================================

Private Sub pvCreateList()
    Const FUNC_NAME     As String = "pvCreateList"
    Dim lStyle          As Long

    On Error GoTo EH
    If m_hList <> 0 Then
        Exit Sub
    End If
    lStyle = WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or LVS_REPORT Or LVS_SHOWSELALWAYS
    If Not m_bMultiSelect Then
        lStyle = lStyle Or LVS_SINGLESEL
    End If
    If m_bOwnerData Then
        lStyle = lStyle Or LVS_OWNERDATA
    End If
    If Not m_bSortHeaders Then
        lStyle = lStyle Or LVS_NOSORTHEADER
    End If
    m_hList = CreateWindowEx(0, StrPtr(STR_CLASS_LISTVIEW), 0, lStyle, _
        0, 0, ScaleWidth, ScaleHeight, UserControl.hWnd, 0, App.hInstance, ByVal 0&)
    If m_hList = 0 Then
        Exit Sub
    End If
    Call SendMessage(m_hList, LVM_SETEXTENDEDLISTVIEWSTYLE, 0, _
        ByVal (LVS_EX_FULLROWSELECT Or LVS_EX_GRIDLINES Or LVS_EX_DOUBLEBUFFER))
    '--- Explorer is what a list looks like on this machine
    Call SetWindowTheme(m_hList, StrPtr(STR_THEME_EXPLORER), 0)
    Exit Sub
EH:
    PrintError FUNC_NAME
    Resume Next
End Sub

'--- the list's own face, emboldened, for whoever asks a row to stand out
Private Sub pvCreateBoldFont()
    Dim hFont           As LongPtr
    Dim uLogFont        As LOGFONTW

    If m_hFontBold <> 0 Then
        Call DeleteObject(m_hFontBold)
        m_hFontBold = 0
    End If
    hFont = SendMessage(m_hList, WM_GETFONT, 0, ByVal 0&)
    If hFont = 0 Then
        Exit Sub
    End If
    If GetObjectAPI(hFont, LenB(uLogFont), uLogFont) = 0 Then
        Exit Sub
    End If
    uLogFont.lfWeight = FW_BOLD
    m_hFontBold = CreateFontIndirect(uLogFont)
End Sub

'--- True when the notification was custom draw, whose answer the caller returns
Private Function pvCustomDraw(ByVal lParam As Long, lReturn As Long) As Boolean
    Dim uDraw           As NMLVCUSTOMDRAW
    Dim clrText         As OLE_COLOR
    Dim clrBack         As OLE_COLOR
    Dim bBold           As Boolean
    Dim bHandled        As Boolean

    Call CopyMemory(uDraw, ByVal lParam, LenB(uDraw))
    If uDraw.Nmcd.Hdr.hWndFrom <> m_hList Or uDraw.Nmcd.Hdr.Code <> NM_CUSTOMDRAW Then
        Exit Function
    End If
    pvCustomDraw = True
    Select Case uDraw.Nmcd.dwDrawStage
    Case CDDS_PREPAINT
        lReturn = CDRF_NOTIFYITEMDRAW
    Case CDDS_ITEMPREPAINT
        lReturn = CDRF_DODEFAULT
        clrText = uDraw.clrText
        clrBack = uDraw.clrTextBk
        RaiseEvent ItemPrePaint(uDraw.Nmcd.dwItemSpec + 1, clrText, clrBack, bBold, bHandled)
        If Not bHandled Then
            Exit Function
        End If
        If bBold And m_hFontBold <> 0 Then
            Call SelectObject(uDraw.Nmcd.hDC, m_hFontBold)
            lReturn = CDRF_NEWFONT
        End If
        If clrText <> uDraw.clrText Or clrBack <> uDraw.clrTextBk Then
            uDraw.clrText = TranslateColor(clrText)
            uDraw.clrTextBk = TranslateColor(clrBack)
            Call CopyMemory(ByVal lParam, uDraw, LenB(uDraw))
        End If
    Case Else
        lReturn = CDRF_DODEFAULT
    End Select
End Function

Private Sub pvNotify(ByVal lParam As Long)
    Dim uHdr            As NMHDR
    Dim uKey            As NMLVKEYDOWN
    Dim nKeyCode        As Integer
    Dim uList           As NMLISTVIEW

    Call CopyMemory(uHdr, ByVal lParam, LenB(uHdr))
    If uHdr.hWndFrom <> m_hList Then
        Exit Sub
    End If
    Select Case uHdr.Code
    Case LVN_GETDISPINFOW
        pvGetDispInfo lParam
    Case NM_CLICK
        RaiseEvent Click
    Case NM_DBLCLK
        RaiseEvent DblClick
    Case NM_RCLICK
        Call CopyMemory(uList, ByVal lParam, LenB(uList))
        RaiseEvent RightClick(uList.iItem + 1)
    Case LVN_KEYDOWN
        Call CopyMemory(uKey, ByVal lParam, LenB(uHdr) + 2)
        nKeyCode = uKey.wVKey
        RaiseEvent KeyDown(nKeyCode, pvGetShiftState())
    Case LVN_COLUMNCLICK
        Call CopyMemory(uList, ByVal lParam, LenB(uList))
        RaiseEvent ColumnClick(uList.iSubItem + 1)
    Case LVN_ITEMCHANGED
        Call CopyMemory(uList, ByVal lParam, LenB(uList))
        If (uList.uChanged And LVIF_STATE) <> 0 Then
            If ((uList.uNewState Xor uList.uOldState) And (LVIS_SELECTED Or LVIS_FOCUSED)) <> 0 Then
                RaiseEvent SelectionChanged
            End If
        End If
    Case LVN_ODSTATECHANGED
        RaiseEvent SelectionChanged
    End Select
End Sub

'--- an owner-data row asks for its text, which goes back in the list's own buffer
Private Sub pvGetDispInfo(ByVal lParam As Long)
    Dim uInfo           As NMLVDISPINFO
    Dim sText           As String
    Dim lLen            As Long
    Dim nTerminator     As Integer

    Call CopyMemory(uInfo, ByVal lParam, LenB(uInfo))
    If (uInfo.Item.Mask And LVIF_TEXT) = 0 Or uInfo.Item.pszText = 0 Or uInfo.Item.cchTextMax <= 0 Then
        Exit Sub
    End If
    RaiseEvent GetItemText(uInfo.Item.iItem + 1, uInfo.Item.iSubItem + 1, sText)
    lLen = Len(sText)
    If lLen > uInfo.Item.cchTextMax - 1 Then
        lLen = uInfo.Item.cchTextMax - 1
    End If
    If lLen > 0 Then
        Call CopyMemory(ByVal uInfo.Item.pszText, ByVal StrPtr(sText), lLen * 2)
    End If
    Call CopyMemory(ByVal uInfo.Item.pszText + lLen * 2, nTerminator, 2)
End Sub

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

Private Function pvAddressOfSubclassProc() As ctxListView
    Set pvAddressOfSubclassProc = InitAddressOfMethod(Me, 5)
End Function

Private Sub pvApplyFont()
    Dim pFont           As IFont

    If m_hList = 0 Then
        Exit Sub
    End If
    Set pFont = UserControl.Font
    Call SendMessage(m_hList, WM_SETFONT, pFont.hFont, ByVal 1&)
    pvCreateBoldFont
End Sub

Private Sub pvInit()
    If Not Ambient.UserMode Then
        Exit Sub
    End If
    InitIPAO m_uIPAO, Me
    '--- host hook goes first so it answers WM_NOTIFYFORMAT while the list is created
    Set m_pHostHook = InitSubclassingThunk(UserControl.hWnd, Me, _
        pvAddressOfSubclassProc.SubclassProc(0, 0, 0, 0, 0))
    pvCreateList
    pvApplyFont
    Set m_pListHook = InitSubclassingThunk(m_hList, Me, _
        pvAddressOfSubclassProc.SubclassProc(0, 0, 0, 0, 0))
End Sub

'=========================================================================
' Control events
'=========================================================================

Private Sub UserControl_InitProperties()
    Set UserControl.Font = Ambient.Font
    m_bMultiSelect = DEF_MULTISELECT
    m_bOwnerData = DEF_OWNERDATA
    m_bSortHeaders = DEF_SORTHEADERS
    pvInit
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Set UserControl.Font = PropBag.ReadProperty("Font", Ambient.Font)
    m_bMultiSelect = PropBag.ReadProperty("MultiSelect", DEF_MULTISELECT)
    m_bOwnerData = PropBag.ReadProperty("OwnerData", DEF_OWNERDATA)
    m_bSortHeaders = PropBag.ReadProperty("SortHeaders", DEF_SORTHEADERS)
    pvInit
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Font", UserControl.Font, Ambient.Font)
    Call PropBag.WriteProperty("MultiSelect", m_bMultiSelect, DEF_MULTISELECT)
    Call PropBag.WriteProperty("OwnerData", m_bOwnerData, DEF_OWNERDATA)
    Call PropBag.WriteProperty("SortHeaders", m_bSortHeaders, DEF_SORTHEADERS)
End Sub

Private Sub UserControl_Resize()
    If m_hList <> 0 Then
        Call MoveWindow(m_hList, 0, 0, ScaleWidth, ScaleHeight, 1)
    End If
End Sub

Private Sub UserControl_EnterFocus()
    Call SetFocusAPI(m_hList)
    SetIPAO m_uIPAO, UserControl.hWnd
End Sub

'=========================================================================
' Base class events
'=========================================================================

Private Sub UserControl_Terminate()
    Set m_pListHook = Nothing
    Set m_pHostHook = Nothing
    TerminateIPAO m_uIPAO
    If m_hFontBold <> 0 Then
        Call DeleteObject(m_hFontBold)
        m_hFontBold = 0
    End If
    If m_hList <> 0 Then
        Call DestroyWindow(m_hList)
        m_hList = 0
    End If
End Sub
