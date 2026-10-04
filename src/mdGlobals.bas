Attribute VB_Name = "mdGlobals"
Option Explicit
Private Const MODULE_NAME As String = "mdGlobals"

'=========================================================================
' Public enums
'=========================================================================

'--- the enum hack, so a pointer is declared for what it is
Public Enum LongPtr
    [_]
End Enum

Public Enum UcsErrorIndexes
    ucsErrNumber
    ucsErrSource
    ucsErrDescription
    ucsErrHelpFile
    ucsErrHelpContext
End Enum

'=========================================================================
' API
'=========================================================================

'--- window messages
Public Const WM_SETFOCUS                    As Long = &H7
Public Const WM_SETREDRAW                   As Long = &HB
Public Const WM_SETFONT                     As Long = &H30
Public Const WM_GETFONT                     As Long = &H31
Public Const WM_NOTIFY                      As Long = &H4E
Public Const WM_KEYDOWN                     As Long = &H100
Public Const WM_HSCROLL                     As Long = &H114
Public Const WM_VSCROLL                     As Long = &H115
Public Const WM_LBUTTONDOWN                 As Long = &H201
Public Const WM_RBUTTONDOWN                 As Long = &H204
Public Const WM_MOUSEWHEEL                  As Long = &H20A
'--- window styles
Public Const WS_TABSTOP                     As Long = &H10000
Public Const WS_VISIBLE                     As Long = &H10000000
Public Const WS_CHILD                       As Long = &H40000000
'--- list-view window styles
Public Const LVS_REPORT                     As Long = &H1
Public Const LVS_SINGLESEL                  As Long = &H4
Public Const LVS_SHOWSELALWAYS              As Long = &H8
Public Const LVS_OWNERDATA                  As Long = &H1000
Public Const LVS_NOSORTHEADER               As Long = &H8000&
'--- extended list-view styles
Public Const LVS_EX_GRIDLINES               As Long = &H1
Public Const LVS_EX_FULLROWSELECT           As Long = &H20
Public Const LVS_EX_DOUBLEBUFFER            As Long = &H10000
'--- for LVITEM.Mask
Public Const LVIF_TEXT                      As Long = &H1
Public Const LVIF_STATE                     As Long = &H8
'--- for LVITEM.State
Public Const LVIS_FOCUSED                   As Long = &H1
Public Const LVIS_SELECTED                  As Long = &H2
'--- for LVM_GETNEXTITEM
Public Const LVNI_FOCUSED                   As Long = &H1
Public Const LVNI_SELECTED                  As Long = &H2
'--- for LVCOLUMN.fmt
Public Const LVCFMT_LEFT                    As Long = &H0
Public Const LVCFMT_RIGHT                   As Long = &H1
Public Const LVCFMT_CENTER                  As Long = &H2
'--- for LVCOLUMN.Mask
Public Const LVCF_FMT                       As Long = &H1
Public Const LVCF_WIDTH                     As Long = &H2
Public Const LVCF_TEXT                      As Long = &H4
Public Const LVCF_SUBITEM                   As Long = &H8
'--- for LVM_GETSUBITEMRECT
Public Const LVIR_BOUNDS                    As Long = 0
'--- for LVM_SETITEMCOUNT
Public Const LVSICF_NOSCROLL                As Long = &H2
'--- list-view messages
Public Const LVM_GETITEMCOUNT               As Long = &H1004
Public Const LVM_DELETEALLITEMS             As Long = &H1009
Public Const LVM_GETNEXTITEM                As Long = &H100C
Public Const LVM_HITTEST                    As Long = &H1012
Public Const LVM_ENSUREVISIBLE              As Long = &H1013
Public Const LVM_REDRAWITEMS                As Long = &H1015
Public Const LVM_DELETECOLUMN               As Long = &H101C
Public Const LVM_GETCOLUMNWIDTH             As Long = &H101D
Public Const LVM_SETCOLUMNWIDTH             As Long = &H101E
Public Const LVM_GETTOPINDEX                As Long = &H1027
Public Const LVM_SETITEMSTATE               As Long = &H102B
Public Const LVM_GETITEMSTATE               As Long = &H102C
Public Const LVM_SETITEMCOUNT               As Long = &H102F
Public Const LVM_GETSELECTEDCOUNT           As Long = &H1032
Public Const LVM_SETEXTENDEDLISTVIEWSTYLE   As Long = &H1036
Public Const LVM_GETSUBITEMRECT             As Long = &H1038
Public Const LVM_INSERTITEM                 As Long = &H104D
Public Const LVM_GETCOLUMN                  As Long = &H105F
Public Const LVM_SETCOLUMN                  As Long = &H1060
Public Const LVM_INSERTCOLUMN               As Long = &H1061
Public Const LVM_GETITEMTEXT                As Long = &H1073
Public Const LVM_SETITEMTEXT                As Long = &H1074
'--- list-view notifications
Public Const LVN_GETDISPINFOW               As Long = -177
Public Const LVN_KEYDOWN                    As Long = -155
Public Const LVN_ODSTATECHANGED             As Long = -115
Public Const LVN_COLUMNCLICK                As Long = -108
Public Const LVN_ITEMCHANGED                As Long = -101
Public Const NM_RCLICK                      As Long = -5
Public Const NM_DBLCLK                      As Long = -3
Public Const NM_CLICK                       As Long = -2
'--- for NM_CUSTOMDRAW, which is how a row is coloured or emboldened
Public Const NM_CUSTOMDRAW                  As Long = -12
Public Const CDRF_DODEFAULT                 As Long = &H0
Public Const CDDS_PREPAINT                  As Long = &H1
Public Const CDRF_NEWFONT                   As Long = &H2
Public Const CDRF_NOTIFYITEMDRAW            As Long = &H20
Public Const CDDS_ITEM                      As Long = &H10000
Public Const CDDS_ITEMPREPAINT              As Long = CDDS_ITEM Or CDDS_PREPAINT
'--- for CreateFontIndirect
Public Const FW_BOLD                        As Long = 700
'--- for InitCommonControlsEx
Private Const ICC_LISTVIEW_CLASSES                      As Long = &H1
Private Const ICC_USEREX_CLASSES                        As Long = &H200
'--- hresults
Private Const S_OK                                      As Long = 0
'--- for invoke
Private Const LOCALE_USER_DEFAULT                       As Long = &H400
'--- for VariantChangeType
Private Const VARIANT_ALPHABOOL                         As Long = 2

Private Declare Function InitCommonControlsEx Lib "comctl32.dll" (iccex As tagInitCommonControlsEx) As Boolean
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Private Declare Function VariantChangeType Lib "oleaut32" (Dest As Variant, Src As Variant, ByVal wFlags As Integer, ByVal vt As VbVarType) As Long
Private Declare Function SysReAllocString Lib "oleaut32" (ByVal pBSTR As LongPtr, ByVal lpsz As LongPtr) As Long
Private Declare Function OleTranslateColor Lib "oleaut32" (ByVal clr As Long, ByVal hPal As LongPtr, lColorRef As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageW" (ByVal hWnd As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, lParam As Any) As LongPtr
Public Declare Function GetClientRect Lib "user32" (ByVal hWnd As LongPtr, lpRect As RECT) As Long
Public Declare Function GetObjectAPI Lib "gdi32" Alias "GetObjectW" (ByVal hObject As LongPtr, ByVal nCount As Long, lpObject As Any) As Long
Public Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectW" (lpLogFont As LOGFONTW) As LongPtr
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As LongPtr) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hDC As LongPtr, ByVal hObject As LongPtr) As LongPtr
Public Declare Function SetWindowTheme Lib "uxtheme" (ByVal hWnd As LongPtr, ByVal pszSubAppName As LongPtr, ByVal pszSubIdList As LongPtr) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Public Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExW" (ByVal dwExStyle As Long, ByVal lpClassName As LongPtr, ByVal lpWindowName As LongPtr, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As LongPtr, ByVal hMenu As LongPtr, ByVal hInstance As LongPtr, lpParam As Any) As LongPtr
Public Declare Function DestroyWindow Lib "user32" (ByVal hWnd As LongPtr) As Long
Public Declare Function MoveWindow Lib "user32" (ByVal hWnd As LongPtr, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Public Declare Function SetFocusAPI Lib "user32" Alias "SetFocus" (ByVal hWnd As LongPtr) As LongPtr
Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer

Private Type tagInitCommonControlsEx
   lngSize              As Long
   lngICC               As Long
End Type

Private Type DISPPARAMS
    rgPointerToVariantArray As Long
    rgPointerToLongNamedArgs As Long
    cArgs               As Long
    cNamedArgs          As Long
End Type

Private Type EXCEPINFO
    wCode               As Integer
    wReserved           As Integer
    Source              As String
    Description         As String
    HelpFile            As String
    dwHelpContext       As Long
    pvReserved          As Long
    pfnDeferredFillIn   As Long
    sCode               As Long
End Type

Public Type RECT
    Left                As Long
    Top                 As Long
    Right               As Long
    Bottom              As Long
End Type

Public Type APIPOINT
    X                   As Long
    Y                   As Long
End Type

Public Type APIMSG
    hWnd                As Long
    lMessage            As Long
    wParam              As Long
    lParam              As Long
    lTime               As Long
    pt                  As APIPOINT
End Type

Public Type LOGFONTW
    lfHeight            As Long
    lfWidth             As Long
    lfEscapement        As Long
    lfOrientation       As Long
    lfWeight            As Long
    lfItalic            As Byte
    lfUnderline         As Byte
    lfStrikeOut         As Byte
    lfCharSet           As Byte
    lfOutPrecision      As Byte
    lfClipPrecision     As Byte
    lfQuality           As Byte
    lfPitchAndFamily    As Byte
    lfFaceName(0 To 31) As Integer
End Type

Public Type LVCOLUMN
    Mask                As Long
    fmt                 As Long
    cx                  As Long
    pszText             As LongPtr
    cchTextMax          As Long
    iSubItem            As Long
    iImage              As Long
    iOrder              As Long
End Type

Public Type LVITEM
    Mask                As Long
    iItem               As Long
    iSubItem            As Long
    State               As Long
    stateMask           As Long
    pszText             As LongPtr
    cchTextMax          As Long
    iImage              As Long
    lParam              As LongPtr
    iIndent             As Long
End Type

Public Type LVHITTESTINFO
    pt                  As APIPOINT
    Flags               As Long
    iItem               As Long
    iSubItem            As Long
End Type

Public Type NMHDR
    hWndFrom            As Long
    IdFrom              As Long
    Code                As Long
End Type

Public Type NMCUSTOMDRAW
    Hdr                 As NMHDR
    dwDrawStage         As Long
    hDC                 As Long
    rc                  As RECT
    dwItemSpec          As Long
    uItemState          As Long
    lItemlParam         As Long
End Type

Public Type NMLISTVIEW
    Hdr                 As NMHDR
    iItem               As Long
    iSubItem            As Long
    uNewState           As Long
    uOldState           As Long
    uChanged            As Long
    ptAction            As APIPOINT
    lParam              As LongPtr
End Type

Public Type NMLVDISPINFO
    Hdr                 As NMHDR
    Item                As LVITEM
End Type

Public Type NMLVKEYDOWN
    Hdr                 As NMHDR
    wVKey               As Integer
    Flags               As Long
End Type

Public Type NMLVCUSTOMDRAW
    Nmcd                As NMCUSTOMDRAW
    clrText             As Long
    clrTextBk           As Long
    iSubItem            As Long
End Type

'=========================================================================
' Constants and member variables
'=========================================================================

Public Const STR_APP_NAME      As String = "Ucs SQL Monitor"

'=========================================================================
' Error handling
'=========================================================================

Private Sub PrintError(sFunc As String)
    Debug.Print MODULE_NAME & "." & sFunc & ": " & Error
End Sub

Public Function PushError(Optional vErr As Variant) As Variant
    vErr = Array(Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
    PushError = vErr
End Function

Public Sub PopRaiseError(vErr As Variant, sModule As String, sFunction As String)
    Err.Raise vErr(ucsErrNumber), , vErr(ucsErrDescription) & " [" & sModule & "." & sFunction & "]"
End Sub

Public Sub PopPrintError(vErr As Variant, sModule As String, sFunction As String)
    Debug.Print "Error " & vErr(ucsErrNumber) & ": " & vErr(ucsErrDescription) & " [" & sModule & "." & sFunction & "]"
End Sub

'=========================================================================
' Functions
'=========================================================================

Public Function CreateRecordset(ParamArray FldDesc()) As Recordset
    Dim vFldDesc        As Variant
    
    vFldDesc = FldDesc
    Set CreateRecordset = CreateRecordsetArray(vFldDesc)
End Function

Public Function CreateRecordsetArray(FldDesc As Variant) As Recordset
    Const FUNC_NAME     As String = "CreateRecordsetArray"
    Dim lIdx            As Long
        
    On Error GoTo EH
    Set CreateRecordsetArray = New Recordset
    With CreateRecordsetArray.Fields
        If UBound(FldDesc) < 0 Then
            .Append "ID", adGUID, , adFldIsNullable
        Else
            Do While lIdx < UBound(FldDesc)
                Select Case FldDesc(lIdx + 1)
                Case adVarChar, adChar, adVarWChar, adWChar
                    .Append FldDesc(lIdx), FldDesc(lIdx + 1), FldDesc(lIdx + 2), adFldIsNullable
                    lIdx = lIdx + 3
                Case adDecimal
                    .Append FldDesc(lIdx), FldDesc(lIdx + 1), , adFldIsNullable
                    With .Item(.Count - 1)
                        .Precision = FldDesc(lIdx + 2)
                        .NumericScale = FldDesc(lIdx + 3)
                    End With
                    lIdx = lIdx + 4
                Case Else
                    .Append FldDesc(lIdx), FldDesc(lIdx + 1), , adFldIsNullable
                    lIdx = lIdx + 2
                End Select
            Loop
        End If
    End With
    CreateRecordsetArray.Open
    Exit Function
EH:
    PrintError FUNC_NAME
End Function


Public Function SearchRecordset( _
            rs As Recordset, _
            sCriteria As String) As Boolean
    Const FUNC_NAME     As String = "SearchRecordset"
    
    On Error GoTo EH
    If Not rs Is Nothing Then
        If rs.RecordCount <> 0 Then
            With rs
                .MoveFirst
                .Find sCriteria
                SearchRecordset = Not .EOF
            End With
        End If
    End If
    Exit Function
EH:
    PrintError FUNC_NAME
    '--- fall through and return false
'    resume next
End Function

Public Function Quote(sText As String) As String
    Quote = Replace(Replace(sText, "'", "''"), Chr$(0), vbNullString)
End Function

Public Function C_Str(Value As Variant) As String
    Dim vDest           As Variant
    
    If VarType(Value) = vbString Then
        C_Str = Value
    ElseIf VariantChangeType(vDest, Value, VARIANT_ALPHABOOL, vbString) = 0 Then
        C_Str = vDest
    End If
End Function

Public Function C_Bool(Value As Variant) As Boolean
    Dim vDest           As Variant
    
    If VarType(Value) = vbBoolean Then
        C_Bool = Value
    ElseIf VariantChangeType(vDest, Value, VARIANT_ALPHABOOL, vbBoolean) = 0 Then
        C_Bool = vDest
    End If
End Function

Public Function C_Dbl(Value As Variant) As Double
    Dim vDest           As Variant
    
    If VarType(Value) = vbDouble Then
        C_Dbl = Value
    ElseIf VariantChangeType(vDest, Value, 0, vbDouble) = 0 Then
        C_Dbl = vDest
    End If
End Function

Public Function C_Lng(Value As Variant) As Long
    Dim vDest           As Variant
    
    If VarType(Value) = vbLong Then
        C_Lng = Value
    ElseIf VariantChangeType(vDest, Value, 0, vbLong) = 0 Then
        C_Lng = vDest
    End If
End Function

Public Function Limit( _
            ByVal Value As Double, _
            Optional Min As Variant, _
            Optional Max As Variant) As Double
    Const FUNC_NAME     As String = "Limit"
    
    On Error GoTo EH
    Limit = Value
    If Not IsMissing(Min) Then
        If Value < C_Dbl(Min) Then
            Limit = C_Dbl(Min)
        End If
    End If
    If Not IsMissing(Max) Then
        If Value > C_Dbl(Max) Then
            Limit = C_Dbl(Max)
        End If
    End If
    Exit Function
EH:
    PrintError FUNC_NAME
    Resume Next
End Function

Public Function InitCommonControlsVB() As Boolean
   Dim iccex            As tagInitCommonControlsEx
   
   On Error Resume Next
   Call LoadLibrary("shell32.dll")
   With iccex
       .lngSize = LenB(iccex)
       .lngICC = ICC_LISTVIEW_CLASSES Or ICC_USEREX_CLASSES
   End With
   Call InitCommonControlsEx(iccex)
   InitCommonControlsVB = (Err.Number = 0)
   On Error GoTo 0
End Function

'--- the header and the rows as tab-separated text, all rows or only the selected ones
Public Sub ClipCopy(oList As ctxListView, Optional ByVal SelectedOnly As Boolean)
    Dim lColCount       As Long
    Dim aCells()        As String
    Dim lCol            As Long
    Dim sText           As String
    Dim lRow            As Long

    lColCount = oList.ColumnCount
    If lColCount = 0 Then
        Exit Sub
    End If
    ReDim aCells(0 To lColCount - 1) As String
    For lCol = 1 To lColCount
        aCells(lCol - 1) = oList.ColumnText(lCol)
    Next
    sText = Join(aCells, vbTab) & vbCrLf
    For lRow = 1 To oList.ItemCount
        If Not SelectedOnly Or oList.ItemSelected(lRow) Then
            For lCol = 1 To lColCount
                aCells(lCol - 1) = oList.ItemText(lRow, lCol)
            Next
            sText = sText & Join(aCells, vbTab) & vbCrLf
        End If
    Next
    Clipboard.Clear
    Clipboard.SetText sText
End Sub

'--- a wide string somewhere in memory, up to its terminator, as a String
Public Function GetStringAt(ByVal lPtr As LongPtr) As String
    Call SysReAllocString(VarPtr(GetStringAt), lPtr)
End Function

'--- a system colour is a negative OLE colour and has to be looked up
Public Function TranslateColor(ByVal clrValue As OLE_COLOR) As Long
    If clrValue >= 0 Then
        TranslateColor = clrValue
    ElseIf OleTranslateColor(clrValue, 0, TranslateColor) <> 0 Then
        TranslateColor = 0
    End If
End Function

'--- the coordinate pair an lParam carries, which is signed
Public Function GetXLParam(ByVal lParam As Long) As Long
    Dim nWord           As Integer

    Call CopyMemory(nWord, lParam, 2)
    GetXLParam = nWord
End Function

Public Function GetYLParam(ByVal lParam As Long) As Long
    Dim nWord           As Integer

    Call CopyMemory(nWord, ByVal VarPtr(lParam) + 2, 2)
    GetYLParam = nWord
End Function

Public Function SetTrue(bValue As Boolean) As Boolean
    bValue = True
    SetTrue = True
End Function

Public Function GetShiftState() As ShiftConstants
    GetShiftState = vbShiftMask * -IsKeyPressed(vbKeyShift) _
                Or vbCtrlMask * -IsKeyPressed(vbKeyControl) _
                Or vbAltMask * -IsKeyPressed(vbKeyMenu)
End Function

Public Function IsKeyPressed(ByVal lVirtKey As KeyCodeConstants) As Boolean
    IsKeyPressed = ((GetAsyncKeyState(lVirtKey) And &H8000) = &H8000)
End Function

Public Function SetAbsolutePosition( _
            rs As Recordset, _
            ByVal lPosition As Long) As Boolean
    On Error Resume Next
    rs.AbsolutePosition = lPosition
    SetAbsolutePosition = (Err.Number = 0) And Not rs.EOF And Not rs.BOF
    On Error GoTo 0
End Function

Public Function InitIndexCollection( _
            rs As Recordset, _
            sFld As String, _
            Optional Fld2 As String, _
            Optional Fld3 As String, _
            Optional Fld4 As String, _
            Optional Fld5 As String, _
            Optional ByVal HasDuplicates As Boolean, _
            Optional RetVal As Collection) As Collection
'    Const FUNC_NAME     As String = "InitIndexCollection"
    Dim oFld            As ADODB.Field
    Dim oFld2           As ADODB.Field
    Dim oFld3           As ADODB.Field
    Dim oFld4           As ADODB.Field
    Dim oFld5           As ADODB.Field
    Dim vBmk            As Variant
    Dim pCol            As IVbCollection
    Dim sKey            As String
    
    On Error GoTo EH
    Set RetVal = New Collection
    If MoveRecordset(rs, 0) Then
        With rs
            vBmk = rs.Bookmark
            If LenB(sFld) <> 0 Then
                Set oFld = .Fields(sFld)
            End If
            If LenB(Fld2) <> 0 Then
                Set oFld2 = .Fields(Fld2)
            End If
            If LenB(Fld3) <> 0 Then
                Set oFld3 = .Fields(Fld3)
            End If
            If LenB(Fld4) <> 0 Then
                Set oFld4 = .Fields(Fld4)
            End If
            If LenB(Fld5) <> 0 Then
                Set oFld5 = .Fields(Fld5)
            End If
            If HasDuplicates Then
                Set pCol = RetVal
                If oFld Is Nothing Then
                    Do
                        pCol.Add .Bookmark, "#" & C_Str(.Bookmark)
                        .MoveNext
                    Loop While Not .EOF
                ElseIf oFld2 Is Nothing Then
                    Do
                        pCol.Add .Bookmark, "#" & Trim$(C_Str(oFld.Value))
                        .MoveNext
                    Loop While Not .EOF
                ElseIf oFld3 Is Nothing Then
                    Do
                        pCol.Add .Bookmark, "#" & C_Str(oFld.Value) & "#" & C_Str(oFld2.Value)
                        .MoveNext
                    Loop While Not .EOF
                ElseIf oFld4 Is Nothing Then
                    Do
                        pCol.Add .Bookmark, "#" & C_Str(oFld.Value) & "#" & C_Str(oFld2.Value) & "#" & C_Str(oFld3.Value)
                        .MoveNext
                    Loop While Not .EOF
                ElseIf oFld5 Is Nothing Then
                    Do
                        pCol.Add .Bookmark, "#" & C_Str(oFld.Value) & "#" & C_Str(oFld2.Value) & "#" & C_Str(oFld3.Value) & "#" & C_Str(oFld4.Value)
                        .MoveNext
                    Loop While Not .EOF
                Else
                    Do
                        pCol.Add .Bookmark, "#" & C_Str(oFld.Value) & "#" & C_Str(oFld2.Value) & "#" & C_Str(oFld3.Value) & "#" & C_Str(oFld4.Value) & "#" & C_Str(oFld5.Value)
                        .MoveNext
                    Loop While Not .EOF
                End If
            Else
                If oFld Is Nothing Then
                    Do
                        RetVal.Add .Bookmark, "#" & C_Str(.Bookmark)
                        .MoveNext
                    Loop While Not .EOF
                ElseIf oFld2 Is Nothing Then
                    Do
                        sKey = "#" & Trim$(C_Str(oFld.Value))
                        RemoveCollection RetVal, sKey
                        RetVal.Add .Bookmark, sKey
                        .MoveNext
                    Loop While Not .EOF
                ElseIf oFld3 Is Nothing Then
                    Do
                        sKey = "#" & C_Str(oFld.Value) & "#" & C_Str(oFld2.Value)
                        RemoveCollection RetVal, sKey
                        RetVal.Add .Bookmark, sKey
                        .MoveNext
                    Loop While Not .EOF
                ElseIf oFld4 Is Nothing Then
                    Do
                        sKey = "#" & C_Str(oFld.Value) & "#" & C_Str(oFld2.Value) & "#" & C_Str(oFld3.Value)
                        RemoveCollection RetVal, sKey
                        RetVal.Add .Bookmark, sKey
                        .MoveNext
                    Loop While Not .EOF
                ElseIf oFld5 Is Nothing Then
                    Do
                        sKey = "#" & C_Str(oFld.Value) & "#" & C_Str(oFld2.Value) & "#" & C_Str(oFld3.Value) & "#" & C_Str(oFld4.Value)
                        RemoveCollection RetVal, sKey
                        RetVal.Add .Bookmark, sKey
                        .MoveNext
                    Loop While Not .EOF
                Else
                    Do
                        sKey = "#" & C_Str(oFld.Value) & "#" & C_Str(oFld2.Value) & "#" & C_Str(oFld3.Value) & "#" & C_Str(oFld4.Value) & "#" & C_Str(oFld5.Value)
                        RemoveCollection RetVal, sKey
                        RetVal.Add .Bookmark, sKey
                        .MoveNext
                    Loop While Not .EOF
                End If
            End If
            SetBookmark rs, vBmk
        End With
    End If
    Set InitIndexCollection = RetVal
    Exit Function
EH:
'    If RaiseError(FUNC_NAME & "(sKey=" & sKey & ")") = vbRetry Then
'        Resume
'    End If
End Function

Public Function SetBookmark( _
            rs As Recordset, _
            vBmk As Variant, _
            Optional Key As Variant) As Boolean
    Dim cIndex          As Collection
    Dim vTemp           As Variant
    Dim vResult         As Variant
    
    On Error GoTo QH
    If IsObject(vBmk) Then
        Set cIndex = vBmk
        If Not cIndex Is Nothing Then
            If SearchCollection(cIndex, Key, RetVal:=vTemp) Then
                DispInvoke rs, "Bookmark", VbLet, vTemp
                DispInvoke rs, "Bookmark", VbLet, vTemp
                If DispInvoke(rs, "Bookmark", VbGet, RetVal:=vResult) Then
                    SetBookmark = (vResult = vTemp)
                End If
            End If
        End If
    Else
        rs.Bookmark = vBmk
        rs.Bookmark = vBmk
        SetBookmark = (rs.Bookmark = vBmk)
    End If
QH:
End Function

Public Sub AssignVariant(vDest As Variant, vSrc As Variant)
    If IsObject(vSrc) And IsObject(vDest) Then
        Set vDest = vSrc
    ElseIf Not IsObject(vSrc) And Not IsObject(vDest) Then
        vDest = vSrc
    End If
End Sub

Public Function MoveRecordset(rs As Recordset, lIter As Long) As Boolean
    On Error GoTo EH
    If rs Is Nothing Then
        GoTo EH
    End If
    If rs.RecordCount = 0 Then
        GoTo EH
    End If
    If lIter = 0 Then
        rs.MoveFirst
    Else
        rs.MoveNext
    End If
    If Not rs.EOF And Not rs.BOF Then
        If Not IsObject(rs.Bookmark) Or rs.CursorLocation = adUseServer Or rs.Status = adRecDeleted Then
            lIter = lIter + 1
            MoveRecordset = True
            Exit Function
        End If
    End If
EH:
    lIter = 0
End Function

Public Function SearchCollection(ByVal pCol As Object, Index As Variant, Optional RetVal As Variant) As Boolean
    Const DISPID_VALUE  As Long = 0
    Const VT_BYREF      As Long = &H4000
    Dim pVbCol          As IVbCollection
    Dim vItem           As Variant
    
    If pCol Is Nothing Then
        '--- do nothing
    ElseIf (PeekInt(VarPtr(RetVal)) And VT_BYREF) = 0 Then
        If TypeOf pCol Is IVbCollection Then
            Set pVbCol = pCol
            SearchCollection = pVbCol.Item(Index, RetVal) = S_OK
        Else
            SearchCollection = DispInvoke(pCol, DISPID_VALUE, VbMethod Or VbGet, RetVal:=RetVal, Args:=Index)
        End If
    Else
        If TypeOf pCol Is IVbCollection Then
            Set pVbCol = pCol
            SearchCollection = pVbCol.Item(Index, vItem) = S_OK
        Else
            SearchCollection = DispInvoke(pCol, DISPID_VALUE, VbMethod Or VbGet, RetVal:=vItem, Args:=Index)
        End If
        If SearchCollection Then
            If IsObject(vItem) Then
                Set RetVal = vItem
            Else
                RetVal = vItem
            End If
        End If
    End If
End Function

Public Function RemoveCollection(pVbCol As IVbCollection, Index As Variant) As Boolean
    If Not pVbCol Is Nothing Then
        RemoveCollection = pVbCol.Remove(Index) = 0
    End If
End Function

Public Property Get InIde() As Boolean
    Debug.Assert pvSetTrue(InIde)
End Property

Private Function pvSetTrue(bValue As Boolean) As Boolean
    bValue = True
    pvSetTrue = True
End Function

Public Function DispInvoke( _
            ByVal pDisp As IVbDispatch, _
            Name As Variant, _
            Optional ByVal CallType As VbCallType, _
            Optional Args As Variant, _
            Optional RetVal As Variant) As Boolean
    Const DISPID_PROPERTYPUT As Long = -3
    Const VT_BYREF      As Long = &H4000
    Dim IID_NULL        As VBGUID
    Dim lDispID         As Long
    Dim hResult         As Long
    Dim uParams         As DISPPARAMS
    Dim uInfo           As EXCEPINFO
    Dim aParams()       As Variant
    Dim lNamedParam     As Long
    Dim lIdx            As Long
    Dim lParamCount     As Long
    Dim lArgErr         As Long
    Dim lPtrResult      As Long
    Dim vRetVal         As Variant

    If pDisp Is Nothing Then
        Exit Function
    End If
    '--- get disp id
    If IsNumeric(Name) Then
        lDispID = C_Lng(Name)
    Else
        hResult = pDisp.GetIDsOfNames(IID_NULL, C_Str(Name), 1, LOCALE_USER_DEFAULT, lDispID)
        If hResult < 0 Then
            GoTo QH
        End If
    End If
    If CallType = 0 Then
        CallType = VbCallType.VbMethod Or IIf(Not IsMissing(RetVal), VbCallType.VbGet, 0)
    End If
    '--- process params
    If Not IsMissing(Args) Then
        If IsArray(Args) Then
            lParamCount = UBound(Args) - LBound(Args)
            ReDim aParams(0 To lParamCount) As Variant
            For lIdx = 0 To lParamCount
                Call AssignVariant(aParams(lParamCount - lIdx), Args(lIdx))
            Next
        Else
            ReDim aParams(0 To 0) As Variant
            Call AssignVariant(aParams(0), Args)
        End If
        With uParams
            .cArgs = lParamCount + 1
            .rgPointerToVariantArray = VarPtr(aParams(0))
        End With
        If (CallType And (VbCallType.VbLet Or VbCallType.VbSet)) <> 0 Then
            lNamedParam = DISPID_PROPERTYPUT
            With uParams
                .cNamedArgs = 1
                .rgPointerToLongNamedArgs = VarPtr(lNamedParam)
            End With
        End If
    End If
    If (CallType And VbCallType.VbGet) <> 0 Or (CallType And VbCallType.VbMethod) <> 0 And Not IsMissing(RetVal) Then
        lPtrResult = VarPtr(RetVal)
        If (PeekInt(lPtrResult) And VT_BYREF) = 0 Then
            If IsObject(RetVal) Then
                Set RetVal = Nothing
            Else
                RetVal = Empty
            End If
        Else
            lPtrResult = VarPtr(vRetVal)
            If IsObject(RetVal) Then
                Set vRetVal = Nothing
            Else
                vRetVal = Empty
            End If
        End If
    End If
    hResult = pDisp.Invoke(lDispID, IID_NULL, LOCALE_USER_DEFAULT, CallType, uParams, ByVal lPtrResult, uInfo, lArgErr)
    If hResult < 0 Then
        GoTo QH
    End If
    If lPtrResult = VarPtr(vRetVal) Then
        If IsObject(vRetVal) Then
            Set RetVal = vRetVal
        Else
            RetVal = vRetVal
        End If
    End If
    '--- success
    DispInvoke = True
    Exit Function
QH:
    If VarType(RetVal) = vbVariant Then
        RetVal = Array(hResult, uInfo.sCode, uInfo.Description, uInfo.Source)
    End If
End Function
