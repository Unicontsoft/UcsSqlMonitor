Attribute VB_Name = "mdGlobals"
Option Explicit
Private Const MODULE_NAME As String = "mdGlobals"

'=========================================================================
' API
'=========================================================================

'--- for InitCommonControlsEx
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
       .lngICC = ICC_USEREX_CLASSES
   End With
   Call InitCommonControlsEx(iccex)
   InitCommonControlsVB = (Err.Number = 0)
   On Error GoTo 0
End Function

Public Sub ClipCopy(geCtl As GridEX)
    Dim lIdx            As Long

    For lIdx = 1 To geCtl.RowCount
        geCtl.RowSelected(lIdx) = True
    Next
    Clipboard.Clear
    Clipboard.SetText Replace(geCtl.GetClipString(True), vbCr, vbCrLf)
End Sub

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
