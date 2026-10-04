Attribute VB_Name = "mdIPAO"
'=========================================================================
'
'   Reports Editor
'   Copyright (c) 2026 Unicontsoft
'
'   Custom IOleInPlaceActiveObject implementation
'
'=========================================================================
Option Explicit
DefObj A-Z
Private Const STR_MODULE_NAME As String = "mdIPAO"

'===========================================================================
' Light-weight object definition
'===========================================================================

Public Type IPAOHookStruct
    HookObjPtr      As LongPtr
    CtlName         As String
End Type

'===========================================================================
' API
'===========================================================================

Private Const PTR_SIZE                      As Long = 4
Private Const NULL_PTR                      As Long = 0
Private Const S_OK                          As Long = 0

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As LongPtr)
Private Declare Function ArrPtr Lib "msvbvm60" Alias "VarPtr" (Ptr() As UcsIPAOHook) As LongPtr
Private Declare Function IsEqualGUID Lib "ole32" (iid1 As Any, iid2 As Any) As Long
Private Declare Function vbaObjSetAddref Lib "msvbvm60" Alias "__vbaObjSetAddref" (oDest As Any, ByVal lSrcPtr As LongPtr) As LongPtr
Private Declare Function CoTaskMemAlloc Lib "ole32" (ByVal cb As LongPtr) As LongPtr
Private Declare Function CoTaskMemFree Lib "ole32" (ByVal pv As Long) As Long

Private Type VBGUID
    Data1               As Long
    Data2               As Integer
    Data3               As Integer
    Data4(0 To 7)       As Byte
End Type

Private Type UcsIPAOHook
    lpVTable            As LongPtr
    IPAORealPtr         As LongPtr
    CtlPtr              As LongPtr
    ThisPtr             As LongPtr
    CtlName             As String
End Type
Private Const sizeof_UcsIPAOHook As Long = 24

Private Type SAFEARRAY1D
    cDims               As Integer
    fFeatures           As Integer
    cbElements          As Long
    cLocks              As Long
    pvData              As LongPtr
    cElements           As Long
    lLbound             As Long
End Type

'===========================================================================
' Constants and member variables
'===========================================================================

Private IID_IOleInPlaceActiveObject As VBGUID
Private m_pVTable                   As LongPtr
Private m_uPeekArray                As SAFEARRAY1D
Private m_aPeekBuffer()             As UcsIPAOHook

'===========================================================================
' Error handling
'===========================================================================

Private Function RaiseError(sFunction As String) As VbMsgBoxResult
    PopRaiseError PushError, STR_MODULE_NAME, sFunction
End Function

Private Function PrintError(sFunction As String) As VbMsgBoxResult
    PopPrintError PushError, STR_MODULE_NAME, sFunction
End Function

'===========================================================================
' Functions
'===========================================================================

Public Sub InitIPAO(uHook As IPAOHookStruct, oCtl As Object)
    Const FUNC_NAME     As String = "InitIPAO"
    Dim oIPAOReal       As IOleInPlaceActiveObject
    Dim oExt            As VBControlExtender
    
    On Error GoTo EH
    Set oExt = GetExtendedControl(oCtl)
    If Not oExt Is Nothing Then
        uHook.CtlName = TypeName(oExt.Parent) & "." & oExt.Name
    End If
    Set oIPAOReal = oCtl
    uHook.HookObjPtr = CoTaskMemAlloc(sizeof_UcsIPAOHook)
    If uHook.HookObjPtr = NULL_PTR Then
        Exit Sub
    End If
    If m_uPeekArray.cDims = 0 Then
        With m_uPeekArray
            .cDims = 1
            .fFeatures = 1 ' FADF_AUTO
            .cbElements = sizeof_UcsIPAOHook
        End With
        Call CopyMemory(ByVal ArrPtr(m_aPeekBuffer), VarPtr(m_uPeekArray), PTR_SIZE)
    End If
    m_uPeekArray.pvData = uHook.HookObjPtr
    m_uPeekArray.cElements = 1
    With m_aPeekBuffer(0)
        .lpVTable = pvGetVTable
        .IPAORealPtr = ObjPtr(oIPAOReal)
        .CtlPtr = ObjPtr(oCtl)
        .ThisPtr = uHook.HookObjPtr
        '--- "borrow" CtlName string from uHook
'        Call CopyMemory(ByVal VarPtr(.CtlName), StrPtr(uHook.CtlName), PTR_SIZE)
    End With
    m_uPeekArray.pvData = NULL_PTR
    m_uPeekArray.cElements = 0
    Exit Sub
EH:
    If RaiseError(FUNC_NAME & "(uHook.CtlName=" & uHook.CtlName & ")") = vbRetry Then
        Resume
    End If
End Sub

Public Sub TerminateIPAO(uHook As IPAOHookStruct)
    Const FUNC_NAME     As String = "TerminateIPAO"
    Dim lCtlPtr         As LongPtr
    
    If uHook.HookObjPtr = NULL_PTR Then
        Exit Sub
    End If
    m_uPeekArray.pvData = uHook.HookObjPtr
    m_uPeekArray.cElements = 1
    lCtlPtr = m_aPeekBuffer(0).CtlPtr
    m_uPeekArray.pvData = NULL_PTR
    m_uPeekArray.cElements = 0
    pvSetIPAO lCtlPtr, NULL_PTR, NULL_PTR
    Call CoTaskMemFree(uHook.HookObjPtr)
    uHook.HookObjPtr = NULL_PTR
    Exit Sub
EH:
    If RaiseError(FUNC_NAME & "(uHook.CtlName=" & uHook.CtlName & ")") = vbRetry Then
        Resume
    End If
End Sub

Public Sub SetIPAO(uHook As IPAOHookStruct, ByVal hWnd As LongPtr)
    Const FUNC_NAME     As String = "SetIPAO"
    Dim lCtlPtr         As LongPtr
    Dim lThisPtr        As LongPtr

    On Error GoTo EH
    If uHook.HookObjPtr = NULL_PTR Then
        Exit Sub
    End If
    m_uPeekArray.pvData = uHook.HookObjPtr
    m_uPeekArray.cElements = 1
    lCtlPtr = m_aPeekBuffer(0).CtlPtr
    lThisPtr = m_aPeekBuffer(0).ThisPtr
    m_uPeekArray.pvData = NULL_PTR
    m_uPeekArray.cElements = 0
    pvSetIPAO lCtlPtr, lThisPtr, hWnd
    Exit Sub
EH:
    If RaiseError(FUNC_NAME & "(uHook.CtlName=" & uHook.CtlName & ")") = vbRetry Then
        Resume
    End If
End Sub

Public Sub RestoreIPAO(uHook As IPAOHookStruct)
    Const FUNC_NAME     As String = "RestoreIPAO"
    Dim lCtlPtr         As LongPtr
    Dim lIPAORealPtr    As LongPtr
    
    On Error GoTo EH
    If uHook.HookObjPtr = NULL_PTR Then
        Exit Sub
    End If
    m_uPeekArray.pvData = uHook.HookObjPtr
    m_uPeekArray.cElements = 1
    lCtlPtr = m_aPeekBuffer(0).CtlPtr
    lIPAORealPtr = m_aPeekBuffer(0).IPAORealPtr
    m_uPeekArray.pvData = NULL_PTR
    m_uPeekArray.cElements = 0
    pvSetIPAO lCtlPtr, lIPAORealPtr, NULL_PTR
    Exit Sub
EH:
    If RaiseError(FUNC_NAME & "(uHook.CtlName=" & uHook.CtlName & ")") = vbRetry Then
        Resume
    End If
End Sub

'= private =================================================================

Private Sub pvSetIPAO(ByVal lCtlPtr As LongPtr, ByVal pvActiveObj As LongPtr, ByVal hWnd As LongPtr)
    Const FUNC_NAME         As String = "pvSetIPAO"
    Const OLEIVERB_UIACTIVATE As Long = -4
    Dim oCtl                As Object
    Dim pOleObject          As IOleObject
    Dim pOleInPlaceSite     As IOleInPlaceSite
    Dim pOleInPlaceFrame    As IOleInPlaceFrame
    Dim pOleInPlaceUIWindow As IOleInPlaceUIWindow
    Dim rcPos               As RECT
    Dim rcClip              As RECT
    Dim uFrameInfo          As OLEINPLACEFRAMEINFO
       
    On Error GoTo EH
    If lCtlPtr = NULL_PTR Then
        Exit Sub
    End If
    Set oCtl = pvToObject(lCtlPtr)
    If Not TypeOf oCtl Is IOleObject Then
        Exit Sub
    End If
    Set pOleObject = oCtl
    If pOleObject.GetClientSite(pOleInPlaceSite) <> S_OK Then
        Exit Sub
    End If
    If pOleInPlaceSite Is Nothing Then
        Exit Sub
    End If
    '--- note: moje da grymne s Access Violation
    On Error Resume Next '--- checked
    pOleInPlaceSite.GetWindowContext pOleInPlaceFrame, pOleInPlaceUIWindow, VarPtr(rcPos), VarPtr(rcClip), VarPtr(uFrameInfo)
    On Error GoTo EH
    If Not pOleInPlaceFrame Is Nothing Then
        pOleInPlaceFrame.SetActiveObject pvActiveObj, vbNullString
    End If
    If Not pOleInPlaceUIWindow Is Nothing Then '-- And Not m_bMouseActivate
        pOleInPlaceUIWindow.SetActiveObject pvActiveObj, vbNullString
    End If
    If hWnd <> NULL_PTR Then
        pOleObject.DoVerb OLEIVERB_UIACTIVATE, 0, ObjPtr(pOleInPlaceSite), 0, hWnd, VarPtr(rcPos)
    End If
    Exit Sub
EH:
    If RaiseError(FUNC_NAME) = vbRetry Then
        Resume
    End If
End Sub

Private Function pvGetVTable() As LongPtr
    Dim STR_RELEASE_THUNK       As String: STR_RELEASE_THUNK = "i1QkBItCBIsIUP9RCMIEAA==" ' 13.5.2020 20:15:19
    Const RELEASE_THUNK_SIZE    As Long = 16
    Dim aVTable(0 To 9)     As LongPtr
    
    If m_pVTable = 0 Then
        '--- init guid
        With IID_IOleInPlaceActiveObject
           .Data1 = &H117
           .Data4(0) = &HC0
           .Data4(7) = &H46
        End With
        aVTable(0) = pvToPtr(AddressOf QueryInterface)
        aVTable(1) = pvToPtr(AddressOf AddRef)
        aVTable(2) = ThunkAllocate(STR_RELEASE_THUNK, RELEASE_THUNK_SIZE)
        aVTable(3) = pvToPtr(AddressOf GetWindow)
        aVTable(4) = pvToPtr(AddressOf ContextSensitiveHelp)
        aVTable(5) = pvToPtr(AddressOf TranslateAccelerator)
        aVTable(6) = pvToPtr(AddressOf OnFrameWindowActivate)
        aVTable(7) = pvToPtr(AddressOf OnDocWindowActivate)
        aVTable(8) = pvToPtr(AddressOf ResizeBorder)
        aVTable(9) = pvToPtr(AddressOf EnableModeless)
        m_pVTable = CoTaskMemAlloc(PTR_SIZE * 10)
        Call CopyMemory(ByVal m_pVTable, aVTable(0), PTR_SIZE * 10)
    End If
    pvGetVTable = m_pVTable
End Function

Private Function pvToObject(ByVal lPtr As LongPtr) As Object
    Call vbaObjSetAddref(pvToObject, lPtr)
End Function

Private Function pvToIOleIPAO(ByVal lPtr As LongPtr) As IOleInPlaceActiveObject
    Call vbaObjSetAddref(pvToIOleIPAO, lPtr)
End Function

Private Function pvToPtr(ByVal lPtr As LongPtr) As LongPtr
    pvToPtr = lPtr
End Function

Private Function GetExtendedControl(oCtl As IUnknown) As VBControlExtender
    Const FUNC_NAME     As String = "GetExtendedControl"
    Dim pOleObject      As IOleObject
    Dim pOleControlSite As IOleControlSite
    
    On Error GoTo EH
    If oCtl Is Nothing Then
        Exit Function
    End If
    If Not TypeOf oCtl Is IOleObject Then
        Exit Function
    End If
    Set pOleObject = oCtl
    If pOleObject.GetClientSite(pOleControlSite) <> S_OK Then
        Exit Function
    End If
    If Not pOleControlSite Is Nothing Then
        Set GetExtendedControl = pOleControlSite.GetExtendedControl
    End If
    Exit Function
EH:
    PrintError FUNC_NAME
End Function

' = interface implemenattion ================================================

Private Function AddRef(This As UcsIPAOHook) As Long
    Const FUNC_NAME     As String = "AddRef"
    
    On Error GoTo EH
    AddRef = pvToIOleIPAO(This.IPAORealPtr).AddRef
    Exit Function
EH:
    If PrintError(FUNC_NAME & "(This.CtlName=" & This.CtlName & ")") = vbRetry Then
        Resume
    End If
    Resume Next
End Function

'Private Function Release(This As UcsIPAOHook) As Long
'    Const FUNC_NAME     As String = "Release"
'
'    On Error GoTo EH
'    Release = pvToIOleIPAO(This.IPAORealPtr).Release
'    Exit Function
'EH:
'    If PrintError(FUNC_NAME & "(This.CtlName=" & This.CtlName & ")") = vbRetry Then
'        Resume
'    End If
'    Resume Next
'End Function

Private Function QueryInterface(This As UcsIPAOHook, riid As VBGUID, pvObj As LongPtr) As Long
    Const FUNC_NAME     As String = "QueryInterface"
    
    On Error GoTo EH
    If IsEqualGUID(riid, IID_IOleInPlaceActiveObject) Then
        pvObj = This.ThisPtr
        AddRef This
    Else
        QueryInterface = pvToIOleIPAO(This.IPAORealPtr).QueryInterface(ByVal VarPtr(riid), pvObj)
    End If
    Exit Function
EH:
    If PrintError(FUNC_NAME & "(This.CtlName=" & This.CtlName & ")") = vbRetry Then
        Resume
    End If
    Resume Next
End Function

Private Function GetWindow(This As UcsIPAOHook, phwnd As LongPtr) As Long
    Const FUNC_NAME     As String = "GetWindow"
    
    On Error GoTo EH
    GetWindow = pvToIOleIPAO(This.IPAORealPtr).GetWindow(phwnd)
    Exit Function
EH:
    If PrintError(FUNC_NAME & "(This.CtlName=" & This.CtlName & ")") = vbRetry Then
        Resume
    End If
    Resume Next
End Function

Private Function ContextSensitiveHelp(This As UcsIPAOHook, ByVal fEnterMode As Long) As Long
    Const FUNC_NAME     As String = "ContextSensitiveHelp"
    
    On Error GoTo EH
    ContextSensitiveHelp = pvToIOleIPAO(This.IPAORealPtr).ContextSensitiveHelp(fEnterMode)
    Exit Function
EH:
    If PrintError(FUNC_NAME & "(This.CtlName=" & This.CtlName & ")") = vbRetry Then
        Resume
    End If
    Resume Next
End Function

Private Function TranslateAccelerator(This As UcsIPAOHook, uMsg As APIMSG) As Long
    Const FUNC_NAME     As String = "TranslateAccelerator"
    Dim bInIde          As Boolean: Debug.Assert SetTrue(bInIde)
    Dim oList           As ctxListView
    Dim bHandled        As Boolean
    Dim oCtl            As Object
    Dim pIPAOReal       As IOleInPlaceActiveObject '--- weakref
    
    On Error GoTo EH
    If This.CtlPtr <> NULL_PTR Then
        Set oCtl = pvToObject(This.CtlPtr)
        If TypeOf oCtl Is ctxListView Then
            Set oList = oCtl
            bHandled = oList.frTranslateAccel(uMsg)
        End If
    End If
    If bHandled Then
        TranslateAccelerator = S_OK
    ElseIf bInIde Then
        TranslateAccelerator = pvToIOleIPAO(This.IPAORealPtr).TranslateAccelerator(ByVal VarPtr(uMsg))
    ElseIf This.IPAORealPtr <> NULL_PTR Then
        '--- skip refcounting on This.IPAORealPtr
        Call CopyMemory(pIPAOReal, This.IPAORealPtr, PTR_SIZE)
        TranslateAccelerator = pIPAOReal.TranslateAccelerator(ByVal VarPtr(uMsg))
        Call CopyMemory(pIPAOReal, NULL_PTR, PTR_SIZE)
    End If
    Exit Function
EH:
    If PrintError(FUNC_NAME & "(This.CtlName=" & This.CtlName & ")") = vbRetry Then
        Resume
    End If
    Resume Next
End Function

Private Function OnFrameWindowActivate(This As UcsIPAOHook, ByVal fActivate As Long) As Long
    Const FUNC_NAME     As String = "OnFrameWindowActivate"
    
    On Error GoTo EH
    OnFrameWindowActivate = pvToIOleIPAO(This.IPAORealPtr).OnFrameWindowActivate(fActivate)
    Exit Function
EH:
    If PrintError(FUNC_NAME & "(This.CtlName=" & This.CtlName & ")") = vbRetry Then
        Resume
    End If
    Resume Next
End Function

Private Function OnDocWindowActivate(This As UcsIPAOHook, ByVal fActivate As Long) As Long
    Const FUNC_NAME     As String = "OnDocWindowActivate"
    
    On Error GoTo EH
    OnDocWindowActivate = pvToIOleIPAO(This.IPAORealPtr).OnDocWindowActivate(fActivate)
    Exit Function
EH:
    If PrintError(FUNC_NAME & "(This.CtlName=" & This.CtlName & ")") = vbRetry Then
        Resume
    End If
    Resume Next
End Function

Private Function ResizeBorder(This As UcsIPAOHook, prcBorder As RECT, ByVal puiWindow As IOleInPlaceUIWindow, ByVal fFrameWindow As Long) As Long
    Const FUNC_NAME     As String = "ResizeBorder"
    
    On Error GoTo EH
    ResizeBorder = pvToIOleIPAO(This.IPAORealPtr).ResizeBorder(VarPtr(prcBorder), puiWindow, fFrameWindow)
    Exit Function
EH:
    If PrintError(FUNC_NAME & "(This.CtlName=" & This.CtlName & ")") = vbRetry Then
        Resume
    End If
    Resume Next
End Function

Private Function EnableModeless(This As UcsIPAOHook, ByVal fEnable As Long) As Long
    Const FUNC_NAME     As String = "EnableModeless"
    
    On Error GoTo EH
    EnableModeless = pvToIOleIPAO(This.IPAORealPtr).EnableModeless(fEnable)
    Exit Function
EH:
    If PrintError(FUNC_NAME & "(This.CtlName=" & This.CtlName & ")") = vbRetry Then
        Resume
    End If
    Resume Next
End Function
