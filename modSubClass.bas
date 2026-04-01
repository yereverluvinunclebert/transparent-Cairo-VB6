Attribute VB_Name = "modSubClass"
'---------------------------------------------------------------------------------------
' Module    : modSubClass
' Author    : beededea
' Date      : 27/03/2026
' Purpose   : This module sub classes Form1 to intercept mouse-move and mouse-down messages, performs hit-tests and raises events.
'---------------------------------------------------------------------------------------

Option Explicit

Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" _
    (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" _
    (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, _
     ByVal wParam As Long, ByVal lParam As Long) As Long

Public Const GWL_WNDPROC = -4

Public Const WM_MOUSEMOVE = &H200
Public Const WM_LBUTTONDOWN = &H201

Public gPrevWndProc As Long
Public gFormHwnd As Long
Public gHitTestCollection As Collection


'---------------------------------------------------------------------------------------
' Procedure : SubclassProc
' Author    : beededea
' Date      : 25/03/2026
' Purpose   : Our subclassed event dispatcher.
'             In the absence of real VB6 controls we have to replicate the functionality of determining the bounds of
'             an image that will become a UI control. Our (invisible) form is sub-classed to intercept mouse-move and mouse-down
'             messages, testing for screen vs local form position against a stored image in the hit test collection.
'             If it is within the bounds then raise an event via the eventHost event capture sink.
'---------------------------------------------------------------------------------------
'
Public Function SubclassProc(ByVal hwnd As Long, ByVal Msg As Long, _
    ByVal wParam As Long, ByVal lParam As Long) As Long

    Dim x As Long, y As Long
    Dim lx As Single, ly As Single
    Dim i As Long
    Dim img As cfImageGDIP
    
    On Error GoTo SubclassProc_Error

    x = lParam And &HFFFF&
    y = (lParam \ &H10000) And &HFFFF&
    
    Select Case Msg
        Case WM_MOUSEMOVE, WM_LBUTTONDOWN
        
            ' iterate TOP to BOTTOM through the hit test collection
            For i = gHitTestCollection.Count To 1 Step -1
                Set img = gHitTestCollection(i)
                If img.HitTest(x, y) Then
                    img.ScreenToLocal x, y, lx, ly
                    If Msg = WM_MOUSEMOVE Then
                        img.RaiseMouseMove 0, 0, lx, ly
                    ElseIf Msg = WM_LBUTTONDOWN Then
                        img.RaiseMouseDown 1, 0, lx, ly
                    End If
                    Exit For   ' STOP event consumed
                End If
            Next
    End Select

    SubclassProc = CallWindowProc(gPrevWndProc, hwnd, Msg, wParam, lParam)

    On Error GoTo 0
    Exit Function

SubclassProc_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure SubclassProc of Module Module1"

End Function

