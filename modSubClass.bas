Attribute VB_Name = "modSubClass"
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
Public gImages As Collection


'---------------------------------------------------------------------------------------
' Procedure : SubclassProc
' Author    : beededea
' Date      : 25/03/2026
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Function SubclassProc(ByVal hwnd As Long, ByVal Msg As Long, _
    ByVal wParam As Long, ByVal lParam As Long) As Long

    Dim x As Long, y As Long
    Dim lx As Single, ly As Single
    
    On Error GoTo SubclassProc_Error

    x = lParam And &HFFFF&
    y = (lParam \ &H10000) And &HFFFF&
    
    Dim i As Long
    Dim img As cfImageGDIP
    
    Select Case Msg
    
        Case WM_MOUSEMOVE, WM_LBUTTONDOWN
        
            ' ?? iterate TOP ? BOTTOM
            For i = gImages.Count To 1 Step -1
                
                Set img = gImages(i)
                
                If img.HitTest(x, y) Then
                    
                    img.ScreenToLocal x, y, lx, ly
                    
                    If Msg = WM_MOUSEMOVE Then
                        img.RaiseMouseMove 0, 0, lx, ly
                    ElseIf Msg = WM_LBUTTONDOWN Then
                        img.RaiseMouseDown 1, 0, lx, ly
                    End If
                    
                    Exit For   ' ?? STOP ? event consumed
                    
                End If
                
            Next
            
    End Select

    SubclassProc = CallWindowProc(gPrevWndProc, hwnd, Msg, wParam, lParam)

    On Error GoTo 0
    Exit Function

SubclassProc_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure SubclassProc of Module Module1"

End Function

