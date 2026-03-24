Attribute VB_Name = "Module1"
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
Public gImage As cfImageGDIP   ' << simple version (single image)

Public Function SubclassProc(ByVal hwnd As Long, ByVal Msg As Long, _
    ByVal wParam As Long, ByVal lParam As Long) As Long

    Dim x As Long, y As Long
    Dim lx As Single, ly As Single
    
    x = lParam And &HFFFF&
    y = (lParam \ &H10000) And &HFFFF&
    
    Select Case Msg
    
        Case WM_MOUSEMOVE
            If Not gImage Is Nothing Then
                If gImage.HitTest(x, y) Then
                    gImage.ScreenToLocal x, y, lx, ly
                    gImage.RaiseMouseMove 0, 0, lx, ly
                End If
            End If
            
        Case WM_LBUTTONDOWN
            If Not gImage Is Nothing Then
                If gImage.HitTest(x, y) Then
                    gImage.ScreenToLocal x, y, lx, ly
                    gImage.RaiseMouseDown 1, 0, lx, ly
                End If
            End If
            
    End Select

    SubclassProc = CallWindowProc(gPrevWndProc, hwnd, Msg, wParam, lParam)
End Function

