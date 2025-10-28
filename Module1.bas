Attribute VB_Name = "Module1"
Public Declare Function GetWindowLong Lib "user32.dll" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function SetLayeredWindowAttributes Lib "user32" ( _
                ByVal hwnd As Long, _
                ByVal crKey As Long, _
                ByVal bAlpha As Byte, _
                ByVal dwFlags As Long) As Long
Public Declare Function UpdateLayeredWindow Lib "user32.dll" ( _
                         ByVal hwnd As Long, _
                         ByVal hdcDst As Long, _
                         ByRef pptDst As Any, _
                         ByRef psize As Any, _
                         ByVal hdcSrc As Long, _
                         ByRef pptSrc As Any, _
                         ByVal crKey As Long, _
                         ByRef pblend As Long, _
                         ByVal dwFlags As Long) As Long

' global Windows+ constants START
Public Const ULW_ALPHA = &H2
Public Const WS_EX_LAYERED = &H80000
Public Const WS_EX_TRANSPARENT As Integer = &H20
Public Const GWL_EXSTYLE As Long = -20
Public Const LWA_COLORKEY = &H1
Public Const AB_32Bpp255     As Long = 33488896
Public Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByRef lParam As Any) As Long
Public Declare Function ReleaseCapture Lib "user32.dll" () As Long
Public Const WM_NCLBUTTONDOWN As Long = &HA1

