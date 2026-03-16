Attribute VB_Name = "modAPIDeclarations"
Option Explicit

' vars to obtain correct screen width (to correct VB6 bug) STARTS
Public Const HORZRES = 8
Public Const VERTRES = 10
Public screenTwipsPerPixelX As Long
Public screenTwipsPerPixelY As Long
Public screenWidthTwips As Long
Public screenHeightTwips As Long
Public screenWidthPixels As Long
Public screenHeightPixels As Long

Public gSngOpacity As Single

' functions from user32 to get/set Window characteristics , opacity &c
Public Declare Function GetWindowLong Lib "user32.dll" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function SetLayeredWindowAttributes Lib "user32" ( _
                ByVal hwnd As Long, _
                ByVal crKey As Long, _
                ByVal bAlpha As Byte, _
                ByVal dwFlags As Long) As Long
Public Declare Function UpdateLayeredWindow Lib "user32.dll" (ByVal hwnd As Long, ByVal hdcDst As Long, pptDst As Any, psize As Any, ByVal hdcSrc As Long, pptSrc As Any, crKey As Long, ByVal pblend As Long, ByVal dwFlags As Long) As Long

'public Declare Function AlphaBlend Lib "msimg32.dll" ( _
'    ByVal hdcDest As Long, _
'    ByVal nXOriginDest As Long, _
'    ByVal nYOriginDest As Long, _
'    ByVal nWidthDest As Long, _
'    ByVal nHeightDest As Long, _
'    ByVal hdcSrc As Long, _
'    ByVal nXOriginSrc As Long, _
'    ByVal nYOriginSrc As Long, _
'    ByVal nWidthSrc As Long, _
'    ByVal nHeightSrc As Long, _
'    ByVal blendFunc As Long) As Long

' vars for the above APIs to get/set Window characteristics , opacity &c
Public Const ULW_ALPHA = &H2
Public Const WS_EX_LAYERED = &H80000
Public Const WS_EX_TRANSPARENT As Integer = &H20
Public Const GWL_EXSTYLE As Long = -20
Public Const LWA_COLORKEY = &H1

Public Type POINTAPI
    x As Long
    y As Long
End Type

Public windowSize As POINTAPI
Public apiPoint As POINTAPI

' functions from user32 to capture keydown on a user control to move a form with no visible title bar
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

' vars for the above APIs to capture keydown on a user control to move a form with no visible title bar
Public Const WM_NCLBUTTONDOWN = &HA1
Public Const HTCAPTION = 2

'' TwinBasic only functionality to read rawbyte PNG data from a scripting dictionary into
'public Enum LongPtr
'    [_]
'End Enum
'
''#If twinbasic Then
'    'public Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
'    public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
'        ByVal Destination As LongPtr, _
'        ByVal Source As LongPtr, _
'        ByVal Length As Long)
'
'   ' --- Cairo constants ---
'    'public Const CAIRO_STATUS_SUCCESS As Long = 0
'    '
'    ' --- Our context type for reading PNG data ---
'    public Type PngReadContext
'        dataPtrPixelBuffer As LongPtr
'        DataSize As Long
'        Position As Long
'    End Type
''#End If

' functions from GDI to transfer an image from Cairo (in this case) to the Windows desktop or a window
Public Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long
Public Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Public Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
'public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
'public Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function SetDIBits Lib "gdi32" (ByVal hDC As Long, ByVal hBitmap As Long, ByVal uStartScan As Long, ByVal cScanLines As Long, lpvBits As Any, lpbi As BITMAPINFO, ByVal fuColorUse As Long) As Long
'public Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Public Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long

' vars for the functions above to transfer an image from Cairo (in this case) to the Windows desktop or a window
Public dcMemory As Long
Public hBmpMemory As Long
Public thisHDC As Long
Public hOldBmp As Long
Public dataPtrPixelBuffer As Long

Public cr As Long
Public surfImg As Long
Public surfPng As Long

Public bmpInfo As BITMAPINFO
Public funcBlend32bpp As BLENDFUNCTION

Public Type BITMAPINFOHEADER
    biSize As Long
    biWidth As Long
    biHeight As Long
    biPlanes As Integer
    biBitCount As Integer
    biCompression As Long
    biSizeImage As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed As Long
    biClrImportant As Long
End Type

Public Type BITMAPINFO
    bmpHeader As BITMAPINFOHEADER
    bmiColors As Long
End Type

Public Type BLENDFUNCTION
    BlendOp As Byte
    BlendFlags As Byte
    SourceConstantAlpha As Byte
    AlphaFormat As Byte
End Type

Public Const BI_RGB = 0&
Public Const AC_SRC_OVER = 0
Public Const AC_SRC_ALPHA = 1

' vars for Cairo, unused due to something not operating quite as expected
' public Const CAIRO_FORMAT_ARGB32 = 0&

Public Enum cairo_format_t
    CAIRO_FORMAT_ARGB32 = 0
    CAIRO_FORMAT_RGB24 = 1
    CAIRO_FORMAT_A8 = 2
    CAIRO_FORMAT_A1 = 3
End Enum

' API declaration testing showing how to call a Cairo function embedded within Olaf's RC so I can test and compare same vbCairo and RC functions
' this crashes
' public Declare Function cairo_image_surface_create Lib "E:\vb6\transparent-Cairo-VB6\vb_cairo_sqlite.dll" (ByVal format As cairo_format_t, ByVal width As Long, ByVal height As Long) As Long

' this does not crsh but it does not function
'public Declare Function cairo_image_surface_create Lib "E:\vb6\transparent-Cairo-VB6\vbcairo.dll" (ByVal format As cairo_format_t, ByVal width As Long, ByVal height As Long) As Long



'general vars

'    public Declare Function AlphaBlend Lib "msimg32.dll" ( _
'        ByVal hdcDest As Long, _
'        ByVal nXOriginDest As Long, _
'        ByVal nYOriginDest As Long, _
'        ByVal nWidthDest As Long, _
'        ByVal nHeightDest As Long, _
'        ByVal hdcSrc As Long, _
'        ByVal nXOriginSrc As Long, _
'        ByVal nYOriginSrc As Long, _
'        ByVal nWidthSrc As Long, _
'        ByVal nHeightSrc As Long, _
'        ByVal blendFunc As Long) As Long

Public Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Public Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hDC As Long) As Long

Public Declare Function GdipCreateFromHDC Lib "GdiPlus.dll" (ByVal hDC As Long, GpGraphics As Long) As Long
Public Declare Function GdipDrawImageRectI Lib "GDIPlus" _
    (ByVal Graphics As Long, _
     ByVal image As Long, _
     ByVal x As Long, _
     ByVal y As Long, _
     ByVal Width As Long, _
     ByVal height As Long) As Long
     
Public Declare Function GdipDeleteGraphics Lib "GdiPlus.dll" (ByVal Graphics As Long) As Long

Public imageBitmap As Long
Public gdipFullScreenBitmap As Long

