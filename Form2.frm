VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "Form1"
   ClientHeight    =   4890
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   13710
   LinkTopic       =   "Form1"
   ScaleHeight     =   4890
   ScaleWidth      =   13710
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picbox1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4245
      Left            =   7920
      ScaleHeight     =   4245
      ScaleWidth      =   3555
      TabIndex        =   0
      Top             =   420
      Width           =   3555
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type BITMAPINFOHEADER
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

Private Type BITMAPINFO
    bmiHeader As BITMAPINFOHEADER
    bmiColors As Long
End Type

Private Type BLENDFUNCTION
    BlendOp As Byte
    BlendFlags As Byte
    SourceConstantAlpha As Byte
    AlphaFormat As Byte
End Type

Private Const BI_RGB = 0&
Private Const AC_SRC_OVER = 0
Private Const AC_SRC_ALPHA = 1

Private Declare Function AlphaBlend Lib "msimg32.dll" ( _
    ByVal hdcDest As Long, _
    ByVal nXOriginDest As Long, _
    ByVal nYOriginDest As Long, _
    ByVal nWidthDest As Long, _
    ByVal nHeightDest As Long, _
    ByVal hdcSrc As Long, _
    ByVal nXOriginSrc As Long, _
    ByVal nYOriginSrc As Long, _
    ByVal nWidthSrc As Long, _
    ByVal nHeightSrc As Long, _
    ByVal blendFunc As Long) As Long

Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function SetDIBits Lib "gdi32" (ByVal hDC As Long, ByVal hBitmap As Long, ByVal uStartScan As Long, ByVal cScanLines As Long, lpvBits As Any, lpbi As BITMAPINFO, ByVal fuColorUse As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long

'' === Cairo declarations ===
'Private Enum cairo_format_t
'    CAIRO_FORMAT_ARGB32 = 0
'    CAIRO_FORMAT_RGB24 = 1
'    CAIRO_FORMAT_A8 = 2
'    CAIRO_FORMAT_A1 = 3
'End Enum


'Private Declare Function cairo_image_surface_create Lib "E:\vb6\transparent-Cairo-VB6\VbCairo" (ByVal format As cairo_format_t, ByVal width As Long, ByVal height As Long) As Long
'Private Declare Function cairo_create Lib "E:\vb6\transparent-Cairo-VB6\VbCairo" (ByVal surface As Long) As Long
'Private Declare Sub cairo_destroy Lib "E:\vb6\transparent-Cairo-VB6\VbCairo" (ByVal cr As Long)
'Private Declare Sub cairo_surface_destroy Lib "E:\vb6\transparent-Cairo-VB6\VbCairo" (ByVal surface As Long)
'Private Declare Sub cairo_set_source_surface Lib "E:\vb6\transparent-Cairo-VB6\VbCairo" (ByVal cr As Long, ByVal surface As Long, ByVal x As Double, ByVal y As Double)
'Private Declare Sub cairo_paint Lib "E:\vb6\transparent-Cairo-VB6\VbCairo" (ByVal cr As Long)
'Private Declare Function cairo_image_surface_get_data Lib "E:\vb6\transparent-Cairo-VB6\VbCairo" (ByVal surface As Long) As Long
'Private Declare Function cairo_image_surface_get_width Lib "E:\vb6\transparent-Cairo-VB6\VbCairo" (ByVal surface As Long) As Long
'Private Declare Function cairo_image_surface_get_height Lib "E:\vb6\transparent-Cairo-VB6\VbCairo" (ByVal surface As Long) As Long
'Private Declare Function cairo_image_surface_create_from_png Lib "E:\vb6\transparent-Cairo-VB6\VbCairo" (ByVal filename As String) As Long

Private Sub Form_Load()
    Dim pngPath As String
    
    pngPath = App.Path & "\tardis.png"
    DrawAlphaPng Me.hWnd, pngPath, 100, 100
End Sub





Private Sub Form_Paint()
'    Dim pngPath As String
'    pngPath = App.Path & "\tardis.png"
'    DrawAlphaPng2 Me.hWnd, pngPath, 100, 100
End Sub

Private Sub DrawAlphaPng(ByVal hWnd As Long, ByVal sPngPath As String, ByVal x As Long, ByVal y As Long)
    Dim width As Long, height As Long
    Dim surfImg As Long, cr As Long, surfPng As Long
    Dim dataPtr As Long
    Dim bmi As BITMAPINFO
    Dim blend As BLENDFUNCTION
    Dim hDC As Long, memDC As Long, hBmp As Long, hOldBmp As Long
    

    hDC = Me.hDC
    
    'First we need a Surface, which is a "physical thing" (a real Render-Target)

    ' create a Cairo surface from Load PNG into Cairo surface
    surfPng = cairo_image_surface_create_from_png(sPngPath)
    If surfPng = 0 Then Exit Sub

    width = cairo_image_surface_get_width(surfPng)
    height = cairo_image_surface_get_height(surfPng)

    ' create an offscreen Cairo surface that supports alpha
    'surfImg = cairo_image_surface_create(CAIRO_FORMAT_ARGB32, width, height)
    surfImg = cairo_win32_surface_create(hDC)
    
    ' create a Cairo surface that writes directly to the picture box hardware device context for a PICBOX, note imageboxes do not have a .hDC
    ' a Cairo-Image-Surface is something like an allocated InMemory-Bitmap (a hDIB)
    'psfcFrm = cairo_win32_surface_create(hdcScreen) '  RichClient equivalent Set Srf = Cairo.CreateSurface(200, 100, ImageSurface) or GdipCreateFromHDC
    
    ' create a Cairo context for issuing drawing commands on the surface, in this case we aren't doing any drawing just painting using a PNG
    ' a context is something akin to a hDC in GDI unlike in GDI, where we would "Select" a Bitmap into a hDC first...
    ' with Cairo we can create such a Cairo context "anytime" from any Surface
    cr = cairo_create(surfImg)
        
    cairo_translate cr, 128#, 128#
    cairo_rotate cr, M_PI / 4
    cairo_scale cr, 1 / Sqr(2), 1 / Sqr(2)
    cairo_translate cr, -128#, -128#

    ' set the cairo context using the surface on the form at a defined position, in this case top/left
    cairo_set_source_surface cr, surfPng, 0, 0
        
    ' Paint the PNG (preserves alpha)
    cairo_paint cr

    ' Get pointer to pixel buffer
    dataPtr = cairo_image_surface_get_data(surfImg)

    ' Prepare GDI structures
    
    ' A handle to the Device Context (HDC) is obtained before output is written and then released after elements have been written.
    ' Get a device context compatible with the screen

    'hDC = Me.hDC
    memDC = CreateCompatibleDC(hDC)
    
    ' create a compatible bitmap and return a handle, bmpMemory, providing it a handle to device context allocated memory previously created with CreateCompatibleDC,
    ' providing size information in bmpInfo and setting any attributes to the new bitmap
    hBmp = CreateCompatibleBitmap(hDC, width, height)
    
    ' Make the device context use the bitmap.
    hOldBmp = SelectObject(memDC, hBmp)

    With bmi.bmiHeader
        .biSize = Len(bmi.bmiHeader)
        .biWidth = width
        .biHeight = -height
        .biPlanes = 1
        .biBitCount = 32
        .biCompression = BI_RGB
    End With

    ' Copy Cairo ARGB buffer into HBITMAP
    Call SetDIBits(memDC, hBmp, 0, height, ByVal dataPtr, bmi, 0)

    
    ' use the source image's alpha channel for blending characteristics for opacity
    blend.BlendOp = AC_SRC_OVER
    blend.SourceConstantAlpha = 255
    blend.AlphaFormat = AC_SRC_ALPHA
    
    ' Alpha blend onto window
    Call AlphaBlend(hDC, x, y, width, height, memDC, 0, 0, width, height, VarPtr(blend))

    ' delete temporary objects
    Call SelectObject(memDC, hOldBmp)
    Call DeleteObject(hBmp)
    Call DeleteDC(memDC)
    Call ReleaseDC(hWnd, hDC)

    ' tasks to tidy up, Cairo image, context and surface
    cairo_destroy cr
    cairo_surface_destroy surfImg
    cairo_surface_destroy surfPng
End Sub


