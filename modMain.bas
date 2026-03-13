Attribute VB_Name = "modMain"
'---------------------------------------------------------------------------------------
' Module    : modMain
' Author    : beededea
' Date      : 12/03/2026
' Purpose   :
'---------------------------------------------------------------------------------------

' The Problem:
' ============
'
' When you use Cairo with an HDC target (via cairo_win32_surface_create(hDC)), the resulting surface is not alpha-aware.
' GDI does not support per-pixel alpha — only color values. So when you draw an image with semi-transparent pixels, Cairo
' composites them as if against a black background before blitting them onto your window — causing shadows to appear dark and non-transparent.

' The potential solution:
' =======================

' Cairo does all compositing in a true ARGB32 surface.
' You transfer that buffer to a GDI DIB section.
' AlphaBlend() performs proper per-pixel alpha blending with the window’s background.

' Status:
' =======

' I have a partially working example that can write (with defects) to the DC of a VB6 form or perfectly to the DC of the main (0) desktop window but it only lasts for a second or so before refreshing.
' but it works for that second!

' I need to persist with the original logic to see if I can make the cairo surface with RGB operate using RC6 versions throughout
' Then test the creation of a window with the above

' Tasks?
' ======

' The recommended call to cairo_image_surface_create does not work using other cairo DLLs, placing nothing on any hDC
' When I use vbCairo cairo_win32_surface_create (thisHDC) then a transparent image is placed directly on the device context as required
' meaning that all the rest of the code to place an image on a layered window is not even utilised.

' tried cairo_image_surface_create with any other cairo.dll whilst using vbCairo for all the rest
' tried cairo_image_surface_create with RC cairo.dll whilst using vbCairo for all the rest

' What I believe is happening is that the write to the hDc is working but the image will not persist on the hDC(0) as the explorer process regularly refreshes it.
' When writing to the VB6 form, the image persists but the vbCyan used as a key for transparency still is present as an outline artefact.

' What I think I have to do next:

' try replacing ALL the Cairo calls using RC6 versions to see if the logic as suggested by chatGPT is reasonable
' try an alternative AI model to test the logic

' In fact the cairo_win32_surface_create(thisHDC) function is a tool that I can use to test the writing of an image to a hDC regardless of the rest of the logic.
' It might be the method I use with own-created window.

' investigate operator_clear

' Next Step:
' ==========

' Create the window with CreateWindow or CreateWindowEx, obtain the handle and then the hDC and try writing to that instead.
' then subclass the form for wm_paint
' have a look here for pointers on some of this: https://www.vbforums.com/showthread.php?340850-Write-code-for-a-created-window-CreateWindowEx-RESOLVED

' replace the hwnd with one from the created window

' in Steamydock if we can replace the VB6 form that is made invisible with a user-created window as per APIwindow, then place the GDI+ created icons onto it then we have materially solved the capability to create
' programming 'widgets' on a user-created transparent window

' In the 10/- VB6 RC5 widget we have code to successfully create a hidden window

' you could try using GDIP but writing to hDC(0) would result in the same.

' When working we could get the reading the PNG data from an array operational.

' so far I stalled on reading the PNG data from a byte array from a dictionary using VB6 as the cairo_image_surface_create_from_png_stream function requires population via a callback
' it CAN be done using TB callbacks but is not essential to have this within the VB6 program, it is just a proof of concept. To continue, this will eventually be a single image widget demo.

' to read a byte array from a scripting collection and feed it to a Cairo surface is more involved than it looks
' as cairo_image_surface_create_from_png_stream requires a callback from the function that is a helper function that handles the data for Cairo
' RichClient does this using VB6 which I cannot replicate so using TB to do the same
    
'#If twinbasic Then
'    Dim surf As LongPtr
'    surf = CairoSurfaceFromPngBytes(dict("Tardis"))
'#End If

    ' The code is still here as trial for doing the same in TwinBasic
    
'#If twinbasic Then
'    ' dictionary for the PNG images
'    Set Surfaces = CreateObject("Scripting.Dictionary")
'    Surfaces.CompareMode = 1 ' for case-insenitive Key-Comparisons
'
'    ' Load PNG bytes into memory
'    PNGData = LoadFileToBytes(App.path & "\tardis.png")
'
'    'We need to learn how to create a scaled image using Cairo
'    ' HERE
'
'    ' Store PNG bytes in the dictionary
'    Surfaces.Add "Tardis", PNGData
'#End If

' Once done and working, use the alternative VBA scripting dictionary to replace the MS version.



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

Public Opacity As Single

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
Public hdcScreen As Long
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




'---------------------------------------------------------------------------------------
' Procedure : Main
' Author    : beededea
' Date      : 12/03/2026
' Purpose   : Creates a transparent form with one image and a close button
'             one of the images is generated using Cairo
'---------------------------------------------------------------------------------------
'
Public Sub Main()

    Dim rc As Long
    
    On Error GoTo Main_Error

    rc = createWindow("API Window in VB6", "VbWndClass")
    'MsgBox "Your window exited with code: " & rc
    
    On Error GoTo 0
    Exit Sub

Main_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Main of Module modMain"
End Sub



'Private Sub Command1_Click()
'    Me.Refresh
'End Sub




'---------------------------------------------------------------------------------------
' Procedure : monitorProperties
' Author    : beededea
' Date      : 06/11/2025
' Purpose   : All this subroutine does at the moment is to set the screenTwipsPerPixel
'---------------------------------------------------------------------------------------
'
Public Function monitorProperties()

    ' only calling TwipsPerPixelX/Y once on startup
    On Error GoTo monitorProperties_Error

    screenTwipsPerPixelX = fTwipsPerPixelX
    screenTwipsPerPixelY = fTwipsPerPixelY

    On Error GoTo 0
    Exit Function

monitorProperties_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure monitorProperties of Form frmMain"
    
End Function

'---------------------------------------------------------------------------------------
' Procedure : resolveVB6SizeBug
' Author    : beededea
' Date      : 18/09/2020
' Purpose   : VB6 has a bug - it should return 28800 on my screen but often returns 16200 when a game runs full screen, changing the resolution
'             the screen width determination is incorrect, the API call below resolves this.
'             NOTE: the dock program is the size of the whole screen
'---------------------------------------------------------------------------------------
'
Public Sub resolveVB6SizeBug()

   On Error GoTo resolveVB6SizeBug_Error
    
    ' pixels for Cairo and GDI
    screenHeightPixels = GetDeviceCaps(hdcScreen, VERTRES)
    screenWidthPixels = GetDeviceCaps(hdcScreen, HORZRES)
    
    'twips for VB6 forms and controls
    screenHeightTwips = screenHeightPixels * screenTwipsPerPixelY
    screenWidthTwips = screenWidthPixels * screenTwipsPerPixelX
    
   On Error GoTo 0
   Exit Sub

resolveVB6SizeBug_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure resolveVB6SizeBug of Form dock"
End Sub
'---------------------------------------------------------------------------------------
' Procedure : setWindowCharacteristics
' Author    : beededea
' Date      : 07/04/2020
' Purpose   : update some characteristics for the window we will be updating using UpdateLayeredWindow API
'---------------------------------------------------------------------------------------
'
Public Sub setWindowCharacteristics()

    On Error GoTo setWindowCharacteristics_Error
    
    'set the transparency of the underlying VB6 form with full click through, makes the form completely transparent
    'SetWindowLong hWindowHwnd, GWL_EXSTYLE, GetWindowLong(hWindowHwnd, GWL_EXSTYLE) Or WS_EX_LAYERED ' Or WS_EX_TRANSPARENT
    ' the addition of "Or WS_EX_TRANSPARENT" to SetWindowLong will make the transparent form fully click-through but controls will be unresponsive
     
    ' this brings the form back again and uses the colour key cyan to make the form and any other similar items appear transparent
    'Me.BackColor = vbCyan ' sets the VB6 form to the transparent key colour
    'SetLayeredWindowAttributes hWindowHwnd, vbCyan, 0&, LWA_COLORKEY

    ' Position over desktop
    'Me.Move 0, 0, screenWidthTwips, screenHeightTwips
    
    ' UpdateLayeredWindow structures
    
    ' point structure that specifies the location of the layer updated in UpdateLayeredWindow
    apiPoint.x = 0
    apiPoint.y = 0
    
    ' point structure that specifies the size of the window in pixels
    windowSize.x = screenWidthPixels ' .59 DAEB 26/04/2021 frmMain.frm changed to use pixels alone, removed all unnecessary twip conversion
    windowSize.y = screenHeightPixels  ' .59 DAEB 26/04/2021 frmMain.frm changed to use pixels alone, removed all unnecessary twip conversion
    
    ' use the source image's alpha channel for blending characteristics, opacity &c for use in UpdateLayeredWindow later In SD it is used im in setWindowCharacteristics
    funcBlend32bpp.BlendOp = AC_SRC_OVER
    funcBlend32bpp.SourceConstantAlpha = 255 * Opacity ' set the opacity of the whole bitmap, used to display solidly and for instant autohide
    funcBlend32bpp.AlphaFormat = AC_SRC_ALPHA
    
   On Error GoTo 0
   Exit Sub

setWindowCharacteristics_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure setWindowCharacteristics of module mdlMain.bas"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : createGDIStructures
' Author    : beededea
' Date      : 05/11/2025
' Purpose   : sets bmpInfo object to create a bitmap of the whole screen size and get a handle to the Device Context
'---------------------------------------------------------------------------------------
'
Private Sub createGDIStructures()

    ' sets the bmpInfo object containing data to create a bitmap the whole screen size
    ' used later when creating DIB section of the correct size, width &c
    
    On Error GoTo createGDIStructures_Error

    ' Set the bitmap characteristics for use in SetDIBits later
    With bmpInfo.bmpHeader
        .biSize = Len(bmpInfo.bmpHeader)
        .biWidth = windowSize.x
        .biHeight = -windowSize.y
        .biPlanes = 1
        .biBitCount = 32
        .biCompression = BI_RGB
    End With
    
    ' A device context is a generalized rendering abstraction. It serves as a proxy between your rendering code and the output device.
    ' It allows you to use the same rendering code regardless of the destination; the low-level details are handled for you,
    ' dependant on the output device, including clipping, scaling, and viewport translation.
    
    ' A handle to the Device Context (HDC) is obtained before output is written and then released after elements have been written.
    dcMemory = CreateCompatibleDC(hdcScreen)

    On Error GoTo 0
    Exit Sub

createGDIStructures_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure createGDIStructures of Form frmMain"

End Sub

'---------------------------------------------------------------------------------------
' Procedure : createNewGDIBitmap in Steamydock used within createNewGDIPBitmap
' Author    : beededea
' Date      : 05/11/2025
' Purpose   : Create a gdi bitmap with width and height of what we are going to draw into it. This is the entire drawing area for everything,
'             creating a bitmap in memory that our VB6/GDI application writes to directly. Called each animation interval.
'---------------------------------------------------------------------------------------
'
Private Sub createNewGDIBitmap()

    On Error GoTo createNewGDIBitmap_Error

    ' the existing bitmap deleted
    Call DeleteObject(hBmpMemory) '
    
    ' create a compatible bitmap DDB and return a handle, bmpMemory, providing it a handle to device context dcMemory, allocated memory previously created with CreateCompatibleDC,
    ' providing size information
    hBmpMemory = CreateCompatibleBitmap(hdcScreen, windowSize.x, windowSize.y) ' in SD uses CreateDIBSection within createNewGDIPBitmap
    
    ' Make the device context dcMemory use the bitmap.  hOldBmp is a return value giving a handle which determines success and allows reverting later to release GDI handles
    hOldBmp = SelectObject(dcMemory, hBmpMemory) ' releases memory used by any open GDI handle  in SD used within createNewGDIPBitmap
    
    On Error GoTo 0
    Exit Sub

createNewGDIBitmap_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure createNewGDIBitmap of Form frmMain"
    
End Sub

'---------------------------------------------------------------------------------------
' Procedure : drawAlphaPngCairo
' Author    : beededea
' Date      : 31/10/2025
' Purpose   : draw a PNG image with a transparent background using Cairo only, no GDI+
'---------------------------------------------------------------------------------------
'
Public Sub drawAlphaPngCairo(thisHDC As Long, ByVal hwnd As Long, ByVal sPngPath As String, ByVal x As Long, ByVal y As Long)
    Dim Width As Long
    Dim height As Long

    On Error GoTo drawAlphaPngCairo_Error
    
    Opacity = 0.5

    ' create a Cairo surface from an on-disc PNG
    surfPng = cairo_image_surface_create_from_png(sPngPath)
    If surfPng = 0 Then Exit Sub

    Width = cairo_image_surface_get_width(surfPng)
    height = cairo_image_surface_get_height(surfPng)

    ' Draw the PNG on an intermediate offscreen ARGB Cairo surface that supports alpha (offscreen buffer).

    ' a Cairo-Image-Surface is something like an allocated InMemory-Bitmap (a hDIB)
    ' surfImg = cairo_image_surface_create(CAIRO_FORMAT_ARGB32, width, height)
    
    ' ^^^^^^^^^^^^^^^^^^^^^^^^
    ' the above Cairo function does not work and I believe this is the reason why the logic overall does not work
    
    ' create a Cairo surface that writes directly to a hardware device context
    surfImg = cairo_win32_surface_create(thisHDC) '  RichClient equivalent Set Srf = Cairo.CreateSurface(200, 100, ImageSurface) or GdipCreateFromHDC
    ' although this places a surface on the current Dc I think this is incorrect, should be using cairo_image_surface_create from RC
    
    ' create a Cairo context for issuing drawing commands on the surface, in this case we aren't doing any drawing just painting using a PNG
    ' a context is something akin to a hDC in GDI unlike in GDI, where we would Select a Bitmap into a hDC first.
    ' we can create such a Cairo context anytime from any Surface
    cr = cairo_create(surfImg)
        
    cairo_select_font_face cr, "segoe", CAIRO_FONT_SLANT_NORMAL, CAIRO_FONT_WEIGHT_BOLD
    cairo_set_font_size cr, 32#
    cairo_set_source_rgba cr, 0#, 0#, 1#, 0.3
    cairo_move_to cr, 10#, 50#
    cairo_show_text cr, "TARDIS"

    cairo_translate cr, 128#, 128#
    cairo_rotate cr, M_PI / 4
    cairo_scale cr, 1 / Sqr(2), 1 / Sqr(2)
    cairo_translate cr, -128#, -128#
'
'    ' set the cairo context using the surface on the form at a defined position, in this case top/left
    cairo_set_source_surface cr, surfPng, 0, 0

    ' Paint the PNG (preserves alpha)
    cairo_paint_with_alpha cr, Opacity '   CC.Paint with alpha
    'cairo_paint cr

    ' Get pointer to pixel buffer
    'dataPtrPixelBuffer = cairo_image_surface_get_data(surfImg)

    ' Copy Cairo ARGB pixel buffer into HBITMAP compatible DDB hBmpMemory (usually has better GDI performance than a DIB as used in Steamydock)
    'Call SetDIBits(dcMemory, hBmpMemory, 0, height, ByVal dataPtrPixelBuffer, bmpInfo, 0) '*  in SD, an equivalent of GdipCreateFromHDC used within createNewGDIPBitmap?
    
    ' tasks to tidy up, Cairo image, context and surface
    cairo_destroy cr
    cairo_surface_destroy surfImg
    cairo_surface_destroy surfPng

    On Error GoTo 0
    Exit Sub

drawAlphaPngCairo_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure drawAlphaPngCairo of Form frmMain"
End Sub


'---------------------------------------------------------------------------------------
' Procedure : drawAlphaPngGDIP
' Author    : beededea
' Date      : 31/10/2025
' Purpose   : draw a PNG image with a transparent background using GDI+
'---------------------------------------------------------------------------------------
'
Public Sub drawAlphaPngGDIP(thisHDC As Long, ByVal hwnd As Long, ByVal sPngPath As String, ByVal x As Long, ByVal y As Long)
    Dim Width As Long
    Dim height As Long

    On Error GoTo drawAlphaPngGDIP_Error
    
    Opacity = 0.5

    ' create a Cairo surface from an on-disc PNG
'    surfPng = cairo_image_surface_create_from_png(sPngPath)
'    If surfPng = 0 Then Exit Sub
'
'    width = cairo_image_surface_get_width(surfPng)
'    height = cairo_image_surface_get_height(surfPng)
'
'    ' Draw the PNG on an intermediate offscreen ARGB Cairo surface that supports alpha (offscreen buffer).
'
'    ' a Cairo-Image-Surface is something like an allocated InMemory-Bitmap (a hDIB)
'    ' surfImg = cairo_image_surface_create(CAIRO_FORMAT_ARGB32, width, height)
'
'    ' ^^^^^^^^^^^^^^^^^^^^^^^^
'    ' the above Cairo function does not work and I believe this is the reason why the logic overall does not work
'
'    ' create a Cairo surface that writes directly to a hardware device context
'    surfImg = cairo_win32_surface_create(thisHDC) '  RichClient equivalent Set Srf = Cairo.CreateSurface(200, 100, ImageSurface) or GdipCreateFromHDC
'    ' although this places a surface on the current Dc I think this is incorrect, should be using cairo_image_surface_create from RC
'
'    ' create a Cairo context for issuing drawing commands on the surface, in this case we aren't doing any drawing just painting using a PNG
'    ' a context is something akin to a hDC in GDI unlike in GDI, where we would Select a Bitmap into a hDC first.
'    ' we can create such a Cairo context anytime from any Surface
'    cr = cairo_create(surfImg)
'
'    cairo_select_font_face cr, "segoe", CAIRO_FONT_SLANT_NORMAL, CAIRO_FONT_WEIGHT_BOLD
'    cairo_set_font_size cr, 32#
'    cairo_set_source_rgba cr, 0#, 0#, 1#, 0.3
'    cairo_move_to cr, 10#, 50#
'    cairo_show_text cr, "TARDIS"
'
'    cairo_translate cr, 128#, 128#
'    cairo_rotate cr, M_PI / 4
'    cairo_scale cr, 1 / Sqr(2), 1 / Sqr(2)
'    cairo_translate cr, -128#, -128#
''
''    ' set the cairo context using the surface on the form at a defined position, in this case top/left
'    cairo_set_source_surface cr, surfPng, 0, 0
'
'    ' Paint the PNG (preserves alpha)
'    cairo_paint_with_alpha cr, opacity '   CC.Paint with alpha
'    'cairo_paint cr
'
'    ' Get pointer to pixel buffer
'    'dataPtrPixelBuffer = cairo_image_surface_get_data(surfImg)
'
'    ' Copy Cairo ARGB pixel buffer into HBITMAP compatible DDB hBmpMemory (usually has better GDI performance than a DIB as used in Steamydock)
'    'Call SetDIBits(dcMemory, hBmpMemory, 0, height, ByVal dataPtrPixelBuffer, bmpInfo, 0) '*  in SD, an equivalent of GdipCreateFromHDC used within createNewGDIPBitmap?
'
'    ' tasks to tidy up, Cairo image, context and surface
'    cairo_destroy cr
'    cairo_surface_destroy surfImg
'    cairo_surface_destroy surfPng

    On Error GoTo 0
    Exit Sub

drawAlphaPngGDIP_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure drawAlphaPngGDIP of Form frmMain"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : UpdateLayeredWindowUsingGDIBitmap
' Author    : beededea
' Date      : 05/11/2025
' Purpose   : Calls UpdateLayeredWindow with created GDI bitmap
'---------------------------------------------------------------------------------------
'
Private Sub UpdateLayeredWindowUsingGDIBitmap()
    
    On Error GoTo UpdateLayeredWindowUsingGDIBitmap_Error
    
    ' We can use either AlphaBlend or UpdateLayeredWindow to write the image to the Window, alphaBlend is slower and thus can flicker
    ' Using UpdateLayeredWindow it is handled by the Windows compositor, which can take advantage of hardware acceleration for blending and movement.
        
    'blit the buffer to the window’s HDC with per-pixel alpha blending.
'    Call AlphaBlend(hdcScreen, 100, 100, 1000, 1000, hdcScreen, 0, 0, 1000, 1000, VarPtr(funcBlend32bpp))

    ' the third parameter to UpdateLayeredWindow is a pointer to a structure that specifies the new screen position of the layered window.
    ' If the current position is not changing, pptDst can be NULL. It is null.
            
    'Update the specified window handle (hwnd) with a handle to our bitmap (dc) passing all the required characteristics
    Call UpdateLayeredWindow(hWindowHwnd, hdcScreen, ByVal 0&, windowSize, dcMemory, apiPoint, 0, VarPtr(funcBlend32bpp), ULW_ALPHA)   '*  in SD called whenever a draw is required

    ' releases memory for GDI handles
    Call SelectObject(dcMemory, hOldBmp)
'    ' the existing bitmap deleted
    Call DeleteObject(hBmpMemory) '
    DeleteDC dcMemory
    'ReleaseDC 0, hdcScreen

    On Error GoTo 0
    Exit Sub

UpdateLayeredWindowUsingGDIBitmap_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure UpdateLayeredWindowUsingGDIBitmap of Form frmMain"

End Sub
'---------------------------------------------------------------------------------------
' Procedure : fTwipsPerPixelX
' Author    : Elroy from Vbforums
' Date      : 23/01/2021
' Purpose   : This works even on Tablet PC.  The problem is: when the tablet screen is rotated, the "Screen" object of VB doesn't pick it up.
'---------------------------------------------------------------------------------------
'
Public Function fTwipsPerPixelX() As Single
    Dim hDC As Long: hDC = 0
    Dim lPixelsPerInch As Long: lPixelsPerInch = 0
    
    Const LOGPIXELSX = 88              '  Logical pixels/inch in X
    Const POINTS_PER_INCH As Long = 72 ' A point is defined as 1/72 inches.
    Const TWIPS_PER_POINT As Long = 20 ' Also, by definition.
    '
    On Error GoTo fTwipsPerPixelX_Error
    
    ' 23/01/2021 .01 monitorModule.bas DAEB added if then else if you can't get device context
    hDC = GetDC(0)
    If hDC <> 0 Then
        lPixelsPerInch = GetDeviceCaps(hDC, LOGPIXELSX)
        ReleaseDC 0, hDC
        fTwipsPerPixelX = TWIPS_PER_POINT * (POINTS_PER_INCH / lPixelsPerInch) ' Cancel units to see it.
    Else
        fTwipsPerPixelX = Screen.TwipsPerPixelX
    End If

   On Error GoTo 0
   Exit Function

fTwipsPerPixelX_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure fTwipsPerPixelX of Module Module1"
End Function


'---------------------------------------------------------------------------------------
' Procedure : fTwipsPerPixelY
' Author    : Elroy from Vbforums
' Date      : 23/01/2021
' Purpose   : This works even on Tablet PC.  The problem is: when the tablet screen is rotated, the "Screen" object of VB doesn't pick it up.
'---------------------------------------------------------------------------------------
'
Public Function fTwipsPerPixelY() As Single
    Dim hDC As Long: hDC = 0
    Dim lPixelsPerInch As Long: lPixelsPerInch = 0
    
    Const LOGPIXELSY = 90        '  Logical pixels/inch in Y
    Const POINTS_PER_INCH As Long = 72 ' A point is defined as 1/72 inches.
    Const TWIPS_PER_POINT As Long = 20 ' Also, by definition.
    
    On Error GoTo fTwipsPerPixelY_Error
   
    ' 23/01/2021 .01 monitorModule.bas DAEB added if then else if you can't get device context
    hDC = GetDC(0)
    If hDC <> 0 Then
        lPixelsPerInch = GetDeviceCaps(hDC, LOGPIXELSY)
        ReleaseDC 0, hDC
        fTwipsPerPixelY = TWIPS_PER_POINT * (POINTS_PER_INCH / lPixelsPerInch) ' Cancel units to see it.
    Else
        fTwipsPerPixelY = Screen.TwipsPerPixelY
    End If

   On Error GoTo 0
   Exit Function

fTwipsPerPixelY_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure fTwipsPerPixelY of Module Module1"

End Function


'Private Sub Form_Paint()
'    drawAlphaPngCairo hdcScreen, hWindowHwnd, App.Path & "\tardis.png", 20, 20
'
'End Sub












''---------------------------------------------------------------------------------------
'' Procedure : Form_MouseDown
'' Author    : beededea
'' Date      : 14/08/2023
'' Purpose   : standard form down event to generate the menu across the board
''---------------------------------------------------------------------------------------
''
'Private Sub Form_MouseDown(ByRef Button As Integer, ByRef Shift As Integer, ByRef x As Single, ByRef y As Single)
'   On Error GoTo Form_MouseDown_Error
'
'    If Button = 2 Then
'        Me.PopupMenu mnuPopupMenu, vbPopupMenuRightButton
'    End If
'
'   On Error GoTo 0
'   Exit Sub
'
'Form_MouseDown_Error:
'
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Form_MouseDown of Form widgetPrefs"
'End Sub

'---------------------------------------------------------------------------------------
' Procedure : mnuClose_Click
' Author    : beededea
' Date      : 27/10/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
'Private Sub mnuClose_Click()
'    On Error GoTo mnuClose_Click_Error
'
'    End
'
'    On Error GoTo 0
'    Exit Sub
'
'mnuClose_Click_Error:
'
'     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuClose_Click of Form frmMain"
'End Sub

'---------------------------------------------------------------------------------------
' Procedure : Command_Click
' Author    : beededea
' Date      : 27/10/2025
' Purpose   : closes the form
'---------------------------------------------------------------------------------------
'
'Private Sub Command_Click()
'    On Error GoTo Command_Click_Error
'
'    ' delete temporary objects
'
'    Call SelectObject(dcMemory, hOldBmp) ' releases memory used by any open GDI handles
'    Call DeleteObject(hBmpMemory)
'    Call DeleteDC(dcMemory)
'    Call ReleaseDC(hWindowHwnd, dcMemory)
'
'    End
'
'    On Error GoTo 0
'    Exit Sub
'
'Command_Click_Error:
'
'     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Command_Click of Form frmMain"
'End Sub




'---------------------------------------------------------------------------------------
' Procedure : MouseDownHandler
' Author    : beededea
' Date      : 27/10/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
'Private Sub MouseDownHandler(ByVal ctrlName As String, Optional Button As Integer, Optional Shift As Integer, Optional x As Single, Optional y As Single)
'    On Error GoTo MouseDownHandler_Error
'
'    If Button = 2 Then
'        Me.PopupMenu mnuPopupMenu, vbPopupMenuRightButton
'    Else
'        'MsgBox ctrlName
'        ReleaseCapture
'        SendMessage hWindowHwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
'    End If
'
'    On Error GoTo 0
'    Exit Sub
'
'MouseDownHandler_Error:
'
'     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure MouseDownHandler of Form frmMain"
'End Sub

'---------------------------------------------------------------------------------------
' Procedure : DblClickHandler
' Author    : beededea
' Date      : 27/10/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
'Private Sub DblClickHandler(ByVal ctrlName As String)
'    On Error GoTo DblClickHandler_Error
'
'    MsgBox ctrlName & " double-clicked!"
'
'    On Error GoTo 0
'    Exit Sub
'
'DblClickHandler_Error:
'
'     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure DblClickHandler of Form frmMain"
'End Sub





'#If twinbasic Then
'
'    ' --- The callback prototype ---
'    Private Function CairoPngReadFunc( _
'        ByVal closure As LongPtr, _
'        ByVal data As LongPtr, _
'        ByVal Length As Long) As Long
'
'        Dim ctx As PngReadContext
'        CopyMemory ByVal VarPtr(ctx), ByVal closure, LenB(ctx)
'
'        Dim remaining As Long
'        remaining = ctx.DataSize - ctx.Position
'        If remaining < Length Then Length = remaining
'
'        ' Copy the next chunk of bytes into Cairo's buffer
'        CopyMemory ByVal data, ByVal (ctx.dataPtrPixelBuffer + ctx.Position), Length
'        ctx.Position = ctx.Position + Length
'
'        ' Write updated ctx back
'        CopyMemory ByVal closure, ByVal VarPtr(ctx), LenB(ctx)
'
'        CairoPngReadFunc = CAIRO_STATUS_SUCCESS
'    End Function
'
'
'    Public Sub Example_CreateCairoSurfaceFromDictionary()
'        ' --- Create and populate dictionary ---
'        Dim dict As Object
'        Set dict = CreateObject("Scripting.Dictionary")
'
'        Dim pngBytes() As Byte
'        pngBytes = LoadFileToBytes("C:\example.png")
'
'        dict.Add "Logo", pngBytes
'
'        ' --- Build read context ---
'        Dim ctx As PngReadContext
'        ctx.dataPtrPixelBuffer = VarPtr(dict("Logo")(0))
'        ctx.DataSize = UBound(dict("Logo")) + 1
'        ctx.Position = 0
'
'        ' --- Create surface from memory ---
'        Dim surf As LongPtr
'        surf = cairo_image_surface_create_from_png_stream(AddressOf CairoPngReadFunc, VarPtr(ctx))
'
'        If cairo_surface_status(surf) = CAIRO_STATUS_SUCCESS Then
'            MsgBox "Cairo surface created successfully!"
'        Else
'            MsgBox "Failed to create Cairo surface"
'        End If
'    End Sub
'
'    Private Function LoadFileToBytes(ByVal path As String) As Byte()
'        Dim f As Integer
'        f = FreeFile
'        Open path For Binary As #f
'        ReDim LoadFileToBytes(LOF(f) - 1)
'        Get #f, , LoadFileToBytes
'        Close #f
'    End Function
'
'
'    Public Function CairoSurfaceFromPngBytes(PNGData() As Byte) As LongPtr
'        Dim ctx As PngReadContext
'        ctx.dataPtrPixelBuffer = VarPtr(PNGData(0))
'        ctx.DataSize = UBound(PNGData) + 1
'        ctx.Position = 0
'        CairoSurfaceFromPngBytes = cairo_image_surface_create_from_png_stream(AddressOf CairoPngReadFunc, VarPtr(ctx))
'    End Function
'
'#End If


'---------------------------------------------------------------------------------------
' Procedure : Timer_Timer
' Author    : beededea
' Date      : 07/11/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
'Private Sub Timer_Timer()
'    'Me.Refresh
'
'    ' Create a GDI bitmap with width and height of what we are going to draw into it
'    On Error GoTo Timer_Timer_Error
'
'    'Call createNewGDIBitmap
'
'    ' that's the native VB6/TwinBasic stuff done, now we play with Cairo
'    Call drawAlphaPngCairo(hdcScreen, Me.hwnd, App.Path & "\tardis.png", 20, 20)
'
'    'Call UpdateLayeredWindowUsingGDIBitmap
'
'    On Error GoTo 0
'    Exit Sub
'
'Timer_Timer_Error:
'
'     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Timer_Timer of Form frmMain"
'End Sub



