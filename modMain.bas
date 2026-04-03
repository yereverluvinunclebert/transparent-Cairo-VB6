Attribute VB_Name = "modMain"
'---------------------------------------------------------------------------------------
' Module    : modMain
' Author    : beededea
' Date      : 12/03/2026
' Purpose   : Test program to write images to a transparent form using Cairo and GDI+
'---------------------------------------------------------------------------------------



Option Explicit


  
'---------------------------------------------------------------------------------------
' Procedure : Main
' Author    : beededea
' Date      : 12/03/2026
' Purpose   : Creates a transparent form with one image and a close button
'             one of the images is generated using Cairo
'---------------------------------------------------------------------------------------
'
'Public Sub Main()
'
'    Dim rc As Long
'    Dim useVBForm As Boolean
'
'    On Error GoTo Main_Error
'
'    useVBForm = True
'
'
'
'    If useVBForm = True Then
'        Call vbFormSetup
'    Else
'        'rc = initiateAPIWindow("API Window in VB6", "VbWndClass")
'        'MsgBox "Your window exited with code: " & rc
'    End If
'
'    On Error GoTo 0
'    Exit Sub
'
'Main_Error:
'
'     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Main of Module modMain"
'End Sub





'---------------------------------------------------------------------------------------
' Procedure : setMainFormDimensions
' Author    : beededea
' Date      : 31/10/2025
' Purpose   : set the main form upon which the dock resides to the size of the whole monitor, has to be done in twips
'---------------------------------------------------------------------------------------
'
Public Sub setMainFormDimensions()
    '
    On Error GoTo setMainFormDimensions_Error

    Form1.Height = screenHeightTwips
    Form1.Width = screenWidthTwips

    On Error GoTo 0
    Exit Sub

setMainFormDimensions_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure setMainFormDimensions of Form dock"

End Sub

'---------------------------------------------------------------------------------------
' Procedure : monitorProperties
' Author    : beededea
' Date      : 06/11/2025
' Purpose   : All this subroutine does at the moment is to set the screenTwipsPerPixel
'---------------------------------------------------------------------------------------
'
Public Sub monitorProperties()

    ' only calling TwipsPerPixelX/Y once on startup
    On Error GoTo monitorProperties_Error

    screenTwipsPerPixelX = fTwipsPerPixelX
    screenTwipsPerPixelY = fTwipsPerPixelY

    On Error GoTo 0
    Exit Sub

monitorProperties_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure monitorProperties of Form frmMain"
    
End Sub

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
    
'    Me.Height = Screen.Height '16200 correct
'    Me.Width = Screen.Width ' 16200 < VB6 bug here
    
    ' pixels for Cairo and GDI
    screenHeightPixels = GetDeviceCaps(thisHDC, VERTRES)
    screenWidthPixels = GetDeviceCaps(thisHDC, HORZRES)
    
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
' Purpose   : update some characteristics for the underlying form we will be updating using UpdateLayeredWindow API
'---------------------------------------------------------------------------------------
'
Public Sub setWindowCharacteristics(ByVal rDzOrderMode As String, ByVal thisOpacity As String)

    On Error GoTo setWindowCharacteristics_Error
    
    'set the transparency of the underlying form with click through
    windowLngReturn = GetWindowLong(Form1.hwnd, GWL_EXSTYLE)
    SetWindowLong Form1.hwnd, GWL_EXSTYLE, windowLngReturn Or WS_EX_LAYERED
    
    If rDzOrderMode = "0" Then
        SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE
    ElseIf rDzOrderMode = "1" Then
        SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE
    ElseIf rDzOrderMode = "2" Then
        SetWindowPos Form1.hwnd, HWND_BOTTOM, 0, 0, 0, 0, SWP_NOSIZE
    End If
    
    ' point structure that specifies the location of the layer updated in UpdateLayeredWindow
    apiPoint.x = 0
    apiPoint.y = 0
    
    ' point structure that specifies the size of the window in pixels
    windowSize.x = screenWidthPixels ' .59 DAEB 26/04/2021 frmMain.frm changed to use pixels alone, removed all unnecessary twip conversion
    windowSize.y = screenHeightPixels  ' .59 DAEB 26/04/2021 frmMain.frm changed to use pixels alone, removed all unnecessary twip conversion
    
    ' blending characteristics for opacity
    funcBlend32bpp.AlphaFormat = AC_SRC_ALPHA
    funcBlend32bpp.BlendFlags = 0
    funcBlend32bpp.BlendOp = AC_SRC_OVER
  
    ' set the opacity of the whole dock, used to display solidly and for instant autohide
    funcBlend32bpp.SourceConstantAlpha = 255 * Val(thisOpacity) / 100 ' this calc can be done elsewhere and we just use a passed var
    
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
Public Sub createGDIStructures()

    ' sets the bmpInfo object containing data to create a bitmap the whole screen size
    ' used later when creating DIB section of the correct size, width &c
    
    On Error GoTo createGDIStructures_Error

    ' Set the bitmap characteristics for use in SetDIBits later
    bmpInfo.bmpHeader.biSize = Len(bmpInfo.bmpHeader)
    bmpInfo.bmpHeader.biBitCount = 32
    bmpInfo.bmpHeader.biHeight = Form1.ScaleHeight
    bmpInfo.bmpHeader.biWidth = screenWidthPixels
    bmpInfo.bmpHeader.biPlanes = 1
    bmpInfo.bmpHeader.biSizeImage = bmpInfo.bmpHeader.biWidth * bmpInfo.bmpHeader.biHeight * (bmpInfo.bmpHeader.biBitCount / 8)
    
    ' A device context is a generalized rendering abstraction. It serves as a proxy between your rendering code and the output device.
    ' It allows you to use the same rendering code regardless of the destination; the low-level details are handled for you,
    ' dependant on the output device, including clipping, scaling, and viewport translation.
    
    ' A handle to the Device Context (HDC) is obtained before output is written and then released after elements have been written.
    dcMemory = CreateCompatibleDC(thisHDC)

    On Error GoTo 0
    Exit Sub

createGDIStructures_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure createGDIStructures of Form frmMain"

End Sub

'---------------------------------------------------------------------------------------
' Procedure : createNewGDIPBitmap
' Author    : beededea
' Date      : 05/11/2025
' Purpose   : Create a gdi bitmap with width and height of what we are going to draw into it. This is the entire drawing area for everything,
'             creating a bitmap in memory that our VB6/GDI application writes to directly. Called each animation interval.
'---------------------------------------------------------------------------------------
'
Public Sub createNewGDIPBitmap()

    On Error GoTo createNewGDIPBitmap_Error

    ' the existing bitmap deleted
    Call DeleteObject(hBmpMemory)
    
    ' create a device independent bitmap and return a handle, hBmpMemory, providing it a handle to device context allocated memory
    ' providing size information in bmpInfo and setting any attributes to the new bitmap
    hBmpMemory = CreateDIBSection(dcMemory, bmpInfo, DIB_RGB_COLORS, ByVal 0, 0, 0)
    
    ' Make the device context dcMemory use the bitmap.  hOldBmp is a return value giving a handle which determines success and allows reverting later to release GDI handles
    hOldBmp = SelectObject(dcMemory, hBmpMemory) ' releases memory used by any open GDI handle  in SD used within createNewGDIPBitmap
    
    ' Creates a GDIP graphic object and provides a pointer 'gdipFullScreenBitmap' using a handle to the bitmap graphic section assigned to the device context
    Call GdipCreateFromHDC(dcMemory, gdipFullScreenBitmap) ' dcMemory used to draw upon and place on screen using UpdateLayeredWindow later
    
    On Error GoTo 0
    Exit Sub

createNewGDIPBitmap_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure createNewGDIPBitmap of Form frmMain"
    
End Sub





'---------------------------------------------------------------------------------------
' Procedure : updateDisplayFromDictionary
' Author    : beededea
' Date      : 07/04/2020
' Purpose   : This utility displays using GDI+, one of several image bitmaps extracted from a dictionary collection by key.
'---------------------------------------------------------------------------------------
'
Public Function updateDisplayFromDictionary(ByVal Key As String, Optional Left As Long = 0, Optional Top As Long = 0, Optional Width As Long = -1, Optional Height As Long = -1) As Boolean

    On Error GoTo updateDisplayFromDictionary_Error
    
    'draws the selected image bitmap onto the GDIP full screen
    Call GdipDrawImageRectI(gdipFullScreenBitmap, imageBitmap, Left, Top, Width, Height)
    
    ' The GDIP graphics are now deleted
    Call GdipDeleteGraphics(imageBitmap)
    
   Exit Function

updateDisplayFromDictionary_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure updateDisplayFromDictionary of Form dock"

    On Error GoTo 0
    Exit Function
    
End Function

        
'---------------------------------------------------------------------------------------
' Procedure : readImageFromDictionary
' Author    : beededea
' Date      : 07/04/2020
' Purpose   : This utility extracts a specified image, by key from a named collection using GDI+.
'---------------------------------------------------------------------------------------
'
Public Function readImageFromDictionary(ByVal Key As String) As Long

    On Error GoTo readImageFromDictionary_Error
    
    ' get the stored image from the collection if it exists

'    #If TWINBASIC Then
'        If thisImageList.ImageExists(Key) <> 0 Then
'            readImageFromDictionary = thisImageList.bitmap(Key) ' return value
'        End If
'    #Else
        If thisImageList.Exists(Key) <> 0 Then
            readImageFromDictionary = thisImageList.bitmap(Key) ' return value
        End If
'    #End If
    
   Exit Function

readImageFromDictionary_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure readImageFromDictionary of Form dock"

    On Error GoTo 0
    Exit Function
    
End Function




'---------------------------------------------------------------------------------------
' Procedure : updateScreenUsingGDIPBitmap
' Author    : beededea
' Date      : 05/11/2025
' Purpose   : Calls UpdateLayeredWindow with previously created GDI bitmap
'---------------------------------------------------------------------------------------
'
Public Sub updateScreenUsingGDIPBitmap()
    
    On Error GoTo updateScreenUsingGDIPBitmap_Error

    'Call GdipDeleteGraphics(gdipFullScreenBitmap)  'The GDIP graphics are deleted first
    
    ' Using UpdateLayeredWindow it is handled by the Windows compositor, which can take advantage of hardware acceleration for blending and movement.

    ' the third parameter to UpdateLayeredWindow is a pointer to a structure that specifies the new screen position of the layered window.
    ' If the current position is not changing, pptDst can be NULL. It is null.
            
    'Update the specified window handle (hwnd) with a handle to our bitmap (dc) passing all the required characteristics
    Call UpdateLayeredWindow(hVBFormHwnd, thisHDC, ByVal 0&, windowSize, dcMemory, apiPoint, 0, VarPtr(funcBlend32bpp), ULW_ALPHA)   '*  in SD called whenever a draw is required

    ' releases memory for GDI handles
    'Call SelectObject(dcMemory, hOldBmp)
'    ' the existing bitmap deleted
    'Call DeleteObject(hBmpMemory) '
    'DeleteDC dcMemory
    'ReleaseDC 0, thisHDC

    On Error GoTo 0
    Exit Sub

updateScreenUsingGDIPBitmap_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure updateScreenUsingGDIPBitmap of Form frmMain"

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
'    Call ReleaseDC(hVBFormHwnd, dcMemory)
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
'    'Call createNewGDIPBitmap
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


'---------------------------------------------------------------------------------------
' Procedure : drawAlphaPngCairo
' Author    : beededea
' Date      : 31/10/2025
' Purpose   : draw a PNG image with a transparent background using Cairo only, no GDI+
'---------------------------------------------------------------------------------------
'
Public Sub drawAlphaPngCairo(thisHDC As Long, ByVal hwnd As Long, ByVal sPngPath As String, ByVal x As Long, ByVal y As Long)
    Dim Width As Long
    Dim Height As Long
    
    Dim surfImg As Long, cr As Long, surfPng As Long
    Dim dataPtr As Long
    Dim bmi As BITMAPINFO
    Dim blend As BLENDFUNCTION
    Dim hDC As Long, memDC As Long, hBmp As Long, hOldBmp As Long

    On Error GoTo drawAlphaPngCairo_Error
    
    gSngOpacity = 1

    ' create a Cairo surface from an on-disc PNG
    surfPng = cairo_image_surface_create_from_png(sPngPath)
    If surfPng = 0 Then Exit Sub

    Width = cairo_image_surface_get_width(surfPng)
    Height = cairo_image_surface_get_height(surfPng)

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
        
'    cairo_select_font_face cr, "segoe", CAIRO_FONT_SLANT_NORMAL, CAIRO_FONT_WEIGHT_BOLD
'    cairo_set_font_size cr, 32#
'    cairo_set_source_rgba cr, 0#, 0#, 1#, 0.3
'    cairo_move_to cr, 10#, 50#
'    cairo_show_text cr, "TARDIS"

'    cairo_translate cr, X, Y
'    cairo_rotate cr, M_PI / 4
'    cairo_scale cr, 1 / Sqr(2), 1 / Sqr(2)
'    cairo_translate cr, -X, -Y
'
'    ' set the cairo context using the surface on the form at a defined position, in this case top/left
    cairo_set_source_surface cr, surfPng, 0, 0

    ' Paint the PNG (preserves alpha)
    cairo_paint_with_alpha cr, gSngOpacity '   CC.Paint with alpha
    'cairo_paint cr

    ' Get pointer to pixel buffer
'    dataPtrPixelBuffer = cairo_image_surface_get_data(surfImg)
'
'    ' Copy Cairo ARGB pixel buffer into HBITMAP compatible DDB hBmpMemory (usually has better GDI performance than a DIB as used in Steamydock)
'    Call SetDIBits(dcMemory, hBmpMemory, 0, Height, ByVal dataPtrPixelBuffer, bmpInfo, 0) '*  in SD, an equivalent of GdipCreateFromHDC used within createNewGDIPBitmap?
'
'    ' tasks to tidy up, Cairo image, context and surface
'    cairo_destroy cr
'    cairo_surface_destroy surfImg
'    cairo_surface_destroy surfPng

    ' Prepare GDI structures
    
    ' A handle to the Device Context (HDC) is obtained before output is written and then released after elements have been written.
    ' Get a device context compatible with the screen

    'hDC = Me.hDC
    'memDC = CreateCompatibleDC(thisHDC)
    
    ' create a compatible bitmap and return a handle, bmpMemory, providing it a handle to device context allocated memory previously created with CreateCompatibleDC,
    ' providing size information in bmpInfo and setting any attributes to the new bitmap
    'hBmp = CreateCompatibleBitmap(thisHDC, Width, Height)
    
    ' Make the device context use the bitmap.
    'hOldBmp = SelectObject(memDC, hBmp)

'    With bmi.bmiHeader
'        .biSize = Len(bmi.bmiHeader)
'        .biWidth = Width
'        .biHeight = -Height
'        .biPlanes = 1
'        .biBitCount = 32
'        .biCompression = BI_RGB
'    End With

    ' Copy Cairo ARGB buffer into HBITMAP
    'Call SetDIBits(memDC, hBmp, 0, Height, ByVal dataPtr, bmi, 0)

    
    ' use the source image's alpha channel for blending characteristics for opacity
'    blend.BlendOp = AC_SRC_OVER
'    blend.SourceConstantAlpha = 255
'    blend.AlphaFormat = AC_SRC_ALPHA
    
    
    ' AlphaBlend() performs proper per-pixel alpha blending with the window’s background.
    
    ' Alpha blend onto window
    'Call AlphaBlend(hDC, X, Y, Width, Height, memDC, 0, 0, Width, Height, VarPtr(blend))
        
    'blit the buffer to the window’s HDC with per-pixel alpha blending.
'    Call AlphaBlend(thisHDC, 100, 100, 1000, 1000, thisHDC, 0, 0, 1000, 1000, VarPtr(funcBlend32bpp))

    ' or updatelayeredwindow

    ' delete temporary objects
'    Call SelectObject(memDC, hOldBmp)
'    Call DeleteObject(hBmp)
'    Call DeleteDC(memDC)
'    Call ReleaseDC(hWnd, hDC)

    ' tasks to tidy up, Cairo image, context and surface
    cairo_destroy cr
    cairo_surface_destroy surfImg
    cairo_surface_destroy surfPng


    On Error GoTo 0
    Exit Sub

drawAlphaPngCairo_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure drawAlphaPngCairo of Form frmMain"
End Sub






