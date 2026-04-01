Attribute VB_Name = "modMain"
'---------------------------------------------------------------------------------------
' Module    : modMain
' Author    : beededea
' Date      : 12/03/2026
' Purpose   : Test program to write images to a transparent form using Cairo and GDI+
'---------------------------------------------------------------------------------------

' form1.form_load         The main program entry point

' Main program classes & modules:
'   Form1                 The main VB6 form that provides the main program entry point. Invisible and used solely as a code holder for some routines.
'   cfImageGDIP           The main class that provides the image properties, hit testing and raising events, uses GDI+ to paint.
'   cImageEventHost       Eventhost bridge class to capture/enable withEvents for all the images of type cfImageGDIP within the collection.
'   MenuForm              Not visible at runtime, provides a right click menu to the main form.
'   modAPIDeclarations    Module containing all the API declarations specific to this program not already declared elsewhere.
'   modMain               This module, containing useful public routines used by this program
'   modSubClass           This module sub classes Form1 to intercept mouse-move and mouse-down messages, performs hit-tests and raises events.
'   modWindowAPI          Module that may provide a user-created window rather than a VB6 form (currently unused).

' Supporting classes & modules:
'   cGdipImageList        Wrapper class giving a 'Richclient' imageList interface to a collection whilst using GDI+ to load, resize and read images from it.
'   cTBImageList          As above but using a native TwinBasic collection whilst still using GDI+ to resize images where needed.
'   Dictionary            Christian Buse's Scripting.Dictionary (collection) replacement.
'   mGDIPImageList        Module to support cGdipImageList & cTBImageList, subs, functions and API declarations.

' The critical parts of this program are:

' addSingleImagesToImageList - that loads the PNGs into a dictionary collection (imageList).
' addSingleImagesToFullScreenDisplay - that puts single images on the pre-prepared screen using the addThisImage routine.
' InitialiseImageWidgetsFromXML - that puts mutiple images on screen from an XML definition, using the addThisImage routine.
' SubclassProc - the routine from which all hit tests and event trapping is initiated.

' Description:

' Program uses GDI+ to read transparent PNGs from a folder into a dictionary. The details of each image are stored in an XML file,
' which is read line by line and used to identify and place each transparent image layer on screen in the correct location and order.
' The dictionary is Christian Buse's - VBA Dictionary. The data for each PNG is loaded as a ADODB.Stream object and GDI+ is used to
' resize and place an image into a dictionary and from that, onto the screen. The main VB6 form is invisible and appears unused, but
' it exists as a receptacle to allow GDI+ to paint images directly to the device context associated with the form, ie. the screen.
' The form is sub-classed, ie. messages to and from the window are intercepted to allow manual handling of mouse events, hit-testing
' and raising of events, replicating what would occur if we were using VB6 controls. Hit-testing is performed using a duplicate collection
' loaded with a duplicate collection and the bounds and transparencies on the image are tested to determine which image layer has been
' 'clicked'. Yet another duplicate collection is loaded and tested to act as an event 'sink' bridge to provide event handling for each
' layer.
'
' For a multi-platform graphics alternative, Cairo will eventually be be an option that can be selected to read and place the images
' on screen but GDI+ is being used in the interim as it is easier to implement, allowing program construction and to prove the
' utility of the program until the Cairo code is complete and working. When Cairo is implemented GDI+ will probably still be used to
' resize and load the images into the various collections.

' Credits:

' Olaf Schmidt -
' Christian Buse - VBA Dictionary
' Andrew Heinlein  - creating a custom window
' Joaquim - Color Matrix

' Cairo Problems:
' ===============
'
' When you use Cairo with an HDC target (via cairo_win32_surface_create(hDC)), the resulting surface is not alpha-aware.
' GDI does not support per-pixel alpha — only color values. So when you draw an image with semi-transparent pixels, Cairo
' composites them as if against a black background before blitting them onto your window — causing shadows to appear dark and non-transparent.

' Potential solutions:
' =======================

' A.
' Cairo does all compositing in a true ARGB32 surface.
' You transfer that buffer to a GDI DIB section.
' AlphaBlend() performs proper per-pixel alpha blending with the window’s background.

' B.
' Do the whole thing using GDI+ as a demonstration of overall capability - (tested and working) -
' - when working later, I will slot in the Cairo code

' Status for B:
' =============

' We have a version that now places a Cairo image to the DC of a solid white background of user-created form using createWindowEx - (tested and working)
' We have added the necessary code to place an image on a solid VB6 form using GDI+ (tested and working)
' We have added the necessary code to place a stable image on a fully transparent form using GDI+ (tested and working)
' We have added the capability to store and extract an std picture image from a dictionary (CB's dict) and with my RC collection wrapper (tested and working)
' We have added bitmap capability to my RC collection wrapper (tested and working)
' Added a menu to allow easy program closing (tested and working)
' Added an image class to take an image bitmap and other properties to draw a single image on the screen (tested and working)
' Added routine to read image properties from parsed XML (tested and working).
' Added a load of multiple image objects to the screen using XML (tested and working).
' Added event handling for a single image object.
' Added a collection and class for hit testing to identify each image object's bounds and transparent areas.
' Added a collection and an event class 'sink' to allow the handling and raising of events for each image object.
' Fully documented.
' Fixed bugs in initialise from XML re: missing fields causing errors & missing " px" in metrics.

' Tasks?
' ======

' test with alternative XML file

' test with TB - WIP
' TBimageList ImageExists
' TBimageList bitmap WIP - do we need an option to return a GDI+ bitmap, I think so
' TBimageList RemoveAll - should not matter when compiled by TB

' Create a widget class comprising all the image objects

' later
' rename the class files to match their exposed names


'    bitmap immutable while locked
'
'    If you redraw or replace a bitmap:
'
'        UnlockBitmap
'        Set new bitmap
'        LockBitmap


' tooltips, use Faf's tooltip method for raising tooltips on items not automatically raising them.

' Add a top layer image list for images that do not require responses to click events and allow full click-through.

' Why do we need to utilise a user-created window/form?
'    For testing if for some reason we can't get the Cairo version working.
'    A custom form is lightweight and program size can be reduced, the extra complexity and utility of a VB6 program.
'    Don't do it until the whole lot is tested and working and after we have moved the code not being form-based.


' Status for Potential solution A - Cairo:
' ========================================
   
' All the above work and structure in B. is applicable to Cairo.

' The Cairo function needs to be hacked to move the file load to the startup

' The recommended call to cairo_image_surface_create does not work using other cairo DLLs, placing nothing on any hDC
' When I use vbCairo cairo_win32_surface_create (thisHDC) then a transparent image is placed directly on the device context as required
' meaning that all the rest of the code to place an image on a layered window is not even utilised.

' tried cairo_image_surface_create with any other cairo.dll whilst using vbCairo for all the rest
' tried cairo_image_surface_create with RC cairo.dll whilst using vbCairo for all the rest

' What I believe is happening is that the write to the hDc is working but the image will not persist on the hDC(0) as the explorer process regularly refreshes it.
' When writing to the VB6 form, the image persists but the vbCyan used as a key for transparency still is present as an outline artefact.

' What I think I have to do :

' try replacing ALL the Cairo calls using RC6 versions to see if the logic as suggested by chatGPT is reasonable
' try an alternative AI model to test the logic

' In fact the cairo_win32_surface_create(thisHDC) function is a tool that I can use to test the writing of an image to a hDC regardless of the rest of the logic.
' It might be the method I use with own-created window.

' investigate operator_clear

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

' initiateAPIWindow     The process that does the majority of the Window initialisation determining screen details, initiating the window.
' createAPIWindow       The main process that does the majority of the Window initialisation and other GDI+ configuration
' mainWndProc           The routine where all the Cairo and GDI+ drawing is done, intercepting messages such as WM_PAINT

' Other things to note for the future:

' GdipBitmapLockBits can act as a future zero-copy bridge between GDI+ LockBits and a Cairo surface, used in high-speed gaming...

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
    If thisImageList.Exists(Key) <> 0 Then
        readImageFromDictionary = thisImageList.bitmap(Key) ' return value
    End If
    
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






