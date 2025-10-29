VERSION 5.00
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   ClientHeight    =   5805
   ClientLeft      =   120
   ClientTop       =   120
   ClientWidth     =   11625
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   5805
   ScaleWidth      =   11625
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picbox1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4995
      Left            =   270
      ScaleHeight     =   4995
      ScaleWidth      =   3705
      TabIndex        =   1
      Top             =   240
      Width           =   3705
   End
   Begin VB.CommandButton Command 
      Caption         =   "Close"
      Height          =   630
      Left            =   10290
      TabIndex        =   0
      ToolTipText     =   "Click me to close the window"
      Top             =   4920
      Width           =   1170
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      Height          =   4200
      Left            =   7170
      Picture         =   "frmMain.frx":0000
      ToolTipText     =   "You should be able to drag the whole form by dragging any of the images "
      Top             =   210
      Width           =   4200
   End
   Begin VB.Image Image2 
      Appearance      =   0  'Flat
      Height          =   4575
      Left            =   4410
      Picture         =   "frmMain.frx":12643
      ToolTipText     =   "You should be able to drag the whole form by dragging any of the images "
      Top             =   210
      Width           =   4575
   End
   Begin VB.Menu mnuPopupMenu 
      Caption         =   "Main Menu"
      Visible         =   0   'False
      Begin VB.Menu mnuClose 
         Caption         =   "Close Widget"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long

' vars to obtain correct screen width (to correct VB6 bug) STARTS
Private Const HORZRES = 8
Private Const VERTRES = 10

Private screenTwipsPerPixelX As Long
Private screenTwipsPerPixelY As Long

Private screenWidthTwips As Long
Private screenHeightTwips As Long

Private Declare Function GetWindowLong Lib "user32.dll" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" ( _
                ByVal hWnd As Long, _
                ByVal crKey As Long, _
                ByVal bAlpha As Byte, _
                ByVal dwFlags As Long) As Long

' global Windows+ constants START
Private Const ULW_ALPHA = &H2
Private Const WS_EX_LAYERED = &H80000
Private Const WS_EX_TRANSPARENT As Integer = &H20
Private Const GWL_EXSTYLE As Long = -20
Private Const LWA_COLORKEY = &H1

Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Const WM_NCLBUTTONDOWN = &HA1
Private Const HTCAPTION = 2

' TwinBasic only functionality to read rawbyte PNG data from a scripting dictionary into
Private Enum LongPtr
    [_]
End Enum

'#If twinbasic Then
    'Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
        ByVal Destination As LongPtr, _
        ByVal Source As LongPtr, _
        ByVal Length As Long)

   ' --- Cairo constants ---
    'Private Const CAIRO_STATUS_SUCCESS As Long = 0
    '
    ' --- Our context type for reading PNG data ---
    Private Type PngReadContext
        DataPtr As LongPtr
        DataSize As Long
        Position As Long
    End Type
'#End If

'---------------------------------------------------------------------------------------
' Procedure : Form_Load
' Author    : beededea
' Date      : 27/10/2025
' Purpose   : Creates a transparent form with two images and a close button
'             one of the images is generated using Cairo
'---------------------------------------------------------------------------------------
'
Private Sub Form_Load()

    'First we need a Surface, which is a "physical thing" (a real Render-Target)
    Dim psfcFrm As cairo_surface_t
    Dim psfcImg As cairo_surface_t
    Dim pCr     As cairo_t ' RichClient equivalent -  Dim CC As cCairoContext
    Dim opacity As Single
    Dim lW          As Long
    Dim lH          As Long
    
    Dim Surfaces  As Object
    Dim PNGData() As Byte
        
    On Error GoTo Form_Load_Error

    'set the transparency of the underlying form using a colour key to define the transparency
    Me.BackColor = vbCyan
    picbox1.BackColor = vbCyan
    
    opacity = 0.9
    
    'set the transparency of the underlying form with click through, makes the form completely transparent
    SetWindowLong Me.hWnd, GWL_EXSTYLE, GetWindowLong(Me.hWnd, GWL_EXSTYLE) Or WS_EX_LAYERED ' Or WS_EX_TRANSPARENT
        
    ' this brings the form back again and sets a colour key to make selected items appear transparent
    SetLayeredWindowAttributes Me.hWnd, vbCyan, 0&, LWA_COLORKEY
    
    ' the addition of "Or WS_EX_TRANSPARENT" to SetWindowLong will make the revealed form fully click-through but unresponsive
    
    ' stalled on reading the PNG data from a byte array from a dictionary using VB6 as the cairo_image_surface_create_from_png_stream function requires population via a callback
    ' it CAN be done using TB callbacks but is not essential to have this within the VB6 program, it is just a proof of concept. To continue, this will eventually be a single image widget demo.
    
    ' The code is still here as trial for doing the same in TwinBasic
    
#If twinbasic Then
    ' dictionary for the PNG images
    Set Surfaces = CreateObject("Scripting.Dictionary")
    Surfaces.CompareMode = 1 ' for case-insenitive Key-Comparisons

    ' Load PNG bytes into memory
    PNGData = LoadFileToBytes(App.path & "\tardis.png")
    
    'We need to learn how to create a scaled image using Cairo
    ' HERE

    ' Store PNG bytes in the dictionary
    Surfaces.Add "Tardis", PNGData
#End If
    
' load the best image file quality that can be loaded into an image control, for VB6 a JPG, for TB a PNG
#If twinbasic Then
    Image1.Picture = LoadPicture(App.path & "\player.png")
    Image2.Picture = LoadPicture(App.path & "\twinbasic.png")
#Else
    Image1.Picture = LoadPicture(App.path & "\player.jpg")
    Image2.Picture = LoadPicture(App.path & "\twinbasic.jpg")
#End If
    
    ' only calling TwipsPerPixelX/Y once on startup
    screenTwipsPerPixelX = fTwipsPerPixelX
    screenTwipsPerPixelY = fTwipsPerPixelY
    
    screenHeightTwips = GetDeviceCaps(Me.hDC, VERTRES) * screenTwipsPerPixelY
    screenWidthTwips = GetDeviceCaps(Me.hDC, HORZRES) * screenTwipsPerPixelX
    
    'set the main form upon which the dock resides to the size of the whole monitor, has to be done in twips
'    Me.Height = screenHeightTwips
'    Me.Width = screenWidthTwips
'
'    using Cairo with an existing handle to a window

'    HDC dc = GetDC(windowHandle);
'    cairo_win32_surface_create(dc);
'    ReleaseDC(windowHandle, dc);
'
'    If you do not want to create a window at this stage use the desktop window to acquire a device context.
'
'    Dim hdcScreen As Long
'
'    windowHandleHDC = GetDesktopWindow()
'    hdcScreen = GetDC(windowHandleHDC)
'    cairo_win32_surface_create dc
'    ReleaseDC windowHandleHDC, dc


'    Register a window class with RegisterClass
'    Create the window with CreateWindow or CreateWindowEx
'    Process messages with a message pump by calling GetMessage, TranslateMessage and DispatchMessage
'
'Additionally you will have to implement a function to handle processing of window messages such as WM_PAINT.
    
    ' that's the native VB6/TwinBasic stuff done, now we play with Cairo

    ' create a Cairo surface that writes directly to the picture box hardware device context for a PICBOX, note imageboxes do not have a .hDC
    ' a Cairo-Image-Surface is something like an allocated InMemory-Bitmap (a hDIB)
    psfcFrm = cairo_win32_surface_create(picbox1.hDC) '  RichClient equivalent Set Srf = Cairo.CreateSurface(200, 100, ImageSurface)
    
    ' create a Cairo context for issuing drawing commands on the surface, in this case we aren't doing any drawing just painting using a PNG
    ' a context is something akin to a hDC in GDI unlike in GDI, where we would "Select" a Bitmap into a hDC first...
    ' with Cairo we can create such a Cairo context "anytime" from any Surface
    pCr = cairo_create(psfcFrm) ' RichClient equivalent Set CC = Srf.CreateContext

    ' create a Cairo image object from file
    psfcImg = cairo_image_surface_create_from_png(App.path & "\tardis.png")
    
    ' to read a byte array from a scripting collection and feed it to a Cairo surface is more involved than it looks
    ' as cairo_image_surface_create_from_png_stream requires a callback from the function that is a helper function that handles the data for Cairo
    ' RichClient does this using VB6 which I cannot replicate so using TB to do the same
    
#If twinbasic Then
    Dim surf As LongPtr
    surf = CairoSurfaceFromPngBytes(dict("Tardis"))
#End If
    
    lW = cairo_image_surface_get_width(psfcImg)
    lH = cairo_image_surface_get_height(psfcImg)
    
    cairo_select_font_face pCr, "segoe", CAIRO_FONT_SLANT_NORMAL, CAIRO_FONT_WEIGHT_BOLD
    cairo_set_font_size pCr, 32#
    cairo_set_source_rgba pCr, 0#, 0#, 1#, 0.3
    cairo_move_to pCr, 10#, 50#
    cairo_show_text pCr, "TARDIS"

    cairo_translate pCr, 128#, 128# ' requires doubles
    cairo_rotate pCr, M_PI / 4
    cairo_scale pCr, 1 / Sqr(2), 1 / Sqr(2)
    cairo_translate pCr, -128#, -128#  ' requires doubles
    
      ' set the cairo context using the surface on the form at a defined position, in this case top/left
    cairo_set_source_surface pCr, psfcImg, 1, 1
        
    'now paint to the cairo context
    cairo_paint_with_alpha pCr, opacity '   CC.Paint with alpha
    
    ' tasks to tidy up, Cairo image, context and surface
    
    cairo_surface_destroy psfcImg
    cairo_destroy pCr
    cairo_surface_destroy psfcFrm

    On Error GoTo 0
    Exit Sub

Form_Load_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Form_Load of Form frmMain"
    
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
    
    Const LOGPIXELSX = 88        '  Logical pixels/inch in X
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


'---------------------------------------------------------------------------------------
' Procedure : Image1_DblClick
' Author    : beededea
' Date      : 28/10/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub Image1_DblClick()
    On Error GoTo Image1_DblClick_Error

    Call DblClickHandler(Image1.Name)

    On Error GoTo 0
    Exit Sub

Image1_DblClick_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Image1_DblClick of Form frmMain"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : image1_MouseDown
' Author    : beededea
' Date      : 27/10/2025
' Purpose   : capture a mousedown on the only VB6 control to allow it to be dragged without a titlebar
'---------------------------------------------------------------------------------------
'
Private Sub image1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    On Error GoTo image1_MouseDown_Error

    Call MouseDownHandler(Image1.Name, Button, Shift, x, y)

    On Error GoTo 0
    Exit Sub

image1_MouseDown_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure image1_MouseDown of Form frmMain"

End Sub



'---------------------------------------------------------------------------------------
' Procedure : Image2_DblClick
' Author    : beededea
' Date      : 28/10/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub Image2_DblClick()
    On Error GoTo Image2_DblClick_Error

    Call DblClickHandler(Image2.Name)

    On Error GoTo 0
    Exit Sub

Image2_DblClick_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Image2_DblClick of Form frmMain"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : image2_MouseDown
' Author    : beededea
' Date      : 27/10/2025
' Purpose   : capture a mousedown on the only VB6 control to allow it to be dragged without a titlebar
'---------------------------------------------------------------------------------------
'
Private Sub Image2_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    On Error GoTo image2_MouseDown_Error

    Call MouseDownHandler(Image2.Name, Button, Shift, x, y)

    On Error GoTo 0
    Exit Sub

image2_MouseDown_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure image2_MouseDown of Form frmMain"

End Sub


'---------------------------------------------------------------------------------------
' Procedure : Form_MouseDown
' Author    : beededea
' Date      : 14/08/2023
' Purpose   : standard form down event to generate the menu across the board
'---------------------------------------------------------------------------------------
'
Private Sub Form_MouseDown(ByRef Button As Integer, ByRef Shift As Integer, ByRef x As Single, ByRef y As Single)
   On Error GoTo Form_MouseDown_Error

    If Button = 2 Then
        Me.PopupMenu mnuPopupMenu, vbPopupMenuRightButton
    End If

   On Error GoTo 0
   Exit Sub

Form_MouseDown_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Form_MouseDown of Form widgetPrefs"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : mnuClose_Click
' Author    : beededea
' Date      : 27/10/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub mnuClose_Click()
    On Error GoTo mnuClose_Click_Error

    End

    On Error GoTo 0
    Exit Sub

mnuClose_Click_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuClose_Click of Form frmMain"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : Command_Click
' Author    : beededea
' Date      : 27/10/2025
' Purpose   : closes the form
'---------------------------------------------------------------------------------------
'
Private Sub Command_Click()
    On Error GoTo Command_Click_Error

    End

    On Error GoTo 0
    Exit Sub

Command_Click_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Command_Click of Form frmMain"
End Sub


'---------------------------------------------------------------------------------------
' Procedure : picbox1_DblClick
' Author    : beededea
' Date      : 28/10/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub picbox1_DblClick()
    On Error GoTo picbox1_DblClick_Error

    Call DblClickHandler(picbox1.Name)

    On Error GoTo 0
    Exit Sub

picbox1_DblClick_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure picbox1_DblClick of Form frmMain"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : picbox1_MouseDown
' Author    : beededea
' Date      : 27/10/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub picbox1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    On Error GoTo picbox1_MouseDown_Error

    Call MouseDownHandler(picbox1.Name, Button, Shift, x, y)
    

    On Error GoTo 0
    Exit Sub

picbox1_MouseDown_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure picbox1_MouseDown of Form frmMain"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : MouseDownHandler
' Author    : beededea
' Date      : 27/10/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub MouseDownHandler(ByVal ctrlName As String, Optional Button As Integer, Optional Shift As Integer, Optional x As Single, Optional y As Single)
    On Error GoTo MouseDownHandler_Error

    If Button = 2 Then
        Me.PopupMenu mnuPopupMenu, vbPopupMenuRightButton
    Else
        'MsgBox ctrlName
        ReleaseCapture
        SendMessage Me.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
    End If

    On Error GoTo 0
    Exit Sub

MouseDownHandler_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure MouseDownHandler of Form frmMain"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : DblClickHandler
' Author    : beededea
' Date      : 27/10/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub DblClickHandler(ByVal ctrlName As String)
    On Error GoTo DblClickHandler_Error

    MsgBox ctrlName & " double-clicked!"

    On Error GoTo 0
    Exit Sub

DblClickHandler_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure DblClickHandler of Form frmMain"
End Sub





#If twinbasic Then

    ' --- The callback prototype ---
    Private Function CairoPngReadFunc( _
        ByVal closure As LongPtr, _
        ByVal data As LongPtr, _
        ByVal Length As Long) As Long
    
        Dim ctx As PngReadContext
        CopyMemory ByVal VarPtr(ctx), ByVal closure, LenB(ctx)
    
        Dim remaining As Long
        remaining = ctx.DataSize - ctx.Position
        If remaining < Length Then Length = remaining
    
        ' Copy the next chunk of bytes into Cairo's buffer
        CopyMemory ByVal data, ByVal (ctx.DataPtr + ctx.Position), Length
        ctx.Position = ctx.Position + Length
    
        ' Write updated ctx back
        CopyMemory ByVal closure, ByVal VarPtr(ctx), LenB(ctx)
    
        CairoPngReadFunc = CAIRO_STATUS_SUCCESS
    End Function
    
    
    Public Sub Example_CreateCairoSurfaceFromDictionary()
        ' --- Create and populate dictionary ---
        Dim dict As Object
        Set dict = CreateObject("Scripting.Dictionary")
    
        Dim pngBytes() As Byte
        pngBytes = LoadFileToBytes("C:\example.png")
    
        dict.Add "Logo", pngBytes
    
        ' --- Build read context ---
        Dim ctx As PngReadContext
        ctx.DataPtr = VarPtr(dict("Logo")(0))
        ctx.DataSize = UBound(dict("Logo")) + 1
        ctx.Position = 0
    
        ' --- Create surface from memory ---
        Dim surf As LongPtr
        surf = cairo_image_surface_create_from_png_stream(AddressOf CairoPngReadFunc, VarPtr(ctx))
    
        If cairo_surface_status(surf) = CAIRO_STATUS_SUCCESS Then
            MsgBox "Cairo surface created successfully!"
        Else
            MsgBox "Failed to create Cairo surface"
        End If
    End Sub
    
    Private Function LoadFileToBytes(ByVal path As String) As Byte()
        Dim f As Integer
        f = FreeFile
        Open path For Binary As #f
        ReDim LoadFileToBytes(LOF(f) - 1)
        Get #f, , LoadFileToBytes
        Close #f
    End Function
    
    
    Public Function CairoSurfaceFromPngBytes(PNGData() As Byte) As LongPtr
        Dim ctx As PngReadContext
        ctx.DataPtr = VarPtr(PNGData(0))
        ctx.DataSize = UBound(PNGData) + 1
        ctx.Position = 0
        CairoSurfaceFromPngBytes = cairo_image_surface_create_from_png_stream(AddressOf CairoPngReadFunc, VarPtr(ctx))
    End Function

#End If
