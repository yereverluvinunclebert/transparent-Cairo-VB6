VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   ClientHeight    =   3135
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   209
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   312
   Begin VB.Timer tmrAnimate 
      Interval        =   500
      Left            =   1410
      Top             =   2130
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    'thisHDC = GetDC(0&)  ' <- writing to the desktop dc fully transparent - overwritten shortly after
    hVBFormHwnd = Me.hWnd
    
    
    thisHDC = Me.hDC
        
    ' check the selected monitor properties and determine the number of twips per pixel for this screen
    Call monitorProperties
    
    ' resolve VB6 sizing width bug
    Call resolveVB6SizeBug ' requires MonitorProperties to be in place above to assign a value to screenTwipsPerPixelY
    
    'set the main form upon which the dock resides to the size of the whole monitor, has to be done in twips
    Call setMainFormDimensions
    
    ' Initialises GDI Plus
    Call initialiseGDIPStartup
    
    ' add image to image list
    Call addImageToImageList
    
    imageBitmap = readImageFromDictionary("tardis")
    
    ' sets bmpInfo object to create a bitmap of the whole screen size and get a handle to the Device Context
    Call createGDIStructures
           
    'creates a bitmap section in memory that applications can write to directly
    Call createNewGDIPBitmap
    
    ' update the window with the appropriately sized and qualified image
    Call setWindowCharacteristics ' This is the function that actually changes the display, called by animate timers, must also be here
    


End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = 2 Then 'right click to display a menu
        PopupMenu menuForm.mnuMainMenu, vbPopupMenuRightButton
    End If

End Sub

Private Sub tmrAnimate_Timer()


    ' now we paint the image using GDI+ extracting the image from a previously loaded dictionary, in this case Christian Buse's VBA dictionary replacement
    updateDisplayFromDictionary "tardis", (500), (250), (200), (200)

    ' Calls UpdateLayeredWindow with created GDI bitmap
    Call updateScreenUsingGDIPBitmap

    'Call drawAlphaPngGDIP(500, 250, 200, 200)
    
    ' now we paint the image using Cairo, Cairo HAS to load from file as the process to get Cairo to load from a collection is rather tricky using VB6 (Cairo requires a callback as input)
    'Call drawAlphaPngCairo(thisHDC, Form1.hwnd, App.Path & "\tardis.png", 300, 350)

End Sub
