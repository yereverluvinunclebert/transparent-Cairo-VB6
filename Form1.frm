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
      Enabled         =   0   'False
      Interval        =   500
      Left            =   150
      Top             =   90
   End
   Begin VB.Label lblForm 
      Caption         =   "This form is made invisible at runtime and unused except for timers"
      Height          =   735
      Left            =   390
      TabIndex        =   0
      Top             =   2220
      Width           =   2145
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents thisGDIPimage As cfImageGDIP
Attribute thisGDIPimage.VB_VarHelpID = -1

Private mHosts As Collection

'---------------------------------------------------------------------------------------
' Procedure : Form_Load
' Author    : beededea
' Date      : 18/03/2026
' Purpose   : calls vbFormSetup to allow the program to be VB6 or user-created custom form based
'---------------------------------------------------------------------------------------
'
Private Sub Form_Load()

    On Error GoTo Form_Load_Error
    
    Call vbFormSetup
    
    On Error GoTo 0
    Exit Sub

Form_Load_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Form_Load of Form Form1"

End Sub


'---------------------------------------------------------------------------------------
' Procedure : vbFormSetup
' Author    : beededea
' Date      : 16/03/2026
' Purpose   : Used by VB6 form1.Form_Load or user-created custom form via Main
'---------------------------------------------------------------------------------------
Public Sub vbFormSetup()

    Dim sWidgetOpacity As String: sWidgetOpacity = vbNullString
    Dim sWidgetZOrder As String: sWidgetZOrder = vbNullString

    On Error GoTo vbFormSetup_Error

    hVBFormHwnd = Form1.hwnd
    thisHDC = Form1.hDC
    sWidgetOpacity = "100"
    sWidgetZOrder = "2"
    
    Set thisGDIPimage = New cfImageGDIP
    
    'Set gImage = thisGDIPimage
    
    Set mHosts = New Collection
    Set gImages = New Collection
    
    gFormHwnd = Form1.hwnd
    gPrevWndProc = SetWindowLong(Form1.hwnd, GWL_WNDPROC, AddressOf SubclassProc)
    
    ' check the selected monitor properties and determine the number of twips per pixel for this screen
    Call monitorProperties
    
    ' resolve VB6 sizing width bug
    Call resolveVB6SizeBug ' requires MonitorProperties to be in place above to assign a value to screenTwipsPerPixelY
    
    'set the main form upon which the dock resides to the size of the whole monitor, has to be done in twips
    Call setMainFormDimensions
    
    ' Initialises GDI Plus
    Call initialiseGDIPStartup
    
    ' update the window with the appropriately sized and qualified image
    Call setWindowCharacteristics(sWidgetZOrder, sWidgetOpacity)
    
    ' sets bmpInfo object to create a bitmap of the whole screen size and get a handle to the Device Context
    Call createGDIStructures
    
    ' add single images to image list
    Call addSingleImagesToImageList
                  
    'creates a bitmap section in memory that applications can write to directly
    Call createNewGDIPBitmap ' clears the whole previously drawn image section and any animation can continue
    
    ' now we paint the images using GDI+ extracting the image from a pre-loaded dictionary, in this case Christian Buse's VBA dictionary replacement
    Call addSingleImagesToFullScreenDisplay
    
    'load the XML image data (previously extracted directly from the PSD)
    Call InitialiseImageWidgetsFromXML

    ' Calls UpdateLayeredWindow with created GDI bitmap
    Call updateScreenUsingGDIPBitmap
    
    ' now we paint the image using Cairo, (unfinished) Cairo HAS to load from file as the process to get Cairo to load from a collection is rather tricky using VB6 (Cairo requires a callback as input)
    'Call drawAlphaPngCairo(GetDC(0&), hVBFormHwnd, App.Path & "\player.png", 50, 350)

    On Error GoTo 0
    Exit Sub

vbFormSetup_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure vbFormSetup of Module modMain"

End Sub
'---------------------------------------------------------------------------------------
' Procedure : Form_MouseUp
' Author    : beededea
' Date      : 18/03/2026
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

    On Error GoTo Form_MouseUp_Error

    If Button = 2 Then 'right click to display a menu
        PopupMenu menuForm.mnuMainMenu, vbPopupMenuRightButton
    End If

    On Error GoTo 0
    Exit Sub

Form_MouseUp_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Form_MouseUp of Form Form1"

End Sub

''---------------------------------------------------------------------------------------
'' Procedure : tmrAnimate_Timer
'' Author    : beededea
'' Date      : 18/03/2026
'' Purpose   : Timer WILL be replaced by a custom timer, not currently used
''---------------------------------------------------------------------------------------
''
'Private Sub tmrAnimate_Timer()
'
'
'    ' now we paint the image using GDI+ extracting the image from a previously loaded dictionary, in this case Christian Buse's VBA dictionary replacement
'    'updateDisplayFromDictionary "tardis", (500), (250), (200), (200)
'
'    ' Calls UpdateLayeredWindow with created GDI bitmap
'    'Call updateScreenUsingGDIPBitmap
'
'    'Call drawAlphaPngGDIP(500, 250, 200, 200)
'
'    ' now we paint the image using Cairo, Cairo HAS to load from file as the process to get Cairo to load from a collection is rather tricky using VB6 (Cairo requires a callback as input)
'    'Call drawAlphaPngCairo(GetDC(0&), hVBFormHwnd, App.Path & "\player.png", 300, 350)
'
'    On Error GoTo tmrAnimate_Timer_Error
'
'    On Error GoTo 0
'    Exit Sub
'
'tmrAnimate_Timer_Error:
'
'     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure tmrAnimate_Timer of Form Form1"
'End Sub


Private Sub Form_Unload(Cancel As Integer)
    If gPrevWndProc <> 0 Then
        SetWindowLong Me.hwnd, GWL_WNDPROC, gPrevWndProc
    End If
End Sub


'---------------------------------------------------------------------------------------
' Procedure : thisGDIPimage_MouseDown
' Author    : beededea
' Date      : 24/03/2026
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub thisGDIPimage_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    On Error GoTo thisGDIPimage_MouseDown_Error

    MsgBox "clicked on !"

    On Error GoTo 0
    Exit Sub

thisGDIPimage_MouseDown_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure thisGDIPimage_MouseDown of Form Form1"
End Sub


'---------------------------------------------------------------------------------------
' Procedure : thisGDIPimage_MouseMove
' Author    : beededea
' Date      : 24/03/2026
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub thisGDIPimage_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    On Error GoTo thisGDIPimage_MouseMove_Error

    

    On Error GoTo 0
    Exit Sub

thisGDIPimage_MouseMove_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure thisGDIPimage_MouseMove of Form Form1"
End Sub

    


'---------------------------------------------------------------------------------------
' Procedure : addSingleImagesToImageList
' Author    : beededea
' Date      : 13/03/2026
' Purpose   : addition of any single images required that are not within the PSD-derived XML
'---------------------------------------------------------------------------------------
'
Public Sub addSingleImagesToImageList()
    
    On Error GoTo addSingleImagesToImageList_Error

    thisImageList.AddImage "tardis", App.Path & "\tardis.png"
    thisImageList.AddImage "player", App.Path & "\player.png"
    
    'addImagesToStdCollection
    
    'Call addImagesToStdCollection(thisImageList.Bitmap("player"), 750, 250, 200, 200)
    
    On Error GoTo 0
    Exit Sub

addSingleImagesToImageList_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure addSingleImagesToImageList of Module modMain"

End Sub
    
'---------------------------------------------------------------------------------------
' Procedure : addSingleImagesToFullScreenDisplay
' Author    : beededea
' Date      : 18/03/2026
' Purpose   : paint the images using GDI+ extracting the image from a pre-loaded dictionary, in this case Christian Buse's VBA dictionary replacement
'---------------------------------------------------------------------------------------
'
Public Sub addSingleImagesToFullScreenDisplay()
    
    On Error GoTo addSingleImagesToFullScreenDisplay_Error

    With thisGDIPimage
        .Bitmap = readImageFromDictionary("tardis")
        .Left = 750
        .Top = 250
        .Width = 200
        .height = 200
        .Name = "tardis"
        .Opacity = 100
        .Tooltip = "this image is the Tardis image"
        .Refresh
    End With
    
    With thisGDIPimage
        .Bitmap = readImageFromDictionary("player")
        .Left = 950
        .Top = 250
        .Width = 200
        .height = 200
        .Name = "player"
        .Opacity = 100
        .Tooltip = "this image is the Player image"
        .Refresh
    End With

    On Error GoTo 0
    Exit Sub

addSingleImagesToFullScreenDisplay_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure addSingleImagesToFullScreenDisplay of Module modMain"
End Sub





' ----------------------------------------------------------------
' Procedure Name: InitialiseImageWidgetsFromXML
' Purpose   :  Creates a GDIP surface object from each and every PSD layer in the PSD file.
'              For all the interactive UI elements it creates surface objects with corresponding keynames,
'              locations and sizes as per the original PSD for each layer. The surfaces are populated from PNGs whose metric
'              data are held in an extracted XML file. It creates an instance of each image using the cfImageGDIP class
'
'              The images are stored as GDIP images within an imageList.
'
'              NOTE: The XML data is created using the Create Widget 2.0.js using Photoshop and the PSD found in the RES folder.
'              The XML file is renamed from .KON and placed in the RES folder.
'
' Procedure Kind: Sub
' Procedure Access: Public
' Author: beededea
' Date: 28/08/2025
' ----------------------------------------------------------------
Public Sub InitialiseImageWidgetsFromXML()

    'Dim answer As VbMsgBoxResult
    Dim answerMsg  As String: answerMsg = vbNullString

    Dim num_results As Integer: num_results = 0
    Dim windowWidth As String: windowWidth = vbNullString
    Dim windowHeight As String: windowHeight = vbNullString
    Dim sSrc  As String: sSrc = vbNullString
    Dim sName As String: sName = vbNullString
    Dim sWidth As String: sWidth = vbNullString
    Dim sHeight As String: sHeight = vbNullString
'    Dim sHOffset As String: sHOffset = vbNullString
'    Dim sVOffset As String: sVOffset = vbNullString
    Dim sOpacity As String: sOpacity = vbNullString
    Dim Width As Long: Width = 0
    Dim height As Long: height = 0
    Dim hOffset As Long: hOffset = 0
    Dim vOffset As Long: vOffset = 0
    Dim Opacity As Integer: Opacity = 0
    Dim someOpacity As Double: someOpacity = 0
    Dim xmlFileToLoad As String: xmlFileToLoad = vbNullString
    Dim pngFileToLoad As String: pngFileToLoad = vbNullString
    
    Dim nodeList As MSXML2.IXMLDOMNodeList
    Dim objxmldoc As MSXML2.DOMDocument60
   
    Set objxmldoc = New MSXML2.DOMDocument60
    
    Dim node As MSXML2.IXMLDOMNode
    Dim MainNode As MSXML2.IXMLDOMNode
'    Dim ImageNode As MSXML2.IXMLDOMNode
'    Dim ImageNodes As MSXML2.IXMLDOMNodeList
    
    On Error GoTo InitialiseImageWidgetsFromXML_Error
    
    someOpacity = Val(sOpacity) / 100
    xmlFileToLoad = App.Path & "\RES\imagesXML.xml"
    
    If fFExists(xmlFileToLoad) Then
        objxmldoc.Load xmlFileToLoad
    Else
        MsgBox "The XML file that contains the image data is missing " & xmlFileToLoad
    End If
    
    ' obtain the overall widget width and height
    Set MainNode = objxmldoc.selectSingleNode("widget/window")
            
    windowWidth = MainNode.selectSingleNode("@width").Text
    windowHeight = MainNode.selectSingleNode("@height").Text
'
'    pPSDWidth = CLng(windowWidth)
'    pPSDHeight = CLng(windowHeight)
    
    ' get the image values from the XML data, the num results should be non-zero
    Set nodeList = objxmldoc.selectNodes("widget/window/image")
    num_results = nodeList.Length
    
    ' no results found, go on as normal using the sampling interval
    If num_results = 0 Then
        answerMsg = "1. There is a problem with the XML data file that describes the image, seems to contain no valid data"
        'msgBoxA answerMsg, vbOKOnly + vbExclamation, "XML Warning", True, "InitialiseImageWidgetsFromXMLPollingWarning"
        MsgBox answerMsg
        Exit Sub ' Return
    End If

    If Not nodeList Is Nothing Then
         For Each node In nodeList
         
            sSrc = node.selectSingleNode("@src").Text
            sSrc = Replace(sSrc, "/", "\")
            sName = node.selectSingleNode("@name").Text
            sWidth = node.selectSingleNode("@width").Text
            Width = CLng(Left$(sWidth, (InStr(sWidth, " px") - 1)))
            
            sHeight = node.selectSingleNode("@height").Text
            height = CLng(Left$(sHeight, (InStr(sHeight, " px") - 1)))
            
            hOffset = CLng(node.selectSingleNode("@hOffset").Text)
            vOffset = CLng(node.selectSingleNode("@vOffset").Text)
            Opacity = CInt(node.selectSingleNode("@opacity").Text) / 2.55
            
            If Opacity = 100 Then  ' only handles layers that have an opacity greater than 0 - need to note this for the future, this will cause a problem!
            
               'add each current Layer path and surface object into the global ImageList collection (using LayerPath as the ImageKey)
                pngFileToLoad = App.Path & "\RES\" & sSrc
                If fFExists(pngFileToLoad) Then
                    
                    ' add the image named in the XML to the image list
                    thisImageList.AddImage sName, pngFileToLoad
    
                    ' create an image object and write it to the full screen bitmap
                    With thisGDIPimage
                        .Bitmap = readImageFromDictionary(sName)
                        .Left = hOffset
                        .Top = vOffset
                        .Width = Width
                        .height = height
                        .Name = sName
                        .Opacity = 100
                        .Tooltip = "this image is the " & sName & " image"
                        .Refresh
                    End With
                Else
                    MsgBox "Error, this PNG resource file seems to be missing " & pngFileToLoad
                End If

            End If
          Next node
    End If
   
    'Cleanup
    Set nodeList = Nothing
    Set MainNode = Nothing
    
    On Error GoTo 0
    Exit Sub

InitialiseImageWidgetsFromXML_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure InitialiseImageWidgetsFromXML, line " & Erl & "."

End Sub



'---------------------------------------------------------------------------------------
' Procedure : addImagesToStdCollection
' Author    : beededea
' Date      : 24/03/2026
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Sub addImagesToStdCollection(ByVal bmp As Long, ByVal x As Long, ByVal y As Long, ByVal w As Long, ByVal h As Long)

    Dim img As cfImageGDIP
    
    On Error GoTo addImagesToStdCollection_Error

    Set img = New cfImageGDIP
    
    img.Bitmap = bmp
    img.Left = x
    img.Top = y
    img.Width = w
    img.height = h
    
    ' lock if using alpha hit-test
    'img.LockBitmap
    
    ' store in global list (for hit-testing)
    gImages.Add img
    
    ' create event host
    Dim host As cImageHost
    Set host = New cImageHost
    
    Set host.img = img
    host.Index = mHosts.Count + 1
    
    mHosts.Add host

    On Error GoTo 0
    Exit Sub

addImagesToStdCollection_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure addImagesToStdCollection of Form Form1"

End Sub
