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
'---------------------------------------------------------------------------------------
' Module    : Form1
' Author    : beededea
' Date      : 29/03/2026
' Purpose   :
'---------------------------------------------------------------------------------------

Option Explicit

Private WithEvents thisWidget As cWidgetForm
Attribute thisWidget.VB_VarHelpID = -1
Private WithEvents thisGDIPimage As cImageGDIP
Attribute thisGDIPimage.VB_VarHelpID = -1

Private mEventHostCollection As Collection



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
    widgetFormName = "Transparent Cairo"
    
    Set thisWidget = New cWidgetForm
    Set thisGDIPimage = New cImageGDIP
        
    ' standard VB6 collections used for hit testing and event capture
    Set gHitTestCollection = New Collection
    Set mEventHostCollection = New Collection
    
    ' subclass the form to capture events
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
    
    ' update some characteristics for the underlying invisible form
    Call setWindowCharacteristics(sWidgetZOrder, sWidgetOpacity)
    
    ' sets bmpInfo object to create a bitmap of the whole screen size and get a handle to the Device Context
    Call createGDIStructures
    
    ' create a widget object
    Call addThisForm("Transparent Cairo", 0, 0, 900, 750, 100, vbNullString, True)
    
    ' add single images to image list (dictionary)
    Call addSingleImagesToImageList
                  
    'creates a bitmap section in memory that applications can write to directly
    Call createNewGDIPBitmap ' clears the whole previously drawn image section and any animation can continue
    
    ' now we paint the images using GDI+ extracting the image from a pre-loaded dictionary, in this case Christian Buse's VBA dictionary replacement
    Call addSingleImagesToFullScreenDisplay
    
    'load the XML image data (previously extracted directly from the PSD)
    Call InitialiseImageSurfacesFromXML

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



'---------------------------------------------------------------------------------------
' Procedure : Form_Unload
' Author    : beededea
' Date      : 29/03/2026
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo Form_Unload_Error
    
    Dim img As cImageGDIP

    If gPrevWndProc <> 0 Then
        SetWindowLong Me.hwnd, GWL_WNDPROC, gPrevWndProc
    End If

    For Each img In gHitTestCollection
        img.UnlockBitmap
    Next

    On Error GoTo 0
    Exit Sub

Form_Unload_Error:

    With Err
         If .Number <> 0 Then
            MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Form_Unload of Form Form1"
            Resume Next
          End If
    End With
End Sub

    


'---------------------------------------------------------------------------------------
' Procedure : addSingleImagesToImageList
' Author    : beededea
' Date      : 13/03/2026
' Purpose   : addition of any single images required that are not within the PSD-derived XML
'             inserts the image into a dictionary, in this case Christian Buse's VBA dictionary replacement
'---------------------------------------------------------------------------------------
'
Public Sub addSingleImagesToImageList()
    
    On Error GoTo addSingleImagesToImageList_Error

    thisImageList.AddImage "tardis", App.Path & "\tardis.png"
    thisImageList.AddImage "player", App.Path & "\player.png"
            
    On Error GoTo 0
    Exit Sub

addSingleImagesToImageList_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure addSingleImagesToImageList of Module modMain"

End Sub
    
'---------------------------------------------------------------------------------------
' Procedure : addSingleImagesToFullScreenDisplay
' Author    : beededea
' Date      : 18/03/2026
' Purpose   : paint the images using GDI+ extracting the image from the pre-loaded dictionary/imageList
'             populate the hitlist collection with an image and size/location parameters
'             populate the event collection
'---------------------------------------------------------------------------------------
'
Public Sub addSingleImagesToFullScreenDisplay()
    
    On Error GoTo addSingleImagesToFullScreenDisplay_Error

    Call addThisImage(widgetFormName, "tardis", 750, 250, 200, 200, "tardis", 100, vbNullString, True)
    Call addThisImage(widgetFormName, "player", 950, 250, 200, 200, "player", 100, vbNullString, True)

    On Error GoTo 0
    Exit Sub

addSingleImagesToFullScreenDisplay_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure addSingleImagesToFullScreenDisplay of Module modMain"
End Sub


'---------------------------------------------------------------------------------------
' Procedure : addThisImage
' Author    : beededea
' Date      : 27/03/2026
' Purpose   : From the previously populated image list, creates an image of type cImageGDIP with associated properties,
'             then adds hit testing and event handling for each layer
'---------------------------------------------------------------------------------------
'
Private Sub addThisImage(ByVal thisWidgetFormName As String, ByVal thisKey As String, ByVal thisX As Long, ByVal thisY As Long, ByVal thisWidth As Long, ByVal thisHeight As Long, ByVal thisName As String, ByVal thisOpacity As Integer, ByVal thisTooltip As String, ByVal thisRefresh As Boolean)
    
    Dim thisBitmap As Long: thisBitmap = 0
    
    On Error GoTo addThisImage_Error
    
    ' extract a bitmap from the previously populated image list
    thisBitmap = readImageFromDictionary(thisKey)
    
    ' creates an image of type cImageGDIP with associated properties
    With thisGDIPimage
        .bitmap = thisBitmap
        .Left = thisX
        .Top = thisY
        .Width = thisWidth
        .Height = thisHeight
        .Name = thisName
        .Opacity = thisOpacity
        .Tooltip = thisTooltip
        If thisRefresh = True Then .Refresh
    End With
    
    ' adds hit testing and event handling for each layer
    Call addImagesToHitAndEventCollections(thisBitmap, thisName, thisX, thisY, thisWidth, thisHeight)

    On Error GoTo 0
    Exit Sub

addThisImage_Error:

    With Err
         If .Number <> 0 Then
            MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure addThisImage of Form Form1"
            Resume Next
          End If
    End With
    
End Sub



' ----------------------------------------------------------------
' Procedure : InitialiseImageSurfacesFromXML
' Purpose   :  Creates a GDIP surface object from each and every PSD layer in the PSD file.
'              For all the interactive UI elements it creates surface objects with corresponding keynames,
'              locations and sizes as per the original PSD for each layer. The surfaces are populated from PNGs whose metric
'              data are held in an extracted XML file. It creates an instance of each image using the cImageGDIP class
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
Public Sub InitialiseImageSurfacesFromXML()

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
    Dim Height As Long: Height = 0
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

    On Error GoTo InitialiseImageSurfacesFromXML_Error
    
    someOpacity = Val(sOpacity) / 100
    xmlFileToLoad = App.Path & "\RES\CPUimagesXML.xml"
    'xmlFileToLoad = App.Path & "\RES\clockcalendarimagesXML.xml"
    
    If fFExists(xmlFileToLoad) Then
        objxmldoc.Load xmlFileToLoad
    Else
        MsgBox "The XML file that contains the image data is missing " & xmlFileToLoad
    End If
    
    ' obtain the overall widget width and height
    Set MainNode = objxmldoc.selectSingleNode("widget/window")
            
    windowWidth = MainNode.selectSingleNode("@width").Text
    windowHeight = MainNode.selectSingleNode("@height").Text

    ' get the image values from the XML data, the num results should be non-zero
    Set nodeList = objxmldoc.selectNodes("widget/window/image")
    num_results = nodeList.Length
    
    ' no results found, go on as normal using the sampling interval
    If num_results = 0 Then
        answerMsg = "1. There is a problem with the XML data file that describes the image, seems to contain no valid data"
        'msgBoxA answerMsg, vbOKOnly + vbExclamation, "XML Warning", True, "InitialiseImageSurfacesFromXMLPollingWarning"
        MsgBox answerMsg
        Exit Sub ' Return
    End If

    If Not nodeList Is Nothing Then
         For Each node In nodeList
         
            On Error Resume Next
            
            sSrc = node.selectSingleNode("@src").Text
            sSrc = Replace(sSrc, "/", "\")
            sName = node.selectSingleNode("@name").Text
                        
            sWidth = node.selectSingleNode("@width").Text
            If sWidth <> "" Then
                If InStr(sWidth, " px") Then
                    Width = CLng(Left$(sWidth, (InStr(sWidth, " px") - 1)))
                Else
                    Width = CLng(sWidth)
                End If
            Else
                Width = 0
            End If
            
            sHeight = node.selectSingleNode("@height").Text
            If sHeight <> "" Then
                If InStr(sHeight, " px") Then
                    Height = CLng(Left$(sHeight, (InStr(sHeight, " px") - 1)))
                Else
                    Height = CLng(sHeight)
                End If
            Else
                Height = 0
            End If
            hOffset = CLng(node.selectSingleNode("@hOffset").Text)
            vOffset = CLng(node.selectSingleNode("@vOffset").Text)
            Opacity = CInt(node.selectSingleNode("@opacity").Text) / 2.55
            
            On Error GoTo InitialiseImageSurfacesFromXML_Error
            
            If Opacity > 0 Then ' only handles layers that have an opacity greater than 0 - need to note this for the future, this will cause a problem!
            
               'add each current Layer path and surface object into the global ImageList collection (using LayerPath as the ImageKey)
                pngFileToLoad = App.Path & "\Res\" & sSrc
                If fFExists(pngFileToLoad) Then
                    
                    ' add the image named in the XML to the image list
                    thisImageList.AddImage sName, pngFileToLoad
                        
                    ' create an image object and write it to the full screen bitmap
                    Call addThisImage(widgetFormName, sName, hOffset, vOffset, Width, Height, sName, Opacity, vbNullString, True)

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

InitialiseImageSurfacesFromXML_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure InitialiseImageSurfacesFromXML, line " & Erl & "."

End Sub



'---------------------------------------------------------------------------------------
' Procedure : addImagesToHitAndEventCollections
' Author    : beededea
' Date      : 24/03/2026
' Purpose   : This is the clever bit that adds a single image bitmap to two collections, one for hit testing and capturing
'             event handling.
'
'---------------------------------------------------------------------------------------
'
Public Sub addImagesToHitAndEventCollections(ByVal bmp As Long, ByVal thisName As String, ByVal x As Long, ByVal y As Long, ByVal w As Long, ByVal h As Long)

    Dim img As cImageGDIP
    Dim eventHost As cImageEventHost

    On Error GoTo addImagesToHitAndEventCollections_Error

    Set img = New cImageGDIP
    
    ' add the image bitmap to a collection complete with the size, location characteristics to allow hit testing
    
    img.bitmap = bmp
    img.Left = x
    img.Top = y
    img.Width = w
    img.Height = h
    
    ' lock the bitmap copy to allow alpha hit-testing
    img.LockBitmap
    
    ' store in global hit-testing collection
    gHitTestCollection.Add img
    
                            
    ' Now add the same image bitmap to a separate collection to act as a sink host for event trapping on each distinct image layer
    
    ' create event host of class cImageEventHost
    Set eventHost = New cImageEventHost
    
    ' use an eventhost bridge to give 'withEvents' to the incoming image bitmap, enabling capture of events for each image bitmap in the event collection, ultimately of type cImageGDIP (see cImageEventHost)
    Set eventHost.bubblingEventImg = img
    
    ' pass the index to the class to allow layer identification by ID number.
    eventHost.Index = mEventHostCollection.Count + 1
    
    eventHost.Name = thisName
    
    ' pop each cImageGDIP image wrapped in an event host, now withEvents and event target code into the local event collection,
    mEventHostCollection.Add eventHost

    On Error GoTo 0
    Exit Sub

addImagesToHitAndEventCollections_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure addImagesToHitAndEventCollections of Form Form1"

End Sub


'---------------------------------------------------------------------------------------
' Procedure : addThisForm
' Author    : beededea
' Date      : 27/03/2026
' Purpose   : creates an Form of type cWidgetForm with associated properties,

'---------------------------------------------------------------------------------------
'
Private Sub addThisForm(ByVal thisName As String, ByVal thisX As Long, ByVal thisY As Long, ByVal thisWidth As Long, ByVal thisHeight As Long, ByVal thisOpacity As Integer, ByVal thisTooltip As String, ByVal thisRefresh As Boolean)
    
    On Error GoTo addThisForm_Error
        
    ' creates a widget of type cWidgetForm with associated properties
    With thisWidget
        .Left = thisX
        .Top = thisY
        .Width = thisWidth
        .Height = thisHeight
        .Name = thisName
        .Opacity = thisOpacity
        .Tooltip = "Test widget"
        If thisRefresh = True Then .Refresh
    End With

    On Error GoTo 0
    Exit Sub

addThisForm_Error:

    With Err
         If .Number <> 0 Then
            MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure addThisForm of Form Form1"
            Resume Next
          End If
    End With
    
End Sub
