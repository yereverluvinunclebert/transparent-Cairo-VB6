Attribute VB_Name = "modXMLWidgets"
Option Explicit


'---------------------------------------------------------------------------------------
' Procedure : InitialiseImageWidgetsFromXML
' Author    : olaf schmidt and myself
' Date      : 31/07/2023
' Purpose   :  Using a previously created Cairo imageList, it creates a surface from each and every PSD layer
'              in the PSD file. The surfaces are populated from PNGs whose metric data are held in an XML file.
'              For each previously 'excluded' entry, it adds path X,Y and alpha properties to the excluded collection.
'              It creates an instance of the cwGaugeOverlay class and populates it with the excluded items that will be rendered in the overlay
'              The overlay comprises items that are non-clickable and will not generate events, ie. animated gauge hands, pendulums &c.
'
'              For all the interactive UI elements it creates RichClient widgets with corresponding keynames, locations and sizes as per the original PSD for each layer.
'              NOTE: The XML data is created using the Create Widget 2.0.js using Photoshop and the PSD found in the RES folder.
'              The XML file is renamed from .KON and placed in the RES folder.
'---------------------------------------------------------------------------------------
'
Public Sub InitialiseImageWidgetsFromXML()
    
    
    On Error GoTo InitialiseImageWidgetsFromXML_Error

    'create the Top-Level-Form
        
    ' code to read image XML data instead of directly from PSD, for RichClient5 only
    Call convertImageXMLToRcWidgets
    
    On Error GoTo 0
   Exit Sub

InitialiseImageWidgetsFromXML_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure InitialiseImageWidgetsFromXML of Class Module cfGauge"
End Sub






' ----------------------------------------------------------------
' Procedure Name: convertImageXMLToRcWidgets
' Purpose: code to read XML image data extracted from a PSD file
' Procedure Kind: Sub
' Procedure Access: Private
' Author: beededea
' Date: 28/08/2025
' ----------------------------------------------------------------
Private Sub convertImageXMLToRcWidgets()

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
    Dim objxmldoc As MSXML2.DOMDocument
   
    Set objxmldoc = New MSXML2.DOMDocument
    
    Dim node As MSXML2.IXMLDOMNode
    Dim MainNode As MSXML2.IXMLDOMNode
'    Dim ImageNode As MSXML2.IXMLDOMNode
'    Dim ImageNodes As MSXML2.IXMLDOMNodeList
    
    On Error GoTo convertImageXMLToRcWidgets_Error
    
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
        'msgBoxA answerMsg, vbOKOnly + vbExclamation, "XML Warning", True, "convertImageXMLToRcWidgetsPollingWarning"
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
                    'Cairo.ImageList.AddSurface sName, Cairo.CreateSurface(Width, height, PSSurface, pngFileToLoad)
                    
                    thisImageList.AddImage sName, pngFileToLoad
    
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

convertImageXMLToRcWidgets_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure convertImageXMLToRcWidgets, line " & Erl & "."

End Sub
