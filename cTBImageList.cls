VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cTBImageList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'---------------------------------------------------------------------------------------
' Module    : cTBImageList
' Author    : beededea
' Date      : 04/01/2026
' Purpose   : Replacement of the RichClient imageList using the TB Collection to store and pull various type images.
'             JPG, BMP, PNG tested so far. PNGs will display correctly with transparency in TB
'             using TwinBasic's native controls that have automatic PNG support built in.

' Note: You CANNOT debug nor run this code using VB6. VB6 has NO knowledge of TwinBasic collections
'
' Why is this ImageList useful?

'             * It is a TB collection-backed ImageList with no scripting runtime dependency.
'             * It can store JPG, BMP, PNG files and others too.
'             * It can store images with alpha transparency as TB provides this capability
'             * It can store images of varying size, not just small 16x16 or 32x32 icons as per the old VB6 imageList.
'             * Avoids RichClient dependency.
'             * Avoids GDI+ dependency so potentially platform-independent in TB's multiplatform future.
'             * Avoids runtime obsolescence.
'             * Uses RichClient-familiar syntax to load and unload the images for easy drop-in replacement.
'             * It is quicker to pull images from TB's collection than a dictionary
'             * It is certainly quicker to pull images from a collection than directly from file.
'
' Limitations?
'
'             * You can't currently use a TB collection with some of RichClient's core functions as much of Olaf's code is
'               designed specifically to work behind the scenes with his own RC imageList, eg. CC.RenderSurfaceContent
'
'             Usage:
'
'             Add a public or private variable to a module (BAS) in order to create a new GDI+ image list instance.
'             Public thisImageList As New cTBImageList
'             eg.
'                #If twinbasic Then
'                    ' Wrapper around TwinBasic's collection
'                    Public thisImageList As New cTBImageList
'                #Else
'                    ' GDI+ image list instance
'                    Public thisImageList As New cGdipImageList
'                #End If
'
'             thisImageList.AddImage "about-icon-dark", App.Path & "\Resources\images\about-icon-dark-1010.jpg"
'             Set imgAbout.Picture = thisImageList.Picture("about-icon-dark")
'             thisImageList.Remove "about-icon-dark"
'             dictionaryCount = thisImageList.count
'             If thisImageList.Exists(thiskey) Then ...
'             thisImageList.ImageWidth = 150 ' not implemented yet
'             thisImageList.ImageHeight      ' ditto
'             thisImageList.ImageOpacity = 100 ' ditto
'             thisImageList.Clear  ' clears the imageList
'
'---------------------------------------------------------------------------------------

Option Explicit

Private mCol As Collection ' a standard VB6/TB-style internal collection to enable external enumeration ie. For Each img In TBImageList

' Declarations for the standard TB Collection replacement

Private mDict As Collection

' member property variables
Private mImageWidth As Long
Private mImageHeight As Long
Private mImageOpacity As Integer
Private mImageExists As Boolean

'---------------------------------------------------------------------------------------
' Procedure : Class_Initialize
' Author    : beededea
' Date      : 04/01/2026
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub Class_Initialize()

    On Error GoTo Class_Initialize_Error
    
    Set mDict = New Collection
    
    ' Not yet implemented
    mImageWidth = 0
    mImageHeight = 0
    ImageOpacity = 100
    
    On Error GoTo 0
    Exit Sub

Class_Initialize_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Class_Initialize of Class Module cTBImageList"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : Class_Terminate
' Author    : beededea
' Date      : 04/01/2026
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub Class_Terminate()
    On Error GoTo Class_Terminate_Error

    ' Clear the whole image list
    Clear
        
    Set mDict = Nothing
    
    On Error GoTo 0
    Exit Sub

Class_Terminate_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Class_Terminate of Class Module cTBImageList"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : NewEnum
' Author    : beededea
' Date      : 14/01/2026
' Purpose   : A standard VB6/TB-style internal collection is used to enable external enumeration ie. For Each img In gdipImageList
'---------------------------------------------------------------------------------------
'
Public Property Get NewEnum() As IUnknown
    On Error GoTo NewEnum_Error

    Set NewEnum = mCol.[_NewEnum]

    On Error GoTo 0
    Exit Property

NewEnum_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure NewEnum of Class Module cTBImageList"
End Property

'---------------------------------------------------------------------------------------
' Procedure : Count
' Author    : beededea
' Date      : 04/01/2026
' Purpose   : Obtain the current count of items in the dictionary
'---------------------------------------------------------------------------------------
'
Public Property Get count() As Long
    On Error GoTo Count_Error

    count = mDict.count

    On Error GoTo 0
    Exit Property

Count_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Count of Class Module cTBImageList"
End Property


'---------------------------------------------------------------------------------------
' Procedure : ImageWidth
' Author    : beededea
' Date      : 05/09/2025
' Purpose   : set the ImageHeight in pixels (not twips)
'---------------------------------------------------------------------------------------
'
Public Property Get ImageWidth() As Long
    On Error GoTo ImageWidth_Error

    ImageWidth = mImageWidth

    On Error GoTo 0
    Exit Property

ImageWidth_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure ImageWidth of Class Module cTBImageList"
End Property
'---------------------------------------------------------------------------------------
' Procedure : ImageWidth
' Author    : beededea
' Date      : 05/09/2025
' Purpose   : not yet implemented, experimenting with Cairo to achieve this
'---------------------------------------------------------------------------------------
'
Public Property Let ImageWidth(ByVal newValue As Long)
    On Error GoTo ImageWidth_Error

    If mImageWidth <> newValue Then mImageWidth = newValue Else Exit Property

    On Error GoTo 0
    Exit Property

ImageWidth_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure ImageWidth of Class Module cTBImageList"
End Property


'---------------------------------------------------------------------------------------
' Procedure : ImageHeight
' Author    : beededea
' Date      : 05/09/2025
' Purpose   : get the ImageHeight in pixels (not twips)
'---------------------------------------------------------------------------------------
'
Public Property Get ImageHeight() As Long
    On Error GoTo ImageHeight_Error

    ImageHeight = mImageHeight

    On Error GoTo 0
    Exit Property

ImageHeight_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure ImageHeight of Class Module cTBImageList"
End Property

'---------------------------------------------------------------------------------------
' Procedure : ImageHeight
' Author    : beededea
' Date      : 05/09/2025
' Purpose   : set the ImageHeight in pixels (not twips)
'---------------------------------------------------------------------------------------
'
Public Property Let ImageHeight(ByVal newValue As Long)
    On Error GoTo ImageHeight_Error

    If mImageHeight <> newValue Then mImageHeight = newValue Else Exit Property

    On Error GoTo 0
    Exit Property

ImageHeight_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure ImageHeight of Class Module cTBImageList"
End Property


'---------------------------------------------------------------------------------------
' Procedure : ImageOpacity
' Author    : beededea
' Date      : 05/09/2025
' Purpose   : image Opacity is not yet implemented, experimenting with Cairo to achieve this
'---------------------------------------------------------------------------------------
'
Public Property Get ImageOpacity() As Integer
    On Error GoTo ImageOpacity_Error

    ImageOpacity = mImageOpacity

    On Error GoTo 0
    Exit Property

ImageOpacity_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure ImageOpacity of Class Module cTBImageList"
End Property

'---------------------------------------------------------------------------------------
' Procedure : ImageOpacity
' Author    : beededea
' Date      : 05/09/2025
' Purpose   : image Opacity is not yet implemented, experimenting with Cairo to achieve this
'---------------------------------------------------------------------------------------
'
Public Property Let ImageOpacity(ByVal newValue As Integer)
    On Error GoTo ImageOpacity_Error

    If mImageOpacity <> newValue Then mImageOpacity = newValue Else Exit Property

    On Error GoTo 0
    Exit Property

ImageOpacity_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure ImageOpacity of Class Module cTBImageList"
End Property



'---------------------------------------------------------------------------------------
' Procedure : ImageExists
' Author    : beededea
' Date      : 05/09/2025
' Purpose   : Test to see if the item ImageExists in the dictionary
'---------------------------------------------------------------------------------------
'
Public Property Get ImageExists(key As String) As Boolean
    On Error GoTo ImageExists_Error

        mImageExists = mDict.Exists(key)
        
        ImageExists = mImageExists

    On Error GoTo 0
    Exit Property

ImageExists_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure ImageExists of Class Module cTBImageList"
End Property

'---------------------------------------------------------------------------------------
' Procedure : Remove
' Author    : beededea
' Date      : 04/01/2026
' Purpose   : Remove a single item from the dictionary
'---------------------------------------------------------------------------------------
'
Public Sub Remove(ByVal key As String)
    On Error GoTo Remove_Error

    If mDict.Exists(key) Then
        mDict.Remove key
    End If

    On Error GoTo 0
    Exit Sub

Remove_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Remove of Class Module cTBImageList"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : Clear
' Author    : beededea
' Date      : 04/01/2026
' Purpose   : Clear the whole image list
'---------------------------------------------------------------------------------------
'
Public Sub Clear()
    Dim k As Variant
    
    On Error GoTo Clear_Error

    For Each k In mDict.Keys
        mDict.Remove k
    Next

    On Error GoTo 0
    Exit Sub

Clear_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Clear of Class Module cTBImageList"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : AddImage
' Author    : beededea
' Date      : 04/01/2026
' Purpose   : Adds an image to the dictionary
'---------------------------------------------------------------------------------------
'
Public Sub AddImage(ByVal key As String, ByVal FilePath As String)

    Dim gdipBitmap As Long: gdipBitmap = 0
    Dim pic As StdPicture

    On Error GoTo AddImage_Error

    If mDict.Exists(key) Then Err.Raise 457, , "Key already exists"
    
    If fFExists(FilePath) = False Then
        MsgBox " File not found " & FilePath
        Exit Sub
    End If
    
    ' if the image width or height is non zero then we need to set the size manually (rescale)
    ' in VB6, typically we would use .PaintPicture method but that only works with bitmaps, so is useless here
    ' Instead, we have to use GDI+ to resize the image as that is what we know.
    
    If mImageWidth <> 0 Or mImageHeight <> 0 Then ' resize required
    
        ' read the file as a series of bytes and create a scaled GDI+ bitmap of the image surface using GdipCreateBitmapFromScan0 and context quality attributes, resizing as we go
        gdipBitmap = loadResizeGDIPImage(mDict, key, FilePath, mImageWidth, mImageHeight, , mImageOpacity)
                
        ' convert the Bitmap to a standard picture
        Set pic = fBmpToStdPicture(gdipBitmap, mImageWidth, mImageHeight)
        
        ' add the image to the dictionary collection
        mDict.Add pic, key ' image followed by Key for a TB collection
    
    Else
        ' set the standard picture object to load from a file on disc
        Set pic = LoadPicture(FilePath)
        
        ' add the image to the dictionary collection
        mDict.Add pic, key ' image followed by Key for a TB collection
    End If
    

    
    On Error GoTo 0
    Exit Sub

AddImage_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure AddImage of Class Module cTBImageList"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : Picture
' Author    : beededea
' Date      : 04/01/2026
' Purpose   : extracts the image from the dictionary using similar syntax to Rich Client
'---------------------------------------------------------------------------------------
'
Public Function Picture(ByVal key As String) As StdPicture
    
    On Error GoTo Picture_Error

    If Not mDict.Exists(key) Then Err.Raise 5, , "Image key not found"
    
    Set Picture = mDict(key)

    On Error GoTo 0
    Exit Function

Picture_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Picture of Class Module cTBImageList"
End Function












