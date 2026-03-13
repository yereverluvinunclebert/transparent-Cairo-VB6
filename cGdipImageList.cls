VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cGdipImageList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : cGdipImageList
' Author    : The vast majority from beededea but the basic class-boilerplate and some bug-fixing from chatGPT
' Date      : 04/01/2026
' Purpose   : Replacement of the RichClient imageList using the scripting.dictionary and GDI+ to store and pull VB6 native type images.
'             JPG, BMP, PNG tested so far. PNGs will display correctly with transparency if displayed using GDI+.
'             Alpha PNGs will display with a black background on any standard VB6 control but should display perfectly
'             using TwinBasic's native controls that have automatic PNG support built in.
'
'             Stores the images as HBitmaps within the chosen collection, extracts them, converts them to standard picture objects.
'
'             Why is this ImageList useful?
'
'             * It is a Dictionary-backed ImageList, when using Cristian Buse's dictionary replacement, no scripting runtime dependency.
'             * It can store JPG, BMP, PNG files and others too.
'             * It can store images with alpha transparency as GDI+ provides this capability
'             * It can store images of varying size, not just small 16x16 or 32x32 icons as per the old VB6 imageList.
'             * Avoids RichClient dependency.
'             * Avoids runtime obsolescence.
'             * Uses dependable GDI+ to load and unload the images.
'             * You have full control as it is FOSS.
'             * The usage syntax is very similar to Rich Client's for easy drop-in replacement.
'             * It is quicker to pull images from the dictionary than directly from file.
'             * If used with Elroy's standard picture Ex project it can parse and render alpha images (PNGs &c) directly to VB6 picture/imageboxes.
'
'             Limitations?
'
'             * You can't currently use this imageList with RichClient as Olaf's code is designed specifically to work with his own imageList, eg. CC.RenderSurfaceContent
'             * VB6 still can't handle PNGs with alpha unless you use something like Elroy's standard picture Ex project.
'
'             Usage:
'
'             Add a public or private variable to a module (BAS) in order to create a new GDI+ image list instance.
'             Public thisImageList As New cGdipImageList
'
'             Add a public variable to a module (BAS) to provide an instance counter for each usage of the class.
'             Public gGdipImageListInstanceCount As Long
'
'             thisImageList.AddImage "about-icon-dark", App.Path & "\Resources\images\about-icon-dark-1010.jpg"
'             Set imgAbout.Picture = thisImageList.Picture("about-icon-dark")
'             thisImageList.Remove "about-icon-dark"
'             dictionaryCount = thisImageList.count
'             If thisImageList.Exists(thiskey) Then ...
'             thisImageList.ImageWidth = 150 ' sets the imageWidth, a value of zero causes the image's real width to be used
'             thisImageList.ImageHeight ' ditto
'             thisImageList.ImageOpacity = 100 ' 0-100% prior to loading will set the opacity
'             thisImageList.Clear  ' clears the imageList
'
'---------------------------------------------------------------------------------------

Option Explicit

' ===========================
' Win32 / GDI
' ===========================


Private mCol As Collection ' a standard VB6/TB-style internal collection to enable external enumeration ie. For Each img In gdipImageList

'Private mDict As Object ' You will need a project reference to the Scripting.Dictionary

'-----------------------------------------------------------
' Declarations for the scripting.dictionary replacement START
'-----------------------------------------------------------

Private Type JustMyUDT
    dict As Dictionary
End Type

#Const x64StackIssues = ((Mac = 0) And Win64 And (twinbasic = 0))

#If twinbasic Then
    Private mDict As Collection ' std TB collection
#Else
    Private mDict As Dictionary ' scripting.dictionary
#End If
'-----------------------------------------------------------
' Declarations for the scripting.dictionary replacements END
'-----------------------------------------------------------

' member property variables
Private mImageWidth As Long
Private mImageHeight As Long
Private mImageOpacity As Integer

'---------------------------------------------------------------------------------------
' Procedure : Class_Initialize
' Author    : beededea
' Date      : 04/01/2026
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub Class_Initialize()

    On Error GoTo Class_Initialize_Error
        
    ' old syntax for the MS Scripting dictionary
    'Set mDict = CreateObject("Scripting.Dictionary")
    
    Set mDict = New Dictionary ' if you are using Cristian Buse's Dictionary replacement for the Scripting.Dictionary
    
    If gGdipImageListInstanceCount = 0 Then Call mGDIPImageList.initialiseGDIPStartup
    
    gGdipImageListInstanceCount = gGdipImageListInstanceCount + 1
    
    ' with image dimensions set to 0, the image's default dimensions will be extracted from the image itself
    mImageWidth = 0
    mImageHeight = 0
    ImageOpacity = 100
    
    On Error GoTo 0
    Exit Sub

Class_Initialize_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Class_Initialize of Class Module cGdipImageList"
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
    
    gGdipImageListInstanceCount = gGdipImageListInstanceCount - 1
        
    ' Only Shutdown GDI+ when there are no users of this class left
    If lngGDI Then
        If gGdipImageListInstanceCount = 0 Then Call mGDIPImageList.shutdownGdiplus
    End If
    
    On Error GoTo 0
    Exit Sub

Class_Terminate_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Class_Terminate of Class Module cGdipImageList"
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

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure NewEnum of Class Module cGdipImageList"
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

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Count of Class Module cGdipImageList"
End Property


'---------------------------------------------------------------------------------------
' Procedure : ImageWidth
' Author    : beededea
' Date      : 05/09/2025
' Purpose   : get the imageWidth in pixels (not twips)
'---------------------------------------------------------------------------------------
'
Public Property Get ImageWidth() As Long
    On Error GoTo ImageWidth_Error

    ImageWidth = mImageWidth

    On Error GoTo 0
    Exit Property

ImageWidth_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure ImageWidth of Class Module cGdipImageList"
End Property
'---------------------------------------------------------------------------------------
' Procedure : ImageWidth
' Author    : beededea
' Date      : 05/09/2025
' Purpose   : set the imageWidth in pixels (not twips)
'---------------------------------------------------------------------------------------
'
Public Property Let ImageWidth(ByVal newValue As Long)
    On Error GoTo ImageWidth_Error

    If mImageWidth <> newValue Then mImageWidth = newValue Else Exit Property

    On Error GoTo 0
    Exit Property

ImageWidth_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure ImageWidth of Class Module cGdipImageList"
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

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure ImageHeight of Class Module cGdipImageList"
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

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure ImageHeight of Class Module cGdipImageList"
End Property


'---------------------------------------------------------------------------------------
' Procedure : ImageOpacity
' Author    : beededea
' Date      : 05/09/2025
' Purpose   : image Opacity is not much use when using VB6 resulting in a black background
'             better for use with TwinBasic
'---------------------------------------------------------------------------------------
'
Public Property Get ImageOpacity() As Integer
    On Error GoTo ImageOpacity_Error

    ImageOpacity = mImageOpacity

    On Error GoTo 0
    Exit Property

ImageOpacity_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure ImageOpacity of Class Module cGdipImageList"
End Property

'---------------------------------------------------------------------------------------
' Procedure : ImageOpacity
' Author    : beededea
' Date      : 05/09/2025
' Purpose   : image Opacity is not much use when using VB6 resulting in a black background
'             better for use with TwinBasic
'---------------------------------------------------------------------------------------
'
Public Property Let ImageOpacity(ByVal newValue As Integer)
    On Error GoTo ImageOpacity_Error

    If mImageOpacity <> newValue Then mImageOpacity = newValue Else Exit Property

    On Error GoTo 0
    Exit Property

ImageOpacity_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure ImageOpacity of Class Module cGdipImageList"
End Property

'---------------------------------------------------------------------------------------
' Procedure : Exists
' Author    : beededea
' Date      : 04/01/2026
' Purpose   : Test to see if the item exists in the dictionary
'---------------------------------------------------------------------------------------
'
Public Function Exists(ByVal key As String) As Boolean
    On Error GoTo Exists_Error

    Exists = mDict.Exists(key)

    On Error GoTo 0
    Exit Function

Exists_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Exists of Class Module cGdipImageList"
End Function

'---------------------------------------------------------------------------------------
' Procedure : Remove
' Author    : beededea
' Date      : 04/01/2026
' Purpose   : Remove a single item from the dictionary
' Note      : Dictionary stores GDI+ Bitmap handles (GpBitmap)
'             Must be destroyed with GdipDisposeImage
'---------------------------------------------------------------------------------------
'
Public Sub Remove(ByVal key As String)
    On Error GoTo Remove_Error

    If mDict.Exists(key) Then
        ' GDI+ delete
        GdipDisposeImage mDict(key) 'NOT DeleteObject mDict(Key)
        mDict.Remove key
    End If

    On Error GoTo 0
    Exit Sub

Remove_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Remove of Class Module cGdipImageList"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : Clear
' Author    : beededea
' Date      : 04/01/2026
' Purpose   : Clear the whole image list
' Note      : Dictionary stores GDI+ Bitmap handles (GpBitmap)
'             Must be destroyed with GdipDisposeImage
'---------------------------------------------------------------------------------------
'
Public Sub Clear()
    Dim k As Variant
    On Error GoTo Clear_Error

    For Each k In mDict.Keys
        ' GDI+ delete
        GdipDisposeImage mDict(k) 'NOT DeleteObject mDict(k)
    Next
    mDict.RemoveAll

    On Error GoTo 0
    Exit Sub

Clear_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Clear of Class Module cGdipImageList"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : GetImage
' Author    : beededea
' Date      : 04/01/2026
' Purpose   : Get a specific image by key from the dictionary
' Returns   : GpBitmap (NOT HBITMAP)
'---------------------------------------------------------------------------------------
'
Public Function GetImage(ByVal key As String) As Long
    On Error GoTo GetImage_Error

    If Not mDict.Exists(key) Then Err.Raise 5, , "Image key not found"
    GetImage = mDict(key)

    On Error GoTo 0
    Exit Function

GetImage_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure GetImage of Class Module cGdipImageList"
End Function

'---------------------------------------------------------------------------------------
' Procedure : AddImage
' Author    : beededea
' Date      : 04/01/2026
' Purpose   : Adds an image to the dictionary
'---------------------------------------------------------------------------------------
'
Public Sub AddImage(ByVal key As String, ByVal FilePath As String)

    On Error GoTo AddImage_Error

    If mDict.Exists(key) Then Err.Raise 457, , "Key already exists"
    
    If fFExists(FilePath) Then
        resizeAndLoadImgToDict mDict, key, FilePath, mImageWidth, mImageHeight, , mImageOpacity
    Else
        MsgBox " File not found " & FilePath
    End If

    On Error GoTo 0
    Exit Sub

AddImage_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure AddImage of Class Module cGdipImageList"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : Picture
' Author    : beededea
' Date      : 04/01/2026
' Purpose   : extracts the image from the dictionary using similar syntax to Rich Client
'---------------------------------------------------------------------------------------
'
Public Function Picture(ByVal key As String) As StdPicture
    Dim gpImage As Long: gpImage = 0
    
    On Error GoTo Picture_Error

    If Not mDict.Exists(key) Then Err.Raise 5, , "Image key not found"
    
    gpImage = GetImage(key)
    Set Picture = fBmpToStdPicture(gpImage, mImageWidth, mImageHeight)

    On Error GoTo 0
    Exit Function

Picture_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Picture of Class Module cGdipImageList"
End Function
