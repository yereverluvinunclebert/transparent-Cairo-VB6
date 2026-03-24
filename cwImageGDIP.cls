VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cfImageGDIP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : cwMainImage
' Author    : beededea
' Date      : 20/11/2025
' Purpose   : A currently unused class that will encapsulate the properties of an image.
'
' Usage     : An image is is written to a imagelist elsewhere.
'             The image is extracted from the imageList when required
'             The resulting GDIP bitmap is passed to this class with other characteristics defined and that image is rendered upon the screen.
'---------------------------------------------------------------------------------------

Option Explicit

' local property holders for this class

Private mLeft As Long
Private mTop As Long
Private mHeight As Long
Private mWidth As Long
Private mIndex As Integer
Private mOpacity As Integer
Private mName As String
Private mVisible As Boolean
Private mEnabled As Boolean
'Private mRefresh As Boolean ' function
Private mScaleSize As Single
Private mRotate As Single
Private mTooltip As String
Private mBitmap As Long
Private mRotationPointX As Long
Private mRotationPointY As Long

' public mouseDown event for an object of this class
' VB6 only raises MouseDown/Move/Up automatically for windowed controls it owns.

Public Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)


'---------------------------------------------------------------------------------------
' Procedure : Class_Initialize
' Author    : beededea
' Date      : 13/12/2025
' Purpose   : Only the main animated Images require this.
'---------------------------------------------------------------------------------------
'
Private Sub Class_Initialize()

    'mainImagexo = 0: mainImageyo = 0

    On Error GoTo Class_Initialize_Error

     mLeft = 0
     mTop = 0
     mHeight = 0
     mWidth = 0
     mIndex = 0
     mOpacity = 100
     mName = vbNullString
     mVisible = True
     mEnabled = True
     mScaleSize = 1
     mRotate = 0
     mTooltip = vbNullString
     mBitmap = 0

    On Error GoTo 0
    Exit Sub

Class_Initialize_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Class_Initialize of Class Module cwMainImage"
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
      
    ' The GDIP graphics are now deleted
    Call GdipDeleteGraphics(Bitmap)
    Call GdipDeleteGraphics(mBitmap)
    
    On Error GoTo 0
    Exit Sub

Class_Terminate_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Class_Terminate of Class Module cTBImageList"
End Sub
'---------------------------------------------------------------------------------------
' Property  : Index
' Author    : beededea
' Date      : 20/11/2025
' Purpose   : property to store an index that may be useful (not a control array)
'---------------------------------------------------------------------------------------
'
Public Property Get Index() As Integer
   On Error GoTo IndexGet_Error

   Index = mIndex

   On Error GoTo 0
   Exit Property

IndexGet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property Index of Class Module cwImage"
End Property
'---------------------------------------------------------------------------------------
' Property  : Index
' Author    : beededea
' Date      : 20/11/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let Index(ByVal newValue As Integer)
   On Error GoTo IndexLet_Error

    If mIndex <> newValue Then mIndex = newValue Else Exit Property
    
    If mIndex = True Then
    
    Else
    
    End If

   On Error GoTo 0
   Exit Property

IndexLet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property Index of Class Module cwImage"
End Property
'
'---------------------------------------------------------------------------------------
' Property  : Left
' Author    : beededea
' Date      : 20/11/2025
' Purpose   : The image's X/Hoffset position
'---------------------------------------------------------------------------------------
'
Public Property Get Left() As Long
   On Error GoTo LeftGet_Error

   Left = mLeft

   On Error GoTo 0
   Exit Property

LeftGet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property Left of Class Module cwImage"
End Property
'---------------------------------------------------------------------------------------
' Property  : Left
' Author    : beededea
' Date      : 20/11/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let Left(ByVal newValue As Long)
   On Error GoTo LeftLet_Error

    If mLeft <> newValue Then mLeft = newValue Else Exit Property
    
    If mLeft = True Then
    
    Else
    
    End If




   On Error GoTo 0
   Exit Property

LeftLet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property Left of Class Module cwImage"
End Property


'
'---------------------------------------------------------------------------------------
' Property  : Top
' Author    : beededea
' Date      : 20/11/2025
' Purpose   : The image's Y/Voffset position
'---------------------------------------------------------------------------------------
'
Public Property Get Top() As Long
   On Error GoTo TopGet_Error

   Top = mTop

   On Error GoTo 0
   Exit Property

TopGet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property Top of Class Module cwImage"
End Property
'---------------------------------------------------------------------------------------
' Property  : Top
' Author    : beededea
' Date      : 20/11/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let Top(ByVal newValue As Long)
   On Error GoTo TopLet_Error

    If mTop <> newValue Then mTop = newValue Else Exit Property
    
    If mTop = True Then
    
    Else
    
    End If
    


   On Error GoTo 0
   Exit Property

TopLet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property Top of Class Module cwImage"
End Property


'---------------------------------------------------------------------------------------
' Property  : Enabled
' Author    : beededea
' Date      : 20/11/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get Enabled() As Boolean
   On Error GoTo EnabledGet_Error

   Enabled = mEnabled

   On Error GoTo 0
   Exit Property

EnabledGet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property Enabled of Class Module cwImage"
End Property
'---------------------------------------------------------------------------------------
' Property  : Enabled
' Author    : beededea
' Date      : 20/11/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let Enabled(ByVal newValue As Boolean)
   On Error GoTo EnabledLet_Error

    If mEnabled <> newValue Then mEnabled = newValue Else Exit Property
    
    If mEnabled = True Then
    
    Else
    
    End If
    


   On Error GoTo 0
   Exit Property

EnabledLet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property Enabled of Class Module cwImage"
End Property

'
'---------------------------------------------------------------------------------------
' Property  : Height
' Author    : beededea
' Date      : 20/11/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get height() As Long
   On Error GoTo HeightGet_Error

   height = mHeight

   On Error GoTo 0
   Exit Property

HeightGet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property Height of Class Module cwImage"
End Property
'---------------------------------------------------------------------------------------
' Property  : Height
' Author    : beededea
' Date      : 20/11/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let height(ByVal newValue As Long)
   On Error GoTo HeightLet_Error

    If mHeight <> newValue Then mHeight = newValue Else Exit Property
    
    If mHeight = True Then
    
    Else
    
    End If

   On Error GoTo 0
   Exit Property

HeightLet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property Height of Class Module cwImage"
End Property


'---------------------------------------------------------------------------------------
' Property  : Width
' Author    : beededea
' Date      : 20/11/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get Width() As Long
   On Error GoTo WidthGet_Error

   Width = mWidth

   On Error GoTo 0
   Exit Property

WidthGet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property Width of Class Module cwImage"
End Property

'---------------------------------------------------------------------------------------
' Property  : Width
' Author    : beededea
' Date      : 20/11/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let Width(ByVal newValue As Long)
   On Error GoTo WidthLet_Error

    If mWidth <> newValue Then mWidth = newValue Else Exit Property
    
    If mWidth = True Then
    
    Else
    
    End If

   On Error GoTo 0
   Exit Property

WidthLet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property Width of Class Module cwImage"
End Property



'
'---------------------------------------------------------------------------------------
' Property  : Opacity
' Author    : beededea
' Date      : 20/11/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get Opacity() As Long
   On Error GoTo OpacityGet_Error

   Opacity = mOpacity

   On Error GoTo 0
   Exit Property

OpacityGet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property Opacity of Class Module cwImage"
End Property

'---------------------------------------------------------------------------------------
' Property  : Opacity
' Author    : beededea
' Date      : 20/11/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let Opacity(ByVal newValue As Long)
   On Error GoTo OpacityLet_Error

    If mOpacity <> newValue Then mOpacity = newValue Else Exit Property
    
    If mOpacity = True Then
    
    Else
    
    End If
    


   On Error GoTo 0
   Exit Property

OpacityLet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property Opacity of Class Module cwImage"
End Property




'---------------------------------------------------------------------------------------
' Property  : Name
' Author    : beededea
' Date      : 20/11/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get Name() As String
   On Error GoTo NameGet_Error

   Name = mName

   On Error GoTo 0
   Exit Property

NameGet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property Name of Class Module cwImage"
End Property
'---------------------------------------------------------------------------------------
' Property  : Name
' Author    : beededea
' Date      : 20/11/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let Name(ByVal newValue As String)
   On Error GoTo NameLet_Error

    If mName <> newValue Then mName = newValue Else Exit Property
    


   On Error GoTo 0
   Exit Property

NameLet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property Name of Class Module cwImage"
End Property


'---------------------------------------------------------------------------------------
' Function  : Refresh
' Author    : beededea
' Date      : 20/11/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Function Refresh() As String
   On Error GoTo RefreshGet_Error

    Paint
    
   On Error GoTo 0
   Exit Function

RefreshGet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Function Refresh of Class Module cwImage"
End Function



'---------------------------------------------------------------------------------------
' Property  : Visible
' Author    : beededea
' Date      : 20/11/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get Visible() As Boolean
   On Error GoTo VisibleGet_Error

   Visible = mVisible

   On Error GoTo 0
   Exit Property

VisibleGet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property Visible of Class Module cwImage"
End Property
'---------------------------------------------------------------------------------------
' Property  : Visible
' Author    : beededea
' Date      : 20/11/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let Visible(ByVal newValue As Boolean)
   On Error GoTo VisibleLet_Error

    If mVisible <> newValue Then mVisible = newValue Else Exit Property
    


   On Error GoTo 0
   Exit Property

VisibleLet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property Visible of Class Module cwImage"
End Property





'---------------------------------------------------------------------------------------
' Property  : ScaleSize
' Author    : beededea
' Date      : 20/11/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get ScaleSize() As Single
   On Error GoTo ScaleSizeGet_Error

   ScaleSize = mScaleSize

   On Error GoTo 0
   Exit Property

ScaleSizeGet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property ScaleSize of Class Module cwImage"
End Property
'---------------------------------------------------------------------------------------
' Property  : ScaleSize
' Author    : beededea
' Date      : 20/11/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let ScaleSize(ByVal newValue As Single)
   On Error GoTo ScaleSizeLet_Error

    If mScaleSize <> newValue Then mScaleSize = newValue Else Exit Property
    


   On Error GoTo 0
   Exit Property

ScaleSizeLet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property ScaleSize of Class Module cwImage"
End Property


'---------------------------------------------------------------------------------------
' Property  : Rotate
' Author    : beededea
' Date      : 20/11/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get Rotate() As Single
   On Error GoTo RotateGet_Error

   Rotate = mRotate

   On Error GoTo 0
   Exit Property

RotateGet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property Rotate of Class Module cwImage"
End Property
'---------------------------------------------------------------------------------------
' Property  : Rotate
' Author    : beededea
' Date      : 20/11/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let Rotate(ByVal newValue As Single)
   On Error GoTo RotateLet_Error

    If mRotate <> newValue Then mRotate = newValue Else Exit Property
    


   On Error GoTo 0
   Exit Property

RotateLet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property Rotate of Class Module cwImage"
End Property


'---------------------------------------------------------------------------------------
' Property  : Tooltip
' Author    : beededea
' Date      : 20/11/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get Tooltip() As String
   On Error GoTo TooltipGet_Error

   Tooltip = mTooltip

   On Error GoTo 0
   Exit Property

TooltipGet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property Tooltip of Class Module cwImage"
End Property
'---------------------------------------------------------------------------------------
' Property  : Tooltip
' Author    : beededea
' Date      : 20/11/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let Tooltip(ByVal newValue As String)
   On Error GoTo TooltipLet_Error

    If mTooltip <> newValue Then mTooltip = newValue Else Exit Property
    
   On Error GoTo 0
   Exit Property

TooltipLet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property Tooltip of Class Module cwImage"
End Property


'---------------------------------------------------------------------------------------
' Property  : Bitmap
' Author    : beededea
' Date      : 20/11/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get Bitmap() As Long
   On Error GoTo BitmapGet_Error

   Bitmap = mBitmap

   On Error GoTo 0
   Exit Property

BitmapGet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property Bitmap of Class Module cwImage"
End Property
'---------------------------------------------------------------------------------------
' Property  : Bitmap
' Author    : beededea
' Date      : 20/11/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let Bitmap(ByVal newValue As Long)
   On Error GoTo BitmapLet_Error

    If mBitmap <> newValue Then mBitmap = newValue Else Exit Property
    
   On Error GoTo 0
   Exit Property

BitmapLet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property Bitmap of Class Module cwImage"
End Property


'---------------------------------------------------------------------------------------
' Property  : RotationPointX
' Author    : beededea
' Date      : 20/11/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get RotationPointX() As Long
   On Error GoTo RotationPointXGet_Error

   RotationPointX = mRotationPointX

   On Error GoTo 0
   Exit Property

RotationPointXGet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property RotationPointX of Class Module cwImage"
End Property
'---------------------------------------------------------------------------------------
' Property  : RotationPointX
' Author    : beededea
' Date      : 20/11/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let RotationPointX(ByVal newValue As Long)
   On Error GoTo RotationPointXLet_Error

    If mRotationPointX <> newValue Then mRotationPointX = newValue Else Exit Property
    
   On Error GoTo 0
   Exit Property

RotationPointXLet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property RotationPointX of Class Module cwImage"
End Property


'---------------------------------------------------------------------------------------
' Property  : RotationPointY
' Author    : beededea
' Date      : 20/11/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get RotationPointY() As Long
   On Error GoTo RotationPointYGet_Error

   RotationPointY = mRotationPointY

   On Error GoTo 0
   Exit Property

RotationPointYGet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property RotationPointY of Class Module cwImage"
End Property
'---------------------------------------------------------------------------------------
' Property  : RotationPointY
' Author    : beededea
' Date      : 20/11/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let RotationPointY(ByVal newValue As Long)
   On Error GoTo RotationPointYLet_Error

    If mRotationPointY <> newValue Then mRotationPointY = newValue Else Exit Property
    
   On Error GoTo 0
   Exit Property

RotationPointYLet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property RotationPointY of Class Module cwImage"
End Property

'---------------------------------------------------------------------------------------
' Procedure : HitTest
' Author    : beededea
' Date      : 19/03/2026
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Function HitTest(ByVal x As Long, ByVal y As Long) As Boolean
    On Error GoTo HitTest_Error

    If Not mVisible Then Exit Function
    If Not mEnabled Then Exit Function
    
    ' Compare against your image bounds (and optionally alpha)
     
    If x >= mLeft And x <= (mLeft + mWidth) _
    And y >= mTop And y <= (mTop + mHeight) Then
        HitTest = True
    End If

    On Error GoTo 0
    Exit Function

HitTest_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure HitTest of Class Module cfImageGDIP"
End Function

'---------------------------------------------------------------------------------------
' Procedure : RaiseMouseDown
' Author    : beededea
' Date      : 21/03/2026
' Purpose   :
'---------------------------------------------------------------------------------------
'
Friend Sub RaiseMouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error GoTo RaiseMouseDown_Error

    RaiseEvent MouseDown(Button, Shift, x, y)

    On Error GoTo 0
    Exit Sub

RaiseMouseDown_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure RaiseMouseDown of Class Module cfImageGDIP"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : RaiseMouseMove
' Author    : beededea
' Date      : 21/03/2026
' Purpose   :
'---------------------------------------------------------------------------------------
'
Friend Sub RaiseMouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error GoTo RaiseMouseMove_Error

    RaiseEvent MouseMove(Button, Shift, x, y)

    On Error GoTo 0
    Exit Sub

RaiseMouseMove_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure RaiseMouseMove of Class Module cfImageGDIP"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : ScreenToLocal
' Author    : beededea
' Date      : 19/03/2026
' Purpose   : convert screen from form to image coordinates
'---------------------------------------------------------------------------------------
'
Public Sub ScreenToLocal(ByVal x As Long, ByVal y As Long, ByRef lx As Single, ByRef ly As Single)
    On Error GoTo ScreenToLocal_Error

    lx = x - mLeft
    ly = y - mTop

    On Error GoTo 0
    Exit Sub

ScreenToLocal_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure ScreenToLocal of Class Module cfImageGDIP"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : W_Paint
' Author    : beededea
' Date      : 20/11/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub Paint()
   
    On Error GoTo W_Paint_Error
        
    'draws the selected image bitmap onto the GDIP full screen
    Call GdipDrawImageRectI(gdipFullScreenBitmap, mBitmap, mLeft, mTop, mWidth, mHeight)
    
   On Error GoTo 0
   Exit Sub

W_Paint_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure W_Paint of Class Module cwOverlay"
End Sub
 


