VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cWidgetForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : clsWidgetForm
' Author    : beededea
' Date      : 20/11/2025
' Purpose   : A class that encapsulates the properties of a 'widgetForm'.
'
'---------------------------------------------------------------------------------------

Option Explicit

' local member property holders for this class

Private mLeft As Long
Private mTop As Long
Private mHeight As Long
Private mWidth As Long

Private mOpacity As Integer
Private mName As String
Private mVisible As Boolean
Private mEnabled As Boolean

Private mScaleSize As Single
Private mRotate As Single
Private mTooltip As String

Private mRotationPointX As Long
Private mRotationPointY As Long

Private mLocked As Boolean

'Private mRefresh As Boolean ' function

' public mouseDown events for an object of this class
Public Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)


Implements IVisual

Private m_Base As clsVisualBase


Public Sub AddChild(ByVal child As IVisual)
    m_Base.AddChild child
End Sub

Public Property Let x(ByVal v As Long): m_Base.x = v: End Property
Public Property Let y(ByVal v As Long): m_Base.y = v: End Property

'========================
' Interface forwarding
'========================
Private Sub IVisual_Render(ByVal hdc As Long)
    m_Base.Render hdc
End Sub

Private Function IVisual_GetAbsoluteX() As Long
    IVisual_GetAbsoluteX = m_Base.GetAbsoluteX
End Function

Private Function IVisual_GetAbsoluteY() As Long
    IVisual_GetAbsoluteY = m_Base.GetAbsoluteY
End Function

Private Function IVisual_HitTest(ByVal x As Long, ByVal y As Long) As IVisual
    Set IVisual_HitTest = m_Base.HitTest(x, y)
End Function

Private Sub IVisual_MouseDown(Button As Integer, Shift As Integer, x As Long, y As Long)
    Debug.Print "Form received click"
End Sub


'---------------------------------------------------------------------------------------
' Procedure : Class_Initialize
' Author    : beededea
' Date      : 13/12/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub Class_Initialize()

    On Error GoTo Class_Initialize_Error

     mLeft = 0
     mTop = 0
     mHeight = 0
     mWidth = 0
     mOpacity = 100
     mName = vbNullString
     mVisible = True
     mEnabled = True
     mScaleSize = 1
     mRotate = 0
     mTooltip = vbNullString
     
         Set m_Base = New clsVisualBase

    On Error GoTo 0
    Exit Sub

Class_Initialize_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Class_Initialize of Class Module clsWidgetForm"
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
      
    
    On Error GoTo 0
    Exit Sub

Class_Terminate_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Class_Terminate of Class Module clsWidgetForm"
End Sub




'
'---------------------------------------------------------------------------------------
' Property  : Left
' Author    : beededea
' Date      : 20/11/2025
' Purpose   : The widget's X/Hoffset position
'---------------------------------------------------------------------------------------
'
Public Property Get Left() As Long
   On Error GoTo LeftGet_Error

   Left = mLeft

   On Error GoTo 0
   Exit Property

LeftGet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property Left of Class Module clsWidgetForm"
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

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property Left of Class Module clsWidgetForm"
End Property


'
'---------------------------------------------------------------------------------------
' Property  : Top
' Author    : beededea
' Date      : 20/11/2025
' Purpose   : The widget's Y/Voffset position
'---------------------------------------------------------------------------------------
'
Public Property Get Top() As Long
   On Error GoTo TopGet_Error

   Top = mTop

   On Error GoTo 0
   Exit Property

TopGet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property Top of Class Module clsWidgetForm"
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

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property Top of Class Module clsWidgetForm"
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

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property Enabled of Class Module clsWidgetForm"
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

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property Enabled of Class Module clsWidgetForm"
End Property

'
'---------------------------------------------------------------------------------------
' Property  : Height
' Author    : beededea
' Date      : 20/11/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get Height() As Long
   On Error GoTo HeightGet_Error

   Height = mHeight

   On Error GoTo 0
   Exit Property

HeightGet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property Height of Class Module clsWidgetForm"
End Property
'---------------------------------------------------------------------------------------
' Property  : Height
' Author    : beededea
' Date      : 20/11/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let Height(ByVal newValue As Long)
   On Error GoTo HeightLet_Error

    If mHeight <> newValue Then mHeight = newValue Else Exit Property
    
    If mHeight = True Then
    
    Else
    
    End If

   On Error GoTo 0
   Exit Property

HeightLet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property Height of Class Module clsWidgetForm"
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

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property Width of Class Module clsWidgetForm"
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

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property Width of Class Module clsWidgetForm"
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

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property Opacity of Class Module clsWidgetForm"
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

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property Opacity of Class Module clsWidgetForm"
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

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property Name of Class Module clsWidgetForm"
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

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property Name of Class Module clsWidgetForm"
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

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Function Refresh of Class Module clsWidgetForm"
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

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property Visible of Class Module clsWidgetForm"
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

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property Visible of Class Module clsWidgetForm"
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

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property ScaleSize of Class Module clsWidgetForm"
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

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property ScaleSize of Class Module clsWidgetForm"
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

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property Rotate of Class Module clsWidgetForm"
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

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property Rotate of Class Module clsWidgetForm"
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

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property Tooltip of Class Module clsWidgetForm"
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

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property Tooltip of Class Module clsWidgetForm"
End Property





'---------------------------------------------------------------------------------------
' Property  : Locked
' Author    : beededea
' Date      : 20/11/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get Locked() As Long
   On Error GoTo LockedGet_Error

   Locked = mLocked

   On Error GoTo 0
   Exit Property

LockedGet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property Locked of Class Module clsWidgetForm"
End Property
'---------------------------------------------------------------------------------------
' Property  : Locked
' Author    : beededea
' Date      : 20/11/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let Locked(ByVal newValue As Long)
   On Error GoTo LockedLet_Error

    If mLocked <> newValue Then mLocked = newValue Else Exit Property

   On Error GoTo 0
   Exit Property

LockedLet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property Locked of Class Module clsWidgetForm"
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

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property RotationPointX of Class Module clsWidgetForm"
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

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property RotationPointX of Class Module clsWidgetForm"
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

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property RotationPointY of Class Module clsWidgetForm"
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

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property RotationPointY of Class Module clsWidgetForm"
End Property





'---------------------------------------------------------------------------------------
' Procedure : W_Paint
' Author    : beededea
' Date      : 20/11/2025
' Purpose   : The paint event for the widget using GDI+
'---------------------------------------------------------------------------------------
'
Private Sub Paint()
   
    On Error GoTo W_Paint_Error
        
    'draws the selected widget bitmap onto the GDIP full screen
    'Call GdipDrawImageRectI(gdipFullScreenBitmap, mBitmap, mLeft, mTop, mWidth, mHeight)
    
   On Error GoTo 0
   Exit Sub

W_Paint_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure W_Paint of Class Module cwOverlay"
End Sub
 





