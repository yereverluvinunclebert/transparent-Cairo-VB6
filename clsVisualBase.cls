VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsVisualBase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements IVisual

Private m_X As Long
Private m_Y As Long
Private m_Width As Long
Private m_Height As Long

Private m_Visible As Boolean
Private m_Children As Collection

Private m_Parent As IVisual

Private Sub Class_Initialize()
    Set m_Children = New Collection
    m_Visible = True
End Sub

'========================
' Properties
'========================
Public Property Get x() As Long
    x = m_X
End Property

Public Property Let x(ByVal Value As Long)
    m_X = Value
End Property

Public Property Get y() As Long
    y = m_Y
End Property

Public Property Let y(ByVal Value As Long)
    m_Y = Value
End Property

Public Property Get Width() As Long
    Width = m_Width
End Property

Public Property Let Width(ByVal Value As Long)
    m_Width = Value
End Property

Public Property Get Height() As Long
    Height = m_Height
End Property

Public Property Let Height(ByVal Value As Long)
    m_Height = Value
End Property

Public Property Set Parent(ByVal obj As IVisual)
    Set m_Parent = obj
End Property

Public Property Get Parent() As IVisual
    Set Parent = m_Parent
End Property

'========================
' Hierarchy
'========================
Public Sub AddChild(ByVal child As IVisual)
    m_Children.Add child
    Set child.Parent = Me
End Sub

Public Property Get ChildCount() As Long
    ChildCount = m_Children.Count
End Property

Public Function GetChild(ByVal index As Long) As IVisual
    Set GetChild = m_Children(index)
End Function

'========================
' Absolute Position
'========================
Public Function GetAbsoluteX() As Long
    If m_Parent Is Nothing Then
        GetAbsoluteX = m_X
    Else
        GetAbsoluteX = m_Parent.GetAbsoluteX + m_X
    End If
End Function

Public Function GetAbsoluteY() As Long
    If m_Parent Is Nothing Then
        GetAbsoluteY = m_Y
    Else
        GetAbsoluteY = m_Parent.GetAbsoluteY + m_Y
    End If
End Function

'========================
' Rendering
'========================
Public Sub RenderAll(ByVal hdc As Long)
    If Not m_Visible Then Exit Sub
    
    ' Call back into owner (important)
    RenderSelf hdc
    
    Dim child As IVisual
    For Each child In m_Children
        child.Render hdc
    Next
End Sub

' This is the "override hook"
Public Sub RenderSelf(ByVal hdc As Long)
    ' Default: do nothing
End Sub

'========================
' Hit Testing
'========================
Public Function HitTestAll(ByVal x As Long, ByVal y As Long) As IVisual
    
    If Not m_Visible Then Exit Function
    
    Dim i As Long
    For i = m_Children.Count To 1 Step -1
        
        Dim child As IVisual
        Set child = m_Children(i)
        
        Dim hit As IVisual
        Set hit = child.HitTest(x, y)
        
        If Not hit Is Nothing Then
            Set HitTestAll = hit
            Exit Function
        End If
        
    Next
    
    If x >= GetAbsoluteX And x <= GetAbsoluteX + m_Width _
    And y >= GetAbsoluteY And y <= GetAbsoluteY + m_Height Then
        
        Set HitTestAll = Me
        
    End If
    
End Function

'========================
' Interface Implementation
'========================
Private Sub IVisual_Render(ByVal hdc As Long)
    RenderAll hdc
End Sub

Private Function IVisual_GetAbsoluteX() As Long
    IVisual_GetAbsoluteX = GetAbsoluteX
End Function

Private Function IVisual_GetAbsoluteY() As Long
    IVisual_GetAbsoluteY = GetAbsoluteY
End Function

Private Function IVisual_HitTest(ByVal x As Long, ByVal y As Long) As IVisual
    Set IVisual_HitTest = HitTestAll(x, y)
End Function

Private Sub IVisual_MouseDown(Button As Integer, Shift As Integer, x As Long, y As Long)
    If Not m_Parent Is Nothing Then
        m_Parent.MouseDown Button, Shift, x, y
    End If
End Sub

