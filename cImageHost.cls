VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cImageHost"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : cImageHost
' Author    : beededea
' Date      : 24/03/2026
' Purpose   : event bridge helper class to sink events, WithEvents cannot be used with collections dynamically
'---------------------------------------------------------------------------------------

Option Explicit

Public WithEvents img As cfImageGDIP
Attribute img.VB_VarHelpID = -1
Public Index As Long   ' optional (for identification)

Private Sub Img_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Debug.Print "Image"; Index; "MouseDown"; x; y
    
    MsgBox ("Image " & Index & " MouseDown " & x & " " & y)
End Sub

Private Sub Img_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Debug.Print "Image"; Index; "MouseMove"; x; y
End Sub

