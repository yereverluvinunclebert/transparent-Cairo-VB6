VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cImageEventHost"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : cImageEventHost
' Author    : beededea
' Date      : 24/03/2026
' Purpose   : An event bridge helper class between the collection of images and the cfImageGDIP image types to sink events.
'             WithEvents cannot be used with collections dynamically so instead we associate each image in the event host collection
'             individually and give it withEvents in order to capture/enable events for each image distinctly.
'
'---------------------------------------------------------------------------------------

Option Explicit

Public WithEvents img As cfImageGDIP
Attribute img.VB_VarHelpID = -1
Public Index As Long   ' optional (for identification)
Public Name As String  ' optional (for identification)

'---------------------------------------------------------------------------------------
' Procedure : Img_MouseDown
' Author    : beededea
' Date      : 27/03/2026
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub Img_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error GoTo Img_MouseDown_Error

    Debug.Print "Image"; Index; "MouseDown"; x; y
    
    MsgBox ("MouseDown For " & Name & " Layer No. " & Index & ". Co-ord.s X=" & x & " Y=" & y)
    
    'Call mouseDownEvent("tardis", Button, Shift, x, y)

    On Error GoTo 0
    Exit Sub

Img_MouseDown_Error:

    With Err
         If .Number <> 0 Then
            MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Img_MouseDown of Class Module cImageEventHost"
            Resume Next
          End If
    End With
End Sub

'---------------------------------------------------------------------------------------
' Procedure : Img_MouseMove
' Author    : beededea
' Date      : 27/03/2026
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub Img_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error GoTo Img_MouseMove_Error
    
    'MsgBox ("MouseMove For " & Name & " Layer No. " & Index & ". Co-ord.s X=" & x & " Y=" & y)


    'Debug.Print "Image"; Index; "MouseMove"; x; y

    On Error GoTo 0
    Exit Sub

Img_MouseMove_Error:

    With Err
         If .Number <> 0 Then
            MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Img_MouseMove of Class Module cImageEventHost"
            Resume Next
          End If
    End With
End Sub

