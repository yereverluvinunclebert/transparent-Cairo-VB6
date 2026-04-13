VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IVisual"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : IVisual
' Author    : beededea
' Date      : 07/04/2026
' Purpose   :
'---------------------------------------------------------------------------------------

Option Explicit

Public Parent As IVisual

Public Function GetAbsoluteX() As Long
End Function

Public Function GetAbsoluteY() As Long
End Function

Public Sub Render(ByVal hdc As Long)
End Sub

Public Function HitTest(ByVal x As Long, ByVal y As Long) As IVisual
End Function

Public Sub MouseDown(Button As Integer, Shift As Integer, x As Long, y As Long)
End Sub
