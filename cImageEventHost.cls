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
' Purpose   : An event bridge helper class between the collection of images and the cImageGDIP image types to sink events.
'             WithEvents cannot be used with collections dynamically so instead we associate each image in the event host collection
'             individually and give it withEvents in order to capture/enable events for each image distinctly.
'
'---------------------------------------------------------------------------------------

Option Explicit

Public WithEvents bubblingEventImg As cImageGDIP
Attribute bubblingEventImg.VB_VarHelpID = -1
Public Index As Long   ' optional (for identification)
Public Name As String  ' optional (for identification)

'---------------------------------------------------------------------------------------
' Procedure : bubblingEventImg_MouseDown
' Author    : beededea
' Date      : 27/03/2026
' Purpose   : called by cImageGDIP using friend sub raiseMouseDown raiseEvent
'---------------------------------------------------------------------------------------
'
Private Sub bubblingEventImg_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    Dim Sender As String: Sender = vbNullString
    
    On Error GoTo bubblingEventImg_MouseDown_Error
    
    Sender = Name ' just a string at the moment
        
    'Call mouseDownEvent("tardis", Button, Shift, x, y)

    Select Case LCase$(Sender)
        Case "lockbutton"
            Call lockbutton_mouseDown(Sender)
        Case "prefsbutton"
            Call prefsbutton_mouseDown(Sender)
        Case "tickbutton"
            Call tickbutton_mouseDown(Sender)
        Case "startbutton"
            Call startbutton_mouseDown(Sender)
        Case "middlebutton"
            Call middlebutton_mouseDown(Sender)
        Case "stopbutton"
            Call stopbutton_mouseDown(Sender)
        Case "helpbutton"
            Call helpbutton_mouseDown(Sender)
        Case Else
            'Call otherItem_mouseDown(Sender)
    End Select

    On Error GoTo 0
    Exit Sub

bubblingEventImg_MouseDown_Error:

    With Err
         If .Number <> 0 Then
            MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure bubblingEventImg_MouseDown of Class Module cImageEventHost"
            Resume Next
          End If
    End With
End Sub

'---------------------------------------------------------------------------------------
' Procedure : bubblingEventImg_MouseMove
' Author    : beededea
' Date      : 27/03/2026
' Purpose   : called by cImageGDIP using friend sub raiseMouseMove raiseEvent
'---------------------------------------------------------------------------------------
'
Private Sub bubblingEventImg_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim Sender As String: Sender = vbNullString
    
    On Error GoTo bubblingEventImg_MouseMove_Error
    
    Sender = Name ' just a string at the moment
    
    'MsgBox ("MouseMove For " & Name & " Layer No. " & Index & ". Co-ord.s X=" & x & " Y=" & y)

    Select Case LCase$(Sender)
        Case "lockbutton"
            'Call lockbutton_mouseMove(Sender)
        Case "prefsbutton"
            'Call prefsbutton_mouseMove(Sender)
        Case "tickbutton"
            'Call tickbutton_mouseMove(Sender)
        Case "startbutton"
            'Call startbutton_mouseMove(Sender)
        Case "middlebutton"
            'Call middlebutton_mouseMove(Sender)
        Case "stopbutton"
            'Call stopbutton_mouseMove(Sender)
        Case "helpbutton"
            'Call helpbutton_mouseMove(Sender)
        Case Else
            'Call otherItem_mouseMove(Sender)
    End Select

    On Error GoTo 0
    Exit Sub

bubblingEventImg_MouseMove_Error:

    With Err
         If .Number <> 0 Then
            MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure bubblingEventImg_MouseMove of Class Module cImageEventHost"
            Resume Next
          End If
    End With
End Sub

'---------------------------------------------------------------------------------------
' Procedure : bubblingEventImg_MouseUp
' Author    : beededea
' Date      : 27/03/2026
' Purpose   : called by cImageGDIP using friend sub raiseMouseUp raiseEvent
'---------------------------------------------------------------------------------------
'
Private Sub bubblingEventImg_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim Sender As String: Sender = vbNullString
    
    On Error GoTo bubblingEventImg_MouseUp_Error
    
    Sender = Name ' just a string at the moment
    
    'MsgBox ("MouseUp For " & Name & " Layer No. " & Index & ". Co-ord.s X=" & x & " Y=" & y)

    Select Case LCase$(Sender)
        Case "lockbutton"
            Call lockbutton_mouseUp(Sender)
        Case "prefsbutton"
            'Call prefsbutton_MouseUp(Sender)
        Case "tickbutton"
            'Call tickbutton_MouseUp(Sender)
        Case "startbutton"
            'Call startbutton_MouseUp(Sender)
        Case "middlebutton"
            'Call middlebutton_MouseUp(Sender)
        Case "stopbutton"
            'Call stopbutton_MouseUp(Sender)
        Case "helpbutton"
            'Call helpbutton_MouseUp(Sender)
        Case Else
            'Call otherItem_MouseUp(Sender)
    End Select

    On Error GoTo 0
    Exit Sub

bubblingEventImg_MouseUp_Error:

    With Err
         If .Number <> 0 Then
            MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure bubblingEventImg_MouseUp of Class Module cImageEventHost"
            Resume Next
          End If
    End With
End Sub




'---------------------------------------------------------------------------------------
' Procedure : lockbutton_mouseDown
' Author    : beededea
' Date      : 03/08/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub lockbutton_mouseDown(ByRef Sender As String)
   On Error GoTo lockbutton_mouseDown_Error
   
    'MsgBox ("MouseDown For " & Sender)
   
    
'    If gaugeOverlay.Locked = False Then
'        Sender.Widget.Alpha = 0
'    Else
'        Sender.Widget.Alpha = Val(gsOpacity) / 100
'    End If
'    Sender.Widget.Refresh ' this removes the delay in the lockButton depressing
'    Call toggleWidgetLock
    
   On Error GoTo 0
   Exit Sub

lockbutton_mouseDown_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure lockbutton_mouseDown of Class Module cfGauge"

End Sub

'---------------------------------------------------------------------------------------
' Procedure : lockbutton_mouseUp
' Author    : beededea
' Date      : 03/08/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub lockbutton_mouseUp(ByRef Sender As String)
   On Error GoTo lockbutton_mouseUp_Error
   
    MsgBox ("mouseUp For " & Sender)
   
    
'    If gaugeOverlay.Locked = False Then
'        Sender.Widget.Alpha = 0
'    Else
'        Sender.Widget.Alpha = Val(gsOpacity) / 100
'    End If
'    Sender.Widget.Refresh ' this removes the delay in the lockButton depressing
'    Call toggleWidgetLock
    
   On Error GoTo 0
   Exit Sub

lockbutton_mouseUp_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure lockbutton_mouseUp of Class Module cfGauge"

End Sub

'---------------------------------------------------------------------------------------
' Procedure : prefsbutton_mouseDown
' Author    : beededea
' Date      : 03/08/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub prefsbutton_mouseDown(ByRef Sender As String)
'    Dim fileToPlay As String: fileToPlay = vbNullString
   
    On Error GoTo prefsbutton_mouseDown_Error
    
    MsgBox ("MouseDown For " & Sender)
    

'    Sender.Widget.Alpha = 0
'    'Sender.Widget.Visible = False
'    Sender.Widget.Refresh ' this removes the delay in the Button depressing
'    fileToPlay = "winding.wav"
'    If gsEnableSounds = "1" And fFExists(App.Path & "\resources\sounds\" & fileToPlay) Then
'        PlaySound App.Path & "\resources\sounds\" & fileToPlay, ByVal 0&, SND_FILENAME Or SND_ASYNC
'    End If
'   'MsgBox "prefsbutton_mouseDown"
'   Call makeProgramPreferencesAvailable

   On Error GoTo 0
   Exit Sub

prefsbutton_mouseDown_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure prefsbutton_mouseDown of Class Module cfGauge"

End Sub
'---------------------------------------------------------------------------------------
' Procedure : tickbutton_mouseDown
' Author    : beededea
' Date      : 03/08/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub tickbutton_mouseDown(ByRef Sender As String)
'    Dim fileToPlay As String: fileToPlay = vbNullString
'
    On Error GoTo tickbutton_mouseDown_Error
    
    MsgBox ("MouseDown For " & Sender)

'
'    If gaugeOverlay.PointerAnimate = True Then
'        gaugeOverlay.PointerAnimate = False
'        Sender.Widget.Alpha = Val(gsOpacity) / 100
'        gsPointerAnimate = "0"
'    Else
'        gaugeOverlay.PointerAnimate = True
'        Sender.Widget.Alpha = 0 ' sender is the referring control
'        gsPointerAnimate = "1"
'
'    End If
'    Sender.Widget.Refresh ' this removes the delay in the Button depressing
'
'    fileToPlay = "lock.wav"
'    If gsEnableSounds = "1" And fFExists(App.Path & "\resources\sounds\" & fileToPlay) Then
'        PlaySound App.Path & "\resources\sounds\" & fileToPlay, ByVal 0&, SND_FILENAME Or SND_ASYNC
'    End If
'
'    sPutINISetting "Software\PzCPUGauge", "pointerAnimate", gsPointerAnimate, gsSettingsFile
'    widgetPrefs.cmbTickSwitchPref.ListIndex = Val(gsPointerAnimate)
    
   On Error GoTo 0
   Exit Sub

tickbutton_mouseDown_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure tickbutton_mouseDown of Class Module cfGauge"

End Sub


'---------------------------------------------------------------------------------------
' Procedure : stopbutton_mouseDown
' Author    : beededea
' Date      : 03/08/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub stopbutton_mouseDown(ByRef Sender As String)
   
    On Error GoTo stopbutton_mouseDown_Error

    MsgBox ("MouseDown For " & Sender)

'    Sender.Widget.Alpha = 0
'    Sender.Widget.Refresh
           
   On Error GoTo 0
   Exit Sub

stopbutton_mouseDown_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure stopbutton_mouseDown of Class Module cfGauge"

End Sub
'---------------------------------------------------------------------------------------
' Procedure : middlebutton_mouseDown
' Author    : beededea
' Date      : 03/08/2023
' Purpose   : toggles the display of the multi-core VB6 form
'---------------------------------------------------------------------------------------
'
Private Sub middlebutton_mouseDown(ByRef Sender As String)
    On Error GoTo middlebutton_mouseDown_Error

    MsgBox ("MouseDown For " & Sender)

'    Sender.Widget.Alpha = 0
'    Sender.Widget.Refresh
'
'    If gsMultiCoreEnable = "0" Then
'        gsMultiCoreEnable = "1"
'        frmMultiCore.Show
'        Call startAllCpuTimers
'    Else
'        gsMultiCoreEnable = "0"
'        frmMultiCore.Hide
'        Call stopAllCpuTimers
'    End If
'
'    widgetPrefs.chkMultiCoreEnable.Value = Val(gsMultiCoreEnable)
    
   On Error GoTo 0
   Exit Sub

middlebutton_mouseDown_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure middlebutton_mouseDown of Class Module cfGauge"

End Sub




'---------------------------------------------------------------------------------------
' Procedure : otherItem_mouseDown
' Author    : beededea
' Date      : 03/08/2023
' Purpose   : toggles the display of the multi-core VB6 form
'---------------------------------------------------------------------------------------
'
Private Sub otherItem_mouseDown(ByRef Sender As String)
    On Error GoTo otherItem_mouseDown_Error

    MsgBox ("MouseDown For " & Sender)

'    Sender.Widget.Alpha = 0
'    Sender.Widget.Refresh
'
'    If gsMultiCoreEnable = "0" Then
'        gsMultiCoreEnable = "1"
'        frmMultiCore.Show
'        Call startAllCpuTimers
'    Else
'        gsMultiCoreEnable = "0"
'        frmMultiCore.Hide
'        Call stopAllCpuTimers
'    End If
'
'    widgetPrefs.chkMultiCoreEnable.Value = Val(gsMultiCoreEnable)
    
   On Error GoTo 0
   Exit Sub

otherItem_mouseDown_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure otherItem_mouseDown of Class Module cfGauge"

End Sub




'---------------------------------------------------------------------------------------
' Procedure : startbutton_mouseDown
' Author    : beededea
' Date      : 03/08/2023
' Purpose   : toggles the display of the multi-core VB6 form
'---------------------------------------------------------------------------------------
'
Private Sub startbutton_mouseDown(ByRef Sender As String)
    On Error GoTo startbutton_mouseDown_Error

    MsgBox ("MouseDown For " & Sender)

'    Sender.Widget.Alpha = 0
'    Sender.Widget.Refresh
'
'    If gsMultiCoreEnable = "0" Then
'        gsMultiCoreEnable = "1"
'        frmMultiCore.Show
'        Call startAllCpuTimers
'    Else
'        gsMultiCoreEnable = "0"
'        frmMultiCore.Hide
'        Call stopAllCpuTimers
'    End If
'
'    widgetPrefs.chkMultiCoreEnable.Value = Val(gsMultiCoreEnable)
    
   On Error GoTo 0
   Exit Sub

startbutton_mouseDown_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure startbutton_mouseDown of Class Module cfGauge"

End Sub

'---------------------------------------------------------------------------------------
' Procedure : helpbutton_mouseDown
' Author    : beededea
' Date      : 03/08/2023
' Purpose   : toggles the display of the multi-core VB6 form
'---------------------------------------------------------------------------------------
'
Private Sub helpbutton_mouseDown(ByRef Sender As String)
    On Error GoTo helpbutton_mouseDown_Error

    MsgBox ("MouseDown For " & Sender)

'    Sender.Widget.Alpha = 0
'    Sender.Widget.Refresh
'
'    If gsMultiCoreEnable = "0" Then
'        gsMultiCoreEnable = "1"
'        frmMultiCore.Show
'        Call startAllCpuTimers
'    Else
'        gsMultiCoreEnable = "0"
'        frmMultiCore.Hide
'        Call stopAllCpuTimers
'    End If
'
'    widgetPrefs.chkMultiCoreEnable.Value = Val(gsMultiCoreEnable)
    
   On Error GoTo 0
   Exit Sub

helpbutton_mouseDown_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure helpbutton_mouseDown of Class Module cfGauge"

End Sub

