VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   ClientHeight    =   3135
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   209
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   312
   Begin VB.Timer tmrAnimate 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   150
      Top             =   90
   End
   Begin VB.Label lblForm 
      Caption         =   "This form is made invisible at runtime and unused except for timers"
      Height          =   735
      Left            =   390
      TabIndex        =   0
      Top             =   2220
      Width           =   2145
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



'---------------------------------------------------------------------------------------
' Procedure : Form_Load
' Author    : beededea
' Date      : 18/03/2026
' Purpose   : calls vbFormSetup to allow the program to be VB6 or user-created custom form based
'---------------------------------------------------------------------------------------
'
Private Sub Form_Load()

    On Error GoTo Form_Load_Error
    
    Call vbFormSetup
    
    On Error GoTo 0
    Exit Sub

Form_Load_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Form_Load of Form Form1"

End Sub

'---------------------------------------------------------------------------------------
' Procedure : Form_MouseUp
' Author    : beededea
' Date      : 18/03/2026
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

    On Error GoTo Form_MouseUp_Error

    If Button = 2 Then 'right click to display a menu
        PopupMenu menuForm.mnuMainMenu, vbPopupMenuRightButton
    End If

    On Error GoTo 0
    Exit Sub

Form_MouseUp_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Form_MouseUp of Form Form1"

End Sub

'---------------------------------------------------------------------------------------
' Procedure : tmrAnimate_Timer
' Author    : beededea
' Date      : 18/03/2026
' Purpose   : Timer WILL be replaced by a custom timer, not currently used
'---------------------------------------------------------------------------------------
'
Private Sub tmrAnimate_Timer()


    ' now we paint the image using GDI+ extracting the image from a previously loaded dictionary, in this case Christian Buse's VBA dictionary replacement
    'updateDisplayFromDictionary "tardis", (500), (250), (200), (200)

    ' Calls UpdateLayeredWindow with created GDI bitmap
    'Call updateScreenUsingGDIPBitmap

    'Call drawAlphaPngGDIP(500, 250, 200, 200)
    
    ' now we paint the image using Cairo, Cairo HAS to load from file as the process to get Cairo to load from a collection is rather tricky using VB6 (Cairo requires a callback as input)
    'Call drawAlphaPngCairo(GetDC(0&), hVBFormHwnd, App.Path & "\player.png", 300, 350)

    On Error GoTo tmrAnimate_Timer_Error

    On Error GoTo 0
    Exit Sub

tmrAnimate_Timer_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure tmrAnimate_Timer of Form Form1"
End Sub
