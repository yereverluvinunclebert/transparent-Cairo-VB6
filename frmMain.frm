VERSION 5.00
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   ClientHeight    =   5805
   ClientLeft      =   120
   ClientTop       =   120
   ClientWidth     =   11625
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   5805
   ScaleWidth      =   11625
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picbox1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5295
      Left            =   270
      ScaleHeight     =   5295
      ScaleWidth      =   3765
      TabIndex        =   1
      Top             =   300
      Width           =   3765
   End
   Begin VB.CommandButton Command 
      Caption         =   "Close"
      Height          =   630
      Left            =   10290
      TabIndex        =   0
      ToolTipText     =   "Click me to close the window"
      Top             =   4920
      Width           =   1170
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      Height          =   4200
      Left            =   7170
      Picture         =   "frmMain.frx":0000
      ToolTipText     =   "You should be able to drag the whole form by dragging any of the images "
      Top             =   210
      Width           =   4200
   End
   Begin VB.Image Image2 
      Appearance      =   0  'Flat
      Height          =   4575
      Left            =   4410
      Picture         =   "frmMain.frx":12643
      ToolTipText     =   "You should be able to drag the whole form by dragging any of the images "
      Top             =   210
      Width           =   4575
   End
   Begin VB.Menu mnuPopupMenu 
      Caption         =   "Main Menu"
      Visible         =   0   'False
      Begin VB.Menu mnuClose 
         Caption         =   "Close Widget"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function GetWindowLong Lib "user32.dll" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" ( _
                ByVal hwnd As Long, _
                ByVal crKey As Long, _
                ByVal bAlpha As Byte, _
                ByVal dwFlags As Long) As Long

' global Windows+ constants START
Private Const ULW_ALPHA = &H2
Private Const WS_EX_LAYERED = &H80000
Private Const WS_EX_TRANSPARENT As Integer = &H20
Private Const GWL_EXSTYLE As Long = -20
Private Const LWA_COLORKEY = &H1

Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Const WM_NCLBUTTONDOWN = &HA1
Private Const HTCAPTION = 2

' things to do
' create a surface in the 10/- using the code here as an example

'---------------------------------------------------------------------------------------
' Procedure : Form_Load
' Author    : beededea
' Date      : 27/10/2025
' Purpose   : Creates a transparent form with two images and a close button
'             one of the images is generated using Cairo
'---------------------------------------------------------------------------------------
'
Private Sub Form_Load()
    Dim psfcFrm As cairo_surface_t
    Dim psfcImg As cairo_surface_t
    Dim pCr     As cairo_t
    
    On Error GoTo Form_Load_Error

    'set the transparency of the underlying form using a colour key to define the transparency
    Me.BackColor = vbCyan
    picbox1.BackColor = vbCyan
    
    ' WS_EX_TRANSPARENT makes the form click-through but that applies to ALL controls
    SetWindowLong Me.hwnd, GWL_EXSTYLE, GetWindowLong(Me.hwnd, GWL_EXSTYLE) Or WS_EX_LAYERED ' Or WS_EX_TRANSPARENT
    SetLayeredWindowAttributes Me.hwnd, vbCyan, 0&, LWA_COLORKEY
    
    ' load the image file to the image control
    #If TWINBASIC Then
        Image1.Picture = LoadPicture(App.Path & "\icon.png")
    #Else
        Image1.Picture = LoadPicture(App.Path & "\icon.jpg")
    #End If

    ' create the Cairo surface
    psfcFrm = cairo_win32_surface_create(picbox1.hDC)
    pCr = cairo_create(psfcFrm)

    psfcImg = cairo_image_surface_create_from_png(App.Path & "\tardis.png")
    
    ' paint the PNG image to the surface on the form
    cairo_set_source_surface pCr, psfcImg, 1, 1
    cairo_paint pCr
    
    cairo_surface_destroy psfcImg
    cairo_destroy pCr
    cairo_surface_destroy psfcFrm

    On Error GoTo 0
    Exit Sub

Form_Load_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Form_Load of Form frmMain"
    
End Sub


'---------------------------------------------------------------------------------------
' Procedure : Image1_DblClick
' Author    : beededea
' Date      : 28/10/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub Image1_DblClick()
    On Error GoTo Image1_DblClick_Error

    Call DblClickHandler(Image1.Name)

    On Error GoTo 0
    Exit Sub

Image1_DblClick_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Image1_DblClick of Form frmMain"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : image1_MouseDown
' Author    : beededea
' Date      : 27/10/2025
' Purpose   : capture a mousedown on the only VB6 control to allow it to be dragged without a titlebar
'---------------------------------------------------------------------------------------
'
Private Sub image1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    On Error GoTo image1_MouseDown_Error

    Call MouseDownHandler(Image1.Name, Button, Shift, x, y)

    On Error GoTo 0
    Exit Sub

image1_MouseDown_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure image1_MouseDown of Form frmMain"

End Sub



'---------------------------------------------------------------------------------------
' Procedure : Image2_DblClick
' Author    : beededea
' Date      : 28/10/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub Image2_DblClick()
    On Error GoTo Image2_DblClick_Error

    Call DblClickHandler(Image2.Name)

    On Error GoTo 0
    Exit Sub

Image2_DblClick_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Image2_DblClick of Form frmMain"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : image2_MouseDown
' Author    : beededea
' Date      : 27/10/2025
' Purpose   : capture a mousedown on the only VB6 control to allow it to be dragged without a titlebar
'---------------------------------------------------------------------------------------
'
Private Sub Image2_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    On Error GoTo image2_MouseDown_Error

    Call MouseDownHandler(Image2.Name, Button, Shift, x, y)

    On Error GoTo 0
    Exit Sub

image2_MouseDown_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure image2_MouseDown of Form frmMain"

End Sub


'---------------------------------------------------------------------------------------
' Procedure : Form_MouseDown
' Author    : beededea
' Date      : 14/08/2023
' Purpose   : standard form down event to generate the menu across the board
'---------------------------------------------------------------------------------------
'
Private Sub Form_MouseDown(ByRef Button As Integer, ByRef Shift As Integer, ByRef x As Single, ByRef y As Single)
   On Error GoTo Form_MouseDown_Error

    If Button = 2 Then
        Me.PopupMenu mnuPopupMenu, vbPopupMenuRightButton
    End If

   On Error GoTo 0
   Exit Sub

Form_MouseDown_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Form_MouseDown of Form widgetPrefs"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : mnuClose_Click
' Author    : beededea
' Date      : 27/10/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub mnuClose_Click()
    On Error GoTo mnuClose_Click_Error

    End

    On Error GoTo 0
    Exit Sub

mnuClose_Click_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuClose_Click of Form frmMain"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : Command_Click
' Author    : beededea
' Date      : 27/10/2025
' Purpose   : closes the form
'---------------------------------------------------------------------------------------
'
Private Sub Command_Click()
    On Error GoTo Command_Click_Error

    End

    On Error GoTo 0
    Exit Sub

Command_Click_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Command_Click of Form frmMain"
End Sub


'---------------------------------------------------------------------------------------
' Procedure : picbox1_DblClick
' Author    : beededea
' Date      : 28/10/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub picbox1_DblClick()
    On Error GoTo picbox1_DblClick_Error

    Call DblClickHandler(picbox1.Name)

    On Error GoTo 0
    Exit Sub

picbox1_DblClick_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure picbox1_DblClick of Form frmMain"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : picbox1_MouseDown
' Author    : beededea
' Date      : 27/10/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub picbox1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    On Error GoTo picbox1_MouseDown_Error

    Call MouseDownHandler(picbox1.Name, Button, Shift, x, y)
    

    On Error GoTo 0
    Exit Sub

picbox1_MouseDown_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure picbox1_MouseDown of Form frmMain"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : MouseDownHandler
' Author    : beededea
' Date      : 27/10/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub MouseDownHandler(ByVal ctrlName As String, Optional Button As Integer, Optional Shift As Integer, Optional x As Single, Optional y As Single)
    On Error GoTo MouseDownHandler_Error

    If Button = 2 Then
        Me.PopupMenu mnuPopupMenu, vbPopupMenuRightButton
    Else
        'MsgBox ctrlName
        ReleaseCapture
        SendMessage Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
    End If

    On Error GoTo 0
    Exit Sub

MouseDownHandler_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure MouseDownHandler of Form frmMain"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : DblClickHandler
' Author    : beededea
' Date      : 27/10/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub DblClickHandler(ByVal ctrlName As String)
    On Error GoTo DblClickHandler_Error

    MsgBox ctrlName & " double-clicked!"

    On Error GoTo 0
    Exit Sub

DblClickHandler_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure DblClickHandler of Form frmMain"
End Sub


