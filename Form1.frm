VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3135
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   209
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   312
   Begin VB.Timer tmrAnimate 
      Interval        =   500
      Left            =   1410
      Top             =   2130
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub tmrAnimate_Timer()

    ' now we paint the image using GDI+ extracting the image from a previously loaded dictionary, in this case Christian Buse's VBA dictionary replacement
    Call drawAlphaPngGDIP(500, 250, 200, 200)
    
    ' now we paint the image using Cairo, Cairo HAS to load from file as the process to get Cairo to load from a collection is rather tricky using VB6 (Cairo requires a callback as input)
    Call drawAlphaPngCairo(thisHDC, Form1.hwnd, App.Path & "\tardis.png", 300, 350)

End Sub
