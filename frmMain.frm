VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "arse"
   ClientHeight    =   9165
   ClientLeft      =   150
   ClientTop       =   495
   ClientWidth     =   11625
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   9165
   ScaleWidth      =   11625
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "refresh"
      Height          =   630
      Left            =   10290
      TabIndex        =   1
      ToolTipText     =   "Click me to close the window"
      Top             =   7500
      Width           =   1170
   End
   Begin VB.CommandButton Command 
      Caption         =   "Close"
      Height          =   630
      Left            =   10260
      TabIndex        =   0
      ToolTipText     =   "Click me to close the window"
      Top             =   8370
      Width           =   1170
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
Private Sub Command1_Click()

End Sub
