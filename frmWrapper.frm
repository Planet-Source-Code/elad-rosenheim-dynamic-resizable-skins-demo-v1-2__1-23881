VERSION 5.00
Begin VB.Form frmWrapper 
   Caption         =   "DynaDemo"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "frmWrapper.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4680
End
Attribute VB_Name = "frmWrapper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_GotFocus()
    ' When the user presses the taskbar button, this form gets
    ' the focus, so we shift the focus to the frmPad
    frmPad.SetFocus
End Sub

Private Sub Form_Load()
    ' The wrapper window is practically invisible to the user.
    ' Older version of WindowBlinds have a bug that causes this
    ' window to appear.
    Me.Top = -10000
End Sub

Private Sub Form_Resize()
    
    ' Change frmPad's state according to changes made to this
    ' form using the taskbar
    
    If Me.WindowState = vbMinimized Then
        frmPad.Visible = False
    Else
        frmPad.WindowState = Me.WindowState
        frmPad.Visible = True
    End If

End Sub
