VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'
' The class handles skinnable buttons. The button position
' can be supplied in absolute or relative coordiantes (see later).
'

Private mUpImage As New clsBitmap
Private mDownImage As New clsBitmap

Private mWidth As Long
Private mHeight As Long

Private mX As Long
Private mY As Long
Private mForm As Form

Public Property Get Height() As Long
    Height = mHeight
End Property

Public Property Get Width() As Long
    Width = mWidth
End Property

Public Property Get X() As Long
    X = mX
End Property

Public Property Get Y() As Long
    Y = mY
End Property

Public Property Get DownImageHDC() As Long
    DownImageHDC = mDownImage.hDC
End Property

Public Property Get UpImageHDC() As Long
    UpImageHDC = mUpImage.hDC
End Property

' Initializes the button.
'
' If the x/y parameter is negative, it means that the value is
' relative to the right side of the parent form.
' The Paint() methods compute the actual position of the buttons,
' each time the form is redrawn.
'
Public Sub Init(UpImageFileName As String, _
                DownImageFileName As String, _
                X As Long, Y As Long, _
                ParentForm As Form)
    
    mUpImage.LoadFile UpImageFileName
    mDownImage.LoadFile DownImageFileName
    
    mWidth = mUpImage.Width
    mHeight = mUpImage.Height
    mX = X
    mY = Y
    
    Set mForm = ParentForm
    
End Sub

' Test whether the given (x,y) coordinate is inside the
' button area, and return TRUE is so.
' Used by the parent form in its MouseDown/MouseUp events to
' determine if a button was pressed/released.
Public Function HitTest(X As Long, Y As Long) As Boolean

    If (X >= AbsX() And X < AbsX() + mWidth) And _
       (Y >= AbsY() And Y < AbsY() + mHeight) Then
        HitTest = True
    End If

End Function

Public Sub PaintUpImage()
    ' Note that the 'real' x/y values are used
    mUpImage.Paint mForm.hDC, AbsX(), AbsY()
End Sub

Public Sub PaintDownImage()
    mDownImage.Paint mForm.hDC, AbsX(), AbsY()
End Sub

' Get the real X position of the button.
' If the X coordinate is negative, compute AbsX as the
' distance from the right side
Public Property Get AbsX() As Long

    If mX >= 0 Then
        AbsX = mX
    Else
        AbsX = mForm.ScaleWidth + mX
    End If

End Property

Public Property Get AbsY() As Long

    If mY >= 0 Then
        AbsY = mY
    Else
        AbsY = mForm.ScaleHeight + mY
    End If

End Property

