VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBitmap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Credits:
' Based on Steve McMahon's work (vbAccelerator.com).
' I removed some code that's unneeded for this demo,
' cleaned it up a bit and added LoadResource().
' It's a great, simple, general-use class.

Private mDC As Long         ' Memory DC
Private mBitmap As Long     ' Bitmap handle
Private mOldBitmap As Long  ' "Original" Bitmap handle

Private mWidth As Long
Private mHeight As Long

Public Function LoadFile(FileName As String) As Boolean

    ' Clear up previous DC/bitmap
    ClearAll
    
    mBitmap = LoadImage(API_NULL_HANDLE, FileName, IMAGE_BITMAP, 0, 0, LR_LOADFROMFILE)
    
    If (mBitmap <> API_NULL_HANDLE) Then
        LoadFile = LoadBitmapIntoDC
    End If
    
End Function

' Note: that function will NOT function properly when the
' project is run under the IDE, because then the resources
' loaded will be VB's resources, not your app's resources.
' You may want to watch some interesting side-effects though...
Public Function LoadResource(ResourceID As Long) As Boolean
    
    ' Clear up previous DC/bitmap
    ClearAll
    
    mBitmap = LoadImage(App.hInstance, ResourceID, IMAGE_BITMAP, 0, 0, LR_DEFAULTCOLOR)
    
    If (mBitmap <> API_NULL_HANDLE) Then
        LoadResource = LoadBitmapIntoDC
    End If
    
End Function

Private Function LoadBitmapIntoDC() As Boolean
Dim ScreenDC As Long
Dim BitmapData As BITMAP

    ' Create a compatible memory DC to hold the bitmap
    ScreenDC = GetDC(API_NULL_HANDLE)
    mDC = CreateCompatibleDC(ScreenDC)
    ReleaseDC API_NULL_HANDLE, ScreenDC
    
    If (mDC <> API_NULL_HANDLE) Then
        ' If the DC was created successfully,
        ' select the bitmap into it
        mOldBitmap = SelectObject(mDC, mBitmap)
        
        ' Get the dimensions of the bitmap
        GDIGetObject mBitmap, Len(BitmapData), BitmapData
        mWidth = BitmapData.bmWidth
        mHeight = BitmapData.bmHeight
        
        LoadBitmapIntoDC = True
    End If

End Function

Property Get Width() As Long
    Width = mWidth
End Property

Property Get Height() As Long
    Height = mHeight
End Property

Property Get hDC() As Long
    hDC = mDC
End Property

Public Sub Paint(DestDC As Long, _
                 DestX As Long, _
                 DestY As Long)
    
    BitBlt DestDC, DestX, DestY, _
        mWidth, mHeight, mDC, _
        0, 0, vbSrcCopy
        
End Sub

Private Sub ClearAll()
    
    If (mDC <> API_NULL_HANDLE) Then
        If (mBitmap <> API_NULL_HANDLE) Then
            ' Select the original bitmap into the DC,
            ' and delete our bitmap
            SelectObject mDC, mOldBitmap
            DeleteObject mBitmap
            mBitmap = API_NULL_HANDLE
        End If
        
        ' Delete the memory DC
        DeleteObject mDC
        mDC = API_NULL_HANDLE
    End If

End Sub

Private Sub Class_Terminate()
    ClearAll
End Sub
