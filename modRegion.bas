Attribute VB_Name = "modRegion"
Option Explicit

Public Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Public Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Public Declare Function GetRegionDataByte Lib "gdi32" Alias "GetRegionData" (ByVal hRgn As Long, ByVal dwCount As Long, lpRgnData As Byte) As Long
Public Declare Function GetRegionDataLong Lib "gdi32" Alias "GetRegionData" (ByVal hRgn As Long, ByVal dwCount As Long, lpRgnData As Long) As Long
Public Declare Function ExtCreateRegionByte Lib "gdi32" Alias "ExtCreateRegion" (lpXform As Long, ByVal nCount As Long, lpRgnData As Byte) As Long
Public Declare Function OffsetRgn Lib "gdi32" (ByVal hRgn As Long, ByVal X As Long, ByVal Y As Long) As Long

Public Const RGN_OR = 2

Public Function MakeRegion(hDC As Long, Width As Long, Height As Long) As Long
    
    ' The usual "Make region by bitmap" procedure
    ' TransparentColor is total pink &HFF00FF
    
    Dim X As Long, Y As Long, StartLineX As Long
    Dim FullRegion As Long, LineRegion As Long
    Dim TransparentColor As Long
    Dim InFirstRegion As Boolean
    Dim InLine As Boolean  ' Flags whether we are in a non-tranparent pixel sequence
    
    InFirstRegion = True: InLine = False
    X = Y = StartLineX = 0
    
    TransparentColor = 16711935
    
    For Y = 0 To Height - 1
        For X = 0 To Width
            
            If GetPixel(hDC, X, Y) = TransparentColor Or X = Width Then
                ' We reached a transparent pixel
                If InLine Then
                    InLine = False
                    LineRegion = CreateRectRgn(StartLineX, Y, X, Y + 1)
                    
                    If InFirstRegion Then
                        FullRegion = LineRegion
                        InFirstRegion = False
                    Else
                        CombineRgn FullRegion, FullRegion, LineRegion, RGN_OR
                        ' Always clean up your mess
                        DeleteObject LineRegion
                    End If
                End If
            Else
                ' We reached a non-transparent pixel
                If Not InLine Then
                    InLine = True
                    StartLineX = X
                End If
            End If
        Next
    Next
    
    MakeRegion = FullRegion
End Function


' Given a bitmap, this procedure fills a RegionDataType
' with binary data describing its region.
' This can be described as "Serializing" a region.

Public Sub CreateRegionData(EdgeImage As clsBitmap, EdgeRegion As RegionDataType)
Dim WinRegion As Long
Dim Ret As Long

    ' First create the region for the bitmap
    WinRegion = MakeRegion(EdgeImage.hDC, EdgeImage.Width, EdgeImage.Height)
    
    ' Get the size needed for the region data buffer
    EdgeRegion.DataLength = GetRegionDataLong(WinRegion, 0&, ByVal 0&)

    If EdgeRegion.DataLength = 0 Then
        Err.Raise 1, , "Could not create region data"
    Else
        ' Actually get the data into the buffer - a byte array
        ' of the proper size.
        ' You need 32 bytes more, because the API call attaches
        ' a 32-byte structure called RGNDATAHEADER before the
        ' data itself
        ReDim EdgeRegion.RegionData(EdgeRegion.DataLength + 32)
        
        Ret = GetRegionDataByte(WinRegion, EdgeRegion.DataLength, EdgeRegion.RegionData(0))
    End If
    
    DeleteObject WinRegion

End Sub


' Create a new region by binary region data, then "move" the
' region by the given offset, and combine it with a previous
' region, if supplied.
' The region data was previously created by a call to
' CreateRegionData()
Public Sub MakeRegionWithOffset(Region As RegionDataType, _
                                  XOffset As Long, _
                                  YOffset As Long, _
                                  PrevRegion As Long)
    
Dim NewRegion As Long
    
    ' The API call requires the address of the region data,
    ' so we pass the first cell in the array. VB passes arrays
    ' ByRef, so here's our address.
    NewRegion = ExtCreateRegionByte(ByVal 0&, Region.DataLength, Region.RegionData(0))
    
    OffsetRgn NewRegion, XOffset, YOffset
    
    If PrevRegion = API_NULL_HANDLE Then
        PrevRegion = NewRegion
    Else
        CombineRgn PrevRegion, PrevRegion, NewRegion, RGN_OR
        DeleteObject NewRegion
    End If

End Sub
