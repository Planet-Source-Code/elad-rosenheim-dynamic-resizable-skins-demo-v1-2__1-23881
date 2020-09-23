Attribute VB_Name = "modEdges"
Option Explicit

Public Enum FormEdges
    FE_TOP_LEFT
    FE_TOP_RIGHT
    FE_BOTTOM_LEFT
    FE_BOTTOM_RIGHT
    FE_TOP_H_SEGMENT
    FE_BOTTOM_H_SEGMENT
    FE_RIGHT_V_SEGMENT
    FE_LEFT_V_SEGMENT
    FE_LAST = FE_LEFT_V_SEGMENT
End Enum

Type RegionDataType
    RegionData() As Byte
    DataLength As Long
End Type

Public EdgeImageFileNames(FE_LAST) As String


Public Sub InitEdgeFileNames()

    EdgeImageFileNames(FE_TOP_LEFT) = "top_left.bmp"
    EdgeImageFileNames(FE_TOP_RIGHT) = "top_right.bmp"
    EdgeImageFileNames(FE_BOTTOM_LEFT) = "bottom_left.bmp"
    EdgeImageFileNames(FE_BOTTOM_RIGHT) = "bottom_right.bmp"
    EdgeImageFileNames(FE_TOP_H_SEGMENT) = "hsegment_top.bmp"
    EdgeImageFileNames(FE_BOTTOM_H_SEGMENT) = "hsegment_bottom.bmp"
    EdgeImageFileNames(FE_RIGHT_V_SEGMENT) = "vsegment_right.bmp"
    EdgeImageFileNames(FE_LEFT_V_SEGMENT) = "vsegment_left.bmp"

End Sub


Public Sub AttemptToLoadSkin()
On Error GoTo AttemptToLoadSkin_ErrHandler
    
    ReadSkinPreferences
    
    frmPad.LoadSkin
    frmPad.SetPadSize
    
    Exit Sub
    
AttemptToLoadSkin_ErrHandler:
    
    MsgBox "Unable to load skin. " & vbCrLf & _
        "Reason: " & Err.Description, vbCritical, "Fatal Error"

End Sub

' This function builds the window region of the form's edges -
' the corners and the sides, using the pre-created regions data
' Each created region is combined with the full window region
Public Sub BuildEdgesRegion(WindowRegion As Long, _
                            EdgeImages() As clsBitmap, _
                            EdgeRegions() As RegionDataType, _
                            NumXSlices As Long, _
                            NumYSlices As Long)
Dim i As Long

    ' Make region for top-left corner. That's an easy one
    MakeRegionWithOffset EdgeRegions(FE_TOP_LEFT), 0, 0, WindowRegion

    ' Top-right corner
    MakeRegionWithOffset EdgeRegions(FE_TOP_RIGHT), _
        EdgeImages(FE_TOP_LEFT).Width + (EdgeImages(FE_TOP_H_SEGMENT).Width * NumXSlices), 0, _
        WindowRegion
    
    ' Bottom-left corner
    MakeRegionWithOffset EdgeRegions(FE_BOTTOM_LEFT), 0, _
        EdgeImages(FE_TOP_LEFT).Height + EdgeImages(FE_LEFT_V_SEGMENT).Height * NumYSlices, _
        WindowRegion
    
    ' Bottom-right corner
    MakeRegionWithOffset EdgeRegions(FE_BOTTOM_RIGHT), _
        EdgeImages(FE_TOP_LEFT).Width + (EdgeImages(FE_TOP_H_SEGMENT).Width * NumXSlices), _
        EdgeImages(FE_TOP_LEFT).Height + EdgeImages(FE_LEFT_V_SEGMENT).Height * NumYSlices, _
        WindowRegion

    ' Create the regions for the top and bottom sides,
    ' By the number of X slices.
    For i = 1 To NumXSlices
        MakeRegionWithOffset EdgeRegions(FE_TOP_H_SEGMENT), EdgeImages(FE_TOP_LEFT).Width + EdgeImages(FE_TOP_H_SEGMENT).Width * (i - 1), 0, WindowRegion
        
        MakeRegionWithOffset EdgeRegions(FE_BOTTOM_H_SEGMENT), _
            EdgeImages(FE_TOP_LEFT).Width + EdgeImages(FE_TOP_H_SEGMENT).Width * (i - 1), _
            EdgeImages(FE_TOP_LEFT).Height + EdgeImages(FE_BOTTOM_LEFT).Height + EdgeImages(FE_LEFT_V_SEGMENT).Height * NumYSlices - EdgeImages(FE_BOTTOM_H_SEGMENT).Height, WindowRegion
    Next

    ' Create the regions for the left and right sides,
    ' By the number of Y slices.
    For i = 1 To NumYSlices
        MakeRegionWithOffset EdgeRegions(FE_LEFT_V_SEGMENT), 0, EdgeImages(FE_TOP_LEFT).Height + EdgeImages(FE_LEFT_V_SEGMENT).Height * (i - 1), WindowRegion
        
        MakeRegionWithOffset EdgeRegions(FE_RIGHT_V_SEGMENT), _
            EdgeImages(FE_TOP_LEFT).Width + EdgeImages(FE_TOP_H_SEGMENT).Width * NumXSlices + EdgeImages(FE_TOP_RIGHT).Width - EdgeImages(FE_RIGHT_V_SEGMENT).Width, _
            EdgeImages(FE_TOP_LEFT).Height + EdgeImages(FE_LEFT_V_SEGMENT).Height * (i - 1), _
            WindowRegion
    Next
    
End Sub

' This fucntion is almost identical to MakeEdgesRegion,
' excepts that it actually DRAWS the edges.
Public Sub DrawEdges(DestForm As Form, _
                     EdgeImages() As clsBitmap, _
                     NumXSlices As Long, NumYSlices As Long, _
                     SimpleStyle As Boolean, _
                     Optional DrawActiveState As Boolean = True)
Dim i As Long
        
    EdgeImages(FE_TOP_LEFT).Paint DestForm.hDC, 0, 0

    EdgeImages(FE_TOP_RIGHT).Paint DestForm.hDC, _
        EdgeImages(FE_TOP_LEFT).Width + EdgeImages(FE_TOP_H_SEGMENT).Width * NumXSlices, 0
    
    EdgeImages(FE_BOTTOM_LEFT).Paint DestForm.hDC, _
        0, EdgeImages(FE_TOP_LEFT).Height + EdgeImages(FE_LEFT_V_SEGMENT).Height * NumYSlices
    
    EdgeImages(FE_BOTTOM_RIGHT).Paint DestForm.hDC, _
        EdgeImages(FE_TOP_LEFT).Width + EdgeImages(FE_TOP_H_SEGMENT).Width * NumXSlices, _
        EdgeImages(FE_TOP_LEFT).Height + EdgeImages(FE_LEFT_V_SEGMENT).Height * NumYSlices
    
    For i = 1 To NumXSlices
        EdgeImages(FE_TOP_H_SEGMENT).Paint DestForm.hDC, _
            EdgeImages(FE_TOP_LEFT).Width + EdgeImages(FE_TOP_H_SEGMENT).Width * (i - 1), 0
        EdgeImages(FE_BOTTOM_H_SEGMENT).Paint DestForm.hDC, _
            EdgeImages(FE_TOP_LEFT).Width + EdgeImages(FE_TOP_H_SEGMENT).Width * (i - 1), _
            EdgeImages(FE_TOP_LEFT).Height + EdgeImages(FE_BOTTOM_LEFT).Height + EdgeImages(FE_LEFT_V_SEGMENT).Height * NumYSlices - EdgeImages(FE_BOTTOM_H_SEGMENT).Height
    Next

    For i = 1 To NumYSlices
        EdgeImages(FE_LEFT_V_SEGMENT).Paint DestForm.hDC, 0, _
            EdgeImages(FE_TOP_LEFT).Height + EdgeImages(FE_LEFT_V_SEGMENT).Height * (i - 1)
        EdgeImages(FE_RIGHT_V_SEGMENT).Paint DestForm.hDC, _
            DestForm.ScaleWidth - EdgeImages(FE_RIGHT_V_SEGMENT).Width, _
            EdgeImages(FE_TOP_LEFT).Height + EdgeImages(FE_LEFT_V_SEGMENT).Height * (i - 1)
    Next
    
End Sub

' Added in v1.2 of demo
'
' Save the region data of all edges to a file. When loading the skin later on,
' we won't have to compute the region again (using MakeRegion()), only load
' the file. This method typically yields about 50% speed increase in loading skins.
' In kewlpAd 1.2, which has more complex skins, it's very noticeable!
'
' NOTE: If you change the transparent areas in a bitmap, the file will be outdated.
' That's why I added the 'Load region data from file' checkbox
'
Public Sub SaveEdgeRegions(EdgeRegions() As RegionDataType, _
                           FileName As String)

Dim i As Long

    Open FileName For Binary As #1

    For i = 0 To FE_LAST
        Put 1, , EdgeRegions(i).DataLength
        Put 1, , EdgeRegions(i).RegionData
    Next

    Close
    
End Sub

' Load the edges' region data from file
Public Function LoadEdgeRegions(EdgeRegions() As RegionDataType, _
                                FileName As String) As Boolean

Dim i As Long
    
    If Dir(FileName) = "" Then Exit Function
    
    Open FileName For Binary As #1
    
    For i = 0 To FE_LAST
        Get 1, , EdgeRegions(i).DataLength
        ReDim EdgeRegions(i).RegionData(EdgeRegions(i).DataLength + 32)
        Get 1, , EdgeRegions(i).RegionData
    Next
    
    Close
    
    LoadEdgeRegions = True

End Function

