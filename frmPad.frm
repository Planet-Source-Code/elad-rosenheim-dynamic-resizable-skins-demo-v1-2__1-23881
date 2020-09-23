VERSION 5.00
Begin VB.Form frmPad 
   BorderStyle     =   0  'None
   Caption         =   ""
   ClientHeight    =   4410
   ClientLeft      =   105
   ClientTop       =   105
   ClientWidth     =   5040
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   294
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   336
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picClientArea 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3495
      Left            =   120
      ScaleHeight     =   3495
      ScaleWidth      =   4095
      TabIndex        =   1
      Top             =   480
      Width           =   4095
      Begin VB.CheckBox chkUseTransFile 
         Caption         =   "Load region data from file"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   1680
         Width           =   2415
      End
      Begin VB.CommandButton cmdExit 
         Caption         =   "Exit"
         Height          =   375
         Left            =   600
         TabIndex        =   4
         Top             =   2400
         Width           =   1215
      End
      Begin VB.ListBox lstSkins 
         Height          =   1230
         ItemData        =   "frmPad.frx":0000
         Left            =   120
         List            =   "frmPad.frx":0002
         TabIndex        =   3
         Top             =   360
         Width           =   2415
      End
      Begin VB.Label lblLoadTime 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1320
         TabIndex        =   7
         Top             =   2040
         Width           =   1215
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Skin load time:"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   2040
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Select a Skin:"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   120
         Width           =   1335
      End
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1320
      TabIndex        =   0
      Top             =   1440
      Width           =   2415
   End
End
Attribute VB_Name = "frmPad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

 ' Current number of horizontal/vertical segments
Dim NumXSlices As Long
Dim NumYSlices As Long

' Minimum needed number of X slices so we don't mess-up
' the button positions
Dim MinXSlices As Long
                        
' Width/height of pad w/o any horizontal segment
Dim BaseXSize As Long
Dim BaseYSize As Long

' Used when resizing the window -
' X/Y distance of the mouse pointer from the form's edge
Dim XDistance As Long
Dim YDistance As Long

' Boolean flags - the current state of the form
Dim InXDrag As Boolean ' In horizontal resize
Dim InYDrag As Boolean ' In vertical resize
Dim InFormDrag As Boolean ' In window drag

Dim NoRedraw As Boolean

' Set to TRUE when in ListSkins(), to prevent lstSkins_Click()
' events from being handled while the list is created
Dim InListSkins As Boolean

' Size of right/bottom segments
Dim XEdgeSize As Single
Dim YEdgeSize As Single

' Handler for window dragging & docking
Dim DockHandler As New clsDockingHandler

' Holds the actual edge skin bitmaps
Dim EdgeImages(FE_LAST) As clsBitmap

' Holds the region data for each of the skin bitmaps
Dim EdgeRegions(FE_LAST) As RegionDataType

Dim WindowRegion As Long ' Current window region

' Custom Exit/Minimize buttons
Dim MyExitButton As New clsButton
Dim MyMinButton As New clsButton

' Default size of client area. Used to compute the number of
' x/y segments needed when the program is loaded
Const DEFAULT_CLIENT_SIZE = 250


Private Sub Form_Load()
    
    Set DockHandler.ParentForm = Me
    
    ListSkins

End Sub


' A mouse button press may initiate form dragging or resizing
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If Button = vbLeftButton Then
    
        ' Test whether the user has pressed a "button",
        ' and show the 'down button' image if so
        If MyExitButton.HitTest(CLng(X), CLng(Y)) Then
            MyExitButton.PaintDownImage
            Exit Sub
        
        ElseIf MyMinButton.HitTest(CLng(X), CLng(Y)) Then
            MyMinButton.PaintDownImage
            Exit Sub
        End If
    
        YDistance = Y - Me.ScaleHeight
        XDistance = X - Me.ScaleWidth
        
        ' If the mouse pointer is on the the bottom edge,
        ' flag Y (vertical) drag
        If Abs(YDistance) < YEdgeSize Then
            InYDrag = True
        End If
        
        ' If the mouse pointer is on the the right edge,
        ' flag X drag. Don't start drag if wer'e in the window
        ' title area
        If Abs(XDistance) < XEdgeSize And _
           Y > EdgeImages(FE_TOP_RIGHT).Height Then
            InXDrag = True
        End If
        
        ' If we're in the window title area, start form draggin'
        If (Y <= EdgeImages(FE_TOP_H_SEGMENT).Height) Then
            DockHandler.StartDockDrag X * Screen.TwipsPerPixelX, _
                Y * Screen.TwipsPerPixelY
            InFormDrag = True
        End If
    
    End If

End Sub


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim NewYSlices As Single
Dim NewXSlices As Single
Dim ShowXResizeCursor As Boolean
Dim ShowYResizeCursor As Boolean
Dim ResizingNeeded As Boolean

    If InFormDrag Then
        ' Continue window draggin'
        DockHandler.UpdateDockDrag X * Screen.TwipsPerPixelX, _
            Y * Screen.TwipsPerPixelY
        Exit Sub
    End If
    
    ' Determine what kind of cursor should be shown
    
    If Abs(Y - Me.ScaleHeight) < YEdgeSize Or InYDrag Then
        ShowYResizeCursor = True
    End If
    
    If (Abs(X - Me.ScaleWidth) < XEdgeSize And _
        Y > EdgeImages(FE_TOP_RIGHT).Height) Or InXDrag Then
        
        ShowXResizeCursor = True
    End If
    
    If ShowXResizeCursor And ShowYResizeCursor Then
        Me.MousePointer = vbSizeNWSE
        
    ElseIf ShowXResizeCursor Then
        Me.MousePointer = vbSizeWE
    
    ElseIf ShowYResizeCursor Then
        Me.MousePointer = vbSizeNS
    
    Else
        Me.MousePointer = vbDefault
    End If

    If InXDrag Then
        ' Compute new number of horizontal segments
        NewXSlices = (X - BaseXSize - XDistance) / EdgeImages(FE_TOP_H_SEGMENT).Width
        If NewXSlices < MinXSlices Then NewXSlices = MinXSlices
        
        ' Check if we should actually do the resize. Not every
        ' slightest mouse drag should cause a resize
        If (NewXSlices - NumXSlices >= 0.5) Or _
           (NewXSlices - NumXSlices < -0.5) Then
            
            NumXSlices = NewXSlices
            ResizingNeeded = True
        End If
    End If

    ' Same handling for vertical resize-drag
    If InYDrag Then
        
        NewYSlices = (Y - BaseYSize - YDistance) / EdgeImages(FE_LEFT_V_SEGMENT).Height
        If NewYSlices < 0 Then NewYSlices = 0
        
        If NewYSlices - NumYSlices >= 0.5 Or _
           (NewYSlices - NumYSlices < -0.5) Then
            
            NumYSlices = NewYSlices
            ResizingNeeded = True
        End If
    End If

    If ResizingNeeded Then SetPadSize
    
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    MyExitButton.PaintUpImage
    MyMinButton.PaintUpImage

    ' Test whether the user has released a "button",
    ' and commit the appropriate action if so
    If MyExitButton.HitTest(CLng(X), CLng(Y)) Then
        End
    ElseIf MyMinButton.HitTest(CLng(X), CLng(Y)) Then
        ' This will cause our form too to minimize
        frmWrapper.WindowState = vbMinimized
    End If

    ' Clear window dragging/resizing flags
    InXDrag = False
    InYDrag = False
    InFormDrag = False

End Sub

Private Sub picClientArea_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.MousePointer = vbDefault
End Sub

Public Sub LoadSkin()
Dim i As Long
Dim FileName As String
Dim PrevXSliceSize As Long, PrevYSliceSize As Long

    ' Save for later. You'll see.
    If Not EdgeImages(0) Is Nothing Then
        PrevXSliceSize = EdgeImages(FE_TOP_H_SEGMENT).Width
        PrevYSliceSize = EdgeImages(FE_LEFT_V_SEGMENT).Height
    End If
    
    ' Initialize bitmaps array
    For i = 0 To FE_LAST
        Set EdgeImages(i) = New clsBitmap
    Next
    
    ' Load skin bitmaps. Check that the files actally  exist
    For i = 0 To FE_LAST
        FileName = CurrPrefs.SkinFullPath & EdgeImageFileNames(i)
        
        If Dir(FileName) = "" Then
            Err.Raise 1, , "Image file " & FileName & " not found!"
                        
        ElseIf EdgeImages(i).LoadFile(FileName) = False Then
            Err.Raise 1, , "Could not load image file: " & FileName
        End If
    Next
    
    ' Set back color according to skin's definition, to match
    ' the skin's "look"
    Me.BackColor = SkinPrefs.BackColor
    picClientArea.BackColor = SkinPrefs.BackColor
    
    ' Prevent the checkbox from flickering when changing back color
    chkUseTransFile.Visible = False
    chkUseTransFile.BackColor = SkinPrefs.BackColor
    chkUseTransFile.Visible = True
    
    ' See documentation in start of file for all those variables
    BaseXSize = EdgeImages(FE_TOP_LEFT).Width + EdgeImages(FE_TOP_RIGHT).Width
    BaseYSize = EdgeImages(FE_TOP_LEFT).Height + EdgeImages(FE_BOTTOM_LEFT).Height

    XEdgeSize = EdgeImages(FE_RIGHT_V_SEGMENT).Width
    YEdgeSize = EdgeImages(FE_BOTTOM_H_SEGMENT).Height

    ' Here we compute how much horizontal/vertical segments
    ' sould be drawn
    If PrevXSliceSize <> 0 Then
        ' Skin was changed, match number of x/y slices
        ' according to the currect/previous sizes of the slices
        NumXSlices = Round(NumXSlices * PrevXSliceSize / EdgeImages(FE_TOP_H_SEGMENT).Width)
        NumYSlices = Round(NumYSlices * PrevYSliceSize / EdgeImages(FE_LEFT_V_SEGMENT).Height)
    Else
        ' Program was just loaded, match number of x/y slices
        ' to the default client width/height
        NumXSlices = Round(DEFAULT_CLIENT_SIZE / EdgeImages(FE_TOP_H_SEGMENT).Width)
        NumYSlices = Round(DEFAULT_CLIENT_SIZE / EdgeImages(FE_LEFT_V_SEGMENT).Height)
    End If
    
    ' Position Client Area
    picClientArea.Top = EdgeImages(FE_TOP_LEFT).Height
    picClientArea.left = EdgeImages(FE_LEFT_V_SEGMENT).Width

    ' Initialize exit/minimize buttons
    MyExitButton.Init _
       CurrPrefs.SkinFullPath & "exitbutton_up.bmp", _
       CurrPrefs.SkinFullPath & "exitbutton_down.bmp", _
       SkinPrefs.ExitButtonX, SkinPrefs.ExitButtonY, _
       Me

    MyMinButton.Init _
       CurrPrefs.SkinFullPath & "minbutton_up.bmp", _
       CurrPrefs.SkinFullPath & "minbutton_down.bmp", _
       SkinPrefs.MinButtonX, SkinPrefs.MinButtonY, _
       Me

    ' Limit minimum number of X slices, in order to allow the
    ' buttons to be drawn correctly
    MinXSlices = FindMinXSlices()
    NumXSlices = IIf(MinXSlices > NumXSlices, MinXSlices, NumXSlices)

    ' Create and store region data for each of the skin bitmaps,
    ' for use whenever creating the window region
    Dim LoadedRegionsFromFile As Boolean
    
    ' If the 'load region data from file' box is checked, try loading region data
    ' from a cache file. if the file does not exist yet, we'll create the regions
    ' and save them - for the next time
    If chkUseTransFile.Value Then
        If LoadEdgeRegions(EdgeRegions, CurrPrefs.SkinFullPath & "trans.dat") Then
            LoadedRegionsFromFile = True
        End If
    End If
    
    If Not LoadedRegionsFromFile Then
        For i = 0 To FE_LAST
            CreateRegionData EdgeImages(i), EdgeRegions(i)
        Next
    
        SaveEdgeRegions EdgeRegions, CurrPrefs.SkinFullPath & "trans.dat"
    End If

End Sub

Private Sub Form_Paint()
    
    If Not NoRedraw Then
        DrawEdges Me, EdgeImages, NumXSlices, NumYSlices, False
    
        MyExitButton.PaintUpImage
        MyMinButton.PaintUpImage
    End If

End Sub

Public Sub SetPadSize()
Dim NewSize As Long
    
    ' We don't want form redraws when in middle of new size
    ' setting, before the new region was set
    NoRedraw = True
    
    ' Compute width/height of form accodring to the number of
    ' x/y slices
    Me.Width = (EdgeImages(FE_TOP_LEFT).Width + EdgeImages(FE_TOP_H_SEGMENT).Width * NumXSlices + EdgeImages(FE_TOP_RIGHT).Width) * Screen.TwipsPerPixelX
    Me.Height = (EdgeImages(FE_TOP_LEFT).Height + EdgeImages(FE_LEFT_V_SEGMENT).Height * NumYSlices + EdgeImages(FE_BOTTOM_LEFT).Height) * Screen.TwipsPerPixelY

    ' Compute size of client area
    
    NewSize = EdgeImages(FE_LEFT_V_SEGMENT).Height * NumYSlices
    picClientArea.Height = NewSize
    
    NewSize = (Me.Width / Screen.TwipsPerPixelX) - EdgeImages(FE_LEFT_V_SEGMENT).Width - EdgeImages(FE_RIGHT_V_SEGMENT).Width
    picClientArea.Width = NewSize

    NoRedraw = False
    
    ' Create new window region. Also triggers a redraw, now that
    ' wer'e done setting the new form shape
    BuildWindowRegion

End Sub

Private Sub BuildWindowRegion()
Dim PrevRegion As Long

    PrevRegion = WindowRegion
    
    ' Create initial region that covers the client area
    WindowRegion = CreateRectRgn(picClientArea.left, picClientArea.Top, picClientArea.left + picClientArea.Width, picClientArea.Top + picClientArea.Height)

    ' Add to it the window region of the form edges
    BuildEdgesRegion WindowRegion, EdgeImages, EdgeRegions, NumXSlices, NumYSlices

    ' Finally - set the full region
    SetWindowRgn Me.hwnd, WindowRegion, True
    
    ' Don't forget - delete old window region
    DeleteObject PrevRegion
    
End Sub

' Fill the list of skins.
' Actually it's a list of directories under App.Path
Private Sub ListSkins()
Dim CurrSkinName As String, SkinPos As Long
Dim i As Long

    InListSkins = True
    
    CurrSkinName = Dir(CurrPrefs.SkinsPath, vbDirectory)
     
    Do While CurrSkinName <> ""
    
        If CurrSkinName <> "." And CurrSkinName <> ".." Then
            If (GetAttr(CurrPrefs.SkinsPath & CurrSkinName) And vbDirectory) Then
                lstSkins.AddItem CurrSkinName
            
                If CurrSkinName = CurrPrefs.SkinName Then
                    SkinPos = i
                End If
                
                i = i + 1
            End If
        End If
        
        CurrSkinName = Dir()
    Loop
    
    ' Visually select 'default' skin
    lstSkins.ListIndex = SkinPos
    InListSkins = False

End Sub

Private Sub lstSkins_Click()
Dim stTime, endTime As Single

    If Not InListSkins Then
        CurrPrefs.SkinName = lstSkins.Text
        
        stTime = Timer
        AttemptToLoadSkin
        endTime = Timer
        lblLoadTime.Caption = Format((endTime - stTime) * 1000, "###.##") & " ms"
    End If

End Sub

Private Sub cmdExit_Click()
    End
End Sub

' Added in v1.1 of demo
' Find out the minimum number of horizontal slices
' that allows the buttons to be drawn correctly
Private Function FindMinXSlices() As Long
Dim MinSize As Long
Dim MinButtonSize As Long, ExitButtonSize As Long
    
    If MyMinButton.X >= 0 Then
        ' Button is attached to top-left corner.
        ' Find out the width of the part of the button that
        ' excceds the top-left part width
        MinButtonSize = MyMinButton.X + MyMinButton.Width - _
            EdgeImages(FE_TOP_LEFT).Width
    Else
        ' Button is attached to top-RIGHT corner (its X value
        ' is relative to the right side).
        ' Find out the width of the part of the button that
        ' excceds the top-right part width
        MinButtonSize = Abs(MyMinButton.X) - _
            EdgeImages(FE_TOP_RIGHT).Width
    End If

    ' Same handling for the exit button
    If MyExitButton.X >= 0 Then
        ExitButtonSize = MyExitButton.X + MyExitButton.Width - _
            EdgeImages(FE_TOP_LEFT).Width
    Else
        ExitButtonSize = Abs(MyExitButton.X) - _
            EdgeImages(FE_TOP_RIGHT).Width
    End If
    
    MinSize = IIf(MinButtonSize > ExitButtonSize, MinButtonSize, ExitButtonSize)

    ' Find out how many slices are needed
    FindMinXSlices = RoundUp(MinSize / EdgeImages(FE_TOP_H_SEGMENT).Width)
    
End Function

' Added in v1.1 of demo
' Given a double number, the function always returns a long
' number that is the rounding UP of the double value
Private Function RoundUp(Number As Double) As Long
    RoundUp = IIf(Number - CLng(Number) <> 0, CLng(Number + 0.5), CLng(Number))
End Function
