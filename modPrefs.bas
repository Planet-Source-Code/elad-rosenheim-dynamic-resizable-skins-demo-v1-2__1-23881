Attribute VB_Name = "modPrefs"
Option Explicit

' Current program preferences, especially not skin-specific
Type CurrentPreferncesType
    SkinName As String
    SkinsPath As String
    SkinFullPath As String
    ' Add more as you like here
End Type

' Skin-specific preferences. Supplied by the skin's creator
' in a skin.ini file
Type SkinPreferencesType
    BackColor As Long
    ExitButtonX As Long
    ExitButtonY As Long
    MinButtonX As Long
    MinButtonY As Long
    ' Add more as you like here
End Type

Public CurrPrefs As CurrentPreferncesType

Public SkinPrefs As SkinPreferencesType

Private SkinINIFileName As String

' Read skin-specific preferences from skin.ini file
Public Sub ReadSkinPreferences()
    
    CurrPrefs.SkinFullPath = CurrPrefs.SkinsPath + CurrPrefs.SkinName + "\"
    SkinINIFileName = CurrPrefs.SkinFullPath + "skin.ini"
    
    If Dir(SkinINIFileName) = "" Then
        Err.Raise 1, , "Can't find " & SkinINIFileName & "!"
    End If
    
    With SkinPrefs
        .BackColor = ReadColorFromINI("Skin", "BackColor")
        .ExitButtonX = INIRead("Skin", "ExitButtonX", SkinINIFileName)
        .ExitButtonY = INIRead("Skin", "ExitButtonY", SkinINIFileName)
        .MinButtonX = INIRead("Skin", "MinButtonX", SkinINIFileName)
        .MinButtonY = INIRead("Skin", "MinButtonY", SkinINIFileName)
    End With
    
End Sub

' Reads an RGB color string (in the format RRR,GGG,BBB) from
' the skin.ini file, and returns it as long
Private Function ReadColorFromINI(Section As String, Value As String) As Long
Dim ColorStr As String, ColorArr As Variant
    
    ColorStr = INIRead(Section, Value, SkinINIFileName)
    ColorArr = Split(ColorStr, ",")
    
    If ColorStr = "NO_SUCH_KEY" Then
        ReadColorFromINI = 0
        
    ElseIf UBound(ColorArr) <> 2 Then
        Err.Raise 1, , "Invalid color value for attribute """ & Value & """"
        
    Else
        ReadColorFromINI = RGB(ColorArr(0), _
                               ColorArr(1), _
                               ColorArr(2))
    End If

End Function
