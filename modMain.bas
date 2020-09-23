Attribute VB_Name = "modMain"
Option Explicit

Sub main()

    InitEdgeFileNames
    
    CurrPrefs.SkinName = "Default"
    CurrPrefs.SkinsPath = App.Path + "\Skins\"
    
    AttemptToLoadSkin

    frmWrapper.Show
    
End Sub
