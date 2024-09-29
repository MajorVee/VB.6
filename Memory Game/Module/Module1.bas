Attribute VB_Name = "ThemeInSet"
Public Sub ThemeIN(frm As Form)
    frm.SkinFramework.LoadSkin App.Path & "\Styles\Fresco.msstyles", ""
    frm.SkinFramework.ApplyWindow frm.hwnd
    frm.SkinFramework.ApplyOptions = frm.SkinFramework.ApplyOptions Or xtpSkinApplyMetrics
End Sub
