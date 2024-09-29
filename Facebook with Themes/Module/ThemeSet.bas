Attribute VB_Name = "ThemeINSet"
Public Sub ThemeIN(frm As Form)
If frm.List1.Text = "Blue Sea" Then
    frm.SkinFramework.LoadSkin App.Path & "\Styles\BlueSea.msstyles", ""
    frm.SkinFramework.ApplyWindow frm.hwnd
    frm.SkinFramework.ApplyOptions = frm.SkinFramework.ApplyOptions Or xtpSkinApplyMetrics
    
ElseIf frm.List1.Text = "Mac OS-X" Then
    frm.SkinFramework.LoadSkin App.Path & "\Styles\MacOSX.msstyles", ""
    frm.SkinFramework.ApplyWindow frm.hwnd
    frm.SkinFramework.ApplyOptions = frm.SkinFramework.ApplyOptions Or xtpSkinApplyMetrics
       
ElseIf frm.List1.Text = "Cocoy" Then
    frm.SkinFramework.LoadSkin App.Path & "\Styles\CrunchOrange.msstyles", ""
    frm.SkinFramework.ApplyWindow frm.hwnd
    frm.SkinFramework.ApplyOptions = frm.SkinFramework.ApplyOptions Or xtpSkinApplyMetrics

ElseIf frm.List1.Text = "Fresco" Then
    frm.SkinFramework.LoadSkin App.Path & "\Styles\Fresco.msstyles", ""
    frm.SkinFramework.ApplyWindow frm.hwnd
    frm.SkinFramework.ApplyOptions = frm.SkinFramework.ApplyOptions Or xtpSkinApplyMetrics
     
ElseIf frm.List1.Text = "Fusion VS" Then
    frm.SkinFramework.LoadSkin App.Path & "\Styles\FusionVS.msstyles", ""
    frm.SkinFramework.ApplyWindow frm.hwnd
    frm.SkinFramework.ApplyOptions = frm.SkinFramework.ApplyOptions Or xtpSkinApplyMetrics
  
ElseIf frm.List1.Text = "Green Grass" Then
    frm.SkinFramework.LoadSkin App.Path & "\Styles\GreenGrass.msstyles", ""
    frm.SkinFramework.ApplyWindow frm.hwnd
    frm.SkinFramework.ApplyOptions = frm.SkinFramework.ApplyOptions Or xtpSkinApplyMetrics
    
ElseIf frm.List1.Text = "Manzanas" Then
    frm.SkinFramework.LoadSkin App.Path & "\Styles\Manzanas.msstyles", ""
    frm.SkinFramework.ApplyWindow frm.hwnd
    frm.SkinFramework.ApplyOptions = frm.SkinFramework.ApplyOptions Or xtpSkinApplyMetrics
    
ElseIf frm.List1.Text = "Rogue" Then
    frm.SkinFramework.LoadSkin App.Path & "\Styles\Rogue.msstyles", ""
    frm.SkinFramework.ApplyWindow frm.hwnd
    frm.SkinFramework.ApplyOptions = frm.SkinFramework.ApplyOptions Or xtpSkinApplyMetrics
    
ElseIf frm.List1.Text = "Blink" Then
    frm.SkinFramework.LoadSkin App.Path & "\Styles\Blink.msstyles", ""
    frm.SkinFramework.ApplyWindow frm.hwnd
    frm.SkinFramework.ApplyOptions = frm.SkinFramework.ApplyOptions Or xtpSkinApplyMetrics
    
ElseIf frm.List1.Text = "Boost" Then
    frm.SkinFramework.LoadSkin App.Path & "\Styles\Boost.msstyles", ""
    frm.SkinFramework.ApplyWindow frm.hwnd
    frm.SkinFramework.ApplyOptions = frm.SkinFramework.ApplyOptions Or xtpSkinApplyMetrics
    
ElseIf frm.List1.Text = "BumbleBee" Then
    frm.SkinFramework.LoadSkin App.Path & "\Styles\BumbleBee.msstyles", ""
    frm.SkinFramework.ApplyWindow frm.hwnd
    frm.SkinFramework.ApplyOptions = frm.SkinFramework.ApplyOptions Or xtpSkinApplyMetrics
    
ElseIf frm.List1.Text = "Cosmo" Then
    frm.SkinFramework.LoadSkin App.Path & "\Styles\Cosmo.msstyles", ""
    frm.SkinFramework.ApplyWindow frm.hwnd
    frm.SkinFramework.ApplyOptions = frm.SkinFramework.ApplyOptions Or xtpSkinApplyMetrics

ElseIf frm.List1.Text = "Harvest" Then
    frm.SkinFramework.LoadSkin App.Path & "\Styles\Harvest.msstyles", ""
    frm.SkinFramework.ApplyWindow frm.hwnd
    frm.SkinFramework.ApplyOptions = frm.SkinFramework.ApplyOptions Or xtpSkinApplyMetrics
    
ElseIf frm.List1.Text = "Hex" Then
    frm.SkinFramework.LoadSkin App.Path & "\Styles\Hex.msstyles", ""
    frm.SkinFramework.ApplyWindow frm.hwnd
    frm.SkinFramework.ApplyOptions = frm.SkinFramework.ApplyOptions Or xtpSkinApplyMetrics

ElseIf frm.List1.Text = "Trippin" Then
    frm.SkinFramework.LoadSkin App.Path & "\Styles\Trippin.msstyles", ""
    frm.SkinFramework.ApplyWindow frm.hwnd
    frm.SkinFramework.ApplyOptions = frm.SkinFramework.ApplyOptions Or xtpSkinApplyMetrics
    
ElseIf frm.List1.Text = "PinkLoop" Then
    frm.SkinFramework.LoadSkin App.Path & "\Styles\PinkLoop.msstyles", ""
    frm.SkinFramework.ApplyWindow frm.hwnd
    frm.SkinFramework.ApplyOptions = frm.SkinFramework.ApplyOptions Or xtpSkinApplyMetrics
    
ElseIf frm.List1.Text = "Red Dragon" Then
    frm.SkinFramework.LoadSkin App.Path & "\Styles\Red Dragon.msstyles", ""
    frm.SkinFramework.ApplyWindow frm.hwnd
    frm.SkinFramework.ApplyOptions = frm.SkinFramework.ApplyOptions Or xtpSkinApplyMetrics
    
ElseIf frm.List1.Text = "Vincent" Then
    frm.SkinFramework.LoadSkin App.Path & "\Styles\Vincent.msstyles", ""
    frm.SkinFramework.ApplyWindow frm.hwnd
    frm.SkinFramework.ApplyOptions = frm.SkinFramework.ApplyOptions Or xtpSkinApplyMetrics
    
ElseIf frm.List1.Text = "VS7" Then
    frm.SkinFramework.LoadSkin App.Path & "\Styles\VS7.msstyles", ""
    frm.SkinFramework.ApplyWindow frm.hwnd
    frm.SkinFramework.ApplyOptions = frm.SkinFramework.ApplyOptions Or xtpSkinApplyMetrics
Else
    frm.SkinFramework.LoadSkin App.Path & "", ""
    frm.SkinFramework.ApplyWindow frm.hwnd
    frm.SkinFramework.ApplyOptions = frm.SkinFramework.ApplyOptions Or xtpSkinApplyMetrics
End If
End Sub


