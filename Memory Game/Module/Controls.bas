Attribute VB_Name = "Controls"
Public Sub LoadMoviePreview()
frmVideo.wmpMoviePreview.URL = (App.Path & "\Design\Video\Memory Game.mp4")
End Sub
Public Sub LoadMoviePreviewOff()
frmMain.wmpMoviePreview.URL = ""
End Sub
