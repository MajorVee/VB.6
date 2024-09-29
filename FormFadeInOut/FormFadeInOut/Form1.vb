Public Class Form1
    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.Opacity = 0
        Timer1.Start()
    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        If Not Button1.Enabled Then
            Me.Opacity += 0.05
            If (Me.Opacity >= 0.95) Then
                Me.Opacity = 1
                Timer1.Stop()
                Timer1.Enabled = False
                Button1.Enabled = True
            End If
        Else
            Me.Opacity -= 0.05
            If (Me.Opacity <= 0) Then
                Timer1.Stop()
                Timer1.Enabled = False
                Application.Exit()
            End If
        End If
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Timer1.Enabled = True
        Timer1.Start()
    End Sub
End Class
