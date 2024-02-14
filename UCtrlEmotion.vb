Public Class UCtrlEmotion
    ' Klassisisene meetod kuvatava pildi määramiseks
    Private Sub selectMood(ByVal intMood As Integer)
            pbEmotion.Image = imgLoend.Images(intMood)
        End Sub
        ' Avalik liidesmeetod
        Public Sub setMood(ByVal strMood As String)
        Select Case strMood
            Case "angry"
                Me.selectMood(0)
            Case "cool"
                Me.selectMood(1)
            Case "heart"
                Me.selectMood(2)
            Case "sad"
                Me.selectMood(3)
            Case "shocked"
                Me.selectMood(4)
            Case "amazed"
                Me.selectMood(5)
            Case Else
                Me.selectMood(0)
        End Select
    End Sub
End Class
