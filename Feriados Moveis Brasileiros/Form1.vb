Public Class Form1

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Try
            For i As Integer = 0 To 20
                For Each x In EasterDate.Feriados(Now.Year + i)
                    TextBox1.AppendText(String.Format("{1}: {2:d}{0}", vbCrLf, x.Item2.PadRight(25), x.Item1))
                Next

                TextBox1.AppendText(vbCrLf)
            Next
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

End Class
