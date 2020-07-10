Public Class Form5
    Public frm1 As Form1
    Public Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim s() As String = TextBox1.Text.Split({vbCrLf}, StringSplitOptions.RemoveEmptyEntries)
        Dim t() As String
        For i As Long = 0 To frm1.mw - 1
            For j As Long = 0 To frm1.mw - 1
                frm1.Mask(i, j) = 1
            Next
        Next
        For i As Long = 0 To s.Length - 1
            t = s(i).Split({","}, StringSplitOptions.RemoveEmptyEntries)
            frm1.Mask(t(0), t(1)) = 0
        Next

    End Sub

    Public Sub NumericUpDown1_ValueChanged(sender As Object, e As EventArgs) Handles NumericUpDown1.ValueChanged
        Try
            NumericUpDown1.Maximum = Math.Min(2147483647, frm1.mw * frm1.mh)
        Catch
        End Try
    End Sub

    Public Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim s As String = ""
        Dim r As New Random
        For i As Long = 1 To NumericUpDown1.Value
            s &= r.Next(frm1.mw) & "," & r.Next(frm1.mh) & vbCrLf
        Next
        TextBox1.Text = s
    End Sub

    Public Sub Form5_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        e.Cancel = True
        Me.Hide()
    End Sub
End Class