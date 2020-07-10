Public Class Form10
    Public frm1 As Form1
    Public img As Bitmap
    Public data(,) As Double, DMin As Double, DMax As Double
    Public mw, mh As Long
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If OpenFileDialog1.ShowDialog = DialogResult.OK Then
            Dim s() As String = My.Computer.FileSystem.ReadAllText(OpenFileDialog1.FileName).Split({vbCrLf}, StringSplitOptions.RemoveEmptyEntries)
            mh = s.Length
            mw = s(0).Split({","}, StringSplitOptions.RemoveEmptyEntries).Length
            ReDim data(mw - 1, mh - 1)
            Dim t() As String

            For i As Long = 0 To mh - 1
                t = s(i).Split({","}, StringSplitOptions.RemoveEmptyEntries)
                For j As Long = 0 To mw - 1
                    data(j, i) = Val(t(j))
                Next
            Next
            DMin = data(0, 0)
            DMax = DMin
            For i As Long = 0 To mh - 1
                For j As Long = 0 To mw - 1
                    If DMin > data(j, i) Then
                        DMin = data(j, i)
                    End If
                    If DMax < data(j, i) Then
                        DMax = data(j, i)
                    End If
                Next
            Next

            img = New Bitmap(mw, mh, Imaging.PixelFormat.Format32bppArgb)

            For i As Long = 0 To mh - 1
                For j As Long = 0 To mw - 1

                    img.SetPixel(j, i, Color.FromArgb(frm1.RGB2ARGB(frm1.ColTrs((data(j, i) - DMin) * 20000 / (DMax - DMin) - 12500))))
                Next
            Next
            PictureBox1.Image = img

        End If
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        If SaveFileDialog1.ShowDialog = DialogResult.OK Then
            img.Save(SaveFileDialog1.FileName)
        End If
    End Sub

    Private Sub Form10_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub Form10_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        e.Cancel = True
        Hide()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Clipboard.SetImage(img)
        'BackColor = Color.FromArgb(frm1.RGB2ARGB(frm1.ColTrs(Val(InputBox("", "", "0")))))
    End Sub
End Class