
Public Class Form4
    Public x, y As Long
    Public frm1 As Form1
    Public Sub RefreshEvent()
        If CheckBox1.Checked Then
            Button3_Click(New Object, New EventArgs)
        End If
        If CheckBox2.Checked Then
            My.Computer.FileSystem.WriteAllText("D:\Desktop\文件\kn\14\x0.csv", frm1.t & "," & frm1.x(0, 0) & vbCrLf, True)
            My.Computer.FileSystem.WriteAllText("D:\Desktop\文件\kn\14\xh.csv", frm1.t & "," & frm1.x(0, frm1.mh - 1) & vbCrLf, True)
            Dim s As Double = 0
            For i As Long = 0 To frm1.mh - 1
                s += frm1.x(0, i)
            Next
            s /= frm1.mh
            My.Computer.FileSystem.WriteAllText("D:\Desktop\文件\kn\14\xA.csv", frm1.t & "," & s & vbCrLf, True)
        End If
    End Sub

    Public Sub NumericUpDown1_ValueChanged(sender As Object, e As EventArgs) Handles NumericUpDown1.ValueChanged
        NumericUpDown1.Maximum = frm1.mw - 1
    End Sub

    Public Sub NumericUpDown2_ValueChanged(sender As Object, e As EventArgs) Handles NumericUpDown2.ValueChanged
        NumericUpDown2.Maximum = frm1.mh - 1
    End Sub

    Public Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        CheckedListBox1.Items(0) = "I		光强   			" & frm1.Intensity(x, y)
        CheckedListBox1.Items(1) = "R_i		引发速率   		" & frm1.R_I(x, y)
        CheckedListBox1.Items(2) = "[PI]		光引发剂浓度   		" & frm1.PI(x, y)
        CheckedListBox1.Items(3) = "[I∙]		初级自由基浓度  	" & frm1.IRadical(x, y)
        CheckedListBox1.Items(4) = "[M]		单体浓度   		" & frm1.Monomer(x, y)
        CheckedListBox1.Items(5) = "[M_n∙]		次级自由基浓度  	" & frm1.M_nRadical(x, y)
        CheckedListBox1.Items(6) = "[O_2]		氧气浓度   		" & frm1.O2Conc(x, y)
        CheckedListBox1.Items(7) = "k_p		链增长速率常数  	" & frm1.k_p(x, y)
        CheckedListBox1.Items(8) = "k_t		链终止速率常数  	" & frm1.k_t(x, y)
        CheckedListBox1.Items(9) = "x		转化率   		" & frm1.x(x, y)

    End Sub

    Public Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim MatLabFigEngine As New Figure
        Dim data(,) As Double = frm1.x

        If CheckedListBox1.GetItemChecked(0) Then
            data = frm1.Intensity
        End If
        If CheckedListBox1.GetItemChecked(1) Then
            data = frm1.R_I
        End If
        If CheckedListBox1.GetItemChecked(2) Then
            data = frm1.PI
        End If
        If CheckedListBox1.GetItemChecked(3) Then
            data = frm1.IRadical
        End If
        If CheckedListBox1.GetItemChecked(4) Then
            data = frm1.Monomer
        End If
        If CheckedListBox1.GetItemChecked(5) Then
            data = frm1.M_nRadical
        End If
        If CheckedListBox1.GetItemChecked(6) Then
            data = frm1.O2Conc
        End If
        If CheckedListBox1.GetItemChecked(7) Then
            data = frm1.k_p
        End If
        If CheckedListBox1.GetItemChecked(8) Then
            data = frm1.k_t
        End If
        If CheckedListBox1.GetItemChecked(9) Then
            data = frm1.x
        End If


        Dim s(frm1.mh - 1, 1) As String
        For ly As Long = 0 To frm1.mh - 1
            s(ly, 0) = ly * frm1.ScaleMeterPerPixel
            s(ly, 1) = data(x, ly)
        Next
        MatLabFigEngine.Plot(s)
    End Sub

    Public Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim MatLabFigEngine As New Figure
        Dim data(,) As Double = frm1.x

        If CheckedListBox1.GetItemChecked(0) Then
            data = frm1.Intensity
        End If
        If CheckedListBox1.GetItemChecked(1) Then
            data = frm1.R_I
        End If
        If CheckedListBox1.GetItemChecked(2) Then
            data = frm1.PI
        End If
        If CheckedListBox1.GetItemChecked(3) Then
            data = frm1.IRadical
        End If
        If CheckedListBox1.GetItemChecked(4) Then
            data = frm1.Monomer
        End If
        If CheckedListBox1.GetItemChecked(5) Then
            data = frm1.M_nRadical
        End If
        If CheckedListBox1.GetItemChecked(6) Then
            data = frm1.O2Conc
        End If
        If CheckedListBox1.GetItemChecked(7) Then
            data = frm1.k_p
        End If
        If CheckedListBox1.GetItemChecked(8) Then
            data = frm1.k_t
        End If
        If CheckedListBox1.GetItemChecked(9) Then
            data = frm1.x
        End If


        Dim s(frm1.mh - 1, 1) As String
        For ly As Long = 0 To frm1.mh - 1
            s(ly, 0) = ly * frm1.ScaleMeterPerPixel
            s(ly, 1) = data(x, ly)
        Next
        MatLabFigEngine.Plot(s,,,,,,, True)
    End Sub

    Public Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        x = NumericUpDown1.Value
        y = NumericUpDown2.Value
    End Sub

    Public Sub Form4_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        e.Cancel = True
        Me.Hide()
    End Sub
End Class