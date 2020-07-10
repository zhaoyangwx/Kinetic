Public Class Form2
    Public AutoLoad As Boolean = False
    Public frm1 As Form1
    Public Function DblFromCsv(ByVal FileName As String) As Double(,)
        Try
            Dim Count As Long = 0
            Dim s As String = My.Computer.FileSystem.ReadAllText(FileName)
            Dim t() As String = s.Split({vbCrLf}, StringSplitOptions.RemoveEmptyEntries)
            Dim p() As String
            Count = t.Length
            Dim OptDbl(Count - 1, 1) As Double
            For i As Long = 0 To Count - 1
                p = t(i).Split({","}, StringSplitOptions.RemoveEmptyEntries)
                OptDbl(i, 0) = Val(p(0))
                OptDbl(i, 1) = Val(p(1))
            Next
            'Plot(OptDbl, 500, 4000,,, True,, )
            Return OptDbl
        Catch ex As Exception
            Return Nothing
        End Try

    End Function

    Public Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        frm1.PI_0 = 32.62241
        frm1.M_0 = 5835.969
        frm1.D_O = 0.000000000108
        frm1.O_2_0 = 0.013
        frm1.O2ConcBoundry = frm1.O_2_0
        frm1.k_i = 0.000048
        frm1.k_p0 = 114.5
        frm1.k_t0 = 1337.0
        frm1.A_p = 0.86
        frm1.A_t = 22.8
        frm1.f_cp = 0.075
        frm1.f_ct = 0.0581
        frm1.R = 11
        frm1.k_O = 500000.0
        frm1.Phi = 0.6
        frm1.Epsilon = 188.3069272
        frm1.I_0 = 20
        frm1.A0l = 1078.809661
        frm1.A0s = 1078.809661
        frm1.Temperature = 298.15
        frm1.Alpha_m = 0.000886
        frm1.Alpha_p = 0.00013
        frm1.T_gm = 225.67
        frm1.T_gp = 358.41
        frm1.Rho_m = 1.173650031
        frm1.Rho_p = 1.215612336
        frm1.StepLength = 0.00001
        frm1.ScaleMeterPerPixel = 0.00001
        Button3_Click(sender, e)
        Button2_Click(sender, e)
    End Sub

    Public Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        frm1.Rho_m = Val(TextBox1.Text)
        frm1.Rho_p = Val(TextBox2.Text)
        frm1.Alpha_m = Val(TextBox3.Text)
        frm1.Alpha_p = Val(TextBox4.Text)
        frm1.T_gm = Val(TextBox5.Text)
        frm1.T_gp = Val(TextBox6.Text)
        frm1.A_p = Val(TextBox7.Text)
        frm1.A_t = Val(TextBox8.Text)
        frm1.k_p0 = Val(TextBox9.Text)
        frm1.k_t0 = Val(TextBox10.Text)
        frm1.f_cp = Val(TextBox11.Text)
        frm1.f_ct = Val(TextBox12.Text)
        frm1.R = Val(TextBox13.Text)
        frm1.k_O = Val(TextBox14.Text)
        frm1.Phi = Val(TextBox15.Text)
        frm1.Epsilon = Val(TextBox16.Text)
        frm1.A0l = Val(TextBox17.Text)
        frm1.A0s = Val(TextBox18.Text)
        frm1.I_0 = Val(TextBox19.Text)
        frm1.Temperature = Val(TextBox20.Text)
        frm1.PI_0 = Val(TextBox21.Text)
        frm1.M_0 = Val(TextBox22.Text)
        frm1.D_O = Val(TextBox23.Text)
        frm1.O_2_0 = Val(TextBox24.Text)
        frm1.k_i = Val(TextBox25.Text)
        frm1.StepLength = Val(TextBox26.Text)
        frm1.ScaleMeterPerPixel = Val(TextBox27.Text)
        If Not AutoLoad Then MessageBox.Show("设置成功！")
        Button3_Click(sender, e)
    End Sub

    Public Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        TextBox1.Text = frm1.Rho_m
        TextBox2.Text = frm1.Rho_p
        TextBox3.Text = frm1.Alpha_m
        TextBox4.Text = frm1.Alpha_p
        TextBox5.Text = frm1.T_gm
        TextBox6.Text = frm1.T_gp
        TextBox7.Text = frm1.A_p
        TextBox8.Text = frm1.A_t
        TextBox9.Text = frm1.k_p0
        TextBox10.Text = frm1.k_t0
        TextBox11.Text = frm1.f_cp
        TextBox12.Text = frm1.f_ct
        TextBox13.Text = frm1.R
        TextBox14.Text = frm1.k_O
        TextBox15.Text = frm1.Phi
        TextBox16.Text = frm1.Epsilon
        TextBox17.Text = frm1.A0l
        TextBox18.Text = frm1.A0s
        TextBox19.Text = frm1.I_0
        TextBox20.Text = frm1.Temperature
        TextBox21.Text = frm1.PI_0
        TextBox22.Text = frm1.M_0
        TextBox23.Text = frm1.D_O
        TextBox24.Text = frm1.O_2_0
        TextBox25.Text = frm1.k_i
        TextBox26.Text = frm1.StepLength
        TextBox27.Text = frm1.ScaleMeterPerPixel
    End Sub

    Public Sub Form2_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Button3_Click(sender, e)
        OpenFileDialog1.FileName = My.Computer.FileSystem.CurrentDirectory & "\Setting.csv"
    End Sub

    Public Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        frm1.frm5.Show()
    End Sub

    Public Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        If Not AutoLoad Then AutoLoad = OpenFileDialog1.ShowDialog = DialogResult.OK
        If AutoLoad Then
            AutoLoad = False
            Dim d(,) As Double = DblFromCsv(OpenFileDialog1.FileName)
            TextBox1.Text = d(0, 1)
            TextBox2.Text = d(1, 1)
            TextBox3.Text = d(2, 1)
            TextBox4.Text = d(3, 1)
            TextBox5.Text = d(4, 1)
            TextBox6.Text = d(5, 1)
            TextBox7.Text = d(6, 1)
            TextBox8.Text = d(7, 1)
            TextBox9.Text = d(8, 1)
            TextBox10.Text = d(9, 1)
            TextBox11.Text = d(10, 1)
            TextBox12.Text = d(11, 1)
            TextBox13.Text = d(12, 1)
            TextBox14.Text = d(13, 1)
            TextBox15.Text = d(14, 1)
            TextBox16.Text = d(15, 1)
            TextBox17.Text = d(16, 1)
            TextBox18.Text = d(17, 1)
            TextBox19.Text = d(18, 1)
            TextBox20.Text = d(19, 1)
            TextBox21.Text = d(20, 1)
            TextBox22.Text = d(21, 1)
            TextBox23.Text = d(22, 1)
            TextBox24.Text = d(23, 1)
            TextBox25.Text = d(24, 1)
            TextBox26.Text = d(25, 1)
            TextBox27.Text = d(26, 1)
        End If

    End Sub

    Public Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        TextBox21.Text = "32.62241202"
    End Sub

    Public Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        TextBox21.Text = "66.36848055"
    End Sub

    Public Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        TextBox21.Text = "98.13856064"
    End Sub

    Public Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        TextBox19.Text = "5"
    End Sub

    Public Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click
        TextBox19.Text = "20"
    End Sub

    Public Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click
        TextBox19.Text = "40"
    End Sub

    Public Sub Form2_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        Me.Hide()
        e.Cancel = True
    End Sub

    Public Sub Button12_Click(sender As Object, e As EventArgs) Handles Button12.Click
        frm1.frm8.Show()
    End Sub

    Public Sub Button13_Click(sender As Object, e As EventArgs) Handles Button13.Click
        If SaveFileDialog1.ShowDialog = DialogResult.OK Then
            Dim d(26, 1) As Double
            d(0, 1) = Val(TextBox1.Text)
            d(1, 1) = Val(TextBox2.Text)
            d(2, 1) = Val(TextBox3.Text)
            d(3, 1) = Val(TextBox4.Text)
            d(4, 1) = Val(TextBox5.Text)
            d(5, 1) = Val(TextBox6.Text)
            d(6, 1) = Val(TextBox7.Text)
            d(7, 1) = Val(TextBox8.Text)
            d(8, 1) = Val(TextBox9.Text)
            d(9, 1) = Val(TextBox10.Text)
            d(10, 1) = Val(TextBox11.Text)
            d(11, 1) = Val(TextBox12.Text)
            d(12, 1) = Val(TextBox13.Text)
            d(13, 1) = Val(TextBox14.Text)
            d(14, 1) = Val(TextBox15.Text)
            d(15, 1) = Val(TextBox16.Text)
            d(16, 1) = Val(TextBox17.Text)
            d(17, 1) = Val(TextBox18.Text)
            d(18, 1) = Val(TextBox19.Text)
            d(19, 1) = Val(TextBox20.Text)
            d(20, 1) = Val(TextBox21.Text)
            d(21, 1) = Val(TextBox22.Text)
            d(22, 1) = Val(TextBox23.Text)
            d(23, 1) = Val(TextBox24.Text)
            d(24, 1) = Val(TextBox25.Text)
            d(25, 1) = Val(TextBox26.Text)
            d(26, 1) = Val(TextBox27.Text)
            Dim s As String = ""
            For i As Long = 0 To 26
                s &= d(i, 0) & "," & d(i， 1) & vbCrLf
            Next
            My.Computer.FileSystem.WriteAllText(SaveFileDialog1.FileName, s, False)
        End If

    End Sub
End Class