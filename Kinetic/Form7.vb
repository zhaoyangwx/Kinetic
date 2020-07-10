Public Class Form7
    Public frm1 As Form1

    Public mw As Long = 4, mh As Long = 200

    Public PI(mw - 1, mh - 1) As Double
    Public R_I(mw - 1, mh - 1) As Double
    Public IRadical(mw - 1, mh - 1) As Double
    Public Intensity(mw - 1, mh - 1) As Double
    Public k_p(mw - 1, mh - 1) As Double
    Public Monomer(mw - 1, mh - 1) As Double
    Public M_nRadical(mw - 1, mh - 1) As Double
    Public k_t(mw - 1, mh - 1) As Double
    Public x(mw - 1, mh - 1) As Double
    Public O2Conc(mw - 1, mh - 1) As Double
    Public O2ConcSDz(mw - 1, mh - 1) As Double
    Public O2ConcDDz(mw - 1, mh - 1) As Double
    Public Mask(mw - 1, mh - 1) As Double
    'Public f As Double
    Public t As Double
    Public SourceY As Long = 0
    Public LayerCount As Long = 10
    Public TimePerLayer As Double = 0.1
    Public CurrentLayer As Long = 1
    Public LayerHeight As Long = (mh + 1) / LayerCount
    Public CalcO2Diff As Boolean = False
    Public hasmask As Boolean = False

    Public MMax, MMin, IMax, IMin As Double
    Public bt() As Byte
    Public Working As Boolean = False
    Public fps As Long = 0
    Public Stride As Long
    Public RenderIntv As Long = 100
    Public RenderFramCount As Long = RenderIntv
    Public SaveFrm As Form3
    Function StepLength() As Double
        StepLength = frm1.StepLength
    End Function
    Function ScaleMeterPerPixel() As Double
        ScaleMeterPerPixel = frm1.ScaleMeterPerPixel
    End Function
    Sub ResetMap()
        ReDim PI(mw - 1, mh - 1)
        ReDim IRadical(mw - 1, mh - 1)
        ReDim R_I(mw - 1, mh - 1)
        ReDim Monomer(mw - 1, mh - 1)
        ReDim Intensity(mw - 1, mh - 1)
        ReDim k_p(mw - 1, mh - 1)
        ReDim M_nRadical(mw - 1, mh - 1)
        ReDim k_t(mw - 1, mh - 1)
        ReDim x(mw - 1, mh - 1)
        ReDim O2Conc(mw - 1, mh - 1)
        ReDim O2ConcSDz(mw - 1, mh - 1)
        ReDim O2ConcDDz(mw - 1, mh - 1)
        If hasmask = False Then ReDim Mask(mw - 1, mh - 1)

        For i As Long = 0 To mw - 1
            For j As Long = 0 To mh - 1
                PI(i, j) = frm1.PI_0
                IRadical(i, j) = 0
                Monomer(i, j) = frm1.M_0
                Intensity(i, j) = 0
                k_p(i, j) = frm1.k_p0
                M_nRadical(i, j) = 0
                k_t(i, j) = frm1.k_t0
                x(i, j) = 0
                O2Conc(i, j) = frm1.O_2_0
                R_I(i, j) = 0
                O2ConcSDz(i, j) = 0
                O2ConcDDz(i, j) = 0
                If hasmask = False Then Mask(i, j) = 1
            Next
        Next
        NumericUpDown4.Maximum = mh - 1
        Button7_Click(New Object, New EventArgs)
        ChangeSourcePosition(mh - LayerHeight)
        t = 0
        CurrentLayer = 1
        Button7_Click(New Object, New EventArgs)
    End Sub
    Public Sub ChangeSourcePosition(ByVal p As Long)
        SourceY = p
        If SourceY > 0 Then
            For i As Long = 0 To mw - 1
                For j As Long = 0 To SourceY - 1
                    Intensity(i, j) = 0
                Next
            Next
        End If
        Parallel.For(0, mw, AddressOf RefreshLiquid)
    End Sub
    Sub RefreshMap()
        Dim img1 As New Bitmap(mw, mh, Imaging.PixelFormat.Format24bppRgb)
        Dim img2 As New Bitmap(mw, mh, Imaging.PixelFormat.Format24bppRgb)
        Render(img1, img2)
        PictureBox1.Image = img1
        PictureBox2.Image = img2
    End Sub
    '渲染画面
    Sub Render(ByRef img1 As Bitmap, ByRef img2 As Bitmap)

        MMax = Monomer(0, 0)
        MMin = Monomer(0, 0)
        IMax = Intensity(0, 0)
        IMin = Intensity(0, 0)
        For i As Long = 0 To mw - 1
            For j As Long = 0 To mh - 1
                If MMax < Monomer(i, j) Then MMax = Monomer(i, j)
                If MMin > Monomer(i, j) Then MMin = Monomer(i, j)
                If IMax < Intensity(i, j) Then IMax = Intensity(i, j)
                If IMin > Intensity(i, j) Then IMin = Intensity(i, j)
            Next
        Next
        If MMin = MMax Then
            MMin -= 1
            MMax += 1
        End If
        If IMin = IMax Then
            IMin -= 1
            IMax += 1
        End If
        Label7.Text = Format(MMax, "0.0000E+000")
        Label8.Text = Format(MMin, "0.0000E+000")
        Label9.Text = Format(IMax, "0.0000E+000")
        Label10.Text = Format(IMin, "0.0000E+000")
        Stride = mw * 3
        If Stride Mod 4 <> 0 Then Stride = Stride + 4 - Stride Mod 4
        Dim numBytes As Long = mh * Stride
        Dim c As Imaging.BitmapData = img1.LockBits(New Rectangle(0, 0, mw, mh), Imaging.ImageLockMode.ReadWrite, Imaging.PixelFormat.Format24bppRgb)
        ReDim bt(numBytes - 1)
        Runtime.InteropServices.Marshal.Copy(c.Scan0, bt, 0, numBytes)
        Parallel.For(0, mh, AddressOf RenderConc)
        Runtime.InteropServices.Marshal.Copy(bt, 0, c.Scan0, numBytes)
        img1.UnlockBits(c)

        c = img2.LockBits(New Rectangle(0, 0, mw, mh), Imaging.ImageLockMode.ReadWrite, Imaging.PixelFormat.Format24bppRgb)
        ReDim bt(numBytes - 1)
        Runtime.InteropServices.Marshal.Copy(c.Scan0, bt, 0, numBytes)
        Parallel.For(0, mh, AddressOf RenderInts)
        Runtime.InteropServices.Marshal.Copy(bt, 0, c.Scan0, numBytes)
        img2.UnlockBits(c)
    End Sub
    Sub RenderConc(ByVal j As Long)
        Dim Offset As Long, Col As Long
        For i As Long = 0 To mw - 1
            Try
                Offset = j * Stride + i * 3
                Col = frm1.ColTrs(CInt((Monomer(i, j) - MMin) * 25000 / (MMax - MMin) - 12500))
                bt(Offset + 2) = (Col And 255)
                Col >>= 8
                bt(Offset + 1) = (Col And 255)
                Col >>= 8
                bt(Offset) = Col And 255
            Catch ex As Exception
                MessageBox.Show(ex.ToString)
            End Try
        Next
    End Sub
    Sub RenderInts(ByVal j As Long)
        Dim Offset As Long, Col As Long
        For i As Long = 0 To mw - 1
            Try
                Offset = j * Stride + i * 3
                Col = frm1.ColTrs(CInt((Intensity(i, j) - IMin) * 25000 / (IMax - IMin) - 12500))
                bt(Offset + 2) = (Col And 255)
                Col >>= 8
                bt(Offset + 1) = (Col And 255)
                Col >>= 8
                bt(Offset) = Col
            Catch ex As Exception

            End Try
        Next
    End Sub

    '计算下一步（并行）
    Sub CalcNextFrame()
        For i As Long = 0 To mw - 1
            Intensity(i, SourceY) = frm1.I_0 * Mask(i, SourceY)
        Next
        Parallel.For(0, mw, AddressOf RefreshIntensity)
        Parallel.For(0, mw, AddressOf RefreshR_I)
        Parallel.For(0, mw, AddressOf RefreshPI)
        Parallel.For(0, mw, AddressOf RefreshIRadical)
        Parallel.For(0, mw, AddressOf RefreshMonomer)
        Parallel.For(0, mw, AddressOf RefreshM_nRadical)
        'For j As Long = 0 To mw - 1
        '    O2Conc(j, mh - 1) = O2ConcBoundry
        'Next
        If CalcO2Diff Then
            CalcDblDifferentialO2Conc()
        End If
        Parallel.For(0, mw, AddressOf RefreshO2Conc)
        Parallel.For(0, mw, AddressOf Refreshx)
        Parallel.For(0, mw, AddressOf Refreshk_pk_t)
    End Sub
    Sub RefreshLiquid(i As Long)
        For j As Long = SourceY To mh - 1
            If x(i, j) < Form1.NJD Then
                PI(i, j) = frm1.PI_0
                IRadical(i, j) = 0
                Monomer(i, j) = frm1.M_0
                Intensity(i, j) = 0
                k_p(i, j) = frm1.k_p0
                M_nRadical(i, j) = 0
                k_t(i, j) = frm1.k_t0
                x(i, j) = 0
                O2Conc(i, j) = frm1.O_2_0
                R_I(i, j) = 0
                O2ConcSDz(i, j) = 0
                O2ConcDDz(i, j) = 0
            End If
        Next
    End Sub
    Sub RefreshIntensity(i As Long)
        For j As Long = SourceY + 1 To mh - 1
            Try
                If x(i, j) < frm1.NJD Then
                    Intensity(i, j) = Intensity(i, j - 1) * Math.Exp((-frm1.Epsilon * PI(i, j)) * frm1.ScaleMeterPerPixel - frm1.A0l * frm1.ScaleMeterPerPixel)
                Else
                    Intensity(i, j) = Intensity(i, j - 1) * Math.Exp((-frm1.Epsilon * PI(i, j)) * frm1.ScaleMeterPerPixel - frm1.A0s * frm1.ScaleMeterPerPixel)
                End If
                'Intensity(i, j) = Intensity(i, j - 1) * Math.Exp((-Epsilon * PI(i, j) - BkgAbsorb) * ScaleMeterPerPixel) * Mask(i, j) 
            Catch ex As Exception
                Intensity(i, j) = 0
            End Try
        Next
    End Sub
    Sub RefreshR_I(i As Long)
        For j As Long = SourceY To mh - 1
            R_I(i, j) = frm1.Phi * frm1.Epsilon * PI(i, j) * Intensity(i, j)
        Next
    End Sub
    Sub RefreshPI(i As Long)
        For j As Long = SourceY To mh - 1
            PI(i, j) += -0.5 * R_I(i, j) * frm1.StepLength
        Next
    End Sub
    Sub RefreshIRadical(i As Long)
        For j As Long = SourceY To mh - 1
            IRadical(i, j) += (R_I(i, j) - frm1.k_i * IRadical(i, j) * Monomer(i, j)) * frm1.StepLength
        Next
    End Sub
    Sub RefreshMonomer(i As Long)
        For j As Long = SourceY To mh - 1
            Monomer(i, j) += (-frm1.k_i * IRadical(i, j) * Monomer(i, j) - k_p(i, j) * Monomer(i, j) * M_nRadical(i, j)) * frm1.StepLength
        Next
    End Sub
    Sub RefreshM_nRadical(i As Long)
        For j As Long = SourceY To mh - 1
            M_nRadical(i, j) += (frm1.k_i * IRadical(i, j) * Monomer(i, j) - k_t(i, j) * M_nRadical(i, j) * M_nRadical(i, j) - frm1.k_O * M_nRadical(i, j) * O2Conc(i, j)) * frm1.StepLength
        Next
    End Sub
    Sub CalcDblDifferentialO2Conc()
        'Parallel.For(0, mw, AddressOf CalcDDOLoop1)
        Parallel.For(0, mw, AddressOf CalcDDOLoop2)
    End Sub
    Sub CalcDDOLoop1(i As Long)
        For j As Long = 1 To mh - 2
            O2ConcSDz(i, j) = (O2Conc(i, j + 1) - O2Conc(i, j - 1)) / 2 / frm1.ScaleMeterPerPixel
        Next
    End Sub
    Sub CalcDDOLoop2(i As Long)
        O2ConcDDz(i, 1) = (O2Conc(i, 3) - O2Conc(i, 1)) / 4 / frm1.ScaleMeterPerPixel / frm1.ScaleMeterPerPixel
        O2ConcDDz(i, mh - 2) = (O2Conc(i, mh - 4) - O2Conc(i, mh - 2)) / 4 / frm1.ScaleMeterPerPixel / frm1.ScaleMeterPerPixel
        For j As Long = 2 To mh - 3
            O2ConcDDz(i, j) = (O2Conc(i, j + 2) + O2Conc(i, j - 2) - 2 * O2Conc(i, j)) / 4 / frm1.ScaleMeterPerPixel / frm1.ScaleMeterPerPixel
        Next
    End Sub
    Sub RefreshO2Conc(i As Long)
        For j As Long = 0 To mh - 1
            O2Conc(i, j) += (frm1.D_O * O2ConcDDz(i, j) - frm1.k_O * O2Conc(i, j) * M_nRadical(i, j)) * frm1.StepLength
        Next
    End Sub
    Sub Refreshx(i As Long)
        For j As Long = SourceY To mh - 1
            x(i, j) = 1 - Monomer(i, j) / frm1.M_0
        Next
    End Sub

    Public Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If Timer1.Enabled Then Timer1.Enabled = False
        CheckBox1.Checked = False
        While Working
            Timer1.Enabled = False
        End While
        'a = Val(TextBox1.Text)
        'b = Val(TextBox2.Text)
        mw = Val(NumericUpDown1.Value)
        mh = Val(NumericUpDown2.Value)
        ResetMap()
        RefreshMap()
    End Sub

    Public Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        RenderFramCount = RenderIntv
        Timer1_Tick(sender, e)
    End Sub

    Public Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        Timer1.Enabled = CheckBox1.Checked

    End Sub

    Public Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Timer1.Enabled = False
        CheckBox1.Checked = False
        While Working
        End While
        mw = Val(NumericUpDown1.Value)
        mh = Val(NumericUpDown2.Value)
        SaveFrm.NumericUpDown1.Maximum = mw - 1
        SaveFrm.NumericUpDown2.Maximum = mh - 1
        ResetMap()
        Button8.Enabled = True
    End Sub

    Public Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        ChangeSourcePosition(NumericUpDown4.Value)
    End Sub

    Public Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        frm1.frm2.Show()
    End Sub

    Public Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        If OpenFileDialog1.ShowDialog = DialogResult.OK Then
            hasmask = True
            Dim img As Bitmap = Image.FromFile(OpenFileDialog1.FileName)
            Dim c As Color
            For i As Long = 0 To mw - 1
                For j As Long = 0 To mh - 1
                    c = img.GetPixel(i, j)
                    If (c.R Or c.G Or c.B) = 0 Then
                        Mask(i, j) = 0
                    Else
                        Mask(i, j) = 1
                    End If
                Next
            Next
            Button1.Enabled = True
        End If
    End Sub

    Public Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        SaveFrm.Show()
    End Sub

    Public Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        LayerCount = NumericUpDown5.Value
        TimePerLayer = TextBox1.Text
        LayerHeight = (mh) / LayerCount
        ChangeSourcePosition(mh - LayerHeight)
        CurrentLayer = 1
        Label4.Text = "层数:" & CurrentLayer & "/" & LayerCount
        Label11.Text = "每层时间:" & TimePerLayer
    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        frm1.frm9.Show()
        Me.Hide()
    End Sub


    Public Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        Working = True
        For m As Long = 0 To NumericUpDown3.Value
            t += frm1.StepLength
            fps += 1
            If t > TimePerLayer * CurrentLayer Then
                CurrentLayer += 1
                If CurrentLayer * LayerHeight > mh Then
                    Timer1.Enabled = False
                    SaveFrm.SaveData()
                    CheckBox1.Checked = False
                    RefreshMap()
                    Label1.Text = "Time=" & t
                    Working = False
                    Exit Sub
                End If
                ChangeSourcePosition(mh - (CurrentLayer * LayerHeight))
                Label4.Text = "层数:" & CurrentLayer & "/" & LayerCount
            End If
            CalcNextFrame()
            RenderFramCount += 1
            If RenderFramCount >= RenderIntv Then
                RenderFramCount = 0
            Else
                Continue For
            End If
            RefreshMap()
            Label1.Text = "Time=" & Format(t, "0.000000E+000")
            SaveFrm.SaveData()
        Next
        Working = False
    End Sub

    Public Sub Timer2_Tick(sender As Object, e As EventArgs) Handles Timer2.Tick
        Label12.Top = Math.Min(PictureBox3.Top + PictureBox3.Height, Math.Max(PictureBox3.Top, PictureBox3.Top + (MMax - frm1.M_0 * (1 - frm1.NJD)) * PictureBox3.Height / (MMax - MMin)))
        Label2.Text = "fps=" & fps * 1000 / Timer2.Interval
        fps = 0
    End Sub

    Sub Refreshk_pk_t(i As Long)
        Dim phi_m As Double, f As Double
        For j As Long = 0 To mh - 1
            phi_m = (1 - x(i, j)) / (1 - x(i, j) + frm1.Rho_m / frm1.Rho_p * x(i, j))
            f = frm1.f_m * phi_m + frm1.f_p * (1 - phi_m)
            k_p(i, j) = frm1.k_p0 / (1 + Math.Exp(-frm1.A_p * (1 / frm1.f_cp - 1 / f)))
            k_t(i, j) = frm1.k_t0 / (1 + 1 / (frm1.R * k_p(i, j) * Monomer(i, j) / frm1.k_t0 + Math.Exp(-frm1.A_t * (1 / f - 1 / frm1.f_ct))))
        Next
    End Sub

    Public Sub Form7_Load(sender As Object, e As EventArgs) Handles Me.Load
        SaveFrm = New Form3
        SaveFrm.frm7 = Me
        SaveFrm.PForm = 7
        SaveFrm.Text = "模拟打印_保存设置"
    End Sub
    Private Sub PictureBox1_DoubleClick(sender As Object, e As EventArgs) Handles PictureBox1.DoubleClick
        If PictureBox1.Image Is Nothing Then Exit Sub
        If MessageBox.Show("是否复制到剪贴板？", "提示", MessageBoxButtons.YesNo) = DialogResult.Yes Then Clipboard.SetImage(PictureBox1.Image)
    End Sub

    Private Sub PictureBox2_DoubleClick(sender As Object, e As EventArgs) Handles PictureBox2.DoubleClick
        If PictureBox2.Image Is Nothing Then Exit Sub
        If MessageBox.Show("是否复制到剪贴板？", "提示", MessageBoxButtons.YesNo) = DialogResult.Yes Then Clipboard.SetImage(PictureBox2.Image)
    End Sub
End Class