Option Explicit On
'Option Strict On

Public Class Form1

    Public mw As Long = 4, mh As Long = 200
    Public StepLength As Double = 0.00001
    Public ScaleMeterPerPixel As Double = 0.00001

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

    'Const
    '[PI]_0		54.557		initial photoinitiator concentration		已知
    '[M]_0		4.84E3		initial monomer concentration			已知
    'D_O		1.08e-10	diffusion coefficient of oxygen			凑
    '[O_2]_0	1.3E-2		initial oxygen concentration			凑
    'k_i		4.8e-5		primary radical kinetic constant		凑
    'k_p0		1.145e5		ideal kp					拟合		
    'k_t0		1.337e6		ideal kt					拟合		
    'A_p		1.26		Acceleration control constant			拟合		
    'A_t		22.8		Acceleration control constant			拟合		
    'f_cp		5.17e-2		critical free volume fraction propagation	拟合		
    'f_ct		5.81e-2		critical free volume fraction termination	拟合		
    'R		    11		reaction-diffusion constant			？
    'k_O		5e5		oxygen inhibition reaction constant		文献值
    'φ		    0.6		quantum yield					？
    'ε		    9.66e-1		molar absorption coefficient			？
    'I_0		10|0		UV Intensity(mW∙cm^-2)					测
    'L		    10e-6								设定
    'T		    303.15								测
    'α_m		5e-4		thermal expansion coefficient of monomer	测		PVT/DSA
    'α_p		7.5e-5		thermal expansion coefficient of polymer	测		PVT/DSA
    'T_gm		236.51		glass transition temperature of monomer		测		DSC
    'T_gp		313.12		glass transition temperature of polymer		测		DSC
    'ρ_m		1.103459	density of monomer				测
    'ρ_p		1.320		density of polymer				测
    Public PI_0 As Double = 32.62241
    Public M_0 As Double = 5835.969
    Public D_O As Double = 0.000000000108
    Public O_2_0 As Double = 0.013
    Public O2ConcBoundry As Double = O_2_0
    Public k_i As Double = 0.000048
    Public k_p0 As Double = 114.5 / 60
    Public k_t0 As Double = 1337.0 / 60
    Public A_p As Double = 0.86
    Public A_t As Double = 22.8
    Public f_cp As Double = 0.075
    Public f_ct As Double = 0.0581
    Public R As Double = 11
    Public k_O As Double = 500000.0
    Public Phi As Double = 0.6
    Public Epsilon As Double = 188.3069272
    Public A0l As Double = 1078.809661
    Public A0s As Double = 1078.809661
    Public I_0 As Double = 20
    Public Temperature As Double = 298.15
    Public Alpha_m As Double = 0.000886
    Public Alpha_p As Double = 0.00013
    Public T_gm As Double = 225.67
    Public T_gp As Double = 358.41
    Public Rho_m As Double = 1.173650031
    Public Rho_p As Double = 1.215612336
    Public NJD As Double = 0.17
    Public f_m As Double = 0.025 + Alpha_m * (Temperature - T_gm)
    Public f_p As Double = 0.025 + Alpha_p * (Temperature - T_gp)

    Public MMax, MMin, IMax, IMin As Double
    Public bt() As Byte
    Public Working As Boolean = False
    Public fps As Long = 0
    Public Stride As Long
    Public RenderIntv As Long = 100
    Public RenderFramCount As Long = RenderIntv
    Public StartTask As Boolean = False

    Public DebugO As String = ""
    Public frm2 As Form2
    Public frm3 As Form3
    Public frm4 As Form4
    Public frm5 As Form5
    Public frm6 As Form6
    Public frm7 As Form7
    Public frm8 As Form8
    Public frm9 As Form9
    Public frm10 As Form10

    Function RGB2ARGB(ByVal n As Integer) As Integer
        RGB2ARGB = (n Mod 256) << 16 Or ((n \ 256) Mod 256) << 8 Or n \ 65536 Or 255 << 24
    End Function

    Public Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        Timer1.Enabled = CheckBox1.Checked
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
    '颜色转换为值
    Public Function ColRTrs(ByVal n As Long) As Long 'RGB to n
        Dim r, g, b As Integer
        r = Color.FromArgb(RGB2ARGB(n)).R
        g = Color.FromArgb(RGB2ARGB(n)).G
        b = Color.FromArgb(RGB2ARGB(n)).B
        If r = 255 And g = 0 Then
            ColRTrs = (255 - b) * 5000 / 255
        ElseIf r = 255 And b = 0 Then
            ColRTrs = g * 5000 / 255 + 5000
        ElseIf g = 255 And b = 0 Then
            ColRTrs = (255 - r) * 5000 / 255 + 10000
        ElseIf r = 0 And g = 255 Then
            ColRTrs = b * 5000 / 255 + 15000
        ElseIf r = 0 And b = 255 Then
            ColRTrs = (255 - g) * 5000 / 255 + 20000
        ElseIf g = 0 And b = 255 Then
            ColRTrs = r * 5000 / 255 + 25000
        Else
            ColRTrs = 0
            MessageBox.Show("Error! R" & r & " G" & g & " B" & b)
        End If

        ColRTrs -= 12500
        ColRTrs *= -1
    End Function

    '-12500,12500 = (0,25000)-12500
    '-12500,7500 =(0,20000)-12500
    '值转换为颜色
    Public Function ColTrs(ByVal n As Long) As Long 'n to RGB
        Dim r, g, b As Integer
        n = -n + 12500

        If n >= 0 And n <= 25000 Then
            'b = 0 'b
            'g = 0 'g
            'r = 255 'r

            '255 0 0
            '255 255 0
            '0 255 0
            '0 255 255
            '0 0 255
            '255 0 255
            If n <= 5000 Then
                r = 255 'r
                g = 0 'g
                b = 255 - n * 255 \ 5000 'b

            ElseIf n <= 10000 Then
                r = 255 'r
                g = (n - 5000) * 255 \ 5000 'g
                b = 0 'b

            ElseIf n <= 15000 Then
                r = 255 - (n - 10000) * 255 \ 5000 'r
                g = 255 'g
                b = 0 'b

            ElseIf n <= 20000 Then
                r = 0 'r
                g = 255 'g
                b = (n - 15000) * 255 \ 5000 'b

            ElseIf n <= 25000 Then
                r = 0 'r
                g = 255 - (n - 20000) * 255 \ 5000 'g
                b = 255 'b
            End If
        ElseIf n > 0 Then
            r = 0
            g = 0
            b = 255
        Else
            r = 255
            g = 0
            b = 255
        End If
        ColTrs = RGB(r, g, b)
    End Function

    '更改光源位置
    Public Sub ChangeSourcePosition(ByVal p As Long)
        SourceY = p
        If SourceY > 0 Then
            For i As Long = 0 To mw - 1
                For j As Long = 0 To SourceY - 1
                    Intensity(i, j) = 0
                Next
            Next
        End If
    End Sub

    '调试输出
    Sub DebugOutput()
        Dim oo As String = ""
        For k As Long = 0 To mh - 1
            oo &= Monomer(0, k) & vbCrLf
        Next
        Clipboard.SetText(oo)
    End Sub

    '计算下一步（并行）
    Sub CalcNextFrame()
        For i As Long = 0 To mw - 1
            Intensity(i, SourceY) = I_0
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
    Sub RefreshIntensity(i As Long)
        For j As Long = SourceY + 1 To mh - 1
            Try
                If x(i, j) < NJD Then
                    Intensity(i, j) = Intensity(i, j - 1) * Math.Exp((-Epsilon * PI(i, j)) * ScaleMeterPerPixel - A0l * ScaleMeterPerPixel) * Mask(i, j)
                Else
                    Intensity(i, j) = Intensity(i, j - 1) * Math.Exp((-Epsilon * PI(i, j)) * ScaleMeterPerPixel - A0s * ScaleMeterPerPixel) * Mask(i, j)
                End If
                'Intensity(i, j) = Intensity(i, j - 1) * Math.Exp((-Epsilon * PI(i, j) - A0l) * ScaleMeterPerPixel) * Mask(i, j)
            Catch ex As Exception
                Intensity(i, j) = 0
            End Try
        Next
    End Sub
    Sub RefreshR_I(i As Long)
        For j As Long = 0 To mh - 1
            R_I(i, j) = Phi * Epsilon * PI(i, j) * Intensity(i, j)
        Next
    End Sub
    Sub RefreshPI(i As Long)
        For j As Long = 0 To mh - 1
            PI(i, j) += -0.5 * R_I(i, j) * StepLength
        Next
    End Sub
    Sub RefreshIRadical(i As Long)
        For j As Long = 0 To mh - 1
            IRadical(i, j) += (R_I(i, j) - k_i * IRadical(i, j) * Monomer(i, j)) * StepLength
        Next
    End Sub
    Sub RefreshMonomer(i As Long)
        For j As Long = 0 To mh - 1
            Monomer(i, j) += (-k_i * IRadical(i, j) * Monomer(i, j) - k_p(i, j) * Monomer(i, j) * M_nRadical(i, j)) * StepLength
        Next
    End Sub
    Sub RefreshM_nRadical(i As Long)
        For j As Long = 0 To mh - 1
            M_nRadical(i, j) += (k_i * IRadical(i, j) * Monomer(i, j) - k_t(i, j) * M_nRadical(i, j) * M_nRadical(i, j) - k_O * M_nRadical(i, j) * O2Conc(i, j)) * StepLength
        Next
    End Sub
    Sub CalcDblDifferentialO2Conc()
        'Parallel.For(0, mw, AddressOf CalcDDOLoop1)
        Parallel.For(0, mw, AddressOf CalcDDOLoop2)
    End Sub
    Sub CalcDDOLoop1(i As Long)
        For j As Long = 1 To mh - 2
            O2ConcSDz(i, j) = (O2Conc(i, j + 1) - O2Conc(i, j - 1)) / 2 / ScaleMeterPerPixel
        Next
    End Sub
    Sub CalcDDOLoop2(i As Long)
        O2ConcDDz(i, 1) = (O2Conc(i, 3) - O2Conc(i, 1)) / 4 / ScaleMeterPerPixel / ScaleMeterPerPixel
        O2ConcDDz(i, mh - 2) = (O2Conc(i, mh - 4) - O2Conc(i, mh - 2)) / 4 / ScaleMeterPerPixel / ScaleMeterPerPixel
        For j As Long = 2 To mh - 3
            O2ConcDDz(i, j) = (O2Conc(i, j + 2) + O2Conc(i, j - 2) - 2 * O2Conc(i, j)) / 4 / ScaleMeterPerPixel / ScaleMeterPerPixel
        Next
    End Sub
    Sub RefreshO2Conc(i As Long)
        For j As Long = 0 To mh - 1
            O2Conc(i, j) += (D_O * O2ConcDDz(i, j) - k_O * O2Conc(i, j) * M_nRadical(i, j)) * StepLength
        Next
    End Sub
    Sub Refreshx(i As Long)
        For j As Long = 0 To mh - 1
            x(i, j) = 1 - Monomer(i, j) / M_0
        Next
    End Sub
    Sub Refreshk_pk_t(i As Long)
        Dim phi_m As Double, f As Double
        For j As Long = 0 To mh - 1
            phi_m = (1 - x(i, j)) / (1 - x(i, j) + Rho_m / Rho_p * x(i, j))
            f = f_m * phi_m + f_p * (1 - phi_m)
            k_p(i, j) = k_p0 / (1 + Math.Exp(-A_p * (1 / f_cp - 1 / f)))
            k_t(i, j) = k_t0 / (1 + 1 / (R * k_p(i, j) * Monomer(i, j) / k_t0 + Math.Exp(-A_t * (1 / f - 1 / f_ct))))
        Next
    End Sub

    '渲染画面
    Sub Render(ByRef img1 As Bitmap, ByRef img2 As Bitmap)

        MMax = Monomer(0, 0)
        DebugO &= t & vbTab & MMin & vbTab & vbCrLf
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
                Col = ColTrs(CInt((Monomer(i, j) - MMin) * 25000 / (MMax - MMin) - 12500))
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
                Col = ColTrs(CInt((Intensity(i, j) - IMin) * 25000 / (IMax - IMin) - 12500))
                bt(Offset + 2) = (Col And 255)
                Col >>= 8
                bt(Offset + 1) = (Col And 255)
                Col >>= 8
                bt(Offset) = Col
            Catch ex As Exception

            End Try
        Next
    End Sub

    Public Sub Timer2_Tick(sender As Object, e As EventArgs) Handles Timer2.Tick
        Label2.Text = "fps=" & fps * 1000 / Timer2.Interval
        fps = 0
        Label12.Top = Math.Min(PictureBox3.Top + PictureBox3.Height, Math.Max(PictureBox3.Top, PictureBox3.Top + (MMax - M_0 * (1 - NJD)) * PictureBox3.Height / (MMax - MMin)))
        If StartTask Then
            Start_Task()
            StartTask = False
        End If
    End Sub

    Public Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        RenderFramCount = RenderIntv
        Timer1_Tick(sender, e)
    End Sub

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
        ReDim Mask(mw - 1, mh - 1)

        For i As Long = 0 To mw - 1
            For j As Long = 0 To mh - 1
                PI(i, j) = PI_0
                IRadical(i, j) = 0
                Monomer(i, j) = M_0
                Intensity(i, j) = 0
                k_p(i, j) = k_p0
                M_nRadical(i, j) = 0
                k_t(i, j) = k_t0
                x(i, j) = 0
                O2Conc(i, j) = O_2_0
                R_I(i, j) = 0
                O2ConcSDz(i, j) = 0
                O2ConcDDz(i, j) = 0
                Mask(i, j) = 1
            Next
        Next
        NumericUpDown4.Maximum = mh - 1
        Button7_Click(New Object, New EventArgs)
        ChangeSourcePosition(mh - LayerHeight)
        t = 0
        CurrentLayer = 1
        Button7_Click(New Object, New EventArgs)
    End Sub
    Sub RefreshMap()
        Dim img1 As New Bitmap(mw, mh, Imaging.PixelFormat.Format24bppRgb)
        Dim img2 As New Bitmap(mw, mh, Imaging.PixelFormat.Format24bppRgb)
        Render(img1, img2)
        PictureBox1.Image = img1
        PictureBox2.Image = img2
    End Sub

    Public Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        frm3.Show(Me)

        'Clipboard.SetText(DebugO)
        'MessageBox.Show("已复制到剪贴板")
        'DebugO = ""
    End Sub

    Public Sub Button3_Click_1(sender As Object, e As EventArgs) Handles Button3.Click
        frm2.Show()
        frm4.Show()
    End Sub

    Public Sub NumericUpDown3_ValueChanged(sender As Object, e As EventArgs) Handles NumericUpDown3.ValueChanged
        RenderIntv = NumericUpDown3.Value
    End Sub

    Public Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        ChangeSourcePosition(NumericUpDown4.Value)
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

    Public Sub Button3_Click(sender As Object, e As EventArgs)
        'Timer1.Enabled = False
        'CheckBox1.Checked = False
        'While Working
        'End While
    End Sub

    Public Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        frm6.Visible = True
    End Sub

    Public Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        Working = True
        For m As Long = 0 To 100
            t += StepLength
            fps += 1
            If CheckBox2.Checked And t > TimePerLayer * CurrentLayer Then
                CurrentLayer += 1
                If CurrentLayer * LayerHeight > mh Then
                    Timer1.Enabled = False
                    frm3.SaveData()
                    frm8.RefreshData()
                    frm6.EndTask()
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
            frm3.SaveData()
            frm8.RefreshData()
            frm6.EndTask()
            Try
                frm4.RefreshEvent()
            Catch ex As Exception
            End Try
        Next
        Working = False
    End Sub

    Public Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        frm7.Show()
    End Sub

    Public Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Timer1.Enabled = False
        CheckBox1.Checked = False
        While Working
        End While
        mw = Val(NumericUpDown1.Value)
        mh = Val(NumericUpDown2.Value)
        frm3.NumericUpDown1.Maximum = mw - 1
        frm3.NumericUpDown2.Maximum = mh - 1
        ResetMap()
    End Sub

    Public Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'mw = 100
        'mh = 200
        frm2 = New Form2
        frm2.frm1 = Me
        If My.Computer.FileSystem.FileExists("Setting.csv") Then
            frm2.Show()
            frm2.AutoLoad = True
            frm2.Button5_Click(sender, e)
            frm2.AutoLoad = True
            frm2.Button2_Click(sender, e)
            frm2.AutoLoad = False
            frm2.Hide()
        End If
        ResetMap()
        RefreshMap()
        frm3 = New Form3
        frm3.frm1 = Me
        frm4 = New Form4
        frm4.frm1 = Me
        frm5 = New Form5
        frm5.frm1 = Me
        frm6 = New Form6
        frm6.frm1 = Me
        frm7 = New Form7
        frm7.frm1 = Me
        frm8 = New Form8
        frm8.frm1 = Me
        frm9 = New Form9
        frm9.frm1 = Me
        frm10 = New Form10
        frm10.frm1 = Me
        frm3.ShowDialog()
    End Sub
    Public Sub Start_Task()
        Button1_Click(New Object, New EventArgs)
        CheckBox1.Checked = True
    End Sub

    Public Sub Form1_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        frm3.Dispose()
        frm6.Dispose()
        frm7.Dispose()
        End
    End Sub

    Public Delegate Sub DChgMW(ByVal w As Long)
    Public Delegate Sub DChgMH(ByVal h As Long)
    Public Delegate Sub DChgLayers(ByVal l As Long)
    Public Delegate Sub DChgTimePerLayer(ByVal tpl As Long)

    Public Sub ChgMW(ByVal w As Long)
        Invoke(New DChgMW(AddressOf SChgMW), w)
    End Sub
    Public Sub SChgMW(ByVal w As Long)
        
        NumericUpDown1.Value = w
        Button4_Click(Nothing, Nothing)
    End Sub

    Public Sub ChgMH(ByVal h As Long)
        Invoke(New DChgMH(AddressOf SChgMH), h)
    End Sub
    Public Sub SChgMH(ByVal h As Long)
        NumericUpDown2.Value = h
        Button4_Click(Nothing, Nothing)
    End Sub

    Public Sub ChgLayers(ByVal l As Long)
        Invoke(New DChgLayers(AddressOf SChgLayers), l)
    End Sub
    Public Sub SChgLayers(ByVal l As Long)
        NumericUpDown5.Value = l
        Button7_Click(Nothing, Nothing)
    End Sub

    Public Sub ChgTimePerLayer(ByVal tpl As Long)
        Invoke(New DChgTimePerLayer(AddressOf SChgTimePerLayer), tpl)
    End Sub
    Public Sub SChgTimePerLayer(ByVal tpl As Double)
        TextBox1.Text = tpl.ToString
        Button7_Click(Nothing, Nothing)
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
