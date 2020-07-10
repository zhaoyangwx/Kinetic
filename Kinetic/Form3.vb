Public Class Form3
    Public Intensity(,), R_I(,), PI(,), IRadical(,), Monomer(,), M_nRadical(,), O2Conc(,), k_p(,), k_t(,), x(,) As Double
    Public tmp(,) As Double
    Public frm1 As Form1, frm7 As Object
    Public PForm As Long = 1
    Public pfrm As Object '泛型，首次使用重载

    Public Sub New()
        ' 此调用是设计器所必需的。
        InitializeComponent()

        ' 在 InitializeComponent() 调用之后添加任何初始化。

    End Sub

    Public Sub Form3_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ReDim tmp(1, 0)
    End Sub
    Public Sub MatToFile(ByRef m(,) As Double, ToFileName As String)
        Dim s As String = ""
        For j As Long = 0 To m.GetLength(1) - 1
            For i As Long = 0 To m.GetLength(0) - 1
                s &= m(i, j) & ","
            Next
            s &= vbCrLf
        Next
        My.Computer.FileSystem.WriteAllText(ToFileName, s, False, System.Text.Encoding.ASCII)
    End Sub
    Public Sub TextToFile(ByRef s As String, ToFileName As String)


        My.Computer.FileSystem.WriteAllText(ToFileName, s, True, System.Text.Encoding.ASCII)
        Dim c As Long = tmp.GetLength(1)
        tmp(0, c - 1) = pfrm.t
        tmp(1, c - 1) = pfrm.x(0, 0)
        ReDim Preserve tmp(1, c)

    End Sub
    Public Sub AppendData(ByRef n(,) As Double, ByRef Line(,) As Double)
        ReDim Preserve n(n.GetLength(0) - 1, n.GetLength(1))
        For i As Long = 0 To pfrm.mh - 1
            n(i, n.GetLength(1) - 1) = Line(NumericUpDown1.Value, i)
        Next
    End Sub
    Public Sub SaveData()
        If CheckBox1.Checked Then
            If RadioButton1.Checked Then
                If CheckedListBox1.GetItemChecked(0) Then
                    MatToFile(pfrm.Intensity, TextBox1.Text & "\Intensity_" & Math.Round(pfrm.t, CInt(-Math.Log10(pfrm.StepLength))) & ".csv")
                End If
                If CheckedListBox1.GetItemChecked(1) Then
                    MatToFile(pfrm.R_I, TextBox1.Text & "\R_i_" & Math.Round(pfrm.t, CInt(-Math.Log10(pfrm.StepLength))) & ".csv")
                End If
                If CheckedListBox1.GetItemChecked(2) Then
                    MatToFile(pfrm.PI, TextBox1.Text & "\PI_" & Math.Round(pfrm.t, CInt(-Math.Log10(pfrm.StepLength))) & ".csv")
                End If
                If CheckedListBox1.GetItemChecked(3) Then
                    MatToFile(pfrm.IRadical, TextBox1.Text & "\IRadical_" & Math.Round(pfrm.t, CInt(-Math.Log10(pfrm.StepLength))) & ".csv")
                End If
                If CheckedListBox1.GetItemChecked(4) Then
                    MatToFile(pfrm.Monomer, TextBox1.Text & "\Monomer_" & Math.Round(pfrm.t, CInt(-Math.Log10(pfrm.StepLength))) & ".csv")
                End If
                If CheckedListBox1.GetItemChecked(5) Then
                    MatToFile(pfrm.M_nRadical, TextBox1.Text & "\M_nRadical_" & Math.Round(pfrm.t, CInt(-Math.Log10(pfrm.StepLength))) & ".csv")
                End If
                If CheckedListBox1.GetItemChecked(6) Then
                    MatToFile(pfrm.O2Conc, TextBox1.Text & "\O2Conc_" & Math.Round(pfrm.t, CInt(-Math.Log10(pfrm.StepLength))) & ".csv")
                End If
                If CheckedListBox1.GetItemChecked(7) Then
                    MatToFile(pfrm.k_p, TextBox1.Text & "\k_p_" & Math.Round(pfrm.t, CInt(-Math.Log10(pfrm.StepLength))) & ".csv")
                End If
                If CheckedListBox1.GetItemChecked(8) Then
                    MatToFile(pfrm.k_t, TextBox1.Text & "\k_t_" & Math.Round(pfrm.t, CInt(-Math.Log10(pfrm.StepLength))) & ".csv")
                End If
                If CheckedListBox1.GetItemChecked(9) Then
                    MatToFile(pfrm.x, TextBox1.Text & "\x_" & Math.Round(pfrm.t, CInt(-Math.Log10(pfrm.StepLength))) & ".csv")
                End If
            ElseIf RadioButton2.Checked Then
                If CheckedListBox1.GetItemChecked(0) Then
                    TextToFile(pfrm.t.ToString & "," & pfrm.Intensity(NumericUpDown1.Value, NumericUpDown2.Value).ToString & vbCrLf, TextBox1.Text & "\Intensity.csv")
                End If
                If CheckedListBox1.GetItemChecked(1) Then
                    TextToFile(pfrm.t.ToString & "," & pfrm.R_I(NumericUpDown1.Value, NumericUpDown2.Value).ToString & vbCrLf, TextBox1.Text & "\R_i.csv")
                End If
                If CheckedListBox1.GetItemChecked(2) Then
                    TextToFile(pfrm.t.ToString & "," & pfrm.PI(NumericUpDown1.Value, NumericUpDown2.Value).ToString & vbCrLf, TextBox1.Text & "\PI.csv")
                End If
                If CheckedListBox1.GetItemChecked(3) Then
                    TextToFile(pfrm.t.ToString & "," & pfrm.IRadical(NumericUpDown1.Value, NumericUpDown2.Value).ToString & vbCrLf, TextBox1.Text & "\IRadical.csv")
                End If
                If CheckedListBox1.GetItemChecked(4) Then
                    TextToFile(pfrm.t.ToString & "," & pfrm.Monomer(NumericUpDown1.Value, NumericUpDown2.Value).ToString & vbCrLf, TextBox1.Text & "\Monomer.csv")
                End If
                If CheckedListBox1.GetItemChecked(5) Then
                    TextToFile(pfrm.t.ToString & "," & pfrm.M_nRadical(NumericUpDown1.Value, NumericUpDown2.Value).ToString & vbCrLf, TextBox1.Text & "\M_nRadical.csv")
                End If
                If CheckedListBox1.GetItemChecked(6) Then
                    TextToFile(pfrm.t.ToString & "," & pfrm.O2Conc(NumericUpDown1.Value, NumericUpDown2.Value).ToString & vbCrLf, TextBox1.Text & "\O2Conc.csv")
                End If
                If CheckedListBox1.GetItemChecked(7) Then
                    TextToFile(pfrm.t.ToString & "," & pfrm.k_p(NumericUpDown1.Value, NumericUpDown2.Value).ToString & vbCrLf, TextBox1.Text & "\k_p.csv")
                End If
                If CheckedListBox1.GetItemChecked(8) Then
                    TextToFile(pfrm.t.ToString & "," & pfrm.k_t(NumericUpDown1.Value, NumericUpDown2.Value).ToString & vbCrLf, TextBox1.Text & "\k_t.csv")
                End If
                If CheckedListBox1.GetItemChecked(9) Then
                    TextToFile(pfrm.t.ToString & "," & pfrm.x(NumericUpDown1.Value, NumericUpDown2.Value).ToString & vbCrLf, TextBox1.Text & "\x.csv")
                End If
            Else
                If CheckedListBox1.GetItemChecked(0) Then
                    AppendData(Intensity, pfrm.Intensity)
                End If
                If CheckedListBox1.GetItemChecked(1) Then
                    AppendData(R_I, pfrm.R_I)
                End If
                If CheckedListBox1.GetItemChecked(2) Then
                    AppendData(PI, pfrm.PI)
                End If
                If CheckedListBox1.GetItemChecked(3) Then
                    AppendData(IRadical, pfrm.IRadical)
                End If
                If CheckedListBox1.GetItemChecked(4) Then
                    AppendData(Monomer, pfrm.Monomer)
                End If
                If CheckedListBox1.GetItemChecked(5) Then
                    AppendData(M_nRadical, pfrm.M_nRadical)
                End If
                If CheckedListBox1.GetItemChecked(6) Then
                    AppendData(O2Conc, pfrm.O2Conc)
                End If
                If CheckedListBox1.GetItemChecked(7) Then
                    AppendData(k_p, pfrm.k_p)
                End If
                If CheckedListBox1.GetItemChecked(8) Then
                    AppendData(k_t, pfrm.k_t)
                End If
                If CheckedListBox1.GetItemChecked(9) Then
                    AppendData(x, pfrm.x)
                End If
            End If
        End If

    End Sub

    Public Sub NumericUpDown1_ValueChanged(sender As Object, e As EventArgs) Handles NumericUpDown1.ValueChanged
        NumericUpDown1.Maximum = pfrm.mw - 1
    End Sub

    Public Sub RadioButton3_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton3.CheckedChanged
        If RadioButton3.Checked Then
            Button5_Click(sender, e)
        End If
    End Sub

    Public Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        Dim MatLabFigEngine As New Figure
        Dim k As Long = tmp.GetLength(1) - 1
        Dim Pdata(k - 1, 1) As Double
        For i As Long = 0 To tmp.GetLength(1) - 2
            Pdata(i, 0) = tmp(0, i)
            Pdata(i, 1) = tmp(1, i)
        Next
        MatLabFigEngine.Plot(Pdata)
    End Sub

    Public Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        ReDim tmp(1, 0)
    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        frm1.frm10.Show()
    End Sub

    Public Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        pfrm.NumericUpDown1.Value = 1
        pfrm.NumericUpDown2.Value = 2
        pfrm.NumericUpDown5.Value = 1
        pfrm.TextBox1.Text = "10"
        pfrm.Button4_Click(sender, e)
        pfrm.Button7_Click(sender, e)
    End Sub



    Public Sub NumericUpDown2_ValueChanged(sender As Object, e As EventArgs) Handles NumericUpDown2.ValueChanged
        NumericUpDown2.Maximum = pfrm.mh - 1
    End Sub

    Public Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        If CheckBox1.Checked And TextBox1.Text = "" Then Button1_Click(sender, e)
        CheckedListBox1.Enabled = CheckBox1.Checked
        GroupBox1.Enabled = CheckBox1.Checked
        If PForm = 1 Then
            pfrm = frm1
        Else
            pfrm = frm7
        End If
    End Sub

    Public Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If FolderBrowserDialog1.ShowDialog = DialogResult.OK Then
            TextBox1.Text = FolderBrowserDialog1.SelectedPath
        End If
    End Sub

    Public Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.Hide()
    End Sub

    Public Sub Form3_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        Me.Hide()
        e.Cancel = True
    End Sub

    Public Sub RadioButton1_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton1.CheckedChanged
        NumericUpDown1.Enabled = RadioButton2.Checked
        NumericUpDown2.Enabled = RadioButton2.Checked

    End Sub

    Public Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        SaveData()
    End Sub

    Public Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        If RadioButton3.Checked Then
            If CheckedListBox1.GetItemChecked(0) Then
                ReDim Intensity(pfrm.mh - 1, 0)
            End If
            If CheckedListBox1.GetItemChecked(1) Then
                ReDim R_I(pfrm.mh - 1, 0)
            End If
            If CheckedListBox1.GetItemChecked(2) Then
                ReDim PI(pfrm.mh - 1, 0)
            End If
            If CheckedListBox1.GetItemChecked(3) Then
                ReDim IRadical(pfrm.mh - 1, 0)
            End If
            If CheckedListBox1.GetItemChecked(4) Then
                ReDim Monomer(pfrm.mh - 1, 0)
            End If
            If CheckedListBox1.GetItemChecked(5) Then
                ReDim M_nRadical(pfrm.mh - 1, 0)
            End If
            If CheckedListBox1.GetItemChecked(6) Then
                ReDim O2Conc(pfrm.mh - 1, 0)
            End If
            If CheckedListBox1.GetItemChecked(7) Then
                ReDim k_p(pfrm.mh - 1, 0)
            End If
            If CheckedListBox1.GetItemChecked(8) Then
                ReDim k_t(pfrm.mh - 1, 0)
            End If
            If CheckedListBox1.GetItemChecked(9) Then
                ReDim x(pfrm.mh - 1, 0)
            End If
        End If
    End Sub

    Public Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        If RadioButton3.Checked Then
            If CheckedListBox1.GetItemChecked(0) Then
                MatToFile(Intensity, TextBox1.Text & "\Intensity_x=" & NumericUpDown1.Value & ".csv")
            End If
            If CheckedListBox1.GetItemChecked(1) Then
                MatToFile(R_I, TextBox1.Text & "\R_i_x=" & NumericUpDown1.Value & ".csv")
            End If
            If CheckedListBox1.GetItemChecked(2) Then
                MatToFile(PI, TextBox1.Text & "\PI_x=" & NumericUpDown1.Value & ".csv")
            End If
            If CheckedListBox1.GetItemChecked(3) Then
                MatToFile(IRadical, TextBox1.Text & "\IRadical_x=" & NumericUpDown1.Value & ".csv")
            End If
            If CheckedListBox1.GetItemChecked(4) Then
                MatToFile(Monomer, TextBox1.Text & "\Monomer_x=" & NumericUpDown1.Value & ".csv")
            End If
            If CheckedListBox1.GetItemChecked(5) Then
                MatToFile(M_nRadical, TextBox1.Text & "\M_nRadical_x=" & NumericUpDown1.Value & ".csv")
            End If
            If CheckedListBox1.GetItemChecked(6) Then
                MatToFile(O2Conc, TextBox1.Text & "\O2Conc_x=" & NumericUpDown1.Value & ".csv")
            End If
            If CheckedListBox1.GetItemChecked(7) Then
                MatToFile(k_p, TextBox1.Text & "\k_p_x=" & NumericUpDown1.Value & ".csv")
            End If
            If CheckedListBox1.GetItemChecked(8) Then
                MatToFile(k_t, TextBox1.Text & "\k_t_x=" & NumericUpDown1.Value & ".csv")
            End If
            If CheckedListBox1.GetItemChecked(9) Then
                MatToFile(x, TextBox1.Text & "\x_x=" & NumericUpDown1.Value & ".csv")
            End If
        End If
    End Sub
End Class