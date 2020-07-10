Imports System.ComponentModel

Public Class Form6
    Dim Param(,) As Double, count As Long = 0, pw As Long = 0
    Dim ParName() As String
    Dim CurrentTask As Long = 0
    Public frm1 As Form1
    Public Sub SaveData(ByRef Data(,) As Double, ByVal FilePrefix As String)
        Dim d As Double = 0
        For i As Long = 0 To Data.GetLength(1) - 1
            d += Data(0, i)
        Next
        d /= Data.GetLength(1)
        My.Computer.FileSystem.WriteAllText(FolderBrowserDialog1.SelectedPath & "\" & FilePrefix & "_" & CurrentTask & ".csv", Form1.t & "," & Data(0, 0) & "," & Data(0, Data.GetLength(1) - 1) & "," & d & vbCrLf, True)
    End Sub

    Public Sub NewResult()
        If CheckedListBox1.GetItemChecked(0) Then
            SaveData(frm1.Intensity, "Intensity")
        End If
        If CheckedListBox1.GetItemChecked(1) Then
            SaveData(frm1.R_I, "R_I")
        End If
        If CheckedListBox1.GetItemChecked(2) Then
            SaveData(frm1.PI, "PI")
        End If
        If CheckedListBox1.GetItemChecked(3) Then
            SaveData(frm1.IRadical, "IRadical")
        End If
        If CheckedListBox1.GetItemChecked(4) Then
            SaveData(frm1.Monomer, "Monomer")
        End If
        If CheckedListBox1.GetItemChecked(5) Then
            SaveData(frm1.M_nRadical, "M_nRadical")
        End If
        If CheckedListBox1.GetItemChecked(6) Then
            SaveData(frm1.O2Conc, "O2Conc")
        End If
        If CheckedListBox1.GetItemChecked(7) Then
            SaveData(frm1.k_p, "k_p")
        End If
        If CheckedListBox1.GetItemChecked(8) Then
            SaveData(frm1.k_t, "k_t")
        End If
        If CheckedListBox1.GetItemChecked(9) Then
            SaveData(frm1.x, "x")
        End If
    End Sub

    Public Sub EndTask()
        NewResult()

    End Sub

    Public Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        If FolderBrowserDialog1.ShowDialog = DialogResult.OK Then
            CurrentTask = 0
            Timer1.Enabled = True
            Button1.Enabled = False
            BackgroundWorker1.RunWorkerAsync()
        End If
    End Sub

    Public Sub BackgroundWorker1_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        For i As Long = 0 To count - 1
            CurrentTask = i
            For j As Long = 0 To pw - 1
                Select Case ParName(j).ToLower
                    Case "rho_m"
                        frm1.Rho_m = Param(i, j)
                    Case "rho_p"
                        frm1.Rho_p = Param(i, j)
                    Case "alpha_m"
                        frm1.Alpha_m = Param(i, j)
                    Case "alpha_p"
                        frm1.Alpha_p = Param(i, j)
                    Case "t_gm"
                        frm1.T_gm = Param(i, j)
                    Case "t_gp"
                        frm1.T_gp = Param(i, j)
                    Case "a_p"
                        frm1.A_p = Param(i, j)
                    Case "a_t"
                        frm1.A_t = Param(i, j)
                    Case "k_p0"
                        frm1.k_p0 = Param(i, j)
                    Case "k_t0"
                        frm1.k_t0 = Param(i, j)
                    Case "f_cp"
                        frm1.f_cp = Param(i, j)
                    Case "f_ct"
                        frm1.f_ct = Param(i, j)
                    Case "r"
                        frm1.R = Param(i, j)
                    Case "k_o"
                        frm1.k_O = Param(i, j)
                    Case "phi"
                        frm1.Phi = Param(i, j)
                    Case "epsilon"
                        frm1.Epsilon = Param(i, j)
                    Case "a_0l"
                        frm1.A0l = Param(i, j)
                    Case "a_0s"
                        frm1.A0s = Param(i, j)
                    Case "i_0"
                        frm1.I_0 = Param(i, j)
                    Case "temperature"
                        frm1.Temperature = Param(i, j)
                    Case "pi_0"
                        frm1.PI_0 = Param(i, j)
                    Case "m_0"
                        frm1.M_0 = Param(i, j)
                    Case "d_o"
                        frm1.D_O = Param(i, j)
                    Case "o2_0"
                        frm1.O_2_0 = Param(i, j)
                    Case "k_i"
                        frm1.k_i = Param(i, j)
                    Case "steplength"
                        frm1.StepLength = Param(i, j)
                    Case "scalemeterperpixel"
                        frm1.ScaleMeterPerPixel = Param(i, j)
                    Case "mw"
                        frm1.ChgMW(Param(i, j))
                    Case "mh"
                        frm1.ChgMH(Param(i, j))
                    Case "layers"
                        frm1.ChgLayers(Param(i, j))
                    Case "timeperlayer"
                        frm1.ChgTimePerLayer(Param(i, j))
                End Select
            Next
            While frm1.CheckBox1.Checked
            End While
            frm1.StartTask = True
            While frm1.StartTask
            End While
            While frm1.CheckBox1.Checked
            End While
        Next
    End Sub

    Public Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        ProgressBar1.Maximum = count
        ProgressBar1.Value = CurrentTask + 1
        Text = CurrentTask + 1 & "/" & count
    End Sub

    Public Sub Form6_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Public Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If OpenFileDialog1.ShowDialog = DialogResult.OK Then
            Dim t As String()
            Dim s() As String = My.Computer.FileSystem.ReadAllText(OpenFileDialog1.FileName).Split({vbCrLf}, StringSplitOptions.RemoveEmptyEntries)
            count = s.Count - 1
            ParName = s(0).Split({","}, StringSplitOptions.RemoveEmptyEntries)
            pw = ParName.Count
            ReDim Param(count - 1, pw - 1)
            For j As Long = 0 To count - 1
                t = s(j + 1).Split({","}, StringSplitOptions.RemoveEmptyEntries)
                For i As Long = 0 To pw - 1
                    Param(j, i) = Val(t(i))
                Next
            Next
        End If
    End Sub

    Public Sub BackgroundWorker1_RunWorkerCompleted(sender As Object, e As RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        Button1.Enabled = True
        Timer1.Enabled = False
    End Sub

    Public Sub Form6_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        Me.Hide()
        e.Cancel = True
    End Sub
End Class