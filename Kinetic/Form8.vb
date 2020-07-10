Public Class Form8
    Public frm1 As Form1
    Public matlab As MLApp.MLApp
    Public tmp(,) As Double
    Public PData(,) As Double
    Public CurrDir As String
    Public CoEff(4)() As Double
    Public ETELine As Integer = 0

    Public Function ETE(ByVal t As Double, ByRef Coeff() As Double) As Double
        ETE = Coeff(0) / (Coeff(6) + Math.Exp(Coeff(1) - Coeff(2) * (t + Coeff(7))) + Math.Exp(Coeff(4) - Coeff(5) * (t + Coeff(7)))) + Coeff(3)
    End Function

    Public Function RSquare(ByRef Coeff() As Double, ByRef Data(,) As Double) As Double
        Dim SSR As Double = 0, SST As Double = 0
        Dim Avg As Double = 0
        For i As Long = 0 To Data.GetLength(0) - 1
            Avg += Data(i, 1)
        Next
        Avg /= Data.GetLength(0)
        For i As Long = 0 To Data.GetLength(0) - 1
            SSR += (ETE(Data(i, 0), Coeff) - Avg) ^ 2
            SST += (Data(i, 1) - Avg) ^ 2
        Next
        RSquare = SSR / SST
    End Function

    Public Sub RefreshData()
        If Not Visible Then Exit Sub
        Try
            tmp(0, tmp.GetLength(1) - 1) = frm1.t
            Dim avg As Double = 0
            For i As Long = 0 To frm1.mh - 1
                avg += frm1.x(frm1.frm3.NumericUpDown1.Value, frm1.frm3.NumericUpDown2.Value)
            Next
            avg /= frm1.mh
            tmp(1, tmp.GetLength(1) - 1) = avg
            ReDim Preserve tmp(1, tmp.GetLength(1))
        Catch
        End Try
    End Sub
    Public Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        TextBox1.Text = "t=0:0.01:120;" & vbCrLf
        TextBox1.Text &= "[a,b,c,d,e,f,g]=deal(1.663,1.851,0.1311,-0.01299+0.0007145,7.631,1.399,4.988);  %[1]2%-5 粉" & vbCrLf
        TextBox1.Text &= "x=a./(g+exp(b-c.*(t+2))+exp(e-f.*(t+2)))+d;" & vbCrLf
        TextBox1.Text &= "hold on;" & vbCrLf
        TextBox1.Text &= "plot(t,x,'m','LineWidth',3);" & vbCrLf
        TextBox1.Text &= "x=[1.95 3.9 5.84 7.79 9.74 11.69 13.64 17.54 21.44 33.13 44.82 60.42 62.36 72.11 81.85 91.6];" & vbCrLf
        TextBox1.Text &= "x=x-1.95;" & vbCrLf
        TextBox1.Text &= "y =[0 0.077586207 0.191304348 0.215517241 0.222222222 0.245762712 0.264957265 0.275862069 0.307692308 0.327731092 0.299145299 0.330508475 0.319327731 0.325 0.31092437 0.322033898];" & vbCrLf
        TextBox1.Text &= "scatter(x,y);" & vbCrLf
        ETELine = 0
    End Sub

    Public Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        TextBox1.Text = "t=0:0.01:120;" & vbCrLf
        TextBox1.Text &= "[a,b,c,d,e,f,g]=deal(3.074,4.995,0.1809,-0.0006072-0.000393,9.842,0.9251,14.98);  %[2]1%-20 红" & vbCrLf
        TextBox1.Text &= "x=a./(g+exp(b-c.*(t+2))+exp(e-f.*(t+2)))+d;" & vbCrLf
        TextBox1.Text &= "hold on;" & vbCrLf
        TextBox1.Text &= "plot(t,x,'r','LineWidth',3);" & vbCrLf
        TextBox1.Text &= "x=[1.95 5.84 9.74 13.64 15.59 17.54 19.49 21.43 27.28 29.23 42.87 64.31 70.15 79.9 87.69];" & vbCrLf
        TextBox1.Text &= "x=x-1.95;" & vbCrLf
        TextBox1.Text &= "y =[0 0.019607843 0.07120743 0.117647059 0.127647059 0.137647059 0.159663866 0.161764706 0.201680672 0.201680672 0.197860963 0.205882353 0.201680672 0.201680672 0.205882353];" & vbCrLf
        TextBox1.Text &= "scatter(x,y);" & vbCrLf
        ETELine = 1
    End Sub

    Public Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        TextBox1.Text = "t=0:0.01:120;" & vbCrLf
        TextBox1.Text &= "[a,b,c,d,e,f,g]=deal(2.299,7.596,1.607,-0.02461-0.002053,0.74571,0.0654,4.378);  %[3]2%-20 绿" & vbCrLf
        TextBox1.Text &= "x=a./(g+exp(b-c.*(t+2))+exp(e-f.*(t+2)))+d;" & vbCrLf
        TextBox1.Text &= "hold on;" & vbCrLf
        TextBox1.Text &= "plot(t,x,'Color',[0 0.5 0]','LineWidth',3);" & vbCrLf
        TextBox1.Text &= "x=[1.95 3.9 5.84 9.74 13.64 15.59 17.54 21.44 31.18 37.03 42.87 52.62 62.36 72.11 81.85 91.6];" & vbCrLf
        TextBox1.Text &= "x=x-1.95;" & vbCrLf
        TextBox1.Text &= "y =[0 0.21154385 0.352759877 0.395348837 0.416335791 0.426162791 0.428247306 0.463981849 0.44015882 0.475893364 0.493540052 0.505598622 0.487804878 0.493540052 0.493540052 0.511627907];" & vbCrLf
        TextBox1.Text &= "scatter(x,y);" & vbCrLf
        ETELine = 2
    End Sub

    Public Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        TextBox1.Text = "t=0:0.01:120;" & vbCrLf
        TextBox1.Text &= "[a,b,c,d,e,f,g]=deal(1.086,2.194,0.1116,-0.004814-0.001797,7.059,1.012,3.394);  %[4]3%-20 深蓝" & vbCrLf
        TextBox1.Text &= "x=a./(g+exp(b-c.*(t+2))+exp(e-f.*(t+2)))+d;" & vbCrLf
        TextBox1.Text &= "hold on;" & vbCrLf
        TextBox1.Text &= "plot(t,x,'Color',[0.08 0.17 0.55]','LineWidth',3);" & vbCrLf
        TextBox1.Text &= "x=[1.95 3.9 5.84 7.79 9.74 13.64 17.54 21.43 25.33 29.23 44.82 60.41 70.15 89.64];" & vbCrLf
        TextBox1.Text &= "x=x-1.95;" & vbCrLf
        TextBox1.Text &= "y =[0 0.03125 0.088235294 0.143493761 0.168685121 0.19550173 0.218487395 0.244537815 0.275951557 0.296638655 0.322689076 0.322689076 0.290849673 0.309269162];" & vbCrLf
        TextBox1.Text &= "scatter(x,y);" & vbCrLf
        ETELine = 3
    End Sub

    Public Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        TextBox1.Text = "t=0:0.01:120;" & vbCrLf
        TextBox1.Text &= "[a,b,c,d,e,f,g]=deal(6.645,11.16,1.297,-0.0001906-0.01571,3.219,0.2489,16.35);  %[5]2%-40 蓝" & vbCrLf
        TextBox1.Text &= "x=a./(g+exp(b-c.*(t+4))+exp(e-f.*(t+4)))+d;" & vbCrLf
        TextBox1.Text &= "hold on;" & vbCrLf
        TextBox1.Text &= "plot(t,x,'b','LineWidth',3);" & vbCrLf
        TextBox1.Text &= "x=[1.95 3.9 5.84 7.79 9.74 11.79 13.64 17.54 31.18 40.92 50.67 56.51 62.36 66.26 76 81.85 83.79];" & vbCrLf
        TextBox1.Text &= "x=x-1.95;" & vbCrLf
        TextBox1.Text &= "y =[0 0.017857143 0.103448276 0.310344828 0.35 0.372881356 0.383333333 0.383333333 0.4 0.416666667 0.4 0.4 0.423728814 0.440677966 0.442622951 0.440677966 0.45];" & vbCrLf
        TextBox1.Text &= "scatter(x,y);" & vbCrLf
        ETELine = 4
    End Sub

    Public Sub Form8_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        matlab = CType(Activator.CreateInstance(Type.GetTypeFromProgID("Matlab.Application")), MLApp.MLApp)
        CoEff(0) = {1.663, 1.851, 0.1311, -0.01299 + 0.0007145, 7.631, 1.399, 4.988, 2}
        CoEff(1) = {3.074, 4.995, 0.1809, -0.0006072 - 0.004968, 9.842, 0.9251, 14.98, 2}
        CoEff(2) = {2.299, 7.596, 1.607, -0.02461 - 0.002053, 0.74571, 0.0654, 4.378, 2}
        CoEff(3) = {1.086, 2.194, 0.1116, -0.004814 - 0.001797, 7.059, 1.012, 3.394, 2}
        CoEff(4) = {6.645, 11.16, 1.297, -0.0001906 - 0.01571, 3.219, 0.2489, 16.35, 4}
    End Sub
    Public Sub Plot(ByRef data(,) As Double, Optional ByVal xmin As Double = 0, Optional ByVal xmax As Double = 0, Optional ymin As Double = 0, Optional ByVal ymax As Double = 0, Optional ByVal ReverseX As Boolean = False, Optional ByVal ReverseY As Boolean = False, Optional ByVal Holdon As Boolean = False)
        Dim h As Long = data.GetLength(0)
        Dim command As String = "x=[" & data(0, 0)
        If h > 1 Then
            For i As Long = 1 To h - 1
                command &= ";" & data(i, 0)
            Next
        End If
        command &= "]; y=[" & data(0, 1)
        If h > 1 Then
            For i As Long = 1 To h - 1
                command &= ";" & data(i, 1)
            Next
        End If
        If Holdon Then
            command &= "]; hold on; plot(x,y);"
        Else
            command &= "]; figure1 = figure; axes1 = axes('Parent',figure1); "
            command &= "hold(axes1,'on'); plot(x,y); "
            If xmax > xmin Then command &= "xlim(axes1,[" & xmin & " " & xmax & "]); "
            If ymax > ymin Then command &= "ylim(axes1,[" & ymin & " " & ymax & "]); "
            command &= "box(axes1,'on'); "
            If ReverseX Then command &= "set(axes1,'XDir','reverse');"
            If ReverseY Then command &= "set(axes1,'YDir','reverse');"
        End If
        matlab.Visible = 1
        matlab.Execute(command)
    End Sub
    Public Sub Form8_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        e.Cancel = True
        Hide()
    End Sub

    Public Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        frm1.ChgMW(1)
        frm1.ChgMH(200)
        frm1.ChgLayers(1)
        frm1.ChgTimePerLayer(120)
        frm1.Button1_Click(sender, e)
        ReDim tmp(1, 0)
    End Sub

    Public Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        Button6_Click(sender, e)
        frm1.StartTask = True
    End Sub

    Public Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        If TextBox1.Text <> "" Then
            matlab.Execute(TextBox1.Text)
        End If
    End Sub

    Public Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        matlab.Execute("figure;")
    End Sub
    Public Sub RefreshPData()
        Try
            If tmp.GetLength(1) <= 1 Then Exit Sub
        Catch
            Exit Sub
        End Try

        ReDim PData(tmp.GetLength(1) - 2, 1)
        For i As Long = 0 To tmp.GetLength(1) - 2
            PData(i, 0) = tmp(0, i)
            PData(i, 1) = tmp(1, i)
        Next
    End Sub

    Public Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click
        RefreshPData()
        If tmp.GetLength(1) <= 1 Then Exit Sub
        Plot(PData,,,,,,, True)
    End Sub
    Public Sub SaveCSV(ByRef Data(,) As Double, ByVal FileName As String)
        Dim s As String = ""
        For i As Long = 0 To Data.GetLength(0) - 1
            For j As Long = 0 To 1
                s &= Data(i, 0) & "," & Data(i, 1) & vbCrLf
            Next
        Next
        My.Computer.FileSystem.WriteAllText(FileName, s, False)
    End Sub

    Public Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click
        RefreshPData()
        SaveCSV(PData, CurrDir & "\" & TextBox2.Text & ".csv")
    End Sub

    Public Sub Button12_Click(sender As Object, e As EventArgs) Handles Button12.Click
        If FolderBrowserDialog1.ShowDialog = DialogResult.OK Then
            CurrDir = FolderBrowserDialog1.SelectedPath
        End If
    End Sub

    Public Sub Button14_Click(sender As Object, e As EventArgs) Handles Button14.Click
        RefreshPData()
        If False Then
            TextBox3.Text = RSquare(CoEff(ETELine), PData).ToString
        Else
            Dim s As String = "T=["
            For i As Long = 0 To PData.GetLength(0) - 1
                s &= ETE(PData(i, 0), CoEff(ETELine)) & " "
            Next
            s &= "]; Y=["
            For i As Integer = 0 To PData.GetLength(0) - 1
                s &= PData(i, 1) & " "
            Next
            s &= "]; R=corrcoef(T,Y);" '& " R2 = norm(T -mean(Y))^2/norm(Y - mean(Y))^2;"
            matlab.Execute(s)
            Dim PR(1, 1) As Double, PI(1, 1) As Double ', R2(0, 0) As Double, I2(0, 0) As Double
            matlab.GetFullMatrix("R", "base", PR, PI)
            'matlab.GetFullMatrix("R2", "base", R2, I2)
            TextBox3.Text = PR(0, 1).ToString ' & vbCrLf & R2(0, 0))
        End If


    End Sub

    Public Sub Button13_Click(sender As Object, e As EventArgs) Handles Button13.Click
        If OpenFileDialog1.ShowDialog = DialogResult.OK Then
            PData = frm1.frm2.DblFromCsv(OpenFileDialog1.FileName)
            ReDim tmp(1, PData.GetLength(0) - 1)
            For i As Integer = 0 To PData.GetLength(0) - 1
                tmp(0, i) = PData(i, 0)
                tmp(1, i) = PData(i, 1)
            Next
        End If
    End Sub

    Public Sub Button15_Click(sender As Object, e As EventArgs) Handles Button15.Click
        frm1.frm2.Button5_Click(sender, e)
    End Sub

    Public Sub Button16_Click(sender As Object, e As EventArgs) Handles Button16.Click
        frm1.frm2.Button13_Click(sender, e)
    End Sub
End Class