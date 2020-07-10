Public Class Figure
    Public Sub Plot(ByRef data(,) As String, Optional ByVal xmin As Double = 0, Optional ByVal xmax As Double = 0, Optional ymin As Double = 0, Optional ByVal ymax As Double = 0, Optional ByVal ReverseX As Boolean = False, Optional ByVal ReverseY As Boolean = False, Optional ByVal Holdon As Boolean = False)
        Dim h As Long = data.GetLength(0)
        Dim matlab As MLApp.MLApp = CType(Activator.CreateInstance(Type.GetTypeFromProgID("Matlab.Application")), MLApp.MLApp)
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

        matlab.Visible = 0
        matlab.Execute(command)
    End Sub
    Public Sub Plot(ByRef data(,) As Double, Optional ByVal xmin As Double = 0, Optional ByVal xmax As Double = 0, Optional ymin As Double = 0, Optional ByVal ymax As Double = 0, Optional ByVal ReverseX As Boolean = False, Optional ByVal ReverseY As Boolean = False, Optional ByVal Holdon As Boolean = False)
        Dim h As Long = data.GetLength(0)
        Dim matlab As MLApp.MLApp = Activator.CreateInstance(Type.GetTypeFromProgID("Matlab.Application"))
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
        matlab.Visible = 0
        matlab.Execute(command)
    End Sub
    Public Sub Scatter(ByRef data(,) As String, Optional ByVal xmin As Double = 0, Optional ByVal xmax As Double = 0, Optional ymin As Double = 0, Optional ByVal ymax As Double = 0, Optional ByVal ReverseX As Boolean = False, Optional ByVal ReverseY As Boolean = False)
        Dim h As Long = data.GetLength(0)
        Dim matlab As MLApp.MLApp = CType(Activator.CreateInstance(Type.GetTypeFromProgID("Matlab.Application")), MLApp.MLApp)
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
        command &= "]; figure1 = figure; axes1 = axes('Parent',figure1); "
        command &= "hold(axes1,'on'); scatter(x,y); "
        If xmax > xmin Then command &= "xlim(axes1,[" & xmin & " " & xmax & "]); "
        If ymax > ymin Then command &= "ylim(axes1,[" & ymin & " " & ymax & "]); "
        command &= "box(axes1,'on'); "
        If ReverseX Then command &= "set(axes1,'XDir','reverse');"
        If ReverseY Then command &= "set(axes1,'YDir','reverse');"
        matlab.Visible = 0
        matlab.Execute(command)
    End Sub
End Class
