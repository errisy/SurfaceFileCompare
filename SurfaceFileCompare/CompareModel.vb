Imports System.ComponentModel

Public Class CompareModel
    Implements System.ComponentModel.INotifyPropertyChanged

    Public FileA As String
    Public FileB As String
    Shared Pattern As New System.Text.RegularExpressions.Regex("[A-Z]\d+[\+\-]")
    Public Event PropertyChanged As PropertyChangedEventHandler Implements INotifyPropertyChanged.PropertyChanged
    Public ReadOnly Property OpenFileA As New ViewModelCommand(AddressOf cmdOpenFileA)
    Private Sub cmdOpenFileA(value As Object)
        Dim ofd As New Microsoft.Win32.OpenFileDialog With {.Filter = "Text File|*.txt"}
        If ofd.ShowDialog Then
            FileA = IO.File.ReadAllText(ofd.FileName)
            _DocumentA = New FlowDocument
            _DocumentA.Blocks.Add((Function()
                                       Dim main = New Paragraph
                                       For Each m As System.Text.RegularExpressions.Match In Pattern.Matches(FileA)
                                           main.Inlines.Add(New Run With {.Text = m.Value})
                                       Next
                                       Return main
                                   End Function).Invoke())
            RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs("DocumentA"))
            Compare()
        End If
    End Sub
    Public ReadOnly Property OpenFileB As New ViewModelCommand(AddressOf cmdOpenFileB)
    Private Sub cmdOpenFileB(value As Object)
        Dim ofd As New Microsoft.Win32.OpenFileDialog With {.Filter = "Text File|*.txt"}
        If ofd.ShowDialog Then
            FileB = IO.File.ReadAllText(ofd.FileName)
            _DocumentB = New FlowDocument
            _DocumentB.Blocks.Add((Function()
                                       Dim main = New Paragraph
                                       For Each m As System.Text.RegularExpressions.Match In Pattern.Matches(FileB)
                                           main.Inlines.Add(New Run With {.Text = m.Value})
                                       Next
                                       Return main
                                   End Function).Invoke())
            RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs("DocumentB"))
            Compare()
        End If
    End Sub
    Public Sub Compare()
        If _DocumentA Is Nothing Then Return
        If _DocumentB Is Nothing Then Return
        Try
            Dim paraA As Paragraph = _DocumentA.Blocks(0)
            Dim paraB As Paragraph = _DocumentB.Blocks(0)

            Dim i As Integer = 0
            While i < paraA.Inlines.Count And i < paraB.Inlines.Count
                Dim runA As Run = paraA.Inlines(i)
                Dim runB As Run = paraB.Inlines(i)
                If runA.Text = runB.Text Then
                    runA.Background = Brushes.LightYellow
                    runA.FontWeight = FontWeights.Normal
                    runB.Background = Brushes.LightYellow
                    runB.FontWeight = FontWeights.Normal

                Else
                    runA.Background = Brushes.Red
                    runA.FontWeight = FontWeights.Bold
                    runB.Background = Brushes.Red
                    runB.FontWeight = FontWeights.Bold
                End If
                i += 1
            End While

        Catch ex As Exception

        End Try


    End Sub
    Private _DocumentA As FlowDocument
    Public ReadOnly Property DocumentA As FlowDocument
        Get
            Return _DocumentA
        End Get
    End Property
    Private _DocumentB As FlowDocument
    Public ReadOnly Property DocumentB As FlowDocument
        Get
            Return _DocumentB
        End Get
    End Property
End Class
