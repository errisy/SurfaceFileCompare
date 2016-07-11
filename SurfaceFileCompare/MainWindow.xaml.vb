Class MainWindow
    Protected Overrides Sub OnInitialized(e As EventArgs)
        DataContext = New CompareModel
        MyBase.OnInitialized(e)
    End Sub
End Class
