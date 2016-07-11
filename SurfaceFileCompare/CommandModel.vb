Imports System.ComponentModel

Public Class CommandModel
    Implements System.ComponentModel.INotifyPropertyChanged
    Public Event PropertyChanged As PropertyChangedEventHandler Implements INotifyPropertyChanged.PropertyChanged
    Private _Hello As New ViewModelCommand(AddressOf cmdHello)
    Public ReadOnly Property Hello As ViewModelCommand
        Get
            Return _Hello
        End Get
    End Property
    Private Sub cmdHello(value As Object)
        MsgBox(value.ToString)
    End Sub
End Class

Public Class ViewModelCommand
    Implements ICommand
    Private _Handler As Action(Of Object)
    Public Sub New(handler As Action(Of Object))
        _Handler = handler
    End Sub
    Public Event CanExecuteChanged As EventHandler Implements ICommand.CanExecuteChanged
    Public Sub Execute(parameter As Object) Implements ICommand.Execute
        If _Handler IsNot Nothing Then _Handler(parameter)
    End Sub
    Public Function CanExecute(parameter As Object) As Boolean Implements ICommand.CanExecute
        Return _Handler IsNot Nothing
    End Function
    Public Sub CallByEvent(sender As Object, e As EventArgs)
        If CanExecute(e) Then Execute(e)
    End Sub
End Class


Public Class EventCommand
    Inherits DependencyObject

    Public Shared Function GetMouseDownCommand(ByVal element As DependencyObject) As ViewModelCommand
        If element Is Nothing Then Return Nothing
        Return element.GetValue(MouseDownCommandProperty)
    End Function

    Public Shared Sub SetMouseDownCommand(ByVal element As UIElement, ByVal value As ViewModelCommand)
        If element Is Nothing Then Return
        element.SetValue(MouseDownCommandProperty, value)
    End Sub

    Public Shared ReadOnly MouseDownCommandProperty As _
                           DependencyProperty = DependencyProperty.RegisterAttached("MouseDownCommand",
                           GetType(ViewModelCommand), GetType(EventCommand),
                           New PropertyMetadata(Nothing, New PropertyChangedCallback(AddressOf SharedMouseDownChanged)))
    Private Shared Sub SharedMouseDownChanged(d As DependencyObject, e As DependencyPropertyChangedEventArgs)
        Dim ocmd As ViewModelCommand = e.OldValue
        Dim ncmd As ViewModelCommand = e.NewValue
        If ocmd IsNot Nothing Then
            RemoveHandler DirectCast(d, UIElement).MouseDown, AddressOf ocmd.CallByEvent
        End If
        If ncmd IsNot Nothing Then
            AddHandler DirectCast(d, UIElement).MouseDown, AddressOf ncmd.CallByEvent
        End If
    End Sub
    Public Shared Function GetKeyDownCommand(ByVal element As DependencyObject) As ViewModelCommand
        If element Is Nothing Then Return Nothing
        Return element.GetValue(KeyDownCommandProperty)
    End Function

    Public Shared Sub SetKeyDownCommand(ByVal element As UIElement, ByVal value As ViewModelCommand)
        If element Is Nothing Then Return
        element.SetValue(KeyDownCommandProperty, value)
    End Sub

    Public Shared ReadOnly KeyDownCommandProperty As _
                           DependencyProperty = DependencyProperty.RegisterAttached("KeyDownCommand",
                           GetType(ViewModelCommand), GetType(EventCommand),
                           New PropertyMetadata(Nothing, New PropertyChangedCallback(AddressOf SharedKeyDownChanged)))
    Private Shared Sub SharedKeyDownChanged(d As DependencyObject, e As DependencyPropertyChangedEventArgs)
        Dim ocmd As ViewModelCommand = e.OldValue
        Dim ncmd As ViewModelCommand = e.NewValue
        If ocmd IsNot Nothing Then
            RemoveHandler DirectCast(d, UIElement).KeyDown, AddressOf ocmd.CallByEvent
        End If
        If ncmd IsNot Nothing Then
            AddHandler DirectCast(d, UIElement).KeyDown, AddressOf ncmd.CallByEvent
        End If
    End Sub


    Public Shared Function GetSelectionChangedCommand(ByVal element As Primitives.Selector) As ViewModelCommand
        If element Is Nothing Then Return Nothing
        Return element.GetValue(SelectionChangedCommandProperty)
    End Function

    Public Shared Sub SetSelectionChangedCommand(ByVal element As Primitives.Selector, ByVal value As ViewModelCommand)
        If element Is Nothing Then Return
        element.SetValue(SelectionChangedCommandProperty, value)
    End Sub

    Public Shared ReadOnly SelectionChangedCommandProperty As _
                           DependencyProperty = DependencyProperty.RegisterAttached("SelectionChangedCommand",
                           GetType(ViewModelCommand), GetType(EventCommand),
                           New PropertyMetadata(Nothing, New PropertyChangedCallback(AddressOf SharedSelectionChangedChanged)))
    Private Shared Sub SharedSelectionChangedChanged(d As DependencyObject, e As DependencyPropertyChangedEventArgs)
        Dim ocmd As ViewModelCommand = e.OldValue
        Dim ncmd As ViewModelCommand = e.NewValue
        If ocmd IsNot Nothing Then
            RemoveHandler DirectCast(d, Primitives.Selector).SelectionChanged, AddressOf ocmd.CallByEvent
        End If
        If ncmd IsNot Nothing Then
            AddHandler DirectCast(d, Primitives.Selector).SelectionChanged, AddressOf ncmd.CallByEvent
        End If
    End Sub

End Class

Public Class WindowCommand
    Public Shared Function GetLoadedCommand(ByVal element As Window) As ViewModelCommand
        If element Is Nothing Then Return Nothing
        Return element.GetValue(LoadedCommandProperty)
    End Function

    Public Shared Sub SetLoadedCommand(ByVal element As Window, ByVal value As ViewModelCommand)
        If element Is Nothing Then Return
        element.SetValue(LoadedCommandProperty, value)
    End Sub

    Public Shared ReadOnly LoadedCommandProperty As _
                           DependencyProperty = DependencyProperty.RegisterAttached("LoadedCommand",
                           GetType(ViewModelCommand), GetType(WindowCommand),
                           New PropertyMetadata(Nothing, New PropertyChangedCallback(AddressOf SharedLoadedChanged)))
    Private Shared Sub SharedLoadedChanged(d As DependencyObject, e As DependencyPropertyChangedEventArgs)
        If System.ComponentModel.DesignerProperties.GetIsInDesignMode(d) Then Return
        Dim ocmd As ViewModelCommand = e.OldValue
        Dim ncmd As ViewModelCommand = e.NewValue
        If ocmd IsNot Nothing Then
            RemoveHandler DirectCast(d, Window).Loaded, AddressOf ocmd.CallByEvent
        End If
        If ncmd IsNot Nothing Then
            AddHandler DirectCast(d, Window).Loaded, AddressOf ncmd.CallByEvent
        End If
    End Sub

    Public Shared Function GetClosingCommand(ByVal element As Window) As ViewModelCommand
        If element Is Nothing Then Return Nothing
        Return element.GetValue(ClosingCommandProperty)
    End Function

    Public Shared Sub SetClosingCommand(ByVal element As Window, ByVal value As ViewModelCommand)
        If element Is Nothing Then Return
        element.SetValue(ClosingCommandProperty, value)
    End Sub

    Public Shared ReadOnly ClosingCommandProperty As _
                           DependencyProperty = DependencyProperty.RegisterAttached("ClosingCommand",
                           GetType(ViewModelCommand), GetType(WindowCommand),
                           New PropertyMetadata(Nothing, New PropertyChangedCallback(AddressOf SharedClosingChanged)))
    Private Shared Sub SharedClosingChanged(d As DependencyObject, e As DependencyPropertyChangedEventArgs)
        If System.ComponentModel.DesignerProperties.GetIsInDesignMode(d) Then Return
        Dim ocmd As ViewModelCommand = e.OldValue
        Dim ncmd As ViewModelCommand = e.NewValue
        If ocmd IsNot Nothing Then
            RemoveHandler DirectCast(d, Window).Closing, AddressOf ocmd.CallByEvent
        End If
        If ncmd IsNot Nothing Then
            AddHandler DirectCast(d, Window).Closing, AddressOf ncmd.CallByEvent
        End If
    End Sub
End Class