Public Class MyCalendar

    Public Event DateSelected(d As Date)

    Private Sub Calendar_SelectedDatesChanged(sender As Object, e As SelectionChangedEventArgs)
        Dim cal As Calendar = sender

        RaiseEvent DateSelected(cal.SelectedDate)
        Me.Close()
    End Sub

    Private Sub MetroWindow_Loaded(sender As Object, e As RoutedEventArgs)

    End Sub
End Class
