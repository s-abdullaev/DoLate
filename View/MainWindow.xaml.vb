Imports System
Imports System.Windows
Imports System.Windows.Threading
Imports System.ComponentModel

Public Class NotificationWindow
    Private tasks As List(Of DummyTask)


    Private Sub Window_Loaded(sender As Object, e As RoutedEventArgs)
        tasks = DummyTasks.GetTasks

        lstTasks.ItemsSource = tasks
    End Sub

    Private Sub btnAdd_Click(sender As Object, e As RoutedEventArgs)
        Dim t As DummyTask = New DummyTask() With {.Title = txtTask.Text}

        tasks.Add(t)
        lstTasks.ItemsSource = tasks

    End Sub

    Private Sub btnDelete_Click(sender As Object, e As RoutedEventArgs)
        Dim btn As Button = sender
        Dim delTask As DummyTask = tasks.FirstOrDefault(Function(c) c.Id = btn.Tag)

        If delTask IsNot Nothing Then tasks.Remove(delTask)

        lstTasks.ItemsSource = tasks
    End Sub
End Class


Public Class DummyTask
    Public Property Id As Integer
    Public Property Title As String
    Public Property IsFinished As Boolean
End Class

Public Class DummyTasks
    Public Shared Function GetTasks() As List(Of DummyTask)
        Dim tasks As New List(Of DummyTask)

        tasks.Add(New DummyTask With {.Id = 1, .Title = "hey jude", .IsFinished = True})
        tasks.Add(New DummyTask With {.Id = 2, .Title = "don't make me sad", .IsFinished = False})
        tasks.Add(New DummyTask With {.Id = 3, .Title = "hey jude, jude jude", .IsFinished = False})
        tasks.Add(New DummyTask With {.Id = 4, .Title = "hey jude", .IsFinished = True})
        tasks.Add(New DummyTask With {.Id = 5, .Title = "don't make me sad", .IsFinished = False})
        tasks.Add(New DummyTask With {.Id = 6, .Title = "hey jude, jude jude", .IsFinished = False})

        Return tasks
    End Function
End Class