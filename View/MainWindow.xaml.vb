Imports System
Imports System.Windows
Imports System.Windows.Threading
Imports System.ComponentModel

Public Class NotificationWindow

    'Private outlookClient As New Tasks.OutlookSync
    Private tasks As New List(Of DummyTask)
    Private tasksXdoc As XDocument
    Private curDueDate As Date = DateAdd(DateInterval.Day, 2, Now)

    Private Sub RefreshList()
        Dim qry = (From t In tasks Order By t.DueDate Ascending Select t)
        lstTasks.ItemsSource = qry

        Dim sum As Double = 0

        For Each t In qry
            sum += t.AmountOwed
        Next

        lblDueAmount.Text = "£" & Format(sum, "0.##")
    End Sub

    Private Sub Window_Loaded(sender As Object, e As RoutedEventArgs)
        tasksXdoc = XDocument.Load("Tasks.xml")

        lblTotalPaid.Text = "£" & tasksXdoc.Root.@TotalPaid

        For Each t In tasksXdoc.Root.<Task>
            Dim task As New DummyTask
            task.MyXDoc = tasksXdoc
            task.PopulateFromXML(t)
            tasks.Add(task)
        Next

        RefreshList()
    End Sub

    Private Sub btnAdd_Click(sender As Object, e As RoutedEventArgs)
        If String.IsNullOrWhiteSpace(txtTask.Text) Then Return

        Dim t As DummyTask = New DummyTask() With {.Subject = txtTask.Text, .DueDate = curDueDate, .MyXDoc = tasksXdoc}

        AddTaskItem(t)
    End Sub

    Private Sub btnDelete_Click(sender As Object, e As RoutedEventArgs)
        Dim btn As Button = sender
        Dim d As DummyTask = btn.DataContext

        RemoveTaskItem(d)
    End Sub

    Private Sub AddTaskItem(t As DummyTask)
        tasks.Add(t)

        tasksXdoc.Root.Add(t.GetXML)
        tasksXdoc.Save("Tasks.xml")

        RefreshList()
    End Sub

    Private Sub RemoveTaskItem(d As DummyTask)
        If tasksXdoc.Root.<Task>.Any(Function(c) c.<Id>.Value = d.Id) Then
            tasksXdoc...<Task>.First(Function(c) c.<Id>.Value = d.Id).Remove()
            tasksXdoc.Save("Tasks.xml")
        End If

        tasks.Remove(d)

        RefreshList()
    End Sub

    Private Sub btnCalendar_Click(sender As Object, e As RoutedEventArgs)
        Dim cal As New MyCalendar

        AddHandler cal.DateSelected, Sub(d)
                                         curDueDate = d
                                         btnCalendar.ToolTip = New ToolTip() With {.Content = curDueDate.ToShortDateString}
                                     End Sub

        cal.WindowStartupLocation = Windows.WindowStartupLocation.CenterScreen
        cal.SaveWindowPosition = True
        cal.ShowDialog()
    End Sub

    Private Sub btnSettings_Click(sender As Object, e As RoutedEventArgs)
        Dim winSettings As New Settings
        winSettings.ShowDialog()
    End Sub

    Private Sub btnPiggy_Click(sender As Object, e As RoutedEventArgs)
        Dim finishedTasks = (From t In tasks Where t.IsFinished = True And t.IsOverdue = True Select t).ToList
        Dim amt As Double = Double.Parse(lblDueAmount.Text.Replace("£", ""))
        Dim amt2 As Double = Double.Parse(lblTotalPaid.Text.Replace("£", ""))
        Dim sum As Double = amt2 + amt

        For Each t In finishedTasks
            RemoveTaskItem(t)
        Next

        MessageBox.Show("Thank you for donating " & lblDueAmount.Text & "!")

        lblDueAmount.Text = "£0"
        lblTotalPaid.Text = "£" & sum
        tasksXdoc.Root.@TotalPaid = sum

        tasksXdoc.Save("Tasks.xml")
    End Sub

    Private Sub btnClear_Click(sender As Object, e As RoutedEventArgs)
        Dim finishedTasks = (From t In tasks Where t.IsFinished = True Select t).ToList

        For Each t In finishedTasks
            RemoveTaskItem(t)
        Next

        MessageBox.Show("You have cleared finished jobs!")
    End Sub
End Class


Public Class DummyTask

    Private _isfinished As Boolean = False

    Public Property MyXDoc As XDocument

    Public Property Id As String = Guid.NewGuid.ToString
    Public Property Subject As String
    Public Property Body As String
    Public Property IsFinished As Boolean
        Get
            Return _isfinished
        End Get
        Set(value As Boolean)
            _isfinished = value
            If MyXDoc.Root.<Task>.Any(Function(c) c.<Id>.Value = Id) Then
                MyXDoc.Root.<Task>.First(Function(c) c.<Id>.Value = Id).<IsFinished>.Value = _isfinished

                MyXDoc.Save("Tasks.xml")
            End If
        End Set
    End Property
    Public Property DueDate As Date

    Public ReadOnly Property IsOverdue As Boolean
        Get
            Return DueDate < Now
        End Get
    End Property
    Public ReadOnly Property AmountOwed As Double
        Get
            Dim val As Double = DateDiff(DateInterval.Day, DueDate, Now) * Double.Parse(System.Configuration.ConfigurationManager.AppSettings("donation"))

            If IsFinished Then Return 0
            If val < 0 Then Return 0

            Return val
        End Get
    End Property

    Public ReadOnly Property BGColor As Brush
        Get
            If IsOverdue Then
                Return Brushes.DarkRed
            End If

            Return Brushes.Black
        End Get
    End Property

    Public ReadOnly Property IsAlertVisible As Visibility
        Get
            If IsOverdue Then
                Return Visibility.Visible
            End If

            Return Visibility.Collapsed
        End Get
    End Property

    Public Sub PopulateFromXML(xel As XElement)
        Id = xel.<Id>.Value
        IsFinished = xel.<IsFinished>.Value
        Subject = xel.<Subject>.Value
        Body = xel.<Body>.Value
        DueDate = xel.<DueDate>.Value
    End Sub

    Public ReadOnly Property GetXML As XElement
        Get
            Return <Task>
                       <Id><%= Id %></Id>
                       <IsFinished><%= IsFinished %></IsFinished>
                       <Subject><%= Subject %></Subject>
                       <Body><%= Body %></Body>
                       <DueDate><%= DueDate %></DueDate>
                   </Task>
        End Get
    End Property
End Class

