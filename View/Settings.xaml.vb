Imports System.Configuration

Public Class Settings
    Private Sub Save_Click(sender As Object, e As RoutedEventArgs)
        ConfigurationManager.AppSettings("donation") = sldRate.Value
        Me.Close()
    End Sub

    Private Sub MetroWindow_Loaded(sender As Object, e As RoutedEventArgs)
        sldRate.Value = ConfigurationManager.AppSettings("donation")
        Dim rdBox As RadioButton = Me.FindName("chk" & ConfigurationManager.AppSettings("charityOrganization"))

        If rdBox IsNot Nothing Then rdBox.IsChecked = True
        '
    End Sub

    Private Sub chkUnicef_Checked(sender As Object, e As RoutedEventArgs)
        ConfigurationManager.AppSettings("charityOrganization") = CType(sender, RadioButton).Name.Replace("chk", "")
    End Sub
End Class
