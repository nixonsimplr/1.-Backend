
Namespace My
    
    'This class allows you to handle specific events on the settings class:
    ' The SettingChanging event is raised before a setting's value is changed.
    ' The PropertyChanged event is raised after a setting's value is changed.
    ' The SettingsLoaded event is raised after the setting values are loaded.
    ' The SettingsSaving event is raised before the setting values are saved.
    Partial Friend NotInheritable Class MySettings
        Private Function GetConnectionString() As String
            Dim ds As New DataSet
            Dim dataDirectory As String
            dataDirectory = Windows.Forms.Application.StartupPath
            ds.ReadXml(dataDirectory & "\ibiz.xml")
            Dim table As DataTable
            For Each table In ds.Tables
                Dim row As DataRow
                If table.TableName = "connectionStrings" Then
                    For Each row In table.Rows
                        Return row("connectionString").ToString()
                    Next row
                End If
            Next table
            Return ""
            ds = Nothing
        End Function
        Private Function GetSalesConnectionString() As String
            Dim ds As New DataSet
            Dim dataDirectory As String
            dataDirectory = Windows.Forms.Application.StartupPath
            ds.ReadXml(dataDirectory & "\ibiz.xml")
            Dim table As DataTable
            For Each table In ds.Tables
                Dim row As DataRow
                If table.TableName = "connectionStrings" Then
                    For Each row In table.Rows
                        Return row("connectionString").ToString()
                    Next row
                End If
            Next table
            Return ""
            ds = Nothing
        End Function
        Private Function GetOledbConnectionString() As String
            Dim ds As New DataSet
            Dim dataDirectory As String
            dataDirectory = Windows.Forms.Application.StartupPath
            ds.ReadXml(dataDirectory & "\ibiz.xml")
            Dim table As DataTable
            For Each table In ds.Tables
                Dim row As DataRow
                If table.TableName = "OledbconnectionString" Then
                    For Each row In table.Rows
                        Return row("connectionString").ToString()
                    Next row
                End If
            Next table
            Return ""
        End Function

        Private Sub MySettings_SettingsLoaded(ByVal sender As Object, ByVal e As System.Configuration.SettingsLoadedEventArgs) Handles Me.SettingsLoaded
            Me.Item("ConnectionString") = GetConnectionString()
            Me.Item("SalesConnectionString") = GetSalesConnectionString()
            Me.Item("OledbconnectionString") = GetOledbConnectionString()
        End Sub
    End Class

End Namespace
