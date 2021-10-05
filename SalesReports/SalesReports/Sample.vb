Public Class Sample

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        ConnectDB()
        ExecuteReport("Select * from Customer", "samp")
        DisconnectDB()
    End Sub
End Class