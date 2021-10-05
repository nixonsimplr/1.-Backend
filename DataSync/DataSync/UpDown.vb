Imports OpenNETCF.Desktop.Communication
Imports System.IO
Imports System.Data.SqlClient

Public Class UpDown
    'Sync
    Private WithEvents m_rapi As RAPI
    Private WithEvents m_activesync As ActiveSync
    Public Status As String = ""
    Private sFolder As String

    Public Sub Connect(ByVal ConnectionString As String)
        m_rapi = New RAPI
        m_activesync = m_rapi.ActiveSync
        ConnectDB(ConnectionString)
    End Sub

    Public Sub Disconnect()
        DisconnectDB()
    End Sub

    Private Function GetPDAFolder(ByVal FileName As String) As String
        Dim objStreamReader As StreamReader
        Dim strLine As String
        Dim sDownloadFolder As String = ""
        objStreamReader = New StreamReader(FileName)
        strLine = objStreamReader.ReadLine
        sDownloadFolder = strLine
        objStreamReader.Close()
        Return sDownloadFolder
    End Function

    Public Function StartUpload() As String
        If Not m_rapi.Connected Then
            Status = "Connecting..."
            m_rapi.Connect()
        End If
        Try
            Dim path As String
            path = System.AppDomain.CurrentDomain.BaseDirectory
            If Directory.Exists(path & "\download") = False Then
                Directory.CreateDirectory(path & "\download")
            End If
            If File.Exists(path & "\download\filedet.txt") = True Then
                File.Delete(path & "\download\filedet.txt")
            End If
            Status = "Data processing..."
            Try
                m_rapi.CopyFileFromDevice(path & "\download\filedet.txt", "\filedet.txt", True)
            Catch ex As Exception
                m_rapi.Disconnect()
                Return "Install the application on the PDA and run once before try to download data"
            End Try
            sFolder = GetPDAFolder(path & "\download\filedet.txt")
            If Not Directory.Exists(path & "\upload") Then
                Directory.CreateDirectory(path & "\upload")
            End If
        Catch ex As Exception
            m_rapi.Disconnect()
            Throw New ApplicationException(ex.Message)
            Return ex.Message
        End Try
        Return "Completed"
    End Function

    Public Sub UploadFile(ByVal SQL As String, ByVal TableName As String)
        Dim path As String
        path = System.AppDomain.CurrentDomain.BaseDirectory
        Dim adp As SqlDataAdapter
        adp = GetDataAdapter(SQL)
        Dim ds As DataSet = Nothing
        If File.Exists(path & "\upload\" & TableName & ".xml") = True Then
            File.Delete(path & "\upload\" & TableName & ".xml")
        End If
        If File.Exists(path & "\upload\" & TableName & ".xsd") = True Then
            File.Delete(path & "\upload\" & TableName & ".xsd")
        End If
        Try
            ds = New DataSet(TableName)
            adp.Fill(ds)
            ds.WriteXml(path & "\upload\" & TableName & ".xml")
            ds.WriteXmlSchema(path & "\upload\" & TableName & ".xsd")
        Catch ex As Exception
            MsgBox("Error: " & ex.Message)
        End Try
        If m_rapi.DeviceFileExists(sFolder & "\In\" & TableName & ".xml") Then m_rapi.DeleteDeviceFile(sFolder & "\In\" & TableName & ".xml")
        If m_rapi.DeviceFileExists(sFolder & "\In\" & TableName & ".xsd") Then m_rapi.DeleteDeviceFile(sFolder & "\In\" & TableName & ".xsd")
        m_rapi.CopyFileToDevice(path & "upload\" & TableName & ".xml", sFolder & "\In\" & TableName & ".xml")
        m_rapi.CopyFileToDevice(path & "upload\" & TableName & ".xsd", sFolder & "\In\" & TableName & ".xsd")
        Status = "Completed"

    End Sub

    Public Sub EndUpload()
        m_rapi.Disconnect()
    End Sub

    Private Function CheckAvailable(ByVal SQL As String) As Boolean
        Dim dr As SqlClient.SqlDataReader
        Dim bCheck As Boolean = False
        dr = ReadRecord(Sql)
        bCheck = dr.Read
        dr.Close()
        Return bCheck
    End Function

End Class
