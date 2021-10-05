Imports OpenNETCF.Desktop.Communication
Imports System.IO
Imports System.Data.SqlClient

Public Class UpDown
    'Sync
    Private WithEvents m_rapi As RAPI
    Private WithEvents m_activesync As ActiveSync
    Public Status As String = ""
    Private sFolder As String
    Public arrDownload() As String
    Public arrDownloadImages() As String
    Dim path As String

    Public Function Connect(ByVal ConnectionString As String) As Boolean
        m_rapi = New RAPI
        m_activesync = m_rapi.ActiveSync
        ConnectDB(ConnectionString)
        If m_rapi.DevicePresent = False Then
            Return False
        End If
        Return True
    End Function

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

    Public Function fileNameWithoutThePath(ByVal b As String) As String
        Dim j As Int16
        j = Convert.ToInt16(b.LastIndexOf("\"))
        Return b.Substring(j + 1)
    End Function

    Public Sub UploadItemPhoto(ByVal FromPath As String, ByVal FileName As String)
        Dim path As String
        path = System.AppDomain.CurrentDomain.BaseDirectory
        Dim ds As DataSet = Nothing

        If Directory.Exists(path & "Upload\") = False Then
            Directory.CreateDirectory(path & "Upload\")
        End If

        If Directory.Exists(path & "Upload\Photo\") = False Then
            Directory.CreateDirectory(path & "Upload\Photo\")
        End If

        If File.Exists(path & "Upload\Photo\" & FileName) = True Then
            File.Delete(path & "Upload\Photo\" & FileName)
        End If

        Try
            File.Copy(FromPath & "\" & FileName, path & "upload\Photo\" & FileName)
        Catch ex As Exception
            MsgBox("Error: " & ex.Message)
        End Try


        If m_rapi.DeviceFileExists("\SD Card\Item Photo\" & FileName) Then m_rapi.DeleteDeviceFile("\SD Card\Item Photo\" & FileName)

        m_rapi.CopyFileToDevice(path & "upload\Photo\" & FileName, "\SD Card\Item Photo\" & FileName)

        Status = "Completed"
    End Sub

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
            If ds.Tables.Count > 0 Then
                ds.Tables(0).Dispose()
            End If
            ds.Dispose()
        Catch ex As Exception
            ds.Dispose()
            MsgBox("Error: " & ex.Message)
        End Try
        If m_rapi.DeviceFileExists(sFolder & "\In\" & TableName & ".xml") Then m_rapi.DeleteDeviceFile(sFolder & "\In\" & TableName & ".xml")
        If m_rapi.DeviceFileExists(sFolder & "\In\" & TableName & ".xsd") Then m_rapi.DeleteDeviceFile(sFolder & "\In\" & TableName & ".xsd")
        m_rapi.CopyFileToDevice(path & "upload\" & TableName & ".xml", sFolder & "\In\" & TableName & ".xml")
        m_rapi.CopyFileToDevice(path & "upload\" & TableName & ".xsd", sFolder & "\In\" & TableName & ".xsd")
        Status = "Completed"
    End Sub

    Public Sub EndUpload()
        Dim path As String
        Dim sNewDir As String
        path = System.AppDomain.CurrentDomain.BaseDirectory
        sNewDir = path & "\Upload\Backup\" & Format(Date.Now, "ddMMyyyy hhmmss")
        If Directory.Exists(sNewDir) = False Then
            Directory.CreateDirectory(sNewDir)
        End If
        Dim fname() As String = Directory.GetFiles(path & "\upload")
        For i As Integer = 0 To UBound(fname)
            File.Move(fname(i), sNewDir & "\" & fileNameWithoutThePath(fname(i)))
        Next
        m_rapi.Disconnect()
    End Sub

    Public Sub UploadFileBarcode(ByVal SQL As String, ByVal TableName As String)
        Dim path As String
        path = System.AppDomain.CurrentDomain.BaseDirectory
        Dim adp As SqlDataAdapter
        adp = GetDataAdapter(SQL)
        Dim ds As DataSet = Nothing
        If File.Exists(path & "\upload\Barcode\" & TableName & ".xml") = True Then
            File.Delete(path & "\upload\Barcode\" & TableName & ".xml")
        End If
        If File.Exists(path & "\upload\Barcode\" & TableName & ".xsd") = True Then
            File.Delete(path & "\upload\Barcode\" & TableName & ".xsd")
        End If
        Try

            ds = New DataSet(TableName)
            adp.Fill(ds)
            ds.WriteXml(path & "\upload\Barcode\" & TableName & ".xml")
            ds.WriteXmlSchema(path & "\upload\Barcode\" & TableName & ".xsd")
        Catch ex As Exception
            MsgBox("Error: " & ex.Message)
        End Try
        If m_rapi.DeviceFileExists(sFolder & "\In\Barcode\" & TableName & ".xml") Then m_rapi.DeleteDeviceFile(sFolder & "\In\Barcode\" & TableName & ".xml")
        If m_rapi.DeviceFileExists(sFolder & "\In\Barcode\" & TableName & ".xsd") Then m_rapi.DeleteDeviceFile(sFolder & "\In\Barcode\" & TableName & ".xsd")
        m_rapi.CopyFileToDevice(path & "upload\Barcode\" & TableName & ".xml", sFolder & "\In\Barcode\" & TableName & ".xml")
        m_rapi.CopyFileToDevice(path & "upload\Barcode\" & TableName & ".xsd", sFolder & "\In\Barcode\" & TableName & ".xsd")
        Status = "Completed"
    End Sub

    Private Function CheckAvailable(ByVal SQL As String) As Boolean
        Dim dr As SqlClient.SqlDataReader
        Dim bCheck As Boolean = False
        dr = ReadRecord(Sql)
        bCheck = dr.Read
        dr.Close()
        Return bCheck
    End Function

    Public Function DownloadImages() As String
        If Not m_rapi.Connected Then
            Status = "Connecting..."
            m_rapi.Connect()
            Status = "Connected"
        End If

        Dim path As String
        path = System.AppDomain.CurrentDomain.BaseDirectory
        If Not Directory.Exists(path & "download") Then
            Directory.CreateDirectory(path & "download")
        End If
        If File.Exists(path & "download\filedet.txt") = True Then
            File.Delete(path & "download\filedet.txt")
        End If

        Dim sDownloadFolder As String
        sDownloadFolder = "\SD Card\Signature\"
        Status = "Data processing..."

        For Each sImgName As String In arrDownloadImages
            If File.Exists(path & "download\" & sImgName) = False Then
                If m_rapi.DeviceFileExists(sDownloadFolder & sImgName) Then
                    m_rapi.CopyFileFromDevice(path & "download\" & sImgName, sDownloadFolder & sImgName, True)
                End If
                If m_rapi.DeviceFileExists(sDownloadFolder & sImgName) Then
                    m_rapi.DeleteDeviceFile(sDownloadFolder & sImgName)
                End If
            End If
        Next

        Return "Completed"
    End Function

    Public Function DownloadOrderSign(ByVal sOrderNo As String) As String
        If Not m_rapi.Connected Then
            Status = "Connecting..."
            m_rapi.Connect()
            Status = "Connected"
        End If

        Dim path As String
        path = System.AppDomain.CurrentDomain.BaseDirectory
        If Not Directory.Exists(path & "download") Then
            Directory.CreateDirectory(path & "download")
        End If
        If File.Exists(path & "download\filedet.txt") = True Then
            File.Delete(path & "download\filedet.txt")
        End If

        Dim sDownloadFolder As String
        sDownloadFolder = "\SD Card\Signature\"
        Status = "Data processing..."


        If File.Exists(path & "download\" & sOrderNo & ".bmp") = False Then
            If m_rapi.DeviceFileExists(sDownloadFolder & sOrderNo & ".bmp") Then
                m_rapi.CopyFileFromDevice(path & "download\" & sOrderNo & ".bmp", sDownloadFolder & sOrderNo & ".bmp", True)
            End If
        End If

        If File.Exists(path & "Signature\" & sOrderNo & ".bmp") = False Then
            If m_rapi.DeviceFileExists(sDownloadFolder & sOrderNo & ".bmp") Then
                m_rapi.CopyFileFromDevice(path & "Signature\" & sOrderNo & ".bmp", sDownloadFolder & sOrderNo & ".bmp", True)
            End If
        End If



        Return "Completed"
    End Function

    Public Function Download() As String
        If Not m_rapi.Connected Then
            Status = "Connecting..."
            m_rapi.Connect()
            Status = "Connected"
        End If
        'Try
        Dim path As String
        path = System.AppDomain.CurrentDomain.BaseDirectory
        If Not Directory.Exists(path & "download") Then
            Directory.CreateDirectory(path & "download")
        End If
        If File.Exists(path & "download\filedet.txt") = True Then
            File.Delete(path & "download\filedet.txt")
        End If
        Status = "Downloading..."
        Try
            m_rapi.CopyFileFromDevice(path & "download\filedet.txt", "\filedet.txt", True)
        Catch ex As Exception
            m_rapi.Disconnect()
            Return "No data to download"
        End Try
        Dim sDownloadFolder As String
        sDownloadFolder = GetPDAFolder(path & "download\filedet.txt")
        Status = "Data processing..."
        For Each sTablename As String In arrDownload
            'Try
            If m_rapi.DeviceFileExists(sDownloadFolder & "\Out\" & sTablename & ".xml") Then
                m_rapi.CopyFileFromDevice(path & "download\" & sTablename & ".xml", sDownloadFolder & "\Out\" & sTablename & ".xml", True)
                m_rapi.CopyFileFromDevice(path & "download\" & sTablename & ".xsd", sDownloadFolder & "\Out\" & sTablename & ".xsd", True)
                If sTablename.ToLower <> "system" Then
                    m_rapi.DeleteDeviceFile(sDownloadFolder & "\Out\" & sTablename & ".xml")
                    m_rapi.DeleteDeviceFile(sDownloadFolder & "\Out\" & sTablename & ".xsd")
                End If
                'DownloadTable(sTablename, path & "download\" & sTablename & ".xml", path & "download\" & sTablename & ".xsd")
                'Catch ex As Exception
                'm_rapi.Disconnect()
                'Return "No data to download"
                'End Try
            End If
        Next

        'm_rapi.MoveDeviceFile(path & "download\*.bmp", sDownloadFolder & "\Out\*.bmp")
        'Status = "Download completed"
        'Catch ex As Exception
        'Throw New ApplicationException(ex.Message)
        'End Try

        Return "Completed"
    End Function

    Public Function CopyFile(ByVal FileName As String, ByVal DestinationFileName As String) As Boolean
        If m_rapi.DeviceFileExists(FileName) Then
            m_rapi.CopyFileFromDevice(DestinationFileName, FileName, True)
            Return True
        Else
            Return False
        End If
    End Function

    Public Function DeleteFile(ByVal FileName As String) As Boolean
        Try
            If m_rapi.DeviceFileExists(FileName) Then
                m_rapi.DeleteDeviceFile(FileName)
                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            Return False
        End Try
    End Function

    Public Function GetMDTFile() As Boolean
        If m_rapi.DeviceFileExists(sFolder & "\Out\system.xml") Then
            m_rapi.CopyFileFromDevice(path & "download\system.xml", sFolder & "\Out\system.xml", True)
            m_rapi.CopyFileFromDevice(path & "download\system.xsd", sFolder & "\Out\system.xsd", True)
            Return True
        Else
            Return False
        End If
    End Function


End Class
