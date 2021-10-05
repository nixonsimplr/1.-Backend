Imports System.Data.SqlClient

Module Execute
    Public Sub ExecuteReport(ByVal strSql As String, ByVal RptName As String)
        Dim DA As New SqlDataAdapter(strSql, My.Settings.ConnectionString)
        Dim DS As New DataSet
        DA.Fill(DS)
        Dim ConInfo As New CrystalDecisions.Shared.TableLogOnInfo
        ConInfo.ConnectionInfo.IntegratedSecurity = True
        'ConInfo.ConnectionInfo.UserID = objDataBase.UserName
        'ConInfo.ConnectionInfo.Password = objDataBase.Password
        ConInfo.ConnectionInfo.ServerName = ".\SQLEXPRESS"
        ConInfo.ConnectionInfo.DatabaseName = "Sales"

        Dim strReportPath As String = Application.StartupPath & "\" & RptName & ".rpt"
        If Not IO.File.Exists(strReportPath) Then
            Throw (New Exception("Unable to locate report file:" & vbCrLf & strReportPath))
        End If
        Dim rptDocument As New CrystalDecisions.CrystalReports.Engine.ReportDocument
        rptDocument.Load(strReportPath)
        rptDocument.SetDataSource(DS.Tables(0))
        Dim frm As New ViewReport1

        frm.crvReport.ShowRefreshButton = False
        frm.crvReport.ShowCloseButton = False
        frm.crvReport.ShowGroupTreeButton = False
        frm.crvReport.ReportSource = rptDocument
        frm.Show()
    End Sub
    'Friend Function ViewReport(ByVal sReportName As String, Optional ByVal sSelectionFormula As String = "", Optional ByVal param As String = "") As Boolean

    '    Dim intCounter As Integer
    '    Dim intCounter1 As Integer
    '    'Dim strTableName As String
    '    'Dim objReportsParameters As frmReportsParameters

    '    Dim objReport As New CrystalDecisions.CrystalReports.Engine.ReportDocument
    '    'Dim mySection As CrystalDecisions.CrystalReports.Engine.Section
    '    'Dim mySections As CrystalDecisions.CrystalReports.Engine.Sections

    '    Dim ConInfo As New CrystalDecisions.Shared.TableLogOnInfo

    '    Dim paraValue As New CrystalDecisions.Shared.ParameterDiscreteValue
    '    Dim currValue As CrystalDecisions.Shared.ParameterValues
    '    Dim mySubReportObject As CrystalDecisions.CrystalReports.Engine.SubreportObject
    '    Dim mySubRepDoc As New CrystalDecisions.CrystalReports.Engine.ReportDocument

    '    Dim strParamenters As String = ""
    '    Dim strParValPair() As String
    '    Dim strVal() As String
    '    Dim sFileName As String = ""
    '    Dim index As Integer

    '    Try


    '        'sFileName = DownloadReport(sReportName, m_strReportDir)

    '        objReport.Load(sFileName)

    '        intCounter = objReport.DataDefinition.ParameterFields.Count
    '        If intCounter = 1 Then
    '            If InStr(objReport.DataDefinition.ParameterFields(0).ParameterFieldName, ".", CompareMethod.Text) > 0 Then
    '                intCounter = 0
    '            End If
    '        End If


    '        If intCounter > 0 And Trim(param) <> "" Then

    '            strParValPair = strParamenters.Split("&")
    '            For index = 0 To UBound(strParValPair)
    '                If InStr(strParValPair(index), "=") > 0 Then
    '                    strVal = strParValPair(index).Split("=")
    '                    paraValue.Value = strVal(1)
    '                    currValue = objReport.DataDefinition.ParameterFields(strVal(0)).CurrentValues
    '                    currValue.Add(paraValue)
    '                    objReport.DataDefinition.ParameterFields(strVal(0)).ApplyCurrentValues(currValue)
    '                End If
    '            Next
    '        End If


    '        ConInfo.ConnectionInfo.IntegratedSecurity = True
    '        'ConInfo.ConnectionInfo.UserID = objDataBase.UserName
    '        'ConInfo.ConnectionInfo.Password = objDataBase.Password
    '        ConInfo.ConnectionInfo.ServerName = "ibiz-IBIZ-JAGADISH\SQLEXPRESS"
    '        ConInfo.ConnectionInfo.DatabaseName = "Sales"

    '        For intCounter = 0 To objReport.Database.Tables.Count - 1
    '            objReport.Database.Tables(intCounter).ApplyLogOnInfo(ConInfo)
    '        Next



    '        For index = 0 To objReport.ReportDefinition.Sections.Count - 1
    '            For intCounter = 0 To objReport.ReportDefinition.Sections(index).ReportObjects.Count - 1
    '                With objReport.ReportDefinition.Sections(index)
    '                    If .ReportObjects(intCounter).Kind = CrystalDecisions.Shared.ReportObjectKind.SubreportObject Then
    '                        mySubReportObject = CType(.ReportObjects(intCounter), CrystalDecisions.CrystalReports.Engine.SubreportObject)
    '                        mySubRepDoc = mySubReportObject.OpenSubreport(mySubReportObject.SubreportName)
    '                        For intCounter1 = 0 To mySubRepDoc.Database.Tables.Count - 1
    '                            mySubRepDoc.Database.Tables(intCounter1).ApplyLogOnInfo(ConInfo)
    '                        Next
    '                    End If
    '                End With
    '            Next
    '        Next

    '        If sSelectionFormula.Length > 0 Then
    '            objReport.RecordSelectionFormula = sSelectionFormula
    '        End If
    '        Dim frm As New ViewReport

    '        frm.crvReport.ReportSource = Nothing
    '        frm.crvReport.ReportSource = objReport
    '        frm.crvReport.Show()

    '        Application.DoEvents()

    '        frm.crvReport.Text = sReportName
    '        frm.crvReport.Visible = True
    '        frm.BringToFront()

    '        Return True

    '    Catch ex As System.Exception
    '        MsgBox(ex.Message)
    '    End Try
    'End Function

End Module
