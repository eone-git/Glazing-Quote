Public Class ForGlazingModule
    Dim moduleName As String = "Glazing Quote"
    Dim glazingQuoteObj As New frmGlazingQuote
    Public _ProjectId As Integer = 0
    Public _StageId As Integer = 0
    Public _JobId As Integer = 0

    Public IsFromJobProject As Boolean = False
    Public IsProgressCliam As Boolean = False
    Public _TotalInvoiced As Decimal = 0
    Public _isSaved As Boolean = False
    Public _odrIndex As Integer = 0

    Public Sub GetProjects(ByVal custId As Integer, Optional ByVal isEnabled As Boolean = False)

        Dim dsProjects As DataSet
        SQL = "SELECT   Id,Name FROM  SpilGlazing_Project WHERE CustomerID=" & custId
        dsProjects = New clsSqlConn().GET_DataSet(SQL)
        If dsProjects.Tables(0).Rows.Count > 0 Then
            glazingQuoteObj.cmbCustProject.Enabled = True
        End If

        glazingQuoteObj.cmbCustProject.DataSource = dsProjects.Tables(0)
        glazingQuoteObj.cmbCustProject.ValueMember = "Id"
        glazingQuoteObj.cmbCustProject.DisplayMember = "Name"
        glazingQuoteObj.cmbCustProject.DisplayLayout.Bands(0).Columns(1).Width = 200
        glazingQuoteObj.cmbCustProject.DisplayLayout.Bands(0).Columns(0).Hidden = True
        If _ProjectId > 0 Then
            glazingQuoteObj.cmbCustProject.Value = _ProjectId
        End If
        glazingQuoteObj.cmbCustProject.Enabled = isEnabled
    End Sub

    Public Sub GetStages(ByVal projectId As Integer, Optional ByVal isEnabled As Boolean = False)
        Dim dsProjects As DataSet
        If (projectId = 0) Then
            SQL = "SELECT    Id,StageName FROM   SpilGlazing_ProjectStage"
        Else
            SQL = "SELECT    Id,StageName FROM   SpilGlazing_ProjectStage WHERE ProjectID=" & projectId
        End If

        dsProjects = New clsSqlConn().GET_DataSet(SQL)
        If dsProjects.Tables(0).Rows.Count > 0 Then
            glazingQuoteObj.cmbProjectStage.Enabled = True
        End If

        glazingQuoteObj.cmbProjectStage.DataSource = dsProjects.Tables(0)
        glazingQuoteObj.cmbProjectStage.ValueMember = "Id"
        glazingQuoteObj.cmbProjectStage.DisplayMember = "StageName"
        glazingQuoteObj.cmbProjectStage.DisplayLayout.Bands(0).Columns(1).Width = 200
        glazingQuoteObj.cmbProjectStage.DisplayLayout.Bands(0).Columns(0).Hidden = True
        If _StageId > 0 Then
            glazingQuoteObj.cmbProjectStage.Value = _StageId
        End If
        glazingQuoteObj.cmbProjectStage.Enabled = isEnabled
    End Sub

    Public Sub GetJobs(ByVal projectId As Integer, Optional ByVal isEnabled As Boolean = False)
        Dim dsProjects As DataSet
        If (projectId = 0) Then
            SQL = "SELECT Id,Name FROM  SpilGlazing_Job"
        Else
            SQL = "SELECT Id,Name FROM  SpilGlazing_Job WHERE ProjectId=" & projectId
        End If

        dsProjects = New clsSqlConn().GET_DataSet(SQL)
        If dsProjects.Tables(0).Rows.Count > 0 Then
            glazingQuoteObj.cmbCustJob.Enabled = True
        End If

        glazingQuoteObj.cmbCustJob.DataSource = dsProjects.Tables(0)
        glazingQuoteObj.cmbCustJob.ValueMember = "Id"
        glazingQuoteObj.cmbCustJob.DisplayMember = "Name"
        glazingQuoteObj.cmbCustJob.DisplayLayout.Bands(0).Columns(1).Width = 200
        glazingQuoteObj.cmbCustJob.DisplayLayout.Bands(0).Columns(0).Hidden = True
        If _JobId > 0 Then
            glazingQuoteObj.cmbCustJob.Value = _JobId
        End If
        glazingQuoteObj.cmbCustJob.Enabled = isEnabled
    End Sub

    Private Sub SetJobProjectDetails(ByRef clsCon As clsInvHeader, ByVal orderIdex As Integer)
        '' find doc type
        Dim doctyp As Integer = 0
        If (glazingQuoteObj.pubMeSpilDocTypeID = 4 And clsCon.DocState = 4) Then
            doctyp = 0
        Else
            doctyp = glazingQuoteObj.pubMeSpilDocTypeID
        End If

        Dim projDs As New DataSet()
        SQL = "SELECT * FROM SpilGlazing_ProjectDocument WHERE DocumentId = " & orderIdex & " AND DocumentType = " & doctyp
        projDs = clsCon.GET_DATA_SQL(SQL)
        If Not projDs Is Nothing And projDs.Tables(0).Rows.Count > 0 Then
            GetProjects(clsCon.AccountID)
            GetStages(projDs.Tables(0).Rows(0)("ProjectId"))
            glazingQuoteObj.cmbCustProject.Value = projDs.Tables(0).Rows(0)("ProjectId")
            If projDs.Tables(0).Rows(0)("StageId") > 0 Then
                glazingQuoteObj.cmbProjectStage.Value = projDs.Tables(0).Rows(0)("StageId")
            End If

        End If
        projDs = Nothing
        SQL = "SELECT * FROM SpilGlazing_JobDocument WHERE DocumentId = " & orderIdex & " AND DocumentType = " & doctyp
        projDs = clsCon.GET_DATA_SQL(SQL)
        If Not projDs Is Nothing And projDs.Tables(0).Rows.Count > 0 Then
            GetJobs(projDs.Tables(0).Rows(0)("ProjectId"))
            If projDs.Tables(0).Rows(0)("JobId") > 0 Then
                glazingQuoteObj.cmbCustJob.Value = projDs.Tables(0).Rows(0)("JobId")
            End If
        End If
    End Sub
       
End Class
