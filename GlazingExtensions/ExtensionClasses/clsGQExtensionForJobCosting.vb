Public Class clsGQExtensionForJobCosting
    Dim moduleName As String = "Glazing Quote"
    Dim glazingQuoteObj As frmGlazingQuote

    Public Sub New(ByRef Obj As frmGlazingQuote)
        Me.glazingQuoteObj = Obj

    End Sub

    Public Sub New()

    End Sub

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
        If glazingQuoteObj._ProjectId > 0 Then
            glazingQuoteObj.cmbCustProject.Value = glazingQuoteObj._ProjectId
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
        If glazingQuoteObj._StageId > 0 Then
            glazingQuoteObj.cmbProjectStage.Value = _StageId
        End If
        glazingQuoteObj.cmbProjectStage.Enabled = isEnabled
    End Sub

    Public Sub GetJobs(ByVal projectId As Integer, ByVal customerId As Integer, Optional ByVal isEnabled As Boolean = False)
        Dim dsProjects As DataSet
        If (projectId = 0) Then
            SQL = "SELECT Id,Name FROM  SpilGlazing_Job WHERE (ProjectId=0 OR ProjectId is Null) AND CustomerID=" & customerId

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
        If glazingQuoteObj._JobId > 0 Then
            glazingQuoteObj.cmbCustJob.Value = _JobId
        End If
        glazingQuoteObj.cmbCustJob.Enabled = isEnabled
    End Sub

    Public Sub GetJobsByCustomer(ByVal customerId As Integer, Optional ByVal isEnabled As Boolean = False)
        Dim dsProjects As DataSet
        SQL = "SELECT Id,Name FROM  SpilGlazing_Job WHERE (ProjectId=0 OR ProjectId is Null) AND  CustomerID=" & customerId
        dsProjects = New clsSqlConn().GET_DataSet(SQL)
        If dsProjects.Tables(0).Rows.Count > 0 Then
            glazingQuoteObj.cmbCustJob.Enabled = True
        End If

        glazingQuoteObj.cmbCustJob.DataSource = dsProjects.Tables(0)
        glazingQuoteObj.cmbCustJob.ValueMember = "Id"
        glazingQuoteObj.cmbCustJob.DisplayMember = "Name"
        glazingQuoteObj.cmbCustJob.DisplayLayout.Bands(0).Columns(1).Width = 200
        glazingQuoteObj.cmbCustJob.DisplayLayout.Bands(0).Columns(0).Hidden = True
        If glazingQuoteObj._JobId > 0 Then
            glazingQuoteObj.cmbCustJob.Value = _JobId
        End If
        glazingQuoteObj.cmbCustJob.Enabled = isEnabled
    End Sub

    Public Sub SetJobProjectDetails(ByRef clsCon As clsInvHeader, ByVal orderIdex As Integer)
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
            GetJobs(projDs.Tables(0).Rows(0)("ProjectId"), glazingQuoteObj.cmbAccount.SelectedRow.Cells("DCLink").Value)
            If projDs.Tables(0).Rows(0)("JobId") > 0 Then
                glazingQuoteObj.cmbCustJob.Value = projDs.Tables(0).Rows(0)("JobId")
            End If
        End If
    End Sub

    Public Sub HideProfitCostFields(ByVal state As Boolean)
        If (state = True) Then
            If glazingQuoteObj.pubMeSpilDocTypeID = GlassDocTypes.Quotation Or glazingQuoteObj.pubMeSpilDocTypeID = GlassDocTypes.Estimate Then
                glazingQuoteObj.UG2.DisplayLayout.Bands(0).Columns("Cost").Hidden = False
                glazingQuoteObj.UG2.DisplayLayout.Bands(0).Columns("Profit").Hidden = False
                glazingQuoteObj.UG2.DisplayLayout.Bands(1).Columns("Cost").Hidden = False
                glazingQuoteObj.UG2.DisplayLayout.Bands(1).Columns("Profit").Hidden = False
                glazingQuoteObj.FileToolStripMenuItem.Visible = True
                'jobProjectToolsStripSaparator.Visible = True
                If glazingQuoteObj.pubMeSpilDocTypeID = GlassDocTypes.Quotation Then
                    glazingQuoteObj.ApprovedQuoteToolStripMenuItem.Visible = True
                    glazingQuoteObj.CancelApprovedQuoteToolStripMenuItem.Visible = True
                    glazingQuoteObj.ConvertToQuoteToolStripMenuItem.Visible = False
                Else
                    glazingQuoteObj.ApprovedQuoteToolStripMenuItem.Visible = False
                    glazingQuoteObj.CancelApprovedQuoteToolStripMenuItem.Visible = False
                    glazingQuoteObj.ConvertToQuoteToolStripMenuItem.Visible = True
                End If
            End If

            If glazingQuoteObj.pubMeSpilDocTypeID = GlassDocTypes.Estimate Then
                glazingQuoteObj.lblTaxCode.Visible = False
                'glazingQuoteObj.cmbTaxRate.Visible = False
                glazingQuoteObj.txtOrderNet.Visible = False
                glazingQuoteObj.lblSubTotalExcl.Visible = False
                glazingQuoteObj.lblSubTotalExcl.Visible = False
                glazingQuoteObj.T1.Visible = False
                glazingQuoteObj.lblGstTotal.Visible = False
                glazingQuoteObj.T2.Visible = False
                glazingQuoteObj.txtTaxTotal.Visible = False
                glazingQuoteObj.UG2.DisplayLayout.Bands(0).Columns("TaxCode").Hidden = True
                glazingQuoteObj.UG2.DisplayLayout.Bands(1).Columns("TaxCode").Hidden = True
            End If
        Else
            glazingQuoteObj.UG2.DisplayLayout.Bands(0).Columns("Cost").Hidden = True
            glazingQuoteObj.UG2.DisplayLayout.Bands(0).Columns("Profit").Hidden = True
            glazingQuoteObj.UG2.DisplayLayout.Bands(1).Columns("Cost").Hidden = True
            glazingQuoteObj.UG2.DisplayLayout.Bands(1).Columns("Profit").Hidden = True
            glazingQuoteObj.cmbCustProject.Enabled = False
            glazingQuoteObj.cmbCustJob.Enabled = False
            glazingQuoteObj.cmbProjectStage.Enabled = False

            'jobProjectToolsStripSaparator.Visible = False
            glazingQuoteObj.ApprovedQuoteToolStripMenuItem.Visible = False
            glazingQuoteObj.CancelApprovedQuoteToolStripMenuItem.Visible = False
            glazingQuoteObj.ConvertToQuoteToolStripMenuItem.Visible = False
        End If


        glazingQuoteObj.lblProfit.Visible = state
        glazingQuoteObj.txtProfit.Visible = state
        ' glazingQuoteObj.lblProfitPercentage.Visible = state

    End Sub

    Public Sub SaveJobProjectInfo(ByRef con As clsInvHeader)
        If (IsFromJobProject) Then
            '' find doc type
            Dim doctyp As Integer = 0
            If (pubMeSpilDocTypeID = 4 And con.DocState = 4) Then
                doctyp = 0
            Else
                doctyp = pubMeSpilDocTypeID
            End If

            ''delete jobs
            SQL = "DELETE FROM  SpilGlazing_JobDocument  WHERE DocumentId=" & con.OrderIndex & " AND DocumentType=" & doctyp
            If con.Execute_Sql_Trans(SQL) = 0 Then
                sError = "Error in Spil Glazing Job Document "
                con.Rollback_Trans()
                Exit Sub
            End If
            '' delete projects
            SQL = "DELETE FROM  SpilGlazing_ProjectDocument WHERE DocumentId=" & con.OrderIndex & " AND DocumentType=" & doctyp
            If con.Execute_Sql_Trans(SQL) = 0 Then
                sError = "Error in Spil Glazing _Project  Document "
                con.Rollback_Trans()
                Exit Sub
            End If

            '' save project
            If glazingQuoteObj.cmbCustProject.Value > 0 Then
                SQL = "INSERT INTO SpilGlazing_ProjectDocument (ProjectId,DocumentId,DocumentType,StageId) VALUES (" & glazingQuoteObj.cmbCustProject.Value & ", " & con.OrderIndex & " , " & doctyp & ", " & If(glazingQuoteObj.cmbProjectStage.Value = Nothing, 0, glazingQuoteObj.cmbProjectStage.Value) & " )"

                If con.Execute_Sql_Trans(SQL) = 0 Then
                    sError = "Error in Spil Glazing Project Document "
                    con.Rollback_Trans()
                    Exit Sub
                End If
            End If

            '' save job
            If glazingQuoteObj.cmbCustJob.Value > 0 Then
                SQL = "INSERT INTO SpilGlazing_JobDocument (JobId, DocumentId, DocumentType, ProjectId, StageId) VALUES (" & glazingQuoteObj.cmbCustJob.Value & ", " & con.OrderIndex & " , '" & doctyp & "', " & If(glazingQuoteObj.cmbCustProject.Value = Nothing, 0, glazingQuoteObj.cmbCustProject.Value) & ", " & If(glazingQuoteObj.cmbProjectStage.Value = Nothing, 0, glazingQuoteObj.cmbProjectStage.Value) & " )"

                If con.Execute_Sql_Trans(SQL) = 0 Then
                    sError = "Error in Spil Glazing Project Document "
                    con.Rollback_Trans()
                    Exit Sub
                End If
            End If


        End If
    End Sub

    Public Sub ApprovedQuotation()
        Try
            If MessageBox.Show("Please confirm approval ?", "Quotation", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = Windows.Forms.DialogResult.No Then
                Exit Sub
            End If

            Dim objSQL As New clsSqlConn
            Dim SQLSELECT As String
            Dim SQLUPDATE As String

            Dim projId As Integer = If(glazingQuoteObj.cmbCustProject.Value = Nothing, 0, glazingQuoteObj.cmbCustProject.Value)
            Dim jobId As Integer = If(glazingQuoteObj.cmbCustJob.Value = Nothing, 0, glazingQuoteObj.cmbCustJob.Value)
            Dim stagId As Integer = If(glazingQuoteObj.cmbProjectStage.Value = Nothing, 0, glazingQuoteObj.cmbProjectStage.Value)


            '' find exsisting
            SQLSELECT = "SELECT SpilInvNum.OrderNum FROM SpilGlazing_JobDocument JOIN SpilInvNum ON SpilInvNum.OrderIndex = SpilGlazing_JobDocument.DocumentId WHERE SpilGlazing_JobDocument.JobId = " & jobId & " AND SpilGlazing_JobDocument.Approved_quot = 1 AND SpilGlazing_JobDocument.DocumentType=5; "

            SQLSELECT += "SELECT SpilInvNum.OrderNum FROM SpilGlazing_ProjectDocument JOIN SpilInvNum ON SpilInvNum.OrderIndex = SpilGlazing_ProjectDocument.DocumentId WHERE SpilGlazing_ProjectDocument.ProjectId = " & projId & " AND SpilGlazing_ProjectDocument.Approved_quot = 1 AND SpilGlazing_ProjectDocument.DocumentType = 5; "

            DS = objSQL.GET_DATA_SQL(SQLSELECT)
            If (DS.Tables(0).Rows.Count > 0) Then
                MessageBox.Show("Quotation : " & DS.Tables(0).Rows(0)("OrderNum").ToString() & " is already approved! ", "Quotation", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Exit Sub
            ElseIf (DS.Tables(1).Rows.Count > 0) Then
                MessageBox.Show("Quotation : " & DS.Tables(1).Rows(0)("OrderNum").ToString() & " is already approved! ", "Quotation", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Exit Sub
            End If

            SQLUPDATE = "UPDATE  SpilGlazing_ProjectDocument  SET Approved_quot =" & 1 & " WHERE ProjectId=" & projId & " AND StageId = " & stagId & ";"

            SQLUPDATE += "UPDATE SpilGlazing_JobDocument SET Approved_quot =" & 1 & " WHERE ProjectId=" & projId & " AND StageId = " & stagId & " AND JobId = " & jobId & ";"


            objSQL.Exe_Query(SQLUPDATE)
            MessageBox.Show(" Quotation approved successfully ", "Save", MessageBoxButtons.OK, MessageBoxIcon.Information)

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Exclamation, "SPIL Glass")

        End Try
    End Sub

    Private Sub CancelApprovedQuote()
        If MsgBox("Are youe sure want to cancel this approved Quotation ", MsgBoxStyle.Question + MsgBoxStyle.YesNo + MsgBoxStyle.DefaultButton1, "Confirmation") = MsgBoxResult.Yes Then
            Dim objSQL As New clsSqlConn
            SQL = "UPDATE  SpilGlazing_ProjectDocument  SET Approved_quot =" & 0 & " WHERE DocumentId=" & glazingQuoteObj.pubMeOrderIndex & "; "
            SQL += "UPDATE  SpilGlazing_JobDocument  SET Approved_quot =" & 0 & " WHERE DocumentId=" & glazingQuoteObj.pubMeOrderIndex & ";"

            objSQL.Exe_Query(SQL)
        End If
    End Sub

End Class
