Imports Infragistics.Win.UltraWinGrid

Public Class clsGlazingNotification
    Dim collspPara As New Collection
    Dim colPara As New spParameters
    Dim ob As frmAgentNotifications
    Public Sub New()
    End Sub

    Public Sub New(ByRef ob As frmAgentNotifications)
        Me.ob = ob
    End Sub

    Function SaveNotificationDetails(ByRef objClsInvHeader As clsInvHeader, ByRef quoteJobID As String) As Integer

        Dim objIncident As New clsIncident
        Dim iIncident As Integer = 0
        Dim expDate As Date = objClsInvHeader.DueDate

        Try

            With objIncident
                '.Begin_Trans()

                .Outline = "Job ID :" & quoteJobID & vbCrLf & "Order Number :" & objClsInvHeader.OrderNum
                .CurrentAgent = AgentID
                .Customer = objClsInvHeader.AccountID
                .IncidentActionID = 1
                .IncidentCreated = Date.Now
                .IncidentDueDate = expDate.Date
                .IncidentLastModified = Date.Now
                .IncidentType = 1
                .IncidentTypeGroup = 1
                .IsRequireAck = 1
                .isSendRejectionNotification = 0
                .OurRef = .GetIncidentNo()
                .YourRef = objClsInvHeader.OrderNum

                'Update Incident Master
                iIncident = .Add_Incident
                If iIncident = -1 Then
                    Return 0
                    Exit Function
                End If

                'Update Incident Log 
                Dim iIncidentLog As Integer = 0
                .IncidentID = iIncident
                .IncidentActionDate = expDate.Date
                .IncidentActionID = 1
                .Revolution = .Outline
                .CurrentAgent = AgentID
                .Proxy = 0
                .NewAgentID = AgentID
                iIncidentLog = .Add_Incident_Log
                If iIncidentLog = -1 Then
                    Return 0
                    Exit Function
                End If

                collspPara.Clear()
                colPara.ParaName = "@dNotifyDate"
                colPara.ParaValue = expDate.Date
                collspPara.Add(colPara)

                colPara.ParaName = "@iForAgentID"
                If objClsInvHeader.DueDate < Today.AddDays(3).Date Then
                    colPara.ParaValue = AgentID

                ElseIf objClsInvHeader.DueDate > Today.AddDays(2).Date Then
                    colPara.ParaValue = -1

                End If

                collspPara.Add(colPara)
                colPara.ParaName = "@iIncidentID"
                colPara.ParaValue = iIncident
                collspPara.Add(colPara)

                colPara.ParaName = "@iIncidentLogID"
                colPara.ParaValue = iIncidentLog
                collspPara.Add(colPara)

                colPara.ParaName = "@bRead"
                colPara.ParaValue = 0
                collspPara.Add(colPara)

                newSQLQuary = "INSERT INTO _rtblNotify(dNotifyDate, iForAgentID, iIncidentID, iIncidentLogID,bRead) VALUES (@dNotifyDate,@iForAgentID,@iIncidentID,@iIncidentLogID,@bRead)"

                If .EXE_SQL_Trans_Para_Return(newSQLQuary, collspPara) = 0 Then
                    Return 0
                    Exit Function
                End If

                ' .Commit_Trans()
            End With
            Return iIncident

        Catch ex As Exception
            Return 0

        Finally
            collspPara.Clear()
        End Try
    End Function
    Function GetNotificationDetails(ByRef objClsInvHeader As clsInvHeader, ByRef isExistingOrder As Boolean, ByRef quoteJobID As String) As Integer
        Dim newSQLQuary As String = ""
        With objClsInvHeader
            Try
                If IsNothing(isExistingOrder) = False Then
                    If isExistingOrder = False Then
                        'If IsNothing(.DueDate) = False Then

                        If SaveNotificationDetails(objClsInvHeader, quoteJobID) = 0 Then
                            Return 0
                            Exit Function


                        End If
                        'End If

                    Else
                        colPara.ParaName = "@dDueBy"
                        colPara.ParaValue = .DueDate
                        collspPara.Add(colPara)

                        colPara.ParaName = "@cYourRef"
                        colPara.ParaValue = .OrderNum
                        collspPara.Add(colPara)

                        colPara.ParaName = "@todayDate"
                        colPara.ParaValue = Today.Date
                        collspPara.Add(colPara)

                        'If .EXE_SQL_Trans_Para_Return(SQL, collspPara) = 0 Then
                        '    Return 0
                        '    Exit Function
                        'End If

                        If IsNothing(.DueDate) = False Then
                            newSQLQuary = ""
                            If .DueDate < Today.AddDays(3).Date Then
                                newSQLQuary = " SET dateformat dmy UPDATE _rtblIncidents SET dLastModified = @todayDate, dDueBy = @dDueBy WHERE cYourRef = @cYourRef ; "

                                newSQLQuary += " SET dateformat dmy UPDATE  _rtblIncidentLog SET dActionDate = @dDueBy WHERE iIncidentID IN(SELECT idIncidents " & _
                                                "from _rtblIncidents where cYourRef = @cYourRef );"

                                newSQLQuary += " SET dateformat dmy UPDATE _rtblNotify SET dNotifyDate = @dDueBy WHERE iIncidentID IN(SELECT idIncidents from _rtblIncidents where cYourRef= @cYourRef )"

                            Else
                                newSQLQuary += " SET dateformat dmy UPDATE _rtblNotify SET iForAgentID = '-1', dNotifyDate = @dDueBy  where iIncidentID IN(SELECT idIncidents from _rtblIncidents where cYourRef= @cYourRef )"

                            End If
                        End If

                        If .EXE_SQL_Trans_Para_Return(newSQLQuary, collspPara) = 0 Then
                            MessageBox.Show("Data Base Error")
                            Return 0
                            Exit Function


                        End If
                    End If
                End If
                Return 1
            Catch ex As Exception
                Return 0

            Finally
                collspPara.Clear()

            End Try

        End With
    End Function

    Public Sub CheckNotificationDate()

        Dim newSQLObj As New clsSqlConn
        Dim NewDataSet As DataSet
        Dim NewDataRow As DataRow
        newSQLObj.Begin_Trans()
        Dim newSQLQuary = ""

        Try
            newSQLQuary = "SET dateformat dmy  SELECT _rtblNotify.dNotifyDate, _rtblIncidents.iCurrentAgentID, _rtblIncidents.idIncidents FROM _rtblNotify LEFT JOIN _rtblIncidents ON _rtblNotify.iIncidentID = _rtblIncidents.idIncidents  WHERE _rtblNotify.dNotifyDate > '" & Today.AddDays(-1).Date & "'"

            NewDataSet = newSQLObj.Get_Data_Trans(newSQLQuary)
            If NewDataSet.Tables(0).Rows.Count > 0 Then
                For Each NewDataRow In NewDataSet.Tables(0).Rows

                    If NewDataRow("dNotifyDate") < Today.AddDays(3).Date And NewDataRow("dNotifyDate") > Today.AddDays(-1).Date Then

                        colPara.ParaName = "@iForAgentID"
                        colPara.ParaValue = NewDataRow("iCurrentAgentID")
                        collspPara.Add(colPara)

                        colPara.ParaName = "@idIncidents"
                        colPara.ParaValue = NewDataRow("idIncidents")
                        collspPara.Add(colPara)

                        newSQLQuary = "SET dateformat dmy UPDATE _rtblNotify SET iForAgentID = @iForAgentID where iIncidentID = @idIncidents"
                        If newSQLObj.EXE_SQL_Trans_Para_Return(newSQLQuary, collspPara) = 0 Then
                            MessageBox.Show("Data Base Error")
                            newSQLObj.RollBack()
                            Exit Sub

                        End If
                        collspPara.Clear()
                    End If

                Next
            End If
            newSQLObj.Commit_Trans()
        Catch ex As Exception

        Finally
            collspPara.Clear()
            newSQLObj = Nothing

        End Try
    End Sub


    Public Sub FormatNotification(row As UltraGridRow)
        If row.Cells("NotifyDate").Value < Today.Date Then
            row.CellAppearance.ForeColor = Color.Red

        End If

    End Sub
End Class
