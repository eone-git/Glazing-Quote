Public Class clsGlazingQuoteExtension
    Dim moduleName As String = "Glazing Quote"
    Public reuturnedMessage As String = ""

    Public Sub GQCreatefrmGlazingQuoteObject()
        Try
            glazingQuote = New frmGlazingQuote

        Catch ex As Exception
            GQShowMessage(ex.Message, moduleName, MsgBoxStyle.Critical)

        End Try

    End Sub

    Public Sub GQOpenGlazingQuote(ByRef quoteOrdeIndex As Integer, ByRef isACopy As Boolean)
        Dim glazingQuote As New frmGlazingQuote

        Try
            glazingQuote.quoteOrdeIndex = quoteOrdeIndex
            glazingQuote.isACopy = isACopy
            glazingQuote.InitializeQuotation()
            glazingQuote.Show()
            glazingQuote.WindowState = FormWindowState.Maximized
            glazingQuote.BringToFront()
            glazingQuote.Refresh()
            'If isACopy = True Then
            '    GQDocumentLog(quoteOrdeIndex, QuoteStateValueCOPY)

            'End If
        Catch ex As Exception
            GQShowMessage(ex.Message, moduleName, MsgBoxStyle.Critical)

        End Try
    End Sub

    Public Function GQEmailSentStateUpdate(ByVal quoteOrdeIndex As Integer, ByRef reqQuoteStateValue As Integer, Optional ByRef requestReason As Boolean = False) As Integer
        Dim collspPara As New Collection
        Dim colPara As New spParameters
        Dim newSQLQuery As String = ""
        Dim newSQL As New clsSqlConn
        Dim items As Array
        items = System.Enum.GetValues(GetType(QuoteStateValue))
        Dim messageLable As String = System.Enum.GetName(GetType(QuoteStateValue), reqQuoteStateValue)
        Try
            If IsNothing(messageLable) = False Then
                If messageLable.Length > 1 Then
                    If messageLable.Substring(messageLable.Length - 2) = "ed" Then
                        messageLable = messageLable.Substring(0, messageLable.Length - 2)
                    End If
                End If
            End If
            If requestReason = True Then
                Dim frmGlazingMessageBoxObj As New frmGlazingMessageBox(Me, moduleName, "Reason for " & messageLable.ToLower() & " this quotation")
                frmGlazingMessageBoxObj.ShowDialog()
            End If
            newSQL.Begin_Trans()

            colPara.ParaName = "@QuoteStateID"
            colPara.ParaValue = reqQuoteStateValue
            collspPara.Add(colPara)

            colPara.ParaName = "@OrderIndex"
            colPara.ParaValue = quoteOrdeIndex
            collspPara.Add(colPara)

            sqlQuary = "update spilInvNum set QuoteStateID = @QuoteStateID where OrderIndex = @OrderIndex"
            If newSQL.EXE_SQL_Trans_Para(sqlQuary, collspPara) = 0 Then
                MsgBox("Error in Updating state", MsgBoxStyle.Critical, "SPIL Glass")
                Exit Function
            End If

            If reqQuoteStateValue <> QuoteStateValue.SentAndConfirmationPending Then
                GQDocumentLog(quoteOrdeIndex, (System.Enum.GetName(GetType(QuoteStateValue), reqQuoteStateValue)), newSQL, reuturnedMessage)
            End If

            newSQL.Commit_Trans()
            Return 1

        Catch ex As Exception
            GQShowMessage(ex.Message, moduleName, MsgBoxStyle.Critical)
            Return 0

        Finally
            collspPara.Clear()

        End Try
    End Function

    Public Sub GQShowJobDescription(ByRef orderIndex As Integer)
        Dim glazingQuote As New frmGlazingQuote
        Dim objSQL As New clsSqlConn
        Dim DS_BATCHES As DataSet
        Try

            SQL = "SELECT * FROM GlzQuote_Job_Details WHERE OrderIndex='" & orderIndex & "'"
            With objSQL
                DS_BATCHES = .GET_INSERT_UPDATE(SQL)
                If DS_BATCHES.Tables(0).Rows.Count <> 0 Then
                    If IsNothing(DS_BATCHES.Tables(0).Rows(0).Item("GlzQuoteJobDes")) = False Then
                        If DS_BATCHES.Tables(0).Rows(0).Item("GlzQuoteJobDes") = "" Then
                            GQShowMessage("There is no desctioption for selected quotation", moduleName, MsgBoxStyle.Information)
                            Exit Sub

                        End If
                    Else
                        GQShowMessage("There is no mathing quotation", moduleName, MsgBoxStyle.Information)
                        Exit Sub

                    End If
                Else
                    GQShowMessage("There is no mathing quotation", moduleName, MsgBoxStyle.Information)
                    Exit Sub

                End If
                glazingQuote.OpenDescriptionState(False, DS_BATCHES.Tables(0).Rows(0).Item("GlzQuoteJobDes"))

            End With

        Catch ex As Exception
            GQShowMessage(ex.Message, moduleName, MsgBoxStyle.Critical)

        Finally
            newGlazingQuote = Nothing

        End Try
    End Sub

    Public Function GQDocumentLog(ByRef orderIndex As Integer, ByRef logAction As String, Optional sqlObj As clsSqlConn = Nothing, Optional description1 As String = "")

        Dim objDLItem As New clsDocumentLogEntry
        Dim isNewConection As Boolean = False
        Try
            If IsNothing(sqlObj) = True Then
                sqlObj = New clsSqlConn
                sqlObj.Begin_Trans()
                isNewConection = True
            End If

            objDLItem.iDocID = orderIndex
            objDLItem.iDocTypeID = GlassDocTypes.Quotation
            objDLItem.LogAction = logAction
            objDLItem.DocItemCount = 0
            objDLItem.DocServiceCount = 0
            objDLItem.LogDateTime = Now
            objDLItem.EnteredBy = strUserName
            objDLItem.Description1 = description1
            If objDLItem.AddDocLogWithTrans(sqlObj.Con, sqlObj.Trans) < 1 Then
                sqlObj.Rollback_Trans()
                Return 0
                Exit Function
            End If

            If isNewConection = True Then
                sqlObj.Commit_Trans()

            End If
            Return 1
        Catch ex As Exception
            GQShowMessage(ex.Message, moduleName, MsgBoxStyle.Critical)

        Finally
            isNewConection = False
        End Try
    End Function


End Class
