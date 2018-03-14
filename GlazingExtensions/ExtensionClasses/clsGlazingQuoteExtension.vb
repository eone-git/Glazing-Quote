Public Class clsGlazingQuoteExtension
    Dim glazingQuote = Nothing
    Dim moduleName As String = "Glazing Quote"

    Public Sub GQCreatefrmGlazingQuoteObject()
        Try
            glazingQuote = New frmGlazingQuote
        Catch ex As Exception
            GQShowMessage(ex.Message, moduleName, MsgBoxStyle.Critical)
        End Try


    End Sub

    Public Sub GQOpenGlazingQuote(ByRef quoteOrdeIndex As Integer, ByRef isACopy As Boolean)
        Try
            GQCreatefrmGlazingQuoteObject()
            glazingQuote.quoteOrdeIndex = quoteOrdeIndex
            glazingQuote.isACopy = isACopy
            glazingQuote.Show()
            glazingQuote.WindowState = FormWindowState.Maximized
            glazingQuote.BringToFront()

        Catch ex As Exception
            GQShowMessage(ex.Message, moduleName, MsgBoxStyle.Critical)

        End Try
    End Sub

    Public Sub GQEmailSentStateUpdate(ByVal orderIndex As Integer, ByRef QuoteStateValue As Integer)
        Dim collspPara As New Collection
        Dim colPara As New spParameters
        Dim newSQLQuery As String = ""
        Dim newSQL As New clsSqlConn
        Try
            colPara.ParaName = "@QuoteStateID"
            colPara.ParaValue = QuoteStateValue
            collspPara.Add(colPara)

            colPara.ParaName = "@OrderIndex"
            colPara.ParaValue = orderIndex
            collspPara.Add(colPara)

            sqlQuary = "update spilInvNum set QuoteStateID = @QuoteStateID where OrderIndex = @OrderIndex"
            If newSQL.EXE_SQL_Para(sqlQuary, collspPara) = 0 Then
                MsgBox("Error in Updating state", MsgBoxStyle.Critical, "SPIL Glass")
                Exit Sub
            End If

        Catch ex As Exception
            GQShowMessage(ex.Message, moduleName, MsgBoxStyle.Critical)

        Finally
            collspPara.Clear()

        End Try
    End Sub
    Public Sub GQShowJobDescription(ByRef orderIndex As Integer)
        Dim objSQL As New clsSqlConn
        GQCreatefrmGlazingQuoteObject()
        Try
            newGlazingQuote = New frmGlazingQuote()
            SQL = "SELECT * FROM GlzQuote_Job_Details WHERE OrderIndex='" & orderIndex & "'"
            With objSQL
                DS_BATCHES = .GET_INSERT_UPDATE(SQL)
                glazingQuote.OpenDescriptionState(False, DS_BATCHES.Tables(0).Rows(0).Item("GlzQuoteJobDes"))

            End With

        Catch ex As Exception
            GQShowMessage(ex.Message, moduleName, MsgBoxStyle.Critical)

        Finally
            newGlazingQuote = Nothing

        End Try
    End Sub
End Class
