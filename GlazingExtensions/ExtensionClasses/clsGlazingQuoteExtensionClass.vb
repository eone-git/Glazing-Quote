Public Class clsGlazingQuoteExtensionClass
    Sub QuoteEmailSentStateUpdate(ByVal orderIndex As Integer)
        Dim collspPara As New Collection
        Dim colPara As New spParameters
        Dim newSQLQuery As String = ""
        Dim newSQL As New clsSqlConn
        Dim moduleName As String = "GlazingQuote"
        Try
            colPara.ParaName = "@QuoteStateID"
            colPara.ParaValue = frmGlazingQuote.QuoteStateValue.SentAndConfirmationPending
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
            frmGlazingQuote.ShowMessage(ex.Message, moduleName, MsgBoxStyle.Critical)

        Finally
            collspPara.Clear()

        End Try
    End Sub

End Class
