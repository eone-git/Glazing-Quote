Public Module modGlazingQuoteExtension

    Public Function GQShowMessage(msgBoxMessage As String, formName As String, messageButtons As MessageBoxButtons) As DialogResult
        Try
            If messageButtons = MessageBoxButtons.YesNo Then
                Return MessageBox.Show(msgBoxMessage, formName, MessageBoxButtons.YesNo)
            Else
                Return MessageBox.Show(msgBoxMessage, formName, MessageBoxButtons.OK)
            End If
        Catch ex As Exception
            MsgBox(ex.Message)

        End Try
    End Function

    Public Enum QuateFiedTypesList As Integer
        Text = 1
        Header_Main = 2
        Header_Sub = 3
        Subtotal = 4
        Stock_Item = 5

    End Enum

    Public Enum QuoteStateValue As Integer
        EditMode = 1
        Copy = 2
        SentAndConfirmationPending = 3
        Confirmed = 4
        Reopened = 5
        Cancelled = 6
        Hold = 7

    End Enum


    Public Enum MeasurementsType As Integer
        Imperial = 1
        Metric = 2

    End Enum

    Public measurementsTypeOrigin As Integer = 1
    Public measurementsTypeTemp As Integer = 1

End Module
