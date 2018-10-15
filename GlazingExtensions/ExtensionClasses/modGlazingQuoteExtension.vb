Imports System.Text.RegularExpressions

Public Module modGlazingQuoteExtension

    ''' <summary>
    ''' ## [Message Box Message {String}], [Module/Form Name {String}], [*Message Box Buttons {MessageBoxButtons}], ( [*Message Box Type {String}], *[Method Name {String}] ) ##
    ''' </summary>
    Public Function GQShowMessage(ByRef msgBoxMessage As String, ByRef formName As String, Optional ByRef messageButtons As MessageBoxButtons = MessageBoxButtons.OK, Optional ByRef type As String = "", Optional ByRef methodName As String = "Default") As DialogResult
        Try
            Dim methodNameFull As String
            Dim methodWordCollection As MatchCollection = Regex.Matches(methodName, "[A-Z][a-z]+")
            Dim wordString As String
            Dim isFirstCounter As Boolean = True
            For Each word As Match In methodWordCollection
                If isFirstCounter = True Then
                    isFirstCounter = False
                    wordString = StrConv((word.ToString), VbStrConv.ProperCase)
                Else
                    wordString = word.ToString.ToLower
                End If

                counter = counter + 1
                methodNameFull = methodNameFull & " " & wordString

            Next
            If methodNameFull = " " Then
                methodNameFull = ""
            Else
                methodNameFull = " :" & methodNameFull
            End If

            If type = "" Then
                If messageButtons = MessageBoxButtons.YesNo Then
                    Return MessageBox.Show(msgBoxMessage, formName, MessageBoxButtons.YesNo, MessageBoxIcon.Question)
                Else
                    Return MessageBox.Show(msgBoxMessage, formName, MessageBoxButtons.OK)
                End If
            Else
                If type = "question" Then
                    Return MessageBox.Show(msgBoxMessage, formName & methodNameFull, MessageBoxButtons.YesNo, MessageBoxIcon.Question)
                ElseIf type = "warning" Then
                    Return MessageBox.Show(msgBoxMessage, formName & methodNameFull, MessageBoxButtons.OK, MessageBoxIcon.Warning)
                End If
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

    Public Enum GeographicRegion As Integer
        NorthAmerica = 1
        Europe = 2
        Oceania = 3
    End Enum

    Public geographicRegionID As Integer = 3

End Module
