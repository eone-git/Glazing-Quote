Imports Infragistics.Win.UltraWinGrid

Public Class frmGlazingDocTemplate
    Dim frmGlazingQuote As frmGlazingQuote

    Public Sub New(ByRef frmGlazingQuote As frmGlazingQuote)
        ' This call is required by the designer.
        InitializeComponent()
        Me.frmGlazingQuote = frmGlazingQuote
        ' Add any initialization after the InitializeComponent() call.

    End Sub

    Function GetQuoteData(ByVal sqlQuary As String) As DataSet
        Dim dsQuoteTemp As DataSet
        Try
            Dim objSQL As New clsSqlConn
            With objSQL
                dsQuoteTemp = .GET_INSERT_UPDATE(sqlQuary)
            End With
            Return dsQuoteTemp
        Catch ex As Exception
            modGlazingQuoteExtension.GQShowMessage(ex.Message, Me.Text, MsgBoxStyle.Exclamation)
            Return dsQuoteTemp
        Finally
        End Try
    End Function
    Private Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click
        Dim activeRowIndex = 0
        Try
            If Me.btnSave.Text = "Save" Then
                For Each ugR As UltraGridRow In ugQuoteList.Rows
                    If ugR.Cells("TempName").Value = txttmpName.Text Then
                        If modGlazingQuoteExtension.GQShowMessage("Enterd name already exsit." & vbCrLf & "Do you wont to overwrite the " & txttmpName.Text, Me.Text, MsgBoxStyle.YesNo) = Windows.Forms.DialogResult.Yes Then

                        Else
                            Exit Sub
                        End If
                    End If
                Next
                frmGlazingQuote.SaveTempalteData(dsQuoteTemp, txttmpName.Text)

            ElseIf Me.btnSave.Text = "Rename" Then
                If IsNothing(ugQuoteList.ActiveRow) = False Then
                    activeRowIndex = Me.ugQuoteList.ActiveRow.Index
                    Dim oSQL = New clsSqlConn
                    Dim newSQLQuery As String
                    newSQLQuery = "UPDATE GlzQuote_Temp SET TempName = '" & txttmpName.Text & "'  WHERE TempName = '" & ugQuoteList.ActiveRow.Cells("TempName").Value & "'  "
                    oSQL.GET_INSERT_UPDATE(newSQLQuery)
                End If
            End If
            LoadGrideData()
        Catch ex As Exception
            modGlazingQuoteExtension.GQShowMessage(ex.Message, Me.Text, MsgBoxStyle.Exclamation)
        Finally
            Me.btnSave.Text = "Save"
            Me.ugQuoteList.ActiveRow = Me.ugQuoteList.Rows(activeRowIndex)
        End Try
    End Sub

    Private Sub frmGlazingDocTemplate_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        LoadGrideData()
    End Sub

    Sub LoadGrideData()
        Try
            frmGlazingQuote.openedModeulename = "Temp"
            ugQuoteList.DataSource = GetQuoteData("SELECT TempName FROM GlzQuote_Temp  WHERE IsRecordActive = 1 GROUP BY TempName")

        Catch ex As Exception
            modGlazingQuoteExtension.GQShowMessage(ex.Message, Me.Text, MsgBoxStyle.Exclamation)
        End Try
    End Sub

    Private Sub ugQuoteList_InitializeLayout(sender As Object, e As InitializeLayoutEventArgs) Handles ugQuoteList.InitializeLayout
        e.Layout.Bands(0).ColHeadersVisible = False
        e.Layout.Bands(0).Override.CellClickAction = CellClickAction.RowSelect
    End Sub

    Private Sub btnLoad_Click(sender As Object, e As EventArgs) Handles btnLoad.Click
        If IsNothing(ugQuoteList.ActiveRow) = False Then
            Dim dsQuoteTemp As DataSet
            Dim itmGroupID As Integer = 0
            Try
                dsQuoteTemp = GetQuoteData("SELECT * FROM GlzQuote_Temp WHERE TempName ='" & ugQuoteList.ActiveRow.Cells("TempName").Value & "'")
                If dsQuoteTemp.Tables(0).Rows.Count > 0 Then
                    ClearGride()
                    Dim dr As UltraGridRow
                    frmGlazingQuote.isOpeningQuote = True
                    For Each objQutDetailline In dsQuoteTemp.Tables(0).Rows
                        dr = frmGlazingQuote.UG2.DisplayLayout.Bands(0).AddNew
                        dr.Cells("QuoteFiedType").Value = objQutDetailline("TempQuoteFiedType")
                        frmGlazingQuote.ucmbQuoteLineType.Value = objQutDetailline("TempQuoteFiedType")
                        dr.Cells("LineComments").Value = objQutDetailline("DocDescription")
                        dr.Cells("ItmGroupID").Value = objQutDetailline("RowGroupID")
                        If objQutDetailline("RowGroupID") > itmGroupID Then
                            itmGroupID = objQutDetailline("RowGroupID")
                        End If
                        dr.Cells("TaxRate").Value = frmGlazingQuote.defaultTaxtRate
                        dr.Cells("TaxRateValue").Value = frmGlazingQuote.defaultTaxtRateValue

                    Next
                    frmGlazingQuote.subHeaderID = itmGroupID
                End If

            Catch ex As Exception
                modGlazingQuoteExtension.GQShowMessage(ex.Message, Me.Text, MsgBoxStyle.Exclamation)

            Finally
                frmGlazingQuote.isOpeningQuote = False

            End Try

        End If
    End Sub

    Private Sub frmGlazingDocTemplate_FormClosed(sender As Object, e As FormClosedEventArgs) Handles MyBase.FormClosed
        frmGlazingQuote.openedModeulename = "Temp"
    End Sub

    Public Sub ClearGride()
        Dim index As Integer = frmGlazingQuote.UG2.Rows.Count
        Dim isRowExist As Boolean = False
        Do
            If index > 0 Then
                frmGlazingQuote.UG2.Rows(index - 1).Delete()
                index -= 1
            Else
                Exit Do
            End If
        Loop Until index = 0

        frmGlazingQuote.lblTotExcAmo.Text = Format(0, "0.00")
        frmGlazingQuote.lblTotVatAmo.Text = Format(0, "0.00")
        frmGlazingQuote.lblTotIncAmo.Text = Format(0, "0.00")

    End Sub

    Private Sub btgnClear_Click(sender As Object, e As EventArgs) Handles btgnClear.Click
        ClearGride()
        frmGlazingQuote.AddNewRow("after")
    End Sub
    Sub DeleteTemplate(ByVal templateName As String)
        Try
            Dim oSQL = New clsSqlConn
            Dim newSQLQuery As String
            Dim isOldRecordAvailbale As Boolean = False
            Dim shouldRollBack As Boolean = False
            newSQLQuery = "SELECT TempName FROM GlzQuote_Temp WHERE TempName = '" & templateName & "'AND IsRecordActive = 1"
            If GetQuoteData(newSQLQuery).Tables(0).Rows.Count > 0 Then
                isOldRecordAvailbale = True
            End If

            oSQL.Begin_Trans()

            If isOldRecordAvailbale = True Then
                'Delete existing restor point
                newSQLQuery = "DELETE FROM GlzQuote_Temp  WHERE TempName = '" & templateName & "' AND IsRecordActive = 0"
                If oSQL.Exe_Query_Trans(newSQLQuery) = 0 Then
                    shouldRollBack = True
                End If
            End If

            'create restor point
            newSQLQuery = "UPDATE GlzQuote_Temp SET IsRecordActive = '0'  WHERE TempName = '" & templateName & "'  "
            If oSQL.Exe_Query_Trans(newSQLQuery) = 0 Then
                shouldRollBack = True
            End If

            'Hanndle the databse transaction
            If shouldRollBack = True Then
                modGlazingQuoteExtension.GQShowMessage("Data not saved", Me.Text, MsgBoxStyle.Critical)
                oSQL.Rollback_Trans()
                Exit Sub
            Else
                oSQL.Commit_Trans()
                LoadGrideData()
            End If

        Catch ex As Exception
            modGlazingQuoteExtension.GQShowMessage("Data not saved", Me.Text, MsgBoxStyle.Critical)
        End Try
    End Sub

    Private Sub btnDelete_Click(sender As Object, e As EventArgs) Handles btnDelete.Click
        If modGlazingQuoteExtension.GQShowMessage("Do you wont to delete the " & txttmpName.Text & " ?", "Deleting Template", MsgBoxStyle.YesNo) = Windows.Forms.DialogResult.Yes Then
            DeleteTemplate(txttmpName.Text)
        End If
    End Sub

    Private Sub ugQuoteList_AfterCellActivate(sender As Object, e As EventArgs) Handles ugQuoteList.AfterCellActivate
        txttmpName.Text = ugQuoteList.ActiveCell.Value
    End Sub

    Private Sub tsmDelete_Click(sender As Object, e As EventArgs) Handles tsmDelete.Click
        If IsNothing(ugQuoteList.ActiveRow) = False Then
            If modGlazingQuoteExtension.GQShowMessage("Do you wont to delete the " & ugQuoteList.ActiveRow.Cells("TempName").Value & " ?", "Deleteing Template", MsgBoxStyle.YesNo) = Windows.Forms.DialogResult.Yes Then
                DeleteTemplate(ugQuoteList.ActiveRow.Cells("TempName").Value)
            End If
        End If
    End Sub

    Private Sub ugQuoteList_AfterRowActivate(sender As Object, e As EventArgs) Handles ugQuoteList.AfterRowActivate
        If IsNothing(ugQuoteList.ActiveRow) = False Then
            txttmpName.Text = ugQuoteList.ActiveRow.Cells("TempName").Value
        End If
    End Sub


    Private Sub tsmRename_Click(sender As Object, e As EventArgs) Handles tsmRename.Click
        If IsNothing(ugQuoteList.ActiveRow) = False Then
            'e.Layout.Bands(0).Override.CellClickAction = CellClickAction.RowSelect
            ugQuoteList.PerformAction(UltraGridAction.EnterEditMode, False, False)
            ugQuoteList.ActiveCell = ugQuoteList.ActiveRow.Cells("TempName")
        End If
        btnSave.Text = "Rename"
    End Sub


    Private Sub ugQuoteList_AfterEnterEditMode(sender As Object, e As EventArgs) Handles ugQuoteList.AfterEnterEditMode

    End Sub
End Class