Imports Infragistics.Win.UltraWinGrid

Public Class frmGlazingDocDescription
    Dim obSQL As clsSqlConn = Nothing
    Public DocDesTypeID As Integer = 0
    Dim inEditMode As Boolean = False
    Public DocDesTypeName As String = ""

    Private Sub GetData()
        If IsNothing(DocDesTypeID) = False Then
            Dim sqlQuery As String = "SELECT Text FROM GlzQuote_Texts_Master WHERE TextTypeID =" & DocDesTypeID & ""
            If IsNothing(obSQL) Then
                obSQL = New clsSqlConn
            End If
            Dim dsDocDes As DataSet = obSQL.GET_INSERT_UPDATE(sqlQuery)
            ugDocDes.DataSource = dsDocDes.Tables(0)
        End If
    End Sub

    Private Sub btnDelete_Click(sender As Object, e As EventArgs) Handles btnDelete.Click
        If ugDocDes.Selected.Rows.Count > 0 Then
            If MsgBox("Please confirm the deletion of the selected row.", MsgBoxStyle.Information + MsgBoxStyle.YesNo, "Confirmation") = MsgBoxResult.No Then
                Exit Sub
            End If
            ugDocDes.Selected.Rows(0).Delete(False)
            SaveDocDes()
        End If
    End Sub

    Private Sub SaveDocDes()
        Try
            Dim sqlQuery As String
            If IsNothing(obSQL) Then
                obSQL = New clsSqlConn
            End If
            Cursor.Current = Cursors.WaitCursor
            sqlQuery = "DELETE FROM GlzQuote_Texts_Master WHERE TextTypeID =" & DocDesTypeID & ""
            obSQL.GET_INSERT_UPDATE(sqlQuery)

            For Each drDocDes As UltraGridRow In ugDocDes.Rows
                sqlQuery = "Insert INTO GlzQuote_Texts_Master (TextTypeID,Text) values(" & DocDesTypeID & ", '" & drDocDes.Cells("Text").Text & "')"
                obSQL.GET_INSERT_UPDATE(sqlQuery)
            Next

        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            GetData()
            Cursor.Current = Cursors.Default
            obSQL = Nothing
        End Try
    End Sub

    Private Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click
        SaveDocDes()
        inEditMode = False
        FormElemetsBeahavior()
    End Sub

    Private Sub frmGlazingDocDescription_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'SaveNotes()
        FormElemetsBeahavior()
    End Sub


    Private Sub ugDocDes_InitializeLayout(sender As Object, e As InitializeLayoutEventArgs) Handles ugDocDes.InitializeLayout
        GetData()
        ugDocDes.DisplayLayout.Bands(0).Columns("Text").Header.Caption = DocDesTypeName + " Text"
        ugDocDes.DisplayLayout.Bands(0).Columns("Text").CellActivation = Activation.AllowEdit
        ugDocDes.DisplayLayout.Override.AllowAddNew = AllowAddNew.TemplateOnBottom
        e.Layout.Bands(0).HeaderVisible = False
        ugDocDes.DisplayLayout.Override.CellClickAction = CellClickAction.RowSelect
    End Sub

    Private Sub ugDocDes_AfterCellUpdate(sender As Object, e As CellEventArgs) Handles ugDocDes.AfterCellUpdate
        e.Cell.Row.PerformAutoSize()

    End Sub
   
    Private Sub btnSelect_Click(sender As Object, e As EventArgs) Handles btnSelect.Click
        ReturnSelectedText()
    End Sub

    Sub ReturnSelectedText()
        Dim rowsCollection As SelectedRowsCollection = ugDocDes.Selected.Rows
        If IsNothing(rowsCollection) Then
        Else
            frmGlazingQuote.selectedDocDes = ""
            For Each row As UltraGridRow In rowsCollection

                If frmGlazingQuote.selectedDocDes <> "" Then
                    frmGlazingQuote.selectedDocDes = frmGlazingQuote.selectedDocDes + vbCrLf + row.Cells("Text").Value
                ElseIf frmGlazingQuote.selectedDocDes = "" Then
                    If IsDBNull(row.Cells("Text").Value) = False Then
                        frmGlazingQuote.selectedDocDes = row.Cells("Text").Value
                    End If
                End If
            Next
        End If
        Me.Close()
    End Sub

    Private Sub frmGlazingDocDescription_FormClosed(sender As Object, e As FormClosedEventArgs) Handles MyBase.FormClosed

    End Sub

    Private Sub tsbEdit_Click(sender As Object, e As EventArgs) Handles tsbEdit.Click
        inEditMode = True
        FormElemetsBeahavior()

    End Sub
    Sub FormElemetsBeahavior()
        If inEditMode = False Then
            btnDelete.Visible = False
            btnSave.Visible = False
            tsbEdit.Visible = True
            btnSelect.Visible = True
            tsbEdit.Text = "Edit"

            ugDocDes.DisplayLayout.Override.CellClickAction = CellClickAction.RowSelect

        Else
            btnDelete.Visible = True
            btnSave.Visible = True
            tsbEdit.Visible = False
            btnSelect.Visible = False
            tsbEdit.Text = "Exite edit mode"

            ugDocDes.DisplayLayout.Override.CellClickAction = CellClickAction.EditAndSelectText

        End If
    End Sub

    Private Sub ugDocDes_DoubleClickRow(sender As Object, e As DoubleClickRowEventArgs) Handles ugDocDes.DoubleClickRow
        If inEditMode = False Then
            ReturnSelectedText()
        End If
    End Sub
End Class