Imports Infragistics.Win.UltraWinGrid

Public Class frmGlazingDocStockItemTemplate
    Public stockLink As Integer = 0
    Dim clsGlazingDocStockItemHelperObj As New clsGlazingDocStockItemHelper(Me)

    Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.

    End Sub

    Sub SetItemData()
        Dim newDataSet As DataSet
        Dim ug2ActivRow As UltraGridRow = UG2.ActiveRow
        Dim row As UltraGridRow
        Try
            newDataSet = clsGlazingDocStockItemHelperObj.GetStkItemDetails(stockLink)
            If newDataSet.Tables(0).Rows.Count > 0 Then
                For Each item As DataRow In newDataSet.Tables(0).Rows
                    row = UG2.DisplayLayout.Bands(0).AddNew
                    UG2.ActiveRow = row
                    ug2ActivRow = UG2.ActiveRow
                    ug2ActivRow.Cells("StockLink").Value = item("Stock_Sub_Link")

                Next
            End If
        Catch ex As Exception

        End Try
    End Sub

    Private Sub frmGlazingDocStockItemTemplate_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        FillComboData()
        SetItemData()
        ucmbItemCode.Visible = True
        ucmbItemCode.Enabled = True
        ucmbPriceType.Visible = True
        ucmbPriceType.Enabled = True
    End Sub

    Private Sub UG2_InitializeLayout(sender As Object, e As InitializeLayoutEventArgs) Handles UG2.InitializeLayout
       
    End Sub

    Sub FillComboData()
        clsGlazingDocStockItemHelperObj.SetItemCodeData("Description_1", "StockLink", "Code")
        clsGlazingDocStockItemHelperObj.SetPriceListsData("PriceList", "CAT_ID")
        clsGlazingDocStockItemHelperObj.SetPriceTypeData("TYPE_PRICE", "TYPE_ID")
        'clsGlazingDocStockItemHelperObj.SetPriceTypeData("Description_1", "TYPE_ID", "TYPE_PRICE")
    End Sub

    Private Sub UG2_AfterCellUpdate(sender As Object, e As CellEventArgs) Handles UG2.AfterCellUpdate
        Try
            GrideHandler(e)
        Catch ex As Exception

        End Try
    End Sub

    Sub GrideHandler(Optional ByRef e As CellEventArgs = Nothing)
        Dim activeCell As UltraGridCell

        If IsNothing(e) = False Then
            activeCell = e.Cell
        Else
            activeCell = UG2.ActiveCell
        End If

        Try
            If activeCell.Column.Key = "StockLink" Then
                activeCell.Row.Cells("SimpleCode").Value = e.Cell.Text
                ' activeCell.Row.Cells("Description1").Value = e.Cell.Text

            ElseIf activeCell.Column.Key = "SimpleCode" Then
                If activeCell.Value <> e.Cell.Row.Cells("Description1").Value Then
                    activeCell.Row.Cells("Description1").Value = e.Cell.Text
                End If

            ElseIf activeCell.Column.Key = "Description1" Then
                If activeCell.Value <> e.Cell.Row.Cells("SimpleCode").Value Then
                    activeCell.Row.Cells("SimpleCode").Value = e.Cell.Text
                End If

            End If
        Catch ex As Exception

        End Try
    End Sub

    Private Sub cmbDDPriceListsSpecial_RowSelected(sender As Object, e As RowSelectedEventArgs) Handles cmbDDPriceListsSpecial.RowSelected
        Try
            e.Row.Cells("SimpleCode").Value = e.Row.Cells("").Text
        Catch ex As Exception

        End Try
    End Sub
End Class