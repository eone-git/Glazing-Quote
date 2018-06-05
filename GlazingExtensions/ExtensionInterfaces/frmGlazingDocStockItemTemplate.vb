Imports Infragistics.Win.UltraWinGrid

Public Class frmGlazingDocStockItemTemplate
    Public stockLink As Integer = 0
    Dim clsGlazingDocStockItemHelperObj As New clsGlazingDocStockItemHelper(Me)
    Dim oPriceUnits As clsSOPricingAndUnits
    Public isLoading As Boolean = False
    Public totalAmount As Double = 0
    Public isClosing As Boolean = False
    Public selectedItems As String = ""
    Public selectedNewItems As String = ""
    Public selectedNewItemsDisplay As String = ""

    Sub New(ByRef clsPriceUnitsObj As clsSOPricingAndUnits)

        ' This call is required by the designer.
        InitializeComponent()
        oPriceUnits = clsPriceUnitsObj
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
                    ug2ActivRow.Cells("SimpleCode").Value = item("Stock_Sub_Link")
                    ug2ActivRow.Cells("Description1").Value = item("Stock_Sub_Link")
                    ug2ActivRow.Cells("ItemType").Value = item("iItemType")
                    ug2ActivRow.Cells("IsPriceItem").Value = item("IsPriceItem")

                    ucmbItemCode.Value = item("Stock_Sub_Link")
                    totalAmount = totalAmount + ug2ActivRow.Cells("Price").Value
                    'clsGlazingDocStockItemHelperObj.FillActiveRowFromSelectedProductParameters(ucmbItemCode, UG2.ActiveRow, oPriceUnits)
                    '  clsGlazingDocStockItemHelperObj.SetPriceOnThisRow(UG2.ActiveRow, oPriceUnits)
                Next
            End If
        Catch ex As Exception

        End Try
    End Sub

    Private Sub frmGlazingDocStockItemTemplate_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        FillComboData()
        If selectedItems = "" Then
            SetItemData()

        End If
        ucmbItemCode.Visible = True
        ucmbItemCode.Enabled = True
        ucmbPriceType.Visible = True
        ucmbPriceType.Enabled = True
    End Sub

    Private Sub UG2_InitializeLayout(sender As Object, e As InitializeLayoutEventArgs) Handles UG2.InitializeLayout
       
    End Sub

    Sub FillComboData()
        clsGlazingDocStockItemHelperObj.SetItemCodeData("Description_1", "StockLink", "Code")
        UG2.DisplayLayout.ColumnChooserEnabled = Infragistics.Win.DefaultableBoolean.True
        clsGlazingDocStockItemHelperObj.SetPriceListsData("PriceList", "CAT_ID")
        clsGlazingDocStockItemHelperObj.SetPriceTypeData("TYPE_PRICE", "TYPE_ID")
        'clsGlazingDocStockItemHelperObj.SetPriceTypeData("Description_1", "TYPE_ID", "TYPE_PRICE")
    End Sub

    Private Sub UG2_AfterCellUpdate(sender As Object, e As CellEventArgs) Handles UG2.AfterCellUpdate
     
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
              

            ElseIf activeCell.Column.Key = "Description1" Then
                If activeCell.Value <> e.Cell.Row.Cells("SimpleCode").Text Then
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

    Private Sub ucmbItemCode_RowSelected(sender As Object, e As RowSelectedEventArgs) Handles ucmbItemCode.RowSelected
        Try
            If e.Row.Cells("Description_1").Value <> UG2.ActiveRow.Cells("Description1").Value Then
                UG2.ActiveRow.Cells("Description1").Value = e.Row.Cells("StockLink").Value
                clsGlazingDocStockItemHelperObj.FillActiveRowFromSelectedProductParameters(ucmbItemCode, UG2.ActiveRow, oPriceUnits)
                clsGlazingDocStockItemHelperObj.SetPriceOnThisRow(UG2.ActiveRow, oPriceUnits)
            End If
        Catch ex As Exception

        End Try
    End Sub

    Private Sub ucmbItemDes_RowSelected(sender As Object, e As RowSelectedEventArgs) Handles ucmbItemDes.RowSelected
        Try
            If e.Row.Cells("Code").Value <> UG2.ActiveRow.Cells("SimpleCode").TEXT Then
                UG2.ActiveRow.Cells("SimpleCode").Value = e.Row.Cells("StockLink").Value
                clsGlazingDocStockItemHelperObj.FillActiveRowFromSelectedProductParameters(ucmbItemCode, UG2.ActiveRow, oPriceUnits)
                clsGlazingDocStockItemHelperObj.SetPriceOnThisRow(UG2.ActiveRow, oPriceUnits)
            End If
        Catch ex As Exception

        End Try
    End Sub

    Private Sub ucmbPriceType_RowSelected(sender As Object, e As RowSelectedEventArgs) Handles ucmbPriceType.RowSelected
        Try
            If IsNothing(ucmbPriceType.SelectedRow) = False And isLoading = False Then
                Me.UG2.ActiveRow.Cells("PriceType").Value = ucmbPriceType.Value
                clsGlazingDocStockItemHelperObj.SetPriceOnThisRow(UG2.ActiveRow, oPriceUnits)

            End If
        Catch ex As Exception

        End Try
    End Sub

    Private Sub cmbDDPriceListsTrade_RowSelected(sender As Object, e As RowSelectedEventArgs) Handles cmbDDPriceListsTrade.RowSelected
        If isLoading = False Then
            If IsNothing(cmbDDPriceListsTrade.Value) = False Then
                If Not IsNothing(cmbDDPriceListsTrade.Value) Then
                    me.UG2.ActiveRow.Cells("PriceList").Value = cmbDDPriceListsTrade.Value
                    clsGlazingDocStockItemHelperObj.SetPriceOnThisRow(UG2.ActiveRow, oPriceUnits)
                ElseIf Not IsNothing(cmbDDPriceListsSpecial.Value) Then
                    Me.UG2.ActiveRow.Cells("PriceList").Value = cmbDDPriceListsSpecial.Value
                    clsGlazingDocStockItemHelperObj.SetPriceOnThisRow(UG2.ActiveRow, oPriceUnits)
                End If
            End If
        End If
    End Sub

    Private Sub frmGlazingDocStockItemTemplate_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        Try
            Dim Result As DialogResult = MessageBox.Show("Do you want process the selection", Me.Text, MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question)
            If Result = Windows.Forms.DialogResult.Yes Then
                SetDataAsString()
                selectedItems = selectedNewItems
            ElseIf Result = Windows.Forms.DialogResult.No Then


            ElseIf Result = Windows.Forms.DialogResult.Cancel Then
                e.Cancel = True

            End If
        Catch ex As Exception

        End Try
    End Sub

    Sub SetDataAsString()
        Try
            For Each rows As UltraGridRow In UG2.Rows
                selectedNewItems = selectedNewItems & rows.Cells("StockLink").Value & "," & rows.Cells("PriceList").Value & "," & rows.Cells("PriceType").Value & "," & rows.Cells("Price").Value & ";"

                selectedNewItemsDisplay = selectedNewItemsDisplay & Chr(9) & "*" & rows.Cells("Description1").Value & vbCrLf

            Next
        Catch ex As Exception

        End Try
    End Sub
End Class