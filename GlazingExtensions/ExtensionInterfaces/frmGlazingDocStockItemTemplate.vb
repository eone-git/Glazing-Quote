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
                    UG2.ActiveRow=row
                    ug2ActivRow.Cells("StockLink").Value = item("StockLink")

                Next
            End If
        Catch ex As Exception

        End Try
    End Sub

    Private Sub frmGlazingDocStockItemTemplate_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        SetItemData()

    End Sub
End Class