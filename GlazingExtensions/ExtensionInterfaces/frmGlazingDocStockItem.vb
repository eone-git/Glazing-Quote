#Region "Namespaces"

Imports System.Data.SqlClient
Imports Infragistics.Win.UltraWinGrid
Imports System.IO
Imports GlassInventoryModule
Imports System.Drawing.Imaging
Imports System.Reflection.MethodBase

#End Region

Public Class frmGlazingDocStockItem

#Region "Variables"

    Dim dsDept As New DataSet
    Dim dsTypes As New DataSet
    Private sNotify As String = String.Empty
    Dim con As New SqlConnection(strCon)
    Dim frmGlazingQuote As frmGlazingQuote
    Public selectedRow As UltraGridRow
    Dim oPriceUnits As New clsSOPricingAndUnits
    Dim comboVlaueChanged As Boolean = False
    Public isLoading As Boolean = False
    Dim isTemplateItemLoading As Boolean = False
    Dim priceTypeValue As Integer = 0
    Dim priceListValue As Integer = 0
    Dim itemIsChanging As Boolean = False
    Dim clsGlazingDocStockItemHelperObj As New clsGlazingDocStockItemHelper(Me)
    Dim clsGlazingQuoteExtensionObj As New clsGlazingQuoteExtension
    Dim moduleDefaultsObj As New clsSOModuleDefaults
    Dim oSOModuleDefaults As New clsSOModuleDefaults

    Public templateItemSubItemsPrice As Double = 0
    Public templateItemSubItemsPriceNew As Double = 0

    Dim templateItemSubItemsData As String = ""
    Dim templateItemSubItemsDataEx As String = ""

    'Sub items
    Dim lineComments As String = ""
    Dim lineCommentsNew As String = ""

    Dim lineSubItems As String = ""
    Dim lineSubItemsNew As String = ""

    Dim lineSubItemsToShow As String = ""
    Dim lineSubItemsToShowNew As String = ""

    Dim lineSubItemsTotalExc As Double = 0
    Dim lineSubItemsTotalExcNew As Double = 0

    Dim lineSubItemsTotalTax As Double = 0
    Dim lineSubItemsTotalTaxNew As Double = 0

    Dim lineSubItemsTotalInc As Double = 0
    Dim lineSubItemsTotalIncNew As Double = 0

    Dim totalExclusiveAmount As Double = 0
    Dim totalExclusiveAmountNew As Double = 0

    Dim actualTotalExc As Double = 0
    Dim TaxExempt As Boolean = False

    Dim isSubItemsDiscards As Boolean = False

    Dim glassVolum As Decimal = 0
    Dim SQFeetForPricing As Decimal = 0.0
    Dim withFractions As Boolean = False

#End Region
#Region "Constructor"

    Public Sub New(ByRef frmGlazingQuote As frmGlazingQuote)
        ' This call is required by the designer.
        InitializeComponent()
        Me.frmGlazingQuote = frmGlazingQuote
        selectedRow = frmGlazingQuote.UG2.ActiveRow

        oPriceUnits.oCustomer = New clsCustomer(frmGlazingQuote.cmbAccount.Value)
        oPriceUnits.iDefaultStockPriceListID = moduleDefaultsObj.DefaultTradePriceListID

        If selectedRow.Cells("ItemType").Value > 0 Then
            isLoading = True
        End If
    End Sub

    Public Sub New()

    End Sub

#End Region
#Region "Load Components Data"
    Sub LoadComboData()
        FillPRICE_TYPES()
        Get_StockItems()
        SetThikness()
        clsGlazingDocStockItemHelperObj.SetPriceListsData("PriceList", "CAT_ID")
        clsGlazingDocStockItemHelperObj.SetItemCodeData("Description_1", "StockLink", "Code")
    End Sub

    Public Sub FillPRICE_TYPES()
        Dim newDataset As DataSet = clsGlazingDocStockItemHelperObj.SetPriceTypeData("TYPE_PRICE", "TYPE_ID", "TYPE_PRICE")
        'This is to load form level variable to use at the price type column activation event to validate to show only template related 
        'price types on the second band for template items (to apply grid filters)
        Dim bFound_IsApplyOnlyTemplate_PriceTypesItems = False
        For Each item As DataRow In newDataset.Tables(0).Rows
            If item("IsApplyOnlyTemplate") = 1 Then
                bFound_IsApplyOnlyTemplate_PriceTypesItems = True
                Exit For
            End If
        Next
        'END OF  'price types on the second band for template items (to apply grid filters)


    End Sub

    'Not use
    Sub GetDataSource()
        Dim DA = New SqlDataAdapter(CMD)
        Dim itemTypeID As Integer = 1

        Do
            If itemTypeID > 4 And itemTypeID < 16 Then

            Else
                SQL = ""
                SQL = SQL & "SELECT     StkItem.Description_1, StkItem.cSimpleCode as Code, StkItem.StockLink, " & vbCrLf
                SQL = SQL & "StkItem.ufIIThickness, TaxRate.TaxRate,stkItem.uiIISRVPRICEID, " & vbCrLf
                SQL = SQL & "StkItem.TTI, StkItem.uiIIPRICETYPEID, StkItem.uiIIItemType,StkItem.Description_3, " & vbCrLf
                SQL = SQL & "StkItem.AddDetails,StkItem.TemplType,StkItem.TemplItemCat,StkItem.TemplSelectThickness, " & vbCrLf
                SQL = SQL & "CategoryID " & vbCrLf
                SQL = SQL & ",SubTypeID " & vbCrLf
                SQL = SQL & "FROM StkItem " & vbCrLf
                SQL = SQL & "LEFT OUTER JOIN (SELECT   StockLink, MAX(CategoryID) AS CategoryID " & vbCrLf
                SQL = SQL & "        FROM     spilStkTemplateItems " & vbCrLf
                SQL = SQL & "        GROUP BY spilStkTemplateItems.StockLink) t2 ON StkItem.StockLink = t2.StockLink " & vbCrLf
                SQL = SQL & "LEFT OUTER JOIN " & vbCrLf
                SQL = SQL & "TaxRate ON StkItem.TTI = TaxRate.Code " & vbCrLf
                SQL = SQL & "WHERE  StkItem.uiIIItemType =" & itemTypeID & " and StkItem.ubIIGLASSSERVICE=0 and ItemActive = 1 order by StkItem.Description_2, StkItem.ufIIThickness"
                '1,2,3,4,16

                Dim objSQL As New clsSqlConn
                With objSQL
                    Try
                        dataSource = .GET_INSERT_UPDATE(SQL)
                    Catch ex As Exception

                    End Try
                End With
            End If

            itemTypeID += 1
        Loop Until itemTypeID > 16

    End Sub

    Private Sub Get_StockItems()
        SQL = ""
        SQL = SQL & "SELECT   StkItem.ItemImage, StkItem.Description_1, StkItem.cSimpleCode as Code, StkItem.StockLink, " & vbCrLf
        SQL = SQL & "StkItem.ufIIThickness, TaxRate.TaxRate,stkItem.uiIISRVPRICEID, " & vbCrLf
        SQL = SQL & "StkItem.TTI, StkItem.uiIIPRICETYPEID, StkItem.uiIIItemType,StkItem.Description_3, " & vbCrLf
        SQL = SQL & "StkItem.AddDetails,StkItem.TemplType,StkItem.TemplItemCat,StkItem.TemplSelectThickness, " & vbCrLf
        SQL = SQL & "CategoryID " & vbCrLf
        SQL = SQL & ",SubTypeID " & vbCrLf
        SQL = SQL & "FROM StkItem " & vbCrLf
        SQL = SQL & "LEFT OUTER JOIN (SELECT   StockLink, MAX(CategoryID) AS CategoryID " & vbCrLf
        SQL = SQL & "        FROM     spilStkTemplateItems " & vbCrLf
        SQL = SQL & "        GROUP BY spilStkTemplateItems.StockLink) t2 ON StkItem.StockLink = t2.StockLink " & vbCrLf
        SQL = SQL & "LEFT OUTER JOIN " & vbCrLf
        SQL = SQL & "TaxRate ON StkItem.TTI = TaxRate.Code " & vbCrLf
        SQL = SQL & "WHERE  StkItem.uiIIItemType IN (1,2,3,4,16) and StkItem.ubIIGLASSSERVICE=0 and ItemActive = 1 order by StkItem.Description_2, StkItem.ufIIThickness"

        '" WHERE     StkItem.uiIIItemType =" & GlassItemTypes.Glass & "  AND ItemActive = 1  order by StkItem.Description_2, StkItem.ufIIThickness "
        Dim objSQL As New clsSqlConn
        With objSQL
            Try

                DS_Stock = .GET_INSERT_UPDATE(SQL)
                ucmbItemDes.DataSource = DS_Stock.Tables(0)

                ucmbItemDes.ValueMember = "StockLink"
                ucmbItemDes.DisplayMember = "Description_1"

                ucmbItemDes.DisplayLayout.Bands(0).Columns("Description_1").Width = 400
                ucmbItemDes.DisplayLayout.Bands(0).Columns("Description_1").Header.Caption = "Description"
                ucmbItemDes.DisplayLayout.Bands(0).Columns("Code").Width = 100
                ucmbItemDes.DisplayLayout.Bands(0).Columns("StockLink").Hidden = True
                ucmbItemDes.DisplayLayout.Bands(0).Columns("ufIIThickness").Hidden = True
                ucmbItemDes.DisplayLayout.Bands(0).Columns("TaxRate").Hidden = True
                ucmbItemDes.DisplayLayout.Bands(0).Columns("uiIISRVPRICEID").Hidden = True
                ucmbItemDes.DisplayLayout.Bands(0).Columns("TTI").Hidden = True
                ucmbItemDes.DisplayLayout.Bands(0).Columns("uiIIPRICETYPEID").Hidden = True
                ucmbItemDes.DisplayLayout.Bands(0).Columns("uiIIItemType").Hidden = True
                ucmbItemDes.DisplayLayout.Bands(0).Columns("Description_3").Hidden = True
                ucmbItemDes.DisplayLayout.Bands(0).Columns("ItemImage").Hidden = True

                ucmbItemDes.DisplayLayout.Bands(0).Columns("TemplType").Hidden = True
                ucmbItemDes.DisplayLayout.Bands(0).Columns("TemplItemCat").Hidden = True
                ucmbItemDes.DisplayLayout.Bands(0).Columns("TemplSelectThickness").Hidden = True
                ucmbItemDes.DisplayLayout.Bands(0).Columns("CategoryID").Hidden = True
                ucmbItemDes.DisplayLayout.Bands(0).Columns("SubTypeID").Hidden = True

                ucmbItemDes.DisplayLayout.Bands(0).Columns("AddDetails").Header.Caption = "Additional Details"

                ucmbItemCode.DataSource = Nothing
                Dim dvCode As New DataView(ucmbItemDes.DataSource)
                dvCode.Sort = "Code"

                ucmbItemCode.DataSource = dvCode

                ucmbItemCode.ValueMember = "StockLink"
                ucmbItemCode.DisplayMember = "Code"

                ucmbItemCode.DisplayLayout.Bands(0).Columns("Code").Header.VisiblePosition = 0
                ucmbItemCode.DisplayLayout.Bands(0).Columns("Description_1").Header.VisiblePosition = 1

                ucmbItemCode.DisplayLayout.Bands(0).Columns("Description_1").Width = 400
                ucmbItemCode.DisplayLayout.Bands(0).Columns("Description_1").Header.Caption = "Description"
                ucmbItemCode.DisplayLayout.Bands(0).Columns("Code").Width = 100

                ucmbItemCode.DisplayLayout.Bands(0).Columns("StockLink").Hidden = True
                ucmbItemCode.DisplayLayout.Bands(0).Columns("ufIIThickness").Hidden = True
                ucmbItemCode.DisplayLayout.Bands(0).Columns("TaxRate").Hidden = True
                ucmbItemCode.DisplayLayout.Bands(0).Columns("uiIISRVPRICEID").Hidden = True
                ucmbItemCode.DisplayLayout.Bands(0).Columns("TTI").Hidden = True
                ucmbItemCode.DisplayLayout.Bands(0).Columns("uiIIPRICETYPEID").Hidden = True
                ucmbItemCode.DisplayLayout.Bands(0).Columns("uiIIItemType").Hidden = True

                ucmbItemCode.DisplayLayout.Bands(0).Columns("TemplType").Hidden = True
                ucmbItemCode.DisplayLayout.Bands(0).Columns("TemplItemCat").Hidden = True
                ucmbItemCode.DisplayLayout.Bands(0).Columns("TemplSelectThickness").Hidden = True
                ucmbItemCode.DisplayLayout.Bands(0).Columns("CategoryID").Hidden = True
                ucmbItemCode.DisplayLayout.Bands(0).Columns("SubTypeID").Hidden = True

                ucmbItemCode.DisplayLayout.Bands(0).Columns("Description_3").Hidden = True
                ucmbItemCode.DisplayLayout.Bands(0).Columns("ItemImage").Hidden = True
                ucmbItemCode.DisplayLayout.Bands(0).Columns("AddDetails").Header.Caption = "Additional Details"

            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.Exclamation, "SPIL Glass")
            Finally
                .Dispose()
                objSQL = Nothing
            End Try
        End With
    End Sub

    'Not use
    Private Sub Get_PriceLists()
        SQL = "SELECT     CAT_ID, CAT_DESCRIPTION As PriceList FROM stkPRICE_CATEGORYspil"
        SQL += " SELECT   CAT_ID, CAT_DESCRIPTION As PriceList FROM stkPRICE_SPECIALCATEGORYspil"
        Dim objSQL As New clsSqlConn
        With objSQL
            Try
                DS = New DataSet
                DS = .GET_DATA_SQL(SQL)
                cmbDDPriceListsTrade.DataSource = DS.Tables(0)
                cmbDDPriceListsTrade.DisplayLayout.Bands(0).ColHeadersVisible = False
                cmbDDPriceListsTrade.DisplayMember = "PriceList"
                cmbDDPriceListsTrade.ValueMember = "CAT_ID"
                cmbDDPriceListsTrade.DisplayLayout.Bands(0).Columns("CAT_ID").Hidden = True

                cmbDDPriceListsSpecial.DataSource = DS.Tables(1)
                cmbDDPriceListsSpecial.DisplayLayout.Bands(0).ColHeadersVisible = False
                cmbDDPriceListsSpecial.DisplayMember = "PriceList"
                cmbDDPriceListsSpecial.ValueMember = "CAT_ID"
                cmbDDPriceListsSpecial.DisplayLayout.Bands(0).Columns("CAT_ID").Hidden = True
            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.Exclamation, "SPIL Glass")
            Finally
                .Dispose()
                objSQL = Nothing
            End Try
        End With
    End Sub

#End Region
#Region "Form core funtions"

    Private Sub frmGlazingDocStockItem_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        LoadData()
    End Sub

    Private Sub frmGlazingDocStockItem_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        frmGlazingQuote.isStockItemActive = False
    End Sub

#End Region
#Region "CRUD"

    Sub LoadData()
        Try
            If IsNothing(selectedRow.Cells("Price_Type").Value) = False Then
                priceTypeValue = selectedRow.Cells("Price_Type").Value
            End If

            If IsNothing(selectedRow.Cells("PriceList").Value) = False And IsDBNull(selectedRow.Cells("PriceList").Value) = False Then
                priceListValue = selectedRow.Cells("PriceList").Value
            End If

            LoadComboData()
            frmGlazingQuote.isStockItemActive = True
            BackupOldData()
            LoadEditMode()
            CreateLineComments()
            txtVolume.ReadOnly = True

        Catch ex As Exception
            modGlazingQuoteExtension.GQShowMessage(ex.Message, Me.Text, MsgBoxStyle.Critical, "warning", "LoadData")

        End Try
    End Sub

    Sub LoadEditMode()

        Dim band As UltraGridBand = cmbDDItemType.DisplayLayout.Bands(0)
        band.ColumnFilters.ClearAllFilters()
        band.ColumnFilters("TypeID").FilterConditions.Add(FilterComparisionOperator.Equals, 1)
        band.ColumnFilters("TypeID").LogicalOperator = FilterLogicalOperator.Or
        band.ColumnFilters("TypeID").FilterConditions.Add(FilterComparisionOperator.Equals, 2)
        band.ColumnFilters("TypeID").LogicalOperator = FilterLogicalOperator.Or
        band.ColumnFilters("TypeID").FilterConditions.Add(FilterComparisionOperator.Equals, 3)
        band.ColumnFilters("TypeID").LogicalOperator = FilterLogicalOperator.Or
        band.ColumnFilters("TypeID").FilterConditions.Add(FilterComparisionOperator.Equals, 4)
        ' band.ColumnFilters("TypeID").FilterConditions.Add(FilterComparisionOperator.Equals, 1)

        If IsNothing(frmGlazingQuote.UG2.ActiveRow) = False Then

            If selectedRow.Cells("ItemType").Value > 0 Then
                utxtDocDes.Value = selectedRow.Cells("LineComments").Value
                cmbDDItemType.Value = selectedRow.Cells("ItemType").Value
                ucmbItemCode.Value = selectedRow.Cells("StockLink").Value
                txtHeight.Value = selectedRow.Cells("Height").Value
                ctxtWidth.Value = selectedRow.Cells("Width").Value
                txtVolume.Value = selectedRow.Cells("Volume").Value
                txtQty.Value = selectedRow.Cells("Qty").Value
                unelblTotal.Value = frmGlazingQuote.UG2.ActiveRow.Cells("ItmExcAmount").Value

                If IsNothing(selectedRow.Cells("Price_Type").Value) = False Then
                    ucmbPriceType.Value = priceTypeValue

                End If

                If IsNothing(selectedRow.Cells("PriceList").Value) = False Then
                    cmbDDPriceListsSpecial.Value = priceListValue
                    cmbDDPriceListsTrade.Value = priceListValue

                End If

                If cmbDDItemType.Value = 2 Then
                    Dim subItemsPrice As Double = clsGlazingDocStockItemHelperObj.GetSubItemPrice(frmGlazingQuote.UG2.ActiveRow.Cells("templateData").Value)
                    'txtPrice.Value = subItemsPrice
                    Dim actualAmount As Double = frmGlazingQuote.UG2.ActiveRow.Cells("amount").Value - frmGlazingQuote.UG2.ActiveRow.Cells("tax").Value
                    unelblTotal.Value = actualAmount
                    txtPrice.Value = ((actualAmount / txtVolume.Value) / txtQty.Value) - subItemsPrice
                    'CalculateTotalExc(subItemsPrice, actualAmount)

                Else
                    txtPrice.Value = frmGlazingQuote.UG2.ActiveRow.Cells("Original_Price").Value
                End If
            Else
                cmbDDItemType.Value = 1

            End If
            isLoading = False
        End If
    End Sub

    Sub SaveData()
        Try
            frmGlazingQuote.isStockItemActive = False
            Dim emptyArray As Byte() = New Byte(63) {}

            If cmbDDItemType.Value = GlassItemTypes.Template Or cmbDDItemType.Value = GlassItemTypes.Glass Then
                If isSubItemsDiscards = False Then
                    frmGlazingQuote.UG2.ActiveRow.Cells("templateData").Value = lineSubItems
                    If lineComments = "" Then
                        frmGlazingQuote.UG2.ActiveRow.Cells("LineComments").Value = utxtDocDes.Text
                    Else
                        frmGlazingQuote.UG2.ActiveRow.Cells("LineComments").Value = lineComments
                    End If

                Else
                    frmGlazingQuote.UG2.ActiveRow.Cells("templateData").Value = lineSubItemsNew
                    If lineComments = "" Then
                        frmGlazingQuote.UG2.ActiveRow.Cells("LineComments").Value = utxtDocDes.Text
                    Else
                        frmGlazingQuote.UG2.ActiveRow.Cells("LineComments").Value = lineComments
                    End If
                End If
            Else
                frmGlazingQuote.UG2.ActiveRow.Cells("LineComments").Value = utxtDocDes.Text
            End If
            frmGlazingQuote.UG2.ActiveRow.Cells("Qty").Value = txtQty.Value
            frmGlazingQuote.UG2.ActiveRow.Cells("Width").Value = ctxtWidth.Text
            frmGlazingQuote.UG2.ActiveRow.Cells("Height").Value = txtHeight.Text
            frmGlazingQuote.UG2.ActiveRow.Cells("Volume").Value = txtVolume.Value

            frmGlazingQuote.UG2.ActiveRow.Cells("ServiceItemTotNet").Value = lineSubItemsTotalInc
            frmGlazingQuote.UG2.ActiveRow.Cells("ServiceGross").Value = lineSubItemsTotalExc
            frmGlazingQuote.UG2.ActiveRow.Cells("ServiceItemTax").Value = lineSubItemsTotalTax

            frmGlazingQuote.UG2.ActiveRow.Cells("Price").Value = If(IsNothing(txtPrice.Value) = False, txtPrice.Value, 0)
            frmGlazingQuote.UG2.ActiveRow.Cells("ItmExcAmount").Value = unelblTotal.Text
            'CalculateFinalAmount()
            Dim itemTotalExc = unelblTotal.Text - lineSubItemsTotalExc
            frmGlazingQuote.UG2.ActiveRow.Cells("Amount").Value = itemTotalExc + If(TaxExempt = False, ((itemTotalExc * frmGlazingQuote.dCusTaxRate) / 100), 0) + lineSubItemsTotalInc
            frmGlazingQuote.UG2.ActiveRow.Cells("Net").Value = itemTotalExc + If(TaxExempt = False, ((itemTotalExc * frmGlazingQuote.dCusTaxRate) / 100), 0) + lineSubItemsTotalInc
            frmGlazingQuote.UG2.ActiveRow.Cells("Original_Price").Value = txtPrice.Value

            frmGlazingQuote.UG2.ActiveRow.Cells("ItemType").Value = cmbDDItemType.Value
            'frmGlazingQuote.UG2.ActiveRow.Cells("SimpleCode").Value = ctxtWidth.Text
            frmGlazingQuote.UG2.ActiveRow.Cells("Description1").Value = utxtDocDes.Text
            frmGlazingQuote.UG2.ActiveRow.Cells("SimpleCode").Value = ucmbItemCode.Text
            frmGlazingQuote.UG2.ActiveRow.Cells("Price_Type").Value = ucmbPriceType.Value
            'frmGlazingQuote.UG2.ActiveRow.Cells("TaxExempt").Value = ucmbItemCode.ActiveRow.Cells("TaxExempt").Value

            If IsNothing(cmbDDPriceListsSpecial.Value) = False And cmbDDPriceListsSpecial.Visible = True Then
                frmGlazingQuote.UG2.ActiveRow.Cells("PriceList").Value = cmbDDPriceListsSpecial.Value

            ElseIf IsNothing(cmbDDPriceListsTrade.Value) = False And cmbDDPriceListsTrade.Visible = True Then
                frmGlazingQuote.UG2.ActiveRow.Cells("PriceList").Value = cmbDDPriceListsTrade.Value

            End If

            If IsNothing(uPicBox.Image) = False Then
                frmGlazingQuote.UG2.ActiveRow.Cells("ItemImage").Activate()
                If IsDBNull(ucmbItemCode.SelectedRow.Cells("ItemImage").Value) = False Then
                    frmGlazingQuote.UG2.ActiveRow.Cells("ItemImage").Value = frmGlazingQuote.ByteToImage(ucmbItemCode.SelectedRow.Cells("ItemImage").Value)
                    frmGlazingQuote.UG2.ActiveRow.Cells("isImageAttached").Value = True
                Else

                End If
                frmGlazingQuote.UG2.ActiveRow.Cells("ItemImageByteArray").Value = ucmbItemCode.SelectedRow.Cells("ItemImage").Value
            Else
                frmGlazingQuote.UG2.ActiveRow.Cells("ItemImage").Value = emptyArray
                frmGlazingQuote.UG2.ActiveRow.Cells("isImageAttached").Value = False

            End If

            If IsNothing(ucmbItemCode.SelectedRow) = False Then
                frmGlazingQuote.UG2.ActiveRow.Cells("StockLink").Value = ucmbItemCode.Value

            End If
            frmGlazingQuote.isStockItemClosing = True
            frmGlazingQuote.UG2.ActiveRow.PerformAutoSize()
        Catch ex As Exception
            modGlazingQuoteExtension.GQShowMessage(ex.Message, Me.Text, MsgBoxStyle.Critical, "warning", "SaveData")

        End Try
    End Sub

#End Region
#Region "Components events"

    Private Sub btnAdd_Click(sender As Object, e As EventArgs) Handles btnAdd.Click
        SaveData()
    End Sub

    Private Sub ucmbItemCode_Enter(sender As Object, e As EventArgs) Handles ucmbItemCode.Enter
        Dim band As UltraGridBand = ucmbItemCode.DisplayLayout.Bands(0)
        band.ColumnFilters.ClearAllFilters()
        band.ColumnFilters("uiIIItemType").FilterConditions.Add(FilterComparisionOperator.Equals, cmbDDItemType.Value)

    End Sub

    Private Sub ucmbItemCode_RowSelected(sender As Object, e As RowSelectedEventArgs) Handles ucmbItemCode.RowSelected
        CodeHanddler(e)
        EnableFormElements()
        LoadItemImage(e.Row)
    End Sub

    Private Sub ucmbPriceType_RowSelected(sender As Object, e As RowSelectedEventArgs) Handles ucmbPriceType.RowSelected
        If IsNothing(ucmbPriceType.SelectedRow) = False And isLoading = False Then
            frmGlazingQuote.UG2.ActiveRow.Cells("PriceType").Value = ucmbPriceType.Value
            SetPriceOnThisRow(frmGlazingQuote.UG2.ActiveRow)

        End If
    End Sub

    Private Sub ucmbItemDes_RowSelected(sender As Object, e As RowSelectedEventArgs) Handles ucmbItemDes.RowSelected
        If IsNothing(ucmbItemDes.SelectedRow) = False Then
            If comboVlaueChanged = False And IsNothing(ucmbItemDes.Value) = False Then
                FillComboData(e)
                comboVlaueChanged = False

            End If
        End If
        'FillActiveRowFromSelectedProductParameters(ucmbItemCode, frmGlazingQuote.UG2.ActiveRow)

    End Sub

    Private Sub ucmbItemDes_Enter(sender As Object, e As EventArgs) Handles ucmbItemDes.Enter
        Dim band As UltraGridBand = ucmbItemDes.DisplayLayout.Bands(0)
        band.ColumnFilters.ClearAllFilters()
        band.ColumnFilters("uiIIItemType").FilterConditions.Add(FilterComparisionOperator.Equals, cmbDDItemType.Value)
    End Sub

    Private Sub ucmbPriceType_Enter(sender As Object, e As EventArgs) Handles ucmbPriceType.Enter, cmbDDPriceListsTrade.Enter
        Dim band As UltraGridBand = ucmbPriceType.DisplayLayout.Bands(0)
        band.ColumnFilters.ClearAllFilters()
        band.ColumnFilters("ItemType").FilterConditions.Add(FilterComparisionOperator.Equals, cmbDDItemType.Value)
    End Sub

    Private Sub ctxtWidth_AfterExitEditMode(sender As Object, e As EventArgs) Handles ctxtWidth.AfterExitEditMode
        calculateVolume()
        CalculateTotalExc()
    End Sub

    Private Sub txtHeight_AfterExitEditMode(sender As Object, e As EventArgs) Handles txtHeight.AfterExitEditMode
        calculateVolume()
        CalculateTotalExc()
    End Sub

    Private Sub cmbDDPriceListsTrade_RowSelected(sender As Object, e As RowSelectedEventArgs) Handles cmbDDPriceListsTrade.RowSelected
        If isLoading = False Then
            If IsNothing(cmbDDPriceListsTrade.Value) = False Then
                If Not IsNothing(cmbDDPriceListsTrade.Value) Then
                    frmGlazingQuote.UG2.ActiveRow.Cells("PriceList").Value = cmbDDPriceListsTrade.Value
                    SetPriceOnThisRow(frmGlazingQuote.UG2.ActiveRow)
                ElseIf Not IsNothing(cmbDDPriceListsSpecial.Value) Then
                    frmGlazingQuote.UG2.ActiveRow.Cells("PriceList").Value = cmbDDPriceListsSpecial.Value
                    SetPriceOnThisRow(frmGlazingQuote.UG2.ActiveRow)
                End If
            End If
        End If

    End Sub

    Private Sub cmbDDItemType_RowSelected(sender As Object, e As RowSelectedEventArgs) Handles cmbDDItemType.RowSelected
        ucmbItemCode.Enabled = True
        lblItemCode.Enabled = True
        EnableVolumElements()
        ResetFormElementsValues(cmbDDItemType.Name)
        EnableFormElements()
    End Sub

    Private Sub ucmbItemCode_DoubleClick(sender As Object, e As EventArgs) Handles ucmbItemCode.DoubleClick
        If IsNothing(cmbDDItemType.ActiveRow) = False Then
            If cmbDDItemType.ActiveRow.Cells("TypeID").Value = 2 Then
                If isLoading = False Then
                    LoadTemplateItems()
                End If
            End If
        End If
    End Sub

    Private Sub ucmbItemCode_MouseDoubleClick(sender As Object, e As MouseEventArgs) Handles ucmbItemCode.MouseDoubleClick
        If IsNothing(cmbDDItemType.ActiveRow) = False Then
            If cmbDDItemType.ActiveRow.Cells("TypeID").Value = 2 Then
                If isLoading = False Then
                    LoadTemplateItems()
                End If
            End If
        End If
    End Sub

    Private Sub btnEditTemplateItems_Click(sender As Object, e As EventArgs) Handles btnEditTemplateItems.Click
        Try
            If IsNothing(cmbDDItemType.ActiveRow) = False Then
                If cmbDDItemType.ActiveRow.Cells("TypeID").Value = 2 Then
                    If isLoading = False Then
                        LoadTemplateItems()
                    End If
                End If
            End If
        Catch ex As Exception
            modGlazingQuoteExtension.GQShowMessage(ex.Message, Me.Text, MsgBoxStyle.Critical, "warning", GetCurrentMethod.Name)
        End Try
    End Sub

    Private Sub txtPrice_ValueChanged(sender As Object, e As EventArgs) Handles txtPrice.ValueChanged
        Try
            If Me.txtPrice.IsInEditMode = True And cmbDDItemType.ActiveRow.Cells("TypeID").Value = 2 Or Me.txtPrice.IsInEditMode = True And cmbDDItemType.ActiveRow.Cells("TypeID").Value = 1 Then
                CalculateTotalExc()
            End If
        Catch ex As Exception
            modGlazingQuoteExtension.GQShowMessage(ex.Message, Me.Text, MsgBoxStyle.Critical, "warning", GetCurrentMethod.Name)
        End Try
    End Sub

    Private Sub txtQty_ValueChanged(sender As Object, e As EventArgs) Handles txtQty.ValueChanged
        Try
            If Me.txtQty.IsInEditMode = True Then
                CalculateTotalExc()
            End If
        Catch ex As Exception
            modGlazingQuoteExtension.GQShowMessage(ex.Message, Me.Text, MsgBoxStyle.Critical, "warning", GetCurrentMethod.Name)

        End Try
    End Sub

    Private Sub btnGlassServices_Click(sender As Object, e As EventArgs) Handles btnGlassServices.Click
        Try
            Dim frmGlazingDocStockItemServicesObj As New frmGlazingDocStockItemServices
            frmGlazingDocStockItemServicesObj.mainStockItem = If(IsNothing(Me.ucmbItemCode.Value) = False, Me.ucmbItemCode.Value, 0)

            If withFractions = False Then
                frmGlazingDocStockItemServicesObj.mainStockItemHeight = If(IsNothing(Me.txtHeight.Value) = False, Me.txtHeight.Value, 0)
                frmGlazingDocStockItemServicesObj.mainStockItemWidth = If(IsNothing(Me.ctxtWidth.Value) = False, Me.ctxtWidth.Value, 0)
            Else
                frmGlazingQuote.UG2.ActiveRow.Cells("Dis.Height").Value = utxtHeightWIthFractions.Value
                frmGlazingQuote.UG2.ActiveRow.Cells("Dis.Width").Value = utxiWidthWIthFractions.Value
                frmGlazingDocStockItemServicesObj.mainStockItemHeight = If(IsNothing(Me.utxtHeightWIthFractions.Value) = False, Me.utxtHeightWIthFractions.Value, 0)
                frmGlazingDocStockItemServicesObj.mainStockItemWidth = If(IsNothing(Me.utxiWidthWIthFractions.Value) = False, Me.utxiWidthWIthFractions.Value, 0)
            End If

            frmGlazingDocStockItemServicesObj.zipCode = Me.frmGlazingQuote.txtPhyPostCode.Value
            frmGlazingDocStockItemServicesObj.dblThickness = ucmbItemCode.SelectedRow.Cells("ufIIThickness").Value

            frmGlazingDocStockItemServicesObj.ug2Row = frmGlazingQuote.UG2.ActiveRow
            frmGlazingDocStockItemServicesObj.isTempered = If((ucmbPriceType.Value), True, False)
            frmGlazingDocStockItemServicesObj.cusID = frmGlazingQuote.cmbAccount.Value

            frmGlazingDocStockItemServicesObj.selectedItems = lineSubItems 'frmGlazingQuote.UG2.ActiveRow.Cells("templateData").Value
            frmGlazingDocStockItemServicesObj.selectedItemsDisplay = lineComments 'frmGlazingQuote.UG2.ActiveRow.Cells("LineComments").Value
            frmGlazingDocStockItemServicesObj.totalAmount = templateItemSubItemsPrice 'frmGlazingQuote.UG2.ActiveRow.Cells("ServiceGross").Value
            frmGlazingDocStockItemServicesObj.cusID = frmGlazingQuote.cmbAccount.Value
            frmGlazingDocStockItemServicesObj.mainGlassQty = txtQty.Value

            'Fill Tax code and Tax Rate
            frmGlazingDocStockItemServicesObj.dCusTaxCode = frmGlazingQuote.dCusTaxCode
            frmGlazingDocStockItemServicesObj.dCusTaxRate = frmGlazingQuote.dCusTaxRate

            frmGlazingDocStockItemServicesObj.glassHeight = txtHeight.Value
            frmGlazingDocStockItemServicesObj.glassWidth = ctxtWidth.Value
            frmGlazingDocStockItemServicesObj.glassVolume = txtVolume.Value

            frmGlazingDocStockItemServicesObj.ShowDialog()

            'Get services data
            lineSubItemsNew = frmGlazingDocStockItemServicesObj.selectedItems
            lineCommentsNew = If(IsNothing(utxtDocDes.Text) = False, utxtDocDes.Text, "") & vbCrLf & frmGlazingDocStockItemServicesObj.selectedItemsDisplay
            templateItemSubItemsPriceNew = frmGlazingDocStockItemServicesObj.totalAmount

            If frmGlazingDocStockItemServicesObj.isDiscard = True Then
                templateItemSubItemsPriceNew = lineSubItemsTotalExc
                lineCommentsNew = ""
                'templateItemSubItems = ""
                isSubItemsDiscards = True
                frmGlazingDocStockItemServicesObj.isDiscard = False
            Else
                CalculateTotalPrice()
                isSubItemsDiscards = False
            End If

            lineSubItems = lineSubItemsNew
            lineComments = lineCommentsNew
            templateItemSubItemsPrice = templateItemSubItemsPriceNew

            CalculateTotalExc()

        Catch ex As Exception
            modGlazingQuoteExtension.GQShowMessage(ex.Message, Me.Text, MsgBoxStyle.Critical, "warning", GetCurrentMethod.Name)

        End Try
    End Sub

#End Region
#Region "Components Funtinality"

    Public Sub CodeHanddler(e As RowSelectedEventArgs)
        Try
            If IsNothing(ucmbItemCode.SelectedRow) = False Then

                If comboVlaueChanged = False Then
                    If isLoading = False Then
                        ResetFormElementsValues(ucmbItemCode.Name)
                    End If

                    FillComboData(e)

                    'If e.Row.Cells("TaxExempt").Value = True Then
                    '    frmGlazingQuote.UG2.ActiveRow.Cells("TaxRate").Value = 0
                    '    frmGlazingQuote.UG2.ActiveRow.Cells("TaxRateValue").Value = 0

                    'Else
                    frmGlazingQuote.UG2.ActiveRow.Cells("TaxRate").Value = frmGlazingQuote.defaultTaxtRateValue
                    frmGlazingQuote.UG2.ActiveRow.Cells("TaxRateValue").Value = frmGlazingQuote.defaultTaxtRateValue

                    'End If
                    If e.Row.Cells("uiIIItemType").Value = 2 Then
                        If isLoading = False Then
                            isTemplateItemLoading = True
                            LoadTemplateItems(e)
                        End If
                    End If
                    comboVlaueChanged = False
                End If

                If isLoading = False Then
                    FillActiveRowFromSelectedProductParameters(ucmbItemCode, frmGlazingQuote.UG2.ActiveRow)
                    SetPriceOnThisRow(frmGlazingQuote.UG2.ActiveRow)

                    If isTemplateItemLoading = True Then
                        txtPrice.Value = templateItemSubItemsPrice
                        isTemplateItemLoading = False
                    End If

                End If
            End If
        Catch ex As Exception

        End Try
    End Sub

    Sub FillComboData(e As RowSelectedEventArgs)
        Try
            If e.Owner.Name = "ucmbItemCode" Then
                comboVlaueChanged = True
                ucmbItemDes.Value = ucmbItemCode.Value
                'ucmbPriceType.Value = ucmbItemCode.SelectedRow.Cells("uiIIPRICETYPEID").Value
                utxtDocDes.Value = ucmbItemCode.SelectedRow.Cells("Description_3").Value

            ElseIf e.Owner.Name = "ucmbPriceType" Then
                comboVlaueChanged = True
                ucmbItemCode.Value = ucmbPriceType.SelectedRow.Cells("Description_1").Value
                ucmbItemDes.Value = ucmbPriceType.SelectedRow.Cells("uiIIPRICETYPEID").Value
                utxtDocDes.Value = ucmbPriceType.SelectedRow.Cells("Description_3").Value

            ElseIf e.Owner.Name = "ucmbItemDes" Then
                comboVlaueChanged = True
                ucmbItemCode.Value = ucmbItemDes.Value
                'ucmbItemDes.Value = ucmbItemDes.SelectedRow.Cells("uiIIPRICETYPEID").Value
                utxtDocDes.Value = ucmbItemCode.SelectedRow.Cells("Description_3").Value
            End If
            'Dim emptyRow As UltraGridRow
            'emptyRow.Cells("ItemType").Value = ucmbPriceType.Value
            'emptyRow.Cells("StockLink").Value = ucmbItemCode.SelectedRow.Cells("StockLink").Value
            'emptyRow.Cells("Description_1").Value = ucmbItemCode.SelectedRow.Cells("Description_1").Value
            'emptyRow.Cells("Price_Type").Value = ucmbItemCode.SelectedRow.Cells("uiIIPRICETYPEID").Value
            'emptyRow.Cells("IsPriceItem").Value = True

        Catch ex As Exception
            modGlazingQuoteExtension.GQShowMessage("Data not saved", Me.Text, MsgBoxStyle.Critical)

        End Try
    End Sub

    Public Function VolumCalculator(ByRef hight As Double, ByRef width As Double) As Double
        Try
            Dim Volume As Double = 0.0
            If IsNothing(hight) = False And IsNothing(width) = False Then
                Volume = Math.Round((hight * width) / 1000000, 2, MidpointRounding.AwayFromZero)
                Return Volume
            End If
        Catch ex As Exception
            GQShowMessage(ex.Message, Me.Text, MsgBoxStyle.Critical)
            Return 0
        End Try
    End Function

    Private Sub FillActiveRowFromSelectedProductParameters(ByRef uddSource As UltraCombo, ugRow As UltraGridRow)
        Try
            If uddSource.ActiveRow IsNot Nothing Then

                If uddSource.Name = "ucmbItemDes" Then
                    ugRow.Cells("SimpleCode").Value = uddSource.ActiveRow.Cells("Code").Value
                Else
                    ugRow.Cells("Description1").Value = uddSource.ActiveRow.Cells("Description_1").Value
                End If

                ugRow.Cells("StockLink").Value = uddSource.ActiveRow.Cells("StockLink").Value
                ugRow.Cells("Description2").Value = uddSource.ActiveRow.Cells("Description_3").Value & " " & IIf(IsDBNull(uddSource.ActiveRow.Cells("AddDetails").Value), "", uddSource.ActiveRow.Cells("AddDetails").Value) 'uddSource.ActiveRow.Cells("Description_1").Value

                If ugRow.Band.Index = 0 Then 'Normal Glass Line (Band 0)
                    Select Case cmbDDItemType.Value
                        Case GlassItemTypes.Glass
                            ugRow.Cells("ItemTypeCategory").Value = "G"
                        Case GlassItemTypes.Template
                            ugRow.Cells("ItemTypeCategory").Value = "M"
                        Case GlassItemTypes.Consumable
                            ugRow.Cells("ItemTypeCategory").Value = "C"
                        Case GlassItemTypes.Service
                            ugRow.Cells("ItemTypeCategory").Value = "S"
                        Case GlassItemTypes.Aluminium
                            ugRow.Cells("ItemTypeCategory").Value = "A"
                    End Select
                Else 'Band=1 all template components lines goes as "S" (form OLD Code)
                    ugRow.Cells("ItemTypeCategory").Value = "S"
                End If

                ugRow.Cells("Thickness").Value = uddSource.ActiveRow.Cells("ufIIThickness").Value
                ''  ugRow.Cells("TaxCode").Value = uddSource.ActiveRow.Cells("TTI").Value
                'ugRow.Cells("TaxRate").Value = uddSource.ActiveRow.Cells("TaxRate").Value


                ugRow.Cells("Price").Value = 0

                ugRow.Cells("PriceCat").Value = DBNull.Value
                ugRow.Cells("PriceList").Value = DBNull.Value

                'Set Item Price Type
                If ugRow.Band.Index = 0 Then 'Normal Glass Line (Band 0)
                    If uddSource.ActiveRow.Cells("uiIIPRICETYPEID").Value > 0 Then
                        ugRow.Cells("PriceType").Value = uddSource.ActiveRow.Cells("uiIIPRICETYPEID").Value
                        Dim oPriceType As SpilCommon.PriceType
                        oPriceType = New SpilCommon.PriceType(GlassItemTypes.Glass, uddSource.ActiveRow.Cells("uiIIPRICETYPEID").Value)
                        ugRow.Cells("Toughened").Value = oPriceType.IsToughened
                        oPriceType = Nothing
                    Else
                        For Each uR As UltraGridRow In ucmbPriceType.Rows
                            If uR.Hidden = False Then
                                ugRow.Cells("PriceType").Value = uR.Cells("TYPE_ID").Value
                                Dim oPriceType As SpilCommon.PriceType
                                oPriceType = New SpilCommon.PriceType(GlassItemTypes.Glass, uR.Cells("TYPE_ID").Value)
                                ugRow.Cells("Toughened").Value = oPriceType.IsToughened
                                oPriceType = Nothing
                                Exit For
                            End If
                        Next
                    End If

                Else 'Tempalte Item Band(1)
                    'Template Glass Line (Band 1) So price type has to read from template master setup

                    Dim iPriceTypeID As Integer = oPriceUnits.GetPriceTypeIDforTemplateSubItems(ugRow.ParentRow, ugRow)

                    If iPriceTypeID > 0 Then
                        ugRow.Cells("PriceType").Value = iPriceTypeID
                        Dim oPriceType As SpilCommon.PriceType
                        oPriceType = New SpilCommon.PriceType(GlassItemTypes.Glass, ugRow.Cells("PriceType").Value)
                        ugRow.Cells("Toughened").Value = oPriceType.IsToughened
                        oPriceType = Nothing
                    Else
                        'For Each uR As UltraGridRow In cmbDDPriceType.Rows
                        '    If uR.Hidden = False Then
                        '        ugRow.Cells("PriceType").Value = uR.Cells("TYPE_ID").Value
                        '        Dim oPriceType As SpilCommon.PriceType
                        '        oPriceType = New SpilCommon.PriceType(GlassItemTypes.Glass, uR.Cells("TYPE_ID").Value)
                        '        ugRow.Cells("Toughened").Value = oPriceType.IsToughened
                        '        oPriceType = Nothing
                        '        Exit For
                        '    End If
                        'Next
                    End If

                    'Added as NZ requirement 15-05-2015
                    'If bAutoCreateTempHeader Then
                    '    'Call CreateTemplateHeaderDescription(ugRow.ParentRow)
                    'End If
                    'Added as NZ requirement 15-05-2015

                End If

                'Set Item Price Type
                ugRow.Cells("Measure").Value = IIf(uddSource.ActiveRow.Cells("uiIISRVPRICEID").Value = 0, 1, uddSource.ActiveRow.Cells("uiIISRVPRICEID").Value)
                If ugRow.Cells("Measure").Value <> GlassServiceUnits.Lineal Then '3 Then
                    ugRow.Cells("Method").Activation = Activation.Disabled
                Else
                    ugRow.Cells("Method").Activation = Activation.AllowEdit
                End If

                'this is for Aluminium item which price calculation by SQM 
                If ugRow.Cells("Measure").Value = GlassServiceUnits.Area Then
                    ugRow.Cells("Width").Activation = Activation.AllowEdit
                    ugRow.Cells("Height").Activation = Activation.AllowEdit
                    ugRow.Cells("Width").Appearance.BackColor = Color.White
                    ugRow.Cells("Height").Appearance.BackColor = Color.White
                Else
                    ugRow.Cells("Width").Activation = Activation.Disabled
                    ugRow.Cells("Height").Activation = Activation.Disabled
                    ugRow.Cells("Width").Appearance.BackColor = Color.LightGray
                    ugRow.Cells("Height").Appearance.BackColor = Color.LightGray
                End If
                'this is for Aluminium item which price calculation by SQM 

            Else
                ugRow.Cells("StockLink").Value = DBNull.Value
                ugRow.Cells("Description1").Value = String.Empty
                ugRow.Cells("Description2").Value = String.Empty
                ugRow.Cells("SimpleCode").Value = String.Empty
                ugRow.Cells("Thickness").Value = DBNull.Value
                ugRow.Cells("TaxCode").Value = String.Empty
                ugRow.Cells("TaxRate").Value = DBNull.Value
                ugRow.Cells("Price").Value = 0
                ugRow.Cells("PriceType").Value = DBNull.Value
                ugRow.Cells("Measure").Value = DBNull.Value
                ugRow.Cells("PriceCat").Value = DBNull.Value
                ugRow.Cells("PriceList").Value = DBNull.Value

            End If
            ucmbPriceType.Value = ugRow.Cells("PriceType").Value
        Catch ex As Exception
            modGlazingQuoteExtension.GQShowMessage(ex.Message, Me.Text, MsgBoxStyle.Critical, "warning", GetCurrentMethod.Name)
        End Try

    End Sub

    Sub EnableFormElements()
        Try
            Dim isItemSelected As Boolean = False
            If IsNothing(ucmbItemCode.Value) = False Then
                'If ucmbItemCode.Value = "" Then
                isItemSelected = True
                'End If
            End If

            lblItemDescription.Enabled = True
            ucmbItemDes.Enabled = True

            lblDocDescription.Enabled = True
            utxtDocDes.Enabled = True

            ucmbPriceType.Enabled = True
            lblPriceType.Enabled = True

            cmbDDPriceListsTrade.Enabled = True
            lblPriceList.Enabled = True
            cmbDDPriceListsSpecial.Enabled = True

            txtQty.Enabled = True
            lblQty.Enabled = True

            txtPrice.Enabled = True
            lblPrice.Enabled = True
            EnableVolumElements()
            If cmbDDItemType.Value = 2 Or cmbDDItemType.Value = 1 Or cmbDDItemType.Value = 3 Then
                unelblTotal.Visible = True
                lblTotal.Visible = True
                If isItemSelected = True Then
                    If cmbDDItemType.Value = 2 Then
                        btnEditTemplateItems.Visible = True
                        btnGlassServices.Visible = False
                    Else
                        btnGlassServices.Visible = True
                        btnEditTemplateItems.Visible = False
                    End If
                End If

                'ElseIf cmbDDItemType.Value = 1 And isItemSelected = True Then
                '    lblTotal.Visible = False
                '    unelblTotal.Visible = False

            Else
                If isItemSelected = True Then
                    btnEditTemplateItems.Visible = False
                    btnGlassServices.Visible = False
                End If
                lblTotal.Visible = False
                unelblTotal.Visible = False

            End If
            ''check
            'btnGlassServices.Visible = True


        Catch ex As Exception
            modGlazingQuoteExtension.GQShowMessage(ex.Message, Me.Text, MsgBoxStyle.Critical, "warning", GetCurrentMethod.Name)

        End Try
    End Sub

    Sub EnableVolumElements()
        Try
            Dim enableVisibility As Boolean = False
            If IsNothing(ucmbItemCode.Value) = False Then

                If Me.cmbDDItemType.Value = 3 And ucmbItemCode.Value > 0 Then
                    enableVisibility = False
                Else
                    enableVisibility = True
                End If

                If withFractions = True Then
                    utxtHeightWIthFractions.Enabled = enableVisibility
                    utxiWidthWIthFractions.Enabled = enableVisibility
                    utxtHeightWIthFractions.Value = 0
                    utxiWidthWIthFractions.Value = 0
                Else
                    txtHeight.Enabled = enableVisibility
                    ctxtWidth.Enabled = enableVisibility
                    txtHeight.Value = 0
                    ctxtWidth.Value = 0
                End If

                lblWidth.Enabled = enableVisibility
                lblHeight.Enabled = enableVisibility
                txtVolume.Enabled = enableVisibility
                lblVolume.Enabled = enableVisibility
                txtVolume.Value = 0
            End If

            enableVisibility = True
            If withFractions = True Then
                utxtHeightWIthFractions.Visible = enableVisibility
                utxiWidthWIthFractions.Visible = enableVisibility
                txtHeight.Visible = Not enableVisibility
                ctxtWidth.Visible = Not enableVisibility
            Else
                txtHeight.Visible = enableVisibility
                ctxtWidth.Visible = enableVisibility
                utxtHeightWIthFractions.Visible = Not enableVisibility
                utxiWidthWIthFractions.Visible = Not enableVisibility
            End If


        Catch ex As Exception

        End Try
    End Sub

    Sub ResetFormElementsValues(ByRef controllerName As String)
        If isLoading = False Then
            If controllerName <> "ucmbItemCode" Then
                ucmbItemCode.Value = Nothing
            End If
            ucmbItemDes.Value = Nothing
            utxtDocDes.Value = Nothing
            ucmbPriceType.Value = Nothing
            cmbDDPriceListsTrade.Value = Nothing
            cmbDDPriceListsSpecial.Value = Nothing
            txtPrice.Value = 0
            txtHeight.Value = 0
            ctxtWidth.Value = 0
            txtVolume.Value = 0
            txtQty.Value = 0
            uPicBox.Image = Nothing
            frmGlazingQuoteCellValuesClear()
        End If
    End Sub

    Public Function frmGlazingQuoteCellValuesClear(Optional ByRef gridRow As UltraGridRow = Nothing) As UltraGridRow
        If IsNothing(gridRow) = False Then
            selectedRow = gridRow
        End If
        If isLoading = False Then
            selectedRow.Cells("LineComments").Value = ""
            selectedRow.Cells("Qty").Value = 0
            selectedRow.Cells("Width").Value = 0
            selectedRow.Cells("Height").Value = 0
            selectedRow.Cells("Price").Value = 0
            selectedRow.Cells("ItemType").Value = 0
            selectedRow.Cells("Description1").Value = ""
            selectedRow.Cells("SimpleCode").Value = ""

            selectedRow.Cells("Price_Type").Value = 0
            selectedRow.Cells("StockLink").Value = 0
            selectedRow.Cells("templateData").Value = ""
            selectedRow.Cells("PriceCat").Value = 0
            selectedRow.Cells("PriceList").Value = DBNull.Value
            selectedRow.Cells("IsPriceItem").Value = True

        End If
        Return selectedRow
    End Function

    Sub LoadItemImage(row As UltraGridRow)
        If IsNothing(row) = False Then
            If IsDBNull(row.Cells("ItemImage").Value) = False Then
                ItemImage = row.Cells("ItemImage").Value
                uPicBox.Image = frmGlazingQuote.ByteToImage(row.Cells("ItemImage").Value)
            End If
        End If
    End Sub

    Sub LoadTemplateItems(Optional e As RowSelectedEventArgs = Nothing)
        Dim slectedRow As UltraGridRow
        If IsNothing(e) = True Then
            slectedRow = ucmbItemCode.ActiveRow
        Else
            slectedRow = e.Row
        End If
        Try
            'Dim oItem As New SPIL.Glass.InventoryItem()
            'Dim items As String = ""
            ''Item = oItem.GetItem(CInt(UgMR.Cells("StockLink").Value))
            'Dim clsOrder As New clsOrderLineDetails
            'Dim clsInvDetLine As New clsInvDetailLine
            'Dim List2 As New List(Of clsInvDetailLine)
            'If IsNothing(selectedRow.Cells("templateData").Value) = False And itemIsChanging = False Then
            '    Dim itemLines() As String = selectedRow.Cells("templateData").Value.Split(";")
            '    For Each itemDetails As String In itemLines
            '        clsInvDetLine = New clsInvDetailLine
            '        If itemDetails = "" Then
            '            Exit For
            '        End If
            '        Dim itemDetailsCol() As String = itemDetails.Split(",")
            '        clsInvDetLine.StockLink = itemDetailsCol(0)
            '        clsInvDetLine.M_NO = 0
            '        'clsInvDetLine.cSimpleCode = itemDetailsCol(2)
            '        'clsInvDetLine.Description_1 = itemDetailsCol(3)
            '        clsInvDetLine.ItemType = itemDetailsCol(1)

            '        List2.Add(clsInvDetLine)
            '    Next
            '    'Dim sSPILEDICodes3() As String = sSPILEDICodes("templateData").Value.Split(";")
            '    clsOrder.InvDetailLinesList = List2

            '    'clsOrder.InvDetailLinesList = selectedRow.Cells("templateData").Value
            'Else
            'End If
            'clsInvDetLine.StockLink = slectedRow.Cells("StockLink").Value
            'clsInvDetLine.ItemType = GlassItemTypes.Template
            'clsInvDetLine.M_NO = 0
            'clsInvDetLine.cSimpleCode = slectedRow.Cells("Code").Value
            'clsInvDetLine.Description_1 = slectedRow.Cells("Description_1").Value
            '' clsInvDetLine.MainItem = MainItem

            'Dim List As New List(Of clsInvDetailLine)
            'List.Add(clsInvDetLine)
            'clsOrder.InvDetailLinesList = List

            'If clsOrder.InvDetailLinesList.Count > 0 Then
            'Else
            'End If
            Dim frmGlazingDocStockItemTemplateObj As New frmGlazingDocStockItemTemplate(oPriceUnits)
            frmGlazingDocStockItemTemplateObj.stockLink = slectedRow.Cells("StockLink").Value
            frmGlazingDocStockItemTemplateObj.selectedItems = frmGlazingQuote.UG2.ActiveRow.Cells("templateData").Value
            frmGlazingDocStockItemTemplateObj.selectedItemsDisplay = frmGlazingQuote.UG2.ActiveRow.Cells("LineComments").Value
            frmGlazingDocStockItemTemplateObj.totalAmount = frmGlazingQuote.UG2.ActiveRow.Cells("Price").Value

            Dim Resullt As DialogResult = frmGlazingDocStockItemTemplateObj.ShowDialog()
            templateItemSubItemsPrice = frmGlazingDocStockItemTemplateObj.totalAmount
            Dim templateItemSubItems = frmGlazingDocStockItemTemplateObj.selectedItems
            Dim lineComments = frmGlazingDocStockItemTemplateObj.selectedItemsDisplay
            items = utxtDocDes.Text
            selectedRow.Cells("templateData").Value = templateItemSubItems

            If frmGlazingDocStockItemTemplateObj.isDiscard = True Then
                selectedRow.Cells("LineComments").Value = lineComments
                frmGlazingDocStockItemTemplateObj.isDiscard = False
            Else
                selectedRow.Cells("LineComments").Value = items & vbCrLf & lineComments
            End If

            txtPrice.Value = templateItemSubItemsPrice

        Catch ex As Exception
            modGlazingQuoteExtension.GQShowMessage(ex.Message, Me.Text, MsgBoxStyle.Critical, "warning", GetCurrentMethod.Name)

        End Try
    End Sub

    Sub LanguageHandler(Optional ByRef region As Integer = 1)
        Try
            If region = GeographicRegion.NorthAmerica Then
                lblPrice.Text = "Sales Price"
            Else
                lblPrice.Text = "Price"
            End If
        Catch ex As Exception
            modGlazingQuoteExtension.GQShowMessage(ex.Message, Me.Text, MsgBoxStyle.Critical, "warning", GetCurrentMethod.Name)
        End Try
    End Sub

    Sub BackupOldData()
        Try
            lineComments = frmGlazingQuote.UG2.ActiveRow.Cells("LineComments").Value
            lineSubItems = frmGlazingQuote.UG2.ActiveRow.Cells("templateData").Value
            templateItemSubItemsPrice = frmGlazingQuote.UG2.ActiveRow.Cells("ServiceGross").Value

            lineSubItemsTotalExc = frmGlazingQuote.UG2.ActiveRow.Cells("ServiceGross").Value
            lineSubItemsTotalInc = frmGlazingQuote.UG2.ActiveRow.Cells("ServiceItemTotNet").Value
            lineSubItemsTotalTax = frmGlazingQuote.UG2.ActiveRow.Cells("ServiceItemTax").Value

        Catch ex As Exception
            modGlazingQuoteExtension.GQShowMessage(ex.Message, Me.Text, MsgBoxStyle.Critical, "warning", GetCurrentMethod.Name)

        End Try
    End Sub

    Function CalculateTotalExc(Optional ByRef addPrice As Boolean = False, Optional ByRef serviceExc As Double = 0) As Double
        Try
            Dim TotalExc As Double = unelblTotal.Value
            Dim qty As Double = txtQty.Text
            Dim volume As Double = txtVolume.Text
            Dim actualEnterdPrice As Double = Me.txtPrice.Value
            Dim quantity As Double = txtQty.Value = 0
            Dim isConsumable As Boolean

            If cmbDDItemType.Value = 3 Then
                isConsumable = True
            Else
                isConsumable = False
            End If
            'Get total
            If frmGlazingQuote.UG2.ActiveRow.Cells("ServiceItemTotNet").Value > 0 Or lineSubItemsTotalExc > 0 Or lineCommentsNew <> "" Then
                CalculateTotalPrice()
                TotalExc = totalExclusiveAmountNew
            Else
                TotalExc = clsGlazingDocStockItemHelperObj.CalculateTotalExc(TotalExc, qty, volume, actualEnterdPrice, quantity, serviceExc, 0, isConsumable)
            End If

            'Set total
            If addPrice = False Then
                unelblTotal.Value = TotalExc
            Else
                unelblTotal.Value = TotalExc + txtPrice.Value
            End If

        Catch ex As Exception
            modGlazingQuoteExtension.GQShowMessage(ex.Message, Me.Text, MsgBoxStyle.Critical, "warning", GetCurrentMethod.Name)

        End Try
    End Function

    Function CreateLineComments() As String
        Try
            Dim itemLines() As String = lineComments.Split(vbCrLf)
            Dim counter As Integer = 1
            For Each itemDetails As String In itemLines
                If itemDetails = "" Or counter = 1 Or itemDetails = vbLf Then
                    counter = counter + 1
                    Continue For
                End If

                If lineSubItemsToShow = "" Then
                    lineSubItemsToShow = itemDetails.Replace(vbLf, "")
                Else
                    lineSubItemsToShow = lineSubItemsToShow & itemDetails
                End If
            Next
            Return lineSubItemsToShow
        Catch ex As Exception
            modGlazingQuoteExtension.GQShowMessage(ex.Message, Me.Text, MsgBoxStyle.Critical, "warning", "CreateLineComments")
            Return lineSubItemsToShow
        End Try
    End Function

#End Region
#Region "Calculations"

    Sub calculateVolume()
        Try
            Dim measurementType As Integer
            txtVolume.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double
            If measurementsTypeOrigin <> measurementsTypeTemp Then
                measurementType = measurementsTypeTemp
            Else
                measurementType = measurementsTypeOrigin
            End If

            If measurementType = MeasurementsType.Imperial Then
                'Imperial
                txtVolume.Value = VolumCalculator(txtHeight.Text, ctxtWidth.Text)
            Else
                'Metric
                'txtVolume.Value = VolumCalculator(txtHeight.Text, txtHeight.Text)
            End If

            If IsNothing(Me.txtVolume.Text) = False Then
                frmGlazingQuote.UG2.ActiveRow.Cells("Volume").Value = txtVolume.Value
            Else
                GQShowMessage("Please check the item measurement values.", Me.Text, MsgBoxStyle.Exclamation)
            End If

        Catch ex As Exception
            GQShowMessage(ex.Message, Me.Text, MsgBoxStyle.Critical)
        End Try
    End Sub

    Public Sub SetPriceOnThisRow(ByRef ugRow As UltraGridRow)

        If IsNothing(oSOModuleDefaults) = True Then
            moduleDefaultsObj = oSOModuleDefaults
        End If

        oPriceUnits.oCustomer = New clsCustomer(frmGlazingQuote.cmbAccount.Value)
        oPriceUnits.iDefaultStockPriceListID = moduleDefaultsObj.DefaultTradePriceListID ' iFormDefaultTradePriceListID
        iFormDefaultTradePriceListID = moduleDefaultsObj.DefaultTradePriceListID
        If ugRow.Cells("IsPriceItem").Value = False Then
            oPriceUnits.Set_PriceList_OnActiveRow_NotRelatingToPriceCalc(ugRow)
        Else
            oPriceUnits.GetStockPriceOnActiveRow(ugRow)
        End If

        If ugRow.Cells("PriceCat").Value = "T" Then
            'ugRow.Cells("PriceList").EditorComponent = cmbDDPriceListsTrade
            cmbDDPriceListsTrade.Visible = True
            cmbDDPriceListsTrade.Value = ugRow.Cells("PriceList").Value
        Else
            ugRow.Cells("PriceList").EditorComponent = cmbDDPriceListsSpecial
        End If
        txtPrice.Value = frmGlazingQuote.UG2.ActiveRow.Cells("Price").Value

    End Sub

    Private Function FindThicknessID(ByVal dThickness As Double) As Double

        Try
            Dim a As Integer = 0
            Dim Dr As UltraGridRow

            For Each Dr In DDThickness.Rows
                If CType(Dr.Cells("THICKNESS_UPTO").Value, Double) >= CType(dThickness, Double) Then
                    a = Dr.Cells("THICKNESS_ID").Value
                    Exit For
                End If
            Next

            'If CType(dblntThik, Double) <= CType(MyDataRow(1), Double) Then
            '    a = MyDataRow(0)
            '    Exit For
            'End If

            '>>>>. ERRORRRRR found at MSG
            ''For Each Dr In DDThickness.Rows
            ''    If CType(Dr.Cells("THICKNESS_UPTO").Value, Int16) <= CType(intThik, Int16) Then
            ''        a = Dr.Cells("THICKNESS_ID").Value
            ''        Exit For
            ''    End If
            ''Next
            If a = 0 Then a = 1
            intThick_ID = a
            Return a
        Catch ex As Exception
            modGlazingQuoteExtension.GQShowMessage(ex.Message, Me.Text, MsgBoxStyle.Critical, "warning", GetCurrentMethod.Name)
        End Try
    End Function

    Public Sub SetThikness()
        SQL = "SELECT SRVPRICEID as ID, SRVPRICE_DESCRIPTION as Pricing FROM stkSRVPRICE_TYPESspil"
        SQL += " SELECT *  FROM stkSRVPRICE_TYPE_CATEGORYSspil"
        SQL += " SELECT THICKNESS_ID, THICKNESS_UPTO FROM  stkTHICKNESS_UPTOspil  ORDER BY THICKNESS_ID"
        'SQL += " SELECT SRVPRICEID as ID, SRVPRICE_DESCRIPTION as Measure FROM stkSRVPRICE_TYPESspil"

        Dim objSQL As New clsSqlConn
        With objSQL
            Try
                DS = New DataSet
                DS = .GET_DATA_SQL(SQL)

                ''DDPriceType.DataSource = Nothing
                ''DDPriceType.DataSource = DS.Tables(0)
                ''DDPriceType.ValueMember = "ID"
                ''DDPriceType.DisplayMember = "ID"

                ''DDEdging.DataSource = Nothing
                ''DDEdging.DataSource = DS.Tables(1)
                ''DDEdging.ValueMember = "SRVPRICE_CATID"
                'DDEdging.DisplayMember = "SRVPRICE_CATID"

                DDThickness.DataSource = Nothing
                DDThickness.DataSource = DS.Tables(2)

                'DDMeasure.DataSource = Nothing
                'DDMeasure.DataSource = DS.Tables(3)
                'DDMeasure.ValueMember = "ID"
                'DDMeasure.DisplayMember = "Measure"
                'DDMeasure.DisplayLayout.Bands(0).Columns("ID").Hidden = True

                'DDMathod.DataSource = Nothing
                'DDMathod.DataSource = DS.Tables(1)
                'DDMathod.ValueMember = "SRVPRICE_CATID"
                'DDMathod.DisplayMember = "SRVPRICE_DESCRIPTION"
                'DDMathod.DisplayLayout.Bands(0).Columns("SRVPRICE_DESCRIPTION").Header.Caption = "Method"
                'DDMathod.DisplayLayout.Bands(0).Columns("SRVPRICEID").Hidden = True
                'DDMathod.DisplayLayout.Bands(0).Columns("SRVPRICE_CATID").Hidden = True
                'DDMathod.DisplayLayout.Bands(0).Columns("QTY_SHORT").Hidden = True
                'DDMathod.DisplayLayout.Bands(0).Columns("QTY_LONG").Hidden = True
                'DDMathod.DisplayLayout.Bands(0).Columns("CUT_ALLOW_SHORT").Hidden = True
                'DDMathod.DisplayLayout.Bands(0).Columns("CUT_ALLOW_LONG").Hidden = True


            Catch ex As Exception
                modGlazingQuoteExtension.GQShowMessage(ex.Message, Me.Text, MsgBoxStyle.Critical, "warning", GetCurrentMethod.Name)
            Finally
                DS.Dispose()
                .Con_Close()
                objSQL = Nothing
            End Try
        End With
    End Sub


    'Calcaulate total price with services
    Private Sub CalculateTotalPrice()
        Try
            Dim glassItemDetails As String
            Dim servicePrice As String
            Dim servicePriceExc As String
            Dim ServiceTaxAmount As String
            Dim servicePriceInc As String
            Dim glassPriceInc As String
            Dim itemLines() As String

            If lineSubItems <> "" Or lineSubItemsNew <> "" Then  'And txtVolume.Value <> glassVolum
                glassItemDetails = txtHeight.Value & "," & ctxtWidth.Value & "," & txtQty.Value & "," & txtPrice.Value
                servicePrice = clsGlazingQuoteExtensionObj.CalculateServiceTotalPrice(If(lineSubItemsNew <> "", lineSubItemsNew, lineSubItems), glassItemDetails, False, False)
                itemLines = servicePrice.Split(";")
                servicePriceExc = itemLines(0)
                totalServiceTaxAmount = itemLines(1)
                servicePriceInc = itemLines(2)
                glassPriceInc = txtPrice.Value
                lineSubItemsTotalExc = servicePriceExc
                lineSubItemsTotalInc = servicePriceInc
                lineSubItemsTotalTax = totalServiceTaxAmount
            End If
            'Get total
            totalExclusiveAmountNew = clsGlazingDocStockItemHelperObj.CalculateTotalExc(TotalExc, txtQty.Value, txtVolume.Value, txtPrice.Value, quantity, servicePriceExc)

            ' glassVolum = txtVolume.Value
        Catch ex As Exception
            modGlazingQuoteExtension.GQShowMessage(ex.Message, Me.Text, MsgBoxStyle.Critical, "warning", "CalculateTotalPrice")

        End Try
    End Sub

    Function CalculateFinalAmount()
        Try

        Catch ex As Exception

        End Try
    End Function

#End Region

End Class