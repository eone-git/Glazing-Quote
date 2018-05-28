Imports System.Data.SqlClient
Imports Infragistics.Win.UltraWinGrid
Imports System.IO
Imports GlassInventoryModule
Imports System.Drawing.Imaging

Public Class frmGlazingDocStockItem
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
    Dim priceTypeValue As Integer =0
    Dim priceListValue As Integer = 0
    Dim itemIsChanging As Boolean = False
    Dim clsGlazingDocStockItemHelperObj As New clsGlazingDocStockItemHelper(Me)
    Dim moduleDefaultsObj As New clsSOModuleDefaults
    Dim templateItemSubItemsPrice As Double = 0
    Dim templateItemSubItemsData As String = ""

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
        'frmGlazingQuote.UG2.DisplayLayout.Clone(PropertyCategories.All)
        ' Add any initialization after the InitializeComponent() call.
        'Dim row As UltraGridRow = frmGlazingQuoteClone.Bands(0).AddNew
        'row.ParentCollection.Move(row, activeRowIndex)
        'frmGlazingQuoteClone.ActiveRowScrollRegion.ScrollRowIntoView(row)
        'Me.UG2.Rows(activeRowIndex).Activate()
        'Dim a2 = frmGlazingQuoteClone.Rows.Count
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

    Sub getDataSource()
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

    Private Sub frmGlazingDocStockItem_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        If IsNothing(selectedRow.Cells("Price_Type").Value) = False Then
            priceTypeValue = selectedRow.Cells("Price_Type").Value

        End If

        If IsNothing(selectedRow.Cells("PriceList").Value) = False And IsDBNull(selectedRow.Cells("PriceList").Value) = False Then
            priceListValue = selectedRow.Cells("PriceList").Value

        End If

        LoadComboData()
        frmGlazingQuote.isStockItemActive = True
        LoadEditMode()

    End Sub

    Sub LoadComboData()
        FillPRICE_TYPES()
        Get_StockItems()
        SetThikness()
        clsGlazingDocStockItemHelperObj.SetPriceListsData("PriceList", "CAT_ID")
        clsGlazingDocStockItemHelperObj.SetItemCodeData("Description_1", "StockLink", "Code")

    End Sub

    Sub LoadEditMode()

        Dim band As UltraGridBand = cmbDDItemType.DisplayLayout.Bands(0)
        band.ColumnFilters.ClearAllFilters()
        band.ColumnFilters("TypeID").FilterConditions.Add(FilterComparisionOperator.Equals, 1)
        band.ColumnFilters("TypeID").LogicalOperator = FilterLogicalOperator.Or
        band.ColumnFilters("TypeID").FilterConditions.Add(FilterComparisionOperator.Equals, 3)
        band.ColumnFilters("TypeID").LogicalOperator = FilterLogicalOperator.Or
        band.ColumnFilters("TypeID").FilterConditions.Add(FilterComparisionOperator.Equals, 2)
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
                txtPrice.Value = frmGlazingQuote.UG2.ActiveRow.Cells("Price").Value

                If IsNothing(selectedRow.Cells("Price_Type").Value) = False Then
                    ucmbPriceType.Value = priceTypeValue

                End If

                If IsNothing(selectedRow.Cells("PriceList").Value) = False Then
                    cmbDDPriceListsSpecial.Value = priceListValue
                    cmbDDPriceListsTrade.Value = priceListValue

                End If

            Else
                cmbDDItemType.Value = 1

            End If
            isLoading = False
        End If
    End Sub


    Private Sub btnAdd_Click(sender As Object, e As EventArgs) Handles btnAdd.Click
        frmGlazingQuote.isStockItemActive = False
        Dim emptyArray As Byte() = New Byte(63) {}

        frmGlazingQuote.UG2.ActiveRow.Cells("LineComments").Value = utxtDocDes.Text
        frmGlazingQuote.UG2.ActiveRow.Cells("Qty").Value = txtQty.Value
        frmGlazingQuote.UG2.ActiveRow.Cells("Width").Value = ctxtWidth.Text
        frmGlazingQuote.UG2.ActiveRow.Cells("Height").Value = txtHeight.Text
        'frmGlazingQuote.UG2.ActiveRow.Cells("Volume").Value = txtVolume.Text
        frmGlazingQuote.UG2.ActiveRow.Cells("Price").Value = txtPrice.Text
        frmGlazingQuote.UG2.ActiveRow.Cells("ItemType").Value = cmbDDItemType.Value
        'frmGlazingQuote.UG2.ActiveRow.Cells("SimpleCode").Value = ctxtWidth.Text
        frmGlazingQuote.UG2.ActiveRow.Cells("Description1").Value = utxtDocDes.Text
        frmGlazingQuote.UG2.ActiveRow.Cells("SimpleCode").Value = ucmbItemCode.Text
        frmGlazingQuote.UG2.ActiveRow.Cells("Price_Type").Value = ucmbPriceType.Value

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
        'Dim image1 As Bitmap
        'image1 = New Bitmap("C:\Documents and Settings\All Users\Documents\My Music\music.bmp", True)
        'frmGlazingQuote.UG2.ActiveRow.Cells("ItemImage").Value = New GlassInventoryModule.frmInventoryItem().ImageToByteArray(uPicBox.Image)
        'Dim arrat2 = New GlassInventoryModule.frmInventoryItem().ImageToByteArray(CType(uPicBox.Image, Bitmap))

        'Dim array As Byte() = New Byte(63) {}
        'array.Clear(array, 0, array.Length)

        'Can be 0
        'If frmGlazingQuote.UG2.ActiveRow.Cells("Price").Value = 0 Then

        '    frmGlazingQuote.UG2.ActiveRow.Cells("ItmExcAmount").Value = 0
        '    frmGlazingQuote.UG2.ActiveRow.Cells("Tax").Value = 0
        '    frmGlazingQuote.UG2.ActiveRow.Cells("Amount").Value = 0
        '    frmGlazingQuote.UG2.ActiveRow.Cells("OrgPrice").Value = 0
        '    frmGlazingQuote.UG2.ActiveRow.Cells("Net").Value = 0

        'End If

        If IsNothing(ucmbItemCode.SelectedRow) = False Then
            frmGlazingQuote.UG2.ActiveRow.Cells("StockLink").Value = ucmbItemCode.Value

        End If
        frmGlazingQuote.isStockItemClosing = True
        frmGlazingQuote.UG2.ActiveRow.PerformAutoSize()

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

    Public Sub CodeHanddler(e As RowSelectedEventArgs)
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
                    e.Row.Cells("").Value = templateItemSubItemsData
                End If

            End If
        End If
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
    End Sub


    Private Sub txtHeight_AfterExitEditMode(sender As Object, e As EventArgs) Handles txtHeight.AfterExitEditMode
        calculateVolume()
    End Sub

    Sub calculateVolume()
        If IsNothing(Me.txtHeight.Text) = False And IsNothing(Me.ctxtWidth.Text) = False Then
            If txtHeight.Text = "" Then
                txtHeight.Text = 0
            End If
            If ctxtWidth.Text = "" Then
                ctxtWidth.Value = 0
            End If

            If IsNothing(Me.txtVolume.Text) = False Then
                txtVolume.Value = (Me.txtHeight.Text * Me.ctxtWidth.Text) / 1000000
                frmGlazingQuote.UG2.ActiveRow.Cells("Volume").Value = txtVolume.Value
            End If
        End If
    End Sub

    Dim oSOModuleDefaults As New clsSOModuleDefaults

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
                MsgBox(ex.Message, MsgBoxStyle.Exclamation, "SPIL Glass")
            Finally
                DS.Dispose()
                .Con_Close()
                objSQL = Nothing
            End Try
        End With
    End Sub

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
            MsgBox(ex.Message)
        End Try

    End Sub

    Private Sub cmbDDPriceListsSpecial_RowSelected(sender As Object, e As RowSelectedEventArgs) Handles cmbDDPriceListsSpecial.RowSelected
        ''  SetPriceOnThisRow(frmGlazingQuote.UG2.ActiveRow)

    End Sub


    Private Sub cmbDDPriceListsTrade_RowSelected(sender As Object, e As RowSelectedEventArgs) Handles cmbDDPriceListsTrade.RowSelected
        '' SetPriceOnThisRow(frmGlazingQuote.UG2.ActiveRow)
        '' FillActiveRowFromSelectedProductParameters(ucmbItemCode, frmGlazingQuote.UG2.ActiveRow)
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

    Private Sub ucmbPriceType_Leave(sender As Object, e As EventArgs) Handles ucmbPriceType.Leave
        'frmGlazingQuote.UG2.ActiveRow.Cells("PriceType").Value = ucmbPriceType.Value
        'FillActiveRowFromSelectedProductParameters(ucmbItemCode, frmGlazingQuote.UG2.ActiveRow)
        'frmGlazingQuote.UG2.ActiveRow.Cells("PriceType").Value = ucmbPriceType.Value
        'SetPriceOnThisRow(frmGlazingQuote.UG2.ActiveRow)

    End Sub

    Private Sub cmbDDPriceListsTrade_Leave(sender As Object, e As EventArgs) Handles cmbDDPriceListsTrade.Leave, cmbDDPriceListsSpecial.Leave
        ' '' FillActiveRowFromSelectedProductParameters(ucmbItemCode, frmGlazingQuote.UG2.ActiveRow)
        'If Not IsNothing(cmbDDPriceListsTrade.Value) Then
        '    frmGlazingQuote.UG2.ActiveRow.Cells("PriceList").Value = cmbDDPriceListsTrade.Value
        '    SetPriceOnThisRow(frmGlazingQuote.UG2.ActiveRow)
        'ElseIf Not IsNothing(cmbDDPriceListsSpecial.Value) Then
        '    frmGlazingQuote.UG2.ActiveRow.Cells("PriceList").Value = cmbDDPriceListsSpecial.Value
        '    SetPriceOnThisRow(frmGlazingQuote.UG2.ActiveRow)
        'End If


    End Sub


    Private Sub frmGlazingDocStockItem_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        frmGlazingQuote.isStockItemActive = False

    End Sub

    Private Sub cmbDDItemType_RowSelected(sender As Object, e As RowSelectedEventArgs) Handles cmbDDItemType.RowSelected
        ucmbItemCode.Enabled = True
        lblItemCode.Enabled = True
        EnableVolumElements()
        ResetFormElementsValues(cmbDDItemType.Name)
    End Sub
    Sub EnableFormElements()

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
    End Sub

    Sub EnableVolumElements()
        If IsNothing(ucmbItemCode.Value) = False Then
            If Me.cmbDDItemType.Value = 3 And ucmbItemCode.Value > 0 Then
                txtHeight.Enabled = False
                lblHeight.Enabled = False
                txtHeight.Value = 0

                ctxtWidth.Enabled = False
                lblWidth.Enabled = False
                ctxtWidth.Value = 0

                txtVolume.Enabled = False
                lblVolume.Enabled = False
                txtVolume.Value = 0

            Else
                txtHeight.Enabled = True
                lblHeight.Enabled = True
                txtHeight.Value = 0

                ctxtWidth.Enabled = True
                lblWidth.Enabled = True
                ctxtWidth.Value = 0

                txtVolume.Enabled = True
                lblVolume.Enabled = True
                txtVolume.Value = 0
            End If
        End If
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

    Private Sub ucmbItemCode_ValueChanged(sender As Object, e As EventArgs) Handles ucmbItemCode.ValueChanged

    End Sub

    Sub LoadTemplateItems(Optional e As RowSelectedEventArgs = Nothing)
        Dim slectedRow As UltraGridRow
        If IsNothing(e) = True Then
            slectedRow = ucmbItemCode.ActiveRow
        Else
            slectedRow = e.Row
        End If
        Try
            Dim oItem As New SPIL.Glass.InventoryItem()
            Dim items As String = ""
            'Item = oItem.GetItem(CInt(UgMR.Cells("StockLink").Value))
            Dim clsOrder As New clsOrderLineDetails
            Dim clsInvDetLine As New clsInvDetailLine
            Dim List2 As New List(Of clsInvDetailLine)
            If IsNothing(selectedRow.Cells("templateData").Value) = False And itemIsChanging = False Then
                Dim itemLines() As String = selectedRow.Cells("templateData").Value.Split(";")
                For Each itemDetails As String In itemLines
                    clsInvDetLine = New clsInvDetailLine
                    If itemDetails = "" Then
                        Exit For
                    End If
                    Dim itemDetailsCol() As String = itemDetails.Split(",")
                    clsInvDetLine.StockLink = itemDetailsCol(0)
                    clsInvDetLine.M_NO = 0
                    clsInvDetLine.cSimpleCode = itemDetailsCol(2)
                    clsInvDetLine.Description_1 = itemDetailsCol(3)
                    clsInvDetLine.ItemType = itemDetailsCol(1)

                    List2.Add(clsInvDetLine)
                Next
                'Dim sSPILEDICodes3() As String = sSPILEDICodes("templateData").Value.Split(";")
                clsOrder.InvDetailLinesList = List2

                'clsOrder.InvDetailLinesList = selectedRow.Cells("templateData").Value
            Else

                clsInvDetLine.StockLink = slectedRow.Cells("StockLink").Value
                clsInvDetLine.ItemType = GlassItemTypes.Template
                clsInvDetLine.M_NO = 0
                clsInvDetLine.cSimpleCode = slectedRow.Cells("Code").Value
                clsInvDetLine.Description_1 = slectedRow.Cells("Description_1").Value
                ' clsInvDetLine.MainItem = MainItem

                Dim List As New List(Of clsInvDetailLine)
                List.Add(clsInvDetLine)
                clsOrder.InvDetailLinesList = List
            End If

            Dim frmGlazingDocStockItemTemplateObj As New frmGlazingDocStockItemTemplate(oPriceUnits)
            frmGlazingDocStockItemTemplateObj.stockLink = slectedRow.Cells("StockLink").Value
            Dim Resullt As DialogResult = frmGlazingDocStockItemTemplateObj.ShowDialog()
            templateItemSubItemsPrice = frmGlazingDocStockItemTemplateObj.totalAmount

            lineComments = ""
            For Each lineItem As clsInvDetailLine In clsOrder.InvDetailLinesList
                items += lineItem.StockLink & "," & lineItem.ItemType & "," & lineItem.cSimpleCode & "," & lineItem.Description_1 & ";"
                If lineComments = "" Then
                    lineComments = ucmbItemDes.Text
                    lineComments = lineComments + vbCrLf & Chr(9) & "*" & lineItem.Description_1

                Else
                    lineComments = lineComments + vbCrLf & Chr(9) & "*" & lineItem.Description_1
                End If
            Next
            selectedRow.Cells("templateData").Value = items
            selectedRow.Cells("Price").Value = templateItemSubItemsPrice
        Catch ex As Exception

        End Try
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
End Class