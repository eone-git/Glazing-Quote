Imports Infragistics.Win.UltraWinGrid

Public Class clsGlazingDocStockItemHelper
    Dim moduleName As String = "Stock Item"
    Dim TemplatesItemObj As frmGlazingDocStockItemTemplate
    Dim StockItemObj As frmGlazingDocStockItem
    Dim clsSOModuleDefaultsObj As clsSOModuleDefaults
    Dim clsSOPricingAndUnitsObj As clsSOPricingAndUnits
    Dim clsSqlConnObj As New clsSqlConn
    Dim sqlQuary As String = ""
    Dim newDataSet As DataSet = Nothing

    Public Sub New(ByRef frmGlazingDocStockItemObj As frmGlazingDocStockItem)
        Me.StockItemObj = frmGlazingDocStockItemObj
        'Me.clsSOModuleDefaultsObj = clsSOModuleDefaultsInherit
        'Me.clsSOPricingAndUnitsObj = clsSOPricingAndUnitsInherit

    End Sub

    Public Sub New(ByRef frmGlazingDocStockItemTemplateObj As frmGlazingDocStockItemTemplate)
        Me.TemplatesItemObj = frmGlazingDocStockItemTemplateObj
    End Sub

    Public Sub SetPriceOnThisRow(ByRef ugRow As UltraGridRow)
        clsSOPricingAndUnitsObj.oCustomer = New clsCustomer(frmGlazingQuote.cmbAccount.Value)
        clsSOPricingAndUnitsObj.iDefaultStockPriceListID = clsSOModuleDefaultsObj.DefaultTradePriceListID ' iFormDefaultTradePriceListID
        iFormDefaultTradePriceListID = clsSOModuleDefaultsObj.DefaultTradePriceListID
        If ugRow.Cells("IsPriceItem").Value = False Then
            clsSOPricingAndUnitsObj.Set_PriceList_OnActiveRow_NotRelatingToPriceCalc(ugRow)
        Else
            clsSOPricingAndUnitsObj.GetStockPriceOnActiveRow(ugRow)
        End If

        If ugRow.Cells("PriceCat").Value = "T" Then
            'ugRow.Cells("PriceList").EditorComponent = cmbDDPriceListsTrade
            StockItemObj.cmbDDPriceListsTrade.Visible = True
            StockItemObj.cmbDDPriceListsTrade.Value = ugRow.Cells("PriceList").Value
        Else
            ugRow.Cells("PriceList").EditorComponent = StockItemObj.cmbDDPriceListsSpecial
        End If
        StockItemObj.txtPrice.Value = frmGlazingQuote.UG2.ActiveRow.Cells("Price").Value

    End Sub

    Function GetStkItemDetails(ByRef tempStockLink As Integer) As DataSet
       
        Try
            sqlQuary = "SELECT     StkItem.Description_1, StkItem.cSimpleCode, spilStkTemplateItems.Stock_Sub_Link, StkItem.ufIIThickness, TaxRate.TaxRate, " & _
                "spilStkTemplateItems.UnitPricePercen,  StkItem.uiIISRVPRICEID ,StkItem.uiIITemplPriceID,spilStkTemplateItems.iItemType,spilStkTemplateItems.iLineNo " & _
                ",spilStkTemplateItems.UnitPrice,spilStkTemplateItems.IsPriceItem,spilStkTemplateItems.SRVPRICE_CATID,spilStkTemplateItems.iHeight, " & _
                "spilStkTemplateItems.iWidth,spilStkTemplateItems.Units,spilStkTemplateItems.Motif,spilStkTemplateItems.Comment2, StkItem.TTI, spilStkTemplateItems.PriceTypeID, " & _
                "StkItem_1.uiIITemplPriceID AS TemplatePriceTypeID, StkItem.Description_3,spilStkTemplateItems.fQuantity, StkItem.BOMItem " & _
                "FROM  spilStkTemplateItems INNER JOIN " & _
                "StkItem ON spilStkTemplateItems.Stock_Sub_Link = StkItem.StockLink INNER JOIN " & _
                "TaxRate ON StkItem.TTI = TaxRate.Code INNER JOIN " & _
                "StkItem AS StkItem_1 ON spilStkTemplateItems.StockLink = StkItem_1.StockLink " & _
                "WHERE  (spilStkTemplateItems.StockLink =" & tempStockLink & ") AND spilStkTemplateItems.iItemType<>5 order by spilStkTemplateItems.iLineNo " & _
                "; select BOMItem from StkItem where StockLink =" & tempStockLink & ""

            newDataSet = clsSqlConnObj.GET_DATA_SQL(sqlQuary)
            Return newDataSet

        Catch ex As Exception
            modGlazingQuoteExtension.GQShowMessage(ex.Message, moduleName, MsgBoxStyle.Critical)
            Return newDataSet

        End Try
    End Function

    Public Function GetPriceTypeData() As DataSet
        Try
            sqlQuary = "SELECT TYPE_ID, TYPE_PRICE,TOUGHENED,IsAllocated, IsApplyOnlyTemplate," & _
                        "IsCuttingNeed,IGU_PROCESS, ItemType FROM stkPRICE_TYPESspil WHERE (SEGMENT_ID = 1) ORDER BY SQUENCE"

            newDataSet = clsSqlConnObj.GET_DATA_SQL(sqlQuary)
            Return newDataSet

        Catch ex As Exception
            modGlazingQuoteExtension.GQShowMessage(ex.Message, moduleName, MsgBoxStyle.Critical)
            Return newDataSet

        End Try
    End Function

    Public Sub UIHandler(ByRef dataSetNew As DataSet, ByRef displayMemberVal As String, ByRef valueMemberVal As String, Optional ByRef allowListInput As String = "")
        Dim counter As Integer = 0
        Dim allowList() As String
        Dim columName As String
        Dim isExist As Boolean
        Dim displayMemberValNew = ""

        'Set data
        If IsNothing(displayMemberVal) = False Then
            displayMemberValNew = displayMemberVal

        End If
        allowList = allowListInput.Split(New Char() {";"c})
        If Array.Exists(allowList, Function(s) s = displayMemberValNew) = False Then
            ReDim Preserve allowList(allowList.Length)
            allowList(allowList.Length - 1) = displayMemberVal

        End If

        If IsNothing(StockItemObj) = False Then
            'For StockItem
            StockItemObj.ucmbPriceType.DataSource = dataSetNew.Tables(0)
            StockItemObj.ucmbPriceType.DisplayMember = displayMemberValNew
            StockItemObj.ucmbPriceType.ValueMember = valueMemberVal
            StockItemObj.ucmbPriceType.Refresh()

            'Colums desinger
            StockItemObj.ucmbPriceType.DisplayLayout.Bands(0).ColHeadersVisible = False
            Do While counter < StockItemObj.ucmbPriceType.DisplayLayout.Bands(0).Columns.Count
                columName = StockItemObj.ucmbPriceType.DisplayLayout.Bands(0).Columns(counter).Key
                isExist = Array.Exists(allowList, Function(s) s = columName)
                If IsNothing(isExist) = False Then
                    StockItemObj.ucmbPriceType.DisplayLayout.Bands(0).Columns(counter).Hidden = True
                End If
                counter = counter + 1
            Loop

        ElseIf IsNothing(TemplatesItemObj) = FalseThen Then
            'For TemplatesItemObj
            TemplatesItemObj.ucmbPriceType.DataSource = dataSetNew.Tables(0)
            TemplatesItemObj.ucmbPriceType.DisplayMember = displayMemberValNew
            TemplatesItemObj.ucmbPriceType.ValueMember = valueMemberVal
            StockItemObj.ucmbPriceType.Refresh()

            'Colums desinger
            TemplatesItemObj.ucmbPriceType.DisplayLayout.Bands(0).ColHeadersVisible = False
            Do While counter < TemplatesItemObj.ucmbPriceType.DisplayLayout.Bands(0).Columns.Count
                columName = TemplatesItemObj.ucmbPriceType.DisplayLayout.Bands(0).Columns(counter).Key
                isExist = Array.Exists(allowList, Function(s) s = columName)
                If IsNothing(isExist) = False Then
                    TemplatesItemObj.ucmbPriceType.DisplayLayout.Bands(0).Columns(counter).Hidden = True
                End If
                counter = counter + 1
            Loop
        End If
    End Sub

End Class
