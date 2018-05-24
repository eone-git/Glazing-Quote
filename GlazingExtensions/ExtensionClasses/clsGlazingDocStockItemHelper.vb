Imports Infragistics.Win.UltraWinGrid
Imports System.Linq

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

    Public Function GetPriceTypeData(Optional ByRef sqlQuaryForController As String = "") As DataSet
        Try
            If sqlQuaryForController = "" Then

            End If

            newDataSet = clsSqlConnObj.GET_DATA_SQL(sqlQuaryForController)
            Return newDataSet

        Catch ex As Exception
            modGlazingQuoteExtension.GQShowMessage(ex.Message, moduleName, MsgBoxStyle.Critical)
            Return newDataSet

        End Try
    End Function

    Private Function UIHandler(ByRef allowListInput As String, Optional ByRef displayMemberVal As String = Nothing) As Array
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
        Return allowList

    End Function

    Public Function SetPriceTypeData(ByRef displayMemberVal As String, ByRef valueMemberVal As String, Optional ByRef allowListInput As String = "")
        Try
            Dim sqlQuaryForController As String = "SELECT TYPE_ID, TYPE_PRICE,TOUGHENED,IsAllocated, IsApplyOnlyTemplate," & _
                          "IsCuttingNeed,IGU_PROCESS, ItemType FROM stkPRICE_TYPESspil WHERE (SEGMENT_ID = 1) ORDER BY SQUENCE"
            Dim dataSetNew As DataSet = GetPriceTypeData(sqlQuaryForController)
            Dim allowList() As String = UIHandler(allowListInput, displayMemberVal)
            Dim counter As Integer = 0
            If IsNothing(StockItemObj) = False Then
                'For StockItem
                StockItemObj.ucmbPriceType.DataSource = dataSetNew.Tables(0)
                StockItemObj.ucmbPriceType.DisplayMember = displayMemberVal
                StockItemObj.ucmbPriceType.ValueMember = valueMemberVal
                StockItemObj.ucmbPriceType.Refresh()

                'Colums desinger
                StockItemObj.ucmbPriceType.DisplayLayout.Bands(0).ColHeadersVisible = False
                Do While counter < StockItemObj.ucmbPriceType.DisplayLayout.Bands(0).Columns.Count And IsNothing(allowList) = False

                    columName = StockItemObj.ucmbPriceType.DisplayLayout.Bands(0).Columns(counter).Key
                    isExist = Array.Exists(allowList, Function(s) s = columName)
                    If isExist = False Then
                        StockItemObj.ucmbPriceType.DisplayLayout.Bands(0).Columns(counter).Hidden = True
                    End If
                    counter = counter + 1
                Loop

            ElseIf IsNothing(TemplatesItemObj) = False Then
                'For TemplatesItemObj
                TemplatesItemObj.ucmbPriceType.DataSource = dataSetNew.Tables(0)
                TemplatesItemObj.ucmbPriceType.DisplayMember = displayMemberVal
                TemplatesItemObj.ucmbPriceType.ValueMember = valueMemberVal
                TemplatesItemObj.ucmbPriceType.Refresh()

                'Colums desinger
                TemplatesItemObj.ucmbPriceType.DisplayLayout.Bands(0).ColHeadersVisible = False
                Do While counter < TemplatesItemObj.ucmbPriceType.DisplayLayout.Bands(0).Columns.Count And IsNothing(allowList) = False
                    columName = TemplatesItemObj.ucmbPriceType.DisplayLayout.Bands(0).Columns(counter).Key
                    isExist = Array.Exists(allowList, Function(s) s = columName)
                    If isExist = False Then
                        TemplatesItemObj.ucmbPriceType.DisplayLayout.Bands(0).Columns(counter).Hidden = True
                    End If
                    counter = counter + 1
                Loop

            End If
            Return dataSetNew
        Catch ex As Exception
            MessageBox.Show(ex.Message, moduleName, MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return Nothing
        End Try
    End Function

    Public Function SetPriceListsData(ByRef displayMemberVal As String, ByRef valueMemberVal As String, Optional ByRef allowListInput As String = "") As DataSet
        Try
            Dim sqlQuaryForController As String = "SELECT CAT_ID, CAT_DESCRIPTION As PriceList FROM stkPRICE_CATEGORYspil"
            sqlQuaryForController += " SELECT   CAT_ID, CAT_DESCRIPTION As PriceList FROM stkPRICE_SPECIALCATEGORYspil"
            Dim dataSetNew As DataSet = GetPriceTypeData(sqlQuaryForController)
            Dim displayMemberArray() = allowListInput.Split(New Char() {";"c})
            Dim allowList() As String = UIHandler(allowListInput, "CAT_ID")
            Dim counter As Integer = 0

            If IsNothing(StockItemObj) = False Then
                'For StockItem
                StockItemObj.cmbDDPriceListsTrade.DataSource = dataSetNew.Tables(0)
                StockItemObj.cmbDDPriceListsTrade.DisplayLayout.Bands(0).ColHeadersVisible = False
                StockItemObj.cmbDDPriceListsTrade.DisplayMember = displayMemberVal
                StockItemObj.cmbDDPriceListsTrade.ValueMember = valueMemberVal

                StockItemObj.cmbDDPriceListsSpecial.DataSource = dataSetNew.Tables(1)
                StockItemObj.cmbDDPriceListsSpecial.DisplayLayout.Bands(0).ColHeadersVisible = False
                StockItemObj.cmbDDPriceListsSpecial.DisplayMember = displayMemberVal
                StockItemObj.cmbDDPriceListsSpecial.ValueMember = valueMemberVal

                'Colums desinger
                StockItemObj.cmbDDItemType.DisplayLayout.Bands(0).ColHeadersVisible = False
                Do While counter < StockItemObj.cmbDDItemType.DisplayLayout.Bands(0).Columns.Count And IsNothing(allowList) = False

                    columName = StockItemObj.cmbDDPriceListsSpecial.DisplayLayout.Bands(0).Columns(counter).Key
                    isExist = Array.Exists(allowList, Function(s) s = columName)
                    If isExist = False Then
                        StockItemObj.cmbDDPriceListsSpecial.DisplayLayout.Bands(0).Columns(counter).Hidden = True
                        StockItemObj.cmbDDPriceListsTrade.DisplayLayout.Bands(0).Columns(counter).Hidden = True
          
                    End If
                    counter = counter + 1
                Loop
                StockItemObj.cmbDDPriceListsTrade.DropDownWidth = StockItemObj.cmbDDPriceListsTrade.Width
                StockItemObj.cmbDDPriceListsSpecial.DropDownWidth = StockItemObj.cmbDDPriceListsSpecial.Width

            ElseIf IsNothing(TemplatesItemObj) = False Then
                'For TemplatesItemObj
                TemplatesItemObj.cmbDDPriceListsTrade.DataSource = DS.Tables(0)
                TemplatesItemObj.cmbDDPriceListsTrade.DisplayLayout.Bands(0).ColHeadersVisible = False
                TemplatesItemObj.cmbDDPriceListsTrade.DisplayMember = "PriceList"
                TemplatesItemObj.cmbDDPriceListsTrade.ValueMember = "CAT_ID"

                TemplatesItemObj.cmbDDPriceListsSpecial.DataSource = DS.Tables(1)
                TemplatesItemObj.cmbDDPriceListsSpecial.DisplayLayout.Bands(0).ColHeadersVisible = False
                TemplatesItemObj.cmbDDPriceListsSpecial.DisplayMember = "PriceList"
                TemplatesItemObj.cmbDDPriceListsSpecial.ValueMember = "CAT_ID"

                'Colums desinger
                TemplatesItemObj.cmbDDItemType.DisplayLayout.Bands(0).ColHeadersVisible = False
                Do While counter < TemplatesItemObj.cmbDDItemType.DisplayLayout.Bands(0).Columns.Count And IsNothing(allowList) = False
                    columName = TemplatesItemObj.cmbDDItemType.DisplayLayout.Bands(0).Columns(counter).Key
                    isExist = Array.Exists(allowList, Function(s) s = columName)
                    If isExist = False Then
                        TemplatesItemObj.cmbDDPriceListsSpecial.DisplayLayout.Bands(0).Columns(counter).Hidden = True
                        TemplatesItemObj.cmbDDPriceListsTrade.DisplayLayout.Bands(0).Columns(counter).Hidden = True

                    End If
                    counter = counter + 1
                Loop
                'TemplatesItemObj.cmbDDPriceListsTrade.DropDownWidth = StockItemObj.cmbDDPriceListsTrade.Width
                'TemplatesItemObj.cmbDDPriceListsSpecial.DropDownWidth = StockItemObj.cmbDDPriceListsSpecial.Width

            End If
         
            Return dataSetNew

        Catch ex As Exception
            MessageBox.Show(ex.Message, moduleName, MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return Nothing
        End Try
    End Function

    Public Function SetItemCodeData(ByRef displayMemberVal As String, ByRef valueMemberVal As String, Optional ByRef allowListInput As String = "")
        Try
            Dim sqlQuaryForController As String = ""
            sqlQuaryForController = sqlQuaryForController & "SELECT   StkItem.ItemImage, StkItem.Description_1, StkItem.cSimpleCode as Code, StkItem.StockLink, " & vbCrLf
            sqlQuaryForController = sqlQuaryForController & "StkItem.ufIIThickness, TaxRate.TaxRate,stkItem.uiIISRVPRICEID, " & vbCrLf
            sqlQuaryForController = sqlQuaryForController & "StkItem.TTI, StkItem.uiIIPRICETYPEID, StkItem.uiIIItemType,StkItem.Description_3, " & vbCrLf
            sqlQuaryForController = sqlQuaryForController & "StkItem.AddDetails,StkItem.TemplType,StkItem.TemplItemCat,StkItem.TemplSelectThickness, " & vbCrLf
            sqlQuaryForController = sqlQuaryForController & "CategoryID " & vbCrLf
            sqlQuaryForController = sqlQuaryForController & ",SubTypeID " & vbCrLf
            sqlQuaryForController = sqlQuaryForController & "FROM StkItem " & vbCrLf
            sqlQuaryForController = sqlQuaryForController & "LEFT OUTER JOIN (SELECT   StockLink, MAX(CategoryID) AS CategoryID " & vbCrLf
            sqlQuaryForController = sqlQuaryForController & "        FROM     spilStkTemplateItems " & vbCrLf
            sqlQuaryForController = sqlQuaryForController & "        GROUP BY spilStkTemplateItems.StockLink) t2 ON StkItem.StockLink = t2.StockLink " & vbCrLf
            sqlQuaryForController = sqlQuaryForController & "LEFT OUTER JOIN " & vbCrLf
            sqlQuaryForController = sqlQuaryForController & "TaxRate ON StkItem.TTI = TaxRate.Code " & vbCrLf
            sqlQuaryForController = sqlQuaryForController & "WHERE  StkItem.uiIIItemType IN (1,2,3,4,16) and StkItem.ubIIGLASSSERVICE=0 and ItemActive = 1 order by StkItem.Description_2, StkItem.ufIIThickness"

            Dim dataSetNew As DataSet = GetPriceTypeData(sqlQuaryForController)
            Dim allowList() As String = UIHandler(allowListInput, displayMemberVal)
            Dim counter As Integer = 0
            Dim columName As String
            If IsNothing(StockItemObj) = False Then
                'For StockItem
                StockItemObj.ucmbItemCode.DataSource = dataSetNew.Tables(0)
                StockItemObj.ucmbItemCode.DisplayMember = "Code"
                StockItemObj.ucmbItemCode.ValueMember = valueMemberVal
                StockItemObj.ucmbItemCode.Refresh()

                StockItemObj.ucmbItemDes.DataSource = dataSetNew.Tables(0)
                StockItemObj.ucmbItemDes.DisplayMember = "Description_1"
                StockItemObj.ucmbItemDes.ValueMember = valueMemberVal
                StockItemObj.ucmbItemDes.Refresh()

                'Colums desinger
                Do While counter < StockItemObj.ucmbItemCode.DisplayLayout.Bands(0).Columns.Count And IsNothing(allowList) = False

                    columName = StockItemObj.ucmbItemCode.DisplayLayout.Bands(0).Columns(counter).Key()
                    isExist = Array.Exists(allowList, Function(s) s = columName)
                    If isExist = False Then
                        StockItemObj.ucmbItemCode.DisplayLayout.Bands(0).Columns(counter).Hidden = True
                        StockItemObj.ucmbItemDes.DisplayLayout.Bands(0).Columns(counter).Hidden = True
                    Else
                        StockItemObj.ucmbItemCode.DisplayLayout.Bands(0).Columns("Code").Header.VisiblePosition = 1
                        StockItemObj.ucmbItemCode.DisplayLayout.Bands(0).Columns("Description_1").Header.VisiblePosition = 2
                        StockItemObj.ucmbItemDes.DisplayLayout.Bands(0).Columns("Code").Header.VisiblePosition = 2
                        StockItemObj.ucmbItemDes.DisplayLayout.Bands(0).Columns("Description_1").Header.VisiblePosition = 1
                    End If
                    counter = counter + 1
                Loop

            ElseIf IsNothing(TemplatesItemObj) = False Then
                'For TemplatesItemObj
                TemplatesItemObj.ucmbItemCode.DataSource = dataSetNew.Tables(0)
                TemplatesItemObj.ucmbItemCode.DisplayMember = "Code"
                TemplatesItemObj.ucmbItemCode.ValueMember = valueMemberVal
                TemplatesItemObj.ucmbItemCode.Refresh()

                TemplatesItemObj.ucmbItemDes.DataSource = dataSetNew.Tables(0)
                TemplatesItemObj.ucmbItemDes.DisplayMember = "Description_1"
                TemplatesItemObj.ucmbItemDes.ValueMember = valueMemberVal
                TemplatesItemObj.ucmbItemDes.Refresh()

                'Colums desinger
                Do While counter < TemplatesItemObj.ucmbItemCode.DisplayLayout.Bands(0).Columns.Count And IsNothing(allowList) = False
                    columName = TemplatesItemObj.ucmbItemCode.DisplayLayout.Bands(0).Columns(counter).Key
                    isExist = Array.Exists(allowList, Function(s) s = columName)
                    If isExist = False Then
                        TemplatesItemObj.ucmbItemCode.DisplayLayout.Bands(0).Columns(counter).Hidden = True
                        TemplatesItemObj.ucmbItemDes.DisplayLayout.Bands(0).Columns(counter).Hidden = True
                    Else
                        TemplatesItemObj.ucmbItemCode.DisplayLayout.Bands(0).Columns("Code").Header.VisiblePosition = 1
                        TemplatesItemObj.ucmbItemCode.DisplayLayout.Bands(0).Columns("Description_1").Header.VisiblePosition = 2
                        TemplatesItemObj.ucmbItemDes.DisplayLayout.Bands(0).Columns("Code").Header.VisiblePosition = 2
                        TemplatesItemObj.ucmbItemDes.DisplayLayout.Bands(0).Columns("Description_1").Header.VisiblePosition = 1
                    End If
                    counter = counter + 1
                Loop
            End If
            Return dataSetNew

        Catch ex As Exception
            MessageBox.Show(ex.Message, moduleName, MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return Nothing
        End Try
    End Function
End Class
