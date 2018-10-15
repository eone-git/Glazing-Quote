Imports Infragistics.Win.UltraWinGrid
Imports System.Linq

Public Class clsGlazingDocStockItemHelper
    Dim moduleName As String = "Stock Item"
    Dim TemplatesItemObj As frmGlazingDocStockItemTemplate
    Dim StockItemObj As frmGlazingDocStockItem
    Dim StockItemServicesObj As frmGlazingDocStockItemServices
    Dim clsSOModuleDefaultsObj As clsSOModuleDefaults
    Dim clsSOPricingAndUnitsObj As clsSOPricingAndUnits
    Dim clsSqlConnObj As New clsSqlConn
    Dim sqlQuary As String = ""
    Dim newDataSet As DataSet = Nothing
    Dim oSOModuleDefaults As New clsSOModuleDefaults
    Public dataSetForItem As DataSet

    'Dim oPriceUnits As New clsSOPricingAndUnits

    Public Sub New(ByRef frmGlazingDocStockItemObj As frmGlazingDocStockItem)
        Me.StockItemObj = frmGlazingDocStockItemObj
        'Me.clsSOModuleDefaultsObj = clsSOModuleDefaultsInherit
        'Me.clsSOPricingAndUnitsObj = clsSOPricingAndUnitsInherit

    End Sub

    Public Sub New(ByRef frmGlazingDocStockItemTemplateObj As frmGlazingDocStockItemTemplate)
        Me.TemplatesItemObj = frmGlazingDocStockItemTemplateObj
    End Sub

    Public Sub New(ByRef frmGlazingDocStockItemServicesObj As frmGlazingDocStockItemServices)
        Me.StockItemServicesObj = frmGlazingDocStockItemServicesObj
    End Sub

    Public Sub New()

    End Sub

    'Public Sub SetPriceOnThisRow(ByRef ugRow As UltraGridRow)
    '    clsSOPricingAndUnitsObj.oCustomer = New clsCustomer(frmGlazingQuote.cmbAccount.Value)
    '    clsSOPricingAndUnitsObj.iDefaultStockPriceListID = clsSOModuleDefaultsObj.DefaultTradePriceListID ' iFormDefaultTradePriceListID
    '    iFormDefaultTradePriceListID = clsSOModuleDefaultsObj.DefaultTradePriceListID
    '    If ugRow.Cells("IsPriceItem").Value = False Then
    '        clsSOPricingAndUnitsObj.Set_PriceList_OnActiveRow_NotRelatingToPriceCalc(ugRow)
    '    Else
    '        clsSOPricingAndUnitsObj.GetStockPriceOnActiveRow(ugRow)
    '    End If

    '    If ugRow.Cells("PriceCat").Value = "T" Then
    '        'ugRow.Cells("PriceList").EditorComponent = cmbDDPriceListsTrade
    '        StockItemObj.cmbDDPriceListsTrade.Visible = True
    '        StockItemObj.cmbDDPriceListsTrade.Value = ugRow.Cells("PriceList").Value
    '    Else
    '        ugRow.Cells("PriceList").EditorComponent = StockItemObj.cmbDDPriceListsSpecial
    '    End If
    '    StockItemObj.txtPrice.Value = frmGlazingQuote.UG2.ActiveRow.Cells("Price").Value

    'End Sub

    Function GetStkItemDetails(Optional ByRef tempStockLink As Integer = -1) As DataSet

        Try
            If tempStockLink = -1 Then
                sqlQuary = "SELECT     StkItem.Description_1, StkItem.cSimpleCode, spilStkTemplateItems.Stock_Sub_Link, StkItem.ufIIThickness, TaxRate.TaxRate, " & _
                    "spilStkTemplateItems.UnitPricePercen,  StkItem.uiIISRVPRICEID ,StkItem.uiIITemplPriceID,spilStkTemplateItems.iItemType,spilStkTemplateItems.iLineNo " & _
                    ",spilStkTemplateItems.UnitPrice,spilStkTemplateItems.IsPriceItem,spilStkTemplateItems.SRVPRICE_CATID,spilStkTemplateItems.iHeight, " & _
                    "spilStkTemplateItems.iWidth,spilStkTemplateItems.Units,spilStkTemplateItems.Motif,spilStkTemplateItems.Comment2, StkItem.TTI, spilStkTemplateItems.PriceTypeID, " & _
                    "StkItem_1.uiIITemplPriceID AS TemplatePriceTypeID, StkItem.Description_3,spilStkTemplateItems.fQuantity, StkItem.BOMItem " & _
                    "FROM  spilStkTemplateItems INNER JOIN " & _
                    "StkItem ON spilStkTemplateItems.Stock_Sub_Link = StkItem.StockLink INNER JOIN " & _
                    "TaxRate ON StkItem.TTI = TaxRate.Code INNER JOIN " & _
                    "StkItem AS StkItem_1 ON spilStkTemplateItems.StockLink = StkItem_1.StockLink "

            Else
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
            End If
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
            Dim allowList() As String = UIHandler(allowListInput, displayMemberVal)
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
                'StockItemObj.cmbDDPriceListsTrade.
                StockItemObj.cmbDDPriceListsSpecial.DropDownWidth = StockItemObj.cmbDDPriceListsSpecial.Width

            ElseIf IsNothing(TemplatesItemObj) = False Then
                'For TemplatesItemObj
                TemplatesItemObj.cmbDDPriceListsTrade.DataSource = dataSetNew.Tables(0)
                TemplatesItemObj.cmbDDPriceListsTrade.DisplayLayout.Bands(0).ColHeadersVisible = False
                TemplatesItemObj.cmbDDPriceListsTrade.DisplayMember = displayMemberVal
                TemplatesItemObj.cmbDDPriceListsTrade.ValueMember = valueMemberVal

                TemplatesItemObj.cmbDDPriceListsSpecial.DataSource = dataSetNew.Tables(1)
                TemplatesItemObj.cmbDDPriceListsSpecial.DisplayLayout.Bands(0).ColHeadersVisible = False
                TemplatesItemObj.cmbDDPriceListsSpecial.DisplayMember = displayMemberVal
                TemplatesItemObj.cmbDDPriceListsSpecial.ValueMember = valueMemberVal

                'Colums desinger
                TemplatesItemObj.cmbDDPriceListsSpecial.DisplayLayout.Bands(0).ColHeadersVisible = False
                TemplatesItemObj.cmbDDPriceListsTrade.DisplayLayout.Bands(0).ColHeadersVisible = False

                Do While counter < TemplatesItemObj.cmbDDPriceListsTrade.DisplayLayout.Bands(0).Columns.Count And IsNothing(allowList) = False
                    columName = TemplatesItemObj.cmbDDPriceListsTrade.DisplayLayout.Bands(0).Columns(counter).Key
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

            dataSetForItem = GetPriceTypeData(sqlQuaryForController)
            Dim allowList() As String = UIHandler(allowListInput, displayMemberVal)
            Dim counter As Integer = 0
            Dim columName As String
            If IsNothing(StockItemObj) = False Then
                'For StockItem
                StockItemObj.ucmbItemCode.DataSource = dataSetForItem.Tables(0)
                StockItemObj.ucmbItemCode.DisplayMember = "Code"
                StockItemObj.ucmbItemCode.ValueMember = valueMemberVal
                StockItemObj.ucmbItemCode.Refresh()

                StockItemObj.ucmbItemDes.DataSource = dataSetForItem.Tables(0)
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
                TemplatesItemObj.ucmbItemCode.DataSource = dataSetForItem.Tables(0)
                TemplatesItemObj.ucmbItemCode.DisplayMember = "Code"
                TemplatesItemObj.ucmbItemCode.ValueMember = valueMemberVal
                TemplatesItemObj.ucmbItemCode.Refresh()

                TemplatesItemObj.ucmbItemDes.DataSource = dataSetForItem.Tables(0)
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

    Public Sub SetPriceOnThisRow(ByRef ugRow As UltraGridRow, ByRef oPriceUnits As clsSOPricingAndUnits, Optional ByRef clsModuleDefaultsObj As clsSOModuleDefaults = Nothing)

        Try
            If IsNothing(clsModuleDefaultsObj) = False Then
                oSOModuleDefaults = clsModuleDefaultsObj
            End If

            oPriceUnits.iDefaultStockPriceListID = oSOModuleDefaults.DefaultTradePriceListID ' iFormDefaultTradePriceListID
            iFormDefaultTradePriceListID = oSOModuleDefaults.DefaultTradePriceListID
            'If ugRow.Cells("IsPriceItem").Value = False Then
            '    oPriceUnits.Set_PriceList_OnActiveRow_NotRelatingToPriceCalc(ugRow)
            'Else
            oPriceUnits.GetStockPriceOnActiveRow(ugRow)
            'End If

            SetPriceList(ugRow)


        Catch ex As Exception
            MessageBox.Show(ex.Message, moduleName, MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try
    End Sub

    Sub SetPriceList(ByRef ugRow As UltraGridRow)
        Try
            If ugRow.Cells("PriceCat").Value = "T" Then
                If IsNothing(StockItemObj) = False Then
                    StockItemObj.cmbDDPriceListsTrade.Visible = True
                    StockItemObj.cmbDDPriceListsTrade.Value = ugRow.Cells("PriceList").Value
                ElseIf IsNothing(TemplatesItemObj) = False Then
                    ugRow.Cells("PriceList").EditorComponent = TemplatesItemObj.cmbDDPriceListsTrade
                    'ugRow.Cells("PriceList").Value = ugRow.Cells("PriceList").Value
                    'ugRow.Cells("Price").Value = ugRow.Cells("Price").Value

                End If
            Else
                If IsNothing(StockItemObj) = False Then
                    StockItemObj.cmbDDPriceListsSpecial.Visible = True
                    StockItemObj.cmbDDPriceListsSpecial.Value = ugRow.Cells("PriceList").Value
                ElseIf IsNothing(TemplatesItemObj) = False Then
                    ugRow.Cells("PriceList").EditorComponent = TemplatesItemObj.cmbDDPriceListsSpecial
                End If

            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message, moduleName, MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try
    End Sub

    Public Sub FillActiveRowFromSelectedProductParameters(ByRef uddSource As UltraCombo, ugRow As UltraGridRow, ByRef oPriceUnits As clsSOPricingAndUnits)
        Dim itemTypeValue As String = ""
        If IsNothing(TemplatesItemObj) = False Then
            itemTypeValue = TemplatesItemObj.UG2.ActiveRow.Cells("ItemType").Value
        ElseIf IsNothing(StockItemObj) = False Then
            itemTypeValue = StockItemObj.cmbDDItemType.Value
        End If
        Try
            If uddSource.ActiveRow IsNot Nothing Then

                If uddSource.Name = "ucmbItemDes" Then
                    ugRow.Cells("SimpleCode").Value = uddSource.ActiveRow.Cells("Code").Value
                Else
                    ugRow.Cells("Description1").Value = uddSource.ActiveRow.Cells("Description_1").Value
                End If

                ugRow.Cells("StockLink").Value = uddSource.ActiveRow.Cells("StockLink").Value
                ' ugRow.Cells("Description2").Value = uddSource.ActiveRow.Cells("Description_3").Value & " " & IIf(IsDBNull(uddSource.ActiveRow.Cells("AddDetails").Value), "", uddSource.ActiveRow.Cells("AddDetails").Value) 'uddSource.ActiveRow.Cells("Description_1").Value

                If ugRow.Band.Index = 0 Then 'Normal Glass Line (Band 0)
                    Select Case itemTypeValue
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
                        For Each uR As UltraGridRow In TemplatesItemObj.ucmbPriceType.Rows
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
                    End If

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
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

    End Sub

    Public Function CalculateTotalExc(ByRef TotalExc As Double, ByRef qty As Double, ByRef volume As Double, ByRef actualEnterdPrice As Double, ByRef quantity As Double, Optional ByRef subItemsPrice As Double = 0, Optional ByRef actualAmount As Double = 0, Optional ByRef excludeVolume As Boolean = False) As Double
        Try
            Dim enterdPrice As Double
            Dim totalPrice As Double
            Dim totalSubItemsAmount As Double
            Dim totalGlassValue As Double

            If subItemsPrice <> 0 Then
                totalSubItemsAmount = subItemsPrice
            Else
                totalSubItemsAmount = GetSubItemPrice()
            End If

            If actualAmount = 0 Then
                enterdPrice = actualEnterdPrice
            Else
                enterdPrice = actualAmount - subItemsPrice
            End If

            Dim exTempltePrice As Double = 0

            If excludeVolume = False Then
                totalGlassValue = (enterdPrice * volume) * qty
            Else
                totalGlassValue = (enterdPrice * 1) * qty
            End If

            'Calculate total
            totalPrice = If(IsNumeric(totalSubItemsAmount) = True, totalSubItemsAmount, 0) + If(IsNumeric(totalGlassValue) = True, totalGlassValue, 0)
            'actualTotalExc = totalPrice

            If qty = 0 Or volume = 0 And excludeVolume = False Then
                TotalExc = 0
            Else
                TotalExc = totalPrice
            End If

            Return TotalExc
        Catch ex As Exception
            modGlazingQuoteExtension.GQShowMessage(ex.Message, moduleName & "(Calculate Total Exc)", MsgBoxStyle.Exclamation)
            Return 0
        End Try
    End Function
    Public Function GetSubItemPrice(Optional ByRef selectedItems As String = "") As Double
        Try
            Dim subItemsPrice As Double = 0

            If StockItemObj.templateItemSubItemsPrice > 0 Then
                'Get values from templte sub items grid
                subItemsPrice = StockItemObj.templateItemSubItemsPrice
            Else
                'Get values from quation main grid
                Dim itemLines() As String = selectedItems.Split(";")
                For Each itemDetails As String In itemLines
                    If itemDetails = "" Then
                        Exit For
                    End If
                    Dim itemDetailsCol() As String = itemDetails.Split(",")
                    subItemsPrice = subItemsPrice + itemDetailsCol(3)
                Next
            End If

            Return subItemsPrice
        Catch ex As Exception
            modGlazingQuoteExtension.GQShowMessage(ex.Message, moduleName & "(Item Price)", MsgBoxStyle.Exclamation)
            Return 0
        End Try
    End Function

    Public Function CalculateTaxAmount(ByRef taxRate As Double, ByRef totalExc As Double, Optional ByRef isTaxExclude As Boolean = True) As Double
        Dim taxAmount As Double = 0
        Try
            Dim totalInc As Double
            If isTaxExclude = False Then
                taxRate = If(IsNothing(taxRate) = False, taxRate, 0)
                totalExc = If(IsNothing(totalExc) = False, totalExc, 0)
                taxAmount = (totalExc / 100) * taxRate
            Else
                taxAmount = 0
            End If
            Return taxAmount

        Catch ex As Exception
            Return taxAmount

        End Try
    End Function


End Class
