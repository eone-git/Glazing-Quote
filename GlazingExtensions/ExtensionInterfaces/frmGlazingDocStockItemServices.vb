Imports Infragistics.Win.UltraWinGrid

Public Class frmGlazingDocStockItemServices

#Region "Variables"
    Public mainStockItem As Double
    Public mainStockItemHeight As Double
    Public mainStockItemWidth As Double
    Public dblThickness As Double
    Public zipCode As Integer

    Dim frmGlazingDocStockItemObj As New frmGlazingDocStockItem
    Dim clsGlazingDocStockItemHelperObj As New clsGlazingDocStockItemHelper(Me)
    Dim isSaving As Boolean = False
    Dim isLoading As Boolean = False
    Dim isSaved As Boolean = False
    Dim isLoaded As Boolean = False
    Dim holdValueChage As Boolean = False
    Dim MyDataSet As DataSet
    Dim componentsDatset As DataSet
    Dim oSOModuleDefaults As New clsSOModuleDefaults
    Public oPriceUnitsMe As New clsSOPricingAndUnits
    Public dCusTaxCode As String
    Public dCusTaxRate As Double
    Dim gridActiveRow As UltraGridRow
    Dim gridActiveCell As UltraGridCell
    Public ug2Row As UltraGridRow
    Public isTempered As Boolean
    Public cusID As Integer
    Public selectedItems As String = ""
    Private selectedNewItems As String = ""
    Public selectedItemsDisplay As String = ""
    Private selectedNewItemsDisplay As String = ""
    Private taxRate As Double = 0
    Public totalAmount As Double = 0.0
    Private totalNewAmount As Double = 0.0
    Public isDiscard As Boolean = False
    Public isDiscard2 As Boolean = False
    Public mainGlassQty As Double = 0.0
    Dim activeTaxExempt As Boolean = False
    Dim taxExemptVal As String

    Public glassHeight As Double = 0
    Public glassWidth As Double = 0
    Public glassVolume As Double = 0

#End Region
#Region "Constructor"
    Public Sub New()
        ' This call is required by the designer.
        InitializeComponent()
    End Sub

#End Region
#Region "Form core funtions"

    Private Sub frmGlazingDocStockItemServices_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Try
            LoadData()
        Catch ex As Exception
            modGlazingQuoteExtension.GQShowMessage(ex.Message, Me.Text, MsgBoxStyle.Critical, "warning", "frmGlazingDocStockItemServices_Load")

        End Try
    End Sub

    Private Sub frmGlazingDocStockItemServices_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        Try
            Dim Result As DialogResult = MessageBox.Show("Do you want process the selection", Me.Text, MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question)
            If Result = Windows.Forms.DialogResult.Yes Then
                SetDataAsString()
                selectedItems = selectedNewItems
                selectedItemsDisplay = selectedNewItemsDisplay
                totalAmount = totalNewAmount
            ElseIf Result = Windows.Forms.DialogResult.No Then
                Dim Result2 As DialogResult = MessageBox.Show("Are you sure you want to discard the changes ?", Me.Text, MessageBoxButtons.YesNo, MessageBoxIcon.Question)
                If Result2 = Windows.Forms.DialogResult.Yes Then
                    isDiscard = True
                    Exit Sub
                Else
                    e.Cancel = True
                End If

            ElseIf Result = Windows.Forms.DialogResult.Cancel Then
                e.Cancel = True

            End If
        Catch ex As Exception
            modGlazingQuoteExtension.GQShowMessage(ex.Message, Me.Text, MsgBoxStyle.Critical, "warning", "frmGlazingDocStockItemServices_FormClosing")

        End Try
    End Sub

#End Region
#Region "CRUD"

    Sub SaveData()
        Try

        Catch ex As Exception
            modGlazingQuoteExtension.GQShowMessage(ex.Message, Me.Text, MsgBoxStyle.Critical, "warning", "SaveData")

        End Try
    End Sub

    Sub LoadData()
        Try
            LoadComponenetsData()
            ComponentsHandler()
            SetItemData()

            If activeTaxExempt = False Then
                taxExemptVal = " "
            Else
                taxExemptVal = ", TaxExempt "
            End If


        Catch ex As Exception
            modGlazingQuoteExtension.GQShowMessage(ex.Message, Me.Text, MsgBoxStyle.Critical, "warning", "LoadData")

        End Try
    End Sub

    Sub SetItemData()
        Dim newDataSet As DataSet
        Dim ugActivRow As UltraGridRow = UG.ActiveRow
        Dim row As UltraGridRow
        Dim hasPreVal As Boolean = False
        Try
            If IsNothing(selectedItems) = False Then
                If selectedItems <> "" Then
                    hasPreVal = True
                End If
            End If

            'Dim foundRows() As DataRow
            If hasPreVal = True Then
                ' newDataSet = clsGlazingDocStockItemHelperObj.GetStkItemDetails()

                Dim itemLines() As String = selectedItems.Split(";")
                For Each itemDetails As String In itemLines
                    clsInvDetLine = New clsInvDetailLine
                    If itemDetails = "" Then
                        Exit For
                    End If
                    row = UG.DisplayLayout.Bands(0).AddNew
                    UG.ActiveRow = row
                    Dim itemDetailsCol() As String = itemDetails.Split(",")
                    DDStock.Value = -99
                    DDStock.Value = itemDetailsCol(0)
                    'foundRows = newDataSet.Tables(0).Select("cSimpleCode='" & ucmbItemCode.SelectedRow.Cells("Code").Value & "'")
                    row.Cells("StockLink").Value = itemDetailsCol(0)
                    row.Cells("Code").Value = itemDetailsCol(0)
                    UG.ActiveCell = row.Cells("Code")
                    UGGridColumsHandler()
                    'row.Cells("ItemType").Value = itemDetailsCol()
                    'row.Cells("IsPriceItem").Value = foundRows(1).Item("IsPriceItem")
                    ' row.Cells("Description1").Value = itemDetailsCol(0)
                    ' row.Cells("PriceList").Value = itemDetailsCol(1)
                    ' row.Cells("PriceType").Value = itemDetailsCol(2)
                    'row.Cells("Price").Value = itemDetailsCol(5)

                Next

            Else
                row = UG.DisplayLayout.Bands(0).AddNew
                UG.ActiveRow = row

            End If

        Catch ex As Exception
            modGlazingQuoteExtension.GQShowMessage(ex.Message, Me.Text, MsgBoxStyle.Critical, "warning", "SetItemData")

        End Try
    End Sub


#End Region
#Region "Advance Options"

    Function GeographicRegionBasedSettings() As Boolean
        Try
            Return True
        Catch ex As Exception
            modGlazingQuoteExtension.GQShowMessage(ex.Message, Me.Text, MsgBoxStyle.Critical, "warning", "GeographicRegionBasedSettings")
            Return False

        End Try
    End Function

    Sub SetDataAsString()
        Try
            For Each rows As UltraGridRow In UG.Rows
                selectedNewItems = selectedNewItems & rows.Cells("StockLink").Value & "," & rows.Cells("Code").Value & "," & rows.Cells("Comment").Value & "," & rows.Cells("Units").Value & "," & rows.Cells("TaxExempt").Value & "," & rows.Cells("Price_Excl").Value & "," & rows.Cells("NetAmount").Value & "," & rows.Cells("Pricing_ID").Value & "," & rows.Cells("Price_Excl").Value & "," & rows.Cells("TaxRate").Value & "," & rows.Cells("Units").Value & ";"

                selectedNewItemsDisplay = selectedNewItemsDisplay & Chr(9) & "+ " & rows.Cells("Description_1").Text & vbCrLf

                totalNewAmount = Math.Round(totalNewAmount + rows.Cells("LineNet_Excl").Value, 8, MidpointRounding.AwayFromZero)

            Next
        Catch ex As Exception
            modGlazingQuoteExtension.GQShowMessage(ex.Message, Me.Text, MsgBoxStyle.Critical, "warning", "SetDataAsString")

        End Try
    End Sub

#End Region
#Region "Components Handler"

    Sub ComponentsHandler()
        Try
            ComponentsTextStyling()
            ComponentsVisibility()
            ComponentsFuntinality()
        Catch ex As Exception
            modGlazingQuoteExtension.GQShowMessage(ex.Message, Me.Text, MsgBoxStyle.Critical, "warning", "ComponentsHandler")

        End Try
    End Sub

    Sub ComponentsTextStyling()
        Try

        Catch ex As Exception
            modGlazingQuoteExtension.GQShowMessage(ex.Message, Me.Text, MsgBoxStyle.Critical, "warning", "ComponentsTextStyling")

        End Try
    End Sub

    Sub ComponentsVisibility()
        Try

        Catch ex As Exception
            modGlazingQuoteExtension.GQShowMessage(ex.Message, Me.Text, MsgBoxStyle.Critical, "warning", "ComponentsVisibility")

        End Try
    End Sub

    Sub ComponentsFuntinality()
        Try
            'AddNewRow("after")
        Catch ex As Exception
            modGlazingQuoteExtension.GQShowMessage(ex.Message, Me.Text, MsgBoxStyle.Critical, "warning", "ComponentsFuntinality")

        End Try
    End Sub

    Public Function AddNewRow(ByVal rowPostion As String) As UltraGridRow
        Try
            Dim activeRowIndex As Integer = 0
            If IsNothing(Me.UG.ActiveRow) = False Then
                If rowPostion = "before" Then
                    activeRowIndex = Me.UG.ActiveRow.Index

                ElseIf rowPostion = "after" Then
                    activeRowIndex = Me.UG.ActiveRow.Index + 1

                ElseIf rowPostion = "end" Then
                    activeRowIndex = Me.UG.Rows.Count

                End If
            End If
            Dim row As UltraGridRow = UG.DisplayLayout.Bands(0).AddNew
            row.ParentCollection.Move(row, activeRowIndex)
            Me.UG.ActiveRowScrollRegion.ScrollRowIntoView(row)
            Me.UG.Rows(activeRowIndex).Activate()

            UG.PerformAction(UltraGridAction.EnterEditMode, True, True)
            ' row = UG.GetRow(ChildRow.First)
            UG.PerformAction(UltraGridAction.EnterEditModeAndDropdown, True, True)
            row.Cells("Code").Activate()

            ''Set new values
            'Me.UG.Rows(activeRowIndex).Cells("TaxRate").Value = defaultTaxtRateValue
            'Me.UG.Rows(activeRowIndex).Cells("TaxRateValue").Value = defaultTaxtRateValue
            'Me.UG.Rows(activeRowIndex).Cells("Dis.Width").Value = "0"
            'Me.UG.Rows(activeRowIndex).Cells("Dis.Height").Value = "0"

            Return row

        Catch ex As Exception
            modGlazingQuoteExtension.GQShowMessage(ex.Message, Me.Text, MsgBoxStyle.Critical, "warning", "AddNewRow")

        End Try
    End Function
#End Region
#Region "Load Componenets Data"

    Private Function GET_DATA() As DataSet
        Dim objSQL As New clsSqlConn
        Dim DS = New DataSet
        Dim dstringSQLOne As String = ""
        Dim dstringSQLTwo As String = ""
      
        MyDataSet = New DataSet
        Try
            dstringSQLOne = " SELECT Code AS cSimpleCode, Description_1, StockLink" & taxExemptVal & " FROM StkItem WHERE ItemActive = 1 AND ubIIGLASSSERVICE = 1 AND DefaultGlassService = 0 ORDER BY Code"
            dstringSQLOne += " SELECT Description_1, Code AS cSimpleCode, StockLink" & taxExemptVal & " FROM StkItem WHERE ItemActive = 1 AND ubIIGLASSSERVICE = 1 AND DefaultGlassService = 0 ORDER BY Description_1"
            dstringSQLOne += " SELECT SRVPRICEID as ID, SRVPRICE_DESCRIPTION as Pricing FROM stkSRVPRICE_TYPESspil"
            dstringSQLOne += " SELECT SRVPRICE_CATID as Edging ,SRVPRICE_DESCRIPTION as Description, AutoLinkSrvID FROM stkSRVPRICE_TYPE_CATEGORYSspil"
            If geographicRegionID = GeographicRegion.NorthAmerica Then
                dstringSQLOne += " SELECT ZipCodeID, ZipCode, TaxRate FROM spil_USAZIPCode"
            Else
                dstringSQLOne += " SELECT idTaxRate, Code, Description, TaxRate FROM TaxRate"
            End If
            dstringSQLOne += " SELECT Code as cSimpleCode,Description_1,  StockLink,StkItem.SubTypeID  FROM StkItem INNER JOIN spil_InvSubType ON StkItem.SubTypeID=spil_InvSubType.SubTypeID where ItemActive = 1 and ubIIGLASSSERVICE = 0 and DefaultGlassService=0 AND spil_InvSubType.SubTypeID=" & oSOModuleDefaults.LinkedSrvCatIDFilter & " order by Code"

            dstringSQLTwo = "SELECT SRVPRICEID as ID, SRVPRICE_DESCRIPTION as Pricing FROM stkSRVPRICE_TYPESspil"
            dstringSQLTwo += " SELECT SRVPRICE_CATID as Edging ,SRVPRICE_DESCRIPTION as Description, AutoLinkSrvID FROM stkSRVPRICE_TYPE_CATEGORYSspil"
            dstringSQLTwo += " SELECT THICKNESS_ID, THICKNESS_UPTO FROM  stkTHICKNESS_UPTOspil  ORDER BY THICKNESS_ID"
            dstringSQLTwo += " SELECT *  FROM stkSRVPRICE_TYPE_CATEGORYSspil"

            DS = objSQL.GET_INSERT_UPDATE(dstringSQLOne)
            componentsDatset = DS
            MyDataSet = objSQL.GET_INSERT_UPDATE(dstringSQLTwo)

            Return DS
        Catch ex As Exception
            modGlazingQuoteExtension.GQShowMessage(ex.Message, Me.Text, MsgBoxStyle.Critical, "warning", "GET_DATA")
            Return DS

        End Try
    End Function

    Sub LoadComponenetsData()
        Try
            componentsDatset = GET_DATA()
            LoadDDStockData()
            LoadDDDescriptionData()
            LoadDDServLinkCatData()
            LoadDDTaxTypeData()
            LoadDDEdgingData()
            LoadDDPricingData()
            FindThicknessID(dblThickness)
            ComponentsApperance()
        Catch ex As Exception
            modGlazingQuoteExtension.GQShowMessage(ex.Message, Me.Text, MsgBoxStyle.Critical, "warning", "LoadComponenetsData")

        End Try
    End Sub

    Sub LoadDDStockData()
        Try
            DDStock.DataSource = componentsDatset.Tables(0)
            DDStock.ValueMember = "StockLink"
            DDStock.DisplayMember = "cSimpleCode"
            DDStock.DisplayLayout.Bands(0).Columns("StockLink").Hidden = True
            If activeTaxExempt = True Then
                DDStock.DisplayLayout.Bands(0).Columns("TaxExempt").Hidden = True
            End If
            DDStock.DisplayLayout.Bands(0).Columns("Description_1").Width = 400
            DDStock.DisplayLayout.Bands(0).Columns("cSimpleCode").Width = 200
            DDStock.DisplayLayout.Bands(0).Columns("cSimpleCode").Header.Caption = "Code"
            DDStock.Value = 128
            Dim Text = DDStock.Text
            DDStock.LimitToList = False


        Catch ex As Exception
            modGlazingQuoteExtension.GQShowMessage(ex.Message, Me.Text, MsgBoxStyle.Critical, "warning", "LoadDDStockData")

        End Try
    End Sub

    Sub LoadDDDescriptionData()
        Try
            DDDescription.DataSource = componentsDatset.Tables(1)
            DDDescription.ValueMember = "Description_1"
            DDDescription.DisplayMember = "Description_1"
            DDDescription.DisplayLayout.Bands(0).Columns("StockLink").Hidden = True
            If activeTaxExempt = True Then
                DDDescription.DisplayLayout.Bands(0).Columns("TaxExempt").Hidden = True
            End If

            DDDescription.DisplayLayout.Bands(0).Columns("Description_1").Width = 400
            DDDescription.DisplayLayout.Bands(0).Columns("cSimpleCode").Width = 100
            DDDescription.DisplayLayout.Bands(0).Columns("cSimpleCode").Header.Caption = "Code"
        Catch ex As Exception
            modGlazingQuoteExtension.GQShowMessage(ex.Message, Me.Text, MsgBoxStyle.Critical, "warning", "LoadDDDescriptionData")

        End Try
    End Sub

    Sub LoadDDServLinkCatData()
        Try
            DDServLinkCat.DataSource = componentsDatset.Tables(5)
            DDServLinkCat.ValueMember = "StockLink"
            DDServLinkCat.DisplayMember = "cSimpleCode"
            DDServLinkCat.DisplayLayout.Bands(0).Columns("StockLink").Hidden = True
            DDServLinkCat.DisplayLayout.Bands(0).Columns("SubTypeID").Hidden = True
            DDServLinkCat.DisplayLayout.Bands(0).Columns("Description_1").Width = 50
            DDServLinkCat.DisplayLayout.Bands(0).Columns("cSimpleCode").Width = 50
            DDServLinkCat.DisplayLayout.Bands(0).Columns("cSimpleCode").Header.Caption = "Code"
        Catch ex As Exception
            modGlazingQuoteExtension.GQShowMessage(ex.Message, Me.Text, MsgBoxStyle.Critical, "warning", "LoadDDServLinkCatData")

        End Try
    End Sub

    Sub LoadDDTaxTypeData()
        Try
            DDTaxType.DataSource = componentsDatset.Tables(4)
            If geographicRegionID = GeographicRegion.NorthAmerica Then
                DDTaxType.ValueMember = "TaxRate"
                DDTaxType.DisplayMember = "ZipCodeID"
                DDTaxType.DisplayLayout.Bands(0).Columns("ZipCodeID").Hidden = True
            Else
                DDTaxType.ValueMember = "idTaxRate"
                DDTaxType.DisplayMember = "Code"
                DDTaxType.DisplayLayout.Bands(0).Columns("idTaxRate").Hidden = True
            End If

        Catch ex As Exception
            modGlazingQuoteExtension.GQShowMessage(ex.Message, Me.Text, MsgBoxStyle.Critical, "warning", "LoadDDTaxTypeData")

        End Try
    End Sub

    Sub LoadDDEdgingData()
        Try
            DDEdging.DataSource = componentsDatset.Tables(3)
            DDEdging.ValueMember = "Edging"
            DDEdging.DisplayMember = "Description"
        Catch ex As Exception
            modGlazingQuoteExtension.GQShowMessage(ex.Message, Me.Text, MsgBoxStyle.Critical, "warning", "LoadDDEdgingData")
        End Try
    End Sub

    Sub LoadDDPricingData()
        Try
            DDPricing.DataSource = componentsDatset.Tables(2)
            DDPricing.ValueMember = "ID"
            DDPricing.DisplayMember = "Pricing"
            DDPricing.DisplayLayout.Bands(0).Columns("ID").Hidden = True
        Catch ex As Exception
            modGlazingQuoteExtension.GQShowMessage(ex.Message, Me.Text, MsgBoxStyle.Critical, "warning", "LoadDDPricingData")

        End Try
    End Sub

    Sub ComponentsApperance()
        Try
            Me.Text = "Glass services"
            Me.UG.Dock = DockStyle.Fill

        Catch ex As Exception
            modGlazingQuoteExtension.GQShowMessage(ex.Message, Me.Text, MsgBoxStyle.Critical, "warning", "ComponentsApperance")

        End Try
    End Sub


#End Region
#Region "Components events"
    Private Sub DDTaxType_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles DDTaxType.Click
        Try
            If DDTaxType.ActiveRow.Activated = True Then
                UG.ActiveRow.Cells("TaxType").Value = DDTaxType.ActiveRow.Cells("ZipCode").Value
                UG.ActiveRow.Cells("TaxRate").Value = DDTaxType.ActiveRow.Cells("TotalTaxRate").Value

            End If

        Catch ex As Exception
            Exit Sub
        End Try
    End Sub

    Private Sub UG_BeforeExitEditMode(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.BeforeExitEditModeEventArgs) Handles UG.BeforeExitEditMode
        UGGridColumsHandler(e)
    End Sub

    Private Sub btnOk_Click(sender As Object, e As EventArgs) Handles btnOk.Click
        Try
            Me.Close()
        Catch ex As Exception

        End Try
    End Sub

    Private Sub UG_CellChange(sender As Object, e As CellEventArgs) Handles UG.CellChange
        'UG2GridFunctionalities()
    End Sub

    Private Sub UG_AfterCellUpdate(sender As Object, e As CellEventArgs) Handles UG.AfterCellUpdate
        UG2GridFunctionalities(e)
    End Sub

#End Region
#Region "ComponentsFuntinality"
    Sub UG2GridFunctionalities(Optional ByRef e As CellEventArgs = Nothing)
        Try
            If IsNothing(e) = True Then
                gridActiveRow = Me.UG.ActiveRow
                gridActiveCell = Me.UG.ActiveCell
            Else
                gridActiveRow = e.Cell.Row
                gridActiveCell = e.Cell
            End If

            If isLoading = False Then
                CalaulateTaxAmount()
            End If

        Catch ex As Exception

        End Try
    End Sub

    Public Sub SetItemDetail(ByVal stklnk As Integer, ByVal intIndex As Integer)
        'Use this Function When Adding New Single Service
        Dim objSQL As New clsSqlConn
        Dim DS_ITEMS As DataSet
        Dim dr1 As DataRow
        Dim isTaxExempt As Boolean
        Try
            boAftrCelUpdate = False
            If stklnk <> 0 Then

mylbl:
                If geographicRegionID = GeographicRegion.NorthAmerica Then
                    SQL = "SELECT cSimpleCode, Description_1, Description_2, iInvSegValue1ID, StockLink, uiIISRVPRICEID, ubIIEDGEALLOWANCE" & taxExemptVal & " FROM StkItem WHERE StockLink = " & stklnk & ""
                Else
                    SQL = "SELECT     StkItem.cSimpleCode, StkItem.Description_1, StkItem.Description_2, StkItem.iInvSegValue1ID, " & _
                        " StkItem.StockLink, StkItem.uiIISRVPRICEID, " & _
                        " StkItem.ubIIEDGEALLOWANCE, StkItem.TTI, TaxRate.TaxRate " & _
                        " FROM         StkItem LEFT OUTER JOIN   TaxRate ON StkItem.TTI = TaxRate.Code WHERE  StkItem. StockLink = " & stklnk
                End If
                DS_ITEMS = objSQL.GET_INSERT_UPDATE(SQL)

                If DS_ITEMS.Tables(0).Rows.Count = 0 Then
                    MsgBox("Error in stock Details", MsgBoxStyle.Critical, "SPIL Glass")
                    'objSQL.Rollback_Trans()
                    Exit Sub
                End If

                For Each dr1 In DS_ITEMS.Tables(0).Rows

                    If activeTaxExempt = True Then
                        UG.ActiveRow.Cells("TaxExempt").Value = dr1("TaxExempt")
                        v = dr1("TaxExempt")
                    Else
                        isTaxExempt = False
                    End If

                    If isTaxExempt = True Then
                        UG.ActiveRow.Cells("TaxType").Value = "0"
                        UG.ActiveRow.Cells("TaxRate").Value = 0
                    Else
                        If geographicRegionID = GeographicRegion.NorthAmerica Then
                            UG.ActiveRow.Cells("TaxType").Value = dCusTaxCode
                            UG.ActiveRow.Cells("TaxRate").Value = dCusTaxRate
                        Else
                            If (cust_taxcode = "" Or cust_taxrate = 0) Then
                                UG.ActiveRow.Cells("TaxType").Value = dr1("TTI")
                                UG.ActiveRow.Cells("TaxRate").Value = dr1("TaxRate")
                            Else
                                UG.ActiveRow.Cells("TaxType").Value = cust_taxcode
                                UG.ActiveRow.Cells("TaxRate").Value = cust_taxrate
                            End If
                        End If
                    End If


                    mysegid = 0 ' dr1("iInvSegValue1ID")
                    myStockLink = dr1("StockLink")

                    UG.ActiveRow.Cells("Price_Excl").Value = 0
                    If boEditSrv = False Then
                        UG.ActiveRow.Cells("Units").Value = 0
                    End If

                    ' <Modified by: Administrator at 24/10/2010-9:18:18 AM on machine: GIMHAN-PC>
                    If IsDBNull(UG.ActiveRow.Cells("Edging_ID").Value) = True Then
                        UG.ActiveRow.Cells("Edging").Value = 1
                        UG.ActiveRow.Cells("Edging_ID").Value = 1
                    Else
                        If UG.ActiveRow.Cells("Edging_ID").Value = 0 Or UG.ActiveRow.Cells("Edging_ID").Value = Nothing Then
                            UG.ActiveRow.Cells("Edging").Value = 1
                            UG.ActiveRow.Cells("Edging_ID").Value = 1
                        End If
                    End If
                    ' </Modified by: Administrator at 24/10/2010-9:18:18 AM on machine: GIMHAN-PC>


                    intSrvPriceTp = dr1("uiIISRVPRICEID")

                    For Each MyDataRow In MyDataSet.Tables(0).Rows
                        If MyDataRow(0) = dr1("uiIISRVPRICEID") Then
                            UG.ActiveRow.Cells("Pricing").Value = MyDataRow(1)   'SRVPRICE_DESCRIPTION - Area/Flat/Leanel/None
                            UG.ActiveRow.Cells("Pricing_ID").Value = MyDataRow(0)  'SRVPRICEID   -1/2/3/4
                            Exit For
                        End If
                    Next

                    If IsDBNull(UG.ActiveRow.Cells("Pricing_ID").Value) Then
                        MsgBox("Please assign valid price type for service item " & UG.ActiveRow.Cells("Description_1").Value)
                        UG.ActiveRow.Delete(False)
                        Exit Sub
                    End If
                    If UG.ActiveRow.Cells("Pricing_ID").Value = 0 Then
                        MsgBox("Please assign valid price type for service item " & UG.ActiveRow.Cells("Description_1").Value)
                        Exit Sub
                    End If
                Next

                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

                dr1 = Nothing
                DS_ITEMS = Nothing



            Else

                UG.ActiveRow.Cells("TaxType").Value = ""  '  dr1("TTI")
                UG.ActiveRow.Cells("TaxRate").Value = 0 'dr1("TaxRate")

                mysegid = 2 'dr1("iInvSegValue1ID")
                myStockLink = 0 ' dr1("StockLink")

                UG.ActiveRow.Cells("Price_Excl").Value = 0
                UG.ActiveRow.Cells("Units").Value = 0

                UG.ActiveRow.Cells("Edging").Value = 1
                UG.ActiveRow.Cells("Edging_ID").Value = 1
                'Set Pricing method
                '1= Area
                '2= flat
                '3= leaneal

                intSrvPriceTp = 0 ' dr1("uiIISRVPRICEID")

                UG.ActiveRow.Cells("Pricing").Value = ""
                UG.ActiveRow.Cells("Pricing_ID").Value = 0 ' MyDataRow(0)

            End If

            ' <Modified by: Administrator at 17/04/2011-8:38:07 AM on machine: GIMHAN-PC>
            ''Call Edge_Setup(intSrvPriceTp, 1, UG.ActiveRow.Index)  ' setup data entry screen ,this will setup according to the price category

            ''Call Set_UnitPrice(UG.ActiveRow.Cells("StockLink").Value, UG.ActiveRow.Index)   ' Select Price
            ' </Modified by: Administrator at 17/04/2011-8:38:07 AM on machine: GIMHAN-PC>
            FilterServiceCatgory()
            boAftrCelUpdate = False

            'Call Value_Calculation()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Exclamation, "SPIL Glass")
        Finally
            boAftrCelUpdate = True
        End Try

    End Sub

    Private Sub RowChange()
        Try
            If Me.UG.ActiveCell Is Nothing Then Return

            If UG.ActiveCell.Column.Key = "Code" Then
                UG.ActiveCell.Value = CStr(UG.ActiveCell.Text)

                If DDStock.ActiveRow.Activated = True Then
                    UG.ActiveRow.Cells("Code").Value = DDStock.ActiveRow.Cells("cSimpleCode").Value
                    UG.ActiveRow.Cells("Description_1").Value = DDStock.ActiveRow.Cells("Description_1").Value
                    UG.ActiveRow.Cells("Comment").Value = DDStock.ActiveRow.Cells("Description_1").Value
                    UG.ActiveRow.Cells("StockLink").Value = DDStock.ActiveRow.Cells("StockLink").Value
                    UG.ActiveRow.Cells("Price_Excl").Activation = Activation.Disabled
                    Call SetItemDetail(DDStock.ActiveRow.Cells("StockLink").Value, UG.ActiveRow.Index)

                    If UG.ActiveRow.Cells("Pricing_ID").Value = 1 Then 'Area
                        UG.ActiveRow.Cells("Units").Activation = Activation.Disabled
                    ElseIf UG.ActiveRow.Cells("Pricing_ID").Value = 2 Then 'Flat
                        UG.ActiveRow.Cells("Edging").Activation = Activation.Disabled
                        If boEditSrv = False Then UG.ActiveRow.Cells("Units").Value = 1
                        UG.ActiveRow.Cells("Units").Activation = Activation.AllowEdit
                    ElseIf UG.ActiveRow.Cells("Pricing_ID").Value = 3 Then 'Leanel
                        UG.ActiveRow.Cells("Edging").Activation = Activation.AllowEdit
                        UG.ActiveRow.Cells("Units").Activation = Activation.Disabled
                    ElseIf UG.ActiveRow.Cells("Pricing_ID").Value = 4 Then 'None
                        If boEditSrv = False Then UG.ActiveRow.Cells("Units").Value = 1
                        UG.ActiveRow.Cells("Edging").Activation = Activation.Disabled
                        UG.ActiveRow.Cells("Units").Activation = Activation.Disabled
                    End If
                    Dim myLinkID As Integer = 0
                    Dim ddr As UltraGridRow
                    For Each ddr In DDEdging.DisplayLayout.Rows
                        If ddr.Cells("Edging").Value = UG.ActiveRow.Cells("Edging_ID").Value Then
                            myLinkID = ddr.Cells("AutoLinkSrvID").Value
                            Exit For
                        End If
                    Next
                End If

                If IsDBNull(UG.ActiveCell.Row.Cells("StockLink").Value) = False Then
                    If UG.ActiveCell.Row.Cells("StockLink").Value > 0 Then
                        If GlassInventoryModule.Program.InventoryDatabaseConnection(strSVR_Name, strSVR_DBName, strSVR_UserName, strSVR_PW) = 0 Then
                            MessageBox.Show("Fail to connect inventory Module")
                            Exit Sub
                        End If
                        Dim oItem As SPIL.Glass.InventoryItem
                        oItem = New SPIL.Glass.InventoryItem().GetItem(CInt(UG.ActiveCell.Row.Cells("StockLink").Value))
                        UG.ActiveCell.Row.Cells("IsExternalItem").Value = oItem.IsExternalItem
                    End If
                End If

            ElseIf UG.ActiveCell.Column.Index = 3 Then 'Item Description

                UG.ActiveCell.Value = CStr(UG.ActiveCell.Text)
                If DDDescription.ActiveRow.Activated = True Then
                    UG.ActiveRow.Cells("Code").Value = DDDescription.ActiveRow.Cells("cSimpleCode").Value
                    UG.ActiveRow.Cells("Description_1").Value = DDDescription.ActiveRow.Cells("Description_1").Value
                    UG.ActiveRow.Cells("Comment").Value = DDDescription.ActiveRow.Cells("Description_1").Value
                    UG.ActiveRow.Cells("StockLink").Value = DDDescription.ActiveRow.Cells("StockLink").Value
                    UG.ActiveRow.Cells("Price_Excl").Activation = Activation.Disabled
                    Call SetItemDetail(DDDescription.ActiveRow.Cells("StockLink").Value, UG.ActiveRow.Index)
                    If UG.ActiveRow.Cells("Pricing_ID").Value = 1 Then 'Area
                        UG.ActiveRow.Cells("Units").Activation = Activation.Disabled
                    ElseIf UG.ActiveRow.Cells("Pricing_ID").Value = 2 Then 'Flat
                        UG.ActiveRow.Cells("Edging").Activation = Activation.Disabled
                        If boEditSrv = False Then UG.ActiveRow.Cells("Units").Value = 1
                        UG.ActiveRow.Cells("Units").Activation = Activation.AllowEdit
                    ElseIf UG.ActiveRow.Cells("Pricing_ID").Value = 3 Then 'Leanel
                        UG.ActiveRow.Cells("Edging").Activation = Activation.AllowEdit
                        UG.ActiveRow.Cells("Units").Activation = Activation.Disabled
                    ElseIf UG.ActiveRow.Cells("Pricing_ID").Value = 4 Then 'None
                        If boEditSrv = False Then UG.ActiveRow.Cells("Units").Value = 1
                        UG.ActiveRow.Cells("Edging").Activation = Activation.Disabled
                        UG.ActiveRow.Cells("Units").Activation = Activation.Disabled
                    End If
                    Dim myLinkID As Integer = 0
                    Dim ddr As UltraGridRow
                    For Each ddr In DDEdging.DisplayLayout.Rows
                        If ddr.Cells("Edging").Value = UG.ActiveRow.Cells("Edging_ID").Value Then
                            myLinkID = ddr.Cells("AutoLinkSrvID").Value
                            Exit For
                        End If
                    Next
                End If
                If IsDBNull(UG.ActiveCell.Row.Cells("StockLink").Value) = False Then
                    If UG.ActiveCell.Row.Cells("StockLink").Value > 0 Then
                        If GlassInventoryModule.Program.InventoryDatabaseConnection(strSVR_Name, strSVR_DBName, strSVR_UserName, strSVR_PW) = 0 Then
                            MessageBox.Show("Fail to connect inventory Module")
                            Exit Sub
                        End If
                        Dim oItem As SPIL.Glass.InventoryItem
                        oItem = New SPIL.Glass.InventoryItem().GetItem(CInt(UG.ActiveCell.Row.Cells("StockLink").Value))
                        UG.ActiveCell.Row.Cells("IsExternalItem").Value = oItem.IsExternalItem
                    End If
                End If

            ElseIf UG.ActiveCell.Column.Index = 4 Then 'Pricing
                UG.ActiveCell.Value = CDbl(UG.ActiveCell.Text)
                If DDPricing.ActiveRow.Activated = True Then
                    UG.ActiveRow.Cells("Pricing").Value = DDPricing.ActiveRow.Cells("Pricing").Value
                    UG.ActiveRow.Cells("Pricing_ID").Value = DDPricing.ActiveRow.Cells("ID").Value
                End If

            ElseIf UG.ActiveCell.Column.Index = 7 Then ' Edging
                UG.ActiveCell.Value = CStr(UG.ActiveCell.Text)
                Dim ddr As UltraGridRow
                For Each ddr In DDEdging.DisplayLayout.Rows
                    If ddr.Cells("Description").Value = UG.ActiveCell.Value Then
                        ddr.Activate()
                        ddr.Activated = True
                        Exit For
                    End If
                Next
                If DDEdging.ActiveRow.Activated = True Then
                    UG.ActiveRow.Cells("Edging").Value = DDEdging.ActiveRow.Cells("Description").Value
                    UG.ActiveRow.Cells("Edging_ID").Value = DDEdging.ActiveRow.Cells("Edging").Value
                    Call Edge_Setup(UG.ActiveRow.Cells("Pricing_ID").Value, DDEdging.ActiveRow.Cells("Edging").Value, UG.ActiveRow.Index)
                    Dim myLinkID As Integer = 0
                    For Each ddr In DDEdging.DisplayLayout.Rows
                        If ddr.Cells("Edging").Value = CInt(UG.ActiveRow.Cells("Edging_ID").Value) Then
                            myLinkID = ddr.Cells("AutoLinkSrvID").Value
                            Exit For
                        End If
                    Next
                End If

            ElseIf UG.ActiveCell.Column.Key = "Price_Excl" Then '12 'Price Excl
                UG.ActiveCell.Value = CDbl(UG.ActiveCell.Text)
                If UG.ActiveRow.Cells("OrgPrice").Value = 0 Then
                    UG.ActiveRow.Cells("OrgPrice").Value = UG.ActiveCell.Value
                End If
                If UG.ActiveRow.Cells("OrgPrice").Value > 0 Then
                    UG.ActiveRow.Cells("Disc_Percentage").Value = Math.Round(((UG.ActiveRow.Cells("OrgPrice").Value - UG.ActiveRow.Cells("Price_Excl").Value) * 100) / UG.ActiveRow.Cells("OrgPrice").Value, 2)
                Else
                    UG.ActiveRow.Cells("Disc_Percentage").Value = 0
                End If
                UG.ActiveRow.Cells("SurChrg").Value = 0
                EditPricesOnServices(UG.ActiveRow, CDbl(UG.ActiveRow.Cells("Price_Excl").Value))
            ElseIf UG.ActiveCell.Column.Key = "LinkedSrvCatID" Then

                GetBillOfMaterials()
            End If
        Catch ex As Exception
            Exit Sub
        End Try

    End Sub

    Sub UGGridColumsHandler(Optional ByRef e As BeforeExitEditModeEventArgs = Nothing)
        Try
            If Me.UG.ActiveCell Is Nothing Then Return

            If UG.ActiveCell.Column.Key = "Code" Then
                UG.ActiveCell.Value = CStr(UG.ActiveCell.Text)

                If DDStock.ActiveRow.Activated = True Then
                    UG.ActiveRow.Cells("Code").Value = DDStock.ActiveRow.Cells("cSimpleCode").Value
                    UG.ActiveRow.Cells("Description_1").Value = DDStock.ActiveRow.Cells("Description_1").Value
                    UG.ActiveRow.Cells("Comment").Value = DDStock.ActiveRow.Cells("Description_1").Value
                    UG.ActiveRow.Cells("StockLink").Value = DDStock.ActiveRow.Cells("StockLink").Value
                    UG.ActiveRow.Cells("Price_Excl").Activation = Activation.Disabled
                    Call SetItemDetail(DDStock.ActiveRow.Cells("StockLink").Value, UG.ActiveRow.Index)


                    If UG.ActiveRow.Cells("Pricing_ID").Value = 1 Then 'Area
                        ''UG.ActiveRow.Cells("Edging").Activation = Activation.Disabled
                        UG.ActiveRow.Cells("Units").Activation = Activation.Disabled
                    ElseIf UG.ActiveRow.Cells("Pricing_ID").Value = 2 Then 'Flat
                        UG.ActiveRow.Cells("Edging").Activation = Activation.Disabled
                        If boEditSrv = False Then UG.ActiveRow.Cells("Units").Value = 1
                        UG.ActiveRow.Cells("Units").Activation = Activation.AllowEdit
                    ElseIf UG.ActiveRow.Cells("Pricing_ID").Value = 3 Then 'Leanel
                        UG.ActiveRow.Cells("Edging").Activation = Activation.AllowEdit
                        UG.ActiveRow.Cells("Units").Activation = Activation.Disabled
                    ElseIf UG.ActiveRow.Cells("Pricing_ID").Value = 4 Then 'None
                        If boEditSrv = False Then UG.ActiveRow.Cells("Units").Value = 1
                        UG.ActiveRow.Cells("Edging").Activation = Activation.Disabled
                        UG.ActiveRow.Cells("Units").Activation = Activation.Disabled
                    End If
                    'UG.ActiveRow.Cells("Pricing").Value = MyDataRow(1)   'SRVPRICE_DESCRIPTION - Area/Flat/Leanel/None
                    'UG.ActiveRow.Cells("Pricing_ID").Value = MyDataRow(0)  'SRVPRICEID   -1/2/3/4


                    ' <Added by: Administrator at: 24/10/2010-9:03:36 AM on machine: GIMHAN-PC>
                    Dim myLinkID As Integer = 0
                    Dim ddr As UltraGridRow
                    For Each ddr In DDEdging.DisplayLayout.Rows
                        If ddr.Cells("Edging").Value = UG.ActiveRow.Cells("Edging_ID").Value Then
                            myLinkID = ddr.Cells("AutoLinkSrvID").Value
                            Exit For
                        End If
                    Next

                    'Call AutoServiceAdd(UG.ActiveRow.Cells("StockLink").Value, CInt(UG.ActiveRow.Cells("Edging").Value), myLinkID)
                    ' </Added by: Administrator at: 24/10/2010-9:03:36 AM on machine: GIMHAN-PC>
                End If

                If IsDBNull(UG.ActiveCell.Row.Cells("StockLink").Value) = False Then
                    If UG.ActiveCell.Row.Cells("StockLink").Value > 0 Then
                        If GlassInventoryModule.Program.InventoryDatabaseConnection(strSVR_Name, strSVR_DBName, strSVR_UserName, strSVR_PW) = 0 Then
                            MessageBox.Show("Fail to connect inventory Module")
                            Exit Sub
                        End If
                        Dim oItem As SPIL.Glass.InventoryItem
                        oItem = New SPIL.Glass.InventoryItem().GetItem(CInt(UG.ActiveCell.Row.Cells("StockLink").Value))

                        UG.ActiveCell.Row.Cells("IsExternalItem").Value = oItem.IsExternalItem

                    End If
                End If

            ElseIf UG.ActiveCell.Column.Index = 3 Then 'Item Description

                UG.ActiveCell.Value = CStr(UG.ActiveCell.Text)

                If DDDescription.ActiveRow.Activated = True Then
                    UG.ActiveRow.Cells("Code").Value = DDDescription.ActiveRow.Cells("cSimpleCode").Value
                    UG.ActiveRow.Cells("Description_1").Value = DDDescription.ActiveRow.Cells("Description_1").Value
                    UG.ActiveRow.Cells("Comment").Value = DDDescription.ActiveRow.Cells("Description_1").Value
                    UG.ActiveRow.Cells("StockLink").Value = DDDescription.ActiveRow.Cells("StockLink").Value
                    UG.ActiveRow.Cells("Price_Excl").Activation = Activation.Disabled
                    Call SetItemDetail(DDDescription.ActiveRow.Cells("StockLink").Value, UG.ActiveRow.Index)


                    If UG.ActiveRow.Cells("Pricing_ID").Value = 1 Then 'Area
                        ''UG.ActiveRow.Cells("Edging").Activation = Activation.Disabled
                        UG.ActiveRow.Cells("Units").Activation = Activation.Disabled
                    ElseIf UG.ActiveRow.Cells("Pricing_ID").Value = 2 Then 'Flat
                        UG.ActiveRow.Cells("Edging").Activation = Activation.Disabled
                        If boEditSrv = False Then UG.ActiveRow.Cells("Units").Value = 1
                        UG.ActiveRow.Cells("Units").Activation = Activation.AllowEdit
                    ElseIf UG.ActiveRow.Cells("Pricing_ID").Value = 3 Then 'Leanel
                        UG.ActiveRow.Cells("Edging").Activation = Activation.AllowEdit
                        UG.ActiveRow.Cells("Units").Activation = Activation.Disabled
                    ElseIf UG.ActiveRow.Cells("Pricing_ID").Value = 4 Then 'None
                        If boEditSrv = False Then UG.ActiveRow.Cells("Units").Value = 1
                        UG.ActiveRow.Cells("Edging").Activation = Activation.Disabled
                        UG.ActiveRow.Cells("Units").Activation = Activation.Disabled
                    End If
                    'UG.ActiveRow.Cells("Pricing").Value = MyDataRow(1)   'SRVPRICE_DESCRIPTION - Area/Flat/Leanel/None
                    'UG.ActiveRow.Cells("Pricing_ID").Value = MyDataRow(0)  'SRVPRICEID   -1/2/3/4


                    ' <Added by: Administrator at: 24/10/2010-9:03:36 AM on machine: GIMHAN-PC>
                    Dim myLinkID As Integer = 0
                    Dim ddr As UltraGridRow
                    For Each ddr In DDEdging.DisplayLayout.Rows
                        If ddr.Cells("Edging").Value = UG.ActiveRow.Cells("Edging_ID").Value Then
                            myLinkID = ddr.Cells("AutoLinkSrvID").Value
                            Exit For
                        End If
                    Next

                    'Call AutoServiceAdd(UG.ActiveRow.Cells("StockLink").Value, CInt(UG.ActiveRow.Cells("Edging").Value), myLinkID)
                    ' </Added by: Administrator at: 24/10/2010-9:03:36 AM on machine: GIMHAN-PC>

                End If
                If IsDBNull(UG.ActiveCell.Row.Cells("StockLink").Value) = False Then
                    If UG.ActiveCell.Row.Cells("StockLink").Value > 0 Then
                        If GlassInventoryModule.Program.InventoryDatabaseConnection(strSVR_Name, strSVR_DBName, strSVR_UserName, strSVR_PW) = 0 Then
                            MessageBox.Show("Fail to connect inventory Module")
                            Exit Sub
                        End If
                        Dim oItem As SPIL.Glass.InventoryItem
                        oItem = New SPIL.Glass.InventoryItem().GetItem(CInt(UG.ActiveCell.Row.Cells("StockLink").Value))

                        UG.ActiveCell.Row.Cells("IsExternalItem").Value = oItem.IsExternalItem

                    End If
                End If

            ElseIf UG.ActiveCell.Column.Key = "Pricing" Then 'Pricing

                UG.ActiveCell.Value = CDbl(UG.ActiveCell.Text)

                If DDPricing.ActiveRow.Activated = True Then

                    UG.ActiveRow.Cells("Pricing").Value = DDPricing.ActiveRow.Cells("Pricing").Value
                    UG.ActiveRow.Cells("Pricing_ID").Value = DDPricing.ActiveRow.Cells("ID").Value

                    '  Call Set_UnitPrice(UG.ActiveRow.Cells("StockLink").Value, UG.ActiveRow.Index)   ' Select Price

                End If

            ElseIf UG.ActiveCell.Column.Index = 7 Then ' Edging

                UG.ActiveCell.Value = CStr(UG.ActiveCell.Text)

                Dim ddr As UltraGridRow
                For Each ddr In DDEdging.DisplayLayout.Rows
                    If ddr.Cells("Description").Value = UG.ActiveCell.Value Then
                        ddr.Activate()
                        ddr.Activated = True
                        Exit For
                    End If
                Next

                If DDEdging.ActiveRow.Activated = True Then

                    UG.ActiveRow.Cells("Edging").Value = DDEdging.ActiveRow.Cells("Description").Value
                    UG.ActiveRow.Cells("Edging_ID").Value = DDEdging.ActiveRow.Cells("Edging").Value

                    Call Edge_Setup(UG.ActiveRow.Cells("Pricing_ID").Value, DDEdging.ActiveRow.Cells("Edging").Value, UG.ActiveRow.Index)

                    ' <Added by: Administrator at: 24/10/2010-9:03:12 AM on machine: GIMHAN-PC>
                    Dim myLinkID As Integer = 0
                    'Dim ddr As UltraGridRow
                    For Each ddr In DDEdging.DisplayLayout.Rows
                        If ddr.Cells("Edging").Value = CInt(UG.ActiveRow.Cells("Edging_ID").Value) Then
                            myLinkID = ddr.Cells("AutoLinkSrvID").Value
                            Exit For
                        End If
                    Next

                    'Call AutoServiceAdd(UG.ActiveRow.Cells("StockLink").Value, CInt(Val(UG.ActiveRow.Cells("Edging_ID").Value)), myLinkID)
                    ' </Added by: Administrator at: 24/10/2010-9:03:12 AM on machine: GIMHAN-PC>

                End If

            ElseIf UG.ActiveCell.Column.Key = "Price_Excl" Then '12 'Price Excl
                UG.ActiveCell.Value = CDbl(UG.ActiveCell.Text)
                If UG.ActiveRow.Cells("OrgPrice").Value = 0 Then
                    UG.ActiveRow.Cells("OrgPrice").Value = UG.ActiveCell.Value
                End If
                If UG.ActiveRow.Cells("OrgPrice").Value > 0 Then
                    UG.ActiveRow.Cells("Disc_Percentage").Value = Math.Round(((UG.ActiveRow.Cells("OrgPrice").Value - UG.ActiveRow.Cells("Price_Excl").Value) * 100) / UG.ActiveRow.Cells("OrgPrice").Value, 2)
                Else
                    UG.ActiveRow.Cells("Disc_Percentage").Value = 0
                End If

                UG.ActiveRow.Cells("SurChrg").Value = 0


                ''If boEditSrv = False Then
                ''    AddOREditServiceToGlassItems(UG.ActiveRow, CDbl(UG.ActiveRow.Cells("Price_Excl").Value))
                ''    If UG.ActiveRow.Cells("Pricing_ID").Value <> GlassServiceUnits.None Then
                ''        UG.ActiveRow.Cells("Price_Excl").Activation = Activation.Disabled
                ''    End If
                ''End If


                EditPricesOnServices(UG.ActiveRow, CDbl(UG.ActiveRow.Cells("Price_Excl").Value))


                ''ElseIf UG.ActiveCell.Column.Key = "PriceCat" Then '26 Price Category
                ''    UG.ActiveCell.Value = CStr(UG.ActiveCell.Text)
                ''    If UG.ActiveCell.Row.Cells("PriceCat").Value.ToString = "T" Or UG.ActiveCell.Row.Cells("PriceCat").Value = "S" Then
                ''        'Call Set_UnitPriceByCategory(UG.ActiveRow.Cells("StockLink").Value, UG.ActiveRow.Index, UG.ActiveCell.Row.Cells("PriceCat").Value.ToString)
                ''    Else
                ''        e.Cancel = True
                ''    End If

                ''ElseIf UG.ActiveCell.Column.Index = 23 Then 'Discount %
                ''    UG.ActiveCell.Value = CDbl(UG.ActiveCell.Text)
                ''    UG.ActiveRow.Cells("Price_Excl").Value = UG.ActiveRow.Cells("OrgPrice").Value - ((UG.ActiveRow.Cells("OrgPrice").Value * UG.ActiveRow.Cells("Disc_Percentage").Value) / 100)
            ElseIf UG.ActiveCell.Column.Key = "LinkedSrvCatID" Then
                GetBillOfMaterials()
            End If
            Edge_Setup_NEW(ug2Row, UG.ActiveRow)
            Set_UnitPrice_NEW(UG.ActiveRow, UG.ActiveRow)
        Catch ex As Exception
            'MsgBox(Err.Description)
            modGlazingQuoteExtension.GQShowMessage(ex.Message, Me.Text, MsgBoxStyle.Critical, "warning", "UGGridColumsHandler")

            Exit Sub
        End Try
    End Sub

    Sub CalaulateTaxAmount()
        Try
            Dim netExcAmount As Double
            Dim excAmount As Double
            Dim qty As Double
            Dim taxAmount As Double
            Dim totalExcAmount As Double
            Dim isTaxExempt As Boolean
            If activeTaxExempt = True Then
                isTaxExempt = gridActiveRow.Cells("TaxExempt").Value
            Else
                isTaxExempt = False
            End If

            If isLoading = False Then
                If gridActiveCell.Column.Key = "Price_Excl" Then
                    excAmount = If(IsNothing(gridActiveRow.Cells("Price_Excl").Text) = False, gridActiveCell.Value, 0)
                    qty = gridActiveRow.Cells("Units").Value
                    totalExcAmount = excAmount * qty
                    netExcAmount = totalExcAmount * mainGlassQty
                    taxAmount = clsGlazingDocStockItemHelperObj.CalculateTaxAmount(dCusTaxRate, totalExcAmount, isTaxExempt)

                    gridActiveRow.Cells("TaxRate").Value = dCusTaxRate
                    gridActiveRow.Cells("LineTotal_Excl").Value = totalExcAmount
                    gridActiveRow.Cells("LineNet_Excl").Value = netExcAmount
                    gridActiveRow.Cells("LineTax").Value = taxAmount
                    gridActiveRow.Cells("NetAmount").Value = netExcAmount + taxAmount * mainGlassQty

                End If
            End If
        Catch ex As Exception
            modGlazingQuoteExtension.GQShowMessage(ex.Message, Me.Text, MsgBoxStyle.Critical, "warning", "CalaulateTaxAmount")

        End Try
    End Sub

#End Region
#Region "Calculations"

    Private Function FindThicknessID(ByVal dblntThik As Double) As Integer
        Dim a As Integer = 0
        Try
            For Each MyDataRow In MyDataSet.Tables(2).Rows
                If CType(dblntThik, Double) <= CType(MyDataRow(1), Double) Then
                    a = MyDataRow(0)
                    Exit For
                End If
            Next
            If a = 0 Then a = 1
            intThick_ID = a
            Return a
        Catch ex As Exception
            modGlazingQuoteExtension.GQShowMessage(ex.Message, Me.Text, MsgBoxStyle.Critical, "warning", "FindThicknessID")
            Return a
        End Try
    End Function

    Private Sub Edge_Setup(ByVal priceType As Integer, ByVal SRVPRICE_CATID As Integer, ByVal intIndex As Integer)
        'When Adding a Single Service 
        Dim dblLng As Double
        Dim dblshrt As Double
        Dim dblUnits As Double

        Try

            '1= Area
            '2= flat
            '3= leaneal
            '4=None

            If priceType = 3 Then  '3= Leaneal  - Flat Polish

                If CDbl(decHeight) > CDbl(decWidth) Then
                    dblLng = decHeight
                    dblshrt = decWidth
                Else
                    dblLng = decWidth
                    dblshrt = decHeight
                End If
                MyDataRow = MyDataSet.Tables(3).Rows(SRVPRICE_CATID - 1)

                UG.Rows(intIndex).Cells("QTY_LONG").Value = CType(MyDataRow("QTY_LONG"), Int16)

                UG.Rows(intIndex).Cells("QTY_SHORT").Value = CType(MyDataRow("QTY_SHORT"), Int16)
                'Comment By Others not Chathura
                'dblUnits = System.Math.Round(((dblLng * (UG.Rows(intIndex).Cells("QTY_LONG").Value) * CInt(frmSO.UG2.Rows(intSOLine).Cells(6).Value)) / 1000), 4)
                'dblUnits = dblUnits + System.Math.Round(((dblshrt * (UG.Rows(intIndex).Cells("QTY_SHORT").Value) * CInt(frmSO.UG2.Rows(intSOLine).Cells(6).Value)) / 1000), 4)

                If geographicRegionID = GeographicRegion.NorthAmerica Then
                    dblUnits = System.Math.Round(((dblLng * (UG.Rows(intIndex).Cells("QTY_LONG").Value))), 4)
                    dblUnits = dblUnits + System.Math.Round(((dblshrt * (UG.Rows(intIndex).Cells("QTY_SHORT").Value))), 4)
                Else
                    dblUnits = System.Math.Round(((dblLng * (UG.Rows(intIndex).Cells("QTY_LONG").Value)) / 1000), 4)
                    dblUnits = dblUnits + System.Math.Round(((dblshrt * (UG.Rows(intIndex).Cells("QTY_SHORT").Value)) / 1000), 4)
                End If

                UG.Rows(intIndex).Cells("Units").Value = dblUnits

                UG.ActiveRow.Cells("Units").Activation = Activation.Disabled
                UG.ActiveRow.Cells("Edging").Activation = Activation.AllowEdit

                ' UG.ActiveRow.Cells("Price_Excl").Activation = Activation.Disabled


            ElseIf priceType = 2 Then '2= flat - Hole Works

                UG.Rows(intIndex).Cells("Units").Value = 1
                UG.ActiveRow.Cells("Units").Activation = Activation.AllowEdit
                UG.ActiveRow.Cells("Edging").Activation = Activation.Disabled
                'UG.ActiveRow.Cells("Price_Excl").Activation = Activation.Disabled


            ElseIf priceType = 1 Then '1= Area - Painting
                UG.ActiveRow.Cells("Units").Activation = Activation.Disabled
                ''UG.ActiveRow.Cells("Edging").Activation = Activation.Disabled

                UG.Rows(intIndex).Cells("Units").Value = UG.ActiveRow.Cells("Volume").Value
                ' UG.ActiveRow.Cells("Price_Excl").Activation = Activation.Disabled


            ElseIf priceType = 4 Then '4= None 'Delivery Charges

                UG.Rows(intIndex).Cells("Units").Value = 1
                UG.ActiveRow.Cells("Units").Activation = Activation.Disabled
                UG.ActiveRow.Cells("Edging").Activation = Activation.Disabled
                UG.Rows(intIndex).Cells("Peices").Value = 1

                ' UG.ActiveRow.Cells("Price_Excl").Activation = Activation.Disabled
            End If


        Catch ex As Exception

        Finally

        End Try

    End Sub

    Private Sub EditPricesOnServices(ByRef ugRMainServ As UltraGridRow, ByVal dblUnitPrice As Double)
        Dim drGlass, ugChildRow As UltraGridRow

        For Each drGlass In ugGlassItems.DisplayLayout.Rows

            If drGlass.Band.Index = 0 Then

                Dim x As Integer = drGlass.Cells("UniqueLN").Value

                ugGlassItems.ActiveRow = drGlass

                If drGlass.ChildBands.HasChildRows Then
                    Dim IsServiceFound As Boolean = False

                    If boEditSrv = False Then
                        For Each ugChildRow In drGlass.ChildBands(0).Rows.All
                            If ugChildRow.Cells("StockLink").Value = ugRMainServ.Cells("StockLink").Value And _
                            ugChildRow.Cells("LN").Value = ugRMainServ.Index Then
                                ugChildRow.Cells("IsExternalItem").Value = ugRMainServ.Cells("IsExternalItem").Value
                                ugChildRow.Cells("Price_Excl").Value = dblUnitPrice
                                CalculateLineAmounts(drGlass, ugChildRow)

                                Exit For
                            End If
                        Next
                    Else
                        For Each ugChildRow In drGlass.ChildBands(0).Rows.All
                            If ugChildRow.Cells("StockLink").Value = ugRMainServ.Cells("StockLink").Value Then
                                ugChildRow.Cells("IsExternalItem").Value = ugRMainServ.Cells("IsExternalItem").Value
                                ugChildRow.Cells("Price_Excl").Value = dblUnitPrice
                                CalculateLineAmounts(drGlass, ugChildRow)
                            End If
                        Next
                    End If
                End If
            End If

        Next
    End Sub

    Private Sub CalculateLineAmounts(ByRef ugParentGlassRow As UltraGridRow, ByRef ugChildServiceRow As UltraGridRow)

        Dim disamt As Integer = ugChildServiceRow.Cells("Disc_Amt").Value
        Dim disPercent As Integer = ugChildServiceRow.Cells("Disc_Percentage").Value
        Dim lineNet As Decimal = 0
        lineNet = ((ugChildServiceRow.Cells("SQFeetForPricing").Value * ugChildServiceRow.Cells("Peices").Value) * ugChildServiceRow.Cells("Price_Excl").Value) + ugChildServiceRow.Cells("SurChrg").Value
        disamt = (lineNet * ugChildServiceRow.Cells("Disc_Percentage").Value) / 100
        ugChildServiceRow.Cells("Disc_Amt").Value = disamt
        ugChildServiceRow.Cells("LineNet_Excl").Value = Math.Round((ugChildServiceRow.Cells("SQFeetForPricing").Value * ugChildServiceRow.Cells("Peices").Value * (ugChildServiceRow.Cells("Price_Excl").Value + ugChildServiceRow.Cells("SurChrg").Value) - disamt), 2, MidpointRounding.AwayFromZero)
        ugChildServiceRow.Cells("LineTax").Value = Math.Round((ugChildServiceRow.Cells("LineNet_Excl").Value * ugChildServiceRow.Cells("TaxRate").Value) / 100, 2, MidpointRounding.AwayFromZero)
        ugChildServiceRow.Cells("LineTotal_Excl").Value = Math.Round(ugChildServiceRow.Cells("LineTax").Value + ugChildServiceRow.Cells("LineNet_Excl").Value, 2, MidpointRounding.AwayFromZero)
        ugChildServiceRow.Cells("TotalUnits").Value = Math.Round(ugChildServiceRow.Cells("Units").Value * ugChildServiceRow.Cells("Peices").Value, 4, MidpointRounding.AwayFromZero)

        ''job
        If ugChildServiceRow.Cells("Cost").Value <= 0 Or ugChildServiceRow.Cells("Price_Excl").Value <= 0 Then
            ugChildServiceRow.Cells("Profit").Value = 0
            ugChildServiceRow.Cells("Cost").Value = 0

        Else
            Dim discost As Double = System.Math.Round((ugChildServiceRow.Cells("Disc_Percentage").Value * ugChildServiceRow.Cells("Cost").Value / 100), 2, MidpointRounding.AwayFromZero)
            Dim cost As Double = Math.Round(((ugChildServiceRow.Cells("Cost").Value - discost) * ugChildServiceRow.Cells("SQFeetForPricing").Value), 2, MidpointRounding.AwayFromZero)
            ugChildServiceRow.Cells("Profit").Value = Math.Round(ugChildServiceRow.Cells("LineNet_Excl").Value - (cost), 2, MidpointRounding.AwayFromZero)
        End If

    End Sub

    Private Sub Set_UnitPrice_NEW(ByRef ugParentGlassRow As UltraGridRow, ByRef ugChildServiceRow As UltraGridRow)

        Try
            oPriceUnitsMe.oCustomer = New clsCustomer(cusID)

            Dim PriceString As String = oPriceUnitsMe.GetServicePriceOnActiveRow_ServiceScreen(isTempered _
                                                                                       , ugChildServiceRow.Cells("StockLink").Value _
                                                                                       , FindThicknessID(dblThickness) _
                                                                                       , oSOModuleDefaults.ServicePriceTypeID)
            Dim dblPrice As Double = -1
            Dim MyArr() As String = PriceString.Split(";")

            dblPrice = CDbl(MyArr(0))
            PriceCategory = CStr(MyArr(1))

            ugChildServiceRow.Cells("Price_Excl").Value = dblPrice
            OriginalPrice = dblPrice
            ugChildServiceRow.Cells("PriceCat").Value = PriceCategory
            ugChildServiceRow.Cells("OrgPrice").Value = OriginalPrice

            ugChildServiceRow.Cells("Price_Excl").Value = ugChildServiceRow.Cells("OrgPrice").Value - ((ugChildServiceRow.Cells("OrgPrice").Value * ugChildServiceRow.Cells("Disc_Percentage").Value) / 100)
            ugChildServiceRow.Cells("SurChrg").Value = 0

        Catch ex As Exception
            ugChildServiceRow.Cells("Price_Excl").Value = 0
            OriginalPrice = 0
            ugChildServiceRow.Cells("PriceCat").Value = "T"
            ugChildServiceRow.Cells("OrgPrice").Value = 0
            ugChildServiceRow.Cells("Price_Excl").Value = 0
            ugChildServiceRow.Cells("SurChrg").Value = 0
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub Edge_Setup_NEW(ByRef ugParentGlassRow As UltraGridRow, ByRef ugChildServiceRow As UltraGridRow)
        'When Adding a Single Service 
        Dim dblLng As Double
        Dim dblshrt As Double
        Dim dblUnits As Double
        Dim dblLngForPricing As Double
        Dim dblshrtForPricing As Double
        Dim dblSQFeetForPricing As Double
        ' Dim clsInchCalculationsObj As New clsInchCalculations

        Try


            If ugChildServiceRow.Cells("Pricing_ID").Value = GlassServiceUnits.Lineal Then  '3= Leaneal  - Flat Polish

             
                If CDbl(glassHeight) > CDbl(glassWidth) Then
                    dblLng = CDbl(glassHeight)
                    dblshrt = CDbl(glassWidth)
                Else
                    dblLng = CDbl(glassWidth)
                    dblshrt = CDbl(glassHeight)
                End If

                MyDataRow = MyDataSet.Tables(3).Rows(ugChildServiceRow.Cells("Edging_ID").Value - 1)

                ugChildServiceRow.Cells("QTY_LONG").Value = CType(MyDataRow("QTY_LONG"), Int16)

                ugChildServiceRow.Cells("QTY_SHORT").Value = CType(MyDataRow("QTY_SHORT"), Int16)

                If geographicRegionID = GeographicRegion.NorthAmerica Then
                    dblUnits = System.Math.Round(((dblLng * (UG.Rows(intIndex).Cells("QTY_LONG").Value))), 4)
                    dblUnits = dblUnits + System.Math.Round(((dblshrt * (UG.Rows(intIndex).Cells("QTY_SHORT").Value))), 4)
                Else
                    dblUnits = System.Math.Round(((dblLng * (UG.Rows(intIndex).Cells("QTY_LONG").Value)) / 1000), 4)
                    dblUnits = dblUnits + System.Math.Round(((dblshrt * (UG.Rows(intIndex).Cells("QTY_SHORT").Value)) / 1000), 4)
                End If

                ugChildServiceRow.Cells("Units").Value = dblUnits

                dblSQFeetForPricing = System.Math.Round(((dblLngForPricing * (ugChildServiceRow.Cells("QTY_LONG").Value))), 2, MidpointRounding.AwayFromZero)
                dblSQFeetForPricing = dblSQFeetForPricing + System.Math.Round(((dblshrtForPricing * (ugChildServiceRow.Cells("QTY_SHORT").Value))), 2, MidpointRounding.AwayFromZero)

                ugChildServiceRow.Cells("SQFeetForPricing").Value = dblSQFeetForPricing

                ugChildServiceRow.Cells("Units").Activation = Activation.Disabled
                ugChildServiceRow.Cells("Edging").Activation = Activation.AllowEdit


            ElseIf ugChildServiceRow.Cells("Pricing_ID").Value = GlassServiceUnits.Flat Then '2= flat - Hole Works

                'ugChildServiceRow.Cells("Units").Value = 1
                ugChildServiceRow.Cells("SQFeetForPricing").Value = 1

                ugChildServiceRow.Cells("Units").Activation = Activation.AllowEdit
                ugChildServiceRow.Cells("Edging").Activation = Activation.Disabled

            ElseIf ugChildServiceRow.Cells("Pricing_ID").Value = GlassServiceUnits.Area Then '1= Area - Painting
                UG.ActiveRow.Cells("Units").Activation = Activation.Disabled
                UG.ActiveRow.Cells("Edging").Activation = Activation.Disabled

                'ugChildServiceRow.Cells("Units").Value = UG.ActiveRow.Cells("Volume").Value / UG.ActiveRow.Cells("Peices").Value

                ugChildServiceRow.Cells("Units").Value = glassVolume / ugParentGlassRow.Cells("Qty").Value
                If geographicRegionID = GeographicRegion.NorthAmerica Then
                    ugChildServiceRow.Cells("SQFeetForPricing").Value = System.Math.Round(glassVolume / ugParentGlassRow.Cells("Qty").Value, 2)
                End If

            ElseIf ugChildServiceRow.Cells("Pricing_ID").Value = GlassServiceUnits.None Then '4= None 'Delivery Charges

                ugChildServiceRow.Cells("Units").Value = 1
                ugChildServiceRow.Cells("SQFeetForPricing").Value = 1
                ugChildServiceRow.Cells("Units").Activation = Activation.Disabled
                ugChildServiceRow.Cells("Edging").Activation = Activation.Disabled
                ugChildServiceRow.Cells("Peices").Value = 1
            End If


        Catch ex As Exception
            MsgBox(ex.Message)
        Finally

        End Try

    End Sub

    Sub CalculateTotalExc()
        Try
            Dim TotalExc As Double = UG.ActiveRow.Cells("Price_Excl").Value
            Dim volume As Double = UG.ActiveRow.Cells("Volume").Value
            Dim actualEnterdPrice As Double = UG.ActiveRow.Cells("Price_Excl").Value
            Dim quantity As Double = UG.ActiveRow.Cells("Units").Value

            'Get total
            TotalExc = clsGlazingDocStockItemHelperObj.CalculateTotalExc(TotalExc, qty, volume, actualEnterdPrice, quantity)

            'Set total
            UG.ActiveRow.Cells("Price_Excl").Value = TotalExc

        Catch ex As Exception

        End Try
    End Sub

#End Region
#Region "Extra"

    Private Sub tsbAddNewRow_Click(sender As Object, e As EventArgs) Handles tsbAddNewRow.Click
        AddNewRow("end")
    End Sub

    Private Sub Discard()
        DeleteRows()
    End Sub

    Private Sub DeleteRows()

        Dim R As Integer = UG.Rows.Count
        Do Until R = 0
            UG.Rows(R - 1).Delete(False)
            R = R - 1
        Loop

    End Sub

    Private Sub GetBillOfMaterials()
        Try

            If UG.ActiveCell Is Nothing Then Return
            If IsNothing(UG.ActiveRow.Cells("LinkedSrvCatID").Value) Then Return
            Dim objSQL As New clsSqlConn
            Dim strSQL As String
            Dim dsBOM As DataSet
            UG.ActiveRow.Cells("LineNotes").Value = String.Empty
            If IsNumeric(UG.ActiveRow.Cells("LinkedSrvCatID").Value) Then
                strSQL = "SELECT StkItem.Code as cSimpleCode,StkItem.Description_1,  StkItem.StockLink, StkItem_BOM.QTY_Percentage  FROM StkItem INNER JOIN StkItem_BOM ON StkItem_BOM.BOM_StockLink=StkItem.StockLink  WHERE StkItem.ItemActive = 1 AND StkItem_BOM.StockLink=" & UG.ActiveRow.Cells("LinkedSrvCatID").Value & "   order by StkItem.Code"
                dsBOM = objSQL.GET_INSERT_UPDATE(strSQL)
                Dim Description_1 As String = String.Empty
                If dsBOM.Tables.Count > 0 Then
                    For Each dr As DataRow In dsBOM.Tables(0).Rows
                        Description_1 = Description_1 + dr("Description_1") & ":" & dr("QTY_Percentage") & "%|"
                    Next
                    Description_1 = Description_1.Trim.Remove(Description_1.Length - 2)
                End If
                UG.ActiveRow.Cells("LineNotes").Value = UG.ActiveRow.Cells("LinkedSrvCatID").Text & " (" & Description_1 & ")"
                Description_1 = String.Empty
            End If


        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

    End Sub

    Private Sub FilterServiceCatgory()
        Try


            If Not IsNothing(UG.ActiveRow.Cells("StockLink").Value) Then

                If IsNothing(oSOModuleDefaults) Then
                    oSOModuleDefaults = New clsSOModuleDefaults
                End If

                If oSOModuleDefaults.SrvdLinkedCatIDFilter = 0 Then Exit Sub

                Dim objSQL As New clsSqlConn
                strSQL = "SELECT StkItem.SubTypeID FROM StkItem INNER JOIN spil_InvSubType ON StkItem.SubTypeID=spil_InvSubType.SubTypeID where ItemActive = 1 and ubIIGLASSSERVICE = 1 and DefaultGlassService=0 AND StkItem.StockLink=" & UG.ActiveRow.Cells("StockLink").Value
                Dim SubTypeID As Integer = objSQL.Get_ScalerINTEGER(strSQL)
                If SubTypeID = oSOModuleDefaults.SrvdLinkedCatIDFilter Then
                    Dim band As UltraGridBand = DDServLinkCat.DisplayLayout.Bands(0)
                    band.ColumnFilters.ClearAllFilters()
                    band.ColumnFilters("SubTypeID").FilterConditions.Add(FilterComparisionOperator.Equals, oSOModuleDefaults.LinkedSrvCatIDFilter)
                    If Not IsNothing(UG.ActiveRow.Cells("LinkedSrvCatID").Value) Then
                        If IsDBNull(UG.ActiveRow.Cells("LinkedSrvCatID").Value) Then
                            DoNotActivat = True
                            Dim frmlinkedServices_Select As New frmlinkedServices_Select()
                            frmlinkedServices_Select.Location = New Point(DDServLinkCat.Location.X, DDServLinkCat.Location.Y + 75)
                            frmlinkedServices_Select.ShowDialog()
                            If frmlinkedServices_Select.DialogResult = Windows.Forms.DialogResult.OK Then
                                UG.ActiveRow.Cells("LinkedSrvCatID").Value = frmlinkedServices_Select.linkedServicesStockLink
                                GetBillOfMaterials()
                            End If
                        End If

                    End If
                Else
                    Dim band As UltraGridBand = DDServLinkCat.DisplayLayout.Bands(0)
                    band.ColumnFilters.ClearAllFilters()
                    band.ColumnFilters("SubTypeID").FilterConditions.Add(FilterComparisionOperator.Equals, -1)
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            DoNotActivat = False
        End Try




    End Sub

#End Region

End Class