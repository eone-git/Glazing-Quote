Imports Infragistics.Win.UltraWinGrid

Public Class clsGlazingQuoteExtension
    Dim moduleName As String = "Glazing Quote"
    Public reuturnedMessage As String = ""
    Dim glazQuoteDefaultDataSet As DataSet

    Public _ProjectId As Integer = 0
    Public _StageId As Integer = 0
    Public _JobId As Integer = 0

    Public IsFromJobProject As Boolean = False
    Public IsProgressCliam As Boolean = False
    Public _TotalInvoiced As Decimal = 0
    Public _isSaved As Boolean = False
    Public _odrIndex As Integer = 0

    Public Sub GQCreatefrmGlazingQuoteObject()
        Try
            glazingQuote = New frmGlazingQuote

        Catch ex As Exception
            GQShowMessage(ex.Message, moduleName, MsgBoxStyle.Critical)

        End Try

    End Sub

    Public Sub GQOpenGlazingQuote(Optional ByRef quoteOrdeIndex As Integer = 0, Optional ByRef isACopy As Boolean = False)
        Dim glazingQuote As New frmGlazingQuote

        Try
            glazingQuote.quoteOrdeIndex = quoteOrdeIndex
            glazingQuote.isACopy = isACopy
            'for Job costing
            glazingQuote._ProjectId = _ProjectId
            glazingQuote._StageId = _StageId
            glazingQuote._JobId = _JobId

            glazingQuote.IsFromJobProject = IsFromJobProject
            glazingQuote.IsProgressCliam = IsProgressCliam
            glazingQuote._TotalInvoiced = _TotalInvoiced
            glazingQuote._isSaved = _isSaved
            glazingQuote._odrIndex = _odrIndex

            glazingQuote.InitializeQuotation()
            glazingQuote.Show()
            glazingQuote.WindowState = FormWindowState.Maximized
            glazingQuote.BringToFront()
            glazingQuote.Refresh()
            'If isACopy = True Then
            '    GQDocumentLog(quoteOrdeIndex, QuoteStateValueCOPY)

            'End If
        Catch ex As Exception
            GQShowMessage(ex.Message, moduleName, MsgBoxStyle.Critical)

        End Try
    End Sub

    Public Sub GQOpenGlazingEstimate(Optional ByRef quoteOrdeIndex As Integer = 0, Optional ByRef isACopy As Boolean = False)
        Dim glazingQuote As New frmGlazingQuote

        Try
            glazingQuote.Text = "Glazing Estimate"
            glazingQuote.IsEstimate = True
            glazingQuote.quoteOrdeIndex = quoteOrdeIndex
            glazingQuote.isACopy = isACopy
            'for Job costing
            glazingQuote._ProjectId = _ProjectId
            glazingQuote._StageId = _StageId
            glazingQuote._JobId = _JobId

            glazingQuote.IsFromJobProject = IsFromJobProject
            glazingQuote.IsProgressCliam = IsProgressCliam
            glazingQuote._TotalInvoiced = _TotalInvoiced
            glazingQuote._isSaved = _isSaved
            glazingQuote._odrIndex = _odrIndex

            glazingQuote.InitializeQuotation()
            glazingQuote.Show()
            glazingQuote.WindowState = FormWindowState.Maximized
            glazingQuote.BringToFront()
            glazingQuote.Refresh()
            'If isACopy = True Then
            '    GQDocumentLog(quoteOrdeIndex, QuoteStateValueCOPY)

            'End If
        Catch ex As Exception
            GQShowMessage(ex.Message, moduleName, MsgBoxStyle.Critical)

        End Try
    End Sub

    Public Function GQEmailSentStateUpdate(ByVal quoteOrdeIndex As Integer, ByRef reqQuoteStateValue As Integer, Optional ByRef requestReason As Boolean = False) As Integer
        Dim collspPara As New Collection
        Dim colPara As New spParameters
        Dim newSQLQuery As String = ""
        Dim newSQL As New clsSqlConn
        Dim items As Array
        items = System.Enum.GetValues(GetType(QuoteStateValue))
        Dim messageLable As String = System.Enum.GetName(GetType(QuoteStateValue), reqQuoteStateValue)
        Try
            If IsNothing(messageLable) = False Then
                If messageLable.Length > 1 Then
                    If messageLable.Substring(messageLable.Length - 2) = "ed" Then
                        messageLable = messageLable.Substring(0, messageLable.Length - 2)
                    End If
                End If
            End If
            If requestReason = True Then
                Dim frmGlazingMessageBoxObj As New frmGlazingMessageBox(Me, moduleName, "Reason for " & messageLable.ToLower() & " this quotation")
                frmGlazingMessageBoxObj.ShowDialog()
            End If
            newSQL.Begin_Trans()

            colPara.ParaName = "@QuoteStateID"
            colPara.ParaValue = reqQuoteStateValue
            collspPara.Add(colPara)

            colPara.ParaName = "@OrderIndex"
            colPara.ParaValue = quoteOrdeIndex
            collspPara.Add(colPara)

            sqlQuary = "update spilInvNum set QuoteStateID = @QuoteStateID where OrderIndex = @OrderIndex"
            If newSQL.EXE_SQL_Trans_Para(sqlQuary, collspPara) = 0 Then
                MsgBox("Error in Updating state", MsgBoxStyle.Critical, "SPIL Glass")
                Exit Function
            End If

            If reqQuoteStateValue <> QuoteStateValue.SentAndConfirmationPending Then
                GQDocumentLog(quoteOrdeIndex, (System.Enum.GetName(GetType(QuoteStateValue), reqQuoteStateValue)), newSQL, reuturnedMessage)
            End If

            newSQL.Commit_Trans()
            Return 1

        Catch ex As Exception
            GQShowMessage(ex.Message, moduleName, MsgBoxStyle.Critical)
            Return 0

        Finally
            collspPara.Clear()

        End Try
    End Function

    Public Sub GQShowJobDescription(ByRef orderIndex As Integer)
        Dim glazingQuote As New frmGlazingQuote
        Dim objSQL As New clsSqlConn
        Dim DS_BATCHES As DataSet
        Try

            SQL = "SELECT * FROM GlzQuote_Job_Details WHERE OrderIndex='" & orderIndex & "'"
            With objSQL
                DS_BATCHES = .GET_INSERT_UPDATE(SQL)
                If DS_BATCHES.Tables(0).Rows.Count <> 0 Then
                    If IsNothing(DS_BATCHES.Tables(0).Rows(0).Item("GlzQuoteJobDes")) = False Then
                        If DS_BATCHES.Tables(0).Rows(0).Item("GlzQuoteJobDes") = "" Then
                            GQShowMessage("There is no desctioption for selected quotation", moduleName, MsgBoxStyle.Information)
                            Exit Sub

                        End If
                    Else
                        GQShowMessage("There is no mathing quotation", moduleName, MsgBoxStyle.Information)
                        Exit Sub

                    End If
                Else
                    GQShowMessage("There is no mathing quotation", moduleName, MsgBoxStyle.Information)
                    Exit Sub

                End If
                glazingQuote.OpenDescriptionState(False, DS_BATCHES.Tables(0).Rows(0).Item("GlzQuoteJobDes"))

            End With

        Catch ex As Exception
            GQShowMessage(ex.Message, moduleName, MsgBoxStyle.Critical)

        Finally
            newGlazingQuote = Nothing

        End Try
    End Sub

    Public Function GQDocumentLog(ByRef orderIndex As Integer, ByRef logAction As String, Optional sqlObj As clsSqlConn = Nothing, Optional description1 As String = "")

        Dim objDLItem As New clsDocumentLogEntry
        Dim isNewConection As Boolean = False
        Try
            If IsNothing(sqlObj) = True Then
                sqlObj = New clsSqlConn
                sqlObj.Begin_Trans()
                isNewConection = True
            End If

            objDLItem.iDocID = orderIndex
            objDLItem.iDocTypeID = GlassDocTypes.Quotation
            objDLItem.LogAction = logAction
            objDLItem.DocItemCount = 0
            objDLItem.DocServiceCount = 0
            objDLItem.LogDateTime = Now
            objDLItem.EnteredBy = strUserName
            objDLItem.Description1 = description1
            If objDLItem.AddDocLogWithTrans(sqlObj.Con, sqlObj.Trans) < 1 Then
                sqlObj.Rollback_Trans()
                Return 0
                Exit Function
            End If

            If isNewConection = True Then
                sqlObj.Commit_Trans()

            End If
            Return 1
        Catch ex As Exception
            GQShowMessage(ex.Message, moduleName, MsgBoxStyle.Critical)

        Finally
            isNewConection = False
        End Try
    End Function

    Function GetGlazQuoteDefaultData() As DataSet
        Dim sqlQuary As String
        Dim clsSqlConnObj As New clsSqlConn
        Try
            sqlQuary = "SELECT * FROM GlzQuote_State"
            glazQuoteDefaultDataSet = clsSqlConnObj.GET_INSERT_UPDATE(sqlQuary)
            Return glazQuoteDefaultDataSet

        Catch ex As Exception
            GQShowMessage(ex.Message, moduleName, MsgBoxStyle.Critical)
            Return glazQuoteDefaultDataSet

        Finally
            clsSqlConnObj.Dispose()
        End Try
    End Function

    Public Function GetDefaultByType(ByRef quoteStateType As Integer) As DataRow
        Dim glazQuoteDefaultData As DataSet = Nothing
        Dim glazQuoteDefaultDataRow As DataRow = Nothing
        Try
            glazQuoteDefaultData = GetGlazQuoteDefaultData()
            If glazQuoteDefaultData.Tables(0).Rows.Count Then
                For Each row As DataRow In glazQuoteDefaultData.Tables(0).Rows
                    If row.Item("JQSID") = quoteStateType Then
                        glazQuoteDefaultDataRow = row
                        Exit For
                    End If
                Next
            End If
            Return glazQuoteDefaultDataRow
        Catch ex As Exception
            Return glazQuoteDefaultDataRow
            GQShowMessage(ex.Message, moduleName, MsgBoxStyle.Critical)
        End Try
    End Function

    Public Function GetNextJobumentNumber(ByRef objSQL As clsInvHeader, ByRef documentType As String) As String

        Dim DS_ITEMS As DataSet
        Dim dr1 As DataRow
        Dim strSONo As String = Nothing
        Dim myDocType As Integer = mintDocType

        Try
            strSQL = "SELECT * FROM  spilDocDefault WHERE Line_ID = 1 "
            DS_ITEMS = objSQL.Get_Data_Trans(strSQL)
            If DS_ITEMS.Tables.Count = 0 Then
                MsgBox("Error in Spil DocsDf Table", MsgBoxStyle.Critical, "SPIL Glass")
                strSONo = Nothing
            Else
                For Each dr1 In DS_ITEMS.Tables(0).Rows
                    If documentType = "JobQuote" Then
                        strSONo = dr1("JobQuote_Prefix") & Format(dr1("Next_JobQuote_No"), "000")

                    ElseIf documentType = "GlazingQuote" Then
                        strSONo = dr1("GlazingQuote_Prefix") & Format(dr1("Next_GlazingQuote_No"), "000")

                    End If
                Next
            End If

            If strSONo <> Nothing Then
                If documentType = "GlazingQuote" Then
                    strSQL = "Update spilDocDefault SET Next_GlazingQuote_No = Next_GlazingQuote_No + 1 where Line_ID = 1"

                ElseIf documentType = "JobQuote" Then
                    strSQL = "Update spilDocDefault SET Next_JobQuote_No = Next_JobQuote_No + 1 where Line_ID = 1"

                End If

                If objSQL.Update_DataOnOpenConnection(strSQL) = 0 Then
                    MsgBox("Error in Updating spilDocDefault", MsgBoxStyle.Critical, "SPIL Glass")
                    strSONo = Nothing

                End If

            End If

            Return strSONo

        Catch ex As Exception
            modGlazingQuoteExtension.GQShowMessage(ex.Message, moduleName, MsgBoxStyle.Critical)
            Return strSONo

        End Try


    End Function


    Sub CoverToOldSalesOrder(ByRef jqOrderIndex As Integer, ByRef jqDocType As Integer)
        Try
            Dim mySalesDoc As New frmSO
            Dim gridCell As UltraGridCell
            mySalesDoc.pubMeIsNewRecord = False
            mySalesDoc.pubMeCalledBy = GlassDocCalledBy.EditDocListDoubleClicked
            mySalesDoc.pubMeOrderIndex = jqOrderIndex
            mySalesDoc.pubMeSpilDocTypeID = jqDocType
            mySalesDoc.WindowState = FormWindowState.Minimized
            mySalesDoc.LoadDocumentMasterData()
            mySalesDoc.isFromGlazingQuote = True
            mySalesDoc.OpenSO()
            mySalesDoc.isFromGlazingQuote = False
            mySalesDoc.SetDocumentControlProperties()
            If mySalesDoc.ConvertToSO() = False Then
                Exit Sub
            End If
            For Each dr As UltraGridRow In mySalesDoc.UG2.Rows.GetRowEnumerator(GridRowType.DataRow, Nothing, Nothing)
                gridCell = dr.Cells("LineType")
                If gridCell.Value = QuateFiedTypesList.Header_Main Or gridCell.Value = QuateFiedTypesList.Header_Sub Or gridCell.Value = QuateFiedTypesList.Subtotal Then
                    dr.Delete()
                End If
                If dr.Cells("Qty").Value = 0 Then
                    dr.Cells("Qty").Value = 1
                End If
                dr.Cells("LineType").Value = LineState.Normal
                dr.CellAppearance.ForeColor = Color.Black
            Next
            mySalesDoc.StartPosition = FormStartPosition.CenterScreen
            mySalesDoc.Show()
            mySalesDoc.Refresh()
            mySalesDoc.WindowState = FormWindowState.Maximized
            mySalesDoc.BringToFront()

        Catch ex As Exception
            modGlazingQuoteExtension.GQShowMessage(ex.Message, moduleName, MsgBoxStyle.Critical, "warning", "CoverToOldSalesOrder")

        End Try
    End Sub

    Function SOHandler(ByRef objInvDetailline As clsInvDetailLine) As Boolean
        Try
            Dim lineType As Integer = objInvDetailline.LineType()
            If lineType = QuateFiedTypesList.Header_Main Or lineType = QuateFiedTypesList.Header_Sub Or lineType = QuateFiedTypesList.Subtotal Then
                Return True
            Else
                Return False
            End If
            Return True
        Catch ex As Exception
            Return False

        End Try
    End Function

    Sub AddServiceDataToSalesOrder(ByRef glassRow As UltraGridRow, ByRef frmsoObj As frmSO)
        Try
            Dim newDataSet As DataSet
            Dim row As UltraGridRow
            Dim hasPreVal As Boolean = False

            If IsNothing(glassRow.Cells("templateData").Value) = False Then
                If selectedItems <> "" Then
                    hasPreVal = True
                End If
            End If

            If hasPreVal = True Then
                Dim itemLines() As String = selectedItems.Split(";")
                For Each itemDetails As String In itemLines
                    clsInvDetLine = New clsInvDetailLine
                    If itemDetails = "" Then
                        Exit For
                    End If
                    Dim itemDetailsCol() As String = itemDetails.Split(",")

                    AddServiceLines(glassRow, itemDetailsCol, frmsoObj)

                Next
            End If
        Catch ex As Exception
            modGlazingQuoteExtension.GQShowMessage(ex.Message, moduleName, MsgBoxStyle.Critical, "warning", "AddDataToSalesOrder")
        End Try
    End Sub

    Sub AddServiceLines(ByRef drGlass As UltraGridRow, itemDetailsCol() As String, ByRef frmsoObj As frmSO)
        Try
            Dim dr As UltraGridRow = frmsoObj.UGServices_List.DisplayLayout.Bands(0).AddNew

            'This is only the line NO of service grid
            dr.Cells("LN").Value = dr.Index + 1
            dr.Cells("Band_Index").Value = drGlass.Cells("BandIndex").Value

            If drGlass.Cells("BandIndex").Value = 0 Then
                dr.Cells("Main_No").Value = 0
            ElseIf drGlass.Cells("BandIndex").Value = 1 Then
                dr.Cells("Main_No").Value = drGlass.Cells("MainLNNo").Value
            End If

            dr.Cells("StockLink").Value = itemDetailsCol(0)
            dr.Cells("LineNo").Value = dr.Index

            dr.Cells("StockLink1").Value = drGlass.Cells("StockLink").Value
            dr.Cells("ItemNo1").Value = drGlass.Cells("LN").Value      ' line no of the main Grid

            dr.Cells("iInvServiceID").Value = 0
            dr.Cells("iInvDetailID").Value = drGlass.Cells("iInvDetailID").Value

            dr.Cells("UniqueLN").Value = drGlass.Cells("UniqueLN").Value



            dr.Cells("Code").Value = itemDetailsCol(1)
            'dr.Cells("Description_1").Value = itemDetailsCol(0)
            dr.Cells("Comment").Value = itemDetailsCol(2)
            'dr.Cells("Pricing").Value = ""
            'dr.Cells("Pricing_ID").Value = ""
            'dr.Cells("Edging").Value = ""
            'dr.Cells("Edging_ID").Value = ""
            'dr.Cells("Peices").Value = ""
            dr.Cells("Volume").Value = drGlass.Cells("Volume").Value

            '1-Area   2-Flat,   3-Lineal, 4-None

            'If drSrv.Cells("Pricing_ID").Value = 1 Or drSrv.Cells("Pricing_ID").Value = 2 Or drSrv.Cells("Pricing_ID").Value = 3 Then
            '    dr.Cells("Units").Value = drSrv.Cells("Units").Value * drSrv.Cells("Peices").Value
            '    dr.Cells("SQFeetForPricing").Value = drSrv.Cells("SQFeetForPricing").Value * drSrv.Cells("Peices").Value
            'Else
            '    dr.Cells("Units").Value = drSrv.Cells("Units").Value
            '    dr.Cells("SQFeetForPricing").Value = drSrv.Cells("SQFeetForPricing").Value
            'End If

            dr.Cells("UnitQty").Value = itemDetailsCol(3)

            'dr.Cells("Price_Excl").Value = drSrv.Cells("Price_Excl").Value
            'dr.Cells("SurChrg").Value = drSrv.Cells("SurChrg").Value

            dr.Cells("TaxType").Value = drGlass.Cells("TaxType").Value
            dr.Cells("TaxRate").Value = drGlass.Cells("TaxRate").Value

            'dr.Cells("LineNet_Excl").Value = drSrv.Cells("LineNet_Excl").Value
            'dr.Cells("LineTax").Value = drSrv.Cells("LineTax").Value
            'dr.Cells("LineTotal_Excl").Value = drSrv.Cells("LineTotal_Excl").Value

            'If IsToAddSingleService Then
            '    dr.Cells("Line_Dis_Excl").Value = 1 'drSrv.Cells("Line_Dis_Excl").Value
            'Else
            '    dr.Cells("Line_Dis_Excl").Value = 0 'drSrv.Cells("Line_Dis_Excl").Value
            'End If

            'dr.Cells("QTY_LONG").Value = drSrv.Cells("QTY_LONG").Value
            'dr.Cells("QTY_SHORT").Value = drSrv.Cells("QTY_SHORT").Value

            'dr.Cells("Original_Price").Value = drSrv.Cells("Original_Price").Value

            'dr.Cells("Disc_Amt").Value = drSrv.Cells("Disc_Amt").Value
            'dr.Cells("Disc_Percentage").Value = drSrv.Cells("Disc_Percentage").Value

            'dr.Cells("Std_Cost").Value = drSrv.Cells("Std_Cost").Value

            'dr.Cells("Comment1").Value = IIf(drSrv.Cells("Pricing_ID").Value = 2, drSrv.Cells("Units").Value, "")
            'dr.Cells("Comment2").Value = IIf(drSrv.Cells("Pricing_ID").Value = 3, drSrv.Cells("Edging").Text, "")

            'dr.Cells("PriceCat").Value = drSrv.Cells("PriceCat").Value
            'dr.Cells("OrgPrice").Value = drSrv.Cells("OrgPrice").Value

            'dr.Cells("LineNotes").Value = drSrv.Cells("LineNotes").Value
            'dr.Cells("IsExternalItem").Value = drSrv.Cells("IsExternalItem").Value
            'dr.Cells("LinkedSrvCatID").Value = drSrv.Cells("LinkedSrvCatID").Value

            ' ''job
            'dr.Cells("Cost").Value = drSrv.Cells("Cost").Value
            'dr.Cells("Profit").Value = drSrv.Cells("Profit").Value

            dr.Cells("TaxExempt").Value = itemDetailsCol(4)
        Catch ex As Exception

        End Try
    End Sub

#Region "Units Calculations"
    Public Function CalculateArea(ByRef isInches As Boolean, ByRef itemHight As Decimal, ByRef itemWidth As Decimal, Optional ByRef needToBeRounded As Boolean = False) As String

        Dim itemHightRounded As Decimal = 0
        Dim itemWidthRounded As Decimal = 0

        Dim volume As Decimal = 0
        Dim SQFeetForPricing As Decimal = 0

        Try
            If isInches = True Then
                If needToBeRounded = True Then
                    'itemHightRounded = clsInchCalculationsObj.RoundToNearestEvenNumber(itemHight)
                    'itemWidthRounded = clsInchCalculationsObj.RoundToNearestEvenNumber(itemWidth)

                    'SQFeetForPricing = Math.Round(((itemHightRounded * itemWidthRounded) / 144), 6)
                    'volume = Math.Round(((itemHight * itemWidth) / 144), 6)
                End If
            Else
                volume = (itemWidth * itemHight) / 1000000
            End If

            Return If(IsNumeric(volume) = True, volume, 0) & ";" & If(IsNumeric(SQFeetForPricing) = True, SQFeetForPricing, 0)

        Catch ex As Exception
            modGlazingQuoteExtension.GQShowMessage(ex.Message, moduleName, MsgBoxStyle.Critical, "warning", "SetItemData")

        End Try
    End Function

    Public Function CalculateUnits(ByRef pricingType As Integer, ByRef isInches As Boolean, ByRef itemHight As Decimal, ByRef itemWidth As Decimal, Optional ByRef serviceQty As Integer = 0, Optional ByRef needToBeRounded As Boolean = False) As Double
        Dim Units As Double
        Try
            If pricingType = GlassServiceUnits.Area Then
                Units = (CalculateArea(isInches, itemHight, itemWidth, needToBeRounded).Split(";"))(1)

            ElseIf pricingType = GlassServiceUnits.Lineal Then
                If geographicRegionID = GeographicRegion.NorthAmerica Then
                    Units = (itemWidth + itemHight) * 2
                Else
                    Units = ((itemWidth + itemHight) * 2) / 1000

                End If
            ElseIf pricingType = GlassServiceUnits.Flat Then
                Units = serviceQty

            ElseIf pricingType = GlassServiceUnits.None Then
                Units = -1

            End If
            Return Units
        Catch ex As Exception
            modGlazingQuoteExtension.GQShowMessage(ex.Message, moduleName, MsgBoxStyle.Critical, "warning", "CalculateUnits")

        End Try
    End Function

    Public Function CalculateServiceTotalPrice(ByRef servicesData As String, ByRef glassData As String, ByRef isInches As Boolean, Optional ByRef needToBeRounded As Boolean = False) As String

        Dim pricingType As Integer
        Dim units As Double
        Dim totalUnits As Double
        Dim totalServicePriceExc As Double
        Dim totalServicePriceInc As Double
        Dim totalServiceTaxAmount As Double
        Dim itemPrice As Double
        Dim glassVolume As Double
        Dim glassPriceExc As Double
        Dim glassNewPriceExc As Double
        Dim glassQty As Double
        Dim serviceQty As Double
        Dim itemHight As Double
        Dim itemWidth As Double
        Dim result As String
        Dim taxRate As Double
        Dim clsGlazingDocStockItemHelperObj As New clsGlazingDocStockItemHelper

        Try
            Dim hasPreVal As Boolean = False
            If IsNothing(servicesData) = False Then
                If servicesData <> "" Then
                    hasPreVal = True
                End If
            End If

            '0: "Hight",  1: "Width", 2: "Qty", 3: "Price" - Optional
            Dim mainGlassData() As String = glassData.Split(",")
            Dim itemLines() As String = servicesData.Split(";")
            For Each itemDetails As String In itemLines
                If itemDetails = "" Then
                    Exit For
                End If

                Dim itemDetailsCol() As String = itemDetails.Split(",")
                pricingType = itemDetailsCol(7)
                itemPrice = itemDetailsCol(8)
                glassQty = mainGlassData(2)
                'glassPriceExc = mainGlassData(3)
                itemHight = mainGlassData(0)
                itemWidth = mainGlassData(1)

                units = CalculateUnits(pricingType, isInches, itemHight, itemWidth, serviceQty, needToBeRounded)

                totalServicePriceExc = (If(units = -1, 1, units * glassQty)) * itemPrice
                totalServiceTaxAmount = clsGlazingDocStockItemHelperObj.CalculateTaxAmount(itemDetailsCol(9), totalServicePriceExc, itemDetailsCol(4))
                totalServicePriceInc = totalServicePriceExc + totalServiceTaxAmount 'itemPrice
                'glassNewPriceExc = (glassPriceExc * glassQty) + totalServicePriceExc
            Next
            result = If(IsNumeric(totalServicePriceExc) = True, totalServicePriceExc, 0) & ";" & If(IsNumeric(totalServiceTaxAmount) = True, totalServiceTaxAmount, 0) & ";" & If(IsNumeric(totalServicePriceInc) = True, totalServicePriceInc, 0)

            Return result
        Catch ex As Exception

        End Try
    End Function

#End Region

#Region "Database Handler"

    Public Function DatabaseHandler(ByRef newSQLQuery As String, Optional ByRef clsSqlConnObj As clsSqlConn = Nothing, Optional ByRef collspPara As Collection = Nothing) As DataSet
        Try
            Dim isCompleted As Boolean
            Dim datasetTOSend As DataSet
            Dim withTransaction As Boolean

            If IsNothing(clsSqlConnObj) = True Then
                clsSqlConnObj = New clsSqlConn
                withTransaction = False

            Else
                If newSQLQuery.Contains("= @") Or newSQLQuery.Contains("=@") Or newSQLQuery.Contains(",@") Or newSQLQuery.Contains(", @") Then
                    withTransaction = True
                Else
                    withTransaction = False

                End If
            End If

            If withTransaction = True Then
                If clsSqlConnObj.EXE_SQL_Trans_Para_Return(newSQLQuery, collspPara) = 0 Then
                    modGlazingQuoteExtension.GQShowMessage("Error in Database Handler", moduleName, MsgBoxStyle.Critical)
                    clsSqlConnObj.Rollback_Trans()
                    datasetTOSend = Nothing
                Else
                    datasetTOSend = New DataSet
                End If
            Else
                datasetTOSend = clsSqlConnObj.GET_DATA_SQL(newSQLQuery)
            End If

            Return datasetTOSend
        Catch ex As Exception

        End Try
    End Function

#End Region


End Class
