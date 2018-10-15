Imports Infragistics.Win.UltraWinGrid
Imports System.Data.SqlClient
Imports Infragistics.Win
Imports SPIL.Shapes
Imports System.Drawing.Imaging
Imports System.Runtime.Serialization.Formatters.Binary
Imports System.IO
Imports GlassInventoryModule
Imports GMap.NET.Core
Imports GMap.NET.WindowsForms

Imports GMap.NET.MapProviders
Imports GMap.NET
Imports GMap.NET.WindowsForms.Markers
Imports System.Text.RegularExpressions
Imports System.Xml

Public Class frmGlazingQuote
    Dim oPriceUnits As New clsSOPricingAndUnits
    Dim oSOModuleDefaults As New clsSOModuleDefaults
    Dim oFormCustomer As clsCustomer
    Public pubMeQuoteDocument As clsInvHeader
    Public pubMeIsNewRecord As Boolean
    Public pubMeCalledBy As Integer
    Public pubMeOrderIndex As Integer
    Public pubMeSpilDocTypeID As Integer
    Public pubMeSpilDocStatusID As Integer
    Public pubMemyLineState As Integer
    Public pubMeProductType As Integer
    Public pubMeToughened As Boolean
    Public pubmyProductionState As Integer
    Dim con As New SqlConnection(strCon)
    Dim Trans As SqlTransaction
    Dim oProdDefaults As New clsProdDefaults("")
    Dim isNew As Boolean = True
    'Public Shared UG2ActiveRow As UltraGridRow
    Dim UG2ActiveRow As UltraGridRow
    'Public  UG2ActiveRow As UltraGridRow
    Public Shared selectedDocDes As String
    Public fotterIsActive As Boolean = False
    Public isJobDescriptionActive As Boolean = False

    Public subHeaderID As Integer = 0
    Dim groupStarted As Boolean = False
    Public Shared priceType As String = ""

    Dim isDefaultChanged As Boolean = False

    Dim isExistingOrder As Boolean = False
    Dim canUpdate As Boolean = True
    Public openedModeulename As String = ""

    'Taxed or not
    Public isTaxedPrice As Boolean = False
    Public defaultTaxtRate As Integer = 0
    Public defaultTaxtRateValue As Decimal = 0
    Public sAllowBlankCustomerOrderNo As Boolean = False


    Dim isCopied As Boolean = False
    Dim isPasting As Boolean = False
    Dim preQuoteNumber As String = ""
    Public quoteOrdeIndex As Integer = 0
    Dim quoteOrderNum As String = ""
    Dim quoteDocRepID As Integer = 0

    Dim collDeletedItemLines As New Collection
    Dim rowCopied As New List(Of UltraGridRow)


    Public isOpeningQuote As Boolean = False
    Dim calculateWhilePasting As Boolean = False
    Dim IsCellClearing As Boolean = False
    Public isStockItemActive As Boolean = False
    Dim objClsInvHeader As New clsInvHeader

    Public jobDescription As String = ""

    'map
    Private latitude As Double = 0.0
    Private longitude As Double = 0.0
    Dim emptyArray As Byte() = New Byte(63) {}

    Dim subTotalByGroup As List(Of Double)
    Public isStockItemClosing As Boolean = False

    'Public Modeulename As String = ""

    'starting tweeks
    Dim lineTypeCell As UltraGridCell
    Dim gridActiveRow As UltraGridRow
    Dim gridActiveCell As UltraGridCell
    Dim downKeyPressed As Boolean = False

    Public quoteState As Integer = -1
    Public isACopy As Boolean = False
    Public isCancelled As Boolean = False
    Dim preQuoteState As String = -1
    Dim clsGlazingQuoteExtensionObj As New clsGlazingQuoteExtension()

    Dim clsGQExtensionForJobCostingObj As New clsGQExtensionForJobCosting(Me)
    Public _ProjectId As Integer = 0
    Public _StageId As Integer = 0
    Public _JobId As Integer = 0

    Public IsFromJobProject As Boolean = False
    Public IsProgressCliam As Boolean = False
    Public _TotalInvoiced As Decimal = 0
    Public _isSaved As Boolean = False
    Public _odrIndex As Integer = 0
    Public IsEstimate As Boolean = False

    Dim isSaved As Boolean = False
    Dim isClosing As Boolean = False
    Dim getTaxRateFromCustomer As Boolean = False
    Dim canEditeAmount As Boolean = False
    Dim expireDateCount As Integer = 2
    Dim defaultTaxRateForQuote As Integer = 0
    Dim defaultTaxRateValueForQuote As Double = 0.0
    Dim saveButtonPreState As Boolean
    Dim isSaveable As Boolean

    Public isInches As Boolean = True
    Public needToBeRounded As Boolean = True
    Public isForUSA As Boolean = True
    Dim defaultItemForText As Integer
    Public dCusTaxCode As String = ""
    Public dCusTaxRate As Double

    Dim oFacDefaults As New clsFacilityDefaults(False, UserBranchID)

    Private Sub GET_CUSTOMERS(ByVal value As String)

        Dim IsProspect As Boolean = False
        If value = "Prospect" Then
            IsProspect = True
        Else
            IsProspect = False
        End If

        SQL = "SELECT Name, Account, DCLink, (Physical1 + ' ' + Physical2 + ' ' + Physical3 + ' ' + Physical4 + ' ' + Physical5 + ' ' + Physical5) as DeliveryAddress " & _
            "FROM  Client  LEFT OUTER JOIN " & _
            " spilFuelLevRates ON Client.FLID = spilFuelLevRates.FLID WHERE Client.IsProspect = '" & IsProspect & "'"

        Dim objSQL As New clsSqlConn
        With objSQL
            Try
                DS_BATCHES = .GET_INSERT_UPDATE(SQL)
                cmbAccount.DataSource = DS_BATCHES.Tables(0)
                cmbAccount.DisplayMember = "Name"
                cmbAccount.ValueMember = "DCLink"
                cmbAccount.DisplayLayout.Bands(0).Columns("Name").Width = 300
                cmbAccount.DisplayLayout.Bands(0).Columns("DeliveryAddress").Width = 500
                cmbAccount.DisplayLayout.Bands(0).Columns("Account").Hidden = True
                cmbAccount.DisplayLayout.Bands(0).Columns("DCLink").Hidden = True

            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.Exclamation, "SPIL Glass")
            Finally
                .Dispose()
                objSQL = Nothing
            End Try
        End With

    End Sub

    Private Sub InitializesCustomerDetails()
        Try
            Dim mUgR As UltraGridRow
            mUgR = cmbAccount.ActiveRow
            If Not mUgR Is Nothing Then
                If Set_Customer(mUgR) = 0 Then
                    Exit Sub
                End If
            Else
                txtPhy1.Text = ""
                txtPhy2.Text = ""
                txtPhy3.Text = ""
                txtPhy4.Text = ""
                txtPhy5.Text = ""
                txtPhyPostCode.Text = ""
                cboArea.Text = ""
                txtPost1.Text = ""
                txtPost2.Text = ""
                txtPost3.Text = ""
                txtPost4.Text = ""
                txtPost5.Text = ""
                txtPostCode.Text = ""
                txtContact1.Text = ""
                txtContact2.Text = ""
                txtTele1.Text = ""
                txtTele2.Text = ""
                txtMobile.Text = ""
                cmbContPerson.Text = ""
                txtContPerTel.Text = ""
                txtContEmail.Text = ""
                'cmbAccount.Text = ""
                cmbAccount.Focus()
            End If
        Catch ex As Exception
            Exit Sub
        End Try

    End Sub

    Private Function Set_Customer(ByVal ugR As UltraGridRow) As Integer
        Try
            '0=Customer   1=Prospect
            If cmbCusType.Value = 0 Or cmbCusType.Value = 1 Then

                'Initialize Customer and Price Objects
                oFormCustomer = New clsCustomer(ugR.Cells("DCLink").Value)
                oPriceUnits.oCustomer = oFormCustomer
                oPriceUnits.iDefaultStockPriceListID = oSOModuleDefaults.DefaultTradePriceListID ' iFormDefaultTradePriceListID
                iFormDefaultTradePriceListID = oSOModuleDefaults.DefaultTradePriceListID
                'End of Initialize Customer and Price Objects

                'Fill currency and Exchange Rate
                txtCurrency.Text = oFormCustomer.CurrencyCode
                txtExRate.Text = oFormCustomer.ExRate

                'Fill Tax code and Tax Rate
                dCusTaxCode = oFormCustomer.TaxCode
                dCusTaxRate = oFormCustomer.TaxRate

                'Q1.Text = "(" & oFormCustomer.CurrencyCode & ")"
                'Q2.Text = "(" & oFormCustomer.CurrencyCode & ")"
                'Q3.Text = "(" & oFormCustomer.CurrencyCode & ")"

                T1.Text = "(" & oFormCustomer.CurrencyCode & ")"
                T2.Text = "(" & oFormCustomer.CurrencyCode & ")"
                T3.Text = "(" & oFormCustomer.CurrencyCode & ")"

                UG2.Enabled = True

                If IsDBNull(oFormCustomer.iClassID) = True Then
                    MsgBox("Please assign valid customer group for selected customer", MsgBoxStyle.Exclamation, "Validation")
                    Return 0
                    Exit Function
                End If

                '>>intCustGroup = ugR.Cells("iClassID").Value
                '>>intPriceListID = 0
                '>>intDCLink = ugR.Cells("DCLink").Value
                Call SetContact()

                'cmbCustGroup.Value = oFormCustomer.iClassID

                '>>If ugR.Cells("CashDebtor").Value = True Or ugR.Cells("bCODAccount").Value = True Then
                If oFormCustomer.CashDebtor = True Or oFormCustomer.bCODAccount = True Then
                    IsCustCODorIWG = True
                Else
                    IsCustCODorIWG = False
                End If

                Call Set_Cust_Address()

                Call GET_Projects(ugR.Cells("DCLink").Value)

                ''To fix customer edit issue (StakeGlass)
                cmbSalesRep.Value = oFormCustomer.RepID '>> ugR.Cells("RepID").Value

                If pubMeIsNewRecord Or pubmyProductionState = GlassProdState.Ready2Batch Then
                    cmbDelMethod.Value = oFormCustomer.uiARDeliveryMethod '>> ugR.Cells("uiARDeliveryMethod").Value
                End If

                cboArea.Value = oFormCustomer.iAreasID '>>ugR.Cells("iAreasID").Value
                txtPostCode.Text = oFormCustomer.PostPC ' IIf(IsDBNull(ugR.Cells("PostPC").Value), "", ugR.Cells("PostPC").Value)

                If pubMeIsNewRecord = False Then
                    Return 1
                    Exit Function
                End If

                If pubMeCalledBy = GlassDocCalledBy.NewNCRFromSO Or pubMeCalledBy = GlassDocCalledBy.NewSOFromQuote Then
                    Exit Function
                End If

                ''To fix customer edit issue (StakeGlass)
                ''cmbSalesRep.Value = oFormCustomer.RepID '>> ugR.Cells("RepID").Value
                ''cmbDelMethod.Value = oFormCustomer.uiARDeliveryMethod '>> ugR.Cells("uiARDeliveryMethod").Value
                ''cboArea.Value = oFormCustomer.iAreasID '>>ugR.Cells("iAreasID").Value
                ''txtPostCode.Text = oFormCustomer.PostPC ' IIf(IsDBNull(ugR.Cells("PostPC").Value), "", ugR.Cells("PostPC").Value)


                'If pubMeSpilDocTypeID = GlassDocTypes.CreditNote Then
                '    txtFuelLevPercen.Value = 0
                'Else
                '    'txtFuelLevPercen.Value = IIf(IsDBNull(cmbAccount.ActiveRow.Cells("FuelLevyPercen").Value) = True, 0, cmbAccount.ActiveRow.Cells("FuelLevyPercen").Value)
                '    txtFuelLevPercen.Value = 0
                'End If

                If pubMeSpilDocStatusID = GlassDocState.Archived Then
                    ' GetInvoiceDueDate()
                    Dim DueDateWithAccountTerms() As String
                    Dim objCustomer As New clsCustomer

                    DueDateWithAccountTerms = Split(objCustomer.GetInvoiceDueDate(ugR.Cells("DCLink").Value, txtDueDate.Value), ";")
                    txtDueDate.Value = Convert.ToDateTime(DueDateWithAccountTerms(0))
                End If

                If cmbDelMethod.Value = 1 Then
                    cmbDelivery.Value = oFormCustomer.FLID 'ugR.Cells("FLID").Value
                    If IsNumeric(cmbDelivery.Text) = True Then
                        If Convert.ToDouble(cmbDelivery.Text) > 0 Then
                            'chkFuelLev.Enabled = True
                            'chkFuelLev.Checked = False
                            'chkFuelLev.Text = "Delivery Charge : " & String.Format("{0:0.00}", cmbDelivery.Text) & "$ to be added "
                            'txtFuelLevPercen.Value = Convert.ToDouble(cmbDelivery.Text)
                        Else
                            'chkFuelLev.Enabled = False
                            'chkFuelLev.Checked = False
                            'chkFuelLev.Text = "No Delivery Charge to be added "
                            'txtFuelLevPercen.Value = 0
                        End If
                    End If
                Else
                    'cmbDelivery.Value = 1
                    'chkFuelLev.Enabled = False
                    'chkFuelLev.Checked = False
                    'chkFuelLev.Text = "No Delivery Charge to be added "
                End If



                Dim i As Integer = 0
                For Each mugR As UltraGridRow In UG2.Rows
                    If Not IsDBNull(mugR.Cells("StockLink").Value) Then
                        If mugR.Cells("StockLink").Value > 0 Then
                            i += 1
                        Else
                            mugR.Delete(False)
                        End If
                    End If
                Next

                If i > 0 Then
                    If MessageBox.Show("Do you want to clear all entered lines", "Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = System.Windows.Forms.DialogResult.Yes Then
                        For Each mugR As UltraGridRow In UG2.Rows.All
                            mugR.Delete(False)
                        Next

                        Dim dr As UltraGridRow = UG2.DisplayLayout.Bands(0).AddNew
                        If pubMeSpilDocTypeID = GlassDocTypes.CreditNote Then
                            dr.Cells("ItemType").Value = GlassItemTypes.Service
                        Else
                            dr.Cells("ItemType").Value = 1
                        End If
                        dr.Cells("Method").Value = 1
                        dr.Cells("Measure").Value = 1
                        UG2.ActiveRow.Cells(1).Selected = True

                        If cmbFacility.Text = "" Then
                            dr.Cells("FacilityID").Value = 0
                        Else
                            dr.Cells("FacilityID").Value = IIf(IsDBNull(cmbFacility.Value), 0, cmbFacility.Value)
                        End If


                        UG2.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit

                    End If
                End If



            End If
            Return 1
        Catch ex As Exception
            Return 0
        End Try

    End Function

    Public Sub SetContact()
        Try
            txtContact1.Text = oFormCustomer.Contact_Person ' Dr("Contact_Person")
            txtContact2.Text = oFormCustomer.Delivered_To 'Dr("Delivered_To")
            txtTele1.Text = oFormCustomer.Telephone 'Dr("Telephone")
            txtTele2.Text = oFormCustomer.Telephone2 'Dr("Telephone2")
            txtMobile.Text = oFormCustomer.Fax1 '  Dr("Fax1")
            cmbContPerson.DataSource = oFormCustomer.CustomerContactsDS
            cmbContPerson.DisplayLayout.Bands(0).ColHeadersVisible = False
            cmbContPerson.DisplayMember = "ContName"
            cmbContPerson.Text = oFormCustomer.Contact_Person
            txtContPerTel.Text = oFormCustomer.Telephone
            txtContEmail.Text = oFormCustomer.EMail ' Dr("EMail")
        Catch ex As Exception
            modGlazingQuoteExtension.GQShowMessage(ex.Message, Me.Text, MsgBoxStyle.Critical)

        End Try
    End Sub

    Private Sub Set_Cust_Address()
        Try
            ugDelAddress.DataSource = Nothing
            ugDelAddress.DataSource = oFormCustomer.DeliveryAddressesDS 'DS_BATCHES2.Tables(2)
            ugDelAddress.DisplayLayout.Bands(0).Columns("cDescription").Header.Caption = "Description"
            ugDelAddress.DisplayLayout.Bands(0).Columns("cDelAddress1").Header.Caption = "DelAddress1"
            ugDelAddress.DisplayLayout.Bands(0).Columns("cDelAddress2").Header.Caption = "DelAddress2"
            ugDelAddress.DisplayLayout.Bands(0).Columns("cDelAddress3").Header.Caption = "DelAddress3"
            ugDelAddress.DisplayLayout.Bands(0).Columns("cDelAddress4").Header.Caption = "DelAddress4"
            ugDelAddress.DisplayLayout.Bands(0).Columns("cDelAddress5").Header.Caption = "DelAddress5"
            ugDelAddress.DisplayLayout.Bands(0).Columns("cDelAddressPC").Header.Caption = "Post Code"

            txtPost1.ResetText()
            txtPost2.ResetText()
            txtPost3.ResetText()
            txtPost4.ResetText()
            txtPost5.ResetText()
            txtPostCode.ResetText()

            '>>For Each dr4 In DS_BATCHES2.Tables(0).Rows
            txtPost1.Text = oFormCustomer.Post1 ' IIf(IsDBNull(dr4("Post1")) = True, "", dr4("Post1"))
            txtPost2.Text = oFormCustomer.Post2 ' IIf(IsDBNull(dr4("Post2")) = True, "", dr4("Post2"))
            txtPost3.Text = oFormCustomer.Post3 ' IIf(IsDBNull(dr4("Post3")) = True, "", dr4("Post3"))
            txtPost4.Text = oFormCustomer.Post4 ' IIf(IsDBNull(dr4("Post4")) = True, "", dr4("Post4"))
            txtPost5.Text = oFormCustomer.Post5 ' IIf(IsDBNull(dr4("Post5")) = True, "", dr4("Post5"))
            txtPostCode.Text = oFormCustomer.PostPC ' IIf(IsDBNull(dr4("PostPC")) = True, "", dr4("PostPC"))
            '>>Next

            txtPhy1.ResetText()
            txtPhy2.ResetText()
            txtPhy3.ResetText()
            txtPhy4.ResetText()
            txtPhy5.ResetText()
            txtPhyPostCode.ResetText()

            '>>For Each dr4 In DS_BATCHES2.Tables(1).Rows
            txtPhy1.Text = oFormCustomer.Physical1 ' IIf(IsDBNull(dr4("Physical1")) = True, "", dr4("Physical1"))
            txtPhy2.Text = oFormCustomer.Physical2 ' IIf(IsDBNull(dr4("Physical2")) = True, "", dr4("Physical2"))
            txtPhy3.Text = oFormCustomer.Physical3 ' IIf(IsDBNull(dr4("Physical3")) = True, "", dr4("Physical3"))
            txtPhy4.Text = oFormCustomer.Physical4 ' IIf(IsDBNull(dr4("Physical4")) = True, "", dr4("Physical4"))
            txtPhy5.Text = oFormCustomer.Physical5 ' IIf(IsDBNull(dr4("Physical5")) = True, "", dr4("Physical5"))
            txtPhyPostCode.Text = oFormCustomer.PhysicalPC ' IIf(IsDBNull(dr4("PhysicalPC")) = True, "", dr4("PhysicalPC"))
            '>>Next

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Exclamation, "SPIL Glass")

        End Try
    End Sub

    Private Sub GET_Projects(ByVal vAccountID As Integer)
        Dim DS_Project As DataSet

        If pubMeIsNewRecord Then
            SQL = "SELECT ProjID, ProjectName, ProjectNumber  " & _
                  "FROM spilProjectMaster where Active=1 and AccountID=" & vAccountID & " " & _
                  "order by ProjectName"
        Else 'this is to fill even inactive projects for retreving records at later stage
            SQL = "SELECT ProjID, ProjectName, ProjectNumber  " & _
                  "FROM spilProjectMaster where AccountID=" & vAccountID & " " & _
                  "order by ProjectName"
        End If

        Dim objSQL As New clsSqlConn
        With objSQL
            Try
                DS_Project = .GET_INSERT_UPDATE(SQL)
                cmbCustProject.DataSource = Nothing
                cmbCustProject.DataSource = DS_Project.Tables(0)
                cmbCustProject.DisplayMember = "ProjectName"
                cmbCustProject.ValueMember = "ProjID"
                cmbCustProject.DisplayLayout.Bands(0).Columns(1).Width = cmbCustProject.Width - 100
                cmbCustProject.DisplayLayout.Bands(0).Columns(2).Width = 100
                cmbCustProject.DisplayLayout.Bands(0).Columns(0).Hidden = True
                If DS_Project.Tables(0).Rows.Count > 0 Then
                    cmbCustProject.Enabled = True
                Else
                    cmbCustProject.Enabled = False
                End If

            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.Exclamation, "SPIL Glass")
            Finally
                .Dispose()
                objSQL = Nothing
            End Try
        End With
    End Sub

    '-----------------------------
    Private Sub GET_Del_Method()

        SQL = "SELECT     Counter, Method  FROM DelTbl ORDER BY Counter"   'Delivery methods

        Dim objSQL As New clsSqlConn

        With objSQL
            Try

                DS_BATCHES = .GET_INSERT_UPDATE(SQL)
                'cmbDelMethod.DataSource = DS_BATCHES.Tables(0)
                'cmbDelMethod.DisplayMember = "Method"
                'cmbDelMethod.ValueMember = "Counter"
                'cmbDelMethod.DisplayLayout.Bands(0).Columns(0).Hidden = True
                ''cmbDelMethod.DisplayLayout.Bands(0).Columns(0).Width = 139
                'cmbDelMethod.DisplayLayout.Bands(0).Columns(1).Width = cmbDelMethod.Width '200

            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.Exclamation, "SPIL Glass")

            Finally

                .Dispose()
                objSQL = Nothing

            End Try

        End With

    End Sub

    Private Sub GET_Couriers()
        SQL = "SELECT     CourierID, CourierName  FROM spilCourierMaster ORDER BY CourierName"
        Dim objSQL As New clsSqlConn
        With objSQL
            Try
                DS_BATCHES = .GET_INSERT_UPDATE(SQL)
                'cmbCourier.DataSource = DS_BATCHES.Tables(0)
                'cmbCourier.DisplayMember = "CourierName"
                'cmbCourier.ValueMember = "CourierID"
                'cmbCourier.DisplayLayout.Bands(0).Columns(1).Width = 200
                'cmbCourier.DisplayLayout.Bands(0).Columns(0).Hidden = True

            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.Exclamation, "SPIL Glass")

            Finally

                .Dispose()
                objSQL = Nothing

            End Try

        End With

    End Sub

    Private Sub GET_CUST_GROUPS()

        SQL = "SELECT IdCliClass, Description, Code FROM CliClass"
        Dim objSQL As New clsSqlConn
        With objSQL
            Try

                DS_BATCHES = .GET_INSERT_UPDATE(SQL)
                'cmbCustGroup.DataSource = DS_BATCHES.Tables(0)
                'cmbCustGroup.DisplayMember = "Description"
                'cmbCustGroup.ValueMember = "IdCliClass"
                'cmbCustGroup.DisplayLayout.Bands(0).Columns(0).Hidden = True
                'cmbCustGroup.DisplayLayout.Bands(0).Columns(1).Width = 100
                'cmbCustGroup.DisplayLayout.Bands(0).Columns(2).Width = 50

            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.Exclamation, "SPIL Glass")
            Finally
                .Dispose()
                objSQL = Nothing
            End Try
        End With

    End Sub

    Public Sub GetProjects(ByVal custId As Integer, Optional ByVal isEnabled As Boolean = False)

        Dim dsProjects As DataSet
        SQL = "SELECT   Id,Name FROM  SpilGlazing_Project WHERE CustomerID=" & custId
        dsProjects = New clsSqlConn().GET_DataSet(SQL)
        If dsProjects.Tables(0).Rows.Count > 0 Then
            cmbCustProject.Enabled = True
        End If

        cmbCustProject.DataSource = dsProjects.Tables(0)
        cmbCustProject.ValueMember = "Id"
        cmbCustProject.DisplayMember = "Name"
        cmbCustProject.DisplayLayout.Bands(0).Columns(1).Width = 200
        cmbCustProject.DisplayLayout.Bands(0).Columns(0).Hidden = True
        If _ProjectId > 0 Then
            cmbCustProject.Value = _ProjectId
        End If
        cmbCustProject.Enabled = isEnabled
    End Sub

    Private Sub GET_SalesRep()
        SQL = "SELECT     idSalesRep, Code, Name ,(Code + '  (' + Name + ')' ) as ToShow  FROM SalesRep ORDER BY Code "
        Dim objSQL As New clsSqlConn
        With objSQL
            Try

                DS_BATCHES = .GET_INSERT_UPDATE(SQL)
                cmbSalesRep.DataSource = DS_BATCHES.Tables(0)
                cmbSalesRep.DisplayMember = "ToShow"
                cmbSalesRep.ValueMember = "idSalesRep"
                cmbSalesRep.DisplayLayout.Bands(0).Columns(0).Width = 50
                cmbSalesRep.DisplayLayout.Bands(0).Columns(1).Width = 139
                cmbSalesRep.DisplayLayout.Bands(0).Columns(2).Width = 200
                cmbSalesRep.DisplayLayout.Bands(0).Columns(0).Hidden = True
                cmbSalesRep.DisplayLayout.Bands(0).Columns(3).Hidden = True

            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.Exclamation, "SPIL Glass")
            Finally
                .Dispose()
                objSQL = Nothing
            End Try
        End With
    End Sub

    Public Sub GetStages(ByVal projectId As Integer, Optional ByVal isEnabled As Boolean = False)
        Try
            Dim dsProjects As DataSet
            If (projectId = 0) Then
                SQL = "SELECT    Id,StageName FROM   SpilGlazing_ProjectStage"
            Else
                SQL = "SELECT    Id,StageName FROM   SpilGlazing_ProjectStage WHERE ProjectID=" & projectId
            End If

            dsProjects = New clsSqlConn().GET_DataSet(SQL)
            If dsProjects.Tables(0).Rows.Count > 0 Then
                cmbProjectStage.Enabled = True
            End If

            cmbProjectStage.DataSource = dsProjects.Tables(0)
            cmbProjectStage.ValueMember = "Id"
            cmbProjectStage.DisplayMember = "StageName"
            cmbProjectStage.DisplayLayout.Bands(0).Columns(1).Width = 200
            cmbProjectStage.DisplayLayout.Bands(0).Columns(0).Hidden = True
            If _StageId > 0 Then
                cmbProjectStage.Value = _StageId
            End If
            cmbProjectStage.Enabled = isEnabled
        Catch ex As Exception
            modGlazingQuoteExtension.GQShowMessage(ex.Message, Me.Text, MsgBoxStyle.Critical)

        End Try
    End Sub

    Private Sub GET_ItemsCategories()
        Dim DS_Stock_Main As DataSet

        SQL = "SELECT ItemCatID, ItemCategory " & _
              "FROM spilItemCategory " & _
              "order by ItemCategory"

        Dim objSQL As New clsSqlConn
        With objSQL
            Try

                DS_Stock_Main = .GET_INSERT_UPDATE(SQL)

                'cmbItemCat.DataSource = Nothing
                'cmbItemCat.DataSource = DS_Stock_Main.Tables(0)
                'cmbItemCat.DisplayMember = "ItemCategory"
                'cmbItemCat.ValueMember = "ItemCatID"
                'cmbItemCat.DisplayLayout.Bands(0).Columns(1).Width = cmbItemCat.Width
                'cmbItemCat.DisplayLayout.Bands(0).Columns(0).Hidden = True

                'cmbItemCat.DisplayLayout.Bands(0).ColHeadersVisible = False

                'cmbItemCat.ActiveRow = cmbItemCat.Rows(0)

            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.Exclamation, "SPIL Glass")
            Finally
                .Dispose()
                objSQL = Nothing
            End Try
        End With
    End Sub



    Private Sub GET_AREAS1() '
        Dim DS_BATCHES As DataSet
        SQL = "SELECT  Code, Description FROM Areas1 order by Description"
        Dim objSQL As New clsSqlConn
        With objSQL
            Try
                DS_BATCHES = .GET_INSERT_UPDATE(SQL)
                txtPost4.DataSource = DS_BATCHES.Tables(0)
                txtPost4.DisplayMember = "Description"
                txtPost4.ValueMember = "Description"
                txtPost4.DisplayLayout.Bands(0).Columns(1).Width = cboArea.Width

                txtPhy4.DataSource = DS_BATCHES.Tables(0)
                txtPhy4.DisplayMember = "Description"
                txtPhy4.ValueMember = "Description"

                txtPhy4.DisplayLayout.Bands(0).Columns(1).Width = cboArea.Width

            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.Exclamation, "SPIL Glass")
            Finally
                .Dispose()
                objSQL = Nothing
            End Try
        End With
    End Sub

    Private Sub GET_AREAS() '
        Dim DS_BATCHES As DataSet
        SQL = "SELECT  idAreas, Description FROM Areas order by Description"
        Dim objSQL As New clsSqlConn
        With objSQL
            Try
                DS_BATCHES = .GET_INSERT_UPDATE(SQL)
                cboArea.DataSource = DS_BATCHES.Tables(0)
                cboArea.DisplayMember = "Description"
                cboArea.ValueMember = "idAreas"
                cboArea.DisplayLayout.Bands(0).Columns(0).Hidden = True '.Width = 50
                cboArea.DisplayLayout.Bands(0).Columns(1).Width = cboArea.Width

            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.Exclamation, "SPIL Glass")
            Finally
                .Dispose()
                objSQL = Nothing
            End Try
        End With
    End Sub

    Sub GET_StkItems()
        Dim DS_BATCHES As DataSet
        SQL = "SELECT StockLink, Code, Description_1  FROM StkItem"
        Dim objSQL As New clsSqlConn
        With objSQL
            Try
                DS_BATCHES = .GET_INSERT_UPDATE(SQL)
                ucmbMainStkCmb.DataSource = DS_BATCHES.Tables(0)
                ucmbMainStkCmb.DisplayMember = "cSimpleCode"
                ucmbMainStkCmb.ValueMember = "StockLink"
                ucmbMainStkCmb.DisplayLayout.Bands(0).Columns(0).Hidden = True
                ucmbMainStkCmb.DisplayLayout.AutoFitStyle = AutoFitStyle.ResizeAllColumns

            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.Exclamation, "SPIL Glass")
            Finally
                .Dispose()
                objSQL = Nothing
            End Try
        End With
    End Sub


    Private Sub GetTerms() '
        Dim DS_BATCHES As DataSet
        SQL = "SELECT AccountTerms, AccTermDescription FROM spilCustTerms"
        Dim objSQL As New clsSqlConn
        With objSQL
            Try
                DS_BATCHES = .GET_INSERT_UPDATE(SQL)
                uCmbTerms.DataSource = DS_BATCHES.Tables(0)
                uCmbTerms.DisplayMember = "AccTermDescription"
                uCmbTerms.ValueMember = "AccountTerms"
                uCmbTerms.DisplayLayout.Bands(0).Columns(0).Hidden = True '.Width = 50
                uCmbTerms.DisplayLayout.Bands(0).Columns(1).Width = cboArea.Width

            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.Exclamation, "SPIL Glass")
            Finally
                .Dispose()
                objSQL = Nothing
            End Try
        End With
    End Sub

    Private Function GET_TaxRates() As DataSet
        SQL = "SELECT * FROM TaxRate Order by code"
        Dim objSQL As New clsSqlConn
        Dim dsTax As DataSet
        With objSQL
            Try
                dsTax = .GET_INSERT_UPDATE(SQL)

                ucmbTaxRate.DataSource = Nothing
                ucmbTaxRate.DataSource = dsTax.Tables(0)
                ucmbTaxRate.DisplayMember = "TaxRate"
                ucmbTaxRate.ValueMember = "idTaxRate"
                ucmbTaxRate.DisplayLayout.Bands(0).Columns("idTaxRate").Hidden = True
                'ucmbTaxRate.DisplayLayout.Bands(0).Columns("Code").Hidden = True
                ucmbTaxRate.DisplayLayout.Bands(0).Columns("TaxRate").Hidden = False
                'cmbTaxRate.DisplayLayout.Bands(0).Columns("Description").InvalidValueBehavior = InvalidValueBehavior.RevertValueAndRetainFocus
                ucmbTaxRate.DisplayLayout.Bands(0).ColHeadersVisible = False
                ucmbTaxRate.Value = Nothing
                'If dsTax.Tables(0).Rows.Count > 0 Then
                '    Dim ugRow As UltraGridRow = cmbTaxRate.GetRow(ChildRow.First)
                '    ugRow.Activate()
                'End If
                'ucmbTaxRate.Visible = True
                Return dsTax
            Catch ex As Exception
                Return Nothing
                MsgBox(ex.Message, MsgBoxStyle.Exclamation, "SPIL Glass")
            Finally
                '.Dispose()
                objSQL = Nothing
            End Try
        End With
    End Function

    Public Sub LoadQuoteState()
        Dim DS_ As DataSet
        Try
            SQL = "SELECT JQSID, JQSName, JQSState FROM GlzQuote_State Where JQSState = 1"
            Dim objSQL As New clsSqlConn
            With objSQL
                DS_BATCHES = .GET_INSERT_UPDATE(SQL)
                utxtQuoteState.DataSource = DS_BATCHES.Tables(0)
                utxtQuoteState.DisplayMember = "JQSName"
                utxtQuoteState.ValueMember = "JQSID"
                utxtQuoteState.DisplayLayout.Bands(0).Columns("JQSID").Hidden = True
                utxtQuoteState.DisplayLayout.Bands(0).Columns("JQSState").Hidden = True
                utxtQuoteState.DisplayLayout.Bands(0).ColHeadersVisible = False

                utxtQuoteState.Value = 1
            End With


        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Exclamation, "SPIL Glass")

        End Try

    End Sub
    'Fill Currency and Exchange rate
    Private Sub GetCurrencyData()
        Try
            SQL = "SELECT currencyCode, ExchangeRate FROM  spil_InvCurrency WHERE IsLocalCurrency = 'TRUE'"
            dsCurrency = New clsSqlConn().GET_DataSet(SQL)

            If dsCurrency.Tables(0).Rows.Count > 0 Then
                txtCurrency.Text = dsCurrency.Tables(0).Rows(0)("currencyCode")
                txtExRate.Text = dsCurrency.Tables(0).Rows(0)("ExchangeRate")

                'Q1.Text = "(" & dsCurrency.Tables(0).Rows(0)("currencyCode") & ")"
                'Q2.Text = "(" & dsCurrency.Tables(0).Rows(0)("currencyCode") & ")"
                'Q3.Text = "(" & dsCurrency.Tables(0).Rows(0)("currencyCode") & ")"

                T1.Text = "(" & dsCurrency.Tables(0).Rows(0)("currencyCode") & ")"
                T2.Text = "(" & dsCurrency.Tables(0).Rows(0)("currencyCode") & ")"
                T3.Text = "(" & dsCurrency.Tables(0).Rows(0)("currencyCode") & ")"
            End If
        Catch

        End Try
    End Sub

    '-----------------------------

    Private Sub cmbAccount_ValueChanged(sender As Object, e As EventArgs) Handles cmbAccount.ValueChanged
        Try
            If IsNumeric(cmbAccount.Value) = False Then
                If cmbAccount.Value = "" Then
                    ' modGlazingQuoteExtension.GQShowMessage("Please select a customer", Me.Text, MsgBoxStyle.Question, "warning")

                Else
                    modGlazingQuoteExtension.GQShowMessage("No matching customer", Me.Text, MsgBoxStyle.Question, "warning")
                End If
                UG2.Enabled = False
                'If IsNothing(saveButtonPreState) = False Then
                '    If saveButtonPreState = False Then
                '        saveButtonPreState = mnuSave.Enabled

                '    End If
                'Else
                '    saveButtonPreState = mnuSave.Enabled

                'End If


                'tsbSave.Enabled = False
                'mnuSave.Enabled = False
                Exit Sub
            Else
                'mnuSave.Enabled = saveButtonPreState
                'tsbSave.Enabled = saveButtonPreState

            End If
            LoadQuoteLineType()
            InitializesCustomerDetails()
            UG2.Enabled = True
            Dim row As UltraGridRow

            If quoteOrdeIndex = 0 And cmbAccount.Text <> "" Then
                If getTaxRateFromCustomer = True Then
                    oFormCustomer.GetCustomerData(cmbAccount.Value)
                    Dim newTable As DataTable = ucmbTaxRate.DataSource
                    Dim newDatset As DataSet
                    Dim newRows() As DataRow
                    If IsNothing(newTable) = False Then
                        newRows = newTable.Select("Code = '" & oFormCustomer.TaxCode & "'")
                    Else
                        newDatset = GET_TaxRates()
                        If IsNothing(newDatset) = False Then
                            newRows = newDatset.Tables(0).Select("Code = '" & oFormCustomer.TaxCode & "'")
                        End If
                    End If
                    For Each newRow As DataRow In newRows
                        defaultTaxtRateValue = oFormCustomer.TaxRate
                        defaultTaxtRate = newRow("idTaxRate")
                    Next
                    defaultTaxRateValueForQuote = defaultTaxtRateValue
                    defaultTaxRateForQuote = defaultTaxtRate
                End If

                If Me.UG2.Rows.Count = 0 Then
                    row = AddNewRow("after")
                    row.ParentCollection.Move(row, UG2.ActiveRow.Index)
                    Me.UG2.ActiveRowScrollRegion.ScrollRowIntoView(row)
                ElseIf Me.UG2.Rows.Count > 0 Then
                    Me.UG2.Rows((UG2.Rows.Count) - 1).Cells("QuoteFiedType").DroppedDown = True
                    AfterDefaultTaxePriceChaged(defaultTaxtRate, defaultTaxtRateValue, True)
                End If
            Else
                If getTaxRateFromCustomer = True Then
                    defaultTaxtRateValue = defaultTaxRateValueForQuote
                    defaultTaxtRate = defaultTaxRateForQuote

                End If

            End If
            If cmbAccount.Text <> "" Then
                txtPostelAdd.Enabled = True
            End If

            If IsFromJobProject Then
                clsGQExtensionForJobCostingObj.GetProjects(cmbAccount.SelectedRow.Cells("DCLink").Value, True)
                clsGQExtensionForJobCostingObj.GetJobsByCustomer(cmbAccount.SelectedRow.Cells("DCLink").Value, True)
            End If

            Dim enable As Boolean
            If defaultTaxRateValueForQuote > 0 Then
                enable = True
            Else
                enable = False
            End If
            lblTotVatAmo.Visible = enable
            lblTotIncAmo.Visible = enable
            lblTotalInc.Visible = enable
            lblTotVatAmo.Visible = enable
            T2.Visible = enable
            T3.Visible = enable

        Catch ex As Exception
            modGlazingQuoteExtension.GQShowMessage(ex.Message, Me.Text, MsgBoxStyle.Critical)

        End Try
    End Sub

    Private Sub cmbCusType_ValueChanged(sender As Object, e As EventArgs) Handles cmbCusType.ValueChanged, UltraComboEditor3.ValueChanged
        GET_CUSTOMERS(cmbCusType.Text)
    End Sub

    Private Sub frmGlazingQuote_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Try
            For Each row As UltraGridRow In UG2.Rows
                gridActiveRow = row
                UG2.ActiveRow = row
                lineTypeCell = row.Cells("QuoteFiedType")
                QuoteGirdRowStyling(True)
            Next
            SetTaxedPriceLable(isTaxedPrice)
        Catch ex As Exception

        End Try
    End Sub

    Public Sub InitializeQuotation()
        GetCurrencyData()
        GetQuoteTaxState(-1, Nothing)
        GET_SalesRep()
        GET_AREAS1()
        GET_AREAS()
        GET_StkItems()
        GetTerms()
        cmbCusType.Value = 0
        GET_TaxRates()
        GET_Facility()
        LoadQuoteState()
        LoadExstingQuote()
        TotalValuesBeahavior()
        clsGQExtensionForJobCostingObj.HideProfitCostFields(IsFromJobProject)
        ' GET_Del_Method()
        'GET_Couriers()
        'GET_CUST_GROUPS()
        'GET_ItemsCategories()

    End Sub

    Private Sub GET_Facility()

        SQL = "SELECT     FacilityID, FacilityName  FROM spilPROD_Facility ORDER BY FacilityID"   '

        Dim objSQL As New clsSqlConn

        With objSQL
            Try

                DS_BATCHES = .GET_INSERT_UPDATE(SQL)
                cmbFacility.DataSource = DS_BATCHES.Tables(0)
                cmbFacility.DisplayMember = "FacilityName"
                cmbFacility.ValueMember = "FacilityID"
                cmbFacility.DisplayLayout.Bands(0).ColHeadersVisible = False
                cmbFacility.DisplayLayout.Bands(0).Columns(0).Hidden = True ' Width = 139
                cmbFacility.DisplayLayout.Bands(0).Columns(1).Width = cmbFacility.Width '200

                uddBranch.DataSource = DS_BATCHES.Tables(0)
                uddBranch.DisplayMember = "FacilityName"
                uddBranch.ValueMember = "FacilityID"
                uddBranch.DisplayLayout.Bands(0).Columns(0).Hidden = True ' Width = 139

                uddBranch.DisplayLayout.Bands(0).Columns(1).Header.Caption = "Branch"

                If cmbFacility.Rows.Count > 0 Then
                    cmbFacility.Value = cmbFacility.GetRow(ChildRow.First).Cells("FacilityID").Value
                End If

            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.Exclamation, "SPIL Glass")

            Finally

                .Dispose()
                objSQL = Nothing

            End Try

        End With

    End Sub

    Private Sub setQuateFiedType(sender As Object, e As InitializeLayoutEventArgs) Handles ugQuote.InitializeLayout, UltraGrid2.InitializeLayout
        Try
            e.Layout.AutoFitStyle = AutoFitStyle.ResizeAllColumns
            e.Layout.Override.RowSizing = RowSizing.Fixed
            e.Layout.Override.CellMultiLine = DefaultableBoolean.True

            Dim quateFiedTypesList As ValueList = New ValueList()
            quateFiedTypesList.ValueListItems.Add(1, "Text")
            quateFiedTypesList.ValueListItems.Add(2, "Header-Main")
            quateFiedTypesList.ValueListItems.Add(3, "Header-Sub")
            quateFiedTypesList.ValueListItems.Add(4, "Subtotal")
            quateFiedTypesList.ValueListItems.Add(5, "Stock Item")
            With e.Layout.Bands(0).Columns("QuoteFiedType")
                .Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownValidate
                '.ValueList = quateFiedTypesList
                '.EditorComponent = ucmbQuoteLineType

            End With
        Catch ex As Exception
            modGlazingQuoteExtension.GQShowMessage(ex.Message, Me.Text, MsgBoxStyle.Critical)

        End Try
    End Sub

    Private Sub LoadQuoteLineType()
        Try

        Catch ex As Exception
            modGlazingQuoteExtension.GQShowMessage(ex.Message, Me.Text, MsgBoxStyle.Critical)

        End Try
    End Sub
    Sub QuoteGirdRowFormat(e As CellEventArgs)
        Try
            If Me.UG2.ActiveCell.Column.Key = "QuoteFiedType" Then
                Me.UG2.ActiveRow.Cells("LineComments").Appearance.Reset()
                Me.UG2.ActiveRow.Appearance.FontData.Reset()

                If isExistingOrder = False Then
                    If Me.UG2.ActiveRow.Cells("LineComments").Text <> "" Or Me.UG2.ActiveRow.Cells("Amount").Text <> "0.00" Then
                        If Me.UG2.ActiveCell.Text = "Header-Main" Then
                            If canUpdate = True Then
                                CellValuesClear()
                            End If
                        End If
                    End If
                End If

                If Me.UG2.ActiveCell.Text = "Header-Main" Then
                    Me.UG2.ActiveRow.Appearance.FontData.Bold = DefaultableBoolean.True
                    Me.UG2.ActiveRow.Cells("LineComments").Appearance.FontData.SizeInPoints = 14
                    ColumsVisibility(e, True)

                ElseIf Me.UG2.ActiveCell.Text = "Header-Sub" Then
                    If groupStarted = False Then
                        groupStarted = True
                        subHeaderID = subHeaderID + 1
                    End If
                    Me.UG2.ActiveRow.Appearance.FontData.Bold = DefaultableBoolean.True
                    Me.UG2.ActiveRow.Appearance.FontData.Underline = DefaultableBoolean.True
                    Me.UG2.ActiveRow.Cells("LineComments").Appearance.FontData.SizeInPoints = 12
                    Me.UG2.ActiveRow.Cells("ItmGroupID").Value = subHeaderID
                    ColumsVisibility(e, True)

                ElseIf Me.UG2.ActiveCell.Text = "Subtotal" Then
                    Me.UG2.ActiveRow.Appearance.FontData.Bold = DefaultableBoolean.True
                    Me.UG2.ActiveRow.Cells("Amount").Appearance.BackColor = Color.LightGreen
                    Me.UG2.ActiveRow.Cells("Amount").Appearance.BorderColor = Color.Green
                    ColumsVisibility(e, True)

                    Me.UG2.ActiveRow.Cells("Amount").Hidden = False
                    Me.UG2.ActiveRow.Cells("LineComments").Hidden = True
                    'Me.UG2.ActiveRow.Cells("ItmGroupID").Value = subHeaderID
                    'Me.UG2.ActiveRow.Cells("QuoteFiedType").Value = Subtotal
                    SetSubTotal(e)

                ElseIf Me.UG2.ActiveCell.Text = "Stock Item" Then
                    If isStockItemClosing = False Then
                        'e.Cell.Row.Appearance.BackColor = Color.WhiteSmoke
                        ColumsVisibility(e, False)

                        If IsNothing(UG2.ActiveRow) = False Then
                            'UG2ActiveRow = UG2.ActiveRow

                            If isPasting = False And isOpeningQuote = False Then
                                Dim glazingDocStockItem As New frmGlazingDocStockItem(Me)
                                glazingDocStockItem.ShowDialog()
                            End If

                            Me.UG2.ActiveRow.Cells("Amount").Tag = Me.UG2.ActiveRow.Cells("ItmGroupID").Value
                            ' Me.UG2.ActiveRow.Cells("ItmGroupID").Value = subHeaderID

                        End If
                    Else

                    End If
                Else
                    Me.UG2.ActiveRow.Cells("LineComments").Appearance.Reset()
                    Me.UG2.ActiveRow.Appearance.FontData.Bold = DefaultableBoolean.False
                    Me.UG2.ActiveRow.Appearance.FontData.Underline = DefaultableBoolean.False
                    Me.UG2.ActiveRow.Cells("Amount").Tag = Me.UG2.ActiveRow.Cells("ItmGroupID").Value
                    'Me.UG2.ActiveRow.Cells("ItmGroupID").Value = subHeaderID

                    ColumsVisibility(e, False)
                End If

                If Me.UG2.ActiveCell.Text = "Subtotal" Then
                    'Me.UG2.ActiveCell = Me.UG2.ActiveRow.Cells("Amount")
                Else
                    ' Me.UG2.ActiveCell = Me.UG2.ActiveRow.Cells("LineComments")
                    'Me.UG2.ActiveCell = Me.UG2.ActiveRow.Cells("LineComments")

                End If
                Me.UG2.ActiveRow.PerformAutoSize()
                Me.UG2.PerformAction(Infragistics.Win.UltraWinGrid.UltraGridAction.EnterEditMode, False, False)
            End If
        Catch ex As Exception
            modGlazingQuoteExtension.GQShowMessage(ex.Message, Me.Text, MsgBoxStyle.Critical)

        End Try
    End Sub
    Sub ColumsVisibility(e As CellEventArgs, isHdden As Boolean)
        Try
            Me.UG2.ActiveRow.Cells("Height").Hidden = isHdden
            Me.UG2.ActiveRow.Cells("Width").Hidden = isHdden
            Me.UG2.ActiveRow.Cells("Qty").Hidden = isHdden
            Me.UG2.ActiveRow.Cells("Price").Hidden = isHdden
            Me.UG2.ActiveRow.Cells("Amount").Hidden = isHdden
            Me.UG2.ActiveRow.Cells("LineNotes").Hidden = isHdden
            Me.UG2.ActiveRow.Cells("MarkAs").Hidden = isHdden
            Me.UG2.ActiveRow.Cells("ItemImage").Hidden = isHdden
            Me.UG2.ActiveRow.Cells("TaxRate").Hidden = isHdden
            Me.UG2.ActiveRow.Cells("Shape").Hidden = isHdden
        Catch ex As Exception
            modGlazingQuoteExtension.GQShowMessage(ex.Message, Me.Text, MsgBoxStyle.Critical)

        End Try
    End Sub
    Public Sub CellValuesClear()
        Try
            If isPasting = False Then
                Me.UG2.ActiveRow.Cells("Height").Value = 0
                Me.UG2.ActiveRow.Cells("Width").Value = 0
                Me.UG2.ActiveRow.Cells("Qty").Value = 0
                Me.UG2.ActiveRow.Cells("Price").Value = 0
                Me.UG2.ActiveRow.Cells("Amount").Value = 0
                Me.UG2.ActiveRow.Cells("LineNotes").Value = ""
                Me.UG2.ActiveRow.Cells("MarkAs").Value = ""
                Me.UG2.ActiveRow.Cells("ItemImage").Value = DBNull.Value
                Me.UG2.ActiveRow.Cells("LineComments").Value = ""
            End If
        Catch ex As Exception
            modGlazingQuoteExtension.GQShowMessage(ex.Message, Me.Text, MsgBoxStyle.Critical)

        End Try
    End Sub
    Public Sub CellValuesClearExtended()
        Try
            Dim ugActiveRow As UltraGridRow = UG2.ActiveRow

            ugActiveRow.Cells("ItemTypeCategory").Value = String.Empty
            ugActiveRow.Cells("Description1").Value = String.Empty
            ugActiveRow.Cells("SimpleCode").Value = String.Empty
            ugActiveRow.Cells("Description2").Value = String.Empty

            ugActiveRow.Cells("PriceType").Value = DBNull.Value
            ugActiveRow.Cells("PriceCat").Value = 0
            ugActiveRow.Cells("PriceList").Value = DBNull.Value

            ugActiveRow.Cells("IsPriceItem").Value = True
            ugActiveRow.Cells("Width").Value = 0.0
            ugActiveRow.Cells("Height").Value = 0.0
            ugActiveRow.Cells("Volume").Value = 0.0
            ugActiveRow.Cells("Qty").Value = 0
            ugActiveRow.Cells("Price").Value = 0

            ugActiveRow.Cells("Net").Value = 0
            ugActiveRow.Cells("Tax").Value = 0
            ugActiveRow.Cells("LineTot").Value = 0

            QuoteGirdRowFormatExtended()

        Catch ex As Exception
            modGlazingQuoteExtension.GQShowMessage(ex.Message, Me.Text, MsgBoxStyle.Critical)

        End Try
    End Sub

    Sub QuoteGirdRowFormatExtended()
        Dim ugActiveRow As UltraGridRow = UG2ActiveRow

        Try
            Select Case ugActiveRow.Cells("ItemType").Value
                Case GlassItemTypes.Glass

                    ugActiveRow.Cells("Service").Hidden = False
                    ugActiveRow.Cells("Width").Hidden = False
                    ugActiveRow.Cells("Height").Hidden = False

                Case GlassItemTypes.Template

                    ugActiveRow.Cells("Service").Hidden = True
                    ugActiveRow.Cells("Width").Hidden = False
                    ugActiveRow.Cells("Height").Hidden = False
                Case GlassItemTypes.Consumable

                    ugActiveRow.Cells("Service").Hidden = True
                    ugActiveRow.Cells("Width").Hidden = True
                    ugActiveRow.Cells("Height").Hidden = True
                    ugActiveRow.Cells("Qty").Value = 1

                Case GlassItemTypes.Service

                    ugActiveRow.Cells("Service").Hidden = True
                    ugActiveRow.Cells("Width").Hidden = True
                    ugActiveRow.Cells("Height").Hidden = True
                    ugActiveRow.Cells("Qty").Value = 1

                Case GlassItemTypes.Aluminium
                    ugActiveRow.Cells("Service").Hidden = False
                    ugActiveRow.Cells("Width").Hidden = False
                    ugActiveRow.Cells("Height").Hidden = False
                    ugActiveRow.Cells("Qty").Value = 1

            End Select

        Catch ex As Exception
            modGlazingQuoteExtension.GQShowMessage(ex.Message, Me.Text, MsgBoxStyle.Critical)

        End Try

    End Sub


    Private Sub btnAddRow_Click(sender As Object, e As EventArgs) Handles btnAddRowButton16.Click, Button7.Click
        Try
            Dim row As UltraGridRow = UG2.DisplayLayout.Bands(0).AddNew
            row.ParentCollection.Move(row, UG2.ActiveRow.Index)
            Me.UG2.ActiveRowScrollRegion.ScrollRowIntoView(row)
        Catch ex As Exception
            modGlazingQuoteExtension.GQShowMessage(ex.Message, Me.Text, MsgBoxStyle.Critical)

        End Try
    End Sub

    Private Sub tsbSave_Click(sender As Object, e As EventArgs) Handles tsbSave.Click
        SaveDocument()
    End Sub

    Function SaveDocument(Optional ByRef isPrintPreview As Boolean = False) As Integer
        Dim saveWithoutItem As Boolean = False
        Dim noDescrption As Boolean = False

        Try
            If isCancelled = True Then
                Return 0
                Exit Function
            End If
            If IsNumeric(cmbAccount.Value) = False Then
                modGlazingQuoteExtension.GQShowMessage("Customet not found", Me.Text, MessageBoxButtons.OK, "warning")
                cmbAccount.Focus()
                Return 0
                Exit Function
            End If
            If Me.UG2.Rows.Count < 1 Then

                modGlazingQuoteExtension.GQShowMessage("Quotation don't have any items", Me.Text, MessageBoxButtons.OK, "warning")

                'If modGlazingQuoteExtension.GQShowMessage("Do you wont to save quotation without items?", Me.Text, MessageBoxButtons.YesNo, "question") = Windows.Forms.DialogResult.No Then
                Return 0
                Exit Function
                'Else
                '    If IsNothing(jobDescription) = False Then
                '        If jobDescription = "" Then
                '            noDescrption = True
                '        Else
                '            saveWithoutItem = True
                '        End If
                '    Else
                '        noDescrption = True
                '    End If
                'End If
                'If noDescrption = True Then
                '    modGlazingQuoteExtension.GQShowMessage("You should either add items or job description before saving.", Me.Text, MessageBoxButtons.OK, "warning")
                '    Return 0
                '    noDescrption = False
                '    Exit Function
                'End If
            ElseIf isClosing = False And isPrintPreview = False Then
                If Me.UG2.Rows.Count = 1 Then
                    If Me.UG2.Rows(0).Cells("QuoteFiedType").Value = "" Then
                        modGlazingQuoteExtension.GQShowMessage("Quotation don't have any items", Me.Text, MessageBoxButtons.OK, "warning")
                        Exit Function

                    End If
                End If
                If modGlazingQuoteExtension.GQShowMessage("Do you want to save this Quotation?", Me.Text, MessageBoxButtons.YesNo, "question") = Windows.Forms.DialogResult.No Then
                    Return 0
                    Exit Function
                End If
            End If

            If IsNothing(objClsInvHeader) = True Then
                objClsInvHeader = New clsInvHeader

            End If

            Dim orderIndex As Integer = 0
            Dim recState As Integer = 0
            If OrderValidation() = 0 Then
                Return 0
                Exit Function
            End If

            If IsFromJobProject Then
                If IsEstimate = True Then
                    objClsInvHeader.DocState = GlassDocState.JobEstimate

                    Dim objSQL As New clsSqlConn

                    If _ProjectId = 0 Then

                        duplicatesql = "select COUNT(OrderIndex) as OrderIndex FROM spilInvNum WHERE DocType=5 AND DocState=15 AND QuoteStateID <> 4 AND JobID = " & _JobId.ToString()

                    Else
                        duplicatesql = "Select Sum(u.OrderIndex) FROM" &
                                            "(select COUNT(OrderIndex) as OrderIndex FROM spilInvNum WHERE DocType=5 AND DocState=15 AND QuoteStateID <> 4 AND JobID = " & _JobId.ToString() &
                                            " UNION select COUNT(OrderIndex) as OrderIndex FROM spilInvNum WHERE DocType=5 AND DocState=15 AND QuoteStateID <> 4 AND ProjectID = " & _ProjectId.ToString() & " AND (JobID is null or JobID=0)) AS u"

                    End If


                    With objSQL
                        Dim dupds As DataSet = .GET_DATA_SQL(duplicatesql)
                        If (dupds.Tables(0).Rows(0)(0) <> 0) Then
                            modGlazingQuoteExtension.GQShowMessage("This Job/Project already has a Estimate", Me.Text, MsgBoxStyle.Critical)
                            Exit Function
                        End If

                    End With

                Else
                    objClsInvHeader.DocState = GlassDocState.GlazingQuote
                    Dim objSQL As New clsSqlConn

                    If _ProjectId = 0 Then

                        duplicatesql = "select COUNT(OrderIndex) as OrderIndex FROM spilInvNum WHERE DocType=5 AND DocState=15 AND QuoteStateID <> 4 AND JobID = " & _JobId.ToString()

                    Else
                        duplicatesql = "Select Sum(u.OrderIndex) FROM" &
                                       "(select COUNT(OrderIndex) as OrderIndex FROM spilInvNum WHERE DocType=5 AND DocState=15 AND QuoteStateID <> 4 AND JobID = " & _JobId.ToString() &
                                       " UNION select COUNT(OrderIndex) as OrderIndex FROM spilInvNum WHERE DocType=5 AND DocState=15 AND QuoteStateID <> 4 AND ProjectID = " & _ProjectId.ToString() & " AND (JobID is null or JobID=0)) AS u"

                    End If

                    With objSQL
                        Dim dupds As DataSet = .GET_DATA_SQL(duplicatesql)
                        If (dupds.Tables(0).Rows(0)(0) <> 0) Then
                            modGlazingQuoteExtension.GQShowMessage("This Job/Project already has a Quote", Me.Text, MsgBoxStyle.Critical)
                            Exit Function
                        End If

                    End With
                End If
            End If


            objClsInvHeader.Begin_Trans()

            'Start quote Header
            objClsInvHeader.AccountID = cmbAccount.ActiveRow.Cells("DCLink").Value
            objClsInvHeader.DocType = GlassDocTypes.Quotation
            If IsFromJobProject = False Then
                objClsInvHeader.DocState = GlassDocState.GlazingQuote
            End If
            objClsInvHeader.InvoiceNotes = utxtNoteText.Text
            objClsInvHeader.ApprovedReason = ""
            objClsInvHeader.iAgentID = AgentID

            If isPrintPreview = True Then
                objClsInvHeader.OrderNum = "QuotePrintPreview"
            ElseIf isExistingOrder = False Or isACopy = True Then
                objClsInvHeader.OrderNum = clsGlazingQuoteExtensionObj.GetNextJobumentNumber(objClsInvHeader, "GlazingQuote")
                lblOrderNo.Text = objClsInvHeader.OrderNum
            Else
                objClsInvHeader.OrderNum = lblOrderNo.Text
            End If

            objClsInvHeader.FacilityIDCurrent = cmbFacility.Value
            objClsInvHeader.FacilityIDOrigin = cmbFacility.Value
            objClsInvHeader.iCusType = cmbCusType.Value
            objClsInvHeader.EnteredBy = strUserName

            objClsInvHeader.TaxRate = defaultTaxtRate
            objClsInvHeader.ExtOrderNum = txtCustOrdNo.Text.Trim
            objClsInvHeader.iAgentID = AgentID
            objClsInvHeader.LastEditedBy = AgentID
            objClsInvHeader.LastEditedDateTime = Today.Date
            objClsInvHeader.cAccountName = cmbAccount.ActiveRow.Cells("DCLink").Value

            objClsInvHeader.ContactName = txtContact1.Text
            objClsInvHeader.ContactTelephone = txtContPerTel.Text
            objClsInvHeader.Address1 = txtPhy1.Text
            objClsInvHeader.Address2 = txtPhy2.Text
            objClsInvHeader.Address3 = txtPhy3.Text
            objClsInvHeader.Address4 = txtPhy4.Text
            objClsInvHeader.Address5 = txtPhy5.Text
            objClsInvHeader.Address6 = txtPhyPostCode.Text

            objClsInvHeader.PostAdd1 = txtPost1.Text
            objClsInvHeader.PostAdd2 = txtPost2.Text
            objClsInvHeader.PostAdd3 = txtPost3.Text
            objClsInvHeader.PostAdd4 = txtPost4.Text
            objClsInvHeader.PostAdd5 = txtPost5.Text
            objClsInvHeader.PostPC = txtPostCode.Text

            objClsInvHeader.ContactEmail = txtContEmail.Text
            objClsInvHeader.iAreasID = cboArea.Value
            objClsInvHeader.CreditState = uCmbTerms.Value

            objClsInvHeader.OrderIndex = quoteOrdeIndex

            objClsInvHeader.Delivery_Status = subHeaderID
            objClsInvHeader.OrderDate = Today.Date
            objClsInvHeader.DueDate = txtDueDate.Value
            objClsInvHeader.InvoiceNotes = utxtNoteText.Text
            'objClsInvHeader.Delivery_Status = ""

            'Link the Project, Stage & Job
            objClsInvHeader.ProjectID = _ProjectId
            objClsInvHeader.ProjID = _ProjectId

            objClsInvHeader.Delivery_Status = DeliveryState.UnDelivered 'for delivery
            objClsInvHeader.ProductionState = GlassProdState.None
            objClsInvHeader.QuotedAmt = lblTotExcAmo.Text
            objClsInvHeader.Quoted_Tax = lblTotVatAmo.Text
            objClsInvHeader.Quoted_Incl = lblTotIncAmo.Text
            objClsInvHeader.DocRepID = cmbSalesRep.Value
            objClsInvHeader.ItemCatID = 1

            'objClsInvHeader.iClassID = cmbAccount.ActiveRow.Cells("iClassID").Value

            'objClsInvHeader.OrderNum = objClsInvHeader.GetNextDocumentNumber
            If isPrintPreview = True Then
                'orderIndex = objClsInvHeader.AddHeader2()
            Else
                orderIndex = objClsInvHeader.AddHeader()
                If orderIndex = -1 Or SaveGlzQuoteJobDetails(objClsInvHeader, orderIndex) = 0 Then
                    objClsInvHeader.Rollback_Trans()
                    Exit Function

                End If
            End If

            quoteOrdeIndex = orderIndex
            UpdateOrderHeaderWhileSaving(quoteOrdeIndex)

            Dim objClsInvHeaderDetailLine As New clsInvDetailLine
            Dim rowTypes As GridRowType = GridRowType.DataRow
            Dim band As UltraGridBand = Me.UG2.DisplayLayout.Bands(0)
            Dim enumerator As IEnumerable = band.GetRowEnumerator(rowTypes)
            Dim subTotal As Integer = 0
            Dim row As UltraGridRow

            If isPrintPreview = False Then
                If collDeletedItemLines.Count > 0 Then
                    For c = 1 To collDeletedItemLines.Count
                        If objClsInvHeader.DeleteLinesFromDB(collDeletedItemLines.Item(c), "ItemLines", oProdDefaults) <> 1 Then
                            objClsInvHeader.Rollback_Trans()
                            MsgBox("Error occured while deleting Item lines", MsgBoxStyle.Information, "Validation")
                            Exit Function
                        End If
                    Next

                    'objClsInvHeader.Commit_Trans()
                    Dim objDLItem As New clsDocumentLogEntry
                    objDLItem.iDocID = InvHeaderID
                    objDLItem.iDocTypeID = pubMeSpilDocTypeID
                    objDLItem.LogAction = "Delete Items"
                    objDLItem.DocItemCount = c - 1
                    objDLItem.DocServiceCount = 0
                    objDLItem.LogDateTime = Now
                    objDLItem.EnteredBy = strUserName
                    objDLItem.Description1 = c - 1 & " Line Item(s) deleted"
                    If objDLItem.AddDocLogWithTrans(objClsInvHeader.Con, objClsInvHeader.Trans) = 0 Then
                        objClsInvHeader.Rollback_Trans()
                        Exit Function
                    End If
                    objDLItem = Nothing
                End If
            End If

            If saveWithoutItem = False Then
                ' Dim idInvoiceLines = 0
                For Each row In enumerator
                    'idInvoiceLines = row.Index
                    objClsInvHeaderDetailLine = New clsInvDetailLine

                    '----Starting Item Line identifers----
                    objClsInvHeaderDetailLine.iInvDetailID = row.Cells("iInvDetailID").Value
                    objClsInvHeaderDetailLine.LN = row.Cells("ItmGroupID").Value
                    objClsInvHeaderDetailLine.OrderIndex = orderIndex
                    objClsInvHeaderDetailLine.idInvoiceLines = row.Index

                    '----Ending Item Line identifers----

                    '----Starting Stock Items details----
                    If row.Cells("QuoteFiedType").Value <> QuateFiedTypesList.Stock_Item Then
                        objClsInvHeaderDetailLine.StockLink = defaultItemForText
                        objClsInvHeaderDetailLine.ItemType = 4
                        objClsInvHeaderDetailLine.cDescription = row.Cells("LineComments").Value
                    Else
                        objClsInvHeaderDetailLine.StockLink = row.Cells("StockLink").Text
                        objClsInvHeaderDetailLine.ItemType = row.Cells("ItemType").Text
                        objClsInvHeaderDetailLine.cDescription = row.Cells("Description1").Value
                    End If


                    '----Starting Stock Items details----

                    If row.Cells("Qty").Text = "" Then
                        objClsInvHeaderDetailLine.fQuantity = 0
                    Else
                        objClsInvHeaderDetailLine.fQuantity = row.Cells("Qty").Text
                    End If
                    objClsInvHeaderDetailLine.iHeight = row.Cells("Height").Text
                    objClsInvHeaderDetailLine.iWidth = row.Cells("Width").Text

                    'For fractions--
                    'objClsInvHeaderDetailLine.DisHeight = row.Cells("Dis.Height").Text
                    'objClsInvHeaderDetailLine.DisWidth = row.Cells("Dis.Width").Text
                    'objClsInvHeaderDetailLine.SQFeetForPricing = row.Cells("SQFeetForPricing").Value

                    '------

                    objClsInvHeaderDetailLine.fVolume = row.Cells("Volume").Value

                    '----Starting finance data----
                    If row.Cells("QuoteFiedType").Value <> QuateFiedTypesList.Stock_Item Then
                        objClsInvHeaderDetailLine.PRICE_TYPES_ID = 15
                        objClsInvHeaderDetailLine.PriceList = oSOModuleDefaults.DefaultTradePriceListID
                        objClsInvHeaderDetailLine.Measure = 1
                        objClsInvHeaderDetailLine.PriceCategory = "T"
                    Else
                        objClsInvHeaderDetailLine.PRICE_TYPES_ID = row.Cells("Price_Type").Value
                        'objClsInvHeaderDetailLine.PriceList = row.Cells("PriceList").Value
                        objClsInvHeaderDetailLine.PriceList = IIf(IsDBNull(row.Cells("PriceList").Value), 0, row.Cells("PriceList").Value)
                        objClsInvHeaderDetailLine.Measure = row.Cells("Measure").Value
                        objClsInvHeaderDetailLine.PriceCategory = If((row.Cells("PriceCat").Value = ""), "T", row.Cells("PriceCat").Value)
                    End If

                    objClsInvHeaderDetailLine.OrgPrice = CDbl(row.Cells("OrgPrice").Value)
                    objClsInvHeaderDetailLine.Qty_Suspended = 0
                    objClsInvHeaderDetailLine.Motif = IIf((row.Cells("Motif").Value) = True, 1, 0)
                    objClsInvHeaderDetailLine.IsPriceItem = IIf((row.Cells("IsPriceItem").Value) = True, 1, 0)

                    objClsInvHeaderDetailLine.ProcessedID = GlassInvLineProductionState.UnProcessed

                    objClsInvHeaderDetailLine.fUnitCost = row.Cells("Price").Text
                    objClsInvHeaderDetailLine.fOriginal_Price = row.Cells("Price").Text
                    objClsInvHeaderDetailLine.fDiscount_Amount = row.Cells("DiscAmt").Value
                    objClsInvHeaderDetailLine.fItem_Gross = row.Cells("ItmExcAmount").Value

                    objClsInvHeaderDetailLine.fService_Net = row.Cells("ServiceItemTotNet").Value
                    objClsInvHeaderDetailLine.fService_tax = row.Cells("ServiceItemTax").Value
                    objClsInvHeaderDetailLine.fService_Gross = row.Cells("ServiceGross").Value

                    '<<<<<<<Just in case comment line after this>>>>>>>>
                    objClsInvHeaderDetailLine.OrgPrice = row.Cells("Original_Price").Value


                    objClsInvHeaderDetailLine.Foreign_InvTotIncl = row.Cells("OrgPrice").Value

                    objClsInvHeaderDetailLine.fItem_tax = row.Cells("Tax").Value
                    objClsInvHeaderDetailLine.iTaxTypeID = row.Cells("TaxRate").Value
                    objClsInvHeaderDetailLine.fTaxRate = row.Cells("TaxRateValue").Value
                    objClsInvHeaderDetailLine.fItem_Net = row.Cells("Net").Value
                    objClsInvHeaderDetailLine.fTotal_Amt = row.Cells("Amount").Text
                    '----Ending finance data----

                    objClsInvHeaderDetailLine.LineNotes = row.Cells("LineNotes").Text
                    objClsInvHeaderDetailLine.Comment2 = row.Cells("MarkAs").Text
                    objClsInvHeaderDetailLine.LineType = row.Cells("QuoteFiedType").Value
                    objClsInvHeaderDetailLine.LineComments = row.Cells("LineComments").Text
                    objClsInvHeaderDetailLine.IsPriceItem = row.Cells("IsPriceItem").Value
                    objClsInvHeaderDetailLine.ShapeFileName = IIf(IsDBNull(row.Cells("Shape").Value), "", row.Cells("Shape").Value)

                    objClsInvHeader.AddInvDetailLines(objClsInvHeaderDetailLine)

                Next

                If isPrintPreview = True Then
                    recState = objClsInvHeader.UpdateInvDetailLines2()
                Else
                    recState = objClsInvHeader.UpdateInvDetailLines()
                End If

                If recState = -1 Then
                    objClsInvHeader.Rollback_Trans()
                    Exit Function
                Else

                    If oSOModuleDefaults.UseShapes = True And oFacDefaults.OptimizeApp = GlassOptimizeApp.PerfectCut And isPrintPreview = False Then
                        SQL = "DELETE FROM spilInvNumLines_ShapeDetails WHERE OrderIndex = " & orderIndex & " "
                        If objClsInvHeader.Execute_Sql_Trans(SQL) = 0 Then
                            MsgBox("Error in Perfect Cut Shapes", MsgBoxStyle.Critical, "SPIL Glass")
                            objClsInvHeader.Rollback_Trans()
                            Exit Function
                        End If
                    End If

                    Dim collspPara As New Collection
                    Dim colPara As New spParameters
                    Dim newSQLQuery As String = ""

                    If isPrintPreview = False Then
                        colPara.ParaName = "@QuoteStateID"
                        If isACopy = True Then
                            colPara.ParaValue = QuoteStateValue.EditMode
                        Else
                            colPara.ParaValue = utxtQuoteState.Value
                        End If

                        collspPara.Add(colPara)

                        colPara.ParaName = "@DefaultTaxRateForQuote"
                        colPara.ParaValue = defaultTaxRateForQuote
                        collspPara.Add(colPara)


                        colPara.ParaName = "@DefaultTaxRateValueForQuote"
                        colPara.ParaValue = defaultTaxRateValueForQuote
                        collspPara.Add(colPara)

                        newSQLQuery = "UPDATE spilInvNum SET QuoteStateID = @QuoteStateID, DefaultTaxRateForQuote = @DefaultTaxRateForQuote, DefaultTaxRateValueForQuote = @DefaultTaxRateValueForQuote WHERE OrderIndex = '" & quoteOrdeIndex & "'"

                        If objClsInvHeader.EXE_SQL_Trans_Para_Return(newSQLQuery, collspPara) = 0 Then
                            modGlazingQuoteExtension.GQShowMessage("Erro in item Quote State ID", Me.Text, MsgBoxStyle.Critical)
                            objClsInvHeader.Rollback_Trans()
                            Exit Function
                        End If

                        collspPara.Clear()

                        'Job ID Update
                        Dim collspParaJ As New Collection
                        Dim colParaJ As New spParameters

                        colParaJ.ParaName = "@JobID"
                        colParaJ.ParaValue = _JobId

                        collspParaJ.Add(colParaJ)

                        newSQLQuery = "UPDATE spilInvNum SET JobID = @JobID WHERE OrderIndex = '" & quoteOrdeIndex & "'"

                        If objClsInvHeader.EXE_SQL_Trans_Para_Return(newSQLQuery, collspParaJ) = 0 Then
                            modGlazingQuoteExtension.GQShowMessage("Error in item Quote StageID & JobID", Me.Text, MsgBoxStyle.Critical)
                            objClsInvHeader.Rollback_Trans()
                            Exit Function
                        End If

                        collspParaJ.Clear()

                        'Stage ID Update
                        Dim collspParaS As New Collection
                        Dim colParaS As New spParameters

                        colParaS.ParaName = "@StageID"
                        colParaS.ParaValue = _StageId

                        collspParaS.Add(colParaS)

                        newSQLQuery = "UPDATE spilInvNum SET StageID = @StageID WHERE OrderIndex = '" & quoteOrdeIndex & "'"

                        If objClsInvHeader.EXE_SQL_Trans_Para_Return(newSQLQuery, collspParaS) = 0 Then
                            modGlazingQuoteExtension.GQShowMessage("Error in item Quote StageID & JobID", Me.Text, MsgBoxStyle.Critical)
                            objClsInvHeader.Rollback_Trans()
                            Exit Function
                        End If

                        collspParaS.Clear()

                        'Total Amount Update
                        Dim collspParaI As New Collection
                        Dim colParaI As New spParameters

                        colParaI.ParaName = "@InvTotIncl"
                        colParaI.ParaValue = Convert.ToDecimal(lblTotExcAmo.Text)

                        collspParaI.Add(colParaI)

                        newSQLQuery = "UPDATE spilInvNum SET InvTotIncl = @InvTotIncl WHERE OrderIndex = '" & quoteOrdeIndex & "'"

                        If objClsInvHeader.EXE_SQL_Trans_Para_Return(newSQLQuery, collspParaI) = 0 Then
                            modGlazingQuoteExtension.GQShowMessage("Error in Total Amount", Me.Text, MsgBoxStyle.Critical)
                            objClsInvHeader.Rollback_Trans()
                            Exit Function
                        End If

                        collspParaI.Clear()

                    End If

                    'Picture
                    Dim imageArray = row.Cells("ItemImageByteArray").Value

                    For Each row In enumerator
                        If IsDBNull(row.Cells("ItemImage").Value) = False And IsDBNull(row.Cells("ItemImageByteArray").Value) = False Then
                            imageArray = row.Cells("ItemImageByteArray").Value

                        Else
                            imageArray = emptyArray

                        End If

                        colPara.ParaName = "@ItemImage"
                        colPara.ParaValue = imageArray
                        collspPara.Add(colPara)

                        colPara.ParaName = "@isImageAttached"
                        colPara.ParaValue = row.Cells("isImageAttached").Value
                        collspPara.Add(colPara)

                        colPara.ParaName = "@isShapeAttached"
                        colPara.ParaValue = row.Cells("isShapeAttached").Value
                        collspPara.Add(colPara)

                        colPara.ParaName = "@templateData"
                        If IsNothing(row.Cells("templateData").Value) = True Then
                            colPara.ParaValue = ""
                        Else
                            colPara.ParaValue = row.Cells("templateData").Value
                        End If

                        collspPara.Add(colPara)

                        If isPrintPreview = True Then
                            newSQLQuery = "UPDATE spilInvNumLines2 SET ItemImage = @ItemImage, isImageAttached = @isImageAttached, isShapeAttached = @isShapeAttached, templateData = @templateData WHERE OrderIndex = '" & quoteOrdeIndex & "' AND idInvoiceLines = '" & row.Index & "'"
                        Else
                            newSQLQuery = "UPDATE spilInvNumLines SET ItemImage = @ItemImage, isImageAttached = @isImageAttached, isShapeAttached = @isShapeAttached, templateData = @templateData WHERE OrderIndex = '" & quoteOrdeIndex & "' AND idInvoiceLines = '" & row.Index & "'"
                        End If

                        If objClsInvHeader.EXE_SQL_Trans_Para_Return(newSQLQuery, collspPara) = 0 Then
                            modGlazingQuoteExtension.GQShowMessage("Erro in item pictures", Me.Text, MsgBoxStyle.Critical)
                            objClsInvHeader.Rollback_Trans()
                            Exit Function
                        End If
                        collspPara.Clear()


                        If row.Cells("ShapeDetails").Value.GetType() Is GetType(PCShapeDetails) And isPrintPreview = False Then
                            Dim Shape As PCShapeDetails = CType(row.Cells("ShapeDetails").Value, PCShapeDetails)
                            If Not IsNothing(Shape) Then
                                For Each oLine As clsInvDetailLine In objClsInvHeader.collDetailInvLines
                                    If oLine.idInvoiceLines = row.Index Then
                                        Shape.iInvDetailID = oLine.iInvDetailID
                                        Exit For
                                    End If
                                Next
                                Shape.OrderIndex = orderIndex
                                colPara.ParaName = "@OrderIndex"
                                colPara.ParaValue = Shape.OrderIndex
                                collspPara.Add(colPara)
                                colPara.ParaName = "@iInvDetailID"
                                colPara.ParaValue = Shape.iInvDetailID
                                collspPara.Add(colPara)
                                colPara.ParaName = "@ShapeXML"
                                colPara.ParaValue = Shape.ShapeXML
                                collspPara.Add(colPara)
                                colPara.ParaName = "@ShapeSAX"
                                colPara.ParaValue = Shape.ShapeSAX
                                collspPara.Add(colPara)
                                colPara.ParaName = "@ShapePNG"
                                colPara.ParaValue = Shape.ShapePNG
                                collspPara.Add(colPara)
                                colPara.ParaName = "@ShapeDimensions"
                                colPara.ParaValue = Shape.ShapeDimensions
                                collspPara.Add(colPara)
                                colPara.ParaName = "@ShapeSizes"
                                colPara.ParaValue = Shape.ShapeSizes
                                collspPara.Add(colPara)
                                colPara.ParaName = "@ShapeName"
                                colPara.ParaValue = Shape.ShapeName
                                collspPara.Add(colPara)
                                SQL = "INSERT INTO spilInvNumLines_ShapeDetails (OrderIndex,iInvDetailID,ShapeXML,ShapeSAX,ShapePNG,ShapeDimensions,ShapeSizes,ShapeName) VALUES(@OrderIndex,@iInvDetailID,@ShapeXML,@ShapeSAX,@ShapePNG,@ShapeDimensions,@ShapeSizes,@ShapeName)"
                                If objClsInvHeader.EXE_SQL_Trans_Para_Return(SQL, collspPara) = 0 Then
                                    sError = "Error in Perfect Cut Shapes "
                                    objClsInvHeader.Rollback_Trans()
                                    Exit Function
                                End If
                                collspPara.Clear()
                            End If
                        End If
                    Next
                    'Update Shape table with Line and Order ID
                End If
            End If

            If isPrintPreview = True Then
                SQL = "DELETE FROM spilInvDocs WHERE OrderIndex = " & orderIndex & " "
                If objClsInvHeader.Execute_Sql_Trans(SQL) = 0 Then
                    MsgBox("Error in Attched Documents", MsgBoxStyle.Critical, "SPIL Glass")
                    objClsInvHeader.Rollback_Trans()
                    Exit Function
                End If

                Dim IsDocFound As Boolean = False
                For Each mdr As UltraGridRow In UGDocs.Rows
                    SQL = " set dateformat dmy INSERT INTO spilInvDocs " & _
                            "(OrderIndex, Line_No,  Path   ) VALUES (" & orderIndex & "," & mdr.Index + 1 & ", '" & IIf(IsDBNull(mdr.Cells(1).Value) = True, String.Empty, mdr.Cells(1).Value) & "' )"
                    If objClsInvHeader.Execute_Sql_Trans(SQL) = 0 Then
                        MsgBox("Error in Attched Documents", MsgBoxStyle.Critical, "SPIL Glass")
                        objClsInvHeader.Rollback_Trans()
                        Exit Function
                    End If
                Next

                Dim newGlazingNotification As New clsGlazingQuoteNotification()
                If newGlazingNotification.GetNotificationDetails(objClsInvHeader, isExistingOrder, utxtQuoteJobID.Text) = 0 Then
                    objClsInvHeader.Rollback_Trans()
                    Exit Function
                End If

                If IsNothing(utxtQuoteState.Text) = False And IsNothing(utxtQuoteState.Value) = False Then
                    If isExistingOrder = False Then
                        clsGlazingQuoteExtensionObj.GQDocumentLog(orderIndex, utxtQuoteState.Text, objClsInvHeader, "Quotation created")

                    ElseIf utxtQuoteState.Value <> QuoteStateValue.EditMode Then

                        If utxtQuoteState.Value = QuoteStateValue.Copy Then
                            If clsGlazingQuoteExtensionObj.GQDocumentLog(orderIndex, utxtQuoteState.Text, objClsInvHeader, "this is a copy of " & preQuoteNumber & " quotation") = 0 Then
                                Exit Function
                            End If

                        Else
                            clsGlazingQuoteExtensionObj.GQDocumentLog(orderIndex, utxtQuoteState.Text, objClsInvHeader)

                        End If
                    ElseIf isExistingOrder = False Then
                        If clsGlazingQuoteExtensionObj.GQDocumentLog(orderIndex, utxtQuoteState.Text, objClsInvHeader, "") = 0 Then
                            Exit Function
                        End If
                        If utxtQuoteState.Value = QuoteStateValue.Cancelled Then

                        End If
                    End If
                End If
                clsGQExtensionForJobCostingObj.SaveJobProjectInfo(objClsInvHeader)
            End If
            objClsInvHeader.Commit_Trans()

            If isPrintPreview = False Then
                isACopy = False
                isSaved = True

                If MsgBox("Do you want to print the quotation : " & objClsInvHeader.OrderNum.ToString & " ?", MsgBoxStyle.YesNo + MessageBoxIcon.Question, "Quotation") = MsgBoxResult.Yes Then
                    LoadPrint(objClsInvHeader, orderIndex)
                End If
                If isClosing = fasle Then
                    If MsgBox("Do you want to exit now ?", MessageBoxButtons.YesNo + MessageBoxIcon.Question, "Confirmation") = MsgBoxResult.Yes Then

                        Me.Dispose()
                        Exit Function
                    Else
                        LoadExstingQuote()

                    End If
                End If
            End If
            objClsInvHeader = Nothing
            Return 1
        Catch ex As Exception
            modGlazingQuoteExtension.GQShowMessage(ex.Message, Me.Text, MsgBoxStyle.Critical)
            lblOrderNo.Text = ""
            quoteOrdeIndex = 0
            isSaved = False

        Finally

        End Try

    End Function

    Function SaveGlzQuoteJobDetails(objSQL As clsInvHeader, ByRef quoteOrdeIndex As Integer) As Integer

        Dim collspPara As New Collection
        Dim colPara As New spParameters
        Dim newSQLQuery As String
        Dim glzQuoteJobID As String

        If IsNothing(utxtQuoteJobID.Text) = False Then
            If utxtQuoteJobID.Text <> "" Then
                glzQuoteJobID = utxtQuoteJobID.Text
            End If
        End If

        Try

            If IsNothing(glzQuoteJobID) Then

                'If SaveNotificationDetails(objSQL) = 0 Then
                '    Me.modGlazingQuoteExtensionClass.GQShowMessage("Data not saved", Me.Text, MsgBoxStyle.Critical)
                '    Return 0
                '    Exit Function
                'End If
                glzQuoteJobID = clsGlazingQuoteExtensionObj.GetNextJobumentNumber(objSQL, "JobQuote")
                utxtQuoteJobID.Text = glzQuoteJobID

                newSQLQuery = "INSERT INTO GlzQuote_Job_Details (GlzQuoteJobID, GlzQuoteJobName, GlzQuoteJobDes, OrderIndex) " & _
                "VALUES (@GlzQuoteJobID, @GlzQuoteJobName, @GlzQuoteJobDes, @OrderIndex)"

            Else
                newSQLQuery = "UPDATE GlzQuote_Job_Details SET GlzQuoteJobName = @GlzQuoteJobName, GlzQuoteJobDes = @GlzQuoteJobDes WHERE OrderIndex  = @OrderIndex AND GlzQuoteJobID = @GlzQuoteJobID"

            End If

            colPara.ParaName = "@OrderIndex"
            colPara.ParaValue = quoteOrdeIndex
            collspPara.Add(colPara)

            colPara.ParaName = "@GlzQuoteJobID"
            colPara.ParaValue = glzQuoteJobID
            collspPara.Add(colPara)

            colPara.ParaName = "@GlzQuoteJobName"
            colPara.ParaValue = utxtQuoteJobName.Text
            collspPara.Add(colPara)

            colPara.ParaName = "@GlzQuoteJobDes"

            If IsNothing(jobDescription) = True Then
                jobDescription = " "

            End If

            colPara.ParaValue = jobDescription
            collspPara.Add(colPara)


            If objClsInvHeader.EXE_SQL_Trans_Para_Return(newSQLQuery, collspPara) = 0 Then
                modGlazingQuoteExtension.GQShowMessage("Data not saved", Me.Text, MsgBoxStyle.Critical)
                Return 0
                Exit Function

            Else
                Return 1

            End If


        Catch ex As Exception
            modGlazingQuoteExtension.GQShowMessage(ex.Message, Me.Text, MsgBoxStyle.Critical)
            glzQuoteJobID = Nothing
            utxtQuoteJobID.Text = ""
            lblOrderNo.Text = ""

        Finally
            collspPara.Clear()
        End Try

    End Function

    Function SaveNotificationDetails(ByRef objSQL As clsInvHeader) As Integer


        Dim objIncident As New clsIncident
        Dim iIncident As Integer = 0
        Dim expDate As Date = txtDueDate.Value

        With objIncident
            .Outline = utxtQuoteJobName.Text

            .CurrentAgent = AgentID
            .Customer = cmbAccount.Value
            .IncidentActionID = 1
            .IncidentCreated = Date.Now
            .IncidentDueDate = expDate.AddDays(-2).Date
            .IncidentLastModified = Date.Now
            .IncidentType = 1
            .IncidentTypeGroup = 1
            .IsRequireAck = 1
            .isSendRejectionNotification = 0
            .OurRef = .GetIncidentNo()
            .YourRef = objSQL.OrderNum

            'Update Incident Master
            iIncident = .Add_Incident
            If iIncident = -1 Then
                Return 0
                Exit Function
            End If

            'Update Incident Log 
            Dim iIncidentLog As Integer = 0
            .IncidentID = iIncident
            .IncidentActionDate = expDate.AddDays(-2).Date
            .IncidentActionID = 1
            .Revolution = .Outline
            .CurrentAgent = AgentID
            .Proxy = 0
            .NewAgentID = AgentID
            iIncidentLog = .Add_Incident_Log
            If iIncidentLog = -1 Then
                Return 0
                Exit Function
            End If

            'SQL = "INSERT INTO _rtblNotify(dNotifyDate, iForAgentID, iIncidentID, iIncidentLogID," & _
            '" bRead) VALUES ('" & Format(Date.Now, "MM/dd/yyyy") & "'," & mDr("AgentID") & "," & iIncident & "," & iIncidentLog & ",'0')"
            Dim collspPara As New Collection
            Dim colPara As New spParameters
            SQL = "INSERT INTO _rtblNotify(dNotifyDate, iForAgentID, iIncidentID, iIncidentLogID,bRead) VALUES (@dNotifyDate,@iForAgentID,@iIncidentID,@iIncidentLogID,@bRead)"
            colPara.ParaName = "@dNotifyDate"
            colPara.ParaValue = expDate.AddDays(-2).Date
            collspPara.Add(colPara)
            colPara.ParaName = "@iForAgentID"
            colPara.ParaValue = AgentID
            collspPara.Add(colPara)
            colPara.ParaName = "@iIncidentID"
            colPara.ParaValue = iIncident
            collspPara.Add(colPara)
            colPara.ParaName = "@iIncidentLogID"
            colPara.ParaValue = iIncidentLog
            collspPara.Add(colPara)
            colPara.ParaName = "@bRead"
            colPara.ParaValue = 0
            collspPara.Add(colPara)
            If .EXE_SQL_Trans_Para_Return(SQL, collspPara) = 0 Then
                Return 0
                Exit Function
            End If

        End With
        Return iIncident

        'Dim collspPara As New Collection
        'Dim colPara As New spParameters
        'Dim newSQLQuery As String
        'Dim glzQuoteJobID As String
        'Dim incidentsID As Integer
        'Dim incidentLogID As Integer

        'Try

        '    newSQLQuery = " INSERT INTO _rtblIncidents (cYourRef, iDebtorID, cOutline, dCreated, dLastModified) " & _
        '    "VALUES (@cYourRef, @iDebtorID, @cOutline, @dCreated, @dLastModified)"

        '    colPara.ParaName = "@cYourRef"
        '    colPara.ParaValue = objSQL.OrderNum
        '    collspPara.Add(colPara)

        '    colPara.ParaName = "@iDebtorID"
        '    colPara.ParaValue = objSQL.AccountID
        '    collspPara.Add(colPara)

        '    colPara.ParaName = "@cOutline"
        '    colPara.ParaValue = utxtQuoteJobName.Text
        '    collspPara.Add(colPara)

        '    colPara.ParaName = "@dCreated"
        '    colPara.ParaValue = Today.Date
        '    collspPara.Add(colPara)

        '    colPara.ParaName = "@dLastModified"
        '    colPara.ParaValue = Today.Date
        '    collspPara.Add(colPara)



        '    incidentsID = objClsInvHeader.EXE_SQL_Para_Return_ID(newSQLQuery, collspPara)



        '    If idIncidents = 0 Then
        '        Me.modGlazingQuoteExtensionClass.GQShowMessage("Data not saved", Me.Text, MsgBoxStyle.Critical)
        '        Return 0
        '        Exit Function

        '    Else

        '        newSQLQuery = "set dateformat INSERT INTO _rtblIncidentLog (iIncidentID, cResolution, dActionDate) " & _
        '        "VALUES (@iIncidentID, @cResolution, @dActionDate)"

        '        colPara.ParaName = "@iIncidentID"
        '        colPara.ParaValue = idIncidents
        '        collspPara.Add(colPara)

        '        colPara.ParaName = "@cResolution"
        '        colPara.ParaValue = utxtQuoteJobName.Text
        '        collspPara.Add(colPara)

        '        colPara.ParaName = "@dActionDate"
        '        colPara.ParaValue = objSQL.DueDate
        '        collspPara.Add(colPara)

        '        incidentLogID = objClsInvHeader.EXE_SQL_Para_Return_ID(newSQLQuery, collspPara)

        '        If incidentLogID = 0 Then
        '            Me.modGlazingQuoteExtensionClass.GQShowMessage("Data not saved", Me.Text, MsgBoxStyle.Critical)
        '            Return 0
        '            Exit Function
        '        Else

        '            newSQLQuery = "set dateformat INSERT INTO _rtblNotify (iIncidentID, iIncidentLogID, iForAgentID, dNotifyDate) " & _
        '            "VALUES (@iIncidentID, @cResolution, @dActionDate)"

        '            colPara.ParaName = "@iIncidentID"
        '            colPara.ParaValue = incidentID
        '            collspPara.Add(colPara)

        '            colPara.ParaName = "@iIncidentLogID"
        '            colPara.ParaValue = incidentLogID
        '            collspPara.Add(colPara)

        '            colPara.ParaName = "@iForAgentID"
        '            colPara.ParaValue = AgentID
        '            collspPara.Add(colPara)

        '            colPara.ParaName = "@dNotifyDate"
        '            colPara.ParaValue = objSQL.DueDate
        '            collspPara.Add(colPara)

        '            incidentLogID = objClsInvHeader.EXE_SQL_Para_Return_ID(newSQLQuery, collspPara)

        '            If idIncidents = 0 Then
        '                Me.modGlazingQuoteExtensionClass.GQShowMessage("Data not saved", Me.Text, MsgBoxStyle.Critical)
        '                Return 0
        '                Exit Function
        '            Else


        '            End If

        '        End If

        '       Return incidentsID

        '    End If


        'Catch ex As Exception
        '    modGlazingQuoteExtensionClass.GQShowMessage(ex.Message, Me.Text, MsgBoxStyle.Critical)
        '    glzQuoteJobID = Nothing
        '    utxtQuoteJobID.Text = ""
        '    lblOrderNo.Text = ""

        'Finally
        '    collspPara.Clear()
        'End Try

    End Function

    Sub SaveExtraItemLineData(ByRef quoteOrdeIndex As Integer, ByRef enumerator As IEnumerable)

    End Sub

    Sub LoadPrint(exClsInvHeader As clsInvHeader, ByRef InvHeaderID As Integer)
        Dim objSQL As clsInvHeader

        Try
            If IsNothing(exClsInvHeader) = False Then
                objSQL = exClsInvHeader
            Else
                objSQL = New clsInvHeader
            End If

            If isExistingOrder = True Then
                objSQL.OrderNum = lblOrderNo.Text
                objSQL.ExtOrderNum = txtCustOrdNo.Text
                'objSQL.ExtOrderNum = utxtQuoteJobID.Text
            End If

            frmDocumentPrint_Email.pubMeDocID = quoteOrdeIndex
            frmDocumentPrint_Email.pubMeSpilDocTypeID = GlassDocTypes.Quotation
            frmDocumentPrint_Email.pubMeSpilDocStatusID = GlassDocState.GlazingQuote
            'frmDocumentPrint_Email.pubMeBranchID = cmbFacility.Value
            frmDocumentPrint_Email.pubAccountID = cmbAccount.ActiveRow.Cells("DCLink").Value
            frmDocumentPrint_Email.pubMeDocNumber = objSQL.OrderNum
            frmDocumentPrint_Email.pubSalesRepID = quoteDocRepID
            frmDocumentPrint_Email.pubSalesRepID = cmbSalesRep.Value
            frmDocumentPrint_Email.pubFollowupID = 0
            frmDocumentPrint_Email.pubDocRefNo = objSQL.ExtOrderNum
            If frmDocumentPrint_Email.LoadEmailPara(AgentID) = 0 Then
                frmDocumentPrint_Email.cmdEmail.Enabled = False
                frmDocumentPrint_Email.chkEmail.Enabled = False
            End If
            ' frmDocumentPrint_Email.quoteAttachmentsDataset = UGDocs.DataSource
            frmDocumentPrint_Email.SetDocumentControlProperties()
            frmDocumentPrint_Email.LoadGridData()
            frmDocumentPrint_Email.FillEmailData()
            frmDocumentPrint_Email.StartPosition = FormStartPosition.CenterScreen
            frmDocumentPrint_Email.WindowState = FormWindowState.Normal
            frmDocumentPrint_Email.ShowDialog()

        Catch ex As Exception
            GQShowMessage(ex.Message, Me.Text, MsgBoxStyle.Critical)

        End Try

    End Sub

    'Public Function UpdateQuoteData(SQLNew As clsSqlConn) As Integer
    '    Try
    '        Dim sqlQuary As String = ""
    '        Dim band As UltraGridBand = Me.UG2.DisplayLayout.Bands(0)
    '        Dim row As UltraGridRow
    '        Dim DS = New DataSet
    '        Dim bError As Boolean = False
    '        sqlQuary = "INSERT INTO spilInvNum (AccountID, DocType, DocState, InvoiceNotes) VALUES (" & cmbAccount.ActiveRow.Cells("DCLink").Value & ", 5, 12, '" & utxtNoteText.Text & "') SELECT SCOPE_IDENTITY()"

    '        DS = SQLNew.Exe_Query_Trans(sqlQuary)
    '        For Each row In band.GetRowEnumerator(GridRowType.DataRow)
    '            sqlQuary = "INSERT INTO spilInvNumLines (OrderIndex, LineComments, iHeight, iWidth, fQuantity, fItem_Net, fItem_Gross, Image) VALUES (" & OrderIndex & ", " & row.Cells("LineComments").Value & ", " & row.Cells("Height").Value & ", " & row.Cells("Width").Value & ", " & row.Cells("Qty").Value & ",  " & row.Cells("Price").Value & ", " & row.Cells("ExclPrice").Value & ",  " & row.Cells("ItemImage").Value & ")"
    '            If SQLNew.Exe_Query_Trans(sqlQuary) = 0 Then
    '                bError = True
    '            End If
    '        Next row
    '        If bError Then
    '            Return 0
    '        Else
    '            Return 1
    '        End If
    '    Catch ex As Exception
    '        Return 0
    '    Finally
    '        oSQLQuery = ""
    '    End Try
    'End Function

    Private Sub btnSaveDocDes_Click(sender As Object, e As EventArgs) Handles btnSaveDocDes.Click
        Try
            utxtDocDecHEaderMain.Tabs("Documents").Visible = True
            utxtDocDecHEaderMain.Tabs("docDescription").Visible = False
            Me.UG2.ActiveCell = UG2.Rows(UG2ActiveRow.Index).Cells("LineComments")
            Me.UG2.PerformAction(Infragistics.Win.UltraWinGrid.UltraGridAction.EnterEditMode, False, False)
            UG2.Rows(UG2ActiveRow.Index).Cells("LineComments").Value = utxtDocDec.Text
            utxtDocDec.Text = ""

        Catch ex As Exception
            modGlazingQuoteExtension.GQShowMessage(ex.Message, Me.Text, MsgBoxStyle.Critical)

        End Try
    End Sub

    Private Sub utxtDocDec_KeyDown(sender As Object, e As KeyEventArgs)
        utxtDocDec.Select()
    End Sub

    Private Sub utxtDocDec_PreviewKeyDown(sender As Object, e As PreviewKeyDownEventArgs) Handles utxtDocDec.PreviewKeyDown, UltraFormattedTextEditor3.PreviewKeyDown, utxtNoteText.PreviewKeyDown
        If e.KeyCode = Keys.Tab Then
            utxtDocDec.Text = utxtDocDec.Text + ControlChars.Tab
            utxtDocDec.Select()
        End If
    End Sub

    Private Sub TextBox1_Enter(sender As Object, e As EventArgs) Handles TextBox1.Enter, TextBox2.Enter
        utxtDocDec.Select()
    End Sub

    Sub AddNewRowsValidator(ByVal rowPostion As String)
        Try
            If IsNothing(Me.UG2.ActiveRow) = False Then
                If IsDBNull(Me.UG2.ActiveRow.Cells("QuoteFiedType").Value) = False Then
                    If IsNothing(Me.UG2.ActiveCell) = False Then
                        If Me.UG2.ActiveRow.Cells("QuoteFiedType").Value <> "" Then
                            If Me.UG2.ActiveRow.Cells("QuoteFiedType").Value = QuateFiedTypesList.Text Or Me.UG2.ActiveRow.Cells("QuoteFiedType").Value = QuateFiedTypesList.Stock_Item Then
                                If UG2.ActiveCell.Column.Key = "MarkAs" Then
                                    AddNewRow(rowPostion)
                                End If
                            ElseIf Me.UG2.ActiveRow.Cells("QuoteFiedType").Value = QuateFiedTypesList.Subtotal Then
                                If UG2.ActiveCell.Column.Key = "Amount" Then
                                    AddNewRow(rowPostion)
                                End If
                            Else
                                If UG2.ActiveCell.Column.Key = "LineComments" Then
                                    AddNewRow(rowPostion)
                                End If
                            End If
                        End If
                    End If
                ElseIf UG2.Rows(UG2.ActiveRow.Index).Selected = True Then
                    AddNewRow(rowPostion)
                End If
            End If

        Catch ex As Exception
            modGlazingQuoteExtension.GQShowMessage(ex.Message, Me.Text, MsgBoxStyle.Critical)

        End Try
    End Sub

    Public Function AddNewRow(ByVal rowPostion As String) As UltraGridRow
        Try
            Dim activeRowIndex As Integer = 0
            'Dim ty1 = Me.UG2.Rows(Me.UG2.ActiveRow.Index).Cells("ItmGroupID").Value
            ' Dim ty2 = Me.UG2.Rows(Me.UG2.ActiveRow.Index).Cells("LineComments").Value

            If IsNothing(Me.UG2.ActiveRow) = False Then
                If rowPostion = "before" Then
                    activeRowIndex = Me.UG2.ActiveRow.Index

                ElseIf rowPostion = "after" Then
                    activeRowIndex = Me.UG2.ActiveRow.Index + 1

                ElseIf rowPostion = "end" Then
                    activeRowIndex = Me.UG2.Rows.Count

                End If
            End If

            Dim row As UltraGridRow = UG2.DisplayLayout.Bands(0).AddNew

            row.ParentCollection.Move(row, activeRowIndex)
            Me.UG2.ActiveRowScrollRegion.ScrollRowIntoView(row)
            Me.UG2.Rows(activeRowIndex).Activate()
            Me.UG2.Rows(activeRowIndex).Cells("QuoteFiedType").Activate()

            'Set new values
            Me.UG2.Rows(activeRowIndex).Cells("TaxRate").Value = defaultTaxtRate

            Me.UG2.Rows(activeRowIndex).Cells("TaxRateValue").Value = defaultTaxtRateValue

            'If activeRowIndex > 0 Then
            '    Me.UG2.Rows(activeRowIndex).Cells("ItmGroupID").Value = Me.UG2.Rows(activeRowIndex - 1).Cells("ItmGroupID").Value
            'Else
            '    Me.UG2.Rows(activeRowIndex).Cells("ItmGroupID").Value = 0
            'End If

            Me.UG2.Rows(activeRowIndex).Cells("QuoteFiedType").DroppedDown = True
            Me.UG2.ActiveRow.Cells("QuoteFiedType").Appearance.BackColor = Nothing
            Me.UG2.ActiveRow.Cells("Shape").ToolTipText = "Click to insert a shape"

            Return row

        Catch ex As Exception
            modGlazingQuoteExtension.GQShowMessage(ex.Message, Me.Text, MsgBoxStyle.Critical)

        End Try
    End Function

    Sub RowDataOverride(ByVal RowIndex As Integer)
        Try
            If IsDBNull(UG2.Rows(RowIndex).Cells("QuoteFiedType").Value) = False Then
                'Pasting
                If UG2.Rows(RowIndex).Cells("QuoteFiedType").Value = QuateFiedTypesList.Header_Sub Then
                    Me.UG2.Rows(activeRowIndex).Cells("ItmGroupID").Value = subHeaderID + 1

                Else
                    If RowIndex > 0 Then
                        Me.UG2.Rows(RowIndex).Cells("ItmGroupID").Value = Me.UG2.Rows(RowIndex - 1).Cells("ItmGroupID").Value
                    Else
                        Me.UG2.Rows(RowIndex).Cells("ItmGroupID").Value = 0
                    End If
                End If
                Me.UG2.Rows(RowIndex).Cells("InvLineID").Value = 0
                Me.UG2.Rows(RowIndex).Cells("iInvDetailID").Value = 0

            End If

        Catch ex As Exception
            modGlazingQuoteExtension.GQShowMessage(ex.Message, Me.Text, MsgBoxStyle.Critical)

        End Try
    End Sub


    Private Sub btnCloseDocDes_Click(sender As Object, e As EventArgs) Handles btnCloseDocDes.Click, Button8.Click
        Me.UG2.ActiveCell = UG2.Rows(UG2ActiveRow.Index).Cells("LineComments")
        Me.UG2.PerformAction(Infragistics.Win.UltraWinGrid.UltraGridAction.EnterEditMode, False, False)
    End Sub

    Private Sub btnFooter_Click(sender As Object, e As EventArgs) Handles btnFooter.Click
        Try
            If fotterIsActive = False Then
                fotterIsActive = True
                btnFooter.FlatAppearance.BorderColor = Color.Yellow
                btnFooter.BackColor = Color.Yellow
                btnFooter.ForeColor = Color.Black
                UG2.Visible = False
                btnFooter.Text = "Save (Press F8 to add predefined text)"
                UG2.Visible = True
            End If

            Dim newGlazingNote As New frmGlazingNote(Me)
            newGlazingNote.utxtNoteText.Value = utxtNoteText.Text
            newGlazingNote.ShowDialog()

            If fotterIsActive = True Then
                fotterIsActive = False
                If utxtNoteText.Text <> "" And utxtNoteText.Text <> Nothing Then
                    btnFooter.Text = "Footer Area - Enabled"
                    btnFooter.FlatAppearance.BorderColor = Color.Green
                    btnFooter.BackColor = Color.Green
                Else
                    btnFooter.Text = " Footer Area - Disabled"
                    btnFooter.FlatAppearance.BorderColor = Color.FromArgb(71, 164, 248)
                    btnFooter.BackColor = Color.FromArgb(71, 164, 248)
                End If
                UG2.Visible = True
            End If

        Catch ex As Exception
            fotterIsActive = False
            utxtNoteText.Text = ""
            btnFooter.Text = " Footer Area - Disabled"
            btnFooter.FlatAppearance.BorderColor = Color.FromArgb(71, 164, 248)
            btnFooter.BackColor = Color.FromArgb(71, 164, 248)
            UG2.Visible = True
            modGlazingQuoteExtension.GQShowMessage(ex.Message, Me.Text, MsgBoxStyle.Critical)

        End Try
    End Sub

    Private Sub utxtNoteText_KeyDown(sender As Object, e As KeyEventArgs) Handles utxtNoteText.KeyDown
        GlazingNote(e)
    End Sub

    Public Sub GlazingNote(e As KeyEventArgs)
        Try
            If e.KeyCode = 119 Then
                Dim newGlazingDocDescription As New frmGlazingDocDescription
                newGlazingDocDescription.DocDesTypeName = "Footer Note"
                newGlazingDocDescription.ShowDialog()

                If utxtNoteText.Text <> "" Then
                    utxtNoteText.Value = utxtNoteText.Text + vbCrLf + selectedDocDes
                ElseIf utxtNoteText.Text = "" Then
                    utxtNoteText.Value = selectedDocDes
                End If
            End If
        Catch ex As Exception
            modGlazingQuoteExtension.GQShowMessage(ex.Message, Me.Text, MsgBoxStyle.Critical)

        End Try
    End Sub
    Sub SetSubTotal(e As CellEventArgs)
        Try
            'If groupStarted = True Then
            Dim rowTypes As GridRowType = GridRowType.DataRow
            Dim band As UltraGridBand = Me.UG2.DisplayLayout.Bands(0)
            Dim enumerator As IEnumerable = band.GetRowEnumerator(rowTypes)
            Dim row As UltraGridRow
            Dim subTotal As Decimal = 0.0
            Dim groupID As Integer = Me.UG2.ActiveRow.Cells("ItmGroupID").Value

            Dim exGroupAmount As Decimal = 0.0

            For Each row In enumerator

                'Loop similer group items
                If IsNothing(row.Cells("ItmGroupID").Value) = False Then
                    If groupID = row.Cells("ItmGroupID").Value Then

                        'IF this is an item
                        If IsNothing(row.Cells("QuoteFiedType").Text) = False Then

                            'If row.Cells("Amount").Text > 0 Then
                            '    Dim a = 0
                            '    a = row.Cells("ItmGroupID").Value
                            'End If

                            If row.Cells("QuoteFiedType").Text = "Text" Or row.Cells("QuoteFiedType").Text = "Stock Item" Then
                                If IsNothing(row.Cells("Amount").Text) = False Then
                                    If row.Cells("ItmGroupID").Value > 0 Then
                                        subTotal = subTotal + row.Cells("Amount").Text
                                        'row.Cells("Net").Value = row.Cells("Amount")
                                        'row.Cells("Tax").Value = (row.Cells("Amount").Text * row.Cells("TaxRate").Text) / 100
                                        'row.Cells("ItmExcAmount").Value = row.Cells("ItmIncAmount").Text + row.Cells("Tax").Text
                                    End If

                                End If

                            ElseIf row.Cells("QuoteFiedType").Text = "Subtotal" Then
                                row.Cells("Amount").Value = Math.Round(subTotal, 2, MidpointRounding.AwayFromZero)

                            End If

                        End If

                    End If
                End If

            Next
            groupStarted = False
            'SetTotalAmounts(True)
        Catch ex As Exception
            modGlazingQuoteExtension.GQShowMessage(ex.Message, Me.Text, MsgBoxStyle.Critical)

        End Try

    End Sub

    Sub SetTotalAmounts(ByVal isAmountChanging As Boolean)
        Try
            Exit Sub
            Dim rowTypes As GridRowType = GridRowType.DataRow
            Dim band As UltraGridBand = Me.UG2.DisplayLayout.Bands(0)
            Dim enumerator As IEnumerable = band.GetRowEnumerator(rowTypes)
            Dim row As UltraGridRow
            Dim TotalExc As Decimal = 0.0
            Dim TotalTax As Decimal = 0.0
            Dim TotalInc As Decimal = 0.0

            For Each row In enumerator
                'Dim a2 = row.Index
                If row.Cells("QuoteFiedType").Value <> "" Then
                    If row.Cells("QuoteFiedType").Value = QuateFiedTypesList.Text Or row.Cells("QuoteFiedType").Value = QuateFiedTypesList.Stock_Item Then

                        TotalExc = TotalExc + Convert.ToDecimal(row.Cells("ItmExcAmount").Text)
                        TotalTax = TotalTax + Convert.ToDecimal(row.Cells("Tax").Text)
                        TotalInc = TotalInc + Convert.ToDecimal(row.Cells("Net").Text)
                    End If
                End If
            Next

            lblTotExcAmo.Text = TotalExc
            lblTotVatAmo.Text = TotalTax
            lblTotIncAmo.Text = TotalInc

        Catch ex As Exception
            modGlazingQuoteExtension.GQShowMessage(ex.Message, Me.Text, MsgBoxStyle.Critical)

        End Try
    End Sub

    Sub ClearRowCellsAfterAmountChange()
        Try
            Me.UG2.ActiveRow.Cells("Qty").Value = 0
            Me.UG2.ActiveRow.Cells("Height").Value = 0
            Me.UG2.ActiveRow.Cells("Width").Value = 0
            Me.UG2.ActiveRow.Cells("Volume").Value = 0
            Me.UG2.ActiveRow.Cells("Price").Value = 0
            Me.UG2.ActiveRow.Cells("DiscAmt").Value = 0
            'Me.UG2.ActiveRow.Cells("Net").Value = 0
            Me.UG2.ActiveRow.Cells("TaxRate").Value = 0
            Me.UG2.ActiveRow.Cells("TaxRateValue").Value = 0
            Me.UG2.ActiveRow.Cells("Tax").Value = 0
            'Me.UG2.ActiveRow.Cells("ItmExcAmount").Value = 0

            'If isTaxedPrice = False Then
            'exc
            Me.UG2.ActiveRow.Cells("ItmExcAmount").Value = Me.UG2.ActiveRow.Cells("Amount").Value

            'ElseIf isTaxedPrice = True Then
            '    'inc
            Me.UG2.ActiveRow.Cells("Net").Value = Me.UG2.ActiveRow.Cells("Amount").Value
            '    QuoteGridBackTaxCal()


            'End If

        Catch ex As Exception
            modGlazingQuoteExtension.GQShowMessage(ex.Message, Me.Text, MsgBoxStyle.Critical)

        End Try
    End Sub

    Sub QuoteGridBackTaxCal()
        'Me.UG2.ActiveRow.Cells("Tax").Value = Me.UG2.ActiveRow.Cells("Amount").Value / Me.UG2.ActiveRow.Cells("TaxRate").Value
        'Me.UG2.ActiveRow.Cells("ItmExcAmount").Value = Me.UG2.ActiveRow.Cells("Tax").Value * Me.UG2.ActiveRow.Cells("Amount").Value

    End Sub

    Public Enum RowColumKey As Integer
        Qty = 1
        Height = 2
        Width = 3
        Volume = 4
        Price = 5
        DiscAmt = 6
        Net = 7
        TaxRate = 8
        TaxRateValue = 9
        Tax = 10
        ItmExcAmount = 11

    End Enum


    Private Sub lineTypeNavigator_TextChanged(sender As Object, e As EventArgs) Handles lineTypeNavigator.TextChanged
        Me.UG2.Focus()
        Me.UG2.ActiveCell = Me.UG2.Rows(lineTypeNavigator.Text).Cells("QuoteFiedType")

    End Sub

    Sub LoadExstingQuote()
        'pubMeSODocument = New clsInvHeader(quoteOrdeIndex)
        txtOrdDate.Value = Today.Date
        txtDueDate.Value = Today.AddDays(2).Date
        txtPostelAdd.Enabled = False
        Try
            If quoteOrdeIndex > 0 Then
                isOpeningQuote = True
                Dim index As Integer = Me.UG2.Rows.Count
                Dim isRowExist As Boolean = False

                'clear the gride
                Do
                    If index > 0 Then
                        Me.UG2.Rows(index - 1).Delete()
                        index -= 1
                    Else
                        Exit Do
                    End If
                Loop Until index = 0

                Dim documentData As clsInvHeader = New clsInvHeader(quoteOrdeIndex)
                oFormCustomer = New clsCustomer(documentData.AccountID)
                Dim test2 = documentData.OrderIndex
                'Stop updating the gride values
                canUpdate = False

                Dim sqlQuary As String = "SELECT * FROM  spilInvNum LEFT JOIN GlzQuote_Job_Details ON spilInvNum.OrderIndex = GlzQuote_Job_Details.OrderIndex WHERE spilInvNum.OrderIndex = " & quoteOrdeIndex

                Dim objSQL As New clsSqlConn
                With objSQL

                    Dim ds As DataSet = .GET_INSERT_UPDATE(sqlQuary)
                    If ds.Tables(0).Rows.Count > 0 Then

                        isExistingOrder = True
                        Dim dr As UltraGridRow

                        For Each objQutDetailline In ds.Tables(0).Rows
                            pubMeSpilDocStatusID = objQutDetailline("DocState")
                            cmbCusType.Value = objQutDetailline("iCusType")
                            defaultTaxRateForQuote = objQutDetailline("DefaultTaxRateForQuote")
                            defaultTaxRateValueForQuote = objQutDetailline("DefaultTaxRateValueForQuote")

                            If cmbCusType.Value = 0 Then
                                cmbAccount.Value = objQutDetailline("AccountID")
                                'cmbAccount.ReadOnly = True
                            Else
                                cmbAccount.Value = objQutDetailline("AccountID")
                                'cmbAccount.ReadOnly = True
                            End If

                            'txtCustOrdNo.Focus()

                            cmbAccount.Value = (objQutDetailline("AccountID"))

                            txtOrdDate.Value = objQutDetailline("OrderDate")                                  ' dr1("OrderDate")
                            txtDelDate.Value = objQutDetailline("DeliveryDate")                                ' dr1("DeliveryDate")
                            txtDueDate.Value = objQutDetailline("DueDate")                                       ' dr1("DueDate")
                            txtInvDate.Value = objQutDetailline("InvDate")                                       ' dr1("InvDate")

                            If isACopy = False Then
                                txtCustOrdNo.Text = objQutDetailline("ExtOrderNum")                            ' dr1("ExtOrderNum")
                                lblOrderNo.Text = objQutDetailline("OrderNum")
                                headder = objQutDetailline("OrderNum")
                            Else
                                preQuoteNumber = objQutDetailline("OrderNum")

                            End If

                            subHeaderID = objQutDetailline("FacilityIDCurrent")
                            cmbFacility.Value = objQutDetailline("FacilityIDCurrent")

                            'defaultTaxtRate = objQutDetailline("TaxRate")
                            'lblInNo.Text = objQutDetailline("SimpleCode").ToString
                            'lblLinkedTo.Text = objQutDetailline("Description1").ToString
                            'lblLinkedTo.Text = objQutDetailline("Price_Type").ToString
                            'lblLinkedTo.Text = objQutDetailline("Volume").ToString
                            'Me.Text = "Quotation - " & objQutDetailline("Price_Type")

                            'Contacts
                            'oFormCustomer = New clsCustomer(objQutDetailline("AccountID"))
                            txtPhy1.Value = objQutDetailline("Address1")
                            txtPhy2.Value = objQutDetailline("Address2")
                            txtPhy3.Value = objQutDetailline("Address3")
                            txtPhy4.Value = objQutDetailline("Address4")
                            txtPhy5.Value = objQutDetailline("Address5")
                            txtPhyPostCode.Value = objQutDetailline("Address6")
                            cmbContPerson.Value = objQutDetailline("ContactName")
                            txtContPerTel.Value = objQutDetailline("ContactTelephone")
                            txtContEmail.Value = objQutDetailline("ContactEmail")
                            cboArea.Value = objQutDetailline("iAreasID")
                            uCmbTerms.Value = objQutDetailline("CreditState")
                            utxtNoteText.Value = objQutDetailline("InvoiceNotes")
                            quoteDocRepID = objQutDetailline("DocRepID")
                            utxtQuoteState.Value = objQutDetailline("QuoteStateID")
                            cmbSalesRep.Value = objQutDetailline("DocRepID")
                            If objQutDetailline("Quoted_Incl") > 0 Then
                                isTaxedPrice = True
                            End If

                            If IsDBNull(objQutDetailline("GlzQuoteJobID")) = False Then
                                If isACopy = False Then
                                    utxtQuoteJobID.Text = objQutDetailline("GlzQuoteJobID")

                                End If

                                utxtQuoteJobName.Text = objQutDetailline("GlzQuoteJobName")
                                jobDescription = objQutDetailline("GlzQuoteJobDes")
                                SetJobDescriptionState()

                            End If


                        Next
                    End If
                    loadLineData()

                    SQL = "SELECT     OrderIndex, Line_No, Path FROM spilInvDocs WHERE OrderIndex = " & quoteOrdeIndex & ""
                    ds = New DataSet()
                    ds = .GET_DATA_SQL(SQL)
                    Dim ugR1 As UltraGridRow
                    For Each mDr As DataRow In ds.Tables(0).Rows
                        ugR1 = UGDocs.DisplayLayout.Bands(0).AddNew
                        ugR1.Cells("LineNo").Value = mDr(1)
                        ugR1.Cells("Path").Value = mDr(2)
                    Next

                End With

                Dim customerID As String = cmbAccount.ActiveRow.Cells("DCLink").Value
                If customerID <> Nothing Or customerID <> "" Then
                    'If Set_Customer(cmbAccount.ActiveRow.Cells("DCLink").Value) = 0 Then
                    '    Exit Sub
                    'End If
                End If
                'sqlQuary = "SELECT *  FROM  Client WHERE OrderIndex = " & cmbAccount.Value & ""
                'objSQL = New clsSqlConn
                'With objSQL
                '    Dim ds As DataSet = .GET_INSERT_UPDATE(sqlQuary)
                '    If ds.Tables(0).Rows.Count > 0 Then
                '        For Each objQutDetailline In ds.Tables(0).Rows
                '            txtPhy1.Text = objQutDetailline("Address1")                                      ' IIf(IsDBNull(objQutDetailline("Address1")) = True, "", objQutDetailline("Address1"))
                '            txtPhy2.Text = objQutDetailline("Address2")                                        'IIf(IsDBNull(objQutDetailline("Address2")) = True, "", objQutDetailline("Address2"))
                '            txtPhy3.Text = objQutDetailline("Address3")                                        'IIf(IsDBNull(objQutDetailline("Address3")) = True, "", objQutDetailline("Address3"))
                '            txtPhy4.Text = objQutDetailline("Address4")                                        'IIf(IsDBNull(objQutDetailline("Address4")) = True, "", objQutDetailline("Address4"))
                '            txtPhy5.Text = objQutDetailline("Address5")                                       'IIf(IsDBNull(objQutDetailline("Address5")) = True, "", objQutDetailline("Address5"))
                '            txtPhyPostCode.Text = objQutDetailline("PhysicalPC")                                       'IIf(IsDBNull(objQutDetailline("Address5")) = True, "", objQutDetailline("Address5"))

                '            cmbContPerson.Text = objQutDetailline("Contact_Person")                                     ' IIf(IsDBNull(objQutDetailline("PostAdd1")) = True, "", objQutDetailline("PostAdd1"))
                '            txtContPerTel.Text = objQutDetailline("Telephone")                                       ' IIf(IsDBNull(objQutDetailline("PostAdd2")) = True, "", objQutDetailline("PostAdd2"))
                '            txtContEmail.Text = objQutDetailline("PostAdd3")                                       ' IIf(IsDBNull(objQutDetailline("PostAdd3")) = True, "", objQutDetailline("PostAdd3"))
                '            cboArea.Text = objQutDetailline("PostAdd4")                                       ' IIf(IsDBNull(objQutDetailline("PostAdd4")) = True, "", objQutDetailline("PostAdd4"))
                '        Next
                '    End If
                'End With

                If isACopy = True Then
                    If objQutDetailline("QuoteStateID") <> QuoteStateValue.Copy Then
                        quoteOrdeIndex = 0
                    End If
                    utxtQuoteState.Value = QuoteStateValue.Copy
                End If

                If utxtQuoteState.Value = QuoteStateValue.Cancelled Then
                    isCancelled = True
                End If

                If isCancelled = True Then
                    tsbPrint.Enabled = False
                    tsbSave.Enabled = False
                    mnuSave.Enabled = False
                Else
                    tsbPrint.Enabled = True
                    tsbSave.Enabled = True
                    mnuSave.Enabled = True
                End If

                Dim pubMeSODocument As New clsInvHeader(quoteOrdeIndex)
                clsGQExtensionForJobCostingObj.SetJobProjectDetails(pubMeSODocument, quoteOrdeIndex)
                isOpeningQuote = False

            Else
                cmbAccount.Focus()
                SetExpireDate()

            End If

        Catch ex As Exception
            modGlazingQuoteExtension.GQShowMessage(ex.Message, Me.Text, MsgBoxStyle.Critical)
            Exit Sub

        Finally
            canUpdate = True
            isClosing = False
            isSaved = False
        End Try
    End Sub

    Public Sub loadLineData()
        Try
            Dim objSQL As New clsSqlConn
            With objSQL
                Dim sqlQuary As String = "SELECT * FROM  spilInvNumLines LEFT JOIN StkItem ON spilInvNumLines.StockLink = StkItem.StockLink  WHERE spilInvNumLines.OrderIndex = " & quoteOrdeIndex & " ORDER BY spilInvNumLines.idInvoiceLines"
                Dim ds As DataSet = .GET_INSERT_UPDATE(sqlQuary)
                If ds.Tables(0).Rows.Count > 0 Then
                    isExistingOrder = True
                    Dim dr As UltraGridRow
                    For Each objQutDetailline In ds.Tables(0).Rows

                        dr = AddNewRow("after")

                        If isACopy = False Then
                            'Set identifer as a exsiting item
                            dr.Cells("IsAExistingItem").Value = True

                            '----Starting Item Line identifers----
                            dr.Cells("iInvDetailID").Value = objQutDetailline("iInvDetailID")

                        End If

                        dr.Cells("QuoteFiedType").Value = objQutDetailline("LineType")
                        dr.Cells("ItmGroupID").Value = objQutDetailline("LN")
                        dr.Cells("InvLineID").Value = objQutDetailline("idInvoiceLines")
                        '----Ending Item Line identifers----

                        '----Starting Stock Items details----    
                        dr.Cells("ItemType").Value = objQutDetailline("ItemType")
                        dr.Cells("StockLink").Value = objQutDetailline("StockLink")
                        dr.Cells("Description1").Value = objQutDetailline("cDescription")
                        dr.Cells("Price_Type").Value = objQutDetailline("PRICE_TYPES_ID")
                        dr.Cells("PriceList").Value = objQutDetailline("PriceList")

                        '----Starting Stock Items details----

                        dr.Cells("Qty").Value = objQutDetailline("fQuantity")
                        dr.Cells("Height").Value = objQutDetailline("iHeight")
                        dr.Cells("Width").Value = objQutDetailline("iWidth")
                        dr.Cells("Volume").Value = objQutDetailline("fVolume")

                        '----Starting finance data----
                        dr.Cells("Price").Value = objQutDetailline("fOriginal_Price")
                        dr.Cells("DiscAmt").Value = objQutDetailline("fDiscount_Amount")
                        dr.Cells("ItmExcAmount").Value = objQutDetailline("fItem_Gross")
                        dr.Cells("Tax").Value = objQutDetailline("fItem_tax")
                        dr.Cells("TaxRate").Value = objQutDetailline("iTaxTypeID")
                        dr.Cells("TaxRateValue").Value = objQutDetailline("fTaxRate")
                        dr.Cells("NET").Value = objQutDetailline("fItem_Net")
                        dr.Cells("amount").Value = objQutDetailline("fTotal_Amt")
                        dr.Cells("Original_Price").Value = objQutDetailline("OrgPrice")
                        dr.Cells("OrgPrice").Value = objQutDetailline("Foreign_InvTotIncl")
                        '----Ending finance data----


                        dr.Cells("LineNotes").Value = objQutDetailline("LineNotes")
                        dr.Cells("MarkAs").Value = objQutDetailline("Comment2")
                        dr.Cells("IsPriceItem").Value = objQutDetailline("IsPriceItem")
                        dr.Cells("LineComments").Value = objQutDetailline("LineComments")
                        dr.Cells("templateData").Value = objQutDetailline("templateData")

                        If IsDBNull(objQutDetailline("ItemImage")) = False Then
                            Dim image As Byte() = objQutDetailline("ItemImage")
                            If image.Length = 64 Then

                            Else
                                dr.Cells("ItemImageByteArray").Value = objQutDetailline("ItemImage")
                                dr.Cells("ItemImage").Value = ByteToImage(objQutDetailline("ItemImage"))
                                Me.UG2.ActiveCell.Row.PerformAutoSize()
                                Me.UG2.PerformAction(Infragistics.Win.UltraWinGrid.UltraGridAction.EnterEditMode, False, False)
                                dr.Cells("isImageAttached").Value = True
                            End If
                        End If

                        dr.Cells("Shape").Value = objQutDetailline("ShapeFileName")
                        If Not String.IsNullOrWhiteSpace(objQutDetailline("ShapeFileName")) Then
                            dr.Cells("Shape").Appearance.BackColor = Color.LightBlue
                            dr.Cells("Shape").ButtonAppearance.Image = Global.SPIL_Glass.My.Resources.Resources.shapesadd
                            dr.Cells("Shape").Value = "Added"
                            ''DisableProcessedLineFeilds(dr)
                            dr.Cells("isShapeAttached").Value = True
                        Else
                            dr.Cells("Shape").Appearance.BackColor = Color.Empty
                            dr.Cells("Shape").ButtonAppearance.Image = Nothing

                        End If

                        dr.Cells("ServiceItemTotNet").Value = objQutDetailline("fService_Net")
                        dr.Cells("ServiceItemTax").Value = objQutDetailline("fService_tax")
                        dr.Cells("ServiceGross").Value = objQutDetailline("fService_Gross")

                        Me.UG2.ActiveRow.PerformAutoSize()
                    Next

                    If oSOModuleDefaults.UseShapes = True And oFacDefaults.OptimizeApp = GlassOptimizeApp.PerfectCut Then
                        SQL = "SELECT spd.ID ,spd.OrderIndex,spd.iInvDetailID,spd.ShapeXML,spd.ShapeSAX,spd.ShapePNG,spd.ShapeDimensions,spd.ShapeSizes,spd.ShapeName FROM spilInvNumLines_ShapeDetails spd INNER JOIN spilInvNumLines  ind ON ind.iInvDetailID=spd.iInvDetailID WHERE spd.OrderIndex = " & quoteOrdeIndex & ""
                        Dim perfectDS As New DataSet()
                        perfectDS = .GET_DATA_SQL(SQL)
                        If Not perfectDS.Tables Is Nothing Then
                            If perfectDS.Tables.Count > 0 Then
                                If perfectDS.Tables(0).Rows.Count > 0 Then
                                    Dim myRowId As UltraGridRow
                                    For Each myRowId In UG2.Rows.GetRowEnumerator(GridRowType.DataRow, Nothing, Nothing)
                                        For Each row As DataRow In perfectDS.Tables(0).Rows
                                            If myRowId.Cells("iInvDetailID").Value = CInt(row("iInvDetailID")) Then
                                                Dim shape As New PCShapeDetails()
                                                shape.ID = row("ID")
                                                shape.UniqueLN = myRowId.Cells("UniqueLN").Value
                                                shape.OrderIndex = row("OrderIndex")
                                                shape.iInvDetailID = row("iInvDetailID")
                                                shape.ShapeXML = row("ShapeXML")
                                                shape.ShapeSAX = row("ShapeSAX")
                                                shape.ShapePNG = row("ShapePNG")
                                                shape.ShapeDimensions = row("ShapeDimensions")
                                                shape.ShapeSizes = row("ShapeSizes")
                                                shape.ShapeName = row("ShapeName")
                                                myRowId.Cells("ShapeDetails").Value = shape
                                                Exit For
                                            End If
                                        Next
                                    Next
                                End If
                            End If
                        End If
                    End If

                    QuoteGridSetSubTotal()
                    'SetTotalAmounts(fasle)
                End If
            End With
        Catch ex As Exception
            modGlazingQuoteExtension.GQShowMessage(ex.Message, Me.Text, MsgBoxStyle.Critical)

        Finally
            SQLNew = Nothing

        End Try
    End Sub

    Private Sub tsmiOptions_Click(sender As Object, e As EventArgs) Handles tsmiOptions.Click
        Dim docDefaultSetting As New frmGlazingDocDefaultSetting(Me)
        docDefaultSetting.ShowDialog()


    End Sub

    Function GetQuoteDefaultData() As DataSet
        Try
            Dim sqlQuary As String = ""
            Dim clsSqlConnObj As New clsSqlConn

            sqlQuary = "SELECT GlzQuote_Defaults.isTaxInc, GlzQuote_Defaults.defaultTaxtRate, GlzQuote_Defaults.AllowBlankCustomerOrderNo, GlzQuote_Defaults.getTaxRateFromCustomer, TaxRate.TaxRate " & _
                  "FROM  GlzQuote_Defaults " & _
                  "INNER JOIN TaxRate ON GlzQuote_Defaults.defaultTaxtRate = TaxRate.idTaxRate " & _
                  "WHERE GlzQuote_Defaults.createdBy= " & AgentID

            sqlQuary += "SELECT * FROM GlzQuote_Defaults WHERE GlzQuote_Defaults.createdBy= " & AgentID

            glzQuoteTax = clsSqlConnObj.GET_INSERT_UPDATE(sqlQuary)

            Return glzQuoteTax
        Catch ex As Exception
            Return Nothing
        End Try
    End Function

    Public Sub GetQuoteTaxState(ByRef taxCheckState As Integer, ByRef taxtRate As Integer)
        Dim glzQuoteTax As DataSet
        Dim newTable As DataTable
        Dim hasValues As Boolean = False
        If IsNothing(taxCheckState) = False Then
            'User default tax rate
            If taxCheckState = -1 Then
                Try
                    glzQuoteTax = GetQuoteDefaultData()
                    If IsNothing(GetQuoteDefaultData) = False Then
                        If glzQuoteTax.Tables(0).Rows.Count > 0 Then
                            newTable = glzQuoteTax.Tables(0)
                            hasValues = True
                        ElseIf glzQuoteTax.Tables(1).Rows.Count > 0 Then
                            newTable = glzQuoteTax.Tables(1)
                            hasValues = True
                        Else
                            hasValues = False
                        End If

                        If hasValues = True Then
                            For Each row In newTable.Rows
                                If IsNothing(row("getTaxRateFromCustomer")) = False Then
                                    getTaxRateFromCustomer = row("getTaxRateFromCustomer")
                                End If
                                If IsDBNull(row("isTaxInc")) = False And glzQuoteTax.Tables(0).Rows.Count > 0 Then
                                    isTaxedPrice = row("isTaxInc")
                                    If row("isTaxInc") = True Then
                                        If IsDBNull(row("defaultTaxtRate")) = False Then
                                            If row("defaultTaxtRate") = True Then
                                                defaultTaxtRate = row("defaultTaxtRate")
                                                defaultTaxtRateValue = row("TaxRate")
                                            End If
                                        End If
                                    Else
                                        defaultTaxtRate = 0
                                        defaultTaxtRateValue = 0
                                    End If
                                End If

                                AllowBlankCustomerOrderNo = row("AllowBlankCustomerOrderNo")
                            Next
                        End If
                    Else
                        modGlazingQuoteExtension.GQShowMessage("Error When retriving quote default data", Me.Text, MsgBoxStyle.Critical)

                    End If
                Catch ex As Exception
                    modGlazingQuoteExtension.GQShowMessage(ex.Message, Me.Text, MsgBoxStyle.Critical)
                End Try
            Else
                'If IsNothing(taxtRate) = False Then
                '    defaultTaxtRate = taxtRate
                'End If

                If taxCheckState = 1 Then
                    isTaxedPrice = True
                    defaultTaxtRate = taxtRate

                Else
                    isTaxedPrice = False
                    defaultTaxtRate = 0

                End If
            End If
        Else
            taxCheckState = -1

        End If

    End Sub

    Sub SetTaxedPriceLable(taxedPrice As Boolean)
        Try
            If IsNothing(taxedPrice) = False Then
                lblTaxType.Visible = True
                If taxedPrice = True Then
                    lblTaxType.Text = "Tax inclusive"
                    priceType = "Inc"
                    Me.UG2.Rows.Band.Columns("Amount").Header.Caption = "Amount (Inc)"

                Else
                    lblTaxType.Text = "Tax exclusive"
                    priceType = "Exc"
                    Me.UG2.Rows.Band.Columns("Amount").Header.Caption = "Amount (Exc)"

                End If
            End If

        Catch ex As Exception
            modGlazingQuoteExtension.GQShowMessage(ex.Message, Me.Text, MsgBoxStyle.Critical)

        End Try
    End Sub

    Sub AfterDefaultTaxePriceChaged(ByVal newTaxtRate As Integer, ByVal newTaxtRateValue As Integer, ByVal newIsTaxedPrice As Boolean, Optional ByRef getTaxFromCustomer As Boolean = False, Optional ByRef expDate As Integer = -1)
        Try
            If expDate > -1 Then

            End If

            If isTaxedPrice <> newIsTaxedPrice Then
                isTaxedPrice = newIsTaxedPrice
                isDefaultChange = True
            End If

            If getTaxFromCustomer = True Then
                newTaxtRate = oFormCustomer.TaxRate
            End If

            If newTaxtRate <> defaultTaxtRate And isTaxedPrice = True Then
                defaultTaxtRate = newTaxtRate
                isDefaultChange = True

            End If

            If isDefaultChange = True Then
                Dim rowTypes As GridRowType = GridRowType.DataRow
                Dim band As UltraGridBand = Me.UG2.DisplayLayout.Bands(0)
                Dim enumerator As IEnumerable = band.GetRowEnumerator(rowTypes)
                Dim row As UltraGridRow
                Dim quoteFiedType As Integer = -1
                For Each row In UG2.DisplayLayout.Bands(0).GetRowEnumerator(rowTypes)
                    If IsDBNull(row.Cells("QuoteFiedType").Value) = False Then
                        If row.Cells("QuoteFiedType").Value = "" Then
                            quoteFiedType = -1
                        Else
                            quoteFiedType = row.Cells("QuoteFiedType").Value
                        End If
                        If (quoteFiedType = QuateFiedTypesList.Text Or quoteFiedType = QuateFiedTypesList.Stock_Item) And row.Cells("IsAExistingItem").Value = False Then
                            If isTaxedPrice = True Then
                                defaultTaxtRate = newTaxtRate
                                defaultTaxtRateValue = newTaxtRateValue
                                row.Cells("TaxRate").Value = newTaxtRate
                                row.Cells("TaxRateValue").Value = newTaxtRateValue
                            Else
                                defaultTaxtRate = 0
                                defaultTaxtRateValue = 0
                                row.Cells("TaxRate").Value = 0
                                row.Cells("TaxRateValue").Value = 0
                            End If
                        End If
                    End If
                Next
            End If
            isDefaultChange = False
            SetTaxedPriceLable(isTaxedPrice)
            TotalValuesBeahavior()
            SetTotalAmounts(True)
        Catch ex As Exception
            modGlazingQuoteExtension.GQShowMessage(ex.Message, Me.Text, MsgBoxStyle.Critical)

        End Try

    End Sub

    Private Sub tsmAddRowBefore_Click(sender As Object, e As EventArgs) Handles tsmAddRowBefore.Click
        AddNewRow("before")
    End Sub

    Private Sub tsmAddRowAfter_Click(sender As Object, e As EventArgs) Handles tsmAddRowAfter.Click
        AddNewRow("after")
    End Sub



#Region "UltraGrid : UG2"

#Region "ultraGrid behaviour"

    Private Sub UG2_AfterRowsDeleted(sender As Object, e As EventArgs) Handles UG2.AfterRowsDeleted
        'SetSubTotalByGroup()
        'SetTotalAmounts(True)
        QuoteGridSetSubTotal()
    End Sub
    '------------------------start ultraGrid behaviour------------------------
    Private Sub UG2_InitializeLayout(sender As Object, e As InitializeLayoutEventArgs) Handles UG2.InitializeLayout, UltraGrid1.InitializeLayout

        Try
            UG2.Rows.Band.Override.ExpansionIndicator = ShowExpansionIndicator.CheckOnDisplay
            'setQuateFiedType(sender, e)

            SetCellApperance()
            UG2.DisplayLayout.Override.CellAppearance.TextVAlign = VAlign.Middle
            UG2.DisplayLayout.ColumnChooserEnabled = DefaultableBoolean.True
            UG2.DisplayLayout.Rows.Band.Columns("LineComments").DefaultCellValue = ""

            ' Set the RowSelectorHeaderStyle to ColumnChooserButton.
            Me.UG2.DisplayLayout.Override.RowSelectorHeaderStyle = _
              RowSelectorHeaderStyle.ColumnChooserButton

            ' Enable the RowSelectors. This is necessary because the column chooser
            ' button is displayed over the row selectors in the column headers area.
            Me.UG2.DisplayLayout.Override.RowSelectors = DefaultableBoolean.True

            'e.Layout.Bands(1).ExcludeFromColumnChooser = ExcludeFromColumnChooser.True
            'For Each ugCol As UltraGridColumn In e.Layout.Bands(0).Columns
            '    e.Layout.Bands(0).Columns(ugCol.Key).ExcludeFromColumnChooser = ExcludeFromColumnChooser.True
            'Next
            'e.Layout.Bands(0).Columns("Volume").ExcludeFromColumnChooser = ExcludeFromColumnChooser.False
            'e.Layout.Bands(0).Columns("StockLink").ExcludeFromColumnChooser = ExcludeFromColumnChooser.False
            'e.Layout.Bands(0).Columns("ItmGroupID").ExcludeFromColumnChooser = ExcludeFromColumnChooser.False
            'e.Layout.Bands(0).Columns("ItemImage").Style = ColumnStyle.Edit

            'e.Layout.Bands(0).Columns(1).ColumnChooserCaption = "Column Chooser Caption"

        Catch ex As Exception
            modGlazingQuoteExtension.GQShowMessage(ex.Message, Me.Text, MsgBoxStyle.Critical)

        End Try
    End Sub

    Sub SetCellApperance()
        Try
            UG2.Rows.Band.Columns("QuoteFiedType").CellAppearance.TextHAlign = HAlign.Center
            UG2.Rows.Band.Columns("QuoteFiedType").CellAppearance.TextVAlign = VAlign.Middle
            UG2.Rows.Band.Columns("Height").CellAppearance.TextHAlign = HAlign.Right
            UG2.Rows.Band.Columns("Width").CellAppearance.TextHAlign = HAlign.Right
            UG2.Rows.Band.Columns("Qty").CellAppearance.TextHAlign = HAlign.Right
            UG2.Rows.Band.Columns("Price").CellAppearance.TextHAlign = HAlign.Right
            UG2.Rows.Band.Columns("Amount").CellAppearance.TextHAlign = HAlign.Right
            UG2.Refresh()

        Catch ex As Exception
            modGlazingQuoteExtension.GQShowMessage(ex.Message, Me.Text, MsgBoxStyle.Critical)

        End Try
        'UG2.
        'UG2.Rows.Band.Columns("Qty").CellAppearance.TextHAlign = HAlign.Right

    End Sub

    Private Sub UG2_AfterExitEditMode(sender As Object, e As EventArgs) Handles UG2.AfterExitEditMode, UltraGrid1.AfterExitEditMode
        Try
            If isPasting = False Then
                If QuoteGridValueValidationAfterUpdate() = 0 Then
                    Exit Sub
                End If
            End If

            If Me.UG2.ActiveCell.Column.Key = "Amount" Then
                If IsDBNull(Me.UG2.ActiveRow.Cells("QuoteFiedType").Value) = False Then
                    If Me.UG2.ActiveRow.Cells("QuoteFiedType").Value = QuateFiedTypesList.Text Or Me.UG2.ActiveRow.Cells("QuoteFiedType").Value = QuateFiedTypesList.Stock_Item Then
                        If IsNumeric(Me.UG2.ActiveRow.Cells("Amount").Text) = True Then
                            If Me.UG2.ActiveRow.Cells("Amount").Value <> Convert.ToDecimal(Me.UG2.ActiveRow.Cells("Net").Value) Then
                                If canEditeAmount = True Then
                                    IsCellClearing = True
                                    ClearRowCellsAfterAmountChange()
                                    IsCellClearing = False
                                Else
                                    Me.UG2.ActiveRow.Cells("Amount").Value = Me.UG2.ActiveRow.Cells("Amount").Value
                                End If
                                'SetTotalAmounts(True)
                                isOpeningQuote = True
                                calculatePriceUsingAmount(Me.UG2.ActiveRow)
                                isTaxedPrice = True
                                TotalValuesBeahavior()

                                isOpeningQuote = False
                                QuoteGridSetSubTotal()

                            End If
                        End If
                    End If
                End If
            End If

            If isStockItemActive = False Then
                If isStockItemClosing = False Then
                    If Me.UG2.ActiveCell.Column.Key = "QuoteFiedType" Then
                        Dim ex As CellEventArgs
                        'QuoteGirdRowFormat(ex)

                        'ElseIf e.Cell.Column.Key = "TaxRate" Then
                        'If e.Cell.Row.Cells("Price").Text > 0 Then
                        '    CalculateItemAmount(e)

                    End If
                Else
                    isStockItemClosing = False
                End If
            End If


            If isStockItemActive = False Then
                If Me.UG2.ActiveCell.Column.Key = "LineComments" Then
                    Me.UG2.ActiveCell.Column.Header.Caption = "Doc Description"

                ElseIf Me.UG2.ActiveCell.Column.Key = "QuoteFiedType" Then
                    'Keyboard usage speedup
                    If downKeyPressed = True Then
                        QuoteGridNavigator()
                        QuoteGridFunctions()
                        downKeyPressed = False

                    End If
                End If
            End If
            SetTotalAmounts(True)
        Catch ex As Exception
            modGlazingQuoteExtension.GQShowMessage(ex.Message, Me.Text, MsgBoxStyle.Critical)

        End Try
    End Sub
    '___________________________________End ultraGrid behaviour___________________________________
#End Region

#Region "cell behaviour"
    '------------------------start cell behaviour------------------------
    Private Sub UG2_AfterCellActivate(sender As Object, e As EventArgs) Handles UG2.AfterCellActivate, UltraGrid1.AfterCellActivate
        If isStockItemActive = False Then
            If Me.UG2.ActiveCell.Column.Key = "LineComments" Then
                If Me.UG2.ActiveRow.Cells("QuoteFiedType").Text = "" Then
                    Me.UG2.ActiveRow.Cells("QuoteFiedType").DroppedDown = True
                Else
                    Me.UG2.ActiveCell.Column.Header.Caption = "Doc Description (Press F8 to add predefined text)"
                End If
            End If
        End If
    End Sub

    Private Sub UG2_CellChange(sender As Object, e As CellEventArgs) Handles UG2.CellChange
        'Dim test2 = e.Cell.Text
        'Dim a2 = Me.UG2.ActiveCell.Text
        'If isStockItemActive = False Then
        '    If isStockItemClosing = False Then
        '        If e.Cell.Column.Key = "QuoteFiedType" Then
        '            QuoteGirdRowFormat(e)

        '            'ElseIf e.Cell.Column.Key = "TaxRate" Then
        '            'If e.Cell.Row.Cells("Price").Text > 0 Then
        '            '    CalculateItemAmount(e)

        '        End If
        '    Else
        '        isStockItemClosing = False
        '    End If
        'End If
    End Sub

    Private Sub UG2_AfterCellUpdate(sender As Object, e As CellEventArgs) Handles ugQuote.AfterCellUpdate, UG2.AfterCellUpdate, UltraGrid2.AfterCellUpdate, UltraGrid1.AfterCellUpdate
        Try
            'If (e.Cell.Row.Cells("IsAExistingItem").Value = True And isOpeningQuote = False) And (e.Cell.Column.Key <> "TaxRate" Or e.Cell.Column.Key <> "TaxRateValue" Or e.Cell.Column.Key <> "IsAExistingItem") Then
            'e.Cell.Row.Cells("TaxRate").Value = defaultTaxtRate
            'e.Cell.Row.Cells("TaxRateValue").Value = defaultTaxtRateValue
            'If e.Cell.Row.Cells("IsAExistingItem").Value = True Then
            '    e.Cell.Row.Cells("IsAExistingItem").Value = False
            'End If
            'End If
            If isPasting = False And isOpeningQuote = False Then
                If QuoteGridValueValidationBeforeUpdate() = 0 Then
                    Exit Sub
                End If
            Else
                Exit Sub
            End If

            If isStockItemActive = False And isOpeningQuote = False Then

                If isPasting = True Then
                    If e.Cell.Column.Key = "QuoteFiedType" Or calculateWhilePasting = True Then
                        'QuoteGirdRowFormat(e)
                        Me.UG2.ActiveRow.Cells("LineComments").Activated = True
                    ElseIf e.Cell.Column.Key = "iInvDetailID" Then
                        'e.Cell.Row.Cells("iInvDetailID").Value = 0
                    End If
                End If

                If e.Cell.Column.Key = "LineComments" Then
                    e.Cell.Row.PerformAutoSize()
                    SetCellApperance()

                ElseIf e.Cell.Column.Key = "Width" Or e.Cell.Column.Key = "Height" Then
                    CalculateItemVolume(e)

                ElseIf e.Cell.Column.Key = "Qty" Or e.Cell.Column.Key = "Volume" Or e.Cell.Column.Key = "Price" Or e.Cell.Column.Key = "TaxRate" Or e.Cell.Column.Key = "DiscAmt" Then
                    If isOpeningQuote = False And IsCellClearing = False Then
                        CalculateItemAmount(e)
                        QuoteGridSetSubTotal()
                    End If

                ElseIf e.Cell.Column.Key = "Amount" Then
                    'QuoteGridSetSubTotal()
                    If e.Cell.Row.Cells("QuoteFiedType").Text = "Subtotal" Then
                        'UG2.Rows(headderSubRow).Cells("OrgPrice").Value = e.Cell.Row.Cells("Amount").Value
                        'SetTotalAmounts(True)

                    Else
                        'If isPasting = False Or calculateWhilePasting = True Then
                        SetSubTotal(e)

                        ''End If
                    End If
                ElseIf e.Cell.Column.Key = "OrgPrice" Then
                    If e.Cell.Row.Cells("QuoteFiedType").Value = QuateFiedTypesList.Subtotal Then

                    End If
                ElseIf e.Cell.Column.Key = "OrgPrice" Then

                ElseIf e.Cell.Column.Key = "ServiceItemTotNet" Then
                    GridCellAppearence(e.Cell.Row)

                End If

                If (isExistingOrder = True Or openedModeulename = "Temp") And e.Cell.Text <> "Stock Item" Then
                    'QuoteGirdRowFormat(e)
                End If

            End If
            SetActiveCell(e)
        Catch ex As Exception
            modGlazingQuoteExtension.GQShowMessage(ex.Message, Me.Text, MsgBoxStyle.Critical)

        End Try
    End Sub
    '________________________________________________End cell behaviour________________________________________________
#End Region

#Region "cell behaviour"
    '------------------------------------------------start key behaviour------------------------------------------------

    Private Sub UG2_KeyDown(sender As Object, e As KeyEventArgs) Handles UG2.KeyDown, UltraGrid1.KeyDown
        Try
            KeyCombination(e)
            If e.KeyCode = 119 Then
                If IsNothing(Me.UG2.ActiveCell) = False Then
                    UG2ActiveRow = UG2.ActiveRow
                    Dim selectedColum As Integer
                    If UG2.ActiveCell.Column.Key = "LineComments" Then
                        Dim newGlazingDocDescription As New frmGlazingDocDescription
                        newGlazingDocDescription.DocDesTypeID = UG2.ActiveRow.Cells("QuoteFiedType").Value
                        newGlazingDocDescription.DocDesTypeName = UG2.ActiveRow.Cells("QuoteFiedType").Text
                        newGlazingDocDescription.ShowDialog()
                        If UG2.Rows(UG2ActiveRow.Index).Cells("LineComments").Text <> "" Then
                            UG2.Rows(UG2ActiveRow.Index).Cells("LineComments").Value = UG2.Rows(UG2ActiveRow.Index).Cells("LineComments").Text + vbCrLf + selectedDocDes
                        ElseIf UG2.Rows(UG2ActiveRow.Index).Cells("LineComments").Text = "" Then
                            UG2.Rows(UG2ActiveRow.Index).Cells("LineComments").Value = selectedDocDes
                        Else
                        End If
                        Me.UG2.ActiveCell.SelectAll()
                        selectedDocDes = ""

                    End If
                End If

            ElseIf e.KeyCode = Keys.Tab Then
                AddNewRowsValidator("after")

            ElseIf e.KeyCode = Keys.Down Then
                If IsNothing(UG2.ActiveCell) = False Then
                    If UG2.ActiveCell.Column.Key = "QuoteFiedType" Then
                        downKeyPressed = True
                    End If
                End If

            ElseIf e.KeyCode = Keys.Enter Then
                If IsNothing(UG2.ActiveCell) = False Then
                    If UG2.ActiveCell.Column.Key = "QuoteFiedType" Then
                        downKeyPressed = False
                        QuoteGridNavigator()
                        QuoteGridFunctions()
                    End If
                End If
            End If

        Catch ex As Exception
            modGlazingQuoteExtension.GQShowMessage(ex.Message, Me.Text, MsgBoxStyle.Critical)

        End Try
    End Sub
    '___________________________________Endstart key behaviourbehaviour___________________________________
#End Region

#End Region

    Private Sub CalculateItemVolume(e As CellEventArgs)
        Try
            Dim itemHight As Decimal = 0
            Dim itemWidth As Decimal = 0

            If IsNumeric(e.Cell.Row.Cells("Width").Text) Then
                itemWidth = e.Cell.Row.Cells("Width").Text
            End If
            If IsNumeric(e.Cell.Row.Cells("Height").Text) Then
                itemHight = e.Cell.Row.Cells("Height").Text
            End If

            e.Cell.Row.Cells("Volume").Value = (itemWidth * itemHight) / 1000000

        Catch ex As Exception
            modGlazingQuoteExtension.GQShowMessage(ex.Message, Me.Text, MsgBoxStyle.Critical)

        End Try
    End Sub

    Private Sub CalculateItemAmount(e As CellEventArgs)
        Try
            Dim itemVolume As Decimal = 0
            Dim itemPrice As Decimal = 0
            Dim ItemQty As Decimal = 0
            Dim itemTaxRate As Decimal = 0
            Dim itemDiscount As Decimal = 0
            Dim itemAmount As Decimal = 0

            Dim itemDiscountedPrice As Decimal = 0
            Dim DiscountPrice As Decimal = 0
            Dim itemIncAmount As Decimal = 0
            Dim itemExcAmount As Decimal = 0
            Dim TaxAmount As Decimal = 0

            Dim itemPreExcAmount As Decimal = 0
            Dim itemPreAmount As Decimal = 0
            Dim itemPreTax As Decimal = 0


            If IsNumeric(e.Cell.Row.Cells("Volume").Text) Then
                itemVolume = e.Cell.Row.Cells("Volume").Text

            Else

            End If
            If IsNumeric(e.Cell.Row.Cells("Price").Text) Then
                itemPrice = e.Cell.Row.Cells("Price").Text

            End If

            'If IsNothing(e.Cell.Row.Cells("TaxRate").Value) = False Then
            '    If IsNumeric(e.Cell.Row.Cells("TaxRate").Value) = True Then
            '        ucmbTaxRate.Value = e.Cell.Row.Cells("TaxRate").Value
            '        If IsNothing(ucmbTaxRate.SelectedRow) = False Then
            '            itemTaxRate = ucmbTaxRate.SelectedRow.Cells("TaxRate").Value

            '        End If
            '    End If
            'Else

            If IsNothing(e.Cell.Row.Cells("TaxRate").Text) = False Then
                If e.Cell.Row.Cells("TaxRate").Text > -1 Then
                    itemTaxRate = e.Cell.Row.Cells("TaxRate").Text
                Else
                    itemTaxRate = defaultTaxtRateValue
                End If
            End If

            If IsNumeric(e.Cell.Row.Cells("DiscAmt").Text) Then
                itemDiscount = e.Cell.Row.Cells("DiscAmt").Text
            End If
            If IsNumeric(e.Cell.Row.Cells("Qty").Text) Then
                ItemQty = e.Cell.Row.Cells("Qty").Text
            End If
            If IsNumeric(e.Cell.Row.Cells("Amount").Text) Then
                itemPreAmount = e.Cell.Row.Cells("Amount").Text
            End If
            If IsNumeric(e.Cell.Row.Cells("ItmExcAmount").Text) Then
                itemPreExcAmount = e.Cell.Row.Cells("ItmExcAmount").Text
            End If
            If IsNumeric(e.Cell.Row.Cells("Tax").Text) Then
                itemPreTax = e.Cell.Row.Cells("Tax").Text
            End If

            'Adding discount
            If itemDiscount > 0 And itemPrice > 0 Then
                DiscountPrice = ((itemPrice / 100) * itemDiscount)
                itemDiscountedPrice = itemPrice - DiscountPrice
            End If

            'Exclusive amount
            If itemVolume > 0 Then
                'Glass
                itemExcAmount = ((itemPrice - DiscountPrice) * itemVolume) * ItemQty

            ElseIf ItemQty > 0 Then
                'Other
                itemExcAmount = (itemPrice - DiscountPrice) * ItemQty

            End If

            'for manual amount
            If itemVolume = 0 And ItemQty = 0 Then
                itemExcAmount = e.Cell.Row.Cells("ItmExcAmount").Text

            End If

            If itemTaxRate > -1 And itemExcAmount > -1 Then
                'adding tax
                TaxAmount = ((itemExcAmount / 100) * itemTaxRate)
                itemIncAmount = TaxAmount + itemExcAmount
            Else
                itemIncAmount = itemExcAmount
                'TaxAmount = 0
            End If


            'Total amount
            If isTaxedPrice = True Then
                e.Cell.Row.Cells("Amount").Value = Math.Round(itemIncAmount, 2, MidpointRounding.AwayFromZero)
            Else
                e.Cell.Row.Cells("Amount").Value = Math.Round(itemExcAmount, 2, MidpointRounding.AwayFromZero)

            End If

            e.Cell.Row.Cells("DiscAmt").Value = DiscountPrice
            e.Cell.Row.Cells("ItmExcAmount").Value = itemExcAmount
            e.Cell.Row.Cells("Tax").Value = TaxAmount
            e.Cell.Row.Cells("Net").Value = itemIncAmount

            If isOpeningQuote = False Then
                lblTotExcAmo.Text = Format((Val(lblTotExcAmo.Text - itemPreExcAmount) + itemExcAmount), "0.00")
                lblTotVatAmo.Text = Format((Val(lblTotVatAmo.Text - itemPreTax) + TaxAmount), "0.00")
                lblTotIncAmo.Text = Format((Val(lblTotIncAmo.Text - itemPreAmount) + itemIncAmount), "0.00")
            End If
            calculateWhilePasting = False

        Catch ex As Exception
            modGlazingQuoteExtension.GQShowMessage(ex.Message, Me.Text, MsgBoxStyle.Critical)

        End Try
    End Sub

    Private Sub Button16_Click(sender As Object, e As EventArgs) Handles Button16.Click
        'If IsNothing(UG2.ActiveRow) = False Then
        '    UG2ActiveRow = UG2.ActiveRow
        '    Dim glazingDocStockItem As New frmGlazingDocStockItem(Me)
        '    'glazingDocStockItem.StartPosition = FormStartPosition.CenterParent
        '    'glazingDocStockItem.Location = New Point(UG2.Location.X, UltraTabControl1.Location.Y)
        '    glazingDocStockItem.ShowDialog()

        'End If
        SetCellApperance()

    End Sub

    Private Sub UG2_DoubleClickRow(sender As Object, e As DoubleClickRowEventArgs) Handles UG2.DoubleClickRow
        Try
            If IsNothing(e.Row) = False And IsDBNull(UG2.ActiveRow.Cells("QuoteFiedType").Value) = False Then
                If UG2.ActiveRow.Cells("QuoteFiedType").Value <> "" Then
                    If UG2.ActiveRow.Cells("QuoteFiedType").Value = QuateFiedTypesList.Stock_Item And UG2.ActiveRow.Cells("ItemImage").Activated = False Then
                        Dim glazingDocStockItem As New frmGlazingDocStockItem(Me)
                        'glazingDocStockItem.selectedRow = UG2.ActiveRow
                        ' Dim yesy2 = UG2.ActiveRow.Cells("stockLink").Value
                        glazingDocStockItem.ShowDialog()
                    End If
                End If
            End If

        Catch ex As Exception
            modGlazingQuoteExtension.GQShowMessage(ex.Message, Me.Text, MsgBoxStyle.Critical)

        End Try
    End Sub

    Private Sub UltraTabPageControl13_Paint(sender As Object, e As PaintEventArgs) Handles UltraTabPageControl13.Paint

    End Sub

    Private Sub btnBrowse_Click(sender As Object, e As EventArgs) Handles btnBrowse.Click
        Try
            Dim SelectedFile As String = String.Empty

            If utcQuoteGrids.ActiveTab.Index <> 1 Then Exit Sub

            Dim R As UltraGridRow
            R = UGDocs.DisplayLayout.Bands(0).AddNew

            UGDocs.ActiveRow.Cells("LineNo").Value = UGDocs.ActiveRow.Index + 1
            OpenFileDialog1.ShowDialog()
            R.Cells(1).Activated = True
            R.Cells(1).Value = OpenFileDialog1.FileName

        Catch ex As Exception
            modGlazingQuoteExtension.GQShowMessage(ex.Message, Me.Text, MsgBoxStyle.Critical)

        End Try
    End Sub

    Private Sub UGDocs_ClickCellButton(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles UGDocs.ClickCellButton
        Try
            If e.Cell.Column.Index = 2 Then
                Dim Proc As New System.Diagnostics.Process
                Proc.StartInfo.FileName = UGDocs.ActiveRow.Cells(1).Value.ToString
                Proc.Start()
            End If
        Catch ex As Exception
            modGlazingQuoteExtension.GQShowMessage(ex.Message, Me.Text, MsgBoxStyle.Critical)

        End Try

    End Sub

    Private Sub UGDocs_DoubleClickRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles UGDocs.DoubleClickRow
        Try

            'If e.Cell.Column.Index = 2 Then
            Dim Proc As New System.Diagnostics.Process
            Proc.StartInfo.FileName = UGDocs.ActiveRow.Cells(1).Value.ToString
            Proc.Start()
            'End If
        Catch ex As Exception
            modGlazingQuoteExtension.GQShowMessage(ex.Message, Me.Text, MsgBoxStyle.Critical)

        End Try
    End Sub

    Private Sub tsmiExportTemplate_Click(sender As Object, e As EventArgs) Handles tsmiExportTemplate.Click
        Try
            Dim newGlazingDocTemplate As New frmGlazingDocTemplate(Me)
            newGlazingDocTemplate.ShowDialog()
        Catch ex As Exception
            modGlazingQuoteExtension.GQShowMessage(ex.Message, Me.Text, MsgBoxStyle.Critical)

        End Try
    End Sub

    Public Sub SaveTempalteData(dsQuoteTemp As DataSet, tempName As String)
        Try
            Dim oSQL = New clsSqlConn
            Dim newSQLQuery As String
            oSQL.Begin_Trans()

            Dim i As Integer = 0
            Dim docDescription As String = ""
            For Each ugR As UltraGridRow In UG2.Rows
                If Not IsDBNull(ugR.Cells("quoteFiedType").Value) Then
                    If ugR.Cells("QuoteFiedType").Value = QuateFiedTypesList.Text Or ugR.Cells("QuoteFiedType").Value = QuateFiedTypesList.Stock_Item Then
                        docDescription = ""
                    Else
                        docDescription = ugR.Cells("LineComments").Value
                    End If

                    newSQLQuery = "INSERT INTO GlzQuote_Temp (TempName, TempQuoteFiedType, AgentID, EditedDate, LastEditedBy, DocDescription, RowGroupID) VALUES ('" & tempName & "', " & ugR.Cells("QuoteFiedType").Value & "," & AgentID & "," & Today.Date & "," & AgentID & ",'" & docDescription & "', " & ugR.Cells("ItmGroupID").Value & ")"
                    If oSQL.Exe_Query_Trans(newSQLQuery) = 0 Then
                        modGlazingQuoteExtension.GQShowMessage("Data not saved", Me.Text, MsgBoxStyle.Critical)
                        oSQL.Rollback_Trans()
                        Exit Sub
                    End If


                End If
            Next

            oSQL.Commit_Trans()

        Catch ex As Exception
            modGlazingQuoteExtension.GQShowMessage(ex.Message, Me.Text, MsgBoxStyle.Critical)

        End Try
    End Sub

    Private Sub UG2_BeforeRowsDeleted(sender As Object, e As BeforeRowsDeletedEventArgs) Handles UG2.BeforeRowsDeleted
        Try
            If IsNothing(UG2.ActiveRow) = False Then
                Dim rowSelected As UltraGridRow
                If UG2.Selected.Rows.Count > 0 Then
                    For Each rowSelected In Me.UG2.Selected.Rows
                        If rowSelected.Cells("iInvDetailID").Value > 0 Then
                            collDeletedItemLines.Add(rowSelected.Cells("iInvDetailID").Value)

                        End If
                    Next
                End If
            End If
            e.DisplayPromptMsg = False
        Catch ex As Exception
            modGlazingQuoteExtension.GQShowMessage(ex.Message, Me.Text, MsgBoxStyle.Critical)

        End Try
    End Sub

    Private Sub UG2_BeforeExitEditMode(sender As Object, e As UltraWinGrid.BeforeExitEditModeEventArgs) Handles UG2.BeforeExitEditMode
        Try
            'If UG2.ActiveCell.Column.Key = "Qty" Or UG2.ActiveCell.Column.Key = "Volume" Or UG2.ActiveCell.Column.Key = "Price" Or UG2.ActiveCell.Column.Key = "TaxRate" Or UG2.ActiveCell.Column.Key = "DiscAmt" Then
            '    'If UG2.ActiveRow.Cells("Price").Text > 0 Then
            '    '    CalculateItemAmount()
            '    'End If
            If QuoteGridValueValidationBeforeUpdate() = 0 Then
                e.Cancel = True
                Exit Sub
            Else

            End If


        Catch ex As Exception
            modGlazingQuoteExtension.GQShowMessage(ex.Message, Me.Text, MsgBoxStyle.Critical)

        End Try
    End Sub
    Sub TotalValuesBeahavior()
        Try
            If isTaxedPrice = True Then
                lblTotExcAmo.Visible = True
                lblTotVatAmo.Visible = True
                lblTotIncAmo.Visible = True

                T1.Visible = True
                T2.Visible = True
                T3.Visible = True

                lblTotalExc.Visible = True
                lblTotalVat.Visible = True
                lblTotalInc.Visible = True
            Else
                lblTotExcAmo.Visible = True
                lblTotVatAmo.Visible = False
                lblTotIncAmo.Visible = False

                T1.Visible = True
                T2.Visible = False
                T3.Visible = False

                lblTotalExc.Visible = True
                lblTotalVat.Visible = False
                lblTotalInc.Visible = False
            End If

        Catch ex As Exception
            modGlazingQuoteExtension.GQShowMessage(ex.Message, Me.Text, MsgBoxStyle.Critical)
        End Try
    End Sub

    Sub getTotalAmount()

        Dim totExc As Decimal = 0
        Dim totInc As Decimal = 0
        Dim totvat As Decimal = 0
        Dim totAmount As Decimal = 0

        Try
            For Each ugRow As UltraGridRow In UG2.Rows

                If ugRow.Cells("QuoteFiedType").Value = QuateFiedTypesList.Stock_Item Or ugRow.Cells("QuoteFiedType").Value = QuateFiedTypesList.Text Then

                    If IsNothing(ugRow.Cells("Price").Value) = False Then

                        'no price
                        If ugRow.Cells("Price").Value = 0 Then
                            If isTaxedPrice = False Then
                            Else


                            End If
                        End If

                        'Exclusive
                        If IsNothing(ugRow.Cells("Amount").Value) = False Then

                            'have Amount
                            If ugRow.Cells("Amount").Value > 0 Then
                                totExc += ugRow.Cells("Amount").Value
                            End If


                        Else
                            ugRow.Cells("Amount").Value = 0
                        End If


                        'Have a price
                    Else
                        If IsNothing(ugRow.Cells("Qty").Value) = False Then

                            'have a price no qty
                            If ugRow.Cells("Price").Value > 0 And ugRow.Cells("Qty").Value = 0 Then
                                If IsNothing(ugRow.Cells("Amount").Value) = False Then

                                    'have Amount
                                    If ugRow.Cells("Amount").Value > 0 And isTaxedPrice = False Then
                                        totExc += ugRow.Cells("Amount").Value
                                    End If
                                End If
                            Else

                                'have a price and qty
                                totAmount += ugRow.Cells("Price").Value * ugRow.Cells("Qty").Value
                                totExc += totAmount

                            End If
                        Else
                            ugRow.Cells("Qty").Value = 0
                        End If
                    End If
                Else
                    ugRow.Cells("Price").Value = 0

                End If
                'End If
            Next
        Catch ex As Exception
            modGlazingQuoteExtension.GQShowMessage(ex.Message, Me.Text, MsgBoxStyle.Critical)

        End Try
    End Sub
    Sub GetDefaultTaxtRate()

    End Sub

    Sub CopyRow()
        'UnhideColums()
        'Me.UG2.DisplayLayout.Override.AllowMultiCellOperations = AllowMultiCellOperation.Copy
        'Me.UG2.PerformAction(UltraGridAction.Copy)
        Try

            rowCopied.Clear()
            If IsNothing(UG2.Selected.Rows) = False Then
                If UG2.Selected.Rows.Count > 0 Then
                    For Each rowSelected In Me.UG2.Selected.Rows
                        rowCopied.Add(rowSelected)
                    Next
                End If
            End If
            'HideColums()
            isCopied = True

        Catch ex As Exception
            modGlazingQuoteExtension.GQShowMessage(ex.Message, Me.Text, MsgBoxStyle.Critical)

        End Try
    End Sub

    Sub PastRow()
        Try
            isPasting = True
            Dim copiedRowCount As Integer = rowCopied.Count
            Dim cellCount As Integer
            Dim row As UltraGridRow
            Dim cell As UltraGridCell
            If UG2.Selected.Rows.Count > 0 Then
                For Each row In rowCopied
                    copiedRowCount -= 1
                    cellCount = row.Cells.Count

                    For Each cell In row.Cells
                        cellCount -= 1
                        If cell.Column.Key = "iInvDetailID" Then
                            If Me.UG2.ActiveRow.Cells("iInvDetailID").Value > 0 Then
                                'Updating the row
                            Else
                                Me.UG2.ActiveRow.Cells("iInvDetailID").Value = 0

                            End If

                        ElseIf cell.Column.Key = "ItmGroupID" Then
                            RowDataOverride(Me.UG2.ActiveRow.Index)


                        ElseIf cell.Column.Key = "InvLineID" Then
                            If Me.UG2.ActiveRow.Cells("InvLineID").Value > 0 Then
                                'Updating the row
                            Else
                                Me.UG2.ActiveRow.Cells("InvLineID").Value = 0

                            End If

                        ElseIf cell.Column.Key = "isPastedRow" Then
                            'calculateWhilePasting = True
                            Me.UG2.ActiveRow.Cells(cell.Column.Key).Value = True
                            'calculateWhilePasting = False

                        ElseIf cell.Column.Key = "Shape" Then
                            If IsNothing(Me.UG2.ActiveRow.Cells("Shape")) = False Then
                                If IsNothing(cell.Row.Cells("Shape").ButtonAppearance.Image) = False Then
                                    Me.UG2.ActiveRow.Cells("Shape").Appearance.BackColor = Color.LightBlue
                                    Me.UG2.ActiveRow.Cells("Shape").ButtonAppearance.Image = Global.SPIL_Glass.My.Resources.Resources.shapesadd
                                    Me.UG2.ActiveRow.Cells("Shape").Value = "Add"
                                End If

                            End If

                        Else
                            Me.UG2.ActiveRow.Cells(cell.Column.Key).Value = row.Cells(cell.Column.Key).Value
                            If cell.Column.Key = "QuoteFiedType" Then
                                ucmbQuoteLineType.Value = row.Cells(cell.Column.Key).Value

                            End If
                        End If

                    Next
                    QuoteGridSetSubTotal()
                    If copiedRowCount > 0 Then
                        AddNewRow("after")
                    End If

                Next
            End If
            'UnhideColums()
            'Me.UG2.DisplayLayout.Override.AllowMultiCellOperations = AllowMultiCellOperation.Paste
            'Me.UG2.PerformAction(UltraGridAction.Paste)
            'HideColums()
            isPasting = False

        Catch ex As Exception
            modGlazingQuoteExtension.GQShowMessage(ex.Message, Me.Text, MsgBoxStyle.Critical)

        End Try

    End Sub


    Private Sub tsmDeli_Click(sender As Object, e As EventArgs) Handles tsmiDel.Click
        Me.UG2.DeleteSelectedRows()
    End Sub

    Private Sub tsbPrint_Click(sender As Object, e As EventArgs) Handles tsbPrint.Click
        LoadPrint(Nothing, quoteOrdeIndex)
    End Sub

    Private Sub frmGlazingQuote_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        isClosing = True
        If isSaved = False Then
            If UG2.Rows.Count = 0 Then

            ElseIf UG2.Rows.Count > 0 Then
                If UG2.Rows.Count = 1 Then
                    If UG2.Rows(0).Cells("QuoteFiedType").Value = "" Then
                        UG2.Rows(0).Delete()

                    End If
                Else
                    If modGlazingQuoteExtension.GQShowMessage("Do you want to save data before exit?", Me.Text, MsgBoxStyle.YesNo) = Windows.Forms.DialogResult.Yes Then
                        If SaveDocument() = 0 Then
                            e.Cancel = True
                            Exit Sub
                        End If
                    End If
                End If
            End If
        End If
        quoteOrdeIndex = 0
        Me.Dispose()
        System.GC.Collect()

    End Sub

    Sub UnhideColums()
        Try
            Me.UG2.ActiveRow.Cells("Tax").Column.Hidden = False
            Me.UG2.ActiveRow.Cells("DiscAmt").Column.Hidden = False
            Me.UG2.ActiveRow.Cells("Original_Price").Column.Hidden = False
            Me.UG2.ActiveRow.Cells("iInvDetailID").Column.Hidden = False
            Me.UG2.ActiveRow.Cells("InvLineID").Column.Hidden = False
            Me.UG2.ActiveRow.Cells("OrgPrice").Column.Hidden = False
            Me.UG2.ActiveRow.Cells("TaxRate").Column.Hidden = False
            Me.UG2.ActiveRow.Cells("Volume").Column.Hidden = False
            Me.UG2.ActiveRow.Cells("ItmExcAmount").Column.Hidden = False
        Catch ex As Exception
            modGlazingQuoteExtension.GQShowMessage(ex.Message, Me.Text, MsgBoxStyle.Critical)

        End Try
    End Sub
    Sub HideColums()
        Try
            Me.UG2.ActiveRow.Cells("Tax").Column.Hidden = True
            Me.UG2.ActiveRow.Cells("DiscAmt").Column.Hidden = True
            Me.UG2.ActiveRow.Cells("Original_Price").Column.Hidden = True
            Me.UG2.ActiveRow.Cells("iInvDetailID").Column.Hidden = True
            Me.UG2.ActiveRow.Cells("InvLineID").Column.Hidden = True
            Me.UG2.ActiveRow.Cells("OrgPrice").Column.Hidden = True
            Me.UG2.ActiveRow.Cells("TaxRate").Column.Hidden = True
            Me.UG2.ActiveRow.Cells("Volume").Column.Hidden = True
            Me.UG2.ActiveRow.Cells("ItmExcAmount").Column.Hidden = True
        Catch ex As Exception
            modGlazingQuoteExtension.GQShowMessage(ex.Message, Me.Text, MsgBoxStyle.Critical)

        End Try
    End Sub

    Private Sub tsmiCopy_Click(sender As Object, e As EventArgs) Handles tsmiCopy.Click
        If UG2.Selected.Rows.Count > 0 Then
            CopyRow()
        End If
    End Sub

    Private Sub tsmiPaste_Click(sender As Object, e As EventArgs) Handles tsmiPaste.Click
        If UG2.Selected.Rows.Count > 0 Then
            PastRow()
        End If
    End Sub

    Private Sub mnuSave_Click(sender As Object, e As EventArgs) Handles mnuSave.Click
        SaveDocument()
    End Sub

    Private Sub mnuNewDocument_Click(sender As Object, e As EventArgs) Handles mnuNewDocument.Click
        clsGlazingQuoteExtensionObj.GQOpenGlazingQuote()
    End Sub

    Private Sub UG2_ClickCellButton(sender As Object, e As CellEventArgs) Handles UG2.ClickCellButton
        Try
            If e.Cell.Column.Key = "Shape" Then

                If oFacDefaults.OptimizeApp = GlassOptimizeApp.OptiWay Then

                    Dim oShape As New SPIL.Shapes.frmShapeEdit

                    If Not e.Cell.Row.Cells("Shape").Value Is Nothing And Not IsDBNull(e.Cell.Row.Cells("Shape").Value) Then
                        SPIL.Shapes.GlobalVariables.SPILShapeName = CStr(e.Cell.Row.Cells("Shape").Value)
                        SPIL.Shapes.GlobalVariables.SPILShapeHeight = CStr(e.Cell.Row.Cells("Height").Value)
                        SPIL.Shapes.GlobalVariables.SPILShapeWidth = CStr(e.Cell.Row.Cells("Width").Value)
                    Else
                        SPIL.Shapes.GlobalVariables.SPILShapeName = ""
                        SPIL.Shapes.GlobalVariables.SPILShapeHeight = ""
                        SPIL.Shapes.GlobalVariables.SPILShapeWidth = ""
                    End If

                    oShape.StartPosition = FormStartPosition.CenterScreen
                    oShape.ShowDialog()

                    If SPIL.Shapes.GlobalVariables.SPILShapeHeight <> "" And SPIL.Shapes.GlobalVariables.SPILShapeWidth <> "" And SPIL.Shapes.GlobalVariables.SPILShapeName <> "" Then
                        e.Cell.Row.Cells("Height").Value = CInt(SPIL.Shapes.GlobalVariables.SPILShapeHeight)
                        e.Cell.Row.Cells("Width").Value = CInt(SPIL.Shapes.GlobalVariables.SPILShapeWidth)
                        e.Cell.Row.Cells("Shape").Value = CStr(SPIL.Shapes.GlobalVariables.SPILShapeName)

                        'For Template Items like IGU and Custom Lam
                        If e.Cell.Row.ChildBands.HasChildRows = True Then
                            ''UpdateParametersOnChildRows(e.Cell.Row)
                        End If
                        'end of.. For Template Items like IGU and Custom Lam

                        'oPriceUnits.SetUnitsAndVolumeOnThisRow(e.Cell.Row)
                        'oPriceUnits.CalculateLineAmountsOnthisRow(e.Cell.Row)

                        ''Call SetUnitsAndPriceOnExistingServices(e.Cell.Row)

                        '' Call CalculateOrderTotalAmount()


                    End If

                Else '>> Perfect Cut 

                    'Delete files
                    If My.Computer.FileSystem.FileExists(strAppPath & "\Shape\Image\" & "PC_Credted_PNG_1.png") Then
                        My.Computer.FileSystem.DeleteFile(strAppPath & "\Shape\Image\" & "PC_Credted_PNG_1.png")
                    End If
                    If My.Computer.FileSystem.FileExists(strAppPath & "\Shape\Temp\" & "PC_CredtedXML_1.xml") Then
                        My.Computer.FileSystem.DeleteFile(strAppPath & "\Shape\Temp\" & "PC_CredtedXML_1.xml")
                    End If
                    If My.Computer.FileSystem.FileExists(strAppPath & "\Shape\Temp\" & "PC_CredtedSAX_1.sax") Then
                        My.Computer.FileSystem.DeleteFile(strAppPath & "\Shape\Temp\" & "PC_CredtedSAX_1.sax")
                    End If
                    'Delete files

                    SPIL.Shapes.SecondaryDBConn.secondaryConnectionString = "Server=" & strSVR_Name & ";Database=" & strSVR_DBName & ";User ID=" & strSVR_UserName & ";Password=" & strSVR_PW & ""


                    Dim oShape As New SPIL.Shapes.frmShapeEditPerfectCut

                    'strCon = "Server=" & strSVR_Name & ";Database=" & strSVR_DBName & ";User ID=" & strSVR_UserName & ";Password=" & strSVR_PW & ""
                    'CRM_ConnectionString = "Data Source=" & strSVR_Name & ";Initial Catalog=" & strSVR_DBName & ";Persist " & _
                    '    "Security Info=True;User ID=" & strSVR_UserName & ";Password=" & strSVR_PW & ""



                    If Not e.Cell.Row.Cells("Shape").Value Is Nothing And Not IsDBNull(e.Cell.Row.Cells("Shape").Value) Then
                        SPIL.Shapes.GlobalVariables.SPILShapeName = CStr(e.Cell.Row.Cells("Shape").Value)
                        SPIL.Shapes.GlobalVariables.SPILShapeHeight = CStr(e.Cell.Row.Cells("Height").Value)
                        SPIL.Shapes.GlobalVariables.SPILShapeWidth = CStr(e.Cell.Row.Cells("Width").Value)
                    Else
                        SPIL.Shapes.GlobalVariables.SPILShapeName = ""
                        SPIL.Shapes.GlobalVariables.SPILShapeHeight = ""
                        SPIL.Shapes.GlobalVariables.SPILShapeWidth = ""
                    End If


                    If pubMeSpilDocTypeID = GlassDocTypes.SalesOrder Or pubMeSpilDocTypeID = GlassDocTypes.Quotation Then
                        If e.Cell.Row.Cells("iInvDetailID").Value = 0 Then

                            ''Copy As new Sales Order Modification
                            SPIL.Shapes.GlobalVariables.SPILShapeOrderIndex = 0
                            SPIL.Shapes.GlobalVariables.SPILShapeiInvDetailID = CInt(e.Cell.Row.Index)
                            SPIL.Shapes.GlobalVariables.SPILShapeUniqueLN = CInt(e.Cell.Row.Index)
                        Else
                            SPIL.Shapes.GlobalVariables.SPILShapeOrderIndex = pubMeOrderIndex
                            SPIL.Shapes.GlobalVariables.SPILShapeiInvDetailID = CInt(e.Cell.Row.Index)
                            SPIL.Shapes.GlobalVariables.SPILShapeUniqueLN = CInt(e.Cell.Row.Index)
                        End If

                    Else
                        If pubMeSpilDocTypeID = GlassDocTypes.NCR Then

                            If e.Cell.Row.Cells("iInvDetailID").Value = 0 Then
                                If Not IsNothing(e.Cell.Row.Cells("ShapeDetails").Value) Then
                                    SPIL.Shapes.GlobalVariables.SPILShapeOrderIndex = pubMeOrderIndex
                                    SPIL.Shapes.GlobalVariables.SPILShapeiInvDetailID = CInt(e.Cell.Row.Index)
                                    SPIL.Shapes.GlobalVariables.SPILShapeUniqueLN = CInt(e.Cell.Row.Index)
                                Else
                                    SPIL.Shapes.GlobalVariables.SPILShapeOrderIndex = pubMeOrderIndex
                                    SPIL.Shapes.GlobalVariables.SPILShapeiInvDetailID = CInt(e.Cell.Row.Index)
                                    SPIL.Shapes.GlobalVariables.SPILShapeUniqueLN = CInt(e.Cell.Row.Index)
                                End If


                            Else
                                SPIL.Shapes.GlobalVariables.SPILShapeOrderIndex = pubMeOrderIndex
                                SPIL.Shapes.GlobalVariables.SPILShapeiInvDetailID = CInt(e.Cell.Row.Index)
                                SPIL.Shapes.GlobalVariables.SPILShapeUniqueLN = CInt(e.Cell.Row.Index)
                            End If

                        End If
                    End If

                    oShape.StartPosition = FormStartPosition.CenterScreen
                    Dim IsEdit As Boolean = False
                    Dim Result As DialogResult
                    If IsNothing(e.Cell.Row.Cells("ShapeDetails").Value) Or IsDBNull(e.Cell.Row.Cells("ShapeDetails").Value) Then
                        Result = oShape.ShowDialog()
                        'Dim Status As Integer = oShape.SelectShape()
                        'If Status > 0 Then
                        '    If e.Cell.Row.Cells("ShapeDetails").Value.GetType() Is GetType(PCShapeDetails) Then
                        '        oShape.ShapeDetails = e.Cell.Row.Cells("ShapeDetails").Value
                        '    End If
                        '    oShape.ShowDialog()
                        'Else
                        '    oShape.ShowDialog()
                        'End If
                    ElseIf e.Cell.Row.Cells("Shape").Value = String.Empty Then
                        Result = oShape.ShowDialog()
                        'Dim Status As Integer = oShape.SelectShape()
                        'If Status > 0 Then
                        '    If e.Cell.Row.Cells("ShapeDetails").Value.GetType() Is GetType(PCShapeDetails) Then
                        '        oShape.ShapeDetails = e.Cell.Row.Cells("ShapeDetails").Value
                        '    End If
                        '    oShape.ShowDialog()
                        'Else
                        '    oShape.ShowDialog()
                        'End If
                    Else
                        If e.Cell.Row.Cells("ShapeDetails").Value.GetType() Is GetType(PCShapeDetails) Then
                            oShape.ShapeDetails = e.Cell.Row.Cells("ShapeDetails").Value
                            IsEdit = True
                        End If
                        Result = oShape.ShowDialog()
                    End If

                    If Result = Windows.Forms.DialogResult.OK And Not IsNothing(oShape.ShapeDetails) Then
                        e.Cell.Row.Cells("ShapeDetails").Value = oShape.ShapeDetails
                    ElseIf Result = Windows.Forms.DialogResult.No Then
                        e.Cell.Row.Cells("ShapeDetails").Value = New Object
                        oShape.ShapeDetails = Nothing
                        SPIL.Shapes.GlobalVariables.SPILShapeName = ""
                        SPIL.Shapes.GlobalVariables.SPILShapeHeight = ""
                        SPIL.Shapes.GlobalVariables.SPILShapeWidth = ""
                    ElseIf Result = Windows.Forms.DialogResult.Cancel And IsEdit = False Then
                        e.Cell.Row.Cells("ShapeDetails").Value = New Object
                        oShape.ShapeDetails = Nothing
                        SPIL.Shapes.GlobalVariables.SPILShapeName = ""
                        SPIL.Shapes.GlobalVariables.SPILShapeHeight = ""
                        SPIL.Shapes.GlobalVariables.SPILShapeWidth = ""
                    End If


                    e.Cell.Row.Cells("isShapeAttached").Value = False
                    e.Cell.Row.Cells("Shape").Value = String.Empty
                    e.Cell.Row.Cells("Shape").Appearance.BackColor = Nothing
                    e.Cell.Row.Cells("Shape").ButtonAppearance.Image = Nothing
                    If SPIL.Shapes.GlobalVariables.SPILShapeHeight <> "" And SPIL.Shapes.GlobalVariables.SPILShapeWidth <> "" And SPIL.Shapes.GlobalVariables.SPILShapeName <> "" Then
                        If CInt(SPIL.Shapes.GlobalVariables.SPILShapeHeight) > 0 Then e.Cell.Row.Cells("Height").Value = CInt(SPIL.Shapes.GlobalVariables.SPILShapeHeight)
                        If CInt(SPIL.Shapes.GlobalVariables.SPILShapeWidth) > 0 Then e.Cell.Row.Cells("Width").Value = CInt(SPIL.Shapes.GlobalVariables.SPILShapeWidth)
                        e.Cell.Row.Cells("Shape").Value = CStr(SPIL.Shapes.GlobalVariables.SPILShapeName)
                        e.Cell.Row.Cells("Shape").Value = "Added"
                        e.Cell.Row.Cells("Shape").Appearance.BackColor = Color.LightBlue
                        e.Cell.Row.Cells("Shape").ButtonAppearance.Image = Global.SPIL_Glass.My.Resources.Resources.shapesadd
                        e.Cell.Row.Cells("isShapeAttached").Value = True
                        'For Template Items like IGU and Custom Lam
                        If e.Cell.Row.HasChild Then
                            If e.Cell.Row.ChildBands.HasChildRows = True Then
                                'UpdateParametersOnChildRows(e.Cell.Row)
                            End If
                        End If
                        'end of.. For Template Items like IGU and Custom Lam

                        'oPriceUnits.SetUnitsAndVolumeOnThisRow(e.Cell.Row)
                        'oPriceUnits.CalculateLineAmountsOnthisRow(e.Cell.Row)

                        'Call SetUnitsAndPriceOnExistingServices(e.Cell.Row)

                        'Call CalculateOrderTotalAmount()
                        '' DisableProcessedLineFeilds(e.Cell.Row)
                    End If


                End If

            End If
        Catch ex As Exception
            modGlazingQuoteExtension.GQShowMessage(ex.Message, Me.Text, MsgBoxStyle.Critical)

        End Try
    End Sub


    Public Function ByteToImage(ByVal blob As Byte()) As Bitmap
        Dim mStream As MemoryStream = New MemoryStream()
        Dim bm As Bitmap = Nothing
        If blob.Length > 64 Then
            Dim pData As Byte() = blob
            mStream.Write(pData, 0, Convert.ToInt32(pData.Length))
            bm = New Bitmap(mStream, False)
            mStream.Dispose()
        End If

        Return bm
    End Function

    Private Sub tsmiSavetext_Click(sender As Object, e As EventArgs) Handles tsmiSavetext.Click
        Try
            Dim newSqlQuery As String
            newSqlQuery = "Insert INTO GlzQuote_Texts_Master (TextTypeID,Text) values(" & Me.UG2.ActiveRow.Cells("QuoteFiedType").Value & ", '" & Me.UG2.ActiveCell.Text & "')"
            Dim obSQL = New clsSqlConn
            obSQL.GET_INSERT_UPDATE(newSqlQuery)

        Catch ex As Exception
            modGlazingQuoteExtension.GQShowMessage(ex.Message, Me.Text, MsgBoxStyle.Critical)

        End Try
    End Sub

    Private Sub cmsQuoteGide_Opening(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles cmsQuoteGide.Opening
        Try
            QuoteGrideDropDownReset()
            If utcQuoteGrids.ActiveTab.Index = 0 Then
                If UG2.Selected.Rows.Count > 0 Then
                    cmsQuoteGide.Items("tsmAdd").Visible = True
                    cmsQuoteGide.Items("tsmiCopy").Visible = True
                    cmsQuoteGide.Items("tsmiPaste").Visible = True
                    cmsQuoteGide.Items("tsmiDel").Visible = True

                Else
                    cmsQuoteGide.Items("tsmAddRow").Visible = True

                End If

                If IsNothing(UG2.ActiveCell) = False Then

                    If UG2.ActiveCell.Column.Key = "LineComments" Then
                        If UG2.ActiveCell.IsInEditMode = True Then
                            If UG2.ActiveCell.SelText.Length > 0 Then
                                cmsQuoteGide.Items("tsmiSavetext").Visible = True
                                cmsQuoteGide.Items("tsmAddRow").Visible = False

                            ElseIf UG2.ActiveCell.SelText.Length = 0 Then
                                cmsQuoteGide.Items("tsmAddTotalAmount").Visible = True

                            End If



                        End If
                    ElseIf UG2.ActiveCell.Column.Key = "ItemImage" And UG2.ActiveRow.Cells("isImageAttached").Value = True Then
                        cmsQuoteGide.Items("tsmRemovePicture").Visible = True

                    End If
                End If
                If Clipboard.ContainsData("XML Spreadsheet") = True Then
                    CopyFromExcelToolStripMenuItem.Visible = True
                Else
                    CopyFromExcelToolStripMenuItem.Visible = False
                End If
            End If
        Catch ex As Exception
            modGlazingQuoteExtension.GQShowMessage(ex.Message, Me.Text, MsgBoxStyle.Critical)

        End Try
    End Sub

    Private Sub tsbExit_Click(sender As Object, e As EventArgs) Handles tsbExit.Click
        Me.Dispose()
        Me.Close()
    End Sub

    Private Sub mnuClose_Click(sender As Object, e As EventArgs) Handles mnuClose.Click
        Me.Dispose()
        Me.Close()
    End Sub

    Private Sub utxtNoteText_TextChanged(sender As Object, e As EventArgs) Handles utxtNoteText.TextChanged
        If utxtNoteText.Text <> "" And utxtNoteText.Text <> Nothing Then
            btnFooter.Text = "Footer Area - Enabled"
            btnFooter.FlatAppearance.BorderColor = Color.Green
            btnFooter.BackColor = Color.Green
        Else
            btnFooter.Text = " Footer Area - Disabled"
            btnFooter.FlatAppearance.BorderColor = Color.FromArgb(71, 164, 248)
            btnFooter.BackColor = Color.FromArgb(71, 164, 248)
        End If
    End Sub


    Private Function OrderValidation() As Integer
        Try

            Dim sErrorMsg As String = ""

            If cmbAccount.Value = 0 Then
                sErrorMsg = sErrorMsg + vbCrLf + "Select Customer"
            End If

            If cmbFacility.Value = 0 Then
                sErrorMsg = sErrorMsg + vbCrLf + "Please select the Branch"
                cmbFacility.Enabled = True
            End If

            'If sAllowBlankCustomerOrderNo = False Then
            '    If txtCustOrdNo.Text.Trim = "" Then
            '        sErrorMsg = sErrorMsg + vbCrLf + "Customer order number can not be left blank"

            '    End If
            'End If

            Dim row As UltraGridRow
            Dim emptyLine As Boolean = False
            Dim count As Integer = 0
DeleteRow:
            For Each row In Me.UG2.Rows
                If row.Cells("QuoteFiedType").Value = "" Then
                    'row.Appearance.BackColor = Color.LightCoral
                    'emptyLine = True
                    'count = count + 1
                    row.Delete()
                    GoTo DeleteRow
                End If
            Next
            Dim messageDiff As String = ""
            If emptyLine = True Then
                If count > 1 Then
                    messageDiff = "There are rows with an empty "
                Else
                    messageDiff = "There is a row with an empty "

                End If
                sErrorMsg = sErrorMsg & vbCrLf & messageDiff & Me.UG2.DisplayLayout.Bands(0).Columns("QuoteFiedType").Header.Caption
            End If

            If sErrorMsg = "" Then
                Return 1

            Else
                modGlazingQuoteExtension.GQShowMessage(sErrorMsg, "Form Validation", MsgBoxStyle.Exclamation)
                Return 0

            End If



        Catch ex As Exception
            modGlazingQuoteExtension.GQShowMessage(ex.Message, "Form Validation", MsgBoxStyle.Exclamation)
            Return 0
        End Try
    End Function


    Private Sub UG2_InitializeRow(sender As Object, e As InitializeRowEventArgs) Handles UG2.InitializeRow

    End Sub

    Private Sub UG2_DoubleClickCell(sender As Object, e As DoubleClickCellEventArgs) Handles UG2.DoubleClickCell
        If e.Cell.Column.Key = ("ItemImage") Then
            'e.Cell.Value = ""
            'e.Cell.CancelUpdate()
            '+e.Cell.Style = ColumnStyle.Image
            OpenFileDialogBox()
            e.Cell.Row.PerformAutoSize()
            Me.UG2.PerformAction(Infragistics.Win.UltraWinGrid.UltraGridAction.EnterEditMode, False, False)
            If IsNothing(pbUG2ItemPic.Image) = False Then
                e.Cell.Row.Cells("ItemImageByteArray").Value = New GlassInventoryModule.frmInventoryItem().ImageToByteArray(CType(pbUG2ItemPic.Image, Bitmap))
            End If
        End If
    End Sub

    Private Sub OpenFileDialogBox()
        Try
            Dim dlg As OpenFileDialog = New OpenFileDialog()
            dlg.Filter = ""
            Dim codecs As ImageCodecInfo() = ImageCodecInfo.GetImageEncoders()
            Dim sep As String = String.Empty
            For Each c In codecs
            Next

            dlg.Filter = String.Format("{0}{1}{2} ({3})|{3}", dlg.Filter, sep, "Image files", "*.JPEG;*.JPG;*.JFIF;*.JPEG2000;*.Exif;*.TIFF;*.GIF;*.*.BMP;*.PNG;*.PPM; *.PGM; *.PBM; *.PNM;*.WebP;*.HEIF;*.BAT;*.BPG;")
            dlg.DefaultExt = ".jpg"
            dlg.ShowDialog()
            Dim fileName As String = dlg.FileName
            If fileName = "" OrElse fileName Is Nothing Then
            Else
                If ValidFile(dlg.FileName, 102400) Then
                    Me.UG2.ActiveCell.Value = Image.FromFile(fileName)
                    pbUG2ItemPic.Image = Image.FromFile(fileName)
                    ItemImage = Nothing

                Else
                    Me.UG2.ActiveCell.Value = Image.FromStream(CompreddedImageToByteArray(Image.FromFile(fileName), 50))
                    ItemImage = Nothing
                    pbUG2ItemPic.Image = Image.FromStream(CompreddedImageToByteArray(Image.FromFile(fileName), 50))

                End If

                Me.UG2.ActiveRow.Cells("isImageAttached").Value = True

            End If

        Catch ex As Exception


        End Try
    End Sub

    Public Function ImageToByteArray(ByVal image As System.Drawing.Image) As Byte()

        Dim byteArea As Byte()
        Using msX As New IO.MemoryStream
            pbUG2ItemPic.Image.Save(msX, Imaging.ImageFormat.Bmp)
            byteArea = msX.ToArray()
        End Using



        'Dim format As ImageFormat = image.RawFormat
        'Dim codec As ImageCodecInfo = ImageCodecInfo.GetImageDecoders().(Function(c) c.FormatID = format.Guid)
        'Dim mimeType As String = codec.MimeType
        'Using ms = New MemoryStream()
        '    image.Save(ms, image.RawFormat)
        '   Return ms.ToArray()
        UG2.ActiveRow.Cells("ItemImageByteArray").Value = byteArea
        Return byteArea
        'End Using
    End Function

    Private Function ValidFile(ByVal filename As String, ByVal limitInBytes As Long) As Boolean
        Dim fileSizeInBytes = New FileInfo(filename).Length
        If fileSizeInBytes > limitInBytes Then
            Return False
        Else
            Return True
        End If
    End Function

    Public Function CompreddedImageToByteArray(ByVal img As Image, ByVal quality As Integer) As MemoryStream
        If quality < 0 OrElse quality > 100 Then Throw New ArgumentOutOfRangeException("quality must be between 0 and 100.")
        Dim qualityParam As EncoderParameter = New EncoderParameter(System.Drawing.Imaging.Encoder.Quality, quality)
        Dim jpegCodec As ImageCodecInfo = GetEncoderInfo("image/jpeg")
        Dim encoderParams As EncoderParameters = New EncoderParameters(1)
        encoderParams.Param(0) = qualityParam
        Dim ms = New MemoryStream()
        img.Save(ms, jpegCodec, encoderParams)
        Return ms

    End Function

    Private Shared Function GetEncoderInfo(ByVal mimeType As String) As ImageCodecInfo
        Dim codecs As ImageCodecInfo() = ImageCodecInfo.GetImageEncoders()
        For i As Integer = 0 To codecs.Length - 1
            If codecs(i).MimeType = mimeType Then Return codecs(i)
        Next

        Return Nothing
    End Function

    Private Sub cmCustAddNew_Click(sender As Object, e As EventArgs) Handles cmCustAddNew.Click
        Dim frmCustomerObj As New frmCustomer()
        frmCustomerObj.pubIsSaveNew = True
        frmCustomerObj.CID = String.Empty
        frmCustomerObj.txtCustID.Text = 0

        If cmbCusType.Value = 0 Then
            frmCustomerObj.cmbCustType.Value = 1

        Else
            frmCustomerObj.cmbCustType.Value = 2

        End If
        frmCustomerObj.ShowDialog()
        If frmCustomerObj.pubCustID <> 0 Then
            Call GET_CUSTOMERS(cmbCusType.Text)
            cmbAccount.Value = frmCustomerObj.pubCustID
            frmCustomerObj.Dispose()
        End If
    End Sub


    Private Sub cmCustEdit_Click(sender As Object, e As EventArgs) Handles cmCustEdit.Click
        If cmbAccount.Value <> 0 Then
            Dim custID As Integer = cmbAccount.Value
            frmCustomer.pubIsSaveNew = False
            frmCustomer.CID = cmbAccount.ActiveRow.Cells("Account").Value
            frmCustomer.pubCustID = cmbAccount.Value
            frmCustomer.StartPosition = FormStartPosition.CenterParent
            frmCustomer.ShowDialog()
            cmbAccount.Enabled = True
            Call GET_CUSTOMERS(cmbCusType.Text)
            cmbAccount.Value = custID

        End If
    End Sub

    Private Sub cmCustView_Click(sender As Object, e As EventArgs) Handles cmCustView.Click
        If cmbAccount.Value <> 0 Then
            Dim custID As Integer = cmbAccount.Value

            frmCustomer.pubIsSaveNew = False
            frmCustomer.CID = cmbAccount.ActiveRow.Cells("Account").Value
            frmCustomer.cmdSave.Enabled = False
            frmCustomer.pubCustID = cmbAccount.Value
            frmCustomer.StartPosition = FormStartPosition.CenterParent
            frmCustomer.ShowDialog()

            cmbAccount.Enabled = True
            Call GET_CUSTOMERS(cmbCusType.Value)
            cmbAccount.Value = custID

        End If
    End Sub

    Private Sub cmSameAsPostalAdd_Click(sender As Object, e As EventArgs) Handles cmSameAsPostalAdd.Click
        txtPhy1.Text = txtPost1.Text
        txtPhy2.Text = txtPost2.Text
        txtPhy3.Text = txtPost3.Text
        txtPhy4.Text = txtPost4.Text
        txtPhy5.Text = txtPost5.Text
        txtPhyPostCode.Text = txtPostCode.Text
    End Sub


    Private Sub btnJobDescription_Click(sender As Object, e As EventArgs) Handles btnJobDescription.Click
        OpenDescriptionState(True, jobDescription)
        SetJobDescriptionState()
    End Sub

    Public Sub OpenDescriptionState(ByRef sendObject As Boolean, ByRef jobDescriptionText As String)
        Dim newGlazingJobDescription As frmGlazingNote
        Try
            If sendObject = True Then
                newGlazingJobDescription = New frmGlazingNote(Me)

            Else
                newGlazingJobDescription = New frmGlazingNote()
                newGlazingJobDescription.utxtNoteText.Dock = DockStyle.Fill
                newGlazingJobDescription.utxtNoteText.ReadOnly = True
                newGlazingJobDescription.utxtNoteText.SpellChecker = Nothing

            End If
            newGlazingJobDescription.isJobDescriptionActive = True
            newGlazingJobDescription.utxtNoteText.Value = jobDescriptionText
            newGlazingJobDescription.ShowDialog()

        Catch ex As Exception
            modGlazingQuoteExtension.GQShowMessage(ex.Message, Me.Text, MsgBoxStyle.Critical)

        Finally
            SetJobDescriptionButtonState()
            newGlazingJobDescription = Nothing

        End Try
    End Sub

    Public Sub SetJobDescriptionState()
        If IsNothing(jobDescription) = False Then
            If jobDescription <> "" Then
                btnJobDescription.Text = "Job description - Active"
                isJobDescriptionActive = True
                btnJobDescription.BackColor = Color.Green
                btnJobDescription.FlatAppearance.BorderColor = Color.Green

            Else
                btnJobDescription.Text = "Job description"
                btnJobDescription.BackColor = Color.DodgerBlue
                btnJobDescription.FlatAppearance.BorderColor = Color.DodgerBlue

            End If
        End If
    End Sub

    Private Sub btnExpandMap_Click()
        Try
            Dim GlazingAddessLocatorGmap As frmGlazingAddessLocatorGmap = New frmGlazingAddessLocatorGmap(Me.latitude, Me.longitude)
            Dim addressInSingleLine = CreateOnelineAddress()
            If addressInSingleLine <> "" Then
                GlazingAddessLocatorGmap.utxtAddress.Value = addressInSingleLine
                GlazingAddessLocatorGmap.GetAddress(addressInSingleLine)
            End If

            If GlazingAddessLocatorGmap.ShowDialog() = Windows.Forms.DialogResult.Cancel Then
                Exit Sub
            End If

            Dim placeMark As List(Of Result) = GlazingAddessLocatorGmap.ReturnNewAddress()
            If placeMark IsNot Nothing Then
                'txtPhy2.Text = placeMark.Item(0).address_components.Item(1).long_name
                'txtPhy3.Text = placeMark.Item(0).address_components.Item(2).long_name
                'txtPhy4.Text = placeMark.Item(0).address_components.Item(3).long_name
                'txtPhy5.Text = placeMark.Item(0).address_components.Item(5).long_name
                'txtPhyPostCode.Text = placeMark.Item(0).address_components.Item(7).long_name

                For Each addressItem As AddressComponent In placeMark.Item(0).address_components

                    For Each itemType As String In addressItem.types

                        Select Case itemType

                            Case "premise"
                                txtPhy1.Text = addressItem.long_name

                            Case "street_number"
                                If placeMark.Item(0).address_components.Item(0).types(0) = "premise" Then
                                    txtPhy1.Text = placeMark.Item(0).address_components.Item(0).long_name + addressItem.long_name
                                Else
                                    txtPhy1.Text = addressItem.long_name
                                End If

                            Case "route"
                                txtPhy2.Text = addressItem.long_name

                            Case "sublocality_level_1"
                                txtPhy2.Text = txtPhy2.Text & ", " & addressItem.long_name

                            Case "locality"
                                txtPhy3.Text = addressItem.long_name

                            Case "administrative_area_level_2"
                                txtPhy4.Text = addressItem.long_name

                            Case "administrative_area_level_1"
                                txtPhy5.Text = addressItem.short_name

                            Case "country"

                            Case "postal_code"
                                txtPhyPostCode.Text = addressItem.long_name

                        End Select

                    Next
                Next

                'Dim overlay As GMapOverlay = New GMapOverlay("Address Overlay")

                'Dim point As PointLatLng? = FindLatLngByAddress(placeMark.Item(placeMark.Count - 1).formatted_address)

                'If point IsNot Nothing Then
                '    Dim marker As GMapMarker = New GMarkerGoogle(New PointLatLng(point.Value.Lat, point.Value.Lng), GMarkerGoogleType.blue)
                '    marker.ToolTipMode = MarkerTooltipMode.OnMouseOver
                '    addressMap.Position = New PointLatLng(point.Value.Lat, point.Value.Lng)
                '    overlay.Markers.Add(marker)
                '    Me.latitude = point.Value.Lat
                '    Me.longitude = point.Value.Lng
                'End If
                'addressMap.Overlays.Clear()
                'addressMap.Overlays.Add(overlay)
            End If
        Catch ex As Exception
            modGlazingQuoteExtension.GQShowMessage(ex.Message, Me.Text, MsgBoxStyle.Critical)
        End Try
    End Sub

    Private Sub btnAddressLocator_Click(sender As Object, e As EventArgs) Handles btnAddressLocator.Click
        btnExpandMap_Click()
    End Sub

    Public Function FindLatLngByAddress(ByVal address As String) As PointLatLng?
        Dim address64 As String = EncodeTo64(address)
        Dim status As GeoCoderStatusCode
        Dim point As PointLatLng? = GMapProviders.GoogleMap.GetPoint(address, status)
        If status = GeoCoderStatusCode.G_GEO_SUCCESS AndAlso point IsNot Nothing Then
            Return point

        End If

        Return Nothing
    End Function

    Public Shared Function EncodeTo64(ByVal toEncode As String) As String
        Dim toEncodeAsBytes As Byte() = System.Text.ASCIIEncoding.ASCII.GetBytes(toEncode)
        Dim returnValue As String = System.Convert.ToBase64String(toEncodeAsBytes)
        Return returnValue
    End Function

    Private Sub btnAddressLocator_MouseEnter(sender As Object, e As EventArgs) Handles btnAddressLocator.MouseEnter
        addressMap.Visible = False
    End Sub

    Private Sub btnAddressLocator_MouseLeave(sender As Object, e As EventArgs) Handles btnAddressLocator.MouseLeave
        addressMap.Visible = False
    End Sub

    Private Sub tsmRemovePicture_Click(sender As Object, e As EventArgs) Handles tsmRemovePicture.Click
        UG2.ActiveRow.Cells("ItemImage").Value = DBNull.Value
        UG2.ActiveRow.Cells("ItemImageByteArray").Value = ""
        UG2.ActiveRow.Cells("isImageAttached").Value = False

    End Sub

    Private Sub UG2_AfterEnterEditMode(sender As Object, e As EventArgs) Handles UG2.AfterEnterEditMode
        Me.UG2.ActiveRow.Appearance.BackColor = Nothing
    End Sub

    Sub SetJobDescriptionButtonState()
        Try
            If IsNothing(jobDescription) = True Then
                If jobDescription = "" Then
                    btnJobDescription.Text = "Job description (Inactive)"

                Else
                    btnJobDescription.Text = "Job description (Active)"

                End If
            End If

        Catch ex As Exception

        End Try

    End Sub

    Sub SetSubTotalByGroup()
        Dim ug2Row As UltraGridRow
        Dim groupID As Integer = 0.0
        Dim subtotal As Double = 0.0
        Dim sumOfsubtotal As Double = 0.0

        Dim TotalExcAmount As Double = 0.0
        Dim TotalTaxedtAmount As Double = 0.0
        Dim TotalIncAmount As Double = 0.0

        Try
            For Each ug2Row In UG2.Rows
                If IsNothing(ug2Row.Cells("ItmGroupID").Value) = False Then
                    groupID = ug2Row.Cells("ItmGroupID").Value

                End If

                If ug2Row.Cells("QuoteFiedType").Value = QuateFiedTypesList.Subtotal Then
                    ug2Row.Cells("Amount").Value = Math.Round(subtotal, 2, MidpointRounding.AwayFromZero)

                ElseIf ug2Row.Cells("QuoteFiedType").Value = QuateFiedTypesList.Text Or ug2Row.Cells("QuoteFiedType").Value = QuateFiedTypesList.Stock_Item Then
                    If groupID = ug2Row.Cells("ItmGroupID").Value Then
                        subtotal = subtotal + ug2Row.Cells("Amount").Value
                        sumOfsubtotal = sumOfsubtotal + subtotal

                    Else
                        groupID = ug2Row.Cells("ItmGroupID").Value
                        subtotal = 0

                    End If
                    TotalExcAmount = TotalExc + ug2Row.Cells("ItmExcAmount").Text
                    TotalTaxedtAmount = TotalTax + ug2Row.Cells("Tax").Text
                    TotalIncAmount = TotalInc + ug2Row.Cells("Net").Text

                    If sumOfsubtotal <> TotalIncAmount Then
                        'modGlazingQuoteExtensionClass.GQShowMessage("Error occered while calculating Total Insive Amount", Me.Text, MsgBoxStyle.Critical)
                        'lblTotIncAmo.Text = "0"
                        'Exit Sub
                    End If
                End If
            Next

            'lblTotExcAmo.Text = TotalExc
            'lblTotVatAmo.Text = TotalTax
            'lblTotIncAmo.Text = TotalInc

        Catch ex As Exception
            modGlazingQuoteExtension.GQShowMessage(ex.Message, Me.Text, MsgBoxStyle.Critical)

        End Try
    End Sub

    Sub SetActiveCell(e As CellEventArgs)
        Try
            If e.Cell.Column.Key = "QuoteFiedType" Then
                If e.Cell.Value = QuateFiedTypesList.Text Or e.Cell.Value = QuateFiedTypesList.Stock_Item Or e.Cell.Value = QuateFiedTypesList.Header_Main Or e.Cell.Value = QuateFiedTypesList.Header_Sub Then
                    Me.UG2.ActiveRow.Cells("LineComments").Activated = True

                ElseIf e.Cell.Value = QuateFiedTypesList.Subtotal Then
                    'Me.UG2.ActiveRow.Cells("Amount").Activated = True

                End If

            End If

        Catch ex As Exception
            modGlazingQuoteExtension.GQShowMessage(ex.Message, Me.Text, MsgBoxStyle.Critical)

        End Try
    End Sub


    Private Sub ucmbQuoteLineType_InitializeLayout(sender As Object, e As InitializeLayoutEventArgs) Handles ucmbQuoteLineType.InitializeLayout
        Try
            ucmbQuoteLineType.DisplayLayout.Bands(0).ColHeadersVisible = False

        Catch ex As Exception

        End Try
    End Sub

    Private Sub ucmbQuoteLineType_RowSelected(sender As Object, e As RowSelectedEventArgs) Handles ucmbQuoteLineType.RowSelected
        If IsNothing(ucmbQuoteLineType.ActiveRow) = False Then
            lineTypeCell = ucmbQuoteLineType.ActiveRow.Cells("LineTypeName")
            gridActiveRow = Me.UG2.ActiveRow
            gridActiveCell = Me.UG2.ActiveCell
            'Dim YESY = UG2.ActiveCell.Value
            'If ucmbQuoteLineType.ActiveRow.Cells("LineTypeName").Value = "Text" Then
            '    Me.UG2.ActiveCell = Me.UG2.ActiveRow.Cells("LineComments")
            '    Me.UG2.PerformAction(Infragistics.Win.UltraWinGrid.UltraGridAction.EnterEditMode, False, False)

            'Else
            '    Me.UG2.ActiveRow.Appearance.BackColor = Nothing

            'End If



            QuoteGirdRowStyling()
            If downKeyPressed = False And isOpeningQuote = False Then
                QuoteGridNavigator()
                QuoteGridFunctions()
                'If isStockItemActive = False Then
                '    If Me.UG2.ActiveCell.Column.Key = "QuoteFiedType" Then
                '        QuoteGirdRowFormat(ex)

                '    End If
            End If
        End If
    End Sub

    Sub QuoteGirdRowStyling(Optional ByRef isOpenning As Boolean = False)
        Try
            gridActiveRow.Appearance.Reset()
            gridActiveRow.Appearance.FontData.Reset()
            gridActiveRow.Cells("Amount").Appearance.Reset()
            If isOpenning = False Then
                Me.UG2.ActiveRow.Cells("LineComments").Hidden = False
                Me.UG2.ActiveRow.Cells("LineComments").Appearance.FontData.Reset()

            End If

            If lineTypeCell.Text = "Header-Main" Then
                gridActiveRow.Appearance.FontData.Bold = DefaultableBoolean.True
                gridActiveRow.Cells("LineComments").Appearance.FontData.SizeInPoints = 14
                GrideColumsVisibility(gridActiveRowe, True)

            ElseIf lineTypeCell.Text = "Header-Sub" Then
                gridActiveRow.Appearance.FontData.Bold = DefaultableBoolean.True
                gridActiveRow.Appearance.FontData.Underline = DefaultableBoolean.True
                gridActiveRow.Cells("LineComments").Appearance.FontData.SizeInPoints = 12
                GrideColumsVisibility(gridActiveRow, True)

            ElseIf lineTypeCell.Text = "Subtotal" Then
                gridActiveRow.Appearance.FontData.Bold = DefaultableBoolean.True
                Me.UG2.ActiveRow.Cells("Amount").Appearance.BackColor = Color.LightGreen
                Me.UG2.ActiveRow.Cells("Amount").Appearance.BorderColor = Color.Green
                GrideColumsVisibility(gridActiveRow, True)
                Me.UG2.ActiveRow.Cells("Amount").Hidden = False
                Me.UG2.ActiveRow.Cells("LineComments").Hidden = True

            ElseIf lineTypeCell.Text = "Stock Item" Then
                GrideColumsVisibility(gridActiveRow, False)

            ElseIf lineTypeCell.Text = "Text" Then
                GrideColumsVisibility(gridActiveRow, False)

            End If

            gridActiveRow.PerformAutoSize()
            Me.UG2.PerformAction(Infragistics.Win.UltraWinGrid.UltraGridAction.EnterEditMode, False, False)

        Catch ex As Exception
            modGlazingQuoteExtension.GQShowMessage(ex.Message, "Error in Line type drop down", MsgBoxStyle.Critical)

        End Try
    End Sub

    Sub GrideColumsVisibility(ByRef row As UltraGridRow, ByRef isHdden As Boolean)
        Try
            Me.UG2.ActiveRow.Cells("Height").Hidden = isHdden
            Me.UG2.ActiveRow.Cells("Width").Hidden = isHdden
            Me.UG2.ActiveRow.Cells("Qty").Hidden = isHdden
            Me.UG2.ActiveRow.Cells("Price").Hidden = isHdden
            Me.UG2.ActiveRow.Cells("Amount").Hidden = isHdden
            Me.UG2.ActiveRow.Cells("LineNotes").Hidden = isHdden
            Me.UG2.ActiveRow.Cells("MarkAs").Hidden = isHdden
            Me.UG2.ActiveRow.Cells("ItemImage").Hidden = isHdden
            Me.UG2.ActiveRow.Cells("TaxRate").Hidden = isHdden
            Me.UG2.ActiveRow.Cells("Shape").Hidden = isHdden

        Catch ex As Exception
            modGlazingQuoteExtension.GQShowMessage(ex.Message, Me.Text, MsgBoxStyle.Critical)

        End Try
    End Sub

    Sub QuoteGridNavigator()
        Try
            If lineTypeCell.Text = "Subtotal" Then

            Else
                gridActiveRow.Cells("LineComments").Activated = True

            End If
            gridActiveRow.PerformAutoSize()
            Me.UG2.PerformAction(Infragistics.Win.UltraWinGrid.UltraGridAction.EnterEditMode, False, False)
        Catch ex As Exception
            modGlazingQuoteExtension.GQShowMessage(ex.Message, Me.Text, MsgBoxStyle.Critical)

        End Try
    End Sub

    Public Enum LineTypeKey As Integer
        Text = 1
        HeaderMain = 2
        HeaderSub = 3
        Subtotal = 4
        StockItem = 5

    End Enum

    Sub QuoteGridFunctions()
        Dim amountKeeper As Double = 0
        Try
            If isExistingOrder = False Then
                If gridActiveRow.Cells("LineComments").Value <> "" Or gridActiveRow.Cells("Amount").Value <> 0 Then
                    If gridActiveCell.Text = "Header-Main" Or gridActiveCell.Text = "Header-Sub" Then
                        If canUpdate = True Then
                            CellValuesClear()
                        End If

                    ElseIf gridActiveCell.Text = "Header-Sub" Then
                        amountKeeper = gridActiveRow.Cells("Amount").Value
                        CellValuesClear()
                        gridActiveRow.Cells("Amount").Value = Math.Round(amountKeeper, 2, MidpointRounding.AwayFromZero)

                    End If
                End If
            End If


            If lineTypeCell.Text = "Header-Main" Then

            ElseIf lineTypeCell.Text = "Header-Sub" Then
                If groupStarted = False Then
                    groupStarted = True
                    subHeaderID = subHeaderID + 1
                End If
                gridActiveRow.Cells("ItmGroupID").Value = subHeaderID

            ElseIf lineTypeCell.Text = "Text" Then

            ElseIf lineTypeCell.Text = "Stock Item" Then
                If IsNothing(UG2.ActiveRow) = False Then
                    If isPasting = False And isOpeningQuote = False Then
                        Dim glazingDocStockItem As New frmGlazingDocStockItem(Me)
                        glazingDocStockItem.ShowDialog()

                    End If
                    Me.UG2.ActiveRow.Cells("Amount").Tag = Me.UG2.ActiveRow.Cells("ItmGroupID").Value

                End If

            ElseIf lineTypeCell.Text = "Subtotal" Then

            Else
                gridActiveRow.Cells("ItmGroupID").Value = Me.UG2.Rows(gridActiveRow.Index - 1).Cells("ItmGroupID").Value
            End If

            QuoteGridSetSubTotal()
        Catch ex As Exception
            modGlazingQuoteExtension.GQShowMessage(ex.Message, Me.Text, MsgBoxStyle.Critical)

        End Try
    End Sub

    Sub QuoteGridSetSubTotal()
        Try
            Dim row As UltraGridRow
            Dim subTotal As Decimal = 0.0
            'Dim groupID As Integer = Me.UG2.ActiveRow.Cells("ItmGroupID").Value
            Dim exGroupAmount As Decimal = 0.0
            Dim isGroupstarted = False
            Dim TotalExc As Decimal = 0.0   'ItmExcAmount
            Dim TotalTax As Decimal = 0.0   'Tax
            Dim TotalInc As Decimal = 0.0   'Net
            Dim lineExc As Decimal = 0.0   'ItmExcAmount
            Dim lineTax As Decimal = 0.0   'Tax
            Dim lineInc As Decimal = 0.0   'Net
            Dim lineTaxSerivce As Decimal = 0.0   'Tax Serivce
            Dim TotalTaxSerivce As Decimal = 0.0   'TotalTax Serivce
            Dim calculateAll As Boolean = False
            Dim calcualteLineExc As Decimal = 0.0   ' Exc

            For Each row In Me.UG2.Rows
                'IF this is an item
                If IsNothing(row.Cells("QuoteFiedType").Text) = False Then
                    If row.Cells("QuoteFiedType").Text = "Header-Sub" Then
                        isGroupstarted = True

                        If calculateAll = False Then
                            subTotal = 0.0
                            lineExc = 0.0
                            lineTax = 0.0
                            lineInc = 0.0
                            lineTaxSerivce = 0.0
                            calcualteLineExc = 0.0
                        Else
                            TotalExc = TotalExc + lineExc
                            TotalTax = TotalTax + lineTax
                            TotalInc = TotalInc + lineInc
                            TotalTaxSerivce = TotalTaxSerivce + lineTaxSerivce
                        End If
                    End If

                    If row.Cells("QuoteFiedType").Text = "Text" Or row.Cells("QuoteFiedType").Text = "Stock Item" Then
                        If IsNothing(row.Cells("Amount").Text) = False Then
                            subTotal = subTotal + row.Cells("Amount").Text
                            TotalExc = TotalExc + row.Cells("ItmExcAmount").Text
                            calcualteLineExc = TotalExc + row.Cells("ItmExcAmount").Text
                            TotalTax = TotalTax + row.Cells("Tax").Text
                            TotalInc = TotalInc + row.Cells("Net").Text
                            TotalTaxSerivce = TotalTaxSerivce + row.Cells("ServiceItemTax").Text
                        End If
                    End If

                    If row.Cells("QuoteFiedType").Text = "Subtotal" Then

                        If isGroupstarted = False Then
                            'Calculate all item lines
                            If calculateAll = True Then
                                row.Cells("Amount").Value = Math.Round(subTotal, 2, MidpointRounding.AwayFromZero)
                                row.Cells("ItmExcAmount").Value = Math.Round(calcualteLineExc, 2, MidpointRounding.AwayFromZero)

                                TotalExc = TotalExc + lineExc
                                TotalTax = TotalTax + lineTax
                                TotalInc = TotalInc + lineInc
                                TotalTaxSerivce = TotalTaxSerivce + lineTaxSerivce

                            Else
                                subTotal = 0.0
                                lineExc = 0.0
                                lineTax = 0.0
                                lineInc = 0.0
                                lineTaxSerivce = 0.0
                                calcualteLineExc = 0.0
                            End If

                        Else
                            row.Cells("Amount").Value = Math.Round(subTotal, 2, MidpointRounding.AwayFromZero)
                            row.Cells("ItmExcAmount").Value = Math.Round(calcualteLineExc, 2, MidpointRounding.AwayFromZero)

                            TotalExc = TotalExc + lineExc
                            TotalTax = TotalTax + lineTax
                            TotalInc = TotalInc + lineInc
                            TotalTaxSerivce = TotalTaxSerivce + lineTaxSerivce

                            If calculateAll = False Then
                                subTotal = 0.0
                                lineExc = 0.0
                                lineTax = 0.0
                                lineInc = 0.0
                                lineTaxSerivce = 0.0
                                isGroupstarted = False
                                calcualteLineExc = 0.0

                            End If
                        End If
                        isGroupstarted = False

                    End If

                End If
                TotalExc = TotalExc + lineExc
                TotalTax = TotalTax + lineTax
                TotalInc = TotalInc + lineInc
                TotalTaxSerivce = TotalTaxSerivce + lineTaxSerivce

            Next
            lblTotExcAmo.Text = Format(TotalExc, "0.00")
            lblTotVatAmo.Text = Format(TotalTax + TotalTaxSerivce, "0.00")
            lblTotIncAmo.Text = Format(TotalInc, "0.00")
            ToolTip1.SetToolTip(lblTotalVat, "Main Item/s: " & TotalTax & "  |  Glass Service/s: " & TotalTaxSerivce)
            ToolTip1.SetToolTip(lblTotVatAmo, "Main Item/s: " & TotalTax & "  |  Glass Service/s: " & TotalTaxSerivce)


        Catch ex As Exception
            modGlazingQuoteExtension.GQShowMessage(ex.Message, Me.Text, MsgBoxStyle.Critical)

        End Try

    End Sub
    Function QuoteGridValueValidationBeforeUpdate() As Integer
        Dim isWrongType As Boolean = False
        Try
            If IsNothing(Me.UG2.ActiveCell) = False Then

                If UG2.ActiveCell.Column.DataType = GetType(Int32) Then
                    If Regex.IsMatch(Me.UG2.ActiveCell.Text, "^[0-9 ]+$") = False Then
                        isWrongType = True

                    End If

                ElseIf Me.UG2.ActiveCell.Column.DataType = GetType(Decimal) Then
                    If Regex.IsMatch(Me.UG2.ActiveCell.Text, "^[0-9 ]?[.]?^[0-9 ]") = False Then
                        isWrongType = True

                    End If
                End If

                If isWrongType = True Then
                    modGlazingQuoteExtension.GQShowMessage(UG2.ActiveCell.Column.Header.Caption & " required a numeric value", "Error in entered value", MsgBoxStyle.Critical)
                    Me.UG2.ActiveCell.Appearance.BackColor = Color.FromArgb(198, 101, 101)
                    Me.UG2.ActiveCell.SelectAll()
                    Return 0
                Else
                    Me.UG2.ActiveCell.Appearance.BackColor = Nothing

                End If
                Return 1
            Else
                Return 0
            End If
        Catch ex As Exception
            modGlazingQuoteExtension.GQShowMessage(ex.Message, Me.Text, MsgBoxStyle.Critical)
            Return 0
        End Try
    End Function

    Function QuoteGridValueValidationAfterUpdate() As Integer
        Try
            Me.UG2.ActiveRow.Cells("QuoteFiedType").Appearance.BackColor = Nothing
            If Me.UG2.ActiveRow.Cells("QuoteFiedType").Text = "" Then
                'Me.UG2.ActiveRow.Cells("QuoteFiedType").Appearance.BackColor = Color.FromArgb(198, 101, 101)
                Return 0
                Exit Function

            Else
                If IsNothing(UG2.ActiveRow.Cells("QuoteFiedType").Value) = False Then
                    If IsNumeric(UG2.ActiveRow.Cells("QuoteFiedType").Value) = False Then
                        'Me.UG2.ActiveRow.Cells("QuoteFiedType").Appearance.BackColor = Color.FromArgb(198, 101, 101)
                        Return 0
                        Exit Function
                    End If
                End If
            End If
            Return 1
        Catch ex As Exception
            modGlazingQuoteExtension.GQShowMessage(ex.Message, Me.Text, MsgBoxStyle.Critical)
            Return 0

        End Try
    End Function

    Private Sub UG2_MouseClick(sender As Object, e As MouseEventArgs) Handles UG2.MouseClick
        If IsNothing(Me.UG2.ActiveCell) = False Then
            If Me.UG2.ActiveCell.Column.Key = "QuoteFiedType" Then


            End If
        End If
    End Sub

    Private Sub ucmbQuoteLineType_AfterDropDown(sender As Object, e As EventArgs) Handles ucmbQuoteLineType.AfterDropDown
        downKeyPressed = False

    End Sub
    Sub KeyCombination(ByRef e As KeyEventArgs)
        If e.KeyCode = Keys.A AndAlso e.Modifiers = Keys.Control Then
            If IsNothing(Me.UG2.ActiveCell) = False Then
                If Me.UG2.ActiveCell.IsInEditMode = True Then
                    Me.UG2.ActiveCell.SelectAll()

                End If
            End If
        End If
        If e.KeyCode = Keys.C AndAlso e.Modifiers = Keys.Control Then
            CopyRow()

        End If
        If e.KeyCode = Keys.V AndAlso e.Modifiers = Keys.Control Then
            PastRow()

        End If
    End Sub

    Private Sub tsmAddTotalAmount_Click(sender As Object, e As EventArgs) Handles tsmAddTotalAmount.Click
        Try
            If IsNothing(Me.UG2.ActiveCell) = False Then
                Me.UG2.ActiveCell.Value = Me.UG2.ActiveCell.Text.Insert(Me.UG2.ActiveCell.SelStart, " <Total> ")

            End If

        Catch ex As Exception
            modGlazingQuoteExtension.GQShowMessage(ex.Message, Me.Text, MsgBoxStyle.Critical)

        End Try
    End Sub

    Sub QuoteGrideDropDownReset()
        cmsQuoteGide.Items("tsmAdd").Visible = False
        cmsQuoteGide.Items("tsmiCopy").Visible = False
        cmsQuoteGide.Items("tsmiPaste").Visible = False
        cmsQuoteGide.Items("tsmiDel").Visible = False
        cmsQuoteGide.Items("tsmiSavetext").Visible = False
        cmsQuoteGide.Items("tsmAddTotalAmount").Visible = False
        cmsQuoteGide.Items("tsmRemovePicture").Visible = False
        cmsQuoteGide.Items("tsmAddRow").Visible = False
    End Sub

    Private Sub tsmAddRow_Click(sender As Object, e As EventArgs) Handles tsmAddRow.Click
        AddNewRow("end")
    End Sub

    Function CreateOnelineAddress() As String
        Dim addressInSingleLine As String
        If txtPhy1.Text <> "" Then
            addressInSingleLine = txtPhy1.Text
        End If

        If txtPhy2.Text <> "" Then
            addressInSingleLine = addressInSingleLine + ", " + txtPhy2.Text
        End If

        If txtPhy3.Text <> "" Then
            addressInSingleLine = addressInSingleLine + ", " + txtPhy3.Text
        End If
        Return addressInSingleLine
    End Function

    Sub UpdateOrderHeaderWhileSaving(ByRef orderNumber As Integer)
        Try

        Catch ex As Exception
            modGlazingQuoteExtension.GQShowMessage(ex.Message, Me.Text, MsgBoxStyle.Critical)

        End Try
    End Sub

    Sub UpdateOrderLines()
        Try

        Catch ex As Exception
            modGlazingQuoteExtension.GQShowMessage(ex.Message, Me.Text, MsgBoxStyle.Critical)

        End Try
    End Sub

    Private Sub cmbCustProject_ValueChanged(sender As Object, e As EventArgs) Handles cmbCustProject.ValueChanged
        Try
            _ProjectId = cmbCustProject.SelectedRow.Cells("Id").Value

        Catch ex As Exception
            _ProjectId = 0
        End Try
        clsGQExtensionForJobCostingObj.GetStages(_ProjectId, True)
        clsGQExtensionForJobCostingObj.GetJobs(_ProjectId, cmbAccount.SelectedRow.Cells("DCLink").Value, True)

    End Sub

    Private Sub cmbCustJob_ValueChanged(sender As Object, e As EventArgs) Handles cmbCustJob.ValueChanged
        Try
            _JobId = cmbCustJob.SelectedRow.Cells("Id").Value
        Catch ex As Exception
            _JobId = 0
        End Try
    End Sub

    Private Sub cmbProjectStage_ValueChanged(sender As Object, e As EventArgs) Handles cmbProjectStage.ValueChanged
        Try
            _StageId = cmbProjectStage.SelectedRow.Cells("Id").Value
        Catch ex As Exception
            _StageId = 0
        End Try
    End Sub

    Sub CopyExcel()
        Try
            isPasting = True
            Dim row As UltraGridRow
            Dim ms As MemoryStream = CType(Clipboard.GetData("XML Spreadsheet"), MemoryStream)
            Dim quoteFiedType As Integer
            If IsNothing(ms) = True Then
                Exit Sub
            End If
            Dim b(CInt(ms.Length)) As Byte
            Dim doc As New XmlDocument()
            doc.Load(ms)
            Dim elemList As XmlNodeList = doc.GetElementsByTagName("Row")
            For Each items As XmlNode In elemList
                If UG2.Rows.Count > 0 Then
                    If UG2.Rows(UG2.Rows.Count - 1).Cells("QuoteFiedType").Value = "" Then
                        ' DeleteEmptyRows("last")
                    End If
                End If
                row = AddNewRow("after")

                row.ParentCollection.Move(row, UG2.ActiveRow.Index)
                Me.UG2.ActiveRowScrollRegion.ScrollRowIntoView(row)
                If items.ChildNodes(1).InnerText <> "" Then
                    quoteFiedType = QuateFiedTypesList.Text
                    row.Cells("QuoteFiedType").Value = quoteFiedType
                    row.Cells("LineNotes").Value = GetData(items, 0)
                    row.Cells("LineComments").Value = GetData(items, 1)
                    row.Cells("Height").Value = If(IsNumeric(GetData(items, 2)) = False, 0, Convert.ToInt32(GetData(items, 2)))
                    row.Cells("Width").Value = If(IsNumeric(GetData(items, 3)) = False, 0, Convert.ToInt32(GetData(items, 3)))
                    row.Cells("Qty").Value = If(IsNumeric(GetData(items, 4)) = False, 0, Convert.ToInt32(GetData(items, 4)))

                Else
                    If GetData(items, 0) <> "" Then
                        quoteFiedType = QuateFiedTypesList.Header_Sub
                        row.Cells("QuoteFiedType").Value = quoteFiedType
                        row.Cells("LineComments").Value = GetData(items, 0)

                    Else
                        'row.Delete()
                    End If
                End If

                '' Me.UG2.ActiveRow.PerformAutoSize()

            Next
            isPasting = False
        Catch ex As Exception
            modGlazingQuoteExtension.GQShowMessage(ex.Message, Me.Text, MsgBoxStyle.Critical)

        End Try
    End Sub

    Private Sub CopyFromExcelToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CopyFromExcelToolStripMenuItem.Click
        CopyExcel()
    End Sub

    Private Function GetData(ByRef items As XmlNode, ByRef position As Integer)
        Try
            Dim value As String
            Dim nodeCount = items.ChildNodes.Count
            If nodeCount = position Then
                Return ""
                Exit Function
            End If
            value = items.ChildNodes(position).InnerText
            If IsNothing(value) = False Then
                Return value
            Else
                Return ""
            End If
        Catch ex As Exception
            modGlazingQuoteExtension.GQShowMessage(ex.Message, Me.Text, MsgBoxStyle.Critical)
            Return ""
        End Try
    End Function

    Sub DeleteEmptyRows(ByRef deleteMode As String)
        Try
            If deleteMode = "all" Then
                For Each row As UltraGridRow In UG2.Rows
                    row.Delete()
                Next
            ElseIf deleteMode = "last" Then
                UG2.Rows(UG2.Rows.Count - 1).Delete()
                UG2.Rows(UG2.Rows.Count - 1).ParentCollection.Move(UG2.Rows(UG2.Rows.Count - 1), UG2.ActiveRow.Index)
                Me.UG2.ActiveRowScrollRegion.ScrollRowIntoView(UG2.Rows(UG2.Rows.Count - 1))
            End If
        Catch ex As Exception
            modGlazingQuoteExtension.GQShowMessage(ex.Message, Me.Text, MsgBoxStyle.Critical)

        End Try
    End Sub

    Sub calculatePriceUsingAmount(ByRef row As UltraGridRow)
        Try
            Dim itemAmount As Decimal
            Dim itemPrice As Decimal = 0.0
            Dim itemTaxRate As Decimal = 0.0
            Dim itemTax As Decimal = 0.0
            Dim itemExcAmount As Decimal = 0.0
            Dim itemIncAmount As Decimal = 0.0
            Dim volum As Decimal = 0.0
            Dim sercieIncPrice As Decimal = 0.0

            itemAmount = If(IsNumeric(row.Cells("Amount").Value) = True, row.Cells("Amount").Value, 0)
            subItemIncPrice = If(IsNumeric(row.Cells("ServiceItemTotNet").Value) = True, row.Cells("ServiceItemTotNet").Value, 0)
            'subItemIncPrice = If(IsNumeric(row.Cells("ServiceGross").Value) = True, row.Cells("ServiceGross").Value, 0)

            If subItemIncPrice > 0 Then
                itemAmount = itemAmount - subItemIncPrice
            End If
            If itemAmount > 0 Then
                itemTax = If(IsNumeric(row.Cells("TaxRate").Value) = True, row.Cells("TaxRate").Value, 0)
                'itemTax = If(IsNumeric(row.Cells("TaxRate").Text) = True, row.Cells("TaxRate").Text, 0)
                volum = If(IsNumeric(row.Cells("Volume").Value) = True, row.Cells("Volume").Value, 0)
                If volum = 0 Then
                    volum = (If(IsNumeric(row.Cells("Height").Value) = True, row.Cells("Height").Value, 0) * If(IsNumeric(row.Cells("Width").Value) = True, row.Cells("Width").Value, 0)) / 1000000
                End If
                If volum = 0 Then
                    volum = If(IsNumeric(row.Cells("Qty").Value) = True, row.Cells("Qty").Value, 0)
                End If
                itemPrice = Math.Round(itemAmount / volum, 6, MidpointRounding.AwayFromZero)

                If itemTax > 0 Then
                    itemTaxRate = Math.Round(itemAmount / itemTax, 2, MidpointRounding.AwayFromZero)
                Else
                    itemTaxRate = 0
                End If

                itemIncAmount = Math.Round(itemTaxRate + itemAmount, 2, MidpointRounding.AwayFromZero)
                itemExcAmount = Math.Round(itemAmount, 2, MidpointRounding.AwayFromZero)
                row.Cells("Amount").Value = itemAmount
                row.Cells("Tax").Value = itemTaxRate
                row.Cells("Price").Value = itemPrice
                row.Cells("ItmExcAmount").Value = itemExcAmount
                row.Cells("Net").Value = itemIncAmount
                row.Cells("TaxRateValue").Value = itemTaxRate
            End If
        Catch ex As Exception
            modGlazingQuoteExtension.GQShowMessage(ex.Message, Me.Text, MsgBoxStyle.Critical)

        End Try
    End Sub

    Sub GetExpireDate(Optional ByRef newDateSet As DataSet = Nothing)
        Try
            If IsNothing(newDateSet) = True Then
                newDateSet = GetQuoteDefaultData()
            End If

            If IsNothing(newDateSet) = False Then
                If newDateSet.Tables(1).Rows.Count > 0 Then
                    For Each newRow As DataRow In newDateSet.Tables(1).Rows
                        expireDateCount = newRow.Item("defaultExpiryDate")
                    Next
                End If
            End If
        Catch ex As Exception
            modGlazingQuoteExtension.GQShowMessage(ex.Message, Me.Text, MsgBoxStyle.Critical)

        End Try
    End Sub

    Public Sub SetExpireDate(Optional ByRef newExpireDateCount As Integer = -1)
        Try
            GetExpireDate()
            Dim dueDate As DateTime = txtOrdDate.DateTime
            Dim expireDate As DateTime = txtDueDate.DateTime
            If newExpireDateCount > -1 Then
                txtDueDate.DateTime = dueDate.AddDays(newExpireDateCount)
            Else
                txtDueDate.DateTime = dueDate.AddDays(expireDateCount)
            End If

        Catch ex As Exception
            modGlazingQuoteExtension.GQShowMessage(ex.Message, Me.Text, MsgBoxStyle.Critical)

        End Try
    End Sub


    Sub GridCellAppearence(Optional ByRef dataRow As UltraGridRow = Nothing)
        Dim row As UltraGridRow

        If IsNothing(dataRow) Then
            row = UG2.ActiveRow
        Else
            row = dataRow
        End If

        Try
            If row.Cells("ServiceItemTotNet").Value > 0 Then
                row.Cells("Height").Activation = Activation.NoEdit
                row.Cells("Width").Activation = Activation.NoEdit
            Else
                row.Cells("Height").Activation = Activation.AllowEdit
                row.Cells("Width").Activation = Activation.AllowEdit
            End If

        Catch ex As Exception

        End Try
    End Sub
End Class