Imports Infragistics.Win.UltraWinGrid
Public Class frmGlazingDocDefaultSetting
    Dim isExist As Boolean
    Public isExternalRequest As Boolean = False
    Dim publicVisibleState As Integer = 0
    Dim getGlobalDataByDefault As Integer = 0
    Dim startingTax As Integer
    Dim startingTaxState As Boolean = False
    Private ob As frmGlazingQuote
    Dim isDefaultColChanged As Boolean = False
    Dim isLoading As Boolean = False
    Dim isQuoteStateChanged = False

    Public Sub New(ByRef ob As frmGlazingQuote)

        ' This call is required by the designer.
        InitializeComponent()
        Me.ob = ob

        ' Add any initialization after the InitializeComponent() call.

    End Sub

    Private Sub btnOK_Click(sender As Object, e As EventArgs) Handles btnOK.Click
        SaveData()

    End Sub
    Public Sub SaveData()
        Dim SQLNew As clsSqlConn = New clsSqlConn
        Dim sqlQuary As String = ""
        Dim taxCheckState As Integer

        '----For global settings----
        'If chkGlobalSave.Checked = True Then
        '    Dim newDataset As DataSet = Nothing
        '    newDataset = GetGlzQuoteDefaults("SELECT * FROM  GlzQuote_Defaults WHERE getGlobalDataByDefault = 1")
        '    If IsNothing(newDataset) Then
        '        If newDataset.Tables(0).Rows.Count > 0 Then
        '            If frmGlazingQuote.ShowMessage("Already default tax rate set in globally" & vbCrLf & "Do you want to Overwrite it", Me.Text, MessageBoxButtons.YesNo) = System.Windows.Forms.DialogResult.No Then
        '                Exit Sub
        '            End If
        '        End If
        '    End If
        'End If
        Dim collspPara As New Collection
        Dim colPara As New spParameters
        Dim newSQLQuery As String = ""
        SQLNew.Begin_Trans()
        Try
            If chkTaxInc.CheckState = CheckState.Checked Then
                'Inc
                taxCheckState = 1

            Else
                'Exc
                taxCheckState = 0

            End If


            colPara.ParaName = "@hasDefaultDocStateColor"
            colPara.ParaValue = chkDefaultBackCol.Checked
            collspPara.Add(colPara)

            colPara.ParaName = "@defaultDocStateColor"
            colPara.ParaValue = ucpBackColor.Color.ToArgb
            collspPara.Add(colPara)

            colPara.ParaName = "@createdBy"
            colPara.ParaValue = AgentID
            collspPara.Add(colPara)

            If (startingTax <> ucmbDefaultTax.Value) Or (startingTaxState <> chkTaxInc.Checked) Then

                colPara.ParaName = "@isTaxInc"
                colPara.ParaValue = taxCheckState
                collspPara.Add(colPara)

                colPara.ParaName = "@defaultTaxtRate"
                colPara.ParaValue = ucmbDefaultTax.Value
                collspPara.Add(colPara)

                colPara.ParaName = "@publicVisibleState"
                colPara.ParaValue = publicVisibleState
                collspPara.Add(colPara)

                colPara.ParaName = "@getGlobalDataByDefault"
                colPara.ParaValue = getGlobalDataByDefault
                collspPara.Add(colPara)

                'colPara.ParaName = "@getGlobalDataByDefault"
                'colPara.ParaValue = getGlobalDataByDefault
                'collspPara.Add(colPara)

                If frmGlazingQuote.ShowMessage("Do you want to change the default tax settings?", Me.Text, MessageBoxButtons.YesNo) = System.Windows.Forms.DialogResult.Yes Then
                    If isExist = False Then
                        newSQLQuery = "INSERT INTO GlzQuote_Defaults (isTaxInc, createdBy, defaultTaxtRate, publicVisibleState, getGlobalDataByDefault, hasDefaultDocStateColor,  defaultDocStateColor) " & _
                        "VALUES (@isTaxInc, @createdBy, @defaultTaxtRate, @publicVisibleState, @getGlobalDataByDefault, @hasDefaultDocStateColor, @defaultDocStateColor )"

                    Else
                        newSQLQuery = "UPDATE GlzQuote_Defaults SET isTaxInc = @isTaxInc, defaultTaxtRate = @defaultTaxtRate, " & _
                        " publicVisibleState = @publicVisibleState, getGlobalDataByDefault = @getGlobalDataByDefault, hasDefaultDocStateColor = @hasDefaultDocStateColor, defaultDocStateColor = @defaultDocStateColor  WHERE createdBy = @createdBy"

                    End If
                End If

            Else
                newSQLQuery = "UPDATE GlzQuote_Defaults SET hasDefaultDocStateColor = @hasDefaultDocStateColor, defaultDocStateColor = @defaultDocStateColor  WHERE createdBy = @createdBy"

            End If
            If SQLNew.EXE_SQL_Trans_Para_Return(newSQLQuery, collspPara) = 0 Then
                frmGlazingQuote.ShowMessage("Erro in item pictures", Me.Text, MsgBoxStyle.Critical)
                SQLNew.Rollback_Trans()
                Exit Sub

            End If

            'changes tax values in frmGlazingQuote
            If (startingTax <> ucmbDefaultTax.Value) Or (startingTaxState <> chkTaxInc.Checked) Then
                ob.AfterDefaultTaxePriceChaged(ucmbDefaultTax.Value, ucmbDefaultTax.SelectedRow.Cells("TaxRate").Value, chkTaxInc.CheckState)

            End If
            collspPara.Clear()

            SaveColorOptionGrid(SQLNew)
            SQLNew.Commit_Trans()

            'changes values in frmGlazingQuote
            If isQuoteStateChanged = True Then
                Dim oldValue = ob.utxtQuoteState.Value
                ob.LoadQuoteState()
                ob.utxtQuoteState.Value = oldValue
                If ob.utxtQuoteState.Text = oldValue.ToString Then
                    ob.utxtQuoteState.Text = ""
                    ob.utxtQuoteState.Appearance.ForeColor = Color.Red
                End If

            End If

        Catch ex As Exception
            frmGlazingQuote.ShowMessage(ex.Message, Me.Text, MessageBoxButtons.OK)

        Finally
            SQLNew = Nothing
        End Try
    End Sub
    Public Sub LoadExistingData()

        Try
            isLoading = True
            Dim newSQLQuary As String


            newSQLQuary = "SELECT * FROM  GlzQuote_Defaults WHERE createdBy = " & AgentID
            newSQLQuary += "SELECT * FROM  GlzQuote_State"

            Dim ds As DataSet = GetGlzQuoteDefaults(newSQLQuary)

            If IsNothing(ds) = False Then
                If ds.Tables(0).Rows.Count > 0 Then
                    isExist = True
                    For Each row In ds.Tables(0).Rows

                        '----For global settings----
                        ''User set global Data by default
                        'If row("getGlobalDataByDefault") = True Then
                        '    'Get default settings
                        '    Dim dsGlobal As DataSet = GetGlzQuoteDefaults("SELECT * FROM  GlzQuote_Defaults WHERE publicVisibleState = 1")
                        '    If dsGlobal.Tables(0).Rows.Count > 0 Then
                        '        For Each rowGlobal In ds.Tables(0).Rows
                        '            GetDefaultGlobalSettings(row)
                        '        Next
                        '    Else
                        '        'Set to get global default but no values in DB
                        '        GetDefaultUsreSettings(row)
                        '    End If
                        'Else
                        '    'Getting user default
                        'GetDefaultUsreSettings(row)
                        'End If

                        GetDefaultGlobalSettings(row)

                        chkDefaultBackCol.Checked = Row("hasDefaultDocStateColor")
                        ucpBackColor.Color = Color.FromArgb(Row("defaultDocStateColor"))

                    Next

                Else
                    'no user default
                    Dim dsGlobal As DataSet = GetGlzQuoteDefaults("SELECT * FROM  GlzQuote_Defaults WHERE publicVisibleState = 1")
                    If dsGlobal.Tables(0).Rows.Count > 0 Then
                        'Get default settings
                        For Each rowGlobal In ds.Tables(0).Rows
                            GetDefaultGlobalSettings(rowGlobal)
                        Next
                    End If
                End If
                FillColorOptionGrid(ds.Tables(1))
            End If
            startingTax = ucmbDefaultTax.Value
            startingTaxState = chkTaxInc.Checked
        Catch ex As Exception
            frmGlazingQuote.ShowMessage(ex.Message, Me.Text, MsgBoxStyle.Critical)
        Finally
            isLoading = False
        End Try

    End Sub


    Public Function LoadExistingTaxData() As DataSet
        Try
            Dim ds As DataSet = GetGlzQuoteDefaults("SELECT * FROM  GlzQuote_Defaults WHERE createdBy = " & AgentID)
            If IsNothing(ds) = False Then
                If ds.Tables(0).Rows.Count > 0 Then
                    isExist = True
                    For Each row In ds.Tables(0).Rows
                        'Getting user default
                        If isExternalRequest = False Then
                            GetDefaultUsreSettings(row)
                        End If
                    Next
                End If
            End If
        Catch ex As Exception
            frmGlazingQuote.ShowMessage(ex.Message, Me.Text, MsgBoxStyle.Critical)
        End Try

    End Function


    Sub GetDefaultGlobalSettings(ByVal rowGlobal As DataRow)
        If rowGlobal("isTaxInc") = True Then
            chkTaxInc.CheckState = CheckState.Checked
        Else
            chkTaxInc.CheckState = CheckState.Unchecked
        End If

        ucmbDefaultTax.Value = rowGlobal("defaultTaxtRate")

        '----For global settings----
        'If rowGlobal("publicVisibleState") = 1 Then
        '    chkGlobalSave.CheckState = CheckState.Checked
        'Else
        '    chkGlobalSave.CheckState = CheckState.Unchecked
        'End If
    End Sub

    Sub GetDefaultUsreSettings(ByVal row As DataRow)

        'no default settings in database
        'frmGlazingQuote.ShowMessage("No default settings were found." & vbCrLf & "Loading user default settings ", Me.Text, MessageBoxButtons.OK)

        If row("isTaxInc") = True Then
            chkTaxInc.CheckState = CheckState.Checked
        Else
            chkTaxInc.CheckState = CheckState.Unchecked
        End If

        ucmbDefaultTax.Value = row("defaultTaxRate")

        If row("publicVisibleState") = 1 Then
            chkGlobalSave.CheckState = CheckState.Checked
        Else
            chkGlobalSave.CheckState = CheckState.Unchecked
        End If

        If row("getGlobalDataByDefault") = 1 Then
            chkGlobalSave.CheckState = CheckState.Checked
        Else
            chkGlobalSave.CheckState = CheckState.Unchecked
        End If
    End Sub

    Public Function GetGlzQuoteDefaults(sqlQuary As String) As DataSet

        Dim QuoteDefaultsDataset As DataSet
        Dim objSQL As New clsSqlConn

        With objSQL
            Try
                Dim ds As DataSet = .GET_INSERT_UPDATE(sqlQuary)
                QuoteDefaultsDataset = ds

            Catch ex As Exception
                frmGlazingQuote.ShowMessage(ex.Message, Me.Text, MsgBoxStyle.Critical)

            Finally
                SQLNew = Nothing
            End Try
        End With

        Return QuoteDefaultsDataset

    End Function

    Public Sub LoadTaxData()
        Dim objSQL As New clsSqlConn
        Try
            Dim sqlQuary As String = "SELECT * FROM  TaxRate"
            With objSQL
                Dim ds As DataSet = .GET_INSERT_UPDATE(sqlQuary)
                If ds.Tables(0).Rows.Count > 0 Then
                    ucmbDefaultTax.DataSource = ds.Tables(0)
                    ucmbDefaultTax.DisplayMember = "Description"
                    ucmbDefaultTax.ValueMember = "idTaxRate"
                    ucmbDefaultTax.DisplayLayout.Bands(0).Columns("idTaxRate").Hidden = True
                    ucmbDefaultTax.Refresh()
                End If
            End With
        Catch ex As Exception
            frmGlazingQuote.ShowMessage(ex.Message, Me.Text, MsgBoxStyle.Critical)

        Finally
            objSQL = Nothing
        End Try
    End Sub


    Private Sub frmGlazingDocDefaultSetting_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        LoadTaxData()
        LoadExistingData()
    End Sub

    '----For global settings----
    'Private Sub chkGlobalSave_CheckedChanged(sender As Object, e As EventArgs) Handles chkGlobalSave.CheckedChanged
    '    If chkGlobalSave.Checked = True Then
    '        publicVisibleState = 1
    '    Else
    '        publicVisibleState = 0
    '    End If
    'End Sub

    '----For global settings----
    'Function GetVisbleStateData() As DataSet
    '    Dim QuoteDefaultsDataset As DataSet = Nothing
    '    Dim sqlQuary As String = "SELECT  publicVisibleState, createdBy FROM  GlzQuote_Defaults WHERE publicVisibleState = 1 "
    '    Dim objSQL As New clsSqlConn

    '    With objSQL
    '        Try
    '            Dim ds As DataSet = .GET_INSERT_UPDATE(sqlQuary)
    '            QuoteDefaultsDataset = ds

    '        Catch ex As Exception
    '            frmGlazingQuote.ShowMessage(ex.Message, Me.Text, MsgBoxStyle.Critical)

    '        Finally
    '            SQLNew = Nothing
    '        End Try
    '    End With
    '    Return QuoteDefaultsDataset

    'End Function

    '----For global settings----
    'Private Sub chkLoadGlobalDefault_CheckedChanged(sender As Object, e As EventArgs) Handles chkLoadGlobalDefault.CheckedChanged
    '    If chkGlobalSave.Checked = True Then
    '        getGlobalDataByDefault = 1
    '    Else
    '        getGlobalDataByDefault = 0
    '    End If
    'End Sub

    Private Sub chkTaxInc_CheckedChanged(sender As Object, e As EventArgs) Handles chkTaxInc.CheckedChanged
        'ucmbDefaultTax.Value = 0
    End Sub

    Sub SetQuoteStateDetails()
        Try
            Dim ds As DataSet = GetGlzQuoteDefaults("SELECT * FROM  GlzQuote_Defaults")
            Dim substringsForeColor1() As String
            Dim substringsForeColor2() As String
            Dim substringsForeColor3() As String
            Dim substringsForeColor4() As String
            Dim substringsForeColor5() As String
            Dim substringsForeColor6() As String
            Dim substringsForeColor7() As String

            If IsNothing(ds) = False Then

                If ds.Tables(0).Rows.Count > 6 Then

                    'Set For color
                    substringsForeColor1 = (ds.Tables(0).Rows(0).Item("JQSForeColor")).Split(",")
                    substringsForeColor2 = (ds.Tables(0).Rows(1).Item("JQSForeColor")).Split(",")
                    substringsForeColor3 = (ds.Tables(0).Rows(2).Item("JQSForeColor")).Split(",")
                    substringsForeColor4 = (ds.Tables(0).Rows(3).Item("JQSForeColor")).Split(",")
                    substringsForeColor5 = (ds.Tables(0).Rows(4).Item("JQSForeColor")).Split(",")
                    substringsForeColor6 = (ds.Tables(0).Rows(5).Item("JQSForeColor")).Split(",")
                    substringsForeColor7 = (ds.Tables(0).Rows(6).Item("JQSForeColor")).Split(",")

                    ucpQuoteState1.Color = Color.FromArgb(Convert.ToInt16(substringsForeColor1(0)), Convert.ToInt16(substringsForeColor1(1)), Convert.ToInt16(substringsForeColor1(2)), Convert.ToInt16(substringsForeColor1(3)))
                    ucpQuoteState2.Color = Color.FromArgb(Convert.ToInt16(substringsForeColor2(0)), Convert.ToInt16(substringsForeColor2(1)), Convert.ToInt16(substringsForeColor2(2)), Convert.ToInt16(substringsForeColor2(3)))
                    ucpQuoteState3.Color = Color.FromArgb(Convert.ToInt16(substringsForeColor3(0)), Convert.ToInt16(substringsForeColor3(1)), Convert.ToInt16(substringsForeColor3(2)), Convert.ToInt16(substringsForeColor3(3)))
                    ucpQuoteState4.Color = Color.FromArgb(Convert.ToInt16(substringsForeColor4(0)), Convert.ToInt16(substringsForeColor4(1)), Convert.ToInt16(substringsForeColor4(2)), Convert.ToInt16(substringsForeColor4(3)))
                    ucpQuoteState5.Color = Color.FromArgb(Convert.ToInt16(substringsForeColor5(0)), Convert.ToInt16(substringsForeColor5(1)), Convert.ToInt16(substringsForeColor5(2)), Convert.ToInt16(substringsForeColor5(3)))
                    ucpQuoteState6.Color = Color.FromArgb(Convert.ToInt16(substringsForeColor6(0)), Convert.ToInt16(substringsForeColor6(1)), Convert.ToInt16(substringsForeColor6(2)), Convert.ToInt16(substringsForeColor6(3)))
                    ucpQuoteState7.Color = Color.FromArgb(Convert.ToInt16(substringsForeColor7(0)), Convert.ToInt16(substringsForeColor7(1)), Convert.ToInt16(substringsForeColor7(2)), Convert.ToInt16(substringsForeColor7(3)))

                    'Set Back color
                    substringsBackColor = (ds.Tables(0).Rows(1).Item("JQSBackColor")).Split(",")
                    ucpBackColor.Color = Color.FromArgb(Convert.ToInt16(substringsForeColor1(0)), Convert.ToInt16(substringsForeColor1(1)), Convert.ToInt16(substringsForeColor1(2)), Convert.ToInt16(substringsForeColor1(3)))


                End If
            End If
        Catch ex As Exception

        End Try
    End Sub

    Sub FillColorOptionGrid(ByRef newDataTable As DataTable)
        Try
            Dim newDataRow As DataRow
            Dim row As UltraGridRow
            For Each newDataRow In newDataTable.Rows
                row = Me.ugColorOption.DisplayLayout.Bands(0).AddNew()
                row.Cells("quoteStateName").Value = newDataRow("JQSName")
                row.Cells("foreColor").Value = Color.FromArgb(newDataRow("JQSForeColor"))
                row.Cells("backColor").Value = Color.FromArgb(newDataRow("JQSBackColor"))
                row.Cells("isActive").Value = newDataRow("JQSState")
                row.Cells("JQSID").Value = newDataRow("JQSID")

            Next

        Catch ex As Exception

        End Try

    End Sub

    Sub SaveColorOptionGrid(ByRef newObjSQL As clsSqlConn)
        Try
            Dim row As UltraGridRow
            Dim collspPara As New Collection
            Dim colPara As New spParameters
            Dim newSQLQuery As String = ""
            Dim newColor As Color
            For Each row In Me.ugColorOption.Rows
                colPara.ParaName = "@JQSName"
                colPara.ParaValue = row.Cells("quoteStateName").Value
                collspPara.Add(colPara)

                colPara.ParaName = "@JQSForeColor"
                newColor = row.Cells("foreColor").Value
                colPara.ParaValue = newColor.ToArgb
                collspPara.Add(colPara)

                colPara.ParaName = "@JQSbackColor"
                newColor = row.Cells("backColor").Value
                colPara.ParaValue = newColor.ToArgb
                collspPara.Add(colPara)

                colPara.ParaName = "@JQSState"
                colPara.ParaValue = row.Cells("isActive").Value
                collspPara.Add(colPara)

                If row.Cells("JQSID").Value = 0 Then
                    newSQLQuery = "INSERT INTO GlzQuote_State (JQSName, JQSForeColor, JQSbackColor, JQSState ) " & _
                    "VALUES (@JQSName, @JQSForeColor, @JQSbackColor, @JQSState )"

                Else
                    newSQLQuery = "UPDATE GlzQuote_State SET JQSName = @JQSName, JQSForeColor = @JQSForeColor, JQSbackColor = @JQSbackColor, JQSState = @JQSState " & _
                        "WHERE JQSID = '" & row.Cells("JQSID").Value & "'"

                End If

                If newObjSQL.EXE_SQL_Trans_Para_Return(newSQLQuery, collspPara) = 0 Then
                    frmGlazingQuote.ShowMessage("Erro in item pictures", Me.Text, MsgBoxStyle.Critical)
                    newObjSQL.Rollback_Trans()
                    Exit Sub

                End If
                collspPara.Clear()
            Next


        Catch ex As Exception
            frmGlazingQuote.ShowMessage(ex.Message, Me.Text, MsgBoxStyle.Critical)

        End Try

    End Sub

    Private Sub chkDefaultBackCol_CheckedChanged(sender As Object, e As EventArgs) Handles chkDefaultBackCol.CheckedChanged
        ucpBackColor.Visible = True
        isDefaultColChanged = True
    End Sub

    

    Private Sub ucpBackColor_ColorChanged(sender As Object, e As EventArgs) Handles ucpBackColor.ColorChanged
        If isLoading = False Then
            ColorChanger()
        End If
    End Sub

    Sub ColorChanger()
        Try
            Dim row As UltraGridRow
            Dim collspPara As New Collection
            Dim colPara As New spParameters
            Dim newSQLQuery As String = ""
            Dim colorList As New List(Of Color)

            If chkDefaultBackCol.Checked = True And IsNothing(ugColorOption.Rows) = False Then
                For Each row In Me.ugColorOption.Rows
                    'If isDefaultColChanged = True Then
                    '    colorList.Add(Color.FromArgb(row.Cells("backColor").Value))

                    'End If
                    row.Cells("backColor").Value = ucpBackColor.Color

                Next
                'isDefaultColChanged = False

                'Else
                '    For Each row In Me.ugColorOption.Rows
                '        row.Cells("backColor").Value = colorList.Item(row.Index)

                '    Next
            End If

        Catch ex As Exception
            frmGlazingQuote.ShowMessage(ex.Message, Me.Text, MsgBoxStyle.Critical)

        End Try
    End Sub

    Private Sub ugColorOption_AfterCellUpdate(sender As Object, e As CellEventArgs) Handles ugColorOption.AfterCellUpdate
        Try
            If IsNothing(Me.ugColorOption.ActiveCell) = False Then
                If Me.ugColorOption.ActiveCell.Column.Key = "quoteStateName" Then
                    isQuoteStateChanged = True
                End If
            End If
        Catch ex As Exception
            ob.ShowMessage(ex.Message, Me.Text, MsgBoxStyle.Critical)

        End Try

    End Sub
End Class