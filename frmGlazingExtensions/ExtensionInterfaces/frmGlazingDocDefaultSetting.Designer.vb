<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmGlazingDocDefaultSetting
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim Appearance1 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim UltraGridBand1 As Infragistics.Win.UltraWinGrid.UltraGridBand = New Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1)
        Dim UltraGridColumn1 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("quoteStateName")
        Dim UltraGridColumn2 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("foreColor")
        Dim UltraGridColumn3 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("backColor")
        Dim UltraGridColumn4 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("isActive")
        Dim UltraGridColumn5 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("JQSID")
        Dim Appearance2 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance3 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance4 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance5 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance6 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance7 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance8 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance9 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance10 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance11 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance12 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim UltraDataColumn1 As Infragistics.Win.UltraWinDataSource.UltraDataColumn = New Infragistics.Win.UltraWinDataSource.UltraDataColumn("quoteStateName")
        Dim UltraDataColumn2 As Infragistics.Win.UltraWinDataSource.UltraDataColumn = New Infragistics.Win.UltraWinDataSource.UltraDataColumn("foreColor")
        Dim UltraDataColumn3 As Infragistics.Win.UltraWinDataSource.UltraDataColumn = New Infragistics.Win.UltraWinDataSource.UltraDataColumn("backColor")
        Dim UltraDataColumn4 As Infragistics.Win.UltraWinDataSource.UltraDataColumn = New Infragistics.Win.UltraWinDataSource.UltraDataColumn("isActive")
        Dim UltraDataColumn5 As Infragistics.Win.UltraWinDataSource.UltraDataColumn = New Infragistics.Win.UltraWinDataSource.UltraDataColumn("JQSID")
        Dim Appearance14 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance15 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance16 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance19 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance20 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance21 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance22 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance23 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance24 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance25 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance26 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance27 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim UltraDataColumn6 As Infragistics.Win.UltraWinDataSource.UltraDataColumn = New Infragistics.Win.UltraWinDataSource.UltraDataColumn("Quote State")
        Dim UltraDataColumn7 As Infragistics.Win.UltraWinDataSource.UltraDataColumn = New Infragistics.Win.UltraWinDataSource.UltraDataColumn("Column 1")
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmGlazingDocDefaultSetting))
        Me.btnOK = New System.Windows.Forms.Button()
        Me.btnCancel = New System.Windows.Forms.Button()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.gbColorOptions = New System.Windows.Forms.GroupBox()
        Me.chkDefaultBackCol = New System.Windows.Forms.CheckBox()
        Me.ugColorOption = New Infragistics.Win.UltraWinGrid.UltraGrid()
        Me.UltraDataSource2 = New Infragistics.Win.UltraWinDataSource.UltraDataSource(Me.components)
        Me.ucpBackColor = New Infragistics.Win.UltraWinEditors.UltraColorPicker()
        Me.ucpQuoteState7 = New Infragistics.Win.UltraWinEditors.UltraColorPicker()
        Me.ucpQuoteState6 = New Infragistics.Win.UltraWinEditors.UltraColorPicker()
        Me.lblColOp7 = New System.Windows.Forms.Label()
        Me.lblColOp6 = New System.Windows.Forms.Label()
        Me.lblColOp5 = New System.Windows.Forms.Label()
        Me.lblColOp4 = New System.Windows.Forms.Label()
        Me.lblColOp3 = New System.Windows.Forms.Label()
        Me.lblColOp2 = New System.Windows.Forms.Label()
        Me.lblColOp1 = New System.Windows.Forms.Label()
        Me.ucpQuoteState5 = New Infragistics.Win.UltraWinEditors.UltraColorPicker()
        Me.ucpQuoteState4 = New Infragistics.Win.UltraWinEditors.UltraColorPicker()
        Me.ucpQuoteState3 = New Infragistics.Win.UltraWinEditors.UltraColorPicker()
        Me.ucpQuoteState2 = New Infragistics.Win.UltraWinEditors.UltraColorPicker()
        Me.ucpQuoteState1 = New Infragistics.Win.UltraWinEditors.UltraColorPicker()
        Me.Panel3 = New System.Windows.Forms.Panel()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.ucmbDefaultTax = New Infragistics.Win.UltraWinGrid.UltraCombo()
        Me.lblDefaultTax = New System.Windows.Forms.Label()
        Me.chkLoadGlobalDefault = New System.Windows.Forms.CheckBox()
        Me.chkGlobalSave = New System.Windows.Forms.CheckBox()
        Me.chkTaxInc = New System.Windows.Forms.CheckBox()
        Me.UltraDataSource1 = New Infragistics.Win.UltraWinDataSource.UltraDataSource(Me.components)
        Me.Panel1.SuspendLayout()
        Me.gbColorOptions.SuspendLayout()
        CType(Me.ugColorOption, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.UltraDataSource2, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.ucpBackColor, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.ucpQuoteState7, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.ucpQuoteState6, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.ucpQuoteState5, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.ucpQuoteState4, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.ucpQuoteState3, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.ucpQuoteState2, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.ucpQuoteState1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel3.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        CType(Me.ucmbDefaultTax, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.UltraDataSource1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'btnOK
        '
        Me.btnOK.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnOK.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnOK.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.btnOK.Location = New System.Drawing.Point(237, 3)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(76, 22)
        Me.btnOK.TabIndex = 5
        Me.btnOK.Text = "OK"
        Me.btnOK.UseVisualStyleBackColor = True
        '
        'btnCancel
        '
        Me.btnCancel.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnCancel.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.btnCancel.Location = New System.Drawing.Point(319, 4)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(76, 22)
        Me.btnCancel.TabIndex = 6
        Me.btnCancel.Text = "Cancel"
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.gbColorOptions)
        Me.Panel1.Controls.Add(Me.Panel3)
        Me.Panel1.Controls.Add(Me.GroupBox1)
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel1.Location = New System.Drawing.Point(0, 0)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(398, 433)
        Me.Panel1.TabIndex = 5
        '
        'gbColorOptions
        '
        Me.gbColorOptions.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.gbColorOptions.Controls.Add(Me.chkDefaultBackCol)
        Me.gbColorOptions.Controls.Add(Me.ugColorOption)
        Me.gbColorOptions.Controls.Add(Me.ucpBackColor)
        Me.gbColorOptions.Controls.Add(Me.ucpQuoteState7)
        Me.gbColorOptions.Controls.Add(Me.ucpQuoteState6)
        Me.gbColorOptions.Controls.Add(Me.lblColOp7)
        Me.gbColorOptions.Controls.Add(Me.lblColOp6)
        Me.gbColorOptions.Controls.Add(Me.lblColOp5)
        Me.gbColorOptions.Controls.Add(Me.lblColOp4)
        Me.gbColorOptions.Controls.Add(Me.lblColOp3)
        Me.gbColorOptions.Controls.Add(Me.lblColOp2)
        Me.gbColorOptions.Controls.Add(Me.lblColOp1)
        Me.gbColorOptions.Controls.Add(Me.ucpQuoteState5)
        Me.gbColorOptions.Controls.Add(Me.ucpQuoteState4)
        Me.gbColorOptions.Controls.Add(Me.ucpQuoteState3)
        Me.gbColorOptions.Controls.Add(Me.ucpQuoteState2)
        Me.gbColorOptions.Controls.Add(Me.ucpQuoteState1)
        Me.gbColorOptions.Location = New System.Drawing.Point(12, 119)
        Me.gbColorOptions.Name = "gbColorOptions"
        Me.gbColorOptions.Size = New System.Drawing.Size(375, 279)
        Me.gbColorOptions.TabIndex = 7
        Me.gbColorOptions.TabStop = False
        Me.gbColorOptions.Text = "Quote state settings"
        '
        'chkDefaultBackCol
        '
        Me.chkDefaultBackCol.AutoSize = True
        Me.chkDefaultBackCol.Location = New System.Drawing.Point(6, 231)
        Me.chkDefaultBackCol.Name = "chkDefaultBackCol"
        Me.chkDefaultBackCol.Size = New System.Drawing.Size(130, 17)
        Me.chkDefaultBackCol.TabIndex = 6
        Me.chkDefaultBackCol.Text = "Set default back color"
        Me.chkDefaultBackCol.UseVisualStyleBackColor = True
        '
        'ugColorOption
        '
        Me.ugColorOption.DataSource = Me.UltraDataSource2
        Appearance1.BackColor = System.Drawing.SystemColors.Window
        Appearance1.BorderColor = System.Drawing.SystemColors.InactiveCaption
        Me.ugColorOption.DisplayLayout.Appearance = Appearance1
        Me.ugColorOption.DisplayLayout.AutoFitStyle = Infragistics.Win.UltraWinGrid.AutoFitStyle.ResizeAllColumns
        UltraGridColumn1.Header.Caption = "Quote State Name"
        UltraGridColumn1.Header.VisiblePosition = 0
        UltraGridColumn1.Width = 188
        UltraGridColumn2.Header.Caption = "Font Color"
        UltraGridColumn2.Header.VisiblePosition = 1
        UltraGridColumn2.Width = 53
        UltraGridColumn3.Header.Caption = "Backgorund Color"
        UltraGridColumn3.Header.VisiblePosition = 2
        UltraGridColumn3.Width = 44
        UltraGridColumn4.Header.Caption = "Enable"
        UltraGridColumn4.Header.VisiblePosition = 3
        UltraGridColumn4.Width = 48
        UltraGridColumn5.Header.VisiblePosition = 4
        UltraGridColumn5.Hidden = True
        UltraGridColumn5.Width = 62
        UltraGridBand1.Columns.AddRange(New Object() {UltraGridColumn1, UltraGridColumn2, UltraGridColumn3, UltraGridColumn4, UltraGridColumn5})
        Me.ugColorOption.DisplayLayout.BandsSerializer.Add(UltraGridBand1)
        Me.ugColorOption.DisplayLayout.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid
        Me.ugColorOption.DisplayLayout.CaptionVisible = Infragistics.Win.DefaultableBoolean.[False]
        Appearance2.BackColor = System.Drawing.SystemColors.ActiveBorder
        Appearance2.BackColor2 = System.Drawing.SystemColors.ControlDark
        Appearance2.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical
        Appearance2.BorderColor = System.Drawing.SystemColors.Window
        Me.ugColorOption.DisplayLayout.GroupByBox.Appearance = Appearance2
        Appearance3.ForeColor = System.Drawing.SystemColors.GrayText
        Me.ugColorOption.DisplayLayout.GroupByBox.BandLabelAppearance = Appearance3
        Me.ugColorOption.DisplayLayout.GroupByBox.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid
        Appearance4.BackColor = System.Drawing.SystemColors.ControlLightLight
        Appearance4.BackColor2 = System.Drawing.SystemColors.Control
        Appearance4.BackGradientStyle = Infragistics.Win.GradientStyle.Horizontal
        Appearance4.ForeColor = System.Drawing.SystemColors.GrayText
        Me.ugColorOption.DisplayLayout.GroupByBox.PromptAppearance = Appearance4
        Me.ugColorOption.DisplayLayout.MaxColScrollRegions = 1
        Me.ugColorOption.DisplayLayout.MaxRowScrollRegions = 1
        Appearance5.BackColor = System.Drawing.SystemColors.Window
        Appearance5.ForeColor = System.Drawing.SystemColors.ControlText
        Me.ugColorOption.DisplayLayout.Override.ActiveCellAppearance = Appearance5
        Appearance6.BackColor = System.Drawing.SystemColors.Highlight
        Appearance6.ForeColor = System.Drawing.SystemColors.HighlightText
        Me.ugColorOption.DisplayLayout.Override.ActiveRowAppearance = Appearance6
        Me.ugColorOption.DisplayLayout.Override.AllowColSizing = Infragistics.Win.UltraWinGrid.AllowColSizing.Free
        Me.ugColorOption.DisplayLayout.Override.BorderStyleCell = Infragistics.Win.UIElementBorderStyle.Dotted
        Me.ugColorOption.DisplayLayout.Override.BorderStyleRow = Infragistics.Win.UIElementBorderStyle.Dotted
        Appearance7.BackColor = System.Drawing.SystemColors.Window
        Me.ugColorOption.DisplayLayout.Override.CardAreaAppearance = Appearance7
        Appearance8.BorderColor = System.Drawing.Color.Silver
        Appearance8.TextTrimming = Infragistics.Win.TextTrimming.EllipsisCharacter
        Me.ugColorOption.DisplayLayout.Override.CellAppearance = Appearance8
        Me.ugColorOption.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.EditAndSelectText
        Me.ugColorOption.DisplayLayout.Override.CellPadding = 0
        Me.ugColorOption.DisplayLayout.Override.FixedRowIndicator = Infragistics.Win.UltraWinGrid.FixedRowIndicator.Button
        Appearance9.BackColor = System.Drawing.SystemColors.Control
        Appearance9.BackColor2 = System.Drawing.SystemColors.ControlDark
        Appearance9.BackGradientAlignment = Infragistics.Win.GradientAlignment.Element
        Appearance9.BackGradientStyle = Infragistics.Win.GradientStyle.Horizontal
        Appearance9.BorderColor = System.Drawing.SystemColors.Window
        Me.ugColorOption.DisplayLayout.Override.GroupByRowAppearance = Appearance9
        Appearance10.TextHAlignAsString = "Left"
        Me.ugColorOption.DisplayLayout.Override.HeaderAppearance = Appearance10
        Me.ugColorOption.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti
        Me.ugColorOption.DisplayLayout.Override.HeaderStyle = Infragistics.Win.HeaderStyle.WindowsXPCommand
        Appearance11.BackColor = System.Drawing.SystemColors.Window
        Appearance11.BorderColor = System.Drawing.Color.Silver
        Me.ugColorOption.DisplayLayout.Override.RowAppearance = Appearance11
        Me.ugColorOption.DisplayLayout.Override.RowSelectorHeaderStyle = Infragistics.Win.UltraWinGrid.RowSelectorHeaderStyle.ColumnChooserButton
        Appearance12.BackColor = System.Drawing.SystemColors.ControlLight
        Me.ugColorOption.DisplayLayout.Override.TemplateAddRowAppearance = Appearance12
        Me.ugColorOption.DisplayLayout.ScrollBounds = Infragistics.Win.UltraWinGrid.ScrollBounds.ScrollToFill
        Me.ugColorOption.DisplayLayout.ScrollStyle = Infragistics.Win.UltraWinGrid.ScrollStyle.Immediate
        Me.ugColorOption.DisplayLayout.UseFixedHeaders = True
        Me.ugColorOption.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ugColorOption.Location = New System.Drawing.Point(6, 18)
        Me.ugColorOption.Name = "ugColorOption"
        Me.ugColorOption.Size = New System.Drawing.Size(363, 200)
        Me.ugColorOption.TabIndex = 5
        Me.ugColorOption.Text = "UltraGrid1"
        '
        'UltraDataSource2
        '
        UltraDataColumn1.DefaultValue = ""
        UltraDataColumn2.DataType = GetType(System.Drawing.Color)
        UltraDataColumn3.DataType = GetType(System.Drawing.Color)
        UltraDataColumn4.DataType = GetType(Boolean)
        UltraDataColumn4.DefaultValue = False
        UltraDataColumn5.DataType = GetType(UInteger)
        UltraDataColumn5.DefaultValue = CType(0UI, UInteger)
        Me.UltraDataSource2.Band.Columns.AddRange(New Object() {UltraDataColumn1, UltraDataColumn2, UltraDataColumn3, UltraDataColumn4, UltraDataColumn5})
        '
        'ucpBackColor
        '
        Me.ucpBackColor.Color = System.Drawing.Color.Empty
        Me.ucpBackColor.Location = New System.Drawing.Point(157, 228)
        Me.ucpBackColor.Name = "ucpBackColor"
        Me.ucpBackColor.Size = New System.Drawing.Size(160, 21)
        Me.ucpBackColor.TabIndex = 5
        Me.ucpBackColor.TabStop = False
        Me.ucpBackColor.Visible = False
        '
        'ucpQuoteState7
        '
        Me.ucpQuoteState7.Color = System.Drawing.Color.Empty
        Me.ucpQuoteState7.Location = New System.Drawing.Point(157, 183)
        Me.ucpQuoteState7.Name = "ucpQuoteState7"
        Me.ucpQuoteState7.Size = New System.Drawing.Size(144, 21)
        Me.ucpQuoteState7.TabIndex = 5
        Me.ucpQuoteState7.Visible = False
        '
        'ucpQuoteState6
        '
        Me.ucpQuoteState6.Color = System.Drawing.Color.Empty
        Me.ucpQuoteState6.Location = New System.Drawing.Point(157, 156)
        Me.ucpQuoteState6.Name = "ucpQuoteState6"
        Me.ucpQuoteState6.Size = New System.Drawing.Size(144, 21)
        Me.ucpQuoteState6.TabIndex = 5
        Me.ucpQuoteState6.Visible = False
        '
        'lblColOp7
        '
        Me.lblColOp7.AutoSize = True
        Me.lblColOp7.Location = New System.Drawing.Point(3, 188)
        Me.lblColOp7.Name = "lblColOp7"
        Me.lblColOp7.Size = New System.Drawing.Size(29, 13)
        Me.lblColOp7.TabIndex = 2
        Me.lblColOp7.Text = "Hold"
        Me.lblColOp7.Visible = False
        '
        'lblColOp6
        '
        Me.lblColOp6.AutoSize = True
        Me.lblColOp6.Location = New System.Drawing.Point(3, 161)
        Me.lblColOp6.Name = "lblColOp6"
        Me.lblColOp6.Size = New System.Drawing.Size(57, 13)
        Me.lblColOp6.TabIndex = 2
        Me.lblColOp6.Text = "Cancelled "
        Me.lblColOp6.Visible = False
        '
        'lblColOp5
        '
        Me.lblColOp5.AutoSize = True
        Me.lblColOp5.Location = New System.Drawing.Point(3, 132)
        Me.lblColOp5.Name = "lblColOp5"
        Me.lblColOp5.Size = New System.Drawing.Size(70, 13)
        Me.lblColOp5.TabIndex = 2
        Me.lblColOp5.Text = "Unconfirmed "
        Me.lblColOp5.Visible = False
        '
        'lblColOp4
        '
        Me.lblColOp4.AutoSize = True
        Me.lblColOp4.Location = New System.Drawing.Point(3, 103)
        Me.lblColOp4.Name = "lblColOp4"
        Me.lblColOp4.Size = New System.Drawing.Size(57, 13)
        Me.lblColOp4.TabIndex = 2
        Me.lblColOp4.Text = "Confirmed "
        Me.lblColOp4.Visible = False
        '
        'lblColOp3
        '
        Me.lblColOp3.AutoSize = True
        Me.lblColOp3.Location = New System.Drawing.Point(3, 77)
        Me.lblColOp3.Name = "lblColOp3"
        Me.lblColOp3.Size = New System.Drawing.Size(151, 13)
        Me.lblColOp3.TabIndex = 2
        Me.lblColOp3.Text = "Sent and confirmation pending"
        Me.lblColOp3.Visible = False
        '
        'lblColOp2
        '
        Me.lblColOp2.AutoSize = True
        Me.lblColOp2.Location = New System.Drawing.Point(3, 50)
        Me.lblColOp2.Name = "lblColOp2"
        Me.lblColOp2.Size = New System.Drawing.Size(34, 13)
        Me.lblColOp2.TabIndex = 2
        Me.lblColOp2.Text = "Copy "
        Me.lblColOp2.Visible = False
        '
        'lblColOp1
        '
        Me.lblColOp1.AutoSize = True
        Me.lblColOp1.Location = New System.Drawing.Point(3, 22)
        Me.lblColOp1.Name = "lblColOp1"
        Me.lblColOp1.Size = New System.Drawing.Size(57, 13)
        Me.lblColOp1.TabIndex = 2
        Me.lblColOp1.Text = "Edit mode "
        Me.lblColOp1.Visible = False
        '
        'ucpQuoteState5
        '
        Me.ucpQuoteState5.Color = System.Drawing.Color.Empty
        Me.ucpQuoteState5.Location = New System.Drawing.Point(157, 127)
        Me.ucpQuoteState5.Name = "ucpQuoteState5"
        Me.ucpQuoteState5.Size = New System.Drawing.Size(144, 21)
        Me.ucpQuoteState5.TabIndex = 5
        Me.ucpQuoteState5.Visible = False
        '
        'ucpQuoteState4
        '
        Me.ucpQuoteState4.Color = System.Drawing.Color.Empty
        Me.ucpQuoteState4.Location = New System.Drawing.Point(157, 98)
        Me.ucpQuoteState4.Name = "ucpQuoteState4"
        Me.ucpQuoteState4.Size = New System.Drawing.Size(144, 21)
        Me.ucpQuoteState4.TabIndex = 5
        Me.ucpQuoteState4.Visible = False
        '
        'ucpQuoteState3
        '
        Me.ucpQuoteState3.Color = System.Drawing.Color.Empty
        Me.ucpQuoteState3.Location = New System.Drawing.Point(157, 72)
        Me.ucpQuoteState3.Name = "ucpQuoteState3"
        Me.ucpQuoteState3.Size = New System.Drawing.Size(144, 21)
        Me.ucpQuoteState3.TabIndex = 5
        Me.ucpQuoteState3.Visible = False
        '
        'ucpQuoteState2
        '
        Me.ucpQuoteState2.Color = System.Drawing.Color.Empty
        Me.ucpQuoteState2.Location = New System.Drawing.Point(157, 45)
        Me.ucpQuoteState2.Name = "ucpQuoteState2"
        Me.ucpQuoteState2.Size = New System.Drawing.Size(144, 21)
        Me.ucpQuoteState2.TabIndex = 5
        Me.ucpQuoteState2.Visible = False
        '
        'ucpQuoteState1
        '
        Me.ucpQuoteState1.Color = System.Drawing.Color.Empty
        Me.ucpQuoteState1.Location = New System.Drawing.Point(157, 22)
        Me.ucpQuoteState1.Name = "ucpQuoteState1"
        Me.ucpQuoteState1.Size = New System.Drawing.Size(144, 21)
        Me.ucpQuoteState1.TabIndex = 5
        Me.ucpQuoteState1.Visible = False
        '
        'Panel3
        '
        Me.Panel3.Controls.Add(Me.btnOK)
        Me.Panel3.Controls.Add(Me.btnCancel)
        Me.Panel3.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.Panel3.Location = New System.Drawing.Point(0, 404)
        Me.Panel3.Name = "Panel3"
        Me.Panel3.Size = New System.Drawing.Size(398, 29)
        Me.Panel3.TabIndex = 5
        '
        'GroupBox1
        '
        Me.GroupBox1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox1.Controls.Add(Me.ucmbDefaultTax)
        Me.GroupBox1.Controls.Add(Me.lblDefaultTax)
        Me.GroupBox1.Controls.Add(Me.chkLoadGlobalDefault)
        Me.GroupBox1.Controls.Add(Me.chkGlobalSave)
        Me.GroupBox1.Controls.Add(Me.chkTaxInc)
        Me.GroupBox1.Location = New System.Drawing.Point(12, 12)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(375, 101)
        Me.GroupBox1.TabIndex = 6
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Default Item Price / Rate"
        '
        'ucmbDefaultTax
        '
        Appearance14.BackColor = System.Drawing.SystemColors.Window
        Appearance14.BorderColor = System.Drawing.SystemColors.InactiveCaption
        Me.ucmbDefaultTax.DisplayLayout.Appearance = Appearance14
        Me.ucmbDefaultTax.DisplayLayout.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid
        Me.ucmbDefaultTax.DisplayLayout.CaptionVisible = Infragistics.Win.DefaultableBoolean.[False]
        Appearance15.BackColor = System.Drawing.SystemColors.ActiveBorder
        Appearance15.BackColor2 = System.Drawing.SystemColors.ControlDark
        Appearance15.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical
        Appearance15.BorderColor = System.Drawing.SystemColors.Window
        Me.ucmbDefaultTax.DisplayLayout.GroupByBox.Appearance = Appearance15
        Appearance16.ForeColor = System.Drawing.SystemColors.GrayText
        Me.ucmbDefaultTax.DisplayLayout.GroupByBox.BandLabelAppearance = Appearance16
        Me.ucmbDefaultTax.DisplayLayout.GroupByBox.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid
        Appearance19.BackColor = System.Drawing.SystemColors.ControlLightLight
        Appearance19.BackColor2 = System.Drawing.SystemColors.Control
        Appearance19.BackGradientStyle = Infragistics.Win.GradientStyle.Horizontal
        Appearance19.ForeColor = System.Drawing.SystemColors.GrayText
        Me.ucmbDefaultTax.DisplayLayout.GroupByBox.PromptAppearance = Appearance19
        Me.ucmbDefaultTax.DisplayLayout.MaxColScrollRegions = 1
        Me.ucmbDefaultTax.DisplayLayout.MaxRowScrollRegions = 1
        Appearance20.BackColor = System.Drawing.SystemColors.Window
        Appearance20.ForeColor = System.Drawing.SystemColors.ControlText
        Me.ucmbDefaultTax.DisplayLayout.Override.ActiveCellAppearance = Appearance20
        Appearance21.BackColor = System.Drawing.SystemColors.Highlight
        Appearance21.ForeColor = System.Drawing.SystemColors.HighlightText
        Me.ucmbDefaultTax.DisplayLayout.Override.ActiveRowAppearance = Appearance21
        Me.ucmbDefaultTax.DisplayLayout.Override.BorderStyleCell = Infragistics.Win.UIElementBorderStyle.Dotted
        Me.ucmbDefaultTax.DisplayLayout.Override.BorderStyleRow = Infragistics.Win.UIElementBorderStyle.Dotted
        Appearance22.BackColor = System.Drawing.SystemColors.Window
        Me.ucmbDefaultTax.DisplayLayout.Override.CardAreaAppearance = Appearance22
        Appearance23.BorderColor = System.Drawing.Color.Silver
        Appearance23.TextTrimming = Infragistics.Win.TextTrimming.EllipsisCharacter
        Me.ucmbDefaultTax.DisplayLayout.Override.CellAppearance = Appearance23
        Me.ucmbDefaultTax.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.EditAndSelectText
        Me.ucmbDefaultTax.DisplayLayout.Override.CellPadding = 0
        Appearance24.BackColor = System.Drawing.SystemColors.Control
        Appearance24.BackColor2 = System.Drawing.SystemColors.ControlDark
        Appearance24.BackGradientAlignment = Infragistics.Win.GradientAlignment.Element
        Appearance24.BackGradientStyle = Infragistics.Win.GradientStyle.Horizontal
        Appearance24.BorderColor = System.Drawing.SystemColors.Window
        Me.ucmbDefaultTax.DisplayLayout.Override.GroupByRowAppearance = Appearance24
        Appearance25.TextHAlignAsString = "Left"
        Me.ucmbDefaultTax.DisplayLayout.Override.HeaderAppearance = Appearance25
        Me.ucmbDefaultTax.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti
        Me.ucmbDefaultTax.DisplayLayout.Override.HeaderStyle = Infragistics.Win.HeaderStyle.WindowsXPCommand
        Appearance26.BackColor = System.Drawing.SystemColors.Window
        Appearance26.BorderColor = System.Drawing.Color.Silver
        Me.ucmbDefaultTax.DisplayLayout.Override.RowAppearance = Appearance26
        Me.ucmbDefaultTax.DisplayLayout.Override.RowSelectors = Infragistics.Win.DefaultableBoolean.[False]
        Appearance27.BackColor = System.Drawing.SystemColors.ControlLight
        Me.ucmbDefaultTax.DisplayLayout.Override.TemplateAddRowAppearance = Appearance27
        Me.ucmbDefaultTax.DisplayLayout.ScrollBounds = Infragistics.Win.UltraWinGrid.ScrollBounds.ScrollToFill
        Me.ucmbDefaultTax.DisplayLayout.ScrollStyle = Infragistics.Win.UltraWinGrid.ScrollStyle.Immediate
        Me.ucmbDefaultTax.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy
        Me.ucmbDefaultTax.Location = New System.Drawing.Point(92, 27)
        Me.ucmbDefaultTax.Name = "ucmbDefaultTax"
        Me.ucmbDefaultTax.Size = New System.Drawing.Size(277, 22)
        Me.ucmbDefaultTax.TabIndex = 4
        '
        'lblDefaultTax
        '
        Me.lblDefaultTax.AutoSize = True
        Me.lblDefaultTax.Location = New System.Drawing.Point(3, 32)
        Me.lblDefaultTax.Name = "lblDefaultTax"
        Me.lblDefaultTax.Size = New System.Drawing.Size(79, 13)
        Me.lblDefaultTax.TabIndex = 2
        Me.lblDefaultTax.Text = "Default tax rate"
        '
        'chkLoadGlobalDefault
        '
        Me.chkLoadGlobalDefault.AutoSize = True
        Me.chkLoadGlobalDefault.Location = New System.Drawing.Point(202, 122)
        Me.chkLoadGlobalDefault.Name = "chkLoadGlobalDefault"
        Me.chkLoadGlobalDefault.Size = New System.Drawing.Size(173, 17)
        Me.chkLoadGlobalDefault.TabIndex = 0
        Me.chkLoadGlobalDefault.Text = "Load global settings if available"
        Me.chkLoadGlobalDefault.UseVisualStyleBackColor = True
        Me.chkLoadGlobalDefault.Visible = False
        '
        'chkGlobalSave
        '
        Me.chkGlobalSave.AutoSize = True
        Me.chkGlobalSave.Location = New System.Drawing.Point(202, 135)
        Me.chkGlobalSave.Name = "chkGlobalSave"
        Me.chkGlobalSave.Size = New System.Drawing.Size(107, 17)
        Me.chkGlobalSave.TabIndex = 0
        Me.chkGlobalSave.Text = "Save for all users"
        Me.chkGlobalSave.UseVisualStyleBackColor = True
        Me.chkGlobalSave.Visible = False
        '
        'chkTaxInc
        '
        Me.chkTaxInc.AutoSize = True
        Me.chkTaxInc.Location = New System.Drawing.Point(6, 74)
        Me.chkTaxInc.Name = "chkTaxInc"
        Me.chkTaxInc.Size = New System.Drawing.Size(134, 17)
        Me.chkTaxInc.TabIndex = 0
        Me.chkTaxInc.Text = "Prices are tax inclusive"
        Me.chkTaxInc.UseVisualStyleBackColor = True
        '
        'UltraDataSource1
        '
        Me.UltraDataSource1.Band.Columns.AddRange(New Object() {UltraDataColumn6, UltraDataColumn7})
        '
        'frmGlazingDocDefaultSetting
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btnCancel
        Me.ClientSize = New System.Drawing.Size(398, 433)
        Me.Controls.Add(Me.Panel1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmGlazingDocDefaultSetting"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Options"
        Me.Panel1.ResumeLayout(False)
        Me.gbColorOptions.ResumeLayout(False)
        Me.gbColorOptions.PerformLayout()
        CType(Me.ugColorOption, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.UltraDataSource2, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.ucpBackColor, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.ucpQuoteState7, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.ucpQuoteState6, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.ucpQuoteState5, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.ucpQuoteState4, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.ucpQuoteState3, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.ucpQuoteState2, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.ucpQuoteState1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel3.ResumeLayout(False)
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        CType(Me.ucmbDefaultTax, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.UltraDataSource1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents btnOK As System.Windows.Forms.Button
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents Panel3 As System.Windows.Forms.Panel
    Friend WithEvents chkTaxInc As System.Windows.Forms.CheckBox
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents lblDefaultTax As System.Windows.Forms.Label
    Friend WithEvents ucmbDefaultTax As Infragistics.Win.UltraWinGrid.UltraCombo
    Friend WithEvents chkGlobalSave As System.Windows.Forms.CheckBox
    Friend WithEvents chkLoadGlobalDefault As System.Windows.Forms.CheckBox
    Friend WithEvents gbColorOptions As System.Windows.Forms.GroupBox
    Friend WithEvents ucpQuoteState1 As Infragistics.Win.UltraWinEditors.UltraColorPicker
    Friend WithEvents lblColOp1 As System.Windows.Forms.Label
    Friend WithEvents ucpQuoteState7 As Infragistics.Win.UltraWinEditors.UltraColorPicker
    Friend WithEvents ucpQuoteState6 As Infragistics.Win.UltraWinEditors.UltraColorPicker
    Friend WithEvents ucpQuoteState5 As Infragistics.Win.UltraWinEditors.UltraColorPicker
    Friend WithEvents ucpQuoteState4 As Infragistics.Win.UltraWinEditors.UltraColorPicker
    Friend WithEvents ucpQuoteState3 As Infragistics.Win.UltraWinEditors.UltraColorPicker
    Friend WithEvents ucpQuoteState2 As Infragistics.Win.UltraWinEditors.UltraColorPicker
    Friend WithEvents lblColOp5 As System.Windows.Forms.Label
    Friend WithEvents lblColOp4 As System.Windows.Forms.Label
    Friend WithEvents lblColOp3 As System.Windows.Forms.Label
    Friend WithEvents lblColOp2 As System.Windows.Forms.Label
    Friend WithEvents lblColOp7 As System.Windows.Forms.Label
    Friend WithEvents lblColOp6 As System.Windows.Forms.Label
    Friend WithEvents ucpBackColor As Infragistics.Win.UltraWinEditors.UltraColorPicker
    Friend WithEvents ugColorOption As Infragistics.Win.UltraWinGrid.UltraGrid
    Friend WithEvents UltraDataSource2 As Infragistics.Win.UltraWinDataSource.UltraDataSource
    Friend WithEvents UltraDataSource1 As Infragistics.Win.UltraWinDataSource.UltraDataSource
    Friend WithEvents chkDefaultBackCol As System.Windows.Forms.CheckBox
End Class
