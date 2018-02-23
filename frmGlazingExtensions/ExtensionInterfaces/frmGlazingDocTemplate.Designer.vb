<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmGlazingDocTemplate
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
        Dim Appearance18 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim UltraGridBand1 As Infragistics.Win.UltraWinGrid.UltraGridBand = New Infragistics.Win.UltraWinGrid.UltraGridBand("", -1)
        Dim Appearance19 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance21 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance20 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance24 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance10 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance12 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance25 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance23 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance22 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance11 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance6 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmGlazingDocTemplate))
        Me.btnSave = New System.Windows.Forms.Button()
        Me.txttmpName = New System.Windows.Forms.TextBox()
        Me.lblTName = New System.Windows.Forms.Label()
        Me.btnCancel = New System.Windows.Forms.Button()
        Me.ugQuoteList = New Infragistics.Win.UltraWinGrid.UltraGrid()
        Me.cmQuoteList = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.tsmRename = New System.Windows.Forms.ToolStripMenuItem()
        Me.tsmDelete = New System.Windows.Forms.ToolStripMenuItem()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.btnLoad = New System.Windows.Forms.Button()
        Me.btgnClear = New System.Windows.Forms.Button()
        Me.btnDelete = New System.Windows.Forms.Button()
        CType(Me.ugQuoteList, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.cmQuoteList.SuspendLayout()
        Me.SuspendLayout()
        '
        'btnSave
        '
        Me.btnSave.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnSave.BackColor = System.Drawing.Color.DodgerBlue
        Me.btnSave.FlatAppearance.BorderColor = System.Drawing.Color.DodgerBlue
        Me.btnSave.FlatAppearance.MouseDownBackColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(64, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.btnSave.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.btnSave.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnSave.ForeColor = System.Drawing.Color.White
        Me.btnSave.Location = New System.Drawing.Point(15, 299)
        Me.btnSave.Name = "btnSave"
        Me.btnSave.Size = New System.Drawing.Size(81, 23)
        Me.btnSave.TabIndex = 0
        Me.btnSave.Text = "Save"
        Me.btnSave.UseVisualStyleBackColor = False
        '
        'txttmpName
        '
        Me.txttmpName.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.txttmpName.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txttmpName.Location = New System.Drawing.Point(108, 266)
        Me.txttmpName.Name = "txttmpName"
        Me.txttmpName.Size = New System.Drawing.Size(238, 20)
        Me.txttmpName.TabIndex = 3
        '
        'lblTName
        '
        Me.lblTName.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.lblTName.AutoSize = True
        Me.lblTName.Location = New System.Drawing.Point(12, 271)
        Me.lblTName.Name = "lblTName"
        Me.lblTName.Size = New System.Drawing.Size(82, 13)
        Me.lblTName.TabIndex = 4
        Me.lblTName.Text = "Template Name"
        '
        'btnCancel
        '
        Me.btnCancel.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnCancel.BackColor = System.Drawing.Color.DodgerBlue
        Me.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnCancel.FlatAppearance.BorderColor = System.Drawing.Color.DodgerBlue
        Me.btnCancel.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Red
        Me.btnCancel.FlatAppearance.MouseOverBackColor = System.Drawing.Color.Maroon
        Me.btnCancel.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnCancel.ForeColor = System.Drawing.Color.White
        Me.btnCancel.Location = New System.Drawing.Point(289, 299)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(57, 52)
        Me.btnCancel.TabIndex = 0
        Me.btnCancel.Text = "Close"
        Me.btnCancel.UseVisualStyleBackColor = False
        '
        'ugQuoteList
        '
        Me.ugQuoteList.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.ugQuoteList.ContextMenuStrip = Me.cmQuoteList
        Appearance18.BackColor = System.Drawing.SystemColors.Window
        Appearance18.BorderColor = System.Drawing.SystemColors.InactiveCaption
        Me.ugQuoteList.DisplayLayout.Appearance = Appearance18
        Me.ugQuoteList.DisplayLayout.AutoFitStyle = Infragistics.Win.UltraWinGrid.AutoFitStyle.ResizeAllColumns
        UltraGridBand1.Header.Enabled = False
        Me.ugQuoteList.DisplayLayout.BandsSerializer.Add(UltraGridBand1)
        Me.ugQuoteList.DisplayLayout.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid
        Me.ugQuoteList.DisplayLayout.CaptionVisible = Infragistics.Win.DefaultableBoolean.[False]
        Appearance19.BackColor = System.Drawing.SystemColors.ActiveBorder
        Appearance19.BackColor2 = System.Drawing.SystemColors.ControlDark
        Appearance19.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical
        Appearance19.BorderColor = System.Drawing.SystemColors.Window
        Me.ugQuoteList.DisplayLayout.GroupByBox.Appearance = Appearance19
        Appearance21.ForeColor = System.Drawing.SystemColors.GrayText
        Me.ugQuoteList.DisplayLayout.GroupByBox.BandLabelAppearance = Appearance21
        Me.ugQuoteList.DisplayLayout.GroupByBox.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid
        Appearance20.BackColor = System.Drawing.SystemColors.ControlLightLight
        Appearance20.BackColor2 = System.Drawing.SystemColors.Control
        Appearance20.BackGradientStyle = Infragistics.Win.GradientStyle.Horizontal
        Appearance20.ForeColor = System.Drawing.SystemColors.GrayText
        Me.ugQuoteList.DisplayLayout.GroupByBox.PromptAppearance = Appearance20
        Me.ugQuoteList.DisplayLayout.MaxColScrollRegions = 1
        Me.ugQuoteList.DisplayLayout.MaxRowScrollRegions = 1
        Appearance24.BackColor = System.Drawing.SystemColors.Window
        Appearance24.ForeColor = System.Drawing.SystemColors.ControlText
        Me.ugQuoteList.DisplayLayout.Override.ActiveCellAppearance = Appearance24
        Appearance10.BackColor = System.Drawing.SystemColors.Highlight
        Appearance10.ForeColor = System.Drawing.SystemColors.HighlightText
        Me.ugQuoteList.DisplayLayout.Override.ActiveRowAppearance = Appearance10
        Me.ugQuoteList.DisplayLayout.Override.BorderStyleCell = Infragistics.Win.UIElementBorderStyle.Dotted
        Me.ugQuoteList.DisplayLayout.Override.BorderStyleRow = Infragistics.Win.UIElementBorderStyle.Dotted
        Appearance12.BackColor = System.Drawing.SystemColors.Window
        Me.ugQuoteList.DisplayLayout.Override.CardAreaAppearance = Appearance12
        Appearance25.BorderColor = System.Drawing.Color.Silver
        Appearance25.TextTrimming = Infragistics.Win.TextTrimming.EllipsisCharacter
        Me.ugQuoteList.DisplayLayout.Override.CellAppearance = Appearance25
        Me.ugQuoteList.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.EditAndSelectText
        Me.ugQuoteList.DisplayLayout.Override.CellPadding = 0
        Appearance23.BackColor = System.Drawing.SystemColors.Control
        Appearance23.BackColor2 = System.Drawing.SystemColors.ControlDark
        Appearance23.BackGradientAlignment = Infragistics.Win.GradientAlignment.Element
        Appearance23.BackGradientStyle = Infragistics.Win.GradientStyle.Horizontal
        Appearance23.BorderColor = System.Drawing.SystemColors.Window
        Me.ugQuoteList.DisplayLayout.Override.GroupByRowAppearance = Appearance23
        Appearance22.TextHAlignAsString = "Left"
        Me.ugQuoteList.DisplayLayout.Override.HeaderAppearance = Appearance22
        Me.ugQuoteList.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti
        Me.ugQuoteList.DisplayLayout.Override.HeaderStyle = Infragistics.Win.HeaderStyle.WindowsXPCommand
        Appearance11.BackColor = System.Drawing.SystemColors.Window
        Appearance11.BorderColor = System.Drawing.Color.Silver
        Me.ugQuoteList.DisplayLayout.Override.RowAppearance = Appearance11
        Me.ugQuoteList.DisplayLayout.Override.RowSelectors = Infragistics.Win.DefaultableBoolean.[False]
        Appearance6.BackColor = System.Drawing.SystemColors.ControlLight
        Me.ugQuoteList.DisplayLayout.Override.TemplateAddRowAppearance = Appearance6
        Me.ugQuoteList.DisplayLayout.ScrollBounds = Infragistics.Win.UltraWinGrid.ScrollBounds.ScrollToFill
        Me.ugQuoteList.DisplayLayout.ScrollStyle = Infragistics.Win.UltraWinGrid.ScrollStyle.Immediate
        Me.ugQuoteList.Location = New System.Drawing.Point(12, 29)
        Me.ugQuoteList.Name = "ugQuoteList"
        Me.ugQuoteList.Size = New System.Drawing.Size(334, 231)
        Me.ugQuoteList.TabIndex = 5
        Me.ugQuoteList.Text = "UltraGrid1"
        Me.ugQuoteList.UseFlatMode = Infragistics.Win.DefaultableBoolean.[True]
        Me.ugQuoteList.UseOsThemes = Infragistics.Win.DefaultableBoolean.[True]
        '
        'cmQuoteList
        '
        Me.cmQuoteList.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.tsmRename, Me.tsmDelete})
        Me.cmQuoteList.Name = "cmQuoteList"
        Me.cmQuoteList.Size = New System.Drawing.Size(118, 48)
        '
        'tsmRename
        '
        Me.tsmRename.Name = "tsmRename"
        Me.tsmRename.Size = New System.Drawing.Size(117, 22)
        Me.tsmRename.Text = "Rename"
        '
        'tsmDelete
        '
        Me.tsmDelete.Name = "tsmDelete"
        Me.tsmDelete.Size = New System.Drawing.Size(117, 22)
        Me.tsmDelete.Text = "Delete"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(12, 9)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(102, 13)
        Me.Label1.TabIndex = 4
        Me.Label1.Text = "Available Templates"
        '
        'btnLoad
        '
        Me.btnLoad.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnLoad.BackColor = System.Drawing.Color.DodgerBlue
        Me.btnLoad.FlatAppearance.BorderColor = System.Drawing.Color.DodgerBlue
        Me.btnLoad.FlatAppearance.MouseDownBackColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.btnLoad.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(64, Byte), Integer), CType(CType(64, Byte), Integer))
        Me.btnLoad.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnLoad.ForeColor = System.Drawing.Color.White
        Me.btnLoad.Location = New System.Drawing.Point(102, 299)
        Me.btnLoad.Name = "btnLoad"
        Me.btnLoad.Size = New System.Drawing.Size(81, 23)
        Me.btnLoad.TabIndex = 0
        Me.btnLoad.Text = "Load"
        Me.btnLoad.UseVisualStyleBackColor = False
        '
        'btgnClear
        '
        Me.btgnClear.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btgnClear.BackColor = System.Drawing.Color.DodgerBlue
        Me.btgnClear.FlatAppearance.BorderColor = System.Drawing.Color.DodgerBlue
        Me.btgnClear.FlatAppearance.MouseDownBackColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.btgnClear.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(64, Byte), Integer), CType(CType(64, Byte), Integer))
        Me.btgnClear.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btgnClear.ForeColor = System.Drawing.Color.White
        Me.btgnClear.Location = New System.Drawing.Point(102, 328)
        Me.btgnClear.Name = "btgnClear"
        Me.btgnClear.Size = New System.Drawing.Size(81, 23)
        Me.btgnClear.TabIndex = 0
        Me.btgnClear.Text = "Clear Gride"
        Me.btgnClear.UseVisualStyleBackColor = False
        '
        'btnDelete
        '
        Me.btnDelete.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnDelete.BackColor = System.Drawing.Color.DodgerBlue
        Me.btnDelete.FlatAppearance.BorderColor = System.Drawing.Color.DodgerBlue
        Me.btnDelete.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Red
        Me.btnDelete.FlatAppearance.MouseOverBackColor = System.Drawing.Color.Maroon
        Me.btnDelete.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnDelete.ForeColor = System.Drawing.Color.White
        Me.btnDelete.Location = New System.Drawing.Point(15, 328)
        Me.btnDelete.Name = "btnDelete"
        Me.btnDelete.Size = New System.Drawing.Size(81, 23)
        Me.btnDelete.TabIndex = 0
        Me.btnDelete.Text = "Delete"
        Me.btnDelete.UseVisualStyleBackColor = False
        '
        'frmGlazingDocTemplate
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btnCancel
        Me.ClientSize = New System.Drawing.Size(358, 363)
        Me.Controls.Add(Me.ugQuoteList)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.lblTName)
        Me.Controls.Add(Me.txttmpName)
        Me.Controls.Add(Me.btnDelete)
        Me.Controls.Add(Me.btnCancel)
        Me.Controls.Add(Me.btgnClear)
        Me.Controls.Add(Me.btnLoad)
        Me.Controls.Add(Me.btnSave)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "frmGlazingDocTemplate"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Glazing Quote Template"
        CType(Me.ugQuoteList, System.ComponentModel.ISupportInitialize).EndInit()
        Me.cmQuoteList.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents btnSave As System.Windows.Forms.Button
    Friend WithEvents txttmpName As System.Windows.Forms.TextBox
    Friend WithEvents lblTName As System.Windows.Forms.Label
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents ugQuoteList As Infragistics.Win.UltraWinGrid.UltraGrid
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents btnLoad As System.Windows.Forms.Button
    Friend WithEvents btgnClear As System.Windows.Forms.Button
    Friend WithEvents btnDelete As System.Windows.Forms.Button
    Friend WithEvents cmQuoteList As System.Windows.Forms.ContextMenuStrip
    Friend WithEvents tsmRename As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents tsmDelete As System.Windows.Forms.ToolStripMenuItem
End Class
