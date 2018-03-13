<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmGlazingAddessLocatorGmap
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
        Me.addressMap = New GMap.NET.WindowsForms.GMapControl()
        Me.btnOK = New System.Windows.Forms.Button()
        Me.lblAddress = New System.Windows.Forms.Label()
        Me.utxtAddress = New Infragistics.Win.UltraWinEditors.UltraTextEditor()
        Me.btnSearch = New System.Windows.Forms.Button()
        CType(Me.utxtAddress, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'addressMap
        '
        Me.addressMap.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.addressMap.Bearing = 0.0!
        Me.addressMap.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.addressMap.CanDragMap = True
        Me.addressMap.EmptyTileColor = System.Drawing.Color.Navy
        Me.addressMap.ForeColor = System.Drawing.Color.Black
        Me.addressMap.GrayScaleMode = False
        Me.addressMap.HelperLineOption = GMap.NET.WindowsForms.HelperLineOptions.DontShow
        Me.addressMap.LevelsKeepInMemmory = 5
        Me.addressMap.Location = New System.Drawing.Point(0, 57)
        Me.addressMap.MarkersEnabled = True
        Me.addressMap.MaxZoom = 2
        Me.addressMap.MinZoom = 2
        Me.addressMap.MouseWheelZoomType = GMap.NET.MouseWheelZoomType.MousePositionAndCenter
        Me.addressMap.Name = "addressMap"
        Me.addressMap.NegativeMode = False
        Me.addressMap.PolygonsEnabled = True
        Me.addressMap.RetryLoadTile = 0
        Me.addressMap.RoutesEnabled = True
        Me.addressMap.ScaleMode = GMap.NET.WindowsForms.ScaleModes.[Integer]
        Me.addressMap.SelectedAreaFillColor = System.Drawing.Color.FromArgb(CType(CType(33, Byte), Integer), CType(CType(65, Byte), Integer), CType(CType(105, Byte), Integer), CType(CType(225, Byte), Integer))
        Me.addressMap.ShowTileGridLines = False
        Me.addressMap.Size = New System.Drawing.Size(774, 426)
        Me.addressMap.TabIndex = 27
        Me.addressMap.Zoom = 0.0R
        '
        'btnOK
        '
        Me.btnOK.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnOK.BackColor = System.Drawing.Color.DodgerBlue
        Me.btnOK.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.btnOK.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnOK.ForeColor = System.Drawing.Color.White
        Me.btnOK.Location = New System.Drawing.Point(698, 28)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(75, 23)
        Me.btnOK.TabIndex = 29
        Me.btnOK.Text = "OK"
        Me.btnOK.UseVisualStyleBackColor = False
        '
        'lblAddress
        '
        Me.lblAddress.Location = New System.Drawing.Point(2, 33)
        Me.lblAddress.Name = "lblAddress"
        Me.lblAddress.Size = New System.Drawing.Size(691, 15)
        Me.lblAddress.TabIndex = 0
        Me.lblAddress.Text = "Select a location"
        '
        'utxtAddress
        '
        Me.utxtAddress.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.utxtAddress.Location = New System.Drawing.Point(1, 2)
        Me.utxtAddress.Name = "utxtAddress"
        Me.utxtAddress.Size = New System.Drawing.Size(693, 21)
        Me.utxtAddress.TabIndex = 30
        '
        'btnSearch
        '
        Me.btnSearch.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnSearch.BackColor = System.Drawing.Color.DodgerBlue
        Me.btnSearch.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnSearch.ForeColor = System.Drawing.Color.White
        Me.btnSearch.Location = New System.Drawing.Point(698, 1)
        Me.btnSearch.Name = "btnSearch"
        Me.btnSearch.Size = New System.Drawing.Size(75, 23)
        Me.btnSearch.TabIndex = 29
        Me.btnSearch.Text = "Search"
        Me.btnSearch.UseVisualStyleBackColor = False
        '
        'frmGlazingAddessLocatorGmap
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(774, 483)
        Me.Controls.Add(Me.utxtAddress)
        Me.Controls.Add(Me.lblAddress)
        Me.Controls.Add(Me.btnSearch)
        Me.Controls.Add(Me.btnOK)
        Me.Controls.Add(Me.addressMap)
        Me.MinimizeBox = False
        Me.Name = "frmGlazingAddessLocatorGmap"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        CType(Me.utxtAddress, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Private WithEvents addressMap As GMap.NET.WindowsForms.GMapControl
    Friend WithEvents btnOK As System.Windows.Forms.Button
    Friend WithEvents lblAddress As System.Windows.Forms.Label
    Friend WithEvents utxtAddress As Infragistics.Win.UltraWinEditors.UltraTextEditor
    Friend WithEvents btnSearch As System.Windows.Forms.Button
End Class
