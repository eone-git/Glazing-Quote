<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmGlazingNote
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
        Dim UltraToolbar2 As Infragistics.Win.UltraWinToolbars.UltraToolbar = New Infragistics.Win.UltraWinToolbars.UltraToolbar("uttNoteText")
        Me.uscGlassQuote = New Infragistics.Win.UltraWinSpellChecker.UltraSpellChecker(Me.components)
        Me.brnAdd = New System.Windows.Forms.Button()
        Me.UltraFormattedTextWordWriter1 = New Infragistics.Win.UltraWinFormattedText.WordWriter.UltraFormattedTextWordWriter(Me.components)
        Me.btnSpellCheck = New System.Windows.Forms.Button()
        Me.UltraFormattedTextWordWriter2 = New Infragistics.Win.UltraWinFormattedText.WordWriter.UltraFormattedTextWordWriter(Me.components)
        Me.frmGlazingNote_Fill_Panel = New Infragistics.Win.Misc.UltraPanel()
        Me.utxtNoteText = New Infragistics.Win.UltraWinEditors.UltraTextEditor()
        Me._frmGlazingNote_Toolbars_Dock_Area_Left = New Infragistics.Win.UltraWinToolbars.UltraToolbarsDockArea()
        Me.UltraToolbarsManager1 = New Infragistics.Win.UltraWinToolbars.UltraToolbarsManager(Me.components)
        Me._frmGlazingNote_Toolbars_Dock_Area_Right = New Infragistics.Win.UltraWinToolbars.UltraToolbarsDockArea()
        Me._frmGlazingNote_Toolbars_Dock_Area_Top = New Infragistics.Win.UltraWinToolbars.UltraToolbarsDockArea()
        Me._frmGlazingNote_Toolbars_Dock_Area_Bottom = New Infragistics.Win.UltraWinToolbars.UltraToolbarsDockArea()
        Me.ContextMenuStrip1 = New System.Windows.Forms.ContextMenuStrip(Me.components)
        CType(Me.uscGlassQuote, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.frmGlazingNote_Fill_Panel.ClientArea.SuspendLayout()
        Me.frmGlazingNote_Fill_Panel.SuspendLayout()
        CType(Me.utxtNoteText, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.UltraToolbarsManager1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'uscGlassQuote
        '
        Me.uscGlassQuote.ContainingControl = Me
        '
        'brnAdd
        '
        Me.brnAdd.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.brnAdd.BackColor = System.Drawing.Color.DodgerBlue
        Me.brnAdd.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.brnAdd.FlatAppearance.BorderColor = System.Drawing.Color.DodgerBlue
        Me.brnAdd.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.brnAdd.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.brnAdd.ForeColor = System.Drawing.Color.White
        Me.brnAdd.Location = New System.Drawing.Point(906, 388)
        Me.brnAdd.Name = "brnAdd"
        Me.brnAdd.Size = New System.Drawing.Size(75, 24)
        Me.brnAdd.TabIndex = 2
        Me.brnAdd.Text = "OK"
        Me.brnAdd.UseVisualStyleBackColor = False
        '
        'btnSpellCheck
        '
        Me.btnSpellCheck.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnSpellCheck.BackColor = System.Drawing.Color.DodgerBlue
        Me.btnSpellCheck.FlatAppearance.BorderColor = System.Drawing.Color.DodgerBlue
        Me.btnSpellCheck.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnSpellCheck.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnSpellCheck.ForeColor = System.Drawing.Color.White
        Me.btnSpellCheck.Location = New System.Drawing.Point(825, 388)
        Me.btnSpellCheck.Name = "btnSpellCheck"
        Me.btnSpellCheck.Size = New System.Drawing.Size(75, 24)
        Me.btnSpellCheck.TabIndex = 2
        Me.btnSpellCheck.Text = "Spell Check"
        Me.btnSpellCheck.UseVisualStyleBackColor = False
        Me.btnSpellCheck.Visible = False
        '
        'frmGlazingNote_Fill_Panel
        '
        '
        'frmGlazingNote_Fill_Panel.ClientArea
        '
        Me.frmGlazingNote_Fill_Panel.ClientArea.Controls.Add(Me.utxtNoteText)
        Me.frmGlazingNote_Fill_Panel.ClientArea.Controls.Add(Me.btnSpellCheck)
        Me.frmGlazingNote_Fill_Panel.ClientArea.Controls.Add(Me.brnAdd)
        Me.frmGlazingNote_Fill_Panel.Cursor = System.Windows.Forms.Cursors.Default
        Me.frmGlazingNote_Fill_Panel.Dock = System.Windows.Forms.DockStyle.Fill
        Me.frmGlazingNote_Fill_Panel.Location = New System.Drawing.Point(0, 17)
        Me.frmGlazingNote_Fill_Panel.Name = "frmGlazingNote_Fill_Panel"
        Me.frmGlazingNote_Fill_Panel.Size = New System.Drawing.Size(984, 415)
        Me.frmGlazingNote_Fill_Panel.TabIndex = 0
        '
        'utxtNoteText
        '
        Me.utxtNoteText.Location = New System.Drawing.Point(3, 6)
        Me.utxtNoteText.Multiline = True
        Me.utxtNoteText.Name = "utxtNoteText"
        Me.utxtNoteText.Size = New System.Drawing.Size(978, 380)
        Me.utxtNoteText.TabIndex = 3
        '
        '_frmGlazingNote_Toolbars_Dock_Area_Left
        '
        Me._frmGlazingNote_Toolbars_Dock_Area_Left.AccessibleRole = System.Windows.Forms.AccessibleRole.Grouping
        Me._frmGlazingNote_Toolbars_Dock_Area_Left.BackColor = System.Drawing.SystemColors.Control
        Me._frmGlazingNote_Toolbars_Dock_Area_Left.DockedPosition = Infragistics.Win.UltraWinToolbars.DockedPosition.Left
        Me._frmGlazingNote_Toolbars_Dock_Area_Left.ForeColor = System.Drawing.SystemColors.ControlText
        Me._frmGlazingNote_Toolbars_Dock_Area_Left.Location = New System.Drawing.Point(0, 17)
        Me._frmGlazingNote_Toolbars_Dock_Area_Left.Name = "_frmGlazingNote_Toolbars_Dock_Area_Left"
        Me._frmGlazingNote_Toolbars_Dock_Area_Left.Size = New System.Drawing.Size(0, 415)
        Me._frmGlazingNote_Toolbars_Dock_Area_Left.ToolbarsManager = Me.UltraToolbarsManager1
        '
        'UltraToolbarsManager1
        '
        Me.UltraToolbarsManager1.DesignerFlags = 1
        Me.UltraToolbarsManager1.DockWithinContainer = Me
        Me.UltraToolbarsManager1.DockWithinContainerBaseType = GetType(System.Windows.Forms.Form)
        UltraToolbar2.DockedColumn = 0
        UltraToolbar2.DockedRow = 0
        UltraToolbar2.IsMainMenuBar = True
        UltraToolbar2.Text = "uttNoteText"
        Me.UltraToolbarsManager1.Toolbars.AddRange(New Infragistics.Win.UltraWinToolbars.UltraToolbar() {UltraToolbar2})
        '
        '_frmGlazingNote_Toolbars_Dock_Area_Right
        '
        Me._frmGlazingNote_Toolbars_Dock_Area_Right.AccessibleRole = System.Windows.Forms.AccessibleRole.Grouping
        Me._frmGlazingNote_Toolbars_Dock_Area_Right.BackColor = System.Drawing.SystemColors.Control
        Me._frmGlazingNote_Toolbars_Dock_Area_Right.DockedPosition = Infragistics.Win.UltraWinToolbars.DockedPosition.Right
        Me._frmGlazingNote_Toolbars_Dock_Area_Right.ForeColor = System.Drawing.SystemColors.ControlText
        Me._frmGlazingNote_Toolbars_Dock_Area_Right.Location = New System.Drawing.Point(984, 17)
        Me._frmGlazingNote_Toolbars_Dock_Area_Right.Name = "_frmGlazingNote_Toolbars_Dock_Area_Right"
        Me._frmGlazingNote_Toolbars_Dock_Area_Right.Size = New System.Drawing.Size(0, 415)
        Me._frmGlazingNote_Toolbars_Dock_Area_Right.ToolbarsManager = Me.UltraToolbarsManager1
        '
        '_frmGlazingNote_Toolbars_Dock_Area_Top
        '
        Me._frmGlazingNote_Toolbars_Dock_Area_Top.AccessibleRole = System.Windows.Forms.AccessibleRole.Grouping
        Me._frmGlazingNote_Toolbars_Dock_Area_Top.BackColor = System.Drawing.SystemColors.Control
        Me._frmGlazingNote_Toolbars_Dock_Area_Top.DockedPosition = Infragistics.Win.UltraWinToolbars.DockedPosition.Top
        Me._frmGlazingNote_Toolbars_Dock_Area_Top.ForeColor = System.Drawing.SystemColors.ControlText
        Me._frmGlazingNote_Toolbars_Dock_Area_Top.Location = New System.Drawing.Point(0, 0)
        Me._frmGlazingNote_Toolbars_Dock_Area_Top.Name = "_frmGlazingNote_Toolbars_Dock_Area_Top"
        Me._frmGlazingNote_Toolbars_Dock_Area_Top.Size = New System.Drawing.Size(984, 17)
        Me._frmGlazingNote_Toolbars_Dock_Area_Top.ToolbarsManager = Me.UltraToolbarsManager1
        '
        '_frmGlazingNote_Toolbars_Dock_Area_Bottom
        '
        Me._frmGlazingNote_Toolbars_Dock_Area_Bottom.AccessibleRole = System.Windows.Forms.AccessibleRole.Grouping
        Me._frmGlazingNote_Toolbars_Dock_Area_Bottom.BackColor = System.Drawing.SystemColors.Control
        Me._frmGlazingNote_Toolbars_Dock_Area_Bottom.DockedPosition = Infragistics.Win.UltraWinToolbars.DockedPosition.Bottom
        Me._frmGlazingNote_Toolbars_Dock_Area_Bottom.ForeColor = System.Drawing.SystemColors.ControlText
        Me._frmGlazingNote_Toolbars_Dock_Area_Bottom.Location = New System.Drawing.Point(0, 432)
        Me._frmGlazingNote_Toolbars_Dock_Area_Bottom.Name = "_frmGlazingNote_Toolbars_Dock_Area_Bottom"
        Me._frmGlazingNote_Toolbars_Dock_Area_Bottom.Size = New System.Drawing.Size(984, 0)
        Me._frmGlazingNote_Toolbars_Dock_Area_Bottom.ToolbarsManager = Me.UltraToolbarsManager1
        '
        'ContextMenuStrip1
        '
        Me.ContextMenuStrip1.Name = "ContextMenuStrip1"
        Me.ContextMenuStrip1.Size = New System.Drawing.Size(61, 4)
        '
        'frmGlazingNote
        '
        Me.AcceptButton = Me.brnAdd
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.CancelButton = Me.brnAdd
        Me.ClientSize = New System.Drawing.Size(984, 432)
        Me.Controls.Add(Me.frmGlazingNote_Fill_Panel)
        Me.Controls.Add(Me._frmGlazingNote_Toolbars_Dock_Area_Left)
        Me.Controls.Add(Me._frmGlazingNote_Toolbars_Dock_Area_Right)
        Me.Controls.Add(Me._frmGlazingNote_Toolbars_Dock_Area_Bottom)
        Me.Controls.Add(Me._frmGlazingNote_Toolbars_Dock_Area_Top)
        Me.Name = "frmGlazingNote"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        CType(Me.uscGlassQuote, System.ComponentModel.ISupportInitialize).EndInit()
        Me.frmGlazingNote_Fill_Panel.ClientArea.ResumeLayout(False)
        Me.frmGlazingNote_Fill_Panel.ClientArea.PerformLayout()
        Me.frmGlazingNote_Fill_Panel.ResumeLayout(False)
        CType(Me.utxtNoteText, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.UltraToolbarsManager1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents brnAdd As System.Windows.Forms.Button
    Friend WithEvents UltraFormattedTextWordWriter1 As Infragistics.Win.UltraWinFormattedText.WordWriter.UltraFormattedTextWordWriter
    Friend WithEvents uscGlassQuote As Infragistics.Win.UltraWinSpellChecker.UltraSpellChecker
    Friend WithEvents btnSpellCheck As System.Windows.Forms.Button
    Friend WithEvents frmGlazingNote_Fill_Panel As Infragistics.Win.Misc.UltraPanel
    Friend WithEvents _frmGlazingNote_Toolbars_Dock_Area_Left As Infragistics.Win.UltraWinToolbars.UltraToolbarsDockArea
    Friend WithEvents UltraToolbarsManager1 As Infragistics.Win.UltraWinToolbars.UltraToolbarsManager
    Friend WithEvents _frmGlazingNote_Toolbars_Dock_Area_Right As Infragistics.Win.UltraWinToolbars.UltraToolbarsDockArea
    Friend WithEvents _frmGlazingNote_Toolbars_Dock_Area_Bottom As Infragistics.Win.UltraWinToolbars.UltraToolbarsDockArea
    Friend WithEvents _frmGlazingNote_Toolbars_Dock_Area_Top As Infragistics.Win.UltraWinToolbars.UltraToolbarsDockArea
    Friend WithEvents UltraFormattedTextWordWriter2 As Infragistics.Win.UltraWinFormattedText.WordWriter.UltraFormattedTextWordWriter
    Friend WithEvents ContextMenuStrip1 As System.Windows.Forms.ContextMenuStrip
    Friend WithEvents utxtNoteText As Infragistics.Win.UltraWinEditors.UltraTextEditor
End Class
